//! OPC Package reader and writer for ZIP-based OOXML packages.

use std::collections::{HashMap, HashSet};
use std::io::{Read, Seek, SeekFrom, Write};
use std::path::{Path, PathBuf};
use std::sync::Arc;

use zip::ZipWriter;
use zip::read::ZipArchive;
use zip::write::SimpleFileOptions;

use crate::content_types::ContentTypes;
use crate::error::{OpcError, Result};
use crate::relationship::{Relationships, rel_types};

const EOCD_SIGNATURE: [u8; 4] = [0x50, 0x4b, 0x05, 0x06];
const ZIP64_EOCD_LOCATOR_SIGNATURE: [u8; 4] = [0x50, 0x4b, 0x06, 0x07];
const ZIP64_EOCD_SIGNATURE: [u8; 4] = [0x50, 0x4b, 0x06, 0x06];
const MAX_ZIP_COMMENT_LEN: u64 = u16::MAX as u64;
const EOCD_LEN: u64 = 22;
const ZIP64_EOCD_LOCATOR_LEN: u64 = 20;
const MIN_ZIP64_EOCD_LEN: u64 = 56;
const ZIP_TAIL_LEN: u64 =
    MAX_ZIP_COMMENT_LEN + EOCD_LEN + ZIP64_EOCD_LOCATOR_LEN + MIN_ZIP64_EOCD_LEN;

/// A single part within the OPC package.
#[derive(Debug, Clone)]
pub struct PackagePart {
    /// Part name (URI), e.g. "/word/document.xml"
    pub name: String,
    /// Raw bytes of the part content
    pub data: Vec<u8>,
}

#[derive(Debug, Clone)]
pub(crate) struct ContentTypesSource {
    xml: Vec<u8>,
    content_types: ContentTypes,
}

struct SerializedPackage<'a> {
    content_types: Vec<u8>,
    package_relationships: Vec<u8>,
    part_relationships: Vec<(String, Vec<u8>)>,
    parts: Vec<(&'a str, &'a [u8])>,
}

/// Resource limits applied while expanding an OPC ZIP archive.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PackageReadLimits {
    pub max_entries: usize,
    pub max_part_uncompressed_bytes: u64,
    pub max_total_uncompressed_bytes: u64,
}

impl PackageReadLimits {
    pub const UNBOUNDED: Self = Self {
        max_entries: usize::MAX,
        max_part_uncompressed_bytes: u64::MAX,
        max_total_uncompressed_bytes: u64::MAX,
    };
}

/// An in-memory representation of an OPC package (ZIP archive).
#[derive(Debug, Clone)]
pub struct OpcPackage {
    /// Content types from `[Content_Types].xml`
    pub content_types: ContentTypes,
    /// Package-level relationships from `_rels/.rels`
    pub package_rels: Relationships,
    /// Part-level relationships keyed by the part they belong to.
    /// Key is the part name (e.g. "/word/document.xml"),
    /// value is the parsed relationships.
    pub part_rels: HashMap<String, Relationships>,
    /// All parts keyed by their URI (e.g. "/word/document.xml"). The bytes
    /// are shared, so cloning a package copies pointers, not part content.
    pub parts: HashMap<String, Arc<Vec<u8>>>,
    content_types_source: Option<ContentTypesSource>,
    /// The bytes of each XML entry that was read already holding a character
    /// XML 1.0 cannot carry, keyed by ZIP entry name. Such producer bytes
    /// are written back verbatim, while no writer can add such a character.
    malformed_source_entries: HashMap<String, Vec<u8>>,
}

impl OpcPackage {
    /// Open an OPC package from a file path.
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        let file = std::fs::File::open(path)?;
        Self::from_reader(file)
    }

    /// Open an OPC package from any reader that implements Read + Seek.
    pub fn from_reader<R: Read + Seek>(reader: R) -> Result<Self> {
        Self::from_reader_with_limits(reader, PackageReadLimits::UNBOUNDED)
    }

    /// Open an agile-encrypted OPC package with no archive-expansion limit.
    #[cfg(feature = "agile-encryption")]
    pub fn from_encrypted_reader<R: Read + Seek>(reader: R, password: &str) -> Result<Self> {
        Self::from_encrypted_reader_with_limits(reader, password, PackageReadLimits::UNBOUNDED)
    }

    /// Authenticate and open an agile-encrypted OPC package with expansion limits.
    #[cfg(feature = "agile-encryption")]
    pub fn from_encrypted_reader_with_limits<R: Read + Seek>(
        reader: R,
        password: &str,
        limits: PackageReadLimits,
    ) -> Result<Self> {
        let plaintext = crate::encryption::decrypt_package(reader, password, limits)?;
        Self::from_reader_with_limits(std::io::Cursor::new(plaintext), limits)
    }

    /// Open an OPC package while bounding archive expansion.
    pub fn from_reader_with_limits<R: Read + Seek>(
        mut reader: R,
        limits: PackageReadLimits,
    ) -> Result<Self> {
        if limits.max_entries != usize::MAX {
            let entry_count = pre_index_entry_count(&mut reader)?;
            if entry_count > limits.max_entries as u64 {
                return Err(OpcError::PackageLimitExceeded {
                    kind: "entry count",
                    limit: limits.max_entries as u64,
                });
            }
        }

        let mut archive = ZipArchive::new(reader)?;
        if archive.len() > limits.max_entries {
            return Err(OpcError::PackageLimitExceeded {
                kind: "entry count",
                limit: limits.max_entries as u64,
            });
        }
        let mut raw_parts: HashMap<String, Vec<u8>> = HashMap::new();
        let mut raw_part_identities = HashSet::new();
        let mut total_uncompressed_bytes = 0_u64;

        // Read all entries from the ZIP
        for i in 0..archive.len() {
            let mut entry = archive.by_index(i)?;
            if entry.is_dir() {
                continue;
            }
            if entry.size() > limits.max_part_uncompressed_bytes {
                return Err(OpcError::PackageLimitExceeded {
                    kind: "part size",
                    limit: limits.max_part_uncompressed_bytes,
                });
            }
            if total_uncompressed_bytes
                .checked_add(entry.size())
                .is_none_or(|total| total > limits.max_total_uncompressed_bytes)
            {
                return Err(OpcError::PackageLimitExceeded {
                    kind: "total uncompressed size",
                    limit: limits.max_total_uncompressed_bytes,
                });
            }
            let name = normalize_part_name(entry.name()?.as_ref());
            let name = name.strip_prefix('/').unwrap_or(&name).to_string();
            if !raw_part_identities.insert(name.to_ascii_lowercase()) {
                return Err(OpcError::DuplicatePartName(name));
            }
            let mut data = Vec::new();
            let read_limit = limits
                .max_part_uncompressed_bytes
                .min(limits.max_total_uncompressed_bytes - total_uncompressed_bytes);
            entry
                .by_ref()
                .take(read_limit.saturating_add(1))
                .read_to_end(&mut data)?;
            let data_len = data.len() as u64;
            if data_len > limits.max_part_uncompressed_bytes {
                return Err(OpcError::PackageLimitExceeded {
                    kind: "part size",
                    limit: limits.max_part_uncompressed_bytes,
                });
            }
            total_uncompressed_bytes = total_uncompressed_bytes
                .checked_add(data_len)
                .filter(|total| *total <= limits.max_total_uncompressed_bytes)
                .ok_or(OpcError::PackageLimitExceeded {
                    kind: "total uncompressed size",
                    limit: limits.max_total_uncompressed_bytes,
                })?;
            raw_parts.insert(name, data);
        }

        // Parse [Content_Types].xml
        let ct_xml = deterministic_part_key(&raw_parts, "/[Content_Types].xml")
            .and_then(|name| raw_parts.get(name))
            .ok_or_else(|| OpcError::PartNotFound("[Content_Types].xml".into()))?;
        let content_types = ContentTypes::from_xml(ct_xml)?;

        // Parse package-level relationships: _rels/.rels
        let package_rels = if let Some(rels_xml) =
            deterministic_part_key(&raw_parts, "/_rels/.rels").and_then(|name| raw_parts.get(name))
        {
            Relationships::from_xml(rels_xml)?
        } else {
            Relationships::new()
        };

        // Parse all part-level .rels files
        let mut part_rels = HashMap::new();
        let rels_entries: Vec<String> = raw_parts
            .keys()
            .filter(|name| {
                name.to_ascii_lowercase().ends_with(".rels")
                    && part_identity(name) != part_identity("/_rels/.rels")
            })
            .cloned()
            .collect();

        for rels_path in rels_entries {
            if let Some(xml_data) = raw_parts.get(&rels_path) {
                let rels = Relationships::from_xml(xml_data)?;
                // Convert rels path to the part name it belongs to.
                // e.g. "word/_rels/document.xml.rels" → "/word/document.xml"
                let part_name = rels_path_to_part_name(&rels_path);
                part_rels.insert(part_name, rels);
            }
        }

        // Build parts map with leading "/" normalized
        let mut parts = HashMap::new();
        for (name, data) in &raw_parts {
            if part_identity(name) == part_identity("/[Content_Types].xml")
                || part_identity(name) == part_identity("/_rels/.rels")
                || name.to_ascii_lowercase().ends_with(".rels")
            {
                continue;
            }
            let normalized = if name.starts_with('/') {
                name.clone()
            } else {
                format!("/{name}")
            };
            parts.insert(normalized, Arc::new(data.clone()));
        }

        let malformed_source_entries = raw_parts
            .iter()
            .filter(|(name, data)| {
                is_xml_entry(&content_types, name) && first_invalid_xml_character(data).is_some()
            })
            .map(|(name, data)| (name.clone(), data.clone()))
            .collect();

        Ok(OpcPackage {
            content_types,
            package_rels,
            part_rels,
            parts,
            malformed_source_entries,
            content_types_source: Some(ContentTypesSource {
                xml: ct_xml.clone(),
                content_types: ContentTypes::from_xml(ct_xml)?,
            }),
        })
    }

    /// Save the OPC package to a file path.
    ///
    /// The complete ZIP is built in memory and published through
    /// [`write_atomic_file`], so a failed save leaves an existing file as it
    /// was.
    pub fn save<P: AsRef<Path>>(&self, path: P) -> Result<()> {
        let mut output = std::io::Cursor::new(Vec::new());
        self.write_to(&mut output)?;
        write_atomic_file(
            path.as_ref(),
            output.get_ref(),
            "oxml-opc",
            "invalid package file name",
            "could not allocate package save staging file",
        )?;
        Ok(())
    }

    /// Write the OPC package to any writer.
    pub fn write_to<W: Write + Seek>(&self, writer: W) -> Result<()> {
        Self::write_serialized(writer, self.serialize()?)
    }

    fn write_serialized<W: Write + Seek>(
        writer: W,
        serialized: SerializedPackage<'_>,
    ) -> Result<()> {
        let mut zip = ZipWriter::new(writer);
        let options =
            SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);

        // Write [Content_Types].xml
        zip.start_file("[Content_Types].xml", options)?;
        zip.write_all(&serialized.content_types)?;

        // Write _rels/.rels
        zip.start_file("_rels/.rels", options)?;
        zip.write_all(&serialized.package_relationships)?;

        // Write part-level .rels files. Both loops iterate in sorted order so
        // that saving the same package twice produces byte-identical output;
        // a HashMap's order would vary between runs.
        for (rels_path, rels_xml) in serialized.part_relationships {
            zip.start_file(&rels_path, options)?;
            zip.write_all(&rels_xml)?;
        }

        // Write all parts
        for (name, data) in serialized.parts {
            // Strip leading "/" for ZIP entry name
            let zip_name = name.strip_prefix('/').unwrap_or(name);
            zip.start_file(zip_name, options)?;
            zip.write_all(data)?;
        }

        zip.finish()?;
        Ok(())
    }

    /// Append a password-protected CFB envelope around the deterministic OPC ZIP.
    ///
    /// The output buffer is unchanged if package serialization, encryption, or
    /// allocation fails.
    #[cfg(feature = "agile-encryption")]
    pub fn write_encrypted_to(&self, output: &mut Vec<u8>, password: &str) -> Result<()> {
        let mut plaintext = std::io::Cursor::new(Vec::new());
        self.write_to(&mut plaintext)?;
        crate::encryption::write_encrypted_package(output, &plaintext.into_inner(), password)
    }

    /// Get raw bytes of a part by its URI.
    pub fn get_part(&self, part_name: &str) -> Option<&[u8]> {
        let key = deterministic_part_key(&self.parts, part_name)?;
        self.parts.get(key).map(|data| data.as_slice())
    }

    /// Return whether a case-equivalent normalized part name exists.
    pub fn contains_part(&self, part_name: &str) -> bool {
        deterministic_part_key(&self.parts, part_name).is_some()
    }

    #[cfg(feature = "digital-signatures")]
    pub(crate) fn stored_part_name(&self, part_name: &str) -> Option<&str> {
        deterministic_part_key(&self.parts, part_name).map(String::as_str)
    }

    /// Set (or replace) a part's raw bytes, owned or already shared.
    pub fn set_part(&mut self, part_name: &str, data: impl Into<Arc<Vec<u8>>>) {
        let key = deterministic_part_key(&self.parts, part_name)
            .cloned()
            .unwrap_or_else(|| part_name.to_owned());
        self.parts.insert(key, data.into());
    }

    /// Remove and return a case-equivalent normalized part, with the bytes
    /// it may still share with a clone of this package.
    pub fn remove_part(&mut self, part_name: &str) -> Option<Arc<Vec<u8>>> {
        let key = deterministic_part_key(&self.parts, part_name)?.clone();
        self.parts.remove(&key)
    }

    pub(crate) fn content_types_bytes(&self) -> Result<Vec<u8>> {
        self.content_types.validate_identities()?;
        if let Some(source) = &self.content_types_source
            && source.content_types == self.content_types
        {
            return Ok(source.xml.clone());
        }
        self.content_types.to_xml()
    }

    /// Verify every digital signature discovered through the OPC relationship graph.
    ///
    /// Cryptographic validity authenticates the embedded certificate's key. It does
    /// not establish that the certificate chains to a trusted root. Certificate
    /// trust remains caller policy.
    #[cfg(feature = "digital-signatures")]
    pub fn verify_signatures(&self) -> Result<Vec<crate::SignatureReport>> {
        crate::signature::verify_signatures(self)
    }

    /// Sign a cloned candidate package with the strict RSA-SHA256 profile.
    ///
    /// The private key must be PKCS#8 DER or one unencrypted PRIVATE KEY PEM
    /// block. The certificate must be X.509 DER or one CERTIFICATE PEM block. The live package changes only after the candidate verifies with
    /// complete declared coverage. Certificate-chain trust remains caller
    /// policy.
    #[cfg(feature = "digital-signatures")]
    pub fn sign(
        &mut self,
        private_key_pkcs8_der: &[u8],
        certificate_der: &[u8],
    ) -> Result<crate::SignatureReport> {
        let mut candidate = self.clone();
        let report = crate::signature::create_signature(
            &mut candidate,
            private_key_pkcs8_der,
            certificate_der,
        )?;
        *self = candidate;
        Ok(report)
    }

    /// Get the relationships for a specific part.
    pub fn get_part_rels(&self, part_name: &str) -> Option<&Relationships> {
        let key = deterministic_part_key(&self.part_rels, part_name)?;
        self.part_rels.get(key)
    }

    /// Get mutable relationships for a case-equivalent normalized owner.
    pub fn get_part_rels_mut(&mut self, part_name: &str) -> Option<&mut Relationships> {
        let key = deterministic_part_key(&self.part_rels, part_name)?.clone();
        self.part_rels.get_mut(&key)
    }

    #[cfg(feature = "digital-signatures")]
    pub(crate) fn stored_part_rels_owner(&self, part_name: &str) -> Option<&str> {
        deterministic_part_key(&self.part_rels, part_name).map(String::as_str)
    }

    /// Get or create the relationships for a specific part.
    pub fn get_or_create_part_rels(&mut self, part_name: &str) -> &mut Relationships {
        let key = deterministic_part_key(&self.part_rels, part_name)
            .cloned()
            .unwrap_or_else(|| part_name.to_owned());
        self.part_rels.entry(key).or_default()
    }

    /// Replaces the relationships for a part while preserving the package's
    /// existing spelling for a case-equivalent relationship owner.
    pub fn set_part_rels(&mut self, part_name: &str, relationships: Relationships) {
        let key = deterministic_part_key(&self.part_rels, part_name)
            .cloned()
            .unwrap_or_else(|| part_name.to_owned());
        self.part_rels.insert(key, relationships);
    }

    /// Remove and return relationships for a case-equivalent normalized owner.
    pub fn remove_part_rels(&mut self, part_name: &str) -> Option<Relationships> {
        let key = deterministic_part_key(&self.part_rels, part_name)?.clone();
        self.part_rels.remove(&key)
    }

    /// Resolve the target URI of a relationship relative to its source part.
    ///
    /// `.` and `..` segments are collapsed, so a target such as
    /// `../media/image1.png` on `/word/charts/chart1.xml` resolves to
    /// `/media/image1.png` and matches the key it is stored under. Without
    /// this, parts referenced through a parent directory could never be found.
    pub fn resolve_rel_target(source_part: &str, rel_target: &str) -> String {
        let joined = if rel_target.starts_with('/') {
            rel_target.to_string()
        } else {
            // Directory of the source part, including the trailing slash.
            let dir = match source_part.rfind('/') {
                Some(pos) => &source_part[..=pos],
                None => "/",
            };
            format!("{dir}{rel_target}")
        };
        normalize_part_name(&joined)
    }

    /// Find the main document part URI by looking at package relationships.
    pub fn main_document_part(&self) -> Option<String> {
        self.package_rels
            .get_by_type(rel_types::DOCUMENT)
            .map(|rel| {
                if rel.target.starts_with('/') {
                    rel.target.clone()
                } else {
                    format!("/{}", rel.target)
                }
            })
    }

    /// Create an empty package with the universal OPC content types.
    pub fn new() -> Self {
        OpcPackage {
            content_types: ContentTypes::minimal(),
            package_rels: Relationships::new(),
            part_rels: HashMap::new(),
            parts: HashMap::new(),
            content_types_source: None,
            malformed_source_entries: HashMap::new(),
        }
    }

    /// Create a package whose office document relationship targets a main part.
    pub fn with_main_part(part_name: &str, content_type: &str) -> Self {
        let mut package = Self::new();
        let part_name = part_name.strip_prefix('/').unwrap_or(part_name);
        package.package_rels.add(rel_types::DOCUMENT, part_name);
        package
            .content_types
            .add_override(&format!("/{part_name}"), content_type);
        package
    }

    fn serialize(&self) -> Result<SerializedPackage<'_>> {
        self.validate_graph()?;

        let content_types = self.content_types_bytes()?;
        let package_relationships = self.package_rels.to_xml()?;
        let mut part_relationships = Vec::with_capacity(self.part_rels.len());
        let mut relationship_owners = self.part_rels.iter().collect::<Vec<_>>();
        relationship_owners.sort_by_key(|(owner, _)| owner.as_str());
        for (owner, relationships) in relationship_owners {
            part_relationships.push((part_name_to_rels_path(owner), relationships.to_xml()?));
        }

        let mut parts = self
            .parts
            .iter()
            .map(|(name, data)| (name.as_str(), data.as_slice()))
            .collect::<Vec<_>>();
        parts.sort_by_key(|(name, _)| *name);
        validate_zip_entry_identities(&part_relationships, &parts)?;

        // No writer may publish a character XML 1.0 forbids, so every XML
        // entry is checked here, where every save of every format passes.
        self.reject_written_xml_characters("[Content_Types].xml", &content_types)?;
        self.reject_written_xml_characters("_rels/.rels", &package_relationships)?;
        for (rels_path, rels_xml) in &part_relationships {
            self.reject_written_xml_characters(rels_path, rels_xml)?;
        }
        for (name, data) in &parts {
            if is_xml_entry(&self.content_types, name) {
                self.reject_written_xml_characters(name, data)?;
            }
        }
        Ok(SerializedPackage {
            content_types,
            package_relationships,
            part_relationships,
            parts,
        })
    }

    pub(crate) fn validate_graph(&self) -> Result<()> {
        validate_part_map_identities(&self.parts)?;
        validate_part_map_identities(&self.part_rels)?;
        self.content_types.validate_identities()?;
        self.package_rels.validate_ids()?;
        for relationships in self.part_rels.values() {
            relationships.validate_ids()?;
        }
        Ok(())
    }
}

impl OpcPackage {
    /// Refuses an entry to be written that holds a character XML 1.0 cannot
    /// carry, unless it is the exact bytes a producer wrote into the source.
    fn reject_written_xml_characters(&self, entry: &str, data: &[u8]) -> Result<()> {
        let Some((code_point, line, column)) = first_invalid_xml_character(data) else {
            return Ok(());
        };
        let name = entry.strip_prefix('/').unwrap_or(entry);
        if self.malformed_source_entries.get(name).map(Vec::as_slice) == Some(data) {
            return Ok(());
        }
        Err(OpcError::InvalidXmlCharacter {
            part: format!("/{name}"),
            code_point,
            line,
            column,
        })
    }
}

/// Whether an entry is XML: a relationship or content types entry, or a part
/// whose content type ends in `xml`.
fn is_xml_entry(content_types: &ContentTypes, name: &str) -> bool {
    let lower = name.to_ascii_lowercase();
    lower.ends_with(".rels")
        || lower.ends_with("[content_types].xml")
        || content_types
            .content_type_for(&format!("/{}", name.strip_prefix('/').unwrap_or(name)))
            .is_some_and(|content_type| content_type.to_ascii_lowercase().ends_with("xml"))
}

/// The code point, line and column of the first character outside the XML
/// 1.0 `Char` production in a UTF-8 or UTF-16 entry. The byte scan keeps
/// producer XML available even when its decoder would reject that character.
fn first_invalid_xml_character(data: &[u8]) -> Option<(u32, usize, usize)> {
    if let Some(body) = data.strip_prefix(&[0xFF, 0xFE]) {
        return first_invalid_utf16_character(body, true);
    }
    if let Some(body) = data.strip_prefix(&[0xFE, 0xFF]) {
        return first_invalid_utf16_character(body, false);
    }
    if data.len() >= 4 && data[1] == 0 && data[3] == 0 {
        return first_invalid_utf16_character(data, true);
    }
    if data.len() >= 4 && data[0] == 0 && data[2] == 0 {
        return first_invalid_utf16_character(data, false);
    }
    let (offset, code_point) = data
        .iter()
        .enumerate()
        .find_map(|(index, byte)| match byte {
            0x00..=0x08 | 0x0B | 0x0C | 0x0E..=0x1F => Some((index, u32::from(*byte))),
            0xEF if data.get(index + 1) == Some(&0xBF) => match data.get(index + 2) {
                Some(0xBE) => Some((index, 0xFFFE)),
                Some(0xBF) => Some((index, 0xFFFF)),
                _ => None,
            },
            _ => None,
        })?;
    let body_start = if data.starts_with(&[0xEF, 0xBB, 0xBF]) {
        3
    } else {
        0
    };
    let mut line = 1;
    let mut line_start = body_start;
    let mut index = body_start;
    while index < offset {
        match data[index] {
            b'\r' => {
                line += 1;
                if data.get(index + 1) == Some(&b'\n') {
                    index += 1;
                }
                line_start = index + 1;
            }
            b'\n' => {
                line += 1;
                line_start = index + 1;
            }
            _ => {}
        }
        index += 1;
    }
    let column = data[line_start.min(offset)..offset]
        .iter()
        .filter(|byte| **byte & 0xC0 != 0x80)
        .count()
        + 1;
    Some((code_point, line, column))
}

fn first_invalid_utf16_character(data: &[u8], little_endian: bool) -> Option<(u32, usize, usize)> {
    let units = data.chunks_exact(2).map(|pair| {
        if little_endian {
            u16::from_le_bytes([pair[0], pair[1]])
        } else {
            u16::from_be_bytes([pair[0], pair[1]])
        }
    });
    let mut line = 1;
    let mut column = 1;
    let mut after_cr = false;
    for decoded in char::decode_utf16(units) {
        let code_point = match decoded {
            Ok(character) => u32::from(character),
            Err(error) => u32::from(error.unpaired_surrogate()),
        };
        if matches!(code_point, 0x00..=0x08 | 0x0B | 0x0C | 0x0E..=0x1F | 0xD800..=0xDFFF | 0xFFFE | 0xFFFF)
        {
            return Some((code_point, line, column));
        }
        match code_point {
            0x0D => {
                line += 1;
                column = 1;
                after_cr = true;
            }
            0x0A if after_cr => after_cr = false,
            0x0A => {
                line += 1;
                column = 1;
            }
            _ => {
                column += 1;
                after_cr = false;
            }
        }
    }
    None
}

pub(crate) fn part_identity(part_name: &str) -> String {
    normalize_part_name(part_name).to_ascii_lowercase()
}

fn deterministic_part_key<'a, T>(
    values: &'a HashMap<String, T>,
    part_name: &str,
) -> Option<&'a String> {
    let identity = part_identity(part_name);
    values
        .keys()
        .filter(|candidate| part_identity(candidate) == identity)
        .min()
}

fn validate_part_map_identities<T>(values: &HashMap<String, T>) -> Result<()> {
    let mut identities = HashSet::with_capacity(values.len());
    let mut names = values.keys().collect::<Vec<_>>();
    names.sort_unstable();
    for name in names {
        if !identities.insert(part_identity(name)) {
            return Err(OpcError::DuplicatePartName(name.clone()));
        }
    }
    Ok(())
}

fn validate_zip_entry_identities(
    part_relationships: &[(String, Vec<u8>)],
    parts: &[(&str, &[u8])],
) -> Result<()> {
    let mut identities = HashSet::from([
        part_identity("/[Content_Types].xml"),
        part_identity("/_rels/.rels"),
    ]);
    for (name, _) in part_relationships {
        if !identities.insert(part_identity(name)) {
            return Err(OpcError::DuplicatePartName(name.clone()));
        }
    }
    for (name, _) in parts {
        if !identities.insert(part_identity(name)) {
            return Err(OpcError::DuplicatePartName((*name).to_owned()));
        }
    }
    Ok(())
}

fn pre_index_entry_count<R: Read + Seek>(reader: &mut R) -> Result<u64> {
    let file_len = reader.seek(SeekFrom::End(0))?;
    if file_len < EOCD_LEN {
        return Err(invalid_zip("end of central directory not found"));
    }

    let tail_len = file_len.min(ZIP_TAIL_LEN);
    let tail_start = file_len - tail_len;
    reader.seek(SeekFrom::Start(tail_start))?;
    let mut tail = vec![0; tail_len as usize];
    reader.read_exact(&mut tail)?;

    let last_candidate = tail.len() - EOCD_LEN as usize;
    let entry_count = (0..=last_candidate)
        .rev()
        .filter(|&offset| tail[offset..].starts_with(&EOCD_SIGNATURE))
        .filter_map(|offset| eocd_entry_count(&tail, tail_start, file_len, offset))
        .max();

    entry_count.ok_or_else(|| invalid_zip("valid end of central directory not found"))
}

fn eocd_entry_count(tail: &[u8], tail_start: u64, file_len: u64, offset: usize) -> Option<u64> {
    let eocd = tail.get(offset..offset.checked_add(EOCD_LEN as usize)?)?;
    let comment_len = u64::from(read_u16(eocd, 20)?);
    let absolute_offset = tail_start.checked_add(offset as u64)?;
    if absolute_offset
        .checked_add(EOCD_LEN)?
        .checked_add(comment_len)?
        != file_len
    {
        return None;
    }

    let disk_number = read_u16(eocd, 4)?;
    let central_directory_disk = read_u16(eocd, 6)?;
    if disk_number != 0 || central_directory_disk != 0 {
        return None;
    }

    let entries_on_disk = read_u16(eocd, 8)?;
    let total_entries = read_u16(eocd, 10)?;
    let central_directory_size = read_u32(eocd, 12)?;
    let central_directory_offset = read_u32(eocd, 16)?;
    let uses_zip64 = entries_on_disk == u16::MAX
        || total_entries == u16::MAX
        || central_directory_size == u32::MAX
        || central_directory_offset == u32::MAX;

    if uses_zip64 {
        return zip64_entry_count(tail, tail_start, absolute_offset);
    }
    if entries_on_disk != total_entries
        || !central_directory_bounds_are_plausible(
            absolute_offset,
            u64::from(central_directory_offset),
            u64::from(central_directory_size),
        )
    {
        return None;
    }

    Some(u64::from(total_entries))
}

fn zip64_entry_count(tail: &[u8], tail_start: u64, eocd_offset: u64) -> Option<u64> {
    let locator_offset = eocd_offset.checked_sub(ZIP64_EOCD_LOCATOR_LEN)?;
    let locator = tail_slice(tail, tail_start, locator_offset, ZIP64_EOCD_LOCATOR_LEN)?;
    if !locator.starts_with(&ZIP64_EOCD_LOCATOR_SIGNATURE)
        || read_u32(locator, 4)? != 0
        || read_u32(locator, 16)? != 1
    {
        return None;
    }

    let zip64_offset = read_u64(locator, 8)?;
    let zip64 = tail_slice(tail, tail_start, zip64_offset, MIN_ZIP64_EOCD_LEN)?;
    if !zip64.starts_with(&ZIP64_EOCD_SIGNATURE) {
        return None;
    }
    let record_size = read_u64(zip64, 4)?;
    if record_size < 44
        || zip64_offset.checked_add(12)?.checked_add(record_size)? != locator_offset
        || read_u32(zip64, 16)? != 0
        || read_u32(zip64, 20)? != 0
    {
        return None;
    }

    let entries_on_disk = read_u64(zip64, 24)?;
    let total_entries = read_u64(zip64, 32)?;
    if entries_on_disk != total_entries
        || !central_directory_bounds_are_plausible(
            zip64_offset,
            read_u64(zip64, 48)?,
            read_u64(zip64, 40)?,
        )
    {
        return None;
    }

    Some(total_entries)
}

fn central_directory_bounds_are_plausible(
    directory_end: u64,
    relative_directory_offset: u64,
    directory_size: u64,
) -> bool {
    directory_end
        .checked_sub(directory_size)
        .is_some_and(|actual_directory_offset| relative_directory_offset <= actual_directory_offset)
}

fn tail_slice(tail: &[u8], tail_start: u64, absolute_offset: u64, len: u64) -> Option<&[u8]> {
    let offset = usize::try_from(absolute_offset.checked_sub(tail_start)?).ok()?;
    let len = usize::try_from(len).ok()?;
    tail.get(offset..offset.checked_add(len)?)
}

fn read_u16(bytes: &[u8], offset: usize) -> Option<u16> {
    Some(u16::from_le_bytes(
        bytes.get(offset..offset.checked_add(2)?)?.try_into().ok()?,
    ))
}

fn read_u32(bytes: &[u8], offset: usize) -> Option<u32> {
    Some(u32::from_le_bytes(
        bytes.get(offset..offset.checked_add(4)?)?.try_into().ok()?,
    ))
}

fn read_u64(bytes: &[u8], offset: usize) -> Option<u64> {
    Some(u64::from_le_bytes(
        bytes.get(offset..offset.checked_add(8)?)?.try_into().ok()?,
    ))
}

fn invalid_zip(detail: &'static str) -> OpcError {
    OpcError::Zip(zip::result::ZipError::InvalidArchive(detail.into()))
}

impl Default for OpcPackage {
    fn default() -> Self {
        Self::new()
    }
}

/// Replace the file at `path` with `bytes` through a synced sibling file.
///
/// The bytes are written to `.{file name}.{staging_tag}-{process id}-{attempt}.tmp`
/// beside the file being replaced, synced, and renamed over it, so readers see
/// the old file or the new one and never a partial write. A file name too long
/// for that staging name is shortened in it. Any failure removes the staged
/// file and leaves the destination as it was. A symbolic link at `path` is
/// followed and kept, and the file it names is replaced. On Unix the
/// replacement keeps the permission bits of the file it replaces, from the
/// moment the staged file is created. The owner, extended attributes, and
/// other hard links of that file are not carried over. A file this process
/// cannot open for writing is refused and left as it was.
///
/// A path that names a device or a FIFO, such as `/dev/null`, has no file to
/// replace. It takes the bytes in place, as a plain write would.
///
/// `staging_tag` is a fragment of a file name and must not contain a path
/// separator.
///
/// # Errors
///
/// Returns `invalid_name_message` when `path` names no file, the error of
/// opening the file to replace for writing, such as
/// [`PermissionDenied`](std::io::ErrorKind::PermissionDenied) when it is
/// read-only, an
/// [`InvalidInput`](std::io::ErrorKind::InvalidInput) error when `path` goes
/// through more than 40 symbolic links, `exhausted_message` when every staging
/// name is taken, and the I/O error of a failed write, staging, sync, or
/// rename.
pub fn write_atomic_file(
    path: &Path,
    bytes: &[u8],
    staging_tag: &str,
    invalid_name_message: &'static str,
    exhausted_message: &'static str,
) -> std::io::Result<()> {
    // Renaming over a device or a FIFO would put a regular file in its place.
    // The kernel follows every link here, including /dev/stdout.
    if std::fs::metadata(path).is_ok_and(|metadata| !metadata.is_file() && !metadata.is_dir()) {
        return std::fs::write(path, bytes);
    }
    let path = resolve_symbolic_links(path)?;
    let parent = path.parent().unwrap_or_else(|| Path::new("."));
    let file_name = path.file_name().ok_or_else(|| {
        std::io::Error::new(std::io::ErrorKind::InvalidInput, invalid_name_message)
    })?;
    let existing = std::fs::metadata(&path).ok();
    // A rename needs write access to the directory, not to the file it
    // replaces. Opening the file for writing, without truncating it, refuses
    // the files the in-place write refused, such as a read-only file or
    // another user's file.
    if existing.as_ref().is_some_and(std::fs::Metadata::is_file) {
        std::fs::OpenOptions::new().write(true).open(&path)?;
    }
    // Only Unix mode bits carry over. On Windows the staged file keeps its
    // default attributes.
    let permissions = existing
        .filter(|_| cfg!(unix))
        .map(|metadata| metadata.permissions());
    let mut options = std::fs::OpenOptions::new();
    options.write(true).create_new(true);
    // Creating the staged file with the kept mode means it is never more open
    // than the file it replaces. The umask can narrow that mode, so it is set
    // again once the file exists.
    #[cfg(unix)]
    if let Some(permissions) = &permissions {
        use std::os::unix::fs::{OpenOptionsExt as _, PermissionsExt as _};
        options.mode(permissions.mode());
    }
    // wasm targets have no process id, and std panics when asked for one. The
    // retry below covers two processes that share a staging name.
    let process_id = if cfg!(target_family = "wasm") {
        0
    } else {
        std::process::id()
    };
    for attempt in 0..128_u8 {
        let temporary = parent.join(staging_file_name(
            file_name,
            &format!(".{staging_tag}-{process_id}-{attempt}.tmp"),
        ));
        let mut file = match options.open(&temporary) {
            Ok(file) => file,
            Err(error) if error.kind() == std::io::ErrorKind::AlreadyExists => continue,
            Err(error) => return Err(error),
        };
        let result = permissions
            .clone()
            .map_or(Ok(()), |permissions| file.set_permissions(permissions))
            .and_then(|()| file.write_all(bytes))
            .and_then(|()| file.sync_all());
        drop(file);
        let result = result.and_then(|()| replace_file(&temporary, &path));
        if result.is_err() {
            let _ = std::fs::remove_file(&temporary);
        }
        return result;
    }
    Err(std::io::Error::new(
        std::io::ErrorKind::AlreadyExists,
        exhausted_message,
    ))
}

/// Name a staged file `.{file name}{suffix}`, keeping it within 255 bytes.
///
/// Common file systems refuse a longer name, so a long file name keeps only
/// the prefix that fits, cut on a character boundary.
fn staging_file_name(file_name: &std::ffi::OsStr, suffix: &str) -> std::ffi::OsString {
    let available = 255_usize.saturating_sub(1 + suffix.len());
    let mut name = std::ffi::OsString::from(".");
    if file_name.len() <= available {
        name.push(file_name);
    } else {
        let file_name = file_name.to_string_lossy();
        let mut end = available.min(file_name.len());
        while !file_name.is_char_boundary(end) {
            end -= 1;
        }
        name.push(&file_name[..end]);
    }
    name.push(suffix);
    name
}

/// Follow symbolic links from `path` to the file a save replaces.
///
/// A relative link target resolves against the directory holding the link,
/// and a dangling link resolves to the missing file it names.
fn resolve_symbolic_links(path: &Path) -> std::io::Result<PathBuf> {
    let mut resolved = path.to_path_buf();
    // Up to 40 links are followed, the Linux limit, so 41 paths are checked.
    // A longer chain is treated as a loop.
    for _ in 0..=40 {
        match std::fs::symlink_metadata(&resolved) {
            Ok(metadata) if metadata.file_type().is_symlink() => {
                let target = std::fs::read_link(&resolved)?;
                resolved = match resolved.parent() {
                    Some(parent) => parent.join(target),
                    None => target,
                };
            }
            _ => return Ok(resolved),
        }
    }
    Err(std::io::Error::new(
        std::io::ErrorKind::InvalidInput,
        "too many levels of symbolic links",
    ))
}

#[cfg(not(target_os = "windows"))]
fn replace_file(source: &Path, destination: &Path) -> std::io::Result<()> {
    std::fs::rename(source, destination)
}

#[cfg(target_os = "windows")]
fn replace_file(source: &Path, destination: &Path) -> std::io::Result<()> {
    use std::os::windows::ffi::OsStrExt;

    const MOVEFILE_REPLACE_EXISTING: u32 = 0x1;
    const MOVEFILE_WRITE_THROUGH: u32 = 0x8;

    #[link(name = "kernel32")]
    unsafe extern "system" {
        fn MoveFileExW(
            existing_file_name: *const u16,
            new_file_name: *const u16,
            flags: u32,
        ) -> i32;
    }

    let source: Vec<u16> = source.as_os_str().encode_wide().chain(Some(0)).collect();
    let destination: Vec<u16> = destination
        .as_os_str()
        .encode_wide()
        .chain(Some(0))
        .collect();
    // SAFETY: both path buffers are NUL-terminated and remain alive for the call.
    let replaced = unsafe {
        MoveFileExW(
            source.as_ptr(),
            destination.as_ptr(),
            MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH,
        )
    };
    if replaced == 0 {
        Err(std::io::Error::last_os_error())
    } else {
        Ok(())
    }
}

#[cfg(test)]
fn docx_package() -> OpcPackage {
    let mut package = OpcPackage::with_main_part(
        "word/document.xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
    );
    package.content_types.add_override(
        "/word/styles.xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
    );
    package
}

/// Collapse `.` and `..` segments in an absolute part name.
///
/// A `..` that would escape the package root is dropped, matching how OPC
/// consumers treat over-long parent traversals.
fn normalize_part_name(path: &str) -> String {
    let mut segments: Vec<&str> = Vec::new();
    for segment in path.split('/') {
        match segment {
            "" | "." => {}
            ".." => {
                segments.pop();
            }
            other => segments.push(other),
        }
    }

    let mut out = String::with_capacity(path.len());
    for segment in segments {
        out.push('/');
        out.push_str(segment);
    }
    if out.is_empty() { "/".to_string() } else { out }
}

/// Convert a .rels file path to the part name it belongs to.
/// e.g. "word/_rels/document.xml.rels" → "/word/document.xml"
fn rels_path_to_part_name(rels_path: &str) -> String {
    // Strip the ".rels" suffix once (`trim_end_matches` would strip repeats,
    // turning "a.rels.rels" into "a") and drop only the "_rels" path segment
    // that directly precedes the file name.
    let lowercase = rels_path.to_ascii_lowercase();
    let without_suffix = if lowercase.ends_with(".rels") {
        &rels_path[..rels_path.len() - ".rels".len()]
    } else {
        rels_path
    };
    let lowercase = without_suffix.to_ascii_lowercase();
    let path = match lowercase.rfind("_rels/") {
        Some(pos) => format!(
            "{}{}",
            &without_suffix[..pos],
            &without_suffix[pos + "_rels/".len()..]
        ),
        None => without_suffix.to_string(),
    };
    if path.starts_with('/') {
        path
    } else {
        format!("/{path}")
    }
}

/// Convert a part name to its .rels file path.
/// e.g. "/word/document.xml" → "word/_rels/document.xml.rels"
fn part_name_to_rels_path(part_name: &str) -> String {
    let name = part_name.strip_prefix('/').unwrap_or(part_name);
    if let Some(pos) = name.rfind('/') {
        let dir = &name[..pos];
        let file = &name[pos + 1..];
        format!("{dir}/_rels/{file}.rels")
    } else {
        format!("_rels/{name}.rels")
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn package_zip(entries: &[(&str, &[u8])]) -> std::io::Cursor<Vec<u8>> {
        package_zip_with_comment(entries, &[])
    }

    fn package_zip_with_comment(
        entries: &[(&str, &[u8])],
        comment: &[u8],
    ) -> std::io::Cursor<Vec<u8>> {
        let mut buffer = std::io::Cursor::new(Vec::new());
        {
            let mut zip = ZipWriter::new(&mut buffer);
            let options = SimpleFileOptions::default();
            for (name, data) in entries {
                zip.start_file(name, options).unwrap();
                zip.write_all(data).unwrap();
            }
            zip.set_raw_comment(comment.to_vec().into_boxed_slice())
                .unwrap();
            zip.finish().unwrap();
        }
        buffer.set_position(0);
        buffer
    }

    fn replace_zip_name(bytes: &mut [u8], old: &[u8], raw: &[u8], utf8: bool) {
        assert_eq!(old.len(), raw.len());
        let mut changed = 0;
        for offset in 0..bytes.len().saturating_sub(46) {
            let (header, flags, length) = match &bytes[offset..offset + 4] {
                b"PK\x03\x04" => (30, 6, 26),
                b"PK\x01\x02" => (46, 8, 28),
                _ => continue,
            };
            let name_len = u16::from_le_bytes(
                bytes[offset + length..offset + length + 2]
                    .try_into()
                    .unwrap(),
            ) as usize;
            let start = offset + header;
            if name_len != old.len() || bytes.get(start..start + name_len) != Some(old) {
                continue;
            }
            bytes[start..start + name_len].copy_from_slice(raw);
            let mut value = u16::from_le_bytes(
                bytes[offset + flags..offset + flags + 2]
                    .try_into()
                    .unwrap(),
            );
            value = (value & !(1 << 11)) | if utf8 { 1 << 11 } else { 0 };
            bytes[offset + flags..offset + flags + 2].copy_from_slice(&value.to_le_bytes());
            changed += 1;
        }
        assert_eq!(changed, 2, "both local and central names must be patched");
    }

    #[test]
    fn zip9_utf8_and_cp437_names_preserve_opaque_payloads() {
        for (raw, utf8) in [
            (b"x\xc3\xa9.bin".as_slice(), false),
            (b"x\xc3\xa9.bin".as_slice(), true),
            (b"x\x82.bin".as_slice(), false),
            (b"x\x82.bin".as_slice(), true),
        ] {
            let name = format!("{}.bin", "a".repeat(raw.len() - 4));
            let mut bytes = package_zip(&[
                ("[Content_Types].xml", br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="bin" ContentType="application/octet-stream"/></Types>"#),
                (&name, b"opaque\0\xff"),
            ]).into_inner();
            replace_zip_name(&mut bytes, name.as_bytes(), raw, utf8);
            let before = bytes.clone();
            let package = OpcPackage::from_reader(std::io::Cursor::new(&bytes)).unwrap();
            assert_eq!(
                package.get_part("/xé.bin"),
                Some(b"opaque\0\xff".as_slice())
            );
            let mut saved = std::io::Cursor::new(Vec::new());
            package.write_to(&mut saved).unwrap();
            let reopened =
                OpcPackage::from_reader(std::io::Cursor::new(saved.into_inner())).unwrap();
            assert_eq!(
                reopened.get_part("/xé.bin"),
                Some(b"opaque\0\xff".as_slice())
            );
            assert_eq!(bytes, before);
        }
    }

    #[test]
    fn zip9_decoded_aliases_reach_existing_duplicate_rejection() {
        let mut bytes = package_zip(&[
            (
                "[Content_Types].xml",
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>"#,
            ),
            ("aa.bin", b"first"),
            ("aaa.bin", b"second"),
        ])
        .into_inner();
        replace_zip_name(&mut bytes, b"aa.bin", b"x\x82.bin", false);
        replace_zip_name(&mut bytes, b"aaa.bin", b"x\xc3\xa9.bin", false);
        let before = bytes.clone();
        assert!(
            matches!(OpcPackage::from_reader(std::io::Cursor::new(&bytes)),
            Err(OpcError::DuplicatePartName(name)) if name == "xé.bin")
        );
        assert_eq!(bytes, before);
    }

    #[test]
    fn zip9_ordinary_writer_header_is_deterministic_version_twenty() {
        let mut package = OpcPackage::new();
        package.set_part("opaque.bin", b"opaque\0\xff".to_vec());
        let mut first = std::io::Cursor::new(Vec::new());
        let mut second = std::io::Cursor::new(Vec::new());
        package.write_to(&mut first).unwrap();
        package.write_to(&mut second).unwrap();
        assert_eq!(first.get_ref(), second.get_ref());
        let bytes = first.get_ref();
        let mut central = 0;
        for offset in 0..bytes.len().saturating_sub(46) {
            if &bytes[offset..offset + 4] != b"PK\x01\x02" {
                continue;
            }
            let read =
                |at| u16::from_le_bytes(bytes[offset + at..offset + at + 2].try_into().unwrap());
            assert_eq!(read(4) & 0xff, 20);
            assert_eq!(read(6), 20);
            assert_eq!(read(12), 0);
            assert_eq!(read(14), 33);
            assert_eq!(read(10), 8);
            central += 1;
        }
        assert_eq!(central, 3);
        let reopened = OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
        assert_eq!(
            reopened.get_part("opaque.bin"),
            Some(b"opaque\0\xff".as_slice())
        );
    }

    fn fake_eocd(total_entries: u16) -> [u8; EOCD_LEN as usize] {
        let mut eocd = [0; EOCD_LEN as usize];
        eocd[..4].copy_from_slice(&EOCD_SIGNATURE);
        eocd[8..10].copy_from_slice(&total_entries.to_le_bytes());
        eocd[10..12].copy_from_slice(&total_entries.to_le_bytes());
        eocd
    }

    const MINIMAL_CONTENT_TYPES: &[u8] =
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
</Types>"#;

    #[test]
    fn unchanged_case_variant_content_types_preserve_producer_bytes_and_order() {
        const PRODUCER_CONTENT_TYPES: &[u8] = br#"<?xml version="1.0"?><ct:Types xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types"><ct:Override ContentType="application/example+xml" PartName="/WORD/DOCUMENT.XML"/><ct:Default ContentType="image/png" Extension="PNG"/><ct:Default ContentType="application/xml" Extension="XML"/></ct:Types>"#;
        let package = OpcPackage::from_reader(package_zip(&[
            ("[Content_Types].xml", PRODUCER_CONTENT_TYPES),
            ("word/document.xml", b"<document/>"),
        ]))
        .unwrap();
        assert_eq!(
            package.content_types.content_type_for("/word/document.xml"),
            Some("application/example+xml")
        );
        assert_eq!(
            package.content_types.content_type_for("/word/image.png"),
            Some("image/png")
        );

        let mut output = std::io::Cursor::new(Vec::new());
        package.write_to(&mut output).unwrap();
        output.set_position(0);
        let mut archive = ZipArchive::new(output).unwrap();
        let mut saved = Vec::new();
        archive
            .by_name("[Content_Types].xml")
            .unwrap()
            .read_to_end(&mut saved)
            .unwrap();
        assert_eq!(saved, PRODUCER_CONTENT_TYPES);
    }

    #[test]
    fn package_write_rejects_direct_content_type_case_conflicts() {
        let mut package = OpcPackage::new();
        package
            .content_types
            .defaults
            .insert("PNG".to_string(), "image/first".to_string());
        package
            .content_types
            .defaults
            .insert("png".to_string(), "image/second".to_string());
        let mut output = std::io::Cursor::new(Vec::new());

        assert!(matches!(
            package.write_to(&mut output),
            Err(OpcError::InvalidContentTypes)
        ));
        assert!(output.into_inner().is_empty());
    }

    #[test]
    fn part_access_is_case_insensitive_and_preserves_first_spelling() {
        let mut package = OpcPackage::new();
        package.set_part("/WORD/DOCUMENT.XML", b"first".to_vec());
        package.set_part("word/document.xml", b"second".to_vec());
        assert_eq!(package.parts.len(), 1);
        assert!(package.parts.contains_key("/WORD/DOCUMENT.XML"));
        assert!(package.contains_part("/word/./document.xml"));
        assert_eq!(package.get_part("word/document.xml"), Some(&b"second"[..]));

        package
            .get_or_create_part_rels("/WORD/DOCUMENT.XML")
            .add("urn:test", "target.xml");
        assert_eq!(package.part_rels.len(), 1);
        assert!(
            package
                .get_part_rels("word/document.xml")
                .is_some_and(|relationships| relationships.items.len() == 1)
        );
        assert!(package.remove_part_rels("/word/document.xml").is_some());
        assert_eq!(
            package.remove_part("/word/document.xml").as_deref(),
            Some(&b"second".to_vec())
        );
    }

    #[test]
    fn direct_case_conflicts_and_duplicate_relationship_ids_fail_before_output() {
        let mut package = OpcPackage::new();
        package
            .parts
            .insert("/WORD/DOCUMENT.XML".to_owned(), b"first".to_vec().into());
        package
            .parts
            .insert("/word/document.xml".to_owned(), b"second".to_vec().into());
        let mut output = std::io::Cursor::new(Vec::new());
        assert!(matches!(
            package.write_to(&mut output),
            Err(OpcError::DuplicatePartName(_))
        ));
        assert!(output.into_inner().is_empty());

        let mut package = OpcPackage::new();
        for target in ["first.xml", "second.xml"] {
            package.package_rels.items.push(crate::Relationship {
                id: "producer-id".to_owned(),
                rel_type: "urn:test".to_owned(),
                target: target.to_owned(),
                target_mode: None,
            });
        }
        let mut output = std::io::Cursor::new(Vec::new());
        assert!(matches!(
            package.write_to(&mut output),
            Err(OpcError::InvalidRelationship)
        ));
        assert!(output.into_inner().is_empty());
    }

    #[test]
    fn save_validates_before_truncating_destination() {
        let destination = std::env::temp_dir().join(format!(
            "rdocx-opc-save-sentinel-{}-{}",
            std::process::id(),
            std::thread::current().name().unwrap_or("test")
        ));
        std::fs::write(&destination, b"sentinel").unwrap();

        let mut package = OpcPackage::new();
        for target in ["first.xml", "second.xml"] {
            package.package_rels.items.push(crate::Relationship {
                id: "duplicate".to_owned(),
                rel_type: "urn:test".to_owned(),
                target: target.to_owned(),
                target_mode: None,
            });
        }
        assert!(matches!(
            package.save(&destination),
            Err(OpcError::InvalidRelationship)
        ));
        assert_eq!(std::fs::read(&destination).unwrap(), b"sentinel");
        std::fs::remove_file(destination).unwrap();
    }

    fn save_test_directory(label: &str) -> PathBuf {
        let directory =
            std::env::temp_dir().join(format!("oxml-opc-save-{label}-{}", std::process::id()));
        let _ = std::fs::remove_dir_all(&directory);
        std::fs::create_dir_all(&directory).unwrap();
        directory
    }

    fn staging_files(directory: &Path) -> Vec<PathBuf> {
        std::fs::read_dir(directory)
            .unwrap()
            .map(|entry| entry.unwrap().path())
            .filter(|path| path.extension().is_some_and(|extension| extension == "tmp"))
            .collect()
    }

    fn package_bytes(package: &OpcPackage) -> Vec<u8> {
        let mut output = std::io::Cursor::new(Vec::new());
        package.write_to(&mut output).unwrap();
        output.into_inner()
    }

    #[cfg(unix)]
    #[test]
    fn save_replaces_an_existing_file_by_rename_and_keeps_its_mode() {
        use std::os::unix::fs::{MetadataExt as _, PermissionsExt as _};

        let directory = save_test_directory("rename");
        let destination = directory.join("existing.docx");
        std::fs::write(&destination, b"previous bytes").unwrap();
        std::fs::set_permissions(&destination, std::fs::Permissions::from_mode(0o600)).unwrap();
        let previous_inode = std::fs::metadata(&destination).unwrap().ino();

        let package = docx_package();
        package.save(&destination).unwrap();

        // A new inode shows the file was replaced by rename, not rewritten in place.
        let metadata = std::fs::metadata(&destination).unwrap();
        assert_ne!(metadata.ino(), previous_inode);
        assert_eq!(metadata.permissions().mode() & 0o7777, 0o600);
        assert_eq!(
            std::fs::read(&destination).unwrap(),
            package_bytes(&package)
        );
        assert!(staging_files(&directory).is_empty());
        std::fs::remove_dir_all(directory).unwrap();
    }

    #[cfg(unix)]
    #[test]
    fn save_through_a_symbolic_link_replaces_the_file_it_names() {
        use std::os::unix::fs::PermissionsExt as _;

        let directory = save_test_directory("symlink");
        let links = directory.join("links");
        std::fs::create_dir(&links).unwrap();
        let target = directory.join("target.docx");
        std::fs::write(&target, b"previous bytes").unwrap();
        std::fs::set_permissions(&target, std::fs::Permissions::from_mode(0o640)).unwrap();
        let link = links.join("link.docx");
        std::os::unix::fs::symlink("../target.docx", &link).unwrap();

        let package = docx_package();
        package.save(&link).unwrap();

        assert!(
            std::fs::symlink_metadata(&link)
                .unwrap()
                .file_type()
                .is_symlink()
        );
        assert_eq!(
            std::fs::read_link(&link).unwrap(),
            Path::new("../target.docx")
        );
        assert_eq!(std::fs::read(&target).unwrap(), package_bytes(&package));
        assert_eq!(
            std::fs::metadata(&target).unwrap().permissions().mode() & 0o7777,
            0o640
        );

        // A dangling link names the file the save creates.
        std::fs::remove_file(&target).unwrap();
        package.save(&link).unwrap();
        assert!(
            std::fs::symlink_metadata(&link)
                .unwrap()
                .file_type()
                .is_symlink()
        );
        assert_eq!(std::fs::read(&target).unwrap(), package_bytes(&package));

        // A chain of 40 links is followed, and one more link is refused as a loop.
        std::fs::remove_file(&target).unwrap();
        let mut head = target.clone();
        for index in 0..40 {
            let chained = links.join(format!("chain-{index}.docx"));
            std::os::unix::fs::symlink(&head, &chained).unwrap();
            head = chained;
        }
        package.save(&head).unwrap();
        assert_eq!(std::fs::read(&target).unwrap(), package_bytes(&package));
        let too_long = links.join("chain-40.docx");
        std::os::unix::fs::symlink(&head, &too_long).unwrap();
        assert!(package.save(&too_long).is_err());
        assert!(staging_files(&directory).is_empty());
        assert!(staging_files(&links).is_empty());
        std::fs::remove_dir_all(directory).unwrap();
    }

    #[test]
    fn failed_save_keeps_the_destination_and_leaves_no_staging_file() {
        let directory = save_test_directory("failure");

        // Renaming the staged file over a directory fails after staging.
        let occupied = directory.join("directory.docx");
        std::fs::create_dir(&occupied).unwrap();
        std::fs::write(occupied.join("inner.xml"), b"inner bytes").unwrap();
        assert!(docx_package().save(&occupied).is_err());
        assert_eq!(
            std::fs::read(occupied.join("inner.xml")).unwrap(),
            b"inner bytes"
        );
        assert!(staging_files(&directory).is_empty());

        // Exhausted staging names fail before the destination is touched.
        let destination = directory.join("existing.docx");
        std::fs::write(&destination, b"previous bytes").unwrap();
        for attempt in 0..128_u8 {
            let staging = directory.join(format!(
                ".existing.docx.oxml-opc-{}-{attempt}.tmp",
                std::process::id()
            ));
            std::fs::write(staging, b"occupied").unwrap();
        }
        let error = docx_package().save(&destination).unwrap_err();
        assert!(
            error
                .to_string()
                .contains("could not allocate package save staging file"),
            "{error}"
        );
        assert_eq!(std::fs::read(&destination).unwrap(), b"previous bytes");
        assert_eq!(staging_files(&directory).len(), 128);
        std::fs::remove_dir_all(directory).unwrap();
    }

    #[cfg(unix)]
    #[test]
    fn save_to_a_device_or_fifo_writes_in_place_and_keeps_the_node() {
        use std::os::unix::fs::FileTypeExt as _;

        let package = docx_package();
        package.save("/dev/null").unwrap();
        assert!(
            std::fs::metadata("/dev/null")
                .unwrap()
                .file_type()
                .is_char_device()
        );

        let directory = save_test_directory("fifo");
        let fifo = directory.join("pipe.docx");
        let status = std::process::Command::new("mkfifo")
            .arg(&fifo)
            .status()
            .unwrap();
        assert!(status.success());
        let reader = {
            let fifo = fifo.clone();
            std::thread::spawn(move || std::fs::read(fifo).unwrap())
        };
        package.save(&fifo).unwrap();
        // Checked before the join, so a save that replaced the FIFO fails here
        // instead of leaving the test waiting on a reader that never gets data.
        assert!(
            std::fs::symlink_metadata(&fifo)
                .unwrap()
                .file_type()
                .is_fifo()
        );
        assert_eq!(reader.join().unwrap(), package_bytes(&package));
        assert!(staging_files(&directory).is_empty());
        std::fs::remove_dir_all(directory).unwrap();
    }

    #[cfg(unix)]
    #[test]
    fn save_refuses_a_read_only_file_and_leaves_it_as_it_was() {
        use std::os::unix::fs::PermissionsExt as _;

        let directory = save_test_directory("read-only");
        let destination = directory.join("protected.docx");
        std::fs::write(&destination, b"previous bytes").unwrap();
        // 0o464 leaves write bits for the group but none for the owner, who
        // cannot write the file either.
        for mode in [0o444, 0o464] {
            std::fs::set_permissions(&destination, std::fs::Permissions::from_mode(mode)).unwrap();
            // Root writes a read-only file in place, so it may replace it too.
            if std::fs::OpenOptions::new()
                .write(true)
                .open(&destination)
                .is_ok()
            {
                continue;
            }
            let error = docx_package().save(&destination).unwrap_err();
            assert!(
                matches!(&error, OpcError::Io(error) if error.kind() == std::io::ErrorKind::PermissionDenied),
                "{mode:o}: {error}"
            );
            assert_eq!(std::fs::read(&destination).unwrap(), b"previous bytes");
        }
        assert!(staging_files(&directory).is_empty());
        std::fs::remove_dir_all(directory).unwrap();
    }

    // Unix only: on Windows the full path passes MAX_PATH, and the MoveFileExW
    // rename takes it without the `\\?\` prefix that std adds for long paths.
    #[cfg(unix)]
    #[test]
    fn save_shortens_the_staging_name_of_a_long_file_name() {
        let directory = save_test_directory("long-name");
        // A name of 255 bytes in 235 characters fills the limit of ext4 and APFS.
        // The shortened staging name ends inside the run of three-byte
        // characters, so it has to cut on a character boundary.
        let destination =
            directory.join(format!("{}{}.docx", "a".repeat(220), "\u{6587}".repeat(10)));
        let package = docx_package();
        package.save(&destination).unwrap();
        std::fs::write(&destination, b"previous bytes").unwrap();
        package.save(&destination).unwrap();
        assert_eq!(
            std::fs::read(&destination).unwrap(),
            package_bytes(&package)
        );
        assert!(staging_files(&directory).is_empty());
        std::fs::remove_dir_all(directory).unwrap();
    }

    #[test]
    fn bounded_reader_rejects_too_many_entries() {
        let archive = package_zip(&[
            ("[Content_Types].xml", MINIMAL_CONTENT_TYPES),
            ("word/document.xml", b"<document/>"),
        ]);
        let error = OpcPackage::from_reader_with_limits(
            archive,
            PackageReadLimits {
                max_entries: 1,
                max_part_uncompressed_bytes: 1_024,
                max_total_uncompressed_bytes: 2_048,
            },
        )
        .unwrap_err();

        assert!(matches!(
            error,
            OpcError::PackageLimitExceeded {
                kind: "entry count",
                limit: 1
            }
        ));
    }

    #[test]
    fn later_eocd_signature_in_comment_cannot_hide_real_entry_count() {
        let mut archive = package_zip_with_comment(
            &[
                ("[Content_Types].xml", MINIMAL_CONTENT_TYPES),
                ("word/document.xml", b"<document/>"),
            ],
            &fake_eocd(1),
        );

        assert_eq!(pre_index_entry_count(&mut archive).unwrap(), 2);
    }

    #[test]
    fn truncated_zip64_metadata_returns_an_error_without_panicking() {
        let mut bytes = vec![0; 4 + ZIP64_EOCD_LOCATOR_LEN as usize + EOCD_LEN as usize];
        bytes[..4].copy_from_slice(&ZIP64_EOCD_SIGNATURE);

        let locator_offset = 4;
        bytes[locator_offset..locator_offset + 4].copy_from_slice(&ZIP64_EOCD_LOCATOR_SIGNATURE);
        bytes[locator_offset + 8..locator_offset + 16].copy_from_slice(&0_u64.to_le_bytes());
        bytes[locator_offset + 16..locator_offset + 20].copy_from_slice(&1_u32.to_le_bytes());

        let eocd_offset = locator_offset + ZIP64_EOCD_LOCATOR_LEN as usize;
        bytes[eocd_offset..eocd_offset + 4].copy_from_slice(&EOCD_SIGNATURE);
        bytes[eocd_offset + 8..eocd_offset + 12].fill(0xff);
        bytes[eocd_offset + 12..eocd_offset + 20].fill(0xff);

        let error = pre_index_entry_count(&mut std::io::Cursor::new(bytes)).unwrap_err();
        assert!(matches!(error, OpcError::Zip(_)));
    }

    #[test]
    fn bounded_reader_rejects_oversized_parts_and_totals() {
        let entries = [
            ("[Content_Types].xml", MINIMAL_CONTENT_TYPES),
            ("word/document.xml", b"<document/>".as_slice()),
        ];
        let part_error = OpcPackage::from_reader_with_limits(
            package_zip(&entries),
            PackageReadLimits {
                max_entries: 8,
                max_part_uncompressed_bytes: 16,
                max_total_uncompressed_bytes: 1_024,
            },
        )
        .unwrap_err();
        assert!(matches!(
            part_error,
            OpcError::PackageLimitExceeded {
                kind: "part size",
                limit: 16
            }
        ));

        let total_error = OpcPackage::from_reader_with_limits(
            package_zip(&entries),
            PackageReadLimits {
                max_entries: 8,
                max_part_uncompressed_bytes: 1_024,
                max_total_uncompressed_bytes: MINIMAL_CONTENT_TYPES.len() as u64,
            },
        )
        .unwrap_err();
        assert!(matches!(
            total_error,
            OpcError::PackageLimitExceeded {
                kind: "total uncompressed size",
                ..
            }
        ));
    }

    fn independently_built_pptx() -> std::io::Cursor<Vec<u8>> {
        const CONTENT_TYPES: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>"#;
        const PACKAGE_RELS: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>"#;
        const PRESENTATION_RELS: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>"#;
        const SLIDE_RELS: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>"#;
        const SLIDE_LAYOUT_RELS: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>"#;
        const SLIDE_MASTER_RELS: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>"#;
        const PRESENTATION: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
  <p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>
  <p:sldSz cx="12192000" cy="6858000"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>"#;
        const SLIDE: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>"#;
        const SLIDE_LAYOUT: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>"#;
        const SLIDE_MASTER: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>
  <p:clrMap accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" bg1="lt1" bg2="lt2" folHlink="folHlink" hlink="hlink" tx1="dk1" tx2="dk2"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483648" r:id="rId1"/></p:sldLayoutIdLst>
  <p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles>
</p:sldMaster>"#;
        const THEME: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Minimal">
  <a:themeElements>
    <a:clrScheme name="Minimal">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1F497D"/></a:dk2>
      <a:lt2><a:srgbClr val="EEECE1"/></a:lt2>
      <a:accent1><a:srgbClr val="4F81BD"/></a:accent1>
      <a:accent2><a:srgbClr val="C0504D"/></a:accent2>
      <a:accent3><a:srgbClr val="9BBB59"/></a:accent3>
      <a:accent4><a:srgbClr val="8064A2"/></a:accent4>
      <a:accent5><a:srgbClr val="4BACC6"/></a:accent5>
      <a:accent6><a:srgbClr val="F79646"/></a:accent6>
      <a:hlink><a:srgbClr val="0000FF"/></a:hlink>
      <a:folHlink><a:srgbClr val="800080"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Minimal">
      <a:majorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
      <a:minorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Minimal">
      <a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>
      <a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst>
      <a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
      <a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>"#;

        package_zip(&[
            ("[Content_Types].xml", CONTENT_TYPES),
            ("_rels/.rels", PACKAGE_RELS),
            ("ppt/presentation.xml", PRESENTATION),
            ("ppt/_rels/presentation.xml.rels", PRESENTATION_RELS),
            ("ppt/slides/slide1.xml", SLIDE),
            ("ppt/slides/_rels/slide1.xml.rels", SLIDE_RELS),
            ("ppt/slideLayouts/slideLayout1.xml", SLIDE_LAYOUT),
            (
                "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
                SLIDE_LAYOUT_RELS,
            ),
            ("ppt/slideMasters/slideMaster1.xml", SLIDE_MASTER),
            (
                "ppt/slideMasters/_rels/slideMaster1.xml.rels",
                SLIDE_MASTER_RELS,
            ),
            ("ppt/theme/theme1.xml", THEME),
        ])
    }

    fn pptx_package() -> OpcPackage {
        let mut package =
            OpcPackage::with_main_part("ppt/presentation.xml", crate::content_types::PRESENTATION);
        package
            .content_types
            .add_override("/ppt/slides/slide1.xml", crate::content_types::SLIDE);
        package.content_types.add_override(
            "/ppt/slideLayouts/slideLayout1.xml",
            crate::content_types::SLIDE_LAYOUT,
        );
        package.set_part("/ppt/presentation.xml", b"<p:presentation/>".to_vec());
        package.set_part("/ppt/slides/slide1.xml", b"<p:sld/>".to_vec());
        package.set_part(
            "/ppt/slideLayouts/slideLayout1.xml",
            b"<p:sldLayout/>".to_vec(),
        );
        package
            .get_or_create_part_rels("/ppt/presentation.xml")
            .add(rel_types::SLIDE, "slides/slide1.xml");
        package
            .get_or_create_part_rels("/ppt/slides/slide1.xml")
            .add(rel_types::SLIDE_LAYOUT, "../slideLayouts/slideLayout1.xml");
        package
    }

    #[test]
    fn a_part_holding_a_character_xml_cannot_carry_is_not_written() {
        let mut package = OpcPackage::with_main_part(
            "ppt/presentation.xml",
            "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
        );
        package
            .content_types
            .add_default("bin", "application/octet-stream");
        package.set_part("/ppt/presentation.xml", b"<p:presentation/>".to_vec());
        package.set_part("/ppt/data.bin", vec![0, 1, 0x0b, 0xff]);
        let utf16 = "<a/>"
            .encode_utf16()
            .flat_map(u16::to_le_bytes)
            .collect::<Vec<_>>();
        package.set_part("/customXml/item1.xml", [&[0xFF, 0xFE][..], &utf16].concat());
        let utf16_big_endian = "<a/>"
            .encode_utf16()
            .flat_map(u16::to_be_bytes)
            .collect::<Vec<_>>();
        package.set_part("/customXml/item2.xml", utf16_big_endian.clone());
        let mut buffer = std::io::Cursor::new(Vec::new());
        package
            .write_to(&mut buffer)
            .expect("binary and valid UTF-16 XML remain writable");
        for (data, point) in [
            (
                "<a>\u{1}</a>"
                    .encode_utf16()
                    .flat_map(u16::to_be_bytes)
                    .collect::<Vec<_>>(),
                1,
            ),
            (
                [
                    &[0xFF, 0xFE][..],
                    &"<a>\u{ffff}</a>"
                        .encode_utf16()
                        .flat_map(u16::to_le_bytes)
                        .collect::<Vec<_>>(),
                ]
                .concat(),
                0xFFFF,
            ),
        ] {
            package.set_part("/customXml/item2.xml", data);
            assert!(matches!(
                package.write_to(std::io::Cursor::new(Vec::new())),
                Err(OpcError::InvalidXmlCharacter { part, code_point, line: 1, column: 4 })
                    if part == "/customXml/item2.xml" && code_point == point
            ));
        }
        package.set_part("/customXml/item2.xml", utf16_big_endian);

        for (text, code_point, line, column) in [
            ("<p:presentation>\n  <a>x\u{1}</a>", 0x1, 2, 7),
            ("<p:presentation name=\"\u{b}\"/>", 0xB, 1, 23),
            ("<p:presentation>\u{ffff}</p:presentation>", 0xFFFF, 1, 17),
            ("\u{feff}<p:presentation>\r\n\r<a>\u{1}</a>", 0x1, 3, 4),
            ("<p:presentation>\r<é>\u{c}</é>", 0xC, 2, 4),
        ] {
            package.set_part("/ppt/presentation.xml", text.as_bytes().to_vec());
            let error = package
                .write_to(std::io::Cursor::new(Vec::new()))
                .expect_err("a forbidden character must not be written");
            assert!(
                matches!(
                    &error,
                    OpcError::InvalidXmlCharacter { part, code_point: found, line: l, column: c }
                        if part == "/ppt/presentation.xml"
                            && *found == code_point
                            && *l == line
                            && *c == column
                ),
                "{error}"
            );
        }
        assert_eq!(
            package
                .write_to(std::io::Cursor::new(Vec::new()))
                .unwrap_err()
                .to_string(),
            "/ppt/presentation.xml holds U+000C at line 2, column 4, a character XML 1.0 cannot carry"
        );

        package.set_part("/ppt/presentation.xml", b"<p:presentation/>".to_vec());
        package
            .get_or_create_part_rels("/ppt/presentation.xml")
            .add("urn:type", "target\u{1}.xml");
        assert_eq!(
            package
                .write_to(std::io::Cursor::new(Vec::new()))
                .unwrap_err()
                .to_string(),
            "/ppt/_rels/presentation.xml.rels holds U+0001 at line 3, column 57, a character XML 1.0 cannot carry"
        );
        package.part_rels.clear();

        // A producer part that already held one is written back verbatim,
        // and only a writer's change to it is refused.
        let producer = "<p:presentation>producer\u{1}</p:presentation>".as_bytes();
        package.set_part("/ppt/presentation.xml", b"<p:presentation/>".to_vec());
        let mut valid = std::io::Cursor::new(Vec::new());
        package.write_to(&mut valid).unwrap();
        let mut source = zip::ZipArchive::new(std::io::Cursor::new(valid.into_inner())).unwrap();
        let mut rebuilt = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
        for index in 0..source.len() {
            let mut entry = source.by_index(index).unwrap();
            let name = entry.name().unwrap().into_owned();
            let mut data = Vec::new();
            std::io::Read::read_to_end(&mut entry, &mut data).unwrap();
            if name == "ppt/presentation.xml" {
                data = producer.to_vec();
            }
            rebuilt
                .start_file(name, SimpleFileOptions::default())
                .unwrap();
            rebuilt.write_all(&data).unwrap();
        }
        let source = rebuilt.finish().unwrap().into_inner();
        let mut reopened = OpcPackage::from_reader(std::io::Cursor::new(source)).unwrap();
        let mut saved = std::io::Cursor::new(Vec::new());
        reopened
            .write_to(&mut saved)
            .expect("producer bytes pass through");
        saved.set_position(0);
        assert_eq!(
            OpcPackage::from_reader(saved)
                .unwrap()
                .get_part("/ppt/presentation.xml"),
            Some(producer)
        );
        reopened.set_part(
            "/ppt/presentation.xml",
            "<p:presentation>written\u{1}</p:presentation>"
                .as_bytes()
                .to_vec(),
        );
        assert!(matches!(
            reopened.write_to(std::io::Cursor::new(Vec::new())),
            Err(OpcError::InvalidXmlCharacter { code_point: 1, .. })
        ));
    }

    #[test]
    fn with_main_part_resolves_and_round_trips() {
        let mut package = OpcPackage::with_main_part(
            "ppt/presentation.xml",
            "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
        );
        package.set_part("/ppt/presentation.xml", b"<p:presentation/>".to_vec());

        assert_eq!(
            package.main_document_part().as_deref(),
            Some("/ppt/presentation.xml")
        );
        assert_eq!(
            package
                .content_types
                .content_type_for("/ppt/presentation.xml"),
            Some(
                "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
            )
        );

        let mut buffer = std::io::Cursor::new(Vec::new());
        package.write_to(&mut buffer).unwrap();
        buffer.set_position(0);
        let round_tripped = OpcPackage::from_reader(buffer).unwrap();

        assert_eq!(
            round_tripped.main_document_part().as_deref(),
            Some("/ppt/presentation.xml")
        );
        assert_eq!(
            round_tripped.get_part("/ppt/presentation.xml"),
            Some(b"<p:presentation/>".as_slice())
        );
    }

    #[test]
    fn pptx_package_resolves_main_slide_and_layout_parts() {
        let package = pptx_package();
        let mut buffer = std::io::Cursor::new(Vec::new());
        package.write_to(&mut buffer).unwrap();
        buffer.set_position(0);
        let reopened = OpcPackage::from_reader(buffer).unwrap();

        let presentation_part = reopened.main_document_part().unwrap();
        assert_eq!(presentation_part, "/ppt/presentation.xml");

        let slide_target = &reopened
            .get_part_rels(&presentation_part)
            .unwrap()
            .get_by_type(rel_types::SLIDE)
            .unwrap()
            .target;
        let slide_part = OpcPackage::resolve_rel_target(&presentation_part, slide_target);
        assert_eq!(slide_part, "/ppt/slides/slide1.xml");

        let layout_target = &reopened
            .get_part_rels(&slide_part)
            .unwrap()
            .get_by_type(rel_types::SLIDE_LAYOUT)
            .unwrap()
            .target;
        let layout_part = OpcPackage::resolve_rel_target(&slide_part, layout_target);
        assert_eq!(layout_part, "/ppt/slideLayouts/slideLayout1.xml");

        assert!(reopened.parts.contains_key(&presentation_part));
        assert!(reopened.parts.contains_key(&slide_part));
        assert!(reopened.parts.contains_key(&layout_part));
    }

    #[test]
    fn independently_built_pptx_opens_and_resolves_relationships() {
        let package = OpcPackage::from_reader(independently_built_pptx()).unwrap();

        let presentation_part = package.main_document_part().unwrap();
        assert_eq!(presentation_part, "/ppt/presentation.xml");

        let presentation_master_target = &package
            .get_part_rels(&presentation_part)
            .unwrap()
            .get_by_type(rel_types::SLIDE_MASTER)
            .unwrap()
            .target;
        let presentation_master_part =
            OpcPackage::resolve_rel_target(&presentation_part, presentation_master_target);
        assert_eq!(
            presentation_master_part,
            "/ppt/slideMasters/slideMaster1.xml"
        );

        let slide_target = &package
            .get_part_rels(&presentation_part)
            .unwrap()
            .get_by_type(rel_types::SLIDE)
            .unwrap()
            .target;
        let slide_part = OpcPackage::resolve_rel_target(&presentation_part, slide_target);
        assert_eq!(slide_part, "/ppt/slides/slide1.xml");

        let layout_target = &package
            .get_part_rels(&slide_part)
            .unwrap()
            .get_by_type(rel_types::SLIDE_LAYOUT)
            .unwrap()
            .target;
        let layout_part = OpcPackage::resolve_rel_target(&slide_part, layout_target);
        assert_eq!(layout_part, "/ppt/slideLayouts/slideLayout1.xml");

        let master_target = &package
            .get_part_rels(&layout_part)
            .unwrap()
            .get_by_type(rel_types::SLIDE_MASTER)
            .unwrap()
            .target;
        let master_part = OpcPackage::resolve_rel_target(&layout_part, master_target);
        assert_eq!(master_part, "/ppt/slideMasters/slideMaster1.xml");
        assert_eq!(master_part, presentation_master_part);

        let master_layout_target = &package
            .get_part_rels(&master_part)
            .unwrap()
            .get_by_type(rel_types::SLIDE_LAYOUT)
            .unwrap()
            .target;
        let master_layout_part = OpcPackage::resolve_rel_target(&master_part, master_layout_target);
        assert_eq!(master_layout_part, layout_part);

        let theme_target = &package
            .get_part_rels(&master_part)
            .unwrap()
            .get_by_type(rel_types::THEME)
            .unwrap()
            .target;
        let theme_part = OpcPackage::resolve_rel_target(&master_part, theme_target);
        assert_eq!(theme_part, "/ppt/theme/theme1.xml");

        let master_xml = std::str::from_utf8(package.parts.get(&master_part).unwrap()).unwrap();
        let layout_id = master_xml
            .split_once("<p:sldLayoutId id=\"")
            .unwrap()
            .1
            .split_once('"')
            .unwrap()
            .0
            .parse::<u64>()
            .unwrap();
        assert!(layout_id >= 2_147_483_648);

        assert!(package.parts.contains_key(&presentation_part));
        assert!(package.parts.contains_key(&slide_part));
        assert!(package.parts.contains_key(&layout_part));
        assert!(package.parts.contains_key(&master_part));
        assert!(package.parts.contains_key(&theme_part));
    }

    #[test]
    fn presentation_layout_target_resolves_one_directory_up() {
        assert_eq!(
            OpcPackage::resolve_rel_target(
                "/ppt/slides/slide1.xml",
                "../slideLayouts/slideLayout1.xml"
            ),
            "/ppt/slideLayouts/slideLayout1.xml"
        );
    }

    #[test]
    fn rels_path_conversion() {
        assert_eq!(
            rels_path_to_part_name("word/_rels/document.xml.rels"),
            "/word/document.xml"
        );
        assert_eq!(
            part_name_to_rels_path("/word/document.xml"),
            "word/_rels/document.xml.rels"
        );
    }

    #[test]
    fn resolve_relative_target() {
        assert_eq!(
            OpcPackage::resolve_rel_target("/word/document.xml", "styles.xml"),
            "/word/styles.xml"
        );
        assert_eq!(
            OpcPackage::resolve_rel_target("/word/document.xml", "/word/styles.xml"),
            "/word/styles.xml"
        );
    }

    #[test]
    fn resolve_target_collapses_parent_segments() {
        // Charts and headers routinely reference media through a parent dir.
        assert_eq!(
            OpcPackage::resolve_rel_target("/word/charts/chart1.xml", "../media/image1.png"),
            "/word/media/image1.png"
        );
        assert_eq!(
            OpcPackage::resolve_rel_target("/word/document.xml", "./styles.xml"),
            "/word/styles.xml"
        );
        assert_eq!(
            OpcPackage::resolve_rel_target("/word/document.xml", "../../../etc/passwd"),
            "/etc/passwd"
        );
    }

    #[test]
    fn zip_entry_that_escapes_root_is_clamped_to_root() {
        let archive = package_zip(&[
            ("[Content_Types].xml", MINIMAL_CONTENT_TYPES),
            ("../../etc/passwd", b"not a real password file"),
        ]);

        let package = OpcPackage::from_reader(archive).unwrap();

        assert_eq!(
            package.get_part("/etc/passwd"),
            Some(b"not a real password file".as_slice())
        );
        assert!(package.parts.keys().all(|name| !name.contains("..")));
    }

    #[test]
    fn duplicate_normalized_part_names_are_rejected() {
        for alias in [
            "word/./document.xml",
            "word//document.xml",
            "WORD/DOCUMENT.XML",
        ] {
            let archive = package_zip(&[
                ("[Content_Types].xml", MINIMAL_CONTENT_TYPES),
                ("word/document.xml", b"first"),
                (alias, b"second"),
            ]);

            let error = OpcPackage::from_reader(archive).unwrap_err();
            assert!(matches!(
                error,
                OpcError::DuplicatePartName(name) if name.eq_ignore_ascii_case("word/document.xml")
            ));
        }
    }

    #[test]
    fn loader_resolves_case_equivalent_special_parts_and_relationship_owners() {
        let content_types = br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/WORD/DOCUMENT.XML" ContentType="application/xml"/></Types>"#;
        let package_relationships = format!(
            r#"<Relationships xmlns="{}"><Relationship Id="rId1" Type="{}" Target="WORD/DOCUMENT.XML"/></Relationships>"#,
            "http://schemas.openxmlformats.org/package/2006/relationships",
            rel_types::DOCUMENT
        );
        let part_relationships = format!(
            r#"<Relationships xmlns="{}"><Relationship Id="rId1" Type="urn:test" Target="TARGET.XML"/></Relationships>"#,
            "http://schemas.openxmlformats.org/package/2006/relationships"
        );
        let package = OpcPackage::from_reader(package_zip(&[
            ("[CONTENT_TYPES].XML", content_types),
            ("_RELS/.RELS", package_relationships.as_bytes()),
            ("WORD/DOCUMENT.XML", b"<document/>"),
            (
                "WORD/_RELS/DOCUMENT.XML.RELS",
                part_relationships.as_bytes(),
            ),
        ]))
        .unwrap();

        assert_eq!(
            package.get_part("/word/document.xml"),
            Some(&b"<document/>"[..])
        );
        assert_eq!(package.package_rels.items.len(), 1);
        assert_eq!(
            package
                .get_part_rels("/word/document.xml")
                .unwrap()
                .items
                .len(),
            1
        );
        assert!(package.parts.contains_key("/WORD/DOCUMENT.XML"));
    }

    #[test]
    fn absolute_zip_entry_is_normalized_to_package_root() {
        let archive = package_zip(&[
            ("/[Content_Types].xml", MINIMAL_CONTENT_TYPES),
            ("/absolute/path.xml", b"<absolute/>"),
        ]);

        let package = OpcPackage::from_reader(archive).unwrap();

        assert_eq!(
            package.get_part("/absolute/path.xml"),
            Some(b"<absolute/>".as_slice())
        );
        assert_eq!(
            package
                .parts
                .keys()
                .filter(|name| name.as_str() == "/absolute/path.xml")
                .count(),
            1
        );
    }

    #[test]
    fn rels_path_suffix_is_stripped_once() {
        assert_eq!(
            rels_path_to_part_name("word/_rels/document.xml.rels"),
            "/word/document.xml"
        );
        // A part whose own name ends in ".rels" must keep that segment.
        assert_eq!(
            rels_path_to_part_name("word/_rels/odd.rels.rels"),
            "/word/odd.rels"
        );
    }

    #[test]
    fn saved_packages_are_byte_identical() {
        let mut pkg = docx_package();
        for i in 0..40 {
            pkg.set_part(&format!("/word/media/image{i}.png"), vec![i as u8]);
        }
        pkg.get_or_create_part_rels("/word/document.xml")
            .add(rel_types::STYLES, "styles.xml");

        let write = || {
            let mut buf = std::io::Cursor::new(Vec::new());
            pkg.write_to(&mut buf).unwrap();
            buf.into_inner()
        };
        assert_eq!(write(), write());
    }

    #[test]
    fn new_docx_package() {
        let pkg = docx_package();
        assert!(pkg.main_document_part().is_some());
        assert_eq!(pkg.main_document_part().unwrap(), "/word/document.xml");
    }

    #[test]
    fn round_trip_package() {
        let mut pkg = docx_package();
        pkg.set_part("/word/document.xml", b"<document/>".to_vec());

        // Write to memory
        let mut buf = std::io::Cursor::new(Vec::new());
        pkg.write_to(&mut buf).unwrap();

        // Read back
        buf.set_position(0);
        let pkg2 = OpcPackage::from_reader(buf).unwrap();
        assert_eq!(
            pkg2.get_part("/word/document.xml"),
            Some(b"<document/>".as_slice())
        );
        assert!(pkg2.main_document_part().is_some());
    }
}
