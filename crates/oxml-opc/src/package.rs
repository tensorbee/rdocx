//! OPC Package reader and writer for ZIP-based OOXML packages.

use std::collections::{HashMap, HashSet};
use std::io::{Read, Seek, SeekFrom, Write};
use std::path::Path;

use zip::ZipWriter;
use zip::read::ZipArchive;
use zip::write::SimpleFileOptions;

use crate::content_types::ContentTypes;
use crate::error::{OpcError, Result};
use crate::relationship::{Relationship, Relationships, rel_types};

const PACKAGE_RELATIONSHIPS_PATH: &str = "_rels/.rels";
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

/// A relationship part's producer bytes and the relationships parsed from them.
#[derive(Debug, Clone)]
struct RelationshipsSource {
    xml: Vec<u8>,
    items: Vec<Relationship>,
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
    /// All parts keyed by their URI (e.g. "/word/document.xml").
    pub parts: HashMap<String, Vec<u8>>,
    content_types_source: Option<ContentTypesSource>,
    /// Relationship parts as read, keyed by the ZIP entry name they are
    /// written back to.
    relationship_sources: HashMap<String, RelationshipsSource>,
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
            let name = normalize_part_name(entry.name());
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
        let mut relationship_sources = HashMap::new();
        let package_rels = if let Some(rels_xml) =
            deterministic_part_key(&raw_parts, "/_rels/.rels").and_then(|name| raw_parts.get(name))
        {
            let relationships = Relationships::from_xml(rels_xml)?;
            relationship_sources.insert(
                PACKAGE_RELATIONSHIPS_PATH.to_owned(),
                RelationshipsSource {
                    xml: rels_xml.clone(),
                    items: relationships.items.clone(),
                },
            );
            relationships
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
                relationship_sources.insert(
                    part_name_to_rels_path(&part_name),
                    RelationshipsSource {
                        xml: xml_data.clone(),
                        items: rels.items.clone(),
                    },
                );
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
            parts.insert(normalized, data.clone());
        }

        Ok(OpcPackage {
            content_types,
            package_rels,
            part_rels,
            parts,
            content_types_source: Some(ContentTypesSource {
                xml: ct_xml.clone(),
                content_types: ContentTypes::from_xml(ct_xml)?,
            }),
            relationship_sources,
        })
    }

    /// Save the OPC package to a file path.
    pub fn save<P: AsRef<Path>>(&self, path: P) -> Result<()> {
        let serialized = self.serialize()?;
        let file = std::fs::File::create(path)?;
        Self::write_serialized(file, serialized)
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
        zip.start_file(PACKAGE_RELATIONSHIPS_PATH, options)?;
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
        self.parts.get(key).map(Vec::as_slice)
    }

    /// Return whether a case-equivalent normalized part name exists.
    pub fn contains_part(&self, part_name: &str) -> bool {
        deterministic_part_key(&self.parts, part_name).is_some()
    }

    #[cfg(feature = "digital-signatures")]
    pub(crate) fn stored_part_name(&self, part_name: &str) -> Option<&str> {
        deterministic_part_key(&self.parts, part_name).map(String::as_str)
    }

    /// Set (or replace) a part's raw bytes.
    pub fn set_part(&mut self, part_name: &str, data: Vec<u8>) {
        let key = deterministic_part_key(&self.parts, part_name)
            .cloned()
            .unwrap_or_else(|| part_name.to_owned());
        self.parts.insert(key, data);
    }

    /// Return whether a stored part already says what a typed model serializes to.
    ///
    /// `reserialize` parses a part and serializes the result, which is what the
    /// stored bytes are written as when nothing changed. When that equals
    /// `serialized`, no edit reached the part and its stored bytes can stay,
    /// with their producer formatting, namespace declarations, and unmodelled
    /// content. An absent part, or bytes that no longer parse, do not match.
    pub fn part_matches_serialization(
        &self,
        part_name: &str,
        serialized: &[u8],
        reserialize: impl FnOnce(&[u8]) -> Option<Vec<u8>>,
    ) -> bool {
        self.get_part(part_name).is_some_and(|stored| {
            stored == serialized || reserialize(stored).as_deref() == Some(serialized)
        })
    }

    /// Remove and return a case-equivalent normalized part.
    pub fn remove_part(&mut self, part_name: &str) -> Option<Vec<u8>> {
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

    /// Return a relationship part's producer bytes while its relationships are
    /// unchanged, and the canonical serialization otherwise.
    fn relationships_bytes(
        &self,
        rels_path: &str,
        relationships: &Relationships,
    ) -> Result<Vec<u8>> {
        match self.relationship_sources.get(rels_path) {
            Some(source) if source.items == relationships.items => Ok(source.xml.clone()),
            _ => relationships.to_xml(),
        }
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
    /// The private key must be PKCS#8 DER and the certificate must be X.509
    /// DER. The live package changes only after the candidate verifies with
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
            relationship_sources: HashMap::new(),
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
        let package_relationships =
            self.relationships_bytes(PACKAGE_RELATIONSHIPS_PATH, &self.package_rels)?;
        let mut part_relationships = Vec::with_capacity(self.part_rels.len());
        let mut relationship_owners = self.part_rels.iter().collect::<Vec<_>>();
        relationship_owners.sort_by_key(|(owner, _)| owner.as_str());
        for (owner, relationships) in relationship_owners {
            let rels_path = part_name_to_rels_path(owner);
            let xml = self.relationships_bytes(&rels_path, relationships)?;
            part_relationships.push((rels_path, xml));
        }

        let mut parts = self
            .parts
            .iter()
            .map(|(name, data)| (name.as_str(), data.as_slice()))
            .collect::<Vec<_>>();
        parts.sort_by_key(|(name, _)| *name);
        validate_zip_entry_identities(&part_relationships, &parts)?;
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
    fn unchanged_relationship_parts_are_not_reformatted_on_save() {
        const PACKAGE_RELS: &[u8] = br#"<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Target="word/document.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Id="rId1"/></Relationships>"#;
        const DOCUMENT_RELS: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<pr:Relationships xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships" xmlns:x="urn:producer">
    <pr:Relationship Id="rId7" Type="urn:first" Target="first.xml"/>
    <pr:Relationship Id="rId3" Type="urn:link" Target="https://example.com/?a=1&amp;b=2" TargetMode="External"/>
</pr:Relationships>"#;
        let saved = |package: &OpcPackage, name: &str| {
            let mut output = std::io::Cursor::new(Vec::new());
            package.write_to(&mut output).unwrap();
            output.set_position(0);
            let mut archive = ZipArchive::new(output).unwrap();
            let mut bytes = Vec::new();
            archive
                .by_name(name)
                .unwrap()
                .read_to_end(&mut bytes)
                .unwrap();
            bytes
        };
        let mut package = OpcPackage::from_reader(package_zip(&[
            ("[Content_Types].xml", MINIMAL_CONTENT_TYPES),
            ("_rels/.rels", PACKAGE_RELS),
            ("word/_rels/document.xml.rels", DOCUMENT_RELS),
            ("word/document.xml", b"<document/>"),
            ("word/first.xml", b"<first/>"),
        ]))
        .unwrap();
        assert_eq!(saved(&package, "_rels/.rels"), PACKAGE_RELS);
        assert_eq!(
            saved(&package, "word/_rels/document.xml.rels"),
            DOCUMENT_RELS
        );

        // Removing and restoring the same relationships is not a change.
        let relationships = package.remove_part_rels("/word/document.xml").unwrap();
        package.set_part_rels("/word/document.xml", relationships);
        assert_eq!(
            saved(&package, "word/_rels/document.xml.rels"),
            DOCUMENT_RELS
        );

        // A changed relationship set is written from the model, and only that
        // relationship part moves.
        package
            .get_or_create_part_rels("/word/document.xml")
            .add_with_id("rId8", "urn:second", "second.xml");
        let rewritten = saved(&package, "word/_rels/document.xml.rels");
        assert_ne!(rewritten, DOCUMENT_RELS);
        let reparsed = Relationships::from_xml(&rewritten).unwrap();
        assert_eq!(
            reparsed
                .items
                .iter()
                .map(|relationship| relationship.id.as_str())
                .collect::<Vec<_>>(),
            ["rId7", "rId3", "rId8"]
        );
        assert_eq!(saved(&package, "_rels/.rels"), PACKAGE_RELS);

        package.package_rels.items[0].target = "word/renamed.xml".to_owned();
        let package_rels = saved(&package, "_rels/.rels");
        assert_eq!(
            Relationships::from_xml(&package_rels).unwrap().items[0].target,
            "word/renamed.xml"
        );
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
            package.remove_part("/word/document.xml"),
            Some(b"second".to_vec())
        );
    }

    #[test]
    fn direct_case_conflicts_and_duplicate_relationship_ids_fail_before_output() {
        let mut package = OpcPackage::new();
        package
            .parts
            .insert("/WORD/DOCUMENT.XML".to_owned(), b"first".to_vec());
        package
            .parts
            .insert("/word/document.xml".to_owned(), b"second".to_vec());
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
