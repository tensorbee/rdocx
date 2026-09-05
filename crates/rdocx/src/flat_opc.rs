//! Strict Flat OPC import and deterministic export for Word packages.

use std::collections::{HashMap, HashSet};
use std::io::Read;
use std::path::Path;

use base64::Engine;
use base64::engine::general_purpose::STANDARD as BASE64;
use oxml_core::xml::validate_strict_xml_1_0;
use oxml_opc::content_types::{self, ContentTypes};
use oxml_opc::relationship::{Relationships, rel_types};
use oxml_opc::{OpcPackage, PackageReadLimits};
use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use quick_xml::{Writer, XmlVersion};

use crate::document::{Document, write_atomic_file};
use crate::error::{Error, Result};

const PACKAGE_NAMESPACE: &[u8] = b"http://schemas.microsoft.com/office/2006/xmlPackage";
const RELATIONSHIPS_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/package/2006/relationships";
const DEFAULT_LIMITS: PackageReadLimits = PackageReadLimits {
    max_entries: 4_096,
    max_part_uncompressed_bytes: 64 * 1024 * 1024,
    max_total_uncompressed_bytes: 128 * 1024 * 1024,
};

struct FlatPart {
    name: String,
    content_type: String,
    data: Vec<u8>,
    is_xml: bool,
}

impl Document {
    /// Import a bounded Flat OPC XML package.
    pub fn from_flat_opc_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_flat_opc_bytes_with_limits(bytes, DEFAULT_LIMITS)
    }

    /// Import Flat OPC XML while bounding part count and decoded payload sizes.
    pub fn from_flat_opc_bytes_with_limits(
        bytes: &[u8],
        limits: PackageReadLimits,
    ) -> Result<Self> {
        validate_strict_xml_1_0(bytes)
            .map_err(|error| invalid(format!("invalid XML lexical form: {error:?}")))?;
        let package = parse_flat_package(bytes, limits)?;
        let loses_signed_content_types = package
            .package_rels
            .items
            .iter()
            .any(|relationship| relationship.rel_type == rel_types::DIGITAL_SIGNATURE_ORIGIN);
        let mut document = Self::from_package(package)?;
        document.package_signatures_invalidated |= loses_signed_content_types;
        Ok(document)
    }

    /// Open a bounded Flat OPC XML package from a path.
    pub fn open_flat_opc<P: AsRef<Path>>(path: P) -> Result<Self> {
        let file = std::fs::File::open(path)?;
        let encoded_limit = DEFAULT_LIMITS
            .max_total_uncompressed_bytes
            .saturating_mul(2)
            .saturating_add(1024 * 1024);
        let mut bytes = Vec::new();
        file.take(encoded_limit.saturating_add(1))
            .read_to_end(&mut bytes)?;
        if bytes.len() as u64 > encoded_limit {
            return Err(invalid("Flat OPC input exceeds the encoded-size limit"));
        }
        Self::from_flat_opc_bytes_with_limits(&bytes, DEFAULT_LIMITS)
    }

    /// Serialize a staged copy as deterministic Flat OPC XML.
    pub fn to_flat_opc_bytes(&self) -> Result<Vec<u8>> {
        let mut candidate = self.clone_for_staging();
        candidate.package_signatures_invalidated |=
            candidate.retained_package_signature_would_be_invalidated()?;
        candidate.flush_to_package()?;
        crate::embedded::persist_invalidated_package_signature(
            &mut candidate.package,
            candidate.package_signatures_invalidated,
        )?;
        let bytes = write_flat_package(&candidate.package)?;
        Self::from_flat_opc_bytes(&bytes)?;
        Ok(bytes)
    }

    /// Save a staged copy as Flat OPC XML using atomic replacement.
    pub fn save_flat_opc<P: AsRef<Path>>(&self, path: P) -> Result<()> {
        let bytes = self.to_flat_opc_bytes()?;
        write_atomic_file(
            path.as_ref(),
            &bytes,
            "invalid Flat OPC file name",
            "could not allocate Flat OPC save staging file",
        )?;
        Ok(())
    }
}

fn parse_flat_package(bytes: &[u8], limits: PackageReadLimits) -> Result<OpcPackage> {
    let mut reader = NsReader::from_reader(bytes);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut parts = Vec::new();
    let mut seen_names = HashSet::new();
    let mut saw_root = false;
    let mut closed_root = false;
    let mut total_size = 0_u64;

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| invalid(format!("invalid Flat OPC XML: {error}")))?;
        match event {
            Event::Decl(_) if !saw_root => {}
            Event::Start(ref element) if !saw_root => {
                require_element(namespace, element.local_name().as_ref(), b"package")?;
                require_no_semantic_attributes(&reader, element)?;
                saw_root = true;
            }
            Event::Start(ref element) if saw_root && !closed_root => {
                require_element(namespace, element.local_name().as_ref(), b"part")?;
                if parts.len() >= limits.max_entries {
                    return Err(limit_error("entry count", limits.max_entries as u64));
                }
                let (name, content_type) = part_attributes(&reader, element)?;
                validate_part_name(&name)?;
                if !seen_names.insert(name.clone()) {
                    return Err(invalid(format!("duplicate Flat OPC part name {name}")));
                }
                let is_xml = xml_content_type(&content_type);
                let data = read_part_payload(&mut reader, is_xml, limits, total_size)?;
                total_size = total_size
                    .checked_add(data.len() as u64)
                    .filter(|size| *size <= limits.max_total_uncompressed_bytes)
                    .ok_or_else(|| {
                        limit_error(
                            "total uncompressed size",
                            limits.max_total_uncompressed_bytes,
                        )
                    })?;
                parts.push(FlatPart {
                    name,
                    content_type,
                    data,
                    is_xml,
                });
            }
            Event::End(ref element) if saw_root && !closed_root => {
                require_element(namespace, element.local_name().as_ref(), b"package")?;
                closed_root = true;
            }
            Event::Text(ref text) if text_is_whitespace(text) => {}
            Event::Comment(_) | Event::PI(_) => {}
            Event::Eof => break,
            _ => return Err(invalid("unexpected content in Flat OPC package")),
        }
        buffer.clear();
    }
    if !saw_root || !closed_root {
        return Err(invalid("Flat OPC package root is missing or unclosed"));
    }
    parts_to_package(parts)
}

fn read_part_payload(
    reader: &mut NsReader<&[u8]>,
    expected_xml: bool,
    limits: PackageReadLimits,
    total_size: u64,
) -> Result<Vec<u8>> {
    let mut buffer = Vec::new();
    let mut payload = None;
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| invalid(format!("invalid Flat OPC part: {error}")))?;
        match event {
            Event::Start(ref element) if payload.is_none() => {
                let wanted = if expected_xml {
                    b"xmlData".as_slice()
                } else {
                    b"binaryData".as_slice()
                };
                require_element(namespace, element.local_name().as_ref(), wanted)?;
                require_no_semantic_attributes(reader, element)?;
                let remaining = limits
                    .max_total_uncompressed_bytes
                    .checked_sub(total_size)
                    .ok_or_else(|| {
                        limit_error(
                            "total uncompressed size",
                            limits.max_total_uncompressed_bytes,
                        )
                    })?;
                let payload_limit = limits.max_part_uncompressed_bytes.min(remaining);
                payload = Some(if expected_xml {
                    read_xml_data(reader, payload_limit)?
                } else {
                    read_binary_data(reader, payload_limit)?
                });
            }
            Event::Empty(ref element) if payload.is_none() && !expected_xml => {
                require_element(namespace, element.local_name().as_ref(), b"binaryData")?;
                require_no_semantic_attributes(reader, element)?;
                payload = Some(Vec::new());
            }
            Event::End(ref element) => {
                require_element(namespace, element.local_name().as_ref(), b"part")?;
                return payload.ok_or_else(|| invalid("Flat OPC part has no data element"));
            }
            Event::Text(ref text) if text_is_whitespace(text) => {}
            Event::Comment(_) | Event::PI(_) => {}
            Event::Eof => return Err(invalid("unclosed Flat OPC part")),
            _ => {
                return Err(invalid(
                    "Flat OPC part must contain exactly one data element",
                ));
            }
        }
        if payload.as_ref().is_some_and(|data| {
            total_size.saturating_add(data.len() as u64) > limits.max_total_uncompressed_bytes
        }) {
            return Err(limit_error(
                "total uncompressed size",
                limits.max_total_uncompressed_bytes,
            ));
        }
        buffer.clear();
    }
}

fn read_xml_data(reader: &mut NsReader<&[u8]>, limit: u64) -> Result<Vec<u8>> {
    let mut output = Writer::new(Vec::new());
    let mut buffer = Vec::new();
    let mut depth = 0_usize;
    let mut roots = 0_usize;
    let mut previous_position = reader.buffer_position();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| invalid(format!("invalid Flat OPC xmlData: {error}")))?;
        let package_namespace = namespace_is_package(namespace);
        let position = reader.buffer_position();
        let encoded_event_size = position.saturating_sub(previous_position);
        previous_position = position;
        match event {
            Event::Start(ref element) if depth == 0 => {
                roots += 1;
                depth = 1;
                write_bounded_xml_event(&mut output, event, encoded_event_size, limit)?;
            }
            Event::Start(_) => {
                depth += 1;
                write_bounded_xml_event(&mut output, event, encoded_event_size, limit)?;
            }
            Event::Empty(_) if depth == 0 => {
                roots += 1;
                write_bounded_xml_event(&mut output, event, encoded_event_size, limit)?;
            }
            Event::Empty(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::PI(_)
                if depth > 0 =>
            {
                write_bounded_xml_event(&mut output, event, encoded_event_size, limit)?
            }
            Event::End(ref element) if depth == 0 => {
                if !package_namespace || element.local_name().as_ref() != b"xmlData" {
                    return Err(invalid("expected pkg:xmlData close element"));
                }
                if roots != 1 {
                    return Err(invalid(
                        "Flat OPC xmlData must contain exactly one XML root",
                    ));
                }
                let data = output.into_inner();
                enforce_part_size(data.len() as u64, limit)?;
                validate_strict_xml_1_0(&data).map_err(|error| {
                    invalid(format!("invalid part XML lexical form: {error:?}"))
                })?;
                return Ok(data);
            }
            Event::End(_) => {
                depth -= 1;
                write_bounded_xml_event(&mut output, event, encoded_event_size, limit)?;
            }
            Event::Text(ref text) if depth == 0 && text_is_whitespace(text) => {
                write_bounded_xml_event(&mut output, event, encoded_event_size, limit)?;
            }
            Event::Comment(_) | Event::PI(_) if depth == 0 => {
                write_bounded_xml_event(&mut output, event, encoded_event_size, limit)?;
            }
            Event::Decl(_) | Event::DocType(_) => {
                return Err(invalid(
                    "Flat OPC xmlData forbids declarations and doctypes",
                ));
            }
            Event::Eof => return Err(invalid("unclosed Flat OPC xmlData")),
            _ => return Err(invalid("Flat OPC xmlData has content outside its XML root")),
        }
        enforce_part_size(output.get_ref().len() as u64, limit)?;
        buffer.clear();
    }
}

fn read_binary_data(reader: &mut NsReader<&[u8]>, limit: u64) -> Result<Vec<u8>> {
    let mut encoded = Vec::new();
    let maximum_encoded = limit
        .checked_add(2)
        .and_then(|value| value.checked_div(3))
        .and_then(|value| value.checked_mul(4))
        .unwrap_or(u64::MAX);
    let mut buffer = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| invalid(format!("invalid Flat OPC binaryData: {error}")))?;
        match event {
            Event::Text(ref text) => {
                let bytes: &[u8] = text.as_ref();
                for byte in bytes {
                    if !byte.is_ascii_whitespace() {
                        encoded.push(*byte);
                        if encoded.len() as u64 > maximum_encoded {
                            return Err(limit_error("part size", limit));
                        }
                    }
                }
            }
            Event::End(ref element) => {
                require_element(namespace, element.local_name().as_ref(), b"binaryData")?;
                let data = BASE64
                    .decode(&encoded)
                    .map_err(|error| invalid(format!("invalid Flat OPC base64: {error}")))?;
                enforce_part_size(data.len() as u64, limit)?;
                return Ok(data);
            }
            Event::Comment(_) | Event::PI(_) => {}
            Event::Eof => return Err(invalid("unclosed Flat OPC binaryData")),
            _ => return Err(invalid("Flat OPC binaryData may contain only base64 text")),
        }
        buffer.clear();
    }
}

fn parts_to_package(parts: Vec<FlatPart>) -> Result<OpcPackage> {
    let mut content_types = ContentTypes::minimal();
    let mut package_rels = None;
    let mut part_rels = HashMap::new();
    let mut package_parts = HashMap::new();

    for part in parts {
        if part.name == "/_rels/.rels" {
            require_relationship_part(&part)?;
            if package_rels
                .replace(parse_relationships(&part.data)?)
                .is_some()
            {
                return Err(invalid("duplicate package relationship part"));
            }
        } else if let Some(owner) = relationship_owner(&part.name) {
            require_relationship_part(&part)?;
            if part_rels
                .insert(owner, parse_relationships(&part.data)?)
                .is_some()
            {
                return Err(invalid("duplicate relationship owner"));
            }
        } else if part.content_type == content_types::RELATIONSHIPS || part.name.contains("/_rels/")
        {
            return Err(invalid(format!(
                "malformed Flat OPC relationship part name {}",
                part.name
            )));
        } else {
            content_types.add_override(&part.name, &part.content_type);
            package_parts.insert(part.name, part.data);
        }
    }
    if part_rels
        .keys()
        .any(|owner| !package_parts.contains_key(owner))
    {
        return Err(invalid(
            "Flat OPC relationship part has no owning package part",
        ));
    }
    let mut package = OpcPackage::new();
    package.content_types = content_types;
    package.package_rels =
        package_rels.ok_or_else(|| invalid("package relationships are missing"))?;
    package.part_rels = part_rels;
    package.parts = package_parts;
    Ok(package)
}

fn write_flat_package(package: &OpcPackage) -> Result<Vec<u8>> {
    let mut parts = Vec::new();
    parts.push(FlatPart {
        name: "/_rels/.rels".to_owned(),
        content_type: content_types::RELATIONSHIPS.to_owned(),
        data: package.package_rels.to_xml()?,
        is_xml: true,
    });
    for (owner, relationships) in &package.part_rels {
        parts.push(FlatPart {
            name: relationship_part_name(owner)?,
            content_type: content_types::RELATIONSHIPS.to_owned(),
            data: relationships.to_xml()?,
            is_xml: true,
        });
    }
    for (name, data) in &package.parts {
        validate_part_name(name)?;
        let content_type = package
            .content_types
            .content_type_for(name)
            .ok_or_else(|| invalid(format!("part {name} has no content type")))?;
        parts.push(FlatPart {
            name: name.clone(),
            content_type: content_type.to_owned(),
            data: data.clone(),
            is_xml: xml_content_type(content_type),
        });
    }
    parts.sort_by(|left, right| left.name.cmp(&right.name));
    for pair in parts.windows(2) {
        if pair[0].name == pair[1].name {
            return Err(invalid(format!(
                "duplicate Flat OPC part name {}",
                pair[0].name
            )));
        }
    }

    let mut writer = Writer::new(Vec::new());
    let mut root = BytesStart::new("pkg:package");
    root.push_attribute(("xmlns:pkg", std::str::from_utf8(PACKAGE_NAMESPACE).unwrap()));
    writer.write_event(Event::Start(root)).map_err(xml_error)?;
    for part in parts {
        let mut element = BytesStart::new("pkg:part");
        element.push_attribute(("pkg:name", part.name.as_str()));
        element.push_attribute(("pkg:contentType", part.content_type.as_str()));
        writer
            .write_event(Event::Start(element))
            .map_err(xml_error)?;
        if part.is_xml {
            validate_strict_xml_1_0(&part.data)
                .map_err(|error| invalid(format!("invalid XML part {}: {error:?}", part.name)))?;
            writer
                .write_event(Event::Start(BytesStart::new("pkg:xmlData")))
                .map_err(xml_error)?;
            copy_xml_without_declaration(&part.data, &mut writer)?;
            writer
                .write_event(Event::End(BytesEnd::new("pkg:xmlData")))
                .map_err(xml_error)?;
        } else {
            writer
                .write_event(Event::Start(BytesStart::new("pkg:binaryData")))
                .map_err(xml_error)?;
            let encoded = BASE64.encode(&part.data);
            writer
                .write_event(Event::Text(BytesText::new(&encoded)))
                .map_err(xml_error)?;
            writer
                .write_event(Event::End(BytesEnd::new("pkg:binaryData")))
                .map_err(xml_error)?;
        }
        writer
            .write_event(Event::End(BytesEnd::new("pkg:part")))
            .map_err(xml_error)?;
    }
    writer
        .write_event(Event::End(BytesEnd::new("pkg:package")))
        .map_err(xml_error)?;
    Ok(writer.into_inner())
}

fn copy_xml_without_declaration(xml: &[u8], output: &mut Writer<Vec<u8>>) -> Result<()> {
    let mut reader = quick_xml::Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    loop {
        match reader
            .read_event_into(&mut buffer)
            .map_err(|error| invalid(format!("invalid XML part: {error}")))?
        {
            Event::Decl(_) => {}
            Event::Eof => return Ok(()),
            event => output.write_event(event.into_owned()).map_err(xml_error)?,
        }
        buffer.clear();
    }
}

fn part_attributes(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<(String, String)> {
    let mut name = None;
    let mut content_type = None;
    for attribute in element.attributes() {
        let attribute =
            attribute.map_err(|error| invalid(format!("invalid part attribute: {error}")))?;
        let key = attribute.key.as_ref();
        if key == b"xmlns" || key.starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if !namespace_is_package(namespace) {
            return Err(invalid(
                "Flat OPC part has an attribute outside the package namespace",
            ));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| invalid(format!("invalid part attribute value: {error}")))?
            .into_owned();
        match local.as_ref() {
            b"name" => {
                if name.replace(value).is_some() {
                    return Err(invalid("duplicate Flat OPC part attribute"));
                }
            }
            b"contentType" => {
                if content_type.replace(value).is_some() {
                    return Err(invalid("duplicate Flat OPC part attribute"));
                }
            }
            _ => return Err(invalid("unknown Flat OPC part attribute")),
        }
    }
    let name = name.ok_or_else(|| invalid("Flat OPC part is missing pkg:name"))?;
    let content_type =
        content_type.ok_or_else(|| invalid("Flat OPC part is missing pkg:contentType"))?;
    validate_content_type(&content_type)?;
    Ok((name, content_type))
}

fn require_no_semantic_attributes(
    _reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<()> {
    for attribute in element.attributes() {
        let attribute =
            attribute.map_err(|error| invalid(format!("invalid structural attribute: {error}")))?;
        let key = attribute.key.as_ref();
        if key != b"xmlns" && !key.starts_with(b"xmlns:") {
            return Err(invalid(
                "Flat OPC structural element has unexpected attributes",
            ));
        }
    }
    Ok(())
}

fn require_element(namespace: ResolveResult<'_>, local: &[u8], wanted: &[u8]) -> Result<()> {
    if is_package_element(namespace, local, wanted) {
        Ok(())
    } else {
        Err(invalid(format!(
            "expected pkg:{} in the Flat OPC namespace",
            String::from_utf8_lossy(wanted)
        )))
    }
}

fn is_package_element(namespace: ResolveResult<'_>, local: &[u8], wanted: &[u8]) -> bool {
    namespace_is_package(namespace) && local == wanted
}

fn namespace_is_package(namespace: ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == PACKAGE_NAMESPACE)
}

fn validate_part_name(name: &str) -> Result<()> {
    if !crate::embedded::relationship_target_is_normalized_pack_uri(name)
        || !name.starts_with('/')
        || name.len() == 1
        || name
            .split('/')
            .skip(1)
            .any(|segment| segment.is_empty() || segment == "." || segment == "..")
        || name.eq_ignore_ascii_case("/[Content_Types].xml")
    {
        return Err(invalid(format!("unsafe Flat OPC part name {name}")));
    }
    Ok(())
}

fn relationship_owner(name: &str) -> Option<String> {
    let (directory, file) = name.rsplit_once("/_rels/")?;
    if directory.contains("/_rels/") || file.contains('/') || !file.ends_with(".rels") {
        return None;
    }
    let owner_file = file.strip_suffix(".rels")?;
    if owner_file.is_empty() {
        return None;
    }
    Some(format!("{directory}/{owner_file}"))
}

fn relationship_part_name(owner: &str) -> Result<String> {
    validate_part_name(owner)?;
    let (directory, file) = owner
        .rsplit_once('/')
        .ok_or_else(|| invalid(format!("invalid relationship owner {owner}")))?;
    Ok(format!("{directory}/_rels/{file}.rels"))
}

fn require_relationship_part(part: &FlatPart) -> Result<()> {
    if part.is_xml && part.content_type == content_types::RELATIONSHIPS {
        Ok(())
    } else {
        Err(invalid(format!(
            "relationship part {} must use relationship XML content",
            part.name
        )))
    }
}

fn parse_relationships(xml: &[u8]) -> Result<Relationships> {
    validate_strict_xml_1_0(xml)
        .map_err(|error| invalid(format!("invalid relationship XML lexical form: {error:?}")))?;
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut relationships = Relationships::new();
    let mut identifiers = HashSet::new();
    let mut saw_root = false;
    let mut closed_root = false;
    let mut open_relationship = false;
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| invalid(format!("invalid relationship XML: {error}")))?;
        match event {
            Event::Decl(_) if !saw_root => {}
            Event::Start(ref element) if !saw_root => {
                if !namespace_is(namespace, RELATIONSHIPS_NAMESPACE)
                    || element.local_name().as_ref() != b"Relationships"
                {
                    return Err(invalid("relationship XML has the wrong root"));
                }
                require_only_namespace_attributes(element)?;
                saw_root = true;
            }
            Event::Empty(ref element) if saw_root && !closed_root && !open_relationship => {
                if !namespace_is(namespace, RELATIONSHIPS_NAMESPACE)
                    || element.local_name().as_ref() != b"Relationship"
                {
                    return Err(invalid("relationship XML has an unexpected child"));
                }
                push_relationship(&reader, element, &mut relationships, &mut identifiers)?;
            }
            Event::Start(ref element) if saw_root && !closed_root && !open_relationship => {
                if !namespace_is(namespace, RELATIONSHIPS_NAMESPACE)
                    || element.local_name().as_ref() != b"Relationship"
                {
                    return Err(invalid("relationship XML has an unexpected child"));
                }
                push_relationship(&reader, element, &mut relationships, &mut identifiers)?;
                open_relationship = true;
            }
            Event::End(ref element) if open_relationship => {
                if !namespace_is(namespace, RELATIONSHIPS_NAMESPACE)
                    || element.local_name().as_ref() != b"Relationship"
                {
                    return Err(invalid("relationship XML has an unexpected close element"));
                }
                open_relationship = false;
            }
            Event::End(ref element) if saw_root && !closed_root && !open_relationship => {
                if !namespace_is(namespace, RELATIONSHIPS_NAMESPACE)
                    || element.local_name().as_ref() != b"Relationships"
                {
                    return Err(invalid("relationship XML has an unexpected close element"));
                }
                closed_root = true;
            }
            Event::Text(ref text) if !open_relationship && text_is_whitespace(text) => {}
            Event::Comment(_) | Event::PI(_) => {}
            Event::Eof => break,
            _ => return Err(invalid("relationship XML has unexpected content")),
        }
        buffer.clear();
    }
    if !saw_root || !closed_root {
        return Err(invalid("relationship XML root is missing or unclosed"));
    }
    Ok(relationships)
}

fn push_relationship(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    relationships: &mut Relationships,
    identifiers: &mut HashSet<String>,
) -> Result<()> {
    let mut id = None;
    let mut rel_type = None;
    let mut target = None;
    let mut target_mode = None;
    for attribute in element.attributes() {
        let attribute = attribute
            .map_err(|error| invalid(format!("invalid relationship attribute: {error}")))?;
        let key = attribute.key.as_ref();
        if key == b"xmlns" || key.starts_with(b"xmlns:") {
            continue;
        }
        let (attribute_namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if !matches!(attribute_namespace, ResolveResult::Unbound) {
            return Err(invalid("relationship attributes must be unqualified"));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| invalid(format!("invalid relationship attribute value: {error}")))?
            .into_owned();
        let destination = match local.as_ref() {
            b"Id" => &mut id,
            b"Type" => &mut rel_type,
            b"Target" => &mut target,
            b"TargetMode" => &mut target_mode,
            _ => return Err(invalid("unknown relationship attribute")),
        };
        if destination.replace(value).is_some() {
            return Err(invalid("duplicate relationship attribute"));
        }
    }
    let id = id.ok_or_else(|| invalid("relationship is missing Id"))?;
    if !identifiers.insert(id.clone()) {
        return Err(invalid(format!("duplicate relationship id {id}")));
    }
    let rel_type = rel_type.ok_or_else(|| invalid("relationship is missing Type"))?;
    let target = target.ok_or_else(|| invalid("relationship is missing Target"))?;
    relationships.add_with_id(&id, &rel_type, &target);
    let relationship = relationships
        .items
        .last_mut()
        .ok_or_else(|| invalid("could not retain parsed relationship"))?;
    relationship.target_mode = target_mode;
    Ok(())
}

fn require_only_namespace_attributes(element: &BytesStart<'_>) -> Result<()> {
    for attribute in element.attributes() {
        let attribute =
            attribute.map_err(|error| invalid(format!("invalid namespace attribute: {error}")))?;
        let key = attribute.key.as_ref();
        if key != b"xmlns" && !key.starts_with(b"xmlns:") {
            return Err(invalid("relationship root has unexpected attributes"));
        }
    }
    Ok(())
}

fn namespace_is(namespace: ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == expected)
}

fn xml_content_type(content_type: &str) -> bool {
    content_type == content_types::XML
        || content_type == "text/xml"
        || content_type == content_types::RELATIONSHIPS
        || content_type.ends_with("+xml")
}

fn validate_content_type(content_type: &str) -> Result<()> {
    let Some((kind, subtype)) = content_type.split_once('/') else {
        return Err(invalid("Flat OPC part content type is not a MIME type"));
    };
    if kind.is_empty()
        || subtype.is_empty()
        || subtype.contains('/')
        || content_type
            .bytes()
            .any(|byte| byte.is_ascii_control() || byte.is_ascii_whitespace() || !byte.is_ascii())
    {
        return Err(invalid("Flat OPC part content type is not canonical"));
    }
    Ok(())
}

fn write_bounded_xml_event(
    output: &mut Writer<Vec<u8>>,
    event: Event<'_>,
    encoded_event_size: u64,
    limit: u64,
) -> Result<()> {
    if (output.get_ref().len() as u64)
        .checked_add(encoded_event_size)
        .is_none_or(|size| size > limit)
    {
        return Err(limit_error("part size", limit));
    }
    output.write_event(event.into_owned()).map_err(xml_error)?;
    enforce_part_size(output.get_ref().len() as u64, limit)
}

fn enforce_part_size(size: u64, limit: u64) -> Result<()> {
    if size > limit {
        Err(limit_error("part size", limit))
    } else {
        Ok(())
    }
}

fn limit_error(kind: &'static str, limit: u64) -> Error {
    invalid(format!("Flat OPC {kind} limit exceeded ({limit})"))
}

fn text_is_whitespace(text: &BytesText<'_>) -> bool {
    let bytes: &[u8] = text.as_ref();
    bytes.iter().all(u8::is_ascii_whitespace)
}

fn xml_error(error: std::io::Error) -> Error {
    invalid(format!("could not write Flat OPC XML: {error}"))
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Other(format!("Flat OPC error: {}", message.into()))
}
