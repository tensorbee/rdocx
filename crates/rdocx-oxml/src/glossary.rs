//! WordprocessingML glossary document parts and building-block entries.

use quick_xml::events::{BytesDecl, BytesEnd, BytesRef, BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use quick_xml::{Reader, Writer, XmlVersion};
use std::collections::HashSet;
use std::ops::Range;

use crate::document::CT_Body;
use crate::error::{OxmlError, Result};
use crate::namespace::W_NS;
use crate::numbering::word_prefixes_at;
use crate::properties::{is_word_attribute, is_word_element};
use crate::raw_xml::{capture_element, capture_empty_element};

/// The `w:glossaryDocument` root.
#[derive(Debug, Clone)]
pub struct CT_GlossaryDocument {
    pub doc_parts: Vec<CT_DocPart>,
    raw_xml: Vec<u8>,
}

/// One `w:docPart` building-block entry.
#[derive(Debug, Clone)]
pub struct CT_DocPart {
    pub properties: CT_DocPartPr,
    pub body: CT_Body,
    raw_xml: Vec<u8>,
    body_xml: Vec<u8>,
    source_span: Range<usize>,
    properties_span: Option<Range<usize>>,
    body_span: Range<usize>,
    original_properties: DocPartPrSnapshot,
    original_body: CT_Body,
}

/// The supported properties of one building block.
#[derive(Debug, Clone)]
pub struct CT_DocPartPr {
    pub name: Option<String>,
    pub style: Option<String>,
    pub category: Option<String>,
    pub gallery: Option<String>,
    pub types: Vec<String>,
    pub behaviors: Vec<String>,
    pub description: Option<String>,
    pub guid: Option<String>,
    source: Option<Box<DocPartPrSource>>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct DocPartPrSnapshot {
    name: Option<String>,
    style: Option<String>,
    category: Option<String>,
    gallery: Option<String>,
    types: Vec<String>,
    behaviors: Vec<String>,
    description: Option<String>,
    guid: Option<String>,
}

#[derive(Debug, Clone)]
struct DocPartPrSource {
    raw_xml: Vec<u8>,
    root_start_xml: Vec<u8>,
    root_end_xml: Vec<u8>,
    field_xml: Vec<Option<Vec<u8>>>,
    extra_xml: Vec<(usize, Vec<u8>)>,
    unsupported: bool,
    word_prefixes: Vec<String>,
    output_prefix: String,
    declare_output_prefix: bool,
}

impl CT_DocPartPr {
    fn empty() -> Self {
        Self {
            name: None,
            style: None,
            category: None,
            gallery: None,
            types: Vec::new(),
            behaviors: Vec::new(),
            description: None,
            guid: None,
            source: None,
        }
    }

    fn snapshot(&self) -> DocPartPrSnapshot {
        DocPartPrSnapshot {
            name: self.name.clone(),
            style: self.style.clone(),
            category: self.category.clone(),
            gallery: self.gallery.clone(),
            types: self.types.clone(),
            behaviors: self.behaviors.clone(),
            description: self.description.clone(),
            guid: self.guid.clone(),
        }
    }

    pub fn has_unsupported_content(&self) -> bool {
        self.source
            .as_deref()
            .is_some_and(|source| source.unsupported)
    }

    fn to_xml(&self, original: &DocPartPrSnapshot) -> Result<Vec<u8>> {
        let Some(source) = self.source.as_deref() else {
            return self.to_canonical_xml(None);
        };
        if &self.snapshot() == original {
            return Ok(source.raw_xml.clone());
        }
        self.to_canonical_xml(Some((source, original)))
    }

    fn to_canonical_xml(
        &self,
        retained: Option<(&DocPartPrSource, &DocPartPrSnapshot)>,
    ) -> Result<Vec<u8>> {
        if self.category.is_some() != self.gallery.is_some() {
            return Err(OxmlError::MissingElement(
                "w:category requires both name and gallery values".to_owned(),
            ));
        }
        let mut writer = Writer::new(Vec::new());
        if let Some((source, _)) = retained {
            writer
                .get_mut()
                .extend_from_slice(&root_start_for_changed_output(source)?);
        } else {
            let mut root = BytesStart::new("w:docPartPr");
            root.push_attribute(("xmlns:w", W_NS));
            writer.write_event(Event::Start(root))?;
        }
        for slot in 0..=7 {
            if let Some((source, _)) = retained {
                for (_, raw) in source.extra_xml.iter().filter(|(at, _)| *at == slot) {
                    writer.get_mut().extend_from_slice(raw);
                }
            }
            if slot == 7 {
                continue;
            }
            let unchanged =
                retained.is_some_and(|(_, original)| property_slot_matches(self, original, slot));
            if unchanged
                && let Some(raw) = retained
                    .and_then(|(source, _)| source.field_xml.get(slot))
                    .and_then(Option::as_deref)
            {
                writer.get_mut().extend_from_slice(raw);
                continue;
            }
            if let Some((source, _)) = retained
                && let Some(raw) = source.field_xml.get(slot).and_then(Option::as_deref)
            {
                writer.get_mut().extend_from_slice(&rewrite_property_slot(
                    raw,
                    self,
                    slot,
                    &source.word_prefixes,
                    &source.output_prefix,
                )?);
            } else {
                let prefix = retained.map_or("w", |(source, _)| source.output_prefix.as_str());
                write_property_slot(&mut writer, self, slot, prefix)?;
            }
        }
        if let Some((source, _)) = retained {
            writer.get_mut().extend_from_slice(&source.root_end_xml);
        } else {
            writer.write_event(Event::End(BytesEnd::new("w:docPartPr")))?;
        }
        Ok(writer.into_inner())
    }
}

fn property_slot_matches(value: &CT_DocPartPr, original: &DocPartPrSnapshot, slot: usize) -> bool {
    match slot {
        0 => value.name == original.name,
        1 => value.style == original.style,
        2 => value.category == original.category && value.gallery == original.gallery,
        3 => value.types == original.types,
        4 => value.behaviors == original.behaviors,
        5 => value.description == original.description,
        6 => value.guid == original.guid,
        _ => true,
    }
}

fn write_property_slot<W: std::io::Write>(
    writer: &mut Writer<W>,
    value: &CT_DocPartPr,
    slot: usize,
    prefix: &str,
) -> Result<()> {
    match slot {
        0 => write_optional_value(writer, "name", value.name.as_deref(), prefix)?,
        1 => write_optional_value(writer, "style", value.style.as_deref(), prefix)?,
        2 if value.category.is_some() || value.gallery.is_some() => {
            writer.write_event(Event::Start(BytesStart::new(format!("{prefix}:category"))))?;
            write_optional_value(writer, "name", value.category.as_deref(), prefix)?;
            write_optional_value(writer, "gallery", value.gallery.as_deref(), prefix)?;
            writer.write_event(Event::End(BytesEnd::new(format!("{prefix}:category"))))?;
        }
        3 if !value.types.is_empty() => {
            writer.write_event(Event::Start(BytesStart::new(format!("{prefix}:types"))))?;
            for value in &value.types {
                write_value(writer, "type", value, prefix)?;
            }
            writer.write_event(Event::End(BytesEnd::new(format!("{prefix}:types"))))?;
        }
        4 if !value.behaviors.is_empty() => {
            writer.write_event(Event::Start(BytesStart::new(format!("{prefix}:behaviors"))))?;
            for value in &value.behaviors {
                write_value(writer, "behavior", value, prefix)?;
            }
            writer.write_event(Event::End(BytesEnd::new(format!("{prefix}:behaviors"))))?;
        }
        5 => write_optional_value(writer, "description", value.description.as_deref(), prefix)?,
        6 => write_optional_value(writer, "guid", value.guid.as_deref(), prefix)?,
        _ => {}
    }
    Ok(())
}

fn write_optional_value<W: std::io::Write>(
    writer: &mut Writer<W>,
    name: &str,
    value: Option<&str>,
    prefix: &str,
) -> Result<()> {
    if let Some(value) = value {
        write_value(writer, name, value, prefix)?;
    }
    Ok(())
}

fn write_value<W: std::io::Write>(
    writer: &mut Writer<W>,
    name: &str,
    value: &str,
    prefix: &str,
) -> Result<()> {
    let qname = format!("{prefix}:{name}");
    let mut element = BytesStart::new(qname);
    element.push_attribute((format!("{prefix}:val").as_str(), value));
    writer.write_event(Event::Empty(element))?;
    Ok(())
}

impl CT_DocPart {
    pub fn has_unsupported_content(&self) -> bool {
        self.properties.has_unsupported_content()
            || self
                .body
                .content
                .iter()
                .any(|item| matches!(item, crate::document::BodyContent::RawXml(_)))
    }

    fn is_unchanged(&self) -> bool {
        self.properties.snapshot() == self.original_properties && self.body == self.original_body
    }

    fn to_xml(&self) -> Result<Vec<u8>> {
        if self.is_unchanged() {
            return Ok(self.raw_xml.clone());
        }
        let properties_changed = self.properties.snapshot() != self.original_properties;
        let body_changed = self.body != self.original_body;
        let mut edits = Vec::new();
        if properties_changed {
            let properties = self.properties.to_xml(&self.original_properties)?;
            if let Some(properties_span) = self.properties_span.clone() {
                edits.push((properties_span, properties));
            } else {
                edits.push((self.body_span.start..self.body_span.start, properties));
            }
        }
        if body_changed {
            let body = serialize_doc_part_body(&self.body, &self.body_xml)?;
            edits.push((self.body_span.clone(), body));
        }
        apply_structural_edits(&self.raw_xml, edits)
    }
}

impl CT_GlossaryDocument {
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        validate_document_declarations_and_doctype(xml)?;
        let mut reader = Reader::from_reader(xml);
        reader.config_mut().trim_text(false);
        let mut buffer = Vec::new();
        let mut word_prefixes = Vec::new();
        let mut prefix_scopes = Vec::new();
        let mut stack = Vec::<Option<String>>::new();
        let mut root_seen = false;
        let mut root_closed = false;
        let mut declaration_allowed = true;
        let mut declaration_seen = false;
        let mut doc_parts_count = 0usize;
        let mut doc_parts = Vec::new();

        loop {
            let event_start = reader.buffer_position() as usize;
            match reader.read_event_into(&mut buffer)? {
                Event::Decl(_) => {
                    if !declaration_allowed || declaration_seen {
                        return Err(OxmlError::InvalidValue(
                            "misplaced or duplicate glossary XML declaration".to_owned(),
                        ));
                    }
                    declaration_seen = true;
                    declaration_allowed = false;
                }
                Event::DocType(_) => {
                    return Err(OxmlError::InvalidValue(
                        "glossary XML cannot contain a document type".to_owned(),
                    ));
                }
                Event::Start(element) => {
                    declaration_allowed = false;
                    if root_closed {
                        return Err(OxmlError::InvalidValue(
                            "multiple top-level glossary roots".to_owned(),
                        ));
                    }
                    let local_prefixes = word_prefixes_at(&element, &word_prefixes)?;
                    let local = word_local_name(&element, &local_prefixes);
                    if !root_seen {
                        if local.as_deref() != Some("glossaryDocument") {
                            return Err(OxmlError::MissingElement(
                                "w:glossaryDocument root".to_owned(),
                            ));
                        }
                        root_seen = true;
                        stack.push(local);
                        prefix_scopes.push(std::mem::replace(&mut word_prefixes, local_prefixes));
                    } else if stack.len() == 1
                        && stack[0].as_deref() == Some("glossaryDocument")
                        && local.as_deref() == Some("docParts")
                    {
                        doc_parts_count += 1;
                        if doc_parts_count > 1 {
                            return Err(OxmlError::InvalidValue(
                                "multiple direct w:docParts containers".to_owned(),
                            ));
                        }
                        stack.push(local);
                        prefix_scopes.push(std::mem::replace(&mut word_prefixes, local_prefixes));
                    } else if stack.len() == 2
                        && stack[0].as_deref() == Some("glossaryDocument")
                        && stack[1].as_deref() == Some("docParts")
                        && local.as_deref() == Some("docPart")
                    {
                        let raw = capture_element(&mut reader, &element)?;
                        let source_span = event_start..reader.buffer_position() as usize;
                        doc_parts.push(parse_doc_part(raw, &local_prefixes, source_span)?);
                    } else {
                        stack.push(local);
                        prefix_scopes.push(std::mem::replace(&mut word_prefixes, local_prefixes));
                    }
                }
                Event::Empty(element) => {
                    declaration_allowed = false;
                    if root_closed {
                        return Err(OxmlError::InvalidValue(
                            "multiple top-level glossary roots".to_owned(),
                        ));
                    }
                    let local_prefixes = word_prefixes_at(&element, &word_prefixes)?;
                    let local = word_local_name(&element, &local_prefixes);
                    if !root_seen {
                        return Err(OxmlError::MissingElement(
                            "nonempty w:glossaryDocument root".to_owned(),
                        ));
                    }
                    if stack.len() == 1
                        && stack[0].as_deref() == Some("glossaryDocument")
                        && local.as_deref() == Some("docParts")
                    {
                        doc_parts_count += 1;
                        if doc_parts_count > 1 {
                            return Err(OxmlError::InvalidValue(
                                "multiple direct w:docParts containers".to_owned(),
                            ));
                        }
                    } else if stack.len() == 2
                        && stack[1].as_deref() == Some("docParts")
                        && local.as_deref() == Some("docPart")
                    {
                        let source_span = event_start..reader.buffer_position() as usize;
                        doc_parts.push(parse_doc_part(
                            capture_empty_element(&element)?,
                            &local_prefixes,
                            source_span,
                        )?);
                    }
                }
                Event::End(_) => {
                    declaration_allowed = false;
                    if stack.len() == 1 {
                        root_closed = true;
                    }
                    stack.pop();
                    word_prefixes = prefix_scopes.pop().unwrap_or_default();
                }
                Event::Text(text)
                    if text.as_ref().iter().any(|byte| !byte.is_ascii_whitespace())
                        && (stack.is_empty() || is_glossary_element_only_container(&stack)) =>
                {
                    return Err(OxmlError::InvalidValue(
                        "non-whitespace text in glossary element-only content".to_owned(),
                    ));
                }
                Event::CData(text)
                    if stack.is_empty()
                        || (is_glossary_element_only_container(&stack)
                            && text.as_ref().iter().any(|byte| !byte.is_ascii_whitespace())) =>
                {
                    return Err(OxmlError::InvalidValue(
                        "character data in glossary element-only content".to_owned(),
                    ));
                }
                Event::GeneralRef(reference) => {
                    let non_whitespace = match reference.resolve_char_ref() {
                        Ok(Some(character)) => !character.is_ascii_whitespace(),
                        Ok(None) | Err(_) => true,
                    };
                    if stack.is_empty()
                        || (is_glossary_element_only_container(&stack) && non_whitespace)
                    {
                        return Err(OxmlError::InvalidValue(
                            "character reference in glossary element-only content".to_owned(),
                        ));
                    }
                }
                Event::Eof => break,
                _ => declaration_allowed = false,
            }
            buffer.clear();
        }
        if !root_seen || !root_closed || doc_parts_count != 1 || doc_parts.is_empty() {
            return Err(OxmlError::MissingElement(
                "w:glossaryDocument root with one direct w:docParts".to_owned(),
            ));
        }
        Ok(Self {
            doc_parts,
            raw_xml: xml.to_vec(),
        })
    }

    pub fn to_xml(&self) -> Result<Vec<u8>> {
        let edits = self
            .doc_parts
            .iter()
            .filter(|part| !part.is_unchanged())
            .map(|part| Ok((part.source_span.clone(), part.to_xml()?)))
            .collect::<Result<Vec<_>>>()?;
        apply_structural_edits(&self.raw_xml, edits)
    }
}

fn is_glossary_element_only_container(stack: &[Option<String>]) -> bool {
    matches!(stack, [Some(root)] if root == "glossaryDocument")
        || matches!(
            stack,
            [Some(root), Some(container)]
                if root == "glossaryDocument" && container == "docParts"
        )
}

fn validate_document_declarations_and_doctype(xml: &[u8]) -> Result<()> {
    validate_literal_xml_characters(xml, "glossary")?;
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_comments = true;
    let mut buffer = Vec::new();
    let mut declaration_allowed = true;
    let mut declaration_seen = false;
    let mut stack = Vec::<Option<Vec<u8>>>::new();
    loop {
        let (namespace, event) = reader.read_resolved_event_into(&mut buffer)?;
        let is_word = namespace_is(&namespace, W_NS);
        let event = event.into_owned();
        drop(namespace);
        validate_scanned_xml_event(&reader, &event, "glossary")?;
        match event {
            Event::Decl(declaration) => {
                if !declaration_allowed || declaration_seen {
                    return Err(OxmlError::InvalidValue(
                        "misplaced or duplicate glossary XML declaration".to_owned(),
                    ));
                }
                validate_xml_declaration(&declaration, "glossary")?;
                declaration_seen = true;
                declaration_allowed = false;
            }
            Event::DocType(_) => {
                return Err(OxmlError::InvalidValue(
                    "glossary XML cannot contain a document type".to_owned(),
                ));
            }
            Event::Start(element) => {
                declaration_allowed = false;
                let element_name = element.name();
                let local = local_name(element_name.as_ref());
                stack.push(
                    (is_word && glossary_container_is_element_only(local)).then(|| local.to_vec()),
                );
            }
            Event::Empty(_) => {
                declaration_allowed = false;
            }
            Event::End(_) => {
                declaration_allowed = false;
                stack.pop();
            }
            Event::Text(text)
                if stack.last().is_some_and(Option::is_some)
                    && text.iter().any(|byte| !byte.is_ascii_whitespace()) =>
            {
                return Err(OxmlError::InvalidValue(
                    "non-whitespace text in glossary element-only content".to_owned(),
                ));
            }
            Event::CData(text)
                if stack.last().is_some_and(Option::is_some)
                    && text.iter().any(|byte| !byte.is_ascii_whitespace()) =>
            {
                return Err(OxmlError::InvalidValue(
                    "character data in glossary element-only content".to_owned(),
                ));
            }
            Event::GeneralRef(reference) => {
                let character = require_predefined_or_character_reference(&reference, "glossary")?;
                if stack.last().is_some_and(Option::is_some) && !character.is_ascii_whitespace() {
                    return Err(OxmlError::InvalidValue(
                        "character reference in glossary element-only content".to_owned(),
                    ));
                }
            }
            Event::PI(instruction) if instruction.target().eq_ignore_ascii_case(b"xml") => {
                return Err(OxmlError::InvalidValue(
                    "reserved glossary XML processing instruction".to_owned(),
                ));
            }
            Event::Eof => return Ok(()),
            _ => declaration_allowed = false,
        }
        buffer.clear();
    }
}

fn glossary_container_is_element_only(local: &[u8]) -> bool {
    matches!(
        local,
        b"glossaryDocument"
            | b"docParts"
            | b"docPart"
            | b"docPartPr"
            | b"category"
            | b"types"
            | b"behaviors"
    )
}

fn validate_literal_xml_characters(xml: &[u8], owner: &str) -> Result<()> {
    let xml = std::str::from_utf8(xml)?;
    if xml.chars().all(xml_1_0_character_is_valid) {
        return Ok(());
    }
    Err(OxmlError::InvalidValue(format!(
        "{owner} XML contains a forbidden literal XML 1.0 character"
    )))
}

fn xml_1_0_character_is_valid(character: char) -> bool {
    matches!(character, '\t' | '\n' | '\r')
        || ('\u{20}'..='\u{D7FF}').contains(&character)
        || ('\u{E000}'..='\u{FFFD}').contains(&character)
        || ('\u{10000}'..='\u{10FFFF}').contains(&character)
}

fn validate_scanned_xml_event(
    reader: &NsReader<&[u8]>,
    event: &Event<'_>,
    owner: &str,
) -> Result<()> {
    match event {
        Event::Start(element) | Event::Empty(element) => {
            validate_scanned_element(reader, element, owner)
        }
        Event::End(element) => {
            let name = element.name();
            let prefix = validate_xml_qname(name.as_ref(), owner)?;
            validate_bound_prefix(&reader.resolver().resolve_element(name).0, prefix, owner)
                .map(|_| ())
        }
        Event::PI(instruction) => validate_xml_name(instruction.target(), owner),
        _ => Ok(()),
    }
}

fn validate_scanned_element(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    owner: &str,
) -> Result<()> {
    let element_name = element.name();
    let prefix = validate_xml_qname(element_name.as_ref(), owner)?;
    if prefix == Some(b"xmlns".as_slice()) {
        return Err(OxmlError::InvalidValue(format!(
            "{owner} XML element uses the reserved xmlns prefix"
        )));
    }
    validate_bound_prefix(
        &reader.resolver().resolve_element(element_name).0,
        prefix,
        owner,
    )?;
    let mut expanded_names = HashSet::new();
    for attribute in element.attributes() {
        let attribute = attribute?;
        let name = attribute.key.as_ref();
        let prefix = validate_xml_qname(name, owner)?;
        if attribute.value.contains(&b'<') {
            return Err(OxmlError::InvalidValue(format!(
                "{owner} XML attribute contains a literal less-than sign"
            )));
        }
        let value =
            attribute.decoded_and_normalized_value(XmlVersion::Implicit1_0, element.decoder())?;
        if !value.chars().all(xml_1_0_character_is_valid) {
            return Err(OxmlError::InvalidValue(format!(
                "{owner} XML attribute contains a forbidden XML 1.0 character"
            )));
        }
        if name == b"xmlns" {
            validate_namespace_declaration(None, value.as_bytes(), owner)?;
            continue;
        }
        if prefix == Some(b"xmlns".as_slice()) {
            validate_namespace_declaration(Some(local_name(name)), value.as_bytes(), owner)?;
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let resolved = validate_bound_prefix(&namespace, prefix, owner)?;
        if !expanded_names.insert((resolved, local.as_ref().to_vec())) {
            return Err(OxmlError::InvalidValue(format!(
                "{owner} XML element has duplicate expanded-name attributes"
            )));
        }
    }
    Ok(())
}

fn validate_namespace_declaration(
    prefix: Option<&[u8]>,
    namespace: &[u8],
    owner: &str,
) -> Result<()> {
    const XML_NS: &[u8] = b"http://www.w3.org/XML/1998/namespace";
    const XMLNS_NS: &[u8] = b"http://www.w3.org/2000/xmlns/";
    let valid = match prefix {
        None => namespace != XML_NS && namespace != XMLNS_NS,
        Some(b"xml") => namespace == XML_NS,
        Some(b"xmlns") => false,
        Some(_) => !namespace.is_empty() && namespace != XML_NS && namespace != XMLNS_NS,
    };
    if valid {
        Ok(())
    } else {
        Err(OxmlError::InvalidValue(format!(
            "{owner} XML contains an invalid namespace declaration"
        )))
    }
}

fn validate_bound_prefix(
    namespace: &ResolveResult<'_>,
    prefix: Option<&[u8]>,
    owner: &str,
) -> Result<Option<Vec<u8>>> {
    match namespace {
        ResolveResult::Bound(Namespace(namespace)) => Ok(Some(namespace.to_vec())),
        ResolveResult::Unbound if prefix.is_none() => Ok(None),
        ResolveResult::Unbound => Err(OxmlError::InvalidValue(format!(
            "{owner} XML uses an unbound namespace prefix"
        ))),
        ResolveResult::Unknown(_) => Err(OxmlError::InvalidValue(format!(
            "{owner} XML uses an unbound namespace prefix"
        ))),
    }
}

fn validate_xml_qname<'a>(name: &'a [u8], owner: &str) -> Result<Option<&'a [u8]>> {
    let name = std::str::from_utf8(name)?;
    let mut parts = name.split(':');
    let first = parts.next().unwrap_or_default();
    let second = parts.next();
    if !xml_ncname_is_valid(first)
        || second.is_some_and(|local| !xml_ncname_is_valid(local))
        || parts.next().is_some()
    {
        return Err(OxmlError::InvalidValue(format!(
            "invalid {owner} XML qualified name {name}"
        )));
    }
    Ok(second.map(|_| first.as_bytes()))
}

fn validate_xml_name(name: &[u8], owner: &str) -> Result<()> {
    let name = std::str::from_utf8(name)?;
    let mut characters = name.chars();
    if characters
        .next()
        .is_some_and(|character| character == ':' || xml_ncname_start_character(character))
        && characters.all(|character| character == ':' || xml_ncname_character(character))
    {
        Ok(())
    } else {
        Err(OxmlError::InvalidValue(format!(
            "invalid {owner} XML name {name}"
        )))
    }
}

fn xml_ncname_is_valid(value: &str) -> bool {
    let mut characters = value.chars();
    characters.next().is_some_and(xml_ncname_start_character)
        && characters.all(xml_ncname_character)
}

fn xml_ncname_start_character(character: char) -> bool {
    matches!(
        character,
        'A'..='Z' | '_' | 'a'..='z' | '\u{00C0}'..='\u{00D6}' | '\u{00D8}'..='\u{00F6}'
            | '\u{00F8}'..='\u{02FF}' | '\u{0370}'..='\u{037D}' | '\u{037F}'..='\u{1FFF}'
            | '\u{200C}'..='\u{200D}' | '\u{2070}'..='\u{218F}' | '\u{2C00}'..='\u{2FEF}'
            | '\u{3001}'..='\u{D7FF}' | '\u{F900}'..='\u{FDCF}' | '\u{FDF0}'..='\u{FFFD}'
            | '\u{10000}'..='\u{EFFFF}'
    )
}

fn xml_ncname_character(character: char) -> bool {
    xml_ncname_start_character(character)
        || matches!(character, '-' | '.' | '0'..='9' | '\u{00B7}' | '\u{0300}'..='\u{036F}' | '\u{203F}'..='\u{2040}')
}

fn require_predefined_or_character_reference(
    reference: &BytesRef<'_>,
    owner: &str,
) -> Result<char> {
    if let Some(character) = reference.resolve_char_ref()? {
        if xml_1_0_character_is_valid(character) {
            return Ok(character);
        }
        return Err(OxmlError::InvalidValue(format!(
            "{owner} character reference is not legal in XML 1.0"
        )));
    }
    let name = reference
        .decode()
        .map_err(|error| OxmlError::InvalidValue(format!("invalid {owner} XML: {error}")))?;
    match name.as_ref() {
        "amp" => Ok('&'),
        "lt" => Ok('<'),
        "gt" => Ok('>'),
        "apos" => Ok('\''),
        "quot" => Ok('"'),
        _ => Err(OxmlError::InvalidValue(format!(
            "undeclared {owner} XML entity reference &{name};"
        ))),
    }
}

fn namespace_is(namespace: &ResolveResult<'_>, expected: &str) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(uri)) if *uri == expected.as_bytes())
}

fn local_name(name: &[u8]) -> &[u8] {
    name.rsplit(|byte| *byte == b':').next().unwrap_or(name)
}

fn validate_xml_declaration(declaration: &BytesDecl<'_>, owner: &str) -> Result<()> {
    let content = std::str::from_utf8(declaration.as_ref())?;
    let start = BytesStart::from_content(content, 3);
    let attributes = start
        .attributes()
        .collect::<std::result::Result<Vec<_>, _>>()?;
    if attributes.first().map(|attribute| attribute.key.as_ref()) != Some(b"version".as_ref()) {
        return Err(OxmlError::InvalidValue(format!(
            "{owner} XML declaration must begin with version"
        )));
    }
    let mut encoding_seen = false;
    let mut standalone_seen = false;
    for (index, attribute) in attributes.iter().enumerate() {
        let value = attribute
            .normalized_value(XmlVersion::Explicit1_0)?
            .into_owned();
        match attribute.key.as_ref() {
            b"version" if index == 0 && value == "1.0" => {}
            b"encoding" if !encoding_seen && !standalone_seen && valid_encoding_name(&value) => {
                encoding_seen = true;
            }
            b"standalone" if !standalone_seen && matches!(value.as_str(), "yes" | "no") => {
                standalone_seen = true;
            }
            _ => {
                return Err(OxmlError::InvalidValue(format!(
                    "invalid {owner} XML declaration"
                )));
            }
        }
    }
    Ok(())
}

fn valid_encoding_name(value: &str) -> bool {
    let mut bytes = value.bytes();
    bytes.next().is_some_and(|byte| byte.is_ascii_alphabetic())
        && bytes.all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'.' | b'_' | b'-'))
}

fn parse_doc_part(
    raw_xml: Vec<u8>,
    inherited_prefixes: &[String],
    source_span: Range<usize>,
) -> Result<CT_DocPart> {
    let mut reader = Reader::from_reader(raw_xml.as_slice());
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut prefixes = inherited_prefixes.to_vec();
    let mut prefix_scopes = Vec::new();
    let mut owner_prefixes = inherited_prefixes.to_vec();
    let mut depth = 0usize;
    let mut properties_xml = None;
    let mut properties_span = None;
    let mut body_xml = None;
    let mut body_span = None;
    loop {
        let event_start = reader.buffer_position() as usize;
        match reader.read_event_into(&mut buffer)? {
            Event::Start(element) => {
                let local_prefixes = word_prefixes_at(&element, &prefixes)?;
                let local = word_local_name(&element, &local_prefixes);
                if depth == 0 {
                    owner_prefixes = local_prefixes.clone();
                }
                if depth == 1 && local.as_deref() == Some("docPartPr") {
                    if properties_xml.is_some() {
                        return Err(OxmlError::InvalidValue(
                            "multiple direct w:docPartPr children".to_owned(),
                        ));
                    }
                    if body_xml.is_some() {
                        return Err(OxmlError::InvalidValue(
                            "w:docPartPr must precede w:docPartBody".to_owned(),
                        ));
                    }
                    properties_xml = Some(capture_element(&mut reader, &element)?);
                    properties_span = Some(event_start..reader.buffer_position() as usize);
                } else if depth == 1 && local.as_deref() == Some("docPartBody") {
                    if body_xml.is_some() {
                        return Err(OxmlError::InvalidValue(
                            "multiple direct w:docPartBody children".to_owned(),
                        ));
                    }
                    body_xml = Some(capture_element(&mut reader, &element)?);
                    body_span = Some(event_start..reader.buffer_position() as usize);
                } else {
                    depth += 1;
                    prefix_scopes.push(std::mem::replace(&mut prefixes, local_prefixes));
                }
            }
            Event::Empty(element) => {
                let local_prefixes = word_prefixes_at(&element, &prefixes)?;
                let local = word_local_name(&element, &local_prefixes);
                if depth == 1 && local.as_deref() == Some("docPartPr") {
                    if properties_xml.is_some() {
                        return Err(OxmlError::InvalidValue(
                            "multiple direct w:docPartPr children".to_owned(),
                        ));
                    }
                    if body_xml.is_some() {
                        return Err(OxmlError::InvalidValue(
                            "w:docPartPr must precede w:docPartBody".to_owned(),
                        ));
                    }
                    properties_xml = Some(capture_empty_element(&element)?);
                    properties_span = Some(event_start..reader.buffer_position() as usize);
                } else if depth == 1 && local.as_deref() == Some("docPartBody") {
                    if body_xml.is_some() {
                        return Err(OxmlError::InvalidValue(
                            "multiple direct w:docPartBody children".to_owned(),
                        ));
                    }
                    body_xml = Some(capture_empty_element(&element)?);
                    body_span = Some(event_start..reader.buffer_position() as usize);
                }
            }
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                prefixes = prefix_scopes
                    .pop()
                    .unwrap_or_else(|| inherited_prefixes.to_vec());
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    let body_xml = body_xml.ok_or_else(|| OxmlError::MissingElement("w:docPartBody".to_owned()))?;
    let body_span = body_span.expect("body span accompanies body XML");
    let properties = properties_xml
        .as_deref()
        .map(|xml| parse_properties(xml, &owner_prefixes))
        .transpose()?
        .unwrap_or_else(CT_DocPartPr::empty);
    let body = parse_doc_part_body(&body_xml, &owner_prefixes)?;
    Ok(CT_DocPart {
        original_properties: properties.snapshot(),
        original_body: body.clone(),
        properties,
        body,
        raw_xml,
        body_xml,
        source_span,
        properties_span,
        body_span,
    })
}

fn parse_properties(raw_xml: &[u8], inherited_prefixes: &[String]) -> Result<CT_DocPartPr> {
    let mut reader = Reader::from_reader(raw_xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut prefixes = inherited_prefixes.to_vec();
    let mut root_seen = false;
    let mut slot = 0usize;
    let mut field_xml = vec![None; 7];
    let mut seen_fields = [false; 7];
    let mut extra_xml = Vec::new();
    let mut unsupported = false;
    let mut root_start_xml = Vec::new();
    let mut root_end_xml = Vec::new();
    let mut source_word_prefixes = inherited_prefixes.to_vec();
    let mut value = CT_DocPartPr {
        name: None,
        style: None,
        category: None,
        gallery: None,
        types: Vec::new(),
        behaviors: Vec::new(),
        description: None,
        guid: None,
        source: None,
    };
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(element) => {
                let local_prefixes = word_prefixes_at(&element, &prefixes)?;
                if !root_seen {
                    root_seen = true;
                    root_start_xml = standalone_event(Event::Start(element.into_owned()))?;
                    source_word_prefixes = local_prefixes.clone();
                    prefixes = local_prefixes;
                } else {
                    let raw = capture_element(&mut reader, &element)?;
                    if let Some(field_slot) = property_slot(&element, &local_prefixes) {
                        if std::mem::replace(&mut seen_fields[field_slot], true) {
                            return Err(OxmlError::InvalidValue(format!(
                                "duplicate modeled glossary property in slot {field_slot}"
                            )));
                        }
                        if field_slot < slot {
                            return Err(OxmlError::InvalidValue(format!(
                                "out-of-order modeled glossary property in slot {field_slot}"
                            )));
                        }
                        apply_property_field(&mut value, field_slot, &raw, &local_prefixes)?;
                        field_xml[field_slot] = Some(raw);
                        slot = field_slot + 1;
                    } else {
                        unsupported = true;
                        extra_xml.push((slot, raw));
                    }
                }
            }
            Event::Empty(element) => {
                let local_prefixes = word_prefixes_at(&element, &prefixes)?;
                if !root_seen {
                    root_seen = true;
                    source_word_prefixes = local_prefixes;
                    let name = String::from_utf8_lossy(element.name().as_ref()).into_owned();
                    root_start_xml = capture_empty_element(&element)?;
                    root_start_xml.splice(root_start_xml.len() - 2.., *b">");
                    root_end_xml = standalone_event(Event::End(BytesEnd::new(name)))?;
                } else {
                    let raw = capture_empty_element(&element)?;
                    if let Some(field_slot) = property_slot(&element, &local_prefixes) {
                        if std::mem::replace(&mut seen_fields[field_slot], true) {
                            return Err(OxmlError::InvalidValue(format!(
                                "duplicate modeled glossary property in slot {field_slot}"
                            )));
                        }
                        if field_slot < slot {
                            return Err(OxmlError::InvalidValue(format!(
                                "out-of-order modeled glossary property in slot {field_slot}"
                            )));
                        }
                        apply_property_field(&mut value, field_slot, &raw, &local_prefixes)?;
                        field_xml[field_slot] = Some(raw);
                        slot = field_slot + 1;
                    } else {
                        unsupported = true;
                        extra_xml.push((slot, raw));
                    }
                }
            }
            Event::End(element) if root_seen => {
                root_end_xml = standalone_event(Event::End(element.into_owned()))?;
                break;
            }
            Event::Eof => break,
            event if root_seen => {
                let raw = standalone_event(event.into_owned())?;
                if !raw.is_empty() {
                    extra_xml.push((slot, raw));
                }
            }
            _ => {}
        }
        buffer.clear();
    }
    let output_prefix = changed_output_prefix(&source_word_prefixes, &root_start_xml);
    value.source = Some(Box::new(DocPartPrSource {
        raw_xml: raw_xml.to_vec(),
        root_start_xml,
        root_end_xml,
        field_xml,
        extra_xml,
        unsupported,
        output_prefix,
        declare_output_prefix: !source_word_prefixes
            .iter()
            .any(|prefix| !prefix.is_empty() && !prefix.starts_with('\0')),
        word_prefixes: source_word_prefixes,
    }));
    Ok(value)
}

fn changed_output_prefix(prefixes: &[String], root_start: &[u8]) -> String {
    if let Some(prefix) = prefixes
        .iter()
        .find(|prefix| !prefix.is_empty() && !prefix.starts_with('\0'))
    {
        return prefix.clone();
    }
    for index in 0usize.. {
        let candidate = if index == 0 {
            "w".to_owned()
        } else {
            format!("w{index}")
        };
        let attribute = format!("xmlns:{candidate}");
        if !start_tag_has_attribute(root_start, attribute.as_bytes()) {
            return candidate;
        }
    }
    unreachable!()
}

fn root_start_for_changed_output(source: &DocPartPrSource) -> Result<Vec<u8>> {
    let mut start = source.root_start_xml.clone();
    if !source.declare_output_prefix {
        return Ok(start);
    }
    let insert_at = start
        .iter()
        .rposition(|byte| *byte == b'>')
        .ok_or_else(|| OxmlError::MissingElement("w:docPartPr start tag".to_owned()))?;
    start.splice(
        insert_at..insert_at,
        format!(r#" xmlns:{}="{W_NS}""#, source.output_prefix).bytes(),
    );
    Ok(start)
}

fn rewrite_property_slot(
    raw: &[u8],
    value: &CT_DocPartPr,
    slot: usize,
    word_prefixes: &[String],
    output_prefix: &str,
) -> Result<Vec<u8>> {
    match slot {
        0 => patch_root_word_value(raw, word_prefixes, value.name.as_deref(), output_prefix),
        1 => patch_root_word_value(raw, word_prefixes, value.style.as_deref(), output_prefix),
        2 if value.category.is_none() && value.gallery.is_none() => Ok(Vec::new()),
        2 => patch_container_values(
            raw,
            word_prefixes,
            &[
                ("name", value.category.as_deref()),
                ("gallery", value.gallery.as_deref()),
            ],
            output_prefix,
        ),
        3 if value.types.is_empty() => Ok(Vec::new()),
        3 => {
            patch_repeated_container_values(raw, word_prefixes, "type", &value.types, output_prefix)
        }
        4 if value.behaviors.is_empty() => Ok(Vec::new()),
        4 => patch_repeated_container_values(
            raw,
            word_prefixes,
            "behavior",
            &value.behaviors,
            output_prefix,
        ),
        5 => patch_root_word_value(
            raw,
            word_prefixes,
            value.description.as_deref(),
            output_prefix,
        ),
        6 => patch_root_word_value(raw, word_prefixes, value.guid.as_deref(), output_prefix),
        _ => Ok(raw.to_vec()),
    }
}

fn patch_root_word_value(
    raw: &[u8],
    inherited_prefixes: &[String],
    value: Option<&str>,
    output_prefix: &str,
) -> Result<Vec<u8>> {
    if value.is_none() {
        return Ok(Vec::new());
    }
    let mut reader = Reader::from_reader(raw);
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            event @ (Event::Start(_) | Event::Empty(_)) => {
                let empty = matches!(event, Event::Empty(_));
                let element = match event {
                    Event::Start(element) | Event::Empty(element) => element,
                    _ => unreachable!(),
                };
                let prefixes = word_prefixes_at(&element, inherited_prefixes)?;
                let mut tag = if empty {
                    let mut writer = Writer::new(Vec::new());
                    writer.write_event(Event::Empty(element.clone()))?;
                    writer.into_inner()
                } else {
                    let mut writer = Writer::new(Vec::new());
                    writer.write_event(Event::Start(element.clone()))?;
                    writer.into_inner()
                };
                let mut attribute_key = None;
                for attribute in element.attributes() {
                    let attribute = attribute?;
                    if is_word_attribute(attribute.key.as_ref(), b"val", &prefixes) {
                        attribute_key = Some(attribute.key.as_ref().to_vec());
                        break;
                    }
                }
                match (attribute_key, value) {
                    (Some(key), Some(value)) => {
                        if let Some(range) = lexical_attribute_value_range(&tag, &key) {
                            let escaped = quick_xml::escape::escape(value);
                            tag.splice(range, escaped.as_bytes().iter().copied());
                        }
                    }
                    (Some(_), None) => unreachable!("handled before parsing the element"),
                    (None, Some(value)) => {
                        let (prefix, declare) = usable_word_prefix(&prefixes, &tag, output_prefix);
                        if declare {
                            declare_word_prefix(&mut tag, &prefix)?;
                        }
                        let insert_at = if tag.ends_with(b"/>") {
                            tag.len() - 2
                        } else {
                            tag.len() - 1
                        };
                        let escaped = quick_xml::escape::escape(value);
                        tag.splice(
                            insert_at..insert_at,
                            format!(" {prefix}:val=\"{escaped}\"").bytes(),
                        );
                    }
                    (None, None) => unreachable!("handled before parsing the element"),
                }
                let original_tag_len = reader.buffer_position() as usize;
                let mut output = tag;
                output.extend_from_slice(&raw[original_tag_len..]);
                return Ok(output);
            }
            Event::Eof => return Ok(raw.to_vec()),
            _ => {}
        }
        buffer.clear();
    }
}

fn patch_container_values(
    raw: &[u8],
    inherited_prefixes: &[String],
    desired: &[(&str, Option<&str>)],
    output_prefix: &str,
) -> Result<Vec<u8>> {
    let (raw, output_prefix) = prepare_container_output(raw, inherited_prefixes, output_prefix)?;
    let seen = std::cell::RefCell::new(vec![false; desired.len()]);
    rewrite_container(
        &raw,
        inherited_prefixes,
        |name, child, prefixes, output| {
            let Some(index) = desired.iter().position(|(candidate, _)| *candidate == name) else {
                output.extend_from_slice(child);
                return Ok(());
            };
            let mut seen = seen.borrow_mut();
            if seen[index] {
                return Err(OxmlError::InvalidValue(format!(
                    "duplicate direct w:{} child",
                    desired[index].0
                )));
            }
            for prior in 0..index {
                if !seen[prior] {
                    if let Some(value) = desired[prior].1 {
                        write_value_bytes(output, desired[prior].0, value, &output_prefix);
                    }
                    seen[prior] = true;
                }
            }
            output.extend_from_slice(&patch_root_word_value(
                child,
                prefixes,
                desired[index].1,
                &output_prefix,
            )?);
            seen[index] = true;
            Ok(())
        },
        |output| {
            let seen = seen.borrow();
            for (index, (name, value)) in desired.iter().enumerate() {
                if !seen[index]
                    && let Some(value) = value
                {
                    write_value_bytes(output, name, value, &output_prefix);
                }
            }
        },
    )
}

fn patch_repeated_container_values(
    raw: &[u8],
    inherited_prefixes: &[String],
    child_name: &str,
    desired: &[String],
    output_prefix: &str,
) -> Result<Vec<u8>> {
    let (raw, output_prefix) = prepare_container_output(raw, inherited_prefixes, output_prefix)?;
    let next = std::cell::Cell::new(0usize);
    rewrite_container(
        &raw,
        inherited_prefixes,
        |name, child, prefixes, output| {
            if name != child_name {
                output.extend_from_slice(child);
                return Ok(());
            }
            let index = next.get();
            let value = desired.get(index).map(String::as_str);
            output.extend_from_slice(&patch_root_word_value(
                child,
                prefixes,
                value,
                &output_prefix,
            )?);
            next.set(index + usize::from(value.is_some()));
            Ok(())
        },
        |output| {
            for value in &desired[next.get()..] {
                write_value_bytes(output, child_name, value, &output_prefix);
            }
        },
    )
}

fn rewrite_container<Child, Finish>(
    raw: &[u8],
    inherited_prefixes: &[String],
    mut child: Child,
    mut finish: Finish,
) -> Result<Vec<u8>>
where
    Child: FnMut(&str, &[u8], &[String], &mut Vec<u8>) -> Result<()>,
    Finish: FnMut(&mut Vec<u8>),
{
    let mut reader = Reader::from_reader(raw);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut output = Vec::with_capacity(raw.len());
    let mut prefixes = inherited_prefixes.to_vec();
    let mut prefix_scopes = Vec::new();
    let mut depth = 0usize;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(element) => {
                let local_prefixes = word_prefixes_at(&element, &prefixes)?;
                let local = word_local_name(&element, &local_prefixes);
                if depth == 1
                    && let Some(local) = local
                {
                    let raw_child = capture_element(&mut reader, &element)?;
                    child(&local, &raw_child, &local_prefixes, &mut output)?;
                } else {
                    let mut writer = Writer::new(Vec::new());
                    writer.write_event(Event::Start(element))?;
                    output.extend_from_slice(&writer.into_inner());
                    depth += 1;
                    prefix_scopes.push(std::mem::replace(&mut prefixes, local_prefixes));
                }
            }
            Event::Empty(element) => {
                let local_prefixes = word_prefixes_at(&element, &prefixes)?;
                let local = word_local_name(&element, &local_prefixes);
                let raw_child = capture_empty_element(&element)?;
                if depth == 1
                    && let Some(local) = local
                {
                    child(&local, &raw_child, &local_prefixes, &mut output)?;
                } else {
                    output.extend_from_slice(&raw_child);
                }
            }
            Event::End(element) => {
                if depth == 1 {
                    finish(&mut output);
                }
                let mut writer = Writer::new(Vec::new());
                writer.write_event(Event::End(element))?;
                output.extend_from_slice(&writer.into_inner());
                depth = depth.saturating_sub(1);
                prefixes = prefix_scopes
                    .pop()
                    .unwrap_or_else(|| inherited_prefixes.to_vec());
            }
            Event::Eof => break,
            event => {
                let raw_event = standalone_event(event.into_owned())?;
                output.extend_from_slice(&raw_event);
            }
        }
        buffer.clear();
    }
    Ok(output)
}

fn prepare_container_output(
    raw: &[u8],
    inherited_prefixes: &[String],
    preferred: &str,
) -> Result<(Vec<u8>, String)> {
    let mut reader = Reader::from_reader(raw);
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            event @ (Event::Start(_) | Event::Empty(_)) => {
                let empty = matches!(event, Event::Empty(_));
                let element = match event {
                    Event::Start(element) | Event::Empty(element) => element,
                    _ => unreachable!(),
                };
                let tag_end = reader.buffer_position() as usize;
                let mut start = raw[..tag_end].to_vec();
                let prefixes = word_prefixes_at(&element, inherited_prefixes)?;
                let (prefix, declare) = usable_word_prefix(&prefixes, &start, preferred);
                if declare {
                    declare_word_prefix(&mut start, &prefix)?;
                }
                let mut output = start;
                if empty {
                    let slash = output
                        .iter()
                        .rposition(|byte| *byte == b'/')
                        .ok_or_else(|| {
                            OxmlError::MissingElement("empty container close".to_owned())
                        })?;
                    output.remove(slash);
                    let name = std::str::from_utf8(element.name().as_ref())?.to_owned();
                    output.extend_from_slice(format!("</{name}>").as_bytes());
                } else {
                    output.extend_from_slice(&raw[tag_end..]);
                }
                return Ok((output, prefix));
            }
            Event::Eof => {
                return Err(OxmlError::MissingElement(
                    "glossary property container".to_owned(),
                ));
            }
            _ => {}
        }
        buffer.clear();
    }
}

fn usable_word_prefix(prefixes: &[String], tag: &[u8], preferred: &str) -> (String, bool) {
    if !preferred.is_empty() && prefixes.iter().any(|prefix| prefix == preferred) {
        return (preferred.to_owned(), false);
    }
    if let Some(prefix) = prefixes
        .iter()
        .find(|prefix| !prefix.is_empty() && !prefix.starts_with('\0'))
    {
        return (prefix.clone(), false);
    }
    (changed_output_prefix(&[], tag), true)
}

fn declare_word_prefix(tag: &mut Vec<u8>, prefix: &str) -> Result<()> {
    let insert_at = if tag.ends_with(b"/>") {
        tag.len() - 2
    } else {
        tag.len()
            .checked_sub(1)
            .ok_or_else(|| OxmlError::MissingElement("element start tag".to_owned()))?
    };
    tag.splice(
        insert_at..insert_at,
        format!(r#" xmlns:{prefix}="{W_NS}""#).bytes(),
    );
    Ok(())
}

fn write_value_bytes(output: &mut Vec<u8>, name: &str, value: &str, prefix: &str) {
    let value = quick_xml::escape::escape(value);
    output.extend_from_slice(format!("<{prefix}:{name} {prefix}:val=\"{value}\"/>").as_bytes());
}

fn start_tag_has_attribute(tag: &[u8], wanted: &[u8]) -> bool {
    lexical_attribute_value_range(tag, wanted).is_some()
}

fn lexical_attribute_value_range(tag: &[u8], wanted: &[u8]) -> Option<std::ops::Range<usize>> {
    lexical_attribute_ranges(tag, wanted).map(|(_, value)| value)
}

fn lexical_attribute_ranges(
    tag: &[u8],
    wanted: &[u8],
) -> Option<(std::ops::Range<usize>, std::ops::Range<usize>)> {
    let mut cursor = tag.iter().position(|byte| byte.is_ascii_whitespace())?;
    while cursor < tag.len() {
        let whole_start = cursor;
        while tag.get(cursor).is_some_and(u8::is_ascii_whitespace) {
            cursor += 1;
        }
        if matches!(tag.get(cursor), Some(b'>' | b'/')) {
            return None;
        }
        let name_start = cursor;
        while tag
            .get(cursor)
            .is_some_and(|byte| !byte.is_ascii_whitespace() && !matches!(byte, b'=' | b'>' | b'/'))
        {
            cursor += 1;
        }
        let name = &tag[name_start..cursor];
        while tag.get(cursor).is_some_and(u8::is_ascii_whitespace) {
            cursor += 1;
        }
        if tag.get(cursor) != Some(&b'=') {
            return None;
        }
        cursor += 1;
        while tag.get(cursor).is_some_and(u8::is_ascii_whitespace) {
            cursor += 1;
        }
        let quote = *tag.get(cursor)?;
        if !matches!(quote, b'\"' | b'\'') {
            return None;
        }
        cursor += 1;
        let value_start = cursor;
        while tag.get(cursor) != Some(&quote) {
            cursor += 1;
            if cursor >= tag.len() {
                return None;
            }
        }
        let value_end = cursor;
        cursor += 1;
        if name == wanted {
            return Some((whole_start..cursor, value_start..value_end));
        }
    }
    None
}

fn property_slot(element: &BytesStart<'_>, prefixes: &[String]) -> Option<usize> {
    [
        b"name".as_slice(),
        b"style".as_slice(),
        b"category".as_slice(),
        b"types".as_slice(),
        b"behaviors".as_slice(),
        b"description".as_slice(),
        b"guid".as_slice(),
    ]
    .iter()
    .position(|local| is_word_element(element.name().as_ref(), local, prefixes))
}

fn apply_property_field(
    value: &mut CT_DocPartPr,
    slot: usize,
    raw: &[u8],
    prefixes: &[String],
) -> Result<()> {
    match slot {
        0 => value.name = root_value(raw, prefixes)?,
        1 => value.style = root_value(raw, prefixes)?,
        2 => {
            let (category, gallery) = category_values(raw, prefixes)?;
            value.category = category;
            value.gallery = gallery;
        }
        3 => value.types = nested_values(raw, b"type", prefixes)?,
        4 => value.behaviors = nested_values(raw, b"behavior", prefixes)?,
        5 => value.description = root_value(raw, prefixes)?,
        6 => {
            let guid = root_value(raw, prefixes)?;
            if guid.as_deref().is_some_and(|guid| !valid_guid(guid)) {
                return Err(OxmlError::InvalidValue(
                    "invalid w:guid lexical form".to_owned(),
                ));
            }
            value.guid = guid;
        }
        _ => {}
    }
    Ok(())
}

fn root_value(raw: &[u8], inherited_prefixes: &[String]) -> Result<Option<String>> {
    let mut reader = Reader::from_reader(raw);
    let mut buffer = Vec::new();
    let mut prefixes = inherited_prefixes.to_vec();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(element) | Event::Empty(element) => {
                prefixes = word_prefixes_at(&element, &prefixes)?;
                return Ok(Some(word_value(&element, &prefixes)?.ok_or_else(|| {
                    OxmlError::MissingElement("modeled glossary w:val".to_owned())
                })?));
            }
            Event::Eof => return Ok(None),
            _ => {}
        }
        buffer.clear();
    }
}

fn nested_values(raw: &[u8], wanted: &[u8], inherited_prefixes: &[String]) -> Result<Vec<String>> {
    let mut reader = Reader::from_reader(raw);
    let mut buffer = Vec::new();
    let mut prefixes = inherited_prefixes.to_vec();
    let mut prefix_scopes = Vec::new();
    let mut depth = 0usize;
    let mut values = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(element) => {
                let local_prefixes = word_prefixes_at(&element, &prefixes)?;
                if depth == 1 && is_word_element(element.name().as_ref(), wanted, &local_prefixes) {
                    values.push(word_value(&element, &local_prefixes)?.ok_or_else(|| {
                        OxmlError::MissingElement("modeled glossary child w:val".to_owned())
                    })?);
                }
                depth += 1;
                prefix_scopes.push(std::mem::replace(&mut prefixes, local_prefixes));
            }
            Event::Empty(element) => {
                let local_prefixes = word_prefixes_at(&element, &prefixes)?;
                if depth == 1 && is_word_element(element.name().as_ref(), wanted, &local_prefixes) {
                    values.push(word_value(&element, &local_prefixes)?.ok_or_else(|| {
                        OxmlError::MissingElement("modeled glossary child w:val".to_owned())
                    })?);
                }
            }
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                prefixes = prefix_scopes
                    .pop()
                    .unwrap_or_else(|| inherited_prefixes.to_vec());
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    if values.is_empty() {
        return Err(OxmlError::MissingElement(format!(
            "modeled glossary collection requires at least one w:{} child",
            String::from_utf8_lossy(wanted)
        )));
    }
    if wanted == b"type" && values.iter().any(|value| !valid_doc_part_type(value)) {
        return Err(OxmlError::InvalidValue(
            "invalid w:type enumeration".to_owned(),
        ));
    }
    if wanted == b"behavior" && values.iter().any(|value| !valid_doc_part_behavior(value)) {
        return Err(OxmlError::InvalidValue(
            "invalid w:behavior enumeration".to_owned(),
        ));
    }
    Ok(values)
}

fn category_values(
    raw: &[u8],
    inherited_prefixes: &[String],
) -> Result<(Option<String>, Option<String>)> {
    let mut reader = Reader::from_reader(raw);
    let mut buffer = Vec::new();
    let mut prefixes = inherited_prefixes.to_vec();
    let mut prefix_scopes = Vec::new();
    let mut depth = 0usize;
    let mut last_slot = None;
    let mut values = [None, None];
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(element) => {
                let local_prefixes = word_prefixes_at(&element, &prefixes)?;
                validate_category_child(
                    &element,
                    &local_prefixes,
                    depth,
                    &mut last_slot,
                    &mut values,
                )?;
                depth += 1;
                prefix_scopes.push(std::mem::replace(&mut prefixes, local_prefixes));
            }
            Event::Empty(element) => {
                let local_prefixes = word_prefixes_at(&element, &prefixes)?;
                validate_category_child(
                    &element,
                    &local_prefixes,
                    depth,
                    &mut last_slot,
                    &mut values,
                )?;
            }
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                prefixes = prefix_scopes
                    .pop()
                    .unwrap_or_else(|| inherited_prefixes.to_vec());
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    if values.iter().any(Option::is_none) {
        return Err(OxmlError::MissingElement(
            "w:category requires direct w:name and w:gallery children".to_owned(),
        ));
    }
    if values[1]
        .as_deref()
        .is_some_and(|gallery| !valid_doc_part_gallery(gallery))
    {
        return Err(OxmlError::InvalidValue(
            "invalid w:gallery enumeration".to_owned(),
        ));
    }
    Ok((values[0].take(), values[1].take()))
}

fn valid_doc_part_behavior(value: &str) -> bool {
    matches!(value, "content" | "p" | "pg")
}

fn valid_doc_part_type(value: &str) -> bool {
    matches!(
        value,
        "none" | "normal" | "autoExp" | "toolbar" | "speller" | "formFld" | "bbPlcHdr"
    )
}

fn valid_guid(value: &str) -> bool {
    let bytes = value.as_bytes();
    bytes.len() == 38
        && bytes.first() == Some(&b'{')
        && bytes.last() == Some(&b'}')
        && [9, 14, 19, 24]
            .into_iter()
            .all(|index| bytes.get(index) == Some(&b'-'))
        && bytes.iter().enumerate().all(|(index, byte)| {
            matches!(index, 0 | 9 | 14 | 19 | 24 | 37) || matches!(byte, b'0'..=b'9' | b'A'..=b'F')
        })
}

fn valid_doc_part_gallery(value: &str) -> bool {
    matches!(
        value,
        "placeholder"
            | "any"
            | "default"
            | "docParts"
            | "coverPg"
            | "eq"
            | "ftrs"
            | "hdrs"
            | "pgNum"
            | "tbls"
            | "watermarks"
            | "autoTxt"
            | "txtBox"
            | "pgNumT"
            | "pgNumB"
            | "pgNumMargins"
            | "tblOfContents"
            | "bib"
            | "custQuickParts"
            | "custCoverPg"
            | "custEq"
            | "custFtrs"
            | "custHdrs"
            | "custPgNum"
            | "custTbls"
            | "custWatermarks"
            | "custAutoTxt"
            | "custTxtBox"
            | "custPgNumT"
            | "custPgNumB"
            | "custPgNumMargins"
            | "custTblOfContents"
            | "custBib"
            | "custom1"
            | "custom2"
            | "custom3"
            | "custom4"
            | "custom5"
    )
}

fn validate_category_child(
    element: &BytesStart<'_>,
    prefixes: &[String],
    depth: usize,
    last_slot: &mut Option<usize>,
    values: &mut [Option<String>; 2],
) -> Result<()> {
    if depth != 1 {
        return Ok(());
    }
    let slot = if is_word_element(element.name().as_ref(), b"name", prefixes) {
        Some(0)
    } else if is_word_element(element.name().as_ref(), b"gallery", prefixes) {
        Some(1)
    } else {
        None
    };
    let Some(slot) = slot else {
        return Ok(());
    };
    if values[slot].is_some() {
        return Err(OxmlError::InvalidValue(
            "duplicate modeled w:category child".to_owned(),
        ));
    }
    if last_slot.is_some_and(|last| slot < last) {
        return Err(OxmlError::InvalidValue(
            "out-of-order modeled w:category child".to_owned(),
        ));
    }
    values[slot] = Some(
        word_value(element, prefixes)?
            .ok_or_else(|| OxmlError::MissingElement("modeled glossary child w:val".to_owned()))?,
    );
    *last_slot = Some(slot);
    Ok(())
}

fn word_value(element: &BytesStart<'_>, prefixes: &[String]) -> Result<Option<String>> {
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute?;
        if is_word_attribute(attribute.key.as_ref(), b"val", prefixes) {
            if value.is_some() {
                return Err(OxmlError::InvalidValue(
                    "duplicate modeled glossary w:val attribute".to_owned(),
                ));
            }
            let raw = std::str::from_utf8(&attribute.value)?;
            value = Some(
                quick_xml::escape::unescape(raw)
                    .map_err(|error| {
                        OxmlError::InvalidValue(format!(
                            "invalid modeled glossary w:val attribute: {error}"
                        ))
                    })?
                    .into_owned(),
            );
        }
    }
    Ok(value)
}

fn parse_doc_part_body(raw: &[u8], inherited_prefixes: &[String]) -> Result<CT_Body> {
    let mut reader = Reader::from_reader(raw);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut prefixes = inherited_prefixes.to_vec();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(element) => {
                prefixes = word_prefixes_at(&element, &prefixes)?;
                if is_word_element(element.name().as_ref(), b"docPartBody", &prefixes) {
                    return CT_Body::from_xml_with_prefixes_and_owner_bindings(
                        &mut reader,
                        &prefixes,
                        &[],
                    );
                }
            }
            Event::Empty(element) => {
                prefixes = word_prefixes_at(&element, &prefixes)?;
                if is_word_element(element.name().as_ref(), b"docPartBody", &prefixes) {
                    return Ok(CT_Body {
                        content: Vec::new(),
                        sect_pr: None,
                    });
                }
            }
            Event::Eof => {
                return Err(OxmlError::MissingElement("w:docPartBody".to_owned()));
            }
            _ => {}
        }
        buffer.clear();
    }
}

fn serialize_doc_part_body(body: &CT_Body, source: &[u8]) -> Result<Vec<u8>> {
    let mut writer = Writer::new(Vec::new());
    body.to_xml(&mut writer)?;
    let xml = writer.into_inner();
    let content = xml
        .strip_prefix(b"<w:body>")
        .and_then(|xml| xml.strip_suffix(b"</w:body>"))
        .ok_or_else(|| OxmlError::MissingElement("serialized w:body wrapper".to_owned()))?;
    let (root_start, root_end) = expanded_root_wrapper(source)?;
    let mut output = root_start;
    output.extend_from_slice(content);
    output.extend_from_slice(&root_end);
    Ok(output)
}

fn expanded_root_wrapper(source: &[u8]) -> Result<(Vec<u8>, Vec<u8>)> {
    let mut reader = Reader::from_reader(source);
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(element) => {
                let end = reader.buffer_position() as usize;
                let name = std::str::from_utf8(element.name().as_ref())?.to_owned();
                return Ok((source[..end].to_vec(), format!("</{name}>").into_bytes()));
            }
            Event::Empty(element) => {
                let end = reader.buffer_position() as usize;
                let mut start = source[..end].to_vec();
                let slash = start
                    .iter()
                    .rposition(|byte| *byte == b'/')
                    .ok_or_else(|| OxmlError::MissingElement("empty root close".to_owned()))?;
                start.remove(slash);
                let name = std::str::from_utf8(element.name().as_ref())?.to_owned();
                return Ok((start, format!("</{name}>").into_bytes()));
            }
            Event::Eof => return Err(OxmlError::MissingElement("root element".to_owned())),
            _ => {}
        }
        buffer.clear();
    }
}

fn apply_structural_edits(
    input: &[u8],
    mut edits: Vec<(Range<usize>, Vec<u8>)>,
) -> Result<Vec<u8>> {
    edits.sort_by_key(|(span, _)| (span.start, span.end));
    let replacement_bytes: usize = edits.iter().map(|(_, replacement)| replacement.len()).sum();
    let removed_bytes = edits.iter().map(|(span, _)| span.len()).sum::<usize>();
    let mut output = Vec::with_capacity(input.len() + replacement_bytes - removed_bytes);
    let mut cursor = 0usize;
    for (span, replacement) in edits {
        if span.start > span.end || span.end > input.len() || span.start < cursor {
            return Err(OxmlError::InvalidValue(
                "overlapping or invalid retained glossary source spans".to_owned(),
            ));
        }
        output.extend_from_slice(&input[cursor..span.start]);
        output.extend_from_slice(&replacement);
        cursor = span.end;
    }
    output.extend_from_slice(&input[cursor..]);
    Ok(output)
}

fn word_local_name(element: &BytesStart<'_>, prefixes: &[String]) -> Option<String> {
    let name = element.name();
    let local = name.as_ref().rsplit(|byte| *byte == b':').next()?;
    is_word_element(name.as_ref(), local, prefixes)
        .then(|| String::from_utf8_lossy(local).into_owned())
}

fn standalone_event(event: Event<'static>) -> Result<Vec<u8>> {
    let mut writer = Writer::new(Vec::new());
    if !matches!(event, Event::Decl(_) | Event::DocType(_)) {
        writer.write_event(event)?;
    }
    Ok(writer.into_inner())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::text::CT_P;

    #[test]
    fn glossary_aliases_reparse_and_changed_properties_keep_raw_subtrees() {
        let xml = br#"<?xml version="1.0"?><q:glossaryDocument xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:test"><q:docParts><q:docPart keep="root"><q:docPartPr><q:name q:val="Greeting"/><q:types><q:type q:val="autoExp"/></q:types><q:description q:val="old"/><x:property keep="yes"/></q:docPartPr><q:docPartBody><q:p><q:r><q:t>Hello</q:t></q:r></q:p><x:body keep="yes"/></q:docPartBody></q:docPart></q:docParts></q:glossaryDocument>"#;
        let mut glossary = CT_GlossaryDocument::from_xml(xml).unwrap();
        assert_eq!(
            glossary.doc_parts[0].properties.name.as_deref(),
            Some("Greeting")
        );
        glossary.doc_parts[0].properties.description = Some("new".to_owned());
        let output = glossary.to_xml().unwrap();
        let text = std::str::from_utf8(&output).unwrap();
        assert!(text.contains(r#"<x:property keep="yes"/>"#));
        assert!(text.contains(r#"<x:body keep="yes"/>"#));
        assert!(text.contains(r#"q:val="new""#));
        CT_GlossaryDocument::from_xml(&output).unwrap();
    }

    #[test]
    fn changed_glossary_property_keeps_unmodelled_root_attributes() {
        let xml = br#"<w:glossaryDocument xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:test"><w:docParts><w:docPart><w:docPartPr x:producer="keep"><w:name w:val="entry"/></w:docPartPr><w:docPartBody><w:p/></w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>"#;
        let mut glossary = CT_GlossaryDocument::from_xml(xml).unwrap();
        glossary.doc_parts[0].properties.description = Some("changed".to_owned());
        let output = String::from_utf8(glossary.to_xml().unwrap()).unwrap();
        assert!(output.contains(r#"<w:docPartPr x:producer="keep""#));
        assert!(output.contains(r#"<w:description w:val="changed"/>"#));
    }

    #[test]
    fn metadata_only_glossary_change_keeps_body_bytes_exact() {
        let body = r#"<q:docPartBody x:producer="keep">
  <q:p><q:r><q:t> exact </q:t></q:r></q:p><x:raw/>
</q:docPartBody>"#;
        let xml = format!(
            r#"<q:glossaryDocument xmlns:q="{W_NS}" xmlns:x="urn:test"><q:docParts><q:docPart><q:docPartPr><q:name q:val="entry"/></q:docPartPr>{body}</q:docPart></q:docParts></q:glossaryDocument>"#
        );
        let mut glossary = CT_GlossaryDocument::from_xml(xml.as_bytes()).unwrap();
        glossary.doc_parts[0].properties.description = Some("changed".to_owned());
        let output = String::from_utf8(glossary.to_xml().unwrap()).unwrap();
        assert!(output.contains(body));
    }

    #[test]
    fn sibling_namespace_shadow_does_not_hide_a_later_glossary_entry() {
        let xml = br#"<q:glossaryDocument xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:test"><q:docParts><x:raw xmlns:q="urn:producer"/><q:docPart><q:docPartPr><q:name q:val="entry"/></q:docPartPr><q:docPartBody><q:p/></q:docPartBody></q:docPart></q:docParts></q:glossaryDocument>"#;
        let glossary = CT_GlossaryDocument::from_xml(xml).unwrap();
        assert_eq!(glossary.doc_parts.len(), 1);
        assert_eq!(
            glossary.doc_parts[0].properties.name.as_deref(),
            Some("entry")
        );
    }

    #[test]
    fn changed_glossary_property_preserves_its_unmodelled_xml_surgically() {
        let xml = br#"<q:glossaryDocument xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:test"><q:docParts><q:docPart><q:docPartPr><q:name q:val="entry"/><q:category x:producer="keep"><q:name x:leaf="keep" q:val="old"><x:child/></q:name><q:gallery q:val="autoTxt"/></q:category><q:description x:scalar="keep" q:val="old"><x:description-child/></q:description></q:docPartPr><q:docPartBody/></q:docPart></q:docParts></q:glossaryDocument>"#;
        let mut glossary = CT_GlossaryDocument::from_xml(xml).unwrap();
        glossary.doc_parts[0].properties.category = Some("new".to_owned());
        glossary.doc_parts[0].properties.description = Some("new description".to_owned());
        let output = String::from_utf8(glossary.to_xml().unwrap()).unwrap();
        assert!(output.contains(r#"<q:category x:producer="keep">"#));
        assert!(output.contains(r#"x:leaf="keep" q:val="new"><x:child/>"#));
        assert!(
            output.contains(r#"x:scalar="keep" q:val="new description"><x:description-child/>"#)
        );
        let reopened = CT_GlossaryDocument::from_xml(output.as_bytes()).unwrap();
        assert_eq!(
            reopened.doc_parts[0].properties.category.as_deref(),
            Some("new")
        );
    }

    #[test]
    fn foreign_w_binding_uses_the_existing_alternate_word_prefix_for_changes() {
        let xml = br#"<q:glossaryDocument xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><q:docParts><q:docPart><q:docPartPr xmlns:w="urn:producer"><q:name q:val="entry"/></q:docPartPr><q:docPartBody/></q:docPart></q:docParts></q:glossaryDocument>"#;
        let mut glossary = CT_GlossaryDocument::from_xml(xml).unwrap();
        glossary.doc_parts[0].properties.description = Some("changed".to_owned());
        let output = String::from_utf8(glossary.to_xml().unwrap()).unwrap();
        assert!(output.contains(r#"xmlns:w="urn:producer""#));
        assert!(output.contains(r#"<q:description q:val="changed"/>"#));
        let reopened = CT_GlossaryDocument::from_xml(output.as_bytes()).unwrap();
        assert_eq!(
            reopened.doc_parts[0].properties.description.as_deref(),
            Some("changed")
        );
    }

    #[test]
    fn empty_doc_part_body_does_not_invent_section_properties() {
        let xml = br#"<w:glossaryDocument xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docParts><w:docPart><w:docPartPr><w:name w:val="entry"/></w:docPartPr><w:docPartBody/></w:docPart></w:docParts></w:glossaryDocument>"#;
        let mut glossary = CT_GlossaryDocument::from_xml(xml).unwrap();
        assert!(glossary.doc_parts[0].body.sect_pr.is_none());
        glossary.doc_parts[0].body.add_paragraph(CT_P::new());
        let output = String::from_utf8(glossary.to_xml().unwrap()).unwrap();
        assert!(!output.contains("sectPr"));
        assert!(CT_GlossaryDocument::from_xml(output.as_bytes()).is_ok());
    }

    #[test]
    fn changed_body_preserves_root_attributes_and_namespace_context() {
        let xml = br#"<w:glossaryDocument xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docParts><w:docPart><w:docPartPr><w:name w:val="entry"/></w:docPartPr><w:docPartBody x:producer="keep" xmlns:x="urn:producer"><w:p/><x:retained/></w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>"#;
        let mut glossary = CT_GlossaryDocument::from_xml(xml).unwrap();
        glossary.doc_parts[0].body.add_paragraph(CT_P::new());
        let output = String::from_utf8(glossary.to_xml().unwrap()).unwrap();
        assert!(output.contains(r#"<w:docPartBody x:producer="keep" xmlns:x="urn:producer">"#));
        assert!(output.contains("<x:retained/>"));
        CT_GlossaryDocument::from_xml(output.as_bytes()).unwrap();
    }

    #[test]
    fn changed_property_handles_greater_than_inside_quoted_attribute() {
        let xml = br#"<w:glossaryDocument xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:producer"><w:docParts><w:docPart><w:docPartPr><w:name w:val="entry"/><w:description x:producer="a>b" w:val="old"/></w:docPartPr><w:docPartBody/></w:docPart></w:docParts></w:glossaryDocument>"#;
        let mut glossary = CT_GlossaryDocument::from_xml(xml).unwrap();
        glossary.doc_parts[0].properties.description = Some("changed".to_owned());
        let output = String::from_utf8(glossary.to_xml().unwrap()).unwrap();
        assert!(output.contains(r#"<w:description x:producer="a>b" w:val="changed"/>"#));
        let reopened = CT_GlossaryDocument::from_xml(output.as_bytes()).unwrap();
        assert_eq!(
            reopened.doc_parts[0].properties.description.as_deref(),
            Some("changed")
        );
    }

    #[test]
    fn nested_glossary_values_remain_unmodelled() {
        let wrapped = r#"<x:wrapper><w:type w:val="autoExp"/></x:wrapper>"#;
        let xml = format!(
            r#"<w:glossaryDocument xmlns:w="{W_NS}" xmlns:x="urn:producer"><w:docParts><w:docPart><w:docPartPr><w:name w:val="entry"/><w:types><w:type w:val="normal"/>{wrapped}</w:types></w:docPartPr><w:docPartBody/></w:docPart></w:docParts></w:glossaryDocument>"#
        );
        let mut glossary = CT_GlossaryDocument::from_xml(xml.as_bytes()).unwrap();
        assert_eq!(glossary.doc_parts[0].properties.types, ["normal"]);
        glossary.doc_parts[0].properties.description = Some("changed".to_owned());
        let output = String::from_utf8(glossary.to_xml().unwrap()).unwrap();
        assert!(output.contains(wrapped));
    }

    #[test]
    fn added_repeated_value_uses_container_local_word_prefix() {
        let xml = br#"<q:glossaryDocument xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><q:docParts><q:docPart><q:docPartPr><q:name q:val="entry"/><r:types xmlns:r="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:q="urn:producer"><r:type r:val="bbPlcHdr"/></r:types></q:docPartPr><q:docPartBody/></q:docPart></q:docParts></q:glossaryDocument>"#;
        let mut glossary = CT_GlossaryDocument::from_xml(xml).unwrap();
        glossary.doc_parts[0].properties.types = vec!["bbPlcHdr".to_owned(), "autoExp".to_owned()];
        let output = String::from_utf8(glossary.to_xml().unwrap()).unwrap();
        assert!(output.contains(r#"<r:type r:val="autoExp"/>"#));
        let reopened = CT_GlossaryDocument::from_xml(output.as_bytes()).unwrap();
        assert_eq!(
            reopened.doc_parts[0].properties.types,
            ["bbPlcHdr", "autoExp"]
        );
    }

    #[test]
    fn glossary_entry_without_doc_part_properties_round_trips_without_inventing_it() {
        let xml = br#"<w:glossaryDocument xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docParts><w:docPart><w:docPartBody><w:p/></w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>"#;
        let glossary = CT_GlossaryDocument::from_xml(xml).unwrap();
        assert_eq!(glossary.doc_parts.len(), 1);
        assert_eq!(glossary.to_xml().unwrap(), xml);
    }

    #[test]
    fn glossary_parser_rejects_multiple_top_level_roots() {
        let xml = br#"<w:glossaryDocument xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docParts/></w:glossaryDocument><w:glossaryDocument xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docParts/></w:glossaryDocument>"#;
        assert!(CT_GlossaryDocument::from_xml(xml).is_err());
    }

    #[test]
    fn duplicate_category_child_is_rejected_during_parse() {
        let xml = br#"<w:glossaryDocument xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docParts><w:docPart><w:docPartPr><w:name w:val="entry"/><w:category><w:name w:val="first"/><w:name w:val="duplicate"/></w:category></w:docPartPr><w:docPartBody/></w:docPart></w:docParts></w:glossaryDocument>"#;
        assert!(CT_GlossaryDocument::from_xml(xml).is_err());
    }

    #[test]
    fn glossary_parser_rejects_non_whitespace_text_outside_the_root() {
        let trailing = br#"<w:glossaryDocument xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docParts/></w:glossaryDocument>garbage"#;
        let leading = br#"garbage<w:glossaryDocument xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docParts/></w:glossaryDocument>"#;
        assert!(CT_GlossaryDocument::from_xml(trailing).is_err());
        assert!(CT_GlossaryDocument::from_xml(leading).is_err());
    }

    #[test]
    fn glossary_requires_exactly_one_direct_doc_parts_container() {
        let missing = br#"<w:glossaryDocument xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:glossaryDocument>"#;
        let duplicate = br#"<w:glossaryDocument xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docParts/><w:docParts/></w:glossaryDocument>"#;
        let nested_only = br#"<w:glossaryDocument xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:customXml><w:docParts/></w:customXml></w:glossaryDocument>"#;
        assert!(CT_GlossaryDocument::from_xml(missing).is_err());
        assert!(CT_GlossaryDocument::from_xml(duplicate).is_err());
        assert!(CT_GlossaryDocument::from_xml(nested_only).is_err());
    }

    #[test]
    fn removing_glossary_values_removes_their_complete_property_elements() {
        let xml = br#"<w:glossaryDocument xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docParts><w:docPart><w:docPartPr><w:name w:val="entry"/><w:category><w:name w:val="category"/><w:gallery w:val="autoTxt"/></w:category><w:types><w:type w:val="autoExp"/><w:type w:val="bbPlcHdr"/></w:types><w:behaviors><w:behavior w:val="content"/><w:behavior w:val="pg"/></w:behaviors><w:description w:val="description"/><w:guid w:val="{01234567-89AB-CDEF-0123-456789ABCDEF}"/></w:docPartPr><w:docPartBody/></w:docPart></w:docParts></w:glossaryDocument>"#;
        let mut glossary = CT_GlossaryDocument::from_xml(xml).unwrap();
        let properties = &mut glossary.doc_parts[0].properties;
        properties.category = None;
        properties.gallery = None;
        properties.types.clear();
        properties.behaviors.truncate(1);
        properties.description = None;
        properties.guid = None;
        let output = String::from_utf8(glossary.to_xml().unwrap()).unwrap();
        for forbidden in [
            "<w:category",
            "<w:type",
            "<w:description",
            "<w:guid",
            "w:val=\"pg\"",
        ] {
            assert!(!output.contains(forbidden), "{forbidden}: {output}");
        }
        assert!(output.contains(r#"<w:behavior w:val="content"/>"#));
        CT_GlossaryDocument::from_xml(output.as_bytes()).unwrap();
    }

    #[test]
    fn duplicate_direct_doc_part_properties_or_bodies_are_rejected() {
        let duplicate_properties = br#"<w:glossaryDocument xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docParts><w:docPart><w:docPartPr/><w:docPartPr/><w:docPartBody/></w:docPart></w:docParts></w:glossaryDocument>"#;
        let duplicate_bodies = br#"<w:glossaryDocument xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docParts><w:docPart><w:docPartPr/><w:docPartBody/><w:docPartBody/></w:docPart></w:docParts></w:glossaryDocument>"#;
        assert!(CT_GlossaryDocument::from_xml(duplicate_properties).is_err());
        assert!(CT_GlossaryDocument::from_xml(duplicate_bodies).is_err());
    }

    #[test]
    fn glossary_replacement_uses_structural_spans_not_identical_comment_bytes() {
        let entry = r#"<w:docPart><w:docPartPr><w:name w:val="entry"/><w:description w:val="old"/></w:docPartPr><w:docPartBody/></w:docPart>"#;
        let xml = format!(
            r#"<w:glossaryDocument xmlns:w="{W_NS}"><w:docParts><!--{entry}-->{entry}</w:docParts></w:glossaryDocument>"#
        );
        let mut glossary = CT_GlossaryDocument::from_xml(xml.as_bytes()).unwrap();
        glossary.doc_parts[0].properties.description = Some("changed".to_owned());
        let output = String::from_utf8(glossary.to_xml().unwrap()).unwrap();
        assert!(output.contains(&format!("<!--{entry}-->")));
        assert_eq!(output.matches(r#"w:val="changed""#).count(), 1);

        let property = r#"<w:description w:val="old"/>"#;
        let nested = format!(
            r#"<w:glossaryDocument xmlns:w="{W_NS}"><w:docParts><w:docPart><w:docPartPr><!--{property}--><w:name w:val="entry"/>{property}</w:docPartPr><w:docPartBody/></w:docPart></w:docParts></w:glossaryDocument>"#
        );
        let mut glossary = CT_GlossaryDocument::from_xml(nested.as_bytes()).unwrap();
        glossary.doc_parts[0].properties.description = Some("changed".to_owned());
        let output = String::from_utf8(glossary.to_xml().unwrap()).unwrap();
        assert!(output.contains(&format!("<!--{property}-->")));
        assert_eq!(output.matches(r#"w:val="changed""#).count(), 1);
    }

    #[test]
    fn glossary_parser_rejects_character_references_outside_the_root() {
        let root = format!(
            r#"<w:glossaryDocument xmlns:w="{W_NS}"><w:docParts><w:docPart><w:docPartBody/></w:docPart></w:docParts></w:glossaryDocument>"#
        );
        for xml in [format!("&#65;{root}"), format!("{root}&#65;")] {
            assert!(
                CT_GlossaryDocument::from_xml(xml.as_bytes()).is_err(),
                "{xml}"
            );
        }
        let valid = format!(r#"<?xml version="1.0"?>{root}"#);
        assert!(CT_GlossaryDocument::from_xml(valid.as_bytes()).is_ok());
    }

    #[test]
    fn doc_part_properties_must_precede_the_body() {
        let xml = format!(
            r#"<w:glossaryDocument xmlns:w="{W_NS}"><w:docParts><w:docPart><w:docPartBody/><w:docPartPr><w:name w:val="entry"/></w:docPartPr></w:docPart></w:docParts></w:glossaryDocument>"#
        );
        assert!(CT_GlossaryDocument::from_xml(xml.as_bytes()).is_err());
    }

    #[test]
    fn duplicate_modeled_singleton_glossary_properties_are_rejected() {
        for property in [
            r#"<w:name w:val="entry"/>"#,
            r#"<w:style w:val="Normal"/>"#,
            r#"<w:category><w:name w:val="category"/></w:category>"#,
            r#"<w:types><w:type w:val="autoExp"/></w:types>"#,
            r#"<w:behaviors><w:behavior w:val="content"/></w:behaviors>"#,
            r#"<w:description w:val="description"/>"#,
            r#"<w:guid w:val="guid"/>"#,
        ] {
            let xml = format!(
                r#"<w:glossaryDocument xmlns:w="{W_NS}"><w:docParts><w:docPart><w:docPartPr>{property}{property}</w:docPartPr><w:docPartBody/></w:docPart></w:docParts></w:glossaryDocument>"#
            );
            assert!(
                CT_GlossaryDocument::from_xml(xml.as_bytes()).is_err(),
                "{property}"
            );
        }
    }

    #[test]
    fn out_of_order_modeled_glossary_properties_are_rejected() {
        let xml = format!(
            r#"<w:glossaryDocument xmlns:w="{W_NS}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="entry"/><w:category><w:name w:val="category"/></w:category><w:style w:val="Normal"/></w:docPartPr><w:docPartBody/></w:docPart></w:docParts></w:glossaryDocument>"#
        );
        assert!(CT_GlossaryDocument::from_xml(xml.as_bytes()).is_err());
    }

    #[test]
    fn valueless_modeled_glossary_properties_are_rejected() {
        for properties in [
            r#"<w:name w:val="entry"/><w:description/>"#,
            r#"<w:name w:val="entry"/><w:types><w:type/></w:types>"#,
            r#"<w:name w:val="entry"/><w:category><w:name/></w:category>"#,
        ] {
            let xml = format!(
                r#"<w:glossaryDocument xmlns:w="{W_NS}"><w:docParts><w:docPart><w:docPartPr>{properties}</w:docPartPr><w:docPartBody/></w:docPart></w:docParts></w:glossaryDocument>"#
            );
            assert!(
                CT_GlossaryDocument::from_xml(xml.as_bytes()).is_err(),
                "{properties}"
            );
        }
    }

    #[test]
    fn nested_category_children_require_schema_order_and_singletons() {
        for category in [
            r#"<w:category><w:name w:val="one"/><w:name w:val="two"/></w:category>"#,
            r#"<w:category><w:gallery w:val="one"/><w:gallery w:val="two"/></w:category>"#,
            r#"<w:category><w:gallery w:val="gallery"/><w:name w:val="name"/></w:category>"#,
        ] {
            let xml = format!(
                r#"<w:glossaryDocument xmlns:w="{W_NS}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="entry"/>{category}</w:docPartPr><w:docPartBody/></w:docPart></w:docParts></w:glossaryDocument>"#
            );
            assert!(
                CT_GlossaryDocument::from_xml(xml.as_bytes()).is_err(),
                "{category}"
            );
        }
    }

    #[test]
    fn glossary_declarations_and_doctypes_require_document_positions() {
        let root = format!(
            r#"<w:glossaryDocument xmlns:w="{W_NS}"><w:docParts><w:docPart><w:docPartBody/></w:docPart></w:docParts></w:glossaryDocument>"#
        );
        for xml in [
            format!(r#"{root}<?xml version="1.0"?>"#),
            format!(
                r#"<w:glossaryDocument xmlns:w="{W_NS}"><?xml version="1.0"?><w:docParts/></w:glossaryDocument>"#
            ),
            format!(r#"{root}<!DOCTYPE glossaryDocument>"#),
            format!(r#"<!DOCTYPE glossaryDocument><!DOCTYPE glossaryDocument>{root}"#),
        ] {
            assert!(
                CT_GlossaryDocument::from_xml(xml.as_bytes()).is_err(),
                "{xml}"
            );
        }
    }

    #[test]
    fn glossary_document_type_declarations_fail_closed() {
        let root = format!(
            r#"<w:glossaryDocument xmlns:w="{W_NS}"><w:docParts><w:docPart><w:docPartBody/></w:docPart></w:docParts></w:glossaryDocument>"#
        );
        for (case, declaration) in [
            ("uppercase-simple", "<!DOCTYPE glossaryDocument>"),
            ("lowercase-keyword", "<!doctype glossaryDocument>"),
            ("invalid-root-name", "<!DOCTYPE 1producer>"),
            (
                "external-system-identifier",
                r#"<!DOCTYPE glossaryDocument SYSTEM "urn:producer">"#,
            ),
            (
                "external-public-identifier",
                r#"<!DOCTYPE glossaryDocument PUBLIC "producer" "urn:producer">"#,
            ),
            (
                "internal-subset",
                "<!DOCTYPE glossaryDocument [<!ELEMENT glossaryDocument ANY>]>",
            ),
            (
                "truncated-internal-subset",
                "<!DOCTYPE glossaryDocument [<!ELEMENT glossaryDocument ANY>",
            ),
        ] {
            let xml = format!("{declaration}{root}");
            assert!(
                CT_GlossaryDocument::from_xml(xml.as_bytes()).is_err(),
                "{case}: {xml}"
            );
        }
    }

    #[test]
    fn glossary_xml_declarations_require_valid_pseudo_attributes() {
        let root = format!(
            r#"<w:glossaryDocument xmlns:w="{W_NS}"><w:docParts><w:docPart><w:docPartBody/></w:docPart></w:docParts></w:glossaryDocument>"#
        );
        for declaration in [
            r#"<?xml?>"#,
            r#"<?xml encoding="UTF-8" version="1.0"?>"#,
            r#"<?xml version="1.0" version="1.0"?>"#,
            r#"<?xml version="1.0" encoding="UTF-8" encoding="UTF-8"?>"#,
            r#"<?xml version="1.0" standalone="maybe"?>"#,
            r#"<?xml version="1.0" unknown="value"?>"#,
        ] {
            let xml = format!("{declaration}{root}");
            assert!(
                CT_GlossaryDocument::from_xml(xml.as_bytes()).is_err(),
                "{xml}"
            );
        }
        let valid = format!(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>{root}"#);
        assert!(CT_GlossaryDocument::from_xml(valid.as_bytes()).is_ok());
    }

    #[test]
    fn ooxml_glossary_rejects_xml_1_1_declarations() {
        let xml = format!(
            r#"<?xml version="1.1"?><w:glossaryDocument xmlns:w="{W_NS}"><w:docParts><w:docPart><w:docPartBody/></w:docPart></w:docParts></w:glossaryDocument>"#
        );
        assert!(CT_GlossaryDocument::from_xml(xml.as_bytes()).is_err());
    }

    #[test]
    fn categories_require_name_and_gallery_as_a_complete_pair() {
        for category in [
            r#"<w:category><w:name w:val="category"/></w:category>"#,
            r#"<w:category><w:gallery w:val="gallery"/></w:category>"#,
            r#"<w:category/>"#,
        ] {
            let xml = format!(
                r#"<w:glossaryDocument xmlns:w="{W_NS}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="entry"/>{category}</w:docPartPr><w:docPartBody/></w:docPart></w:docParts></w:glossaryDocument>"#
            );
            assert!(
                CT_GlossaryDocument::from_xml(xml.as_bytes()).is_err(),
                "{category}"
            );
        }
    }

    #[test]
    fn required_glossary_collections_must_not_be_empty() {
        for xml in [
            format!(r#"<w:glossaryDocument xmlns:w="{W_NS}"><w:docParts/></w:glossaryDocument>"#),
            format!(
                r#"<w:glossaryDocument xmlns:w="{W_NS}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="entry"/><w:types/></w:docPartPr><w:docPartBody/></w:docPart></w:docParts></w:glossaryDocument>"#
            ),
            format!(
                r#"<w:glossaryDocument xmlns:w="{W_NS}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="entry"/><w:behaviors/></w:docPartPr><w:docPartBody/></w:docPart></w:docParts></w:glossaryDocument>"#
            ),
        ] {
            assert!(
                CT_GlossaryDocument::from_xml(xml.as_bytes()).is_err(),
                "{xml}"
            );
        }
    }

    #[test]
    fn glossary_gallery_and_behavior_values_require_schema_enumerations() {
        for properties in [
            r#"<w:category><w:name w:val="category"/><w:gallery w:val="not-a-gallery"/></w:category>"#,
            r#"<w:behaviors><w:behavior w:val="not-a-behavior"/></w:behaviors>"#,
        ] {
            let xml = format!(
                r#"<w:glossaryDocument xmlns:w="{W_NS}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="entry"/>{properties}</w:docPartPr><w:docPartBody/></w:docPart></w:docParts></w:glossaryDocument>"#
            );
            assert!(
                CT_GlossaryDocument::from_xml(xml.as_bytes()).is_err(),
                "{properties}"
            );
        }
    }

    #[test]
    fn glossary_types_require_schema_enumerations() {
        let invalid = format!(
            r#"<w:glossaryDocument xmlns:w="{W_NS}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="entry"/><w:types><w:type w:val="not-a-type"/></w:types></w:docPartPr><w:docPartBody/></w:docPart></w:docParts></w:glossaryDocument>"#
        );
        assert!(CT_GlossaryDocument::from_xml(invalid.as_bytes()).is_err());

        let values = [
            "none", "normal", "autoExp", "toolbar", "speller", "formFld", "bbPlcHdr",
        ];
        let types = values
            .iter()
            .map(|value| format!(r#"<w:type w:val="{value}"/>"#))
            .collect::<String>();
        let valid = format!(
            r#"<w:glossaryDocument xmlns:w="{W_NS}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="entry"/><w:types>{types}</w:types></w:docPartPr><w:docPartBody/></w:docPart></w:docParts></w:glossaryDocument>"#
        );
        assert!(CT_GlossaryDocument::from_xml(valid.as_bytes()).is_ok());
    }

    #[test]
    fn glossary_guid_requires_braced_uppercase_lexical_form() {
        for guid in [
            "not-a-guid",
            "01234567-89AB-CDEF-0123-456789ABCDEF",
            "{01234567-89ab-CDEF-0123-456789ABCDEF}",
            "{01234567-89AB-CDEF-0123-456789ABCDE}",
        ] {
            let xml = format!(
                r#"<w:glossaryDocument xmlns:w="{W_NS}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="entry"/><w:guid w:val="{guid}"/></w:docPartPr><w:docPartBody/></w:docPart></w:docParts></w:glossaryDocument>"#
            );
            assert!(
                CT_GlossaryDocument::from_xml(xml.as_bytes()).is_err(),
                "{guid}"
            );
        }
        let valid = format!(
            r#"<w:glossaryDocument xmlns:w="{W_NS}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="entry"/><w:guid w:val="{{01234567-89AB-CDEF-0123-456789ABCDEF}}"/></w:docPartPr><w:docPartBody/></w:docPart></w:docParts></w:glossaryDocument>"#
        );
        assert!(CT_GlossaryDocument::from_xml(valid.as_bytes()).is_ok());
    }

    #[test]
    fn typed_glossary_values_reject_duplicate_or_unescape_failing_attributes() {
        for property in [
            r#"<w:name xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="one" q:val="two"/>"#,
            r#"<w:name w:val="&undefined;"/>"#,
        ] {
            let xml = format!(
                r#"<w:glossaryDocument xmlns:w="{W_NS}"><w:docParts><w:docPart><w:docPartPr>{property}</w:docPartPr><w:docPartBody/></w:docPart></w:docParts></w:glossaryDocument>"#
            );
            assert!(
                CT_GlossaryDocument::from_xml(xml.as_bytes()).is_err(),
                "{property}"
            );
        }
        let malformed = BytesStart::from_content("w:name w:val", 6);
        assert!(word_value(&malformed, &["w".to_owned()]).is_err());
    }

    #[test]
    fn glossary_element_only_containers_reject_character_data() {
        let entry = concat!(
            r#"<w:docPart><w:docPartPr><w:name w:val="entry"/></w:docPartPr>"#,
            r#"<w:docPartBody/></w:docPart>"#,
        );
        for xml in [
            format!(
                r#"<w:glossaryDocument xmlns:w="{W_NS}">text<w:docParts>{entry}</w:docParts></w:glossaryDocument>"#
            ),
            format!(
                r#"<w:glossaryDocument xmlns:w="{W_NS}"><w:docParts><![CDATA[text]]>{entry}</w:docParts></w:glossaryDocument>"#
            ),
            format!(
                r#"<w:glossaryDocument xmlns:w="{W_NS}"><w:docParts>&amp;{entry}</w:docParts></w:glossaryDocument>"#
            ),
        ] {
            assert!(
                CT_GlossaryDocument::from_xml(xml.as_bytes()).is_err(),
                "{xml}"
            );
        }
        for content in ["&#x20;", "<![CDATA[ ]]>"] {
            let xml = format!(
                r#"<w:glossaryDocument xmlns:w="{W_NS}"><w:docParts>{content}{entry}</w:docParts></w:glossaryDocument>"#
            );
            assert!(
                CT_GlossaryDocument::from_xml(xml.as_bytes()).is_ok(),
                "{xml}"
            );
        }
    }

    #[test]
    fn nested_glossary_element_only_containers_reject_character_data() {
        for entry in [
            r#"<w:docPart>text<w:docPartPr><w:name w:val="entry"/></w:docPartPr><w:docPartBody/></w:docPart>"#,
            r#"<w:docPart><w:docPartPr>text<w:name w:val="entry"/></w:docPartPr><w:docPartBody/></w:docPart>"#,
            r#"<w:docPart><w:docPartPr><w:name w:val="entry"/><w:category><![CDATA[text]]><w:name w:val="cat"/><w:gallery w:val="autoTxt"/></w:category></w:docPartPr><w:docPartBody/></w:docPart>"#,
            r#"<w:docPart><w:docPartPr><w:name w:val="entry"/><w:types>&amp;<w:type w:val="normal"/></w:types></w:docPartPr><w:docPartBody/></w:docPart>"#,
            r#"<w:docPart><w:docPartPr><w:name w:val="entry"/><w:behaviors>text<w:behavior w:val="content"/></w:behaviors></w:docPartPr><w:docPartBody/></w:docPart>"#,
        ] {
            let xml = format!(
                r#"<w:glossaryDocument xmlns:w="{W_NS}"><w:docParts>{entry}</w:docParts></w:glossaryDocument>"#
            );
            assert!(
                CT_GlossaryDocument::from_xml(xml.as_bytes()).is_err(),
                "{entry}"
            );
        }
    }

    #[test]
    fn glossary_forbidden_comment_bodies_fail_closed() {
        let xml = format!(
            r#"<w:glossaryDocument xmlns:w="{W_NS}"><!--producer--invalid--><w:docParts><w:docPart><w:docPartPr><w:name w:val="entry"/></w:docPartPr><w:docPartBody/></w:docPart></w:docParts></w:glossaryDocument>"#
        );
        assert!(CT_GlossaryDocument::from_xml(xml.as_bytes()).is_err());
    }

    #[test]
    fn glossary_processing_instruction_targets_require_xml_names() {
        for instruction in ["<?XML version=\"1.0\"?>", "<?1producer value?>"] {
            let xml = format!(
                r#"{instruction}<w:glossaryDocument xmlns:w="{W_NS}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="entry"/></w:docPartPr><w:docPartBody/></w:docPart></w:docParts></w:glossaryDocument>"#
            );
            assert!(
                CT_GlossaryDocument::from_xml(xml.as_bytes()).is_err(),
                "{instruction}"
            );
        }
    }

    #[test]
    fn glossary_xml_names_bindings_and_expanded_attributes_fail_closed() {
        for malformed in [
            r#"<1producer/>"#,
            r#"<producer:item/>"#,
            r#"<w:name 1producer="value" w:val="entry"/>"#,
            r#"<w:name producer:value="opaque" w:val="entry"/>"#,
            r#"<w:name xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="one" q:val="two"/>"#,
        ] {
            let xml = format!(
                r#"<w:glossaryDocument xmlns:w="{W_NS}"><w:docParts><w:docPart><w:docPartPr>{malformed}<w:name w:val="entry"/></w:docPartPr><w:docPartBody/></w:docPart></w:docParts></w:glossaryDocument>"#
            );
            assert!(
                CT_GlossaryDocument::from_xml(xml.as_bytes()).is_err(),
                "{malformed}"
            );
        }
    }

    #[test]
    fn glossary_nested_references_and_literal_xml_characters_fail_closed() {
        for malformed in [
            "<x:raw>&undefined;</x:raw>".to_owned(),
            "<x:raw>&#xFFFE;</x:raw>".to_owned(),
            "<x:raw>producer\u{1}</x:raw>".to_owned(),
            "<x:raw><![CDATA[producer\u{FFFE}]]></x:raw>".to_owned(),
        ] {
            let xml = format!(
                r#"<w:glossaryDocument xmlns:w="{W_NS}" xmlns:x="urn:producer"><w:docParts><w:docPart><w:docPartPr><w:name w:val="entry"/>{malformed}</w:docPartPr><w:docPartBody/></w:docPart></w:docParts></w:glossaryDocument>"#
            );
            assert!(
                CT_GlossaryDocument::from_xml(xml.as_bytes()).is_err(),
                "{malformed:?}"
            );
        }
    }
}
