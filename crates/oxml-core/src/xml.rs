//! Shared OOXML namespace and attribute helpers.

use std::collections::HashSet;

use quick_xml::XmlVersion;
use quick_xml::errors::{Error as QuickXmlError, IllFormedError, SyntaxError};
use quick_xml::events::{BytesDecl, BytesRef, BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use crate::error::Result;

/// A strict XML 1.0 lexical validation failure.
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum XmlLexicalError {
    /// The input is not UTF-8.
    InvalidUtf8,
    /// The XML declaration does not satisfy the XML 1.0 grammar.
    InvalidDeclaration(String),
    /// The input contains a literal character forbidden by XML 1.0.
    ForbiddenLiteralCharacter,
    /// An element, attribute, or processing instruction name is invalid.
    InvalidName(String),
    /// A namespace declaration or qualified-name binding is invalid.
    InvalidNamespace(String),
    /// Two attributes resolve to the same namespace and local name.
    DuplicateExpandedAttribute,
    /// An entity or character reference is invalid.
    InvalidReference(String),
    /// A processing instruction target is invalid or reserved.
    InvalidProcessingInstruction(String),
    /// A comment is not lexically valid.
    InvalidComment(String),
}

/// Validate the format-neutral lexical rules required by strict OOXML readers.
pub fn validate_strict_xml_1_0(xml: &[u8]) -> std::result::Result<(), XmlLexicalError> {
    let text = std::str::from_utf8(xml).map_err(|_| XmlLexicalError::InvalidUtf8)?;
    if !text.chars().all(xml_1_0_character_is_valid) {
        return Err(XmlLexicalError::ForbiddenLiteralCharacter);
    }

    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_comments = true;
    let mut buffer = Vec::new();
    loop {
        let (_, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(classify_reader_error)?;
        let event = event.into_owned();
        validate_lexical_event(&reader, &event)?;
        if matches!(event, Event::Eof) {
            return Ok(());
        }
        buffer.clear();
    }
}

fn classify_reader_error(error: QuickXmlError) -> XmlLexicalError {
    let message = error.to_string();
    match error {
        QuickXmlError::Syntax(SyntaxError::UnclosedComment)
        | QuickXmlError::IllFormed(IllFormedError::DoubleHyphenInComment) => {
            XmlLexicalError::InvalidComment(message)
        }
        QuickXmlError::Syntax(SyntaxError::UnclosedPI) => {
            XmlLexicalError::InvalidProcessingInstruction(message)
        }
        QuickXmlError::Syntax(SyntaxError::UnclosedXmlDecl)
        | QuickXmlError::IllFormed(IllFormedError::MissingDeclVersion(_))
        | QuickXmlError::IllFormed(IllFormedError::UnknownVersion) => {
            XmlLexicalError::InvalidDeclaration(message)
        }
        QuickXmlError::IllFormed(IllFormedError::UnclosedReference) | QuickXmlError::Escape(_) => {
            XmlLexicalError::InvalidReference(message)
        }
        QuickXmlError::Namespace(_) => XmlLexicalError::InvalidNamespace(message),
        _ => XmlLexicalError::InvalidName(message),
    }
}

fn validate_lexical_event(
    reader: &NsReader<&[u8]>,
    event: &Event<'_>,
) -> std::result::Result<(), XmlLexicalError> {
    match event {
        Event::Decl(declaration) => validate_xml_declaration(declaration),
        Event::Start(element) | Event::Empty(element) => validate_element(reader, element),
        Event::End(element) => {
            let name = element.name();
            let prefix = validate_xml_qname(name.as_ref())?;
            validate_bound_prefix(&reader.resolver().resolve_element(name).0, prefix).map(|_| ())
        }
        Event::GeneralRef(reference) => validate_reference(reference).map(|_| ()),
        Event::PI(instruction) => validate_processing_instruction(instruction.target()),
        _ => Ok(()),
    }
}

fn validate_xml_declaration(
    declaration: &BytesDecl<'_>,
) -> std::result::Result<(), XmlLexicalError> {
    if declaration
        .xml_version()
        .map_err(|error| XmlLexicalError::InvalidDeclaration(error.to_string()))?
        != XmlVersion::Explicit1_0
    {
        return Err(XmlLexicalError::InvalidDeclaration(
            "version must be 1.0".to_owned(),
        ));
    }
    let content = std::str::from_utf8(declaration.as_ref())
        .map_err(|error| XmlLexicalError::InvalidDeclaration(error.to_string()))?;
    let start = BytesStart::from_content(content, 3);
    let attributes = start
        .attributes()
        .collect::<std::result::Result<Vec<_>, _>>()
        .map_err(|error| XmlLexicalError::InvalidDeclaration(error.to_string()))?;
    if attributes.first().map(|attribute| attribute.key.as_ref()) != Some(b"version".as_ref()) {
        return Err(XmlLexicalError::InvalidDeclaration(
            "must begin with version".to_owned(),
        ));
    }
    let mut encoding_seen = false;
    let mut standalone_seen = false;
    for (index, attribute) in attributes.iter().enumerate() {
        let value = std::str::from_utf8(attribute.value.as_ref())
            .map_err(|error| XmlLexicalError::InvalidDeclaration(error.to_string()))?;
        match attribute.key.as_ref() {
            b"version" if index == 0 && value == "1.0" => {}
            b"encoding" if !encoding_seen && !standalone_seen && encoding_name_is_valid(value) => {
                encoding_seen = true;
            }
            b"standalone" if !standalone_seen && matches!(value, "yes" | "no") => {
                standalone_seen = true;
            }
            _ => {
                return Err(XmlLexicalError::InvalidDeclaration(
                    "attributes are invalid, duplicated, or out of order".to_owned(),
                ));
            }
        }
    }
    Ok(())
}

fn encoding_name_is_valid(value: &str) -> bool {
    let mut bytes = value.bytes();
    bytes.next().is_some_and(|byte| byte.is_ascii_alphabetic())
        && bytes.all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'.' | b'_' | b'-'))
}

fn validate_element(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> std::result::Result<(), XmlLexicalError> {
    let element_name = element.name();
    let prefix = validate_xml_qname(element_name.as_ref())?;
    if prefix == Some(b"xmlns".as_slice()) {
        return Err(XmlLexicalError::InvalidNamespace(
            "element uses the reserved xmlns prefix".to_owned(),
        ));
    }
    validate_bound_prefix(&reader.resolver().resolve_element(element_name).0, prefix)?;

    let mut expanded_names = HashSet::new();
    for attribute in element.attributes() {
        let attribute =
            attribute.map_err(|error| XmlLexicalError::InvalidName(error.to_string()))?;
        let name = attribute.key.as_ref();
        let prefix = validate_xml_qname(name)?;
        if attribute.value.contains(&b'<') {
            return Err(XmlLexicalError::InvalidReference(
                "attribute contains a literal less-than sign".to_owned(),
            ));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, element.decoder())
            .map_err(|error| XmlLexicalError::InvalidReference(error.to_string()))?;
        if !value.chars().all(xml_1_0_character_is_valid) {
            return Err(XmlLexicalError::ForbiddenLiteralCharacter);
        }
        if name == b"xmlns" {
            validate_namespace_declaration(None, value.as_bytes())?;
            continue;
        }
        if prefix == Some(b"xmlns".as_slice()) {
            validate_namespace_declaration(Some(local_name(name)), value.as_bytes())?;
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let resolved = validate_bound_prefix(&namespace, prefix)?;
        if !expanded_names.insert((resolved, local.as_ref().to_vec())) {
            return Err(XmlLexicalError::DuplicateExpandedAttribute);
        }
    }
    Ok(())
}

fn validate_namespace_declaration(
    prefix: Option<&[u8]>,
    namespace: &[u8],
) -> std::result::Result<(), XmlLexicalError> {
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
        Err(XmlLexicalError::InvalidNamespace(
            "invalid namespace declaration".to_owned(),
        ))
    }
}

fn validate_bound_prefix(
    namespace: &ResolveResult<'_>,
    prefix: Option<&[u8]>,
) -> std::result::Result<Option<Vec<u8>>, XmlLexicalError> {
    match namespace {
        ResolveResult::Bound(Namespace(namespace)) => Ok(Some(namespace.to_vec())),
        ResolveResult::Unbound if prefix.is_none() => Ok(None),
        ResolveResult::Unbound => Err(XmlLexicalError::InvalidNamespace(format!(
            "unbound namespace prefix {}",
            String::from_utf8_lossy(prefix.unwrap_or_default())
        ))),
        ResolveResult::Unknown(prefix) => Err(XmlLexicalError::InvalidNamespace(format!(
            "unbound namespace prefix {}",
            String::from_utf8_lossy(prefix)
        ))),
    }
}

fn validate_xml_qname(name: &[u8]) -> std::result::Result<Option<&[u8]>, XmlLexicalError> {
    let name = std::str::from_utf8(name)
        .map_err(|error| XmlLexicalError::InvalidName(error.to_string()))?;
    let mut parts = name.split(':');
    let first = parts.next().unwrap_or_default();
    let second = parts.next();
    if !xml_ncname_is_valid(first)
        || second.is_some_and(|local| !xml_ncname_is_valid(local))
        || parts.next().is_some()
    {
        return Err(XmlLexicalError::InvalidName(format!(
            "qualified name {name}"
        )));
    }
    Ok(second.map(|_| first.as_bytes()))
}

fn validate_xml_name(name: &[u8]) -> std::result::Result<(), XmlLexicalError> {
    let name = std::str::from_utf8(name)
        .map_err(|error| XmlLexicalError::InvalidName(error.to_string()))?;
    let mut characters = name.chars();
    if characters
        .next()
        .is_some_and(|character| character == ':' || xml_ncname_start_character(character))
        && characters.all(|character| character == ':' || xml_ncname_character(character))
    {
        Ok(())
    } else {
        Err(XmlLexicalError::InvalidName(format!("name {name}")))
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

fn xml_1_0_character_is_valid(character: char) -> bool {
    matches!(character, '\t' | '\n' | '\r')
        || ('\u{20}'..='\u{D7FF}').contains(&character)
        || ('\u{E000}'..='\u{FFFD}').contains(&character)
        || ('\u{10000}'..='\u{10FFFF}').contains(&character)
}

fn validate_reference(reference: &BytesRef<'_>) -> std::result::Result<char, XmlLexicalError> {
    if let Some(character) = reference
        .resolve_char_ref()
        .map_err(|error| XmlLexicalError::InvalidReference(error.to_string()))?
    {
        if xml_1_0_character_is_valid(character) {
            return Ok(character);
        }
        return Err(XmlLexicalError::InvalidReference(
            "character reference is not legal in XML 1.0".to_owned(),
        ));
    }
    let name = reference
        .decode()
        .map_err(|error| XmlLexicalError::InvalidReference(error.to_string()))?;
    match name.as_ref() {
        "amp" => Ok('&'),
        "lt" => Ok('<'),
        "gt" => Ok('>'),
        "apos" => Ok('\''),
        "quot" => Ok('"'),
        _ => Err(XmlLexicalError::InvalidReference(format!(
            "undeclared entity reference &{name};"
        ))),
    }
}

fn validate_processing_instruction(target: &[u8]) -> std::result::Result<(), XmlLexicalError> {
    validate_xml_name(target).map_err(|error| match error {
        XmlLexicalError::InvalidName(message) => {
            XmlLexicalError::InvalidProcessingInstruction(message)
        }
        other => XmlLexicalError::InvalidProcessingInstruction(format!("{other:?}")),
    })?;
    if target.eq_ignore_ascii_case(b"xml") {
        return Err(XmlLexicalError::InvalidProcessingInstruction(
            "reserved XML target".to_owned(),
        ));
    }
    Ok(())
}

/// Relationships namespace.
pub const R_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/// Markup Compatibility namespace.
pub const MC_NS: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

/// Return the local portion of a possibly prefixed XML name.
pub fn local_name(name: &[u8]) -> &[u8] {
    match name.iter().position(|&byte| byte == b':') {
        Some(pos) => &name[pos + 1..],
        None => name,
    }
}

/// Check whether an XML name has the expected local portion.
pub fn matches_local_name(name: &[u8], expected: &[u8]) -> bool {
    local_name(name) == expected
}

/// Return a named attribute value, matching with or without a prefix.
pub fn get_attr(element: &BytesStart<'_>, name: &[u8]) -> Option<String> {
    element
        .attributes()
        .flatten()
        .find(|attr| matches_local_name(attr.key.as_ref(), name))
        .and_then(|attr| std::str::from_utf8(&attr.value).ok().map(str::to_owned))
}

/// Return non-`vt` prefixed namespace declarations needed by raw XML children.
pub(crate) fn extra_namespace_declarations(
    element: &BytesStart<'_>,
) -> Result<Vec<(String, String)>> {
    let mut declarations = Vec::new();
    for attribute in element.attributes() {
        let attribute = attribute?;
        let key = attribute.key.as_ref();
        if key.starts_with(b"xmlns:") && key != b"xmlns:vt" {
            let name = std::str::from_utf8(key)?.to_owned();
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, element.decoder())?
                .into_owned();
            declarations.push((name, value));
        }
    }
    Ok(declarations)
}

#[cfg(test)]
mod tests {
    use quick_xml::Reader;
    use quick_xml::events::Event;

    use super::*;

    #[test]
    fn local_names_match_with_or_without_a_prefix() {
        assert_eq!(local_name(b"w:document"), b"document");
        assert_eq!(local_name(b"document"), b"document");
        assert!(matches_local_name(b"p:sld", b"sld"));
        assert!(!matches_local_name(b"p:sld", b"slide"));
    }

    #[test]
    fn attributes_match_with_or_without_a_prefix() {
        let mut reader = Reader::from_str(r#"<item r:id="rId7" plain="value"/>"#);
        let mut buf = Vec::new();
        let Event::Empty(element) = reader.read_event_into(&mut buf).unwrap() else {
            panic!("expected empty element");
        };

        assert_eq!(get_attr(&element, b"id").as_deref(), Some("rId7"));
        assert_eq!(get_attr(&element, b"plain").as_deref(), Some("value"));
        assert_eq!(get_attr(&element, b"missing"), None);
    }

    #[test]
    fn strict_xml_1_0_validator_rejects_every_shared_lexical_class() {
        let malformed = [
            (&b"\xff"[..], XmlLexicalError::InvalidUtf8),
            (
                &b"<?xml encoding=\"UTF-8\"?><root/>"[..],
                XmlLexicalError::InvalidDeclaration(String::new()),
            ),
            (
                &b"<?xml version=\"1.0\" encoding=\"UT&#70;-8\"?><root/>"[..],
                XmlLexicalError::InvalidDeclaration(String::new()),
            ),
            (
                &b"<root>\x01</root>"[..],
                XmlLexicalError::ForbiddenLiteralCharacter,
            ),
            (
                &b"<1root/>"[..],
                XmlLexicalError::InvalidName(String::new()),
            ),
            (
                &b"<p:root/>"[..],
                XmlLexicalError::InvalidNamespace(String::new()),
            ),
            (
                &b"<root xmlns:a=\"urn:same\" xmlns:b=\"urn:same\" a:id=\"1\" b:id=\"2\"/>"[..],
                XmlLexicalError::DuplicateExpandedAttribute,
            ),
            (
                &b"<root>&undefined;</root>"[..],
                XmlLexicalError::InvalidReference(String::new()),
            ),
            (
                &b"<?XML value?><root/>"[..],
                XmlLexicalError::InvalidProcessingInstruction(String::new()),
            ),
            (
                &b"<root><!--bad--comment--></root>"[..],
                XmlLexicalError::InvalidComment(String::new()),
            ),
        ];

        for (xml, expected) in malformed {
            let actual = validate_strict_xml_1_0(xml);
            assert!(
                actual.as_ref().is_err_and(
                    |error| std::mem::discriminant(error) == std::mem::discriminant(&expected)
                ),
                "{xml:?}: {actual:?}"
            );
        }
    }
}
