//! OOXML namespace constants.

use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, QName, ResolveResult};
use quick_xml::reader::NsReader;

use crate::error::Result;

/// WordprocessingML main namespace
pub const W_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
/// WordprocessingML namespace prefix
pub const W_PREFIX: &[u8] = b"w";
/// Transitional OfficeMath namespace.
pub const M_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/math";
/// Canonical OfficeMath namespace prefix.
pub const M_PREFIX: &[u8] = b"m";

pub use oxml_core::xml::{MC_NS, R_NS, matches_local_name};

/// Whether elements named one of `locals` nest more than `limit` levels deep
/// in `xml`, found without recursing.
///
/// An element counts whatever its namespace, which can only over-count.
/// Malformed XML counts up to the error, so the parser still reports it.
pub(crate) fn nesting_exceeds(xml: &[u8], locals: &[&[u8]], limit: usize) -> bool {
    let mut reader = quick_xml::Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut open_counts = Vec::new();
    let mut depth = 0usize;
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(start)) => {
                let counts = locals
                    .iter()
                    .any(|local| matches_local_name(start.name().as_ref(), local));
                if counts {
                    depth += 1;
                    if depth > limit {
                        return true;
                    }
                }
                open_counts.push(counts);
            }
            Ok(Event::End(_)) => {
                if open_counts.pop() == Some(true) {
                    depth -= 1;
                }
            }
            Ok(Event::Eof) | Err(_) => return false,
            Ok(_) => {}
        }
        buffer.clear();
    }
}

/// The prefixes that an `mc:Ignorable` or `mc:MustUnderstand` attribute in
/// `xml` lists without a namespace declaration in scope.
///
/// Each finding is the attribute's local name and the prefix, once per pair,
/// in document order. Markup Compatibility (ECMA-376 Part 3) requires every
/// listed prefix to be declared, so a part with a finding is not conformant.
pub fn undeclared_compatibility_prefixes(xml: &[u8]) -> Result<Vec<(String, String)>> {
    let mut reader = NsReader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut findings: Vec<(String, String)> = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(element) | Event::Empty(element) => {
                for attribute in element.attributes() {
                    let attribute = attribute?;
                    let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
                    let local = std::str::from_utf8(local.as_ref())?;
                    if namespace != ResolveResult::Bound(Namespace(MC_NS.as_bytes()))
                        || !matches!(local, "Ignorable" | "MustUnderstand")
                    {
                        continue;
                    }
                    let value = attribute
                        .decoded_and_normalized_value(XmlVersion::Implicit1_0, element.decoder())?;
                    for prefix in value.split_ascii_whitespace() {
                        let probe = format!("{prefix}:probe");
                        let (bound, _) = reader.resolver().resolve_element(QName(probe.as_bytes()));
                        if !matches!(bound, ResolveResult::Bound(_))
                            && !findings
                                .iter()
                                .any(|(name, known)| name == local && known == prefix)
                        {
                            findings.push((local.to_owned(), prefix.to_owned()));
                        }
                    }
                }
            }
            Event::Eof => return Ok(findings),
            _ => {}
        }
        buffer.clear();
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn compatibility_prefixes_must_be_declared_in_scope() {
        let xml = br#"<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:compat="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" compat:Ignorable="w14 w15"><w:comment w:id="0" xmlns:w16="urn:w16" compat:Ignorable="w16 w14" compat:MustUnderstand="w17"/><w:comment w:id="1" compat:Ignorable="w16"/></w:comments>"#;
        assert_eq!(
            undeclared_compatibility_prefixes(xml).unwrap(),
            [
                ("Ignorable".to_owned(), "w14".to_owned()),
                ("MustUnderstand".to_owned(), "w17".to_owned()),
                ("Ignorable".to_owned(), "w16".to_owned()),
            ]
        );

        // An attribute named Ignorable outside the Markup Compatibility
        // namespace lists nothing.
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="urn:other" mc:Ignorable="w14" Ignorable="w14"/>"#;
        assert_eq!(undeclared_compatibility_prefixes(xml).unwrap(), []);
        assert!(undeclared_compatibility_prefixes(b"<a><b></a>").is_err());
    }
}
