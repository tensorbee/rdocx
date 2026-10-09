use std::collections::HashMap;

use oxml_core::OxmlError;
use oxml_drawing::namespace::A_NS;
use quick_xml::events::{BytesStart, Event};
use quick_xml::{Reader, Writer, XmlVersion};

/// PresentationML main namespace URI.
pub const P_NS: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";

/// Fixed PresentationML prefix used when writing XML.
pub const P_PREFIX: &str = "p";

/// Office document relationships namespace URI.
pub const R_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/// Fixed office document relationships prefix used when writing XML.
pub const R_PREFIX: &str = "r";

/// Markup Compatibility namespace URI.
pub const MC_NS: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

pub(crate) const FIXED_MODEL_PREFIXES: &[&str] = &["p", "a", "r"];
pub(crate) const FIXED_SHAPE_TREE_PREFIXES: &[&str] = &["p", "a", "r", "mc"];

#[derive(Clone, Debug, Default)]
pub(crate) struct NamespaceBindings {
    default: Option<String>,
    prefixes: HashMap<String, String>,
}

impl NamespaceBindings {
    pub(crate) fn from_entries(entries: &[(String, String)]) -> Self {
        let mut bindings = Self::default();
        for (prefix, uri) in entries {
            if prefix.is_empty() {
                bindings.default = Some(uri.clone());
            } else {
                bindings.prefixes.insert(prefix.clone(), uri.clone());
            }
        }
        bindings
    }

    pub(crate) fn with_start(&self, start: &BytesStart<'_>) -> Result<Self, OxmlError> {
        let mut bindings = self.clone();
        for (name, value) in all_attributes(start)? {
            if name == "xmlns" {
                bindings.default = Some(value);
            } else if let Some(prefix) = name.strip_prefix("xmlns:") {
                bindings.prefixes.insert(prefix.to_owned(), value);
            }
        }
        Ok(bindings)
    }

    pub(crate) fn element_uri<'a>(&'a self, name: &[u8]) -> Option<&'a str> {
        match qname_prefix(name) {
            Some(prefix) => self.prefixes.get(prefix).map(String::as_str),
            None => self.default.as_deref(),
        }
    }

    pub(crate) fn attribute_uri<'a>(&'a self, name: &[u8]) -> Option<&'a str> {
        qname_prefix(name).and_then(|prefix| self.prefixes.get(prefix).map(String::as_str))
    }

    pub(crate) fn entries(&self) -> Vec<(String, String)> {
        let mut entries: Vec<_> = self
            .prefixes
            .iter()
            .map(|(prefix, uri)| (prefix.clone(), uri.clone()))
            .collect();
        sort_namespace_entries(&mut entries);
        if let Some(uri) = &self.default {
            entries.push((String::new(), uri.clone()));
        }
        entries
    }

    pub(crate) fn reject_writer_conflicts(&self, fixed_prefixes: &[&str]) -> Result<(), OxmlError> {
        for prefix in fixed_prefixes {
            let expected = canonical_uri(prefix).expect("fixed prefixes have canonical URIs");
            if let Some(actual) = self.prefixes.get(*prefix)
                && actual != expected
            {
                return Err(OxmlError::InvalidValue(format!(
                    "xmlns:{prefix} conflicts with the fixed writer namespace"
                )));
            }
        }
        Ok(())
    }
}

pub(crate) fn all_attributes(start: &BytesStart<'_>) -> Result<Vec<(String, String)>, OxmlError> {
    let mut attributes = Vec::new();
    for attribute in start.attributes() {
        let attribute = attribute?;
        let name = std::str::from_utf8(attribute.key.as_ref())?.to_owned();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, start.decoder())?
            .replace(['\r', '\n', '\t'], " ");
        attributes.push((name, value));
    }
    Ok(attributes)
}

pub(crate) fn non_visual_drawing_id(xml: &[u8]) -> Option<u32> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer).ok()? {
            Event::Start(start) | Event::Empty(start) => {
                return all_attributes(&start)
                    .ok()?
                    .into_iter()
                    .find(|(name, _)| name == "id")
                    .and_then(|(_, value)| value.parse().ok());
            }
            Event::Eof => return None,
            _ => {}
        }
        buffer.clear();
    }
}

pub(crate) fn non_visual_drawing_name(start: &BytesStart<'_>) -> Result<Option<String>, OxmlError> {
    Ok(all_attributes(start)?
        .into_iter()
        .find(|(name, _)| name == "name")
        .map(|(_, value)| value))
}

pub(crate) fn set_non_visual_drawing_name(xml: &mut Vec<u8>, name: &str) -> Result<(), OxmlError> {
    oxml_core::xml::reject_non_xml_characters("shape name", name)?;
    set_non_visual_drawing_attribute(xml, "name", Some(name))
}

/// Returns one unqualified attribute of a `p:cNvPr` fragment, decoded.
pub(crate) fn non_visual_drawing_attribute(xml: &[u8], name: &str) -> Option<String> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer).ok()? {
            Event::Start(start) | Event::Empty(start) => {
                return all_attributes(&start)
                    .ok()?
                    .into_iter()
                    .find_map(|(key, value)| (key == name).then_some(value));
            }
            Event::Eof => return None,
            _ => {}
        }
        buffer.clear();
    }
}

/// Sets one unqualified attribute of a `p:cNvPr` fragment in place, appends
/// it when absent, or removes it for `None`, keeping every child as written.
pub(crate) fn set_non_visual_drawing_attribute(
    xml: &mut Vec<u8>,
    name: &str,
    value: Option<&str>,
) -> Result<(), OxmlError> {
    let mut reader = Reader::from_reader(xml.as_slice());
    let mut buffer = Vec::new();
    loop {
        let event = reader.read_event_into(&mut buffer)?;
        let is_empty = matches!(&event, Event::Empty(_));
        match event {
            Event::Start(start) | Event::Empty(start) => {
                let end = reader.buffer_position() as usize;
                let qualified_name = std::str::from_utf8(start.name().as_ref())?.to_owned();
                let mut replacement = BytesStart::new(qualified_name);
                let mut replaced = false;
                for attribute in start.attributes().with_checks(false) {
                    let attribute = attribute?;
                    if attribute.key.as_ref() == name.as_bytes() {
                        if let Some(value) = value {
                            replacement.push_attribute((name, value));
                        }
                        replaced = true;
                    } else {
                        replacement.push_attribute(attribute);
                    }
                }
                if let (false, Some(value)) = (replaced, value) {
                    replacement.push_attribute((name, value));
                }
                let mut writer = Writer::new(Vec::new());
                if is_empty {
                    writer.write_event(Event::Empty(replacement))?;
                } else {
                    writer.write_event(Event::Start(replacement))?;
                    writer.get_mut().extend_from_slice(&xml[end..]);
                }
                *xml = writer.into_inner();
                return Ok(());
            }
            Event::Eof => {
                return Err(OxmlError::MissingElement(
                    "non-visual drawing properties".to_owned(),
                ));
            }
            _ => {}
        }
        buffer.clear();
    }
}

/// The `a:ext` URI of the Office 2019 decorative flag on `p:cNvPr`.
const DECORATIVE_EXTENSION_URI: &str = "{C183D7F6-B498-43B3-948B-1728B52AA6E4}";

/// The extension PowerPoint writes for a shape marked decorative, which
/// screen readers then skip.
const DECORATIVE_EXTENSION: &str = r#"<a:ext uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}"><adec:decorative xmlns:adec="http://schemas.microsoft.com/office/drawing/2017/decorative" val="1"/></a:ext>"#;

/// Returns whether a `p:cNvPr` fragment carries the decorative extension with
/// a true `val`.
pub(crate) fn non_visual_drawing_decorative(xml: &[u8]) -> bool {
    let Some((_, children, _)) = split_element(xml) else {
        return false;
    };
    children
        .iter()
        .filter(|child| element_local_name(child).as_deref() == Some(b"extLst".as_slice()))
        .filter_map(|list| split_element(list))
        .flat_map(|(_, extensions, _)| extensions)
        .filter(|extension| {
            element_attribute(extension, "uri").as_deref() == Some(DECORATIVE_EXTENSION_URI)
        })
        .filter_map(|extension| split_element(&extension).map(|(_, flags, _)| flags))
        .flatten()
        .any(|flag| {
            element_local_name(&flag).as_deref() == Some(b"decorative".as_slice())
                && matches!(
                    element_attribute(&flag, "val").as_deref(),
                    Some("1" | "true")
                )
        })
}

/// Marks a `p:cNvPr` fragment decorative, or removes the mark, keeping its
/// other children and extensions in place.
///
/// The extension goes into the existing `a:extLst`, or into a new one after
/// the hyperlinks, where the schema places it. Removing the last extension
/// removes the list.
pub(crate) fn set_non_visual_drawing_decorative(
    xml: &mut Vec<u8>,
    decorative: bool,
) -> Result<(), OxmlError> {
    let (start, mut children, end) = split_element(xml)
        .ok_or_else(|| OxmlError::MissingElement("non-visual drawing properties".to_owned()))?;
    let list_index = children
        .iter()
        .position(|child| element_local_name(child).as_deref() == Some(b"extLst".as_slice()));
    let is_decorative = |extension: &Vec<u8>| {
        element_attribute(extension, "uri").as_deref() == Some(DECORATIVE_EXTENSION_URI)
    };
    match (list_index, decorative) {
        (Some(index), _) => {
            let (list_start, mut extensions, list_end) = split_element(&children[index])
                .ok_or_else(|| OxmlError::MissingElement("a:extLst".to_owned()))?;
            extensions.retain(|extension| !is_decorative(extension));
            if decorative {
                extensions.push(DECORATIVE_EXTENSION.as_bytes().to_vec());
            }
            if extensions
                .iter()
                .any(|extension| element_local_name(extension).is_some())
            {
                children[index] = join_element(&list_start, &extensions, &list_end);
            } else {
                children.remove(index);
            }
        }
        (None, true) => {
            children.push(format!("<a:extLst>{DECORATIVE_EXTENSION}</a:extLst>").into_bytes())
        }
        (None, false) => return Ok(()),
    }
    *xml = join_element(&start, &children, &end);
    Ok(())
}

/// One element split into its start tag, its immediate children and text as
/// written, and its end tag.
type ElementParts = (Vec<u8>, Vec<Vec<u8>>, Vec<u8>);

/// Splits one element into its parts. An empty element gets a start tag and
/// an end tag, so children can be added.
fn split_element(xml: &[u8]) -> Option<ElementParts> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let (start, name, is_empty) = loop {
        let before = reader.buffer_position() as usize;
        match reader.read_event_into(&mut buffer).ok()? {
            Event::Start(start) => {
                break (
                    xml[before..reader.buffer_position() as usize].to_vec(),
                    start.name().as_ref().to_vec(),
                    false,
                );
            }
            Event::Empty(start) => {
                let mut writer = Writer::new(Vec::new());
                writer.write_event(Event::Start(start.to_owned())).ok()?;
                break (writer.into_inner(), start.name().as_ref().to_vec(), true);
            }
            Event::Eof => return None,
            _ => {}
        }
        buffer.clear();
    };
    let end = [b"</".as_slice(), &name, b">"].concat();
    if is_empty {
        return Some((start, Vec::new(), end));
    }
    let mut children = Vec::new();
    loop {
        buffer.clear();
        let before = reader.buffer_position() as usize;
        match reader.read_event_into(&mut buffer).ok()? {
            Event::Start(child) => {
                let child_name = child.name().as_ref().to_vec();
                reader
                    .read_to_end_into(quick_xml::name::QName(&child_name), &mut Vec::new())
                    .ok()?;
            }
            Event::End(_) => return Some((start, children, end)),
            Event::Eof => return None,
            _ => {}
        }
        children.push(xml[before..reader.buffer_position() as usize].to_vec());
    }
}

fn join_element(start: &[u8], children: &[Vec<u8>], end: &[u8]) -> Vec<u8> {
    let mut xml = start.to_vec();
    for child in children {
        xml.extend_from_slice(child);
    }
    xml.extend_from_slice(end);
    xml
}

/// The local name of a raw child that is an element, or `None` for text,
/// comments, and other nodes.
fn element_local_name(xml: &[u8]) -> Option<Vec<u8>> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    match reader.read_event_into(&mut buffer).ok()? {
        Event::Start(start) | Event::Empty(start) => {
            Some(oxml_core::xml::local_name(start.name().as_ref()).to_vec())
        }
        _ => None,
    }
}

fn element_attribute(xml: &[u8], name: &str) -> Option<String> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    match reader.read_event_into(&mut buffer).ok()? {
        Event::Start(start) | Event::Empty(start) => all_attributes(&start)
            .ok()?
            .into_iter()
            .find_map(|(key, value)| (key == name).then_some(value)),
        _ => None,
    }
}

pub(crate) fn root_attributes(
    start: &BytesStart<'_>,
    fixed_prefixes: &[&str],
) -> Result<Vec<(String, String)>, OxmlError> {
    Ok(all_attributes(start)?
        .into_iter()
        .filter(|(name, _)| !is_fixed_xmlns(name, fixed_prefixes))
        .collect())
}

pub(crate) fn self_contained_attributes(
    start: &BytesStart<'_>,
    fixed_prefixes: &[&str],
    inherited: &[(String, String)],
) -> Result<Vec<(String, String)>, OxmlError> {
    let mut attributes = root_attributes(start, fixed_prefixes)?;
    for (prefix, uri) in inherited {
        let name = if prefix.is_empty() {
            "xmlns".to_owned()
        } else {
            format!("xmlns:{prefix}")
        };
        if is_fixed_xmlns(&name, fixed_prefixes)
            || attributes.iter().any(|(existing, _)| existing == &name)
        {
            continue;
        }
        attributes.push((name, uri.clone()));
    }
    Ok(attributes)
}

pub(crate) fn is_fixed_xmlns(name: &str, fixed_prefixes: &[&str]) -> bool {
    name.strip_prefix("xmlns:")
        .is_some_and(|prefix| fixed_prefixes.contains(&prefix))
}

fn canonical_uri(prefix: &str) -> Option<&'static str> {
    match prefix {
        "p" => Some(P_NS),
        "a" => Some(A_NS),
        "r" => Some(R_NS),
        "mc" => Some(MC_NS),
        _ => None,
    }
}

fn qname_prefix(name: &[u8]) -> Option<&str> {
    let position = name.iter().position(|byte| *byte == b':')?;
    std::str::from_utf8(&name[..position]).ok()
}

fn sort_namespace_entries(entries: &mut [(String, String)]) {
    entries.sort_unstable_by(|left, right| left.0.cmp(&right.0));
}

#[cfg(test)]
mod tests {
    use super::{non_visual_drawing_id, set_non_visual_drawing_name, sort_namespace_entries};

    #[test]
    fn namespace_entries_have_deterministic_prefix_order() {
        let mut entries = vec![
            ("z".to_owned(), "urn:z".to_owned()),
            ("a".to_owned(), "urn:a".to_owned()),
            ("m".to_owned(), "urn:m".to_owned()),
        ];
        sort_namespace_entries(&mut entries);

        assert_eq!(
            entries,
            vec![
                ("a".to_owned(), "urn:a".to_owned()),
                ("m".to_owned(), "urn:m".to_owned()),
                ("z".to_owned(), "urn:z".to_owned()),
            ]
        );
    }

    #[test]
    fn non_visual_id_accepts_any_prefix_and_only_an_unqualified_id() {
        assert_eq!(
            non_visual_drawing_id(br#"<q:cNvPr xmlns:q="urn:p" id="42" q:id="7"/>"#),
            Some(42)
        );
        assert_eq!(
            non_visual_drawing_id(br#"<q:cNvPr xmlns:q="urn:p" q:id="7"/>"#),
            None
        );
        assert_eq!(
            non_visual_drawing_id(
                br#"<q:cNvPr xmlns:q="urn:p"><x:cNvPr xmlns:x="urn:extension" id="99"/></q:cNvPr>"#
            ),
            None
        );
    }

    #[test]
    fn non_visual_name_rewrite_escapes_the_value_and_preserves_children() {
        let mut xml = br#"<q:cNvPr xmlns:q="urn:p" id="42" name="old" producer="one&#x20;two"><x:raw xmlns:x="urn:x">one &amp; two</x:raw><!--note--></q:cNvPr>"#.to_vec();

        set_non_visual_drawing_name(&mut xml, "A & B \"quoted\"").unwrap();

        assert_eq!(
            xml,
            br#"<q:cNvPr xmlns:q="urn:p" id="42" name="A &amp; B &quot;quoted&quot;" producer="one&#x20;two"><x:raw xmlns:x="urn:x">one &amp; two</x:raw><!--note--></q:cNvPr>"#
        );
    }
}
