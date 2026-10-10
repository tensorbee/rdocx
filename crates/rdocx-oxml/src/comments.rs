//! Word comment elements: `CT_Comments` and `CT_Comment`.

use std::io::Write;

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer, XmlVersion};

use crate::error::{OxmlError, Result};
use crate::namespace::{W_NS, matches_local_name};
use crate::numbering::word_prefixes_at;
use crate::properties::is_word_element;
use crate::raw_xml::{capture_element, capture_empty_element};
use crate::text::{CT_P, declare_w14_on_part_root};

const W14_NS: &str = "http://schemas.microsoft.com/office/word/2010/wordml";

/// One entry in a Word comments part.
#[derive(Debug, Clone, PartialEq)]
pub struct CT_Comment {
    pub id: i32,
    pub author: Option<String>,
    pub date: Option<String>,
    pub initials: Option<String>,
    pub paragraphs: Vec<CT_P>,
    /// `w14:paraId` values aligned with `paragraphs`.
    pub paragraph_ids: Vec<Option<String>>,
    /// Unmodelled attributes, including producer namespace declarations.
    pub extra_attributes: Vec<(String, String)>,
    /// Unmodelled children retained at paragraph boundaries.
    pub extra_xml: Vec<(usize, Vec<u8>)>,
}

/// The typed contents of `word/comments.xml`.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct CT_Comments {
    pub comments: Vec<CT_Comment>,
    /// Namespace declarations and compatibility attributes from the root.
    pub root_attributes: Vec<(String, String)>,
    /// Unmodelled root children retained at comment boundaries.
    pub extra_xml: Vec<(usize, Vec<u8>)>,
}

impl CT_Comments {
    pub fn new() -> Self {
        Self::default()
    }

    /// Parse a complete comments part.
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        crate::content_control::validate_story_part_nesting(xml)?;
        let mut reader = Reader::from_reader(xml);
        reader.config_mut().trim_text(true);
        let mut comments = Vec::new();
        let mut root_attributes = Vec::new();
        let mut extra_xml = Vec::new();
        let mut word_prefixes = Vec::new();
        let mut w14_prefixes = Vec::new();
        let mut saw_root = false;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref element)) => {
                    let prefixes = word_prefixes_at(element, &word_prefixes)?;
                    let name = element.name();
                    if is_word_element(name.as_ref(), b"comments", &prefixes) {
                        root_attributes = capture_attributes(element, &[], &prefixes)?;
                        word_prefixes = prefixes;
                        w14_prefixes = namespace_prefixes_at(element, &w14_prefixes, W14_NS)?;
                        saw_root = true;
                    } else if is_word_element(name.as_ref(), b"comment", &prefixes) {
                        comments.push(parse_comment(
                            &mut reader,
                            element,
                            &prefixes,
                            &w14_prefixes,
                        )?);
                    } else if saw_root {
                        extra_xml.push((comments.len(), capture_element(&mut reader, element)?));
                    } else {
                        reader.read_to_end_into(name, &mut Vec::new())?;
                    }
                }
                Ok(Event::Empty(ref element)) => {
                    let prefixes = word_prefixes_at(element, &word_prefixes)?;
                    let name = element.name();
                    if is_word_element(name.as_ref(), b"comments", &prefixes) {
                        root_attributes = capture_attributes(element, &[], &prefixes)?;
                        w14_prefixes = namespace_prefixes_at(element, &w14_prefixes, W14_NS)?;
                        saw_root = true;
                    } else if is_word_element(name.as_ref(), b"comment", &prefixes) {
                        comments.push(parse_empty_comment(element, &prefixes)?);
                    } else if saw_root {
                        extra_xml.push((comments.len(), capture_empty_element(element)?));
                    }
                }
                Ok(Event::Eof) => break,
                Err(error) => return Err(error.into()),
                _ => {}
            }
            buf.clear();
        }

        if !saw_root {
            return Err(OxmlError::MissingElement("comments root".to_owned()));
        }
        Ok(Self {
            comments,
            root_attributes,
            extra_xml,
        })
    }

    /// Serialize with fixed WordprocessingML prefixes and schema child order.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        let mut writer = Writer::new(Vec::new());
        writer.write_event(Event::Decl(BytesDecl::new(
            "1.0",
            Some("UTF-8"),
            Some("yes"),
        )))?;

        // Retained root attributes keep their source order. The fixed `w` and
        // `w14` declarations are added ahead of them only when the source
        // lacks them, so a producer declaration listed in `mc:Ignorable`
        // stays on the root whether or not a paragraph id still uses it.
        let declared = |name: &str| {
            self.root_attributes
                .iter()
                .any(|(candidate, _)| candidate == name)
        };
        let mut root = BytesStart::new("w:comments");
        if !declared("xmlns:w") {
            root.push_attribute(("xmlns:w", W_NS));
        }
        if !declared("xmlns:w14")
            && self
                .comments
                .iter()
                .any(|comment| comment.paragraph_ids.iter().any(Option::is_some))
        {
            root.push_attribute(("xmlns:w14", W14_NS));
        }
        push_preserved_attributes(&mut root, &self.root_attributes, true);
        writer.write_event(Event::Start(root))?;

        write_raw_at(&mut writer, &self.extra_xml, 0)?;
        for (index, comment) in self.comments.iter().enumerate() {
            write_comment(&mut writer, comment)?;
            write_raw_at(&mut writer, &self.extra_xml, index + 1)?;
        }
        writer.write_event(Event::End(BytesEnd::new("w:comments")))?;
        let mut xml = writer.into_inner();
        declare_w14_on_part_root(&mut xml)?;
        Ok(xml)
    }
}

fn parse_comment(
    reader: &mut Reader<&[u8]>,
    start: &BytesStart<'_>,
    prefixes: &[String],
    inherited_w14_prefixes: &[String],
) -> Result<CT_Comment> {
    let (id, author, date, initials, extra_attributes) = comment_attributes(start, prefixes)?;
    let mut paragraphs = Vec::new();
    let mut paragraph_ids = Vec::new();
    let mut extra_xml = Vec::new();
    let w14_prefixes = namespace_prefixes_at(start, inherited_w14_prefixes, W14_NS)?;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) => {
                let child_prefixes = word_prefixes_at(element, prefixes)?;
                if is_word_element(element.name().as_ref(), b"p", &child_prefixes) {
                    let child_w14_prefixes = namespace_prefixes_at(element, &w14_prefixes, W14_NS)?;
                    paragraph_ids.push(namespace_attribute(
                        element,
                        b"paraId",
                        &child_w14_prefixes,
                    )?);
                    paragraphs.push(CT_P::from_xml_with_prefixes_and_root(
                        reader,
                        &child_prefixes,
                        Some(element),
                    )?);
                } else {
                    extra_xml.push((paragraphs.len(), capture_element(reader, element)?));
                }
            }
            Ok(Event::Empty(ref element)) => {
                let child_prefixes = word_prefixes_at(element, prefixes)?;
                if is_word_element(element.name().as_ref(), b"p", &child_prefixes) {
                    let child_w14_prefixes = namespace_prefixes_at(element, &w14_prefixes, W14_NS)?;
                    paragraph_ids.push(namespace_attribute(
                        element,
                        b"paraId",
                        &child_w14_prefixes,
                    )?);
                    paragraphs.push(CT_P::from_empty_root(element, &child_prefixes)?);
                } else {
                    extra_xml.push((paragraphs.len(), capture_empty_element(element)?));
                }
            }
            Ok(Event::End(ref element))
                if matches_local_name(element.name().as_ref(), b"comment") =>
            {
                break;
            }
            Ok(Event::Eof) => break,
            Err(error) => return Err(error.into()),
            _ => {}
        }
        buf.clear();
    }

    Ok(CT_Comment {
        id,
        author,
        date,
        initials,
        paragraphs,
        paragraph_ids,
        extra_attributes,
        extra_xml,
    })
}

fn parse_empty_comment(start: &BytesStart<'_>, prefixes: &[String]) -> Result<CT_Comment> {
    let (id, author, date, initials, extra_attributes) = comment_attributes(start, prefixes)?;
    Ok(CT_Comment {
        id,
        author,
        date,
        initials,
        paragraphs: Vec::new(),
        paragraph_ids: Vec::new(),
        extra_attributes,
        extra_xml: Vec::new(),
    })
}

type CommentAttributes = (
    i32,
    Option<String>,
    Option<String>,
    Option<String>,
    Vec<(String, String)>,
);

fn comment_attributes(start: &BytesStart<'_>, prefixes: &[String]) -> Result<CommentAttributes> {
    let prefixes = word_prefixes_at(start, prefixes)?;
    let id = required_i32_attribute(start, b"id", &prefixes)?;
    let author = word_attribute(start, b"author", &prefixes)?;
    let date = word_attribute(start, b"date", &prefixes)?;
    let initials = word_attribute(start, b"initials", &prefixes)?;
    let extra_attributes =
        capture_attributes(start, &[b"id", b"author", b"date", b"initials"], &prefixes)?;
    Ok((id, author, date, initials, extra_attributes))
}

fn required_i32_attribute(
    start: &BytesStart<'_>,
    local: &[u8],
    prefixes: &[String],
) -> Result<i32> {
    let value = word_attribute(start, local, prefixes)?.ok_or_else(|| {
        OxmlError::MissingElement(format!(
            "comment {} attribute",
            String::from_utf8_lossy(local)
        ))
    })?;
    value.parse().map_err(OxmlError::from)
}

fn word_attribute(
    start: &BytesStart<'_>,
    local: &[u8],
    prefixes: &[String],
) -> Result<Option<String>> {
    for attribute in start.attributes() {
        let attribute = attribute?;
        if is_word_attribute(attribute.key.as_ref(), local, prefixes) {
            return Ok(Some(
                attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, start.decoder())?
                    .into_owned(),
            ));
        }
    }
    Ok(None)
}

fn is_word_attribute(key: &[u8], local: &[u8], prefixes: &[String]) -> bool {
    let Some(separator) = key.iter().position(|byte| *byte == b':') else {
        return false;
    };
    key.get(separator + 1..) == Some(local)
        && prefixes
            .iter()
            .any(|prefix| prefix.as_bytes() == &key[..separator])
}

fn capture_attributes(
    start: &BytesStart<'_>,
    modelled: &[&[u8]],
    prefixes: &[String],
) -> Result<Vec<(String, String)>> {
    let mut attributes = Vec::new();
    for attribute in start.attributes() {
        let attribute = attribute?;
        if modelled
            .iter()
            .any(|local| is_word_attribute(attribute.key.as_ref(), local, prefixes))
        {
            continue;
        }
        let name = std::str::from_utf8(attribute.key.as_ref())?.to_owned();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, start.decoder())?
            .into_owned();
        attributes.push((name, value));
    }
    Ok(attributes)
}

fn write_comment<W: Write>(writer: &mut Writer<W>, comment: &CT_Comment) -> Result<()> {
    let mut start = BytesStart::new("w:comment");
    if let Some(author) = &comment.author {
        start.push_attribute(("w:author", author.as_str()));
    }
    if let Some(date) = &comment.date {
        start.push_attribute(("w:date", date.as_str()));
    }
    if let Some(initials) = &comment.initials {
        start.push_attribute(("w:initials", initials.as_str()));
    }
    let mut id = itoa::Buffer::new();
    start.push_attribute(("w:id", id.format(comment.id)));
    push_preserved_attributes(&mut start, &comment.extra_attributes, false);
    writer.write_event(Event::Start(start))?;

    write_raw_at(writer, &comment.extra_xml, 0)?;
    for (index, paragraph) in comment.paragraphs.iter().enumerate() {
        paragraph.to_xml_with_para_id(
            writer,
            comment.paragraph_ids.get(index).and_then(Option::as_deref),
        )?;
        write_raw_at(writer, &comment.extra_xml, index + 1)?;
    }
    writer.write_event(Event::End(BytesEnd::new("w:comment")))?;
    Ok(())
}

fn push_preserved_attributes(
    start: &mut BytesStart<'_>,
    attributes: &[(String, String)],
    root: bool,
) {
    for (name, value) in attributes {
        // The serializer writes the fixed `w` and `w14` prefixes, so their
        // root declarations bind the namespaces it writes.
        let value = match name.as_str() {
            "xmlns:w" if root => W_NS,
            "xmlns:w14" if root => W14_NS,
            _ => value.as_str(),
        };
        start.push_attribute((name.as_str(), value));
    }
}

fn namespace_prefixes_at(
    start: &BytesStart<'_>,
    inherited: &[String],
    namespace: &str,
) -> Result<Vec<String>> {
    let mut prefixes = inherited.to_vec();
    for attribute in start.attributes() {
        let attribute = attribute?;
        let name = attribute.key.as_ref();
        let prefix = if name == b"xmlns" {
            b"".as_slice()
        } else if let Some(prefix) = name.strip_prefix(b"xmlns:") {
            prefix
        } else {
            continue;
        };
        let prefix = std::str::from_utf8(prefix)?.to_owned();
        prefixes.retain(|candidate| candidate != &prefix);
        let value =
            attribute.decoded_and_normalized_value(XmlVersion::Implicit1_0, start.decoder())?;
        if value.as_ref() == namespace {
            prefixes.push(prefix);
        }
    }
    Ok(prefixes)
}

fn namespace_attribute(
    start: &BytesStart<'_>,
    local: &[u8],
    prefixes: &[String],
) -> Result<Option<String>> {
    for attribute in start.attributes() {
        let attribute = attribute?;
        let key = attribute.key.as_ref();
        let Some(separator) = key.iter().position(|byte| *byte == b':') else {
            continue;
        };
        if key.get(separator + 1..) == Some(local)
            && prefixes
                .iter()
                .any(|prefix| prefix.as_bytes() == &key[..separator])
        {
            return Ok(Some(
                attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, start.decoder())?
                    .into_owned(),
            ));
        }
    }
    Ok(None)
}

fn write_raw_at<W: Write>(
    writer: &mut Writer<W>,
    extra_xml: &[(usize, Vec<u8>)],
    position: usize,
) -> Result<()> {
    for (at, raw) in extra_xml {
        if *at == position {
            writer.get_mut().write_all(raw)?;
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn comments_accept_aliases_and_write_fixed_prefixes_in_schema_order() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:comments xmlns:x="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:ext="urn:producer" ext:root="kept"><ext:before/><x:comment ext:item="kept" x:id="7" x:author="Ada" x:date="2026-08-17T10:00:00Z" x:initials="AL"><ext:beforeParagraph/><x:p><x:r><x:t>Review this</x:t></x:r></x:p><ext:afterParagraph/></x:comment><ext:after/></x:comments>"#;

        let comments = CT_Comments::from_xml(xml).expect("aliased comments should parse");
        assert_eq!(comments.comments.len(), 1);
        assert_eq!(comments.comments[0].id, 7);
        assert_eq!(comments.comments[0].author.as_deref(), Some("Ada"));
        assert_eq!(comments.comments[0].initials.as_deref(), Some("AL"));
        assert_eq!(comments.comments[0].paragraphs[0].text(), "Review this");

        let output = String::from_utf8(comments.to_xml().expect("comments should serialize"))
            .expect("comments XML should be UTF-8");
        assert!(output.contains("<w:comments"));
        assert!(output.contains(
            r#"<w:comment w:author="Ada" w:date="2026-08-17T10:00:00Z" w:initials="AL" w:id="7""#
        ));
        assert!(output.contains("<ext:beforeParagraph/><w:p>"));
        assert!(output.contains("</w:p><ext:afterParagraph/>"));
        assert!(output.contains("<ext:before/><w:comment"));
        assert!(output.contains("</w:comment><ext:after/>"));
        assert!(output.contains(r#"ext:root="kept""#));
        assert!(output.contains(r#"ext:item="kept""#));
    }

    #[test]
    fn a_retained_text_identity_stays_bound_without_a_paragraph_identity() {
        // The root declares `w14` for an authored paragraph identity only, so
        // a paragraph that carries `w14:textId` alone relies on the check of
        // the written part.
        let xml = br#"<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:comment w:id="1"><w:p w14:textId="5E6F7A8B"><w:r><w:t>thread</w:t></w:r></w:p></w:comment></w:comments>"#;
        let comments = CT_Comments::from_xml(xml).unwrap();
        let output = comments.to_xml().unwrap();
        let reparsed = CT_Comments::from_xml(&output)
            .unwrap_or_else(|error| panic!("{error}: {}", String::from_utf8_lossy(&output)));
        assert_eq!(reparsed.to_xml().unwrap(), output);
    }

    #[test]
    fn malformed_comment_id_is_rejected() {
        let xml = br#"<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="not-a-number"><w:p/></w:comment></w:comments>"#;
        assert!(CT_Comments::from_xml(xml).is_err());
    }

    #[test]
    fn paragraph_ids_follow_the_bound_extension_namespace() {
        let xml = br#"<x:comments xmlns:x="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:p14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:ext="urn:producer"><x:comment x:id="1"><x:p p14:paraId="0000000A" ext:paraId="foreign"><x:r><x:t>thread</x:t></x:r></x:p></x:comment></x:comments>"#;
        let comments = CT_Comments::from_xml(xml).unwrap();
        assert_eq!(
            comments.comments[0].paragraph_ids[0].as_deref(),
            Some("0000000A")
        );
        let output = String::from_utf8(comments.to_xml().unwrap()).unwrap();
        assert!(output.contains("w14:paraId=\"0000000A\""));
        assert!(!output.contains("w14:paraId=\"foreign\""));
    }

    /// #160: an empty Word comments root that lists `w14` in `mc:Ignorable`
    /// lost the `w14` declaration once a comment without a paragraph id was
    /// written into it, and the kept declarations moved behind `xmlns:w`.
    #[test]
    fn rewritten_root_keeps_its_declarations_in_source_order() {
        let comment = |para_id: Option<&str>| {
            let mut paragraph = CT_P::new();
            paragraph.add_run("note");
            CT_Comment {
                id: 0,
                author: Some("Ada".to_owned()),
                date: None,
                initials: None,
                paragraphs: vec![paragraph],
                paragraph_ids: vec![para_id.map(str::to_owned)],
                extra_attributes: Vec::new(),
                extra_xml: Vec::new(),
            }
        };
        let word_root = concat!(
            r#"<w:comments xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" "#,
            r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" "#,
            r#"xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" "#,
            r#"xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" "#,
            r#"mc:Ignorable="w14 w15">"#,
        );
        let mut comments = CT_Comments::from_xml(word_root.replace('>', "/>").as_bytes()).unwrap();
        comments.comments.push(comment(None));
        let output = String::from_utf8(comments.to_xml().unwrap()).unwrap();
        assert!(
            output.contains(&format!("{word_root}<w:comment ")),
            "{output}"
        );

        // Declarations the serializer needs and the source lacks come first.
        let aliased_root = concat!(
            r#"<comments xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main" "#,
            r#"xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" "#,
            r#"xmlns:p14="http://schemas.microsoft.com/office/word/2010/wordml" "#,
            r#"mc:Ignorable="p14"/>"#,
        );
        let mut comments = CT_Comments::from_xml(aliased_root.as_bytes()).unwrap();
        comments.comments.push(comment(Some("0000000A")));
        let output = String::from_utf8(comments.to_xml().unwrap()).unwrap();
        let expected = concat!(
            r#"<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" "#,
            r#"xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" "#,
            r#"xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main" "#,
            r#"xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" "#,
            r#"xmlns:p14="http://schemas.microsoft.com/office/word/2010/wordml" "#,
            r#"mc:Ignorable="p14"><w:comment "#,
        );
        assert!(output.contains(expected), "{output}");
    }
}
