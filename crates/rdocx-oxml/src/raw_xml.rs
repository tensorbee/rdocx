//! Raw XML capture utilities for preserving unknown elements during round-trip.

use quick_xml::events::{BytesStart, Event};
use quick_xml::{Reader, Writer};

use crate::error::Result;

/// Capture a full XML subtree (from after the start tag through the matching end tag)
/// and return it as raw bytes. The returned bytes include the start tag, all children,
/// and the end tag.
pub fn capture_element(reader: &mut Reader<&[u8]>, start: &BytesStart) -> Result<Vec<u8>> {
    let mut writer = Writer::new(Vec::new());

    // Write the start tag
    writer.write_event(Event::Start(start.to_owned()))?;

    // Track nesting depth for the tag name
    let tag_name = start.name().as_ref().to_vec();
    let mut depth = 1u32;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                if e.name().as_ref() == tag_name {
                    depth += 1;
                }
                writer.write_event(Event::Start(e.to_owned()))?;
            }
            Ok(Event::End(ref e)) => {
                if e.name().as_ref() == tag_name {
                    depth -= 1;
                }
                writer.write_event(Event::End(e.to_owned()))?;
                if depth == 0 {
                    break;
                }
            }
            Ok(Event::Empty(ref e)) => {
                writer.write_event(Event::Empty(e.to_owned()))?;
            }
            Ok(Event::Text(ref e)) => {
                writer.write_event(Event::Text(e.to_owned().into_owned()))?;
            }
            Ok(Event::CData(ref e)) => {
                writer.write_event(Event::CData(e.to_owned().into_owned()))?;
            }
            Ok(Event::Comment(ref e)) => {
                writer.write_event(Event::Comment(e.to_owned().into_owned()))?;
            }
            Ok(Event::PI(ref e)) => {
                writer.write_event(Event::PI(e.to_owned().into_owned()))?;
            }
            Ok(Event::Decl(ref e)) => {
                writer.write_event(Event::Decl(e.to_owned().into_owned()))?;
            }
            Ok(Event::DocType(ref e)) => {
                writer.write_event(Event::DocType(e.to_owned().into_owned()))?;
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(e.into()),
        }
        buf.clear();
    }

    Ok(writer.into_inner())
}

/// Capture an empty (self-closing) element as raw bytes.
pub fn capture_empty_element(e: &BytesStart) -> Result<Vec<u8>> {
    let mut writer = Writer::new(Vec::new());
    writer.write_event(Event::Empty(e.to_owned()))?;
    Ok(writer.into_inner())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn capture_simple_element() {
        let xml = r#"<root><child>text</child><sibling/></root>"#;
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();

        // Read past <root>
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.name().as_ref() == b"child" => {
                    let captured = capture_element(&mut reader, e).unwrap();
                    let s = String::from_utf8(captured).unwrap();
                    assert!(s.contains("<child>"));
                    assert!(s.contains("text"));
                    assert!(s.contains("</child>"));
                    return;
                }
                _ => {}
            }
            buf.clear();
        }
    }

    #[test]
    fn capture_empty() {
        let xml = r#"<item attr="val"/>"#;
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(ref e)) => {
                    let captured = capture_empty_element(e).unwrap();
                    let s = String::from_utf8(captured).unwrap();
                    assert!(s.contains("item"));
                    assert!(s.contains("attr"));
                    return;
                }
                _ => {}
            }
            buf.clear();
        }
    }

    #[test]
    fn capture_nested_element() {
        let xml = r#"<outer><inner><deep>data</deep></inner></outer>"#;
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.name().as_ref() == b"outer" => {
                    let captured = capture_element(&mut reader, e).unwrap();
                    let s = String::from_utf8(captured).unwrap();
                    assert!(s.contains("<outer>"));
                    assert!(s.contains("<inner>"));
                    assert!(s.contains("<deep>"));
                    assert!(s.contains("data"));
                    assert!(s.contains("</deep>"));
                    assert!(s.contains("</inner>"));
                    assert!(s.contains("</outer>"));
                    return;
                }
                _ => {}
            }
            buf.clear();
        }
    }
}
