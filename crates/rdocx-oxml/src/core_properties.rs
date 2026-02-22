//! Dublin Core metadata from `docProps/core.xml`.

use quick_xml::Reader;
use quick_xml::events::Event;

use crate::error::Result;

/// Document metadata from `docProps/core.xml` (Dublin Core).
#[derive(Debug, Clone, Default, PartialEq)]
pub struct CoreProperties {
    /// Document title (`dc:title`).
    pub title: Option<String>,
    /// Document creator/author (`dc:creator`).
    pub creator: Option<String>,
    /// Subject (`dc:subject`).
    pub subject: Option<String>,
    /// Description/comments (`dc:description`).
    pub description: Option<String>,
    /// Keywords (`cp:keywords`).
    pub keywords: Option<String>,
    /// Last modified by (`cp:lastModifiedBy`).
    pub last_modified_by: Option<String>,
    /// Date created (`dcterms:created`).
    pub created: Option<String>,
    /// Date modified (`dcterms:modified`).
    pub modified: Option<String>,
}

impl CoreProperties {
    /// Parse `docProps/core.xml` content.
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        let mut reader = Reader::from_reader(xml);
        reader.config_mut().trim_text(true);

        let mut props = CoreProperties::default();
        let mut buf = Vec::new();
        let mut current_tag: Option<String> = None;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let name = e.name();
                    let local = local_name(name.as_ref());
                    match local {
                        "title" | "creator" | "subject" | "description" | "keywords"
                        | "lastModifiedBy" | "created" | "modified" => {
                            current_tag = Some(local.to_string());
                        }
                        _ => {
                            current_tag = None;
                        }
                    }
                }
                Ok(Event::Text(ref e)) => {
                    if let Some(ref tag) = current_tag {
                        let text = e.unescape().unwrap_or_default().to_string();
                        if !text.is_empty() {
                            match tag.as_str() {
                                "title" => props.title = Some(text),
                                "creator" => props.creator = Some(text),
                                "subject" => props.subject = Some(text),
                                "description" => props.description = Some(text),
                                "keywords" => props.keywords = Some(text),
                                "lastModifiedBy" => props.last_modified_by = Some(text),
                                "created" => props.created = Some(text),
                                "modified" => props.modified = Some(text),
                                _ => {}
                            }
                        }
                    }
                }
                Ok(Event::End(_)) => {
                    current_tag = None;
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(props)
    }

    /// Serialize to `docProps/core.xml` bytes.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        use quick_xml::Writer;
        use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText};

        let mut writer = Writer::new_with_indent(Vec::new(), b' ', 2);

        writer.write_event(Event::Decl(BytesDecl::new(
            "1.0",
            Some("UTF-8"),
            Some("yes"),
        )))?;

        let mut root = BytesStart::new("cp:coreProperties");
        root.push_attribute((
            "xmlns:cp",
            "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
        ));
        root.push_attribute(("xmlns:dc", "http://purl.org/dc/elements/1.1/"));
        root.push_attribute(("xmlns:dcterms", "http://purl.org/dc/terms/"));
        root.push_attribute(("xmlns:dcmitype", "http://purl.org/dc/dcmitype/"));
        root.push_attribute(("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"));
        writer.write_event(Event::Start(root))?;

        fn write_element<W: std::io::Write>(
            writer: &mut Writer<W>,
            tag: &str,
            value: &Option<String>,
        ) -> Result<()> {
            if let Some(val) = value {
                writer.write_event(Event::Start(BytesStart::new(tag)))?;
                writer.write_event(Event::Text(BytesText::new(val)))?;
                writer.write_event(Event::End(BytesEnd::new(tag)))?;
            }
            Ok(())
        }

        fn write_date_element<W: std::io::Write>(
            writer: &mut Writer<W>,
            tag: &str,
            value: &Option<String>,
        ) -> Result<()> {
            if let Some(val) = value {
                let mut e = BytesStart::new(tag);
                e.push_attribute(("xsi:type", "dcterms:W3CDTF"));
                writer.write_event(Event::Start(e))?;
                writer.write_event(Event::Text(BytesText::new(val)))?;
                writer.write_event(Event::End(BytesEnd::new(tag)))?;
            }
            Ok(())
        }

        write_element(&mut writer, "dc:title", &self.title)?;
        write_element(&mut writer, "dc:subject", &self.subject)?;
        write_element(&mut writer, "dc:creator", &self.creator)?;
        write_element(&mut writer, "cp:keywords", &self.keywords)?;
        write_element(&mut writer, "dc:description", &self.description)?;
        write_element(&mut writer, "cp:lastModifiedBy", &self.last_modified_by)?;
        write_date_element(&mut writer, "dcterms:created", &self.created)?;
        write_date_element(&mut writer, "dcterms:modified", &self.modified)?;

        writer.write_event(Event::End(BytesEnd::new("cp:coreProperties")))?;

        Ok(writer.into_inner())
    }
}

/// Extract the local name (after the last `:`) from a qualified XML name.
fn local_name(name: &[u8]) -> &str {
    let s = std::str::from_utf8(name).unwrap_or("");
    if let Some(pos) = s.rfind(':') {
        &s[pos + 1..]
    } else {
        s
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_core_properties() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/"
                   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Test Document</dc:title>
  <dc:creator>John Doe</dc:creator>
  <dc:subject>Testing</dc:subject>
  <dc:description>A test document</dc:description>
  <cp:keywords>test, document</cp:keywords>
  <cp:lastModifiedBy>Jane Doe</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">2024-01-15T10:30:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2024-06-20T14:00:00Z</dcterms:modified>
</cp:coreProperties>"#;

        let props = CoreProperties::from_xml(xml).unwrap();
        assert_eq!(props.title, Some("Test Document".to_string()));
        assert_eq!(props.creator, Some("John Doe".to_string()));
        assert_eq!(props.subject, Some("Testing".to_string()));
        assert_eq!(props.description, Some("A test document".to_string()));
        assert_eq!(props.keywords, Some("test, document".to_string()));
        assert_eq!(props.last_modified_by, Some("Jane Doe".to_string()));
        assert_eq!(props.created, Some("2024-01-15T10:30:00Z".to_string()));
        assert_eq!(props.modified, Some("2024-06-20T14:00:00Z".to_string()));
    }

    #[test]
    fn round_trip_core_properties() {
        let props = CoreProperties {
            title: Some("My Title".to_string()),
            creator: Some("Author".to_string()),
            subject: None,
            description: None,
            keywords: Some("rust, docx".to_string()),
            last_modified_by: None,
            created: Some("2024-01-01T00:00:00Z".to_string()),
            modified: Some("2024-06-01T00:00:00Z".to_string()),
        };

        let xml = props.to_xml().unwrap();
        let parsed = CoreProperties::from_xml(&xml).unwrap();

        assert_eq!(parsed.title, props.title);
        assert_eq!(parsed.creator, props.creator);
        assert_eq!(parsed.keywords, props.keywords);
        assert_eq!(parsed.created, props.created);
        assert_eq!(parsed.modified, props.modified);
    }
}
