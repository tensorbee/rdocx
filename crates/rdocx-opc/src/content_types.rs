//! Parsing and writing of `[Content_Types].xml`.

use std::collections::HashMap;

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};

use crate::error::{OpcError, Result};

/// A single content type entry — either a Default (by extension) or an Override (by part name).
#[derive(Debug, Clone, PartialEq)]
pub enum ContentType {
    Default {
        extension: String,
        content_type: String,
    },
    Override {
        part_name: String,
        content_type: String,
    },
}

/// Parsed `[Content_Types].xml`.
#[derive(Debug, Clone)]
pub struct ContentTypes {
    pub defaults: HashMap<String, String>,
    pub overrides: HashMap<String, String>,
}

impl ContentTypes {
    /// Parse from XML bytes.
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        let mut reader = Reader::from_reader(xml);
        reader.config_mut().trim_text(true);

        let mut defaults = HashMap::new();
        let mut overrides = HashMap::new();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(ref e)) => match e.name().as_ref() {
                    b"Default" => {
                        let mut ext = None;
                        let mut ct = None;
                        for attr in e.attributes() {
                            let attr = attr?;
                            match attr.key.as_ref() {
                                b"Extension" => {
                                    ext = Some(std::str::from_utf8(&attr.value)?.to_string());
                                }
                                b"ContentType" => {
                                    ct = Some(std::str::from_utf8(&attr.value)?.to_string());
                                }
                                _ => {}
                            }
                        }
                        match (ext, ct) {
                            (Some(e), Some(c)) => {
                                defaults.insert(e, c);
                            }
                            _ => return Err(OpcError::InvalidContentTypes),
                        }
                    }
                    b"Override" => {
                        let mut pn = None;
                        let mut ct = None;
                        for attr in e.attributes() {
                            let attr = attr?;
                            match attr.key.as_ref() {
                                b"PartName" => {
                                    pn = Some(std::str::from_utf8(&attr.value)?.to_string());
                                }
                                b"ContentType" => {
                                    ct = Some(std::str::from_utf8(&attr.value)?.to_string());
                                }
                                _ => {}
                            }
                        }
                        match (pn, ct) {
                            (Some(p), Some(c)) => {
                                overrides.insert(p, c);
                            }
                            _ => return Err(OpcError::InvalidContentTypes),
                        }
                    }
                    _ => {}
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(ContentTypes {
            defaults,
            overrides,
        })
    }

    /// Serialize to XML bytes.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        let mut writer = Writer::new_with_indent(Vec::new(), b' ', 2);

        writer.write_event(Event::Decl(BytesDecl::new(
            "1.0",
            Some("UTF-8"),
            Some("yes"),
        )))?;

        let mut types_start = BytesStart::new("Types");
        types_start.push_attribute((
            "xmlns",
            "http://schemas.openxmlformats.org/package/2006/content-types",
        ));
        writer.write_event(Event::Start(types_start))?;

        // Write defaults sorted for deterministic output
        let mut sorted_defaults: Vec<_> = self.defaults.iter().collect();
        sorted_defaults.sort_by_key(|(k, _)| (*k).clone());
        for (ext, ct) in sorted_defaults {
            let mut elem = BytesStart::new("Default");
            elem.push_attribute(("Extension", ext.as_str()));
            elem.push_attribute(("ContentType", ct.as_str()));
            writer.write_event(Event::Empty(elem))?;
        }

        // Write overrides sorted for deterministic output
        let mut sorted_overrides: Vec<_> = self.overrides.iter().collect();
        sorted_overrides.sort_by_key(|(k, _)| (*k).clone());
        for (pn, ct) in sorted_overrides {
            let mut elem = BytesStart::new("Override");
            elem.push_attribute(("PartName", pn.as_str()));
            elem.push_attribute(("ContentType", ct.as_str()));
            writer.write_event(Event::Empty(elem))?;
        }

        writer.write_event(Event::End(BytesEnd::new("Types")))?;

        Ok(writer.into_inner())
    }

    /// Look up the content type for a given part name.
    pub fn content_type_for(&self, part_name: &str) -> Option<&str> {
        // Check overrides first
        if let Some(ct) = self.overrides.get(part_name) {
            return Some(ct.as_str());
        }
        // Fall back to defaults by extension
        if let Some(dot_pos) = part_name.rfind('.') {
            let ext = &part_name[dot_pos + 1..];
            if let Some(ct) = self.defaults.get(ext) {
                return Some(ct.as_str());
            }
        }
        None
    }

    /// Add a default content type for an extension (e.g., "png" -> "image/png").
    pub fn add_default(&mut self, extension: &str, content_type: &str) {
        self.defaults
            .entry(extension.to_string())
            .or_insert_with(|| content_type.to_string());
    }

    /// Add an override content type for a specific part name.
    pub fn add_override(&mut self, part_name: &str, content_type: &str) {
        self.overrides
            .insert(part_name.to_string(), content_type.to_string());
    }

    /// Create a new ContentTypes with the standard DOCX defaults.
    pub fn new_docx() -> Self {
        let mut defaults = HashMap::new();
        defaults.insert(
            "rels".to_string(),
            "application/vnd.openxmlformats-package.relationships+xml".to_string(),
        );
        defaults.insert("xml".to_string(), "application/xml".to_string());

        let mut overrides = HashMap::new();
        overrides.insert(
            "/word/document.xml".to_string(),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
                .to_string(),
        );
        overrides.insert(
            "/word/styles.xml".to_string(),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml".to_string(),
        );

        ContentTypes {
            defaults,
            overrides,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn round_trip_content_types() {
        let ct = ContentTypes::new_docx();
        let xml = ct.to_xml().unwrap();
        let parsed = ContentTypes::from_xml(&xml).unwrap();
        assert_eq!(parsed.defaults.len(), ct.defaults.len());
        assert_eq!(parsed.overrides.len(), ct.overrides.len());
        assert_eq!(
            parsed.content_type_for("/word/document.xml"),
            Some(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
            )
        );
    }

    #[test]
    fn lookup_by_extension() {
        let ct = ContentTypes::new_docx();
        assert_eq!(
            ct.content_type_for("/word/_rels/document.xml.rels"),
            Some("application/vnd.openxmlformats-package.relationships+xml")
        );
    }
}
