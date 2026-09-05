//! Parsing and writing of `[Content_Types].xml`.

use std::collections::HashMap;

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};

use crate::error::{OpcError, Result};

pub const RELATIONSHIPS: &str = "application/vnd.openxmlformats-package.relationships+xml";
pub const XML: &str = "application/xml";

pub const CORE_PROPERTIES: &str = "application/vnd.openxmlformats-package.core-properties+xml";
pub const EXTENDED_PROPERTIES: &str =
    "application/vnd.openxmlformats-officedocument.extended-properties+xml";
pub const CUSTOM_PROPERTIES: &str =
    "application/vnd.openxmlformats-officedocument.custom-properties+xml";
pub const THEME: &str = "application/vnd.openxmlformats-officedocument.theme+xml";
pub const CHART: &str = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
pub const WORD_GLOSSARY: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml";
pub const WORD_DOCUMENT: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
pub const WORD_DOCUMENT_MACRO_ENABLED: &str =
    "application/vnd.ms-word.document.macroEnabled.main+xml";
pub const WORD_TEMPLATE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml";
pub const WORD_TEMPLATE_MACRO_ENABLED: &str =
    "application/vnd.ms-word.template.macroEnabledTemplate.main+xml";

pub const PRESENTATION: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";
pub const PRESENTATION_MACRO_ENABLED: &str =
    "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml";
pub const PRESENTATION_TEMPLATE: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml";
pub const PRESENTATION_TEMPLATE_MACRO_ENABLED: &str =
    "application/vnd.ms-powerpoint.template.macroEnabled.main+xml";
pub const SLIDESHOW: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml";
pub const SLIDESHOW_MACRO_ENABLED: &str =
    "application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml";
pub const SLIDE: &str = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
pub const SLIDE_LAYOUT: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml";
pub const SLIDE_MASTER: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml";
pub const NOTES_SLIDE: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml";
pub const NOTES_MASTER: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml";
pub const PRES_PROPS: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.presProps+xml";
pub const VIEW_PROPS: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml";
pub const TABLE_STYLES: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml";
pub const HANDOUT_MASTER: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml";
pub const POWERPOINT_COMMENTS: &str = "application/vnd.ms-powerpoint.comments+xml";
pub const POWERPOINT_AUTHORS: &str = "application/vnd.ms-powerpoint.authors+xml";

pub const WORKBOOK: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
pub const EMBEDDED_WORKBOOK: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
pub const WORKSHEET: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
pub const SHARED_STRINGS: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
pub const STYLES: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";

/// A single content type entry, either a Default by extension or an Override by part name.
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
#[derive(Debug, Clone, PartialEq, Eq)]
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

    /// Create the minimal content types shared by every OPC package.
    pub fn minimal() -> Self {
        let mut defaults = HashMap::new();
        defaults.insert("rels".to_string(), RELATIONSHIPS.to_string());
        defaults.insert("xml".to_string(), XML.to_string());

        ContentTypes {
            defaults,
            overrides: HashMap::new(),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn docx_content_types() -> ContentTypes {
        let mut content_types = ContentTypes::minimal();
        content_types.add_override(
            "/word/document.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
        );
        content_types.add_override(
            "/word/styles.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
        );
        content_types
    }

    #[test]
    fn minimal_content_types_contain_only_universal_defaults() {
        let content_types = ContentTypes::minimal();

        assert_eq!(content_types.defaults.len(), 2);
        assert_eq!(
            content_types.defaults.get("rels").map(String::as_str),
            Some("application/vnd.openxmlformats-package.relationships+xml")
        );
        assert_eq!(
            content_types.defaults.get("xml").map(String::as_str),
            Some("application/xml")
        );
        assert!(content_types.overrides.is_empty());
    }

    #[test]
    fn round_trip_content_types() {
        let ct = docx_content_types();
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
        let ct = docx_content_types();
        assert_eq!(
            ct.content_type_for("/word/_rels/document.xml.rels"),
            Some("application/vnd.openxmlformats-package.relationships+xml")
        );
    }
}
