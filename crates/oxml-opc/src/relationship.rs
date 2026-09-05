//! Parsing and writing of `.rels` relationship files.

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};

use crate::error::{OpcError, Result};

/// Well-known OOXML relationship types.
pub mod rel_types {
    // Package-level relationships.
    pub const CORE_PROPERTIES: &str =
        "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
    pub const THUMBNAIL: &str =
        "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail";
    pub const DIGITAL_SIGNATURE_ORIGIN: &str =
        "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin";
    pub const DIGITAL_SIGNATURE: &str =
        "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature";

    // Shared officeDocument relationships.
    pub const DOCUMENT: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
    pub const EXTENDED_PROPERTIES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
    pub const CUSTOM_PROPERTIES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties";
    pub const STYLES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
    pub const NUMBERING: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
    pub const HEADER: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
    pub const FOOTER: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
    pub const IMAGE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
    pub const AUDIO: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio";
    pub const VIDEO: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video";
    pub const POWERPOINT_MEDIA: &str =
        "http://schemas.microsoft.com/office/2007/relationships/media";
    pub const SETTINGS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";
    pub const FONT_TABLE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable";
    pub const THEME: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
    pub const HYPERLINK: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
    pub const FOOTNOTES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";
    pub const ENDNOTES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes";
    pub const COMMENTS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
    pub const GLOSSARY_DOCUMENT: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument";
    pub const CHART: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
    pub const DIAGRAM_DATA: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData";
    pub const DIAGRAM_LAYOUT: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout";
    pub const DIAGRAM_QUICK_STYLE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle";
    pub const DIAGRAM_COLORS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors";
    pub const DIAGRAM_DRAWING: &str =
        "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing";
    pub const PACKAGE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package";
    pub const OLE_OBJECT: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject";
    pub const CONTROL: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/control";
    pub const STRICT_OLE_OBJECT: &str =
        "http://purl.oclc.org/ooxml/officeDocument/relationships/oleObject";
    pub const STRICT_CONTROL: &str =
        "http://purl.oclc.org/ooxml/officeDocument/relationships/control";
    pub const ACTIVEX_CONTROL_BINARY: &str =
        "http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary";
    pub const VBA_PROJECT: &str =
        "http://schemas.microsoft.com/office/2006/relationships/vbaProject";
    pub const VBA_PROJECT_SIGNATURE: &str =
        "http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature";
    pub const VBA_PROJECT_SIGNATURE_AGILE: &str =
        "http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignatureAgile";

    // SpreadsheetML relationships.
    pub const WORKSHEET: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
    pub const SHARED_STRINGS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";

    // PresentationML relationships.
    pub const SLIDE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
    pub const SLIDE_LAYOUT: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
    pub const SLIDE_MASTER: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";
    pub const NOTES_SLIDE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";
    pub const NOTES_MASTER: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster";
    pub const PRES_PROPS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps";
    pub const VIEW_PROPS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps";
    pub const TABLE_STYLES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles";
    pub const HANDOUT_MASTER: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/handoutMaster";
    pub const POWERPOINT_COMMENTS: &str =
        "http://schemas.microsoft.com/office/2018/10/relationships/comments";
    pub const POWERPOINT_AUTHORS: &str =
        "http://schemas.microsoft.com/office/2018/10/relationships/authors";
}

/// A single relationship entry.
#[derive(Debug, Clone, PartialEq)]
pub struct Relationship {
    pub id: String,
    pub rel_type: String,
    pub target: String,
    pub target_mode: Option<String>,
}

/// A collection of relationships parsed from a `.rels` file.
#[derive(Debug, Clone, Default)]
pub struct Relationships {
    pub items: Vec<Relationship>,
    next_id: u32,
}

impl Relationships {
    pub fn new() -> Self {
        Self {
            items: Vec::new(),
            next_id: 1,
        }
    }

    /// Parse from XML bytes.
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        let mut reader = Reader::from_reader(xml);
        reader.config_mut().trim_text(true);

        let mut items = Vec::new();
        let mut max_id: u32 = 0;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(ref e)) if e.name().as_ref() == b"Relationship" => {
                    let mut id = None;
                    let mut rel_type = None;
                    let mut target = None;
                    let mut target_mode = None;

                    for attr in e.attributes() {
                        let attr = attr?;
                        match attr.key.as_ref() {
                            b"Id" => {
                                let val = std::str::from_utf8(&attr.value)?.to_string();
                                // Extract numeric suffix for next_id tracking
                                if let Some(num_str) = val.strip_prefix("rId")
                                    && let Ok(n) = num_str.parse::<u32>()
                                {
                                    max_id = max_id.max(n);
                                }
                                id = Some(val);
                            }
                            b"Type" => {
                                rel_type = Some(std::str::from_utf8(&attr.value)?.to_string());
                            }
                            b"Target" => {
                                target = Some(std::str::from_utf8(&attr.value)?.to_string());
                            }
                            b"TargetMode" => {
                                target_mode = Some(std::str::from_utf8(&attr.value)?.to_string());
                            }
                            _ => {}
                        }
                    }

                    match (id, rel_type, target) {
                        (Some(id), Some(rel_type), Some(target)) => {
                            items.push(Relationship {
                                id,
                                rel_type,
                                target,
                                target_mode,
                            });
                        }
                        _ => return Err(OpcError::InvalidRelationship),
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(Relationships {
            items,
            next_id: next_relationship_number(max_id),
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

        let mut rels_start = BytesStart::new("Relationships");
        rels_start.push_attribute((
            "xmlns",
            "http://schemas.openxmlformats.org/package/2006/relationships",
        ));
        writer.write_event(Event::Start(rels_start))?;

        for rel in &self.items {
            let mut elem = BytesStart::new("Relationship");
            elem.push_attribute(("Id", rel.id.as_str()));
            elem.push_attribute(("Type", rel.rel_type.as_str()));
            elem.push_attribute(("Target", rel.target.as_str()));
            if let Some(ref mode) = rel.target_mode {
                elem.push_attribute(("TargetMode", mode.as_str()));
            }
            writer.write_event(Event::Empty(elem))?;
        }

        writer.write_event(Event::End(BytesEnd::new("Relationships")))?;

        Ok(writer.into_inner())
    }

    /// Find a relationship by its ID.
    pub fn get_by_id(&self, id: &str) -> Option<&Relationship> {
        self.items.iter().find(|r| r.id == id)
    }

    /// Find the first relationship matching a given type.
    pub fn get_by_type(&self, rel_type: &str) -> Option<&Relationship> {
        self.items.iter().find(|r| r.rel_type == rel_type)
    }

    /// Find all relationships matching a given type.
    pub fn get_all_by_type(&self, rel_type: &str) -> Vec<&Relationship> {
        self.items
            .iter()
            .filter(|r| r.rel_type == rel_type)
            .collect()
    }

    /// Add a new relationship and return its generated ID.
    pub fn add(&mut self, rel_type: &str, target: &str) -> String {
        let id = self.allocate_id();
        self.items.push(Relationship {
            id: id.clone(),
            rel_type: rel_type.to_string(),
            target: target.to_string(),
            target_mode: None,
        });
        id
    }

    /// Add an externally-targeted relationship (e.g. a hyperlink URL) and
    /// return its generated ID.
    pub fn add_external(&mut self, rel_type: &str, target: &str) -> String {
        let id = self.allocate_id();
        self.items.push(Relationship {
            id: id.clone(),
            rel_type: rel_type.to_string(),
            target: target.to_string(),
            target_mode: Some("External".to_string()),
        });
        id
    }

    /// Add a relationship with a specific ID.
    ///
    /// If a relationship with this ID already exists, it is replaced.
    /// The `next_id` counter is updated to avoid future collisions.
    pub fn add_with_id(&mut self, id: &str, rel_type: &str, target: &str) {
        self.items.retain(|r| r.id != id);
        self.items.push(Relationship {
            id: id.to_string(),
            rel_type: rel_type.to_string(),
            target: target.to_string(),
            target_mode: None,
        });
        if let Some(num) = id.strip_prefix("rId").and_then(|s| s.parse::<u32>().ok())
            && num >= self.next_id
        {
            self.next_id = next_relationship_number(num);
        }
    }

    fn allocate_id(&mut self) -> String {
        let mut candidate = self.next_id;
        let numeric_space = usize::try_from(u32::MAX).unwrap_or(usize::MAX);
        let attempts = self.items.len().saturating_add(1).min(numeric_space);
        for _ in 0..attempts {
            let id = format!("rId{candidate}");
            let next = next_relationship_number(candidate);
            if self.items.iter().all(|relationship| relationship.id != id) {
                self.next_id = next;
                return id;
            }
            candidate = next;
        }

        for ordinal in 1u128.. {
            let id = format!("rIdGenerated{ordinal}");
            if self.items.iter().all(|relationship| relationship.id != id) {
                return id;
            }
        }
        unreachable!("the generated relationship id space is unbounded")
    }
}

fn next_relationship_number(current: u32) -> u32 {
    current.checked_add(1).unwrap_or(1).max(1)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn relationship_and_content_type_constants_are_unique_and_well_formed() {
        const PACKAGE_PREFIX: &str =
            "http://schemas.openxmlformats.org/package/2006/relationships/";
        const OFFICE_PREFIX: &str =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";

        let package_relationships = [rel_types::CORE_PROPERTIES, rel_types::THUMBNAIL];
        let office_relationships = [
            rel_types::DOCUMENT,
            rel_types::STYLES,
            rel_types::NUMBERING,
            rel_types::HEADER,
            rel_types::FOOTER,
            rel_types::IMAGE,
            rel_types::SETTINGS,
            rel_types::FONT_TABLE,
            rel_types::THEME,
            rel_types::HYPERLINK,
            rel_types::FOOTNOTES,
            rel_types::ENDNOTES,
            rel_types::GLOSSARY_DOCUMENT,
            rel_types::CHART,
            rel_types::DIAGRAM_DATA,
            rel_types::DIAGRAM_LAYOUT,
            rel_types::DIAGRAM_QUICK_STYLE,
            rel_types::DIAGRAM_COLORS,
            rel_types::PACKAGE,
            rel_types::OLE_OBJECT,
            rel_types::CONTROL,
            rel_types::WORKSHEET,
            rel_types::SHARED_STRINGS,
            rel_types::EXTENDED_PROPERTIES,
            rel_types::CUSTOM_PROPERTIES,
            rel_types::SLIDE,
            rel_types::SLIDE_LAYOUT,
            rel_types::SLIDE_MASTER,
            rel_types::NOTES_SLIDE,
            rel_types::NOTES_MASTER,
            rel_types::PRES_PROPS,
            rel_types::VIEW_PROPS,
            rel_types::TABLE_STYLES,
            rel_types::HANDOUT_MASTER,
        ];
        let strict_relationships = [rel_types::STRICT_OLE_OBJECT, rel_types::STRICT_CONTROL];
        let microsoft_relationships = [
            rel_types::ACTIVEX_CONTROL_BINARY,
            rel_types::VBA_PROJECT,
            rel_types::VBA_PROJECT_SIGNATURE,
            rel_types::VBA_PROJECT_SIGNATURE_AGILE,
        ];

        let mut relationship_values = std::collections::HashSet::new();
        for value in package_relationships {
            assert!(value.starts_with(PACKAGE_PREFIX));
            assert!(!value.chars().any(char::is_whitespace));
            assert!(relationship_values.insert(value));
        }
        for value in office_relationships {
            assert!(value.starts_with(OFFICE_PREFIX));
            assert!(!value.chars().any(char::is_whitespace));
            assert!(relationship_values.insert(value));
        }
        for value in strict_relationships {
            assert!(value.starts_with("http://purl.oclc.org/ooxml/officeDocument/relationships/"));
            assert!(!value.chars().any(char::is_whitespace));
            assert!(relationship_values.insert(value));
        }
        for value in microsoft_relationships {
            assert!(value.starts_with("http://schemas.microsoft.com/office/2006/relationships/"));
            assert!(!value.chars().any(char::is_whitespace));
            assert!(relationship_values.insert(value));
        }
        for value in [
            rel_types::POWERPOINT_COMMENTS,
            rel_types::POWERPOINT_AUTHORS,
            rel_types::DIAGRAM_DRAWING,
        ] {
            assert!(value.starts_with("http://schemas.microsoft.com/office/"));
            assert!(!value.chars().any(char::is_whitespace));
            assert!(relationship_values.insert(value));
        }

        let content_type_values = [
            crate::content_types::RELATIONSHIPS,
            crate::content_types::XML,
            crate::content_types::CORE_PROPERTIES,
            crate::content_types::EXTENDED_PROPERTIES,
            crate::content_types::CUSTOM_PROPERTIES,
            crate::content_types::THEME,
            crate::content_types::CHART,
            crate::content_types::WORD_GLOSSARY,
            crate::content_types::WORD_DOCUMENT,
            crate::content_types::WORD_DOCUMENT_MACRO_ENABLED,
            crate::content_types::WORD_TEMPLATE,
            crate::content_types::WORD_TEMPLATE_MACRO_ENABLED,
            crate::content_types::PRESENTATION,
            crate::content_types::SLIDESHOW,
            crate::content_types::SLIDE,
            crate::content_types::SLIDE_LAYOUT,
            crate::content_types::SLIDE_MASTER,
            crate::content_types::NOTES_SLIDE,
            crate::content_types::NOTES_MASTER,
            crate::content_types::PRES_PROPS,
            crate::content_types::VIEW_PROPS,
            crate::content_types::TABLE_STYLES,
            crate::content_types::HANDOUT_MASTER,
            crate::content_types::POWERPOINT_COMMENTS,
            crate::content_types::POWERPOINT_AUTHORS,
            crate::content_types::WORKBOOK,
            crate::content_types::EMBEDDED_WORKBOOK,
            crate::content_types::WORKSHEET,
            crate::content_types::SHARED_STRINGS,
            crate::content_types::STYLES,
        ];

        let mut content_types = std::collections::HashSet::new();
        for value in content_type_values {
            let (kind, subtype) = value.split_once('/').expect("valid MIME type");
            assert_eq!(kind, "application");
            assert!(!subtype.is_empty());
            assert!(!subtype.contains('/'));
            assert!(!value.chars().any(char::is_whitespace));
            assert!(content_types.insert(value));
        }
    }

    #[test]
    fn round_trip_relationships() {
        let mut rels = Relationships::new();
        rels.add(rel_types::DOCUMENT, "word/document.xml");
        rels.add(rel_types::STYLES, "word/styles.xml");

        let xml = rels.to_xml().unwrap();
        let parsed = Relationships::from_xml(&xml).unwrap();

        assert_eq!(parsed.items.len(), 2);
        assert_eq!(parsed.items[0].id, "rId1");
        assert_eq!(parsed.items[0].target, "word/document.xml");
        assert_eq!(parsed.items[1].id, "rId2");
    }

    #[test]
    fn find_by_type() {
        let mut rels = Relationships::new();
        rels.add(rel_types::DOCUMENT, "word/document.xml");
        rels.add(rel_types::STYLES, "word/styles.xml");

        let doc = rels.get_by_type(rel_types::DOCUMENT).unwrap();
        assert_eq!(doc.target, "word/document.xml");
    }

    #[test]
    fn high_numeric_relationship_ids_roll_over_without_collision() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId4294967294" Type="type-a" Target="a.xml"/>
</Relationships>"#;
        let mut relationships = Relationships::from_xml(xml).unwrap();
        assert_eq!(relationships.add("type-b", "b.xml"), "rId4294967295");
        assert_eq!(relationships.add("type-c", "c.xml"), "rId1");
        assert_eq!(
            relationships.add_external("type-d", "https://example.com"),
            "rId2"
        );
        let ids = relationships
            .items
            .iter()
            .map(|relationship| relationship.id.as_str())
            .collect::<std::collections::HashSet<_>>();
        assert_eq!(ids.len(), relationships.items.len());
        assert!(!ids.contains("rId0"));
    }

    #[test]
    fn parsed_u32_max_relationship_id_rolls_to_first_free_positive_id() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId4294967295" Type="type-max" Target="max.xml"/>
  <Relationship Id="rId1" Type="type-one" Target="one.xml"/>
</Relationships>"#;
        let mut relationships = Relationships::from_xml(xml).unwrap();
        assert_eq!(relationships.add("type-two", "two.xml"), "rId2");
        relationships.add_with_id("rId4294967295", "type-max", "replacement.xml");
        assert_eq!(relationships.add("type-three", "three.xml"), "rId3");
    }
}
