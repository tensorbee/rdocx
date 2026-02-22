//! Parsing and writing of `.rels` relationship files.

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};

use crate::error::{OpcError, Result};

/// Well-known OOXML relationship types.
pub mod rel_types {
    pub const DOCUMENT: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
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
    pub const CHART: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
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
            next_id: max_id + 1,
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
        let id = format!("rId{}", self.next_id);
        self.next_id += 1;
        self.items.push(Relationship {
            id: id.clone(),
            rel_type: rel_type.to_string(),
            target: target.to_string(),
            target_mode: None,
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
            self.next_id = num + 1;
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

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
}
