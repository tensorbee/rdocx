//! Header and footer elements: `CT_HdrFtr`.

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};

use crate::error::Result;
use crate::namespace::{W_NS, matches_local_name};
use crate::raw_xml::{capture_element, capture_empty_element};
use crate::text::CT_P;

/// `CT_HdrFtr` — Content of a header or footer part.
///
/// Contains paragraphs (and potentially tables, same as a document body).
#[derive(Debug, Clone, PartialEq)]
pub struct CT_HdrFtr {
    pub paragraphs: Vec<CT_P>,
    /// Extra namespace declarations captured from the root element.
    pub extra_namespaces: Vec<(String, String)>,
    /// Unknown child elements captured as raw XML.
    pub extra_xml: Vec<Vec<u8>>,
}

#[allow(non_snake_case)]
impl CT_HdrFtr {
    pub fn new() -> Self {
        CT_HdrFtr {
            paragraphs: Vec::new(),
            extra_namespaces: Vec::new(),
            extra_xml: Vec::new(),
        }
    }

    /// Get the combined text of all paragraphs.
    pub fn text(&self) -> String {
        self.paragraphs
            .iter()
            .map(|p| p.text())
            .collect::<Vec<_>>()
            .join("\n")
    }

    /// Parse from XML bytes (the content of header*.xml or footer*.xml).
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        let mut reader = Reader::from_reader(xml);
        reader.config_mut().trim_text(true);

        let mut paragraphs = Vec::new();
        let mut extra_namespaces = Vec::new();
        let mut extra_xml = Vec::new();
        let mut buf = Vec::new();

        let known_ns: &[&[u8]] = &[b"xmlns:w", b"xmlns:r", b"xmlns"];

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let name = e.name();
                    if matches_local_name(name.as_ref(), b"p") {
                        paragraphs.push(CT_P::from_xml(&mut reader)?);
                    } else if matches_local_name(name.as_ref(), b"hdr")
                        || matches_local_name(name.as_ref(), b"ftr")
                    {
                        // Capture extra namespace declarations from root element
                        for attr in e.attributes().flatten() {
                            let key = attr.key.as_ref();
                            if (key.starts_with(b"xmlns:") || key == b"xmlns")
                                && !known_ns.contains(&key)
                            {
                                let key_str = std::str::from_utf8(key)
                                    .unwrap_or("")
                                    .to_string();
                                let val_str = std::str::from_utf8(&attr.value)
                                    .unwrap_or("")
                                    .to_string();
                                extra_namespaces.push((key_str, val_str));
                            }
                        }
                    } else {
                        // Capture unknown elements as raw XML
                        extra_xml.push(capture_element(&mut reader, e)?);
                    }
                }
                Ok(Event::Empty(ref e)) => {
                    let name = e.name();
                    if !matches_local_name(name.as_ref(), b"hdr")
                        && !matches_local_name(name.as_ref(), b"ftr")
                    {
                        extra_xml.push(capture_empty_element(e)?);
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(CT_HdrFtr {
            paragraphs,
            extra_namespaces,
            extra_xml,
        })
    }

    /// Serialize to XML bytes as a header.
    pub fn to_xml_header(&self) -> Result<Vec<u8>> {
        self.to_xml_root("w:hdr")
    }

    /// Serialize to XML bytes as a footer.
    pub fn to_xml_footer(&self) -> Result<Vec<u8>> {
        self.to_xml_root("w:ftr")
    }

    fn to_xml_root(&self, root_tag: &str) -> Result<Vec<u8>> {
        let mut writer = Writer::new_with_indent(Vec::new(), b' ', 2);

        writer.write_event(Event::Decl(BytesDecl::new(
            "1.0",
            Some("UTF-8"),
            Some("yes"),
        )))?;

        let mut start = BytesStart::new(root_tag);
        start.push_attribute(("xmlns:w", W_NS));
        start.push_attribute((
            "xmlns:r",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        ));

        // Always emit xmlns:wp for drawing elements
        let wp_ns = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        let mut has_wp = false;
        for (key, _) in &self.extra_namespaces {
            if key == "xmlns:wp" {
                has_wp = true;
                break;
            }
        }
        if !has_wp {
            start.push_attribute(("xmlns:wp", wp_ns));
        }

        // Replay captured extra namespaces
        for (key, val) in &self.extra_namespaces {
            start.push_attribute((key.as_str(), val.as_str()));
        }

        writer.write_event(Event::Start(start))?;

        for p in &self.paragraphs {
            p.to_xml(&mut writer)?;
        }

        // Write captured unknown elements
        for raw in &self.extra_xml {
            writer.get_mut().extend_from_slice(raw);
        }

        writer.write_event(Event::End(BytesEnd::new(root_tag)))?;

        Ok(writer.into_inner())
    }
}

impl Default for CT_HdrFtr {
    fn default() -> Self {
        Self::new()
    }
}

/// Header/footer reference type in section properties.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HdrFtrType {
    /// Default header/footer
    Default,
    /// First page header/footer
    First,
    /// Even page header/footer
    Even,
}

impl HdrFtrType {
    pub fn from_str(s: &str) -> Self {
        match s {
            "first" => Self::First,
            "even" => Self::Even,
            _ => Self::Default,
        }
    }

    pub fn to_str(self) -> &'static str {
        match self {
            Self::Default => "default",
            Self::First => "first",
            Self::Even => "even",
        }
    }
}

/// A header or footer reference (stored in section properties).
#[derive(Debug, Clone, PartialEq)]
pub struct HdrFtrRef {
    /// The type (default, first, even)
    pub hdr_ftr_type: HdrFtrType,
    /// Relationship ID
    pub rel_id: String,
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn round_trip_header() {
        let mut hdr = CT_HdrFtr::new();
        let mut p = CT_P::new();
        p.add_run("Page Header");
        hdr.paragraphs.push(p);

        let xml = hdr.to_xml_header().unwrap();
        let parsed = CT_HdrFtr::from_xml(&xml).unwrap();
        assert_eq!(parsed.paragraphs.len(), 1);
        assert_eq!(parsed.text(), "Page Header");
    }

    #[test]
    fn round_trip_footer() {
        let mut ftr = CT_HdrFtr::new();
        let mut p = CT_P::new();
        p.add_run("Page Footer");
        ftr.paragraphs.push(p);

        let xml = ftr.to_xml_footer().unwrap();
        let parsed = CT_HdrFtr::from_xml(&xml).unwrap();
        assert_eq!(parsed.text(), "Page Footer");
    }

    #[test]
    fn empty_header() {
        let hdr = CT_HdrFtr::new();
        let xml = hdr.to_xml_header().unwrap();
        let parsed = CT_HdrFtr::from_xml(&xml).unwrap();
        assert_eq!(parsed.paragraphs.len(), 0);
    }
}
