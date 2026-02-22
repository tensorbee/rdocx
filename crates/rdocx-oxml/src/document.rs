//! Document-level elements: `CT_Document` and `CT_Body`.

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};

use crate::error::Result;
use crate::header_footer::{HdrFtrRef, HdrFtrType};
use crate::namespace::{W_NS, matches_local_name};
use crate::properties::get_val_attr;
use crate::raw_xml::{capture_element, capture_empty_element};
use crate::shared::{ST_PageOrientation, ST_SectionType};
use crate::table::CT_Tbl;
use crate::text::CT_P;
use crate::units::Twips;

/// Content that can appear in a document body (paragraphs and tables).
#[derive(Debug, Clone, PartialEq)]
pub enum BodyContent {
    Paragraph(CT_P),
    Table(CT_Tbl),
    /// Raw XML for unknown elements (bookmarks, SDTs, mc:AlternateContent, etc.)
    RawXml(Vec<u8>),
}

/// Column definition for multi-column layouts.
#[derive(Debug, Clone, PartialEq)]
pub struct CT_Column {
    /// Column width in twips
    pub width: Option<Twips>,
    /// Space after this column in twips
    pub space: Option<Twips>,
}

/// `CT_Columns` — Column layout configuration.
#[derive(Debug, Clone, PartialEq)]
pub struct CT_Columns {
    /// Number of columns (if equal width)
    pub num: Option<u32>,
    /// Space between columns in twips (when equal width)
    pub space: Option<Twips>,
    /// Whether columns are equal width
    pub equal_width: Option<bool>,
    /// Separator line between columns
    pub sep: Option<bool>,
    /// Individual column definitions (for unequal widths)
    pub columns: Vec<CT_Column>,
}

impl Default for CT_Columns {
    fn default() -> Self {
        CT_Columns {
            num: Some(1),
            space: Some(Twips(720)),
            equal_width: Some(true),
            sep: None,
            columns: Vec::new(),
        }
    }
}

/// `CT_SectPr` — Section properties (page size, margins, columns, orientation).
#[derive(Debug, Clone, PartialEq)]
#[allow(non_snake_case)]
pub struct CT_SectPr {
    /// Page width in twips
    pub page_width: Option<Twips>,
    /// Page height in twips
    pub page_height: Option<Twips>,
    /// Page orientation
    pub orientation: Option<ST_PageOrientation>,
    /// Top margin in twips
    pub margin_top: Option<Twips>,
    /// Right margin in twips
    pub margin_right: Option<Twips>,
    /// Bottom margin in twips
    pub margin_bottom: Option<Twips>,
    /// Left margin in twips
    pub margin_left: Option<Twips>,
    /// Gutter margin in twips
    pub gutter: Option<Twips>,
    /// Header distance from top edge in twips
    pub header_distance: Option<Twips>,
    /// Footer distance from bottom edge in twips
    pub footer_distance: Option<Twips>,
    /// Section break type
    pub section_type: Option<ST_SectionType>,
    /// Column layout
    pub columns: Option<CT_Columns>,
    /// Title page (different first page header/footer)
    pub title_pg: Option<bool>,
    /// Header references
    pub header_refs: Vec<HdrFtrRef>,
    /// Footer references
    pub footer_refs: Vec<HdrFtrRef>,
    /// Unknown child elements captured as raw XML.
    pub extra_xml: Vec<Vec<u8>>,
}

#[allow(non_snake_case)]
impl CT_SectPr {
    /// Default US Letter page with 1-inch margins.
    pub fn default_letter() -> Self {
        CT_SectPr {
            page_width: Some(Twips(12240)),   // 8.5"
            page_height: Some(Twips(15840)),  // 11"
            orientation: Some(ST_PageOrientation::Portrait),
            margin_top: Some(Twips(1440)),    // 1"
            margin_right: Some(Twips(1440)),  // 1"
            margin_bottom: Some(Twips(1440)), // 1"
            margin_left: Some(Twips(1440)),   // 1"
            gutter: Some(Twips(0)),
            header_distance: Some(Twips(720)),
            footer_distance: Some(Twips(720)),
            section_type: None,
            columns: None,
            title_pg: None,
            header_refs: Vec::new(),
            footer_refs: Vec::new(),
            extra_xml: Vec::new(),
        }
    }

    /// Default A4 page with 1-inch margins.
    pub fn default_a4() -> Self {
        CT_SectPr {
            page_width: Some(Twips(11906)),   // 210mm
            page_height: Some(Twips(16838)),  // 297mm
            orientation: Some(ST_PageOrientation::Portrait),
            margin_top: Some(Twips(1440)),
            margin_right: Some(Twips(1440)),
            margin_bottom: Some(Twips(1440)),
            margin_left: Some(Twips(1440)),
            gutter: Some(Twips(0)),
            header_distance: Some(Twips(720)),
            footer_distance: Some(Twips(720)),
            section_type: None,
            columns: None,
            title_pg: None,
            header_refs: Vec::new(),
            footer_refs: Vec::new(),
            extra_xml: Vec::new(),
        }
    }

    pub fn from_xml(reader: &mut Reader<&[u8]>) -> Result<Self> {
        let mut sect = CT_SectPr {
            page_width: None,
            page_height: None,
            orientation: None,
            margin_top: None,
            margin_right: None,
            margin_bottom: None,
            margin_left: None,
            gutter: None,
            header_distance: None,
            footer_distance: None,
            section_type: None,
            columns: None,
            title_pg: None,
            header_refs: Vec::new(),
            footer_refs: Vec::new(),
            extra_xml: Vec::new(),
        };
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(ref e)) => {
                    let name = e.name();
                    if matches_local_name(name.as_ref(), b"pgSz") {
                        for attr in e.attributes() {
                            let attr = attr?;
                            let key = attr.key.as_ref();
                            let val_str = std::str::from_utf8(&attr.value)?;
                            if matches_local_name(key, b"w") {
                                sect.page_width = Some(Twips(val_str.parse()?));
                            } else if matches_local_name(key, b"h") {
                                sect.page_height = Some(Twips(val_str.parse()?));
                            } else if matches_local_name(key, b"orient") {
                                sect.orientation =
                                    Some(ST_PageOrientation::from_str(val_str)?);
                            }
                        }
                    } else if matches_local_name(name.as_ref(), b"pgMar") {
                        for attr in e.attributes() {
                            let attr = attr?;
                            let key = attr.key.as_ref();
                            let val: i32 = std::str::from_utf8(&attr.value)?.parse()?;
                            if matches_local_name(key, b"top") {
                                sect.margin_top = Some(Twips(val));
                            } else if matches_local_name(key, b"right")
                                || matches_local_name(key, b"end")
                            {
                                sect.margin_right = Some(Twips(val));
                            } else if matches_local_name(key, b"bottom") {
                                sect.margin_bottom = Some(Twips(val));
                            } else if matches_local_name(key, b"left")
                                || matches_local_name(key, b"start")
                            {
                                sect.margin_left = Some(Twips(val));
                            } else if matches_local_name(key, b"gutter") {
                                sect.gutter = Some(Twips(val));
                            } else if matches_local_name(key, b"header") {
                                sect.header_distance = Some(Twips(val));
                            } else if matches_local_name(key, b"footer") {
                                sect.footer_distance = Some(Twips(val));
                            }
                        }
                    } else if matches_local_name(name.as_ref(), b"type") {
                        if let Some(val) = get_val_attr(e)? {
                            sect.section_type = Some(ST_SectionType::from_str(&val)?);
                        }
                    } else if matches_local_name(name.as_ref(), b"cols") {
                        sect.columns = Some(Self::parse_cols_empty(e)?);
                    } else if matches_local_name(name.as_ref(), b"headerReference") {
                        let mut hdr_type = HdrFtrType::Default;
                        let mut rel_id = String::new();
                        for attr in e.attributes() {
                            let attr = attr?;
                            let key = attr.key.as_ref();
                            let val = std::str::from_utf8(&attr.value)?;
                            if matches_local_name(key, b"type") {
                                hdr_type = HdrFtrType::from_str(val);
                            } else if matches_local_name(key, b"id") {
                                rel_id = val.to_string();
                            }
                        }
                        if !rel_id.is_empty() {
                            sect.header_refs.push(HdrFtrRef {
                                hdr_ftr_type: hdr_type,
                                rel_id,
                            });
                        }
                    } else if matches_local_name(name.as_ref(), b"footerReference") {
                        let mut ftr_type = HdrFtrType::Default;
                        let mut rel_id = String::new();
                        for attr in e.attributes() {
                            let attr = attr?;
                            let key = attr.key.as_ref();
                            let val = std::str::from_utf8(&attr.value)?;
                            if matches_local_name(key, b"type") {
                                ftr_type = HdrFtrType::from_str(val);
                            } else if matches_local_name(key, b"id") {
                                rel_id = val.to_string();
                            }
                        }
                        if !rel_id.is_empty() {
                            sect.footer_refs.push(HdrFtrRef {
                                hdr_ftr_type: ftr_type,
                                rel_id,
                            });
                        }
                    } else if matches_local_name(name.as_ref(), b"titlePg") {
                        sect.title_pg = Some(true);
                    } else {
                        // Capture unknown empty elements
                        sect.extra_xml.push(capture_empty_element(e)?);
                    }
                }
                Ok(Event::Start(ref e)) => {
                    let name = e.name();
                    if matches_local_name(name.as_ref(), b"cols") {
                        sect.columns = Some(Self::parse_cols_start(reader, e)?);
                    } else {
                        // Capture unknown start elements as raw XML
                        sect.extra_xml.push(capture_element(reader, e)?);
                    }
                }
                Ok(Event::End(ref e))
                    if matches_local_name(e.name().as_ref(), b"sectPr") =>
                {
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(sect)
    }

    fn parse_cols_attrs(e: &BytesStart) -> Result<CT_Columns> {
        let mut cols = CT_Columns::default();
        for attr in e.attributes() {
            let attr = attr?;
            let key = attr.key.as_ref();
            let val_str = std::str::from_utf8(&attr.value)?;
            if matches_local_name(key, b"num") {
                cols.num = Some(val_str.parse()?);
            } else if matches_local_name(key, b"space") {
                cols.space = Some(Twips(val_str.parse()?));
            } else if matches_local_name(key, b"equalWidth") {
                cols.equal_width = Some(val_str == "1" || val_str == "true");
            } else if matches_local_name(key, b"sep") {
                cols.sep = Some(val_str == "1" || val_str == "true");
            }
        }
        Ok(cols)
    }

    fn parse_cols_empty(e: &BytesStart) -> Result<CT_Columns> {
        Self::parse_cols_attrs(e)
    }

    fn parse_cols_start(reader: &mut Reader<&[u8]>, e: &BytesStart) -> Result<CT_Columns> {
        let mut cols = Self::parse_cols_attrs(e)?;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(ref e))
                    if matches_local_name(e.name().as_ref(), b"col") =>
                {
                    let mut width = None;
                    let mut space = None;
                    for attr in e.attributes() {
                        let attr = attr?;
                        let key = attr.key.as_ref();
                        let val: i32 = std::str::from_utf8(&attr.value)?.parse()?;
                        if matches_local_name(key, b"w") {
                            width = Some(Twips(val));
                        } else if matches_local_name(key, b"space") {
                            space = Some(Twips(val));
                        }
                    }
                    cols.columns.push(CT_Column { width, space });
                }
                Ok(Event::End(ref e))
                    if matches_local_name(e.name().as_ref(), b"cols") =>
                {
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(cols)
    }

    pub fn to_xml<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut buf = itoa::Buffer::new();
        writer.write_event(Event::Start(BytesStart::new("w:sectPr")))?;

        // headerReference elements
        for hdr in &self.header_refs {
            let mut e = BytesStart::new("w:headerReference");
            e.push_attribute(("w:type", hdr.hdr_ftr_type.to_str()));
            e.push_attribute(("r:id", hdr.rel_id.as_str()));
            writer.write_event(Event::Empty(e))?;
        }

        // footerReference elements
        for ftr in &self.footer_refs {
            let mut e = BytesStart::new("w:footerReference");
            e.push_attribute(("w:type", ftr.hdr_ftr_type.to_str()));
            e.push_attribute(("r:id", ftr.rel_id.as_str()));
            writer.write_event(Event::Empty(e))?;
        }

        // type (section break type)
        if let Some(st) = self.section_type {
            let mut e = BytesStart::new("w:type");
            e.push_attribute(("w:val", st.to_str()));
            writer.write_event(Event::Empty(e))?;
        }

        // pgSz
        if self.page_width.is_some() || self.page_height.is_some() || self.orientation.is_some() {
            let mut e = BytesStart::new("w:pgSz");
            if let Some(w) = self.page_width {
                e.push_attribute(("w:w", buf.format(w.0)));
            }
            if let Some(h) = self.page_height {
                e.push_attribute(("w:h", buf.format(h.0)));
            }
            if let Some(orient) = self.orientation {
                if orient == ST_PageOrientation::Landscape {
                    e.push_attribute(("w:orient", orient.to_str()));
                }
            }
            writer.write_event(Event::Empty(e))?;
        }

        // pgMar
        if self.margin_top.is_some()
            || self.margin_right.is_some()
            || self.margin_bottom.is_some()
            || self.margin_left.is_some()
        {
            let mut e = BytesStart::new("w:pgMar");
            if let Some(t) = self.margin_top {
                e.push_attribute(("w:top", buf.format(t.0)));
            }
            if let Some(r) = self.margin_right {
                e.push_attribute(("w:right", buf.format(r.0)));
            }
            if let Some(b) = self.margin_bottom {
                e.push_attribute(("w:bottom", buf.format(b.0)));
            }
            if let Some(l) = self.margin_left {
                e.push_attribute(("w:left", buf.format(l.0)));
            }
            if let Some(g) = self.gutter {
                e.push_attribute(("w:gutter", buf.format(g.0)));
            }
            if let Some(h) = self.header_distance {
                e.push_attribute(("w:header", buf.format(h.0)));
            }
            if let Some(f) = self.footer_distance {
                e.push_attribute(("w:footer", buf.format(f.0)));
            }
            writer.write_event(Event::Empty(e))?;
        }

        // cols
        if let Some(ref cols) = self.columns {
            if cols.columns.is_empty() {
                // Simple equal-width columns
                let mut e = BytesStart::new("w:cols");
                if let Some(num) = cols.num {
                    e.push_attribute(("w:num", buf.format(num)));
                }
                if let Some(space) = cols.space {
                    e.push_attribute(("w:space", buf.format(space.0)));
                }
                if let Some(eq) = cols.equal_width {
                    if !eq {
                        e.push_attribute(("w:equalWidth", "0"));
                    }
                }
                if let Some(sep) = cols.sep {
                    if sep {
                        e.push_attribute(("w:sep", "1"));
                    }
                }
                writer.write_event(Event::Empty(e))?;
            } else {
                // Individual column definitions
                let mut e = BytesStart::new("w:cols");
                if let Some(num) = cols.num {
                    e.push_attribute(("w:num", buf.format(num)));
                }
                if let Some(eq) = cols.equal_width {
                    e.push_attribute(("w:equalWidth", if eq { "1" } else { "0" }));
                }
                if let Some(sep) = cols.sep {
                    if sep {
                        e.push_attribute(("w:sep", "1"));
                    }
                }
                writer.write_event(Event::Start(e))?;

                for col in &cols.columns {
                    let mut ce = BytesStart::new("w:col");
                    if let Some(w) = col.width {
                        ce.push_attribute(("w:w", buf.format(w.0)));
                    }
                    if let Some(s) = col.space {
                        ce.push_attribute(("w:space", buf.format(s.0)));
                    }
                    writer.write_event(Event::Empty(ce))?;
                }

                writer.write_event(Event::End(BytesEnd::new("w:cols")))?;
            }
        }

        // titlePg
        if let Some(true) = self.title_pg {
            writer.write_event(Event::Empty(BytesStart::new("w:titlePg")))?;
        }

        // Write captured unknown elements
        for raw in &self.extra_xml {
            writer.get_mut().write_all(raw)?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:sectPr")))?;
        Ok(())
    }
}

/// `CT_Body` — The document body containing paragraphs, tables, and section properties.
#[derive(Debug, Clone, PartialEq)]
#[allow(non_snake_case)]
pub struct CT_Body {
    /// Mixed content: paragraphs and tables in document order.
    pub content: Vec<BodyContent>,
    pub sect_pr: Option<CT_SectPr>,
}

#[allow(non_snake_case)]
impl CT_Body {
    pub fn new() -> Self {
        CT_Body {
            content: Vec::new(),
            sect_pr: Some(CT_SectPr::default_letter()),
        }
    }

    /// Get an iterator over only the paragraphs.
    pub fn paragraphs(&self) -> impl Iterator<Item = &CT_P> {
        self.content.iter().filter_map(|c| match c {
            BodyContent::Paragraph(p) => Some(p),
            _ => None,
        })
    }

    /// Get a mutable iterator over only the paragraphs.
    pub fn paragraphs_mut(&mut self) -> impl Iterator<Item = &mut CT_P> {
        self.content.iter_mut().filter_map(|c| match c {
            BodyContent::Paragraph(p) => Some(p),
            _ => None,
        })
    }

    /// Get an iterator over only the tables.
    pub fn tables(&self) -> impl Iterator<Item = &CT_Tbl> {
        self.content.iter().filter_map(|c| match c {
            BodyContent::Table(t) => Some(t),
            _ => None,
        })
    }

    /// Get a mutable iterator over only the tables.
    pub fn tables_mut(&mut self) -> impl Iterator<Item = &mut CT_Tbl> {
        self.content.iter_mut().filter_map(|c| match c {
            BodyContent::Table(t) => Some(t),
            _ => None,
        })
    }

    /// Add a paragraph to the body.
    pub fn add_paragraph(&mut self, p: CT_P) {
        self.content.push(BodyContent::Paragraph(p));
    }

    /// Add a table to the body.
    pub fn add_table(&mut self, tbl: CT_Tbl) {
        self.content.push(BodyContent::Table(tbl));
    }

    /// Get the number of body content elements (paragraphs + tables).
    pub fn content_count(&self) -> usize {
        self.content.len()
    }

    /// Insert a paragraph at the given index.
    ///
    /// Panics if `index > content_count()`.
    pub fn insert_paragraph(&mut self, index: usize, p: CT_P) {
        self.content.insert(index, BodyContent::Paragraph(p));
    }

    /// Insert a table at the given index.
    ///
    /// Panics if `index > content_count()`.
    pub fn insert_table(&mut self, index: usize, tbl: CT_Tbl) {
        self.content.insert(index, BodyContent::Table(tbl));
    }

    /// Find the index of the first paragraph whose text contains the given substring.
    pub fn find_paragraph_index(&self, text: &str) -> Option<usize> {
        self.content.iter().position(|c| match c {
            BodyContent::Paragraph(p) => p.text().contains(text),
            _ => false,
        })
    }

    /// Remove and return the content at the given index, or `None` if out of bounds.
    pub fn remove(&mut self, index: usize) -> Option<BodyContent> {
        if index < self.content.len() {
            Some(self.content.remove(index))
        } else {
            None
        }
    }

    /// Get a reference to the content at the given index.
    pub fn get(&self, index: usize) -> Option<&BodyContent> {
        self.content.get(index)
    }

    /// Get a mutable reference to the content at the given index.
    pub fn get_mut(&mut self, index: usize) -> Option<&mut BodyContent> {
        self.content.get_mut(index)
    }

    pub fn from_xml(reader: &mut Reader<&[u8]>) -> Result<Self> {
        let mut content = Vec::new();
        let mut sect_pr = None;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let name = e.name();
                    if matches_local_name(name.as_ref(), b"p") {
                        content.push(BodyContent::Paragraph(CT_P::from_xml(reader)?));
                    } else if matches_local_name(name.as_ref(), b"tbl") {
                        content.push(BodyContent::Table(CT_Tbl::from_xml(reader)?));
                    } else if matches_local_name(name.as_ref(), b"sectPr") {
                        sect_pr = Some(CT_SectPr::from_xml(reader)?);
                    } else {
                        // Capture unknown elements as raw XML
                        content.push(BodyContent::RawXml(capture_element(reader, e)?));
                    }
                }
                Ok(Event::Empty(ref e)) => {
                    let name = e.name();
                    if !matches_local_name(name.as_ref(), b"body") {
                        content.push(BodyContent::RawXml(capture_empty_element(e)?));
                    }
                }
                Ok(Event::End(ref e)) if matches_local_name(e.name().as_ref(), b"body") => {
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(CT_Body { content, sect_pr })
    }

    pub fn to_xml<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_event(Event::Start(BytesStart::new("w:body")))?;

        for item in &self.content {
            match item {
                BodyContent::Paragraph(p) => p.to_xml(writer)?,
                BodyContent::Table(t) => t.to_xml(writer)?,
                BodyContent::RawXml(raw) => {
                    writer.get_mut().write_all(raw)?;
                }
            }
        }

        if let Some(ref sect) = self.sect_pr {
            sect.to_xml(writer)?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:body")))?;
        Ok(())
    }
}

impl Default for CT_Body {
    fn default() -> Self {
        Self::new()
    }
}

/// `CT_Document` — The root document element.
#[derive(Debug, Clone, PartialEq)]
#[allow(non_snake_case)]
pub struct CT_Document {
    pub body: CT_Body,
    /// Extra namespace declarations captured from the original document element.
    /// Each entry is (prefix, uri), e.g. ("xmlns:wp14", "http://...").
    pub extra_namespaces: Vec<(String, String)>,
    /// Raw XML for `<w:background>` element if present.
    pub background_xml: Option<Vec<u8>>,
}

#[allow(non_snake_case)]
impl CT_Document {
    pub fn new() -> Self {
        CT_Document {
            body: CT_Body::new(),
            extra_namespaces: Vec::new(),
            background_xml: None,
        }
    }

    /// Parse from XML bytes (the content of word/document.xml).
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        let mut reader = Reader::from_reader(xml);
        reader.config_mut().trim_text(true);

        let mut body = None;
        let mut extra_namespaces = Vec::new();
        let mut background_xml = None;
        let mut buf = Vec::new();

        // Known namespace prefixes that we always emit ourselves
        let known_ns: &[&[u8]] = &[b"xmlns:w", b"xmlns:r", b"xmlns:mc", b"xmlns"];

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let name = e.name();
                    if matches_local_name(name.as_ref(), b"body") {
                        body = Some(CT_Body::from_xml(&mut reader)?);
                    } else if matches_local_name(name.as_ref(), b"document") {
                        // Capture extra namespace declarations from the document element
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
                        // Continue into document element
                    } else if matches_local_name(name.as_ref(), b"background") {
                        background_xml = Some(capture_element(&mut reader, e)?);
                    } else {
                        reader.read_to_end_into(name, &mut Vec::new())?;
                    }
                }
                Ok(Event::Empty(ref e)) => {
                    if matches_local_name(e.name().as_ref(), b"background") {
                        background_xml = Some(capture_empty_element(e)?);
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(CT_Document {
            body: body.unwrap_or_default(),
            extra_namespaces,
            background_xml,
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

        let mut doc_start = BytesStart::new("w:document");
        doc_start.push_attribute(("xmlns:w", W_NS));
        doc_start.push_attribute((
            "xmlns:r",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        ));
        doc_start.push_attribute((
            "xmlns:mc",
            "http://schemas.openxmlformats.org/markup-compatibility/2006",
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
            doc_start.push_attribute(("xmlns:wp", wp_ns));
        }

        // Replay captured extra namespaces
        for (key, val) in &self.extra_namespaces {
            doc_start.push_attribute((key.as_str(), val.as_str()));
        }

        writer.write_event(Event::Start(doc_start))?;

        // Write background element if present
        if let Some(ref bg) = self.background_xml {
            writer.get_mut().extend_from_slice(bg);
        }

        self.body.to_xml(&mut writer)?;

        writer.write_event(Event::End(BytesEnd::new("w:document")))?;

        Ok(writer.into_inner())
    }
}

impl Default for CT_Document {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn round_trip_document() {
        let mut doc = CT_Document::new();
        let mut p = CT_P::new();
        p.add_run("Hello World");
        doc.body.add_paragraph(p);

        let xml = doc.to_xml().unwrap();
        let parsed = CT_Document::from_xml(&xml).unwrap();

        let paras: Vec<_> = parsed.body.paragraphs().collect();
        assert_eq!(paras.len(), 1);
        assert_eq!(paras[0].text(), "Hello World");
    }

    #[test]
    fn round_trip_with_section() {
        let doc = CT_Document::new();
        let xml = doc.to_xml().unwrap();
        let parsed = CT_Document::from_xml(&xml).unwrap();
        assert!(parsed.body.sect_pr.is_some());
        let sect = parsed.body.sect_pr.unwrap();
        assert_eq!(sect.page_width, Some(Twips(12240)));
    }

    #[test]
    fn round_trip_landscape() {
        let mut doc = CT_Document::new();
        let sect = doc.body.sect_pr.as_mut().unwrap();
        sect.orientation = Some(ST_PageOrientation::Landscape);
        sect.page_width = Some(Twips(15840)); // 11"
        sect.page_height = Some(Twips(12240)); // 8.5"

        let xml = doc.to_xml().unwrap();
        let parsed = CT_Document::from_xml(&xml).unwrap();
        let sect = parsed.body.sect_pr.unwrap();
        assert_eq!(sect.orientation, Some(ST_PageOrientation::Landscape));
        assert_eq!(sect.page_width, Some(Twips(15840)));
    }

    #[test]
    fn round_trip_columns() {
        let mut doc = CT_Document::new();
        let sect = doc.body.sect_pr.as_mut().unwrap();
        sect.columns = Some(CT_Columns {
            num: Some(2),
            space: Some(Twips(720)),
            equal_width: Some(true),
            sep: Some(true),
            columns: Vec::new(),
        });

        let xml = doc.to_xml().unwrap();
        let parsed = CT_Document::from_xml(&xml).unwrap();
        let cols = parsed.body.sect_pr.unwrap().columns.unwrap();
        assert_eq!(cols.num, Some(2));
        assert_eq!(cols.space, Some(Twips(720)));
        assert_eq!(cols.sep, Some(true));
    }

    #[test]
    fn round_trip_section_type() {
        let mut doc = CT_Document::new();
        let sect = doc.body.sect_pr.as_mut().unwrap();
        sect.section_type = Some(ST_SectionType::Continuous);
        sect.title_pg = Some(true);

        let xml = doc.to_xml().unwrap();
        let parsed = CT_Document::from_xml(&xml).unwrap();
        let sect = parsed.body.sect_pr.unwrap();
        assert_eq!(sect.section_type, Some(ST_SectionType::Continuous));
        assert_eq!(sect.title_pg, Some(true));
    }

    #[test]
    fn insert_paragraph_at_beginning() {
        let mut body = CT_Body::new();
        let mut p1 = CT_P::new();
        p1.add_run("First");
        body.add_paragraph(p1);

        let mut p0 = CT_P::new();
        p0.add_run("Inserted");
        body.insert_paragraph(0, p0);

        assert_eq!(body.content_count(), 2);
        match &body.content[0] {
            BodyContent::Paragraph(p) => assert_eq!(p.text(), "Inserted"),
            _ => panic!("expected paragraph"),
        }
        match &body.content[1] {
            BodyContent::Paragraph(p) => assert_eq!(p.text(), "First"),
            _ => panic!("expected paragraph"),
        }
    }

    #[test]
    fn insert_paragraph_in_middle() {
        let mut body = CT_Body::new();
        let mut p1 = CT_P::new();
        p1.add_run("First");
        body.add_paragraph(p1);
        let mut p2 = CT_P::new();
        p2.add_run("Third");
        body.add_paragraph(p2);

        let mut mid = CT_P::new();
        mid.add_run("Middle");
        body.insert_paragraph(1, mid);

        assert_eq!(body.content_count(), 3);
        let texts: Vec<_> = body.paragraphs().map(|p| p.text()).collect();
        assert_eq!(texts, vec!["First", "Middle", "Third"]);
    }

    #[test]
    fn find_paragraph_index_match() {
        let mut body = CT_Body::new();
        let mut p1 = CT_P::new();
        p1.add_run("Hello World");
        body.add_paragraph(p1);
        let mut p2 = CT_P::new();
        p2.add_run("INSERT_HERE");
        body.add_paragraph(p2);

        assert_eq!(body.find_paragraph_index("INSERT_HERE"), Some(1));
        assert_eq!(body.find_paragraph_index("NONEXISTENT"), None);
    }

    #[test]
    fn remove_content() {
        let mut body = CT_Body::new();
        let mut p1 = CT_P::new();
        p1.add_run("First");
        body.add_paragraph(p1);
        let mut p2 = CT_P::new();
        p2.add_run("Second");
        body.add_paragraph(p2);

        let removed = body.remove(0);
        assert!(removed.is_some());
        assert_eq!(body.content_count(), 1);
        match &body.content[0] {
            BodyContent::Paragraph(p) => assert_eq!(p.text(), "Second"),
            _ => panic!("expected paragraph"),
        }

        // Out of bounds
        assert!(body.remove(5).is_none());
    }

    #[test]
    fn get_and_get_mut() {
        let mut body = CT_Body::new();
        let mut p = CT_P::new();
        p.add_run("Test");
        body.add_paragraph(p);

        assert!(body.get(0).is_some());
        assert!(body.get(1).is_none());

        if let Some(BodyContent::Paragraph(p)) = body.get_mut(0) {
            p.add_run(" Modified");
        }
        match body.get(0).unwrap() {
            BodyContent::Paragraph(p) => assert_eq!(p.text(), "Test Modified"),
            _ => panic!("expected paragraph"),
        }
    }

    #[test]
    fn sect_pr_section_type_and_orientation_round_trip() {
        let mut doc = CT_Document::new();
        let sect = doc.body.sect_pr.as_mut().unwrap();
        sect.section_type = Some(ST_SectionType::NextPage);
        sect.orientation = Some(ST_PageOrientation::Landscape);
        sect.page_width = Some(Twips(15840));
        sect.page_height = Some(Twips(12240));

        let xml = doc.to_xml().unwrap();
        let parsed = CT_Document::from_xml(&xml).unwrap();
        let sect2 = parsed.body.sect_pr.unwrap();
        assert_eq!(sect2.section_type, Some(ST_SectionType::NextPage));
        assert_eq!(sect2.orientation, Some(ST_PageOrientation::Landscape));
        assert_eq!(sect2.page_width, Some(Twips(15840)));
        assert_eq!(sect2.page_height, Some(Twips(12240)));
    }

    #[test]
    fn sect_pr_all_section_types() {
        for section_type in [
            ST_SectionType::NextPage,
            ST_SectionType::Continuous,
            ST_SectionType::EvenPage,
            ST_SectionType::OddPage,
        ] {
            let mut doc = CT_Document::new();
            let sect = doc.body.sect_pr.as_mut().unwrap();
            sect.section_type = Some(section_type);

            let xml = doc.to_xml().unwrap();
            let parsed = CT_Document::from_xml(&xml).unwrap();
            let sect2 = parsed.body.sect_pr.unwrap();
            assert_eq!(sect2.section_type, Some(section_type),
                "section type round-trip failed for {section_type:?}");
        }
    }

    #[test]
    fn sect_pr_in_paragraph_ppr_round_trip() {
        // Section breaks inside paragraph properties (pPr/sectPr)
        let mut doc = CT_Document::new();
        let mut p = CT_P::new();
        p.add_run("Section break paragraph");
        let mut ppr = crate::properties::CT_PPr::default();
        let mut sect = CT_SectPr::default_letter();
        sect.section_type = Some(ST_SectionType::NextPage);
        sect.orientation = Some(ST_PageOrientation::Landscape);
        sect.page_width = Some(Twips(15840));
        sect.page_height = Some(Twips(12240));
        ppr.sect_pr = Some(sect);
        p.properties = Some(ppr);
        doc.body.add_paragraph(p);

        let xml = doc.to_xml().unwrap();
        let parsed = CT_Document::from_xml(&xml).unwrap();

        let paras: Vec<_> = parsed.body.paragraphs().collect();
        assert_eq!(paras.len(), 1);
        let ppr2 = paras[0].properties.as_ref().unwrap();
        let sect2 = ppr2.sect_pr.as_ref().unwrap();
        assert_eq!(sect2.section_type, Some(ST_SectionType::NextPage));
        assert_eq!(sect2.orientation, Some(ST_PageOrientation::Landscape));
    }
}
