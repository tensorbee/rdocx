//! Drawing elements for inline and anchor images: `CT_Drawing`, `CT_Inline`, `CT_Anchor`.

use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};

use crate::error::Result;
use crate::namespace::matches_local_name;
use crate::raw_xml::capture_element;
use crate::units::Emu;

/// Namespaces used in drawing markup.
pub mod drawing_ns {
    pub const WP: &str = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    pub const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
    pub const PIC: &str = "http://schemas.openxmlformats.org/drawingml/2006/picture";
    pub const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
}

/// Horizontal relative-from for anchor positioning.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ST_RelativeFromH {
    Page,
    Margin,
    Column,
    Character,
    LeftMargin,
    RightMargin,
    InsideMargin,
    OutsideMargin,
}

impl ST_RelativeFromH {
    pub fn from_str(s: &str) -> Self {
        match s {
            "page" => Self::Page,
            "margin" => Self::Margin,
            "column" => Self::Column,
            "character" => Self::Character,
            "leftMargin" => Self::LeftMargin,
            "rightMargin" => Self::RightMargin,
            "insideMargin" => Self::InsideMargin,
            "outsideMargin" => Self::OutsideMargin,
            _ => Self::Page,
        }
    }

    pub fn to_str(self) -> &'static str {
        match self {
            Self::Page => "page",
            Self::Margin => "margin",
            Self::Column => "column",
            Self::Character => "character",
            Self::LeftMargin => "leftMargin",
            Self::RightMargin => "rightMargin",
            Self::InsideMargin => "insideMargin",
            Self::OutsideMargin => "outsideMargin",
        }
    }
}

/// Vertical relative-from for anchor positioning.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ST_RelativeFromV {
    Page,
    Margin,
    Paragraph,
    Line,
    TopMargin,
    BottomMargin,
    InsideMargin,
    OutsideMargin,
}

impl ST_RelativeFromV {
    pub fn from_str(s: &str) -> Self {
        match s {
            "page" => Self::Page,
            "margin" => Self::Margin,
            "paragraph" => Self::Paragraph,
            "line" => Self::Line,
            "topMargin" => Self::TopMargin,
            "bottomMargin" => Self::BottomMargin,
            "insideMargin" => Self::InsideMargin,
            "outsideMargin" => Self::OutsideMargin,
            _ => Self::Page,
        }
    }

    pub fn to_str(self) -> &'static str {
        match self {
            Self::Page => "page",
            Self::Margin => "margin",
            Self::Paragraph => "paragraph",
            Self::Line => "line",
            Self::TopMargin => "topMargin",
            Self::BottomMargin => "bottomMargin",
            Self::InsideMargin => "insideMargin",
            Self::OutsideMargin => "outsideMargin",
        }
    }
}

/// Wrapping type for anchored drawings.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum WrapType {
    None,
    Square,
    TopAndBottom,
}

/// Horizontal alignment for anchored drawings.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AnchorAlignH {
    Left,
    Center,
    Right,
    Inside,
    Outside,
}

impl AnchorAlignH {
    pub fn from_str(s: &str) -> Option<Self> {
        match s {
            "left" => Some(Self::Left),
            "center" => Some(Self::Center),
            "right" => Some(Self::Right),
            "inside" => Some(Self::Inside),
            "outside" => Some(Self::Outside),
            _ => None,
        }
    }
}

/// Vertical alignment for anchored drawings.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AnchorAlignV {
    Top,
    Center,
    Bottom,
    Inside,
    Outside,
}

impl AnchorAlignV {
    pub fn from_str(s: &str) -> Option<Self> {
        match s {
            "top" => Some(Self::Top),
            "center" => Some(Self::Center),
            "bottom" => Some(Self::Bottom),
            "inside" => Some(Self::Inside),
            "outside" => Some(Self::Outside),
            _ => None,
        }
    }
}

/// `CT_Anchor` — An anchored (floating) drawing element.
#[derive(Debug, Clone, PartialEq)]
pub struct CT_Anchor {
    /// Whether the drawing is behind document text.
    pub behind_doc: bool,
    /// Horizontal position offset in EMUs.
    pub pos_h_offset: Emu,
    /// Horizontal relative-from.
    pub pos_h_relative_from: ST_RelativeFromH,
    /// Optional horizontal alignment.
    pub pos_h_align: Option<AnchorAlignH>,
    /// Vertical position offset in EMUs.
    pub pos_v_offset: Emu,
    /// Vertical relative-from.
    pub pos_v_relative_from: ST_RelativeFromV,
    /// Optional vertical alignment.
    pub pos_v_align: Option<AnchorAlignV>,
    /// Width in EMUs.
    pub extent_cx: Emu,
    /// Height in EMUs.
    pub extent_cy: Emu,
    /// Extra distance between surrounding text and the top edge.
    pub dist_t: Emu,
    /// Extra distance between surrounding text and the bottom edge.
    pub dist_b: Emu,
    /// Extra distance between surrounding text and the left edge.
    pub dist_l: Emu,
    /// Extra distance between surrounding text and the right edge.
    pub dist_r: Emu,
    /// Wrapping type.
    pub wrap: WrapType,
    /// Relationship ID referencing the image part.
    pub embed_id: String,
    /// Z-order relative height.
    pub relative_height: u32,
    /// Optional description/alt text.
    pub description: Option<String>,
    /// Optional name.
    pub name: Option<String>,
    /// Raw XML bytes for the entire wp:anchor element (used for round-trip preservation).
    /// When present, to_xml uses this instead of structured serialization.
    pub raw_xml: Option<Vec<u8>>,
}

impl CT_Anchor {
    /// Create an anchor for a full-page background image.
    pub fn background(embed_id: &str, page_width_emu: i64, page_height_emu: i64) -> Self {
        CT_Anchor {
            behind_doc: true,
            pos_h_offset: Emu(0),
            pos_h_relative_from: ST_RelativeFromH::Page,
            pos_h_align: None,
            pos_v_offset: Emu(0),
            pos_v_relative_from: ST_RelativeFromV::Page,
            pos_v_align: None,
            extent_cx: Emu(page_width_emu),
            extent_cy: Emu(page_height_emu),
            dist_t: Emu(0),
            dist_b: Emu(0),
            dist_l: Emu(0),
            dist_r: Emu(0),
            wrap: WrapType::None,
            embed_id: embed_id.to_string(),
            relative_height: 0,
            description: Some("Background".to_string()),
            name: Some("Background".to_string()),
            raw_xml: None,
        }
    }

    pub fn from_xml(reader: &mut Reader<&[u8]>, start: &BytesStart) -> Result<Self> {
        let mut behind_doc = false;
        let mut relative_height = 0u32;
        let mut dist_t = Emu(0);
        let mut dist_b = Emu(0);
        let mut dist_l = Emu(0);
        let mut dist_r = Emu(0);

        // Parse attributes from the <wp:anchor> start tag
        for attr in start.attributes() {
            let attr = attr?;
            let key = attr.key.as_ref();
            let val = std::str::from_utf8(&attr.value)?;
            if key == b"behindDoc" {
                behind_doc = val == "1" || val == "true";
            } else if key == b"relativeHeight" {
                relative_height = val.parse().unwrap_or(0);
            } else if key == b"distT" {
                dist_t = Emu(val.parse().unwrap_or(0));
            } else if key == b"distB" {
                dist_b = Emu(val.parse().unwrap_or(0));
            } else if key == b"distL" {
                dist_l = Emu(val.parse().unwrap_or(0));
            } else if key == b"distR" {
                dist_r = Emu(val.parse().unwrap_or(0));
            }
        }

        let mut pos_h_offset = Emu(0);
        let mut pos_h_relative_from = ST_RelativeFromH::Page;
        let mut pos_h_align = None;
        let mut pos_v_offset = Emu(0);
        let mut pos_v_relative_from = ST_RelativeFromV::Page;
        let mut pos_v_align = None;
        let mut extent_cx = Emu(0);
        let mut extent_cy = Emu(0);
        let mut embed_id = String::new();
        let mut description = None;
        let mut name = None;
        let mut wrap = WrapType::None;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(ref e)) => {
                    let ename = e.name();
                    if matches_local_name(ename.as_ref(), b"extent") {
                        for attr in e.attributes() {
                            let attr = attr?;
                            let key = attr.key.as_ref();
                            let val = std::str::from_utf8(&attr.value)?;
                            if key == b"cx" {
                                extent_cx = Emu(val.parse()?);
                            } else if key == b"cy" {
                                extent_cy = Emu(val.parse()?);
                            }
                        }
                    } else if matches_local_name(ename.as_ref(), b"docPr") {
                        for attr in e.attributes() {
                            let attr = attr?;
                            let key = attr.key.as_ref();
                            let val = std::str::from_utf8(&attr.value)?;
                            if key == b"descr" {
                                description = Some(val.to_string());
                            } else if key == b"name" {
                                name = Some(val.to_string());
                            }
                        }
                    } else if matches_local_name(ename.as_ref(), b"blip") {
                        for attr in e.attributes() {
                            let attr = attr?;
                            let key = attr.key.as_ref();
                            let val = std::str::from_utf8(&attr.value)?;
                            if matches_local_name(key, b"embed") {
                                embed_id = val.to_string();
                            }
                        }
                    } else if matches_local_name(ename.as_ref(), b"simplePos") {
                        // Ignore simplePos
                    } else if matches_local_name(ename.as_ref(), b"wrapNone") {
                        // Already default
                    } else if matches_local_name(ename.as_ref(), b"wrapSquare") {
                        wrap = WrapType::Square;
                    } else if matches_local_name(ename.as_ref(), b"wrapTopAndBottom") {
                        wrap = WrapType::TopAndBottom;
                    }
                }
                Ok(Event::Start(ref e)) => {
                    let ename = e.name();
                    if matches_local_name(ename.as_ref(), b"positionH") {
                        for attr in e.attributes() {
                            let attr = attr?;
                            if attr.key.as_ref() == b"relativeFrom" {
                                pos_h_relative_from =
                                    ST_RelativeFromH::from_str(std::str::from_utf8(&attr.value)?);
                            }
                        }
                        // Read child <wp:posOffset>
                        let mut inner_buf = Vec::new();
                        loop {
                            match reader.read_event_into(&mut inner_buf) {
                                Ok(Event::Start(ref ie))
                                    if matches_local_name(ie.name().as_ref(), b"posOffset") =>
                                {
                                    let text = reader.read_text(ie.name()).unwrap_or_default();
                                    pos_h_offset = Emu(text.trim().parse().unwrap_or(0));
                                }
                                Ok(Event::Start(ref ie))
                                    if matches_local_name(ie.name().as_ref(), b"align") =>
                                {
                                    let text = reader.read_text(ie.name()).unwrap_or_default();
                                    pos_h_align = AnchorAlignH::from_str(text.trim());
                                }
                                Ok(Event::End(ref ie))
                                    if matches_local_name(ie.name().as_ref(), b"positionH") =>
                                {
                                    break;
                                }
                                Ok(Event::Eof) => break,
                                Err(e) => return Err(e.into()),
                                _ => {}
                            }
                            inner_buf.clear();
                        }
                    } else if matches_local_name(ename.as_ref(), b"positionV") {
                        for attr in e.attributes() {
                            let attr = attr?;
                            if attr.key.as_ref() == b"relativeFrom" {
                                pos_v_relative_from =
                                    ST_RelativeFromV::from_str(std::str::from_utf8(&attr.value)?);
                            }
                        }
                        let mut inner_buf = Vec::new();
                        loop {
                            match reader.read_event_into(&mut inner_buf) {
                                Ok(Event::Start(ref ie))
                                    if matches_local_name(ie.name().as_ref(), b"posOffset") =>
                                {
                                    let text = reader.read_text(ie.name()).unwrap_or_default();
                                    pos_v_offset = Emu(text.trim().parse().unwrap_or(0));
                                }
                                Ok(Event::Start(ref ie))
                                    if matches_local_name(ie.name().as_ref(), b"align") =>
                                {
                                    let text = reader.read_text(ie.name()).unwrap_or_default();
                                    pos_v_align = AnchorAlignV::from_str(text.trim());
                                }
                                Ok(Event::End(ref ie))
                                    if matches_local_name(ie.name().as_ref(), b"positionV") =>
                                {
                                    break;
                                }
                                Ok(Event::Eof) => break,
                                Err(e) => return Err(e.into()),
                                _ => {}
                            }
                            inner_buf.clear();
                        }
                    } else if matches_local_name(ename.as_ref(), b"blip") {
                        for attr in e.attributes() {
                            let attr = attr?;
                            let key = attr.key.as_ref();
                            let val = std::str::from_utf8(&attr.value)?;
                            if matches_local_name(key, b"embed") {
                                embed_id = val.to_string();
                            }
                        }
                        reader.read_to_end_into(ename, &mut Vec::new())?;
                    } else if matches_local_name(ename.as_ref(), b"docPr") {
                        for attr in e.attributes() {
                            let attr = attr?;
                            let key = attr.key.as_ref();
                            let val = std::str::from_utf8(&attr.value)?;
                            if key == b"descr" {
                                description = Some(val.to_string());
                            } else if key == b"name" {
                                name = Some(val.to_string());
                            }
                        }
                        reader.read_to_end_into(ename, &mut Vec::new())?;
                    } else if matches_local_name(ename.as_ref(), b"wrapSquare") {
                        wrap = WrapType::Square;
                        reader.read_to_end_into(ename, &mut Vec::new())?;
                    } else if matches_local_name(ename.as_ref(), b"wrapTopAndBottom") {
                        wrap = WrapType::TopAndBottom;
                        reader.read_to_end_into(ename, &mut Vec::new())?;
                    } else {
                        // Continue into nested elements (graphic, graphicData, pic, etc.)
                    }
                }
                Ok(Event::End(ref e)) if matches_local_name(e.name().as_ref(), b"anchor") => {
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(CT_Anchor {
            behind_doc,
            pos_h_offset,
            pos_h_relative_from,
            pos_h_align,
            pos_v_offset,
            pos_v_relative_from,
            pos_v_align,
            extent_cx,
            extent_cy,
            dist_t,
            dist_b,
            dist_l,
            dist_r,
            wrap,
            embed_id,
            relative_height,
            description,
            name,
            raw_xml: None, // Will be set by CT_Drawing::from_xml
        })
    }

    pub fn to_xml<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        // If we have raw XML from parsing, use it for perfect round-trip
        if let Some(ref raw) = self.raw_xml {
            writer.get_mut().write_all(raw)?;
            return Ok(());
        }

        let mut buf = itoa::Buffer::new();
        let mut anchor = BytesStart::new("wp:anchor");
        anchor.push_attribute(("behindDoc", if self.behind_doc { "1" } else { "0" }));
        anchor.push_attribute(("simplePos", "0"));
        anchor.push_attribute(("relativeHeight", buf.format(self.relative_height)));
        let dist_t = self.dist_t.0.to_string();
        let dist_b = self.dist_b.0.to_string();
        let dist_l = self.dist_l.0.to_string();
        let dist_r = self.dist_r.0.to_string();
        anchor.push_attribute(("distT", dist_t.as_str()));
        anchor.push_attribute(("distB", dist_b.as_str()));
        anchor.push_attribute(("distL", dist_l.as_str()));
        anchor.push_attribute(("distR", dist_r.as_str()));
        anchor.push_attribute(("locked", "0"));
        anchor.push_attribute(("layoutInCell", "1"));
        anchor.push_attribute(("allowOverlap", "1"));
        writer.write_event(Event::Start(anchor))?;

        // wp:simplePos
        let mut sp = BytesStart::new("wp:simplePos");
        sp.push_attribute(("x", "0"));
        sp.push_attribute(("y", "0"));
        writer.write_event(Event::Empty(sp))?;

        // wp:positionH
        let mut pos_h = BytesStart::new("wp:positionH");
        pos_h.push_attribute(("relativeFrom", self.pos_h_relative_from.to_str()));
        writer.write_event(Event::Start(pos_h))?;
        writer.write_event(Event::Start(BytesStart::new("wp:posOffset")))?;
        writer.write_event(Event::Text(BytesText::new(
            &self.pos_h_offset.0.to_string(),
        )))?;
        writer.write_event(Event::End(BytesEnd::new("wp:posOffset")))?;
        writer.write_event(Event::End(BytesEnd::new("wp:positionH")))?;

        // wp:positionV
        let mut pos_v = BytesStart::new("wp:positionV");
        pos_v.push_attribute(("relativeFrom", self.pos_v_relative_from.to_str()));
        writer.write_event(Event::Start(pos_v))?;
        writer.write_event(Event::Start(BytesStart::new("wp:posOffset")))?;
        writer.write_event(Event::Text(BytesText::new(
            &self.pos_v_offset.0.to_string(),
        )))?;
        writer.write_event(Event::End(BytesEnd::new("wp:posOffset")))?;
        writer.write_event(Event::End(BytesEnd::new("wp:positionV")))?;

        // wp:extent
        let mut extent = BytesStart::new("wp:extent");
        extent.push_attribute(("cx", buf.format(self.extent_cx.0)));
        extent.push_attribute(("cy", buf.format(self.extent_cy.0)));
        writer.write_event(Event::Empty(extent))?;

        // wp:wrapNone
        writer.write_event(Event::Empty(BytesStart::new("wp:wrapNone")))?;

        // wp:docPr
        let mut doc_pr = BytesStart::new("wp:docPr");
        doc_pr.push_attribute(("id", "1"));
        doc_pr.push_attribute(("name", self.name.as_deref().unwrap_or("Picture")));
        if let Some(ref desc) = self.description {
            doc_pr.push_attribute(("descr", desc.as_str()));
        }
        writer.write_event(Event::Empty(doc_pr))?;

        // a:graphic (same pic:pic structure as inline)
        write_graphic_element(
            writer,
            &self.embed_id,
            self.extent_cx,
            self.extent_cy,
            self.name.as_deref(),
        )?;

        writer.write_event(Event::End(BytesEnd::new("wp:anchor")))?;
        Ok(())
    }
}

/// `CT_Inline` — An inline drawing (image) element.
#[derive(Debug, Clone, PartialEq)]
pub struct CT_Inline {
    /// Width in EMUs
    pub extent_cx: Emu,
    /// Height in EMUs
    pub extent_cy: Emu,
    /// Relationship ID referencing the image part
    pub embed_id: String,
    /// Optional description/alt text
    pub description: Option<String>,
    /// Optional name
    pub name: Option<String>,
    /// Raw XML bytes for the entire wp:inline element (used for round-trip preservation).
    /// When present, to_xml uses this instead of structured serialization.
    pub raw_xml: Option<Vec<u8>>,
}

impl CT_Inline {
    pub fn new(embed_id: &str, width_emu: i64, height_emu: i64) -> Self {
        CT_Inline {
            extent_cx: Emu(width_emu),
            extent_cy: Emu(height_emu),
            embed_id: embed_id.to_string(),
            description: None,
            name: None,
            raw_xml: None,
        }
    }

    pub fn from_xml(reader: &mut Reader<&[u8]>) -> Result<Self> {
        let mut cx = Emu(0);
        let mut cy = Emu(0);
        let mut embed_id = String::new();
        let mut description = None;
        let mut name = None;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(ref e)) => {
                    let ename = e.name();
                    if matches_local_name(ename.as_ref(), b"extent") {
                        for attr in e.attributes() {
                            let attr = attr?;
                            let key = attr.key.as_ref();
                            let val = std::str::from_utf8(&attr.value)?;
                            if key == b"cx" {
                                cx = Emu(val.parse()?);
                            } else if key == b"cy" {
                                cy = Emu(val.parse()?);
                            }
                        }
                    } else if matches_local_name(ename.as_ref(), b"docPr") {
                        for attr in e.attributes() {
                            let attr = attr?;
                            let key = attr.key.as_ref();
                            let val = std::str::from_utf8(&attr.value)?;
                            if key == b"descr" {
                                description = Some(val.to_string());
                            } else if key == b"name" {
                                name = Some(val.to_string());
                            }
                        }
                    } else if matches_local_name(ename.as_ref(), b"blip") {
                        for attr in e.attributes() {
                            let attr = attr?;
                            let key = attr.key.as_ref();
                            let val = std::str::from_utf8(&attr.value)?;
                            if matches_local_name(key, b"embed") {
                                embed_id = val.to_string();
                            }
                        }
                    }
                }
                Ok(Event::Start(ref e)) => {
                    let ename = e.name();
                    if matches_local_name(ename.as_ref(), b"blip") {
                        for attr in e.attributes() {
                            let attr = attr?;
                            let key = attr.key.as_ref();
                            let val = std::str::from_utf8(&attr.value)?;
                            if matches_local_name(key, b"embed") {
                                embed_id = val.to_string();
                            }
                        }
                        reader.read_to_end_into(ename, &mut Vec::new())?;
                    } else if matches_local_name(ename.as_ref(), b"docPr") {
                        for attr in e.attributes() {
                            let attr = attr?;
                            let key = attr.key.as_ref();
                            let val = std::str::from_utf8(&attr.value)?;
                            if key == b"descr" {
                                description = Some(val.to_string());
                            } else if key == b"name" {
                                name = Some(val.to_string());
                            }
                        }
                        reader.read_to_end_into(ename, &mut Vec::new())?;
                    } else if !matches_local_name(ename.as_ref(), b"inline") {
                        // Continue parsing nested elements (graphic, graphicData, pic, etc.)
                    }
                }
                Ok(Event::End(ref e)) if matches_local_name(e.name().as_ref(), b"inline") => {
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(CT_Inline {
            extent_cx: cx,
            extent_cy: cy,
            embed_id,
            description,
            name,
            raw_xml: None, // Will be set by CT_Drawing::from_xml
        })
    }

    pub fn to_xml<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        // If we have raw XML from parsing, use it for perfect round-trip
        if let Some(ref raw) = self.raw_xml {
            writer.get_mut().write_all(raw)?;
            return Ok(());
        }

        // wp:inline
        let mut buf = itoa::Buffer::new();
        let mut inline = BytesStart::new("wp:inline");
        inline.push_attribute(("distT", "0"));
        inline.push_attribute(("distB", "0"));
        inline.push_attribute(("distL", "0"));
        inline.push_attribute(("distR", "0"));
        writer.write_event(Event::Start(inline))?;

        // wp:extent
        let mut extent = BytesStart::new("wp:extent");
        extent.push_attribute(("cx", buf.format(self.extent_cx.0)));
        extent.push_attribute(("cy", buf.format(self.extent_cy.0)));
        writer.write_event(Event::Empty(extent))?;

        // wp:docPr
        let mut doc_pr = BytesStart::new("wp:docPr");
        doc_pr.push_attribute(("id", "1"));
        doc_pr.push_attribute(("name", self.name.as_deref().unwrap_or("Picture")));
        if let Some(ref desc) = self.description {
            doc_pr.push_attribute(("descr", desc.as_str()));
        }
        writer.write_event(Event::Empty(doc_pr))?;

        // a:graphic
        write_graphic_element(
            writer,
            &self.embed_id,
            self.extent_cx,
            self.extent_cy,
            self.name.as_deref(),
        )?;

        writer.write_event(Event::End(BytesEnd::new("wp:inline")))?;

        Ok(())
    }
}

/// Write the `a:graphic` > `a:graphicData` > `pic:pic` structure (shared by inline and anchor).
fn write_graphic_element<W: std::io::Write>(
    writer: &mut Writer<W>,
    embed_id: &str,
    cx: Emu,
    cy: Emu,
    name: Option<&str>,
) -> Result<()> {
    let mut buf = itoa::Buffer::new();
    let mut graphic = BytesStart::new("a:graphic");
    graphic.push_attribute(("xmlns:a", drawing_ns::A));
    writer.write_event(Event::Start(graphic))?;

    let mut gd = BytesStart::new("a:graphicData");
    gd.push_attribute(("uri", drawing_ns::PIC));
    writer.write_event(Event::Start(gd))?;

    let mut pic = BytesStart::new("pic:pic");
    pic.push_attribute(("xmlns:pic", drawing_ns::PIC));
    writer.write_event(Event::Start(pic))?;

    // pic:nvPicPr
    writer.write_event(Event::Start(BytesStart::new("pic:nvPicPr")))?;
    let mut cnvpr = BytesStart::new("pic:cNvPr");
    cnvpr.push_attribute(("id", "0"));
    cnvpr.push_attribute(("name", name.unwrap_or("Picture")));
    writer.write_event(Event::Empty(cnvpr))?;
    writer.write_event(Event::Empty(BytesStart::new("pic:cNvPicPr")))?;
    writer.write_event(Event::End(BytesEnd::new("pic:nvPicPr")))?;

    // pic:blipFill
    writer.write_event(Event::Start(BytesStart::new("pic:blipFill")))?;
    let mut blip = BytesStart::new("a:blip");
    blip.push_attribute(("r:embed", embed_id));
    writer.write_event(Event::Empty(blip))?;
    writer.write_event(Event::Start(BytesStart::new("a:stretch")))?;
    writer.write_event(Event::Empty(BytesStart::new("a:fillRect")))?;
    writer.write_event(Event::End(BytesEnd::new("a:stretch")))?;
    writer.write_event(Event::End(BytesEnd::new("pic:blipFill")))?;

    // pic:spPr
    writer.write_event(Event::Start(BytesStart::new("pic:spPr")))?;
    writer.write_event(Event::Start(BytesStart::new("a:xfrm")))?;
    let mut off = BytesStart::new("a:off");
    off.push_attribute(("x", "0"));
    off.push_attribute(("y", "0"));
    writer.write_event(Event::Empty(off))?;
    let mut ext = BytesStart::new("a:ext");
    ext.push_attribute(("cx", buf.format(cx.0)));
    ext.push_attribute(("cy", buf.format(cy.0)));
    writer.write_event(Event::Empty(ext))?;
    writer.write_event(Event::End(BytesEnd::new("a:xfrm")))?;
    let mut prst = BytesStart::new("a:prstGeom");
    prst.push_attribute(("prst", "rect"));
    writer.write_event(Event::Start(prst))?;
    writer.write_event(Event::Empty(BytesStart::new("a:avLst")))?;
    writer.write_event(Event::End(BytesEnd::new("a:prstGeom")))?;
    writer.write_event(Event::End(BytesEnd::new("pic:spPr")))?;

    writer.write_event(Event::End(BytesEnd::new("pic:pic")))?;
    writer.write_event(Event::End(BytesEnd::new("a:graphicData")))?;
    writer.write_event(Event::End(BytesEnd::new("a:graphic")))?;

    Ok(())
}

/// `CT_Drawing` — A drawing element that wraps inline or anchor images.
#[derive(Debug, Clone, PartialEq)]
pub struct CT_Drawing {
    pub inline: Option<CT_Inline>,
    pub anchor: Option<CT_Anchor>,
}

impl CT_Drawing {
    pub fn inline(inline: CT_Inline) -> Self {
        CT_Drawing {
            inline: Some(inline),
            anchor: None,
        }
    }

    pub fn anchor(anchor: CT_Anchor) -> Self {
        CT_Drawing {
            inline: None,
            anchor: Some(anchor),
        }
    }

    pub fn from_xml(reader: &mut Reader<&[u8]>) -> Result<Self> {
        let mut inline = None;
        let mut anchor = None;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let name = e.name();
                    if matches_local_name(name.as_ref(), b"inline") {
                        // Capture full raw XML, then re-parse for structured fields
                        let raw = capture_element(reader, e)?;
                        let mut re_reader = Reader::from_reader(raw.as_slice());
                        re_reader.config_mut().trim_text(true);
                        // Skip to the <wp:inline> start
                        let mut rbuf = Vec::new();
                        loop {
                            match re_reader.read_event_into(&mut rbuf) {
                                Ok(Event::Start(ref ie))
                                    if matches_local_name(ie.name().as_ref(), b"inline") =>
                                {
                                    let mut inl = CT_Inline::from_xml(&mut re_reader)?;
                                    inl.raw_xml = Some(raw);
                                    inline = Some(inl);
                                    break;
                                }
                                Ok(Event::Eof) => break,
                                Err(e) => return Err(e.into()),
                                _ => {}
                            }
                            rbuf.clear();
                        }
                    } else if matches_local_name(name.as_ref(), b"anchor") {
                        // Capture full raw XML, then re-parse for structured fields
                        let raw = capture_element(reader, e)?;
                        let mut re_reader = Reader::from_reader(raw.as_slice());
                        re_reader.config_mut().trim_text(true);
                        let mut rbuf = Vec::new();
                        loop {
                            match re_reader.read_event_into(&mut rbuf) {
                                Ok(Event::Start(ref ie))
                                    if matches_local_name(ie.name().as_ref(), b"anchor") =>
                                {
                                    let mut anc = CT_Anchor::from_xml(&mut re_reader, ie)?;
                                    anc.raw_xml = Some(raw);
                                    anchor = Some(anc);
                                    break;
                                }
                                Ok(Event::Eof) => break,
                                Err(e) => return Err(e.into()),
                                _ => {}
                            }
                            rbuf.clear();
                        }
                    } else {
                        reader.read_to_end_into(name, &mut Vec::new())?;
                    }
                }
                Ok(Event::End(ref e)) if matches_local_name(e.name().as_ref(), b"drawing") => {
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(CT_Drawing { inline, anchor })
    }

    pub fn to_xml<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let drawing = BytesStart::new("w:drawing");
        writer.write_event(Event::Start(drawing))?;

        if let Some(ref inl) = self.inline {
            inl.to_xml(writer)?;
        }
        if let Some(ref anc) = self.anchor {
            anc.to_xml(writer)?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:drawing")))?;
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn parse_drawing(xml: &str) -> CT_Drawing {
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if matches_local_name(e.name().as_ref(), b"drawing") => {
                    break;
                }
                _ => {}
            }
            buf.clear();
        }
        CT_Drawing::from_xml(&mut reader).unwrap()
    }

    #[test]
    fn round_trip_inline_drawing() {
        let inline = CT_Inline {
            extent_cx: Emu(914400), // 1 inch
            extent_cy: Emu(457200), // 0.5 inch
            embed_id: "rId5".to_string(),
            description: Some("A test image".to_string()),
            name: Some("TestPic".to_string()),
            raw_xml: None,
        };

        let drawing = CT_Drawing::inline(inline);

        let mut output = Vec::new();
        let mut writer = Writer::new(&mut output);
        drawing.to_xml(&mut writer).unwrap();
        let xml = String::from_utf8(output).unwrap();

        let parsed = parse_drawing(&xml);
        let inl = parsed.inline.unwrap();
        assert_eq!(inl.extent_cx, Emu(914400));
        assert_eq!(inl.extent_cy, Emu(457200));
        assert_eq!(inl.embed_id, "rId5");
    }

    #[test]
    fn ct_anchor_background_constructor() {
        let anchor = CT_Anchor::background("rId1", 7772400, 10058400);
        assert!(anchor.behind_doc);
        assert_eq!(anchor.pos_h_offset, Emu(0));
        assert_eq!(anchor.pos_v_offset, Emu(0));
        assert_eq!(anchor.pos_h_relative_from, ST_RelativeFromH::Page);
        assert_eq!(anchor.pos_v_relative_from, ST_RelativeFromV::Page);
        assert_eq!(anchor.extent_cx, Emu(7772400));
        assert_eq!(anchor.extent_cy, Emu(10058400));
        assert_eq!(anchor.embed_id, "rId1");
        assert_eq!(anchor.relative_height, 0);
    }

    #[test]
    fn ct_anchor_round_trip_xml() {
        let anchor = CT_Anchor::background("rId3", 7772400, 10058400);

        let drawing = CT_Drawing::anchor(anchor);
        let mut output = Vec::new();
        let mut writer = Writer::new(&mut output);
        drawing.to_xml(&mut writer).unwrap();
        let xml = String::from_utf8(output).unwrap();

        let parsed = parse_drawing(&xml);
        assert!(parsed.anchor.is_some());
        let anc = parsed.anchor.unwrap();
        assert!(anc.behind_doc);
        assert_eq!(anc.pos_h_offset, Emu(0));
        assert_eq!(anc.pos_v_offset, Emu(0));
        assert_eq!(anc.pos_h_relative_from, ST_RelativeFromH::Page);
        assert_eq!(anc.pos_v_relative_from, ST_RelativeFromV::Page);
        assert_eq!(anc.extent_cx, Emu(7772400));
        assert_eq!(anc.extent_cy, Emu(10058400));
        assert_eq!(anc.embed_id, "rId3");
    }

    #[test]
    fn ct_drawing_with_anchor_and_inline() {
        // A drawing can have either inline or anchor (not both in practice, but test both paths)
        let inline = CT_Inline::new("rId1", 914400, 457200);
        let d1 = CT_Drawing::inline(inline);
        assert!(d1.inline.is_some());
        assert!(d1.anchor.is_none());

        let anchor = CT_Anchor::background("rId2", 7772400, 10058400);
        let d2 = CT_Drawing::anchor(anchor);
        assert!(d2.inline.is_none());
        assert!(d2.anchor.is_some());
    }
}
