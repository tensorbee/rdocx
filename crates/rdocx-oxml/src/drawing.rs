//! Drawing elements for inline and anchor images: `CT_Drawing`, `CT_Inline`, `CT_Anchor`.

use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use quick_xml::{Reader, Writer, XmlVersion};

use crate::error::Result;
use crate::namespace::matches_local_name;
use crate::numbering::namespace_bindings;
use crate::raw_xml::capture_element;
use crate::units::Emu;

/// Namespaces used in drawing markup.
pub mod drawing_ns {
    pub const WP: &str = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    pub const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
    pub const PIC: &str = "http://schemas.openxmlformats.org/drawingml/2006/picture";
    pub const C: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
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
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum WrapType {
    /// Text flows over or under the drawing without reserving space.
    #[default]
    None,
    /// Text keeps clear of the drawing's frame.
    Square,
    /// Text clears the drawing above and below, using the full width.
    TopAndBottom,
    /// Text follows the drawing's outline.
    Tight,
    /// Text follows the outline and may enter its concave regions.
    Through,
}

impl WrapType {
    /// The wrapping element's local name, or `None` when there is no wrap.
    fn element_name(self) -> &'static str {
        match self {
            WrapType::None => "wp:wrapNone",
            WrapType::Square => "wp:wrapSquare",
            WrapType::TopAndBottom => "wp:wrapTopAndBottom",
            WrapType::Tight => "wp:wrapTight",
            WrapType::Through => "wp:wrapThrough",
        }
    }
}

/// The wrap mode a wrapping element names, or `None` if it is not one.
fn wrap_type_of(name: &[u8]) -> Option<WrapType> {
    if matches_local_name(name, b"wrapNone") {
        Some(WrapType::None)
    } else if matches_local_name(name, b"wrapSquare") {
        Some(WrapType::Square)
    } else if matches_local_name(name, b"wrapTopAndBottom") {
        Some(WrapType::TopAndBottom)
    } else if matches_local_name(name, b"wrapTight") {
        Some(WrapType::Tight)
    } else if matches_local_name(name, b"wrapThrough") {
        Some(WrapType::Through)
    } else {
        None
    }
}

/// Horizontal alignment for an anchored drawing.
///
/// An anchor positions itself either by an offset or by an alignment. When an
/// alignment is given the offset is not used.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AnchorAlignH {
    Left,
    Center,
    Right,
    Inside,
    Outside,
}

impl AnchorAlignH {
    /// Parse an alignment. An unrecognised value reads as no alignment, which
    /// falls back to the offset rather than inventing a position.
    pub fn parse(s: &str) -> Option<Self> {
        match s {
            "left" => Some(Self::Left),
            "center" => Some(Self::Center),
            "right" => Some(Self::Right),
            "inside" => Some(Self::Inside),
            "outside" => Some(Self::Outside),
            _ => None,
        }
    }

    pub fn to_str(self) -> &'static str {
        match self {
            Self::Left => "left",
            Self::Center => "center",
            Self::Right => "right",
            Self::Inside => "inside",
            Self::Outside => "outside",
        }
    }
}

/// Vertical alignment for an anchored drawing.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AnchorAlignV {
    Top,
    Center,
    Bottom,
    Inside,
    Outside,
}

impl AnchorAlignV {
    /// Parse an alignment. An unrecognised value reads as no alignment.
    pub fn parse(s: &str) -> Option<Self> {
        match s {
            "top" => Some(Self::Top),
            "center" => Some(Self::Center),
            "bottom" => Some(Self::Bottom),
            "inside" => Some(Self::Inside),
            "outside" => Some(Self::Outside),
            _ => None,
        }
    }

    pub fn to_str(self) -> &'static str {
        match self {
            Self::Top => "top",
            Self::Center => "center",
            Self::Bottom => "bottom",
            Self::Inside => "inside",
            Self::Outside => "outside",
        }
    }
}

/// `CT_Anchor` — An anchored (floating) drawing element.
#[derive(Debug, Clone, PartialEq)]
pub struct CT_Anchor {
    /// Non-visual drawing properties identifier (`wp:docPr/@id`).
    pub doc_pr_id: u32,
    /// Whether the drawing is behind document text.
    pub behind_doc: bool,
    /// Horizontal position offset in EMUs.
    pub pos_h_offset: Emu,
    /// Horizontal relative-from.
    pub pos_h_relative_from: ST_RelativeFromH,
    /// Horizontal alignment, used instead of the offset when present.
    pub pos_h_align: Option<AnchorAlignH>,
    /// Vertical position offset in EMUs.
    pub pos_v_offset: Emu,
    /// Vertical relative-from.
    pub pos_v_relative_from: ST_RelativeFromV,
    /// Vertical alignment, used instead of the offset when present.
    pub pos_v_align: Option<AnchorAlignV>,
    /// Width in EMUs.
    pub extent_cx: Emu,
    /// Height in EMUs.
    pub extent_cy: Emu,
    /// Space kept between the drawing and the text wrapping around it.
    pub dist_t: Emu,
    pub dist_b: Emu,
    pub dist_l: Emu,
    pub dist_r: Emu,
    /// Wrapping type.
    pub wrap: WrapType,
    /// Relationship ID referencing the image part.
    pub embed_id: String,
    /// Relationship ID referencing an externally linked image.
    pub link_id: Option<String>,
    /// Relationship ID referencing a ChartML part.
    pub chart_rel_id: Option<String>,
    /// Z-order relative height.
    pub relative_height: u32,
    /// Optional description/alt text.
    pub description: Option<String>,
    /// Optional name.
    pub name: Option<String>,
    /// Raw XML bytes for the entire wp:anchor element (used for round-trip preservation).
    /// When present, to_xml uses this instead of structured serialization.
    pub raw_xml: Option<Vec<u8>>,
    /// Shape content, when the anchor holds a `wps:wsp` rather than a picture.
    pub shape: Option<CT_Shape>,
}

/// A `wps:wsp` shape: preset geometry, an optional fill, and optional text.
///
/// Word writes these inside `mc:AlternateContent`, with a VML fallback beside
/// them. Only what we can actually draw is modelled here.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct CT_Shape {
    /// Preset geometry name from `a:prstGeom@prst`, e.g. "rect" or "line".
    pub preset: Option<String>,
    /// Solid fill colour as RRGGBB, from the shape's own fill rather than its
    /// outline. `None` covers both `a:noFill` and a fill we cannot resolve,
    /// such as a theme colour.
    pub solid_fill: Option<String>,
    /// Paragraphs of the shape's text box, from `wps:txbx/w:txbxContent`.
    pub text: Vec<crate::text::CT_P>,
}

/// Parse the DrawingML held inside a captured `mc:AlternateContent` block.
///
/// Word writes a shape as a compatibility block: the modern DrawingML sits in
/// `mc:Choice` and a VML fallback in `mc:Fallback`. We read the Choice and
/// ignore the Fallback.
///
/// The caller keeps the raw bytes for write back, so whatever comes out of
/// here must not be serialised again or the element ends up duplicated.
pub fn parse_alternate_content(raw: &[u8], inherited_prefixes: &[String]) -> Option<CT_Drawing> {
    let mut reader = Reader::from_reader(raw);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut in_choice = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let name = e.name();
                let local = name.as_ref();
                if matches_local_name(local, b"Choice") {
                    in_choice = true;
                } else if in_choice && matches_local_name(local, b"drawing") {
                    let prefixes =
                        crate::numbering::word_prefixes_at(e, inherited_prefixes).ok()?;
                    return CT_Drawing::from_xml_with_prefixes(&mut reader, &prefixes, &[]).ok();
                }
            }
            Ok(Event::End(ref e)) if matches_local_name(e.name().as_ref(), b"Choice") => {
                in_choice = false;
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    None
}

/// Pull the preset geometry and solid fill out of a captured `wps:spPr`.
///
/// The fill has to be told apart from the outline colour. Both are written as
/// `a:srgbClr`, and the outline sits inside `a:ln`, so anything at or below an
/// `a:ln` is skipped.
fn parse_shape_props(raw: &[u8]) -> (Option<String>, Option<String>) {
    let mut reader = Reader::from_reader(raw);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut preset = None;
    let mut fill = None;
    let mut ln_depth = 0usize;
    let mut in_solid_fill = false;

    // Only a Start can open a scope. A self-closing a:ln or a:solidFill has no
    // children, so it must not change the depth or the fill flag.
    loop {
        let event = reader.read_event_into(&mut buf);
        match event {
            Ok(Event::Start(ref e)) => {
                let name = e.name();
                let local = name.as_ref();
                if matches_local_name(local, b"ln") {
                    ln_depth += 1;
                } else if matches_local_name(local, b"solidFill") {
                    in_solid_fill = true;
                } else if matches_local_name(local, b"prstGeom") {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"prst" {
                            preset = std::str::from_utf8(&attr.value).ok().map(str::to_string);
                        }
                    }
                } else if matches_local_name(local, b"srgbClr")
                    && in_solid_fill
                    && ln_depth == 0
                    && fill.is_none()
                {
                    // srgbClr can carry children such as a:alpha, so it turns
                    // up as a Start as well as an Empty.
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"val" {
                            fill = std::str::from_utf8(&attr.value).ok().map(str::to_string);
                        }
                    }
                }
            }
            Ok(Event::Empty(ref e)) => {
                let name = e.name();
                let local = name.as_ref();
                if matches_local_name(local, b"prstGeom") {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"prst" {
                            preset = std::str::from_utf8(&attr.value).ok().map(str::to_string);
                        }
                    }
                } else if matches_local_name(local, b"srgbClr")
                    && in_solid_fill
                    && ln_depth == 0
                    && fill.is_none()
                {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"val" {
                            fill = std::str::from_utf8(&attr.value).ok().map(str::to_string);
                        }
                    }
                }
            }
            Ok(Event::End(ref e)) => {
                let name = e.name();
                let local = name.as_ref();
                if matches_local_name(local, b"ln") {
                    ln_depth = ln_depth.saturating_sub(1);
                } else if matches_local_name(local, b"solidFill") {
                    in_solid_fill = false;
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    (preset, fill)
}

impl CT_Anchor {
    /// Create an anchor for a full-page background image.
    pub fn background(embed_id: &str, page_width_emu: i64, page_height_emu: i64) -> Self {
        CT_Anchor {
            doc_pr_id: 1,
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
            link_id: None,
            chart_rel_id: None,
            relative_height: 0,
            description: Some("Background".to_string()),
            name: Some("Background".to_string()),
            raw_xml: None,
            shape: None,
        }
    }

    /// Create an anchored drawing whose payload is a native Word chart.
    pub fn new_chart(chart_rel_id: &str, width_emu: i64, height_emu: i64) -> Self {
        let mut anchor = Self::background("", width_emu, height_emu);
        anchor.behind_doc = false;
        anchor.chart_rel_id = Some(chart_rel_id.to_owned());
        anchor.description = None;
        anchor.name = Some("Chart".to_owned());
        anchor
    }

    pub fn from_xml(reader: &mut NsReader<&[u8]>, start: &BytesStart) -> Result<Self> {
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
        let mut wrap = WrapType::None;
        let mut extent_cx = Emu(0);
        let mut extent_cy = Emu(0);
        let mut shape: Option<CT_Shape> = None;
        let mut description = None;
        let mut name = None;
        let mut doc_pr_id = None;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(ref e)) => {
                    let ename = e.name();
                    if matches_local_name(ename.as_ref(), b"extent") {
                        for attr in e.attributes() {
                            let attr = attr?;
                            let key = attr.key.as_ref();
                            let val = attr.decoded_and_normalized_value(
                                XmlVersion::Implicit1_0,
                                e.decoder(),
                            )?;
                            if key == b"cx" {
                                extent_cx = Emu(val.parse()?);
                            } else if key == b"cy" {
                                extent_cy = Emu(val.parse()?);
                            }
                        }
                    } else if canonical_wp_element(reader, e, b"docPr") {
                        for attr in e.attributes() {
                            let attr = attr?;
                            let key = attr.key.as_ref();
                            let val = attr.decoded_and_normalized_value(
                                XmlVersion::Implicit1_0,
                                e.decoder(),
                            )?;
                            if key == b"descr" {
                                description = Some(val.to_string());
                            } else if key == b"name" {
                                name = Some(val.to_string());
                            } else if key == b"id" {
                                doc_pr_id = Some(val.parse()?);
                            }
                        }
                    } else if matches_local_name(ename.as_ref(), b"simplePos") {
                        // Ignore simplePos
                    } else if let Some(parsed) = wrap_type_of(ename.as_ref()) {
                        wrap = parsed;
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
                                    let text = reader
                                        .read_text(ie.name())
                                        .map(|t| crate::xml_text::decode_escaped(&t))
                                        .unwrap_or_default();
                                    pos_h_offset = Emu(text.trim().parse().unwrap_or(0));
                                }
                                Ok(Event::Start(ref ie))
                                    if matches_local_name(ie.name().as_ref(), b"align") =>
                                {
                                    let text = reader
                                        .read_text(ie.name())
                                        .map(|t| crate::xml_text::decode_escaped(&t))
                                        .unwrap_or_default();
                                    pos_h_align = AnchorAlignH::parse(text.trim());
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
                    } else if let Some(parsed) = wrap_type_of(ename.as_ref()) {
                        // Expanded spelling. wrapSquare and the outline wraps
                        // carry children we do not model, so skip the subtree.
                        wrap = parsed;
                        reader.read_to_end_into(ename, &mut Vec::new())?;
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
                                    let text = reader
                                        .read_text(ie.name())
                                        .map(|t| crate::xml_text::decode_escaped(&t))
                                        .unwrap_or_default();
                                    pos_v_offset = Emu(text.trim().parse().unwrap_or(0));
                                }
                                Ok(Event::Start(ref ie))
                                    if matches_local_name(ie.name().as_ref(), b"align") =>
                                {
                                    let text = reader
                                        .read_text(ie.name())
                                        .map(|t| crate::xml_text::decode_escaped(&t))
                                        .unwrap_or_default();
                                    pos_v_align = AnchorAlignV::parse(text.trim());
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
                    } else if matches_local_name(ename.as_ref(), b"spPr") {
                        // Capture the shape properties and read geometry and
                        // fill out of them separately, so the fill colour is
                        // not confused with the outline colour.
                        let raw = capture_ns_element(reader, e)?;
                        let (preset, solid_fill) = parse_shape_props(&raw);
                        let s = shape.get_or_insert_with(CT_Shape::default);
                        s.preset = preset;
                        s.solid_fill = solid_fill;
                    } else if matches_local_name(ename.as_ref(), b"txbxContent") {
                        // A shape's text box holds ordinary w:p paragraphs.
                        let mut inner_buf = Vec::new();
                        let mut paragraphs = Vec::new();
                        loop {
                            match reader.read_event_into(&mut inner_buf) {
                                Ok(Event::Start(ref ie))
                                    if matches_local_name(ie.name().as_ref(), b"p") =>
                                {
                                    let raw = capture_ns_element(reader, ie)?;
                                    let mut paragraph_reader = Reader::from_reader(raw.as_slice());
                                    let mut paragraph_buffer = Vec::new();
                                    loop {
                                        match paragraph_reader
                                            .read_event_into(&mut paragraph_buffer)?
                                        {
                                            Event::Start(ref paragraph_start)
                                                if matches_local_name(
                                                    paragraph_start.name().as_ref(),
                                                    b"p",
                                                ) =>
                                            {
                                                paragraphs.push(crate::text::CT_P::from_xml(
                                                    &mut paragraph_reader,
                                                )?);
                                                break;
                                            }
                                            Event::Eof => break,
                                            _ => {}
                                        }
                                        paragraph_buffer.clear();
                                    }
                                }
                                Ok(Event::End(ref ie))
                                    if matches_local_name(ie.name().as_ref(), b"txbxContent") =>
                                {
                                    break;
                                }
                                Ok(Event::Eof) => break,
                                Err(e) => return Err(e.into()),
                                _ => {}
                            }
                            inner_buf.clear();
                        }
                        shape.get_or_insert_with(CT_Shape::default).text = paragraphs;
                    } else if matches_local_name(ename.as_ref(), b"blip") {
                        reader.read_to_end_into(ename, &mut Vec::new())?;
                    } else if canonical_wp_element(reader, e, b"docPr") {
                        for attr in e.attributes() {
                            let attr = attr?;
                            let key = attr.key.as_ref();
                            let val = attr.decoded_and_normalized_value(
                                XmlVersion::Implicit1_0,
                                e.decoder(),
                            )?;
                            if key == b"descr" {
                                description = Some(val.to_string());
                            } else if key == b"name" {
                                name = Some(val.to_string());
                            } else if key == b"id" {
                                doc_pr_id = Some(val.parse()?);
                            }
                        }
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
            doc_pr_id: doc_pr_id
                .ok_or_else(|| crate::OxmlError::MissingElement("wp:docPr/@id".to_owned()))?,
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
            embed_id: String::new(),
            link_id: None,
            chart_rel_id: None,
            relative_height,
            description,
            name,
            raw_xml: None, // Will be set by CT_Drawing::from_xml
            shape,
        })
    }

    pub fn to_xml<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        // If we have raw XML from parsing, use it for perfect round-trip
        if let Some(ref raw) = self.raw_xml {
            writer.write_indent()?;
            writer.get_mut().write_all(raw)?;
            return Ok(());
        }

        let payload = drawing_payload(&self.embed_id, self.chart_rel_id.as_deref())?;

        let mut buf = itoa::Buffer::new();
        let mut anchor = BytesStart::new("wp:anchor");
        anchor.push_attribute(("behindDoc", if self.behind_doc { "1" } else { "0" }));
        anchor.push_attribute(("simplePos", "0"));
        anchor.push_attribute(("relativeHeight", buf.format(self.relative_height)));
        // A zero distance is the default, and an absent attribute means the
        // same thing. Emitting zeros would change the bytes of every anchor
        // that never asked for a wrap distance, for no difference in meaning.
        let dist_t = self.dist_t.0.to_string();
        let dist_b = self.dist_b.0.to_string();
        let dist_l = self.dist_l.0.to_string();
        let dist_r = self.dist_r.0.to_string();
        if self.dist_t.0 != 0 {
            anchor.push_attribute(("distT", dist_t.as_str()));
        }
        if self.dist_b.0 != 0 {
            anchor.push_attribute(("distB", dist_b.as_str()));
        }
        if self.dist_l.0 != 0 {
            anchor.push_attribute(("distL", dist_l.as_str()));
        }
        if self.dist_r.0 != 0 {
            anchor.push_attribute(("distR", dist_r.as_str()));
        }
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
        // An anchor positions by an alignment or by an offset, never both.
        if let Some(align) = self.pos_h_align {
            writer.write_event(Event::Start(BytesStart::new("wp:align")))?;
            writer.write_event(Event::Text(BytesText::new(align.to_str())))?;
            writer.write_event(Event::End(BytesEnd::new("wp:align")))?;
        } else {
            writer.write_event(Event::Start(BytesStart::new("wp:posOffset")))?;
            writer.write_event(Event::Text(BytesText::new(
                &self.pos_h_offset.0.to_string(),
            )))?;
            writer.write_event(Event::End(BytesEnd::new("wp:posOffset")))?;
        }
        writer.write_event(Event::End(BytesEnd::new("wp:positionH")))?;

        // wp:positionV
        let mut pos_v = BytesStart::new("wp:positionV");
        pos_v.push_attribute(("relativeFrom", self.pos_v_relative_from.to_str()));
        writer.write_event(Event::Start(pos_v))?;
        if let Some(align) = self.pos_v_align {
            writer.write_event(Event::Start(BytesStart::new("wp:align")))?;
            writer.write_event(Event::Text(BytesText::new(align.to_str())))?;
            writer.write_event(Event::End(BytesEnd::new("wp:align")))?;
        } else {
            writer.write_event(Event::Start(BytesStart::new("wp:posOffset")))?;
            writer.write_event(Event::Text(BytesText::new(
                &self.pos_v_offset.0.to_string(),
            )))?;
            writer.write_event(Event::End(BytesEnd::new("wp:posOffset")))?;
        }
        writer.write_event(Event::End(BytesEnd::new("wp:positionV")))?;

        // wp:extent
        let mut extent = BytesStart::new("wp:extent");
        extent.push_attribute(("cx", buf.format(self.extent_cx.0)));
        extent.push_attribute(("cy", buf.format(self.extent_cy.0)));
        writer.write_event(Event::Empty(extent))?;

        // The wrapping element, in the sequence position wp:wrapNone held.
        writer.write_event(Event::Empty(BytesStart::new(self.wrap.element_name())))?;

        // wp:docPr
        let mut doc_pr = BytesStart::new("wp:docPr");
        doc_pr.push_attribute(("id", buf.format(self.doc_pr_id)));
        doc_pr.push_attribute(("name", self.name.as_deref().unwrap_or("Picture")));
        if let Some(ref desc) = self.description {
            doc_pr.push_attribute(("descr", desc.as_str()));
        }
        writer.write_event(Event::Empty(doc_pr))?;

        write_graphic_element(
            writer,
            payload,
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
    /// Non-visual drawing properties identifier (`wp:docPr/@id`).
    pub doc_pr_id: u32,
    /// Width in EMUs
    pub extent_cx: Emu,
    /// Height in EMUs
    pub extent_cy: Emu,
    /// Relationship ID referencing the image part
    pub embed_id: String,
    /// Relationship ID referencing an externally linked image.
    pub link_id: Option<String>,
    /// Relationship ID referencing a ChartML part.
    pub chart_rel_id: Option<String>,
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
            doc_pr_id: 1,
            extent_cx: Emu(width_emu),
            extent_cy: Emu(height_emu),
            embed_id: embed_id.to_string(),
            link_id: None,
            chart_rel_id: None,
            description: None,
            name: None,
            raw_xml: None,
        }
    }

    /// Create an inline drawing whose payload is a native Word chart.
    pub fn new_chart(chart_rel_id: &str, width_emu: i64, height_emu: i64) -> Self {
        CT_Inline {
            doc_pr_id: 1,
            extent_cx: Emu(width_emu),
            extent_cy: Emu(height_emu),
            embed_id: String::new(),
            link_id: None,
            chart_rel_id: Some(chart_rel_id.to_owned()),
            description: None,
            name: Some("Chart".to_owned()),
            raw_xml: None,
        }
    }

    pub fn from_xml(reader: &mut NsReader<&[u8]>) -> Result<Self> {
        let mut cx = Emu(0);
        let mut cy = Emu(0);
        let mut description = None;
        let mut name = None;
        let mut doc_pr_id = None;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(ref e)) => {
                    let ename = e.name();
                    if matches_local_name(ename.as_ref(), b"extent") {
                        for attr in e.attributes() {
                            let attr = attr?;
                            let key = attr.key.as_ref();
                            let val = attr.decoded_and_normalized_value(
                                XmlVersion::Implicit1_0,
                                e.decoder(),
                            )?;
                            if key == b"cx" {
                                cx = Emu(val.parse()?);
                            } else if key == b"cy" {
                                cy = Emu(val.parse()?);
                            }
                        }
                    } else if canonical_wp_element(reader, e, b"docPr") {
                        for attr in e.attributes() {
                            let attr = attr?;
                            let key = attr.key.as_ref();
                            let val = attr.decoded_and_normalized_value(
                                XmlVersion::Implicit1_0,
                                e.decoder(),
                            )?;
                            if key == b"descr" {
                                description = Some(val.to_string());
                            } else if key == b"name" {
                                name = Some(val.to_string());
                            } else if key == b"id" {
                                doc_pr_id = Some(val.parse()?);
                            }
                        }
                    }
                }
                Ok(Event::Start(ref e)) => {
                    let ename = e.name();
                    if matches_local_name(ename.as_ref(), b"blip") {
                        reader.read_to_end_into(ename, &mut Vec::new())?;
                    } else if canonical_wp_element(reader, e, b"docPr") {
                        for attr in e.attributes() {
                            let attr = attr?;
                            let key = attr.key.as_ref();
                            let val = attr.decoded_and_normalized_value(
                                XmlVersion::Implicit1_0,
                                e.decoder(),
                            )?;
                            if key == b"descr" {
                                description = Some(val.to_string());
                            } else if key == b"name" {
                                name = Some(val.to_string());
                            } else if key == b"id" {
                                doc_pr_id = Some(val.parse()?);
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
            doc_pr_id: doc_pr_id
                .ok_or_else(|| crate::OxmlError::MissingElement("wp:docPr/@id".to_owned()))?,
            extent_cx: cx,
            extent_cy: cy,
            embed_id: String::new(),
            link_id: None,
            chart_rel_id: None,
            description,
            name,
            raw_xml: None, // Will be set by CT_Drawing::from_xml
        })
    }

    pub fn to_xml<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        // If we have raw XML from parsing, use it for perfect round-trip
        if let Some(ref raw) = self.raw_xml {
            writer.write_indent()?;
            writer.get_mut().write_all(raw)?;
            return Ok(());
        }

        let payload = drawing_payload(&self.embed_id, self.chart_rel_id.as_deref())?;

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
        doc_pr.push_attribute(("id", buf.format(self.doc_pr_id)));
        doc_pr.push_attribute(("name", self.name.as_deref().unwrap_or("Picture")));
        if let Some(ref desc) = self.description {
            doc_pr.push_attribute(("descr", desc.as_str()));
        }
        writer.write_event(Event::Empty(doc_pr))?;

        // a:graphic
        write_graphic_element(
            writer,
            payload,
            self.extent_cx,
            self.extent_cy,
            self.name.as_deref(),
        )?;

        writer.write_event(Event::End(BytesEnd::new("wp:inline")))?;

        Ok(())
    }
}

#[derive(Clone, Copy)]
enum DrawingPayload<'a> {
    Picture(&'a str),
    Chart(&'a str),
}

fn drawing_payload<'a>(
    embed_id: &'a str,
    chart_rel_id: Option<&'a str>,
) -> Result<DrawingPayload<'a>> {
    match (
        !embed_id.is_empty(),
        chart_rel_id.filter(|id| !id.is_empty()),
    ) {
        (true, None) => Ok(DrawingPayload::Picture(embed_id)),
        (false, Some(id)) => Ok(DrawingPayload::Chart(id)),
        (true, Some(_)) => Err(crate::OxmlError::InvalidValue(
            "a Word drawing cannot contain both picture and chart relationships".to_owned(),
        )),
        (false, None) => Err(crate::OxmlError::InvalidValue(
            "a Word drawing requires exactly one picture or chart relationship".to_owned(),
        )),
    }
}

fn chart_relationship_id(xml: &[u8]) -> Result<Option<String>> {
    let mut reader = NsReader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut chart_graphic_data_depth = None;
    loop {
        let (namespace, event) = reader.read_resolved_event_into(&mut buffer)?;
        match event {
            Event::Start(ref element) => {
                if chart_graphic_data_depth.is_none()
                    && namespace_matches(&namespace, drawing_ns::A, b"a")
                    && matches_local_name(element.name().as_ref(), b"graphicData")
                    && graphic_data_uri_is_chart(element)?
                {
                    chart_graphic_data_depth = Some(depth);
                } else if chart_graphic_data_depth.is_some_and(|parent| depth == parent + 1)
                    && namespace_matches(&namespace, drawing_ns::C, b"c")
                    && matches_local_name(element.name().as_ref(), b"chart")
                {
                    return chart_element_relationship_id(&reader, element);
                }
                depth = depth.checked_add(1).ok_or_else(|| {
                    crate::OxmlError::InvalidValue("drawing XML depth overflow".to_owned())
                })?;
            }
            Event::Empty(ref element)
                if chart_graphic_data_depth.is_some_and(|parent| depth == parent + 1)
                    && namespace_matches(&namespace, drawing_ns::C, b"c")
                    && matches_local_name(element.name().as_ref(), b"chart") =>
            {
                return chart_element_relationship_id(&reader, element);
            }
            Event::End(ref element) => {
                depth = depth.saturating_sub(1);
                if chart_graphic_data_depth == Some(depth)
                    && namespace_matches(&namespace, drawing_ns::A, b"a")
                    && matches_local_name(element.name().as_ref(), b"graphicData")
                {
                    chart_graphic_data_depth = None;
                }
            }
            Event::Eof => return Ok(None),
            _ => {}
        }
        buffer.clear();
    }
}

fn image_relationship_ids(xml: &[u8]) -> Result<(String, Option<String>)> {
    let mut reader = NsReader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut graphic_depth = None;
    let mut graphic_data_depth = None;
    let mut picture_depth = None;
    let mut blip_fill_depth = None;
    let mut picture_count = 0usize;
    let mut blip_count = 0usize;
    let mut relationship_ids = None;
    let mut ambiguous = false;
    loop {
        let (namespace, event) = reader.read_resolved_event_into(&mut buffer)?;
        match event {
            Event::Start(ref element) => {
                if depth == 1
                    && namespace_matches(&namespace, drawing_ns::A, b"a")
                    && matches_local_name(element.name().as_ref(), b"graphic")
                {
                    graphic_depth = Some(depth);
                } else if graphic_depth.is_some_and(|parent| depth == parent + 1)
                    && namespace_matches(&namespace, drawing_ns::A, b"a")
                    && matches_local_name(element.name().as_ref(), b"graphicData")
                    && graphic_data_uri_is_picture(element)?
                {
                    graphic_data_depth = Some(depth);
                } else if graphic_data_depth.is_some_and(|parent| depth == parent + 1)
                    && namespace_matches(&namespace, drawing_ns::PIC, b"pic")
                    && matches_local_name(element.name().as_ref(), b"pic")
                {
                    picture_count += 1;
                    ambiguous |= picture_count > 1;
                    picture_depth = Some(depth);
                } else if picture_depth.is_some_and(|parent| depth == parent + 1)
                    && namespace_matches(&namespace, drawing_ns::PIC, b"pic")
                    && matches_local_name(element.name().as_ref(), b"blipFill")
                {
                    blip_fill_depth = Some(depth);
                } else if blip_fill_depth.is_some_and(|parent| depth == parent + 1)
                    && namespace_matches(&namespace, drawing_ns::A, b"a")
                    && matches_local_name(element.name().as_ref(), b"blip")
                {
                    blip_count += 1;
                    ambiguous |= blip_count > 1;
                    match blip_relationship_ids(&reader, element)? {
                        Some(ids) if blip_count == 1 => relationship_ids = Some(ids),
                        Some(_) | None => ambiguous = true,
                    }
                }
                depth = depth.checked_add(1).ok_or_else(|| {
                    crate::OxmlError::InvalidValue("drawing XML depth overflow".to_owned())
                })?;
            }
            Event::Empty(ref element) => {
                if graphic_data_depth.is_some_and(|parent| depth == parent + 1)
                    && namespace_matches(&namespace, drawing_ns::PIC, b"pic")
                    && matches_local_name(element.name().as_ref(), b"pic")
                {
                    picture_count += 1;
                    ambiguous |= picture_count > 1;
                } else if blip_fill_depth.is_some_and(|parent| depth == parent + 1)
                    && namespace_matches(&namespace, drawing_ns::A, b"a")
                    && matches_local_name(element.name().as_ref(), b"blip")
                {
                    blip_count += 1;
                    ambiguous |= blip_count > 1;
                    match blip_relationship_ids(&reader, element)? {
                        Some(ids) if blip_count == 1 => relationship_ids = Some(ids),
                        Some(_) | None => ambiguous = true,
                    }
                }
            }
            Event::End(ref element) => {
                depth = depth.saturating_sub(1);
                if blip_fill_depth == Some(depth)
                    && namespace_matches(&namespace, drawing_ns::PIC, b"pic")
                    && matches_local_name(element.name().as_ref(), b"blipFill")
                {
                    blip_fill_depth = None;
                }
                if picture_depth == Some(depth)
                    && namespace_matches(&namespace, drawing_ns::PIC, b"pic")
                    && matches_local_name(element.name().as_ref(), b"pic")
                {
                    picture_depth = None;
                }
                if graphic_data_depth == Some(depth)
                    && namespace_matches(&namespace, drawing_ns::A, b"a")
                    && matches_local_name(element.name().as_ref(), b"graphicData")
                {
                    graphic_data_depth = None;
                }
                if graphic_depth == Some(depth)
                    && namespace_matches(&namespace, drawing_ns::A, b"a")
                    && matches_local_name(element.name().as_ref(), b"graphic")
                {
                    graphic_depth = None;
                }
            }
            Event::Eof => {
                return Ok(if ambiguous || picture_count != 1 || blip_count != 1 {
                    (String::new(), None)
                } else {
                    relationship_ids.unwrap_or_default()
                });
            }
            _ => {}
        }
        buffer.clear();
    }
}

fn blip_relationship_ids(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<Option<(String, Option<String>)>> {
    let mut embed_id = None;
    let mut link_id = None;
    for attribute in element.attributes() {
        let attribute = attribute?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if !namespace_matches(&namespace, drawing_ns::R, b"r") {
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())?
            .into_owned();
        let duplicate = match local.as_ref() {
            b"embed" => embed_id.replace(value).is_some(),
            b"link" => link_id.replace(value).is_some(),
            _ => false,
        };
        if duplicate {
            return Ok(None);
        }
    }
    Ok(Some((embed_id.unwrap_or_default(), link_id)))
}

fn graphic_data_uri_is_chart(element: &BytesStart<'_>) -> Result<bool> {
    for attribute in element.attributes() {
        let attribute = attribute?;
        if attribute.key.as_ref() == b"uri" {
            return Ok(attribute.value.as_ref() == drawing_ns::C.as_bytes());
        }
    }
    Ok(false)
}

fn graphic_data_uri_is_picture(element: &BytesStart<'_>) -> Result<bool> {
    for attribute in element.attributes() {
        let attribute = attribute?;
        if attribute.key.as_ref() == b"uri" {
            return Ok(attribute.value.as_ref() == drawing_ns::PIC.as_bytes());
        }
    }
    Ok(false)
}

fn chart_element_relationship_id(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<Option<String>> {
    for attribute in element.attributes() {
        let attribute = attribute?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if namespace_matches(&namespace, drawing_ns::R, b"r") && local.as_ref() == b"id" {
            return Ok(Some(
                attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())?
                    .into_owned(),
            ));
        }
    }
    Ok(None)
}

fn namespace_matches(namespace: &ResolveResult<'_>, expected: &str, _conventional: &[u8]) -> bool {
    match namespace {
        ResolveResult::Bound(Namespace(uri)) => *uri == expected.as_bytes(),
        ResolveResult::Unknown(_) => false,
        ResolveResult::Unbound => false,
    }
}

fn canonical_wp_element(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    expected_local: &[u8],
) -> bool {
    let (namespace, local) = reader.resolver().resolve_element(element.name());
    namespace_matches(&namespace, drawing_ns::WP, b"wp") && local.as_ref() == expected_local
}

fn capture_ns_element(reader: &mut NsReader<&[u8]>, start: &BytesStart<'_>) -> Result<Vec<u8>> {
    let mut writer = Writer::new(Vec::new());
    writer.write_event(Event::Start(start.to_owned()))?;
    let tag_name = start.name().as_ref().to_vec();
    let mut depth = 1u32;
    let mut buffer = Vec::new();
    loop {
        let event = reader.read_event_into(&mut buffer)?;
        match event {
            Event::Start(ref element) => {
                if element.name().as_ref() == tag_name {
                    depth += 1;
                }
                writer.write_event(Event::Start(element.to_owned()))?;
            }
            Event::End(ref element) => {
                if element.name().as_ref() == tag_name {
                    depth -= 1;
                }
                writer.write_event(Event::End(element.to_owned()))?;
                if depth == 0 {
                    break;
                }
            }
            Event::Empty(ref element) => {
                writer.write_event(Event::Empty(element.to_owned()))?;
            }
            Event::Text(ref text) => {
                writer.write_event(Event::Text(text.to_owned().into_owned()))?;
            }
            Event::CData(ref text) => {
                writer.write_event(Event::CData(text.to_owned().into_owned()))?;
            }
            Event::Comment(ref text) => {
                writer.write_event(Event::Comment(text.to_owned().into_owned()))?;
            }
            Event::PI(ref text) => {
                writer.write_event(Event::PI(text.to_owned().into_owned()))?;
            }
            Event::Decl(ref declaration) => {
                writer.write_event(Event::Decl(declaration.to_owned().into_owned()))?;
            }
            Event::DocType(ref declaration) => {
                writer.write_event(Event::DocType(declaration.to_owned().into_owned()))?;
            }
            Event::GeneralRef(ref reference) => {
                writer.write_event(Event::GeneralRef(reference.to_owned().into_owned()))?;
            }
            Event::Eof => break,
        }
        buffer.clear();
    }
    Ok(writer.into_inner())
}

/// Write the `a:graphic` payload shared by inline and anchored drawings.
fn write_graphic_element<W: std::io::Write>(
    writer: &mut Writer<W>,
    payload: DrawingPayload<'_>,
    cx: Emu,
    cy: Emu,
    name: Option<&str>,
) -> Result<()> {
    let mut buf = itoa::Buffer::new();
    let mut graphic = BytesStart::new("a:graphic");
    graphic.push_attribute(("xmlns:a", drawing_ns::A));
    writer.write_event(Event::Start(graphic))?;

    let mut gd = BytesStart::new("a:graphicData");
    gd.push_attribute((
        "uri",
        match payload {
            DrawingPayload::Picture(_) => drawing_ns::PIC,
            DrawingPayload::Chart(_) => drawing_ns::C,
        },
    ));
    writer.write_event(Event::Start(gd))?;

    if let DrawingPayload::Chart(chart_rel_id) = payload {
        let mut chart = BytesStart::new("c:chart");
        chart.push_attribute(("xmlns:c", drawing_ns::C));
        chart.push_attribute(("xmlns:r", drawing_ns::R));
        chart.push_attribute(("r:id", chart_rel_id));
        writer.write_event(Event::Empty(chart))?;
        writer.write_event(Event::End(BytesEnd::new("a:graphicData")))?;
        writer.write_event(Event::End(BytesEnd::new("a:graphic")))?;
        return Ok(());
    }

    let DrawingPayload::Picture(embed_id) = payload else {
        unreachable!();
    };

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
        Self::from_xml_with_prefixes(reader, &[], &[])
    }

    /// Parse a `w:drawing` whose start tag the caller has consumed.
    ///
    /// `owner_bindings` are the namespace declarations made on that start tag.
    /// The preserved inline or anchor subtree is written back without its
    /// `w:drawing` owner, so it keeps the declarations it uses.
    pub(crate) fn from_xml_with_prefixes(
        reader: &mut Reader<&[u8]>,
        prefixes: &[String],
        owner_bindings: &[(String, String)],
    ) -> Result<Self> {
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
                        let scoped_raw = crate::text::raw_with_external_bindings(
                            &raw,
                            &namespace_bindings(prefixes),
                        )?;
                        let mut re_reader = NsReader::from_reader(scoped_raw.as_slice());
                        re_reader.config_mut().trim_text(true);
                        // Skip to the <wp:inline> start
                        let mut rbuf = Vec::new();
                        loop {
                            match re_reader.read_event_into(&mut rbuf) {
                                Ok(Event::Start(ref ie))
                                    if matches_local_name(ie.name().as_ref(), b"inline") =>
                                {
                                    let mut inl = CT_Inline::from_xml(&mut re_reader)?;
                                    inl.doc_pr_id = non_visual_drawing_id(&scoped_raw)?;
                                    (inl.embed_id, inl.link_id) =
                                        image_relationship_ids(&scoped_raw)?;
                                    inl.chart_rel_id = chart_relationship_id(&scoped_raw)?;
                                    inl.raw_xml = Some(crate::text::raw_with_external_bindings(
                                        &raw,
                                        owner_bindings,
                                    )?);
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
                        let scoped_raw = crate::text::raw_with_external_bindings(
                            &raw,
                            &namespace_bindings(prefixes),
                        )?;
                        let mut re_reader = NsReader::from_reader(scoped_raw.as_slice());
                        re_reader.config_mut().trim_text(true);
                        let mut rbuf = Vec::new();
                        loop {
                            match re_reader.read_event_into(&mut rbuf) {
                                Ok(Event::Start(ref ie))
                                    if matches_local_name(ie.name().as_ref(), b"anchor") =>
                                {
                                    let mut anc = CT_Anchor::from_xml(&mut re_reader, ie)?;
                                    anc.doc_pr_id = non_visual_drawing_id(&scoped_raw)?;
                                    (anc.embed_id, anc.link_id) =
                                        image_relationship_ids(&scoped_raw)?;
                                    anc.chart_rel_id = chart_relationship_id(&scoped_raw)?;
                                    anc.raw_xml = Some(crate::text::raw_with_external_bindings(
                                        &raw,
                                        owner_bindings,
                                    )?);
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

fn non_visual_drawing_id(xml: &[u8]) -> Result<u32> {
    let mut reader = NsReader::from_reader(xml);
    let mut buffer = Vec::new();
    loop {
        let (namespace, event) = reader.read_resolved_event_into(&mut buffer)?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element)
                if namespace_matches(&namespace, drawing_ns::WP, b"wp")
                    && element.local_name().as_ref() == b"docPr" =>
            {
                for attribute in element.attributes() {
                    let attribute = attribute?;
                    let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
                    if matches!(namespace, ResolveResult::Unbound) && local.as_ref() == b"id" {
                        return Ok(attribute
                            .decoded_and_normalized_value(
                                XmlVersion::Implicit1_0,
                                element.decoder(),
                            )?
                            .parse()?);
                    }
                }
                return Err(crate::OxmlError::InvalidValue(
                    "wp:docPr requires an unqualified id attribute".to_owned(),
                ));
            }
            Event::Eof => {
                return Err(crate::OxmlError::MissingElement("wp:docPr".to_owned()));
            }
            _ => {}
        }
        buffer.clear();
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn parse_drawing(xml: &str) -> CT_Drawing {
        let mut valid_xml = xml.to_owned();
        if !valid_xml.contains("<wp:docPr") {
            let owner = ["<wp:inline", "<wp:anchor"]
                .into_iter()
                .find_map(|start| valid_xml.find(start));
            if let Some(owner) = owner {
                let insertion = valid_xml[owner..].find('>').unwrap() + owner + 1;
                valid_xml.insert_str(
                    insertion,
                    &format!(r#"<wp:docPr xmlns:wp="{}" id="1"/>"#, drawing_ns::WP),
                );
            }
        }
        let mut reader = Reader::from_str(&valid_xml);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        let prefixes = loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if matches_local_name(e.name().as_ref(), b"drawing") => {
                    break crate::numbering::word_prefixes_at(e, &[]).unwrap();
                }
                _ => {}
            }
            buf.clear();
        };
        CT_Drawing::from_xml_with_prefixes(&mut reader, &prefixes, &[]).unwrap()
    }

    fn parse_inline_direct(xml: &str) -> CT_Inline {
        let mut reader = NsReader::from_str(xml);
        let mut buffer = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer).unwrap() {
                Event::Start(ref element)
                    if matches_local_name(element.name().as_ref(), b"inline") =>
                {
                    return CT_Inline::from_xml(&mut reader).unwrap();
                }
                Event::Eof => panic!("inline start not found"),
                _ => {}
            }
            buffer.clear();
        }
    }

    fn parse_anchor_direct(xml: &str) -> CT_Anchor {
        let mut reader = NsReader::from_str(xml);
        let mut buffer = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer).unwrap() {
                Event::Start(ref element)
                    if matches_local_name(element.name().as_ref(), b"anchor") =>
                {
                    return CT_Anchor::from_xml(&mut reader, element).unwrap();
                }
                Event::Eof => panic!("anchor start not found"),
                _ => {}
            }
            buffer.clear();
        }
    }

    #[test]
    fn namespaces_declared_on_the_drawing_element_travel_with_its_raw_xml() {
        use crate::document::CT_Document;
        use crate::namespace::W_NS;
        use crate::text::RunContent;

        const DECLARATIONS: &str = r#"xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#;
        let graphic = r#"<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rId10"/></pic:blipFill></pic:pic></a:graphicData></a:graphic>"#;
        let inline = format!(
            r#"<wp:inline><wp:extent cx="1" cy="1"/><wp:docPr id="1" name="Inline"/>{graphic}</wp:inline>"#
        );
        let anchor = format!(
            r#"<wp:anchor behindDoc="0"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="1" cy="1"/><wp:wrapNone/><wp:docPr id="2" name="Anchor"/>{graphic}</wp:anchor>"#
        );
        let drawing_of = |document: &CT_Document| {
            let paragraph = document.body.paragraphs().next().expect("paragraph");
            let Some(RunContent::Drawing(drawing)) = paragraph.runs[0].content.first() else {
                panic!("drawing run");
            };
            drawing.clone()
        };
        let parts = |drawing: &CT_Drawing| match (&drawing.inline, &drawing.anchor) {
            (Some(inline), _) => (inline.embed_id.clone(), inline.raw_xml.clone().unwrap()),
            (_, Some(anchor)) => (anchor.embed_id.clone(), anchor.raw_xml.clone().unwrap()),
            _ => panic!("inline or anchor"),
        };

        for subtree in [&inline, &anchor] {
            let xml = format!(
                r#"<w:document xmlns:w="{W_NS}"><w:body><w:p><w:r><w:drawing {DECLARATIONS}>{subtree}</w:drawing></w:r></w:p></w:body></w:document>"#
            );
            let document = CT_Document::from_xml(xml.as_bytes()).unwrap();
            let saved = document.to_xml().unwrap();
            let reopened = CT_Document::from_xml(&saved).unwrap();
            for parsed in [&document, &reopened] {
                let (embed_id, raw) = parts(&drawing_of(parsed));
                assert_eq!(embed_id, "rId10", "{}", String::from_utf8_lossy(&saved));
                let raw = String::from_utf8(raw).unwrap();
                assert!(raw.contains("xmlns:a="), "{raw}");
                assert!(raw.contains("xmlns:pic="), "{raw}");
            }

            // Declarations on the part root leave the preserved subtree untouched.
            let xml = format!(
                r#"<w:document xmlns:w="{W_NS}" {DECLARATIONS}><w:body><w:p><w:r><w:drawing>{subtree}</w:drawing></w:r></w:p></w:body></w:document>"#
            );
            let document = CT_Document::from_xml(xml.as_bytes()).unwrap();
            let (embed_id, raw) = parts(&drawing_of(&document));
            assert_eq!(embed_id, "rId10");
            assert!(!String::from_utf8(raw).unwrap().contains("xmlns"));
        }
    }

    #[test]
    fn round_trip_inline_drawing() {
        let inline = CT_Inline {
            doc_pr_id: 1,
            extent_cx: Emu(914400), // 1 inch
            extent_cy: Emu(457200), // 0.5 inch
            embed_id: "rId5".to_string(),
            link_id: None,
            chart_rel_id: None,
            description: Some("A test image".to_string()),
            name: Some("TestPic".to_string()),
            raw_xml: None,
        };

        let drawing = CT_Drawing::inline(inline);

        let mut output = Vec::new();
        let mut writer = Writer::new(&mut output);
        drawing.to_xml(&mut writer).unwrap();
        let xml = String::from_utf8(output).unwrap().replacen(
            "<w:drawing",
            &format!(
                r#"<w:drawing xmlns:r="{}" xmlns:wp="{}""#,
                drawing_ns::R,
                drawing_ns::WP
            ),
            1,
        );

        let parsed = parse_drawing(&xml);
        let inl = parsed.inline.unwrap();
        assert_eq!(inl.extent_cx, Emu(914400));
        assert_eq!(inl.extent_cy, Emu(457200));
        assert_eq!(inl.embed_id, "rId5");
    }

    #[test]
    fn linked_image_relationships_are_retained_for_inline_and_anchor_drawings() {
        let inline = parse_drawing(concat!(
            r#"<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" "#,
            r#"xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" "#,
            r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" "#,
            r#"xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" "#,
            r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
            r#"<wp:inline><wp:extent cx="10" cy="20"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:link="rId7"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>"#,
        ));
        assert_eq!(inline.inline.unwrap().link_id.as_deref(), Some("rId7"));

        let anchor = parse_drawing(concat!(
            r#"<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" "#,
            r#"xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" "#,
            r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" "#,
            r#"xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" "#,
            r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
            r#"<wp:anchor><wp:extent cx="10" cy="20"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:link="rId8"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing>"#,
        ));
        assert_eq!(anchor.anchor.unwrap().link_id.as_deref(), Some("rId8"));
    }

    #[test]
    fn linked_image_relationships_require_the_office_relationship_namespace() {
        let foreign = parse_drawing(concat!(
            r#"<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" "#,
            r#"xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" "#,
            r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" "#,
            r#"xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">"#,
            r#"<wp:inline><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip xmlns:x="urn:producer" x:link="rIdBad"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>"#,
        ));
        assert!(foreign.inline.unwrap().link_id.is_none());

        let aliased = parse_drawing(concat!(
            r#"<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" "#,
            r#"xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" "#,
            r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" "#,
            r#"xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">"#,
            r#"<wp:inline><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip xmlns:q="http://schemas.openxmlformats.org/officeDocument/2006/relationships" q:link="rId9"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>"#,
        ));
        assert_eq!(aliased.inline.unwrap().link_id.as_deref(), Some("rId9"));
    }

    #[test]
    fn public_nested_drawing_parsers_do_not_type_unresolved_relationship_attributes() {
        let inline = parse_inline_direct(
            r#"<wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:docPr id="9"/><a:blip xmlns:a="urn:a" xmlns:ext="urn:producer" ext:embed="rIdBad" ext:link="rIdBad"/></wp:inline>"#,
        );
        assert_eq!(inline.doc_pr_id, 9);
        assert!(inline.embed_id.is_empty());
        assert!(inline.link_id.is_none());

        let anchor = parse_anchor_direct(
            r#"<wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:docPr id="10"/><a:blip xmlns:a="urn:a" xmlns:ext="urn:producer" ext:embed="rIdBad" ext:link="rIdBad"/></wp:anchor>"#,
        );
        assert_eq!(anchor.doc_pr_id, 10);
        assert!(anchor.embed_id.is_empty());
        assert!(anchor.link_id.is_none());
    }

    #[test]
    fn public_nested_drawing_parsers_ignore_foreign_doc_pr_decoys() {
        let inline = parse_inline_direct(concat!(
            r#"<root xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:x="urn:producer">"#,
            r#"<wp:inline><x:docPr id="99"/><wp:docPr id="9"/></wp:inline></root>"#,
        ));
        assert_eq!(inline.doc_pr_id, 9);

        let anchor = parse_anchor_direct(concat!(
            r#"<root xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:x="urn:producer">"#,
            r#"<wp:anchor><x:docPr id="99"/><wp:docPr id="10"/></wp:anchor></root>"#,
        ));
        assert_eq!(anchor.doc_pr_id, 10);
    }

    #[test]
    fn numeric_character_references_decode_in_every_doc_pr_parser() {
        let inline = parse_inline_direct(
            r#"<wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:docPr id="&#57;"/></wp:inline>"#,
        );
        assert_eq!(inline.doc_pr_id, 9);

        let anchor = parse_anchor_direct(
            r#"<wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:docPr id="&#49;0"/></wp:anchor>"#,
        );
        assert_eq!(anchor.doc_pr_id, 10);

        let wrapped = parse_drawing(concat!(
            r#"<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" "#,
            r#"xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">"#,
            r#"<wp:inline><wp:docPr id="1&#49;"/></wp:inline></w:drawing>"#,
        ));
        assert_eq!(wrapped.inline.unwrap().doc_pr_id, 11);
    }

    #[test]
    fn relationship_character_references_decode_for_picture_and_chart_payloads() {
        let picture = parse_drawing(concat!(
            r#"<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" "#,
            r#"xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" "#,
            r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" "#,
            r#"xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" "#,
            r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
            r#"<wp:inline><wp:docPr id="1"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="r&#73;d7" r:link="rId&#56;"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>"#,
        ));
        let picture = picture.inline.expect("inline picture");
        assert_eq!(picture.embed_id, "rId7");
        assert_eq!(picture.link_id.as_deref(), Some("rId8"));

        let chart = parse_drawing(concat!(
            r#"<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" "#,
            r#"xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" "#,
            r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" "#,
            r#"xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" "#,
            r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
            r#"<wp:inline><wp:docPr id="2"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId&#57;"/></a:graphicData></a:graphic></wp:inline></w:drawing>"#,
        ));
        let chart = chart.inline.expect("inline chart");
        assert!(chart.embed_id.is_empty());
        assert_eq!(chart.chart_rel_id.as_deref(), Some("rId9"));
    }

    #[test]
    fn undeclared_conventional_picture_prefixes_do_not_gain_namespace_semantics() {
        let drawing = parse_drawing(concat!(
            r#"<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" "#,
            r#"xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" "#,
            r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" "#,
            r#"xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">"#,
            r#"<wp:inline><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rIdBad" r:link="rIdBad"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>"#,
        ));
        let inline = drawing.inline.expect("inline drawing");
        assert!(inline.embed_id.is_empty());
        assert!(inline.link_id.is_none());

        let drawing = parse_drawing(concat!(
            r#"<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" "#,
            r#"xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" "#,
            r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
            r#"<wp:inline><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rIdBad"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>"#,
        ));
        assert!(drawing.inline.expect("inline drawing").embed_id.is_empty());

        let xml = concat!(
            r#"<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">"#,
            r#"<wp:inline><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rIdPublic"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>"#,
        );
        let mut reader = Reader::from_str(xml);
        let mut buffer = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer).unwrap() {
                Event::Start(element)
                    if matches_local_name(element.name().as_ref(), b"drawing") =>
                {
                    break;
                }
                Event::Eof => panic!("drawing start not found"),
                _ => {}
            }
            buffer.clear();
        }
        let error = CT_Drawing::from_xml(&mut reader).unwrap_err();
        assert!(matches!(error, crate::OxmlError::MissingElement(_)));
    }

    #[test]
    fn picture_relationships_require_the_schema_owned_graphic_payload() {
        let drawing = parse_drawing(concat!(
            r#"<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" "#,
            r#"xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" "#,
            r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" "#,
            r#"xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" "#,
            r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
            r#"<wp:inline><wp:extent cx="10" cy="20"/><a:extLst><a:ext><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rIdExtension"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></a:ext></a:extLst></wp:inline></w:drawing>"#,
        ));
        assert!(drawing.inline.expect("inline drawing").embed_id.is_empty());

        let drawing = parse_drawing(concat!(
            r#"<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" "#,
            r#"xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" "#,
            r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" "#,
            r#"xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" "#,
            r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
            r#"<wp:inline><a:graphic><a:graphicData uri="urn:producer"><pic:pic><pic:blipFill><a:blip r:embed="rIdWrongPayload"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>"#,
        ));
        assert!(drawing.inline.expect("inline drawing").embed_id.is_empty());
    }

    #[test]
    fn ambiguous_picture_payloads_do_not_merge_relationship_facts() {
        let two_pictures = parse_drawing(concat!(
            r#"<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" "#,
            r#"xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" "#,
            r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" "#,
            r#"xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" "#,
            r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
            r#"<wp:inline><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rId1"/></pic:blipFill></pic:pic><pic:pic><pic:blipFill><a:blip r:link="rId2"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>"#,
        ));
        let inline = two_pictures.inline.expect("inline drawing");
        assert!(inline.embed_id.is_empty());
        assert!(inline.link_id.is_none());

        let two_blips = parse_drawing(concat!(
            r#"<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" "#,
            r#"xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" "#,
            r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" "#,
            r#"xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" "#,
            r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
            r#"<wp:inline><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rId1"/><a:blip r:link="rId2"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>"#,
        ));
        let inline = two_blips.inline.expect("inline drawing");
        assert!(inline.embed_id.is_empty());
        assert!(inline.link_id.is_none());
    }

    #[test]
    fn shape_bitmap_fills_do_not_become_picture_relationships() {
        let drawing = parse_drawing(concat!(
            r#"<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" "#,
            r#"xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" "#,
            r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" "#,
            r#"xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" "#,
            r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
            r#"<wp:anchor><wps:wsp><wps:spPr><a:blipFill><a:blip r:embed="rIdFill"/></a:blipFill></wps:spPr></wps:wsp></wp:anchor></w:drawing>"#,
        ));
        let anchor = drawing.anchor.expect("shape anchor");
        assert!(anchor.shape.is_some());
        assert!(anchor.embed_id.is_empty());
        assert!(anchor.link_id.is_none());
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
        let xml = String::from_utf8(output).unwrap().replacen(
            "<w:drawing",
            &format!(
                r#"<w:drawing xmlns:r="{}" xmlns:wp="{}""#,
                drawing_ns::R,
                drawing_ns::WP
            ),
            1,
        );

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

    // F-X015, wrap and alignment model.

    const W_NS_URI: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const WP_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

    fn parse_anchor(inner: &str) -> CT_Anchor {
        let xml = format!(
            r#"<w:drawing xmlns:w="{W_NS_URI}" xmlns:wp="{WP_NS}"><wp:anchor {inner}</wp:anchor></w:drawing>"#
        );
        let drawing = parse_drawing(&xml);
        drawing.anchor.expect("an anchor")
    }

    const ANCHOR_TAIL: &str = r#"<wp:positionH relativeFrom="margin"><wp:align>right</wp:align></wp:positionH>
        <wp:positionV relativeFrom="paragraph"><wp:align>bottom</wp:align></wp:positionV>
        <wp:extent cx="914400" cy="457200"/>"#;

    #[test]
    fn every_wrap_element_parses_to_its_own_mode() {
        let cases = [
            ("wrapNone", WrapType::None),
            ("wrapSquare", WrapType::Square),
            ("wrapTopAndBottom", WrapType::TopAndBottom),
            ("wrapTight", WrapType::Tight),
            ("wrapThrough", WrapType::Through),
        ];

        for (element, expected) in cases {
            // Empty spelling.
            let anchor = parse_anchor(&format!(
                r#"distT="0" distB="0" distL="0" distR="0">{ANCHOR_TAIL}<wp:{element}/>"#
            ));
            assert_eq!(anchor.wrap, expected, "{element} as an empty element");

            // Expanded spelling, with a child subtree we do not model.
            let anchor = parse_anchor(&format!(
                r#"distT="0" distB="0" distL="0" distR="0">{ANCHOR_TAIL}<wp:{element}><wp:wrapPolygon/></wp:{element}>"#
            ));
            assert_eq!(anchor.wrap, expected, "{element} as an expanded element");
        }
    }

    #[test]
    fn anchor_alignments_and_distances_are_read() {
        let anchor = parse_anchor(&format!(
            r#"distT="10" distB="20" distL="30" distR="40">{ANCHOR_TAIL}<wp:wrapSquare/>"#
        ));

        assert_eq!(anchor.dist_t, Emu(10));
        assert_eq!(anchor.dist_b, Emu(20));
        assert_eq!(anchor.dist_l, Emu(30));
        assert_eq!(anchor.dist_r, Emu(40));
        assert_eq!(anchor.pos_h_align, Some(AnchorAlignH::Right));
        assert_eq!(anchor.pos_v_align, Some(AnchorAlignV::Bottom));
    }

    #[test]
    fn an_unknown_alignment_reads_as_no_alignment() {
        // Falling back to the offset beats inventing a position.
        let anchor = parse_anchor(
            r#"distT="0" distB="0" distL="0" distR="0"><wp:positionH relativeFrom="margin"><wp:align>sideways</wp:align></wp:positionH>
            <wp:positionV relativeFrom="paragraph"><wp:posOffset>5000</wp:posOffset></wp:positionV>
            <wp:extent cx="914400" cy="457200"/><wp:wrapSquare/>"#,
        );

        assert_eq!(anchor.pos_h_align, None);
        assert_eq!(anchor.pos_v_align, None);
        assert_eq!(anchor.pos_v_offset, Emu(5000));
    }

    #[test]
    fn an_anchor_round_trips_its_wrap_distances_and_alignments() {
        for wrap in [
            WrapType::None,
            WrapType::Square,
            WrapType::TopAndBottom,
            WrapType::Tight,
            WrapType::Through,
        ] {
            let mut built = CT_Anchor::background("rId7", 914400, 457200);
            built.wrap = wrap;
            built.dist_t = Emu(11);
            built.dist_b = Emu(22);
            built.dist_l = Emu(33);
            built.dist_r = Emu(44);
            built.pos_h_align = Some(AnchorAlignH::Center);
            built.pos_v_align = Some(AnchorAlignV::Top);

            let mut writer = Writer::new(Vec::new());
            built.to_xml(&mut writer).expect("serialises");
            let bytes = writer.into_inner();
            let xml = format!(
                r#"<w:drawing xmlns:w="{W_NS_URI}" xmlns:wp="{WP_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">{}</w:drawing>"#,
                String::from_utf8(bytes).expect("utf8")
            );

            let parsed = parse_drawing(&xml).anchor.expect("an anchor");

            assert_eq!(parsed.wrap, wrap, "wrap survives");
            assert_eq!(parsed.dist_t, Emu(11));
            assert_eq!(parsed.dist_b, Emu(22));
            assert_eq!(parsed.dist_l, Emu(33));
            assert_eq!(parsed.dist_r, Emu(44));
            assert_eq!(parsed.pos_h_align, Some(AnchorAlignH::Center));
            assert_eq!(parsed.pos_v_align, Some(AnchorAlignV::Top));
        }
    }

    #[test]
    fn a_parsed_anchor_re_emits_its_original_bytes() {
        // The capture path is what keeps an unmodelled subtree intact. This
        // story must not disturb it.
        let xml = format!(
            r#"<w:drawing xmlns:w="{W_NS_URI}" xmlns:wp="{WP_NS}"><wp:anchor distT="5" distB="6" distL="7" distR="8">{ANCHOR_TAIL}<wp:wrapSquare wrapText="bothSides"/><wp:unmodelled foo="bar"/></wp:anchor></w:drawing>"#
        );
        let drawing = parse_drawing(&xml);
        let anchor = drawing.anchor.as_ref().expect("an anchor");
        assert!(anchor.raw_xml.is_some(), "the anchor bytes are captured");

        let mut writer = Writer::new(Vec::new());
        anchor.to_xml(&mut writer).expect("serialises");
        let emitted = String::from_utf8(writer.into_inner()).expect("utf8");
        assert!(
            emitted.contains("wp:unmodelled"),
            "the unmodelled subtree survives, got {emitted}"
        );
        assert!(
            emitted.contains(r#"wrapText="bothSides""#),
            "attributes we do not model survive, got {emitted}"
        );
    }

    // F-157 start-feature stubs. These assertions describe the missing chart
    // relationship payload before the implementation adds its typed seam.

    #[test]
    fn word_chart_drawing_writes_schema_order_and_fixed_prefixes() {
        let inline = CT_Inline::new_chart("rId9", 4_572_000, 2_743_200);
        let mut writer = Writer::new(Vec::new());
        inline.to_xml(&mut writer).expect("serialises");
        let xml = String::from_utf8(writer.into_inner()).expect("utf8");

        assert!(
            xml.contains(r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId9"/>"#),
            "Word chart drawing payload is not implemented: {xml}"
        );
        let positions = [
            "<wp:extent",
            "<wp:docPr",
            "<a:graphic",
            "<a:graphicData",
            "<c:chart",
        ]
        .map(|tag| xml.find(tag).expect("required schema child"));
        assert!(positions.windows(2).all(|pair| pair[0] < pair[1]));

        let anchor = CT_Anchor::new_chart("rId10", 4_572_000, 2_743_200);
        let mut writer = Writer::new(Vec::new());
        anchor
            .to_xml(&mut writer)
            .expect("anchored chart serialises");
        let xml = String::from_utf8(writer.into_inner()).expect("utf8");
        assert!(xml.contains(r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId10"/>"#));

        let mut ambiguous = CT_Inline::new("rId1", 100, 100);
        ambiguous.chart_rel_id = Some("rId2".to_owned());
        let mut writer = Writer::new(Vec::new());
        assert!(ambiguous.to_xml(&mut writer).is_err());
        assert!(writer.into_inner().is_empty());

        let mut ambiguous = CT_Anchor::new_chart("rId2", 100, 100);
        ambiguous.embed_id = "rId1".to_owned();
        let mut writer = Writer::new(Vec::new());
        assert!(ambiguous.to_xml(&mut writer).is_err());
        assert!(writer.into_inner().is_empty());
    }

    #[test]
    fn opened_chart_drawing_preserves_unmodelled_xml() {
        let xml = format!(
            r#"<w:drawing xmlns:w="{W_NS_URI}" xmlns:wp="{WP_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><wp:inline><wp:extent cx="4572000" cy="2743200"/><wp:docPr id="1" name="Chart 1"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId4"/><x:producer xmlns:x="urn:producer" keep="yes"/></a:graphicData></a:graphic></wp:inline></w:drawing>"#
        );
        let drawing = parse_drawing(&xml);
        let inline = drawing.inline.expect("inline chart");

        assert_eq!(inline.chart_rel_id.as_deref(), Some("rId4"));
        let mut writer = Writer::new(Vec::new());
        inline.to_xml(&mut writer).expect("raw chart re-emits");
        assert_eq!(writer.into_inner(), inline.raw_xml.expect("captured bytes"));

        let picture_uri = xml.replace(drawing_ns::C, drawing_ns::PIC);
        let parsed = parse_drawing(&picture_uri).inline.expect("inline drawing");
        assert_eq!(
            parsed.chart_rel_id, None,
            "c:chart is typed only inside ChartML graphicData"
        );

        let aliased = format!(
            r#"<w:drawing xmlns:w="{W_NS_URI}" xmlns:wp="{WP_NS}" xmlns:a="{}"><wp:inline><wp:extent cx="1" cy="1"/><a:graphic><a:graphicData uri="{}"><q:chart xmlns:q="{}" xmlns:rel="{}" rel:id="rId8"/></a:graphicData></a:graphic></wp:inline></w:drawing>"#,
            drawing_ns::A,
            drawing_ns::C,
            drawing_ns::C,
            drawing_ns::R,
        );
        assert_eq!(
            parse_drawing(&aliased)
                .inline
                .expect("aliased inline chart")
                .chart_rel_id
                .as_deref(),
            Some("rId8")
        );

        let foreign = aliased
            .replace(
                &format!(r#"xmlns:q="{}""#, drawing_ns::C),
                r#"xmlns:q="urn:foreign""#,
            )
            .replace(
                &format!(r#"xmlns:rel="{}""#, drawing_ns::R),
                r#"xmlns:rel="urn:foreign-rel""#,
            );
        assert_eq!(
            parse_drawing(&foreign)
                .inline
                .expect("foreign inline drawing")
                .chart_rel_id,
            None
        );

        let empty_container = format!(
            r#"<w:drawing xmlns:w="{W_NS_URI}" xmlns:wp="{WP_NS}" xmlns:a="{}"><wp:inline><wp:extent cx="1" cy="1"/><a:graphic><a:graphicData uri="{}"/><q:chart xmlns:q="{}" xmlns:rel="{}" rel:id="rId8"/></a:graphic></wp:inline></w:drawing>"#,
            drawing_ns::A,
            drawing_ns::C,
            drawing_ns::C,
            drawing_ns::R,
        );
        assert_eq!(
            parse_drawing(&empty_container)
                .inline
                .expect("empty chart container")
                .chart_rel_id,
            None
        );

        let nested = aliased
            .replace("<q:chart", "<x:wrapper xmlns:x=\"urn:producer\"><q:chart")
            .replace("/></a:graphicData>", "/></x:wrapper></a:graphicData>");
        assert_eq!(
            parse_drawing(&nested)
                .inline
                .expect("nested chart lookalike")
                .chart_rel_id,
            None
        );
    }
}
