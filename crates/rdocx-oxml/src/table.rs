//! Table elements: `CT_Tbl`, `CT_Row`, `CT_Tc` and related types.

use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};

use crate::borders::CT_BorderEdge;
use crate::error::Result;
use crate::namespace::matches_local_name;
use crate::properties::{CT_Shd, get_val_attr};
#[cfg(test)]
use crate::shared::ST_Border;
use crate::shared::ST_Jc;
use crate::text::CT_P;
use crate::units::Twips;

// ---- Table border types ----

/// `CT_TblBorders` — Table-level borders.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct CT_TblBorders {
    pub top: Option<CT_BorderEdge>,
    pub bottom: Option<CT_BorderEdge>,
    pub left: Option<CT_BorderEdge>,
    pub right: Option<CT_BorderEdge>,
    pub inside_h: Option<CT_BorderEdge>,
    pub inside_v: Option<CT_BorderEdge>,
}

impl CT_TblBorders {
    pub fn from_xml(reader: &mut Reader<&[u8]>) -> Result<Self> {
        let mut borders = CT_TblBorders::default();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(ref e)) => {
                    let name = e.name();
                    let edge = CT_BorderEdge::from_xml_attrs(e)?;
                    if matches_local_name(name.as_ref(), b"top") {
                        borders.top = Some(edge);
                    } else if matches_local_name(name.as_ref(), b"bottom") {
                        borders.bottom = Some(edge);
                    } else if matches_local_name(name.as_ref(), b"left")
                        || matches_local_name(name.as_ref(), b"start")
                    {
                        borders.left = Some(edge);
                    } else if matches_local_name(name.as_ref(), b"right")
                        || matches_local_name(name.as_ref(), b"end")
                    {
                        borders.right = Some(edge);
                    } else if matches_local_name(name.as_ref(), b"insideH") {
                        borders.inside_h = Some(edge);
                    } else if matches_local_name(name.as_ref(), b"insideV") {
                        borders.inside_v = Some(edge);
                    }
                }
                Ok(Event::End(ref e))
                    if matches_local_name(e.name().as_ref(), b"tblBorders")
                        || matches_local_name(e.name().as_ref(), b"tcBorders") =>
                {
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(borders)
    }

    pub fn to_xml<W: std::io::Write>(&self, writer: &mut Writer<W>, tag: &str) -> Result<()> {
        writer.write_event(Event::Start(BytesStart::new(tag)))?;
        if let Some(ref e) = self.top {
            e.to_xml(writer, "w:top")?;
        }
        if let Some(ref e) = self.left {
            e.to_xml(writer, "w:left")?;
        }
        if let Some(ref e) = self.bottom {
            e.to_xml(writer, "w:bottom")?;
        }
        if let Some(ref e) = self.right {
            e.to_xml(writer, "w:right")?;
        }
        if let Some(ref e) = self.inside_h {
            e.to_xml(writer, "w:insideH")?;
        }
        if let Some(ref e) = self.inside_v {
            e.to_xml(writer, "w:insideV")?;
        }
        writer.write_event(Event::End(BytesEnd::new(tag)))?;
        Ok(())
    }

    pub fn is_empty(&self) -> bool {
        self.top.is_none()
            && self.bottom.is_none()
            && self.left.is_none()
            && self.right.is_none()
            && self.inside_h.is_none()
            && self.inside_v.is_none()
    }
}

/// Table cell margin (a single edge width).
#[derive(Debug, Clone, Default, PartialEq)]
pub struct CT_TblCellMar {
    pub top: Option<Twips>,
    pub bottom: Option<Twips>,
    pub left: Option<Twips>,
    pub right: Option<Twips>,
}

impl CT_TblCellMar {
    fn parse_edge(e: &BytesStart) -> Result<Option<Twips>> {
        for attr in e.attributes() {
            let attr = attr?;
            if matches_local_name(attr.key.as_ref(), b"w") {
                let val: i32 = std::str::from_utf8(&attr.value)?.parse()?;
                return Ok(Some(Twips(val)));
            }
        }
        Ok(None)
    }

    pub fn from_xml(reader: &mut Reader<&[u8]>) -> Result<Self> {
        let mut mar = CT_TblCellMar::default();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(ref e)) => {
                    let name = e.name();
                    if matches_local_name(name.as_ref(), b"top") {
                        mar.top = Self::parse_edge(e)?;
                    } else if matches_local_name(name.as_ref(), b"bottom") {
                        mar.bottom = Self::parse_edge(e)?;
                    } else if matches_local_name(name.as_ref(), b"left")
                        || matches_local_name(name.as_ref(), b"start")
                    {
                        mar.left = Self::parse_edge(e)?;
                    } else if matches_local_name(name.as_ref(), b"right")
                        || matches_local_name(name.as_ref(), b"end")
                    {
                        mar.right = Self::parse_edge(e)?;
                    }
                }
                Ok(Event::End(ref e)) if matches_local_name(e.name().as_ref(), b"tblCellMar") => {
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(mar)
    }

    pub fn to_xml<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_event(Event::Start(BytesStart::new("w:tblCellMar")))?;

        fn write_edge<W: std::io::Write>(
            writer: &mut Writer<W>,
            tag: &str,
            val: Twips,
        ) -> Result<()> {
            let mut buf = itoa::Buffer::new();
            let mut e = BytesStart::new(tag);
            e.push_attribute(("w:w", buf.format(val.0)));
            e.push_attribute(("w:type", "dxa"));
            writer.write_event(Event::Empty(e))?;
            Ok(())
        }

        if let Some(t) = self.top {
            write_edge(writer, "w:top", t)?;
        }
        if let Some(l) = self.left {
            write_edge(writer, "w:left", l)?;
        }
        if let Some(b) = self.bottom {
            write_edge(writer, "w:bottom", b)?;
        }
        if let Some(r) = self.right {
            write_edge(writer, "w:right", r)?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:tblCellMar")))?;
        Ok(())
    }
}

// ---- Table width ----

/// Table width specification.
#[derive(Debug, Clone, PartialEq)]
pub struct CT_TblWidth {
    /// Width value
    pub w: i32,
    /// Width type: "dxa" (twips), "pct" (50ths of a percent), "auto", "nil"
    pub width_type: String,
}

impl CT_TblWidth {
    pub fn dxa(twips: i32) -> Self {
        CT_TblWidth {
            w: twips,
            width_type: "dxa".to_string(),
        }
    }

    pub fn pct(fiftieths: i32) -> Self {
        CT_TblWidth {
            w: fiftieths,
            width_type: "pct".to_string(),
        }
    }

    pub fn auto() -> Self {
        CT_TblWidth {
            w: 0,
            width_type: "auto".to_string(),
        }
    }

    pub fn from_xml_attrs(e: &BytesStart) -> Result<Self> {
        let mut w = 0;
        let mut width_type = "dxa".to_string();

        for attr in e.attributes() {
            let attr = attr?;
            let key = attr.key.as_ref();
            let val = std::str::from_utf8(&attr.value)?;
            if matches_local_name(key, b"w") {
                w = val.parse().unwrap_or(0);
            } else if matches_local_name(key, b"type") {
                width_type = val.to_string();
            }
        }

        Ok(CT_TblWidth { w, width_type })
    }

    pub fn write_xml<W: std::io::Write>(&self, writer: &mut Writer<W>, tag: &str) -> Result<()> {
        let mut buf = itoa::Buffer::new();
        let mut e = BytesStart::new(tag);
        e.push_attribute(("w:w", buf.format(self.w)));
        e.push_attribute(("w:type", self.width_type.as_str()));
        writer.write_event(Event::Empty(e))?;
        Ok(())
    }
}

// ---- Table grid column ----

/// `CT_TblGridCol` — A column definition in the table grid.
#[derive(Debug, Clone, PartialEq)]
pub struct CT_TblGridCol {
    /// Column width in twips
    pub width: Twips,
}

// ---- Table properties ----

/// `CT_TblPr` — Table properties.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct CT_TblPr {
    /// Table style ID
    pub style_id: Option<String>,
    /// Table width
    pub width: Option<CT_TblWidth>,
    /// Table alignment
    pub jc: Option<ST_Jc>,
    /// Table borders
    pub borders: Option<CT_TblBorders>,
    /// Default cell margins
    pub cell_margin: Option<CT_TblCellMar>,
    /// Table layout: "fixed" or "autofit"
    pub layout: Option<String>,
    /// Table indent from left margin
    pub indent: Option<CT_TblWidth>,
    /// Table shading/background
    pub shading: Option<CT_Shd>,
    /// "Look" flags for conditional formatting (firstRow, lastRow, etc.)
    pub look: Option<String>,
}

#[allow(non_snake_case)]
impl CT_TblPr {
    pub fn from_xml(reader: &mut Reader<&[u8]>) -> Result<Self> {
        let mut pr = CT_TblPr::default();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(ref e)) => {
                    let name = e.name();
                    if matches_local_name(name.as_ref(), b"tblStyle") {
                        pr.style_id = get_val_attr(e)?;
                    } else if matches_local_name(name.as_ref(), b"tblW") {
                        pr.width = Some(CT_TblWidth::from_xml_attrs(e)?);
                    } else if matches_local_name(name.as_ref(), b"jc") {
                        if let Some(val) = get_val_attr(e)? {
                            pr.jc = Some(ST_Jc::from_str(&val)?);
                        }
                    } else if matches_local_name(name.as_ref(), b"tblLayout") {
                        if let Some(val) = get_val_attr(e)? {
                            pr.layout = Some(val);
                        }
                    } else if matches_local_name(name.as_ref(), b"tblInd") {
                        pr.indent = Some(CT_TblWidth::from_xml_attrs(e)?);
                    } else if matches_local_name(name.as_ref(), b"shd") {
                        pr.shading = Some(CT_Shd::from_xml_attrs(e)?);
                    } else if matches_local_name(name.as_ref(), b"tblLook")
                        && let Some(val) = get_val_attr(e)?
                    {
                        pr.look = Some(val);
                    }
                }
                Ok(Event::Start(ref e)) => {
                    let name = e.name();
                    if matches_local_name(name.as_ref(), b"tblBorders") {
                        pr.borders = Some(CT_TblBorders::from_xml(reader)?);
                    } else if matches_local_name(name.as_ref(), b"tblCellMar") {
                        pr.cell_margin = Some(CT_TblCellMar::from_xml(reader)?);
                    } else {
                        reader.read_to_end_into(name, &mut Vec::new())?;
                    }
                }
                Ok(Event::End(ref e)) if matches_local_name(e.name().as_ref(), b"tblPr") => {
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(pr)
    }

    pub fn to_xml<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_event(Event::Start(BytesStart::new("w:tblPr")))?;

        if let Some(ref style_id) = self.style_id {
            let mut e = BytesStart::new("w:tblStyle");
            e.push_attribute(("w:val", style_id.as_str()));
            writer.write_event(Event::Empty(e))?;
        }

        if let Some(ref width) = self.width {
            width.write_xml(writer, "w:tblW")?;
        }

        if let Some(jc) = self.jc {
            let mut e = BytesStart::new("w:jc");
            e.push_attribute(("w:val", jc.to_str()));
            writer.write_event(Event::Empty(e))?;
        }

        if let Some(ref indent) = self.indent {
            indent.write_xml(writer, "w:tblInd")?;
        }

        if let Some(ref borders) = self.borders
            && !borders.is_empty()
        {
            borders.to_xml(writer, "w:tblBorders")?;
        }

        if let Some(ref shd) = self.shading {
            shd.write_xml(writer, "w:shd")?;
        }

        if let Some(ref layout) = self.layout {
            let mut e = BytesStart::new("w:tblLayout");
            e.push_attribute(("w:type", layout.as_str()));
            writer.write_event(Event::Empty(e))?;
        }

        if let Some(ref cell_margin) = self.cell_margin {
            cell_margin.to_xml(writer)?;
        }

        if let Some(ref look) = self.look {
            let mut e = BytesStart::new("w:tblLook");
            e.push_attribute(("w:val", look.as_str()));
            writer.write_event(Event::Empty(e))?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:tblPr")))?;
        Ok(())
    }
}

// ---- Table grid ----

/// `CT_TblGrid` — Defines the column structure of a table.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct CT_TblGrid {
    pub columns: Vec<CT_TblGridCol>,
}

#[allow(non_snake_case)]
impl CT_TblGrid {
    pub fn from_xml(reader: &mut Reader<&[u8]>) -> Result<Self> {
        let mut columns = Vec::new();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(ref e)) => {
                    if matches_local_name(e.name().as_ref(), b"gridCol") {
                        let mut width = Twips(0);
                        for attr in e.attributes() {
                            let attr = attr?;
                            if matches_local_name(attr.key.as_ref(), b"w") {
                                width = Twips(std::str::from_utf8(&attr.value)?.parse()?);
                            }
                        }
                        columns.push(CT_TblGridCol { width });
                    }
                }
                Ok(Event::End(ref e)) if matches_local_name(e.name().as_ref(), b"tblGrid") => {
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(CT_TblGrid { columns })
    }

    pub fn to_xml<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut buf = itoa::Buffer::new();
        writer.write_event(Event::Start(BytesStart::new("w:tblGrid")))?;

        for col in &self.columns {
            let mut e = BytesStart::new("w:gridCol");
            e.push_attribute(("w:w", buf.format(col.width.0)));
            writer.write_event(Event::Empty(e))?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:tblGrid")))?;
        Ok(())
    }
}

// ---- Row properties ----

/// Vertical merge state for a cell.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VMerge {
    /// Start of a vertical merge group
    Restart,
    /// Continuation of the merge group above
    Continue,
}

/// `CT_TrPr` — Table row properties.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct CT_TrPr {
    /// Row height in twips
    pub height: Option<Twips>,
    /// Row height rule: "exact" or "atLeast"
    pub height_rule: Option<String>,
    /// Repeat as header row on each page
    pub header: Option<bool>,
    /// Row alignment
    pub jc: Option<ST_Jc>,
    /// Allow row to break across pages
    pub cant_split: Option<bool>,
}

#[allow(non_snake_case)]
impl CT_TrPr {
    pub fn from_xml(reader: &mut Reader<&[u8]>) -> Result<Self> {
        let mut pr = CT_TrPr::default();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(ref e)) => {
                    let name = e.name();
                    if matches_local_name(name.as_ref(), b"trHeight") {
                        for attr in e.attributes() {
                            let attr = attr?;
                            let key = attr.key.as_ref();
                            let val = std::str::from_utf8(&attr.value)?;
                            if matches_local_name(key, b"val") {
                                pr.height = Some(Twips(val.parse()?));
                            } else if matches_local_name(key, b"hRule") {
                                pr.height_rule = Some(val.to_string());
                            }
                        }
                    } else if matches_local_name(name.as_ref(), b"tblHeader") {
                        pr.header = Some(true);
                    } else if matches_local_name(name.as_ref(), b"jc") {
                        if let Some(val) = get_val_attr(e)? {
                            pr.jc = Some(ST_Jc::from_str(&val)?);
                        }
                    } else if matches_local_name(name.as_ref(), b"cantSplit") {
                        pr.cant_split = Some(true);
                    }
                }
                Ok(Event::End(ref e)) if matches_local_name(e.name().as_ref(), b"trPr") => {
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(pr)
    }

    pub fn to_xml<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        if self.is_empty() {
            return Ok(());
        }

        writer.write_event(Event::Start(BytesStart::new("w:trPr")))?;

        if let Some(ref cant_split) = self.cant_split
            && *cant_split
        {
            writer.write_event(Event::Empty(BytesStart::new("w:cantSplit")))?;
        }

        if let Some(height) = self.height {
            let mut buf = itoa::Buffer::new();
            let mut e = BytesStart::new("w:trHeight");
            e.push_attribute(("w:val", buf.format(height.0)));
            if let Some(ref rule) = self.height_rule {
                e.push_attribute(("w:hRule", rule.as_str()));
            }
            writer.write_event(Event::Empty(e))?;
        }

        if let Some(true) = self.header {
            writer.write_event(Event::Empty(BytesStart::new("w:tblHeader")))?;
        }

        if let Some(jc) = self.jc {
            let mut e = BytesStart::new("w:jc");
            e.push_attribute(("w:val", jc.to_str()));
            writer.write_event(Event::Empty(e))?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:trPr")))?;
        Ok(())
    }

    fn is_empty(&self) -> bool {
        self.height.is_none()
            && self.header.is_none()
            && self.jc.is_none()
            && self.cant_split.is_none()
    }
}

// ---- Cell properties ----

/// Vertical alignment within a cell.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ST_VerticalJc {
    Top,
    Center,
    Bottom,
}

impl ST_VerticalJc {
    pub fn from_str(s: &str) -> Self {
        match s {
            "center" => Self::Center,
            "bottom" => Self::Bottom,
            _ => Self::Top,
        }
    }

    pub fn to_str(self) -> &'static str {
        match self {
            Self::Top => "top",
            Self::Center => "center",
            Self::Bottom => "bottom",
        }
    }
}

/// `CT_TcPr` — Table cell properties.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct CT_TcPr {
    /// Cell width
    pub width: Option<CT_TblWidth>,
    /// Horizontal merge (number of grid columns spanned)
    pub grid_span: Option<u32>,
    /// Vertical merge
    pub v_merge: Option<VMerge>,
    /// Cell borders
    pub borders: Option<CT_TblBorders>,
    /// Cell shading
    pub shading: Option<CT_Shd>,
    /// Vertical alignment
    pub v_align: Option<ST_VerticalJc>,
    /// No-wrap text
    pub no_wrap: Option<bool>,
    /// Text direction
    pub text_direction: Option<String>,
}

#[allow(non_snake_case)]
impl CT_TcPr {
    pub fn from_xml(reader: &mut Reader<&[u8]>) -> Result<Self> {
        let mut pr = CT_TcPr::default();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(ref e)) => {
                    let name = e.name();
                    if matches_local_name(name.as_ref(), b"tcW") {
                        pr.width = Some(CT_TblWidth::from_xml_attrs(e)?);
                    } else if matches_local_name(name.as_ref(), b"gridSpan") {
                        if let Some(val) = get_val_attr(e)? {
                            pr.grid_span = Some(val.parse()?);
                        }
                    } else if matches_local_name(name.as_ref(), b"vMerge") {
                        if let Some(val) = get_val_attr(e)? {
                            pr.v_merge = Some(if val == "restart" {
                                VMerge::Restart
                            } else {
                                VMerge::Continue
                            });
                        } else {
                            // Empty vMerge means "continue"
                            pr.v_merge = Some(VMerge::Continue);
                        }
                    } else if matches_local_name(name.as_ref(), b"vAlign") {
                        if let Some(val) = get_val_attr(e)? {
                            pr.v_align = Some(ST_VerticalJc::from_str(&val));
                        }
                    } else if matches_local_name(name.as_ref(), b"shd") {
                        pr.shading = Some(CT_Shd::from_xml_attrs(e)?);
                    } else if matches_local_name(name.as_ref(), b"noWrap") {
                        pr.no_wrap = Some(true);
                    } else if matches_local_name(name.as_ref(), b"textDirection")
                        && let Some(val) = get_val_attr(e)?
                    {
                        pr.text_direction = Some(val);
                    }
                }
                Ok(Event::Start(ref e)) => {
                    let name = e.name();
                    if matches_local_name(name.as_ref(), b"tcBorders") {
                        pr.borders = Some(CT_TblBorders::from_xml(reader)?);
                    } else {
                        reader.read_to_end_into(name, &mut Vec::new())?;
                    }
                }
                Ok(Event::End(ref e)) if matches_local_name(e.name().as_ref(), b"tcPr") => {
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(pr)
    }

    pub fn to_xml<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        if self.is_empty() {
            return Ok(());
        }

        writer.write_event(Event::Start(BytesStart::new("w:tcPr")))?;

        if let Some(ref width) = self.width {
            width.write_xml(writer, "w:tcW")?;
        }

        if let Some(grid_span) = self.grid_span
            && grid_span > 1
        {
            let mut buf = itoa::Buffer::new();
            let mut e = BytesStart::new("w:gridSpan");
            e.push_attribute(("w:val", buf.format(grid_span)));
            writer.write_event(Event::Empty(e))?;
        }

        if let Some(ref vm) = self.v_merge {
            let mut e = BytesStart::new("w:vMerge");
            match vm {
                VMerge::Restart => e.push_attribute(("w:val", "restart")),
                VMerge::Continue => {} // empty element
            }
            writer.write_event(Event::Empty(e))?;
        }

        if let Some(ref borders) = self.borders
            && !borders.is_empty()
        {
            borders.to_xml(writer, "w:tcBorders")?;
        }

        if let Some(ref shd) = self.shading {
            shd.write_xml(writer, "w:shd")?;
        }

        if let Some(true) = self.no_wrap {
            writer.write_event(Event::Empty(BytesStart::new("w:noWrap")))?;
        }

        if let Some(ref va) = self.v_align {
            let mut e = BytesStart::new("w:vAlign");
            e.push_attribute(("w:val", va.to_str()));
            writer.write_event(Event::Empty(e))?;
        }

        if let Some(ref td) = self.text_direction {
            let mut e = BytesStart::new("w:textDirection");
            e.push_attribute(("w:val", td.as_str()));
            writer.write_event(Event::Empty(e))?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:tcPr")))?;
        Ok(())
    }

    fn is_empty(&self) -> bool {
        self.width.is_none()
            && self.grid_span.is_none()
            && self.v_merge.is_none()
            && self.borders.is_none()
            && self.shading.is_none()
            && self.v_align.is_none()
            && self.no_wrap.is_none()
            && self.text_direction.is_none()
    }
}

// ---- Table cell ----

/// Content that can appear inside a table cell.
#[derive(Debug, Clone, PartialEq)]
pub enum CellContent {
    /// A paragraph.
    Paragraph(CT_P),
    /// A nested table.
    Table(CT_Tbl),
}

/// `CT_Tc` — A table cell containing paragraphs and possibly nested tables.
#[derive(Debug, Clone, PartialEq)]
pub struct CT_Tc {
    pub properties: Option<CT_TcPr>,
    /// Cell content (paragraphs and nested tables).
    pub content: Vec<CellContent>,
}

#[allow(non_snake_case)]
impl CT_Tc {
    pub fn new() -> Self {
        CT_Tc {
            properties: None,
            // OOXML requires at least one paragraph per cell
            content: vec![CellContent::Paragraph(CT_P::new())],
        }
    }

    /// Get all paragraphs in this cell (excludes nested tables).
    pub fn paragraphs(&self) -> Vec<&CT_P> {
        self.content
            .iter()
            .filter_map(|c| match c {
                CellContent::Paragraph(p) => Some(p),
                CellContent::Table(_) => None,
            })
            .collect()
    }

    /// Get mutable reference to paragraphs (backward compatibility).
    pub fn paragraphs_mut(&mut self) -> Vec<&mut CT_P> {
        self.content
            .iter_mut()
            .filter_map(|c| match c {
                CellContent::Paragraph(p) => Some(p),
                CellContent::Table(_) => None,
            })
            .collect()
    }

    pub fn text(&self) -> String {
        self.paragraphs()
            .iter()
            .map(|p| p.text())
            .collect::<Vec<_>>()
            .join("\n")
    }

    pub fn from_xml(reader: &mut Reader<&[u8]>) -> Result<Self> {
        let mut properties = None;
        let mut content = Vec::new();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let name = e.name();
                    if matches_local_name(name.as_ref(), b"tcPr") {
                        properties = Some(CT_TcPr::from_xml(reader)?);
                    } else if matches_local_name(name.as_ref(), b"p") {
                        content.push(CellContent::Paragraph(CT_P::from_xml(reader)?));
                    } else if matches_local_name(name.as_ref(), b"tbl") {
                        content.push(CellContent::Table(CT_Tbl::from_xml(reader)?));
                    } else {
                        reader.read_to_end_into(name, &mut Vec::new())?;
                    }
                }
                Ok(Event::End(ref e)) if matches_local_name(e.name().as_ref(), b"tc") => {
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(CT_Tc {
            properties,
            content,
        })
    }

    pub fn to_xml<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_event(Event::Start(BytesStart::new("w:tc")))?;

        if let Some(ref props) = self.properties {
            props.to_xml(writer)?;
        }

        for item in &self.content {
            match item {
                CellContent::Paragraph(p) => p.to_xml(writer)?,
                CellContent::Table(tbl) => tbl.to_xml(writer)?,
            }
        }

        writer.write_event(Event::End(BytesEnd::new("w:tc")))?;
        Ok(())
    }
}

impl Default for CT_Tc {
    fn default() -> Self {
        Self::new()
    }
}

// ---- Table row ----

/// `CT_Row` — A table row containing cells.
#[derive(Debug, Clone, PartialEq)]
pub struct CT_Row {
    pub properties: Option<CT_TrPr>,
    pub cells: Vec<CT_Tc>,
}

#[allow(non_snake_case)]
impl CT_Row {
    pub fn new() -> Self {
        CT_Row {
            properties: None,
            cells: Vec::new(),
        }
    }

    pub fn from_xml(reader: &mut Reader<&[u8]>) -> Result<Self> {
        let mut properties = None;
        let mut cells = Vec::new();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let name = e.name();
                    if matches_local_name(name.as_ref(), b"trPr") {
                        properties = Some(CT_TrPr::from_xml(reader)?);
                    } else if matches_local_name(name.as_ref(), b"tc") {
                        cells.push(CT_Tc::from_xml(reader)?);
                    } else {
                        reader.read_to_end_into(name, &mut Vec::new())?;
                    }
                }
                Ok(Event::End(ref e)) if matches_local_name(e.name().as_ref(), b"tr") => {
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(CT_Row { properties, cells })
    }

    pub fn to_xml<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_event(Event::Start(BytesStart::new("w:tr")))?;

        if let Some(ref props) = self.properties {
            props.to_xml(writer)?;
        }

        for cell in &self.cells {
            cell.to_xml(writer)?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:tr")))?;
        Ok(())
    }
}

impl Default for CT_Row {
    fn default() -> Self {
        Self::new()
    }
}

// ---- Table ----

/// `CT_Tbl` — A table element containing rows.
#[derive(Debug, Clone, PartialEq)]
pub struct CT_Tbl {
    pub properties: Option<CT_TblPr>,
    pub grid: Option<CT_TblGrid>,
    pub rows: Vec<CT_Row>,
}

#[allow(non_snake_case)]
impl CT_Tbl {
    pub fn new() -> Self {
        CT_Tbl {
            properties: None,
            grid: None,
            rows: Vec::new(),
        }
    }

    pub fn from_xml(reader: &mut Reader<&[u8]>) -> Result<Self> {
        let mut properties = None;
        let mut grid = None;
        let mut rows = Vec::new();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let name = e.name();
                    if matches_local_name(name.as_ref(), b"tblPr") {
                        properties = Some(CT_TblPr::from_xml(reader)?);
                    } else if matches_local_name(name.as_ref(), b"tblGrid") {
                        grid = Some(CT_TblGrid::from_xml(reader)?);
                    } else if matches_local_name(name.as_ref(), b"tr") {
                        rows.push(CT_Row::from_xml(reader)?);
                    } else {
                        reader.read_to_end_into(name, &mut Vec::new())?;
                    }
                }
                Ok(Event::End(ref e)) if matches_local_name(e.name().as_ref(), b"tbl") => {
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(CT_Tbl {
            properties,
            grid,
            rows,
        })
    }

    pub fn to_xml<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_event(Event::Start(BytesStart::new("w:tbl")))?;

        if let Some(ref props) = self.properties {
            props.to_xml(writer)?;
        }

        if let Some(ref grid) = self.grid {
            grid.to_xml(writer)?;
        }

        for row in &self.rows {
            row.to_xml(writer)?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:tbl")))?;
        Ok(())
    }
}

impl Default for CT_Tbl {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn parse_table(xml: &str) -> CT_Tbl {
        let full = format!("<w:tbl>{xml}</w:tbl>");
        let mut reader = Reader::from_str(&full);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if matches_local_name(e.name().as_ref(), b"tbl") => break,
                _ => {}
            }
            buf.clear();
        }
        CT_Tbl::from_xml(&mut reader).unwrap()
    }

    #[test]
    fn parse_simple_table() {
        let tbl = parse_table(
            r#"<w:tblPr><w:tblW w:w="5000" w:type="dxa"/></w:tblPr>
               <w:tblGrid><w:gridCol w:w="2500"/><w:gridCol w:w="2500"/></w:tblGrid>
               <w:tr>
                 <w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc>
                 <w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc>
               </w:tr>
               <w:tr>
                 <w:tc><w:p><w:r><w:t>A2</w:t></w:r></w:p></w:tc>
                 <w:tc><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc>
               </w:tr>"#,
        );
        assert_eq!(tbl.rows.len(), 2);
        assert_eq!(tbl.rows[0].cells.len(), 2);
        assert_eq!(tbl.rows[0].cells[0].text(), "A1");
        assert_eq!(tbl.rows[1].cells[1].text(), "B2");

        let grid = tbl.grid.unwrap();
        assert_eq!(grid.columns.len(), 2);
        assert_eq!(grid.columns[0].width, Twips(2500));

        let pr = tbl.properties.unwrap();
        assert_eq!(pr.width.as_ref().unwrap().w, 5000);
    }

    #[test]
    fn parse_cell_merge() {
        let tbl = parse_table(
            r#"<w:tblGrid><w:gridCol w:w="2500"/><w:gridCol w:w="2500"/></w:tblGrid>
               <w:tr>
                 <w:tc>
                   <w:tcPr><w:gridSpan w:val="2"/></w:tcPr>
                   <w:p><w:r><w:t>Merged</w:t></w:r></w:p>
                 </w:tc>
               </w:tr>
               <w:tr>
                 <w:tc>
                   <w:tcPr><w:vMerge w:val="restart"/></w:tcPr>
                   <w:p><w:r><w:t>VM Start</w:t></w:r></w:p>
                 </w:tc>
                 <w:tc><w:p/></w:tc>
               </w:tr>
               <w:tr>
                 <w:tc>
                   <w:tcPr><w:vMerge/></w:tcPr>
                   <w:p/>
                 </w:tc>
                 <w:tc><w:p/></w:tc>
               </w:tr>"#,
        );

        // First row: horizontal merge
        assert_eq!(
            tbl.rows[0].cells[0].properties.as_ref().unwrap().grid_span,
            Some(2)
        );

        // Second row: vertical merge start
        assert_eq!(
            tbl.rows[1].cells[0].properties.as_ref().unwrap().v_merge,
            Some(VMerge::Restart)
        );

        // Third row: vertical merge continue
        assert_eq!(
            tbl.rows[2].cells[0].properties.as_ref().unwrap().v_merge,
            Some(VMerge::Continue)
        );
    }

    #[test]
    fn parse_table_borders() {
        let tbl = parse_table(
            r#"<w:tblPr>
                 <w:tblBorders>
                   <w:top w:val="single" w:sz="4" w:color="000000"/>
                   <w:bottom w:val="single" w:sz="4" w:color="000000"/>
                   <w:left w:val="single" w:sz="4" w:color="000000"/>
                   <w:right w:val="single" w:sz="4" w:color="000000"/>
                   <w:insideH w:val="single" w:sz="4" w:color="000000"/>
                   <w:insideV w:val="single" w:sz="4" w:color="000000"/>
                 </w:tblBorders>
               </w:tblPr>
               <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
               <w:tr><w:tc><w:p/></w:tc></w:tr>"#,
        );

        let borders = tbl.properties.unwrap().borders.unwrap();
        assert_eq!(borders.top.unwrap().val, ST_Border::Single);
        assert_eq!(borders.inside_h.unwrap().val, ST_Border::Single);
        assert_eq!(borders.inside_v.unwrap().val, ST_Border::Single);
    }

    #[test]
    fn parse_cell_shading() {
        let tbl = parse_table(
            r#"<w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
               <w:tr>
                 <w:tc>
                   <w:tcPr><w:shd w:val="clear" w:fill="FFFF00"/></w:tcPr>
                   <w:p/>
                 </w:tc>
               </w:tr>"#,
        );

        let shd = tbl.rows[0].cells[0]
            .properties
            .as_ref()
            .unwrap()
            .shading
            .as_ref()
            .unwrap();
        assert_eq!(shd.fill, Some("FFFF00".to_string()));
    }

    #[test]
    fn parse_row_properties() {
        let tbl = parse_table(
            r#"<w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
               <w:tr>
                 <w:trPr>
                   <w:trHeight w:val="720" w:hRule="exact"/>
                   <w:tblHeader/>
                 </w:trPr>
                 <w:tc><w:p/></w:tc>
               </w:tr>"#,
        );

        let tr_pr = tbl.rows[0].properties.as_ref().unwrap();
        assert_eq!(tr_pr.height, Some(Twips(720)));
        assert_eq!(tr_pr.height_rule, Some("exact".to_string()));
        assert_eq!(tr_pr.header, Some(true));
    }

    #[test]
    fn round_trip_table() {
        let mut tbl = CT_Tbl::new();
        tbl.properties = Some(CT_TblPr {
            width: Some(CT_TblWidth::dxa(9000)),
            borders: Some(CT_TblBorders {
                top: Some(CT_BorderEdge {
                    val: ST_Border::Single,
                    sz: Some(4),
                    space: Some(0),
                    color: Some("000000".to_string()),
                }),
                bottom: Some(CT_BorderEdge {
                    val: ST_Border::Single,
                    sz: Some(4),
                    space: Some(0),
                    color: Some("000000".to_string()),
                }),
                ..Default::default()
            }),
            ..Default::default()
        });
        tbl.grid = Some(CT_TblGrid {
            columns: vec![
                CT_TblGridCol { width: Twips(4500) },
                CT_TblGridCol { width: Twips(4500) },
            ],
        });

        let mut row = CT_Row::new();
        let mut cell1 = CT_Tc::new();
        cell1.paragraphs_mut()[0].add_run("Hello");
        let mut cell2 = CT_Tc::new();
        cell2.paragraphs_mut()[0].add_run("World");
        row.cells.push(cell1);
        row.cells.push(cell2);
        tbl.rows.push(row);

        // Serialize
        let mut output = Vec::new();
        let mut writer = Writer::new(&mut output);
        tbl.to_xml(&mut writer).unwrap();
        let xml = String::from_utf8(output).unwrap();

        // Parse back
        let parsed = parse_table(
            xml.strip_prefix("<w:tbl>")
                .unwrap()
                .strip_suffix("</w:tbl>")
                .unwrap(),
        );

        assert_eq!(parsed.rows.len(), 1);
        assert_eq!(parsed.rows[0].cells.len(), 2);
        assert_eq!(parsed.rows[0].cells[0].text(), "Hello");
        assert_eq!(parsed.rows[0].cells[1].text(), "World");

        let grid = parsed.grid.unwrap();
        assert_eq!(grid.columns.len(), 2);
        assert_eq!(grid.columns[0].width, Twips(4500));

        let borders = parsed.properties.unwrap().borders.unwrap();
        assert!(borders.top.is_some());
        assert!(borders.bottom.is_some());
    }

    #[test]
    fn nested_table_xml_round_trip() {
        use crate::text::CT_P;

        // Build a cell containing a paragraph + a nested table
        let mut outer_cell = CT_Tc::new();
        outer_cell.paragraphs_mut()[0].add_run("Before table");

        let mut nested_tbl = CT_Tbl::new();
        nested_tbl.grid = Some(CT_TblGrid {
            columns: vec![CT_TblGridCol { width: Twips(2000) }],
        });
        let mut nested_row = CT_Row::new();
        let mut nested_cell = CT_Tc::new();
        nested_cell.paragraphs_mut()[0].add_run("Nested content");
        nested_row.cells.push(nested_cell);
        nested_tbl.rows.push(nested_row);

        outer_cell.content.push(CellContent::Table(nested_tbl));

        let mut after = CT_P::new();
        after.add_run("After table");
        outer_cell.content.push(CellContent::Paragraph(after));

        // Serialize
        let mut output = Vec::new();
        let mut writer = Writer::new(&mut output);
        outer_cell.to_xml(&mut writer).unwrap();
        let xml = String::from_utf8(output).unwrap();

        // Should contain nested <w:tbl>
        assert!(xml.contains("<w:tbl>"));
        assert!(xml.contains("Nested content"));

        // Parse back
        let inner_xml = xml
            .strip_prefix("<w:tc>")
            .unwrap()
            .strip_suffix("</w:tc>")
            .unwrap();
        let full_xml = format!(
            "<w:tc xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">{inner_xml}</w:tc>"
        );
        let mut reader = Reader::from_str(&full_xml);
        reader.config_mut().trim_text(true);
        // Skip start tag
        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) if e.local_name().as_ref() == b"tc" => break,
                _ => {}
            }
        }
        let parsed = CT_Tc::from_xml(&mut reader).unwrap();

        // Check structure: 2 paragraphs + 1 nested table
        assert_eq!(parsed.paragraphs().len(), 2);
        assert_eq!(parsed.paragraphs()[0].text(), "Before table");
        assert_eq!(parsed.paragraphs()[1].text(), "After table");

        // Check nested table
        let tables: Vec<_> = parsed
            .content
            .iter()
            .filter_map(|c| match c {
                CellContent::Table(t) => Some(t),
                _ => None,
            })
            .collect();
        assert_eq!(tables.len(), 1);
        assert_eq!(tables[0].rows.len(), 1);
        assert_eq!(tables[0].rows[0].cells[0].text(), "Nested content");
    }

    #[test]
    fn paragraphs_method_backward_compat() {
        let mut cell = CT_Tc::new();
        // Cell starts with one empty paragraph
        assert_eq!(cell.paragraphs().len(), 1);

        // Add a run to existing paragraph
        cell.paragraphs_mut()[0].add_run("First");

        // Add a nested table (should not appear in paragraphs())
        let nested = CT_Tbl::new();
        cell.content.push(CellContent::Table(nested));

        // Add another paragraph
        let mut p = CT_P::new();
        p.add_run("Second");
        cell.content.push(CellContent::Paragraph(p));

        // paragraphs() should return only the 2 CT_P items
        assert_eq!(cell.paragraphs().len(), 2);
        assert_eq!(cell.paragraphs()[0].text(), "First");
        assert_eq!(cell.paragraphs()[1].text(), "Second");

        // text() should concat paragraph text with newline separator
        assert_eq!(cell.text(), "First\nSecond");
    }
}
