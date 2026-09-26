//! Table — a block-level container for rows and cells of content.

use rdocx_oxml::borders::CT_BorderEdge;
use rdocx_oxml::drawing::AnchorAlignH;
use rdocx_oxml::properties::CT_Shd;
use rdocx_oxml::shared::ST_Jc;
pub use rdocx_oxml::table::VMerge;
use rdocx_oxml::table::{
    CT_Row, CT_Tbl, CT_TblBorders, CT_TblCellMar, CT_TblLook, CT_TblPPr, CT_TblPr, CT_TblWidth,
    CT_Tc, CT_TcPr, CT_TrPr, CellContent, ST_TblAnchor, ST_TblOverlap, ST_VerticalJc, ST_YAlign,
};
use rdocx_oxml::text::CT_P;

use crate::content_control::ContentControlRef;
use crate::document::{DrawingHorizontalAlignment, DrawingVerticalAlignment};
use crate::paragraph::{Paragraph, ParagraphRef};
use crate::{Error, Length, Result};

/// Vertical alignment within a table cell.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VerticalAlignment {
    Top,
    Center,
    Bottom,
}

impl VerticalAlignment {
    fn to_st(self) -> ST_VerticalJc {
        match self {
            Self::Top => ST_VerticalJc::Top,
            Self::Center => ST_VerticalJc::Center,
            Self::Bottom => ST_VerticalJc::Bottom,
        }
    }

    /// Map the OOXML value onto the three alignments this facade exposes.
    ///
    /// `both` has no facade spelling and reads back as `Top`, which is how it
    /// lays out. The cell keeps its source value in `CT_TcPr`.
    fn from_st(st: ST_VerticalJc) -> Self {
        match st {
            ST_VerticalJc::Center => Self::Center,
            ST_VerticalJc::Bottom => Self::Bottom,
            ST_VerticalJc::Top | _ => Self::Top,
        }
    }
}

/// A complete table width mode.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum TableWidth {
    /// Let Word choose the table width from its content and container.
    Auto,
    /// Use an exact physical width.
    Fixed(Length),
    /// Use a percentage of the containing width, from 0 through 100.
    Percentage(f64),
}

/// The table layout algorithm written to `w:tblLayout`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableLayout {
    /// Let Word resize columns from their content.
    AutoFit,
    /// Keep the authored grid widths fixed.
    Fixed,
}

/// What a floating table's position is measured from.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableAnchor {
    /// The surrounding text margin.
    Margin,
    /// The page edge.
    Page,
    /// The surrounding text.
    Text,
}

impl TableAnchor {
    fn to_st(self) -> ST_TblAnchor {
        match self {
            Self::Margin => ST_TblAnchor::Margin,
            Self::Page => ST_TblAnchor::Page,
            Self::Text => ST_TblAnchor::Text,
        }
    }

    fn from_st(value: ST_TblAnchor) -> Self {
        match value {
            ST_TblAnchor::Margin => Self::Margin,
            ST_TblAnchor::Page => Self::Page,
            ST_TblAnchor::Text => Self::Text,
        }
    }
}

/// The horizontal placement of a floating table.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum TableFloatX {
    /// Align against the horizontal anchor.
    Align(DrawingHorizontalAlignment),
    /// Offset from the horizontal anchor.
    Offset(Length),
}

/// The vertical placement of a floating table.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum TableFloatY {
    /// Keep the table in the vertical flow rather than floating it.
    Inline,
    /// Align against the vertical anchor.
    Align(DrawingVerticalAlignment),
    /// Offset from the vertical anchor.
    Offset(Length),
}

/// The clearance a floating table keeps from the text that flows around it.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct TableTextDistance {
    pub top: Length,
    pub right: Length,
    pub bottom: Length,
    pub left: Length,
}

/// A complete floating table position, written to `w:tblpPr`.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct TableFloatPosition {
    pub horizontal_anchor: TableAnchor,
    pub vertical_anchor: TableAnchor,
    pub horizontal: TableFloatX,
    pub vertical: TableFloatY,
    pub distance_from_text: TableTextDistance,
}

/// Whether a floating table may overlap another float.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableOverlap {
    /// Word moves this table rather than letting it overlap.
    Never,
    /// Word allows the overlap.
    Allow,
}

impl TableOverlap {
    fn to_st(self) -> ST_TblOverlap {
        match self {
            Self::Never => ST_TblOverlap::Never,
            Self::Allow => ST_TblOverlap::Overlap,
        }
    }

    fn from_st(value: ST_TblOverlap) -> Self {
        match value {
            ST_TblOverlap::Never => Self::Never,
            ST_TblOverlap::Overlap => Self::Allow,
        }
    }
}

/// One edge in the table border model.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableBorderEdge {
    Top,
    Bottom,
    Left,
    Right,
    InsideHorizontal,
    InsideVertical,
}

/// One edge in the cell border model.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CellBorderEdge {
    /// The top edge.
    Top,
    /// The bottom edge.
    Bottom,
    /// The left edge.
    Left,
    /// The right edge.
    Right,
    /// Horizontal edges between cells.
    InsideHorizontal,
    /// Vertical edges between cells.
    InsideVertical,
}

/// A checked row height and its Word height rule.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum RowHeight {
    /// The row may grow beyond this minimum.
    AtLeast(Length),
    /// The row uses exactly this height.
    Exact(Length),
}

/// Text flow written to `w:textDirection` for a table cell.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CellTextDirection {
    /// Horizontal text from left to right, with lines from top to bottom.
    LeftToRightTopToBottom,
    /// Vertical text from top to bottom, with lines from right to left.
    TopToBottomRightToLeft,
    /// Vertical text from bottom to top, with lines from left to right.
    BottomToTopLeftToRight,
    /// Vertically oriented left-to-right text with top-to-bottom lines.
    LeftToRightTopToBottomVertical,
    /// Vertically oriented top-to-bottom text with right-to-left lines.
    TopToBottomRightToLeftVertical,
    /// Vertically oriented top-to-bottom text with left-to-right lines.
    TopToBottomLeftToRightVertical,
}

impl CellTextDirection {
    fn from_str(value: &str) -> Option<Self> {
        match value {
            "lrTb" => Some(Self::LeftToRightTopToBottom),
            "tbRl" => Some(Self::TopToBottomRightToLeft),
            "btLr" => Some(Self::BottomToTopLeftToRight),
            "lrTbV" => Some(Self::LeftToRightTopToBottomVertical),
            "tbRlV" => Some(Self::TopToBottomRightToLeftVertical),
            "tbLrV" => Some(Self::TopToBottomLeftToRightVertical),
            _ => None,
        }
    }

    fn to_str(self) -> &'static str {
        match self {
            Self::LeftToRightTopToBottom => "lrTb",
            Self::TopToBottomRightToLeft => "tbRl",
            Self::BottomToTopLeftToRight => "btLr",
            Self::LeftToRightTopToBottomVertical => "lrTbV",
            Self::TopToBottomRightToLeftVertical => "tbRlV",
            Self::TopToBottomLeftToRightVertical => "tbLrV",
        }
    }
}

/// Direct conditional table-style regions on a row or cell.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct TableConditionalFormatting {
    /// Apply first-row conditional formatting.
    pub first_row: bool,
    /// Apply last-row conditional formatting.
    pub last_row: bool,
    /// Apply first-column conditional formatting.
    pub first_column: bool,
    /// Apply last-column conditional formatting.
    pub last_column: bool,
    /// Apply odd vertical-band conditional formatting.
    pub odd_vertical_band: bool,
    /// Apply even vertical-band conditional formatting.
    pub even_vertical_band: bool,
    /// Apply odd horizontal-band conditional formatting.
    pub odd_horizontal_band: bool,
    /// Apply even horizontal-band conditional formatting.
    pub even_horizontal_band: bool,
    /// Apply the first-row and last-column corner formatting.
    pub first_row_last_column: bool,
    /// Apply the first-row and first-column corner formatting.
    pub first_row_first_column: bool,
    /// Apply the last-row and last-column corner formatting.
    pub last_row_last_column: bool,
    /// Apply the last-row and first-column corner formatting.
    pub last_row_first_column: bool,
}

impl TableConditionalFormatting {
    pub(crate) fn to_value(self) -> String {
        [
            self.first_row,
            self.last_row,
            self.first_column,
            self.last_column,
            self.odd_vertical_band,
            self.even_vertical_band,
            self.odd_horizontal_band,
            self.even_horizontal_band,
            self.first_row_last_column,
            self.first_row_first_column,
            self.last_row_last_column,
            self.last_row_first_column,
        ]
        .into_iter()
        .map(|enabled| if enabled { '1' } else { '0' })
        .collect()
    }

    pub(crate) fn from_value(value: &str) -> Option<Self> {
        let bits = value.as_bytes();
        if bits.len() != 12 || bits.iter().any(|bit| !matches!(bit, b'0' | b'1')) {
            return None;
        }
        let enabled = |index: usize| bits[index] == b'1';
        Some(Self {
            first_row: enabled(0),
            last_row: enabled(1),
            first_column: enabled(2),
            last_column: enabled(3),
            odd_vertical_band: enabled(4),
            even_vertical_band: enabled(5),
            odd_horizontal_band: enabled(6),
            even_horizontal_band: enabled(7),
            first_row_last_column: enabled(8),
            first_row_first_column: enabled(9),
            last_row_last_column: enabled(10),
            last_row_first_column: enabled(11),
        })
    }
}

/// Conditional table-style regions selected by `w:tblLook`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct TableLook {
    /// Apply first-row conditional formatting.
    pub first_row: bool,
    /// Apply last-row conditional formatting.
    pub last_row: bool,
    /// Apply first-column conditional formatting.
    pub first_column: bool,
    /// Apply last-column conditional formatting.
    pub last_column: bool,
    /// Apply horizontal row banding.
    pub horizontal_banding: bool,
    /// Apply vertical column banding.
    pub vertical_banding: bool,
}

impl TableLook {
    /// The legacy `w:tblLook/@w:val` bitmask for this selection.
    ///
    /// The two banding fields are inverted, because the mask records the
    /// suppression bits `noHBand` and `noVBand`. Four uppercase hex digits is
    /// the form Word writes.
    fn to_mask(self) -> String {
        let mut mask = 0u16;
        for (enabled, bit) in [
            (self.first_row, 0x0020),
            (self.last_row, 0x0040),
            (self.first_column, 0x0080),
            (self.last_column, 0x0100),
            (!self.horizontal_banding, 0x0200),
            (!self.vertical_banding, 0x0400),
        ] {
            if enabled {
                mask |= bit;
            }
        }
        format!("{mask:04X}")
    }
}

/// Default margins applied to every table cell.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct TableCellMargins {
    /// Top cell margin.
    pub top: Option<Length>,
    /// Right cell margin.
    pub right: Option<Length>,
    /// Bottom cell margin.
    pub bottom: Option<Length>,
    /// Left cell margin.
    pub left: Option<Length>,
}

/// An immutable table border edge.
#[derive(Debug, Clone, Copy)]
pub struct TableBorderRef<'a> {
    inner: &'a CT_BorderEdge,
}

impl TableBorderRef<'_> {
    /// The OOXML border style name, including explicit `none` edges.
    pub fn style(self) -> &'static str {
        self.inner.val.to_str()
    }

    /// Border width in eighths of a point.
    pub fn size_eighths_pt(self) -> Option<u32> {
        self.inner.sz
    }

    /// Border color, normally six hexadecimal digits or `auto`.
    pub fn color(&self) -> Option<&str> {
        self.inner.color.as_deref()
    }
}

fn checked_band_size(name: &str, value: u32) -> Result<u32> {
    if value == 0 {
        return Err(Error::Other(format!("{name} must be at least one")));
    }
    Ok(value)
}

fn checked_table_twips(name: &str, value: Length) -> Result<i32> {
    if value.to_emu() < 0 {
        return Err(Error::Other(format!("table {name} cannot be negative")));
    }
    i32::try_from(value.to_emu() / 635)
        .map_err(|_| Error::Other(format!("table {name} exceeds the signed twip range")))
}

/// Check a twip measurement that Word allows to be negative.
///
/// `w:tblpX` and `w:tblpY` place a float on either side of their anchor, so
/// the nonnegative rule in [`checked_table_twips`] does not apply to them.
fn checked_table_offset_twips(name: &str, value: Length) -> Result<i32> {
    i32::try_from(value.to_emu() / 635)
        .map_err(|_| Error::Other(format!("table {name} exceeds the signed twip range")))
}

/// Validate a complete width mode before it reaches the document.
fn checked_table_width(name: &str, width: TableWidth) -> Result<CT_TblWidth> {
    Ok(match width {
        TableWidth::Auto => CT_TblWidth::auto(),
        TableWidth::Fixed(value) => CT_TblWidth::dxa(checked_table_twips(name, value)?),
        TableWidth::Percentage(percent) => {
            if !percent.is_finite() || !(0.0..=100.0).contains(&percent) {
                return Err(Error::Other(format!(
                    "table {name} percentage must be finite and between 0 and 100"
                )));
            }
            CT_TblWidth::pct((percent * 50.0) as i32)
        }
    })
}

/// Project a stored width onto the public mode, or `None` for a spelling the
/// facade does not author, such as `nil`.
fn table_width_from_ct(width: &CT_TblWidth) -> Option<TableWidth> {
    match width.width_type.as_str() {
        "auto" => Some(TableWidth::Auto),
        "dxa" => Some(TableWidth::Fixed(Length::twips(width.w))),
        "pct" => Some(TableWidth::Percentage(width.w as f64 / 50.0)),
        _ => None,
    }
}

/// Reject an empty accessible string rather than writing a blank attribute.
fn checked_table_text(name: &str, value: &str) -> Result<String> {
    if value.trim().is_empty() {
        return Err(Error::Other(format!("table {name} cannot be empty")));
    }
    Ok(value.to_owned())
}

fn float_position_to_ct(position: TableFloatPosition) -> Result<CT_TblPPr> {
    let mut resolved = CT_TblPPr {
        horz_anchor: Some(position.horizontal_anchor.to_st()),
        vert_anchor: Some(position.vertical_anchor.to_st()),
        left_from_text: Some(rdocx_oxml::Twips(checked_table_twips(
            "left distance from text",
            position.distance_from_text.left,
        )?)),
        right_from_text: Some(rdocx_oxml::Twips(checked_table_twips(
            "right distance from text",
            position.distance_from_text.right,
        )?)),
        top_from_text: Some(rdocx_oxml::Twips(checked_table_twips(
            "top distance from text",
            position.distance_from_text.top,
        )?)),
        bottom_from_text: Some(rdocx_oxml::Twips(checked_table_twips(
            "bottom distance from text",
            position.distance_from_text.bottom,
        )?)),
        ..CT_TblPPr::default()
    };
    match position.horizontal {
        TableFloatX::Align(alignment) => {
            resolved.tbl_p_x_spec = Some(AnchorAlignH::from(alignment));
        }
        TableFloatX::Offset(offset) => {
            resolved.tbl_p_x = Some(rdocx_oxml::Twips(checked_table_offset_twips(
                "horizontal float offset",
                offset,
            )?));
        }
    }
    match position.vertical {
        TableFloatY::Inline => resolved.tbl_p_y_spec = Some(ST_YAlign::Inline),
        TableFloatY::Align(alignment) => {
            resolved.tbl_p_y_spec = Some(match alignment {
                DrawingVerticalAlignment::Top => ST_YAlign::Top,
                DrawingVerticalAlignment::Center => ST_YAlign::Center,
                DrawingVerticalAlignment::Bottom => ST_YAlign::Bottom,
                DrawingVerticalAlignment::Inside => ST_YAlign::Inside,
                DrawingVerticalAlignment::Outside => ST_YAlign::Outside,
            });
        }
        TableFloatY::Offset(offset) => {
            resolved.tbl_p_y = Some(rdocx_oxml::Twips(checked_table_offset_twips(
                "vertical float offset",
                offset,
            )?));
        }
    }
    Ok(resolved)
}

/// Project a parsed `w:tblpPr` onto the public position.
///
/// An alignment spec wins over an offset, which is Word's own rule. An absent
/// anchor reads as `margin` and an absent offset reads as zero, which are
/// Word's defaults, so a partial `w:tblpPr` still reads as a complete
/// position.
fn float_position_from_ct(position: &CT_TblPPr) -> TableFloatPosition {
    let horizontal = match position.tbl_p_x_spec {
        Some(AnchorAlignH::Left) => TableFloatX::Align(DrawingHorizontalAlignment::Left),
        Some(AnchorAlignH::Center) => TableFloatX::Align(DrawingHorizontalAlignment::Center),
        Some(AnchorAlignH::Right) => TableFloatX::Align(DrawingHorizontalAlignment::Right),
        Some(AnchorAlignH::Inside) => TableFloatX::Align(DrawingHorizontalAlignment::Inside),
        Some(AnchorAlignH::Outside) => TableFloatX::Align(DrawingHorizontalAlignment::Outside),
        None => TableFloatX::Offset(Length::twips(
            position.tbl_p_x.map(|value| value.0).unwrap_or(0),
        )),
    };
    let vertical = match position.tbl_p_y_spec {
        Some(ST_YAlign::Inline) => TableFloatY::Inline,
        Some(ST_YAlign::Top) => TableFloatY::Align(DrawingVerticalAlignment::Top),
        Some(ST_YAlign::Center) => TableFloatY::Align(DrawingVerticalAlignment::Center),
        Some(ST_YAlign::Bottom) => TableFloatY::Align(DrawingVerticalAlignment::Bottom),
        Some(ST_YAlign::Inside) => TableFloatY::Align(DrawingVerticalAlignment::Inside),
        Some(ST_YAlign::Outside) => TableFloatY::Align(DrawingVerticalAlignment::Outside),
        None => TableFloatY::Offset(Length::twips(
            position.tbl_p_y.map(|value| value.0).unwrap_or(0),
        )),
    };
    let distance =
        |value: Option<rdocx_oxml::Twips>| Length::twips(value.map_or(0, |twips| twips.0));
    TableFloatPosition {
        horizontal_anchor: position
            .horz_anchor
            .map_or(TableAnchor::Margin, TableAnchor::from_st),
        vertical_anchor: position
            .vert_anchor
            .map_or(TableAnchor::Margin, TableAnchor::from_st),
        horizontal,
        vertical,
        distance_from_text: TableTextDistance {
            top: distance(position.top_from_text),
            right: distance(position.right_from_text),
            bottom: distance(position.bottom_from_text),
            left: distance(position.left_from_text),
        },
    }
}

fn checked_table_color(name: &str, value: &str) -> Result<String> {
    if value.eq_ignore_ascii_case("auto")
        || (value.len() == 6 && value.bytes().all(|byte| byte.is_ascii_hexdigit()))
    {
        Ok(value.to_owned())
    } else {
        Err(Error::Other(format!(
            "table {name} must be 'auto' or six hexadecimal digits"
        )))
    }
}

fn checked_table_border(
    style: crate::BorderStyle,
    size_eighths_pt: u32,
    color: &str,
) -> Result<CT_BorderEdge> {
    if size_eighths_pt > 96 || (style != crate::BorderStyle::None && size_eighths_pt == 0) {
        return Err(Error::Other(
            "table border width must be 1 through 96 for a visible edge, or 0 through 96 for an invisible edge"
                .to_owned(),
        ));
    }
    Ok(CT_BorderEdge {
        val: style.to_st(),
        sz: Some(size_eighths_pt),
        space: Some(0),
        color: Some(checked_table_color("border color", color)?),
        extra_attributes: Vec::new(),
        nil: false,
    })
}

pub(crate) fn row_cell_ranges(row: &CT_Row, grid_columns: usize) -> Result<Vec<(usize, usize)>> {
    let before = row
        .properties
        .as_ref()
        .and_then(|properties| properties.grid_before)
        .unwrap_or(0) as usize;
    let after = row
        .properties
        .as_ref()
        .and_then(|properties| properties.grid_after)
        .unwrap_or(0) as usize;
    let limit = grid_columns.checked_sub(after).ok_or_else(|| {
        Error::Other("table row grid omissions exceed the active grid".to_owned())
    })?;
    if before > limit || row.cells.is_empty() {
        return Err(Error::Other(
            "table row does not cover a valid active-grid range".to_owned(),
        ));
    }
    let mut start = before;
    let mut ranges = Vec::with_capacity(row.cells.len());
    for cell in &row.cells {
        let span = cell
            .properties
            .as_ref()
            .and_then(|properties| properties.grid_span)
            .unwrap_or(1) as usize;
        if span == 0 {
            return Err(Error::Other(
                "table cell grid span must be positive".to_owned(),
            ));
        }
        let end = start
            .checked_add(span)
            .ok_or_else(|| Error::Other("table cell grid span overflows".to_owned()))?;
        if end > limit {
            return Err(Error::Other(
                "table row cells exceed the active grid".to_owned(),
            ));
        }
        ranges.push((start, end));
        start = end;
    }
    if start != limit {
        return Err(Error::Other(
            "table row cells do not cover the active grid".to_owned(),
        ));
    }
    Ok(ranges)
}

pub(crate) fn validate_table_topology(table: &CT_Tbl) -> Result<()> {
    let grid = table
        .grid
        .as_ref()
        .filter(|grid| !grid.columns.is_empty())
        .ok_or_else(|| Error::Other("table has no active grid".to_owned()))?;
    if grid.columns.iter().any(|column| column.width.0 <= 0) {
        return Err(Error::Other(
            "table grid columns must have positive widths".to_owned(),
        ));
    }
    let grid_columns = grid.columns.len();
    let mut previous_vertical_ranges = std::collections::BTreeSet::new();
    for row in &table.rows {
        let ranges = row_cell_ranges(row, grid_columns)?;
        let mut current_vertical_ranges = std::collections::BTreeSet::new();
        for (cell, range) in row.cells.iter().zip(ranges) {
            match cell
                .properties
                .as_ref()
                .and_then(|properties| properties.v_merge.as_ref())
            {
                Some(VMerge::Restart) => {
                    current_vertical_ranges.insert(range);
                }
                Some(VMerge::Continue) => {
                    if !previous_vertical_ranges.contains(&range) {
                        return Err(Error::Other(
                            "vertical merge continuation has no matching cell above".to_owned(),
                        ));
                    }
                    current_vertical_ranges.insert(range);
                }
                None => {}
            }
        }
        previous_vertical_ranges = current_vertical_ranges;
    }
    Ok(())
}

fn cell_is_discardable_for_grid_edit(cell: &CT_Tc) -> bool {
    let empty_content = matches!(
        cell.content.as_slice(),
        [CellContent::Paragraph(paragraph)] if paragraph == &CT_P::new()
    );
    let width_only = cell.properties.as_ref().is_none_or(|properties| {
        properties.grid_span.is_none()
            && properties.h_merge.is_none()
            && properties.v_merge.is_none()
            && properties.borders.is_none()
            && properties.shading.is_none()
            && properties.v_align.is_none()
            && properties.no_wrap.is_none()
            && properties.cell_margin.is_none()
            && properties.text_direction.is_none()
            && properties.cnf_style.is_none()
            && properties.extra_xml.is_empty()
    });
    empty_content && cell.extra_xml.is_empty() && width_only
}

fn empty_grid_cell(width: i32) -> CT_Tc {
    let mut cell = CT_Tc::new();
    cell.properties = Some(CT_TcPr {
        width: Some(CT_TblWidth::dxa(width)),
        ..Default::default()
    });
    cell
}

// ---- Mutable Table ----

/// A mutable reference to a table in a document.
pub struct Table<'a> {
    pub(crate) inner: &'a mut CT_Tbl,
}

impl<'a> Table<'a> {
    /// Set the table style by ID.
    pub fn style(mut self, style_id: &str) -> Self {
        self.set_style(style_id);
        self
    }

    /// Set the table style by ID in place.
    pub fn set_style(&mut self, style_id: &str) {
        self.ensure_tbl_pr().style_id = Some(style_id.to_string());
    }

    /// Set the table width in twips (dxa).
    pub fn width(mut self, length: Length) -> Self {
        self.set_width(length);
        self
    }

    /// Set the table width in place.
    pub fn set_width(&mut self, length: Length) {
        self.ensure_tbl_pr().width = Some(CT_TblWidth::dxa(length.as_twips().0));
    }

    /// Set a checked auto, fixed, or percentage table width.
    ///
    /// Validation finishes before the table is changed.
    pub fn set_width_mode(&mut self, width: TableWidth) -> Result<()> {
        let width = checked_table_width("width", width)?;
        self.ensure_tbl_pr().width = Some(width);
        Ok(())
    }

    /// Set or remove the floating table position written to `w:tblpPr`.
    ///
    /// The position is validated before the table is changed, so an invalid
    /// value leaves the document bytes untouched.
    pub fn set_float_position(&mut self, position: Option<TableFloatPosition>) -> Result<()> {
        let resolved = position.map(float_position_to_ct).transpose()?;
        self.ensure_tbl_pr().float_position = resolved.map(Box::new);
        Ok(())
    }

    /// Set or remove the float overlap policy written to `w:tblOverlap`.
    pub fn set_overlap(&mut self, overlap: Option<TableOverlap>) {
        self.ensure_tbl_pr().overlap = overlap.map(TableOverlap::to_st);
    }

    /// Set or remove the bidirectional visual column order.
    pub fn set_bidi_visual(&mut self, value: Option<bool>) {
        self.ensure_tbl_pr().bidi_visual = value;
    }

    /// Set or remove the gap between adjacent cell content boxes.
    pub fn set_cell_spacing(&mut self, spacing: Option<Length>) -> Result<()> {
        let resolved = spacing
            .map(|value| checked_table_twips("cell spacing", value))
            .transpose()?;
        self.ensure_tbl_pr().cell_spacing = resolved.map(CT_TblWidth::dxa);
        Ok(())
    }

    /// Set or remove the accessible table caption.
    pub fn set_caption(&mut self, caption: Option<&str>) -> Result<()> {
        let resolved = caption
            .map(|value| checked_table_text("caption", value))
            .transpose()?;
        self.ensure_tbl_pr().caption = resolved;
        Ok(())
    }

    /// Set or remove the accessible table description.
    pub fn set_description(&mut self, description: Option<&str>) -> Result<()> {
        let resolved = description
            .map(|value| checked_table_text("description", value))
            .transpose()?;
        self.ensure_tbl_pr().description = resolved;
        Ok(())
    }

    /// Set the table indentation from the left margin in place.
    pub fn set_indent(&mut self, length: Length) {
        self.ensure_tbl_pr().indent = Some(CT_TblWidth::dxa(length.as_twips().0));
    }

    /// Set a checked nonnegative table indentation.
    pub fn set_indent_checked(&mut self, length: Length) -> Result<()> {
        let twips = checked_table_twips("indentation", length)?;
        self.ensure_tbl_pr().indent = Some(CT_TblWidth::dxa(twips));
        Ok(())
    }

    /// Set the table width as a percentage (0–100).
    pub fn width_pct(mut self, percent: f64) -> Self {
        self.set_width_pct(percent);
        self
    }

    /// Set the table width as a percentage in place.
    pub fn set_width_pct(&mut self, percent: f64) {
        // OOXML uses 50ths of a percent
        self.ensure_tbl_pr().width = Some(CT_TblWidth::pct((percent * 50.0) as i32));
    }

    /// Set table alignment.
    pub fn alignment(mut self, jc: crate::paragraph::Alignment) -> Self {
        self.set_alignment(jc);
        self
    }

    /// Set table alignment in place.
    pub fn set_alignment(&mut self, jc: crate::paragraph::Alignment) {
        use crate::paragraph::Alignment;
        let st_jc = match jc {
            Alignment::Left => ST_Jc::Left,
            Alignment::Center => ST_Jc::Center,
            Alignment::Right => ST_Jc::Right,
            Alignment::Justify => ST_Jc::Both,
        };
        self.ensure_tbl_pr().jc = Some(st_jc);
    }

    /// Set borders on all edges and internal gridlines.
    pub fn borders(mut self, style: crate::BorderStyle, size_eighths_pt: u32, color: &str) -> Self {
        self.set_borders(style, size_eighths_pt, color);
        self
    }

    /// Set borders on all edges and internal gridlines in place.
    pub fn set_borders(&mut self, style: crate::BorderStyle, size_eighths_pt: u32, color: &str) {
        let edge = CT_BorderEdge {
            val: style.to_st(),
            sz: Some(size_eighths_pt),
            space: Some(0),
            color: Some(color.to_string()),
            extra_attributes: Vec::new(),
            nil: false,
        };
        let borders = self
            .ensure_tbl_pr()
            .borders
            .get_or_insert_with(CT_TblBorders::default);
        borders.top = Some(edge.clone());
        borders.bottom = Some(edge.clone());
        borders.left = Some(edge.clone());
        borders.right = Some(edge.clone());
        borders.inside_h = Some(edge.clone());
        borders.inside_v = Some(edge);
    }

    /// Set every table edge after validating its width and color.
    pub fn set_all_borders_checked(
        &mut self,
        style: crate::BorderStyle,
        size_eighths_pt: u32,
        color: &str,
    ) -> Result<()> {
        let edge = checked_table_border(style, size_eighths_pt, color)?;
        let borders = self
            .ensure_tbl_pr()
            .borders
            .get_or_insert_with(CT_TblBorders::default);
        borders.top = Some(edge.clone());
        borders.bottom = Some(edge.clone());
        borders.left = Some(edge.clone());
        borders.right = Some(edge.clone());
        borders.inside_h = Some(edge.clone());
        borders.inside_v = Some(edge);
        Ok(())
    }

    /// Set one explicit table edge without replacing the other edges or raw
    /// producer extensions.
    pub fn set_border_checked(
        &mut self,
        position: TableBorderEdge,
        style: crate::BorderStyle,
        size_eighths_pt: u32,
        color: &str,
    ) -> Result<()> {
        let edge = checked_table_border(style, size_eighths_pt, color)?;
        let borders = self
            .ensure_tbl_pr()
            .borders
            .get_or_insert_with(CT_TblBorders::default);
        match position {
            TableBorderEdge::Top => borders.top = Some(edge),
            TableBorderEdge::Bottom => borders.bottom = Some(edge),
            TableBorderEdge::Left => borders.left = Some(edge),
            TableBorderEdge::Right => borders.right = Some(edge),
            TableBorderEdge::InsideHorizontal => borders.inside_h = Some(edge),
            TableBorderEdge::InsideVertical => borders.inside_v = Some(edge),
        }
        Ok(())
    }

    /// Set default cell margins.
    pub fn cell_margins(
        mut self,
        top: Length,
        right: Length,
        bottom: Length,
        left: Length,
    ) -> Self {
        self.set_cell_margins(top, right, bottom, left);
        self
    }

    /// Set default cell margins in place.
    pub fn set_cell_margins(&mut self, top: Length, right: Length, bottom: Length, left: Length) {
        self.ensure_tbl_pr().cell_margin = Some(CT_TblCellMar {
            top: Some(top.as_twips()),
            right: Some(right.as_twips()),
            bottom: Some(bottom.as_twips()),
            left: Some(left.as_twips()),
        });
    }

    /// Set checked nonnegative default cell margins.
    pub fn set_cell_margins_checked(
        &mut self,
        top: Length,
        right: Length,
        bottom: Length,
        left: Length,
    ) -> Result<()> {
        let margins = CT_TblCellMar {
            top: Some(rdocx_oxml::Twips(checked_table_twips(
                "top cell margin",
                top,
            )?)),
            right: Some(rdocx_oxml::Twips(checked_table_twips(
                "right cell margin",
                right,
            )?)),
            bottom: Some(rdocx_oxml::Twips(checked_table_twips(
                "bottom cell margin",
                bottom,
            )?)),
            left: Some(rdocx_oxml::Twips(checked_table_twips(
                "left cell margin",
                left,
            )?)),
        };
        self.ensure_tbl_pr().cell_margin = Some(margins);
        Ok(())
    }

    /// Set the table layout to fixed or auto.
    pub fn layout_fixed(mut self) -> Self {
        self.set_layout_fixed();
        self
    }

    /// Set the table layout to fixed in place.
    pub fn set_layout_fixed(&mut self) {
        self.ensure_tbl_pr().layout = Some("fixed".to_string());
    }

    /// Set the complete table layout mode.
    pub fn set_layout(&mut self, layout: TableLayout) {
        self.ensure_tbl_pr().layout = Some(
            match layout {
                TableLayout::AutoFit => "autofit",
                TableLayout::Fixed => "fixed",
            }
            .to_owned(),
        );
    }

    /// Set checked table shading.
    pub fn set_shading_checked(&mut self, fill_color: &str) -> Result<()> {
        let fill = checked_table_color("shading color", fill_color)?;
        self.ensure_tbl_pr().shading = Some(CT_Shd {
            val: "clear".to_owned(),
            color: Some("auto".to_owned()),
            fill: Some(fill),
            ..Default::default()
        });
        Ok(())
    }

    /// Select the conditional regions supplied by the table style.
    ///
    /// Both forms are written. `Table::look` still reads the legacy `w:val`
    /// bitmask as a fallback, and Word writes both, so writing only the
    /// booleans leaves the two forms free to disagree.
    pub fn set_look(&mut self, look: TableLook) {
        self.ensure_tbl_pr().look = Some(CT_TblLook {
            val: Some(look.to_mask()),
            first_row: Some(look.first_row),
            last_row: Some(look.last_row),
            first_column: Some(look.first_column),
            last_column: Some(look.last_column),
            no_h_band: Some(!look.horizontal_banding),
            no_v_band: Some(!look.vertical_banding),
        });
    }

    /// Remove the conditional region selection.
    pub fn clear_look(&mut self) {
        if let Some(properties) = self.inner.properties.as_mut() {
            properties.look = None;
        }
    }

    /// Set the number of rows in each horizontal conditional band.
    ///
    /// A band is a count of rows, not a length, so no unit conversion applies.
    /// Zero is rejected because a band of no rows selects nothing.
    pub fn set_row_band_size(&mut self, rows: u32) -> Result<()> {
        let rows = checked_band_size("table style row band size", rows)?;
        self.ensure_tbl_pr().row_band_size = Some(rows);
        Ok(())
    }

    /// Set the number of columns in each vertical conditional band.
    pub fn set_column_band_size(&mut self, columns: u32) -> Result<()> {
        let columns = checked_band_size("table style column band size", columns)?;
        self.ensure_tbl_pr().column_band_size = Some(columns);
        Ok(())
    }

    /// Remove both conditional band sizes, restoring the one-row default.
    pub fn clear_band_sizes(&mut self) {
        if let Some(properties) = self.inner.properties.as_mut() {
            properties.row_band_size = None;
            properties.column_band_size = None;
        }
    }

    /// Replace the complete active grid and synchronize the fixed table width
    /// and every covering cell width.
    ///
    /// The input must cover the existing grid exactly. Invalid lengths, row
    /// spans, omissions, and sums are rejected before mutation.
    pub fn set_grid_widths(&mut self, widths: &[Length]) -> Result<()> {
        let Some(grid) = self.inner.grid.as_ref() else {
            return Err(Error::Other("table has no active grid".to_owned()));
        };
        if widths.len() != grid.columns.len() || widths.is_empty() {
            return Err(Error::Other(format!(
                "table grid requires exactly {} positive column widths",
                grid.columns.len()
            )));
        }

        let widths = widths
            .iter()
            .enumerate()
            .map(|(index, width)| {
                let value = checked_table_twips(&format!("grid column {index}"), *width)?;
                if value == 0 {
                    return Err(Error::Other(format!(
                        "table grid column {index} must be positive"
                    )));
                }
                Ok(value)
            })
            .collect::<Result<Vec<_>>>()?;
        let table_width = widths.iter().try_fold(0_i32, |total, width| {
            total.checked_add(*width).ok_or_else(|| {
                Error::Other("table grid width exceeds the signed twip range".to_owned())
            })
        })?;

        let mut cell_widths = Vec::with_capacity(self.inner.rows.len());
        for (row_index, row) in self.inner.rows.iter().enumerate() {
            let before = row
                .properties
                .as_ref()
                .and_then(|properties| properties.grid_before)
                .unwrap_or(0) as usize;
            let after = row
                .properties
                .as_ref()
                .and_then(|properties| properties.grid_after)
                .unwrap_or(0) as usize;
            let limit = widths.len().checked_sub(after).ok_or_else(|| {
                Error::Other(format!(
                    "table row {row_index} grid omissions exceed the grid"
                ))
            })?;
            if before > limit {
                return Err(Error::Other(format!(
                    "table row {row_index} grid omissions exceed the grid"
                )));
            }
            let mut grid_index = before;
            let mut row_widths = Vec::with_capacity(row.cells.len());
            for (cell_index, cell) in row.cells.iter().enumerate() {
                let span = cell
                    .properties
                    .as_ref()
                    .and_then(|properties| properties.grid_span)
                    .unwrap_or(1);
                if span == 0 {
                    return Err(Error::Other(format!(
                        "table row {row_index} cell {cell_index} has a zero grid span"
                    )));
                }
                let span = span as usize;
                let end = grid_index.checked_add(span).ok_or_else(|| {
                    Error::Other(format!(
                        "table row {row_index} cell {cell_index} span overflows"
                    ))
                })?;
                if end > limit {
                    return Err(Error::Other(format!(
                        "table row {row_index} cells exceed the active grid"
                    )));
                }
                let cell_width =
                    widths[grid_index..end]
                        .iter()
                        .try_fold(0_i32, |total, width| {
                            total.checked_add(*width).ok_or_else(|| {
                                Error::Other(format!(
                                    "table row {row_index} cell {cell_index} width overflows"
                                ))
                            })
                        })?;
                row_widths.push(cell_width);
                grid_index = end;
            }
            if grid_index != limit {
                return Err(Error::Other(format!(
                    "table row {row_index} cells do not cover the active grid"
                )));
            }
            cell_widths.push(row_widths);
        }

        for (column, width) in self
            .inner
            .grid
            .as_mut()
            .expect("active grid was validated")
            .columns
            .iter_mut()
            .zip(&widths)
        {
            column.width = rdocx_oxml::Twips(*width);
        }
        self.ensure_tbl_pr().width = Some(CT_TblWidth::dxa(table_width));
        for (row, row_widths) in self.inner.rows.iter_mut().zip(cell_widths) {
            for (cell, width) in row.cells.iter_mut().zip(row_widths) {
                cell.properties.get_or_insert_with(CT_TcPr::default).width =
                    Some(CT_TblWidth::dxa(width));
            }
        }
        Ok(())
    }

    /// Set the leading and trailing grid omissions for one row.
    ///
    /// Increasing an omission removes only untouched empty edge cells.
    /// Decreasing it inserts empty cells with the corresponding grid widths.
    /// The complete table is validated before the candidate replaces the live
    /// value.
    pub fn set_row_grid_omissions(
        &mut self,
        row_index: usize,
        before: Option<u32>,
        after: Option<u32>,
    ) -> Result<()> {
        let mut candidate = self.inner.clone();
        let grid = candidate
            .grid
            .as_ref()
            .ok_or_else(|| Error::Other("table has no active grid".to_owned()))?;
        let grid_widths = grid
            .columns
            .iter()
            .map(|column| column.width.0)
            .collect::<Vec<_>>();
        row_cell_ranges(
            candidate
                .rows
                .get(row_index)
                .ok_or_else(|| Error::Other(format!("table row {row_index} does not exist")))?,
            grid_widths.len(),
        )?;
        let row = candidate
            .rows
            .get_mut(row_index)
            .ok_or_else(|| Error::Other(format!("table row {row_index} does not exist")))?;
        let current_before = row
            .properties
            .as_ref()
            .and_then(|properties| properties.grid_before)
            .unwrap_or(0) as usize;
        let current_after = row
            .properties
            .as_ref()
            .and_then(|properties| properties.grid_after)
            .unwrap_or(0) as usize;
        let target_before = before.unwrap_or(0) as usize;
        let target_after = after.unwrap_or(0) as usize;
        if target_before
            .checked_add(target_after)
            .is_none_or(|omitted| omitted >= grid_widths.len())
        {
            return Err(Error::Other(
                "table row omissions must leave at least one active column".to_owned(),
            ));
        }

        if target_before > current_before {
            for _ in 0..(target_before - current_before) {
                let cell = row.cells.first().ok_or_else(|| {
                    Error::Other("table row has no leading cell to omit".to_owned())
                })?;
                if !cell_is_discardable_for_grid_edit(cell) {
                    return Err(Error::Other(
                        "table row cannot omit a nonempty leading cell".to_owned(),
                    ));
                }
                row.cells.remove(0);
            }
        } else {
            for column in (target_before..current_before).rev() {
                row.cells.insert(0, empty_grid_cell(grid_widths[column]));
            }
        }

        if target_after > current_after {
            for _ in 0..(target_after - current_after) {
                let cell = row.cells.last().ok_or_else(|| {
                    Error::Other("table row has no trailing cell to omit".to_owned())
                })?;
                if !cell_is_discardable_for_grid_edit(cell) {
                    return Err(Error::Other(
                        "table row cannot omit a nonempty trailing cell".to_owned(),
                    ));
                }
                row.cells.pop();
            }
        } else {
            let first_column = grid_widths.len() - current_after;
            let last_column = grid_widths.len() - target_after;
            for width in &grid_widths[first_column..last_column] {
                row.cells.push(empty_grid_cell(*width));
            }
        }

        let properties = row.properties.get_or_insert_with(CT_TrPr::default);
        properties.grid_before = before;
        properties.grid_after = after;
        validate_table_topology(&candidate)?;
        *self.inner = candidate;
        Ok(())
    }

    /// Set one cell's horizontal grid span and reconcile untouched cells that
    /// enter or leave the span.
    pub fn set_cell_grid_span_checked(
        &mut self,
        row_index: usize,
        cell_index: usize,
        span: Option<u32>,
    ) -> Result<()> {
        let desired = span.unwrap_or(1) as usize;
        if desired == 0 {
            return Err(Error::Other(
                "table cell grid span must be positive".to_owned(),
            ));
        }
        let mut candidate = self.inner.clone();
        let grid_columns = candidate
            .grid
            .as_ref()
            .map(|grid| grid.columns.len())
            .filter(|count| *count > 0)
            .ok_or_else(|| Error::Other("table has no active grid".to_owned()))?;
        let ranges = row_cell_ranges(
            candidate
                .rows
                .get(row_index)
                .ok_or_else(|| Error::Other(format!("table row {row_index} does not exist")))?,
            grid_columns,
        )?;
        let (start, end) = *ranges.get(cell_index).ok_or_else(|| {
            Error::Other(format!(
                "table row {row_index} cell {cell_index} does not exist"
            ))
        })?;
        let current = end - start;
        let grid_widths = candidate
            .grid
            .as_ref()
            .expect("active grid was validated")
            .columns
            .iter()
            .map(|column| column.width.0)
            .collect::<Vec<_>>();
        let desired_end = start
            .checked_add(desired)
            .filter(|end| *end <= grid_widths.len())
            .ok_or_else(|| Error::Other("horizontal merge exceeds the row grid".to_owned()))?;
        let desired_width = grid_widths[start..desired_end]
            .iter()
            .try_fold(0_i32, |total, width| total.checked_add(*width))
            .ok_or_else(|| Error::Other("horizontal merge width overflows".to_owned()))?;
        let row = &mut candidate.rows[row_index];
        if desired > current {
            let mut covered = current;
            while covered < desired {
                let next = row.cells.get(cell_index + 1).ok_or_else(|| {
                    Error::Other("horizontal merge exceeds the row grid".to_owned())
                })?;
                if !cell_is_discardable_for_grid_edit(next) {
                    return Err(Error::Other(
                        "horizontal merge cannot consume a nonempty cell".to_owned(),
                    ));
                }
                let next_span = next
                    .properties
                    .as_ref()
                    .and_then(|properties| properties.grid_span)
                    .unwrap_or(1) as usize;
                covered = covered
                    .checked_add(next_span)
                    .ok_or_else(|| Error::Other("horizontal merge span overflows".to_owned()))?;
                if covered > desired {
                    return Err(Error::Other(
                        "horizontal merge cuts through an existing span".to_owned(),
                    ));
                }
                row.cells.remove(cell_index + 1);
            }
        } else if desired < current {
            for (offset, width) in grid_widths[(start + desired)..end].iter().enumerate() {
                row.cells
                    .insert(cell_index + 1 + offset, empty_grid_cell(*width));
            }
        }
        let properties = row.cells[cell_index]
            .properties
            .get_or_insert_with(CT_TcPr::default);
        properties.grid_span = (desired > 1).then_some(desired as u32);
        properties.width = Some(CT_TblWidth::dxa(desired_width));
        validate_table_topology(&candidate)?;
        *self.inner = candidate;
        Ok(())
    }

    /// Set or clear one cell's vertical merge state after validating the
    /// complete merge topology.
    pub fn set_cell_vertical_merge(
        &mut self,
        row_index: usize,
        cell_index: usize,
        merge: Option<VMerge>,
    ) -> Result<()> {
        let mut candidate = self.inner.clone();
        let cell = candidate
            .rows
            .get_mut(row_index)
            .and_then(|row| row.cells.get_mut(cell_index))
            .ok_or_else(|| {
                Error::Other(format!(
                    "table row {row_index} cell {cell_index} does not exist"
                ))
            })?;
        cell.properties.get_or_insert_with(CT_TcPr::default).v_merge = merge;
        validate_table_topology(&candidate)?;
        *self.inner = candidate;
        Ok(())
    }

    /// Set one grid column's width and keep the table, grid, and covering cell
    /// widths synchronized.
    ///
    /// A cell that spans the changed grid column receives the sum of every
    /// grid column it covers. Returns `false` without changing the table when
    /// `column` is outside the grid, `width` is negative, a row's spans exceed
    /// the grid, or a width total overflows.
    pub fn set_column_width(&mut self, column: usize, width: Length) -> bool {
        if width.as_twips().0 < 0 {
            return false;
        }
        let Some(grid) = self.inner.grid.as_ref() else {
            return false;
        };
        if column >= grid.columns.len() {
            return false;
        }

        let mut grid_widths: Vec<i32> = grid.columns.iter().map(|item| item.width.0).collect();
        grid_widths[column] = width.as_twips().0;
        let Some(table_width) = grid_widths
            .iter()
            .try_fold(0_i32, |total, item| total.checked_add(*item))
        else {
            return false;
        };

        let mut cell_widths = Vec::with_capacity(self.inner.rows.len());
        for row in &self.inner.rows {
            let mut grid_index = 0_usize;
            let mut row_widths = Vec::with_capacity(row.cells.len());
            for cell in &row.cells {
                let span = cell
                    .properties
                    .as_ref()
                    .and_then(|properties| properties.grid_span)
                    .unwrap_or(1)
                    .max(1) as usize;
                let Some(end) = grid_index.checked_add(span) else {
                    return false;
                };
                if end > grid_widths.len() {
                    return false;
                }
                if (grid_index..end).contains(&column) {
                    let Some(cell_width) = grid_widths[grid_index..end]
                        .iter()
                        .try_fold(0_i32, |total, item| total.checked_add(*item))
                    else {
                        return false;
                    };
                    row_widths.push(Some(cell_width));
                } else {
                    row_widths.push(None);
                }
                grid_index = end;
            }
            cell_widths.push(row_widths);
        }

        self.inner.grid.as_mut().unwrap().columns[column].width = width.as_twips();
        self.ensure_tbl_pr().width = Some(CT_TblWidth::dxa(table_width));
        for (row, widths) in self.inner.rows.iter_mut().zip(cell_widths) {
            for (cell, cell_width) in row.cells.iter_mut().zip(widths) {
                if let Some(cell_width) = cell_width {
                    cell.properties.get_or_insert_with(CT_TcPr::default).width =
                        Some(CT_TblWidth::dxa(cell_width));
                }
            }
        }
        true
    }

    /// Get the number of rows.
    pub fn row_count(&self) -> usize {
        self.inner.rows.len()
    }

    /// Get a mutable reference to a row by index.
    pub fn row(&mut self, index: usize) -> Option<Row<'_>> {
        self.inner.rows.get_mut(index).map(|r| Row { inner: r })
    }

    /// Get a mutable reference to a cell at (row, col).
    pub fn cell(&mut self, row: usize, col: usize) -> Option<Cell<'_>> {
        self.inner
            .rows
            .get_mut(row)
            .and_then(|r| r.cells.get_mut(col))
            .map(|c| Cell { inner: c })
    }

    fn ensure_tbl_pr(&mut self) -> &mut CT_TblPr {
        self.inner.properties.get_or_insert_with(CT_TblPr::default)
    }
}

// ---- Mutable Row ----

/// A mutable reference to a table row.
pub struct Row<'a> {
    pub(crate) inner: &'a mut CT_Row,
}

impl<'a> Row<'a> {
    /// Set the row height.
    pub fn height(mut self, length: Length) -> Self {
        self.set_height(length);
        self
    }

    /// Set the row height in place.
    pub fn set_height(&mut self, length: Length) {
        let pr = self.ensure_tr_pr();
        pr.height = Some(length.as_twips());
        pr.height_rule = Some("atLeast".to_string());
    }

    /// Set exact row height.
    pub fn height_exact(mut self, length: Length) -> Self {
        self.set_height_exact(length);
        self
    }

    /// Set exact row height in place.
    pub fn set_height_exact(&mut self, length: Length) {
        let pr = self.ensure_tr_pr();
        pr.height = Some(length.as_twips());
        pr.height_rule = Some("exact".to_string());
    }

    /// Set a checked minimum or exact row height.
    pub fn set_height_checked(&mut self, height: RowHeight) -> Result<()> {
        let (length, rule) = match height {
            RowHeight::AtLeast(length) => (length, "atLeast"),
            RowHeight::Exact(length) => (length, "exact"),
        };
        let twips = checked_table_twips("row height", length)?;
        let properties = self.ensure_tr_pr();
        properties.height = Some(rdocx_oxml::Twips(twips));
        properties.height_rule = Some(rule.to_owned());
        Ok(())
    }

    /// Mark this row as a header row (repeats on each page).
    pub fn header(mut self) -> Self {
        self.set_header();
        self
    }

    /// Mark this row as a header row in place.
    pub fn set_header(&mut self) {
        self.ensure_tr_pr().header = Some(true);
    }

    /// Set an explicit header toggle or remove the direct value.
    pub fn set_header_value(&mut self, value: Option<bool>) {
        self.ensure_tr_pr().header = value;
    }

    /// Prevent this row from splitting across pages.
    pub fn cant_split(mut self) -> Self {
        self.set_cant_split();
        self
    }

    /// Prevent this row from splitting across pages in place.
    pub fn set_cant_split(&mut self) {
        self.ensure_tr_pr().cant_split = Some(true);
    }

    /// Set an explicit split-policy toggle or remove the direct value.
    pub fn set_cant_split_value(&mut self, value: Option<bool>) {
        self.ensure_tr_pr().cant_split = value;
    }

    /// Set direct row alignment.
    pub fn set_alignment(&mut self, alignment: crate::paragraph::Alignment) {
        use crate::paragraph::Alignment;
        self.ensure_tr_pr().jc = Some(match alignment {
            Alignment::Left => ST_Jc::Left,
            Alignment::Center => ST_Jc::Center,
            Alignment::Right => ST_Jc::Right,
            Alignment::Justify => ST_Jc::Both,
        });
    }

    /// Set direct conditional table-style regions on this row.
    pub fn set_conditional_formatting(&mut self, regions: TableConditionalFormatting) {
        self.ensure_tr_pr().cnf_style = Some(regions.to_value());
    }

    /// Set or remove the width of this row's omitted leading grid columns.
    pub fn set_width_before(&mut self, width: Option<TableWidth>) -> Result<()> {
        let resolved = width
            .map(|value| checked_table_width("leading row width", value))
            .transpose()?;
        self.ensure_tr_pr().width_before = resolved;
        Ok(())
    }

    /// Set or remove the width of this row's omitted trailing grid columns.
    pub fn set_width_after(&mut self, width: Option<TableWidth>) -> Result<()> {
        let resolved = width
            .map(|value| checked_table_width("trailing row width", value))
            .transpose()?;
        self.ensure_tr_pr().width_after = resolved;
        Ok(())
    }

    /// Set or remove this row's gap between adjacent cell content boxes.
    pub fn set_cell_spacing(&mut self, spacing: Option<Length>) -> Result<()> {
        let resolved = spacing
            .map(|value| checked_table_twips("row cell spacing", value))
            .transpose()?;
        self.ensure_tr_pr().cell_spacing = resolved.map(CT_TblWidth::dxa);
        Ok(())
    }

    /// Set or remove the hidden toggle written to `w:hidden`.
    pub fn set_hidden(&mut self, value: Option<bool>) {
        self.ensure_tr_pr().hidden = value;
    }

    /// Get a mutable reference to a cell by index.
    pub fn cell(&mut self, index: usize) -> Option<Cell<'_>> {
        self.inner.cells.get_mut(index).map(|c| Cell { inner: c })
    }

    /// Get the number of cells in this row.
    pub fn cell_count(&self) -> usize {
        self.inner.cells.len()
    }

    fn ensure_tr_pr(&mut self) -> &mut CT_TrPr {
        self.inner.properties.get_or_insert_with(CT_TrPr::default)
    }
}

// ---- Mutable Cell ----

/// A mutable reference to a table cell.
pub struct Cell<'a> {
    pub(crate) inner: &'a mut CT_Tc,
}

impl<'a> Cell<'a> {
    /// Get the combined text of all paragraphs in this cell.
    pub fn text(&self) -> String {
        self.inner.text()
    }

    /// Set the text of the first paragraph (replacing existing content).
    pub fn set_text(&mut self, text: &str) {
        use rdocx_oxml::table::CellContent;
        // Find first paragraph or create one
        let first_para = self.inner.content.iter_mut().find_map(|c| {
            if let CellContent::Paragraph(p) = c {
                Some(p)
            } else {
                None
            }
        });
        if let Some(para) = first_para {
            para.runs.clear();
            if !text.is_empty() {
                para.add_run(text);
            }
        } else {
            let mut p = CT_P::new();
            if !text.is_empty() {
                p.add_run(text);
            }
            self.inner.content.insert(0, CellContent::Paragraph(p));
        }
    }

    /// Add a paragraph to the cell and return a mutable reference.
    pub fn add_paragraph(&mut self, text: &str) -> Paragraph<'_> {
        use rdocx_oxml::table::CellContent;
        let mut p = CT_P::new();
        if !text.is_empty() {
            p.add_run(text);
        }
        self.inner.content.push(CellContent::Paragraph(p));
        let para = self.inner.content.last_mut().unwrap();
        if let CellContent::Paragraph(p) = para {
            Paragraph { inner: p }
        } else {
            unreachable!()
        }
    }

    /// Add an inline image to the cell using a pre-embedded relationship ID.
    ///
    /// Obtain the `rel_id` by calling [`crate::Document::embed_image`] first, then
    /// pass it here along with the desired display dimensions. This matches
    /// the python-docx `run.add_picture()` pattern.
    pub fn add_picture(&mut self, rel_id: &str, width: Length, height: Length) {
        use rdocx_oxml::drawing::{CT_Drawing, CT_Inline};
        use rdocx_oxml::table::CellContent;
        use rdocx_oxml::text::{CT_R, RunContent};

        let inline = CT_Inline::new(rel_id, width.to_emu(), height.to_emu());
        let drawing = CT_Drawing::inline(inline);
        let run = CT_R {
            alt_drawings: Vec::new(),
            properties: None,
            content: vec![RunContent::Drawing(drawing)],
            extra_xml: Vec::new(),
            extra_xml_positions: Vec::new(),
        };
        let mut p = CT_P::new();
        p.runs.push(run);
        self.inner.content.push(CellContent::Paragraph(p));
    }

    /// Remove the first empty paragraph from the cell.
    ///
    /// OOXML creates a default empty paragraph when a cell is instantiated.
    /// Call this before adding content to avoid a spurious blank line at the
    /// top of the cell — mirrors the `add_html_block` behaviour in python-docx.
    pub fn remove_first_empty_paragraph(&mut self) {
        use rdocx_oxml::table::CellContent;
        if let Some(pos) = self.inner.content.iter().position(|c| {
            if let CellContent::Paragraph(p) = c {
                p.text().trim().is_empty()
            } else {
                false
            }
        }) {
            self.inner.content.remove(pos);
        }
    }

    /// Get an iterator over immutable paragraph references.
    pub fn paragraphs(&self) -> impl Iterator<Item = ParagraphRef<'_>> {
        self.inner
            .paragraphs()
            .into_iter()
            .map(|p| ParagraphRef { inner: p })
    }

    /// Get the number of paragraphs in the cell.
    pub fn paragraph_count(&self) -> usize {
        self.inner.paragraphs().len()
    }

    /// Get an immutable paragraph by index.
    pub fn paragraph(&self, index: usize) -> Option<ParagraphRef<'_>> {
        self.inner
            .paragraphs()
            .get(index)
            .map(|inner| ParagraphRef { inner })
    }

    /// Get a mutable paragraph by index.
    pub fn paragraph_mut(&mut self, index: usize) -> Option<Paragraph<'_>> {
        self.inner
            .paragraphs_mut()
            .into_iter()
            .nth(index)
            .map(|inner| Paragraph { inner })
    }

    /// Set cell width.
    pub fn width(mut self, length: Length) -> Self {
        self.set_width(length);
        self
    }

    /// Set cell width in place.
    pub fn set_width(&mut self, length: Length) {
        self.ensure_tc_pr().width = Some(CT_TblWidth::dxa(length.as_twips().0));
    }

    /// Set a checked nonnegative cell width.
    pub fn set_width_checked(&mut self, length: Length) -> Result<()> {
        let width = checked_table_twips("cell width", length)?;
        self.ensure_tc_pr().width = Some(CT_TblWidth::dxa(width));
        Ok(())
    }

    /// Set cell background shading color.
    pub fn shading(mut self, fill_color: &str) -> Self {
        self.set_shading(fill_color);
        self
    }

    /// Set cell background shading color in place.
    pub fn set_shading(&mut self, fill_color: &str) {
        self.ensure_tc_pr().shading = Some(CT_Shd {
            val: "clear".to_string(),
            color: Some("auto".to_string()),
            fill: Some(fill_color.to_string()),
            ..Default::default()
        });
    }

    /// Set checked cell shading.
    pub fn set_shading_checked(&mut self, fill_color: &str) -> Result<()> {
        let fill = checked_table_color("cell shading color", fill_color)?;
        self.ensure_tc_pr().shading = Some(CT_Shd {
            val: "clear".to_owned(),
            color: Some("auto".to_owned()),
            fill: Some(fill),
            ..Default::default()
        });
        Ok(())
    }

    /// Set one cell border without replacing other edges or raw extensions.
    pub fn set_border_checked(
        &mut self,
        position: CellBorderEdge,
        style: crate::BorderStyle,
        size_eighths_pt: u32,
        color: &str,
    ) -> Result<()> {
        let edge = checked_table_border(style, size_eighths_pt, color)?;
        let borders = self
            .ensure_tc_pr()
            .borders
            .get_or_insert_with(CT_TblBorders::default);
        match position {
            CellBorderEdge::Top => borders.top = Some(edge),
            CellBorderEdge::Bottom => borders.bottom = Some(edge),
            CellBorderEdge::Left => borders.left = Some(edge),
            CellBorderEdge::Right => borders.right = Some(edge),
            CellBorderEdge::InsideHorizontal => borders.inside_h = Some(edge),
            CellBorderEdge::InsideVertical => borders.inside_v = Some(edge),
        }
        Ok(())
    }

    /// Set checked nonnegative per-cell margins.
    pub fn set_margins_checked(
        &mut self,
        top: Length,
        right: Length,
        bottom: Length,
        left: Length,
    ) -> Result<()> {
        let margins = CT_TblCellMar {
            top: Some(rdocx_oxml::Twips(checked_table_twips(
                "top cell margin",
                top,
            )?)),
            right: Some(rdocx_oxml::Twips(checked_table_twips(
                "right cell margin",
                right,
            )?)),
            bottom: Some(rdocx_oxml::Twips(checked_table_twips(
                "bottom cell margin",
                bottom,
            )?)),
            left: Some(rdocx_oxml::Twips(checked_table_twips(
                "left cell margin",
                left,
            )?)),
        };
        self.ensure_tc_pr().cell_margin = Some(margins);
        Ok(())
    }

    /// Set vertical alignment within the cell.
    pub fn vertical_alignment(mut self, align: VerticalAlignment) -> Self {
        self.set_vertical_alignment(align);
        self
    }

    /// Set vertical alignment within the cell in place.
    pub fn set_vertical_alignment(&mut self, align: VerticalAlignment) {
        self.ensure_tc_pr().v_align = Some(align.to_st());
    }

    /// Set horizontal merge (gridSpan). This cell spans `span` columns.
    pub fn grid_span(mut self, span: u32) -> Self {
        self.set_grid_span(span);
        self
    }

    /// Set horizontal merge span in place.
    pub fn set_grid_span(&mut self, span: u32) {
        self.ensure_tc_pr().grid_span = Some(span);
    }

    /// Start a vertical merge group (this cell is the top of the merged range).
    pub fn v_merge_restart(mut self) -> Self {
        self.set_v_merge_restart();
        self
    }

    /// Start a vertical merge group in place.
    pub fn set_v_merge_restart(&mut self) {
        self.ensure_tc_pr().v_merge = Some(VMerge::Restart);
    }

    /// Continue a vertical merge group (this cell merges with the one above).
    pub fn v_merge_continue(mut self) -> Self {
        self.set_v_merge_continue();
        self
    }

    /// Continue a vertical merge group in place.
    pub fn set_v_merge_continue(&mut self) {
        self.ensure_tc_pr().v_merge = Some(VMerge::Continue);
    }

    /// Set no-wrap for text in this cell.
    pub fn no_wrap(mut self) -> Self {
        self.set_no_wrap();
        self
    }

    /// Set no-wrap for text in this cell in place.
    pub fn set_no_wrap(&mut self) {
        self.ensure_tc_pr().no_wrap = Some(true);
    }

    /// Set an explicit no-wrap toggle or remove the direct value.
    pub fn set_no_wrap_value(&mut self, value: Option<bool>) {
        self.ensure_tc_pr().no_wrap = value;
    }

    /// Set or clear the direct cell text direction.
    pub fn set_text_direction(&mut self, direction: Option<CellTextDirection>) {
        self.ensure_tc_pr().text_direction = direction.map(|value| value.to_str().to_owned());
    }

    /// Set direct conditional table-style regions on this cell.
    pub fn set_conditional_formatting(&mut self, regions: TableConditionalFormatting) {
        self.ensure_tc_pr().cnf_style = Some(regions.to_value());
    }

    /// Add a nested table inside this cell.
    pub fn add_table(&mut self, rows: usize, cols: usize) -> Table<'_> {
        use rdocx_oxml::table::{
            CT_Row, CT_Tbl, CT_TblGrid, CT_TblGridCol, CT_TblPr, CT_TblWidth, CT_Tc, CellContent,
        };
        use rdocx_oxml::units::Twips;

        // Default nested table column width: use equal splits of 4500tw (~3.125").
        // Clamped so a zero-column request cannot divide by zero.
        let col_width = Twips(4500 / cols.max(1) as i32);

        let grid = CT_TblGrid {
            columns: (0..cols)
                .map(|_| CT_TblGridCol { width: col_width })
                .collect(),
            ..Default::default()
        };

        let mut tbl = CT_Tbl::new();
        tbl.properties = Some(CT_TblPr {
            width: Some(CT_TblWidth::dxa(col_width.0 * cols as i32)),
            ..Default::default()
        });
        tbl.grid = Some(grid);

        for _ in 0..rows {
            let mut row = CT_Row::new();
            for _ in 0..cols {
                row.cells.push(CT_Tc::new());
            }
            tbl.rows.push(row);
        }

        self.inner.content.push(CellContent::Table(tbl));
        match self.inner.content.last_mut().unwrap() {
            CellContent::Table(t) => Table { inner: t },
            _ => unreachable!(),
        }
    }

    /// Add a nonempty nested table while retaining the required trailing cell
    /// paragraph.
    pub fn add_table_checked(&mut self, rows: usize, cols: usize) -> Result<Table<'_>> {
        use rdocx_oxml::table::{CT_TblGrid, CT_TblGridCol};
        use rdocx_oxml::units::Twips;

        if rows == 0 || cols == 0 {
            return Err(Error::Other(
                "nested table dimensions must be positive".to_owned(),
            ));
        }
        let columns = i32::try_from(cols)
            .map_err(|_| Error::Other("nested table column count is too large".to_owned()))?;
        let col_width = Twips(4500 / columns);
        if col_width.0 == 0 {
            return Err(Error::Other(
                "nested table columns must be at least one twip wide".to_owned(),
            ));
        }

        let mut table = CT_Tbl::new();
        table.properties = Some(CT_TblPr {
            width: Some(CT_TblWidth::dxa(col_width.0 * columns)),
            ..Default::default()
        });
        table.grid = Some(CT_TblGrid {
            columns: (0..cols)
                .map(|_| CT_TblGridCol { width: col_width })
                .collect(),
            ..Default::default()
        });
        for _ in 0..rows {
            let mut row = CT_Row::new();
            row.cells.extend((0..cols).map(|_| CT_Tc::new()));
            table.rows.push(row);
        }
        validate_table_topology(&table)?;

        let table_index = self.inner.content.len();
        self.inner.content.push(CellContent::Table(table));
        self.inner.content.push(CellContent::Paragraph(CT_P::new()));
        match self.inner.content.get_mut(table_index) {
            Some(CellContent::Table(table)) => Ok(Table { inner: table }),
            _ => unreachable!(),
        }
    }

    fn ensure_tc_pr(&mut self) -> &mut CT_TcPr {
        self.inner.properties.get_or_insert_with(CT_TcPr::default)
    }
}

// ---- Immutable references ----

/// An immutable reference to a table.
pub struct TableRef<'a> {
    pub(crate) inner: &'a CT_Tbl,
}

impl<'a> TableRef<'a> {
    /// Get the number of rows.
    pub fn row_count(&self) -> usize {
        self.inner.rows.len()
    }

    /// Get the number of columns (from the grid definition).
    pub fn column_count(&self) -> usize {
        self.inner
            .grid
            .as_ref()
            .map(|g| g.columns.len())
            .unwrap_or(0)
    }

    /// Whether the table contains direct content other than rows.
    pub fn has_unsupported_content(&self) -> bool {
        !self.inner.extra_xml.is_empty() || !self.inner.content_controls.is_empty()
    }

    /// Whether table properties retain facts outside the public reader model.
    pub fn has_unmodeled_properties(&self) -> bool {
        self.inner.properties.as_ref().is_some_and(|properties| {
            properties.change.is_some()
                || !properties.revision_xml.is_empty()
                || !properties.extra_xml.is_empty()
                || properties
                    .borders
                    .as_ref()
                    .is_some_and(|borders| !borders.extra_xml.is_empty())
        }) || self
            .inner
            .grid
            .as_ref()
            .is_some_and(|grid| !grid.extra_xml.is_empty())
    }

    /// Whether the table preserves a historical grid revision.
    ///
    /// The historical grid is round-trip metadata. Active columns remain the
    /// sole input to current table layout.
    pub fn has_grid_change(&self) -> bool {
        self.inner
            .grid
            .as_ref()
            .is_some_and(|grid| grid.grid_change_xml.is_some())
    }

    /// Get an immutable row reference.
    pub fn row(&self, index: usize) -> Option<RowRef<'_>> {
        self.inner.rows.get(index).map(|r| RowRef { inner: r })
    }

    /// Get a cell reference at (row, col).
    pub fn cell(&self, row: usize, col: usize) -> Option<CellRef<'_>> {
        self.inner
            .rows
            .get(row)
            .and_then(|r| r.cells.get(col))
            .map(|c| CellRef { inner: c })
    }

    /// Get the table style ID, if set.
    pub fn style_id(&self) -> Option<&str> {
        self.inner
            .properties
            .as_ref()
            .and_then(|pr| pr.style_id.as_deref())
    }

    /// Get table alignment, if set.
    pub fn alignment(&self) -> Option<crate::paragraph::Alignment> {
        use crate::paragraph::Alignment;
        self.inner
            .properties
            .as_ref()
            .and_then(|pr| pr.jc)
            .map(|value| match value {
                ST_Jc::Center => Alignment::Center,
                ST_Jc::Right | ST_Jc::End => Alignment::Right,
                ST_Jc::Both | ST_Jc::Distribute => Alignment::Justify,
                _ => Alignment::Left,
            })
    }

    /// Get the table width when stored as twips.
    pub fn width(&self) -> Option<Length> {
        let width = self.inner.properties.as_ref()?.width.as_ref()?;
        (width.width_type == "dxa").then(|| Length::twips(width.w))
    }

    /// Get the complete authored table width mode.
    pub fn width_mode(&self) -> Option<TableWidth> {
        table_width_from_ct(self.inner.properties.as_ref()?.width.as_ref()?)
    }

    /// Get the authored floating table position.
    pub fn float_position(&self) -> Option<TableFloatPosition> {
        self.inner
            .properties
            .as_ref()?
            .float_position
            .as_deref()
            .map(float_position_from_ct)
    }

    /// Get the authored float overlap policy.
    pub fn overlap(&self) -> Option<TableOverlap> {
        self.inner
            .properties
            .as_ref()?
            .overlap
            .map(TableOverlap::from_st)
    }

    /// Get the authored bidirectional visual column order.
    pub fn bidi_visual(&self) -> Option<bool> {
        self.inner.properties.as_ref()?.bidi_visual
    }

    /// Get the authored gap between adjacent cell content boxes.
    pub fn cell_spacing(&self) -> Option<TableWidth> {
        table_width_from_ct(self.inner.properties.as_ref()?.cell_spacing.as_ref()?)
    }

    /// Get the accessible table caption.
    pub fn caption(&self) -> Option<&str> {
        self.inner.properties.as_ref()?.caption.as_deref()
    }

    /// Get the accessible table description.
    pub fn description(&self) -> Option<&str> {
        self.inner.properties.as_ref()?.description.as_deref()
    }

    /// Get the authored table indentation when stored as twips.
    pub fn indent(&self) -> Option<Length> {
        let indent = self.inner.properties.as_ref()?.indent.as_ref()?;
        (indent.width_type == "dxa").then(|| Length::twips(indent.w))
    }

    /// Get the authored table layout mode.
    pub fn layout(&self) -> Option<TableLayout> {
        match self.inner.properties.as_ref()?.layout.as_deref()? {
            "fixed" => Some(TableLayout::Fixed),
            "autofit" => Some(TableLayout::AutoFit),
            _ => None,
        }
    }

    /// Get the direct table shading fill.
    pub fn shading_fill(&self) -> Option<&str> {
        self.inner
            .properties
            .as_ref()?
            .shading
            .as_ref()?
            .fill
            .as_deref()
    }

    /// Get one direct table border, including an explicit invisible edge.
    pub fn border(&self, position: TableBorderEdge) -> Option<TableBorderRef<'_>> {
        let borders = self.inner.properties.as_ref()?.borders.as_ref()?;
        let inner = match position {
            TableBorderEdge::Top => borders.top.as_ref(),
            TableBorderEdge::Bottom => borders.bottom.as_ref(),
            TableBorderEdge::Left => borders.left.as_ref(),
            TableBorderEdge::Right => borders.right.as_ref(),
            TableBorderEdge::InsideHorizontal => borders.inside_h.as_ref(),
            TableBorderEdge::InsideVertical => borders.inside_v.as_ref(),
        }?;
        Some(TableBorderRef { inner })
    }

    /// Get the direct default cell margins.
    pub fn cell_margins(&self) -> Option<TableCellMargins> {
        let margins = self.inner.properties.as_ref()?.cell_margin.as_ref()?;
        Some(TableCellMargins {
            top: margins.top.map(|value| Length::twips(value.0)),
            right: margins.right.map(|value| Length::twips(value.0)),
            bottom: margins.bottom.map(|value| Length::twips(value.0)),
            left: margins.left.map(|value| Length::twips(value.0)),
        })
    }

    /// Get the active grid column widths in source order.
    pub fn grid_widths(&self) -> Vec<Length> {
        self.inner
            .grid
            .as_ref()
            .map(|grid| {
                grid.columns
                    .iter()
                    .map(|column| Length::twips(column.width.0))
                    .collect()
            })
            .unwrap_or_default()
    }

    /// Get the number of rows in each horizontal conditional band.
    pub fn row_band_size(&self) -> Option<u32> {
        self.inner.properties.as_ref()?.row_band_size
    }

    /// Get the number of columns in each vertical conditional band.
    pub fn column_band_size(&self) -> Option<u32> {
        self.inner.properties.as_ref()?.column_band_size
    }

    /// Get the selected conditional table-style regions.
    pub fn look(&self) -> Option<TableLook> {
        let look = self.inner.properties.as_ref()?.look.as_ref()?;
        let mask = look
            .val
            .as_deref()
            .and_then(|value| u16::from_str_radix(value, 16).ok());
        let enabled = |explicit: Option<bool>, bit: u16| {
            explicit.unwrap_or_else(|| mask.is_some_and(|value| value & bit != 0))
        };
        Some(TableLook {
            first_row: enabled(look.first_row, 0x20),
            last_row: enabled(look.last_row, 0x40),
            first_column: enabled(look.first_column, 0x80),
            last_column: enabled(look.last_column, 0x100),
            horizontal_banding: !enabled(look.no_h_band, 0x200),
            vertical_banding: !enabled(look.no_v_band, 0x400),
        })
    }

    /// Whether explicit table width is present.
    pub fn has_width(&self) -> bool {
        self.inner
            .properties
            .as_ref()
            .is_some_and(|properties| properties.width.is_some())
    }

    /// Whether explicit table borders are present.
    pub fn has_borders(&self) -> bool {
        self.inner
            .properties
            .as_ref()
            .and_then(|properties| properties.borders.as_ref())
            .is_some_and(|borders| !borders.is_empty())
    }

    /// Whether explicit table shading is present.
    pub fn has_shading(&self) -> bool {
        self.inner
            .properties
            .as_ref()
            .is_some_and(|properties| properties.shading.is_some())
    }

    /// Whether an explicit table layout mode is present.
    pub fn has_layout(&self) -> bool {
        self.inner
            .properties
            .as_ref()
            .is_some_and(|properties| properties.layout.is_some())
    }

    /// Whether default table-cell margins are present.
    pub fn has_cell_margins(&self) -> bool {
        self.inner
            .properties
            .as_ref()
            .is_some_and(|properties| properties.cell_margin.is_some())
    }

    /// Whether an explicit table indentation is present.
    pub fn has_indent(&self) -> bool {
        self.inner
            .properties
            .as_ref()
            .is_some_and(|properties| properties.indent.is_some())
    }

    /// Whether table-style conditional formatting flags are present.
    pub fn has_style_look(&self) -> bool {
        self.inner
            .properties
            .as_ref()
            .is_some_and(|properties| properties.look.is_some())
    }
}

/// An immutable reference to a table row.
pub struct RowRef<'a> {
    pub(crate) inner: &'a CT_Row,
}

impl<'a> RowRef<'a> {
    /// Get the number of cells.
    pub fn cell_count(&self) -> usize {
        self.inner.cells.len()
    }

    /// Whether the row contains direct content other than cells.
    pub fn has_unsupported_content(&self) -> bool {
        !self.inner.extra_xml.is_empty() || !self.inner.content_controls.is_empty()
    }

    /// Whether row properties retain facts outside the public reader model.
    pub fn has_unmodeled_properties(&self) -> bool {
        self.inner.properties.as_ref().is_some_and(|properties| {
            !properties.revision_markers.is_empty()
                || !properties.revision_xml.is_empty()
                || !properties.extra_xml.is_empty()
        })
    }

    /// Number of table grid columns omitted before the first cell.
    pub fn grid_before(&self) -> Option<u32> {
        self.inner
            .properties
            .as_ref()
            .and_then(|properties| properties.grid_before)
    }

    /// Number of table grid columns omitted after the last cell.
    pub fn grid_after(&self) -> Option<u32> {
        self.inner
            .properties
            .as_ref()
            .and_then(|properties| properties.grid_after)
    }

    /// Width of this row's omitted leading grid columns.
    pub fn width_before(&self) -> Option<TableWidth> {
        table_width_from_ct(self.inner.properties.as_ref()?.width_before.as_ref()?)
    }

    /// Width of this row's omitted trailing grid columns.
    pub fn width_after(&self) -> Option<TableWidth> {
        table_width_from_ct(self.inner.properties.as_ref()?.width_after.as_ref()?)
    }

    /// This row's gap between adjacent cell content boxes.
    pub fn cell_spacing(&self) -> Option<TableWidth> {
        table_width_from_ct(self.inner.properties.as_ref()?.cell_spacing.as_ref()?)
    }

    /// The authored hidden toggle, including an explicit false.
    pub fn hidden(&self) -> Option<bool> {
        self.inner.properties.as_ref()?.hidden
    }

    /// Get a cell reference by index.
    pub fn cell(&self, index: usize) -> Option<CellRef<'_>> {
        self.inner.cells.get(index).map(|c| CellRef { inner: c })
    }

    /// Check if this row is a header row.
    pub fn is_header(&self) -> bool {
        self.inner
            .properties
            .as_ref()
            .and_then(|pr| pr.header)
            .unwrap_or(false)
    }

    /// Get the direct minimum or exact height.
    pub fn height(&self) -> Option<RowHeight> {
        let properties = self.inner.properties.as_ref()?;
        let height = Length::twips(properties.height?.0);
        match properties.height_rule.as_deref().unwrap_or("atLeast") {
            "atLeast" => Some(RowHeight::AtLeast(height)),
            "exact" => Some(RowHeight::Exact(height)),
            _ => None,
        }
    }

    /// Get the direct repeating-header toggle, including explicit false.
    pub fn header_value(&self) -> Option<bool> {
        self.inner
            .properties
            .as_ref()
            .and_then(|properties| properties.header)
    }

    /// Get the direct split-policy toggle, including explicit false.
    pub fn cant_split_value(&self) -> Option<bool> {
        self.inner
            .properties
            .as_ref()
            .and_then(|properties| properties.cant_split)
    }

    /// Get direct row alignment.
    pub fn alignment(&self) -> Option<crate::paragraph::Alignment> {
        use crate::paragraph::Alignment;
        self.inner
            .properties
            .as_ref()
            .and_then(|properties| properties.jc)
            .map(|value| match value {
                ST_Jc::Center => Alignment::Center,
                ST_Jc::Right | ST_Jc::End => Alignment::Right,
                ST_Jc::Both | ST_Jc::Distribute => Alignment::Justify,
                _ => Alignment::Left,
            })
    }

    /// Get direct conditional table-style regions.
    pub fn conditional_formatting(&self) -> Option<TableConditionalFormatting> {
        self.inner
            .properties
            .as_ref()?
            .cnf_style
            .as_deref()
            .and_then(TableConditionalFormatting::from_value)
    }

    /// Whether the row carries explicit presentation formatting.
    pub fn has_formatting(&self) -> bool {
        self.inner.table_property_exception.is_some()
            || self.inner.properties.as_ref().is_some_and(|properties| {
                properties.height.is_some()
                    || properties.height_rule.is_some()
                    || properties.header.is_some()
                    || properties.jc.is_some()
                    || properties.grid_before.is_some()
                    || properties.grid_after.is_some()
                    || properties.width_before.is_some()
                    || properties.width_after.is_some()
                    || properties.cell_spacing.is_some()
                    || properties.hidden.is_some()
                    || properties.cant_split.is_some()
                    || properties.cnf_style.is_some()
            })
    }
}

/// An immutable reference to a table cell.
pub struct CellRef<'a> {
    pub(crate) inner: &'a CT_Tc,
}

/// One direct child of a table cell, in source order.
#[non_exhaustive]
pub enum CellItemRef<'a> {
    /// A cell paragraph.
    Paragraph(ParagraphRef<'a>),
    /// A nested table.
    Table(TableRef<'a>),
    /// A cell-level content control.
    ContentControl(ContentControlRef<'a>),
    /// A preserved cell child that rdocx does not model.
    UnsupportedXml(&'a [u8]),
}

impl<'a> CellRef<'a> {
    /// Whether cell properties retain facts outside the public reader model.
    pub fn has_unmodeled_properties(&self) -> bool {
        self.inner.properties.as_ref().is_some_and(|properties| {
            !properties.extra_xml.is_empty()
                || properties
                    .borders
                    .as_ref()
                    .is_some_and(|borders| !borders.extra_xml.is_empty())
        })
    }

    /// Get the combined text of all paragraphs.
    pub fn text(&self) -> String {
        self.inner.text()
    }

    /// Iterate over direct cell items in source order.
    ///
    /// Unlike [`Self::paragraphs`], this retains nested tables, content
    /// controls, and preserved unmodelled XML at their original boundaries.
    pub fn items(&self) -> impl Iterator<Item = CellItemRef<'_>> {
        let mut items = Vec::with_capacity(self.inner.content.len() + self.inner.extra_xml.len());
        for index in 0..=self.inner.content.len() {
            items.extend(
                self.inner
                    .extra_xml
                    .iter()
                    .filter(|(at, _)| *at == index)
                    .map(|(_, raw)| CellItemRef::UnsupportedXml(raw.as_slice())),
            );
            if let Some(content) = self.inner.content.get(index) {
                items.push(match content {
                    CellContent::Paragraph(paragraph) => {
                        CellItemRef::Paragraph(ParagraphRef { inner: paragraph })
                    }
                    CellContent::Table(table) => CellItemRef::Table(TableRef { inner: table }),
                    CellContent::ContentControl(control) => {
                        CellItemRef::ContentControl(ContentControlRef { inner: control })
                    }
                });
            }
        }
        items.into_iter()
    }

    /// Get paragraph references.
    pub fn paragraphs(&self) -> impl Iterator<Item = ParagraphRef<'_>> {
        self.inner
            .paragraphs()
            .into_iter()
            .map(|p| ParagraphRef { inner: p })
    }

    /// Get the number of paragraphs in the cell.
    pub fn paragraph_count(&self) -> usize {
        self.inner.paragraphs().len()
    }

    /// Get an immutable paragraph by index.
    pub fn paragraph(&self, index: usize) -> Option<ParagraphRef<'_>> {
        self.inner
            .paragraphs()
            .get(index)
            .map(|inner| ParagraphRef { inner })
    }

    /// Get the cell width when stored as twips.
    pub fn width(&self) -> Option<Length> {
        let width = self.inner.properties.as_ref()?.width.as_ref()?;
        (width.width_type == "dxa").then(|| Length::twips(width.w))
    }

    /// Get the grid span, if set.
    pub fn grid_span(&self) -> Option<u32> {
        self.inner.properties.as_ref().and_then(|pr| pr.grid_span)
    }

    /// Whether the cell uses the legacy horizontal-merge property.
    pub fn has_horizontal_merge(&self) -> bool {
        self.inner
            .properties
            .as_ref()
            .is_some_and(|properties| properties.h_merge.is_some())
    }

    /// Get the vertical merge state, if set.
    pub fn v_merge(&self) -> Option<&VMerge> {
        self.inner
            .properties
            .as_ref()
            .and_then(|pr| pr.v_merge.as_ref())
    }

    /// Get the shading fill color, if set.
    pub fn shading_fill(&self) -> Option<&str> {
        self.inner
            .properties
            .as_ref()
            .and_then(|pr| pr.shading.as_ref())
            .and_then(|shd| shd.fill.as_deref())
    }

    /// Get one direct cell border, including an explicit invisible edge.
    pub fn border(&self, position: CellBorderEdge) -> Option<TableBorderRef<'_>> {
        let borders = self.inner.properties.as_ref()?.borders.as_ref()?;
        let inner = match position {
            CellBorderEdge::Top => borders.top.as_ref(),
            CellBorderEdge::Bottom => borders.bottom.as_ref(),
            CellBorderEdge::Left => borders.left.as_ref(),
            CellBorderEdge::Right => borders.right.as_ref(),
            CellBorderEdge::InsideHorizontal => borders.inside_h.as_ref(),
            CellBorderEdge::InsideVertical => borders.inside_v.as_ref(),
        }?;
        Some(TableBorderRef { inner })
    }

    /// Get direct per-cell margins.
    pub fn margins(&self) -> Option<TableCellMargins> {
        let margins = self.inner.properties.as_ref()?.cell_margin.as_ref()?;
        Some(TableCellMargins {
            top: margins.top.map(|value| Length::twips(value.0)),
            right: margins.right.map(|value| Length::twips(value.0)),
            bottom: margins.bottom.map(|value| Length::twips(value.0)),
            left: margins.left.map(|value| Length::twips(value.0)),
        })
    }

    /// Get the direct cell text direction.
    pub fn text_direction(&self) -> Option<CellTextDirection> {
        self.inner
            .properties
            .as_ref()?
            .text_direction
            .as_deref()
            .and_then(CellTextDirection::from_str)
    }

    /// Get the direct no-wrap toggle, including explicit false.
    pub fn no_wrap_value(&self) -> Option<bool> {
        self.inner
            .properties
            .as_ref()
            .and_then(|properties| properties.no_wrap)
    }

    /// Get direct conditional table-style regions.
    pub fn conditional_formatting(&self) -> Option<TableConditionalFormatting> {
        self.inner
            .properties
            .as_ref()?
            .cnf_style
            .as_deref()
            .and_then(TableConditionalFormatting::from_value)
    }

    /// Whether explicit cell width is present.
    pub fn has_width(&self) -> bool {
        self.inner
            .properties
            .as_ref()
            .is_some_and(|properties| properties.width.is_some())
    }

    /// Whether explicit cell borders are present.
    pub fn has_borders(&self) -> bool {
        self.inner
            .properties
            .as_ref()
            .and_then(|properties| properties.borders.as_ref())
            .is_some_and(|borders| !borders.is_empty())
    }

    /// Whether explicit cell shading is present.
    pub fn has_shading(&self) -> bool {
        self.inner
            .properties
            .as_ref()
            .is_some_and(|properties| properties.shading.is_some())
    }

    /// Whether the cell disables wrapping.
    pub fn has_wrapping_formatting(&self) -> bool {
        self.inner
            .properties
            .as_ref()
            .is_some_and(|properties| properties.no_wrap.is_some())
    }

    /// Whether the cell carries explicit per-cell margins.
    pub fn has_cell_margins(&self) -> bool {
        self.inner
            .properties
            .as_ref()
            .is_some_and(|properties| properties.cell_margin.is_some())
    }

    /// Whether the cell carries an explicit text direction.
    pub fn has_text_direction(&self) -> bool {
        self.inner
            .properties
            .as_ref()
            .is_some_and(|properties| properties.text_direction.is_some())
    }

    /// Whether the cell carries table-style conditional formatting flags.
    pub fn has_conditional_formatting(&self) -> bool {
        self.inner
            .properties
            .as_ref()
            .is_some_and(|properties| properties.cnf_style.is_some())
    }

    /// Get the vertical alignment, if set.
    pub fn vertical_alignment(&self) -> Option<VerticalAlignment> {
        self.inner
            .properties
            .as_ref()
            .and_then(|pr| pr.v_align)
            .map(VerticalAlignment::from_st)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use rdocx_oxml::shared::ST_Border;
    use rdocx_oxml::table::{CT_Row, CT_TblGrid, CT_TblGridCol, CT_TblLook};
    use rdocx_oxml::units::Twips;

    #[test]
    fn reader_exposes_table_row_and_cell_completeness_facts() {
        let border = CT_BorderEdge::new(ST_Border::Single);
        let mut inner = CT_Tbl::new();
        inner.extra_xml.push((0, b"<w:custom/>".to_vec()));
        inner.grid = Some(CT_TblGrid {
            extra_xml: vec![b"<w:customGrid/>".to_vec()],
            ..Default::default()
        });
        inner.properties = Some(CT_TblPr {
            width: Some(CT_TblWidth::dxa(1_000)),
            borders: Some(CT_TblBorders {
                top: Some(border.clone()),
                extra_xml: vec![b"<w:diagonalDown/>".to_vec()],
                ..Default::default()
            }),
            shading: Some(CT_Shd {
                val: "clear".to_owned(),
                color: None,
                fill: Some("FFFFFF".to_owned()),
                ..Default::default()
            }),
            layout: Some("fixed".to_owned()),
            cell_margin: Some(CT_TblCellMar::default()),
            indent: Some(CT_TblWidth::dxa(100)),
            look: Some(CT_TblLook::default()),
            extra_xml: vec![(3, br#"<ext:span xmlns:ext="urn:producer"/>"#.to_vec())],
            ..Default::default()
        });

        let mut row = CT_Row::new();
        row.extra_xml.push((0, b"<w:customRow/>".to_vec()));
        row.properties = Some(CT_TrPr {
            height: Some(Twips(240)),
            grid_before: Some(1),
            grid_after: Some(2),
            extra_xml: vec![(1, br#"<w:divId w:val="1"/>"#.to_vec())],
            ..Default::default()
        });

        let mut cell = CT_Tc::new();
        cell.properties = Some(CT_TcPr {
            width: Some(CT_TblWidth::dxa(1_000)),
            h_merge: Some("restart".to_owned()),
            v_merge: Some(VMerge::Restart),
            borders: Some(CT_TblBorders {
                top: Some(border),
                extra_xml: vec![b"<w:diagonalDown/>".to_vec()],
                ..Default::default()
            }),
            shading: Some(CT_Shd {
                val: "clear".to_owned(),
                color: None,
                fill: Some("FFFFFF".to_owned()),
                ..Default::default()
            }),
            no_wrap: Some(true),
            cell_margin: Some(CT_TblCellMar::default()),
            text_direction: Some("btLr".to_owned()),
            cnf_style: Some("100000000000".to_owned()),
            extra_xml: vec![(0, b"<w:fitText/>".to_vec())],
            ..Default::default()
        });
        row.cells.push(cell);
        inner.rows.push(row);

        let table = TableRef { inner: &inner };
        assert!(table.has_unsupported_content());
        assert!(table.has_unmodeled_properties());
        assert!(table.has_width());
        assert!(table.has_borders());
        assert!(table.has_shading());
        assert!(table.has_layout());
        assert!(table.has_cell_margins());
        assert!(table.has_indent());
        assert!(table.has_style_look());

        let row = table.row(0).expect("table row");
        assert!(row.has_unsupported_content());
        assert!(row.has_unmodeled_properties());
        assert!(row.has_formatting());
        assert_eq!(row.grid_before(), Some(1));
        assert_eq!(row.grid_after(), Some(2));

        let cell = row.cell(0).expect("table cell");
        assert!(cell.has_unmodeled_properties());
        assert!(cell.has_horizontal_merge());
        assert!(matches!(cell.v_merge(), Some(VMerge::Restart)));
        assert!(cell.has_width());
        assert!(cell.has_borders());
        assert!(cell.has_shading());
        assert!(cell.has_wrapping_formatting());
        assert!(cell.has_cell_margins());
        assert!(cell.has_text_direction());
        assert!(cell.has_conditional_formatting());
    }

    #[test]
    fn row_formatting_reports_header_and_grid_offset_facts() {
        for properties in [
            CT_TrPr {
                header: Some(true),
                ..Default::default()
            },
            CT_TrPr {
                grid_before: Some(1),
                ..Default::default()
            },
            CT_TrPr {
                grid_after: Some(1),
                ..Default::default()
            },
        ] {
            let mut row = CT_Row::new();
            row.properties = Some(properties);
            assert!(RowRef { inner: &row }.has_formatting());
        }
    }

    #[test]
    fn table_ref_reports_preserved_grid_change() {
        let historical = br#"<w:tblGridChange w:id="4"><w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid></w:tblGridChange>"#;
        let mut document = crate::Document::new();
        {
            let table = document.add_table(1, 1);
            table.inner.grid.as_mut().unwrap().grid_change_xml = Some(historical.to_vec());
        }

        let package = document.to_bytes().expect("table package saves");
        let reopened = crate::Document::from_bytes(&package).expect("table package reopens");

        assert!(reopened.table(0).unwrap().has_grid_change());
        assert_eq!(
            reopened
                .table(0)
                .unwrap()
                .inner
                .grid
                .as_ref()
                .unwrap()
                .grid_change_xml
                .as_deref(),
            Some(historical.as_slice())
        );
    }

    #[test]
    fn table_column_width_updates_grid_table_and_spanning_cells() {
        let mut inner = CT_Tbl::new();
        inner.grid = Some(CT_TblGrid {
            columns: vec![
                CT_TblGridCol {
                    width: Twips(1_000),
                },
                CT_TblGridCol {
                    width: Twips(2_000),
                },
                CT_TblGridCol {
                    width: Twips(3_000),
                },
            ],
            ..Default::default()
        });
        let mut row = CT_Row::new();
        let mut spanning_cell = CT_Tc::new();
        spanning_cell.properties = Some(CT_TcPr {
            grid_span: Some(2),
            ..CT_TcPr::default()
        });
        row.cells.push(spanning_cell);
        row.cells.push(CT_Tc::new());
        inner.rows.push(row);

        let mut table = Table { inner: &mut inner };
        assert!(table.set_column_width(1, Length::twips(4_000)));

        let properties = table.inner.properties.as_ref().unwrap();
        assert_eq!(properties.width, Some(CT_TblWidth::dxa(8_000)));
        let first_cell = &table.inner.rows[0].cells[0];
        assert_eq!(
            first_cell.properties.as_ref().unwrap().width,
            Some(CT_TblWidth::dxa(5_000))
        );
    }

    #[test]
    fn table_column_width_rejects_negative_geometry_without_mutation() {
        let mut inner = CT_Tbl::new();
        inner.grid = Some(CT_TblGrid {
            columns: vec![
                CT_TblGridCol {
                    width: Twips(1_000),
                },
                CT_TblGridCol {
                    width: Twips(2_000),
                },
            ],
            ..Default::default()
        });
        inner.properties = Some(CT_TblPr {
            width: Some(CT_TblWidth::dxa(3_000)),
            ..CT_TblPr::default()
        });
        let before = inner.clone();

        let mut table = Table { inner: &mut inner };
        assert!(!table.set_column_width(0, Length::twips(-1)));
        assert_eq!(*table.inner, before);
    }

    #[test]
    fn checked_grid_omission_rejects_malformed_existing_topology_without_mutation() {
        let mut inner = CT_Tbl::new();
        inner.grid = Some(CT_TblGrid {
            columns: vec![
                CT_TblGridCol {
                    width: Twips(1_000),
                },
                CT_TblGridCol {
                    width: Twips(1_000),
                },
            ],
            ..Default::default()
        });
        let mut row = CT_Row::new();
        row.properties = Some(CT_TrPr {
            grid_after: Some(3),
            ..Default::default()
        });
        row.cells.push(CT_Tc::new());
        inner.rows.push(row);
        let before = inner.clone();

        let mut table = Table { inner: &mut inner };
        assert!(table.set_row_grid_omissions(0, None, None).is_err());
        assert_eq!(*table.inner, before);
    }
}
