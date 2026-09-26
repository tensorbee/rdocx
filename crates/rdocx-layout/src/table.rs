//! Table layout: column widths, cell content, merge handling.

use rdocx_oxml::borders::CT_BorderEdge;
use rdocx_oxml::content_control::{CT_Sdt, SdtContent};
use rdocx_oxml::document::CT_DocGrid;
use rdocx_oxml::drawing::{AnchorAlignH, AnchorAlignV, ST_RelativeFromH, ST_RelativeFromV};
use rdocx_oxml::shared::{ST_Border, ST_Jc};
use rdocx_oxml::styles::{CT_Styles, TableStyleRegion};
use rdocx_oxml::table::{
    CT_Row, CT_Tbl, CT_TblBorders, CT_TblGrid, CT_TblPPr, CT_TblPr, CT_TblWidth, CT_Tc, CT_TrPr,
    ST_TblAnchor, ST_TblOverlap, ST_VerticalJc, ST_YAlign, VMerge,
};

use crate::WordStory;
use crate::block::{
    CellBlockSemantics, CellSemantics, ParagraphBlock, ParagraphSemantics, RowSemantics,
    TableSemantics,
};
use crate::engine::SourceRegistry;
use crate::input::{LayoutInput, MediaRegistry};
use crate::style_resolver::NumberingState;
use oxml_layout::{Color, Diagnostic, FontManager, Result, StructureId};

const CONTROL_PATH_COMPONENT: usize = usize::MAX;

pub(crate) fn layout_table_rows<'a>(
    table: &'a CT_Tbl,
    path: &[usize],
) -> Vec<(&'a CT_Row, Vec<usize>)> {
    let mut rows = Vec::new();
    for boundary in 0..=table.rows.len() {
        for (control_index, (_, _, control)) in table
            .content_controls
            .iter()
            .enumerate()
            .filter(|(_, (position, _, _))| *position == boundary)
        {
            let mut control_path = path.to_vec();
            control_path.extend([CONTROL_PATH_COMPONENT, boundary, control_index]);
            collect_control_rows(control, &control_path, &mut rows);
        }
        if let Some(row) = table.rows.get(boundary) {
            let mut row_path = path.to_vec();
            row_path.push(boundary);
            rows.push((row, row_path));
        }
    }
    rows
}

fn collect_control_rows<'a>(
    control: &'a CT_Sdt,
    path: &[usize],
    rows: &mut Vec<(&'a CT_Row, Vec<usize>)>,
) {
    for (content_index, content) in control.content.iter().enumerate() {
        let mut content_path = path.to_vec();
        content_path.push(content_index);
        match content {
            SdtContent::Row(row) => rows.push((row, content_path)),
            SdtContent::ContentControl(control) => {
                collect_control_rows(control, &content_path, rows)
            }
            _ => {}
        }
    }
}

pub(crate) fn layout_row_cells<'a>(
    row: &'a CT_Row,
    path: &[usize],
) -> Vec<(&'a CT_Tc, Vec<usize>)> {
    let mut cells = Vec::new();
    for boundary in 0..=row.cells.len() {
        for (control_index, (_, _, control)) in row
            .content_controls
            .iter()
            .enumerate()
            .filter(|(_, (position, _, _))| *position == boundary)
        {
            let mut control_path = path.to_vec();
            control_path.extend([CONTROL_PATH_COMPONENT, boundary, control_index]);
            collect_control_cells(control, &control_path, &mut cells);
        }
        if let Some(cell) = row.cells.get(boundary) {
            let mut cell_path = path.to_vec();
            cell_path.push(boundary);
            cells.push((cell, cell_path));
        }
    }
    cells
}

fn collect_control_cells<'a>(
    control: &'a CT_Sdt,
    path: &[usize],
    cells: &mut Vec<(&'a CT_Tc, Vec<usize>)>,
) {
    for (content_index, content) in control.content.iter().enumerate() {
        let mut content_path = path.to_vec();
        content_path.push(content_index);
        match content {
            SdtContent::Cell(cell) => cells.push((cell, content_path)),
            SdtContent::ContentControl(control) => {
                collect_control_cells(control, &content_path, cells)
            }
            _ => {}
        }
    }
}

/// Where a floating table sits, lowered from `w:tblpPr` onto the anchor frames
/// the paginator already resolves for a floating drawing.
///
/// This is deliberately the positioning half of `AnchoredDrawing`, field for
/// field, so `resolve_anchor_h` and `resolve_anchor_v` take it without a
/// signature change. The content half does not apply, because the content is
/// the table's own rows.
#[derive(Debug, Clone)]
pub struct FloatingTable {
    /// Frame the horizontal offset is measured from.
    pub rel_h: ST_RelativeFromH,
    /// Horizontal offset in points.
    pub off_h: f64,
    /// Horizontal alignment, used instead of the offset when present.
    pub align_h: Option<AnchorAlignH>,
    /// Frame the vertical offset is measured from.
    pub rel_v: ST_RelativeFromV,
    /// Vertical offset in points.
    pub off_v: f64,
    /// Vertical alignment, used instead of the offset when present.
    pub align_v: Option<AnchorAlignV>,
    /// Clearance kept between the table and the text flowing around it, in
    /// points.
    pub dist_top: f64,
    pub dist_bottom: f64,
    pub dist_left: f64,
    pub dist_right: f64,
    /// Whether `w:tblOverlap` lets this float overlap another one.
    pub overlap_allowed: bool,
}

/// Map `w:horzAnchor` onto the horizontal frame a drawing anchor names.
///
/// Three of the eight variants are reachable, because `ST_TblAnchor` spells
/// only three. An absent anchor reads as `margin`, which is Word's default.
fn floating_frame_h(anchor: Option<ST_TblAnchor>) -> ST_RelativeFromH {
    match anchor {
        Some(ST_TblAnchor::Page) => ST_RelativeFromH::Page,
        Some(ST_TblAnchor::Text) => ST_RelativeFromH::Column,
        Some(ST_TblAnchor::Margin) | None => ST_RelativeFromH::Margin,
    }
}

/// Map `w:vertAnchor` onto the vertical frame a drawing anchor names.
fn floating_frame_v(anchor: Option<ST_TblAnchor>) -> ST_RelativeFromV {
    match anchor {
        Some(ST_TblAnchor::Page) => ST_RelativeFromV::Page,
        Some(ST_TblAnchor::Text) => ST_RelativeFromV::Paragraph,
        Some(ST_TblAnchor::Margin) | None => ST_RelativeFromV::Margin,
    }
}

/// Lower the resolved table properties onto a float, when they declare one.
///
/// `tblpYSpec="inline"` is how `w:tblpPr` spells "not floating", and an absent
/// `w:tblpPr` means the same, so both leave the table in the flow. This is the
/// one place that decides whether a table floats, so the engine asks it rather
/// than repeating the rule.
pub(crate) fn floating_table(properties: &CT_TblPr) -> Option<FloatingTable> {
    let position: &CT_TblPPr = properties.float_position.as_deref()?;
    if position.tbl_p_y_spec == Some(ST_YAlign::Inline) {
        return None;
    }
    let align_v = match position.tbl_p_y_spec {
        Some(ST_YAlign::Top) => Some(AnchorAlignV::Top),
        Some(ST_YAlign::Center) => Some(AnchorAlignV::Center),
        Some(ST_YAlign::Bottom) => Some(AnchorAlignV::Bottom),
        Some(ST_YAlign::Inside) => Some(AnchorAlignV::Inside),
        Some(ST_YAlign::Outside) => Some(AnchorAlignV::Outside),
        Some(ST_YAlign::Inline) | None => None,
    };
    // An absent measurement is zero, which is the schema default for each of
    // these attributes and what the public reader already reports.
    let points =
        |value: Option<rdocx_oxml::Twips>| value.map_or(0.0, |twips| twips.0 as f64 / 20.0);
    Some(FloatingTable {
        rel_h: floating_frame_h(position.horz_anchor),
        off_h: points(position.tbl_p_x),
        align_h: position.tbl_p_x_spec,
        rel_v: floating_frame_v(position.vert_anchor),
        off_v: points(position.tbl_p_y),
        align_v,
        dist_top: points(position.top_from_text),
        dist_bottom: points(position.bottom_from_text),
        dist_left: points(position.left_from_text),
        dist_right: points(position.right_from_text),
        overlap_allowed: properties.overlap != Some(ST_TblOverlap::Never),
    })
}

/// A laid-out table.
#[derive(Debug, Clone)]
pub struct TableBlock {
    /// Logical table node allocated before pagination.
    pub structure_id: Option<StructureId>,
    /// Column widths in points.
    pub col_widths: Vec<f64>,
    /// Laid-out rows.
    pub rows: Vec<TableRow>,
    /// Indices of rows that are header rows (repeat on page break).
    pub header_row_indices: Vec<usize>,
    /// Total table width in points.
    pub table_width: f64,
    /// Table indent from left margin in points.
    pub table_indent: f64,
    /// Table-level borders (used as fallback for cell borders).
    pub borders: Option<CT_TblBorders>,
    /// Whether `w:bidiVisual` reverses visual column placement.
    ///
    /// Only the painting order is reversed. `col_widths` and every cell's
    /// `col_index` stay logical, which is what keeps cell ownership, the
    /// structure tree and the body fragments in reading order.
    pub bidi_visual: bool,
    /// Where `w:tblpPr` floats this table, or `None` for a table in the flow.
    ///
    /// Boxed, like the `CT_TblPPr` it is lowered from, so a table stays cheap
    /// on the stack. Test threads build whole documents by value against a
    /// 2 MiB ceiling, and `TableBlock` nests inside itself through `CellBlock`.
    pub floating: Option<Box<FloatingTable>>,
}

impl TableBlock {
    /// Total content height of all rows.
    pub fn content_height(&self) -> f64 {
        self.rows.iter().map(|r| r.height).sum()
    }

    /// Total height (same as content for tables, no before/after spacing).
    pub fn total_height(&self) -> f64 {
        self.content_height()
    }
}

/// A laid-out table row.
#[derive(Debug, Clone)]
pub struct TableRow {
    /// Logical row node allocated before pagination.
    pub structure_id: Option<StructureId>,
    /// Cells in this row.
    pub cells: Vec<TableCell>,
    /// Row height in points.
    pub height: f64,
    /// Whether this row is a header row.
    pub is_header: bool,
    /// Distance in points from the table origin to this row's first painted
    /// cell, resolved from the row's omitted grid columns and their width.
    ///
    /// A bidirectional row measures the omission on its own leading side,
    /// which is the trailing side of the logical grid.
    pub offset_left: f64,
}

/// One source-ordered block inside a table cell.
#[derive(Debug, Clone)]
pub enum CellBlock {
    /// A laid-out paragraph.
    Paragraph(ParagraphBlock),
    /// A recursively laid-out nested table.
    Table(TableBlock),
}

impl CellBlock {
    /// Total block height in points.
    pub fn total_height(&self) -> f64 {
        match self {
            Self::Paragraph(paragraph) => paragraph.total_height(),
            Self::Table(table) => table.total_height(),
        }
    }
}

/// A laid-out table cell.
#[derive(Debug, Clone)]
pub struct TableCell {
    /// Logical cell node allocated before pagination.
    pub structure_id: Option<StructureId>,
    /// Source-ordered paragraph and nested-table blocks.
    pub blocks: Vec<CellBlock>,
    /// Cell width in points (may span multiple grid columns).
    pub width: f64,
    /// Cell height in points (set to row height).
    pub height: f64,
    /// Number of grid columns this cell spans.
    pub grid_span: u32,
    /// Whether this cell is part of a vertical merge continuation (render no content).
    pub is_vmerge_continue: bool,
    /// Whether this cell begins a vertical merge.
    pub starts_vmerge: bool,
    /// Height of the complete vertical-merge span.
    pub merged_height: f64,
    /// Whether the same grid span continues in the next row.
    pub merge_with_below: bool,
    /// Whether overflowing cell content must be clipped to its painted span.
    pub clip_content: bool,
    /// Column index in the grid.
    pub col_index: usize,
    /// Cell-level borders.
    pub borders: Option<CT_TblBorders>,
    /// Cell background shading color.
    pub shading: Option<Color>,
    /// Cell margin left in points.
    pub margin_left: f64,
    /// Cell margin right in points.
    pub margin_right: f64,
    /// Cell margin top in points.
    pub margin_top: f64,
    /// Cell margin bottom in points.
    pub margin_bottom: f64,
    /// Points of horizontal border band above the content, which the row
    /// height includes.
    pub border_band_top: f64,
    /// Points of horizontal border band below the content. Only a cell that
    /// reaches the table's last row has one, the table's bottom border.
    pub border_band_bottom: f64,
    /// Whether this cell is in the first row.
    pub is_first_row: bool,
    /// Whether this cell is in the last row.
    pub is_last_row: bool,
    /// Vertical alignment of content within the cell.
    pub v_align: Option<ST_VerticalJc>,
    /// Degrees the cell's content box rotates for `w:tcPr/w:textDirection`.
    ///
    /// `None` is the ordinary horizontal cell, which takes the placement
    /// arithmetic it always had with no group wrapper.
    pub rotation: Option<f64>,
}

/// The diagnostic an upright stacked East Asian direction records.
///
/// Upright stacking is out of scope and stays visible as rotated text, which
/// is the fallback `docs/hld/08-rendering-spec.md` already documents for the
/// DrawingML shape path, so the product says one thing about it.
pub(crate) const UPRIGHT_STACK_DIAGNOSTIC: &str =
    "east Asian vertical text rendered as rotated vertical text";

/// The rotation in degrees a `w:textDirection` value projects onto.
///
/// `lrTb` and any unmodelled value return `None`, which is today's horizontal
/// path. `tbRl` and `tbRlV` rotate 90 degrees, and `btLr`, `lrTbV` and
/// `tbLrV` rotate -90.
pub(crate) fn text_direction_rotation(value: &str) -> Option<f64> {
    match value {
        "tbRl" | "tbRlV" => Some(90.0),
        "btLr" | "lrTbV" | "tbLrV" => Some(-90.0),
        _ => None,
    }
}

/// The rotation one cell's `w:tcPr/w:textDirection` projects onto.
pub(crate) fn cell_rotation(cell: &CT_Tc) -> Option<f64> {
    cell.properties
        .as_ref()
        .and_then(|properties| properties.text_direction.as_deref())
        .and_then(text_direction_rotation)
}

/// Whether a `w:textDirection` value asks for upright stacked East Asian text.
pub(crate) fn text_direction_stacks_upright(value: &str) -> bool {
    matches!(value, "lrTbV" | "tbRlV" | "tbLrV")
}

/// The same-centre transposed content box a rotated cell is laid out in.
///
/// Width and height swap about the box centre, so rotating the laid-out
/// result about that same centre lands it back inside the cell.
pub(crate) fn transposed_box(x: f64, y: f64, width: f64, height: f64) -> (f64, f64, f64, f64) {
    let center_x = x + width / 2.0;
    let center_y = y + height / 2.0;
    (
        center_x - height / 2.0,
        center_y - width / 2.0,
        height,
        width,
    )
}

/// Lay out a table into a TableBlock.
pub fn layout_table(
    tbl: &CT_Tbl,
    available_width: f64,
    styles: &CT_Styles,
    input: &LayoutInput,
    media: &MediaRegistry,
    fm: &mut FontManager,
    num_state: &mut NumberingState,
    diagnostics: &mut Vec<Diagnostic>,
    doc_grid: Option<&CT_DocGrid>,
) -> Result<TableBlock> {
    layout_table_inner(
        tbl,
        available_width,
        styles,
        input,
        media,
        fm,
        num_state,
        diagnostics,
        None,
        &WordStory::Document,
        &[],
        doc_grid,
    )
    .map(|(block, _)| block)
}

pub(crate) fn layout_table_with_provenance(
    tbl: &CT_Tbl,
    available_width: f64,
    styles: &CT_Styles,
    input: &LayoutInput,
    media: &MediaRegistry,
    fm: &mut FontManager,
    num_state: &mut NumberingState,
    diagnostics: &mut Vec<Diagnostic>,
    sources: Option<&SourceRegistry>,
    story: &WordStory,
    path: &[usize],
    doc_grid: Option<&CT_DocGrid>,
) -> Result<(TableBlock, TableSemantics)> {
    layout_table_inner(
        tbl,
        available_width,
        styles,
        input,
        media,
        fm,
        num_state,
        diagnostics,
        sources,
        story,
        path,
        doc_grid,
    )
}

#[allow(clippy::too_many_arguments)]
fn layout_table_inner(
    tbl: &CT_Tbl,
    available_width: f64,
    styles: &CT_Styles,
    input: &LayoutInput,
    media: &MediaRegistry,
    fm: &mut FontManager,
    num_state: &mut NumberingState,
    diagnostics: &mut Vec<Diagnostic>,
    sources: Option<&SourceRegistry>,
    story: &WordStory,
    path: &[usize],
    doc_grid: Option<&CT_DocGrid>,
) -> Result<(TableBlock, TableSemantics)> {
    let direct_width = tbl
        .properties
        .as_ref()
        .is_some_and(|properties| properties.width.is_some());
    let direct_alignment = tbl
        .properties
        .as_ref()
        .is_some_and(|properties| properties.jc.is_some());
    let mut resolved_table = tbl.clone();
    let mut resolved_properties = resolve_base_table_properties(tbl, styles);
    // The authored width type, captured before the direct width is dropped
    // below. Autofit engages against what the author declared, and a direct
    // `w:tblW` is authored even though the declared grid, not the width,
    // drives the ordinary path.
    let authored_width_type = resolved_properties
        .width
        .as_ref()
        .map(|width| width.width_type.clone());
    if direct_width {
        resolved_properties.width = None;
    }
    if direct_alignment {
        resolved_properties.jc = None;
    }
    resolved_table.properties = Some(resolved_properties);
    let tbl = &resolved_table;
    let source_rows = layout_table_rows(tbl, path);
    let bidi_visual = tbl
        .properties
        .as_ref()
        .and_then(|properties| properties.bidi_visual)
        .unwrap_or(false);
    let floating = tbl
        .properties
        .as_ref()
        .and_then(floating_table)
        .map(Box::new);
    // 1. Compute column widths. Content-driven autofit engages only for an
    //    auto or absent width with an autofit or absent layout mode.
    let col_widths = match autofit_column_widths(
        tbl,
        authored_width_type.as_deref(),
        available_width,
        styles,
        input,
        media,
        fm,
        num_state,
        path,
        doc_grid,
    )? {
        Some(widths) => widths,
        None => compute_column_widths(tbl.grid.as_ref(), available_width, tbl, path),
    };
    let table_width: f64 = col_widths.iter().sum();

    // Table indent
    let authored_indent = tbl
        .properties
        .as_ref()
        .and_then(|p| p.indent.as_ref())
        .map(|ind| {
            if ind.width_type == "dxa" {
                ind.w as f64 / 20.0 // twips to pt
            } else {
                0.0
            }
        })
        .unwrap_or(0.0);
    let table_indent = match tbl.properties.as_ref().and_then(|properties| properties.jc) {
        Some(ST_Jc::Center) => ((available_width - table_width) / 2.0).max(0.0),
        Some(ST_Jc::Right | ST_Jc::End) => (available_width - table_width).max(0.0),
        _ => authored_indent,
    };

    // Direct table borders win. Table-style borders are the fallback.
    let table_borders = tbl
        .properties
        .as_ref()
        .and_then(|properties| properties.borders.clone())
        .or_else(|| {
            let mut style_id = tbl.properties.as_ref()?.style_id.as_deref()?;
            let mut visited = std::collections::HashSet::new();
            while visited.insert(style_id) {
                let style = styles.get_by_id(style_id)?;
                if let Some(properties) = &style.table_properties
                    && let Some(borders) = &properties.borders
                {
                    return Some(borders.clone());
                }
                style_id = style.based_on.as_deref()?;
            }
            None
        });

    // Default cell margins
    let default_cell_margin = tbl.properties.as_ref().and_then(|p| p.cell_margin.as_ref());
    let cell_margin_left = default_cell_margin
        .and_then(|m| m.left)
        .map(|t| t.to_pt())
        .unwrap_or(5.4); // Word default ~108 twips
    let cell_margin_right = default_cell_margin
        .and_then(|m| m.right)
        .map(|t| t.to_pt())
        .unwrap_or(5.4);
    let cell_margin_top = default_cell_margin
        .and_then(|m| m.top)
        .map(|t| t.to_pt())
        .unwrap_or(0.0);
    let cell_margin_bottom = default_cell_margin
        .and_then(|m| m.bottom)
        .map(|t| t.to_pt())
        .unwrap_or(0.0);
    // A percentage cell gap is a percentage of the table, not of the caller's
    // width, and a gap with no length resolves to none.
    let table_cell_spacing = tbl
        .properties
        .as_ref()
        .and_then(|properties| properties.cell_spacing.as_ref())
        .map(|spacing| {
            table_width_to_pt(Some(spacing), table_width)
                .unwrap_or(0.0)
                .max(0.0)
        });

    let num_rows = source_rows.len();
    let mut header_row_indices = Vec::new();
    let mut rows = Vec::new();
    let mut row_semantics = Vec::new();
    let mut exact_rows = Vec::new();

    for (row_idx, (row, row_path)) in source_rows.iter().enumerate() {
        let row_properties =
            resolve_row_properties(tbl, styles, row, row_idx, num_rows, col_widths.len().max(1));
        let is_header = row_properties.header.unwrap_or(false);
        if is_header {
            header_row_indices.push(row_idx);
        }

        // The row's own gap wins over the table's, and each cell carries half
        // of it so adjacent content boxes are one whole gap apart.
        let row_cell_spacing = row_properties
            .cell_spacing
            .as_ref()
            .map(|spacing| {
                table_width_to_pt(Some(spacing), table_width)
                    .unwrap_or(0.0)
                    .max(0.0)
            })
            .or(table_cell_spacing)
            .unwrap_or(0.0);
        let half_spacing = row_cell_spacing / 2.0;

        // Omitted edge columns move the row's own origin. The table origin,
        // the table width and every other row stay where they are.
        let grid_before = (row_properties.grid_before.unwrap_or(0) as usize).min(col_widths.len());
        let grid_after = (row_properties.grid_after.unwrap_or(0) as usize)
            .min(col_widths.len().saturating_sub(grid_before));
        let omitted_before = table_width_to_pt(row_properties.width_before.as_ref(), table_width)
            .unwrap_or_else(|| col_widths.iter().take(grid_before).sum())
            .max(0.0);
        let omitted_after = table_width_to_pt(row_properties.width_after.as_ref(), table_width)
            .unwrap_or_else(|| col_widths.iter().rev().take(grid_after).sum())
            .max(0.0);
        let offset_left = if bidi_visual {
            omitted_after
        } else {
            omitted_before
        };

        let mut cells = Vec::new();
        let mut cell_semantics = Vec::new();
        let mut col_index = grid_before;

        let source_cells = layout_row_cells(row, row_path);
        for (cell, cell_path) in &source_cells {
            let grid_span = cell
                .properties
                .as_ref()
                .and_then(|p| p.grid_span)
                .unwrap_or(1);

            let is_vmerge_continue = cell
                .properties
                .as_ref()
                .and_then(|p| p.v_merge)
                .map(|vm| vm == VMerge::Continue)
                .unwrap_or(false);
            let starts_vmerge =
                cell.properties.as_ref().and_then(|p| p.v_merge) == Some(VMerge::Restart);

            let style_cell = resolve_table_style_cell(
                tbl,
                styles,
                row_idx,
                col_index,
                num_rows,
                col_widths.len(),
                &cell_conditional_selectors(row, cell),
            );

            // Direct cell borders overlay table-style region borders.
            let mut cell_borders = style_cell.borders;
            if let Some(direct) = cell.properties.as_ref().and_then(|p| p.borders.as_ref()) {
                overlay_borders(&mut cell_borders, direct);
            }
            let cell_shading = cell
                .properties
                .as_ref()
                .and_then(|p| p.shading.as_ref())
                .or(style_cell.shading.as_ref())
                .and_then(|shd| shd.fill.as_ref())
                .filter(|f| f.as_str() != "auto")
                .map(|f| Color::from_hex(f));

            // A cell's own `w:tcMar` replaces the table's margins edge by edge.
            let own_margin = cell
                .properties
                .as_ref()
                .and_then(|properties| properties.cell_margin.as_ref());
            let margin = |own: Option<rdocx_oxml::Twips>, table: f64| {
                own.map_or(table, |twips| twips.to_pt()) + half_spacing
            };
            let cell_margin_left = margin(own_margin.and_then(|m| m.left), cell_margin_left);
            let cell_margin_right = margin(own_margin.and_then(|m| m.right), cell_margin_right);
            let cell_margin_top = margin(own_margin.and_then(|m| m.top), cell_margin_top);
            let cell_margin_bottom = margin(own_margin.and_then(|m| m.bottom), cell_margin_bottom);

            // Calculate cell width from spanned columns
            let cell_width: f64 = (col_index..col_index + grid_span as usize)
                .filter_map(|i| col_widths.get(i))
                .sum();

            let content_width = (cell_width - cell_margin_left - cell_margin_right).max(0.0);

            // A rotated cell lays its content out in a same-centre transposed
            // box, so the measure runs along the cell's height rather than its
            // width. The painted box is the cell's height less its left and
            // right margins, because those margins sit across the transposed
            // box, so the measure subtracts the same pair. A row that declares
            // a height gives the measure exactly. An auto-height row grows to
            // the text, so the cell lays out unwrapped and the row becomes the
            // length it produced, which is then the measure the paginator
            // paints into.
            let cell_direction = cell
                .properties
                .as_ref()
                .and_then(|properties| properties.text_direction.as_deref());
            let rotation = cell_rotation(cell);
            if cell_direction.is_some_and(text_direction_stacks_upright) {
                crate::engine::push_unique_diagnostic(
                    diagnostics,
                    UPRIGHT_STACK_DIAGNOSTIC.to_owned(),
                );
            }
            let declared_height = row_properties.height.map(|h| h.to_pt()).unwrap_or(0.0);
            let layout_width = match rotation {
                Some(_) if declared_height > 0.0 => {
                    (declared_height - cell_margin_left - cell_margin_right).max(1.0)
                }
                Some(_) => VERTICAL_AUTO_MEASURE,
                None => content_width,
            };

            // Layout cell content (paragraphs and nested tables)
            let (blocks, block_semantics) = if is_vmerge_continue {
                (Vec::new(), Vec::new())
            } else {
                layout_cell_content(
                    &cell.content,
                    layout_width,
                    styles,
                    input,
                    media,
                    fm,
                    num_state,
                    diagnostics,
                    sources,
                    story,
                    cell_path,
                    style_cell.paragraph_properties.as_ref(),
                    style_cell.run_properties.as_ref(),
                    doc_grid,
                )?
            };

            // A rotated cell contributes the transposed box's measure to the
            // row, because its line direction runs down the cell. The left and
            // right margins are added back, so the row height the paginator
            // then strips them from is exactly the measure this laid out at.
            let content_height: f64 = if rotation.is_some() && !is_vmerge_continue {
                measured_content_width(&blocks) + cell_margin_left + cell_margin_right
            } else {
                blocks.iter().map(CellBlock::total_height).sum::<f64>()
                    + cell_margin_top
                    + cell_margin_bottom
            };

            let v_align = cell.properties.as_ref().and_then(|p| p.v_align);

            cells.push(TableCell {
                structure_id: None,
                blocks,
                width: cell_width,
                height: content_height,
                grid_span,
                is_vmerge_continue,
                starts_vmerge,
                merged_height: content_height,
                merge_with_below: false,
                clip_content: false,
                col_index,
                borders: cell_borders,
                shading: cell_shading,
                margin_left: cell_margin_left,
                margin_right: cell_margin_right,
                margin_top: cell_margin_top,
                margin_bottom: cell_margin_bottom,
                border_band_top: 0.0,
                border_band_bottom: 0.0,
                is_first_row: row_idx == 0,
                is_last_row: row_idx == num_rows - 1,
                v_align,
                rotation,
            });
            cell_semantics.push(CellSemantics {
                blocks: block_semantics,
            });

            col_index += grid_span as usize;
        }

        let specified_height = row_properties.height.map(|h| h.to_pt()).unwrap_or(0.0);
        let exact =
            row_properties.height_rule.as_deref() == Some("exact") && specified_height > 0.0;
        exact_rows.push(exact);
        for cell in &mut cells {
            cell.clip_content = exact && !cell.is_vmerge_continue;
        }
        // Word keeps a cell's top and bottom margins outside a minimum height,
        // which bounds the content between them. An exact height holds the
        // band and the top margin, and Word adds the bottom margin below it.
        let row_cells = || cells.iter().filter(|cell| !cell.starts_vmerge);
        let row_height = if exact {
            specified_height
                + row_cells()
                    .map(|cell| cell.margin_bottom)
                    .fold(0.0f64, f64::max)
        } else {
            row_cells()
                .map(|cell| {
                    cell.height
                        .max(specified_height + cell.margin_top + cell.margin_bottom)
                })
                .fold(specified_height, f64::max)
        };

        rows.push(TableRow {
            structure_id: None,
            cells,
            height: row_height,
            is_header,
            offset_left,
        });
        row_semantics.push(RowSemantics {
            cells: cell_semantics,
        });
    }

    // Resolve vertical merges over exact logical grid spans. Only the last
    // non-exact row grows when the restart content exceeds the full span.
    let mut spans = Vec::new();
    for row_index in 0..rows.len() {
        for cell_index in 0..rows[row_index].cells.len() {
            let cell = &rows[row_index].cells[cell_index];
            if !cell.starts_vmerge {
                continue;
            }
            let start_col = cell.col_index;
            let grid_span = cell.grid_span;
            let mut last_row = row_index;
            while let Some(next_row) = rows.get(last_row + 1) {
                let continues = next_row.cells.iter().any(|next| {
                    next.is_vmerge_continue
                        && next.col_index == start_col
                        && next.grid_span == grid_span
                });
                if !continues {
                    break;
                }
                last_row += 1;
            }
            let required = cell.height;
            let available = rows[row_index..=last_row]
                .iter()
                .map(|row| row.height)
                .sum::<f64>();
            if required > available
                && let Some(grow_row) = (row_index..=last_row)
                    .rev()
                    .find(|candidate| !exact_rows[*candidate])
            {
                rows[grow_row].height += required - available;
            }
            spans.push((row_index, cell_index, last_row, required));
        }
    }

    for row_index in 0..rows.len() {
        let continuing_spans = rows
            .get(row_index + 1)
            .map(|next_row| {
                next_row
                    .cells
                    .iter()
                    .filter(|cell| cell.is_vmerge_continue)
                    .map(|cell| (cell.col_index, cell.grid_span))
                    .collect::<Vec<_>>()
            })
            .unwrap_or_default();
        for cell in &mut rows[row_index].cells {
            cell.merge_with_below = continuing_spans.contains(&(cell.col_index, cell.grid_span));
        }
    }

    // Word gives a horizontal border its own height between two rows rather
    // than drawing it over their content. An exact height already includes
    // the band above its row, and the last row also carries the band below.
    let bands = border_bands(&rows, table_borders.as_ref());
    let table_bottom_band = |last_row: usize| {
        if last_row + 1 == num_rows {
            bands[num_rows]
        } else {
            0.0
        }
    };
    for (row_index, row) in rows.iter_mut().enumerate() {
        if !exact_rows[row_index] {
            row.height += bands[row_index];
        }
        row.height += table_bottom_band(row_index);
        for cell in &mut row.cells {
            cell.border_band_top = bands[row_index];
            cell.border_band_bottom = table_bottom_band(row_index);
        }
    }

    let row_heights = rows.iter().map(|row| row.height).collect::<Vec<_>>();
    for (row, height) in rows.iter_mut().zip(&row_heights) {
        for cell in &mut row.cells {
            cell.height = *height;
            cell.merged_height = *height;
        }
    }
    for (row_index, cell_index, last_row, required) in spans {
        // Word paints the bottom edge of a merge from the last cell it covers,
        // not from the cell that starts it.
        if last_row > row_index {
            let (col_index, grid_span) = {
                let restart = &rows[row_index].cells[cell_index];
                (restart.col_index, restart.grid_span)
            };
            let bottom = rows[last_row]
                .cells
                .iter()
                .find(|cell| {
                    cell.is_vmerge_continue
                        && cell.col_index == col_index
                        && cell.grid_span == grid_span
                })
                .and_then(|cell| cell.borders.as_ref()?.bottom.clone());
            rows[row_index].cells[cell_index]
                .borders
                .get_or_insert_with(CT_TblBorders::default)
                .bottom = bottom;
        }
        let restart = &mut rows[row_index].cells[cell_index];
        restart.merged_height = row_heights[row_index..=last_row].iter().sum();
        restart.is_last_row = last_row + 1 == num_rows;
        restart.border_band_bottom = table_bottom_band(last_row);
        // The painter draws the merged content between the band above the
        // merge and the table's bottom band, so exact rows clip to that box.
        restart.clip_content =
            required > restart.merged_height - restart.border_band_top - restart.border_band_bottom;
    }

    Ok((
        TableBlock {
            structure_id: None,
            col_widths,
            rows,
            header_row_indices,
            table_width,
            table_indent,
            borders: table_borders,
            bidi_visual,
            floating,
        },
        TableSemantics {
            rows: row_semantics,
        },
    ))
}

/// The border one side of a cell draws, or `None` when it draws none.
///
/// The cell's own edge wins over the table's, except that a `none` cell edge
/// on the outside of the table falls back to the table's outer edge.
pub(crate) fn resolved_cell_edge<'a>(
    cell_edge: Option<&'a CT_BorderEdge>,
    table_edge: Option<&'a CT_BorderEdge>,
    outer_edge: bool,
) -> Option<&'a CT_BorderEdge> {
    let edge = match cell_edge {
        Some(edge) if edge.val == ST_Border::None && outer_edge => table_edge?,
        Some(edge) => edge,
        None => table_edge?,
    };
    (edge.val != ST_Border::None).then_some(edge)
}

/// The height in points a horizontal border takes between two rows, as Word 16
/// reserves it. A single line is its `w:sz` eighths of a point. A compound
/// style is its lines and gaps, some as wide as the size and some at a width
/// Word fixes whatever the size, and a wave is a fixed height.
fn border_band(edge: &CT_BorderEdge) -> f64 {
    let width = edge.sz.unwrap_or(4) as f64 / 8.0;
    match edge.val {
        ST_Border::Double => 3.0 * width,
        ST_Border::Triple => 5.0 * width,
        ST_Border::ThinThickSmallGap | ST_Border::ThickThinSmallGap => width + 1.5,
        ST_Border::ThinThickMediumGap | ST_Border::ThickThinMediumGap => 2.0 * width,
        ST_Border::ThinThickLargeGap | ST_Border::ThickThinLargeGap => width + 2.25,
        ST_Border::ThreeDEmboss | ST_Border::ThreeDEngrave if width < 3.0 => width + 1.5,
        ST_Border::ThreeDEmboss | ST_Border::ThreeDEngrave => width + 3.0,
        ST_Border::Wave => 3.0,
        ST_Border::DoubleWave => 5.25,
        _ => width,
    }
}

/// The horizontal border band above each row, then the one below the last.
///
/// The band between two rows is the widest edge that meets there, from the
/// bottom edges of the row above and the top edges of the row below, so a
/// border the two rows share counts once. Word counts the edges of every
/// cell, the cells a vertical merge covers included, although it paints none
/// of them inside the merge.
fn border_bands(rows: &[TableRow], table_borders: Option<&CT_TblBorders>) -> Vec<f64> {
    let mut bands = vec![0.0f64; rows.len() + 1];
    for (row_index, row) in rows.iter().enumerate() {
        let first_row = row_index == 0;
        let last_row = row_index + 1 == rows.len();
        let table_top = table_borders.and_then(|borders| {
            if first_row {
                borders.top.as_ref()
            } else {
                borders.inside_h.as_ref()
            }
        });
        let table_bottom = table_borders.and_then(|borders| {
            if last_row {
                borders.bottom.as_ref()
            } else {
                borders.inside_h.as_ref()
            }
        });
        for cell in &row.cells {
            let cell_borders = cell.borders.as_ref();
            if let Some(edge) = resolved_cell_edge(
                cell_borders.and_then(|borders| borders.top.as_ref()),
                table_top,
                first_row,
            ) {
                bands[row_index] = bands[row_index].max(border_band(edge));
            }
            if let Some(edge) = resolved_cell_edge(
                cell_borders.and_then(|borders| borders.bottom.as_ref()),
                table_bottom,
                last_row,
            ) {
                bands[row_index + 1] = bands[row_index + 1].max(border_band(edge));
            }
        }
    }
    bands
}

/// The band a page break gives a row at the top of a page, or at the bottom
/// of one. Word closes a table on every page it crosses with the table's own
/// top and bottom borders, so the row's cell edges resolve against those, as
/// on the table's first and last row.
pub(crate) fn page_break_band(
    row: &TableRow,
    table_borders: Option<&CT_TblBorders>,
    top: bool,
) -> f64 {
    fn edge(borders: &CT_TblBorders, top: bool) -> Option<&CT_BorderEdge> {
        if top {
            borders.top.as_ref()
        } else {
            borders.bottom.as_ref()
        }
    }
    let table_edge = table_borders.and_then(|borders| edge(borders, top));
    row.cells
        .iter()
        .filter_map(|cell| {
            let own = cell.borders.as_ref().and_then(|borders| edge(borders, top));
            resolved_cell_edge(own, table_edge, true)
        })
        .map(border_band)
        .fold(0.0, f64::max)
}

/// `row` as the first row of the page a table break carries it to. Its band
/// becomes the one the top of a page gives it, which Word adds to the row
/// whatever its height rule, and its cells paint the table's top border.
pub(crate) fn row_opening_page(row: &TableRow, table_borders: Option<&CT_TblBorders>) -> TableRow {
    let band = page_break_band(row, table_borders, true);
    let growth = band - row.cells.first().map_or(0.0, |cell| cell.border_band_top);
    let mut opened = row.clone();
    opened.height += growth;
    for cell in &mut opened.cells {
        cell.height += growth;
        cell.merged_height += growth;
        cell.border_band_top = band;
        cell.is_first_row = true;
    }
    opened
}

/// `row` as the last row of a page before a table break, carrying below its
/// content the band and the edges that the table's bottom border gives it. A
/// merge that goes on to the next page keeps its box.
pub(crate) fn row_closing_page(row: &TableRow, table_borders: Option<&CT_TblBorders>) -> TableRow {
    let band = page_break_band(row, table_borders, false);
    let mut closed = row.clone();
    closed.height += band;
    for cell in &mut closed.cells {
        if cell.merge_with_below {
            continue;
        }
        cell.height += band;
        cell.merged_height += band;
        cell.border_band_bottom = band;
        cell.is_last_row = true;
    }
    closed
}

fn resolve_base_table_properties(table: &CT_Tbl, styles: &CT_Styles) -> CT_TblPr {
    let direct = table.properties.as_ref();
    let selected_style_id = direct
        .and_then(|properties| properties.style_id.as_deref())
        .or_else(|| {
            styles
                .get_default(rdocx_oxml::styles::StyleType::Table)
                .map(|style| style.style_id.as_str())
        });
    let mut chain = Vec::new();
    let mut current = selected_style_id;
    let mut visited = std::collections::HashSet::new();
    while let Some(style_id) = current.filter(|style_id| visited.insert(*style_id)) {
        let Some(style) = styles.get_by_id(style_id) else {
            break;
        };
        chain.push(style);
        current = style.based_on.as_deref();
    }

    let mut resolved = CT_TblPr {
        style_id: selected_style_id.map(str::to_owned),
        ..CT_TblPr::default()
    };
    for style in chain.into_iter().rev() {
        if let Some(properties) = &style.table_properties {
            overlay_table_properties(&mut resolved, properties);
        }
    }
    if let Some(properties) = direct {
        overlay_table_properties(&mut resolved, properties);
    }
    resolved
}

/// Resolve one row's effective properties base-first.
///
/// The table style's base `w:trPr` applies first, then every conditional
/// region that scopes a whole row, in the same ascending priority the cell
/// layers use, then the row's own direct properties.
fn resolve_row_properties(
    table: &CT_Tbl,
    styles: &CT_Styles,
    row: &CT_Row,
    row_index: usize,
    row_count: usize,
    column_count: usize,
) -> CT_TrPr {
    let mut resolved = CT_TrPr::default();
    if let Some(mut style_id) = table
        .properties
        .as_ref()
        .and_then(|properties| properties.style_id.as_deref())
        .or_else(|| {
            styles
                .get_default(rdocx_oxml::styles::StyleType::Table)
                .map(|style| style.style_id.as_str())
        })
    {
        // Most derived first, so `rev()` below applies from the base outwards.
        let mut chain = Vec::new();
        let mut visited = std::collections::HashSet::new();
        while visited.insert(style_id) {
            let Some(style) = styles.get_by_id(style_id) else {
                break;
            };
            chain.push(style);
            let Some(base) = style.based_on.as_deref() else {
                break;
            };
            style_id = base;
        }
        for style in chain.iter().rev() {
            if let Some(properties) = &style.table_row_properties {
                overlay_row_properties(&mut resolved, properties);
            }
        }
        let selectors = row
            .properties
            .as_ref()
            .and_then(|properties| properties.cnf_style.as_deref())
            .map(|value| vec![value])
            .unwrap_or_default();
        for region in applicable_table_regions(
            table,
            row_index,
            0,
            row_count,
            column_count,
            band_size(
                table
                    .properties
                    .as_ref()
                    .and_then(|properties| properties.row_band_size),
            ),
            1,
            &selectors,
        )
        .into_iter()
        .filter(|region| region_scopes_a_whole_row(*region))
        {
            for style in chain.iter().rev() {
                for conditional in style
                    .conditional_table_styles
                    .iter()
                    .filter(|conditional| conditional.region == Some(region))
                {
                    if let Some(properties) = &conditional.row_properties {
                        overlay_row_properties(&mut resolved, properties);
                    }
                }
            }
        }
    }
    if let Some(direct) = row.properties.as_ref() {
        overlay_row_properties(&mut resolved, direct);
    }
    resolved
}

/// Whether a conditional region applies to every cell of a row.
///
/// A column or corner region formats part of a row, so its `w:trPr` cannot
/// decide that row's height or grid offsets.
fn region_scopes_a_whole_row(region: TableStyleRegion) -> bool {
    matches!(
        region,
        TableStyleRegion::WholeTable
            | TableStyleRegion::Band1Horz
            | TableStyleRegion::Band2Horz
            | TableStyleRegion::FirstRow
            | TableStyleRegion::LastRow
    )
}

/// Absent means one row or column per band, which is what Word assumes.
fn band_size(size: Option<u32>) -> usize {
    size.unwrap_or(1).max(1) as usize
}

fn overlay_row_properties(target: &mut CT_TrPr, source: &CT_TrPr) {
    if source.height.is_some() {
        target.height = source.height;
    }
    if source.height_rule.is_some() {
        target.height_rule.clone_from(&source.height_rule);
    }
    if source.header.is_some() {
        target.header = source.header;
    }
    if source.jc.is_some() {
        target.jc = source.jc;
    }
    if source.grid_before.is_some() {
        target.grid_before = source.grid_before;
    }
    if source.grid_after.is_some() {
        target.grid_after = source.grid_after;
    }
    if source.width_before.is_some() {
        target.width_before.clone_from(&source.width_before);
    }
    if source.width_after.is_some() {
        target.width_after.clone_from(&source.width_after);
    }
    if source.cell_spacing.is_some() {
        target.cell_spacing.clone_from(&source.cell_spacing);
    }
    if source.hidden.is_some() {
        target.hidden = source.hidden;
    }
    if source.cant_split.is_some() {
        target.cant_split = source.cant_split;
    }
    if source.cnf_style.is_some() {
        target.cnf_style.clone_from(&source.cnf_style);
    }
}

/// Resolve a table measurement onto points, or `None` when the spelling
/// carries no length, which is `auto` or `nil`.
///
/// `percentage_base` is what a `pct` measurement is a percentage of. That is
/// the caller's width for a table width and the table's own width for a row
/// or cell measurement inside it.
fn table_width_to_pt(width: Option<&CT_TblWidth>, percentage_base: f64) -> Option<f64> {
    let width = width?;
    match width.width_type.as_str() {
        "dxa" => Some(width.w as f64 / 20.0),
        "pct" => Some(percentage_base * width.w as f64 / 5000.0),
        _ => None,
    }
}

fn overlay_table_properties(target: &mut CT_TblPr, source: &CT_TblPr) {
    if source.style_id.is_some() {
        target.style_id.clone_from(&source.style_id);
    }
    if source.row_band_size.is_some() {
        target.row_band_size = source.row_band_size;
    }
    if source.column_band_size.is_some() {
        target.column_band_size = source.column_band_size;
    }
    if source.width.is_some() {
        target.width.clone_from(&source.width);
    }
    if source.jc.is_some() {
        target.jc = source.jc;
    }
    if let Some(borders) = &source.borders {
        overlay_borders(&mut target.borders, borders);
    }
    if let Some(margins) = &source.cell_margin {
        let target = target.cell_margin.get_or_insert_default();
        if margins.top.is_some() {
            target.top = margins.top;
        }
        if margins.bottom.is_some() {
            target.bottom = margins.bottom;
        }
        if margins.left.is_some() {
            target.left = margins.left;
        }
        if margins.right.is_some() {
            target.right = margins.right;
        }
    }
    if source.layout.is_some() {
        target.layout.clone_from(&source.layout);
    }
    if source.float_position.is_some() {
        target.float_position.clone_from(&source.float_position);
    }
    if source.overlap.is_some() {
        target.overlap = source.overlap;
    }
    if source.bidi_visual.is_some() {
        target.bidi_visual = source.bidi_visual;
    }
    if source.cell_spacing.is_some() {
        target.cell_spacing.clone_from(&source.cell_spacing);
    }
    if source.indent.is_some() {
        target.indent.clone_from(&source.indent);
    }
    if source.shading.is_some() {
        target.shading.clone_from(&source.shading);
    }
    if let Some(look) = &source.look {
        let target = target.look.get_or_insert_default();
        if look.val.is_some() {
            target.val.clone_from(&look.val);
        }
        if look.first_row.is_some() {
            target.first_row = look.first_row;
        }
        if look.last_row.is_some() {
            target.last_row = look.last_row;
        }
        if look.first_column.is_some() {
            target.first_column = look.first_column;
        }
        if look.last_column.is_some() {
            target.last_column = look.last_column;
        }
        if look.no_h_band.is_some() {
            target.no_h_band = look.no_h_band;
        }
        if look.no_v_band.is_some() {
            target.no_v_band = look.no_v_band;
        }
    }
}

/// The line-length measure an auto-height rotated cell lays out against.
///
/// A rotated cell's line direction runs down the cell, so its measure is the
/// row height. An auto-height row has no height until its content produces
/// one, and Word grows such a row to the text rather than wrapping it, so the
/// cell lays out against a measure only a forced break ends a line inside.
/// Measuring it against the column width instead would wrap on the stacking
/// axis, and the stack would then be taller than the column is wide.
const VERTICAL_AUTO_MEASURE: f64 = AUTOFIT_MAX_TRIAL_WIDTH;

/// The trial width a maximum content measurement is taken against.
///
/// Wide enough that only a forced break ends a line, so the longest line is
/// the paragraph's natural width.
const AUTOFIT_MAX_TRIAL_WIDTH: f64 = 10_000.0;

/// The trial width a minimum content measurement is taken against.
///
/// One point, so every breakable opportunity is taken and the longest line is
/// the longest unbreakable run of text.
const AUTOFIT_MIN_TRIAL_WIDTH: f64 = 1.0;

/// Compute content-driven column widths, or `None` when autofit does not
/// engage and the declared grid stands.
///
/// Autofit engages only when the effective `w:tblLayout` is autofit or absent
/// **and** the effective `w:tblW` type is `auto` or absent. That is narrower
/// than the literal ECMA default, which applies autofit whenever
/// `w:tblLayout` is absent. The narrowing is deliberate: it keeps an authored
/// `dxa` or `pct` table on the declared grid, so adopting the wider predicate
/// stays a separate reviewed change rather than a side effect of this one.
///
/// Measurement runs the production cell path twice, once at a wide trial
/// width for the maximum content width and once at a minimal trial width for
/// the minimum. It consumes a clone of the numbering state and discards its
/// diagnostics, because the production pass that follows emits both for real.
fn autofit_column_widths(
    tbl: &CT_Tbl,
    authored_width_type: Option<&str>,
    available_width: f64,
    styles: &CT_Styles,
    input: &LayoutInput,
    media: &MediaRegistry,
    fm: &mut FontManager,
    num_state: &NumberingState,
    path: &[usize],
    doc_grid: Option<&CT_DocGrid>,
) -> Result<Option<Vec<f64>>> {
    let properties = tbl.properties.as_ref();
    // ECMA makes autofit the default when `w:tblLayout` is absent, but this
    // engagement is deliberately narrower and requires the element. An absent
    // layout is the shape almost every producer writes, 131 of the 141 tables
    // in the Word corpus, and treating it as autofit changes the reference
    // page count that `scripts/docx_authoring_conformance.py --private-required`
    // pins. Adopting the literal default is its own story with its own
    // reviewed geometry delta.
    let autofit_layout = matches!(
        properties.and_then(|properties| properties.layout.as_deref()),
        Some("autofit")
    );
    let auto_width = !matches!(authored_width_type, Some(kind) if kind != "auto");
    if !autofit_layout || !auto_width || available_width <= 0.0 {
        return Ok(None);
    }

    let source_rows = layout_table_rows(tbl, path);
    let column_count = tbl
        .grid
        .as_ref()
        .map(|grid| grid.columns.len())
        .filter(|count| *count > 0)
        .or_else(|| {
            source_rows.first().map(|(row, row_path)| {
                layout_row_cells(row, row_path)
                    .iter()
                    .map(|(cell, _)| {
                        cell.properties
                            .as_ref()
                            .and_then(|properties| properties.grid_span)
                            .unwrap_or(1) as usize
                    })
                    .sum::<usize>()
            })
        })
        .filter(|count| *count > 0)
        .unwrap_or(0);
    if column_count == 0 {
        return Ok(None);
    }

    let default_cell_margin = properties.and_then(|properties| properties.cell_margin.as_ref());
    let horizontal_margin = default_cell_margin
        .and_then(|margin| margin.left)
        .map_or(5.4, |value| value.to_pt())
        + default_cell_margin
            .and_then(|margin| margin.right)
            .map_or(5.4, |value| value.to_pt());

    let mut minima = vec![0.0f64; column_count];
    let mut maxima = vec![0.0f64; column_count];
    let row_count = source_rows.len();
    for (row_index, (row, row_path)) in source_rows.iter().enumerate() {
        // Resolved, not direct, so measurement assigns cells to the same grid
        // columns the production pass will.
        let mut col_index =
            (resolve_row_properties(tbl, styles, row, row_index, row_count, column_count)
                .grid_before
                .unwrap_or(0) as usize)
                .min(column_count);
        for (cell, cell_path) in &layout_row_cells(row, row_path) {
            let grid_span = cell
                .properties
                .as_ref()
                .and_then(|properties| properties.grid_span)
                .unwrap_or(1)
                .max(1) as usize;
            let end = (col_index + grid_span).min(column_count);
            if col_index >= end {
                col_index = end;
                continue;
            }
            let style_cell = resolve_table_style_cell(
                tbl,
                styles,
                row_index,
                col_index,
                row_count,
                column_count,
                &cell_conditional_selectors(row, cell),
            );
            let (minimum, mut maximum) = match declared_nested_grid_width(cell) {
                // Laying a nested table out at two trial widths would make it
                // autofit twice as well, so a table nested `n` deep would cost
                // three to the `n`. Its declared grid is the answer here, and
                // the production pass below still measures it for real.
                Some(declared) => {
                    let width = declared.min(available_width).max(0.0) + horizontal_margin;
                    (width, width)
                }
                None => {
                    // A rotated cell is measured in its transposed box, so the
                    // width it needs is the stacked height of its lines, not
                    // their length. Its line length belongs to the row height.
                    let rotated = cell_rotation(cell).is_some();
                    let mut measure = |trial_width: f64| -> Result<f64> {
                        let mut measurement_state = num_state.clone();
                        let mut measurement_diagnostics = Vec::new();
                        let (blocks, _) = layout_cell_content(
                            &cell.content,
                            trial_width,
                            styles,
                            input,
                            media,
                            fm,
                            &mut measurement_state,
                            &mut measurement_diagnostics,
                            None,
                            &WordStory::Document,
                            cell_path,
                            style_cell.paragraph_properties.as_ref(),
                            style_cell.run_properties.as_ref(),
                            doc_grid,
                        )?;
                        Ok(if rotated {
                            blocks.iter().map(CellBlock::total_height).sum::<f64>()
                        } else {
                            measured_content_width(&blocks)
                        })
                    };
                    if rotated {
                        // The stacked height runs the other way to a content
                        // width across the two trials: a one-point trial puts
                        // one word on every line and makes the stack as tall
                        // as it can be. Measuring a rotated cell at the narrow
                        // trial would hand its minimum the largest number it
                        // can produce, so it is measured once, at the width
                        // its lines will actually have.
                        let width = measure(AUTOFIT_MAX_TRIAL_WIDTH)?.max(0.0) + horizontal_margin;
                        (width, width)
                    } else {
                        let minimum =
                            measure(AUTOFIT_MIN_TRIAL_WIDTH)?.max(0.0) + horizontal_margin;
                        let maximum =
                            measure(AUTOFIT_MAX_TRIAL_WIDTH)?.max(0.0) + horizontal_margin;
                        (minimum, maximum)
                    }
                }
            };
            maximum = maximum.max(minimum);
            // A cell's own preferred width narrows the maximum, but never
            // below what the content needs to render its longest word.
            if let Some(preferred) = table_width_to_pt(
                cell.properties
                    .as_ref()
                    .and_then(|properties| properties.width.as_ref()),
                available_width,
            ) {
                maximum = preferred.clamp(minimum, maximum);
            }
            let span = (end - col_index) as f64;
            for column in col_index..end {
                minima[column] = minima[column].max(minimum / span);
                maxima[column] = maxima[column].max(maximum / span);
            }
            col_index = end;
        }
    }

    for column in 0..column_count {
        maxima[column] = maxima[column].max(minima[column]);
    }
    let total_min: f64 = minima.iter().sum();
    let total_max: f64 = maxima.iter().sum();
    if total_max < 0.01 {
        // Nothing measurable, so the declared grid is a better answer than a
        // table of zero-width columns.
        return Ok(None);
    }
    let widths = if total_max <= available_width {
        maxima
    } else if total_min >= available_width {
        let scale = available_width / total_min;
        minima.iter().map(|width| width * scale).collect()
    } else {
        let slack = (available_width - total_min) / (total_max - total_min);
        minima
            .iter()
            .zip(&maxima)
            .map(|(minimum, maximum)| minimum + (maximum - minimum) * slack)
            .collect()
    };
    Ok(Some(widths))
}

/// The widest declared grid among the tables nested directly in one cell, or
/// `None` when the cell holds no nested table.
///
/// This is what keeps autofit measurement linear in nesting depth. It is the
/// declared grid rather than a measured width, so a nested table's own
/// content does not widen the column that holds it.
fn declared_nested_grid_width(cell: &CT_Tc) -> Option<f64> {
    cell.content
        .iter()
        .filter_map(|item| match item {
            rdocx_oxml::table::CellContent::Table(nested) => Some(
                nested
                    .grid
                    .as_ref()
                    .map(|grid| {
                        grid.columns
                            .iter()
                            .map(|column| column.width.to_pt())
                            .sum::<f64>()
                    })
                    .unwrap_or(0.0),
            ),
            _ => None,
        })
        .reduce(f64::max)
}

/// The widest single line any block in a measured cell produced.
fn measured_content_width(blocks: &[CellBlock]) -> f64 {
    blocks
        .iter()
        .map(|block| match block {
            CellBlock::Paragraph(paragraph) => {
                paragraph.indent_left
                    + paragraph.indent_right
                    + paragraph
                        .lines
                        .iter()
                        .map(|line| line.width)
                        .fold(0.0f64, f64::max)
            }
            CellBlock::Table(table) => table.table_indent + table.table_width,
        })
        .fold(0.0f64, f64::max)
}

/// Compute column widths from CT_TblGrid, shrinking to the available width if
/// the declared grid overflows it.
///
/// A grid narrower than the text column keeps its declared width: Word renders
/// a deliberately narrow table at the size the author chose rather than
/// stretching it to the margins, and so do we.
fn compute_column_widths(
    grid: Option<&CT_TblGrid>,
    available_width: f64,
    table: &CT_Tbl,
    path: &[usize],
) -> Vec<f64> {
    let requested_width = table
        .properties
        .as_ref()
        .and_then(|properties| properties.width.as_ref())
        .and_then(|width| match width.width_type.as_str() {
            "dxa" if width.w > 0 => Some(width.w as f64 / 20.0),
            "pct" if width.w > 0 => Some(available_width * width.w as f64 / 5000.0),
            _ => None,
        })
        .map(|width| width.min(available_width));
    let target_width = requested_width.unwrap_or(available_width);
    match grid {
        Some(g) if !g.columns.is_empty() => {
            let widths: Vec<f64> = g.columns.iter().map(|c| c.width.to_pt()).collect();
            let total: f64 = widths.iter().sum();
            if total < 0.01 {
                // All zero widths — distribute equally based on column count
                let n = g.columns.len();
                vec![target_width / n as f64; n]
            } else if total > target_width + 1.0 || requested_width.is_some() {
                // Honor an explicit table width, or shrink an overflowing grid.
                let scale = target_width / total;
                widths.iter().map(|w| w * scale).collect()
            } else {
                widths
            }
        }
        _ => {
            // No grid defined — infer column count from the first row
            let num_cols = layout_table_rows(table, path)
                .first()
                .map(|(row, path)| {
                    layout_row_cells(row, path)
                        .iter()
                        .map(|(cell, _)| {
                            cell.properties
                                .as_ref()
                                .and_then(|p| p.grid_span)
                                .unwrap_or(1) as usize
                        })
                        .sum::<usize>()
                })
                .unwrap_or(1)
                .max(1);
            vec![target_width / num_cols as f64; num_cols]
        }
    }
}

/// Layout content within a table cell (paragraphs and nested tables).
///
/// Nested tables remain recursive blocks in source order.
fn layout_cell_content(
    content: &[rdocx_oxml::table::CellContent],
    available_width: f64,
    styles: &CT_Styles,
    input: &LayoutInput,
    media: &MediaRegistry,
    fm: &mut FontManager,
    num_state: &mut NumberingState,
    diagnostics: &mut Vec<Diagnostic>,
    sources: Option<&SourceRegistry>,
    story: &WordStory,
    cell_path: &[usize],
    table_style_ppr: Option<&rdocx_oxml::properties::CT_PPr>,
    table_style_rpr: Option<&rdocx_oxml::properties::CT_RPr>,
    doc_grid: Option<&CT_DocGrid>,
) -> Result<(Vec<CellBlock>, Vec<CellBlockSemantics>)> {
    use crate::engine;
    use rdocx_oxml::table::CellContent;

    let mut blocks = Vec::new();
    let mut semantics = Vec::new();
    for (content_index, item) in content.iter().enumerate() {
        let mut source_path = cell_path.to_vec();
        source_path.push(content_index);
        match item {
            CellContent::Paragraph(para) => {
                let source = sources.and_then(|sources| sources.id(story, &source_path));
                let (block, reflow_direction) = engine::layout_paragraph_with_source_in_table(
                    para,
                    available_width,
                    styles,
                    input,
                    media,
                    fm,
                    num_state,
                    diagnostics,
                    source,
                    table_style_ppr,
                    table_style_rpr,
                    doc_grid,
                )?;
                blocks.push(CellBlock::Paragraph(block));
                semantics.push(CellBlockSemantics::Paragraph(ParagraphSemantics {
                    source_node: source,
                    structure_id: None,
                    reflow_direction,
                }));
            }
            CellContent::Table(tbl) => {
                // Recursively lay out the nested table
                let (nested, nested_semantics) = layout_table_inner(
                    tbl,
                    available_width,
                    styles,
                    input,
                    media,
                    fm,
                    num_state,
                    diagnostics,
                    sources,
                    story,
                    &source_path,
                    doc_grid,
                )?;
                blocks.push(CellBlock::Table(nested));
                semantics.push(CellBlockSemantics::Table(nested_semantics));
            }
            CellContent::ContentControl(control) => layout_control_cell_content(
                control,
                available_width,
                styles,
                input,
                media,
                fm,
                num_state,
                diagnostics,
                sources,
                story,
                &source_path,
                table_style_ppr,
                table_style_rpr,
                doc_grid,
                &mut blocks,
                &mut semantics,
            )?,
        }
    }
    // Two consecutive paragraphs of a cell are one flow, so Word keeps the
    // larger of their facing spacing there as it does in the body. The space
    // before is reduced here to what it adds below the space after above it,
    // which keeps the cell height a plain sum. A nested table breaks the run.
    if !input.do_not_use_html_paragraph_auto_spacing {
        for index in 1..blocks.len() {
            if let [CellBlock::Paragraph(previous), CellBlock::Paragraph(next)] =
                &mut blocks[index - 1..=index]
            {
                next.space_before = (next.space_before - previous.space_after).max(0.0);
            }
        }
    }
    Ok((blocks, semantics))
}

#[allow(clippy::too_many_arguments)]
fn layout_control_cell_content(
    control: &CT_Sdt,
    available_width: f64,
    styles: &CT_Styles,
    input: &LayoutInput,
    media: &MediaRegistry,
    fm: &mut FontManager,
    num_state: &mut NumberingState,
    diagnostics: &mut Vec<Diagnostic>,
    sources: Option<&SourceRegistry>,
    story: &WordStory,
    path: &[usize],
    table_style_ppr: Option<&rdocx_oxml::properties::CT_PPr>,
    table_style_rpr: Option<&rdocx_oxml::properties::CT_RPr>,
    doc_grid: Option<&CT_DocGrid>,
    blocks: &mut Vec<CellBlock>,
    semantics: &mut Vec<CellBlockSemantics>,
) -> Result<()> {
    use crate::engine;

    for (content_index, content) in control.content.iter().enumerate() {
        let mut source_path = path.to_vec();
        source_path.push(content_index);
        match content {
            SdtContent::Paragraph(paragraph) => {
                let source = sources.and_then(|sources| sources.id(story, &source_path));
                let (block, reflow_direction) = engine::layout_paragraph_with_source_in_table(
                    paragraph,
                    available_width,
                    styles,
                    input,
                    media,
                    fm,
                    num_state,
                    diagnostics,
                    source,
                    table_style_ppr,
                    table_style_rpr,
                    doc_grid,
                )?;
                blocks.push(CellBlock::Paragraph(block));
                semantics.push(CellBlockSemantics::Paragraph(ParagraphSemantics {
                    source_node: source,
                    structure_id: None,
                    reflow_direction,
                }));
            }
            SdtContent::Table(table) => {
                let (nested, nested_semantics) = layout_table_inner(
                    table,
                    available_width,
                    styles,
                    input,
                    media,
                    fm,
                    num_state,
                    diagnostics,
                    sources,
                    story,
                    &source_path,
                    doc_grid,
                )?;
                blocks.push(CellBlock::Table(nested));
                semantics.push(CellBlockSemantics::Table(nested_semantics));
            }
            SdtContent::ContentControl(control) => layout_control_cell_content(
                control,
                available_width,
                styles,
                input,
                media,
                fm,
                num_state,
                diagnostics,
                sources,
                story,
                &source_path,
                table_style_ppr,
                table_style_rpr,
                doc_grid,
                blocks,
                semantics,
            )?,
            SdtContent::Row(_)
            | SdtContent::Cell(_)
            | SdtContent::Run(_)
            | SdtContent::RawXml(_) => {}
        }
    }
    Ok(())
}

#[derive(Default)]
struct ResolvedTableCellStyle {
    paragraph_properties: Option<rdocx_oxml::properties::CT_PPr>,
    run_properties: Option<rdocx_oxml::properties::CT_RPr>,
    borders: Option<CT_TblBorders>,
    shading: Option<rdocx_oxml::properties::CT_Shd>,
}

/// The conditional-region selectors that apply to one cell.
///
/// Word writes `w:cnfStyle` on the row, on the cell and on every paragraph
/// inside the cell, and a bit set anywhere selects the region. The cell's own
/// paragraphs are collected here because a table style resolves once per cell,
/// before its paragraphs are laid out.
fn cell_conditional_selectors<'a>(row: &'a CT_Row, cell: &'a CT_Tc) -> Vec<&'a str> {
    let mut selectors = Vec::new();
    if let Some(value) = row
        .properties
        .as_ref()
        .and_then(|properties| properties.cnf_style.as_deref())
    {
        selectors.push(value);
    }
    if let Some(value) = cell
        .properties
        .as_ref()
        .and_then(|properties| properties.cnf_style.as_deref())
    {
        selectors.push(value);
    }
    for item in &cell.content {
        if let rdocx_oxml::table::CellContent::Paragraph(paragraph) = item
            && let Some(value) = paragraph
                .properties
                .as_ref()
                .and_then(|properties| properties.cnf_style.as_deref())
        {
            selectors.push(value);
        }
    }
    selectors
}

fn resolve_table_style_cell(
    table: &CT_Tbl,
    styles: &CT_Styles,
    row: usize,
    column: usize,
    row_count: usize,
    column_count: usize,
    selectors: &[&str],
) -> ResolvedTableCellStyle {
    let Some(mut style_id) = table
        .properties
        .as_ref()
        .and_then(|p| p.style_id.as_deref())
        .or_else(|| {
            styles
                .get_default(rdocx_oxml::styles::StyleType::Table)
                .map(|style| style.style_id.as_str())
        })
    else {
        return ResolvedTableCellStyle::default();
    };
    // Most derived first, so `rev()` below applies from the base outwards.
    let mut chain = Vec::new();
    let mut visited = std::collections::HashSet::new();
    while visited.insert(style_id) {
        let Some(style) = styles.get_by_id(style_id) else {
            break;
        };
        chain.push(style);
        let Some(base) = style.based_on.as_deref() else {
            break;
        };
        style_id = base;
    }

    let mut resolved = ResolvedTableCellStyle::default();
    // The style's own property layers, applied base first.
    for style in chain.iter().rev() {
        if let Some(properties) = &style.ppr {
            resolved
                .paragraph_properties
                .get_or_insert_with(rdocx_oxml::properties::CT_PPr::default)
                .merge_from(properties);
        }
        if let Some(properties) = &style.rpr {
            resolved
                .run_properties
                .get_or_insert_with(rdocx_oxml::properties::CT_RPr::default)
                .merge_from(properties);
        }
        if let Some(borders) = style
            .table_properties
            .as_ref()
            .and_then(|properties| properties.borders.as_ref())
        {
            overlay_borders(&mut resolved.borders, borders);
        }
        if let Some(shading) = style
            .table_properties
            .as_ref()
            .and_then(|properties| properties.shading.as_ref())
        {
            resolved.shading = Some(shading.clone());
        }
    }

    // `table` already carries the style chain's table properties, resolved by
    // `resolve_base_table_properties` before the rows are laid out, so the
    // band sizes here are the resolved ones.
    let row_band_size = band_size(
        table
            .properties
            .as_ref()
            .and_then(|properties| properties.row_band_size),
    );
    let column_band_size = band_size(
        table
            .properties
            .as_ref()
            .and_then(|properties| properties.column_band_size),
    );

    // Regions apply in ascending priority. The whole `basedOn` chain is
    // flattened for one region before the next region starts, so a base
    // style's `firstRow` still beats a derived style's `wholeTable`.
    for region in applicable_table_regions(
        table,
        row,
        column,
        row_count,
        column_count,
        row_band_size,
        column_band_size,
        selectors,
    ) {
        for style in chain.iter().rev() {
            for conditional in style
                .conditional_table_styles
                .iter()
                .filter(|conditional| conditional.region == Some(region))
            {
                apply_conditional_region(&mut resolved, conditional);
            }
        }
    }
    resolved
}

/// Overlay one conditional region's layers onto the resolved cell style.
///
/// The region's `w:trPr` is modeled and round-tripped but not applied. Row
/// geometry from a conditional region belongs to F-268a.
fn apply_conditional_region(
    resolved: &mut ResolvedTableCellStyle,
    conditional: &rdocx_oxml::styles::CT_TblStylePr,
) {
    if let Some(properties) = &conditional.paragraph_properties {
        resolved
            .paragraph_properties
            .get_or_insert_with(rdocx_oxml::properties::CT_PPr::default)
            .merge_from(properties);
    }
    if let Some(properties) = &conditional.run_properties {
        resolved
            .run_properties
            .get_or_insert_with(rdocx_oxml::properties::CT_RPr::default)
            .merge_from(properties);
    }
    if let Some(borders) = conditional
        .cell_properties
        .as_ref()
        .and_then(|properties| properties.borders.as_ref())
        .or_else(|| {
            conditional
                .table_properties
                .as_ref()
                .and_then(|properties| properties.borders.as_ref())
        })
    {
        overlay_borders(&mut resolved.borders, borders);
    }
    if let Some(shading) = conditional
        .cell_properties
        .as_ref()
        .and_then(|properties| properties.shading.as_ref())
        .or_else(|| {
            conditional
                .table_properties
                .as_ref()
                .and_then(|properties| properties.shading.as_ref())
        })
    {
        resolved.shading = Some(shading.clone());
    }
}

fn applicable_table_regions(
    table: &CT_Tbl,
    row: usize,
    column: usize,
    row_count: usize,
    column_count: usize,
    row_band_size: usize,
    column_band_size: usize,
    selectors: &[&str],
) -> Vec<TableStyleRegion> {
    let cnf = |index: usize| {
        selectors
            .iter()
            .any(|value| value.as_bytes().get(index) == Some(&b'1'))
    };
    let look = table
        .properties
        .as_ref()
        .and_then(|properties| properties.look.as_ref());
    let enabled = |explicit: Option<bool>, mask: u16, default: bool| {
        explicit.unwrap_or_else(|| {
            look.and_then(|look| look.val.as_deref())
                .and_then(|value| u16::from_str_radix(value, 16).ok())
                .map_or(default, |value| value & mask != 0)
        })
    };
    let heads_rows = enabled(look.and_then(|look| look.first_row), 0x20, false);
    let heads_columns = enabled(look.and_then(|look| look.first_column), 0x80, false);
    let first_row = (heads_rows && row == 0) || cnf(0);
    let last_row = (enabled(look.and_then(|look| look.last_row), 0x40, false)
        && row + 1 == row_count)
        || cnf(1);
    let first_column = (heads_columns && column == 0) || cnf(2);
    let last_column = (enabled(look.and_then(|look| look.last_column), 0x100, false)
        && column + 1 == column_count)
        || cnf(3);
    let no_h_band = enabled(look.and_then(|look| look.no_h_band), 0x200, false);
    let no_v_band = enabled(look.and_then(|look| look.no_v_band), 0x400, false);

    let mut regions = vec![TableStyleRegion::WholeTable];
    // Banding counts whole bands of the resolved size, and starts after the
    // header row the look designates. The header row itself is in no band.
    if cnf(6) {
        regions.push(TableStyleRegion::Band1Horz);
    } else if cnf(7) {
        regions.push(TableStyleRegion::Band2Horz);
    } else if !no_h_band && let Some(offset) = row.checked_sub(usize::from(heads_rows)) {
        regions.push(if (offset / row_band_size).is_multiple_of(2) {
            TableStyleRegion::Band1Horz
        } else {
            TableStyleRegion::Band2Horz
        });
    }
    if cnf(4) {
        regions.push(TableStyleRegion::Band1Vert);
    } else if cnf(5) {
        regions.push(TableStyleRegion::Band2Vert);
    } else if !no_v_band && let Some(offset) = column.checked_sub(usize::from(heads_columns)) {
        regions.push(if (offset / column_band_size).is_multiple_of(2) {
            TableStyleRegion::Band1Vert
        } else {
            TableStyleRegion::Band2Vert
        });
    }
    if first_column {
        regions.push(TableStyleRegion::FirstCol);
    }
    if last_column {
        regions.push(TableStyleRegion::LastCol);
    }
    if first_row {
        regions.push(TableStyleRegion::FirstRow);
    }
    if last_row {
        regions.push(TableStyleRegion::LastRow);
    }
    if cnf(9) {
        regions.push(TableStyleRegion::NwCell);
    } else if cnf(8) {
        regions.push(TableStyleRegion::NeCell);
    } else if cnf(11) {
        regions.push(TableStyleRegion::SwCell);
    } else if cnf(10) {
        regions.push(TableStyleRegion::SeCell);
    } else {
        match (first_row, last_row, first_column, last_column) {
            (true, _, true, _) => regions.push(TableStyleRegion::NwCell),
            (true, _, _, true) => regions.push(TableStyleRegion::NeCell),
            (_, true, true, _) => regions.push(TableStyleRegion::SwCell),
            (_, true, _, true) => regions.push(TableStyleRegion::SeCell),
            _ => {}
        }
    }
    // Declaration order is priority order, so sorting is what makes the
    // precedence a property of the type rather than of the push order above.
    regions.sort_unstable();
    regions.dedup();
    regions
}

fn overlay_borders(target: &mut Option<CT_TblBorders>, source: &CT_TblBorders) {
    let target = target.get_or_insert_with(CT_TblBorders::default);
    if source.top.is_some() {
        target.top = source.top.clone();
    }
    if source.bottom.is_some() {
        target.bottom = source.bottom.clone();
    }
    if source.left.is_some() {
        target.left = source.left.clone();
    }
    if source.right.is_some() {
        target.right = source.right.clone();
    }
    if source.inside_h.is_some() {
        target.inside_h = source.inside_h.clone();
    }
    if source.inside_v.is_some() {
        target.inside_v = source.inside_v.clone();
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use rdocx_oxml::table::{
        CT_Row, CT_TblCellMar, CT_TblGrid, CT_TblGridCol, CT_TblLook, CT_TblPr, CT_TblWidth, CT_Tc,
        CT_TcPr, CT_TrPr,
    };
    use rdocx_oxml::units::Twips;

    fn layout_with_defaults(table: &CT_Tbl, width: f64) -> TableBlock {
        let styles = CT_Styles::default();
        layout_with_styles(table, width, &styles)
    }

    fn layout_with_styles(table: &CT_Tbl, width: f64, styles: &CT_Styles) -> TableBlock {
        let input = LayoutInput {
            revision_view: crate::input::RevisionView::Accepted,
            automatic_hyphenation: false,
            mirror_margins: false,
            gutter_at_top: false,
            do_not_use_html_paragraph_auto_spacing: false,
            default_tab_stop: None,
            math_properties: None,
            document: rdocx_oxml::document::CT_Document {
                body: rdocx_oxml::document::CT_Body {
                    content: Vec::new(),
                    sect_pr: None,
                },
                extra_namespaces: Vec::new(),
                background_xml: None,
                background_extra_xml: Vec::new(),
            },
            styles: styles.clone(),
            numbering: None,
            headers: std::collections::HashMap::new(),
            footers: std::collections::HashMap::new(),
            images: std::collections::HashMap::new(),
            charts: std::collections::HashMap::new(),
            chart_theme: oxml_drawing::theme::CT_OfficeStyleSheet::office_default(),
            chart_color_map: oxml_drawing::color::ColorMap::default(),
            hyperlink_urls: std::collections::HashMap::new(),
            footnotes: None,
            endnotes: None,
            core_properties: None,
            theme: None,
            fonts: Vec::new(),
        };
        let media = MediaRegistry::new(&input.images);
        let mut font_manager = FontManager::new();
        let mut numbering = NumberingState::new();
        layout_table(
            table,
            width,
            styles,
            &input,
            &media,
            &mut font_manager,
            &mut numbering,
            &mut Vec::new(),
            None,
        )
        .unwrap()
    }

    #[test]
    fn narrow_grid_keeps_its_declared_width() {
        let tbl = CT_Tbl::new();
        let grid = CT_TblGrid {
            columns: vec![
                CT_TblGridCol { width: Twips(2880) }, // 2 inches = 144pt
                CT_TblGridCol { width: Twips(2880) },
            ],
            ..Default::default()
        };

        // 288pt total in a 468pt text column: the author asked for a narrow
        // table, so it must not be stretched to the margins.
        let widths = compute_column_widths(Some(&grid), 468.0, &tbl, &[]);

        assert_eq!(widths.len(), 2);
        let total: f64 = widths.iter().sum();
        assert!((total - 288.0).abs() < 1.0, "got {total}");
    }

    #[test]
    fn historical_table_grid_never_changes_active_column_widths() {
        let table = CT_Tbl::new();
        let grid = CT_TblGrid {
            columns: vec![
                CT_TblGridCol { width: Twips(1440) },
                CT_TblGridCol { width: Twips(2880) },
            ],
            grid_change_xml: Some(
                br#"<w:tblGridChange w:id="4"><w:tblGrid><w:gridCol w:w="9000"/><w:gridCol w:w="9000"/></w:tblGrid></w:tblGridChange>"#
                    .to_vec(),
            ),
            ..CT_TblGrid::default()
        };

        assert_eq!(
            compute_column_widths(Some(&grid), 468.0, &table, &[]),
            vec![72.0, 144.0]
        );
    }

    #[test]
    fn overflowing_grid_is_scaled_down_to_fit() {
        let tbl = CT_Tbl::new();
        let grid = CT_TblGrid {
            columns: vec![
                CT_TblGridCol { width: Twips(7200) }, // 5 inches = 360pt
                CT_TblGridCol { width: Twips(7200) },
            ],
            ..Default::default()
        };

        // 720pt total will not fit a 468pt column, so scale it down.
        let widths = compute_column_widths(Some(&grid), 468.0, &tbl, &[]);

        let total: f64 = widths.iter().sum();
        assert!((total - 468.0).abs() < 1.0, "got {total}");
        // Proportions are preserved.
        assert!((widths[0] - widths[1]).abs() < 0.01);
    }

    #[test]
    fn column_widths_no_grid() {
        let tbl = CT_Tbl::new();
        let widths = compute_column_widths(None, 468.0, &tbl, &[]);
        assert_eq!(widths.len(), 1);
        assert!((widths[0] - 468.0).abs() < 0.01);
    }

    #[test]
    fn column_widths_zero_grid() {
        let tbl = CT_Tbl::new();
        let grid = CT_TblGrid {
            columns: vec![
                CT_TblGridCol { width: Twips(0) },
                CT_TblGridCol { width: Twips(0) },
                CT_TblGridCol { width: Twips(0) },
            ],
            ..Default::default()
        };
        let widths = compute_column_widths(Some(&grid), 468.0, &tbl, &[]);
        assert_eq!(widths.len(), 3);
        for w in &widths {
            assert!((w - 156.0).abs() < 0.01);
        }
    }

    #[test]
    fn column_widths_inferred_from_rows() {
        use rdocx_oxml::table::{CT_Row, CT_Tc};
        let mut tbl = CT_Tbl::new();
        let mut row = CT_Row::new();
        row.cells.push(CT_Tc::new());
        row.cells.push(CT_Tc::new());
        row.cells.push(CT_Tc::new());
        tbl.rows.push(row);
        let widths = compute_column_widths(None, 300.0, &tbl, &[]);
        assert_eq!(widths.len(), 3);
        for w in &widths {
            assert!((w - 100.0).abs() < 0.01);
        }
    }

    #[test]
    fn nested_tables_remain_recursive_cell_blocks() {
        use rdocx_oxml::table::{CT_Row, CT_Tbl, CT_Tc, CellContent};

        // Build an outer table with one cell containing a nested table
        let mut outer = CT_Tbl::new();
        outer.grid = Some(CT_TblGrid {
            columns: vec![CT_TblGridCol { width: Twips(4680) }], // 3.25"
            ..Default::default()
        });

        let mut outer_row = CT_Row::new();
        let mut outer_cell = CT_Tc::new();
        outer_cell.paragraphs_mut()[0].add_run("Before nested");

        // Nested table with 2 columns
        let mut nested = CT_Tbl::new();
        nested.grid = Some(CT_TblGrid {
            columns: vec![
                CT_TblGridCol { width: Twips(2000) },
                CT_TblGridCol { width: Twips(2000) },
            ],
            ..Default::default()
        });
        let mut nr = CT_Row::new();
        let mut nc1 = CT_Tc::new();
        nc1.paragraphs_mut()[0].add_run("N1");
        let mut nc2 = CT_Tc::new();
        nc2.paragraphs_mut()[0].add_run("N2");
        nr.cells.push(nc1);
        nr.cells.push(nc2);
        nested.rows.push(nr);

        outer_cell.content.push(CellContent::Table(nested));
        outer_row.cells.push(outer_cell);
        outer.rows.push(outer_row);

        // Layout with default styles
        let styles = rdocx_oxml::styles::CT_Styles::default();
        let input = crate::input::LayoutInput {
            revision_view: crate::input::RevisionView::Accepted,
            automatic_hyphenation: false,
            mirror_margins: false,
            gutter_at_top: false,
            do_not_use_html_paragraph_auto_spacing: false,
            default_tab_stop: None,
            math_properties: None,
            document: rdocx_oxml::document::CT_Document {
                body: rdocx_oxml::document::CT_Body {
                    content: Vec::new(),
                    sect_pr: None,
                },
                extra_namespaces: Vec::new(),
                background_xml: None,
                background_extra_xml: Vec::new(),
            },
            styles: styles.clone(),
            numbering: None,
            headers: std::collections::HashMap::new(),
            footers: std::collections::HashMap::new(),
            images: std::collections::HashMap::new(),
            charts: std::collections::HashMap::new(),
            chart_theme: oxml_drawing::theme::CT_OfficeStyleSheet::office_default(),
            chart_color_map: oxml_drawing::color::ColorMap::default(),
            hyperlink_urls: std::collections::HashMap::new(),
            footnotes: None,
            endnotes: None,
            core_properties: None,
            theme: None,
            fonts: Vec::new(),
        };

        let mut fm = FontManager::new();
        let mut num_state = crate::style_resolver::NumberingState::new();
        let mut diagnostics = Vec::new();
        let media = MediaRegistry::new(&input.images);

        let result = layout_table(
            &outer,
            234.0,
            &styles,
            &input,
            &media,
            &mut fm,
            &mut num_state,
            &mut diagnostics,
            None,
        );
        assert!(result.is_ok());
        let block = result.unwrap();

        // Outer table should have 1 row, 1 cell
        assert_eq!(block.rows.len(), 1);
        assert_eq!(block.rows[0].cells.len(), 1);

        // The nested table remains a distinct block at its source-order slot.
        let cell = &block.rows[0].cells[0];
        assert_eq!(cell.blocks.len(), 2);
        assert!(matches!(cell.blocks[0], CellBlock::Paragraph(_)));
        assert!(matches!(cell.blocks[1], CellBlock::Table(_)));

        // This table declares neither a width nor a layout mode. Autofit
        // engagement requires the `w:tblLayout` element, so the width comes
        // from the declared grid and never exceeds the caller's width.
        assert!(block.table_width > 0.0);
        assert!(block.table_width <= 234.0);
    }

    #[test]
    fn compound_borders_reserve_the_band_word_gives_their_lines_and_gaps() {
        // Word 16 measurements: the points a top, inside or bottom border
        // takes at w:sz 4 and at w:sz 24.
        for (style, at_4, at_24) in [
            (ST_Border::Single, 0.5, 3.0),
            (ST_Border::Dotted, 0.5, 3.0),
            (ST_Border::Outset, 0.5, 3.0),
            (ST_Border::Double, 1.5, 9.0),
            (ST_Border::Triple, 2.5, 15.0),
            (ST_Border::ThinThickSmallGap, 2.0, 4.5),
            (ST_Border::ThickThinSmallGap, 2.0, 4.5),
            (ST_Border::ThinThickMediumGap, 1.0, 6.0),
            (ST_Border::ThickThinMediumGap, 1.0, 6.0),
            (ST_Border::ThinThickLargeGap, 2.75, 5.25),
            (ST_Border::ThickThinLargeGap, 2.75, 5.25),
            (ST_Border::ThreeDEmboss, 2.0, 6.0),
            (ST_Border::ThreeDEngrave, 2.0, 6.0),
            (ST_Border::Wave, 3.0, 3.0),
            (ST_Border::DoubleWave, 5.25, 5.25),
        ] {
            for (sz, band) in [(4, at_4), (24, at_24)] {
                let mut edge = CT_BorderEdge::new(style);
                edge.sz = Some(sz);
                assert_eq!(border_band(&edge), band, "{style:?} at w:sz {sz}");
            }
        }
    }

    #[test]
    fn a_merge_over_exact_rows_clips_to_the_box_between_its_bands() {
        // Two exact 20 point rows under 3 point borders hold a 43 point merge
        // whose painted box is 37 points, so 39 points of content must clip.
        let edge = || {
            let mut edge = CT_BorderEdge::new(ST_Border::Single);
            edge.sz = Some(24);
            edge
        };
        let mut table = CT_Tbl::new();
        table.properties = Some(CT_TblPr {
            borders: Some(CT_TblBorders {
                top: Some(edge()),
                bottom: Some(edge()),
                left: Some(edge()),
                right: Some(edge()),
                inside_h: Some(edge()),
                inside_v: Some(edge()),
                ..Default::default()
            }),
            ..Default::default()
        });
        table.grid = Some(CT_TblGrid {
            columns: vec![CT_TblGridCol { width: Twips(3000) }],
            ..Default::default()
        });
        let exact_row = || {
            let mut row = CT_Row::new();
            row.properties = Some(CT_TrPr {
                height: Some(Twips(400)),
                height_rule: Some("exact".to_owned()),
                ..Default::default()
            });
            row
        };
        let mut restart = CT_Tc::new();
        restart.properties = Some(CT_TcPr {
            v_merge: Some(VMerge::Restart),
            ..Default::default()
        });
        restart.content.clear();
        for text in ["one", "two", "three"] {
            let mut paragraph = rdocx_oxml::text::CT_P::new();
            paragraph.properties = Some(rdocx_oxml::properties::CT_PPr {
                line_spacing: Some(Twips(260)),
                line_rule: Some("exact".to_owned()),
                space_before: Some(Twips(0)),
                space_after: Some(Twips(0)),
                ..Default::default()
            });
            paragraph.add_run(text);
            restart
                .content
                .push(rdocx_oxml::table::CellContent::Paragraph(paragraph));
        }
        let mut first = exact_row();
        first.cells.push(restart);
        let mut continuation = CT_Tc::new();
        continuation.properties = Some(CT_TcPr {
            v_merge: Some(VMerge::Continue),
            ..Default::default()
        });
        let mut second = exact_row();
        second.cells.push(continuation);
        table.rows = vec![first, second];

        let block = layout_with_defaults(&table, 150.0);

        let restart = &block.rows[0].cells[0];
        assert_eq!(restart.merged_height, 43.0);
        assert_eq!(
            (restart.border_band_top, restart.border_band_bottom),
            (3.0, 3.0)
        );
        assert!(restart.clip_content);
    }

    #[test]
    fn vertical_merges_and_row_height_rules_share_the_exact_grid_span() {
        let mut table = CT_Tbl::new();
        table.grid = Some(CT_TblGrid {
            columns: vec![
                CT_TblGridCol { width: Twips(600) },
                CT_TblGridCol { width: Twips(600) },
            ],
            ..Default::default()
        });

        let mut exact_row = CT_Row::new();
        exact_row.properties = Some(CT_TrPr {
            height: Some(Twips(200)),
            height_rule: Some("exact".to_owned()),
            ..Default::default()
        });
        let mut restart = CT_Tc::new();
        restart.properties = Some(CT_TcPr {
            grid_span: Some(2),
            v_merge: Some(VMerge::Restart),
            ..Default::default()
        });
        restart.paragraphs_mut()[0].add_run(
            "merged content wraps across enough words to require both rows and grow only a minimum row",
        );
        exact_row.cells.push(restart);

        let mut minimum_row = CT_Row::new();
        minimum_row.properties = Some(CT_TrPr {
            height: Some(Twips(200)),
            height_rule: Some("atLeast".to_owned()),
            ..Default::default()
        });
        let mut continuation = CT_Tc::new();
        continuation.properties = Some(CT_TcPr {
            grid_span: Some(2),
            v_merge: Some(VMerge::Continue),
            ..Default::default()
        });
        minimum_row.cells.push(continuation);
        table.rows = vec![exact_row, minimum_row];

        let block = layout_with_defaults(&table, 60.0);
        assert_eq!(block.rows[0].height, 10.0, "exact row must stay pinned");
        assert!(block.rows[1].height >= 10.0);
        let restart = &block.rows[0].cells[0];
        assert_eq!(restart.grid_span, 2);
        assert!(restart.merge_with_below);
        assert_eq!(
            restart.merged_height,
            block.rows[0].height + block.rows[1].height
        );
        assert!(
            restart.is_last_row,
            "merge ends on the table's outer bottom"
        );
        assert!(block.rows[1].cells[0].is_vmerge_continue);

        let mut minimum_merge = CT_Tbl::new();
        minimum_merge.grid = Some(CT_TblGrid {
            columns: vec![CT_TblGridCol { width: Twips(600) }],
            ..Default::default()
        });
        let mut restart_row = CT_Row::new();
        let mut restart = CT_Tc::new();
        restart.properties = Some(CT_TcPr {
            v_merge: Some(VMerge::Restart),
            ..Default::default()
        });
        restart.paragraphs_mut()[0]
            .add_run("merged content grows the final eligible row in this span");
        restart_row.cells.push(restart);
        let mut final_row = CT_Row::new();
        final_row.properties = Some(CT_TrPr {
            height: Some(Twips(200)),
            height_rule: Some("atLeast".to_owned()),
            ..Default::default()
        });
        let mut continuation = CT_Tc::new();
        continuation.properties = Some(CT_TcPr {
            v_merge: Some(VMerge::Continue),
            ..Default::default()
        });
        final_row.cells.push(continuation);
        minimum_merge.rows = vec![restart_row, final_row];

        let minimum_block = layout_with_defaults(&minimum_merge, 30.0);
        assert_eq!(minimum_block.rows[0].height, 0.0);
        assert!(
            minimum_block.rows[1].height > 10.0,
            "restart content grows the final non-exact row"
        );
    }

    #[test]
    fn table_style_cascade_resolves_borders_and_paragraph_spacing() {
        let styles = CT_Styles::from_xml(
            format!(
                r#"<w:styles xmlns:w="{}"><w:style w:type="table" w:styleId="Base"><w:pPr><w:spacing w:after="80"/></w:pPr><w:tblPr><w:tblBorders><w:left w:val="single" w:sz="8" w:color="AA0000"/></w:tblBorders></w:tblPr></w:style><w:style w:type="table" w:styleId="Dense"><w:basedOn w:val="Base"/><w:pPr><w:spacing w:after="40"/></w:pPr><w:tblStylePr w:type="firstRow"><w:pPr><w:spacing w:after="0"/></w:pPr><w:tcPr><w:tcBorders><w:top w:val="double" w:sz="12" w:color="0000AA"/></w:tcBorders><w:shd w:val="clear" w:fill="DDEEFF"/></w:tcPr></w:tblStylePr><w:tblStylePr w:type="firstCol"><w:tcPr><w:shd w:val="clear" w:fill="CCFFCC"/></w:tcPr></w:tblStylePr></w:style></w:styles>"#,
                rdocx_oxml::namespace::W_NS
            )
            .as_bytes(),
        )
        .unwrap();
        let mut table = CT_Tbl::new();
        table.properties = Some(CT_TblPr {
            style_id: Some("Dense".to_owned()),
            look: Some(CT_TblLook {
                first_row: Some(false),
                first_column: Some(false),
                no_h_band: Some(true),
                no_v_band: Some(true),
                ..Default::default()
            }),
            ..Default::default()
        });
        table.grid = Some(CT_TblGrid {
            columns: vec![CT_TblGridCol { width: Twips(1200) }],
            ..Default::default()
        });
        for (index, text) in ["header", "body"].into_iter().enumerate() {
            let mut row = CT_Row::new();
            let mut cell = CT_Tc::new();
            if index == 0 {
                row.properties = Some(CT_TrPr {
                    cnf_style: Some("100000000000".to_owned()),
                    ..Default::default()
                });
            } else {
                cell.properties = Some(CT_TcPr {
                    cnf_style: Some("001000000000".to_owned()),
                    ..Default::default()
                });
            }
            cell.paragraphs_mut()[0].add_run(text);
            row.cells.push(cell);
            table.rows.push(row);
        }

        let block = layout_with_styles(&table, 60.0, &styles);
        let CellBlock::Paragraph(header) = &block.rows[0].cells[0].blocks[0] else {
            panic!("header paragraph");
        };
        let CellBlock::Paragraph(body) = &block.rows[1].cells[0].blocks[0] else {
            panic!("body paragraph");
        };
        assert_eq!(header.space_after, 0.0);
        assert_eq!(body.space_after, 2.0);
        assert_eq!(
            block.rows[0].cells[0].shading,
            Some(Color::from_hex("DDEEFF"))
        );
        assert_eq!(
            block.rows[1].cells[0].shading,
            Some(Color::from_hex("CCFFCC"))
        );
        let header_borders = block.rows[0].cells[0].borders.as_ref().unwrap();
        assert_eq!(header_borders.top.as_ref().unwrap().sz, Some(12));
        assert_eq!(
            header_borders.top.as_ref().unwrap().color.as_deref(),
            Some("0000AA")
        );
        assert_eq!(
            header_borders.left.as_ref().unwrap().color.as_deref(),
            Some("AA0000")
        );
    }

    #[test]
    fn table_without_an_explicit_style_uses_the_authored_default() {
        let styles = CT_Styles::from_xml(
            format!(
                r#"<w:styles xmlns:w="{}"><w:style w:type="table" w:styleId="CorpusTable" w:default="1"><w:tblPr><w:shd w:val="clear" w:fill="D9EAF7"/></w:tblPr></w:style></w:styles>"#,
                rdocx_oxml::namespace::W_NS
            )
            .as_bytes(),
        )
        .unwrap();
        let resolved = resolve_table_style_cell(&CT_Tbl::new(), &styles, 0, 0, 1, 1, &[]);
        assert_eq!(
            resolved
                .shading
                .as_ref()
                .and_then(|shading| shading.fill.as_deref()),
            Some("D9EAF7")
        );

        let styles = CT_Styles::from_xml(
            format!(
                r#"<w:styles xmlns:w="{}"><w:style w:type="table" w:styleId="Sized" w:default="1"><w:tblPr><w:tblW w:w="2000" w:type="dxa"/><w:jc w:val="center"/><w:tblCellMar><w:left w:w="200" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style></w:styles>"#,
                rdocx_oxml::namespace::W_NS
            )
            .as_bytes(),
        )
        .unwrap();
        let mut table = CT_Tbl::new();
        table.grid = Some(CT_TblGrid {
            columns: vec![CT_TblGridCol { width: Twips(4000) }],
            ..CT_TblGrid::default()
        });
        let mut row = CT_Row::new();
        row.cells.push(CT_Tc::new());
        table.rows.push(row);
        let laid_out = layout_with_styles(&table, 300.0, &styles);
        assert!((laid_out.table_width - 100.0).abs() < 0.01);
        assert!((laid_out.table_indent - 100.0).abs() < 0.01);
        assert!((laid_out.rows[0].cells[0].margin_left - 10.0).abs() < 0.01);
        assert!((laid_out.rows[0].cells[0].margin_right - 5.0).abs() < 0.01);

        table.properties = Some(CT_TblPr {
            width: Some(CT_TblWidth::dxa(3000)),
            jc: Some(ST_Jc::Left),
            indent: Some(CT_TblWidth::dxa(400)),
            cell_margin: Some(CT_TblCellMar {
                left: Some(Twips(40)),
                ..CT_TblCellMar::default()
            }),
            ..CT_TblPr::default()
        });
        let overlaid = layout_with_styles(&table, 300.0, &styles);
        assert!((overlaid.table_width - 200.0).abs() < 0.01);
        assert!((overlaid.table_indent - 20.0).abs() < 0.01);
        assert!((overlaid.rows[0].cells[0].margin_left - 2.0).abs() < 0.01);
        assert!((overlaid.rows[0].cells[0].margin_right - 5.0).abs() < 0.01);

        let styles = CT_Styles::from_xml(
            format!(
                r#"<w:styles xmlns:w="{}"><w:style w:type="table" w:styleId="Base"><w:tblPr><w:tblLook w:firstRow="1" w:noHBand="1"/></w:tblPr></w:style><w:style w:type="table" w:styleId="Derived" w:default="1"><w:basedOn w:val="Base"/><w:tblPr><w:tblLook w:lastRow="1"/></w:tblPr></w:style></w:styles>"#,
                rdocx_oxml::namespace::W_NS
            )
            .as_bytes(),
        )
        .unwrap();
        let inherited = resolve_base_table_properties(&CT_Tbl::new(), &styles);
        let look = inherited.look.unwrap();
        assert_eq!(look.first_row, Some(true));
        assert_eq!(look.last_row, Some(true));
        assert_eq!(look.no_h_band, Some(true));
    }
}
