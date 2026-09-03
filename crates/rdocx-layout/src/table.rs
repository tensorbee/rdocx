//! Table layout: column widths, cell content, merge handling.

use rdocx_oxml::styles::CT_Styles;
use rdocx_oxml::table::{CT_Tbl, CT_TblBorders, CT_TblGrid, ST_VerticalJc, VMerge};

use crate::WordStory;
use crate::block::{
    CellBlockSemantics, CellSemantics, ParagraphBlock, ParagraphSemantics, RowSemantics,
    TableSemantics,
};
use crate::engine::SourceRegistry;
use crate::input::{LayoutInput, MediaRegistry};
use crate::style_resolver::NumberingState;
use oxml_layout::{Color, Diagnostic, FontManager, Result, StructureId};

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
    /// Whether this cell is in the first row.
    pub is_first_row: bool,
    /// Whether this cell is in the last row.
    pub is_last_row: bool,
    /// Vertical alignment of content within the cell.
    pub v_align: Option<ST_VerticalJc>,
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
    )
}

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
) -> Result<(TableBlock, TableSemantics)> {
    // 1. Compute column widths
    let col_widths = compute_column_widths(tbl.grid.as_ref(), available_width, tbl);
    let table_width: f64 = col_widths.iter().sum();

    // Table indent
    let table_indent = tbl
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

    let num_rows = tbl.rows.len();
    let mut header_row_indices = Vec::new();
    let mut rows = Vec::new();
    let mut row_semantics = Vec::new();
    let mut exact_rows = Vec::new();

    for (row_idx, row) in tbl.rows.iter().enumerate() {
        let is_header = row
            .properties
            .as_ref()
            .and_then(|p| p.header)
            .unwrap_or(false);
        if is_header {
            header_row_indices.push(row_idx);
        }

        let mut cells = Vec::new();
        let mut cell_semantics = Vec::new();
        let mut col_index = 0usize;

        for (cell_index, cell) in row.cells.iter().enumerate() {
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
                row.properties
                    .as_ref()
                    .and_then(|properties| properties.cnf_style.as_deref()),
                cell.properties
                    .as_ref()
                    .and_then(|properties| properties.cnf_style.as_deref()),
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

            // Calculate cell width from spanned columns
            let cell_width: f64 = (col_index..col_index + grid_span as usize)
                .filter_map(|i| col_widths.get(i))
                .sum();

            let content_width = (cell_width - cell_margin_left - cell_margin_right).max(0.0);

            // Layout cell content (paragraphs and nested tables)
            let (blocks, block_semantics) = if is_vmerge_continue {
                (Vec::new(), Vec::new())
            } else {
                layout_cell_content(
                    &cell.content,
                    content_width,
                    styles,
                    input,
                    media,
                    fm,
                    num_state,
                    diagnostics,
                    sources,
                    story,
                    path,
                    row_idx,
                    cell_index,
                    style_cell.paragraph_properties.as_ref(),
                )?
            };

            let content_height: f64 = blocks.iter().map(CellBlock::total_height).sum::<f64>()
                + cell_margin_top
                + cell_margin_bottom;

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
                is_first_row: row_idx == 0,
                is_last_row: row_idx == num_rows - 1,
                v_align,
            });
            cell_semantics.push(CellSemantics {
                blocks: block_semantics,
            });

            col_index += grid_span as usize;
        }

        let max_cell_height = cells
            .iter()
            .filter(|cell| !cell.starts_vmerge)
            .map(|cell| cell.height)
            .fold(0.0f64, f64::max);
        let specified_height = row
            .properties
            .as_ref()
            .and_then(|p| p.height)
            .map(|h| h.to_pt())
            .unwrap_or(0.0);
        let exact = row
            .properties
            .as_ref()
            .and_then(|properties| properties.height_rule.as_deref())
            == Some("exact")
            && specified_height > 0.0;
        exact_rows.push(exact);
        for cell in &mut cells {
            cell.clip_content = exact && !cell.is_vmerge_continue;
        }
        let row_height = if exact {
            specified_height
        } else {
            max_cell_height.max(specified_height)
        };

        rows.push(TableRow {
            structure_id: None,
            cells,
            height: row_height,
            is_header,
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

    let row_heights = rows.iter().map(|row| row.height).collect::<Vec<_>>();
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
            cell.height = row_heights[row_index];
            cell.merged_height = row_heights[row_index];
            cell.merge_with_below = continuing_spans.contains(&(cell.col_index, cell.grid_span));
        }
    }
    for (row_index, cell_index, last_row, required) in spans {
        let restart = &mut rows[row_index].cells[cell_index];
        restart.merged_height = row_heights[row_index..=last_row].iter().sum();
        restart.is_last_row = last_row + 1 == num_rows;
        restart.clip_content = required > restart.merged_height;
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
        },
        TableSemantics {
            rows: row_semantics,
        },
    ))
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
    tbl: &CT_Tbl,
) -> Vec<f64> {
    match grid {
        Some(g) if !g.columns.is_empty() => {
            let widths: Vec<f64> = g.columns.iter().map(|c| c.width.to_pt()).collect();
            let total: f64 = widths.iter().sum();
            if total < 0.01 {
                // All zero widths — distribute equally based on column count
                let n = g.columns.len();
                vec![available_width / n as f64; n]
            } else if total > available_width + 1.0 {
                // Overflows the text column: scale down so it fits the page.
                let scale = available_width / total;
                widths.iter().map(|w| w * scale).collect()
            } else {
                widths
            }
        }
        _ => {
            // No grid defined — infer column count from the first row
            let num_cols = tbl
                .rows
                .first()
                .map(|r| {
                    r.cells
                        .iter()
                        .map(|c| {
                            c.properties.as_ref().and_then(|p| p.grid_span).unwrap_or(1) as usize
                        })
                        .sum::<usize>()
                })
                .unwrap_or(1)
                .max(1);
            vec![available_width / num_cols as f64; num_cols]
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
    table_path: &[usize],
    row_index: usize,
    cell_index: usize,
    table_style_ppr: Option<&rdocx_oxml::properties::CT_PPr>,
) -> Result<(Vec<CellBlock>, Vec<CellBlockSemantics>)> {
    use crate::engine;
    use rdocx_oxml::table::CellContent;

    let mut blocks = Vec::new();
    let mut semantics = Vec::new();
    for (content_index, item) in content.iter().enumerate() {
        let mut source_path = table_path.to_vec();
        source_path.extend([row_index, cell_index, content_index]);
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
                )?;
                blocks.push(CellBlock::Table(nested));
                semantics.push(CellBlockSemantics::Table(nested_semantics));
            }
            CellContent::ContentControl(_) => {}
        }
    }
    Ok((blocks, semantics))
}

#[derive(Default)]
struct ResolvedTableCellStyle {
    paragraph_properties: Option<rdocx_oxml::properties::CT_PPr>,
    borders: Option<CT_TblBorders>,
    shading: Option<rdocx_oxml::properties::CT_Shd>,
}

fn resolve_table_style_cell(
    table: &CT_Tbl,
    styles: &CT_Styles,
    row: usize,
    column: usize,
    row_count: usize,
    column_count: usize,
    row_cnf_style: Option<&str>,
    cell_cnf_style: Option<&str>,
) -> ResolvedTableCellStyle {
    let Some(mut style_id) = table
        .properties
        .as_ref()
        .and_then(|p| p.style_id.as_deref())
    else {
        return ResolvedTableCellStyle::default();
    };
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
    for style in chain.into_iter().rev() {
        if let Some(properties) = &style.ppr {
            resolved
                .paragraph_properties
                .get_or_insert_with(rdocx_oxml::properties::CT_PPr::default)
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
        for region in applicable_table_regions(
            table,
            row,
            column,
            row_count,
            column_count,
            row_cnf_style,
            cell_cnf_style,
        ) {
            for conditional in style
                .conditional_table_styles
                .iter()
                .filter(|conditional| conditional.region == region)
            {
                if let Some(properties) = &conditional.paragraph_properties {
                    resolved
                        .paragraph_properties
                        .get_or_insert_with(rdocx_oxml::properties::CT_PPr::default)
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
        }
    }
    resolved
}

fn applicable_table_regions(
    table: &CT_Tbl,
    row: usize,
    column: usize,
    row_count: usize,
    column_count: usize,
    row_cnf_style: Option<&str>,
    cell_cnf_style: Option<&str>,
) -> Vec<&'static str> {
    let cnf = |index: usize| {
        [row_cnf_style, cell_cnf_style]
            .into_iter()
            .flatten()
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
    let first_row =
        (enabled(look.and_then(|look| look.first_row), 0x20, false) && row == 0) || cnf(0);
    let last_row = (enabled(look.and_then(|look| look.last_row), 0x40, false)
        && row + 1 == row_count)
        || cnf(1);
    let first_column =
        (enabled(look.and_then(|look| look.first_column), 0x80, false) && column == 0) || cnf(2);
    let last_column = (enabled(look.and_then(|look| look.last_column), 0x100, false)
        && column + 1 == column_count)
        || cnf(3);
    let no_h_band = enabled(look.and_then(|look| look.no_h_band), 0x200, false);
    let no_v_band = enabled(look.and_then(|look| look.no_v_band), 0x400, false);

    let mut regions = vec!["wholeTable"];
    if cnf(6) {
        regions.push("band1Horz");
    } else if cnf(7) {
        regions.push("band2Horz");
    } else if !no_h_band {
        regions.push(if row.is_multiple_of(2) {
            "band1Horz"
        } else {
            "band2Horz"
        });
    }
    if cnf(4) {
        regions.push("band1Vert");
    } else if cnf(5) {
        regions.push("band2Vert");
    } else if !no_v_band {
        regions.push(if column.is_multiple_of(2) {
            "band1Vert"
        } else {
            "band2Vert"
        });
    }
    if first_column {
        regions.push("firstCol");
    }
    if last_column {
        regions.push("lastCol");
    }
    if first_row {
        regions.push("firstRow");
    }
    if last_row {
        regions.push("lastRow");
    }
    if cnf(9) {
        regions.push("nwCell");
    } else if cnf(8) {
        regions.push("neCell");
    } else if cnf(11) {
        regions.push("swCell");
    } else if cnf(10) {
        regions.push("seCell");
    } else {
        match (first_row, last_row, first_column, last_column) {
            (true, _, true, _) => regions.push("nwCell"),
            (true, _, _, true) => regions.push("neCell"),
            (_, true, true, _) => regions.push("swCell"),
            (_, true, _, true) => regions.push("seCell"),
            _ => {}
        }
    }
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
        CT_Row, CT_TblGrid, CT_TblGridCol, CT_TblLook, CT_TblPr, CT_Tc, CT_TcPr, CT_TrPr,
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
        let widths = compute_column_widths(Some(&grid), 468.0, &tbl);

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
            compute_column_widths(Some(&grid), 468.0, &table),
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
        let widths = compute_column_widths(Some(&grid), 468.0, &tbl);

        let total: f64 = widths.iter().sum();
        assert!((total - 468.0).abs() < 1.0, "got {total}");
        // Proportions are preserved.
        assert!((widths[0] - widths[1]).abs() < 0.01);
    }

    #[test]
    fn column_widths_no_grid() {
        let tbl = CT_Tbl::new();
        let widths = compute_column_widths(None, 468.0, &tbl);
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
        let widths = compute_column_widths(Some(&grid), 468.0, &tbl);
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
        let widths = compute_column_widths(None, 300.0, &tbl);
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

        // Table width should match available width
        assert!((block.table_width - 234.0).abs() < 1.0);
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
}
