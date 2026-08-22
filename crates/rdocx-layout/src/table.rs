//! Table layout: column widths, cell content, merge handling.

use rdocx_oxml::styles::CT_Styles;
use rdocx_oxml::table::{CT_Tbl, CT_TblBorders, CT_TblGrid, ST_VerticalJc, VMerge};

use crate::WordStory;
use crate::block::ParagraphBlock;
use crate::engine::SourceRegistry;
use crate::input::{LayoutInput, MediaRegistry};
use crate::style_resolver::NumberingState;
use oxml_layout::{Color, Diagnostic, FontManager, Result};

/// A laid-out table.
#[derive(Debug, Clone)]
pub struct TableBlock {
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
    /// Cells in this row.
    pub cells: Vec<TableCell>,
    /// Row height in points.
    pub height: f64,
    /// Whether this row is a header row.
    pub is_header: bool,
}

/// A laid-out table cell.
#[derive(Debug, Clone)]
pub struct TableCell {
    /// Cell content (paragraph blocks).
    pub paragraphs: Vec<ParagraphBlock>,
    /// Nested tables laid out inside this cell, as
    /// `(paragraph position, block)`: the block renders after that many of
    /// `paragraphs` (document order is preserved for mixed content).
    pub nested: Vec<(usize, TableBlock)>,
    /// Cell width in points (may span multiple grid columns).
    pub width: f64,
    /// Cell height in points (set to row height).
    pub height: f64,
    /// Number of grid columns this cell spans.
    pub grid_span: u32,
    /// Whether this cell is part of a vertical merge continuation (render no content).
    pub is_vmerge_continue: bool,
    /// Whether this cell starts a vertical merge (vMerge=restart).
    pub starts_vmerge: bool,
    /// Total height this cell spans: the sum of the spanned row heights for
    /// a merge-starting cell, otherwise equal to `height`. The renderer
    /// draws shading, side borders, and vertically-aligned content over
    /// this span.
    pub merged_height: f64,
    /// The same grid column continues a vertical merge in the next row —
    /// the renderer suppresses this cell's bottom border.
    pub merge_with_below: bool,
    /// Column index in the grid.
    pub col_index: usize,
    /// Cell-level borders.
    pub borders: Option<CT_TblBorders>,
    /// Cell background shading color.
    pub shading: Option<Color>,
    /// Cell margin left in points.
    pub margin_left: f64,
    /// Cell margin top in points.
    pub margin_top: f64,
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
) -> Result<TableBlock> {
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

    // Table-level borders: direct properties first, then the referenced
    // table style (following basedOn up a bounded chain) — receipts and
    // forms commonly carry all their grid lines in the style.
    let table_borders = tbl
        .properties
        .as_ref()
        .and_then(|p| p.borders.clone())
        .or_else(|| {
            let mut id = tbl.properties.as_ref()?.style_id.as_deref()?;
            for _ in 0..8 {
                let style = styles.get_by_id(id)?;
                if let Some(borders) = &style.table_borders {
                    return Some(borders.clone());
                }
                id = style.based_on.as_deref()?;
            }
            None
        });

    // Table-style paragraph properties, merged base-first along the
    // basedOn chain: they cascade onto every paragraph inside the table
    // (styles like Table Grid carry `spacing after=0` here).
    let table_style_ppr: Option<rdocx_oxml::properties::CT_PPr> = {
        let mut merged: Option<rdocx_oxml::properties::CT_PPr> = None;
        if let Some(start) = tbl.properties.as_ref().and_then(|p| p.style_id.as_deref()) {
            let mut chain = Vec::new();
            let mut id = start;
            for _ in 0..8 {
                let Some(style) = styles.get_by_id(id) else { break };
                chain.push(style);
                match style.based_on.as_deref() {
                    Some(next) => id = next,
                    None => break,
                }
            }
            for style in chain.iter().rev() {
                if let Some(ppr) = &style.ppr {
                    merged
                        .get_or_insert_with(rdocx_oxml::properties::CT_PPr::default)
                        .merge_from(ppr);
                }
            }
        }
        merged
    };

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
    let mut exact_rows: Vec<bool> = Vec::new();

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
            let starts_vmerge = cell
                .properties
                .as_ref()
                .and_then(|p| p.v_merge)
                .map(|vm| vm == VMerge::Restart)
                .unwrap_or(false);

            // Cell-level borders and shading
            let cell_borders = cell.properties.as_ref().and_then(|p| p.borders.clone());
            let cell_shading = cell
                .properties
                .as_ref()
                .and_then(|p| p.shading.as_ref())
                .and_then(|shd| shd.fill.as_ref())
                .filter(|f| f.as_str() != "auto")
                .map(|f| Color::from_hex(f));

            // Calculate cell width from spanned columns
            let cell_width: f64 = (col_index..col_index + grid_span as usize)
                .filter_map(|i| col_widths.get(i))
                .sum();

            let content_width = (cell_width - cell_margin_left - cell_margin_right).max(0.0);

            // Layout cell content (paragraphs and nested tables)
            let (paragraphs, nested) = if is_vmerge_continue {
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
                    table_style_ppr.as_ref(),
                )?
            };

            let content_height: f64 = paragraphs.iter().map(|p| p.total_height()).sum::<f64>()
                + nested
                    .iter()
                    .map(|(_, t)| t.total_height())
                    .sum::<f64>()
                + cell_margin_top
                + cell_margin_bottom;

            let v_align = cell.properties.as_ref().and_then(|p| p.v_align);

            cells.push(TableCell {
                paragraphs,
                nested,
                width: cell_width,
                height: content_height,
                grid_span,
                is_vmerge_continue,
                starts_vmerge,
                merged_height: content_height,
                merge_with_below: false,
                col_index,
                borders: cell_borders,
                shading: cell_shading,
                margin_left: cell_margin_left,
                margin_top: cell_margin_top,
                is_first_row: row_idx == 0,
                is_last_row: row_idx == num_rows - 1,
                v_align,
            });

            col_index += grid_span as usize;
        }

        // Row height from cells that do NOT start a vertical merge — a
        // merge-starting cell's content spans several rows, so counting it
        // here would balloon this row; its needs are distributed below.
        let max_cell_height = cells
            .iter()
            .filter(|c| !c.starts_vmerge)
            .map(|c| c.height)
            .fold(0.0f64, f64::max);
        let fallback_height = cells.iter().map(|c| c.height).fold(0.0f64, f64::max);
        let specified_height = row
            .properties
            .as_ref()
            .and_then(|p| p.height)
            .map(|h| h.to_pt())
            .unwrap_or(0.0);
        // hRule="exact" pins the row height regardless of content (Word
        // clips overflow) — dense forms rely on it row by row.
        let exact = row
            .properties
            .as_ref()
            .and_then(|p| p.height_rule.as_deref())
            == Some("exact");
        exact_rows.push(exact && specified_height > 0.0);
        let row_height = if exact && specified_height > 0.0 {
            specified_height
        } else if max_cell_height > 0.0 || specified_height > 0.0 {
            max_cell_height.max(specified_height)
        } else {
            // Every cell starts a merge (or the row is empty): fall back so
            // the row is not zero-height.
            fallback_height
        };

        // Cell heights are frozen after the vertical-merge pass below.
        rows.push(TableRow {
            cells,
            height: row_height,
            is_header,
        });
    }

    // Vertical merges: find each merge-starting cell's span (continue cells
    // at the same grid column in the following rows), grow the last spanned
    // row when the merged content needs more room than the span offers,
    // then freeze cell heights and the merge geometry for the renderer.
    let row_count = rows.len();
    let mut spans: Vec<(usize, usize, usize)> = Vec::new(); // (row, cell, last_row)
    for r in 0..row_count {
        for ci in 0..rows[r].cells.len() {
            if !rows[r].cells[ci].starts_vmerge {
                continue;
            }
            let col = rows[r].cells[ci].col_index;
            let mut last = r;
            while last + 1 < row_count
                && rows[last + 1]
                    .cells
                    .iter()
                    .any(|c| c.col_index == col && c.is_vmerge_continue)
            {
                last += 1;
            }
            let needed = rows[r].cells[ci].height; // content height, unforced
            let have: f64 = (r..=last).map(|i| rows[i].height).sum();
            // Exact-height rows never grow (Word clips merged content too).
            if needed > have && !exact_rows.get(last).copied().unwrap_or(false) {
                rows[last].height += needed - have;
            }
            spans.push((r, ci, last));
        }
    }
    let heights: Vec<f64> = rows.iter().map(|x| x.height).collect();
    for r in 0..row_count {
        let next_continue_cols: Vec<usize> = if r + 1 < row_count {
            rows[r + 1]
                .cells
                .iter()
                .filter(|c| c.is_vmerge_continue)
                .map(|c| c.col_index)
                .collect()
        } else {
            Vec::new()
        };
        for cell in &mut rows[r].cells {
            cell.height = heights[r];
            cell.merged_height = heights[r];
            cell.merge_with_below =
                cell.merge_with_below || next_continue_cols.contains(&cell.col_index);
        }
    }
    for &(r, ci, last) in &spans {
        rows[r].cells[ci].merged_height = heights[r..=last].iter().sum();
    }

    Ok(TableBlock {
        col_widths,
        rows,
        header_row_indices,
        table_width,
        table_indent,
        borders: table_borders,
    })
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
/// For nested tables, we lay out the table and flatten its cell paragraphs
/// into the parent cell's paragraph blocks.
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
    table_ppr: Option<&rdocx_oxml::properties::CT_PPr>,
) -> Result<(Vec<ParagraphBlock>, Vec<(usize, TableBlock)>)> {
    use crate::engine;
    use rdocx_oxml::table::CellContent;

    let mut blocks = Vec::new();
    let mut nested = Vec::new();
    for (content_index, item) in content.iter().enumerate() {
        let mut source_path = table_path.to_vec();
        source_path.extend([row_index, cell_index, content_index]);
        match item {
            CellContent::Paragraph(para) => {
                let source = sources.and_then(|sources| sources.id(story, &source_path));
                let block = engine::layout_paragraph_with_source_in_table(
                    para,
                    available_width,
                    styles,
                    input,
                    media,
                    fm,
                    num_state,
                    diagnostics,
                    source,
                    table_ppr,
                )?;
                blocks.push(block);
            }
            CellContent::Table(tbl) => {
                // Recursively lay out the nested table and keep it as a
                // block anchored at the current paragraph position, so the
                // paginator renders real rows, borders, and shading instead
                // of the old flattened paragraph stream.
                let block = layout_table_inner(
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
                nested.push((blocks.len(), block));
            }
            CellContent::ContentControl(_) => {}
        }
    }
    Ok((blocks, nested))
}

#[cfg(test)]
mod tests {
    use super::*;
    use rdocx_oxml::table::{CT_TblGrid, CT_TblGridCol};
    use rdocx_oxml::units::Twips;

    #[test]
    fn narrow_grid_keeps_its_declared_width() {
        let tbl = CT_Tbl::new();
        let grid = CT_TblGrid {
            columns: vec![
                CT_TblGridCol { width: Twips(2880) }, // 2 inches = 144pt
                CT_TblGridCol { width: Twips(2880) },
            ],
        };

        // 288pt total in a 468pt text column: the author asked for a narrow
        // table, so it must not be stretched to the margins.
        let widths = compute_column_widths(Some(&grid), 468.0, &tbl);

        assert_eq!(widths.len(), 2);
        let total: f64 = widths.iter().sum();
        assert!((total - 288.0).abs() < 1.0, "got {total}");
    }

    #[test]
    fn overflowing_grid_is_scaled_down_to_fit() {
        let tbl = CT_Tbl::new();
        let grid = CT_TblGrid {
            columns: vec![
                CT_TblGridCol { width: Twips(7200) }, // 5 inches = 360pt
                CT_TblGridCol { width: Twips(7200) },
            ],
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
    fn nested_table_layout_dimensions() {
        use rdocx_oxml::table::{CT_Row, CT_Tbl, CT_Tc, CellContent};

        // Build an outer table with one cell containing a nested table
        let mut outer = CT_Tbl::new();
        outer.grid = Some(CT_TblGrid {
            columns: vec![CT_TblGridCol { width: Twips(4680) }], // 3.25"
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
            document: rdocx_oxml::document::CT_Document {
                body: rdocx_oxml::document::CT_Body {
                    content: Vec::new(),
                    sect_pr: None,
                },
                extra_namespaces: Vec::new(),
                background_xml: None,
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

        // The outer paragraph stays a paragraph block; the nested table is
        // kept as a real block anchored after it (no more flattening).
        let cell = &block.rows[0].cells[0];
        assert_eq!(cell.paragraphs.len(), 1, "outer paragraph only");
        assert_eq!(cell.nested.len(), 1, "one nested table block");
        let (pos, nested) = &cell.nested[0];
        assert_eq!(*pos, 1, "nested table renders after the paragraph");
        assert_eq!(nested.rows.len(), 1);
        assert_eq!(nested.rows[0].cells.len(), 2);
        assert!(
            cell.height >= cell.paragraphs[0].total_height() + nested.total_height(),
            "cell height covers paragraph + nested table"
        );

        // Table width should match available width
        assert!((block.table_width - 234.0).abs() < 1.0);
    }
}
