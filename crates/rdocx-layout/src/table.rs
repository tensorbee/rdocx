//! Table layout: column widths, cell content, merge handling.

use rdocx_oxml::styles::CT_Styles;
use rdocx_oxml::table::{CT_Tbl, CT_TblBorders, CT_TblGrid, ST_VerticalJc, VMerge};

use crate::block::ParagraphBlock;
use crate::error::Result;
use crate::font::FontManager;
use crate::input::LayoutInput;
use crate::style_resolver::NumberingState;

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
    /// Cell width in points (may span multiple grid columns).
    pub width: f64,
    /// Cell height in points (set to row height).
    pub height: f64,
    /// Number of grid columns this cell spans.
    pub grid_span: u32,
    /// Whether this cell is part of a vertical merge continuation (render no content).
    pub is_vmerge_continue: bool,
    /// Column index in the grid.
    pub col_index: usize,
    /// Cell-level borders.
    pub borders: Option<CT_TblBorders>,
    /// Cell background shading color.
    pub shading: Option<crate::output::Color>,
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
    fm: &mut FontManager,
    num_state: &mut NumberingState,
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

    // Table-level borders
    let table_borders = tbl.properties.as_ref().and_then(|p| p.borders.clone());

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

        for cell in &row.cells {
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

            // Cell-level borders and shading
            let cell_borders = cell.properties.as_ref().and_then(|p| p.borders.clone());
            let cell_shading = cell
                .properties
                .as_ref()
                .and_then(|p| p.shading.as_ref())
                .and_then(|shd| shd.fill.as_ref())
                .filter(|f| f.as_str() != "auto")
                .map(|f| crate::output::Color::from_hex(f));

            // Calculate cell width from spanned columns
            let cell_width: f64 = (col_index..col_index + grid_span as usize)
                .filter_map(|i| col_widths.get(i))
                .sum();

            let content_width = (cell_width - cell_margin_left - cell_margin_right).max(0.0);

            // Layout cell content (paragraphs and nested tables)
            let paragraphs = if is_vmerge_continue {
                Vec::new()
            } else {
                layout_cell_content(&cell.content, content_width, styles, input, fm, num_state)?
            };

            let content_height: f64 = paragraphs.iter().map(|p| p.total_height()).sum::<f64>()
                + cell_margin_top
                + cell_margin_bottom;

            let v_align = cell.properties.as_ref().and_then(|p| p.v_align);

            cells.push(TableCell {
                paragraphs,
                width: cell_width,
                height: content_height,
                grid_span,
                is_vmerge_continue,
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

        // Row height is max of all cell heights and any specified height
        let max_cell_height = cells.iter().map(|c| c.height).fold(0.0f64, f64::max);
        let specified_height = row
            .properties
            .as_ref()
            .and_then(|p| p.height)
            .map(|h| h.to_pt())
            .unwrap_or(0.0);
        let row_height = max_cell_height.max(specified_height);

        // Set all cell heights to match row height
        for cell in &mut cells {
            cell.height = row_height;
        }

        rows.push(TableRow {
            cells,
            height: row_height,
            is_header,
        });
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

/// Compute column widths from CT_TblGrid, scaling to available width if needed.
fn compute_column_widths(
    grid: Option<&CT_TblGrid>,
    available_width: f64,
    tbl: &CT_Tbl,
) -> Vec<f64> {
    match grid {
        Some(g) if !g.columns.is_empty() => {
            let widths: Vec<f64> = g.columns.iter().map(|c| c.width.to_pt()).collect();
            let total: f64 = widths.iter().sum();
            if total > 0.01 && (total - available_width).abs() > 1.0 {
                // Scale to fit available width
                let scale = available_width / total;
                widths.iter().map(|w| w * scale).collect()
            } else if total < 0.01 {
                // All zero widths — distribute equally based on column count
                let n = g.columns.len();
                vec![available_width / n as f64; n]
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
    fm: &mut FontManager,
    num_state: &mut NumberingState,
) -> Result<Vec<ParagraphBlock>> {
    use crate::engine;
    use rdocx_oxml::table::CellContent;

    let mut blocks = Vec::new();
    for item in content {
        match item {
            CellContent::Paragraph(para) => {
                let block =
                    engine::layout_paragraph(para, available_width, styles, input, fm, num_state)?;
                blocks.push(block);
            }
            CellContent::Table(tbl) => {
                // Recursively lay out the nested table
                let _nested = layout_table(tbl, available_width, styles, input, fm, num_state)?;
                // For now, flatten: render nested table cell content as paragraph blocks
                // (Full nested table rendering would require the paginator to handle tables within cells)
                for row in &_nested.rows {
                    for cell in &row.cells {
                        if !cell.is_vmerge_continue {
                            blocks.extend(cell.paragraphs.iter().cloned());
                        }
                    }
                }
            }
        }
    }
    Ok(blocks)
}

#[cfg(test)]
mod tests {
    use super::*;
    use rdocx_oxml::table::{CT_TblGrid, CT_TblGridCol};
    use rdocx_oxml::units::Twips;

    #[test]
    fn column_widths_from_grid() {
        let tbl = CT_Tbl::new();
        let grid = CT_TblGrid {
            columns: vec![
                CT_TblGridCol { width: Twips(2880) }, // 2 inches
                CT_TblGridCol { width: Twips(2880) },
            ],
        };
        let widths = compute_column_widths(Some(&grid), 468.0, &tbl);
        assert_eq!(widths.len(), 2);
        // 2880tw = 144pt, total = 288pt, scaled to 468pt
        let total: f64 = widths.iter().sum();
        assert!((total - 468.0).abs() < 1.0);
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
            hyperlink_urls: std::collections::HashMap::new(),
            footnotes: None,
            endnotes: None,
            core_properties: None,
            theme: None,
            fonts: Vec::new(),
        };

        let mut fm = crate::font::FontManager::new();
        let mut num_state = crate::style_resolver::NumberingState::new();

        let result = layout_table(&outer, 234.0, &styles, &input, &mut fm, &mut num_state);
        assert!(result.is_ok());
        let block = result.unwrap();

        // Outer table should have 1 row, 1 cell
        assert_eq!(block.rows.len(), 1);
        assert_eq!(block.rows[0].cells.len(), 1);

        // Cell should have paragraphs from both the outer paragraph and flattened nested content
        let cell = &block.rows[0].cells[0];
        // At least: "Before nested" + "N1" + "N2" = 3 paragraph blocks
        assert!(
            cell.paragraphs.len() >= 3,
            "Expected at least 3 paragraph blocks from outer + nested content, got {}",
            cell.paragraphs.len()
        );

        // Table width should match available width
        assert!((block.table_width - 234.0).abs() < 1.0);
    }
}
