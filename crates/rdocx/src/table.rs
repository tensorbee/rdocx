//! Table — a block-level container for rows and cells of content.

use rdocx_oxml::borders::CT_BorderEdge;
use rdocx_oxml::properties::CT_Shd;
use rdocx_oxml::shared::ST_Jc;
use rdocx_oxml::table::{
    CT_Row, CT_Tbl, CT_TblBorders, CT_TblCellMar, CT_TblPr, CT_TblWidth, CT_Tc, CT_TcPr, CT_TrPr,
    ST_VerticalJc, VMerge,
};
use rdocx_oxml::text::CT_P;

use crate::Length;
use crate::paragraph::{Paragraph, ParagraphRef};

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

    fn from_st(st: ST_VerticalJc) -> Self {
        match st {
            ST_VerticalJc::Top => Self::Top,
            ST_VerticalJc::Center => Self::Center,
            ST_VerticalJc::Bottom => Self::Bottom,
        }
    }
}

// ---- Mutable Table ----

/// A mutable reference to a table in a document.
pub struct Table<'a> {
    pub(crate) inner: &'a mut CT_Tbl,
}

impl<'a> Table<'a> {
    /// Set the table style by ID.
    pub fn style(mut self, style_id: &str) -> Self {
        self.ensure_tbl_pr().style_id = Some(style_id.to_string());
        self
    }

    /// Set the table width in twips (dxa).
    pub fn width(mut self, length: Length) -> Self {
        self.ensure_tbl_pr().width = Some(CT_TblWidth::dxa(length.as_twips().0));
        self
    }

    /// Set the table width as a percentage (0–100).
    pub fn width_pct(mut self, percent: f64) -> Self {
        // OOXML uses 50ths of a percent
        self.ensure_tbl_pr().width = Some(CT_TblWidth::pct((percent * 50.0) as i32));
        self
    }

    /// Set table alignment.
    pub fn alignment(mut self, jc: crate::paragraph::Alignment) -> Self {
        use crate::paragraph::Alignment;
        let st_jc = match jc {
            Alignment::Left => ST_Jc::Left,
            Alignment::Center => ST_Jc::Center,
            Alignment::Right => ST_Jc::Right,
            Alignment::Justify => ST_Jc::Both,
        };
        self.ensure_tbl_pr().jc = Some(st_jc);
        self
    }

    /// Set borders on all edges and internal gridlines.
    pub fn borders(mut self, style: crate::BorderStyle, size_eighths_pt: u32, color: &str) -> Self {
        let edge = CT_BorderEdge {
            val: style.to_st_border(),
            sz: Some(size_eighths_pt),
            space: Some(0),
            color: Some(color.to_string()),
        };
        self.ensure_tbl_pr().borders = Some(CT_TblBorders {
            top: Some(edge.clone()),
            bottom: Some(edge.clone()),
            left: Some(edge.clone()),
            right: Some(edge.clone()),
            inside_h: Some(edge.clone()),
            inside_v: Some(edge),
        });
        self
    }

    /// Set default cell margins.
    pub fn cell_margins(
        mut self,
        top: Length,
        right: Length,
        bottom: Length,
        left: Length,
    ) -> Self {
        self.ensure_tbl_pr().cell_margin = Some(CT_TblCellMar {
            top: Some(top.as_twips()),
            right: Some(right.as_twips()),
            bottom: Some(bottom.as_twips()),
            left: Some(left.as_twips()),
        });
        self
    }

    /// Set the table layout to fixed or auto.
    pub fn layout_fixed(mut self) -> Self {
        self.ensure_tbl_pr().layout = Some("fixed".to_string());
        self
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
        let pr = self.ensure_tr_pr();
        pr.height = Some(length.as_twips());
        pr.height_rule = Some("atLeast".to_string());
        self
    }

    /// Set exact row height.
    pub fn height_exact(mut self, length: Length) -> Self {
        let pr = self.ensure_tr_pr();
        pr.height = Some(length.as_twips());
        pr.height_rule = Some("exact".to_string());
        self
    }

    /// Mark this row as a header row (repeats on each page).
    pub fn header(mut self) -> Self {
        self.ensure_tr_pr().header = Some(true);
        self
    }

    /// Prevent this row from splitting across pages.
    pub fn cant_split(mut self) -> Self {
        self.ensure_tr_pr().cant_split = Some(true);
        self
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

    /// Get an iterator over immutable paragraph references.
    pub fn paragraphs(&self) -> impl Iterator<Item = ParagraphRef<'_>> {
        self.inner
            .paragraphs()
            .into_iter()
            .map(|p| ParagraphRef { inner: p })
    }

    /// Set cell width.
    pub fn width(mut self, length: Length) -> Self {
        self.ensure_tc_pr().width = Some(CT_TblWidth::dxa(length.as_twips().0));
        self
    }

    /// Set cell background shading color.
    pub fn shading(mut self, fill_color: &str) -> Self {
        self.ensure_tc_pr().shading = Some(CT_Shd {
            val: "clear".to_string(),
            color: Some("auto".to_string()),
            fill: Some(fill_color.to_string()),
        });
        self
    }

    /// Set vertical alignment within the cell.
    pub fn vertical_alignment(mut self, align: VerticalAlignment) -> Self {
        self.ensure_tc_pr().v_align = Some(align.to_st());
        self
    }

    /// Set horizontal merge (gridSpan). This cell spans `span` columns.
    pub fn grid_span(mut self, span: u32) -> Self {
        self.ensure_tc_pr().grid_span = Some(span);
        self
    }

    /// Start a vertical merge group (this cell is the top of the merged range).
    pub fn v_merge_restart(mut self) -> Self {
        self.ensure_tc_pr().v_merge = Some(VMerge::Restart);
        self
    }

    /// Continue a vertical merge group (this cell merges with the one above).
    pub fn v_merge_continue(mut self) -> Self {
        self.ensure_tc_pr().v_merge = Some(VMerge::Continue);
        self
    }

    /// Set no-wrap for text in this cell.
    pub fn no_wrap(mut self) -> Self {
        self.ensure_tc_pr().no_wrap = Some(true);
        self
    }

    /// Add a nested table inside this cell.
    pub fn add_table(&mut self, rows: usize, cols: usize) -> Table<'_> {
        use rdocx_oxml::table::{
            CT_Row, CT_Tbl, CT_TblGrid, CT_TblGridCol, CT_TblPr, CT_TblWidth, CT_Tc, CellContent,
        };
        use rdocx_oxml::units::Twips;

        // Default nested table column width: use equal splits of 4500tw (~3.125")
        let col_width = Twips(4500 / cols as i32);

        let grid = CT_TblGrid {
            columns: (0..cols)
                .map(|_| CT_TblGridCol { width: col_width })
                .collect(),
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
}

/// An immutable reference to a table cell.
pub struct CellRef<'a> {
    pub(crate) inner: &'a CT_Tc,
}

impl<'a> CellRef<'a> {
    /// Get the combined text of all paragraphs.
    pub fn text(&self) -> String {
        self.inner.text()
    }

    /// Get paragraph references.
    pub fn paragraphs(&self) -> impl Iterator<Item = ParagraphRef<'_>> {
        self.inner
            .paragraphs()
            .into_iter()
            .map(|p| ParagraphRef { inner: p })
    }

    /// Get the grid span, if set.
    pub fn grid_span(&self) -> Option<u32> {
        self.inner.properties.as_ref().and_then(|pr| pr.grid_span)
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

    /// Get the vertical alignment, if set.
    pub fn vertical_alignment(&self) -> Option<VerticalAlignment> {
        self.inner
            .properties
            .as_ref()
            .and_then(|pr| pr.v_align)
            .map(VerticalAlignment::from_st)
    }
}
