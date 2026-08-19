//! Table — a block-level container for rows and cells of content.

use rdocx_oxml::borders::CT_BorderEdge;
use rdocx_oxml::properties::CT_Shd;
use rdocx_oxml::shared::ST_Jc;
use rdocx_oxml::table::{
    CT_Row, CT_Tbl, CT_TblBorders, CT_TblCellMar, CT_TblPr, CT_TblWidth, CT_Tc, CT_TcPr, CT_TrPr,
    CellContent, ST_VerticalJc, VMerge,
};
use rdocx_oxml::text::CT_P;

use crate::Length;
use crate::content_control::ContentControlRef;
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

    /// Set the table indentation from the left margin in place.
    pub fn set_indent(&mut self, length: Length) {
        self.ensure_tbl_pr().indent = Some(CT_TblWidth::dxa(length.as_twips().0));
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
        };
        self.ensure_tbl_pr().borders = Some(CT_TblBorders {
            top: Some(edge.clone()),
            bottom: Some(edge.clone()),
            left: Some(edge.clone()),
            right: Some(edge.clone()),
            inside_h: Some(edge.clone()),
            inside_v: Some(edge),
        });
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

    /// Set the table layout to fixed or auto.
    pub fn layout_fixed(mut self) -> Self {
        self.set_layout_fixed();
        self
    }

    /// Set the table layout to fixed in place.
    pub fn set_layout_fixed(&mut self) {
        self.ensure_tbl_pr().layout = Some("fixed".to_string());
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

    /// Mark this row as a header row (repeats on each page).
    pub fn header(mut self) -> Self {
        self.set_header();
        self
    }

    /// Mark this row as a header row in place.
    pub fn set_header(&mut self) {
        self.ensure_tr_pr().header = Some(true);
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
        });
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

/// One direct child of a table cell, in source order.
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

#[cfg(test)]
mod tests {
    use super::*;
    use rdocx_oxml::table::{CT_Row, CT_TblGrid, CT_TblGridCol};
    use rdocx_oxml::units::Twips;

    #[test]
    fn cell_items_preserve_paragraph_table_control_and_raw_order() {
        let xml = br#"<w:tc xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:foreign"><w:p><w:r><w:t>first</w:t></w:r></w:p><x:raw/><w:tbl><w:tr><w:tc><w:p><w:r><w:t>nested</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sdt><w:sdtContent><w:p><w:r><w:t>inside</w:t></w:r></w:p></w:sdtContent></w:sdt><w:p><w:r><w:t>last</w:t></w:r></w:p></w:tc>"#;
        let mut reader = quick_xml::Reader::from_reader(xml.as_slice());
        let mut buffer = Vec::new();
        let cell = match reader.read_event_into(&mut buffer).unwrap() {
            quick_xml::events::Event::Start(_) => CT_Tc::from_xml(&mut reader).unwrap(),
            event => panic!("expected cell start, got {event:?}"),
        };
        let cell = CellRef { inner: &cell };

        let items = cell
            .items()
            .map(|item| match item {
                CellItemRef::Paragraph(paragraph) => format!("paragraph:{}", paragraph.text()),
                CellItemRef::Table(table) => format!("table:{}", table.row_count()),
                CellItemRef::ContentControl(control) => format!("control:{}", control.text()),
                CellItemRef::UnsupportedXml(raw) => {
                    format!("raw:{}", std::str::from_utf8(raw).unwrap())
                }
            })
            .collect::<Vec<_>>();

        assert_eq!(
            items,
            [
                "paragraph:first",
                "raw:<x:raw/>",
                "table:1",
                "control:inside",
                "paragraph:last",
            ]
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
}
