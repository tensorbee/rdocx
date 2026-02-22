//! Pagination: distribute blocks across pages with constraints.
//!
//! Handles page breaks, widow/orphan control, keep-with-next,
//! keep-lines-together, and header/footer placement.

use crate::block::{LayoutBlock, ParagraphBlock};
use crate::font::FontManager;
use crate::line::{LayoutLine, LineItem};
use crate::output::{Color, GlyphRun, OutlineEntry, PageFrame, Point, PositionedElement, Rect};

use rdocx_oxml::shared::{ST_Border, ST_Jc, ST_Underline};

/// A resolved border edge: (thickness in pt, color, optional dash pattern as (dash, gap)).
type BorderEdge = (f64, Color, Option<(f64, f64)>);

/// Page geometry derived from section properties.
#[derive(Debug, Clone, Copy)]
pub struct PageGeometry {
    pub page_width: f64,
    pub page_height: f64,
    pub margin_top: f64,
    pub margin_right: f64,
    pub margin_bottom: f64,
    pub margin_left: f64,
    pub header_distance: f64,
    pub footer_distance: f64,
}

impl PageGeometry {
    /// Content area width.
    pub fn content_width(&self) -> f64 {
        self.page_width - self.margin_left - self.margin_right
    }

    /// Content area height.
    pub fn content_height(&self) -> f64 {
        self.page_height - self.margin_top - self.margin_bottom
    }
}

impl Default for PageGeometry {
    fn default() -> Self {
        // US Letter with 1" margins
        PageGeometry {
            page_width: 612.0,
            page_height: 792.0,
            margin_top: 72.0,
            margin_right: 72.0,
            margin_bottom: 72.0,
            margin_left: 72.0,
            header_distance: 36.0,
            footer_distance: 36.0,
        }
    }
}

/// Header/footer content already laid out as paragraph blocks.
pub struct HeaderFooterContent {
    pub header_blocks: Vec<ParagraphBlock>,
    pub footer_blocks: Vec<ParagraphBlock>,
    /// First-page header blocks (used when title_pg is true).
    pub first_header_blocks: Vec<ParagraphBlock>,
    /// First-page footer blocks (used when title_pg is true).
    pub first_footer_blocks: Vec<ParagraphBlock>,
}

/// A section with its blocks, geometry, and header/footer content.
pub struct Section {
    pub blocks: Vec<LayoutBlock>,
    pub geometry: PageGeometry,
    pub header_footer: Option<HeaderFooterContent>,
    /// Whether this section uses a different first page header/footer.
    pub title_pg: bool,
}

/// Paginate across multiple sections, each with its own geometry and header/footer.
pub fn paginate_sections(
    sections: &[Section],
    fm: &FontManager,
) -> (Vec<PageFrame>, Vec<OutlineEntry>) {
    if sections.is_empty() {
        return (vec![PageFrame {
            page_number: 1,
            width: 612.0,
            height: 792.0,
            elements: Vec::new(),
        }], Vec::new());
    }

    // For a single section, delegate to the existing paginate function
    if sections.len() == 1 {
        let s = &sections[0];
        return paginate(&s.blocks, s.geometry, s.header_footer.as_ref(), s.title_pg, fm);
    }

    // Multi-section pagination
    let mut all_pages = Vec::new();
    let mut all_outlines = Vec::new();
    let mut page_offset = 0;

    for section in sections {
        let (mut pages, mut outlines) = paginate(
            &section.blocks,
            section.geometry,
            section.header_footer.as_ref(),
            section.title_pg,
            fm,
        );

        // Adjust page numbers and outline page indices
        for page in &mut pages {
            page.page_number += page_offset;
        }
        for outline in &mut outlines {
            outline.page_index += page_offset;
        }

        page_offset += pages.len();
        all_pages.append(&mut pages);
        all_outlines.append(&mut outlines);
    }

    // If a section produced no pages (empty blocks), we might have duplicates
    // Renumber pages sequentially
    for (i, page) in all_pages.iter_mut().enumerate() {
        page.page_number = i + 1;
    }

    (all_pages, all_outlines)
}

/// Paginate a sequence of blocks into pages.
pub fn paginate(
    blocks: &[LayoutBlock],
    geometry: PageGeometry,
    header_footer: Option<&HeaderFooterContent>,
    title_pg: bool,
    _fm: &FontManager,
) -> (Vec<PageFrame>, Vec<OutlineEntry>) {
    let mut pager = Pager::new(geometry, header_footer, title_pg);

    for (block_idx, block) in blocks.iter().enumerate() {
        // Check for page break before
        if block.page_break_before() && pager.has_content() {
            pager.finish_page();
        }

        match block {
            LayoutBlock::Paragraph(para) => {
                // Record heading outline entry before rendering
                if let (Some(level), Some(title)) = (para.heading_level, &para.heading_text) {
                    pager.outlines.push(OutlineEntry {
                        title: title.clone(),
                        level,
                        page_index: pager.page_number - 1,
                        y_position: pager.geometry.margin_top + pager.cursor_y,
                    });
                }
                paginate_paragraph(para, block_idx, blocks, &mut pager);
            }
            LayoutBlock::Table(table) => {
                let table_x = geometry.margin_left + table.table_indent;
                let tbl_borders = table.borders.as_ref();

                for (row_idx, row) in table.rows.iter().enumerate() {
                    if pager.cursor_y + row.height > pager.content_height && pager.has_content() {
                        pager.finish_page();

                        // Repeat header rows
                        for &hdr_idx in &table.header_row_indices {
                            if hdr_idx < row_idx {
                                let hdr_row = &table.rows[hdr_idx];
                                render_table_row(
                                    hdr_row,
                                    &table.col_widths,
                                    table_x,
                                    pager.geometry.margin_top + pager.cursor_y,
                                    &pager.geometry,
                                    tbl_borders,
                                    &mut pager.elements,
                                );
                                pager.cursor_y += hdr_row.height;
                                pager.mark_content();
                            }
                        }
                    }

                    render_table_row(
                        row,
                        &table.col_widths,
                        table_x,
                        pager.geometry.margin_top + pager.cursor_y,
                        &pager.geometry,
                        tbl_borders,
                        &mut pager.elements,
                    );
                    pager.cursor_y += row.height;
                    pager.mark_content();
                }
            }
        }
    }

    pager.flush()
}

/// Helper struct to track page state during pagination.
struct Pager<'a> {
    pages: Vec<PageFrame>,
    elements: Vec<PositionedElement>,
    cursor_y: f64,
    page_number: usize,
    content_height: f64,
    geometry: PageGeometry,
    header_footer: Option<&'a HeaderFooterContent>,
    has_content_flag: bool,
    outlines: Vec<OutlineEntry>,
    /// Whether the current page is the first page of the section.
    is_first_page: bool,
    /// Whether this section uses different first page header/footer.
    title_pg: bool,
}

impl<'a> Pager<'a> {
    fn new(geometry: PageGeometry, header_footer: Option<&'a HeaderFooterContent>, title_pg: bool) -> Self {
        Pager {
            pages: Vec::new(),
            elements: Vec::new(),
            cursor_y: 0.0,
            page_number: 1,
            content_height: geometry.content_height(),
            geometry,
            header_footer,
            has_content_flag: false,
            outlines: Vec::new(),
            is_first_page: true,
            title_pg,
        }
    }

    fn has_content(&self) -> bool {
        self.has_content_flag
    }

    fn mark_content(&mut self) {
        self.has_content_flag = true;
    }

    fn finish_page(&mut self) {
        let mut all_elements = Vec::new();

        if let Some(hf) = self.header_footer {
            // Choose header blocks: first-page or default
            let header_blocks = if self.is_first_page && self.title_pg {
                &hf.first_header_blocks
            } else {
                &hf.header_blocks
            };
            if !header_blocks.is_empty() {
                let header_y = self.geometry.header_distance;
                render_hf_blocks(header_blocks, &self.geometry, header_y, &mut all_elements);
            }
        }

        all_elements.append(&mut self.elements);

        if let Some(hf) = self.header_footer {
            // Choose footer blocks: first-page or default
            let footer_blocks = if self.is_first_page && self.title_pg {
                &hf.first_footer_blocks
            } else {
                &hf.footer_blocks
            };
            if !footer_blocks.is_empty() {
                let footer_height: f64 = footer_blocks.iter().map(|b| b.content_height()).sum();
                let footer_y =
                    self.geometry.page_height - self.geometry.footer_distance - footer_height;
                render_hf_blocks(footer_blocks, &self.geometry, footer_y, &mut all_elements);
            }
        }

        self.pages.push(PageFrame {
            page_number: self.page_number,
            width: self.geometry.page_width,
            height: self.geometry.page_height,
            elements: all_elements,
        });
        self.page_number += 1;
        self.cursor_y = 0.0;
        self.has_content_flag = false;
        self.is_first_page = false;
    }

    fn flush(mut self) -> (Vec<PageFrame>, Vec<OutlineEntry>) {
        // Always create at least one page
        if self.has_content() || self.pages.is_empty() {
            self.finish_page();
        }
        (self.pages, self.outlines)
    }
}

/// Paginate a single paragraph, handling splitting across pages.
fn paginate_paragraph(
    para: &ParagraphBlock,
    block_idx: usize,
    blocks: &[LayoutBlock],
    pager: &mut Pager,
) {
    let space_before = if pager.cursor_y == 0.0 {
        0.0
    } else {
        para.space_before
    };

    // Check if paragraph fits on current page
    let total_needed = space_before + para.content_height();
    let remaining = pager.content_height - pager.cursor_y;

    if total_needed > remaining && pager.has_content() {
        // Paragraph doesn't fit. Decide: move whole or split.
        if para.keep_lines || para.lines.len() <= 2 {
            pager.finish_page();
            // Re-call with fresh page
            paginate_paragraph(para, block_idx, blocks, pager);
            return;
        }

        let available_for_lines = remaining - space_before;
        let lines_that_fit = count_lines_that_fit(&para.lines, available_for_lines);

        if para.widow_control && lines_that_fit < 2 {
            // Can't fit enough lines — move whole paragraph
            pager.finish_page();
            paginate_paragraph(para, block_idx, blocks, pager);
            return;
        }

        let lines_remaining = para.lines.len() - lines_that_fit;
        if para.widow_control && lines_remaining < 2 && lines_that_fit >= 3 {
            // Would leave orphan — move one line to next page
            let split_at = lines_that_fit - 1;
            render_para_split(para, split_at, space_before, pager);
            return;
        }

        if lines_that_fit > 0 {
            render_para_split(para, lines_that_fit, space_before, pager);
            return;
        }

        // No lines fit (shouldn't happen since we checked has_content above)
        pager.finish_page();
        paginate_paragraph(para, block_idx, blocks, pager);
        return;
    }

    // Paragraph fits OR we're at the top of a page
    // If it doesn't fit and we're at the top, we must split line by line
    if total_needed > pager.content_height && pager.cursor_y == 0.0 {
        // Paragraph is taller than a page; split line by line
        let lines_that_fit = count_lines_that_fit(&para.lines, pager.content_height);
        if lines_that_fit > 0 && lines_that_fit < para.lines.len() {
            render_para_split(para, lines_that_fit, 0.0, pager);
            return;
        }
    }

    // Check keep-with-next
    if para.keep_next && block_idx + 1 < blocks.len() {
        let next_first = match &blocks[block_idx + 1] {
            LayoutBlock::Paragraph(p) => p.lines.first().map(|l| l.height).unwrap_or(0.0),
            LayoutBlock::Table(t) => t.rows.first().map(|r| r.height).unwrap_or(0.0),
        };
        if pager.cursor_y + space_before + para.content_height() + next_first > pager.content_height
            && pager.has_content()
        {
            pager.finish_page();
        }
    }

    // Render the paragraph
    let space = if pager.cursor_y == 0.0 {
        0.0
    } else {
        para.space_before
    };
    pager.cursor_y += space;

    if let Some(shading) = para.shading {
        pager.elements.push(PositionedElement::FilledRect {
            rect: Rect {
                x: pager.geometry.margin_left + para.indent_left,
                y: pager.geometry.margin_top + pager.cursor_y,
                width: pager.geometry.content_width() - para.indent_left - para.indent_right,
                height: para.content_height(),
            },
            color: shading,
        });
    }

    // Render paragraph borders
    if let Some(ref borders) = para.borders {
        let border_x = pager.geometry.margin_left + para.indent_left;
        let border_y = pager.geometry.margin_top + pager.cursor_y;
        let border_w = pager.geometry.content_width() - para.indent_left - para.indent_right;
        let border_h = para.content_height();
        render_border_edges(borders, border_x, border_y, border_w, border_h, &mut pager.elements);
    }

    render_paragraph_lines(&para.lines, para, &pager.geometry, pager.cursor_y, &mut pager.elements);
    pager.cursor_y += para.content_height();
    pager.cursor_y += para.space_after;
    pager.mark_content();
}

/// Split a paragraph at the given line index, rendering first part on current page
/// and continuing the rest on a new page (recursively if needed).
fn render_para_split(
    para: &ParagraphBlock,
    split_at: usize,
    space_before: f64,
    pager: &mut Pager,
) {
    // Render lines before split on current page
    pager.cursor_y += space_before;
    render_paragraph_lines(
        &para.lines[..split_at],
        para,
        &pager.geometry,
        pager.cursor_y,
        &mut pager.elements,
    );
    pager.mark_content();
    pager.finish_page();

    // Handle remaining lines, which may themselves need splitting
    let remaining_lines = &para.lines[split_at..];
    let remaining_height: f64 = remaining_lines.iter().map(|l| l.height).sum();

    if remaining_height > pager.content_height {
        // Still too tall — split again
        let lines_that_fit = count_lines_that_fit(remaining_lines, pager.content_height);
        if lines_that_fit > 0 && lines_that_fit < remaining_lines.len() {
            // Build a temporary para with remaining lines
            let temp_para = ParagraphBlock {
                lines: remaining_lines.to_vec(),
                space_before: 0.0,
                space_after: para.space_after,
                borders: para.borders.clone(),
                shading: para.shading,
                indent_left: para.indent_left,
                indent_right: para.indent_right,
                jc: para.jc,
                keep_next: para.keep_next,
                keep_lines: false,
                page_break_before: false,
                widow_control: para.widow_control,
                heading_level: None,
                heading_text: None,
            };
            render_para_split(&temp_para, lines_that_fit, 0.0, pager);
            return;
        }
    }

    // Remaining fits on the new page
    render_paragraph_lines(remaining_lines, para, &pager.geometry, 0.0, &mut pager.elements);
    pager.cursor_y = remaining_height + para.space_after;
    pager.mark_content();
}

/// Count how many lines fit in the remaining space.
fn count_lines_that_fit(lines: &[LayoutLine], available: f64) -> usize {
    let mut used = 0.0;
    for (i, line) in lines.iter().enumerate() {
        used += line.height;
        if used > available {
            return i;
        }
    }
    lines.len()
}

/// Render paragraph lines as positioned elements.
fn render_paragraph_lines(
    lines: &[LayoutLine],
    para: &ParagraphBlock,
    geometry: &PageGeometry,
    start_y: f64,
    elements: &mut Vec<PositionedElement>,
) {
    let mut y = start_y;
    for line in lines {
        let baseline_y = geometry.margin_top + y + line.ascent;

        // Compute x offset based on justification
        let text_width: f64 = line.items.iter().map(|item| item.width()).sum();
        let remaining_width = line.available_width - text_width;

        // For justified text (Both), compute extra space per gap
        let justify_extra = if para.jc == Some(ST_Jc::Both) && !line.is_last && remaining_width > 0.0 {
            // Count inter-word gaps: spaces between items + spaces within text segments
            let gap_count = count_word_gaps(&line.items);
            if gap_count > 0 {
                remaining_width / gap_count as f64
            } else {
                0.0
            }
        } else {
            0.0
        };

        let x_offset = match para.jc {
            Some(ST_Jc::Center) => {
                geometry.margin_left + line.indent_left + remaining_width / 2.0
            }
            Some(ST_Jc::Right) | Some(ST_Jc::End) => {
                geometry.margin_left + line.indent_left + remaining_width
            }
            Some(ST_Jc::Both) if !line.is_last && justify_extra > 0.0 => {
                // Justified: start from left margin (extra space distributed in gaps)
                geometry.margin_left + line.indent_left
            }
            _ => geometry.margin_left + line.indent_left,
        };

        let mut x = x_offset;
        let mut _accumulated_extra = 0.0;

        for item in &line.items {
            match item {
                LineItem::Text(seg) | LineItem::Marker(seg) => {
                    let adjusted_baseline = baseline_y - seg.baseline_offset;

                    // For justified text, compute the extra width from spaces in this segment
                    let segment_spaces = if justify_extra > 0.0 {
                        seg.text.chars().filter(|c| *c == ' ').count()
                    } else {
                        0
                    };
                    let segment_extra = segment_spaces as f64 * justify_extra;
                    let effective_width = seg.width + segment_extra;

                    // Render highlight background
                    if let Some(hl_color) = seg.highlight {
                        elements.push(PositionedElement::FilledRect {
                            rect: Rect {
                                x,
                                y: geometry.margin_top + y,
                                width: effective_width,
                                height: line.height,
                            },
                            color: hl_color,
                        });
                    }

                    // Render text, adjusting advances for justified text
                    let advances = if justify_extra > 0.0 && segment_spaces > 0 {
                        // Widen advances for space glyphs
                        distribute_justify_advances(&seg.text, &seg.advances, justify_extra)
                    } else {
                        seg.advances.clone()
                    };

                    elements.push(PositionedElement::Text(GlyphRun {
                        origin: Point { x, y: adjusted_baseline },
                        font_id: seg.font_id,
                        font_size: seg.font_size,
                        glyph_ids: seg.glyph_ids.clone(),
                        advances,
                        text: seg.text.clone(),
                        color: seg.color,
                        bold: seg.bold,
                        italic: seg.italic,
                        field_kind: seg.field_kind,
                        footnote_id: seg.footnote_id,
                    }));

                    // Render underline
                    if let Some(ul_style) = seg.underline {
                        if ul_style != ST_Underline::None {
                            let ul_y = adjusted_baseline + seg.descent * 0.3;
                            let ul_thickness = match ul_style {
                                ST_Underline::Thick => seg.font_size / 12.0,
                                ST_Underline::Double => seg.font_size / 24.0,
                                _ => seg.font_size / 18.0,
                            };
                            elements.push(PositionedElement::Line {
                                start: Point { x, y: ul_y },
                                end: Point { x: x + effective_width, y: ul_y },
                                width: ul_thickness,
                                color: seg.color,
                                dash_pattern: None,
                            });
                            // Second line for double underline
                            if ul_style == ST_Underline::Double {
                                let ul_y2 = ul_y + ul_thickness * 2.5;
                                elements.push(PositionedElement::Line {
                                    start: Point { x, y: ul_y2 },
                                    end: Point { x: x + effective_width, y: ul_y2 },
                                    width: ul_thickness,
                                    color: seg.color,
                                    dash_pattern: None,
                                });
                            }
                        }
                    }

                    // Render strikethrough
                    if seg.strike {
                        let strike_y = adjusted_baseline - seg.ascent * 0.3;
                        let strike_thickness = seg.font_size / 24.0;
                        elements.push(PositionedElement::Line {
                            start: Point { x, y: strike_y },
                            end: Point { x: x + effective_width, y: strike_y },
                            width: strike_thickness,
                            color: seg.color,
                            dash_pattern: None,
                        });
                    }

                    // Render double strikethrough
                    if seg.dstrike {
                        let strike_y = adjusted_baseline - seg.ascent * 0.3;
                        let strike_thickness = seg.font_size / 24.0;
                        let gap = strike_thickness * 2.0;
                        elements.push(PositionedElement::Line {
                            start: Point { x, y: strike_y - gap / 2.0 },
                            end: Point { x: x + effective_width, y: strike_y - gap / 2.0 },
                            width: strike_thickness,
                            color: seg.color,
                            dash_pattern: None,
                        });
                        elements.push(PositionedElement::Line {
                            start: Point { x, y: strike_y + gap / 2.0 },
                            end: Point { x: x + effective_width, y: strike_y + gap / 2.0 },
                            width: strike_thickness,
                            color: seg.color,
                            dash_pattern: None,
                        });
                    }

                    // Render hyperlink annotation
                    if let Some(ref url) = seg.hyperlink_url {
                        elements.push(PositionedElement::LinkAnnotation {
                            rect: Rect {
                                x,
                                y: geometry.margin_top + y,
                                width: effective_width,
                                height: line.height,
                            },
                            url: url.clone(),
                        });
                    }

                    _accumulated_extra += segment_extra;
                    x += effective_width;
                }
                LineItem::Tab { width, leader } => {
                    if let Some(leader_seg) = leader {
                        // Render the pre-shaped leader text
                        let baseline_y = geometry.margin_top + y + line.ascent;
                        elements.push(PositionedElement::Text(GlyphRun {
                            origin: Point { x, y: baseline_y },
                            font_id: leader_seg.font_id,
                            font_size: leader_seg.font_size,
                            glyph_ids: leader_seg.glyph_ids.clone(),
                            advances: leader_seg.advances.clone(),
                            text: leader_seg.text.clone(),
                            color: leader_seg.color,
                            bold: leader_seg.bold,
                            italic: leader_seg.italic,
                            field_kind: None,
                            footnote_id: None,
                        }));
                    }
                    x += width;
                }
                LineItem::Image {
                    width,
                    height,
                    embed_id,
                } => {
                    // Image positioned at current x, top-aligned with line
                    elements.push(PositionedElement::Image {
                        rect: Rect {
                            x,
                            y: geometry.margin_top + y,
                            width: *width,
                            height: *height,
                        },
                        data: Vec::new(),
                        content_type: String::new(),
                        embed_id: Some(embed_id.clone()),
                    });
                    x += width;
                }
            }
        }

        y += line.height;
    }
}

/// Render header/footer blocks.
fn render_hf_blocks(
    blocks: &[ParagraphBlock],
    geometry: &PageGeometry,
    start_y: f64,
    elements: &mut Vec<PositionedElement>,
) {
    let mut y = start_y - geometry.margin_top; // Convert to relative
    for para in blocks {
        render_paragraph_lines(&para.lines, para, geometry, y, elements);
        y += para.content_height();
    }
}

/// Render a table row.
fn render_table_row(
    row: &crate::table::TableRow,
    _col_widths: &[f64],
    table_x: f64,
    row_y: f64,
    geometry: &PageGeometry,
    table_borders: Option<&rdocx_oxml::table::CT_TblBorders>,
    elements: &mut Vec<PositionedElement>,
) {
    let mut cell_x = table_x;
    let num_cells = row.cells.len();

    for (cell_idx, cell) in row.cells.iter().enumerate() {
        // Render cell shading
        if let Some(ref shading) = cell.shading {
            elements.push(PositionedElement::FilledRect {
                rect: Rect {
                    x: cell_x,
                    y: row_y,
                    width: cell.width,
                    height: cell.height,
                },
                color: *shading,
            });
        }

        // Render cell borders
        render_cell_borders(
            cell_x,
            row_y,
            cell.width,
            cell.height,
            &cell.borders,
            table_borders,
            cell_idx,
            num_cells,
            cell.is_first_row,
            cell.is_last_row,
            elements,
        );

        if !cell.is_vmerge_continue {
            // Render cell content
            let cell_margin_top = cell.margin_top;
            let cell_margin_left = cell.margin_left;

            // Compute vertical alignment offset
            let content_height: f64 = cell.paragraphs.iter().map(|p| p.total_height()).sum();
            let v_offset = match cell.v_align {
                Some(rdocx_oxml::table::ST_VerticalJc::Center) => {
                    ((cell.height - cell_margin_top - content_height) / 2.0).max(0.0)
                }
                Some(rdocx_oxml::table::ST_VerticalJc::Bottom) => {
                    (cell.height - cell_margin_top - content_height).max(0.0)
                }
                _ => 0.0, // Top or unspecified
            };

            let mut para_y = row_y - geometry.margin_top + cell_margin_top + v_offset;
            for para in &cell.paragraphs {
                render_paragraph_lines(
                    &para.lines,
                    para,
                    &PageGeometry {
                        margin_left: cell_x + cell_margin_left,
                        ..*geometry
                    },
                    para_y,
                    elements,
                );
                para_y += para.total_height();
            }
        }
        cell_x += cell.width;
    }
}

/// Render borders for a table cell.
fn render_cell_borders(
    x: f64,
    y: f64,
    w: f64,
    h: f64,
    cell_borders: &Option<rdocx_oxml::table::CT_TblBorders>,
    table_borders: Option<&rdocx_oxml::table::CT_TblBorders>,
    cell_idx: usize,
    num_cells: usize,
    is_first_row: bool,
    is_last_row: bool,
    elements: &mut Vec<PositionedElement>,
) {
    // Determine effective border for each edge (cell overrides table)
    let get_edge = |cell_edge: Option<&rdocx_oxml::borders::CT_BorderEdge>,
                    table_edge: Option<&rdocx_oxml::borders::CT_BorderEdge>|
     -> Option<BorderEdge> {
        let edge = cell_edge.or(table_edge)?;
        if edge.val == ST_Border::None {
            return None;
        }
        let thickness = edge.sz.unwrap_or(4) as f64 / 8.0; // sz is in 1/8 pt
        let color = edge
            .color
            .as_ref()
            .filter(|c| c.as_str() != "auto")
            .map(|c| Color::from_hex(c))
            .unwrap_or(Color::BLACK);
        let dash = border_dash_pattern(edge.val, thickness);
        Some((thickness, color, dash))
    };

    // Top border: use table top for first row, table insideH otherwise
    let table_top = table_borders.and_then(|b| {
        if is_first_row {
            b.top.as_ref()
        } else {
            b.inside_h.as_ref()
        }
    });
    let cell_top = cell_borders.as_ref().and_then(|b| b.top.as_ref());
    if let Some((thickness, color, dash_pattern)) = get_edge(cell_top, table_top) {
        elements.push(PositionedElement::Line {
            start: Point { x, y },
            end: Point { x: x + w, y },
            width: thickness,
            color,
            dash_pattern,
        });
    }

    // Bottom border: use table bottom for last row, table insideH otherwise
    let table_bottom = table_borders.and_then(|b| {
        if is_last_row {
            b.bottom.as_ref()
        } else {
            b.inside_h.as_ref()
        }
    });
    let cell_bottom = cell_borders.as_ref().and_then(|b| b.bottom.as_ref());
    if let Some((thickness, color, dash_pattern)) = get_edge(cell_bottom, table_bottom) {
        elements.push(PositionedElement::Line {
            start: Point { x, y: y + h },
            end: Point { x: x + w, y: y + h },
            width: thickness,
            color,
            dash_pattern,
        });
    }

    // Left border: use table left for first cell, table insideV otherwise
    let table_left = table_borders.and_then(|b| {
        if cell_idx == 0 {
            b.left.as_ref()
        } else {
            b.inside_v.as_ref()
        }
    });
    let cell_left = cell_borders.as_ref().and_then(|b| b.left.as_ref());
    if let Some((thickness, color, dash_pattern)) = get_edge(cell_left, table_left) {
        elements.push(PositionedElement::Line {
            start: Point { x, y },
            end: Point { x, y: y + h },
            width: thickness,
            color,
            dash_pattern,
        });
    }

    // Right border: use table right for last cell, table insideV otherwise
    let table_right = table_borders.and_then(|b| {
        if cell_idx == num_cells - 1 {
            b.right.as_ref()
        } else {
            b.inside_v.as_ref()
        }
    });
    let cell_right = cell_borders.as_ref().and_then(|b| b.right.as_ref());
    if let Some((thickness, color, dash_pattern)) = get_edge(cell_right, table_right) {
        elements.push(PositionedElement::Line {
            start: Point { x: x + w, y },
            end: Point { x: x + w, y: y + h },
            width: thickness,
            color,
            dash_pattern,
        });
    }
}

/// Render paragraph border edges as positioned lines.
fn render_border_edges(
    borders: &rdocx_oxml::borders::CT_PBdr,
    x: f64,
    y: f64,
    w: f64,
    h: f64,
    elements: &mut Vec<PositionedElement>,
) {
    let render_edge = |edge: &rdocx_oxml::borders::CT_BorderEdge,
                       start: Point,
                       end: Point,
                       elements: &mut Vec<PositionedElement>| {
        if edge.val == ST_Border::None {
            return;
        }
        let thickness = edge.sz.unwrap_or(4) as f64 / 8.0; // sz is in eighths of a point
        let color = edge
            .color
            .as_ref()
            .filter(|c| c.as_str() != "auto")
            .map(|c| Color::from_hex(c))
            .unwrap_or(Color::BLACK);
        let dash_pattern = border_dash_pattern(edge.val, thickness);

        if edge.val == ST_Border::Double {
            // Double border: emit two parallel lines
            let gap = thickness * 2.0;
            let dx = end.x - start.x;
            let dy = end.y - start.y;
            let len = (dx * dx + dy * dy).sqrt();
            let (nx, ny) = if len > 0.0 { (-dy / len, dx / len) } else { (0.0, 1.0) };
            let offset = gap / 2.0;
            elements.push(PositionedElement::Line {
                start: Point { x: start.x + nx * offset, y: start.y + ny * offset },
                end: Point { x: end.x + nx * offset, y: end.y + ny * offset },
                width: thickness,
                color,
                dash_pattern: None,
            });
            elements.push(PositionedElement::Line {
                start: Point { x: start.x - nx * offset, y: start.y - ny * offset },
                end: Point { x: end.x - nx * offset, y: end.y - ny * offset },
                width: thickness,
                color,
                dash_pattern: None,
            });
        } else {
            elements.push(PositionedElement::Line {
                start,
                end,
                width: thickness,
                color,
                dash_pattern,
            });
        }
    };

    if let Some(ref edge) = borders.top {
        let space = edge.space.unwrap_or(0) as f64;
        render_edge(
            edge,
            Point { x, y: y - space },
            Point { x: x + w, y: y - space },
            elements,
        );
    }
    if let Some(ref edge) = borders.bottom {
        let space = edge.space.unwrap_or(0) as f64;
        render_edge(
            edge,
            Point { x, y: y + h + space },
            Point { x: x + w, y: y + h + space },
            elements,
        );
    }
    if let Some(ref edge) = borders.left {
        let space = edge.space.unwrap_or(0) as f64;
        render_edge(
            edge,
            Point { x: x - space, y },
            Point { x: x - space, y: y + h },
            elements,
        );
    }
    if let Some(ref edge) = borders.right {
        let space = edge.space.unwrap_or(0) as f64;
        render_edge(
            edge,
            Point { x: x + w + space, y },
            Point { x: x + w + space, y: y + h },
            elements,
        );
    }
}

/// Map a border style to a dash pattern (dash_on, dash_off) in points.
/// Returns None for solid lines (Single, Thick, Double, etc.).
fn border_dash_pattern(style: ST_Border, thickness: f64) -> Option<(f64, f64)> {
    match style {
        ST_Border::Dashed => Some((3.0 * thickness, 2.0 * thickness)),
        ST_Border::Dotted => Some((thickness, thickness)),
        ST_Border::DotDash | ST_Border::DotDotDash => Some((3.0 * thickness, thickness)),
        _ => None,
    }
}

/// Count inter-word gap positions in a line (spaces within text segments).
fn count_word_gaps(items: &[LineItem]) -> usize {
    let mut count = 0;
    for item in items {
        match item {
            LineItem::Text(seg) | LineItem::Marker(seg) => {
                count += seg.text.chars().filter(|c| *c == ' ').count();
            }
            LineItem::Tab { .. } => {
                count += 1;
            }
            _ => {}
        }
    }
    count
}

/// Distribute extra justify space across advances by widening space-character advances.
fn distribute_justify_advances(text: &str, advances: &[f64], extra_per_gap: f64) -> Vec<f64> {
    let chars: Vec<char> = text.chars().collect();
    let mut result = advances.to_vec();

    if chars.len() == result.len() {
        // 1:1 char-to-glyph mapping
        for (i, &ch) in chars.iter().enumerate() {
            if ch == ' ' {
                result[i] += extra_per_gap;
            }
        }
    } else {
        // Fallback: distribute evenly across all glyphs
        let total_extra = extra_per_gap * text.chars().filter(|c| *c == ' ').count() as f64;
        if !result.is_empty() {
            let per_glyph = total_extra / result.len() as f64;
            for a in &mut result {
                *a += per_glyph;
            }
        }
    }

    result
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::block::ParagraphBlock;
    use crate::line::LayoutLine;

    fn make_line(height: f64) -> LayoutLine {
        LayoutLine {
            items: vec![],
            width: 100.0,
            ascent: height * 0.77,
            descent: height * 0.23,
            height,
            indent_left: 0.0,
            available_width: 468.0,
            is_last: true,
        }
    }

    fn make_para(line_count: usize, line_height: f64) -> ParagraphBlock {
        let mut lines = Vec::new();
        for _ in 0..line_count {
            lines.push(make_line(line_height));
        }
        ParagraphBlock {
            lines,
            space_before: 0.0,
            space_after: 0.0,
            borders: None,
            shading: None,
            indent_left: 0.0,
            indent_right: 0.0,
            jc: None,
            keep_next: false,
            keep_lines: false,
            page_break_before: false,
            widow_control: true,
            heading_level: None,
            heading_text: None,
        }
    }

    #[test]
    fn single_page_layout() {
        let fm = FontManager::new();
        let blocks = vec![LayoutBlock::Paragraph(make_para(3, 14.0))];
        let geom = PageGeometry::default();
        let (pages, _outlines) = paginate(&blocks, geom, None, false, &fm);
        assert_eq!(pages.len(), 1);
        assert_eq!(pages[0].page_number, 1);
    }

    #[test]
    fn multi_page_overflow() {
        let fm = FontManager::new();
        // 648pt content height / 14pt per line ≈ 46 lines per page
        let blocks = vec![LayoutBlock::Paragraph(make_para(100, 14.0))];
        let geom = PageGeometry::default();
        let (pages, _outlines) = paginate(&blocks, geom, None, false, &fm);
        assert!(pages.len() >= 2);
    }

    #[test]
    fn forced_page_break() {
        let fm = FontManager::new();
        let mut para2 = make_para(3, 14.0);
        para2.page_break_before = true;
        let blocks = vec![
            LayoutBlock::Paragraph(make_para(3, 14.0)),
            LayoutBlock::Paragraph(para2),
        ];
        let geom = PageGeometry::default();
        let (pages, _outlines) = paginate(&blocks, geom, None, false, &fm);
        assert_eq!(pages.len(), 2);
    }

    #[test]
    fn page_dimensions() {
        let fm = FontManager::new();
        let blocks = vec![LayoutBlock::Paragraph(make_para(1, 14.0))];
        let geom = PageGeometry::default();
        let (pages, _outlines) = paginate(&blocks, geom, None, false, &fm);
        assert!((pages[0].width - 612.0).abs() < 0.01);
        assert!((pages[0].height - 792.0).abs() < 0.01);
    }

    fn make_text_line(height: f64, underline: Option<ST_Underline>, strike: bool) -> LayoutLine {
        use crate::line::TextSegment;
        let seg = TextSegment {
            text: "Hello".to_string(),
            font_id: crate::output::FontId(0),
            font_size: 12.0,
            glyph_ids: vec![1, 2, 3],
            advances: vec![6.0, 6.0, 6.0],
            width: 40.0,
            ascent: height * 0.77,
            descent: height * 0.23,
            color: Color::BLACK,
            bold: false,
            italic: false,
            underline,
            strike,
            dstrike: false,
            highlight: None,
            baseline_offset: 0.0,
            hyperlink_url: None,
            field_kind: None,
            footnote_id: None,
        };
        LayoutLine {
            items: vec![LineItem::Text(seg)],
            width: 40.0,
            ascent: height * 0.77,
            descent: height * 0.23,
            height,
            indent_left: 0.0,
            available_width: 468.0,
            is_last: true,
        }
    }

    #[test]
    fn underline_renders_line_element() {
        let fm = FontManager::new();
        let para = ParagraphBlock {
            lines: vec![make_text_line(14.0, Some(ST_Underline::Single), false)],
            space_before: 0.0,
            space_after: 0.0,
            borders: None,
            shading: None,
            indent_left: 0.0,
            indent_right: 0.0,
            jc: None,
            keep_next: false,
            keep_lines: false,
            page_break_before: false,
            widow_control: true,
            heading_level: None,
            heading_text: None,
        };
        let blocks = vec![LayoutBlock::Paragraph(para)];
        let (pages, _outlines) = paginate(&blocks, PageGeometry::default(), None, false, &fm);
        // Should have Text + Line (underline)
        let lines: Vec<_> = pages[0]
            .elements
            .iter()
            .filter(|e| matches!(e, PositionedElement::Line { .. }))
            .collect();
        assert_eq!(lines.len(), 1, "expected 1 underline line");
    }

    #[test]
    fn strikethrough_renders_line_element() {
        let fm = FontManager::new();
        let para = ParagraphBlock {
            lines: vec![make_text_line(14.0, None, true)],
            space_before: 0.0,
            space_after: 0.0,
            borders: None,
            shading: None,
            indent_left: 0.0,
            indent_right: 0.0,
            jc: None,
            keep_next: false,
            keep_lines: false,
            page_break_before: false,
            widow_control: true,
            heading_level: None,
            heading_text: None,
        };
        let blocks = vec![LayoutBlock::Paragraph(para)];
        let (pages, _outlines) = paginate(&blocks, PageGeometry::default(), None, false, &fm);
        let lines: Vec<_> = pages[0]
            .elements
            .iter()
            .filter(|e| matches!(e, PositionedElement::Line { .. }))
            .collect();
        assert_eq!(lines.len(), 1, "expected 1 strikethrough line");
    }

    #[test]
    fn highlight_renders_filled_rect() {
        use crate::line::TextSegment;
        let fm = FontManager::new();
        let seg = TextSegment {
            text: "Hi".to_string(),
            font_id: crate::output::FontId(0),
            font_size: 12.0,
            glyph_ids: vec![1],
            advances: vec![10.0],
            width: 20.0,
            ascent: 10.0,
            descent: 3.0,
            color: Color::BLACK,
            bold: false,
            italic: false,
            underline: None,
            strike: false,
            dstrike: false,
            highlight: Some(Color { r: 1.0, g: 1.0, b: 0.0, a: 1.0 }),
            baseline_offset: 0.0,
            hyperlink_url: None,
            field_kind: None,
            footnote_id: None,
        };
        let line = LayoutLine {
            items: vec![LineItem::Text(seg)],
            width: 20.0,
            ascent: 10.0,
            descent: 3.0,
            height: 13.0,
            indent_left: 0.0,
            available_width: 468.0,
            is_last: true,
        };
        let para = ParagraphBlock {
            lines: vec![line],
            space_before: 0.0,
            space_after: 0.0,
            borders: None,
            shading: None,
            indent_left: 0.0,
            indent_right: 0.0,
            jc: None,
            keep_next: false,
            keep_lines: false,
            page_break_before: false,
            widow_control: true,
            heading_level: None,
            heading_text: None,
        };
        let blocks = vec![LayoutBlock::Paragraph(para)];
        let (pages, _outlines) = paginate(&blocks, PageGeometry::default(), None, false, &fm);
        let rects: Vec<_> = pages[0]
            .elements
            .iter()
            .filter(|e| matches!(e, PositionedElement::FilledRect { .. }))
            .collect();
        assert_eq!(rects.len(), 1, "expected 1 highlight rect");
    }

    #[test]
    fn paragraph_borders_render_lines() {
        use rdocx_oxml::borders::{CT_BorderEdge, CT_PBdr};
        let fm = FontManager::new();
        let para = ParagraphBlock {
            lines: vec![make_line(14.0)],
            space_before: 0.0,
            space_after: 0.0,
            borders: Some(CT_PBdr {
                top: Some(CT_BorderEdge {
                    val: ST_Border::Single,
                    sz: Some(4),
                    space: Some(1),
                    color: Some("000000".to_string()),
                }),
                bottom: Some(CT_BorderEdge {
                    val: ST_Border::Single,
                    sz: Some(4),
                    space: Some(1),
                    color: Some("000000".to_string()),
                }),
                ..Default::default()
            }),
            shading: None,
            indent_left: 0.0,
            indent_right: 0.0,
            jc: None,
            keep_next: false,
            keep_lines: false,
            page_break_before: false,
            widow_control: true,
            heading_level: None,
            heading_text: None,
        };
        let blocks = vec![LayoutBlock::Paragraph(para)];
        let (pages, _outlines) = paginate(&blocks, PageGeometry::default(), None, false, &fm);
        let lines: Vec<_> = pages[0]
            .elements
            .iter()
            .filter(|e| matches!(e, PositionedElement::Line { .. }))
            .collect();
        assert_eq!(lines.len(), 2, "expected 2 border lines (top + bottom)");
    }

    #[test]
    fn paragraph_shading_renders_filled_rect() {
        let fm = FontManager::new();
        let para = ParagraphBlock {
            lines: vec![make_line(14.0)],
            space_before: 0.0,
            space_after: 0.0,
            borders: None,
            shading: Some(Color { r: 1.0, g: 1.0, b: 0.0, a: 1.0 }),
            indent_left: 0.0,
            indent_right: 0.0,
            jc: None,
            keep_next: false,
            keep_lines: false,
            page_break_before: false,
            widow_control: true,
            heading_level: None,
            heading_text: None,
        };
        let blocks = vec![LayoutBlock::Paragraph(para)];
        let (pages, _outlines) = paginate(&blocks, PageGeometry::default(), None, false, &fm);
        let rects: Vec<_> = pages[0]
            .elements
            .iter()
            .filter(|e| matches!(e, PositionedElement::FilledRect { .. }))
            .collect();
        assert_eq!(rects.len(), 1, "expected 1 paragraph shading rect");
    }

    #[test]
    fn double_underline_renders_two_lines() {
        let fm = FontManager::new();
        let para = ParagraphBlock {
            lines: vec![make_text_line(14.0, Some(ST_Underline::Double), false)],
            space_before: 0.0,
            space_after: 0.0,
            borders: None,
            shading: None,
            indent_left: 0.0,
            indent_right: 0.0,
            jc: None,
            keep_next: false,
            keep_lines: false,
            page_break_before: false,
            widow_control: true,
            heading_level: None,
            heading_text: None,
        };
        let blocks = vec![LayoutBlock::Paragraph(para)];
        let (pages, _outlines) = paginate(&blocks, PageGeometry::default(), None, false, &fm);
        let lines: Vec<_> = pages[0]
            .elements
            .iter()
            .filter(|e| matches!(e, PositionedElement::Line { .. }))
            .collect();
        assert_eq!(lines.len(), 2, "expected 2 lines for double underline");
    }

    fn make_justified_line(text: &str, seg_width: f64, is_last: bool) -> LayoutLine {
        use crate::line::TextSegment;
        let seg = TextSegment {
            text: text.to_string(),
            font_id: crate::output::FontId(0),
            font_size: 12.0,
            glyph_ids: vec![1; text.len()],
            advances: vec![seg_width / text.len() as f64; text.len()],
            width: seg_width,
            ascent: 10.0,
            descent: 3.0,
            color: Color::BLACK,
            bold: false,
            italic: false,
            underline: None,
            strike: false,
            dstrike: false,
            highlight: None,
            baseline_offset: 0.0,
            hyperlink_url: None,
            field_kind: None,
            footnote_id: None,
        };
        LayoutLine {
            items: vec![LineItem::Text(seg)],
            width: seg_width,
            ascent: 10.0,
            descent: 3.0,
            height: 13.0,
            indent_left: 0.0,
            available_width: 468.0,
            is_last,
        }
    }

    #[test]
    fn hyperlink_emits_link_annotation() {
        use crate::line::TextSegment;
        let fm = FontManager::new();
        let seg = TextSegment {
            text: "Click me".to_string(),
            font_id: crate::output::FontId(0),
            font_size: 12.0,
            glyph_ids: vec![1, 2, 3],
            advances: vec![8.0, 8.0, 8.0],
            width: 60.0,
            ascent: 10.0,
            descent: 3.0,
            color: Color::BLACK,
            bold: false,
            italic: false,
            underline: None,
            strike: false,
            dstrike: false,
            highlight: None,
            baseline_offset: 0.0,
            hyperlink_url: Some("https://example.com".to_string()),
            field_kind: None,
            footnote_id: None,
        };
        let line = LayoutLine {
            items: vec![LineItem::Text(seg)],
            width: 60.0,
            ascent: 10.0,
            descent: 3.0,
            height: 13.0,
            indent_left: 0.0,
            available_width: 468.0,
            is_last: true,
        };
        let para = ParagraphBlock {
            lines: vec![line],
            space_before: 0.0,
            space_after: 0.0,
            borders: None,
            shading: None,
            indent_left: 0.0,
            indent_right: 0.0,
            jc: None,
            keep_next: false,
            keep_lines: false,
            page_break_before: false,
            widow_control: true,
            heading_level: None,
            heading_text: None,
        };
        let blocks = vec![LayoutBlock::Paragraph(para)];
        let (pages, _outlines) = paginate(&blocks, PageGeometry::default(), None, false, &fm);
        let annotations: Vec<_> = pages[0]
            .elements
            .iter()
            .filter(|e| matches!(e, PositionedElement::LinkAnnotation { .. }))
            .collect();
        assert_eq!(annotations.len(), 1, "expected 1 link annotation");
        if let PositionedElement::LinkAnnotation { url, .. } = annotations[0] {
            assert_eq!(url, "https://example.com");
        }
    }

    #[test]
    fn justified_text_fills_line_width() {
        let fm = FontManager::new();
        // Line with "Hello World" (1 space = 1 gap), width 200 out of 468 available
        let para = ParagraphBlock {
            lines: vec![
                make_justified_line("Hello World", 200.0, false),
                make_justified_line("End.", 40.0, true),
            ],
            space_before: 0.0,
            space_after: 0.0,
            borders: None,
            shading: None,
            indent_left: 0.0,
            indent_right: 0.0,
            jc: Some(ST_Jc::Both),
            keep_next: false,
            keep_lines: false,
            page_break_before: false,
            widow_control: true,
            heading_level: None,
            heading_text: None,
        };

        let blocks = vec![LayoutBlock::Paragraph(para)];
        let (pages, _outlines) = paginate(&blocks, PageGeometry::default(), None, false, &fm);

        // The first line's text run should have widened advances
        let first_text = pages[0].elements.iter().find_map(|e| {
            if let PositionedElement::Text(run) = e {
                Some(run)
            } else {
                None
            }
        });
        assert!(first_text.is_some());
        let run = first_text.unwrap();
        // The total advance should be wider than the original 200pt
        let total_advance: f64 = run.advances.iter().sum();
        assert!(
            total_advance > 200.0,
            "justified text should be wider than original: {total_advance}"
        );
    }

    #[test]
    fn justified_last_line_stays_left_aligned() {
        let fm = FontManager::new();
        let para = ParagraphBlock {
            lines: vec![
                make_justified_line("Hello World Test", 200.0, false),
                make_justified_line("End.", 40.0, true),
            ],
            space_before: 0.0,
            space_after: 0.0,
            borders: None,
            shading: None,
            indent_left: 0.0,
            indent_right: 0.0,
            jc: Some(ST_Jc::Both),
            keep_next: false,
            keep_lines: false,
            page_break_before: false,
            widow_control: true,
            heading_level: None,
            heading_text: None,
        };

        let blocks = vec![LayoutBlock::Paragraph(para)];
        let (pages, _outlines) = paginate(&blocks, PageGeometry::default(), None, false, &fm);

        // Find the second text run (last line)
        let text_runs: Vec<_> = pages[0]
            .elements
            .iter()
            .filter_map(|e| {
                if let PositionedElement::Text(run) = e {
                    Some(run)
                } else {
                    None
                }
            })
            .collect();

        assert!(text_runs.len() >= 2);
        // Last line should NOT be stretched — advances should sum to original width
        let last_advance: f64 = text_runs[1].advances.iter().sum();
        assert!(
            (last_advance - 40.0).abs() < 0.1,
            "last line should stay at original width: {last_advance}"
        );
    }

    #[test]
    fn justified_single_word_not_stretched() {
        let fm = FontManager::new();
        // A line with a single word (no spaces) should not be stretched
        let para = ParagraphBlock {
            lines: vec![
                make_justified_line("Superlongword", 100.0, false),
                make_justified_line("End.", 40.0, true),
            ],
            space_before: 0.0,
            space_after: 0.0,
            borders: None,
            shading: None,
            indent_left: 0.0,
            indent_right: 0.0,
            jc: Some(ST_Jc::Both),
            keep_next: false,
            keep_lines: false,
            page_break_before: false,
            widow_control: true,
            heading_level: None,
            heading_text: None,
        };

        let blocks = vec![LayoutBlock::Paragraph(para)];
        let (pages, _outlines) = paginate(&blocks, PageGeometry::default(), None, false, &fm);

        let first_text = pages[0].elements.iter().find_map(|e| {
            if let PositionedElement::Text(run) = e {
                Some(run)
            } else {
                None
            }
        });
        assert!(first_text.is_some());
        let run = first_text.unwrap();
        let total_advance: f64 = run.advances.iter().sum();
        // No spaces → no stretching
        assert!(
            (total_advance - 100.0).abs() < 0.1,
            "single word should not be stretched: {total_advance}"
        );
    }
}
