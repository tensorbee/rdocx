//! Pagination: distribute blocks across pages with constraints.
//!
//! Handles page breaks, widow/orphan control, keep-with-next,
//! keep-lines-together, and header/footer placement.

use crate::block::{AnchoredContent, AnchoredDrawing, LayoutBlock, ParagraphBlock, ShapePreset};
use std::collections::HashMap;
use std::sync::Arc;

use oxml_layout::{
    Align, Color, FontManager, GlyphRun, LayoutLine, LineItem, MediaId, NoteRef, NoteStream,
    OutlineEntry, PageFrame, Point, PositionedElement, Rect, Underline, break_into_lines,
};

use rdocx_oxml::drawing::{
    AnchorAlignH, AnchorAlignV, ST_RelativeFromH, ST_RelativeFromV, WrapType,
};
use rdocx_oxml::shared::ST_Border;

use crate::input::{ImageData, MediaRegistry};
use crate::notes::{
    NOTE_INDENT, NOTE_SEPARATOR_OFFSET, NoteLayout, NoteRegistry, SEPARATOR_WIDTH_FRACTION,
};

/// A wrapping drawing that has been placed on the page being built.
#[derive(Debug, Clone, Copy)]
struct PlacedWrap {
    rect: Rect,
    wrap: WrapType,
    dist_top: f64,
    dist_bottom: f64,
    dist_left: f64,
    dist_right: f64,
}

impl PlacedWrap {
    /// Top of the band this drawing keeps text out of.
    fn keep_out_top(&self) -> f64 {
        self.rect.y - self.dist_top
    }

    /// Bottom of the band this drawing keeps text out of.
    fn keep_out_bottom(&self) -> f64 {
        self.rect.y + self.rect.height + self.dist_bottom
    }
}

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
    /// Even-page header blocks.
    pub even_header_blocks: Vec<ParagraphBlock>,
    /// Even-page footer blocks.
    pub even_footer_blocks: Vec<ParagraphBlock>,
    /// Whether Word selects the even header and footer variants.
    pub even_headers_active: bool,
    /// Default-header watermark, already positioned in page coordinates.
    pub watermark: Option<oxml_layout::GroupElement>,
    /// First-page-header watermark.
    pub first_watermark: Option<oxml_layout::GroupElement>,
    /// Even-page-header watermark.
    pub even_watermark: Option<oxml_layout::GroupElement>,
}

/// A section with its blocks, geometry, and header/footer content.
pub struct Section {
    pub blocks: Vec<LayoutBlock>,
    pub geometry: PageGeometry,
    pub header_footer: Option<HeaderFooterContent>,
    /// Whether this section uses a different first page header/footer.
    pub title_pg: bool,
    /// Displayed page number assigned to the first page of this section.
    pub page_number_start: Option<usize>,
}

/// Paginate across multiple sections, each with its own geometry and header/footer.
pub fn paginate_sections(
    sections: &[Section],
    fm: &FontManager,
    media: &MediaRegistry,
    notes: &NoteRegistry,
) -> (Vec<Arc<PageFrame>>, Vec<OutlineEntry>) {
    let media = media.media();
    if sections.is_empty() {
        return (
            vec![Arc::new(PageFrame::new(1, 612.0, 792.0, Vec::new()))],
            Vec::new(),
        );
    }

    // For a single section, delegate to the existing paginate function
    if sections.len() == 1 {
        let s = &sections[0];
        return paginate_with_media(
            &s.blocks,
            s.geometry,
            s.header_footer.as_ref(),
            s.title_pg,
            fm,
            media,
            notes,
            1,
            s.page_number_start.unwrap_or(1),
        );
    }

    // Multi-section pagination
    let mut all_pages = Vec::new();
    let mut all_outlines = Vec::new();
    let mut page_offset = 0;
    let mut next_section_page_number = 1usize;

    for section in sections {
        let section_page_number = section
            .page_number_start
            .unwrap_or(next_section_page_number);
        let (mut pages, mut outlines) = paginate_with_media(
            &section.blocks,
            section.geometry,
            section.header_footer.as_ref(),
            section.title_pg,
            fm,
            media,
            notes,
            page_offset + 1,
            section_page_number,
        );

        next_section_page_number = section_page_number.saturating_add(pages.len());
        page_offset += pages.len();
        all_pages.append(&mut pages);
        all_outlines.append(&mut outlines);
    }

    // If a section produced no pages (empty blocks), we might have duplicates
    // Renumber pages sequentially
    for (i, page) in all_pages.iter_mut().enumerate() {
        Arc::make_mut(page).page_number = i + 1;
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
    media: &MediaRegistry,
    notes: &NoteRegistry,
) -> (Vec<Arc<PageFrame>>, Vec<OutlineEntry>) {
    paginate_with_media(
        blocks,
        geometry,
        header_footer,
        title_pg,
        _fm,
        media.media(),
        notes,
        1,
        1,
    )
}

/// Where a pass placed the wrapping drawings whose vertical anchor is their own
/// paragraph, keyed by block index and the drawing's index within that block.
///
/// The key is stable across passes because both passes walk the same block list
/// in the same order.
type ResolvedWraps = HashMap<(usize, usize), (usize, PlacedWrap)>;

/// Whether any block anchors a wrapping drawing to its own paragraph or line.
///
/// A document without one paginates in a single pass, which is every sample and
/// every corpus document today.
fn has_paragraph_relative_wrap(blocks: &[LayoutBlock]) -> bool {
    blocks.iter().any(|block| {
        let LayoutBlock::Paragraph(para) = block else {
            return false;
        };
        para.anchored.iter().any(is_paragraph_relative_wrap)
    })
}

/// The filter both the look-ahead and the two-pass predicate agree on.
fn is_paragraph_relative_wrap(anchored: &AnchoredDrawing) -> bool {
    anchored.wrap != WrapType::None
        && matches!(
            anchored.rel_v,
            ST_RelativeFromV::Paragraph | ST_RelativeFromV::Line
        )
}

fn paginate_with_media(
    blocks: &[LayoutBlock],
    geometry: PageGeometry,
    header_footer: Option<&HeaderFooterContent>,
    title_pg: bool,
    _fm: &FontManager,
    media: &HashMap<MediaId, ImageData>,
    notes: &NoteRegistry,
    first_page_number: usize,
    first_header_page_number: usize,
) -> (Vec<Arc<PageFrame>>, Vec<OutlineEntry>) {
    let context = PassContext {
        geometry,
        header_footer,
        title_pg,
        fm: _fm,
        media,
        notes,
        first_page_number,
        first_header_page_number,
    };
    let first = paginate_pass(blocks, &context, &ResolvedWraps::new());

    // A paragraph-relative drawing has no vertical position until its own
    // paragraph is placed, so the first pass cannot offer one to the text above
    // it. The second pass can, because the first recorded where each landed.
    //
    // Two passes, not a fixed point. The second pass reflows earlier text, which
    // can move the drawing's own paragraph, so the rect it offered may be
    // slightly stale. Iterating is not guaranteed to terminate: growing a
    // paragraph can push a drawing to the next page, which shrinks the
    // paragraph, which pulls it back.
    if !has_paragraph_relative_wrap(blocks) {
        return (first.pages, first.outlines);
    }

    let second = paginate_pass(blocks, &context, &first.resolved);
    (second.pages, second.outlines)
}

/// One pagination pass, and what it learned about paragraph-relative wraps.
struct PassResult {
    pages: Vec<Arc<PageFrame>>,
    outlines: Vec<OutlineEntry>,
    resolved: ResolvedWraps,
}

/// Everything a pass needs that is the same for both passes.
///
/// Built once and borrowed twice, so the eight-value argument list appears in
/// one place rather than at each call.
struct PassContext<'a> {
    geometry: PageGeometry,
    header_footer: Option<&'a HeaderFooterContent>,
    title_pg: bool,
    fm: &'a FontManager,
    media: &'a HashMap<MediaId, ImageData>,
    notes: &'a NoteRegistry,
    first_page_number: usize,
    first_header_page_number: usize,
}

fn paginate_pass(
    blocks: &[LayoutBlock],
    context: &PassContext,
    resolved_in: &ResolvedWraps,
) -> PassResult {
    let geometry = context.geometry;
    let mut pager = Pager::new(
        geometry,
        context.header_footer,
        context.title_pg,
        context.media,
        context.notes,
        context.fm,
        resolved_in,
        context.first_page_number,
        context.first_header_page_number,
    );

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
                    if pager.cursor_y + row.height > pager.available_height() && pager.has_content()
                    {
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
                                    pager.page_number,
                                    tbl_borders,
                                    &mut pager.elements,
                                    pager.media,
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
                        pager.page_number,
                        tbl_borders,
                        &mut pager.elements,
                        pager.media,
                    );
                    pager.cursor_y += row.height;
                    pager.mark_content();
                }
            }
        }
    }

    let resolved = std::mem::take(&mut pager.resolved_out);
    let (pages, outlines) = pager.flush();
    PassResult {
        pages,
        outlines,
        resolved,
    }
}

/// Helper struct to track page state during pagination.
struct Pager<'a> {
    pages: Vec<Arc<PageFrame>>,
    elements: Vec<PositionedElement>,
    /// Anchored drawings marked behindDoc. Held apart from the normal element
    /// list so they can be emitted before everything else on the page, which
    /// is what puts them underneath the text.
    behind_elements: Vec<PositionedElement>,
    cursor_y: f64,
    page_number: usize,
    header_page_number: usize,
    content_height: f64,
    geometry: PageGeometry,
    header_footer: Option<&'a HeaderFooterContent>,
    has_content_flag: bool,
    outlines: Vec<OutlineEntry>,
    /// Whether the current page is the first page of the section.
    is_first_page: bool,
    /// Whether this section uses different first page header/footer.
    title_pg: bool,
    media: &'a HashMap<MediaId, ImageData>,
    /// Every note the document defines, laid out once before pagination.
    notes: &'a NoteRegistry,
    /// Notes first referenced by a line placed on the page being built, in
    /// reference order. Line counts are decided when the page is finished,
    /// since that is when the leftover height is known.
    page_note_ids: Vec<NoteRef>,
    /// Note content that did not fit on the previous page, as (id, next line).
    /// Placed before this page's own notes, and drawn without a marker.
    pending_notes: Vec<(NoteRef, usize)>,
    /// Re-breaking a paragraph around a drawing needs the shaper.
    fm: &'a FontManager,
    /// Rectangles of the wrapping drawings already placed on this page, with
    /// the wrap mode and text distances each one asks for.
    page_wraps: Vec<PlacedWrap>,
    /// Where the body's last mark sits, ignoring trailing paragraph spacing.
    ///
    /// `cursor_y` includes the space after the final paragraph, and that space
    /// collapses at a page break. Measuring the note area from `cursor_y`
    /// would let it eat into the height that was reserved, which is enough to
    /// push a note off the page its own reference sits on.
    ink_bottom: f64,
    /// Where the previous pass placed each paragraph-relative wrapping drawing.
    /// Empty on the first pass, which is what makes that pass identical to a
    /// single-pass run.
    resolved_in: &'a ResolvedWraps,
    /// Where this pass is placing them, for the pass that follows.
    resolved_out: ResolvedWraps,
}

impl<'a> Pager<'a> {
    fn new(
        geometry: PageGeometry,
        header_footer: Option<&'a HeaderFooterContent>,
        title_pg: bool,
        media: &'a HashMap<MediaId, ImageData>,
        notes: &'a NoteRegistry,
        fm: &'a FontManager,
        resolved_in: &'a ResolvedWraps,
        first_page_number: usize,
        first_header_page_number: usize,
    ) -> Self {
        Pager {
            pages: Vec::new(),
            elements: Vec::new(),
            behind_elements: Vec::new(),
            cursor_y: 0.0,
            page_number: first_page_number,
            header_page_number: first_header_page_number,
            content_height: geometry.content_height(),
            geometry,
            header_footer,
            has_content_flag: false,
            outlines: Vec::new(),
            is_first_page: true,
            title_pg,
            media,
            notes,
            page_note_ids: Vec::new(),
            pending_notes: Vec::new(),
            fm,
            page_wraps: Vec::new(),
            ink_bottom: 0.0,
            resolved_in,
            resolved_out: ResolvedWraps::new(),
        }
    }

    fn has_content(&self) -> bool {
        self.has_content_flag
    }

    /// Height the note area needs for a given set of notes, in full.
    ///
    /// Zero when there are none, so a page without notes keeps every point of
    /// its content height.
    fn reserve_for(&self, carried: &[(NoteRef, usize)], fresh: &[NoteRef]) -> f64 {
        if carried.is_empty() && fresh.is_empty() {
            return 0.0;
        }
        let carried_height: f64 = carried
            .iter()
            .filter_map(|(id, first)| {
                self.notes
                    .get(*id, self.geometry.content_width())
                    .map(|note| note.height_from(*first))
            })
            .sum();
        let fresh_height: f64 = fresh
            .iter()
            .filter_map(|id| {
                self.notes
                    .get(*id, self.geometry.content_width())
                    .map(NoteLayout::height)
            })
            .sum();
        NOTE_SEPARATOR_OFFSET + carried_height + fresh_height
    }

    /// The note area currently committed for the page being built.
    fn reserved_height(&self) -> f64 {
        self.reserve_for(&self.pending_notes, &self.page_note_ids)
    }

    /// Content height still usable by body text on this page.
    fn available_height(&self) -> f64 {
        (self.content_height - self.reserved_height()).max(0.0)
    }

    /// What the note area would cost if `lines` were placed on this page,
    /// without committing to placing them.
    ///
    /// A paragraph is measured before anyone knows which page it lands on, so
    /// its notes must be priced without being claimed. Claiming first and
    /// moving the paragraph afterwards leaves the note stranded on the page
    /// before its own reference.
    fn available_height_for(&self, lines: &[LayoutLine]) -> f64 {
        let mut fresh = self.page_note_ids.clone();
        for line in lines {
            for id in page_foot_notes_in_line(line) {
                if self.notes.get(id, self.geometry.content_width()).is_some()
                    && !fresh.contains(&id)
                    && !self.pending_notes.iter().any(|(pending, _)| *pending == id)
                {
                    fresh.push(id);
                }
            }
        }
        (self.content_height - self.reserve_for(&self.pending_notes, &fresh)).max(0.0)
    }

    /// Record the footnotes referenced by lines about to be placed.
    ///
    /// Endnotes are ignored here. They are emitted at the document end, so
    /// they cost the page carrying their reference nothing.
    fn claim_notes(&mut self, lines: &[LayoutLine]) {
        for id in lines.iter().flat_map(page_foot_notes_in_line) {
            {
                if self.notes.get(id, self.geometry.content_width()).is_some()
                    && !self.page_note_ids.contains(&id)
                    && !self.pending_notes.iter().any(|(pending, _)| *pending == id)
                {
                    self.page_note_ids.push(id);
                }
            }
        }
    }

    /// How many of `lines` fit, once the note area their references demand is
    /// taken out of the page.
    ///
    /// A line is admitted only if the whole note area still fits after it, so
    /// a note is not split merely because body text was greedy. The one
    /// exception is a page that has placed nothing yet: there the line goes
    /// down regardless and the note splits, because a page that admits neither
    /// body nor note makes no progress and pagination would not terminate.
    fn count_lines_that_fit_with_notes(&self, lines: &[LayoutLine], start_y: f64) -> usize {
        let mut fresh = self.page_note_ids.clone();
        let mut used = 0.0;

        for (index, line) in lines.iter().enumerate() {
            for id in page_foot_notes_in_line(line) {
                if self.notes.get(id, self.geometry.content_width()).is_some()
                    && !fresh.contains(&id)
                    && !self.pending_notes.iter().any(|(pending, _)| *pending == id)
                {
                    fresh.push(id);
                }
            }

            let reserve = self.reserve_for(&self.pending_notes, &fresh);
            if start_y + used + line.height > self.content_height - reserve + 0.01 {
                let page_is_empty = !self.has_content() && used == 0.0 && index == 0;
                if !page_is_empty {
                    return index;
                }
            }
            used += line.height;
        }

        lines.len()
    }

    fn mark_content(&mut self) {
        self.has_content_flag = true;
    }

    /// Resolve the wrapping drawings a paragraph carries, without placing
    /// them. Measuring a paragraph needs to know what it must flow around
    /// before anything is committed to the page.
    fn wrap_rects_for(
        &self,
        anchored: &[AnchoredDrawing],
        para_top: f64,
        indent_left: f64,
    ) -> Vec<PlacedWrap> {
        anchored
            .iter()
            .filter(|a| a.wrap != WrapType::None)
            .map(|a| PlacedWrap {
                rect: Rect {
                    x: resolve_anchor_h(
                        a.rel_h,
                        a.off_h,
                        a.align_h,
                        a.width,
                        &self.geometry,
                        indent_left,
                    ),
                    y: resolve_anchor_v(
                        a.rel_v,
                        a.off_v,
                        a.align_v,
                        a.height,
                        &self.geometry,
                        para_top,
                    ),
                    width: a.width,
                    height: a.height,
                },
                wrap: a.wrap,
                dist_top: a.dist_top,
                dist_bottom: a.dist_bottom,
                dist_left: a.dist_left,
                dist_right: a.dist_right,
            })
            .collect()
    }

    /// Wrapping drawings anchored to blocks after `block_idx`, positioned well
    /// enough to flow this block's text around them.
    ///
    /// A drawing anchored to a later paragraph still pushes earlier text aside,
    /// and Word documents do this routinely: the arrow beside a paragraph is
    /// often anchored to the paragraph after it. Where its position comes from
    /// depends on the frame it is measured against.
    ///
    /// A drawing framed by the page or a margin is positioned here, because
    /// that needs nothing from the block that owns it. A drawing framed by its
    /// own paragraph has no position until that paragraph is placed, so the
    /// first pass offers nothing for it and the second offers what the first
    /// recorded, for the drawings the first put on the page being built now.
    fn lookahead_wraps(&self, block_idx: usize, blocks: &[LayoutBlock]) -> Vec<PlacedWrap> {
        let mut out = Vec::new();
        let mut height = self.cursor_y;

        for (offset, block) in blocks.iter().enumerate().skip(block_idx + 1) {
            if block.page_break_before() || height > self.content_height {
                break;
            }
            height += block.space_before() + block.content_height() + block.space_after();

            let LayoutBlock::Paragraph(para) = block else {
                continue;
            };
            for (anchor_idx, a) in para.anchored.iter().enumerate() {
                if a.wrap == WrapType::None {
                    continue;
                }
                if is_paragraph_relative_wrap(a) {
                    // Resolved by the previous pass, or not at all. The page
                    // check is what stops a drawing that landed overleaf from
                    // pushing this page's text aside.
                    if let Some((page, placed)) = self.resolved_in.get(&(offset, anchor_idx))
                        && *page == self.page_number
                    {
                        out.push(*placed);
                    }
                    continue;
                }
                out.extend(self.wrap_rects_for(std::slice::from_ref(a), 0.0, para.indent_left));
            }
        }

        out
    }

    /// Place the drawings anchored to a paragraph whose top sits at `para_top`,
    /// measured from the top of the content area.
    ///
    /// `block_idx` identifies the owning block, so a paragraph-relative
    /// wrapping drawing can be recorded for the pass that follows this one.
    fn place_anchored(
        &mut self,
        anchored: &[AnchoredDrawing],
        para_top: f64,
        indent_left: f64,
        block_idx: usize,
    ) {
        for (anchor_idx, a) in anchored.iter().enumerate() {
            let x = resolve_anchor_h(
                a.rel_h,
                a.off_h,
                a.align_h,
                a.width,
                &self.geometry,
                indent_left,
            );
            let y = resolve_anchor_v(
                a.rel_v,
                a.off_v,
                a.align_v,
                a.height,
                &self.geometry,
                para_top,
            );
            let rect = Rect {
                x,
                y,
                width: a.width,
                height: a.height,
            };

            if a.wrap != WrapType::None {
                let placed = PlacedWrap {
                    rect,
                    wrap: a.wrap,
                    dist_top: a.dist_top,
                    dist_bottom: a.dist_bottom,
                    dist_left: a.dist_left,
                    dist_right: a.dist_right,
                };
                if is_paragraph_relative_wrap(a) {
                    self.resolved_out
                        .insert((block_idx, anchor_idx), (self.page_number, placed));
                }
                self.page_wraps.push(placed);
            }

            let mut produced: Vec<PositionedElement> = Vec::new();

            match &a.content {
                AnchoredContent::Image { media_id } => {
                    let image = self.media.get(media_id);
                    produced.push(PositionedElement::Image {
                        rect,
                        data: image.map_or_else(Vec::new, |image| image.data.clone()),
                        content_type: image
                            .map_or_else(String::new, |image| image.content_type.clone()),
                        media_id: *media_id,
                    });
                }
                AnchoredContent::Group(group) => {
                    let mut positioned = group.clone();
                    positioned.transform = positioned.transform.then(oxml_layout::Transform {
                        e: rect.x,
                        f: rect.y,
                        ..oxml_layout::Transform::IDENTITY
                    });
                    produced.push(PositionedElement::Group(positioned));
                }
                AnchoredContent::Shape { preset, fill, text } => {
                    // A shape with no fill draws no body. That is not a gap:
                    // Word uses unfilled rectangles as plain text boxes.
                    match (preset, fill) {
                        (ShapePreset::Rect, Some(color)) => {
                            produced.push(PositionedElement::FilledRect {
                                rect,
                                color: *color,
                            });
                        }
                        (ShapePreset::Line, Some(color)) => {
                            // A line shape's extent describes its bounding box,
                            // so the stroke runs corner to corner.
                            produced.push(PositionedElement::Line {
                                start: Point { x, y },
                                end: Point {
                                    x: x + a.width,
                                    y: y + a.height,
                                },
                                width: 1.0,
                                color: *color,
                                dash_pattern: None,
                            });
                        }
                        _ => {}
                    }
                    produced.extend(render_shape_text(text, &self.geometry, rect, self.media));
                }
            }

            if a.behind_doc {
                self.behind_elements.append(&mut produced);
            } else {
                self.elements.append(&mut produced);
            }
        }
    }

    /// Draw the note area for the page being built, and carry what did not
    /// fit onto the next one.
    ///
    /// Notes sit above the bottom margin and grow upward, so the body text
    /// above them was already kept clear by `available_height`.
    fn place_page_notes(&mut self) {
        let mut queue: Vec<(NoteRef, usize, bool)> = self
            .pending_notes
            .drain(..)
            .map(|(id, first)| (id, first, true))
            .collect();
        queue.extend(self.page_note_ids.drain(..).map(|id| (id, 0usize, false)));

        if queue.is_empty() {
            return;
        }

        let opens_with_continuation = queue[0].2;
        let available = (self.content_height - self.ink_bottom - NOTE_SEPARATOR_OFFSET).max(0.0);

        // Decide how much of each note this page can hold.
        let mut placed: Vec<(NoteRef, usize, usize, bool)> = Vec::new();
        let mut used = 0.0;
        let mut carried: Vec<(NoteRef, usize)> = Vec::new();

        for (id, first, continued) in queue {
            let Some(note) = self.notes.get(id, self.geometry.content_width()) else {
                continue;
            };
            if !carried.is_empty() {
                // An earlier note already ran out of room, so everything
                // after it waits too, or the notes would be reordered.
                carried.push((id, first));
                continue;
            }

            let mut count = 0;
            for line in note.lines.iter().skip(first) {
                if used + line.height > available + 0.01 {
                    break;
                }
                used += line.height;
                count += 1;
            }

            if count > 0 {
                placed.push((id, first, count, continued));
            }
            if first + count < note.lines.len() {
                carried.push((id, first + count));
            }
        }

        self.pending_notes = carried;

        if placed.is_empty() {
            return;
        }

        let total: f64 = placed
            .iter()
            .filter_map(|(id, first, count, _)| {
                self.notes
                    .get(*id, self.geometry.content_width())
                    .map(|n| n.height_of(*first, *count))
            })
            .sum();

        let separator_y =
            self.geometry.page_height - self.geometry.margin_bottom - total - NOTE_SEPARATOR_OFFSET;

        // A page opening with carried content gets the full-width rule, which
        // is how Word says "this continues from the previous page". A document
        // that never defined one keeps the short rule.
        let separator_width = if opens_with_continuation && self.notes.has_continuation_separator()
        {
            self.geometry.content_width()
        } else {
            self.geometry.content_width() * SEPARATOR_WIDTH_FRACTION
        };

        self.elements.push(PositionedElement::Line {
            start: Point {
                x: self.geometry.margin_left,
                y: separator_y,
            },
            end: Point {
                x: self.geometry.margin_left + separator_width,
                y: separator_y,
            },
            width: 0.5,
            color: Color::BLACK,
            dash_pattern: None,
        });

        let mut cursor_y = separator_y + NOTE_SEPARATOR_OFFSET;
        for (id, first, count, continued) in placed {
            let Some(note) = self.notes.get(id, self.geometry.content_width()) else {
                continue;
            };
            cursor_y += draw_note(
                &mut self.elements,
                &self.geometry,
                note,
                first,
                count,
                continued,
                cursor_y,
                self.page_number,
            );
        }
    }

    fn finish_page(&mut self) {
        self.place_page_notes();
        let mut all_elements = Vec::new();

        if let Some(hf) = self.header_footer {
            let watermark = if self.is_first_page && self.title_pg {
                hf.first_watermark.as_ref()
            } else if hf.even_headers_active && self.header_page_number.is_multiple_of(2) {
                hf.even_watermark.as_ref()
            } else {
                hf.watermark.as_ref()
            };
            if let Some(watermark) = watermark {
                all_elements.push(PositionedElement::Group(watermark.clone()));
            }
        }

        // behindDoc drawings render underneath everything else on the page.
        all_elements.append(&mut self.behind_elements);

        if let Some(hf) = self.header_footer {
            // Choose header blocks: first-page or default
            let header_blocks = if self.is_first_page && self.title_pg {
                &hf.first_header_blocks
            } else if hf.even_headers_active && self.header_page_number.is_multiple_of(2) {
                &hf.even_header_blocks
            } else {
                &hf.header_blocks
            };
            if !header_blocks.is_empty() {
                let header_y = self.geometry.header_distance;
                render_hf_blocks(
                    header_blocks,
                    &self.geometry,
                    header_y,
                    self.page_number,
                    &mut all_elements,
                    self.media,
                );
            }
        }

        all_elements.append(&mut self.elements);

        if let Some(hf) = self.header_footer {
            // Choose footer blocks: first-page or default
            let footer_blocks = if self.is_first_page && self.title_pg {
                &hf.first_footer_blocks
            } else if hf.even_headers_active && self.header_page_number.is_multiple_of(2) {
                &hf.even_footer_blocks
            } else {
                &hf.footer_blocks
            };
            if !footer_blocks.is_empty() {
                let footer_height: f64 = footer_blocks.iter().map(|b| b.content_height()).sum();
                let footer_y =
                    self.geometry.page_height - self.geometry.footer_distance - footer_height;
                render_hf_blocks(
                    footer_blocks,
                    &self.geometry,
                    footer_y,
                    self.page_number,
                    &mut all_elements,
                    self.media,
                );
            }
        }

        self.pages.push(Arc::new(PageFrame::new(
            self.page_number,
            self.geometry.page_width,
            self.geometry.page_height,
            all_elements,
        )));
        self.page_number += 1;
        self.header_page_number += 1;
        self.cursor_y = 0.0;
        self.page_wraps.clear();
        self.ink_bottom = 0.0;
        self.has_content_flag = false;
        self.is_first_page = false;
    }

    fn flush(mut self) -> (Vec<Arc<PageFrame>>, Vec<OutlineEntry>) {
        // Always create at least one page
        if self.has_content() || self.pages.is_empty() {
            self.finish_page();
        }
        // A note that ran past the last page of body text still has to land
        // somewhere, so keep making pages until the queue drains. Each page
        // places at least one note line, so this terminates.
        while !self.pending_notes.is_empty() {
            let before = self.pending_notes.clone();
            self.finish_page();
            if self.pending_notes == before {
                // Every page places at least one note line, so this is
                // unreachable. It exists so a future change that breaks that
                // guarantee stops rather than spins, and the assertion makes
                // it loud in tests instead of silently losing note text.
                debug_assert!(
                    false,
                    "a page placed no note content, dropping {:?}",
                    self.pending_notes
                );
                break;
            }
        }
        (self.pages, self.outlines)
    }
}

/// Draw one note, or one slice of one, with its top edge at `top`.
///
/// Returns the height consumed. Shared by the page foot and the document end
/// so the two regions cannot drift apart in how a note looks.
fn draw_note(
    elements: &mut Vec<PositionedElement>,
    geometry: &PageGeometry,
    note: &NoteLayout,
    first: usize,
    count: usize,
    continued: bool,
    top: f64,
    page_number: usize,
) -> f64 {
    let baseline = top + note.lines.get(first).map_or(0.0, |line| line.ascent);

    // A continuation does not repeat the marker.
    if !continued {
        elements.push(PositionedElement::Text(GlyphRun {
            origin: Point {
                x: geometry.margin_left,
                y: baseline - note.marker_rise,
            },
            font_id: note.marker.font_id,
            font_size: note.marker.font_size,
            glyph_ids: note.marker.glyph_ids.clone(),
            advances: note.marker.advances.clone(),
            text: note.marker.text.clone(),
            source: None,
            color: note.marker.color,
            bold: note.marker.bold,
            italic: note.marker.italic,
            field_kind: None,
            note: None,
        }));
    }

    let mut cursor_y = top;
    for line in note.lines.iter().skip(first).take(count) {
        let line_baseline = cursor_y + line.ascent;
        let mut x = geometry.margin_left + NOTE_INDENT;
        for item in &line.items {
            let (segment, advance) = match item {
                LineItem::Text(seg) | LineItem::Marker(seg) => (Some(seg), seg.width),
                LineItem::Tab { width, .. }
                | LineItem::Image { width, .. }
                | LineItem::Group { width, .. } => (None, *width),
                _ => (None, item.width()),
            };
            if let Some(seg) = segment {
                let adjusted_baseline = line_baseline - seg.baseline_offset;
                if let Some(color) = seg.highlight {
                    elements.push(PositionedElement::FilledRect {
                        rect: Rect {
                            x,
                            y: cursor_y,
                            width: seg.width,
                            height: line.height,
                        },
                        color,
                    });
                }
                elements.push(PositionedElement::Text(GlyphRun {
                    origin: Point {
                        x,
                        y: adjusted_baseline,
                    },
                    font_id: seg.font_id,
                    font_size: seg.font_size,
                    glyph_ids: seg.glyph_ids.clone(),
                    advances: seg.advances.clone(),
                    text: seg.text.clone(),
                    source: seg.source,
                    color: seg.color,
                    bold: seg.bold,
                    italic: seg.italic,
                    field_kind: None,
                    note: None,
                }));
                if let Some(underline) = seg.underline {
                    let underline_y = adjusted_baseline + seg.descent * 0.3;
                    let thickness = match underline {
                        Underline::Thick => seg.font_size / 12.0,
                        Underline::Double => seg.font_size / 24.0,
                        _ => seg.font_size / 18.0,
                    };
                    elements.push(PositionedElement::Line {
                        start: Point { x, y: underline_y },
                        end: Point {
                            x: x + seg.width,
                            y: underline_y,
                        },
                        width: thickness,
                        color: seg.color,
                        dash_pattern: None,
                    });
                    if underline == Underline::Double {
                        let second_y = underline_y + thickness * 2.5;
                        elements.push(PositionedElement::Line {
                            start: Point { x, y: second_y },
                            end: Point {
                                x: x + seg.width,
                                y: second_y,
                            },
                            width: thickness,
                            color: seg.color,
                            dash_pattern: None,
                        });
                    }
                }
                if seg.strike {
                    let strike_y = adjusted_baseline - seg.ascent * 0.3;
                    let thickness = seg.font_size / 24.0;
                    elements.push(PositionedElement::Line {
                        start: Point { x, y: strike_y },
                        end: Point {
                            x: x + seg.width,
                            y: strike_y,
                        },
                        width: thickness,
                        color: seg.color,
                        dash_pattern: None,
                    });
                }
                if seg.dstrike {
                    let strike_y = adjusted_baseline - seg.ascent * 0.3;
                    let thickness = seg.font_size / 24.0;
                    let gap = thickness * 2.0;
                    for y in [strike_y - gap / 2.0, strike_y + gap / 2.0] {
                        elements.push(PositionedElement::Line {
                            start: Point { x, y },
                            end: Point {
                                x: x + seg.width,
                                y,
                            },
                            width: thickness,
                            color: seg.color,
                            dash_pattern: None,
                        });
                    }
                }
            }
            x += advance;
        }
        cursor_y += line.height;
    }

    for range in &note.revision_ranges {
        let visible_start = range.start.max(first);
        let visible_end = range.end.min(first + count);
        if visible_start >= visible_end {
            continue;
        }
        let offset = note
            .lines
            .iter()
            .skip(first)
            .take(visible_start - first)
            .map(|line| line.height)
            .sum::<f64>();
        let height = note
            .lines
            .iter()
            .skip(visible_start)
            .take(visible_end - visible_start)
            .map(|line| line.height)
            .sum::<f64>();
        render_change_bar_at(top + offset, height, geometry, page_number, elements);
    }

    cursor_y - top
}

/// Append the document's endnotes as pages after the last body page.
///
/// Endnotes are flow content read at the end, not marginalia, so they start at
/// the top of a fresh page and carry no separator rule. There is no body text
/// on these pages for a rule to divide them from.
pub fn append_endnote_pages(
    pages: &mut Vec<Arc<PageFrame>>,
    notes: &NoteRegistry,
    geometry: PageGeometry,
) {
    // First-reference order across the document, which is the order a reader
    // met them in.
    let mut ordered: Vec<NoteRef> = Vec::new();
    for page in pages.iter() {
        for element in &page.elements {
            if let PositionedElement::Text(run) = element
                && let Some(note) = run.note
                && note.stream == NoteStream::Endnote
                && notes.get(note, geometry.content_width()).is_some()
                && !ordered.contains(&note)
            {
                ordered.push(note);
            }
        }
    }

    if ordered.is_empty() {
        return;
    }

    let content_height = geometry.content_height();
    let mut elements: Vec<PositionedElement> = Vec::new();
    let mut cursor_y = 0.0;
    let mut page_number = pages.len() + 1;

    let mut flush = |elements: &mut Vec<PositionedElement>, page_number: &mut usize| {
        pages.push(Arc::new(PageFrame::new(
            *page_number,
            geometry.page_width,
            geometry.page_height,
            std::mem::take(elements),
        )));
        *page_number += 1;
    };

    for note_ref in ordered {
        let Some(note) = notes.get(note_ref, geometry.content_width()) else {
            continue;
        };

        let mut first = 0;
        let mut continued = false;
        while first < note.lines.len() {
            let mut count = 0;
            let mut used = cursor_y;
            for line in note.lines.iter().skip(first) {
                if used + line.height > content_height + 0.01 {
                    break;
                }
                used += line.height;
                count += 1;
            }

            if count == 0 {
                // Nothing more fits on this page. Start a fresh one, unless
                // the page is already empty. An empty page that still cannot
                // take a line means the line is taller than the page, so it is
                // placed and allowed to overflow rather than looping forever.
                // Overflowing beats dropping the text, and body text on a page
                // of its own behaves the same way.
                if cursor_y == 0.0 {
                    count = 1;
                } else {
                    flush(&mut elements, &mut page_number);
                    cursor_y = 0.0;
                    continue;
                }
            }

            cursor_y += draw_note(
                &mut elements,
                &geometry,
                note,
                first,
                count,
                continued,
                geometry.margin_top + cursor_y,
                page_number,
            );
            first += count;
            continued = true;
        }
    }

    if !elements.is_empty() {
        flush(&mut elements, &mut page_number);
    }
}

/// The notes referenced by the segments on one line.
fn notes_in_line(line: &LayoutLine) -> impl Iterator<Item = NoteRef> + '_ {
    line.items.iter().filter_map(|item| match item {
        LineItem::Text(seg) | LineItem::Marker(seg) => seg.note,
        _ => None,
    })
}

/// The notes on one line that belong at the foot of its page.
///
/// Endnotes are excluded. They are emitted at the document end and take no
/// height from the page their reference sits on.
fn page_foot_notes_in_line(line: &LayoutLine) -> impl Iterator<Item = NoteRef> + '_ {
    notes_in_line(line).filter(|note| note.stream == NoteStream::Footnote)
}

/// Paginate a single paragraph, handling splitting across pages.
/// Shift a positioned element by a fixed offset.
///
/// Paragraph rendering always lays out against the page margins, so a text box
/// is rendered at the margin first and then moved to where the shape sits.
fn translate_element(element: &mut PositionedElement, dx: f64, dy: f64) {
    match element {
        PositionedElement::Text(run) => {
            run.origin.x += dx;
            run.origin.y += dy;
        }
        PositionedElement::Line { start, end, .. } => {
            start.x += dx;
            start.y += dy;
            end.x += dx;
            end.y += dy;
        }
        PositionedElement::FilledRect { rect, .. }
        | PositionedElement::Image { rect, .. }
        | PositionedElement::LinkAnnotation { rect, .. } => {
            rect.x += dx;
            rect.y += dy;
        }
        _ => {}
    }
}

/// Render a shape's text box inside `rect`.
///
/// The paragraphs arrive already laid out at the shape's width. They are
/// rendered as if they sat at the left margin and then translated onto the
/// shape, which keeps all the justification and indent handling in one place.
fn render_shape_text(
    text: &[ParagraphBlock],
    geometry: &PageGeometry,
    rect: Rect,
    media: &HashMap<MediaId, ImageData>,
) -> Vec<PositionedElement> {
    if text.is_empty() {
        return Vec::new();
    }

    let mut local = Vec::new();
    let mut y = 0.0;
    for para in text {
        render_paragraph_lines(&para.lines, para, geometry, y, &mut local, media);
        y += para.content_height();
    }

    // render_paragraph_lines works in content-area coordinates, so undo the
    // margin it applied and then move onto the shape.
    let dx = rect.x - geometry.margin_left;
    let dy = rect.y - geometry.margin_top;
    for element in &mut local {
        translate_element(element, dx, dy);
    }
    local
}

/// Resolve a horizontal anchor offset against the frame it is measured from.
///
/// An offset says nothing on its own. The same number lands somewhere
/// different depending on the frame, and treating every offset as a page
/// coordinate put anchored drawings in the corner of the sheet.
fn frame_h(rel: ST_RelativeFromH, g: &PageGeometry, indent_left: f64) -> (f64, f64) {
    let text_width = g.page_width - g.margin_left - g.margin_right;
    match rel {
        ST_RelativeFromH::Page | ST_RelativeFromH::LeftMargin => (0.0, g.page_width),
        ST_RelativeFromH::RightMargin | ST_RelativeFromH::OutsideMargin => {
            (g.page_width - g.margin_right, g.margin_right)
        }
        ST_RelativeFromH::InsideMargin => (g.margin_left, g.margin_left),
        // A character-relative offset starts where the text does on the line.
        ST_RelativeFromH::Character => (g.margin_left + indent_left, text_width),
        // Margin and column both start at the left edge of the text area.
        // Multiple columns are not laid out yet, so the two coincide.
        ST_RelativeFromH::Margin | ST_RelativeFromH::Column => (g.margin_left, text_width),
    }
}

fn resolve_anchor_h(
    rel: ST_RelativeFromH,
    off: f64,
    align: Option<AnchorAlignH>,
    width: f64,
    g: &PageGeometry,
    indent_left: f64,
) -> f64 {
    let (start, size) = frame_h(rel, g, indent_left);
    match align {
        // Inside and outside mean binding-side and outer-edge, which differ on
        // facing pages. Facing pages are not modelled, so the odd-page reading
        // stands in for both.
        Some(AnchorAlignH::Left | AnchorAlignH::Inside) => start,
        Some(AnchorAlignH::Center) => start + (size - width) / 2.0,
        Some(AnchorAlignH::Right | AnchorAlignH::Outside) => start + size - width,
        None => start + off,
    }
}

/// Resolve a vertical anchor offset against the frame it is measured from.
///
/// `para_top` is the top of the anchoring paragraph, measured from the top of
/// the content area.
fn frame_v(rel: ST_RelativeFromV, g: &PageGeometry, para_top: f64) -> (f64, f64) {
    let text_height = g.page_height - g.margin_top - g.margin_bottom;
    match rel {
        ST_RelativeFromV::Page | ST_RelativeFromV::TopMargin => (0.0, g.page_height),
        ST_RelativeFromV::BottomMargin | ST_RelativeFromV::OutsideMargin => {
            (g.page_height - g.margin_bottom, g.margin_bottom)
        }
        ST_RelativeFromV::Margin | ST_RelativeFromV::InsideMargin => (g.margin_top, text_height),
        // Paragraph and line are both relative to where this paragraph landed.
        // Per-line anchoring would need the line box, which is finer than we
        // track here, so the paragraph top stands in for both.
        ST_RelativeFromV::Paragraph | ST_RelativeFromV::Line => {
            (g.margin_top + para_top, text_height)
        }
    }
}

fn resolve_anchor_v(
    rel: ST_RelativeFromV,
    off: f64,
    align: Option<AnchorAlignV>,
    height: f64,
    g: &PageGeometry,
    para_top: f64,
) -> f64 {
    let (start, size) = frame_v(rel, g, para_top);
    match align {
        Some(AnchorAlignV::Top | AnchorAlignV::Inside) => start,
        Some(AnchorAlignV::Center) => start + (size - height) / 2.0,
        Some(AnchorAlignV::Bottom | AnchorAlignV::Outside) => start + size - height,
        None => start + off,
    }
}

/// Re-break a paragraph so its text flows around the wrapping drawings that
/// share its band of the page.
///
/// `para_top` is where the paragraph's content starts, measured from the top of
/// the content area. Returns `None` when nothing applies, so the caller keeps
/// the paragraph it already has.
fn reflow_around_wraps(
    para: &ParagraphBlock,
    wraps: &[PlacedWrap],
    para_top: f64,
    geometry: &PageGeometry,
    fm: &FontManager,
) -> Option<ParagraphBlock> {
    let reflow = para.reflow.as_ref()?;
    if wraps.is_empty() {
        return None;
    }

    let mut lines = para.lines.clone();
    let mut offset_top = 0.0;

    // Two passes. The first reserves against the paragraph as laid out, the
    // second against the heights the first produced, which is what settles a
    // drawing that only overlaps once the text has moved.
    for _ in 0..2 {
        let mut prefix: Vec<f64> = Vec::new();
        let mut suffix: Vec<f64> = Vec::new();

        // Vertical clearance is resolved first, because it moves the lines the
        // horizontal reservations are then measured against.
        let paragraph_top = geometry.margin_top + para_top;
        let mut next_offset_top: f64 = 0.0;
        for wrap in wraps.iter().filter(|w| w.wrap == WrapType::TopAndBottom) {
            if wrap.keep_out_top() <= paragraph_top + 1.0 && wrap.keep_out_bottom() > paragraph_top
            {
                next_offset_top = next_offset_top.max(wrap.keep_out_bottom() - paragraph_top);
            }
        }

        for wrap in wraps.iter().filter(|w| w.wrap != WrapType::TopAndBottom) {
            // Square, and the outline wraps approximated as square.
            let text_left = geometry.margin_left + para.indent_left;
            let text_right = geometry.page_width - geometry.margin_right - para.indent_right;
            let drawing_centre = wrap.rect.x + wrap.rect.width / 2.0;
            let on_the_left = drawing_centre < (text_left + text_right) / 2.0;

            let reserve = if on_the_left {
                (wrap.rect.x + wrap.rect.width + wrap.dist_right - text_left).max(0.0)
            } else {
                (text_right - (wrap.rect.x - wrap.dist_left)).max(0.0)
            };
            if reserve <= 0.0 {
                continue;
            }

            let mut line_top = geometry.margin_top + para_top + next_offset_top;
            for (index, line) in lines.iter().enumerate() {
                let line_bottom = line_top + line.height;
                if line_bottom > wrap.keep_out_top() && line_top < wrap.keep_out_bottom() {
                    let target = if on_the_left {
                        &mut prefix
                    } else {
                        &mut suffix
                    };
                    if target.len() <= index {
                        target.resize(index + 1, 0.0);
                    }
                    target[index] += reserve;
                }
                line_top = line_bottom;
            }
        }

        if prefix.is_empty() && suffix.is_empty() && next_offset_top == 0.0 {
            return None;
        }

        let mut params = reflow.params.clone();
        params.line_prefix_widths = prefix;
        params.line_suffix_widths = suffix;

        let Ok(reflowed) = break_into_lines(&reflow.items, &params, fm) else {
            return None;
        };
        lines = reflowed;
        offset_top = next_offset_top;
    }

    let mut adjusted = para.clone();
    adjusted.lines = lines;
    adjusted.content_offset_top = offset_top;
    // The paragraph has been reflowed. Re-entering would reserve twice.
    adjusted.reflow = None;
    Some(adjusted)
}

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

    // Flow the paragraph around anything floating in its band of the page,
    // before anything is measured. A reflow changes the paragraph's height, so
    // doing it after the fitting decision would measure the wrong thing.
    let reflowed = {
        let para_top = pager.cursor_y + space_before;
        let mut wraps = pager.page_wraps.clone();
        wraps.extend(pager.wrap_rects_for(&para.anchored, para_top, para.indent_left));
        wraps.extend(pager.lookahead_wraps(block_idx, blocks));
        reflow_around_wraps(para, &wraps, para_top, &pager.geometry, pager.fm)
    };
    let para = reflowed.as_ref().unwrap_or(para);

    // Check if paragraph fits on current page. The note area its references
    // will demand is priced in, but not claimed: the paragraph may yet move to
    // the next page, and its notes must move with it.
    let total_needed = space_before + para.content_height();
    let remaining = pager.available_height_for(&para.lines) - pager.cursor_y;

    if total_needed > remaining && pager.has_content() {
        // Paragraph doesn't fit. Decide: move whole or split.
        if para.keep_lines || para.lines.len() <= 2 {
            pager.finish_page();
            // Re-call with fresh page
            paginate_paragraph(para, block_idx, blocks, pager);
            return;
        }

        // The lines start below any drawing the paragraph must clear, so the
        // counter has to be told where they actually begin.
        let lines_that_fit = pager.count_lines_that_fit_with_notes(
            &para.lines,
            pager.cursor_y + space_before + para.content_offset_top,
        );

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
            render_para_split(para, split_at, space_before, pager, block_idx);
            return;
        }

        if lines_that_fit > 0 {
            render_para_split(para, lines_that_fit, space_before, pager, block_idx);
            return;
        }

        // No lines fit (shouldn't happen since we checked has_content above)
        pager.finish_page();
        paginate_paragraph(para, block_idx, blocks, pager);
        return;
    }

    // Paragraph fits OR we're at the top of a page
    // If it doesn't fit and we're at the top, we must split line by line
    if total_needed > pager.available_height_for(&para.lines) && pager.cursor_y == 0.0 {
        // Paragraph is taller than a page; split line by line
        let lines_that_fit =
            pager.count_lines_that_fit_with_notes(&para.lines, para.content_offset_top);
        if lines_that_fit > 0 && lines_that_fit < para.lines.len() {
            render_para_split(para, lines_that_fit, 0.0, pager, block_idx);
            return;
        }
    }

    // Check keep-with-next
    if para.keep_next && block_idx + 1 < blocks.len() {
        let next_first = match &blocks[block_idx + 1] {
            LayoutBlock::Paragraph(p) => p.lines.first().map(|l| l.height).unwrap_or(0.0),
            LayoutBlock::Table(t) => t.rows.first().map(|r| r.height).unwrap_or(0.0),
        };
        if pager.cursor_y + space_before + para.content_height() + next_first
            > pager.available_height_for(&para.lines)
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
        render_border_edges(
            borders,
            border_x,
            border_y,
            border_w,
            border_h,
            &mut pager.elements,
        );
    }

    // Anchored drawings resolve against the paragraph's position, so place
    // them now that the page and the cursor are settled.
    pager.place_anchored(&para.anchored, pager.cursor_y, para.indent_left, block_idx);

    render_paragraph_lines(
        &para.lines,
        para,
        &pager.geometry,
        pager.cursor_y,
        &mut pager.elements,
        pager.media,
    );
    render_change_bar(
        para,
        pager.cursor_y,
        para.content_height(),
        &pager.geometry,
        pager.page_number,
        &mut pager.elements,
    );
    pager.claim_notes(&para.lines);
    pager.cursor_y += para.content_height();
    pager.ink_bottom = pager.cursor_y;
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
    block_idx: usize,
) {
    // Render lines before split on current page
    pager.cursor_y += space_before;
    // A split paragraph anchors its drawings to where it starts.
    pager.place_anchored(&para.anchored, pager.cursor_y, para.indent_left, block_idx);
    render_paragraph_lines(
        &para.lines[..split_at],
        para,
        &pager.geometry,
        pager.cursor_y,
        &mut pager.elements,
        pager.media,
    );
    let first_height = para.content_offset_top
        + para.lines[..split_at]
            .iter()
            .map(|line| line.height)
            .sum::<f64>();
    render_change_bar(
        para,
        pager.cursor_y,
        first_height,
        &pager.geometry,
        pager.page_number,
        &mut pager.elements,
    );
    // Only the lines placed on this page count toward its notes. The rest of
    // the paragraph, and any note it references, belong to the next page.
    pager.claim_notes(&para.lines[..split_at]);
    pager.ink_bottom = pager.cursor_y
        + para.content_offset_top
        + para.lines[..split_at].iter().map(|l| l.height).sum::<f64>();
    pager.mark_content();
    pager.finish_page();

    // Handle remaining lines, which may themselves need splitting
    let remaining_lines = &para.lines[split_at..];
    let remaining_height: f64 = remaining_lines.iter().map(|l| l.height).sum();

    if remaining_height > pager.available_height_for(remaining_lines) {
        // Still too tall — split again
        let lines_that_fit = pager.count_lines_that_fit_with_notes(remaining_lines, 0.0);
        if lines_that_fit > 0 && lines_that_fit < remaining_lines.len() {
            // Build a temporary para with remaining lines
            let temp_para = ParagraphBlock {
                // The anchors were placed with the first part of the
                // paragraph, so the continuation must not place them again.
                anchored: Vec::new(),
                has_visible_revision: para.has_visible_revision,
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
                // The continuation was already reflowed as part of the whole
                // paragraph, so it must not be reflowed again.
                reflow: None,
                content_offset_top: 0.0,
            };
            render_para_split(&temp_para, lines_that_fit, 0.0, pager, block_idx);
            return;
        }
    }

    // Remaining fits on the new page
    render_paragraph_lines(
        remaining_lines,
        para,
        &pager.geometry,
        0.0,
        &mut pager.elements,
        pager.media,
    );
    render_change_bar(
        para,
        0.0,
        remaining_height,
        &pager.geometry,
        pager.page_number,
        &mut pager.elements,
    );
    pager.claim_notes(remaining_lines);
    pager.ink_bottom = remaining_height;
    pager.cursor_y = remaining_height + para.space_after;
    pager.mark_content();
}

/// Render paragraph lines as positioned elements.
fn render_paragraph_lines(
    lines: &[LayoutLine],
    para: &ParagraphBlock,
    geometry: &PageGeometry,
    start_y: f64,
    elements: &mut Vec<PositionedElement>,
    media: &HashMap<MediaId, ImageData>,
) {
    // A drawing this paragraph must clear rather than flow beside pushes its
    // first line down. `content_height` already counts the same offset.
    let mut y = start_y + para.content_offset_top;
    for line in lines {
        let baseline_y = geometry.margin_top + y + line.ascent;

        // Compute x offset based on justification
        let text_width: f64 = line.items.iter().map(|item| item.width()).sum();
        let remaining_width = line.available_width - text_width;

        // For justified text (Both), compute extra space per gap
        let justify_extra =
            if para.jc == Some(Align::Justify) && !line.is_last && remaining_width > 0.0 {
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
            Some(Align::Center) => geometry.margin_left + line.indent_left + remaining_width / 2.0,
            Some(Align::End) => geometry.margin_left + line.indent_left + remaining_width,
            Some(Align::Justify) if !line.is_last && justify_extra > 0.0 => {
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
                        origin: Point {
                            x,
                            y: adjusted_baseline,
                        },
                        font_id: seg.font_id,
                        font_size: seg.font_size,
                        glyph_ids: seg.glyph_ids.clone(),
                        advances,
                        text: seg.text.clone(),
                        source: seg.source,
                        color: seg.color,
                        bold: seg.bold,
                        italic: seg.italic,
                        field_kind: seg.field_kind,
                        note: seg.note,
                    }));

                    // Render underline
                    if let Some(ul_style) = seg.underline {
                        let ul_y = adjusted_baseline + seg.descent * 0.3;
                        let ul_thickness = match ul_style {
                            Underline::Thick => seg.font_size / 12.0,
                            Underline::Double => seg.font_size / 24.0,
                            _ => seg.font_size / 18.0,
                        };
                        elements.push(PositionedElement::Line {
                            start: Point { x, y: ul_y },
                            end: Point {
                                x: x + effective_width,
                                y: ul_y,
                            },
                            width: ul_thickness,
                            color: seg.color,
                            dash_pattern: None,
                        });
                        // Second line for double underline
                        if ul_style == Underline::Double {
                            let ul_y2 = ul_y + ul_thickness * 2.5;
                            elements.push(PositionedElement::Line {
                                start: Point { x, y: ul_y2 },
                                end: Point {
                                    x: x + effective_width,
                                    y: ul_y2,
                                },
                                width: ul_thickness,
                                color: seg.color,
                                dash_pattern: None,
                            });
                        }
                    }

                    // Render strikethrough
                    if seg.strike {
                        let strike_y = adjusted_baseline - seg.ascent * 0.3;
                        let strike_thickness = seg.font_size / 24.0;
                        elements.push(PositionedElement::Line {
                            start: Point { x, y: strike_y },
                            end: Point {
                                x: x + effective_width,
                                y: strike_y,
                            },
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
                            start: Point {
                                x,
                                y: strike_y - gap / 2.0,
                            },
                            end: Point {
                                x: x + effective_width,
                                y: strike_y - gap / 2.0,
                            },
                            width: strike_thickness,
                            color: seg.color,
                            dash_pattern: None,
                        });
                        elements.push(PositionedElement::Line {
                            start: Point {
                                x,
                                y: strike_y + gap / 2.0,
                            },
                            end: Point {
                                x: x + effective_width,
                                y: strike_y + gap / 2.0,
                            },
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
                            source: None,
                            color: leader_seg.color,
                            bold: leader_seg.bold,
                            italic: leader_seg.italic,
                            field_kind: None,
                            note: None,
                        }));
                    }
                    x += width;
                }
                LineItem::Image {
                    width,
                    height,
                    media_id,
                } => {
                    let image = media.get(media_id);
                    // Image positioned at current x, top-aligned with line
                    elements.push(PositionedElement::Image {
                        rect: Rect {
                            x,
                            y: geometry.margin_top + y,
                            width: *width,
                            height: *height,
                        },
                        data: image.map_or_else(Vec::new, |image| image.data.clone()),
                        content_type: image
                            .map_or_else(String::new, |image| image.content_type.clone()),
                        media_id: *media_id,
                    });
                    x += width;
                }
                LineItem::Group { width, group, .. } => {
                    let mut positioned = group.clone();
                    positioned.transform = positioned.transform.then(oxml_layout::Transform {
                        e: x,
                        f: geometry.margin_top + y,
                        ..oxml_layout::Transform::IDENTITY
                    });
                    elements.push(PositionedElement::Group(positioned));
                    x += width;
                }
                _ => x += item.width(),
            }
        }

        y += line.height;
    }
}

fn render_change_bar(
    para: &ParagraphBlock,
    start_y: f64,
    height: f64,
    geometry: &PageGeometry,
    page_number: usize,
    elements: &mut Vec<PositionedElement>,
) {
    if !para.has_visible_revision || !height.is_finite() || height <= 0.0 {
        return;
    }
    render_change_bar_at(
        geometry.margin_top + start_y,
        height,
        geometry,
        page_number,
        elements,
    );
}

fn render_change_bar_at(
    start_y: f64,
    height: f64,
    geometry: &PageGeometry,
    page_number: usize,
    elements: &mut Vec<PositionedElement>,
) {
    if !height.is_finite() || height <= 0.0 {
        return;
    }
    let x = if page_number.is_multiple_of(2) {
        geometry.margin_left / 2.0
    } else {
        geometry.page_width - geometry.margin_right / 2.0
    };
    if !x.is_finite() || !start_y.is_finite() || !(start_y + height).is_finite() {
        return;
    }
    elements.push(PositionedElement::Line {
        start: Point { x, y: start_y },
        end: Point {
            x,
            y: start_y + height,
        },
        width: 1.5,
        color: Color::BLACK,
        dash_pattern: None,
    });
}

/// Render header/footer blocks.
fn render_hf_blocks(
    blocks: &[ParagraphBlock],
    geometry: &PageGeometry,
    start_y: f64,
    page_number: usize,
    elements: &mut Vec<PositionedElement>,
    media: &HashMap<MediaId, ImageData>,
) {
    let mut y = start_y - geometry.margin_top; // Convert to relative
    for para in blocks {
        render_paragraph_lines(&para.lines, para, geometry, y, elements, media);
        render_change_bar(
            para,
            y,
            para.content_height(),
            geometry,
            page_number,
            elements,
        );
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
    page_number: usize,
    table_borders: Option<&rdocx_oxml::table::CT_TblBorders>,
    elements: &mut Vec<PositionedElement>,
    media: &HashMap<MediaId, ImageData>,
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
                    media,
                );
                render_change_bar(
                    para,
                    para_y,
                    para.content_height(),
                    geometry,
                    page_number,
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
            let (nx, ny) = if len > 0.0 {
                (-dy / len, dx / len)
            } else {
                (0.0, 1.0)
            };
            let offset = gap / 2.0;
            elements.push(PositionedElement::Line {
                start: Point {
                    x: start.x + nx * offset,
                    y: start.y + ny * offset,
                },
                end: Point {
                    x: end.x + nx * offset,
                    y: end.y + ny * offset,
                },
                width: thickness,
                color,
                dash_pattern: None,
            });
            elements.push(PositionedElement::Line {
                start: Point {
                    x: start.x - nx * offset,
                    y: start.y - ny * offset,
                },
                end: Point {
                    x: end.x - nx * offset,
                    y: end.y - ny * offset,
                },
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
            Point {
                x: x + w,
                y: y - space,
            },
            elements,
        );
    }
    if let Some(ref edge) = borders.bottom {
        let space = edge.space.unwrap_or(0) as f64;
        render_edge(
            edge,
            Point {
                x,
                y: y + h + space,
            },
            Point {
                x: x + w,
                y: y + h + space,
            },
            elements,
        );
    }
    if let Some(ref edge) = borders.left {
        let space = edge.space.unwrap_or(0) as f64;
        render_edge(
            edge,
            Point { x: x - space, y },
            Point {
                x: x - space,
                y: y + h,
            },
            elements,
        );
    }
    if let Some(ref edge) = borders.right {
        let space = edge.space.unwrap_or(0) as f64;
        render_edge(
            edge,
            Point {
                x: x + w + space,
                y,
            },
            Point {
                x: x + w + space,
                y: y + h,
            },
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
    use oxml_layout::LayoutLine;

    fn empty_media() -> MediaRegistry {
        MediaRegistry::new(&HashMap::new())
    }

    fn make_line(height: f64) -> LayoutLine {
        LayoutLine {
            items: vec![],
            width: 100.0,
            ascent: height * 0.77,
            descent: height * 0.23,
            line_gap: 0.0,
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
            anchored: Vec::new(),
            has_visible_revision: false,
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
            reflow: None,
            content_offset_top: 0.0,
        }
    }

    #[test]
    fn single_page_layout() {
        let fm = FontManager::new();
        let blocks = vec![LayoutBlock::Paragraph(make_para(3, 14.0))];
        let geom = PageGeometry::default();
        let (pages, _outlines) = paginate(
            &blocks,
            geom,
            None,
            false,
            &fm,
            &empty_media(),
            &NoteRegistry::default(),
        );
        assert_eq!(pages.len(), 1);
        assert_eq!(pages[0].page_number, 1);
    }

    #[test]
    fn multi_page_overflow() {
        let fm = FontManager::new();
        // 648pt content height / 14pt per line ≈ 46 lines per page
        let blocks = vec![LayoutBlock::Paragraph(make_para(100, 14.0))];
        let geom = PageGeometry::default();
        let (pages, _outlines) = paginate(
            &blocks,
            geom,
            None,
            false,
            &fm,
            &empty_media(),
            &NoteRegistry::default(),
        );
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
        let (pages, _outlines) = paginate(
            &blocks,
            geom,
            None,
            false,
            &fm,
            &empty_media(),
            &NoteRegistry::default(),
        );
        assert_eq!(pages.len(), 2);
    }

    #[test]
    fn page_dimensions() {
        let fm = FontManager::new();
        let blocks = vec![LayoutBlock::Paragraph(make_para(1, 14.0))];
        let geom = PageGeometry::default();
        let (pages, _outlines) = paginate(
            &blocks,
            geom,
            None,
            false,
            &fm,
            &empty_media(),
            &NoteRegistry::default(),
        );
        assert!((pages[0].width - 612.0).abs() < 0.01);
        assert!((pages[0].height - 792.0).abs() < 0.01);
    }

    fn make_text_line(height: f64, underline: Option<Underline>, strike: bool) -> LayoutLine {
        use oxml_layout::TextSegment;
        let seg = TextSegment {
            text: "Hello".to_string(),
            source: None,
            font_id: oxml_layout::FontId(0),
            font_size: 12.0,
            glyph_ids: vec![1, 2, 3],
            advances: vec![6.0, 6.0, 6.0],
            width: 40.0,
            ascent: height * 0.77,
            descent: height * 0.23,
            line_gap: 0.0,
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
            note: None,
        };
        LayoutLine {
            items: vec![LineItem::Text(seg)],
            width: 40.0,
            ascent: height * 0.77,
            descent: height * 0.23,
            line_gap: 0.0,
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
            anchored: Vec::new(),
            has_visible_revision: false,
            lines: vec![make_text_line(14.0, Some(Underline::Single), false)],
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
            reflow: None,
            content_offset_top: 0.0,
        };
        let blocks = vec![LayoutBlock::Paragraph(para)];
        let (pages, _outlines) = paginate(
            &blocks,
            PageGeometry::default(),
            None,
            false,
            &fm,
            &empty_media(),
            &NoteRegistry::default(),
        );
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
            anchored: Vec::new(),
            has_visible_revision: false,
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
            reflow: None,
            content_offset_top: 0.0,
        };
        let blocks = vec![LayoutBlock::Paragraph(para)];
        let (pages, _outlines) = paginate(
            &blocks,
            PageGeometry::default(),
            None,
            false,
            &fm,
            &empty_media(),
            &NoteRegistry::default(),
        );
        let lines: Vec<_> = pages[0]
            .elements
            .iter()
            .filter(|e| matches!(e, PositionedElement::Line { .. }))
            .collect();
        assert_eq!(lines.len(), 1, "expected 1 strikethrough line");
    }

    #[test]
    fn highlight_renders_filled_rect() {
        use oxml_layout::TextSegment;
        let fm = FontManager::new();
        let seg = TextSegment {
            text: "Hi".to_string(),
            source: None,
            font_id: oxml_layout::FontId(0),
            font_size: 12.0,
            glyph_ids: vec![1],
            advances: vec![10.0],
            width: 20.0,
            ascent: 10.0,
            descent: 3.0,
            line_gap: 0.0,
            color: Color::BLACK,
            bold: false,
            italic: false,
            underline: None,
            strike: false,
            dstrike: false,
            highlight: Some(Color {
                r: 1.0,
                g: 1.0,
                b: 0.0,
                a: 1.0,
            }),
            baseline_offset: 0.0,
            hyperlink_url: None,
            field_kind: None,
            note: None,
        };
        let line = LayoutLine {
            items: vec![LineItem::Text(seg)],
            width: 20.0,
            ascent: 10.0,
            descent: 3.0,
            line_gap: 0.0,
            height: 13.0,
            indent_left: 0.0,
            available_width: 468.0,
            is_last: true,
        };
        let para = ParagraphBlock {
            anchored: Vec::new(),
            has_visible_revision: false,
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
            reflow: None,
            content_offset_top: 0.0,
        };
        let blocks = vec![LayoutBlock::Paragraph(para)];
        let (pages, _outlines) = paginate(
            &blocks,
            PageGeometry::default(),
            None,
            false,
            &fm,
            &empty_media(),
            &NoteRegistry::default(),
        );
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
            anchored: Vec::new(),
            has_visible_revision: false,
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
            reflow: None,
            content_offset_top: 0.0,
        };
        let blocks = vec![LayoutBlock::Paragraph(para)];
        let (pages, _outlines) = paginate(
            &blocks,
            PageGeometry::default(),
            None,
            false,
            &fm,
            &empty_media(),
            &NoteRegistry::default(),
        );
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
            anchored: Vec::new(),
            has_visible_revision: false,
            lines: vec![make_line(14.0)],
            space_before: 0.0,
            space_after: 0.0,
            borders: None,
            shading: Some(Color {
                r: 1.0,
                g: 1.0,
                b: 0.0,
                a: 1.0,
            }),
            indent_left: 0.0,
            indent_right: 0.0,
            jc: None,
            keep_next: false,
            keep_lines: false,
            page_break_before: false,
            widow_control: true,
            heading_level: None,
            heading_text: None,
            reflow: None,
            content_offset_top: 0.0,
        };
        let blocks = vec![LayoutBlock::Paragraph(para)];
        let (pages, _outlines) = paginate(
            &blocks,
            PageGeometry::default(),
            None,
            false,
            &fm,
            &empty_media(),
            &NoteRegistry::default(),
        );
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
            anchored: Vec::new(),
            has_visible_revision: false,
            lines: vec![make_text_line(14.0, Some(Underline::Double), false)],
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
            reflow: None,
            content_offset_top: 0.0,
        };
        let blocks = vec![LayoutBlock::Paragraph(para)];
        let (pages, _outlines) = paginate(
            &blocks,
            PageGeometry::default(),
            None,
            false,
            &fm,
            &empty_media(),
            &NoteRegistry::default(),
        );
        let lines: Vec<_> = pages[0]
            .elements
            .iter()
            .filter(|e| matches!(e, PositionedElement::Line { .. }))
            .collect();
        assert_eq!(lines.len(), 2, "expected 2 lines for double underline");
    }

    fn make_justified_line(text: &str, seg_width: f64, is_last: bool) -> LayoutLine {
        use oxml_layout::TextSegment;
        let seg = TextSegment {
            text: text.to_string(),
            source: None,
            font_id: oxml_layout::FontId(0),
            font_size: 12.0,
            glyph_ids: vec![1; text.len()],
            advances: vec![seg_width / text.len() as f64; text.len()],
            width: seg_width,
            ascent: 10.0,
            descent: 3.0,
            line_gap: 0.0,
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
            note: None,
        };
        LayoutLine {
            items: vec![LineItem::Text(seg)],
            width: seg_width,
            ascent: 10.0,
            descent: 3.0,
            line_gap: 0.0,
            height: 13.0,
            indent_left: 0.0,
            available_width: 468.0,
            is_last,
        }
    }

    #[test]
    fn hyperlink_emits_link_annotation() {
        use oxml_layout::TextSegment;
        let fm = FontManager::new();
        let seg = TextSegment {
            text: "Click me".to_string(),
            source: None,
            font_id: oxml_layout::FontId(0),
            font_size: 12.0,
            glyph_ids: vec![1, 2, 3],
            advances: vec![8.0, 8.0, 8.0],
            width: 60.0,
            ascent: 10.0,
            descent: 3.0,
            line_gap: 0.0,
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
            note: None,
        };
        let line = LayoutLine {
            items: vec![LineItem::Text(seg)],
            width: 60.0,
            ascent: 10.0,
            descent: 3.0,
            line_gap: 0.0,
            height: 13.0,
            indent_left: 0.0,
            available_width: 468.0,
            is_last: true,
        };
        let para = ParagraphBlock {
            anchored: Vec::new(),
            has_visible_revision: false,
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
            reflow: None,
            content_offset_top: 0.0,
        };
        let blocks = vec![LayoutBlock::Paragraph(para)];
        let (pages, _outlines) = paginate(
            &blocks,
            PageGeometry::default(),
            None,
            false,
            &fm,
            &empty_media(),
            &NoteRegistry::default(),
        );
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
            anchored: Vec::new(),
            has_visible_revision: false,
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
            jc: Some(Align::Justify),
            keep_next: false,
            keep_lines: false,
            page_break_before: false,
            widow_control: true,
            heading_level: None,
            heading_text: None,
            reflow: None,
            content_offset_top: 0.0,
        };

        let blocks = vec![LayoutBlock::Paragraph(para)];
        let (pages, _outlines) = paginate(
            &blocks,
            PageGeometry::default(),
            None,
            false,
            &fm,
            &empty_media(),
            &NoteRegistry::default(),
        );

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
            anchored: Vec::new(),
            has_visible_revision: false,
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
            jc: Some(Align::Justify),
            keep_next: false,
            keep_lines: false,
            page_break_before: false,
            widow_control: true,
            heading_level: None,
            heading_text: None,
            reflow: None,
            content_offset_top: 0.0,
        };

        let blocks = vec![LayoutBlock::Paragraph(para)];
        let (pages, _outlines) = paginate(
            &blocks,
            PageGeometry::default(),
            None,
            false,
            &fm,
            &empty_media(),
            &NoteRegistry::default(),
        );

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
            anchored: Vec::new(),
            has_visible_revision: false,
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
            jc: Some(Align::Justify),
            keep_next: false,
            keep_lines: false,
            page_break_before: false,
            widow_control: true,
            heading_level: None,
            heading_text: None,
            reflow: None,
            content_offset_top: 0.0,
        };

        let blocks = vec![LayoutBlock::Paragraph(para)];
        let (pages, _outlines) = paginate(
            &blocks,
            PageGeometry::default(),
            None,
            false,
            &fm,
            &empty_media(),
            &NoteRegistry::default(),
        );

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

    /// A wp:anchor offset means nothing without the frame it is measured from.
    /// Treating every offset as a page coordinate put anchored drawings in the
    /// corner of the sheet instead of beside their paragraph.
    #[test]
    fn anchor_offsets_resolve_against_their_frame() {
        let g = PageGeometry::default(); // 612 x 792, 72pt margins
        let para_top = 100.0;
        let off = 10.0;

        assert_eq!(
            resolve_anchor_h(ST_RelativeFromH::Page, off, None, 0.0, &g, 0.0),
            10.0
        );
        assert_eq!(
            resolve_anchor_h(ST_RelativeFromH::LeftMargin, off, None, 0.0, &g, 0.0),
            10.0
        );
        assert_eq!(
            resolve_anchor_h(ST_RelativeFromH::Margin, off, None, 0.0, &g, 0.0),
            82.0,
            "margin-relative starts at the left margin"
        );
        assert_eq!(
            resolve_anchor_h(ST_RelativeFromH::Column, off, None, 0.0, &g, 0.0),
            82.0,
            "column-relative starts at the text area"
        );
        assert_eq!(
            resolve_anchor_h(ST_RelativeFromH::RightMargin, off, None, 0.0, &g, 0.0),
            550.0,
            "right-margin-relative starts at the right margin edge"
        );
        assert_eq!(
            resolve_anchor_h(ST_RelativeFromH::Character, off, None, 0.0, &g, 36.0),
            118.0,
            "character-relative includes the paragraph indent"
        );

        assert_eq!(
            resolve_anchor_v(ST_RelativeFromV::Page, off, None, 0.0, &g, para_top),
            10.0
        );
        assert_eq!(
            resolve_anchor_v(ST_RelativeFromV::TopMargin, off, None, 0.0, &g, para_top),
            10.0
        );
        assert_eq!(
            resolve_anchor_v(ST_RelativeFromV::Margin, off, None, 0.0, &g, para_top),
            82.0
        );
        assert_eq!(
            resolve_anchor_v(ST_RelativeFromV::Paragraph, off, None, 0.0, &g, para_top),
            182.0,
            "paragraph-relative follows the paragraph down the page"
        );
        assert_eq!(
            resolve_anchor_v(ST_RelativeFromV::Line, off, None, 0.0, &g, para_top),
            182.0
        );
        assert_eq!(
            resolve_anchor_v(ST_RelativeFromV::BottomMargin, off, None, 0.0, &g, para_top),
            730.0
        );
    }

    /// The same offset must land somewhere different once the paragraph moves.
    /// This is the property the old code could not express at all.
    #[test]
    fn paragraph_relative_anchor_tracks_the_paragraph() {
        let g = PageGeometry::default();
        let near_top = resolve_anchor_v(ST_RelativeFromV::Paragraph, 5.0, None, 0.0, &g, 0.0);
        let further_down = resolve_anchor_v(ST_RelativeFromV::Paragraph, 5.0, None, 0.0, &g, 300.0);
        assert_eq!(near_top, 77.0);
        assert_eq!(further_down, 377.0);
        assert!(further_down > near_top);
    }

    // F-X016, alignment placement and text wrapping.

    #[test]
    fn an_aligned_anchor_resolves_against_its_frame() {
        let g = PageGeometry::default();
        let width = 100.0;
        let height = 50.0;

        // Margin frame: the text area.
        let text_left = g.margin_left;
        let text_width = g.page_width - g.margin_left - g.margin_right;

        assert_eq!(
            resolve_anchor_h(
                ST_RelativeFromH::Margin,
                999.0,
                Some(AnchorAlignH::Left),
                width,
                &g,
                0.0
            ),
            text_left,
            "an alignment replaces the offset rather than adding to it"
        );
        assert_eq!(
            resolve_anchor_h(
                ST_RelativeFromH::Margin,
                0.0,
                Some(AnchorAlignH::Right),
                width,
                &g,
                0.0
            ),
            text_left + text_width - width
        );
        assert_eq!(
            resolve_anchor_h(
                ST_RelativeFromH::Margin,
                0.0,
                Some(AnchorAlignH::Center),
                width,
                &g,
                0.0
            ),
            text_left + (text_width - width) / 2.0
        );

        // Page frame, vertical axis.
        assert_eq!(
            resolve_anchor_v(
                ST_RelativeFromV::Page,
                0.0,
                Some(AnchorAlignV::Top),
                height,
                &g,
                0.0
            ),
            0.0
        );
        assert_eq!(
            resolve_anchor_v(
                ST_RelativeFromV::Page,
                0.0,
                Some(AnchorAlignV::Bottom),
                height,
                &g,
                0.0
            ),
            g.page_height - height
        );
    }

    #[test]
    fn an_anchor_without_an_alignment_still_uses_its_offset() {
        // This is what keeps every existing baseline still.
        let g = PageGeometry::default();
        assert_eq!(
            resolve_anchor_h(ST_RelativeFromH::Page, 10.0, None, 100.0, &g, 0.0),
            10.0
        );
        assert_eq!(
            resolve_anchor_h(ST_RelativeFromH::Margin, 10.0, None, 100.0, &g, 0.0),
            g.margin_left + 10.0
        );
        assert_eq!(
            resolve_anchor_v(ST_RelativeFromV::Paragraph, 5.0, None, 50.0, &g, 300.0),
            g.margin_top + 300.0 + 5.0
        );
    }

    // F-X019, paragraph-relative drawings in later blocks should wrap.

    fn wrapping_drawing(rel_v: ST_RelativeFromV) -> AnchoredDrawing {
        AnchoredDrawing {
            behind_doc: false,
            rel_h: ST_RelativeFromH::Margin,
            off_h: 0.0,
            rel_v,
            off_v: 0.0,
            width: 100.0,
            height: 50.0,
            wrap: WrapType::Square,
            dist_top: 0.0,
            dist_bottom: 0.0,
            dist_left: 0.0,
            dist_right: 0.0,
            align_h: None,
            align_v: None,
            content: AnchoredContent::Image {
                media_id: MediaId(1),
            },
        }
    }

    fn para_anchoring(rel_v: ST_RelativeFromV) -> LayoutBlock {
        let mut para = make_para(1, 14.0);
        para.anchored = vec![wrapping_drawing(rel_v)];
        LayoutBlock::Paragraph(para)
    }

    #[test]
    fn the_two_pass_predicate_matches_only_paragraph_relative_wraps() {
        assert!(!has_paragraph_relative_wrap(&[LayoutBlock::Paragraph(
            make_para(3, 14.0)
        )]));
        assert!(!has_paragraph_relative_wrap(&[para_anchoring(
            ST_RelativeFromV::Page
        )]));
        assert!(has_paragraph_relative_wrap(&[para_anchoring(
            ST_RelativeFromV::Paragraph
        )]));
        assert!(has_paragraph_relative_wrap(&[para_anchoring(
            ST_RelativeFromV::Line
        )]));

        // A paragraph-relative drawing that does not wrap pushes nothing
        // aside, so it must not buy the document a second pass.
        let mut still = wrapping_drawing(ST_RelativeFromV::Paragraph);
        still.wrap = WrapType::None;
        let mut para = make_para(1, 14.0);
        para.anchored = vec![still];
        assert!(!has_paragraph_relative_wrap(&[LayoutBlock::Paragraph(
            para
        )]));
    }

    #[test]
    fn pass_one_ignores_paragraph_relative_anchors() {
        let fm = FontManager::new();
        let media = HashMap::new();
        let notes = NoteRegistry::default();
        let empty = ResolvedWraps::new();
        let blocks = vec![
            LayoutBlock::Paragraph(make_para(3, 14.0)),
            para_anchoring(ST_RelativeFromV::Paragraph),
            para_anchoring(ST_RelativeFromV::Page),
        ];
        let pager = Pager::new(
            PageGeometry::default(),
            None,
            false,
            &media,
            &notes,
            &fm,
            &empty,
            1,
            1,
        );

        // With nothing resolved, the look-ahead offers the page-relative
        // drawing and nothing else, which is what it did before this story.
        assert_eq!(pager.lookahead_wraps(0, &blocks).len(), 1);
    }

    #[test]
    fn the_lookahead_offers_a_resolved_rect_only_on_its_own_page() {
        let fm = FontManager::new();
        let media = HashMap::new();
        let notes = NoteRegistry::default();
        let blocks = vec![
            LayoutBlock::Paragraph(make_para(3, 14.0)),
            para_anchoring(ST_RelativeFromV::Paragraph),
        ];
        let placed = PlacedWrap {
            rect: Rect {
                x: 100.0,
                y: 200.0,
                width: 100.0,
                height: 50.0,
            },
            wrap: WrapType::Square,
            dist_top: 0.0,
            dist_bottom: 0.0,
            dist_left: 0.0,
            dist_right: 0.0,
        };

        for (recorded_page, expected) in [(1usize, 1usize), (2, 0)] {
            let mut resolved = ResolvedWraps::new();
            resolved.insert((1, 0), (recorded_page, placed));
            let pager = Pager::new(
                PageGeometry::default(),
                None,
                false,
                &media,
                &notes,
                &fm,
                &resolved,
                1,
                1,
            );

            // The pager is building page one. A drawing the previous pass put
            // on page two must not push page one's text aside.
            assert_eq!(
                pager.lookahead_wraps(0, &blocks).len(),
                expected,
                "recorded on page {recorded_page}"
            );
        }
    }

    #[test]
    fn a_placed_paragraph_relative_wrap_is_recorded_for_the_next_pass() {
        let fm = FontManager::new();
        let media = HashMap::new();
        let notes = NoteRegistry::default();
        let empty = ResolvedWraps::new();
        let blocks = vec![
            LayoutBlock::Paragraph(make_para(3, 14.0)),
            para_anchoring(ST_RelativeFromV::Paragraph),
            para_anchoring(ST_RelativeFromV::Page),
        ];

        let context = PassContext {
            geometry: PageGeometry::default(),
            header_footer: None,
            title_pg: false,
            fm: &fm,
            media: &media,
            notes: &notes,
            first_page_number: 1,
            first_header_page_number: 1,
        };
        let pass = paginate_pass(&blocks, &context, &empty);

        // Only the paragraph-relative one is recorded. The page-relative one
        // needs no second pass to be known.
        assert_eq!(pass.resolved.len(), 1);
        let (page, placed) = pass.resolved.get(&(1, 0)).expect("block one, anchor zero");
        assert_eq!(*page, 1);
        assert_eq!(placed.wrap, WrapType::Square);
    }
}
