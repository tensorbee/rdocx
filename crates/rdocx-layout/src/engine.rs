//! Layout engine orchestrator: ties all phases together.

use std::collections::{HashMap, VecDeque};

use rdocx_oxml::borders::{CT_PBdr, CT_TabStop};
use rdocx_oxml::content_control::{CT_Sdt, SdtContent};
use rdocx_oxml::document::{BodyContent, CT_SectPr};
use rdocx_oxml::drawing::WrapType;
use rdocx_oxml::header_footer::{HdrFtrType, VmlWatermark};
use rdocx_oxml::numbering::ST_LvlSuffix;
use rdocx_oxml::properties::{CT_PPr, CT_RPr, CT_Shd};
use rdocx_oxml::revision::{CT_Revision, RevisionContent, RevisionKind};
use rdocx_oxml::shared::ST_HighlightColor;
use rdocx_oxml::styles::CT_Styles;
use rdocx_oxml::table::{CT_Row, CT_Tbl, CT_Tc, CellContent};
use rdocx_oxml::text::{
    BreakType, CT_P, CT_R, Field, FieldArgument, RunContent, hyperlink_revision_index,
};

use crate::block::{self, LayoutBlock, ParagraphBlock};
use crate::convert;
use crate::input::{LayoutInput, MediaRegistry, RevisionView};
use crate::notes::NoteRegistry;
use crate::paginator::{self, HeaderFooterContent, PageGeometry};
use crate::style_resolver::{self, NumberingState};
use crate::table;
use crate::{WordSourcePath, WordStory};
use oxml_layout::{
    Color, Diagnostic, DocumentMetadata, FieldKind, FontId, FontManager, GlyphRun, GroupElement,
    InlineItem, LayoutResult, LineItem, NoteRef, NoteStream, PageFrame, Point, PositionedElement,
    Rect, Result, SourceNodeId, SourceSpan, TextSegment, Transform, Underline, break_into_lines,
};

#[derive(Clone, Copy)]
struct ProjectedRun<'a> {
    run: &'a CT_R,
    boundary: usize,
    raw_order: RawOrder,
    ordinary_run_index: Option<usize>,
    hyperlink_index: Option<usize>,
    force_underline: bool,
    force_strike: bool,
}

#[derive(Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
enum RawOrder {
    BeforeRaw,
    Raw(usize),
    AfterRaw,
}

/// Immutable source identities allocated once before layout starts.
pub(crate) struct SourceRegistry {
    nodes: Vec<WordSourcePath>,
    ids: HashMap<WordSourcePath, SourceNodeId>,
}

impl SourceRegistry {
    fn for_input(input: &LayoutInput) -> Self {
        let mut registry = Self {
            nodes: Vec::new(),
            ids: HashMap::new(),
        };

        for (body_index, content) in input.document.body.content.iter().enumerate() {
            match content {
                BodyContent::Paragraph(_) => registry.insert(WordSourcePath {
                    story: WordStory::Document,
                    children: vec![body_index],
                }),
                BodyContent::Table(table) => {
                    registry.collect_table(table, &WordStory::Document, &[body_index])
                }
                BodyContent::ContentControl(_) | BodyContent::RawXml(_) => {}
            }
        }

        let mut headers = input.headers.iter().collect::<Vec<_>>();
        headers.sort_unstable_by_key(|(relationship_id, _)| *relationship_id);
        for (relationship_id, header) in headers {
            let story = WordStory::Header {
                relationship_id: relationship_id.clone(),
            };
            for paragraph_index in 0..header.paragraphs.len() {
                registry.insert(WordSourcePath {
                    story: story.clone(),
                    children: vec![paragraph_index],
                });
            }
        }

        let mut footers = input.footers.iter().collect::<Vec<_>>();
        footers.sort_unstable_by_key(|(relationship_id, _)| *relationship_id);
        for (relationship_id, footer) in footers {
            let story = WordStory::Footer {
                relationship_id: relationship_id.clone(),
            };
            for paragraph_index in 0..footer.paragraphs.len() {
                registry.insert(WordSourcePath {
                    story: story.clone(),
                    children: vec![paragraph_index],
                });
            }
        }

        for (story_kind, stream) in [
            (NoteStream::Footnote, input.footnotes.as_ref()),
            (NoteStream::Endnote, input.endnotes.as_ref()),
        ]
        .into_iter()
        .filter_map(|(story, stream)| stream.map(|stream| (story, stream)))
        {
            for note in &stream.footnotes {
                if stream.get_by_id(note.id).is_none() {
                    continue;
                }
                let story = match story_kind {
                    NoteStream::Footnote => WordStory::Footnote { id: note.id },
                    NoteStream::Endnote => WordStory::Endnote { id: note.id },
                };
                for paragraph_index in 0..note.paragraphs.len() {
                    registry.insert(WordSourcePath {
                        story: story.clone(),
                        children: vec![paragraph_index],
                    });
                }
            }
        }

        registry
    }

    fn collect_table(&mut self, table: &CT_Tbl, story: &WordStory, prefix: &[usize]) {
        for (row_index, row) in table.rows.iter().enumerate() {
            for (cell_index, cell) in row.cells.iter().enumerate() {
                for (content_index, content) in cell.content.iter().enumerate() {
                    let mut children = prefix.to_vec();
                    children.extend([row_index, cell_index, content_index]);
                    match content {
                        CellContent::Paragraph(_) => self.insert(WordSourcePath {
                            story: story.clone(),
                            children,
                        }),
                        CellContent::Table(table) => {
                            self.collect_table(table, story, &children);
                        }
                        CellContent::ContentControl(_) => {}
                    }
                }
            }
        }
    }

    fn insert(&mut self, path: WordSourcePath) {
        if self.ids.contains_key(&path) {
            return;
        }
        let index = u32::try_from(self.nodes.len() + 1)
            .expect("a layout result cannot contain more than u32::MAX source paragraphs");
        let id = SourceNodeId::new(index).expect("source ids are one based");
        self.nodes.push(path.clone());
        self.ids.insert(path, id);
    }

    pub(crate) fn id(&self, story: &WordStory, children: &[usize]) -> Option<SourceNodeId> {
        self.ids
            .get(&WordSourcePath {
                story: story.clone(),
                children: children.to_vec(),
            })
            .copied()
    }

    fn into_nodes(self) -> Vec<WordSourcePath> {
        self.nodes
    }
}

fn project_paragraph_runs(para: &CT_P, view: RevisionView) -> Vec<ProjectedRun<'_>> {
    let mut projected = Vec::new();
    for boundary in 0..=para.runs.len() {
        for (_, slot, revision) in para.revisions.iter().filter(|(at, _, _)| *at == boundary) {
            let hyperlink_index = hyperlink_revision_index(*slot);
            let raw_order = match hyperlink_index {
                Some(index) => {
                    if let Some(raw_before) = para
                        .hyperlinks
                        .get(index)
                        .and_then(|hyperlink| hyperlink.preserved_raw_before)
                    {
                        RawOrder::Raw(raw_before)
                    } else if para
                        .hyperlinks
                        .get(index)
                        .is_some_and(|hyperlink| boundary == hyperlink.run_end)
                    {
                        RawOrder::BeforeRaw
                    } else {
                        RawOrder::AfterRaw
                    }
                }
                None => RawOrder::Raw(*slot),
            };
            project_revision_runs(
                revision,
                view,
                boundary,
                raw_order,
                hyperlink_index,
                false,
                false,
                &mut projected,
            );
        }
        if let Some(run) = para.runs.get(boundary) {
            projected.push(ProjectedRun {
                run,
                boundary,
                raw_order: RawOrder::AfterRaw,
                ordinary_run_index: Some(boundary),
                hyperlink_index: None,
                force_underline: false,
                force_strike: false,
            });
        }
    }
    projected
}

fn project_revision_runs<'a>(
    revision: &'a CT_Revision,
    view: RevisionView,
    boundary: usize,
    raw_order: RawOrder,
    hyperlink_index: Option<usize>,
    inherited_underline: bool,
    inherited_strike: bool,
    projected: &mut Vec<ProjectedRun<'a>>,
) {
    let included = match view {
        RevisionView::Tracked => true,
        RevisionView::Accepted => matches!(
            revision.kind(),
            RevisionKind::Insertion | RevisionKind::MoveTo
        ),
    };
    if !included {
        return;
    }

    let force_underline = inherited_underline
        || (view == RevisionView::Tracked
            && matches!(
                revision.kind(),
                RevisionKind::Insertion | RevisionKind::MoveTo
            ));
    let force_strike = inherited_strike
        || (view == RevisionView::Tracked
            && matches!(
                revision.kind(),
                RevisionKind::Deletion | RevisionKind::MoveFrom
            ));
    let runs = match revision.content() {
        RevisionContent::Runs(runs) => runs.as_slice(),
        RevisionContent::Marker => &[],
        RevisionContent::PriorRunProperties(_)
        | RevisionContent::PriorParagraphProperties(_)
        | RevisionContent::PriorTableProperties(_)
        | RevisionContent::PriorSectionProperties(_) => return,
    };

    for run_boundary in 0..=runs.len() {
        for (_, nested) in revision
            .nested_revisions()
            .iter()
            .filter(|(at, _)| *at == run_boundary)
        {
            project_revision_runs(
                nested,
                view,
                boundary,
                raw_order,
                hyperlink_index,
                force_underline,
                force_strike,
                projected,
            );
        }
        if let Some(run) = runs.get(run_boundary) {
            projected.push(ProjectedRun {
                run,
                boundary,
                raw_order,
                ordinary_run_index: None,
                hyperlink_index,
                force_underline,
                force_strike,
            });
        }
    }
}

fn projected_paragraph_text(para: &CT_P, view: RevisionView) -> String {
    project_paragraph_runs(para, view)
        .iter()
        .map(|projected| projected.run.text())
        .collect()
}

fn projected_content_char_starts(run: &CT_R) -> Vec<usize> {
    let mut starts = Vec::with_capacity(run.content.len());
    let mut char_offset = 0usize;
    for content in &run.content {
        starts.push(char_offset);
        char_offset += match content {
            RunContent::Text(text) | RunContent::DeletedText(text) => text.text.chars().count(),
            RunContent::Tab | RunContent::Break(_) => 1,
            RunContent::Field(field) => field
                .projected_text()
                .map_or(0, |text| text.chars().count()),
            RunContent::Drawing(_)
            | RunContent::FootnoteRef { .. }
            | RunContent::EndnoteRef { .. }
            | RunContent::CommentReference { .. } => 0,
        };
    }
    debug_assert_eq!(char_offset, run.text().chars().count());
    starts
}

fn paragraph_has_visible_revision(para: &CT_P) -> bool {
    let property_revision = para.properties.as_ref().is_some_and(|properties| {
        properties.numbering_revision.is_some()
            || properties.change.is_some()
            || properties
                .sect_pr
                .as_ref()
                .is_some_and(|section| section.change.is_some())
            || properties
                .rpr
                .as_ref()
                .is_some_and(run_properties_have_revision)
    });
    property_revision
        || para
            .runs
            .iter()
            .filter_map(|run| run.properties.as_ref())
            .any(run_properties_have_revision)
        || para
            .revisions
            .iter()
            .any(|(_, _, revision)| revision_is_visible(revision))
}

fn run_properties_have_revision(properties: &rdocx_oxml::properties::CT_RPr) -> bool {
    properties.change.is_some() || !properties.revision_markers.is_empty()
}

fn revision_is_visible(revision: &CT_Revision) -> bool {
    match revision.content() {
        RevisionContent::Runs(runs) => {
            runs.iter().any(|run| {
                run.content.iter().any(|content| match content {
                    RunContent::Text(text) | RunContent::DeletedText(text) => !text.text.is_empty(),
                    RunContent::CommentReference { .. } => false,
                    RunContent::Tab
                    | RunContent::Break(_)
                    | RunContent::Drawing(_)
                    | RunContent::Field(_)
                    | RunContent::FootnoteRef { .. }
                    | RunContent::EndnoteRef { .. } => true,
                })
            }) || revision
                .nested_revisions()
                .iter()
                .any(|(_, nested)| revision_is_visible(nested))
        }
        RevisionContent::Marker => revision
            .nested_revisions()
            .iter()
            .any(|(_, nested)| revision_is_visible(nested)),
        RevisionContent::PriorRunProperties(_)
        | RevisionContent::PriorParagraphProperties(_)
        | RevisionContent::PriorTableProperties(_)
        | RevisionContent::PriorSectionProperties(_) => true,
    }
}

/// The layout engine.
pub struct Engine {
    font_manager: FontManager,
    paragraph_cache_context: Option<ParagraphCacheContext>,
    paragraph_cache: VecDeque<ParagraphCacheEntry>,
    paragraph_cache_bytes: usize,
    paragraph_cache_hits: usize,
    paragraph_cache_builds: usize,
    pending_paragraph_cache: Option<VecDeque<ParagraphCacheEntry>>,
    pending_paragraph_cache_bytes: usize,
    #[cfg(test)]
    pending_paragraph_cache_peak_entries: usize,
    #[cfg(test)]
    pending_paragraph_cache_peak_bytes: usize,
    paragraph_cache_reads_enabled: bool,
    /// Pages from the previous relayout of a single-section document, so an
    /// edit repaginates only the pages around the change. See
    /// [`paginator::PaginationCache`].
    pagination_cache: Option<paginator::PaginationCache>,
    /// Bumped whenever the caller/document font set changes; folded into the
    /// pagination environment fingerprint.
    fonts_generation: u64,
    /// Laid-out tables, keyed on source fingerprint, body position and the
    /// source-node table (cached blocks carry baked source spans).
    table_cache: HashMap<u64, TableCacheEntry>,
    /// Laid-out header/footer content per section source, same key regime.
    hf_cache: HashMap<u64, HfCacheEntry>,
    /// Post-field-substitution page pairs `(pristine, substituted)` from the
    /// previous relayout; while the field environment is unchanged, a page
    /// that is still the same shared pristine Arc reuses its substituted
    /// form instead of being unshared and reshaped.
    subst_prev: Vec<(std::sync::Arc<PageFrame>, std::sync::Arc<PageFrame>)>,
    /// Field environment of `subst_prev`: total pages + bookmark targets.
    subst_env: u64,
}

/// A laid-out table plus what a cache hit must replay, mirroring
/// [`ParagraphCacheEntry`]. Tables whose cells render numbering markers or
/// note references are never stored.
struct TableCacheEntry {
    block: crate::table::TableBlock,
    diagnostics: Vec<Diagnostic>,
    font_trace: Vec<FontId>,
}

/// Laid-out header/footer content for one section source, plus replay data.
struct HfCacheEntry {
    content: Option<paginator::HeaderFooterContent>,
    diagnostics: Vec<Diagnostic>,
    font_trace: Vec<FontId>,
}

#[derive(Clone, PartialEq)]
struct ParagraphCacheContext {
    styles: CT_Styles,
    theme: Option<rdocx_oxml::theme::Theme>,
}

impl ParagraphCacheContext {
    fn for_input(input: &LayoutInput) -> Self {
        Self {
            styles: input.styles.clone(),
            theme: input.theme.clone(),
        }
    }
}

#[derive(Clone, PartialEq)]
struct ParagraphCacheKey {
    paragraph: CT_P,
    content_width_bits: u64,
    revision_view: RevisionView,
}

struct ParagraphCacheEntry {
    /// Collision-tolerant prefilter for lookups; `key` stays the authority.
    fp: u64,
    key: ParagraphCacheKey,
    block: ParagraphBlock,
    diagnostics: Vec<Diagnostic>,
    font_trace: Vec<FontId>,
    bytes: usize,
}

#[cfg(not(target_arch = "wasm32"))]
thread_local! {
    /// RDOCX_TIMING profiling: per-layout accumulated sub-phase costs of the
    /// block-building walk, in ms. Slots: 0 safe-check, 1 fingerprint,
    /// 2 cache scan, 3 hit path (clone+rebind+replay), 4 miss layout,
    /// 5 staging, 6 tables, 7 headers/footers.
    static BLOCK_TIMERS: std::cell::RefCell<[f64; 8]> =
        const { std::cell::RefCell::new([0.0; 8]) };
}

#[cfg(not(target_arch = "wasm32"))]
fn timer_add(slot: usize, since: std::time::Instant) {
    BLOCK_TIMERS.with(|cell| {
        cell.borrow_mut()[slot] += since.elapsed().as_secs_f64() * 1000.0;
    });
}

const PARAGRAPH_CACHE_MAX_ENTRIES: usize = 4096;
const PARAGRAPH_CACHE_MAX_BYTES: usize = 64 * 1024 * 1024;
const CACHE_SOURCE_NODE: SourceNodeId = match SourceNodeId::new(1) {
    Some(node) => node,
    None => panic!("one is a valid source node id"),
};

impl Default for Engine {
    fn default() -> Self {
        Self::new()
    }
}

impl Engine {
    fn with_font_manager(font_manager: FontManager) -> Self {
        Self {
            font_manager,
            paragraph_cache_context: None,
            paragraph_cache: VecDeque::new(),
            paragraph_cache_bytes: 0,
            paragraph_cache_hits: 0,
            paragraph_cache_builds: 0,
            pending_paragraph_cache: None,
            pending_paragraph_cache_bytes: 0,
            #[cfg(test)]
            pending_paragraph_cache_peak_entries: 0,
            #[cfg(test)]
            pending_paragraph_cache_peak_bytes: 0,
            paragraph_cache_reads_enabled: false,
            pagination_cache: None,
            fonts_generation: 0,
            table_cache: HashMap::new(),
            hf_cache: HashMap::new(),
            subst_prev: Vec::new(),
            subst_env: 0,
        }
    }

    pub fn new() -> Self {
        Self::with_font_manager(FontManager::new())
    }

    /// Create an engine that resolves fonts without system font discovery.
    pub fn new_deterministic() -> Result<Self> {
        Ok(Self::with_font_manager(FontManager::new_deterministic()?))
    }

    /// Create an engine whose font universe is supplied entirely by the
    /// layout input, without bundled or system-font discovery.
    pub(crate) fn new_with_caller_fonts() -> Self {
        Self::with_font_manager(FontManager::new_with_fonts(Vec::new()))
    }

    /// Lay out the entire document.
    pub fn layout(&mut self, input: &LayoutInput) -> Result<LayoutResult> {
        self.layout_inner(input, None)
    }

    /// Lay out the document and retain its result-local Word source table.
    pub(crate) fn layout_with_provenance(
        &mut self,
        input: &LayoutInput,
    ) -> Result<(LayoutResult, Vec<WordSourcePath>)> {
        let sources = SourceRegistry::for_input(input);
        let result = self.layout_inner(input, Some(&sources))?;
        Ok((result, sources.into_nodes()))
    }

    fn layout_inner(
        &mut self,
        input: &LayoutInput,
        sources: Option<&SourceRegistry>,
    ) -> Result<LayoutResult> {
        // Load user-provided / DOCX-embedded fonts (highest priority). An exact
        // unchanged set is a no-op in a reusable engine.
        let fonts_changed = self.font_manager.load_additional_fonts(&input.fonts);
        self.font_manager.begin_layout();

        let paragraph_context = ParagraphCacheContext::for_input(input);
        if fonts_changed {
            self.paragraph_cache.clear();
            self.paragraph_cache_bytes = 0;
        }
        let context_matches =
            !fonts_changed && self.paragraph_cache_context.as_ref() == Some(&paragraph_context);
        self.paragraph_cache_reads_enabled = context_matches;
        if fonts_changed {
            self.fonts_generation = self.fonts_generation.wrapping_add(1);
        }
        if !context_matches {
            // Styles/theme/fonts changed (or first layout): every
            // cross-relayout cache keyed on that context is stale. This is
            // the styles/numbering/theme invalidation boundary F-X038 keys
            // paragraph reuse on, applied to page and block reuse as well
            // (numbering definitions additionally join the pagination
            // environment fingerprint below).
            self.pagination_cache = None;
            self.table_cache.clear();
            self.hf_cache.clear();
            self.subst_prev.clear();
        }
        self.pending_paragraph_cache = Some(VecDeque::new());
        self.pending_paragraph_cache_bytes = 0;
        #[cfg(test)]
        {
            self.pending_paragraph_cache_peak_entries = 0;
            self.pending_paragraph_cache_peak_bytes = 0;
        }

        let result = self.layout_transaction(input, sources);
        let pending = self.pending_paragraph_cache.take().unwrap_or_default();
        self.pending_paragraph_cache_bytes = 0;
        self.paragraph_cache_reads_enabled = false;
        if result.is_ok() {
            if !context_matches {
                self.paragraph_cache.clear();
                self.paragraph_cache_bytes = 0;
                self.paragraph_cache_context = Some(paragraph_context);
            }
            for entry in pending {
                self.publish_paragraph_cache_entry(entry);
            }
        }
        let current_fonts = self
            .font_manager
            .current_layout_fonts()
            .iter()
            .copied()
            .collect::<std::collections::HashSet<_>>();
        self.paragraph_cache.retain(|entry| {
            entry
                .font_trace
                .iter()
                .all(|font_id| current_fonts.contains(font_id))
        });
        self.paragraph_cache_bytes = self.paragraph_cache.iter().map(|entry| entry.bytes).sum();
        self.font_manager.retain_current_fonts();
        result
    }

    fn layout_transaction(
        &mut self,
        input: &LayoutInput,
        sources: Option<&SourceRegistry>,
    ) -> Result<LayoutResult> {
        let styles = &input.styles;
        let mut num_state = NumberingState::new();
        let media = MediaRegistry::new(&input.images);
        let mut diagnostics = Vec::new();

        // Re-breaking a paragraph around a floating drawing needs its line
        // breaking inputs kept alive past layout. Nearly no document has a
        // drawing that wraps, so the state is dropped again unless one does.
        let document_wraps = document_has_wrapping_drawing(input);

        // Get final section properties (body-level sectPr)
        let final_sect_pr = input
            .document
            .body
            .sect_pr
            .as_ref()
            .cloned()
            .unwrap_or_else(CT_SectPr::default_letter);

        // Build sections: each section has blocks + geometry + header/footer
        let mut sections: Vec<paginator::Section> = Vec::new();
        let mut current_blocks: Vec<LayoutBlock> = Vec::new();
        // Pagination identity per block, aligned with the block list of a
        // single-section document (the only shape the pagination cache
        // accepts; a section break mid-list leaves this misaligned, unused).
        let mut block_fps: Vec<u64> = Vec::new();
        // Reused pages and cached table/header blocks carry baked
        // result-local SourceSpans; they stay valid only while the whole
        // source-node table (ids and paths) is unchanged.
        let source_fold = sources.map_or(0, |s| fingerprint_source_nodes(&s.nodes));
        let mut current_sect_pr: Option<CT_SectPr> = None; // Will be set from paragraph sect_pr

        #[cfg(not(target_arch = "wasm32"))]
        let t_blocks = std::time::Instant::now();
        #[cfg(not(target_arch = "wasm32"))]
        BLOCK_TIMERS.with(|cell| *cell.borrow_mut() = [0.0; 8]);
        for (body_index, content) in input.document.body.content.iter().enumerate() {
            match content {
                BodyContent::Paragraph(para) => {
                    // Check if this paragraph ends a section (has sect_pr)
                    let para_sect_pr = para.properties.as_ref().and_then(|p| p.sect_pr.clone());

                    let sect_pr_for_layout = para_sect_pr
                        .as_ref()
                        .or(current_sect_pr.as_ref())
                        .unwrap_or(&final_sect_pr);
                    let geometry = sect_pr_to_geometry(sect_pr_for_layout);

                    let source =
                        sources.and_then(|sources| sources.id(&WordStory::Document, &[body_index]));
                    let mut para_block = self.layout_body_paragraph(
                        para,
                        geometry.content_width(),
                        styles,
                        input,
                        &media,
                        &mut num_state,
                        &mut diagnostics,
                        source,
                    )?;

                    if !document_wraps {
                        para_block.reflow = None;
                    }

                    // Detect heading style for outline generation
                    if let Some(level) = detect_heading_level(para, styles) {
                        para_block.heading_level = Some(level);
                        para_block.heading_text =
                            Some(projected_paragraph_text(para, input.revision_view));
                    }

                    let block = LayoutBlock::Paragraph(para_block);
                    block_fps.push(combine_fp(
                        fingerprint_paragraph(
                            para,
                            geometry.content_width(),
                            input.revision_view as u8,
                        ),
                        pagination_salt(&block),
                    ));
                    current_blocks.push(block);

                    // If this paragraph has sect_pr, it ends a section
                    if let Some(sect_pr) = para_sect_pr {
                        let geometry = sect_pr_to_geometry(&sect_pr);
                        let header_footer = self.layout_header_footer_cached(
                            &sect_pr,
                            input,
                            styles,
                            &media,
                            &mut num_state,
                            &mut diagnostics,
                            sources,
                            source_fold,
                        )?;
                        let title_pg = sect_pr.title_pg.unwrap_or(false);
                        sections.push(paginator::Section {
                            blocks: std::mem::take(&mut current_blocks),
                            geometry,
                            header_footer,
                            title_pg,
                            page_number_start: section_page_number_start(&sect_pr),
                        });
                        current_sect_pr = Some(sect_pr);
                    }
                }
                BodyContent::Table(tbl) => {
                    let sect_pr_for_layout = current_sect_pr.as_ref().unwrap_or(&final_sect_pr);
                    let geometry = sect_pr_to_geometry(sect_pr_for_layout);

                    #[cfg(not(target_arch = "wasm32"))]
                    let t_table = std::time::Instant::now();
                    let table_src_fp =
                        fingerprint_table(tbl, geometry.content_width(), input.revision_view as u8);
                    let table_key = {
                        let mut fp = Fingerprint::new();
                        fp.eat(&table_src_fp.to_le_bytes());
                        fp.eat(&(body_index as u64).to_le_bytes());
                        fp.eat(&source_fold.to_le_bytes());
                        fp.finish()
                    };
                    let table_block = if let Some(hit) = self.table_cache.get(&table_key) {
                        diagnostics.extend(hit.diagnostics.iter().cloned());
                        self.font_manager.replay_layout_font_trace(&hit.font_trace);
                        hit.block.clone()
                    } else {
                        let diagnostics_start = diagnostics.len();
                        self.font_manager.begin_paragraph_font_trace();
                        let block = table::layout_table_with_provenance(
                            tbl,
                            geometry.content_width(),
                            styles,
                            input,
                            &media,
                            &mut self.font_manager,
                            &mut num_state,
                            &mut diagnostics,
                            sources,
                            &WordStory::Document,
                            &[body_index],
                        )?;
                        let font_trace = self.font_manager.finish_paragraph_font_trace();
                        // A cache hit skips layout_table, so NumberingState
                        // would not advance through numbered cell paragraphs
                        // and note markers would freeze: such tables stay
                        // uncached, exactly like marker paragraphs.
                        if let Some(font_trace) = font_trace
                            && !table_renders_shared_state(&block)
                        {
                            if self.table_cache.len() >= 256 {
                                self.table_cache.clear();
                            }
                            self.table_cache.insert(
                                table_key,
                                TableCacheEntry {
                                    block: block.clone(),
                                    diagnostics: diagnostics[diagnostics_start..].to_vec(),
                                    font_trace,
                                },
                            );
                        }
                        block
                    };
                    let block = LayoutBlock::Table(table_block);
                    block_fps.push(combine_fp(table_src_fp, pagination_salt(&block)));
                    current_blocks.push(block);
                    #[cfg(not(target_arch = "wasm32"))]
                    timer_add(6, t_table);
                }
                _ => {} // Skip RawXml elements during layout
            }
        }

        // Remaining blocks belong to the final section
        let final_geometry = sect_pr_to_geometry(&final_sect_pr);
        #[cfg(not(target_arch = "wasm32"))]
        let t_hf = std::time::Instant::now();
        let final_hf = self.layout_header_footer_cached(
            &final_sect_pr,
            input,
            styles,
            &media,
            &mut num_state,
            &mut diagnostics,
            sources,
            source_fold,
        )?;
        #[cfg(not(target_arch = "wasm32"))]
        timer_add(7, t_hf);
        let final_title_pg = final_sect_pr.title_pg.unwrap_or(false);
        sections.push(paginator::Section {
            blocks: current_blocks,
            geometry: final_geometry,
            header_footer: final_hf,
            title_pg: final_title_pg,
            page_number_start: section_page_number_start(&final_sect_pr),
        });

        // Lay the notes out once per width, before pagination, so the paginator
        // can reserve exactly the height it will later draw. A note is broken to
        // the measure of the section carrying its reference, so every section's
        // width is registered. The endnote pages that follow the last body page
        // are drawn against `final_geometry`, which belongs to the section
        // pushed just above and is therefore already in this list.
        #[cfg(not(target_arch = "wasm32"))]
        let blocks_ms = t_blocks.elapsed().as_secs_f64() * 1000.0;
        #[cfg(not(target_arch = "wasm32"))]
        let t_notes = std::time::Instant::now();
        let content_widths: Vec<f64> = sections
            .iter()
            .map(|section| section.geometry.content_width())
            .collect();
        let notes = NoteRegistry::build(
            input,
            styles,
            &media,
            &mut self.font_manager,
            &mut num_state,
            &content_widths,
            &mut diagnostics,
            sources,
        )?;

        // Paginate across all sections
        // Everything pagination reads besides the block stream; a change in
        // any of it invalidates the whole pagination cache. Styles and theme
        // invalidate via the paragraph-cache context gate in layout_inner;
        // numbering definitions, note parts, header/footer parts, section
        // properties, fonts, revision view and the source-node table are
        // folded here.
        let env_fp = {
            let mut fp = Fingerprint::new();
            fp.eat_debug(&final_sect_pr);
            let mut ids: Vec<&String> = input.headers.keys().collect();
            ids.sort();
            for id in ids {
                fp.eat(id.as_bytes());
                fp.eat_hdr_ftr(&input.headers[id]);
            }
            let mut ids: Vec<&String> = input.footers.keys().collect();
            ids.sort();
            for id in ids {
                fp.eat(id.as_bytes());
                fp.eat_hdr_ftr(&input.footers[id]);
            }
            fp.eat_notes(&input.footnotes);
            fp.eat_notes(&input.endnotes);
            fp.eat_debug(&input.numbering);
            fp.eat(&[input.revision_view as u8]);
            fp.eat(&self.fonts_generation.to_le_bytes());
            fp.eat(&source_fold.to_le_bytes());
            fp.finish()
        };
        #[cfg(not(target_arch = "wasm32"))]
        let notes_ms = t_notes.elapsed().as_secs_f64() * 1000.0;
        #[cfg(not(target_arch = "wasm32"))]
        let t_pag = std::time::Instant::now();
        let (mut pages, outlines) = paginator::paginate_sections_cached(
            &sections,
            &self.font_manager,
            &media,
            &notes,
            (sections.len() == 1).then_some(block_fps.as_slice()),
            env_fp,
            &mut self.pagination_cache,
        );

        // Endnotes read at the end of the document, so they follow the last
        // body page rather than sitting at the foot of their reference's page.
        paginator::append_endnote_pages(&mut pages, &notes, final_geometry);

        // Post-pagination pass: record bookmark targets and substitute fields.
        let total_pages = pages.len();
        let bookmark_pages = pages
            .iter()
            .flat_map(|page| {
                page.elements.iter().filter_map(move |element| {
                    let PositionedElement::Text(run) = element else {
                        return None;
                    };
                    let Some(FieldKind::Target(target)) = run.field_kind else {
                        return None;
                    };
                    Some((target, page.page_number))
                })
            })
            .collect::<HashMap<_, _>>();
        // Everything substitution reads besides the page itself; while it is
        // unchanged, a page that is still the same shared pristine page as
        // last relayout substitutes to the same result.
        let subst_env = {
            let mut fp = Fingerprint::new();
            fp.eat(&(total_pages as u64).to_le_bytes());
            let mut targets: Vec<(usize, usize)> =
                bookmark_pages.iter().map(|(t, p)| (*t, *p)).collect();
            targets.sort_unstable();
            for (t, p) in targets {
                fp.eat(&(t as u64).to_le_bytes());
                fp.eat(&(p as u64).to_le_bytes());
            }
            fp.finish()
        };
        let pristine: Vec<std::sync::Arc<PageFrame>> = pages.clone();
        for (i, page) in pages.iter_mut().enumerate() {
            // Pages are shared with the relayout caches; only unshare the
            // ones substitution would actually rewrite, so field-free pages
            // stay a pointer copy across relayouts.
            let has_field = page.elements.iter().any(|element| {
                matches!(element, PositionedElement::Text(run) if run.field_kind.is_some())
            });
            if !has_field {
                continue;
            }
            if subst_env == self.subst_env
                && let Some((old_pristine, old_subst)) = self.subst_prev.get(i)
                && std::sync::Arc::ptr_eq(old_pristine, &pristine[i])
            {
                *page = std::sync::Arc::clone(old_subst);
                continue;
            }
            let page = std::sync::Arc::make_mut(page);
            let page_num = page.page_number;
            substitute_fields(
                &mut page.elements,
                page_num,
                total_pages,
                &bookmark_pages,
                &mut self.font_manager,
            );
        }
        self.subst_prev = pristine.into_iter().zip(pages.iter().cloned()).collect();
        self.subst_env = subst_env;

        // Post-pagination pass: apply page background color
        apply_page_background(&mut pages, input);

        // Remap persistent manager ids to result-local ids and omit faces that
        // are no longer present in the current layout.
        #[cfg(not(target_arch = "wasm32"))]
        if std::env::var("RDOCX_TIMING").is_ok() {
            eprintln!(
                "timing: blocks {blocks_ms:.0} ms, notes {notes_ms:.0} ms, paginate+post {:.0} ms",
                t_pag.elapsed().as_secs_f64() * 1000.0
            );
            BLOCK_TIMERS.with(|cell| {
                let t = cell.borrow();
                eprintln!(
                    "timing: blocks split - safe {:.1}, fp {:.1}, scan {:.1}, hit {:.1}, miss {:.1}, stage {:.1}, tables {:.1}, hf {:.1}",
                    t[0], t[1], t[2], t[3], t[4], t[5], t[6], t[7]
                );
            });
        }
        let fonts = if self.font_manager.every_loaded_font_is_current() {
            self.font_manager.all_font_data()
        } else {
            let current_fonts = self.font_manager.current_layout_fonts().to_vec();
            canonicalize_layout_fonts(&mut pages, &self.font_manager, &current_fonts)?
        };

        // Convert core properties to document metadata
        let metadata = input.core_properties.as_ref().map(|cp| DocumentMetadata {
            title: cp.title.clone(),
            author: cp.creator.clone(),
            subject: cp.subject.clone(),
            keywords: cp.keywords.clone(),
            creator: Some("rdocx".to_string()),
        });

        let mut result = LayoutResult::from_shared(pages, fonts, metadata, outlines);
        result.diagnostics = diagnostics;
        Ok(result)
    }

    #[allow(clippy::too_many_arguments)]
    fn layout_body_paragraph(
        &mut self,
        paragraph: &CT_P,
        content_width: f64,
        styles: &CT_Styles,
        input: &LayoutInput,
        media: &MediaRegistry,
        numbering: &mut NumberingState,
        diagnostics: &mut Vec<Diagnostic>,
        source_node: Option<SourceNodeId>,
    ) -> Result<ParagraphBlock> {
        #[cfg(not(target_arch = "wasm32"))]
        let t = std::time::Instant::now();
        let cache_safe = paragraph_is_cache_safe(paragraph, styles);
        #[cfg(not(target_arch = "wasm32"))]
        timer_add(0, t);
        if !cache_safe {
            return layout_paragraph_with_source(
                paragraph,
                content_width,
                styles,
                input,
                media,
                &mut self.font_manager,
                numbering,
                diagnostics,
                source_node,
            );
        }

        // A cheap fingerprint prefilters the scan; the typed key equality
        // below stays the authority, so a hash collision cannot alias two
        // different paragraphs. This also keeps the hot lookup path free of
        // the CT_P clone the key used to need.
        #[cfg(not(target_arch = "wasm32"))]
        let t = std::time::Instant::now();
        let fp = fingerprint_paragraph(paragraph, content_width, input.revision_view as u8);
        #[cfg(not(target_arch = "wasm32"))]
        timer_add(1, t);
        #[cfg(not(target_arch = "wasm32"))]
        let t = std::time::Instant::now();
        let found = if self.paragraph_cache_reads_enabled {
            self.paragraph_cache.iter().position(|entry| {
                entry.fp == fp
                    && entry.key.content_width_bits == content_width.to_bits()
                    && entry.key.revision_view == input.revision_view
                    && entry.key.paragraph == *paragraph
            })
        } else {
            None
        };
        #[cfg(not(target_arch = "wasm32"))]
        timer_add(2, t);
        if let Some(index) = found {
            #[cfg(not(target_arch = "wasm32"))]
            let t = std::time::Instant::now();
            let entry = &self.paragraph_cache[index];
            let mut block = entry.block.clone();
            diagnostics.extend(entry.diagnostics.iter().cloned());
            let font_trace = entry.font_trace.clone();
            rebind_paragraph_source(&mut block, source_node);
            self.font_manager.replay_layout_font_trace(&font_trace);
            // No LRU refresh: VecDeque::remove is O(len) and a 700-paragraph
            // document pays it per paragraph per relayout. Eviction order
            // degrades to insertion order, which only matters once the cache
            // is actually full.
            self.paragraph_cache_hits += 1;
            #[cfg(not(target_arch = "wasm32"))]
            timer_add(3, t);
            return Ok(block);
        }

        let diagnostics_start = diagnostics.len();
        #[cfg(not(target_arch = "wasm32"))]
        let t = std::time::Instant::now();
        self.font_manager.begin_paragraph_font_trace();
        let block_result = layout_paragraph_with_source(
            paragraph,
            content_width,
            styles,
            input,
            media,
            &mut self.font_manager,
            numbering,
            diagnostics,
            Some(CACHE_SOURCE_NODE),
        );
        let font_trace = self.font_manager.finish_paragraph_font_trace();
        #[cfg(not(target_arch = "wasm32"))]
        timer_add(4, t);
        let mut block = block_result?;
        self.paragraph_cache_builds += 1;

        #[cfg(not(target_arch = "wasm32"))]
        let t = std::time::Instant::now();
        let cached_diagnostics = diagnostics[diagnostics_start..].to_vec();
        if let Some(font_trace) = font_trace {
            let key = ParagraphCacheKey {
                paragraph: paragraph.clone(),
                content_width_bits: content_width.to_bits(),
                revision_view: input.revision_view,
            };
            let bytes = paragraph_cache_entry_bytes(
                &key.paragraph,
                &block,
                &cached_diagnostics,
                font_trace.len(),
            );
            self.stage_paragraph_cache_entry(ParagraphCacheEntry {
                fp,
                key,
                block: block.clone(),
                diagnostics: cached_diagnostics,
                font_trace,
                bytes,
            });
        }

        #[cfg(not(target_arch = "wasm32"))]
        timer_add(5, t);
        rebind_paragraph_source(&mut block, source_node);
        Ok(block)
    }

    #[cfg(test)]
    fn paragraph_cache_counts(&self) -> (usize, usize) {
        (self.paragraph_cache_hits, self.paragraph_cache_builds)
    }

    fn publish_paragraph_cache_entry(&mut self, entry: ParagraphCacheEntry) {
        if entry.bytes > PARAGRAPH_CACHE_MAX_BYTES {
            return;
        }
        while self.paragraph_cache.len() >= PARAGRAPH_CACHE_MAX_ENTRIES
            || self.paragraph_cache_bytes.saturating_add(entry.bytes) > PARAGRAPH_CACHE_MAX_BYTES
        {
            let Some(evicted) = self.paragraph_cache.pop_front() else {
                break;
            };
            self.paragraph_cache_bytes = self.paragraph_cache_bytes.saturating_sub(evicted.bytes);
        }
        self.paragraph_cache_bytes += entry.bytes;
        self.paragraph_cache.push_back(entry);
    }

    fn stage_paragraph_cache_entry(&mut self, entry: ParagraphCacheEntry) {
        if entry.bytes > PARAGRAPH_CACHE_MAX_BYTES {
            return;
        }
        let Some(pending) = self.pending_paragraph_cache.as_mut() else {
            return;
        };
        while pending.len() >= PARAGRAPH_CACHE_MAX_ENTRIES
            || self
                .pending_paragraph_cache_bytes
                .saturating_add(entry.bytes)
                > PARAGRAPH_CACHE_MAX_BYTES
        {
            let Some(evicted) = pending.pop_front() else {
                break;
            };
            self.pending_paragraph_cache_bytes = self
                .pending_paragraph_cache_bytes
                .saturating_sub(evicted.bytes);
        }
        self.pending_paragraph_cache_bytes += entry.bytes;
        pending.push_back(entry);
        #[cfg(test)]
        {
            self.pending_paragraph_cache_peak_entries =
                self.pending_paragraph_cache_peak_entries.max(pending.len());
            self.pending_paragraph_cache_peak_bytes = self
                .pending_paragraph_cache_peak_bytes
                .max(self.pending_paragraph_cache_bytes);
        }
    }
}

fn paragraph_is_cache_safe(paragraph: &CT_P, styles: &CT_Styles) -> bool {
    if !paragraph.hyperlinks.is_empty()
        || !paragraph.comment_ranges.is_empty()
        || !paragraph.bookmark_markers.is_empty()
        || !paragraph.extra_xml.is_empty()
        || !paragraph.content_controls.is_empty()
        || !paragraph.revisions.is_empty()
    {
        return false;
    }

    let style_id = paragraph
        .properties
        .as_ref()
        .and_then(|properties| properties.style_id.as_deref());
    let resolved = style_resolver::resolve_paragraph_properties(style_id, styles);
    if resolved.num_id.is_some()
        || paragraph.properties.as_ref().is_some_and(|properties| {
            properties.num_id.is_some()
                || properties.sect_pr.is_some()
                || properties.numbering_revision.is_some()
                || !properties.numbering_revision_xml.is_empty()
                || properties.change.is_some()
                || !properties.revision_xml.is_empty()
                || properties.rpr.as_ref().is_some_and(|rpr| {
                    !rpr.revision_markers.is_empty()
                        || rpr.change.is_some()
                        || !rpr.revision_xml.is_empty()
                        || !rpr.revision_xml_positions.is_empty()
                })
        })
    {
        return false;
    }

    paragraph.runs.iter().all(|run| {
        run.alt_drawings.is_empty()
            && run.extra_xml.is_empty()
            && run.extra_xml_positions.is_empty()
            && run.properties.as_ref().is_none_or(|rpr| {
                rpr.revision_markers.is_empty()
                    && rpr.change.is_none()
                    && rpr.revision_xml.is_empty()
                    && rpr.revision_xml_positions.is_empty()
            })
            && run.content.iter().all(|content| {
                matches!(
                    content,
                    RunContent::Text(_) | RunContent::Tab | RunContent::Break(_)
                )
            })
    })
}

fn canonicalize_layout_fonts(
    pages: &mut [std::sync::Arc<PageFrame>],
    font_manager: &FontManager,
    current_fonts: &[FontId],
) -> Result<Vec<oxml_layout::FontData>> {
    fn collect(
        elements: &[PositionedElement],
        remap: &mut HashMap<FontId, FontId>,
        order: &mut Vec<FontId>,
    ) {
        for element in elements {
            match element {
                PositionedElement::Text(run) => {
                    if let std::collections::hash_map::Entry::Vacant(entry) =
                        remap.entry(run.font_id)
                    {
                        let local = FontId(order.len() as u32);
                        entry.insert(local);
                        order.push(run.font_id);
                    }
                }
                PositionedElement::Group(group) => collect(&group.children, remap, order),
                _ => {}
            }
        }
    }

    fn rewrite(elements: &mut [PositionedElement], remap: &HashMap<FontId, FontId>) {
        for element in elements {
            match element {
                PositionedElement::Text(run) => {
                    run.font_id = remap[&run.font_id];
                }
                PositionedElement::Group(group) => rewrite(&mut group.children, remap),
                _ => {}
            }
        }
    }

    let mut remap = HashMap::new();
    let mut order = Vec::with_capacity(current_fonts.len());
    for &font_id in current_fonts {
        if let std::collections::hash_map::Entry::Vacant(entry) = remap.entry(font_id) {
            let local = FontId(order.len() as u32);
            entry.insert(local);
            order.push(font_id);
        }
    }
    for page in pages.iter() {
        collect(&page.elements, &mut remap, &mut order);
    }
    let mut fonts = Vec::with_capacity(order.len());
    for persistent_id in order {
        let mut font = font_manager.font_data(persistent_id)?;
        font.id = remap[&persistent_id];
        fonts.push(font);
    }
    // Pages are shared with the relayout caches; when the remap is the
    // identity (persistent ids already dense and in order, the usual case
    // for a retained engine) leave them shared instead of unsharing all.
    if remap.iter().any(|(persistent, local)| persistent != local) {
        for page in pages {
            rewrite(&mut std::sync::Arc::make_mut(page).elements, &remap);
        }
    }
    Ok(fonts)
}

impl Engine {
    /// [`layout_header_footer`] behind a content-keyed cache, so a relayout
    /// does not re-shape the same header/footer parts on every edit.
    ///
    /// Entries whose blocks render numbering markers or note references are
    /// not cached: a hit skips `layout_header_footer`, which would stop
    /// NumberingState from advancing and freeze note-marker numbering. The
    /// key folds the source-node table because header paragraphs carry
    /// result-local source spans.
    #[allow(clippy::too_many_arguments)]
    fn layout_header_footer_cached(
        &mut self,
        sect_pr: &CT_SectPr,
        input: &LayoutInput,
        styles: &CT_Styles,
        media: &MediaRegistry,
        num_state: &mut NumberingState,
        diagnostics: &mut Vec<Diagnostic>,
        sources: Option<&SourceRegistry>,
        source_fold: u64,
    ) -> Result<Option<paginator::HeaderFooterContent>> {
        let key = {
            let mut fp = Fingerprint::new();
            fp.eat_debug(&sect_pr.header_refs);
            fp.eat_debug(&sect_pr.footer_refs);
            for href in &sect_pr.header_refs {
                if let Some(part) = input.headers.get(&href.rel_id) {
                    fp.eat_hdr_ftr(part);
                }
            }
            for fref in &sect_pr.footer_refs {
                if let Some(part) = input.footers.get(&fref.rel_id) {
                    fp.eat_hdr_ftr(part);
                }
            }
            let geometry = sect_pr_to_geometry(sect_pr);
            fp.eat(&geometry.content_width().to_bits().to_le_bytes());
            fp.eat(&[input.revision_view as u8]);
            fp.eat(&source_fold.to_le_bytes());
            fp.finish()
        };
        if let Some(hit) = self.hf_cache.get(&key) {
            diagnostics.extend(hit.diagnostics.iter().cloned());
            self.font_manager.replay_layout_font_trace(&hit.font_trace);
            return Ok(hit.content.clone());
        }
        let diagnostics_start = diagnostics.len();
        self.font_manager.begin_paragraph_font_trace();
        let built = layout_header_footer(
            sect_pr,
            input,
            styles,
            media,
            &mut self.font_manager,
            num_state,
            diagnostics,
            sources,
        )?;
        let font_trace = self.font_manager.finish_paragraph_font_trace();
        let cacheable = built.as_ref().is_none_or(|hf| {
            ![
                &hf.header_blocks,
                &hf.footer_blocks,
                &hf.first_header_blocks,
                &hf.first_footer_blocks,
                &hf.even_header_blocks,
                &hf.even_footer_blocks,
            ]
            .iter()
            .any(|blocks| para_blocks_render_shared_state(blocks))
        });
        if let Some(font_trace) = font_trace
            && cacheable
        {
            if self.hf_cache.len() >= 64 {
                self.hf_cache.clear();
            }
            self.hf_cache.insert(
                key,
                HfCacheEntry {
                    content: built.clone(),
                    diagnostics: diagnostics[diagnostics_start..].to_vec(),
                    font_trace,
                },
            );
        }
        Ok(built)
    }
}

fn rebind_text_source(text: &mut TextSegment, source_node: Option<SourceNodeId>) {
    match (text.source.as_mut(), source_node) {
        (Some(source), Some(node)) => source.node = node,
        (Some(_), None) => text.source = None,
        (None, _) => {}
    }
}

fn rebind_paragraph_source(block: &mut ParagraphBlock, source_node: Option<SourceNodeId>) {
    for line in &mut block.lines {
        for item in &mut line.items {
            match item {
                LineItem::Text(text) | LineItem::Marker(text) => {
                    rebind_text_source(text, source_node)
                }
                LineItem::Tab {
                    leader: Some(leader),
                    ..
                } => rebind_text_source(leader, source_node),
                _ => {}
            }
        }
    }
    if let Some(reflow) = block.reflow.as_mut() {
        for item in &mut reflow.items {
            match item {
                InlineItem::Text(text) | InlineItem::Marker(text) => {
                    rebind_text_source(text, source_node)
                }
                _ => {}
            }
        }
    }
}

fn paragraph_cache_entry_bytes(
    paragraph: &CT_P,
    block: &ParagraphBlock,
    diagnostics: &[Diagnostic],
    font_trace_len: usize,
) -> usize {
    fn option_string_bytes(value: &Option<String>) -> usize {
        value.as_ref().map_or(0, String::capacity)
    }
    fn shading_bytes(shading: &CT_Shd) -> usize {
        shading
            .val
            .capacity()
            .saturating_add(option_string_bytes(&shading.color))
            .saturating_add(option_string_bytes(&shading.fill))
    }
    fn run_properties_bytes(properties: &CT_RPr) -> usize {
        [
            &properties.style_id,
            &properties.font_ascii,
            &properties.font_hansi,
            &properties.font_east_asia,
            &properties.font_cs,
            &properties.font_ascii_theme,
            &properties.font_hansi_theme,
            &properties.color,
            &properties.color_theme,
            &properties.vert_align,
        ]
        .into_iter()
        .map(option_string_bytes)
        .fold(0usize, usize::saturating_add)
        .saturating_add(properties.shading.as_ref().map_or(0, shading_bytes))
    }
    fn border_bytes(borders: &CT_PBdr) -> usize {
        [
            &borders.top,
            &borders.bottom,
            &borders.left,
            &borders.right,
            &borders.between,
            &borders.bar,
        ]
        .into_iter()
        .map(|edge| {
            edge.as_ref()
                .and_then(|edge| edge.color.as_ref())
                .map_or(0, String::capacity)
        })
        .fold(0usize, usize::saturating_add)
    }
    fn paragraph_properties_bytes(properties: &CT_PPr) -> usize {
        option_string_bytes(&properties.style_id)
            .saturating_add(option_string_bytes(&properties.line_rule))
            .saturating_add(properties.borders.as_ref().map_or(0, border_bytes))
            .saturating_add(properties.tabs.as_ref().map_or(0, |tabs| {
                tabs.tabs
                    .capacity()
                    .saturating_mul(std::mem::size_of::<CT_TabStop>())
            }))
            .saturating_add(properties.shading.as_ref().map_or(0, shading_bytes))
            .saturating_add(properties.rpr.as_ref().map_or(0, run_properties_bytes))
    }
    fn paragraph_key_bytes(paragraph: &CT_P) -> usize {
        paragraph
            .runs
            .capacity()
            .saturating_mul(std::mem::size_of::<CT_R>())
            .saturating_add(
                paragraph
                    .runs
                    .iter()
                    .map(|run| {
                        run.content
                            .capacity()
                            .saturating_mul(std::mem::size_of::<RunContent>())
                            .saturating_add(
                                run.content
                                    .iter()
                                    .map(|content| match content {
                                        RunContent::Text(text) => text.text.capacity(),
                                        _ => 0,
                                    })
                                    .fold(0usize, usize::saturating_add),
                            )
                            .saturating_add(run.properties.as_ref().map_or(0, run_properties_bytes))
                    })
                    .fold(0usize, usize::saturating_add),
            )
            .saturating_add(
                paragraph
                    .properties
                    .as_ref()
                    .map_or(0, paragraph_properties_bytes),
            )
    }
    fn text_bytes(text: &TextSegment) -> usize {
        text.text.capacity()
            + text.glyph_ids.capacity() * std::mem::size_of::<u16>()
            + text.advances.capacity() * std::mem::size_of::<f64>()
            + text.hyperlink_url.as_ref().map_or(0, String::capacity)
    }
    fn inline_bytes(item: &InlineItem) -> usize {
        match item {
            InlineItem::Text(text) | InlineItem::Marker(text) => text_bytes(text),
            InlineItem::Group { .. } => usize::MAX,
            _ => 0,
        }
    }
    fn line_item_bytes(item: &LineItem) -> usize {
        match item {
            LineItem::Text(text) | LineItem::Marker(text) => text_bytes(text),
            LineItem::Tab { leader, .. } => leader.as_ref().map_or(0, text_bytes),
            LineItem::Group { .. } => usize::MAX,
            _ => 0,
        }
    }

    let paragraph_bytes = paragraph_key_bytes(paragraph);
    let line_bytes = block
        .lines
        .capacity()
        .saturating_mul(std::mem::size_of::<oxml_layout::LayoutLine>())
        .saturating_add(
            block
                .lines
                .iter()
                .map(|line| {
                    line.items
                        .capacity()
                        .saturating_mul(std::mem::size_of::<LineItem>())
                        .saturating_add(
                            line.items
                                .iter()
                                .map(line_item_bytes)
                                .fold(0usize, usize::saturating_add),
                        )
                })
                .fold(0usize, usize::saturating_add),
        );
    let reflow_bytes = block.reflow.as_ref().map_or(0, |reflow| {
        std::mem::size_of_val(reflow.as_ref())
            .saturating_add(
                reflow
                    .items
                    .capacity()
                    .saturating_mul(std::mem::size_of::<InlineItem>()),
            )
            .saturating_add(
                reflow
                    .items
                    .iter()
                    .map(inline_bytes)
                    .fold(0usize, usize::saturating_add),
            )
            .saturating_add(
                reflow
                    .params
                    .tab_stops
                    .capacity()
                    .saturating_mul(std::mem::size_of::<oxml_layout::TabStop>()),
            )
            .saturating_add(
                reflow
                    .params
                    .line_prefix_widths
                    .capacity()
                    .saturating_mul(std::mem::size_of::<f64>()),
            )
            .saturating_add(
                reflow
                    .params
                    .line_suffix_widths
                    .capacity()
                    .saturating_mul(std::mem::size_of::<f64>()),
            )
    });
    let diagnostic_bytes = diagnostics
        .len()
        .saturating_mul(std::mem::size_of::<Diagnostic>())
        .saturating_add(
            diagnostics
                .iter()
                .map(|diagnostic| diagnostic.message.capacity())
                .fold(0usize, usize::saturating_add),
        );
    std::mem::size_of::<ParagraphCacheEntry>()
        .saturating_add(paragraph_bytes)
        .saturating_add(line_bytes)
        .saturating_add(reflow_bytes)
        .saturating_add(if block.anchored.is_empty() {
            0
        } else {
            usize::MAX
        })
        .saturating_add(block.heading_text.as_ref().map_or(0, String::capacity))
        .saturating_add(block.borders.as_ref().map_or(0, border_bytes))
        .saturating_add(font_trace_len * std::mem::size_of::<FontId>())
        .saturating_add(diagnostic_bytes)
}

/// Apply page background color from `w:background` element to all pages.
fn apply_page_background(pages: &mut [std::sync::Arc<PageFrame>], input: &LayoutInput) {
    let bg_xml = match &input.document.background_xml {
        Some(xml) => xml,
        None => return,
    };

    // Parse w:color attribute from background XML
    let xml_str = std::str::from_utf8(bg_xml).unwrap_or("");
    let color = extract_background_color(xml_str);
    let color = match color {
        Some(c) => c,
        None => return,
    };

    // Insert a full-page FilledRect at position 0 on every page (renders underneath everything)
    for page in pages.iter_mut() {
        let page = std::sync::Arc::make_mut(page);
        page.elements.insert(
            0,
            PositionedElement::FilledRect {
                rect: Rect {
                    x: 0.0,
                    y: 0.0,
                    width: page.width,
                    height: page.height,
                },
                color,
            },
        );
    }
}

/// Extract the background color hex from w:background XML.
fn extract_background_color(xml: &str) -> Option<Color> {
    // Look for w:color="RRGGBB" or color="RRGGBB"
    for attr in ["w:color=\"", "color=\""] {
        if let Some(start) = xml.find(attr) {
            let val_start = start + attr.len();
            if let Some(end) = xml[val_start..].find('"') {
                let hex = &xml[val_start..val_start + end];
                if hex.len() == 6 && hex != "auto" {
                    return Some(Color::from_hex(hex));
                }
            }
        }
    }
    None
}

/// Replace field placeholder GlyphRuns with actual values.
fn substitute_fields(
    elements: &mut Vec<PositionedElement>,
    page_number: usize,
    total_pages: usize,
    bookmark_pages: &HashMap<usize, usize>,
    fm: &mut FontManager,
) {
    for element in elements.iter_mut() {
        if let PositionedElement::Text(run) = element
            && let Some(fk) = run.field_kind
        {
            let value = match fk {
                FieldKind::Page => page_number.to_string(),
                FieldKind::NumPages => total_pages.to_string(),
                FieldKind::TargetPage(target) => bookmark_pages
                    .get(&target)
                    .map(usize::to_string)
                    .unwrap_or_else(|| run.text.clone()),
                FieldKind::Target(_) => continue,
            };
            // Re-shape the text with the actual value
            if let Ok(shaped) = fm.shape_text(run.font_id, &value, run.font_size) {
                run.text = value;
                run.glyph_ids = shaped.glyph_ids;
                run.advances = shaped.advances;
            }
        }
    }
    elements.retain(|element| {
        !matches!(
            element,
            PositionedElement::Text(run)
                if matches!(run.field_kind, Some(FieldKind::Target(_)))
        )
    });
}

/// Detect if a paragraph has a heading style, returning the level (1-9).
fn detect_heading_level(para: &CT_P, styles: &CT_Styles) -> Option<u32> {
    let style_id = para.properties.as_ref()?.style_id.as_deref()?;
    // Check if style ID matches "Heading1" .. "Heading9"
    if let Some(rest) = style_id.strip_prefix("Heading") {
        return rest.parse::<u32>().ok().filter(|n| (1..=9).contains(n));
    }
    // Also check style name in the styles definitions
    if let Some(style_def) = styles.get_by_id(style_id)
        && let Some(ref name) = style_def.name
        && let Some(rest) = name.strip_prefix("heading ")
    {
        return rest.parse::<u32>().ok().filter(|n| (1..=9).contains(n));
    }
    None
}

/// Lay out a single paragraph into a ParagraphBlock.
pub fn layout_paragraph(
    para: &CT_P,
    available_width: f64,
    styles: &CT_Styles,
    input: &LayoutInput,
    media: &MediaRegistry,
    fm: &mut FontManager,
    num_state: &mut NumberingState,
    diagnostics: &mut Vec<Diagnostic>,
) -> Result<ParagraphBlock> {
    layout_paragraph_with_source(
        para,
        available_width,
        styles,
        input,
        media,
        fm,
        num_state,
        diagnostics,
        None,
    )
}

pub(crate) fn layout_paragraph_with_source(
    para: &CT_P,
    available_width: f64,
    styles: &CT_Styles,
    input: &LayoutInput,
    media: &MediaRegistry,
    fm: &mut FontManager,
    num_state: &mut NumberingState,
    diagnostics: &mut Vec<Diagnostic>,
    source_node: Option<SourceNodeId>,
) -> Result<ParagraphBlock> {
    layout_paragraph_with_source_in_table(
        para,
        available_width,
        styles,
        input,
        media,
        fm,
        num_state,
        diagnostics,
        source_node,
        None,
    )
}

/// [`layout_paragraph_with_source`] with the table-style paragraph
/// properties of the enclosing table, merged between docDefaults and the
/// paragraph style per the OOXML cascade.
#[allow(clippy::too_many_arguments)]
pub(crate) fn layout_paragraph_with_source_in_table(
    para: &CT_P,
    available_width: f64,
    styles: &CT_Styles,
    input: &LayoutInput,
    media: &MediaRegistry,
    fm: &mut FontManager,
    num_state: &mut NumberingState,
    diagnostics: &mut Vec<Diagnostic>,
    source_node: Option<SourceNodeId>,
    table_ppr: Option<&rdocx_oxml::properties::CT_PPr>,
) -> Result<ParagraphBlock> {
    // Resolve paragraph properties
    let para_style_id = para.properties.as_ref().and_then(|p| p.style_id.as_deref());

    let resolved_ppr =
        style_resolver::resolve_paragraph_properties_in_table(para_style_id, styles, table_ppr);

    let mut effective_ppr = resolved_ppr;

    // A numbering level carries paragraph properties of its own, mainly the
    // indentation for that level. They sit between the style and direct
    // formatting, so merge them before the direct properties rather than
    // after. Without this every level of a list draws at the same indent.
    let direct_ppr = para.properties.as_ref();
    let list_num_id = direct_ppr.and_then(|p| p.num_id).or(effective_ppr.num_id);
    let list_ilvl = direct_ppr
        .and_then(|p| p.num_ilvl)
        .or(effective_ppr.num_ilvl)
        .unwrap_or(0);
    if let (Some(num_id), Some(numbering)) = (list_num_id, input.numbering.as_ref())
        && let Some(lvl_ppr) =
            style_resolver::level_paragraph_properties(num_id, list_ilvl, numbering)
    {
        merge_direct_ppr(&mut effective_ppr, lvl_ppr);
    }

    // Merge direct paragraph properties
    if let Some(direct_ppr) = direct_ppr {
        merge_direct_ppr(&mut effective_ppr, direct_ppr);
    }

    // Convert paragraph properties to layout values
    let space_before = effective_ppr.space_before.map(|t| t.to_pt()).unwrap_or(0.0);
    let space_after = effective_ppr.space_after.map(|t| t.to_pt()).unwrap_or(0.0);
    let ind_left = effective_ppr.ind_left.map(|t| t.to_pt()).unwrap_or(0.0);
    let ind_right = effective_ppr.ind_right.map(|t| t.to_pt()).unwrap_or(0.0);
    let keep_next = effective_ppr.keep_next.unwrap_or(false);
    let keep_lines = effective_ppr.keep_lines.unwrap_or(false);
    let page_break_before = effective_ppr.page_break_before.unwrap_or(false);
    let widow_control = effective_ppr.widow_control.unwrap_or(true);
    let jc = convert::alignment(effective_ppr.jc);

    // Parse shading color
    let shading = effective_ppr
        .shading
        .as_ref()
        .and_then(|shd| shd.fill.as_ref())
        .filter(|f| f != &"auto")
        .map(|f| Color::from_hex(f));

    // Convert runs to inline items
    let mut inline_items = Vec::new();

    // Handle numbering marker
    if let (Some(num_id), Some(numbering)) = (effective_ppr.num_id, input.numbering.as_ref()) {
        let ilvl = effective_ppr.num_ilvl.unwrap_or(0);
        if let Some(marker) = style_resolver::generate_marker(num_id, ilvl, numbering, num_state) {
            // Shape the marker text
            let marker_rpr = marker.marker_rpr;
            let marker_font_size = marker_rpr.sz.map(|hp| hp.to_pt()).unwrap_or_else(|| {
                style_resolver::resolve_run_properties(para_style_id, None, styles)
                    .sz
                    .map(|hp| hp.to_pt())
                    .unwrap_or(11.0)
            });
            let marker_bold = marker_rpr.bold.unwrap_or(false);
            let marker_italic = marker_rpr.italic.unwrap_or(false);
            let marker_font_family = marker_rpr.font_ascii.as_deref();

            // Bullet glyphs are not in every font either, so the marker gets
            // the same coverage check as body text.
            if let Ok(font_id) = fm.resolve_font_for_text(
                marker_font_family,
                marker_bold,
                marker_italic,
                &marker.marker_text,
            ) && let Ok(shaped) = fm.shape_text(font_id, &marker.marker_text, marker_font_size)
            {
                let metrics = fm.metrics(font_id, marker_font_size)?;
                let color = marker_rpr
                    .color
                    .as_ref()
                    .map(|c| Color::from_hex(c))
                    .unwrap_or(Color::BLACK);

                inline_items.push(InlineItem::Marker(TextSegment {
                    text: marker.marker_text,
                    source: None,
                    font_id,
                    font_size: marker_font_size,
                    glyph_ids: shaped.glyph_ids,
                    advances: shaped.advances,
                    width: shaped.width,
                    ascent: metrics.ascent,
                    descent: metrics.descent,
                    line_gap: 0.0,
                    color,
                    bold: marker_bold,
                    italic: marker_italic,
                    underline: None,
                    strike: false,
                    dstrike: false,
                    highlight: None,
                    baseline_offset: 0.0,
                    hyperlink_url: None,
                    field_kind: None,
                    note: None,
                }));

                match marker.suffix {
                    ST_LvlSuffix::Tab => inline_items.push(InlineItem::Tab),
                    ST_LvlSuffix::Space => {
                        let shaped = fm.shape_text(font_id, " ", marker_font_size)?;
                        inline_items.push(InlineItem::Text(TextSegment {
                            text: " ".to_owned(),
                            source: None,
                            font_id,
                            font_size: marker_font_size,
                            glyph_ids: shaped.glyph_ids,
                            advances: shaped.advances,
                            width: shaped.width,
                            ascent: metrics.ascent,
                            descent: metrics.descent,
                            line_gap: 0.0,
                            color,
                            bold: marker_bold,
                            italic: marker_italic,
                            underline: None,
                            strike: false,
                            dstrike: false,
                            highlight: None,
                            baseline_offset: 0.0,
                            hyperlink_url: None,
                            field_kind: None,
                            note: None,
                        }));
                    }
                    ST_LvlSuffix::Nothing => {}
                }
            }
        }
    }

    // Build hyperlink URL map: run index → URL
    let mut run_hyperlink_url: std::collections::HashMap<usize, String> =
        std::collections::HashMap::new();
    for hl in &para.hyperlinks {
        if let Some(ref rel_id) = hl.rel_id
            && let Some(url) = input.hyperlink_urls.get(rel_id)
        {
            for run_idx in hl.run_start..hl.run_end {
                run_hyperlink_url.insert(run_idx, url.clone());
            }
        }
    }

    // Process ordinary and revision-wrapped runs in their preserved order.
    let mut marker_boundary = None;
    let mut marker_raw_before = None;
    let mut projection_char_offset = 0usize;
    for projected in project_paragraph_runs(para, input.revision_view) {
        let run = projected.run;
        let projected_run_start = projection_char_offset;
        projection_char_offset += run.text().chars().count();
        if marker_boundary != Some(projected.boundary) {
            marker_boundary = Some(projected.boundary);
            marker_raw_before = None;
        }
        push_targeted_bookmark_markers(
            &mut inline_items,
            para,
            projected.boundary,
            marker_raw_before,
            projected.raw_order,
            input,
            fm,
        )?;
        marker_raw_before = Some(projected.raw_order);
        let current_hyperlink_url = projected
            .ordinary_run_index
            .and_then(|run_index| run_hyperlink_url.get(&run_index).cloned())
            .or_else(|| {
                projected
                    .hyperlink_index
                    .and_then(|index| para.hyperlinks.get(index))
                    .and_then(|hyperlink| hyperlink.rel_id.as_deref())
                    .and_then(|rel_id| input.hyperlink_urls.get(rel_id).cloned())
            });

        let run_style_id = run.properties.as_ref().and_then(|p| p.style_id.as_deref());

        let resolved_rpr =
            style_resolver::resolve_run_properties(para_style_id, run_style_id, styles);

        // Merge direct run properties
        let mut effective_rpr = resolved_rpr;
        if let Some(ref direct_rpr) = run.properties {
            effective_rpr.merge_from(direct_rpr);
        }

        // Skip hidden text
        if effective_rpr.vanish == Some(true) {
            continue;
        }

        let mut font_size = effective_rpr.sz.map(|hp| hp.to_pt()).unwrap_or(11.0);
        let bold = effective_rpr.bold.unwrap_or(false);
        let italic = effective_rpr.italic.unwrap_or(false);

        // Resolve font family: theme font takes priority when no explicit font is set
        let font_family = resolve_font_family(&effective_rpr, input.theme.as_ref());

        // Resolve color: theme color takes priority over literal color value
        let color = resolve_run_color(&effective_rpr, input.theme.as_ref());

        // Decoration properties
        let underline = if projected.force_underline {
            Some(Underline::Single)
        } else {
            convert::underline(effective_rpr.underline)
        };
        let strike = projected.force_strike || effective_rpr.strike.unwrap_or(false);
        let dstrike = effective_rpr.dstrike.unwrap_or(false);
        let highlight = effective_rpr.highlight.and_then(highlight_to_color);

        // Superscript/subscript handling
        let mut baseline_offset = 0.0;
        if let Some(ref va) = effective_rpr.vert_align {
            match va.as_str() {
                "superscript" => {
                    // Reduce font size to ~58% and raise baseline
                    let original_size = font_size;
                    font_size *= 0.58;
                    baseline_offset = original_size * 0.33; // raise by 1/3 of original size
                }
                "subscript" => {
                    // Reduce font size to ~58% and lower baseline
                    let original_size = font_size;
                    font_size *= 0.58;
                    baseline_offset = -(original_size * 0.14); // lower
                }
                _ => {}
            }
        }

        // Position offset (in half-points, positive=raise)
        if let Some(pos) = effective_rpr.position {
            baseline_offset += pos as f64 / 2.0; // half-points to points
        }

        // Resolved against the run's own text, so a family without glyphs for
        // this script is replaced by one that has them.
        let font_id =
            fm.resolve_font_for_text(font_family.as_deref(), bold, italic, &run.text())?;
        let metrics = fm.metrics(font_id, font_size)?;

        let content_char_starts = projected_content_char_starts(run);
        for (content_index, content) in run.content.iter().enumerate() {
            let content_char_start = projected_run_start + content_char_starts[content_index];
            match content {
                RunContent::Text(ct_text) | RunContent::DeletedText(ct_text) => {
                    let text = if effective_rpr.caps == Some(true) {
                        ct_text.text.to_uppercase()
                    } else {
                        ct_text.text.clone()
                    };

                    if text.is_empty() {
                        continue;
                    }

                    let mut shaped = fm.shape_text(font_id, &text, font_size)?;
                    let source = if text == ct_text.text {
                        source_node.and_then(|node| {
                            let char_start = u32::try_from(content_char_start).ok()?;
                            let char_end =
                                u32::try_from(content_char_start + ct_text.text.chars().count())
                                    .ok()?;
                            Some(SourceSpan {
                                node,
                                char_start,
                                char_end,
                            })
                        })
                    } else {
                        None
                    };

                    // Apply character spacing from run properties (in twips)
                    if let Some(spacing) = effective_rpr.spacing {
                        let extra = spacing.to_pt();
                        for advance in &mut shaped.advances {
                            *advance += extra;
                        }
                        shaped.width += extra * shaped.advances.len() as f64;
                    }

                    inline_items.extend(convert::text_segments(TextSegment {
                        text,
                        source,
                        font_id,
                        font_size,
                        glyph_ids: shaped.glyph_ids,
                        advances: shaped.advances,
                        width: shaped.width,
                        ascent: metrics.ascent,
                        descent: metrics.descent,
                        line_gap: 0.0,
                        color,
                        bold,
                        italic,
                        underline,
                        strike,
                        dstrike,
                        highlight,
                        baseline_offset,
                        hyperlink_url: current_hyperlink_url.clone(),
                        field_kind: None,
                        note: None,
                    }));
                }
                RunContent::Tab => {
                    inline_items.push(InlineItem::Tab);
                }
                RunContent::Break(bt) => match bt {
                    BreakType::Line => inline_items.push(InlineItem::LineBreak),
                    BreakType::Page => inline_items.push(InlineItem::PageBreak),
                    BreakType::Column => inline_items.push(InlineItem::ColumnBreak),
                },
                RunContent::Drawing(drawing) => {
                    if let Some(ref inline) = drawing.inline {
                        let width = inline.extent_cx.to_pt();
                        let height = inline.extent_cy.to_pt();
                        if let Some(relationship_id) = inline.chart_rel_id.as_deref() {
                            inline_items.push(InlineItem::Group {
                                width,
                                height,
                                group: render_word_chart(
                                    relationship_id,
                                    width,
                                    height,
                                    input,
                                    fm,
                                    diagnostics,
                                )?,
                            });
                        } else {
                            inline_items.push(InlineItem::Image {
                                width,
                                height,
                                media_id: media.id_for_relationship(&inline.embed_id),
                            });
                        }
                    }
                }
                RunContent::Field(field) => {
                    let (computed_value, field_kind) = match field.instruction.name.as_str() {
                        "PAGE" => (Some("99".to_owned()), Some(FieldKind::Page)),
                        "NUMPAGES" => (Some("99".to_owned()), Some(FieldKind::NumPages)),
                        "REF" => {
                            let Some(bookmark) = field_text_argument(field, 0) else {
                                continue;
                            };
                            if let Some(text) = bookmark_text(input, bookmark) {
                                (Some(text), None)
                            } else {
                                diagnostics.push(Diagnostic {
                                    message: format!(
                                        "REF target {bookmark} was not found, stored display retained"
                                    ),
                                });
                                (None, None)
                            }
                        }
                        "PAGEREF" => {
                            let Some(bookmark) = field_text_argument(field, 0) else {
                                continue;
                            };
                            if bookmark_text(input, bookmark).is_none() {
                                diagnostics.push(Diagnostic {
                                    message: format!(
                                        "PAGEREF target {bookmark} was not found, stored display retained"
                                    ),
                                });
                                (None, None)
                            } else if let Some(target) = page_ref_id(input, bookmark) {
                                (Some("99".to_owned()), Some(FieldKind::TargetPage(target)))
                            } else {
                                (None, None)
                            }
                        }
                        _ => (None, None),
                    };
                    let stored_segments = field.cached_display_segments();
                    let segments = if let Some(value) = computed_value.as_deref() {
                        let stored_properties = stored_segments
                            .first()
                            .and_then(|(_, properties)| *properties);
                        vec![(value, stored_properties)]
                    } else {
                        stored_segments
                    };
                    for (value, stored_properties) in segments {
                        let segment_style_id =
                            stored_properties.and_then(|properties| properties.style_id.as_deref());
                        let mut segment_rpr = if stored_properties.is_some() {
                            style_resolver::resolve_run_properties(
                                para_style_id,
                                segment_style_id,
                                styles,
                            )
                        } else {
                            effective_rpr.clone()
                        };
                        if let Some(properties) = stored_properties {
                            segment_rpr.merge_from(properties);
                        }
                        if segment_rpr.vanish == Some(true) {
                            continue;
                        }
                        let mut segment_font_size =
                            segment_rpr.sz.map(|hp| hp.to_pt()).unwrap_or(11.0);
                        let segment_bold = segment_rpr.bold.unwrap_or(false);
                        let segment_italic = segment_rpr.italic.unwrap_or(false);
                        let segment_font_family =
                            resolve_font_family(&segment_rpr, input.theme.as_ref());
                        let segment_color = resolve_run_color(&segment_rpr, input.theme.as_ref());
                        let segment_underline = if projected.force_underline {
                            Some(Underline::Single)
                        } else {
                            convert::underline(segment_rpr.underline)
                        };
                        let segment_strike =
                            projected.force_strike || segment_rpr.strike.unwrap_or(false);
                        let segment_dstrike = segment_rpr.dstrike.unwrap_or(false);
                        let segment_highlight = segment_rpr.highlight.and_then(highlight_to_color);
                        let mut segment_baseline_offset = 0.0;
                        if let Some(vertical) = segment_rpr.vert_align.as_deref() {
                            match vertical {
                                "superscript" => {
                                    let original_size = segment_font_size;
                                    segment_font_size *= 0.58;
                                    segment_baseline_offset = original_size * 0.33;
                                }
                                "subscript" => {
                                    let original_size = segment_font_size;
                                    segment_font_size *= 0.58;
                                    segment_baseline_offset = -(original_size * 0.14);
                                }
                                _ => {}
                            }
                        }
                        if let Some(position) = segment_rpr.position {
                            segment_baseline_offset += position as f64 / 2.0;
                        }
                        let segment_font_id = fm.resolve_font_for_text(
                            segment_font_family.as_deref(),
                            segment_bold,
                            segment_italic,
                            value,
                        )?;
                        let segment_metrics = fm.metrics(segment_font_id, segment_font_size)?;

                        let mut start = 0usize;
                        for (index, character) in value
                            .char_indices()
                            .chain(std::iter::once((value.len(), '\0')))
                        {
                            let control = match character {
                                '\t' => Some(InlineItem::Tab),
                                '\n' => Some(InlineItem::LineBreak),
                                '\u{000c}' => Some(InlineItem::PageBreak),
                                '\u{000b}' => Some(InlineItem::ColumnBreak),
                                '\0' if index == value.len() => None,
                                _ => continue,
                            };
                            if start < index {
                                let mut text = value[start..index].to_owned();
                                if segment_rpr.caps == Some(true) {
                                    text = text.to_uppercase();
                                }
                                let mut shaped =
                                    fm.shape_text(segment_font_id, &text, segment_font_size)?;
                                if let Some(spacing) = segment_rpr.spacing {
                                    let extra = spacing.to_pt();
                                    for advance in &mut shaped.advances {
                                        *advance += extra;
                                    }
                                    shaped.width += extra * shaped.advances.len() as f64;
                                }
                                inline_items.extend(convert::text_segments(TextSegment {
                                    text,
                                    source: None,
                                    font_id: segment_font_id,
                                    font_size: segment_font_size,
                                    glyph_ids: shaped.glyph_ids,
                                    advances: shaped.advances,
                                    width: shaped.width,
                                    ascent: segment_metrics.ascent,
                                    descent: segment_metrics.descent,
                                    line_gap: 0.0,
                                    color: segment_color,
                                    bold: segment_bold,
                                    italic: segment_italic,
                                    underline: segment_underline,
                                    strike: segment_strike,
                                    dstrike: segment_dstrike,
                                    highlight: segment_highlight,
                                    baseline_offset: segment_baseline_offset,
                                    hyperlink_url: current_hyperlink_url.clone(),
                                    field_kind,
                                    note: None,
                                }));
                            }
                            if let Some(control) = control {
                                inline_items.push(control);
                                start = index + character.len_utf8();
                            }
                        }
                    }
                }
                RunContent::FootnoteRef { id } | RunContent::EndnoteRef { id } => {
                    // The two streams number independently, so the marker has
                    // to carry which one it came from.
                    let stream = match content {
                        RunContent::EndnoteRef { .. } => NoteStream::Endnote,
                        _ => NoteStream::Footnote,
                    };
                    // Render as superscript number
                    let marker = id.to_string();
                    let sup_size = font_size * 0.58;
                    let sup_offset = font_size * 0.33; // raise baseline
                    let shaped = fm.shape_text(font_id, &marker, sup_size)?;
                    let sup_metrics = fm.metrics(font_id, sup_size)?;
                    let revision_marker = input.revision_view == RevisionView::Tracked
                        && projected.ordinary_run_index.is_none();
                    inline_items.push(InlineItem::Text(TextSegment {
                        text: marker,
                        source: None,
                        font_id,
                        font_size: sup_size,
                        glyph_ids: shaped.glyph_ids,
                        advances: shaped.advances,
                        width: shaped.width,
                        ascent: sup_metrics.ascent,
                        descent: sup_metrics.descent,
                        line_gap: 0.0,
                        color,
                        bold,
                        italic,
                        underline: revision_marker.then_some(underline).flatten(),
                        strike: revision_marker && strike,
                        dstrike: revision_marker && dstrike,
                        highlight: revision_marker.then_some(highlight).flatten(),
                        baseline_offset: sup_offset,
                        hyperlink_url: None,
                        field_kind: None,
                        note: Some(NoteRef { stream, id: *id }),
                    }));
                }
                RunContent::CommentReference { .. } => {}
            }
        }
    }

    let final_marker_lower = (marker_boundary == Some(para.runs.len()))
        .then_some(marker_raw_before)
        .flatten();
    push_targeted_bookmark_markers(
        &mut inline_items,
        para,
        para.runs.len(),
        final_marker_lower,
        RawOrder::AfterRaw,
        input,
        fm,
    )?;

    // Line breaking
    let line_params = convert::line_break_params(&effective_ppr, available_width);

    let mut lines = break_into_lines(&inline_items, &line_params, fm)?;
    convert::restore_word_line_heights(&mut lines, &effective_ppr);

    // An empty paragraph renders no elements, which leaves an interactive
    // caller with no way to click into it or draw a caret there. Give its
    // line one zero-width empty text segment carrying the paragraph's
    // resolved default font and its source identity (char span 0..0).
    // Emitted whether or not a source registry is in use, so ordinary and
    // provenance layouts stay structurally identical.
    if inline_items.is_empty()
        && let Some(line) = lines.first_mut()
        && line.items.is_empty()
    {
        let mut default_rpr = style_resolver::resolve_run_properties(para_style_id, None, styles);
        // The paragraph mark's own run properties (pPr/rPr) size an empty
        // paragraph in Word — a 7pt mark makes a ~8.4pt line, not 11pt.
        if let Some(mark_rpr) = para.properties.as_ref().and_then(|p| p.rpr.as_ref()) {
            if mark_rpr.sz.is_some() {
                default_rpr.sz = mark_rpr.sz;
            }
            if mark_rpr.bold.is_some() {
                default_rpr.bold = mark_rpr.bold;
            }
            if mark_rpr.italic.is_some() {
                default_rpr.italic = mark_rpr.italic;
            }
            if mark_rpr.font_ascii.is_some() {
                default_rpr.font_ascii = mark_rpr.font_ascii.clone();
            }
        }
        let font_size = default_rpr.sz.map(|hp| hp.to_pt()).unwrap_or(11.0);
        let bold = default_rpr.bold.unwrap_or(false);
        let italic = default_rpr.italic.unwrap_or(false);
        if let Ok(font_id) =
            fm.resolve_font_for_text(default_rpr.font_ascii.as_deref(), bold, italic, " ")
            && let Ok(metrics) = fm.metrics(font_id, font_size)
        {
            line.items.push(LineItem::Text(TextSegment {
                text: String::new(),
                source: source_node.map(|node| oxml_layout::SourceSpan {
                    node,
                    char_start: 0,
                    char_end: 0,
                }),
                font_id,
                font_size,
                glyph_ids: Vec::new(),
                advances: Vec::new(),
                width: 0.0,
                ascent: metrics.ascent,
                descent: metrics.descent,
                line_gap: 0.0,
                color: Color::BLACK,
                bold,
                italic,
                underline: None,
                strike: false,
                dstrike: false,
                highlight: None,
                baseline_offset: 0.0,
                hyperlink_url: None,
                field_kind: None,
                note: None,
            }));
            // The line was sized before this segment existed, so it still
            // carries the 12pt no-metrics fallback; give it the real font
            // metrics (an empty 7pt paragraph is ~8.4pt tall, not 12).
            line.ascent = metrics.ascent;
            line.descent = metrics.descent;
        }
    }
    if inline_items.is_empty() {
        convert::restore_word_line_heights(&mut lines, &effective_ppr);
    }

    let mut result = block::build_paragraph_block(
        lines,
        space_before,
        space_after,
        effective_ppr.borders,
        shading,
        ind_left,
        ind_right,
        jc,
        keep_next,
        keep_lines,
        page_break_before,
        widow_control,
    );
    result.has_visible_revision =
        input.revision_view == RevisionView::Tracked && paragraph_has_visible_revision(para);
    result.anchored =
        collect_anchored_drawings(para, styles, input, media, fm, num_state, diagnostics)?;
    // `inline_items` is finished with here and would otherwise be dropped, so
    // handing it to the reflow costs nothing but the memory it already holds.
    // `Engine::layout` frees it again unless the document wraps.
    result.reflow = Some(Box::new(block::ParagraphReflow {
        items: inline_items,
        params: line_params,
    }));
    Ok(result)
}

fn push_targeted_bookmark_markers(
    items: &mut Vec<InlineItem>,
    paragraph: &CT_P,
    run_index: usize,
    after_raw: Option<RawOrder>,
    through_raw: RawOrder,
    input: &LayoutInput,
    fm: &mut FontManager,
) -> Result<()> {
    let mut font_id = None;
    for marker in paragraph.bookmark_markers.iter().filter(|marker| {
        marker.is_start()
            && marker.run_index() == run_index
            && after_raw.is_none_or(|after| RawOrder::Raw(marker.raw_before()) > after)
            && RawOrder::Raw(marker.raw_before()) <= through_raw
            && marker.name().is_some_and(|name| {
                document_has_page_ref(input, name) && bookmark_text(input, name).is_some()
            })
    }) {
        if let Some(target) = marker.name().and_then(|name| page_ref_id(input, name)) {
            let resolved_font = match font_id {
                Some(font_id) => font_id,
                None => {
                    let resolved = fm.resolve_font_for_text(None, false, false, " ")?;
                    font_id = Some(resolved);
                    resolved
                }
            };
            push_bookmark_marker(items, target, resolved_font);
        }
    }
    Ok(())
}

fn push_bookmark_marker(items: &mut Vec<InlineItem>, target: usize, font_id: oxml_layout::FontId) {
    items.push(InlineItem::Text(TextSegment {
        text: "\u{2060}".to_owned(),
        source: None,
        font_id,
        font_size: 1.0,
        glyph_ids: vec![0],
        advances: vec![0.0],
        width: 0.0,
        ascent: 0.0,
        descent: 0.0,
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
        field_kind: Some(FieldKind::Target(target)),
        note: None,
    }));
}

fn page_ref_id(input: &LayoutInput, name: &str) -> Option<usize> {
    let mut names = Vec::<&str>::new();
    visit_document_paragraphs(input, &mut |paragraph| {
        for projected in project_paragraph_runs(paragraph, input.revision_view) {
            let run = projected.run;
            for content in &run.content {
                let RunContent::Field(field) = content else {
                    continue;
                };
                if field.instruction.name != "PAGEREF" {
                    continue;
                }
                let Some(bookmark) = field_text_argument(field, 0) else {
                    continue;
                };
                if !names.contains(&bookmark) {
                    names.push(bookmark);
                }
            }
        }
    });
    names.iter().position(|candidate| *candidate == name)
}

fn field_text_argument(field: &Field, index: usize) -> Option<&str> {
    match field.instruction.arguments.get(index) {
        Some(FieldArgument::Text(value)) => Some(value),
        Some(FieldArgument::Nested(_)) | None => None,
    }
}

fn document_has_page_ref(input: &LayoutInput, name: &str) -> bool {
    page_ref_id(input, name).is_some()
}

fn visit_document_paragraphs<'a>(input: &'a LayoutInput, visit: &mut impl FnMut(&'a CT_P)) {
    for content in &input.document.body.content {
        match content {
            BodyContent::Paragraph(paragraph) => visit(paragraph),
            BodyContent::Table(table) => visit_table_paragraphs(table, visit),
            BodyContent::ContentControl(control) => visit_control_paragraphs(control, visit),
            BodyContent::RawXml(_) => {}
        }
    }
}

fn visit_table_paragraphs<'a>(table: &'a CT_Tbl, visit: &mut impl FnMut(&'a CT_P)) {
    for (_, _, control) in &table.content_controls {
        visit_control_paragraphs(control, visit);
    }
    for row in &table.rows {
        visit_row_paragraphs(row, visit);
    }
}

fn visit_row_paragraphs<'a>(row: &'a CT_Row, visit: &mut impl FnMut(&'a CT_P)) {
    for (_, _, control) in &row.content_controls {
        visit_control_paragraphs(control, visit);
    }
    for cell in &row.cells {
        visit_cell_paragraphs(cell, visit);
    }
}

fn visit_cell_paragraphs<'a>(cell: &'a CT_Tc, visit: &mut impl FnMut(&'a CT_P)) {
    for content in &cell.content {
        match content {
            CellContent::Paragraph(paragraph) => visit(paragraph),
            CellContent::Table(table) => visit_table_paragraphs(table, visit),
            CellContent::ContentControl(control) => visit_control_paragraphs(control, visit),
        }
    }
}

fn visit_control_paragraphs<'a>(control: &'a CT_Sdt, visit: &mut impl FnMut(&'a CT_P)) {
    for content in &control.content {
        match content {
            SdtContent::Paragraph(paragraph) => visit(paragraph),
            SdtContent::Table(table) => visit_table_paragraphs(table, visit),
            SdtContent::Row(row) => visit_row_paragraphs(row, visit),
            SdtContent::Cell(cell) => visit_cell_paragraphs(cell, visit),
            SdtContent::ContentControl(control) => visit_control_paragraphs(control, visit),
            SdtContent::Run(_) | SdtContent::RawXml(_) => {}
        }
    }
}

fn bookmark_text(input: &LayoutInput, name: &str) -> Option<String> {
    type BodyRunPosition = (usize, usize, RawOrder);
    type BookmarkStart<'a> = (Option<&'a str>, BodyRunPosition);

    let mut starts: HashMap<i32, Vec<BookmarkStart<'_>>> = HashMap::new();
    let mut ends: HashMap<i32, Vec<BodyRunPosition>> = HashMap::new();
    for (body_index, content) in input.document.body.content.iter().enumerate() {
        let BodyContent::Paragraph(paragraph) = content else {
            continue;
        };
        for marker in &paragraph.bookmark_markers {
            let Some(id) = marker.id() else {
                continue;
            };
            if marker.run_index() > paragraph.runs.len() {
                return None;
            }
            let position = (
                body_index,
                marker.run_index(),
                RawOrder::Raw(marker.raw_before()),
            );
            if marker.is_start() {
                starts
                    .entry(id)
                    .or_default()
                    .push((marker.name(), position));
            } else {
                ends.entry(id).or_default().push(position);
            }
        }
    }
    let candidates = starts
        .iter()
        .filter_map(|(id, starts)| {
            let ends = ends.get(id)?;
            (starts.len() == 1 && starts[0].0 == Some(name) && ends.len() == 1)
                .then_some((starts[0].1, ends[0]))
        })
        .collect::<Vec<_>>();
    if candidates.len() != 1 {
        return None;
    }
    let (start, end) = candidates[0];
    if start > end {
        return None;
    }
    let mut parts = Vec::new();
    for body_index in start.0..=end.0 {
        let BodyContent::Paragraph(paragraph) = &input.document.body.content[body_index] else {
            continue;
        };
        parts.push(
            project_paragraph_runs(paragraph, input.revision_view)
                .iter()
                .filter(|projected| {
                    let position = (body_index, projected.boundary, projected.raw_order);
                    position >= start && position < end
                })
                .map(|projected| projected.run.text())
                .collect::<String>(),
        );
    }
    Some(parts.join("\n"))
}

/// Whether any drawing in the document body wraps text around itself.
///
/// A document without one can never reach the reflow path, so it does not pay
/// for it.
fn document_has_wrapping_drawing(input: &LayoutInput) -> bool {
    fn paragraph_wraps(para: &CT_P, view: RevisionView) -> bool {
        project_paragraph_runs(para, view).iter().any(|projected| {
            let run = projected.run;
            run.content
                .iter()
                .filter_map(|rc| match rc {
                    RunContent::Drawing(d) => Some(d),
                    _ => None,
                })
                .chain(run.alt_drawings.iter())
                .any(|drawing| {
                    drawing
                        .anchor
                        .as_ref()
                        .is_some_and(|anchor| anchor.wrap != WrapType::None)
                })
        })
    }

    input
        .document
        .body
        .content
        .iter()
        .any(|content| match content {
            BodyContent::Paragraph(para) => paragraph_wraps(para, input.revision_view),
            BodyContent::Table(table) => table
                .rows
                .iter()
                .flat_map(|row| row.cells.iter())
                .flat_map(|cell| cell.content.iter())
                .any(|content| match content {
                    rdocx_oxml::table::CellContent::Paragraph(para) => {
                        paragraph_wraps(para, input.revision_view)
                    }
                    // A drawing inside a nested table is rare enough that the
                    // conservative answer is to look no deeper.
                    rdocx_oxml::table::CellContent::Table(_) => false,
                    rdocx_oxml::table::CellContent::ContentControl(_) => false,
                }),
            _ => false,
        })
}

fn render_word_chart(
    relationship_id: &str,
    width: f64,
    height: f64,
    input: &LayoutInput,
    fm: &mut FontManager,
    diagnostics: &mut Vec<Diagnostic>,
) -> Result<GroupElement> {
    let bounds = Rect {
        x: 0.0,
        y: 0.0,
        width,
        height,
    };
    let rendered = match input.charts.get(relationship_id) {
        Some(Ok(chart)) => oxml_chart::render_chart(
            &chart.chart,
            bounds,
            &input.chart_theme,
            &input.chart_color_map,
            fm,
        )
        .map_err(|error| error.to_string()),
        Some(Err(message)) => Err(message.clone()),
        None => Err("relationship was not resolved from the document part".to_owned()),
    };
    match rendered {
        Ok(group) => Ok(group),
        Err(detail) => {
            diagnostics.push(Diagnostic {
                message: format!("Word chart relationship {relationship_id}: {detail}"),
            });
            oxml_chart::render_chart_placeholder(bounds, fm)
                .map_err(|error| oxml_layout::LayoutError::Layout(error.to_string()))
        }
    }
}

/// Collect the floating drawings anchored to a paragraph.
///
/// The offsets stay paired with the frame they are measured from. Resolving
/// them here is not possible: a paragraph-relative offset needs the laid-out
/// position of the paragraph, which only the paginator knows.
///
/// A shape's text box is laid out here rather than later, because breaking it
/// into lines needs the font manager.
fn collect_anchored_drawings(
    para: &CT_P,
    styles: &CT_Styles,
    input: &LayoutInput,
    media: &MediaRegistry,
    fm: &mut FontManager,
    num_state: &mut NumberingState,
    diagnostics: &mut Vec<Diagnostic>,
) -> Result<Vec<block::AnchoredDrawing>> {
    let mut out = Vec::new();

    // Drawings written plainly, and drawings recovered from an
    // mc:AlternateContent block, are both anchored the same way.
    for projected in project_paragraph_runs(para, input.revision_view) {
        let run = projected.run;
        let plain = run.content.iter().filter_map(|rc| match rc {
            RunContent::Drawing(d) => Some(d),
            _ => None,
        });
        for drawing in plain.chain(run.alt_drawings.iter()) {
            let Some(anchor) = drawing.anchor.as_ref() else {
                continue;
            };

            // A picture also carries a pic:spPr, so a parsed shape alone does
            // not mean this is a shape. An embed id is what makes it a
            // picture, and that takes precedence.
            let shape = if anchor.embed_id.is_empty() && anchor.chart_rel_id.is_none() {
                anchor.shape.as_ref()
            } else {
                None
            };

            let content = if let Some(relationship_id) = anchor.chart_rel_id.as_deref() {
                block::AnchoredContent::Group(render_word_chart(
                    relationship_id,
                    anchor.extent_cx.to_pt(),
                    anchor.extent_cy.to_pt(),
                    input,
                    fm,
                    diagnostics,
                )?)
            } else {
                match shape {
                    Some(shape) => {
                        // A shape's text box wraps at the shape width.
                        let mut text = Vec::new();
                        for p in &shape.text {
                            text.push(layout_paragraph(
                                p,
                                anchor.extent_cx.to_pt(),
                                styles,
                                input,
                                media,
                                fm,
                                num_state,
                                diagnostics,
                            )?);
                        }
                        block::AnchoredContent::Shape {
                            preset: block::ShapePreset::from_prst(shape.preset.as_deref()),
                            fill: shape.solid_fill.as_deref().map(Color::from_hex),
                            text,
                        }
                    }
                    None if anchor.embed_id.is_empty() => continue,
                    None => block::AnchoredContent::Image {
                        media_id: media.id_for_relationship(&anchor.embed_id),
                    },
                }
            };

            out.push(block::AnchoredDrawing {
                behind_doc: anchor.behind_doc,
                rel_h: anchor.pos_h_relative_from,
                off_h: anchor.pos_h_offset.to_pt(),
                rel_v: anchor.pos_v_relative_from,
                off_v: anchor.pos_v_offset.to_pt(),
                width: anchor.extent_cx.to_pt(),
                height: anchor.extent_cy.to_pt(),
                wrap: anchor.wrap,
                dist_top: anchor.dist_t.to_pt(),
                dist_bottom: anchor.dist_b.to_pt(),
                dist_left: anchor.dist_l.to_pt(),
                dist_right: anchor.dist_r.to_pt(),
                align_h: anchor.pos_h_align,
                align_v: anchor.pos_v_align,
                content,
            });
        }
    }
    Ok(out)
}

/// Merge direct paragraph properties (only fields explicitly set in the XML).
fn merge_direct_ppr(effective: &mut CT_PPr, direct: &CT_PPr) {
    // Don't merge style_id — that was already used for resolution
    if direct.jc.is_some() {
        effective.jc = direct.jc;
    }
    if direct.space_before.is_some() {
        effective.space_before = direct.space_before;
    }
    if direct.space_after.is_some() {
        effective.space_after = direct.space_after;
    }
    if direct.line_spacing.is_some() {
        effective.line_spacing = direct.line_spacing;
    }
    if direct.line_rule.is_some() {
        effective.line_rule = direct.line_rule.clone();
    }
    if direct.ind_left.is_some() {
        effective.ind_left = direct.ind_left;
    }
    if direct.ind_right.is_some() {
        effective.ind_right = direct.ind_right;
    }
    if direct.ind_first_line.is_some() {
        effective.ind_first_line = direct.ind_first_line;
    }
    if direct.ind_hanging.is_some() {
        effective.ind_hanging = direct.ind_hanging;
    }
    if direct.keep_next.is_some() {
        effective.keep_next = direct.keep_next;
    }
    if direct.keep_lines.is_some() {
        effective.keep_lines = direct.keep_lines;
    }
    if direct.page_break_before.is_some() {
        effective.page_break_before = direct.page_break_before;
    }
    if direct.widow_control.is_some() {
        effective.widow_control = direct.widow_control;
    }
    if direct.borders.is_some() {
        effective.borders = direct.borders.clone();
    }
    if direct.tabs.is_some() {
        effective.tabs = direct.tabs.clone();
    }
    if direct.shading.is_some() {
        effective.shading = direct.shading.clone();
    }
    if direct.num_id.is_some() {
        effective.num_id = direct.num_id;
    }
    if direct.num_ilvl.is_some() {
        effective.num_ilvl = direct.num_ilvl;
    }
}

/// Convert section properties to page geometry.
fn sect_pr_to_geometry(sect_pr: &CT_SectPr) -> PageGeometry {
    PageGeometry {
        page_width: sect_pr.page_width.map(|t| t.to_pt()).unwrap_or(612.0),
        page_height: sect_pr.page_height.map(|t| t.to_pt()).unwrap_or(792.0),
        margin_top: sect_pr.margin_top.map(|t| t.to_pt()).unwrap_or(72.0),
        margin_right: sect_pr.margin_right.map(|t| t.to_pt()).unwrap_or(72.0),
        margin_bottom: sect_pr.margin_bottom.map(|t| t.to_pt()).unwrap_or(72.0),
        margin_left: sect_pr.margin_left.map(|t| t.to_pt()).unwrap_or(72.0),
        header_distance: sect_pr.header_distance.map(|t| t.to_pt()).unwrap_or(36.0),
        footer_distance: sect_pr.footer_distance.map(|t| t.to_pt()).unwrap_or(36.0),
    }
}

fn section_page_number_start(sect_pr: &CT_SectPr) -> Option<usize> {
    for raw in &sect_pr.extra_xml {
        let Some((name, raw_attributes)) = raw_root_start_tag(raw) else {
            continue;
        };
        let Some(attributes) = parse_raw_attributes(raw_attributes) else {
            continue;
        };
        if xml_local_name(name) != b"pgNumType"
            || !raw_name_has_namespace(name, &attributes, rdocx_oxml::namespace::W_NS, false)
        {
            continue;
        }
        let (_, value) = attributes.iter().find(|(attribute_name, _)| {
            xml_local_name(attribute_name) == b"start"
                && raw_name_has_namespace(
                    attribute_name,
                    &attributes,
                    rdocx_oxml::namespace::W_NS,
                    true,
                )
        })?;
        return decode_xml_attribute(value)?.parse().ok();
    }
    None
}

fn xml_local_name(name: &[u8]) -> &[u8] {
    name.rsplit(|byte| *byte == b':').next().unwrap_or(name)
}

fn raw_root_start_tag(raw: &[u8]) -> Option<(&[u8], &[u8])> {
    let mut cursor = 0usize;
    while cursor < raw.len() && raw[cursor].is_ascii_whitespace() {
        cursor += 1;
    }
    if raw.get(cursor) != Some(&b'<') {
        return None;
    }
    cursor += 1;
    if matches!(raw.get(cursor), Some(b'!' | b'?' | b'/')) {
        return None;
    }
    let name_start = cursor;
    while cursor < raw.len()
        && !raw[cursor].is_ascii_whitespace()
        && !matches!(raw[cursor], b'>' | b'/')
    {
        cursor += 1;
    }
    if cursor == name_start {
        return None;
    }
    let name_end = cursor;
    let attributes_start = cursor;
    let mut quote = None;
    while cursor < raw.len() {
        match (quote, raw[cursor]) {
            (None, b'\'' | b'"') => quote = Some(raw[cursor]),
            (Some(expected), found) if expected == found => quote = None,
            (None, b'>') => {
                return Some((&raw[name_start..name_end], &raw[attributes_start..cursor]));
            }
            _ => {}
        }
        cursor += 1;
    }
    None
}

fn parse_raw_attributes(attributes: &[u8]) -> Option<Vec<(&[u8], &[u8])>> {
    let mut parsed = Vec::new();
    let mut cursor = 0usize;
    while cursor < attributes.len() {
        while cursor < attributes.len() && attributes[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor == attributes.len() || attributes[cursor] == b'/' {
            break;
        }
        let name_start = cursor;
        while cursor < attributes.len()
            && !attributes[cursor].is_ascii_whitespace()
            && attributes[cursor] != b'='
        {
            cursor += 1;
        }
        let name_end = cursor;
        while cursor < attributes.len() && attributes[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if attributes.get(cursor) != Some(&b'=') {
            cursor = cursor.saturating_add(1);
            continue;
        }
        cursor += 1;
        while cursor < attributes.len() && attributes[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        let quote = *attributes.get(cursor)?;
        if !matches!(quote, b'\'' | b'"') {
            return None;
        }
        cursor += 1;
        let value_start = cursor;
        while cursor < attributes.len() && attributes[cursor] != quote {
            cursor += 1;
        }
        let value_end = cursor;
        cursor += 1;
        parsed.push((
            &attributes[name_start..name_end],
            &attributes[value_start..value_end],
        ));
    }
    Some(parsed)
}

fn raw_name_has_namespace(
    name: &[u8],
    attributes: &[(&[u8], &[u8])],
    expected: &str,
    is_attribute: bool,
) -> bool {
    let prefix = name
        .iter()
        .rposition(|byte| *byte == b':')
        .map(|separator| &name[..separator]);
    let namespace = match prefix {
        Some(prefix) => attributes.iter().find_map(|(attribute_name, value)| {
            attribute_name
                .strip_prefix(b"xmlns:")
                .is_some_and(|declared| declared == prefix)
                .then_some(*value)
        }),
        None if !is_attribute => attributes
            .iter()
            .find_map(|(attribute_name, value)| (*attribute_name == b"xmlns").then_some(*value)),
        None => None,
    };
    match namespace {
        Some(namespace) => decode_xml_attribute(namespace).is_some_and(|value| value == expected),
        None => prefix == Some(b"w".as_slice()) && expected == rdocx_oxml::namespace::W_NS,
    }
}

fn decode_xml_attribute(value: &[u8]) -> Option<String> {
    let value = std::str::from_utf8(value).ok()?;
    let mut decoded = String::with_capacity(value.len());
    let mut cursor = 0usize;
    while let Some(relative_start) = value[cursor..].find('&') {
        let entity_start = cursor + relative_start;
        decoded.push_str(&value[cursor..entity_start]);
        let entity_end = entity_start + value[entity_start..].find(';')?;
        let entity = &value[entity_start + 1..entity_end];
        match entity {
            "amp" => decoded.push('&'),
            "apos" => decoded.push('\''),
            "gt" => decoded.push('>'),
            "lt" => decoded.push('<'),
            "quot" => decoded.push('"'),
            numeric if numeric.starts_with("#x") => {
                decoded.push(char::from_u32(
                    u32::from_str_radix(&numeric[2..], 16).ok()?,
                )?);
            }
            numeric if numeric.starts_with('#') => {
                decoded.push(char::from_u32(numeric[1..].parse().ok()?)?);
            }
            _ => return None,
        }
        cursor = entity_end + 1;
    }
    decoded.push_str(&value[cursor..]);
    Some(decoded)
}

/// Lay out header and footer content (both Default and First-page).
fn layout_header_footer(
    sect_pr: &CT_SectPr,
    input: &LayoutInput,
    styles: &CT_Styles,
    media: &MediaRegistry,
    fm: &mut FontManager,
    num_state: &mut NumberingState,
    diagnostics: &mut Vec<Diagnostic>,
    sources: Option<&SourceRegistry>,
) -> Result<Option<HeaderFooterContent>> {
    let mut has_content = false;
    let mut header_blocks = Vec::new();
    let mut footer_blocks = Vec::new();
    let mut first_header_blocks = Vec::new();
    let mut first_footer_blocks = Vec::new();
    let mut even_header_blocks = Vec::new();
    let mut even_footer_blocks = Vec::new();
    let mut watermark = None;
    let mut first_watermark = None;
    let mut even_watermark = None;
    let even_headers_active = sect_pr
        .header_refs
        .iter()
        .any(|reference| reference.hdr_ftr_type == HdrFtrType::Even);

    let geometry = sect_pr_to_geometry(sect_pr);
    let width = geometry.content_width();

    for href in &sect_pr.header_refs {
        let (target_blocks, target_watermark) = match href.hdr_ftr_type {
            HdrFtrType::Default => (&mut header_blocks, &mut watermark),
            HdrFtrType::First => (&mut first_header_blocks, &mut first_watermark),
            HdrFtrType::Even => (&mut even_header_blocks, &mut even_watermark),
        };
        if let Some(hdr) = input.headers.get(&href.rel_id) {
            if target_watermark.is_none()
                && let Some(projected) = hdr.watermarks().first()
            {
                *target_watermark = layout_watermark(
                    projected,
                    &href.rel_id,
                    input,
                    media,
                    fm,
                    geometry,
                    diagnostics,
                )?;
            }
            let story = WordStory::Header {
                relationship_id: href.rel_id.clone(),
            };
            for (paragraph_index, para) in hdr.paragraphs.iter().enumerate() {
                let source = sources.and_then(|sources| sources.id(&story, &[paragraph_index]));
                let block = layout_paragraph_with_source(
                    para,
                    width,
                    styles,
                    input,
                    media,
                    fm,
                    num_state,
                    diagnostics,
                    source,
                )?;
                target_blocks.push(block);
            }
            has_content = true;
        }
    }

    for fref in &sect_pr.footer_refs {
        let target_blocks = match fref.hdr_ftr_type {
            HdrFtrType::Default => &mut footer_blocks,
            HdrFtrType::First => &mut first_footer_blocks,
            HdrFtrType::Even => &mut even_footer_blocks,
        };
        if let Some(ftr) = input.footers.get(&fref.rel_id) {
            let story = WordStory::Footer {
                relationship_id: fref.rel_id.clone(),
            };
            for (paragraph_index, para) in ftr.paragraphs.iter().enumerate() {
                let source = sources.and_then(|sources| sources.id(&story, &[paragraph_index]));
                let block = layout_paragraph_with_source(
                    para,
                    width,
                    styles,
                    input,
                    media,
                    fm,
                    num_state,
                    diagnostics,
                    source,
                )?;
                target_blocks.push(block);
            }
            has_content = true;
        }
    }

    if has_content {
        Ok(Some(HeaderFooterContent {
            header_blocks,
            footer_blocks,
            first_header_blocks,
            first_footer_blocks,
            even_header_blocks,
            even_footer_blocks,
            even_headers_active,
            watermark,
            first_watermark,
            even_watermark,
        }))
    } else {
        Ok(None)
    }
}

fn layout_watermark(
    watermark: &VmlWatermark,
    header_relationship_id: &str,
    input: &LayoutInput,
    media: &MediaRegistry,
    fm: &mut FontManager,
    geometry: PageGeometry,
    diagnostics: &mut Vec<Diagnostic>,
) -> Result<Option<GroupElement>> {
    let (width, height, rotation, opacity) = match watermark {
        VmlWatermark::Text {
            width_pt,
            height_pt,
            rotation_degrees,
            opacity,
            ..
        }
        | VmlWatermark::Image {
            width_pt,
            height_pt,
            rotation_degrees,
            opacity,
            ..
        } => (*width_pt, *height_pt, *rotation_degrees, *opacity),
    };
    let translate = Transform {
        e: geometry.margin_left + (geometry.content_width() - width) / 2.0,
        f: geometry.margin_top + (geometry.content_height() - height) / 2.0,
        ..Transform::IDENTITY
    };
    let transform = Transform::rotate_about(rotation, width / 2.0, height / 2.0).then(translate);
    let children = match watermark {
        VmlWatermark::Text {
            text,
            color,
            font_family,
            ..
        } => {
            let Some(color) = vml_color(color) else {
                diagnostics.push(Diagnostic {
                    message: format!("VML watermark colour {color:?} is unsupported"),
                });
                return Ok(None);
            };
            let estimated = width / (text.chars().count().max(1) as f64 * 0.62);
            let font_size = (height * 0.62).min(estimated).max(1.0);
            let font_id = fm.resolve_font_for_text(
                font_family.as_deref().or(Some("Calibri")),
                false,
                false,
                text,
            )?;
            let shaped = fm.shape_text(font_id, text, font_size)?;
            let metrics = fm.metrics(font_id, font_size)?;
            vec![PositionedElement::Text(GlyphRun {
                origin: Point {
                    x: (width - shaped.width) / 2.0,
                    y: (height + metrics.ascent - metrics.descent) / 2.0,
                },
                font_id,
                font_size,
                glyph_ids: shaped.glyph_ids,
                advances: shaped.advances,
                text: text.clone(),
                source: None,
                color,
                bold: false,
                italic: false,
                field_kind: None,
                note: None,
            })]
        }
        VmlWatermark::Image {
            relationship_id, ..
        } => {
            let scoped_id = format!("{header_relationship_id}\0{relationship_id}");
            let Some(image) = input.images.get(&scoped_id) else {
                diagnostics.push(Diagnostic {
                    message: format!(
                        "VML watermark image relationship {relationship_id} in header {header_relationship_id} was not resolved"
                    ),
                });
                return Ok(None);
            };
            let data = image.data.clone();
            vec![PositionedElement::Image {
                rect: Rect {
                    x: 0.0,
                    y: 0.0,
                    width,
                    height,
                },
                content_type: image.content_type.clone(),
                media_id: media.id_for_relationship(&scoped_id),
                data,
            }]
        }
    };
    Ok(Some(GroupElement {
        transform,
        clip: None,
        opacity,
        effects: Vec::new(),
        children,
    }))
}

fn vml_color(value: &str) -> Option<Color> {
    let normalized = value.trim().to_ascii_lowercase();
    let hex = match normalized.as_str() {
        "black" => "000000",
        "silver" => "c0c0c0",
        "gray" | "grey" => "808080",
        "white" => "ffffff",
        "maroon" => "800000",
        "red" => "ff0000",
        "purple" => "800080",
        "fuchsia" | "magenta" => "ff00ff",
        "green" => "008000",
        "lime" => "00ff00",
        "olive" => "808000",
        "yellow" => "ffff00",
        "navy" => "000080",
        "blue" => "0000ff",
        "teal" => "008080",
        "aqua" | "cyan" => "00ffff",
        _ => normalized.trim_start_matches('#'),
    };
    (hex.len() == 6 && hex.bytes().all(|byte| byte.is_ascii_hexdigit()))
        .then(|| Color::from_hex(hex))
}

/// Resolve the effective font family for a run, considering theme fonts.
///
/// Priority: explicit font_ascii > theme font > None (use default).
fn resolve_font_family(
    rpr: &rdocx_oxml::properties::CT_RPr,
    theme: Option<&rdocx_oxml::theme::Theme>,
) -> Option<String> {
    // Explicit font name takes priority
    if rpr.font_ascii.is_some() {
        return rpr.font_ascii.clone();
    }

    // Resolve theme font reference
    if let (Some(theme_ref), Some(theme)) = (&rpr.font_ascii_theme, theme) {
        let font = match theme_ref.as_str() {
            "majorAscii" | "majorHAnsi" | "majorBidi" | "majorEastAsia" => {
                theme.major_font.as_deref()
            }
            "minorAscii" | "minorHAnsi" | "minorBidi" | "minorEastAsia" => {
                theme.minor_font.as_deref()
            }
            _ => None,
        };
        if let Some(f) = font {
            return Some(f.to_string());
        }
    }

    None
}

/// Resolve the effective color for a run, considering theme colors.
///
/// Priority: literal color (non-auto) > theme color > black.
fn resolve_run_color(
    rpr: &rdocx_oxml::properties::CT_RPr,
    theme: Option<&rdocx_oxml::theme::Theme>,
) -> Color {
    // If theme color is specified, resolve it from the theme
    if let Some(ref theme_name) = rpr.color_theme
        && let Some(theme) = theme
        && let Some(hex) = theme.colors.get(theme_name)
    {
        return Color::from_hex(hex);
    }

    // Fall back to literal color value
    rpr.color
        .as_ref()
        .filter(|c| c.as_str() != "auto")
        .map(|c| Color::from_hex(c))
        .unwrap_or(Color::BLACK)
}

/// Convert a highlight color enum to an RGBA Color.
fn highlight_to_color(h: ST_HighlightColor) -> Option<Color> {
    match h {
        ST_HighlightColor::None => None,
        ST_HighlightColor::Black => Some(Color {
            r: 0.0,
            g: 0.0,
            b: 0.0,
            a: 1.0,
        }),
        ST_HighlightColor::Blue => Some(Color {
            r: 0.0,
            g: 0.0,
            b: 1.0,
            a: 1.0,
        }),
        ST_HighlightColor::Cyan => Some(Color {
            r: 0.0,
            g: 1.0,
            b: 1.0,
            a: 1.0,
        }),
        ST_HighlightColor::DarkBlue => Some(Color {
            r: 0.0,
            g: 0.0,
            b: 0.545,
            a: 1.0,
        }),
        ST_HighlightColor::DarkCyan => Some(Color {
            r: 0.0,
            g: 0.545,
            b: 0.545,
            a: 1.0,
        }),
        ST_HighlightColor::DarkGray => Some(Color {
            r: 0.663,
            g: 0.663,
            b: 0.663,
            a: 1.0,
        }),
        ST_HighlightColor::DarkGreen => Some(Color {
            r: 0.0,
            g: 0.392,
            b: 0.0,
            a: 1.0,
        }),
        ST_HighlightColor::DarkMagenta => Some(Color {
            r: 0.545,
            g: 0.0,
            b: 0.545,
            a: 1.0,
        }),
        ST_HighlightColor::DarkRed => Some(Color {
            r: 0.545,
            g: 0.0,
            b: 0.0,
            a: 1.0,
        }),
        ST_HighlightColor::DarkYellow => Some(Color {
            r: 0.545,
            g: 0.545,
            b: 0.0,
            a: 1.0,
        }),
        ST_HighlightColor::Green => Some(Color {
            r: 0.0,
            g: 1.0,
            b: 0.0,
            a: 1.0,
        }),
        ST_HighlightColor::LightGray => Some(Color {
            r: 0.827,
            g: 0.827,
            b: 0.827,
            a: 1.0,
        }),
        ST_HighlightColor::Magenta => Some(Color {
            r: 1.0,
            g: 0.0,
            b: 1.0,
            a: 1.0,
        }),
        ST_HighlightColor::Red => Some(Color {
            r: 1.0,
            g: 0.0,
            b: 0.0,
            a: 1.0,
        }),
        ST_HighlightColor::White => Some(Color {
            r: 1.0,
            g: 1.0,
            b: 1.0,
            a: 1.0,
        }),
        ST_HighlightColor::Yellow => Some(Color {
            r: 1.0,
            g: 1.0,
            b: 0.0,
            a: 1.0,
        }),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::input::ImageData;
    use oxml_layout::MediaId;
    use std::collections::HashMap;

    // --- restartable pagination -------------------------------------------

    fn many_paragraph_input(n: usize) -> LayoutInput {
        let mut doc = rdocx_oxml::document::CT_Document::new();
        for i in 0..n {
            let mut p = CT_P::new();
            if i % 30 == 5 {
                p.properties.get_or_insert_with(Default::default).style_id =
                    Some("Heading1".to_string());
                p.add_run(&format!("Heading {i}"));
            } else {
                p.add_run(&format!(
                    "Paragraph {i}: the quick brown fox jumps over the lazy dog, \
                     again and again, until the page runs out of room and the \
                     paginator has to start another one to hold the rest."
                ));
            }
            doc.body.add_paragraph(p);
        }
        let mut input = make_input_with_text("");
        input.document = doc;
        input
    }

    fn set_paragraph_text(input: &mut LayoutInput, idx: usize, text: &str) {
        let BodyContent::Paragraph(p) = &mut input.document.body.content[idx] else {
            panic!("expected paragraph at {idx}");
        };
        *p = CT_P::new();
        p.add_run(text);
    }

    fn pages_debug(result: &LayoutResult) -> Vec<String> {
        result.pages.iter().map(|p| format!("{p:?}")).collect()
    }

    /// Attach a text header and a "Page {PAGE}" footer to the input's body
    /// section, the common shape that exercises the header/footer cache and
    /// per-page field substitution reuse.
    fn attach_page_footer(input: &mut LayoutInput) {
        use rdocx_oxml::header_footer::{CT_HdrFtr, HdrFtrRef, HdrFtrType};
        use rdocx_oxml::text::{CT_R, Field, RunContent};

        let mut header = CT_HdrFtr::new();
        let mut hp = CT_P::new();
        hp.add_run("Equivalence header");
        header.paragraphs.push(hp);
        input.headers.insert("rIdH1".to_string(), header);

        let mut footer = CT_HdrFtr::new();
        let mut fp = CT_P::new();
        fp.add_run("Page ");
        let mut run = CT_R::new("");
        run.content = vec![RunContent::Field(Field::new(" PAGE ", "1"))];
        fp.runs.push(run);
        footer.paragraphs.push(fp);
        input.footers.insert("rIdF1".to_string(), footer);

        let sect_pr = input
            .document
            .body
            .sect_pr
            .get_or_insert_with(rdocx_oxml::document::CT_SectPr::default_letter);
        sect_pr.header_refs.push(HdrFtrRef {
            hdr_ftr_type: HdrFtrType::Default,
            rel_id: "rIdH1".to_string(),
        });
        sect_pr.footer_refs.push(HdrFtrRef {
            hdr_ftr_type: HdrFtrType::Default,
            rel_id: "rIdF1".to_string(),
        });
    }

    /// The invariant restartable pagination must never break: a relayout on
    /// an engine holding cached pages produces byte-for-byte what a fresh
    /// engine computes from scratch.
    fn assert_cached_relayout_matches_fresh(edit: impl Fn(&mut LayoutInput)) {
        assert_cached_relayout_matches_fresh_on(many_paragraph_input(120), edit);
    }

    fn assert_cached_relayout_matches_fresh_on(
        input: LayoutInput,
        edit: impl Fn(&mut LayoutInput),
    ) {
        let mut input = input;
        let mut engine = Engine::new_deterministic().expect("deterministic engine");
        let first = engine.layout(&input).expect("first layout");
        assert!(
            first.pages.len() >= 4,
            "test document must span several pages, got {}",
            first.pages.len()
        );
        edit(&mut input);
        let cached = engine.layout(&input).expect("cached relayout");
        let fresh = Engine::new_deterministic()
            .expect("deterministic engine")
            .layout(&input)
            .expect("fresh layout");
        assert_eq!(cached.pages.len(), fresh.pages.len(), "page count");
        let (a, b) = (pages_debug(&cached), pages_debug(&fresh));
        for (i, (a, b)) in a.iter().zip(&b).enumerate() {
            assert_eq!(a, b, "page {} differs between cached and fresh", i + 1);
        }
        assert_eq!(
            format!("{:?}", cached.outlines),
            format!("{:?}", fresh.outlines),
            "outlines"
        );
    }

    #[test]
    fn cached_repagination_matches_fresh_on_a_tail_edit() {
        assert_cached_relayout_matches_fresh(|input| {
            set_paragraph_text(input, 110, "changed near the end");
        });
    }

    #[test]
    fn cached_repagination_matches_fresh_on_a_middle_edit() {
        assert_cached_relayout_matches_fresh(|input| {
            set_paragraph_text(
                input,
                60,
                "changed in the middle, with enough new text that this \
                 paragraph re-breaks into a different number of lines than \
                 it had before the edit came in and moved everything around",
            );
        });
    }

    #[test]
    fn cached_repagination_matches_fresh_on_a_first_page_edit() {
        assert_cached_relayout_matches_fresh(|input| {
            set_paragraph_text(input, 0, "changed on page one");
        });
    }

    #[test]
    fn cached_repagination_matches_fresh_on_insert_and_delete() {
        assert_cached_relayout_matches_fresh(|input| {
            let mut p = CT_P::new();
            p.add_run("a brand new paragraph pushed into the middle");
            input
                .document
                .body
                .content
                .insert(60, BodyContent::Paragraph(p));
        });
        assert_cached_relayout_matches_fresh(|input| {
            input.document.body.content.remove(60);
        });
    }

    #[test]
    fn cached_repagination_matches_fresh_without_any_edit() {
        assert_cached_relayout_matches_fresh(|_| {});
    }

    #[test]
    fn cached_repagination_matches_fresh_when_a_heading_moves_pages() {
        // Outline entries are filtered by page index during reuse; growing a
        // paragraph right before a heading exercises that filtering.
        assert_cached_relayout_matches_fresh(|input| {
            set_paragraph_text(
                input,
                64,
                "grown just before the heading at index 65 so the heading \
                 slides toward the next page: the quick brown fox jumps over \
                 the lazy dog and keeps going for a good while longer, well \
                 past where the old paragraph used to stop, adding lines",
            );
        });
    }

    #[test]
    fn cached_repagination_matches_fresh_with_a_page_number_footer() {
        // Middle edit: page numbers unchanged, so reused pages take the
        // substituted-page shortcut — output must still be byte-identical.
        let mut input = many_paragraph_input(120);
        attach_page_footer(&mut input);
        assert_cached_relayout_matches_fresh_on(input, |input| {
            set_paragraph_text(input, 60, "changed in the middle");
        });
        // Insert enough text to change the page count: the substitution
        // environment changes and every field page must be redone.
        let mut input = many_paragraph_input(120);
        attach_page_footer(&mut input);
        assert_cached_relayout_matches_fresh_on(input, |input| {
            for i in 0..6 {
                let mut p = CT_P::new();
                p.add_run(
                    "a long inserted paragraph that adds real height to the \
                     document so the total page count moves, invalidating \
                     every page-number field after the insertion point",
                );
                input
                    .document
                    .body
                    .content
                    .insert(60 + i, BodyContent::Paragraph(p));
            }
        });
    }

    #[test]
    fn editing_the_header_part_invalidates_cached_pages() {
        let mut input = many_paragraph_input(120);
        attach_page_footer(&mut input);
        assert_cached_relayout_matches_fresh_on(input, |input| {
            let header = input.headers.get_mut("rIdH1").expect("header part");
            header.paragraphs[0] = {
                let mut p = CT_P::new();
                p.add_run("Rewritten header text");
                p
            };
        });
    }

    #[test]
    fn multi_section_documents_bypass_the_pagination_cache() {
        let input = make_two_section_input(6 * 1440, false);
        let mut engine = Engine::new_deterministic().expect("deterministic engine");
        let first = engine.layout(&input).expect("first layout");
        assert!(
            engine.pagination_cache.is_none(),
            "multi-section documents must not populate the pagination cache"
        );
        let second = engine.layout(&input).expect("second layout");
        assert_eq!(pages_debug(&first), pages_debug(&second));
    }

    #[test]
    fn editing_a_footnote_definition_invalidates_cached_pages() {
        // The paragraph fingerprints do not change when only the footnote
        // part changes, so this relies entirely on the environment
        // fingerprint.
        let mut input = make_input_with_footnote(&["original note text"]);
        let mut engine = Engine::new_deterministic().expect("deterministic engine");
        engine.layout(&input).expect("first layout");
        let replacement = make_input_with_footnote(&["rewritten note text"]);
        input.footnotes = replacement.footnotes.clone();
        let cached = engine.layout(&input).expect("cached relayout");
        let fresh = Engine::new_deterministic()
            .expect("deterministic engine")
            .layout(&input)
            .expect("fresh layout");
        assert_eq!(pages_debug(&cached), pages_debug(&fresh));
    }

    #[test]
    fn cached_repagination_matches_fresh_with_provenance() {
        let mut input = many_paragraph_input(120);
        let mut engine = Engine::new_deterministic().expect("deterministic engine");
        engine.layout_with_provenance(&input).expect("first layout");
        set_paragraph_text(&mut input, 60, "changed in the middle");
        let (cached, cached_nodes) = engine
            .layout_with_provenance(&input)
            .expect("cached relayout");
        let (fresh, fresh_nodes) = Engine::new_deterministic()
            .expect("deterministic engine")
            .layout_with_provenance(&input)
            .expect("fresh layout");
        assert_eq!(pages_debug(&cached), pages_debug(&fresh));
        assert_eq!(format!("{cached_nodes:?}"), format!("{fresh_nodes:?}"));
    }

    #[test]
    fn provenance_stays_correct_when_the_paragraph_count_changes() {
        // Inserting a paragraph shifts every later source id; the source-node
        // fold must refuse page reuse so no stale span survives.
        let mut input = many_paragraph_input(120);
        let mut engine = Engine::new_deterministic().expect("deterministic engine");
        engine.layout_with_provenance(&input).expect("first layout");
        let mut p = CT_P::new();
        p.add_run("a brand new paragraph pushed into the middle");
        input
            .document
            .body
            .content
            .insert(60, BodyContent::Paragraph(p));
        let (cached, cached_nodes) = engine
            .layout_with_provenance(&input)
            .expect("cached relayout");
        let (fresh, fresh_nodes) = Engine::new_deterministic()
            .expect("deterministic engine")
            .layout_with_provenance(&input)
            .expect("fresh layout");
        assert_eq!(pages_debug(&cached), pages_debug(&fresh));
        assert_eq!(format!("{cached_nodes:?}"), format!("{fresh_nodes:?}"));
    }

    #[test]
    fn changing_a_style_definition_invalidates_cached_pages() {
        // The F-X040 requirement: style/numbering/theme definitions join the
        // reuse boundary. A doc-defaults size change must not serve pages
        // laid out under the old definitions.
        let mut input = many_paragraph_input(120);
        let mut engine = Engine::new_deterministic().expect("deterministic engine");
        engine.layout(&input).expect("first layout");
        input.styles.doc_defaults = Some(rdocx_oxml::styles::CT_DocDefaults {
            rpr: Some(rdocx_oxml::CT_RPr {
                sz: Some(rdocx_oxml::HalfPoint(32)),
                ..Default::default()
            }),
            ppr: None,
        });
        let cached = engine.layout(&input).expect("cached relayout");
        let fresh = Engine::new_deterministic()
            .expect("deterministic engine")
            .layout(&input)
            .expect("fresh layout");
        assert_eq!(pages_debug(&cached), pages_debug(&fresh));
    }

    #[test]
    fn revision_views_project_wrapped_runs_in_document_order() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:r><w:t>A</w:t></w:r>
            <w:ins w:id="1" w:author="Ada"><w:r><w:t>I1</w:t></w:r><w:del w:id="2" w:author="Ben"><w:r><w:delText>D</w:delText></w:r></w:del><w:r><w:t>I2</w:t></w:r></w:ins>
            <w:del w:id="3" w:author="Cy"><w:r><w:delText>X</w:delText></w:r></w:del>
            <w:moveFrom w:id="4" w:author="Dee"><w:r><w:t>F</w:t></w:r></w:moveFrom>
            <w:moveTo w:id="5" w:author="Eve"><w:r><w:t>T</w:t></w:r></w:moveTo>
            <w:r><w:t>Z</w:t></w:r>
        </w:p></w:body></w:document>"#;
        let document = rdocx_oxml::CT_Document::from_xml(xml).expect("revision document parses");
        let BodyContent::Paragraph(paragraph) = &document.body.content[0] else {
            panic!("expected paragraph");
        };

        let accepted = project_paragraph_runs(paragraph, RevisionView::Accepted)
            .iter()
            .map(|projected| projected.run.text())
            .collect::<Vec<_>>();
        assert_eq!(accepted, ["A", "I1", "I2", "T", "Z"]);

        let tracked = project_paragraph_runs(paragraph, RevisionView::Tracked)
            .iter()
            .map(|projected| projected.run.text())
            .collect::<Vec<_>>();
        assert_eq!(tracked, ["A", "I1", "D", "I2", "X", "F", "T", "Z"]);
    }

    #[test]
    fn nested_only_revision_wrappers_project_their_visible_runs() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:ins w:id="1" w:author="Ada"><w:moveTo w:id="2" w:author="Ben"><w:r><w:t>nested</w:t></w:r></w:moveTo></w:ins>
        </w:p></w:body></w:document>"#;
        let document = rdocx_oxml::CT_Document::from_xml(xml).expect("revision document parses");
        let BodyContent::Paragraph(paragraph) = &document.body.content[0] else {
            panic!("expected paragraph");
        };
        for view in [RevisionView::Accepted, RevisionView::Tracked] {
            assert_eq!(projected_paragraph_text(paragraph, view), "nested");
        }
        assert!(paragraph_has_visible_revision(paragraph));
    }

    #[test]
    fn pageref_target_follows_an_earlier_revision_at_the_same_boundary() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:ins w:id="1" w:author="Ada"><w:r><w:t>before</w:t></w:r></w:ins>
            <w:bookmarkStart w:id="7" w:name="target"/>
            <w:fldSimple w:instr=" PAGEREF target "><w:r><w:t>1</w:t></w:r></w:fldSimple>
            <w:bookmarkEnd w:id="7"/>
        </w:p></w:body></w:document>"#;
        let document = rdocx_oxml::CT_Document::from_xml(xml).expect("revision document parses");
        let mut input = make_input_with_text("");
        input.document = document;
        let BodyContent::Paragraph(paragraph) = &input.document.body.content[0] else {
            panic!("expected paragraph");
        };
        let media = MediaRegistry::new(&input.images);
        let mut fonts = FontManager::new_deterministic().expect("bundled fonts load");
        let mut numbering = NumberingState::new();
        let mut diagnostics = Vec::new();
        let block = layout_paragraph(
            paragraph,
            468.0,
            &input.styles,
            &input,
            &media,
            &mut fonts,
            &mut numbering,
            &mut diagnostics,
        )
        .expect("paragraph lays out");
        let items = &block.reflow.expect("reflow items retained").items;
        let revision_index = items
            .iter()
            .position(|item| matches!(item, InlineItem::Text(text) if text.text == "before"))
            .expect("revision text");
        let target_index = items
            .iter()
            .position(|item| {
                matches!(item, InlineItem::Text(text) if matches!(text.field_kind, Some(FieldKind::Target(_))))
            })
            .expect("PAGEREF target");
        assert!(revision_index < target_index);
    }

    #[test]
    fn derived_revision_text_uses_the_selected_projection() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:bookmarkStart w:id="7" w:name="target"/>
            <w:ins w:id="1" w:author="Ada"><w:r><w:t>new</w:t></w:r></w:ins>
            <w:del w:id="2" w:author="Ben"><w:r><w:delText>old</w:delText></w:r></w:del>
            <w:bookmarkEnd w:id="7"/>
        </w:p></w:body></w:document>"#;
        let document = rdocx_oxml::CT_Document::from_xml(xml).expect("revision document parses");
        let mut input = make_input_with_text("");
        input.document = document;

        assert_eq!(bookmark_text(&input, "target").as_deref(), Some("new"));
        input.revision_view = RevisionView::Tracked;
        assert_eq!(bookmark_text(&input, "target").as_deref(), Some("newold"));
    }

    #[test]
    fn bookmark_after_a_terminal_hyperlink_revision_excludes_that_revision() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p>
            <w:hyperlink r:id="rId1"><w:r><w:t>link</w:t></w:r><w:ins w:id="1" w:author="Ada"><w:r><w:t>before bookmark</w:t></w:r></w:ins></w:hyperlink>
            <w:bookmarkStart w:id="7" w:name="target"/><w:r><w:t>inside</w:t></w:r><w:bookmarkEnd w:id="7"/>
        </w:p></w:body></w:document>"#;
        let document = rdocx_oxml::CT_Document::from_xml(xml).expect("revision document parses");
        let mut input = make_input_with_text("");
        input.document = document;

        assert_eq!(bookmark_text(&input, "target").as_deref(), Some("inside"));
    }

    #[test]
    fn revision_only_hyperlink_keeps_its_link_annotation() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p>
            <w:hyperlink r:id="rId1"><w:ins w:id="1" w:author="Ada"><w:r><w:t>linked revision</w:t></w:r></w:ins></w:hyperlink>
        </w:p></w:body></w:document>"#;
        let mut document =
            rdocx_oxml::CT_Document::from_xml(xml).expect("revision document parses");
        let BodyContent::Paragraph(paragraph) = &mut document.body.content[0] else {
            panic!("expected paragraph");
        };
        paragraph.hyperlinks[0].rel_id = Some("rId2".to_owned());
        let serialized = String::from_utf8(document.to_xml().expect("document serializes"))
            .expect("document XML is UTF-8");
        assert!(serialized.contains("r:id=\"rId2\""), "{serialized}");
        assert!(!serialized.contains("r:id=\"rId1\""), "{serialized}");
        let mut input = make_input_with_text("");
        input.document = document;
        input
            .hyperlink_urls
            .insert("rId2".to_owned(), "https://example.com".to_owned());

        let output = Engine::new_deterministic()
            .expect("bundled fonts load")
            .layout(&input)
            .expect("revision hyperlink lays out");
        assert!(output.pages[0].elements.iter().any(|element| {
            matches!(element, PositionedElement::LinkAnnotation { url, .. }
                if url == "https://example.com")
        }));
    }

    #[test]
    fn derived_revision_text_keeps_order_after_comment_run_removal() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:bookmarkStart w:id="7" w:name="target"/>
            <w:commentRangeStart w:id="5"/><w:r><w:commentReference w:id="5"/></w:r>
            <w:ins w:id="1" w:author="Ada"><w:r><w:t>inside</w:t></w:r></w:ins>
            <w:bookmarkEnd w:id="7"/>
        </w:p></w:body></w:document>"#;
        let mut document =
            rdocx_oxml::CT_Document::from_xml(xml).expect("revision document parses");
        let BodyContent::Paragraph(paragraph) = &mut document.body.content[0] else {
            panic!("expected paragraph");
        };
        paragraph.remove_comment_anchors(&[5]);
        let mut input = make_input_with_text("");
        input.document = document;

        assert_eq!(bookmark_text(&input, "target").as_deref(), Some("inside"));
    }

    #[test]
    fn heading_text_uses_the_selected_revision_projection() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:ins w:id="1" w:author="Ada"><w:r><w:t>new</w:t></w:r></w:ins>
            <w:del w:id="2" w:author="Ben"><w:r><w:delText>old</w:delText></w:r></w:del>
        </w:p></w:body></w:document>"#;
        let document = rdocx_oxml::CT_Document::from_xml(xml).expect("revision document parses");
        let BodyContent::Paragraph(paragraph) = &document.body.content[0] else {
            panic!("expected paragraph");
        };
        assert_eq!(
            projected_paragraph_text(paragraph, RevisionView::Accepted),
            "new"
        );
        assert_eq!(
            projected_paragraph_text(paragraph, RevisionView::Tracked),
            "newold"
        );
    }

    #[test]
    fn revised_floating_anchors_follow_the_selected_projection() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><w:body><w:p>
            <w:ins w:id="1" w:author="Ada"><w:r><w:drawing><wp:anchor behindDoc="0">
              <wp:positionH relativeFrom="margin"><wp:align>right</wp:align></wp:positionH>
              <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
              <wp:extent cx="914400" cy="457200"/><wp:wrapSquare wrapText="bothSides"/>
              <a:graphic><a:graphicData><wps:wsp><wps:spPr><a:prstGeom prst="rect"/></wps:spPr></wps:wsp></a:graphicData></a:graphic>
            </wp:anchor></w:drawing></w:r></w:ins>
        </w:p></w:body></w:document>"#;
        let document = rdocx_oxml::CT_Document::from_xml(xml).expect("revision document parses");
        let mut input = make_input_with_text("");
        input.document = document;
        assert!(document_has_wrapping_drawing(&input));

        input.revision_view = RevisionView::Tracked;
        let BodyContent::Paragraph(paragraph) = &input.document.body.content[0] else {
            panic!("expected paragraph");
        };
        let media = MediaRegistry::new(&input.images);
        let mut fonts = FontManager::new_deterministic().expect("bundled fonts load");
        let mut numbering = NumberingState::new();
        let mut diagnostics = Vec::new();
        let anchored = collect_anchored_drawings(
            paragraph,
            &input.styles,
            &input,
            &media,
            &mut fonts,
            &mut numbering,
            &mut diagnostics,
        )
        .expect("tracked anchor collection succeeds");
        assert_eq!(anchored.len(), 1);
    }

    #[test]
    fn tracked_revision_decorations_override_only_underline_and_strike() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:ins w:id="1" w:author="Ada"><w:r><w:rPr><w:rFonts w:ascii="Liberation Sans"/><w:b/><w:i/><w:u w:val="double"/><w:dstrike/><w:color w:val="AA0000"/><w:highlight w:val="yellow"/></w:rPr><w:t>inserted</w:t></w:r></w:ins>
            <w:del w:id="2" w:author="Ben"><w:r><w:rPr><w:u w:val="double"/><w:color w:val="0000AA"/></w:rPr><w:delText>deleted</w:delText></w:r></w:del>
            <w:ins w:id="3" w:author="Ada"><w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:footnoteReference w:id="11"/></w:r></w:ins>
            <w:del w:id="4" w:author="Ben"><w:r><w:endnoteReference w:id="12"/></w:r></w:del>
        </w:p></w:body></w:document>"#;
        let document = rdocx_oxml::CT_Document::from_xml(xml).expect("revision document parses");
        let mut input = make_input_with_text("");
        input.document = document;
        input.revision_view = RevisionView::Tracked;
        let BodyContent::Paragraph(paragraph) = &input.document.body.content[0] else {
            panic!("expected paragraph");
        };
        let media = MediaRegistry::new(&input.images);
        let mut fonts = FontManager::new_deterministic().expect("bundled fonts load");
        let mut numbering = NumberingState::new();
        let mut diagnostics = Vec::new();
        let block = layout_paragraph(
            paragraph,
            468.0,
            &input.styles,
            &input,
            &media,
            &mut fonts,
            &mut numbering,
            &mut diagnostics,
        )
        .expect("tracked paragraph lays out");
        let segments = block
            .lines
            .iter()
            .flat_map(|line| &line.items)
            .filter_map(|item| match item {
                oxml_layout::LineItem::Text(segment) => Some(segment),
                _ => None,
            })
            .collect::<Vec<_>>();
        let inserted = segments
            .iter()
            .find(|segment| segment.text == "inserted")
            .expect("inserted segment");
        assert_eq!(inserted.underline, Some(Underline::Single));
        assert!(!inserted.strike);
        assert!(inserted.dstrike);
        assert!(inserted.bold && inserted.italic);
        assert_eq!(inserted.color, Color::from_hex("AA0000"));
        assert_eq!(inserted.highlight, Some(Color::from_hex("FFFF00")));

        let deleted = segments
            .iter()
            .find(|segment| segment.text == "deleted")
            .expect("deleted segment");
        assert_eq!(deleted.underline, Some(Underline::Double));
        assert!(deleted.strike);
        assert_eq!(deleted.color, Color::from_hex("0000AA"));

        let inserted_note = segments
            .iter()
            .find(|segment| segment.text == "11")
            .expect("inserted note marker");
        assert_eq!(inserted_note.underline, Some(Underline::Single));
        assert_eq!(inserted_note.highlight, Some(Color::from_hex("FFFF00")));
        let deleted_note = segments
            .iter()
            .find(|segment| segment.text == "12")
            .expect("deleted note marker");
        assert!(deleted_note.strike);

        let mut accepted_input = input.clone();
        accepted_input.revision_view = RevisionView::Accepted;
        let BodyContent::Paragraph(accepted_paragraph) = &accepted_input.document.body.content[0]
        else {
            panic!("expected paragraph");
        };
        let accepted_media = MediaRegistry::new(&accepted_input.images);
        let accepted_block = layout_paragraph(
            accepted_paragraph,
            468.0,
            &accepted_input.styles,
            &accepted_input,
            &accepted_media,
            &mut fonts,
            &mut numbering,
            &mut diagnostics,
        )
        .expect("accepted paragraph lays out");
        let accepted_note = accepted_block
            .lines
            .iter()
            .flat_map(|line| &line.items)
            .filter_map(|item| match item {
                oxml_layout::LineItem::Text(segment) if segment.text == "11" => Some(segment),
                _ => None,
            })
            .next()
            .expect("accepted note marker");
        assert_eq!(accepted_note.underline, None);
        assert!(!accepted_note.strike && !accepted_note.dstrike);
        assert_eq!(accepted_note.highlight, None);
    }

    #[test]
    fn a_split_changed_paragraph_draws_one_margin_bar_on_each_page() {
        let changed = "changed ".repeat(3_000);
        let xml = format!(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:ins w:id="1" w:author="Ada"><w:r><w:t xml:space="preserve">{changed}</w:t></w:r></w:ins></w:p></w:body></w:document>"#
        );
        let document =
            rdocx_oxml::CT_Document::from_xml(xml.as_bytes()).expect("revision document parses");
        let mut input = make_input_with_text("");
        input.document = document;
        let accepted = Engine::new_deterministic()
            .expect("bundled fonts load")
            .layout(&input)
            .expect("accepted document lays out");
        input.revision_view = RevisionView::Tracked;
        let output = Engine::new_deterministic()
            .expect("bundled fonts load")
            .layout(&input)
            .expect("tracked document lays out");
        let geometry = PageGeometry::default();
        assert!(output.pages.len() > 1);
        assert_eq!(accepted.pages.len(), output.pages.len());
        for (accepted_page, page) in accepted.pages.iter().zip(&output.pages) {
            let accepted_text = accepted_page
                .elements
                .iter()
                .filter_map(|element| match element {
                    PositionedElement::Text(text) => Some(text),
                    _ => None,
                })
                .collect::<Vec<_>>();
            let tracked_text = page
                .elements
                .iter()
                .filter_map(|element| match element {
                    PositionedElement::Text(text) => Some(text),
                    _ => None,
                })
                .collect::<Vec<_>>();
            assert_eq!(accepted_text, tracked_text);
            let bars = page
                .elements
                .iter()
                .filter_map(|element| match element {
                    PositionedElement::Line {
                        start,
                        end,
                        width,
                        dash_pattern,
                        ..
                    } if (*width - 1.5).abs() < f64::EPSILON
                        && dash_pattern.is_none()
                        && start.x == end.x
                        && (start.x < geometry.margin_left
                            || start.x > geometry.page_width - geometry.margin_right) =>
                    {
                        Some((*start, *end))
                    }
                    _ => None,
                })
                .collect::<Vec<_>>();
            assert_eq!(bars.len(), 1, "page {}", page.page_number);
            let (start, end) = bars[0];
            assert!(start.x.is_finite() && start.y.is_finite() && end.y.is_finite());
            assert!(end.y > start.y);
            if page.page_number.is_multiple_of(2) {
                assert!(start.x < geometry.margin_left);
            } else {
                assert!(start.x > geometry.page_width - geometry.margin_right);
            }
        }
    }

    fn page_change_bar_count(page: &PageFrame) -> usize {
        let geometry = PageGeometry::default();
        page.elements
            .iter()
            .filter(|element| {
                matches!(element, PositionedElement::Line { start, end, width, .. }
                    if (*width - 1.5).abs() < f64::EPSILON
                        && start.x == end.x
                        && (start.x < geometry.margin_left
                            || start.x > geometry.page_width - geometry.margin_right))
            })
            .count()
    }

    #[test]
    fn tracked_header_paragraph_draws_a_change_bar() {
        use rdocx_oxml::header_footer::{CT_HdrFtr, HdrFtrRef, HdrFtrType};

        let mut input = make_input_with_text("body");
        input.revision_view = RevisionView::Tracked;
        input.document.body.sect_pr = Some(CT_SectPr::default_letter());
        input
            .document
            .body
            .sect_pr
            .as_mut()
            .expect("section properties")
            .header_refs
            .push(HdrFtrRef {
                hdr_ftr_type: HdrFtrType::Default,
                rel_id: "rIdHeader".to_owned(),
            });
        let header = CT_HdrFtr::from_xml(
            br#"<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:ins w:id="1" w:author="Ada"><w:r><w:t>changed header</w:t></w:r></w:ins></w:p></w:hdr>"#,
        )
        .expect("header parses");
        input.headers.insert("rIdHeader".to_owned(), header);

        let output = Engine::new_deterministic()
            .expect("bundled fonts load")
            .layout(&input)
            .expect("tracked header lays out");
        assert_eq!(page_change_bar_count(&output.pages[0]), 1);
    }

    #[test]
    fn tracked_note_paragraph_draws_a_change_bar() {
        let mut input = make_input_with_footnote(&["plain"]);
        input.revision_view = RevisionView::Tracked;
        let changed_note = rdocx_oxml::CT_Document::from_xml(
            br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:ins w:id="1" w:author="Ada"><w:r><w:t>added note</w:t></w:r></w:ins><w:del w:id="2" w:author="Ben"><w:r><w:delText>removed note</w:delText></w:r></w:del></w:p></w:body></w:document>"#,
        )
        .expect("note paragraph parses");
        let BodyContent::Paragraph(paragraph) = &changed_note.body.content[0] else {
            panic!("expected paragraph");
        };
        input.footnotes.as_mut().expect("footnote stream").footnotes[0].paragraphs =
            vec![paragraph.clone()];

        let output = Engine::new_deterministic()
            .expect("bundled fonts load")
            .layout(&input)
            .expect("tracked note lays out");
        assert_eq!(page_change_bar_count(&output.pages[0]), 1);
        let decoration_widths = output.pages[0]
            .elements
            .iter()
            .filter_map(|element| match element {
                PositionedElement::Line {
                    start, end, width, ..
                } if start.y == end.y && (*width - 0.5).abs() > f64::EPSILON => Some(*width),
                _ => None,
            })
            .collect::<Vec<_>>();
        assert!(
            decoration_widths
                .iter()
                .any(|width| (*width - 11.0 / 18.0).abs() < 0.001),
            "tracked insertion underline missing: {decoration_widths:?}"
        );
        assert!(
            decoration_widths
                .iter()
                .any(|width| (*width - 11.0 / 24.0).abs() < 0.001),
            "tracked deletion strike missing: {decoration_widths:?}"
        );
    }

    #[test]
    fn property_only_revisions_mark_the_tracked_paragraph() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:pPrChange w:id="1" w:author="Ada"><w:pPr><w:jc w:val="right"/></w:pPr></w:pPrChange></w:pPr><w:r><w:t>current</w:t></w:r></w:p></w:body></w:document>"#;
        let document = rdocx_oxml::CT_Document::from_xml(xml).expect("revision document parses");
        let BodyContent::Paragraph(paragraph) = &document.body.content[0] else {
            panic!("expected paragraph");
        };
        assert!(paragraph_has_visible_revision(paragraph));
    }

    #[test]
    fn empty_revision_wrappers_do_not_mark_the_tracked_paragraph() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:ins w:id="1" w:author="Ada"/></w:p><w:p><w:ins w:id="2" w:author="Ben"><w:r><w:t/></w:r></w:ins></w:p></w:body></w:document>"#;
        let document = rdocx_oxml::CT_Document::from_xml(xml).expect("revision document parses");
        for content in &document.body.content {
            let BodyContent::Paragraph(paragraph) = content else {
                panic!("expected paragraph");
            };
            assert!(!paragraph_has_visible_revision(paragraph));
        }
    }

    fn make_input_with_text(text: &str) -> LayoutInput {
        let mut doc = rdocx_oxml::document::CT_Document::new();
        let mut p = CT_P::new();
        p.add_run(text);
        doc.body.add_paragraph(p);

        LayoutInput {
            revision_view: crate::input::RevisionView::Accepted,
            document: doc,
            styles: CT_Styles::new_default(),
            numbering: None,
            headers: HashMap::new(),
            footers: HashMap::new(),
            images: HashMap::new(),
            charts: HashMap::new(),
            chart_theme: oxml_drawing::theme::CT_OfficeStyleSheet::office_default(),
            chart_color_map: oxml_drawing::color::ColorMap::default(),
            core_properties: None,
            hyperlink_urls: HashMap::new(),
            footnotes: None,
            endnotes: None,
            theme: None,
            fonts: Vec::new(),
        }
    }

    #[test]
    fn warm_relayout_matches_cold_and_rebuilds_only_changed_safe_paragraphs() {
        let mut input = make_input_with_text("first cache-safe paragraph");
        for text in ["second cache-safe paragraph", "third cache-safe paragraph"] {
            let mut paragraph = CT_P::new();
            paragraph.add_run(text);
            input.document.body.add_paragraph(paragraph);
        }

        let mut warm_engine = Engine::new_deterministic().expect("bundled fonts load");
        let cold = warm_engine
            .layout_with_provenance(&input)
            .expect("cold layout succeeds");
        let after_cold = warm_engine.paragraph_cache_counts();

        let BodyContent::Paragraph(changed) = &mut input.document.body.content[1] else {
            panic!("second body item is a paragraph");
        };
        changed.runs[0].content = vec![RunContent::Text(rdocx_oxml::text::CT_Text::new(
            "changed cache-safe paragraph",
        ))];

        let warm = warm_engine
            .layout_with_provenance(&input)
            .expect("warm relayout succeeds");
        let after_warm = warm_engine.paragraph_cache_counts();
        let cold_after_edit = Engine::new_deterministic()
            .expect("bundled fonts load")
            .layout_with_provenance(&input)
            .expect("independent cold relayout succeeds");

        assert_eq!(format!("{:?}", warm.0), format!("{:?}", cold_after_edit.0));
        assert_eq!(warm.1, cold_after_edit.1);
        assert_eq!(after_cold, (0, 3));
        assert_eq!(after_warm, (2, 4));
        assert_ne!(output_text(&cold.0), output_text(&warm.0));
    }

    #[test]
    fn warm_relayout_rebinds_font_tables_and_ids_to_the_current_result() {
        let mut input = make_input_with_text("font identity changes");
        {
            let BodyContent::Paragraph(paragraph) = &mut input.document.body.content[0] else {
                panic!("body paragraph");
            };
            paragraph.runs[0]
                .properties
                .get_or_insert_default()
                .font_ascii = Some("Carlito".to_owned());
        }

        let mut warm_engine = Engine::new_deterministic().expect("bundled fonts load");
        warm_engine.layout(&input).expect("prime warm font state");
        let BodyContent::Paragraph(paragraph) = &mut input.document.body.content[0] else {
            panic!("body paragraph");
        };
        paragraph.runs[0]
            .properties
            .get_or_insert_default()
            .font_ascii = Some("Caladea".to_owned());

        let warm = warm_engine.layout(&input).expect("warm relayout succeeds");
        let cold = Engine::new_deterministic()
            .expect("bundled fonts load")
            .layout(&input)
            .expect("cold relayout succeeds");
        assert_eq!(format!("{warm:?}"), format!("{cold:?}"));
        assert_eq!(format!("{:?}", warm.fonts), format!("{:?}", cold.fonts));
    }

    #[test]
    fn warm_relayout_canonicalizes_the_same_fonts_in_new_resolution_order() {
        let mut input = make_input_with_text("first family");
        let BodyContent::Paragraph(first) = &mut input.document.body.content[0] else {
            panic!("body paragraph");
        };
        first.runs[0].properties.get_or_insert_default().font_ascii = Some("Carlito".to_owned());
        let mut second = CT_P::new();
        second
            .add_run("second family")
            .properties
            .get_or_insert_default()
            .font_ascii = Some("Caladea".to_owned());
        input.document.body.add_paragraph(second);

        let mut warm_engine = Engine::new_deterministic().expect("bundled fonts load");
        warm_engine.layout(&input).expect("prime original order");
        input.document.body.content.swap(0, 1);

        let warm = warm_engine.layout(&input).expect("warm reordered layout");
        let cold = Engine::new_deterministic()
            .expect("bundled fonts load")
            .layout(&input)
            .expect("cold reordered layout");
        assert_eq!(format!("{warm:?}"), format!("{cold:?}"));
        assert_eq!(format!("{:?}", warm.fonts), format!("{:?}", cold.fonts));
    }

    #[test]
    fn shared_layout_context_changes_cannot_serve_stale_blocks() {
        let mut input = make_input_with_text("context-sensitive cache identity");
        let mut warm_engine = Engine::new_deterministic().expect("bundled fonts load");
        warm_engine.layout(&input).expect("prime context cache");

        let normal = input
            .styles
            .styles
            .iter_mut()
            .find(|style| style.is_default)
            .expect("default style");
        normal.rpr.get_or_insert_default().font_ascii = Some("Caladea".to_owned());
        let warm = warm_engine.layout(&input).expect("warm style mutation");
        let cold = Engine::new_deterministic()
            .expect("bundled fonts load")
            .layout(&input)
            .expect("cold style mutation");
        assert_eq!(format!("{warm:?}"), format!("{cold:?}"));

        input.numbering = Some(rdocx_oxml::numbering::CT_Numbering::new());
        let warm = warm_engine.layout(&input).expect("warm numbering mutation");
        let cold = Engine::new_deterministic()
            .expect("bundled fonts load")
            .layout(&input)
            .expect("cold numbering mutation");
        assert_eq!(format!("{warm:?}"), format!("{cold:?}"));

        input.theme = Some(rdocx_oxml::theme::Theme::default());
        let warm = warm_engine.layout(&input).expect("warm theme mutation");
        let cold = Engine::new_deterministic()
            .expect("bundled fonts load")
            .layout(&input)
            .expect("cold theme mutation");
        assert_eq!(format!("{warm:?}"), format!("{cold:?}"));

        input
            .hyperlink_urls
            .insert("rIdLink".to_owned(), "https://example.com".to_owned());
        input.images.insert(
            "rIdImage".to_owned(),
            crate::input::ImageData {
                data: vec![1, 2, 3],
                content_type: "image/png".to_owned(),
            },
        );
        let warm = warm_engine
            .layout(&input)
            .expect("warm relationship and image mutation");
        let cold = Engine::new_deterministic()
            .expect("bundled fonts load")
            .layout(&input)
            .expect("cold relationship and image mutation");
        assert_eq!(format!("{warm:?}"), format!("{cold:?}"));

        input.fonts.push(oxml_layout::FontFile {
            family: "Embedded".to_owned(),
            data: oxml_layout::bundled_fonts::bundled_font_data()[0]
                .1
                .to_vec(),
        });
        let warm = warm_engine.layout(&input).expect("warm font mutation");
        let cold = Engine::new_deterministic()
            .expect("bundled fonts load")
            .layout(&input)
            .expect("cold font mutation");
        assert_eq!(format!("{warm:?}"), format!("{cold:?}"));

        let contextual = rdocx_oxml::CT_Document::from_xml(
            br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:hyperlink r:id="rIdLink"><w:r><w:t>link</w:t></w:r></w:hyperlink></w:p><w:p><w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p></w:body></w:document>"#,
        )
        .expect("contextual paragraphs parse");
        for content in &contextual.body.content {
            let BodyContent::Paragraph(paragraph) = content else {
                continue;
            };
            assert!(!paragraph_is_cache_safe(paragraph, &input.styles));
        }
    }

    #[test]
    fn alternate_content_drawings_bypass_paragraph_reuse() {
        let document = rdocx_oxml::CT_Document::from_xml(
            br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><w:body><w:p><w:r><w:t>ordinary text</w:t></w:r><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor behindDoc="0"><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="914400" cy="457200"/><a:graphic><a:graphicData><wps:wsp><wps:spPr><a:prstGeom prst="rect"/></wps:spPr></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice></mc:AlternateContent></w:r></w:p></w:body></w:document>"#,
        )
        .expect("AlternateContent drawing parses");
        let BodyContent::Paragraph(paragraph) = &document.body.content[0] else {
            panic!("body paragraph");
        };
        assert!(!paragraph.runs[1].alt_drawings.is_empty());
        assert!(!paragraph_is_cache_safe(
            paragraph,
            &CT_Styles::new_default()
        ));
    }

    #[test]
    fn warm_provenance_rebinds_to_current_word_source_nodes() {
        let mut input = make_input_with_text("first paragraph");
        for text in ["second paragraph", "third paragraph"] {
            let mut paragraph = CT_P::new();
            paragraph.add_run(text);
            input.document.body.add_paragraph(paragraph);
        }
        let mut engine = Engine::new_deterministic().expect("bundled fonts load");
        engine
            .layout_with_provenance(&input)
            .expect("prime paragraph cache");

        let moved = input.document.body.content.remove(2);
        input.document.body.content.insert(0, moved);
        let mut inserted = CT_P::new();
        inserted.add_run("new paragraph");
        input
            .document
            .body
            .content
            .insert(1, BodyContent::Paragraph(inserted));
        let (layout, sources) = engine
            .layout_with_provenance(&input)
            .expect("warm provenance layout");

        for page in &layout.pages {
            oxml_layout::walk(&page.elements, &mut |element, _| {
                let PositionedElement::Text(run) = element else {
                    return;
                };
                let Some(span) = run.source else {
                    return;
                };
                let path = &sources[span.node.get() as usize - 1];
                assert_eq!(path.story, WordStory::Document);
                let BodyContent::Paragraph(paragraph) =
                    &input.document.body.content[path.children[0]]
                else {
                    panic!("source path resolves to a body paragraph");
                };
                let text = paragraph.text();
                let resolved = text
                    .chars()
                    .skip(span.char_start as usize)
                    .take((span.char_end - span.char_start) as usize)
                    .collect::<String>();
                assert_eq!(resolved, run.text);
            });
        }
        assert_eq!(engine.paragraph_cache_counts(), (3, 4));
    }

    #[test]
    fn cold_and_warm_diagnostics_are_identical() {
        let mut input = make_input_with_text("");
        input.document = rdocx_oxml::CT_Document::from_xml(
            br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>cache-safe prefix</w:t></w:r></w:p><w:p><w:fldSimple w:instr="REF missing"><w:r><w:t>stored</w:t></w:r></w:fldSimple></w:p></w:body></w:document>"#,
        )
        .expect("diagnostic document parses");
        let mut engine = Engine::new_deterministic().expect("bundled fonts load");
        let cold = engine.layout(&input).expect("cold layout succeeds");
        let warm = engine.layout(&input).expect("warm layout succeeds");
        assert!(!cold.diagnostics.is_empty());
        assert_eq!(cold.diagnostics, warm.diagnostics);
        assert_eq!(engine.paragraph_cache_counts(), (1, 1));

        let (valid_family, valid_bytes) = oxml_layout::bundled_fonts::bundled_font_data()[0];
        let (invalid_family, invalid_source) = oxml_layout::bundled_fonts::bundled_font_data()[4];
        let mut invalid_bytes = invalid_source.to_vec();
        let table_count = u16::from_be_bytes([invalid_bytes[4], invalid_bytes[5]]) as usize;
        let head_offset = (0..table_count)
            .find_map(|table| {
                let record = 12 + table * 16;
                (&invalid_bytes[record..record + 4] == b"head").then(|| {
                    u32::from_be_bytes(
                        invalid_bytes[record + 8..record + 12]
                            .try_into()
                            .expect("head offset"),
                    ) as usize
                })
            })
            .expect("font has head table");
        invalid_bytes[head_offset + 18..head_offset + 20].copy_from_slice(&0u16.to_be_bytes());

        let mut failing = Engine::with_font_manager(FontManager::new_with_fonts(vec![(
            valid_family.to_owned(),
            valid_bytes.to_vec(),
        )]));
        let mut failing_input = make_input_with_text("cache-safe successful prefix");
        let BodyContent::Paragraph(prefix) = &mut failing_input.document.body.content[0] else {
            panic!("prefix paragraph");
        };
        prefix.runs[0].properties.get_or_insert_default().font_ascii =
            Some(valid_family.to_owned());
        let mut later = CT_P::new();
        later
            .add_run("late font failure")
            .properties
            .get_or_insert_default()
            .font_ascii = Some(invalid_family.to_owned());
        failing_input.document.body.add_paragraph(later);
        failing_input.fonts.push(oxml_layout::FontFile {
            family: invalid_family.to_owned(),
            data: invalid_bytes,
        });
        assert!(failing.layout(&failing_input).is_err());
        assert!(failing.paragraph_cache.is_empty());
        assert_eq!(failing.paragraph_cache_counts(), (0, 1));
    }

    #[test]
    fn paragraph_relayout_cache_is_bounded() {
        let mut input = make_input_with_text("bounded paragraph 0");
        for index in 1..(PARAGRAPH_CACHE_MAX_ENTRIES + 20) {
            let mut paragraph = CT_P::new();
            paragraph.add_run(&format!("bounded paragraph {index}"));
            input.document.body.add_paragraph(paragraph);
        }
        let mut engine = Engine::new_deterministic().expect("bundled fonts load");
        engine.layout(&input).expect("bounded layout succeeds");
        assert!(engine.paragraph_cache.len() <= PARAGRAPH_CACHE_MAX_ENTRIES);
        assert!(engine.paragraph_cache_bytes <= PARAGRAPH_CACHE_MAX_BYTES);
        assert!(engine.pending_paragraph_cache_peak_entries <= PARAGRAPH_CACHE_MAX_ENTRIES);
        assert!(engine.pending_paragraph_cache_peak_bytes <= PARAGRAPH_CACHE_MAX_BYTES);
    }

    #[test]
    fn transactional_paragraph_staging_is_bounded_before_publication() {
        let mut input = make_input_with_text("staged paragraph 0");
        for index in 1..(PARAGRAPH_CACHE_MAX_ENTRIES * 2) {
            let mut paragraph = CT_P::new();
            paragraph.add_run(&format!("staged paragraph {index}"));
            input.document.body.add_paragraph(paragraph);
        }
        let mut engine = Engine::new_deterministic().expect("bundled fonts load");
        engine
            .layout(&input)
            .expect("transactional layout succeeds");
        // With the larger entry cap the byte ceiling can bind first; either
        // way staging must stay bounded well below the paragraph count.
        assert!(engine.pending_paragraph_cache_peak_entries <= PARAGRAPH_CACHE_MAX_ENTRIES);
        assert!(
            engine.pending_paragraph_cache_peak_entries < PARAGRAPH_CACHE_MAX_ENTRIES * 2,
            "staging must evict, got {}",
            engine.pending_paragraph_cache_peak_entries
        );
        assert!(engine.pending_paragraph_cache_peak_bytes <= PARAGRAPH_CACHE_MAX_BYTES);
    }

    #[test]
    fn paragraph_relayout_cache_enforces_the_reflow_byte_ceiling() {
        let input = make_input_with_text("reflow accounting template");
        let mut engine = Engine::new_deterministic().expect("bundled fonts load");
        engine.layout(&input).expect("template layout succeeds");
        let template = engine
            .paragraph_cache
            .front()
            .expect("safe paragraph cached")
            .block
            .clone();
        let mut retained = match &template.lines[0].items[0] {
            LineItem::Text(text) => text.clone(),
            other => panic!("expected text line item, got {other:?}"),
        };
        retained.advances = vec![0.0; PARAGRAPH_CACHE_MAX_BYTES / 8 + 1];

        let mut block = template;
        block.reflow = Some(Box::new(block::ParagraphReflow {
            items: vec![InlineItem::Text(retained)],
            params: oxml_layout::LineBreakParams::default(),
        }));
        let BodyContent::Paragraph(paragraph) = &input.document.body.content[0] else {
            panic!("body paragraph");
        };
        let bytes = paragraph_cache_entry_bytes(paragraph, &block, &[], 0);
        assert!(bytes > PARAGRAPH_CACHE_MAX_BYTES);

        engine.paragraph_cache.clear();
        engine.paragraph_cache_bytes = 0;
        engine.publish_paragraph_cache_entry(ParagraphCacheEntry {
            fp: 0,
            key: ParagraphCacheKey {
                paragraph: paragraph.clone(),
                content_width_bits: PageGeometry::default().content_width().to_bits(),
                revision_view: RevisionView::Accepted,
            },
            block,
            diagnostics: Vec::new(),
            font_trace: Vec::new(),
            bytes,
        });
        assert!(engine.paragraph_cache.is_empty());
        assert_eq!(engine.paragraph_cache_bytes, 0);
    }

    #[test]
    fn tab_heavy_paragraph_in_wrapping_document_counts_reflow_parameter_buffers() {
        use rdocx_oxml::borders::{CT_TabStop, CT_Tabs};
        use rdocx_oxml::drawing::{AnchorAlignH, WrapType};
        use rdocx_oxml::shared::ST_TabJc;
        use rdocx_oxml::units::Twips;

        let mut input =
            make_wrapping_document(WrapType::Square, Some(AnchorAlignH::Left), 100.0, 40.0, 5.0);
        let retained_per_stop =
            std::mem::size_of::<CT_TabStop>() + std::mem::size_of::<oxml_layout::TabStop>();
        let stop_count = PARAGRAPH_CACHE_MAX_BYTES / retained_per_stop + 1;
        let mut paragraph = CT_P::new();
        paragraph.properties = Some(CT_PPr {
            tabs: Some(CT_Tabs {
                tabs: (0..stop_count)
                    .map(|_| CT_TabStop::new(ST_TabJc::Left, Twips(720)))
                    .collect(),
            }),
            ..CT_PPr::default()
        });
        paragraph.add_run("cache-safe paragraph with many owned tab definitions");
        assert!(paragraph_is_cache_safe(&paragraph, &input.styles));
        input.document.body.add_paragraph(paragraph);

        let mut engine = Engine::new_deterministic().expect("bundled fonts load");
        engine.layout(&input).expect("tab-heavy layout succeeds");
        assert!(engine.paragraph_cache.is_empty());
        assert_eq!(engine.paragraph_cache_bytes, 0);
    }

    #[test]
    fn paragraph_relayout_cache_counts_all_reflow_parameter_vectors() {
        let input = make_input_with_text("reflow parameter accounting template");
        let mut engine = Engine::new_deterministic().expect("bundled fonts load");
        engine.layout(&input).expect("template layout succeeds");
        let mut block = engine
            .paragraph_cache
            .front()
            .expect("safe paragraph cached")
            .block
            .clone();
        let BodyContent::Paragraph(paragraph) = &input.document.body.content[0] else {
            panic!("body paragraph");
        };
        let baseline = paragraph_cache_entry_bytes(paragraph, &block, &[], 0);
        let reflow = block.reflow.as_mut().expect("cache retains reflow inputs");
        reflow.params.tab_stops = vec![
            oxml_layout::TabStop {
                pos_pt: 36.0,
                align: oxml_layout::TabAlign::Left,
                leader: None,
            };
            3
        ];
        reflow.params.line_prefix_widths = vec![0.0; 5];
        reflow.params.line_suffix_widths = vec![0.0; 7];
        let with_parameters = paragraph_cache_entry_bytes(paragraph, &block, &[], 0);
        let expected =
            3 * std::mem::size_of::<oxml_layout::TabStop>() + 12 * std::mem::size_of::<f64>();
        assert_eq!(with_parameters - baseline, expected);
    }

    #[test]
    fn paragraph_relayout_cache_counts_fixed_storage_in_owned_keys() {
        let input = make_input_with_text("key accounting template");
        let mut engine = Engine::new_deterministic().expect("bundled fonts load");
        engine.layout(&input).expect("template layout succeeds");
        let block = engine
            .paragraph_cache
            .front()
            .expect("safe paragraph cached")
            .block
            .clone();

        let mut paragraph = CT_P::new();
        let mut run = CT_R::new("");
        let content_count = PARAGRAPH_CACHE_MAX_BYTES / std::mem::size_of::<RunContent>() + 1;
        run.content = std::iter::repeat_n(RunContent::Tab, content_count).collect();
        paragraph.runs.push(run);
        assert!(paragraph_is_cache_safe(&paragraph, &input.styles));
        let bytes = paragraph_cache_entry_bytes(&paragraph, &block, &[], 0);
        assert!(bytes > PARAGRAPH_CACHE_MAX_BYTES);

        engine.paragraph_cache.clear();
        engine.paragraph_cache_bytes = 0;
        engine.publish_paragraph_cache_entry(ParagraphCacheEntry {
            fp: 0,
            key: ParagraphCacheKey {
                paragraph,
                content_width_bits: PageGeometry::default().content_width().to_bits(),
                revision_view: RevisionView::Accepted,
            },
            block,
            diagnostics: Vec::new(),
            font_trace: Vec::new(),
            bytes,
        });
        assert!(engine.paragraph_cache.is_empty());
        assert_eq!(engine.paragraph_cache_bytes, 0);
    }

    #[test]
    fn every_sourced_glyph_run_resolves_to_its_exact_word_text() {
        use rdocx_oxml::document::CT_SectPr;
        use rdocx_oxml::footnotes::{CT_Footnote, CT_Footnotes, NoteType};
        use rdocx_oxml::header_footer::{CT_HdrFtr, HdrFtrRef};

        let body_text = "Body ASCII 🚀 界 wraps across several exact source slices ".repeat(5);
        let mut input = make_input_with_text(&body_text);

        let mut outer = CT_Tbl::new();
        let mut outer_row = CT_Row::new();
        let mut outer_cell = CT_Tc::new();
        outer_cell.paragraphs_mut()[0].add_run("outer cell");
        let mut nested = CT_Tbl::new();
        let mut nested_row = CT_Row::new();
        let mut nested_cell = CT_Tc::new();
        nested_cell.paragraphs_mut()[0].add_run("nested cell");
        nested_row.cells.push(nested_cell);
        nested.rows.push(nested_row);
        outer_cell.content.push(CellContent::Table(nested));
        outer_row.cells.push(outer_cell);
        outer.rows.push(outer_row);
        input.document.body.add_table(outer);

        let mut references = CT_P::new();
        let mut reference_run = CT_R::new("");
        reference_run.content = vec![
            RunContent::FootnoteRef { id: 4 },
            RunContent::EndnoteRef { id: 9 },
        ];
        references.runs.push(reference_run);
        input.document.body.add_paragraph(references);

        let mut header = CT_HdrFtr::new();
        let mut header_paragraph = CT_P::new();
        header_paragraph.add_run("header text");
        header.paragraphs.push(header_paragraph);
        input.headers.insert("rIdHeader".to_owned(), header);

        let mut footer = CT_HdrFtr::new();
        let mut footer_paragraph = CT_P::new();
        footer_paragraph.add_run("footer text");
        footer.paragraphs.push(footer_paragraph);
        input.footers.insert("rIdFooter".to_owned(), footer);

        let mut section = CT_SectPr::default_letter();
        section.header_refs.push(HdrFtrRef {
            hdr_ftr_type: HdrFtrType::Default,
            rel_id: "rIdHeader".to_owned(),
        });
        section.footer_refs.push(HdrFtrRef {
            hdr_ftr_type: HdrFtrType::Default,
            rel_id: "rIdFooter".to_owned(),
        });
        input.document.body.sect_pr = Some(section);

        let mut footnote_paragraph = CT_P::new();
        footnote_paragraph.add_run("footnote text");
        input.footnotes = Some(CT_Footnotes {
            footnotes: vec![CT_Footnote {
                id: 4,
                note_type: NoteType::Normal,
                paragraphs: vec![footnote_paragraph],
            }],
        });
        let mut endnote_paragraph = CT_P::new();
        endnote_paragraph.add_run("endnote text");
        input.endnotes = Some(CT_Footnotes {
            footnotes: vec![CT_Footnote {
                id: 9,
                note_type: NoteType::Normal,
                paragraphs: vec![endnote_paragraph],
            }],
        });

        let expected = HashMap::from([
            (
                WordSourcePath {
                    story: WordStory::Document,
                    children: vec![0],
                },
                body_text,
            ),
            (
                WordSourcePath {
                    story: WordStory::Document,
                    children: vec![1, 0, 0, 0],
                },
                "outer cell".to_owned(),
            ),
            (
                WordSourcePath {
                    story: WordStory::Document,
                    children: vec![1, 0, 0, 1, 0, 0, 0],
                },
                "nested cell".to_owned(),
            ),
            (
                WordSourcePath {
                    story: WordStory::Header {
                        relationship_id: "rIdHeader".to_owned(),
                    },
                    children: vec![0],
                },
                "header text".to_owned(),
            ),
            (
                WordSourcePath {
                    story: WordStory::Footer {
                        relationship_id: "rIdFooter".to_owned(),
                    },
                    children: vec![0],
                },
                "footer text".to_owned(),
            ),
            (
                WordSourcePath {
                    story: WordStory::Footnote { id: 4 },
                    children: vec![0],
                },
                "footnote text".to_owned(),
            ),
            (
                WordSourcePath {
                    story: WordStory::Endnote { id: 9 },
                    children: vec![0],
                },
                "endnote text".to_owned(),
            ),
        ]);

        let result = crate::layout_document_deterministic_with_provenance(&input)
            .expect("layout with provenance");
        let mut seen = std::collections::HashSet::new();
        for run in result.layout.pages.iter().flat_map(|page| {
            page.elements.iter().filter_map(|element| match element {
                PositionedElement::Text(run) => Some(run),
                _ => None,
            })
        }) {
            let Some(span) = run.source else {
                continue;
            };
            let path = result.source_node(span.node).expect("source node resolves");
            let source_text = expected.get(path).expect("source path belongs to fixture");
            let selected = source_text
                .chars()
                .skip(span.char_start as usize)
                .take((span.char_end - span.char_start) as usize)
                .collect::<String>();
            assert_eq!(selected, run.text, "mismatch at {path:?}");
            seen.insert(path.clone());
        }
        assert_eq!(
            seen.len(),
            expected.len(),
            "every supported story is sourced"
        );
        for path in expected.keys() {
            assert!(seen.contains(path), "missing source path {path:?}");
        }
    }

    #[test]
    fn repeated_text_and_repeated_stories_keep_distinct_source_nodes() {
        use rdocx_oxml::header_footer::{CT_HdrFtr, HdrFtrRef};

        let repeated = "duplicate phrase ".repeat(220);
        let mut input = make_input_with_text(&repeated);
        let mut second = CT_P::new();
        second.add_run(&repeated);
        input.document.body.add_paragraph(second);

        let mut header = CT_HdrFtr::new();
        let mut paragraph = CT_P::new();
        paragraph.add_run("repeated header");
        header.paragraphs.push(paragraph);
        input.headers.insert("rIdRepeated".to_owned(), header);
        input
            .document
            .body
            .sect_pr
            .as_mut()
            .expect("default section")
            .header_refs
            .push(HdrFtrRef {
                hdr_ftr_type: HdrFtrType::Default,
                rel_id: "rIdRepeated".to_owned(),
            });

        let result = crate::layout_document_deterministic_with_provenance(&input)
            .expect("layout repeated stories");
        assert!(result.layout.pages.len() > 1, "header must be reused");
        let mut first_body = std::collections::HashSet::new();
        let mut second_body = std::collections::HashSet::new();
        let mut header_nodes = std::collections::HashSet::new();
        let mut header_runs = 0usize;
        for run in result.layout.pages.iter().flat_map(|page| {
            page.elements.iter().filter_map(|element| match element {
                PositionedElement::Text(run) => Some(run),
                _ => None,
            })
        }) {
            let Some(source) = run.source else {
                continue;
            };
            match result.source_node(source.node).expect("source resolves") {
                WordSourcePath {
                    story: WordStory::Document,
                    children,
                } if children == &[0] => {
                    first_body.insert(source.node);
                }
                WordSourcePath {
                    story: WordStory::Document,
                    children,
                } if children == &[1] => {
                    second_body.insert(source.node);
                }
                WordSourcePath {
                    story: WordStory::Header { relationship_id },
                    children,
                } if relationship_id == "rIdRepeated" && children == &[0] => {
                    header_nodes.insert(source.node);
                    header_runs += 1;
                }
                _ => {}
            }
        }
        assert_eq!(first_body.len(), 1);
        assert_eq!(second_body.len(), 1);
        assert_ne!(
            first_body, second_body,
            "duplicate paragraphs must not alias"
        );
        assert_eq!(header_nodes.len(), 1, "repeated header reuses one node");
        assert!(header_runs > 1, "header must be emitted more than once");
    }

    #[test]
    fn accepted_and_tracked_views_record_projection_local_ranges() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>A</w:t></w:r><w:del w:id="1" w:author="Ada"><w:r><w:delText>B</w:delText></w:r></w:del><w:ins w:id="2" w:author="Ada"><w:r><w:t>C</w:t></w:r></w:ins></w:p></w:body></w:document>"#;
        let document = rdocx_oxml::CT_Document::from_xml(xml).expect("revision XML parses");
        let BodyContent::Paragraph(paragraph) = &document.body.content[0] else {
            panic!("expected paragraph");
        };

        for (view, expected) in [
            (RevisionView::Accepted, "AC"),
            (RevisionView::Tracked, "ABC"),
        ] {
            assert_eq!(projected_paragraph_text(paragraph, view), expected);
            let mut input = make_input_with_text("");
            input.document = document.clone();
            input.revision_view = view;
            let result = crate::layout_document_deterministic_with_provenance(&input)
                .expect("revision layout with provenance");
            assert_eq!(result.revision_view, view);
            let mut selected = String::new();
            for run in result.layout.pages.iter().flat_map(|page| {
                page.elements.iter().filter_map(|element| match element {
                    PositionedElement::Text(run) => Some(run),
                    _ => None,
                })
            }) {
                let Some(span) = run.source else {
                    continue;
                };
                assert!(matches!(
                    result.source_node(span.node),
                    Some(WordSourcePath {
                        story: WordStory::Document,
                        children,
                    }) if children == &[0]
                ));
                let exact = expected
                    .chars()
                    .skip(span.char_start as usize)
                    .take((span.char_end - span.char_start) as usize)
                    .collect::<String>();
                assert_eq!(exact, run.text);
                selected.push_str(&run.text);
            }
            assert_eq!(selected, expected);
        }
    }

    #[test]
    fn field_projection_ownership_disambiguates_repeated_literal_ranges() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>DATE</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>a</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:body></w:document>"#;
        let document = rdocx_oxml::CT_Document::from_xml(xml).expect("complex field parses");
        let BodyContent::Paragraph(parsed) = &document.body.content[0] else {
            panic!("expected paragraph");
        };
        let [RunContent::Field(complex)] = parsed.runs[0].content.as_slice() else {
            panic!("expected projected complex field");
        };

        let cases = [
            (
                vec![
                    RunContent::Field(complex.clone()),
                    RunContent::Text(rdocx_oxml::text::CT_Text::new("a")),
                    RunContent::Field(Field::new("DATE", "a")),
                ],
                "aa",
                vec![("a", 1, 2)],
            ),
            (
                vec![
                    RunContent::Text(rdocx_oxml::text::CT_Text::new("a")),
                    RunContent::Field(complex.clone()),
                    RunContent::Text(rdocx_oxml::text::CT_Text::new("aa")),
                    RunContent::Field(Field::new("DATE", "a")),
                    RunContent::Text(rdocx_oxml::text::CT_Text::new("a")),
                ],
                "aaaaa",
                vec![("a", 0, 1), ("aa", 2, 4), ("a", 4, 5)],
            ),
        ];

        for (content, expected_projection, expected_literals) in cases {
            let mut input = make_input_with_text("");
            let BodyContent::Paragraph(paragraph) = &mut input.document.body.content[0] else {
                panic!("expected paragraph");
            };
            let mut run = CT_R::new("");
            run.content = content;
            assert_eq!(run.text(), expected_projection);
            paragraph.runs = vec![run];

            let result = crate::layout_document_deterministic_with_provenance(&input)
                .expect("mixed field layout");
            let sourced = result
                .layout
                .pages
                .iter()
                .flat_map(|page| &page.elements)
                .filter_map(|element| match element {
                    PositionedElement::Text(run) => run
                        .source
                        .map(|span| (run.text.as_str(), span.char_start, span.char_end)),
                    _ => None,
                })
                .collect::<Vec<_>>();
            assert_eq!(sourced, expected_literals);
        }
    }

    #[test]
    fn generated_or_transformed_text_remains_unattributed() {
        use rdocx_oxml::borders::{CT_TabStop, CT_Tabs};
        use rdocx_oxml::footnotes::{CT_Footnote, CT_Footnotes, NoteType};
        use rdocx_oxml::numbering::{
            CT_AbstractNum, CT_Lvl, CT_Num, CT_Numbering, ST_NumberFormat,
        };
        use rdocx_oxml::properties::CT_RPr;
        use rdocx_oxml::shared::{ST_TabJc, ST_TabLeader};
        use rdocx_oxml::units::Twips;

        let mut input = make_input_with_text("ordinary");

        let mut transformed = CT_P::new();
        let mut caps = CT_R::new("straße");
        caps.properties = Some(CT_RPr {
            caps: Some(true),
            ..Default::default()
        });
        transformed.runs.push(caps);
        input.document.body.add_paragraph(transformed);

        let mut generated = CT_P::new();
        generated.properties = Some(CT_PPr {
            tabs: Some(CT_Tabs {
                tabs: vec![CT_TabStop {
                    val: ST_TabJc::Left,
                    pos: Twips(3600),
                    leader: Some(ST_TabLeader::Dot),
                    source_occurrence: None,
                }],
            }),
            ..Default::default()
        });
        let mut generated_run = CT_R::new("");
        generated_run.content = vec![
            RunContent::Text(rdocx_oxml::text::CT_Text::new("left")),
            RunContent::Tab,
            RunContent::Text(rdocx_oxml::text::CT_Text::new("right")),
            RunContent::Field(Field::new("PAGE", "7")),
            RunContent::Text(rdocx_oxml::text::CT_Text::new("after")),
            RunContent::FootnoteRef { id: 4 },
        ];
        generated.runs.push(generated_run);
        input.document.body.add_paragraph(generated);

        let mut list = CT_P::new();
        list.properties = Some(CT_PPr {
            num_id: Some(1),
            num_ilvl: Some(0),
            ..Default::default()
        });
        list.add_run("listed");
        input.document.body.add_paragraph(list);
        let mut level = CT_Lvl::new(0);
        level.start = Some(1);
        level.num_fmt = Some(ST_NumberFormat::Decimal);
        level.lvl_text = Some("%1.".to_owned());
        let mut abstract_num = CT_AbstractNum::new(1);
        abstract_num.levels.push(level);
        input.numbering = Some(CT_Numbering {
            abstract_nums: vec![abstract_num],
            nums: vec![CT_Num {
                num_id: 1,
                abstract_num_id: 1,
                extra_xml: Vec::new(),
                extra_attributes: Vec::new(),
            }],
            root_attributes: Vec::new(),
            extra_xml: Vec::new(),
        });

        let mut note = CT_P::new();
        note.add_run("note body");
        input.footnotes = Some(CT_Footnotes {
            footnotes: vec![CT_Footnote {
                id: 4,
                note_type: NoteType::Normal,
                paragraphs: vec![note],
            }],
        });

        let result = crate::layout_document_deterministic_with_provenance(&input)
            .expect("generated text layout");
        let runs = result
            .layout
            .pages
            .iter()
            .flat_map(|page| &page.elements)
            .filter_map(|element| match element {
                PositionedElement::Text(run) => Some(run),
                _ => None,
            })
            .collect::<Vec<_>>();
        assert!(
            runs.iter()
                .any(|run| run.text == "ordinary" && run.source.is_some())
        );
        assert!(
            runs.iter()
                .any(|run| run.text == "after" && run.source.is_some())
        );
        assert!(
            runs.iter()
                .any(|run| run.text == "STRASSE" && run.source.is_none())
        );
        assert!(
            runs.iter()
                .any(|run| run.text == "1." && run.source.is_none())
        );
        assert!(runs.iter().any(|run| {
            !run.text.is_empty()
                && run.text.chars().all(|character| character == '.')
                && run.source.is_none()
        }));
        assert!(
            runs.iter()
                .any(|run| run.field_kind == Some(FieldKind::Page) && run.source.is_none())
        );
        assert!(
            runs.iter()
                .any(|run| run.note.is_some() && run.source.is_none())
        );
    }

    #[test]
    fn existing_low_level_layout_functions_keep_identical_output() {
        let input = make_input_with_text("compatibility 🚀 text that wraps ".repeat(30).as_str());
        let ordinary = crate::layout_document_deterministic(&input).expect("ordinary layout");
        let mut sourced = crate::layout_document_deterministic_with_provenance(&input)
            .expect("provenance layout")
            .into_layout_result();
        for page in &mut sourced.pages {
            let page = std::sync::Arc::make_mut(page);
            for element in &mut page.elements {
                if let PositionedElement::Text(run) = element {
                    run.source = None;
                }
            }
        }
        assert_eq!(format!("{ordinary:?}"), format!("{sourced:?}"));
    }

    #[test]
    fn caller_font_and_deterministic_provenance_variants_return_complete_maps() {
        let mut input = make_input_with_text("caller font provenance");
        let BodyContent::Paragraph(paragraph) = &mut input.document.body.content[0] else {
            panic!("expected paragraph");
        };
        paragraph.runs[0].properties = Some(rdocx_oxml::properties::CT_RPr {
            font_ascii: Some("Caller Carlito".to_owned()),
            font_hansi: Some("Caller Carlito".to_owned()),
            ..Default::default()
        });
        input.fonts.push(oxml_layout::FontFile {
            family: "Caller Carlito".to_owned(),
            data: include_bytes!("../../oxml-layout/fonts/Carlito-Regular.ttf").to_vec(),
        });

        let normal = crate::layout_document_with_provenance(&input).expect("caller font layout");
        let deterministic = crate::layout_document_deterministic_with_provenance(&input)
            .expect("deterministic caller font layout");
        for result in [&normal, &deterministic] {
            assert!(
                result
                    .layout
                    .fonts
                    .iter()
                    .any(|font| font.data.as_ref() == input.fonts[0].data.as_slice()),
                "the caller-provided font bytes shaped the result"
            );
            let runs = result
                .layout
                .pages
                .iter()
                .flat_map(|page| &page.elements)
                .filter_map(|element| match element {
                    PositionedElement::Text(run) if run.source.is_some() => Some(run),
                    _ => None,
                })
                .collect::<Vec<_>>();
            assert!(!runs.is_empty(), "caller-font text is sourced");
            assert_eq!(
                runs.iter().map(|run| run.text.as_str()).collect::<String>(),
                "caller font provenance"
            );
            for run in runs {
                let source = run.source.expect("run is sourced");
                assert!(matches!(
                    result.source_node(source.node),
                    Some(WordSourcePath {
                        story: WordStory::Document,
                        children,
                    }) if children == &[0]
                ));
            }
        }
    }

    #[test]
    fn layout_simple_document() {
        let input = make_input_with_text("Hello World");
        let result = Engine::new().layout(&input);
        // On systems without fonts, this may fail — that's OK
        if let Ok(result) = result {
            assert!(!result.pages.is_empty());
            assert_eq!(result.pages[0].page_number, 1);
            assert!((result.pages[0].width - 612.0).abs() < 0.01);
        }
    }

    #[test]
    fn layout_empty_document() {
        let mut doc = rdocx_oxml::document::CT_Document::new();
        doc.body.add_paragraph(CT_P::new());

        let input = LayoutInput {
            revision_view: crate::input::RevisionView::Accepted,
            document: doc,
            styles: CT_Styles::new_default(),
            numbering: None,
            headers: HashMap::new(),
            footers: HashMap::new(),
            images: HashMap::new(),
            charts: HashMap::new(),
            chart_theme: oxml_drawing::theme::CT_OfficeStyleSheet::office_default(),
            chart_color_map: oxml_drawing::color::ColorMap::default(),
            core_properties: None,
            hyperlink_urls: HashMap::new(),
            footnotes: None,
            endnotes: None,
            theme: None,
            fonts: Vec::new(),
        };

        let result = Engine::new().layout(&input);
        if let Ok(result) = result {
            assert_eq!(result.pages.len(), 1);
        }
    }

    #[test]
    fn empty_shapeless_anchor_keeps_the_pre_cutover_omission() {
        let input = make_input_with_text("");
        let mut paragraph = CT_P::new();
        paragraph.add_run("").content = vec![RunContent::Drawing(
            rdocx_oxml::drawing::CT_Drawing::anchor(rdocx_oxml::drawing::CT_Anchor::background(
                "", 914_400, 914_400,
            )),
        )];
        let mut font_manager = FontManager::new();
        let mut numbering_state = NumberingState::new();
        let mut diagnostics = Vec::new();
        let media = MediaRegistry::new(&input.images);

        let anchored = collect_anchored_drawings(
            &paragraph,
            &input.styles,
            &input,
            &media,
            &mut font_manager,
            &mut numbering_state,
            &mut diagnostics,
        )
        .expect("empty shapeless anchor collection should succeed");

        assert!(anchored.is_empty());
    }

    #[test]
    fn colliding_media_ids_keep_inline_and_anchored_image_bytes_distinct() {
        let mut input = make_input_with_text("");
        input.images.insert(
            "rIdInline".to_string(),
            ImageData {
                data: vec![1, 2, 3],
                content_type: "image/png".to_string(),
            },
        );
        input.images.insert(
            "rIdAnchor".to_string(),
            ImageData {
                data: vec![4, 5, 6],
                content_type: "image/jpeg".to_string(),
            },
        );

        let media = MediaRegistry::with_hasher(&input.images, |_| MediaId(7));
        let inline_id = media.id_for_relationship("rIdInline");
        let anchor_id = media.id_for_relationship("rIdAnchor");
        assert_ne!(inline_id, anchor_id);

        let line = oxml_layout::LayoutLine {
            items: vec![oxml_layout::LineItem::Image {
                width: 12.0,
                height: 10.0,
                media_id: inline_id,
            }],
            width: 12.0,
            ascent: 10.0,
            descent: 0.0,
            line_gap: 0.0,
            height: 10.0,
            indent_left: 0.0,
            available_width: 468.0,
            is_last: true,
        };
        let mut paragraph = block::build_paragraph_block(
            vec![line],
            0.0,
            0.0,
            None,
            None,
            0.0,
            0.0,
            None,
            false,
            false,
            false,
            true,
        );
        paragraph.anchored.push(block::AnchoredDrawing {
            behind_doc: false,
            rel_h: rdocx_oxml::drawing::ST_RelativeFromH::Page,
            off_h: 20.0,
            rel_v: rdocx_oxml::drawing::ST_RelativeFromV::Page,
            off_v: 20.0,
            width: 12.0,
            height: 10.0,
            wrap: rdocx_oxml::drawing::WrapType::None,
            dist_top: 0.0,
            dist_bottom: 0.0,
            dist_left: 0.0,
            dist_right: 0.0,
            align_h: None,
            align_v: None,
            content: block::AnchoredContent::Image {
                media_id: anchor_id,
            },
        });
        let sections = [paginator::Section {
            blocks: vec![LayoutBlock::Paragraph(paragraph)],
            geometry: PageGeometry::default(),
            header_footer: None,
            title_pg: false,
            page_number_start: None,
        }];

        let (pages, _) = paginator::paginate_sections(
            &sections,
            &FontManager::new(),
            &media,
            &NoteRegistry::default(),
        );
        let images = pages[0]
            .elements
            .iter()
            .filter_map(|element| match element {
                PositionedElement::Image {
                    data,
                    content_type,
                    media_id,
                    ..
                } => Some((data.as_slice(), content_type.as_str(), *media_id)),
                _ => None,
            })
            .collect::<Vec<_>>();

        assert!(images.contains(&(b"\x01\x02\x03".as_slice(), "image/png", inline_id)));
        assert!(images.contains(&(b"\x04\x05\x06".as_slice(), "image/jpeg", anchor_id)));
    }

    #[test]
    fn watermark_image_uses_the_collision_safe_media_registry_id() {
        let mut input = make_input_with_text("body");
        input.images.insert(
            "rIdHeader\0rIdOrdinary".to_owned(),
            ImageData {
                data: vec![1],
                content_type: "image/png".to_owned(),
            },
        );
        input.images.insert(
            "rIdHeader\0rIdWatermark".to_owned(),
            ImageData {
                data: vec![2],
                content_type: "image/png".to_owned(),
            },
        );
        let media = MediaRegistry::with_hasher(&input.images, |_| MediaId(7));
        let expected = media.id_for_relationship("rIdHeader\0rIdWatermark");
        let mut font_manager = FontManager::new();
        let mut diagnostics = Vec::new();
        let group = layout_watermark(
            &VmlWatermark::Image {
                relationship_id: "rIdWatermark".to_owned(),
                width_pt: 72.0,
                height_pt: 36.0,
                rotation_degrees: 0.0,
                opacity: 0.5,
            },
            "rIdHeader",
            &input,
            &media,
            &mut font_manager,
            PageGeometry::default(),
            &mut diagnostics,
        )
        .unwrap()
        .unwrap();
        let PositionedElement::Image { media_id, data, .. } = &group.children[0] else {
            panic!("expected watermark image");
        };
        assert_eq!(*media_id, expected);
        assert_eq!(data, &[2]);
        assert!(diagnostics.is_empty());
    }

    #[test]
    fn group_inline_item_breaks_and_positions_like_an_image() {
        let child = PositionedElement::FilledRect {
            rect: Rect {
                x: 2.0,
                y: 3.0,
                width: 4.0,
                height: 5.0,
            },
            color: Color::BLACK,
        };
        let group = GroupElement {
            transform: oxml_layout::Transform::IDENTITY,
            clip: None,
            opacity: 1.0,
            effects: Vec::new(),
            children: vec![child.clone()],
        };
        let line = oxml_layout::LayoutLine {
            items: vec![oxml_layout::LineItem::Group {
                width: 80.0,
                height: 40.0,
                group,
            }],
            width: 80.0,
            ascent: 40.0,
            descent: 0.0,
            line_gap: 0.0,
            height: 40.0,
            indent_left: 0.0,
            available_width: 468.0,
            is_last: true,
        };
        let paragraph = block::build_paragraph_block(
            vec![line],
            0.0,
            0.0,
            None,
            None,
            0.0,
            0.0,
            None,
            false,
            false,
            false,
            true,
        );
        let sections = [paginator::Section {
            blocks: vec![LayoutBlock::Paragraph(paragraph)],
            geometry: PageGeometry::default(),
            header_footer: None,
            title_pg: false,
            page_number_start: None,
        }];
        let media = MediaRegistry::new(&HashMap::new());
        let (pages, _) = paginator::paginate_sections(
            &sections,
            &FontManager::new(),
            &media,
            &NoteRegistry::default(),
        );

        let PositionedElement::Group(actual) = &pages[0].elements[0] else {
            panic!("group line item should become a positioned group");
        };
        assert_eq!((actual.transform.e, actual.transform.f), (72.0, 72.0));
        assert_eq!(actual.children, vec![child]);
    }

    #[test]
    fn layout_with_heading_style() {
        let mut doc = rdocx_oxml::document::CT_Document::new();
        let mut p = CT_P::new();
        p.properties = Some(CT_PPr {
            style_id: Some("Heading1".to_string()),
            ..Default::default()
        });
        p.add_run("Chapter 1");
        doc.body.add_paragraph(p);

        let input = LayoutInput {
            revision_view: crate::input::RevisionView::Accepted,
            document: doc,
            styles: CT_Styles::new_default(),
            numbering: None,
            headers: HashMap::new(),
            footers: HashMap::new(),
            images: HashMap::new(),
            charts: HashMap::new(),
            chart_theme: oxml_drawing::theme::CT_OfficeStyleSheet::office_default(),
            chart_color_map: oxml_drawing::color::ColorMap::default(),
            core_properties: None,
            hyperlink_urls: HashMap::new(),
            footnotes: None,
            endnotes: None,
            theme: None,
            fonts: Vec::new(),
        };

        let result = Engine::new().layout(&input);
        if let Ok(result) = result {
            assert!(!result.pages.is_empty());
            // Should produce one outline entry for Heading1
            assert_eq!(result.outlines.len(), 1);
            assert_eq!(result.outlines[0].title, "Chapter 1");
            assert_eq!(result.outlines[0].level, 1);
            assert_eq!(result.outlines[0].page_index, 0);
        }
    }

    #[test]
    fn layout_nested_headings_produce_outlines() {
        let mut doc = rdocx_oxml::document::CT_Document::new();

        // H1
        let mut h1 = CT_P::new();
        h1.properties = Some(CT_PPr {
            style_id: Some("Heading1".to_string()),
            ..Default::default()
        });
        h1.add_run("Chapter 1");
        doc.body.add_paragraph(h1);

        // H2 under H1
        let mut h2 = CT_P::new();
        h2.properties = Some(CT_PPr {
            style_id: Some("Heading2".to_string()),
            ..Default::default()
        });
        h2.add_run("Section 1.1");
        doc.body.add_paragraph(h2);

        // Another H1
        let mut h1b = CT_P::new();
        h1b.properties = Some(CT_PPr {
            style_id: Some("Heading1".to_string()),
            ..Default::default()
        });
        h1b.add_run("Chapter 2");
        doc.body.add_paragraph(h1b);

        let input = LayoutInput {
            revision_view: crate::input::RevisionView::Accepted,
            document: doc,
            styles: CT_Styles::new_default(),
            numbering: None,
            headers: HashMap::new(),
            footers: HashMap::new(),
            images: HashMap::new(),
            charts: HashMap::new(),
            chart_theme: oxml_drawing::theme::CT_OfficeStyleSheet::office_default(),
            chart_color_map: oxml_drawing::color::ColorMap::default(),
            core_properties: None,
            hyperlink_urls: HashMap::new(),
            footnotes: None,
            endnotes: None,
            theme: None,
            fonts: Vec::new(),
        };

        let result = Engine::new().layout(&input);
        if let Ok(result) = result {
            assert_eq!(result.outlines.len(), 3);
            assert_eq!(result.outlines[0].level, 1);
            assert_eq!(result.outlines[0].title, "Chapter 1");
            assert_eq!(result.outlines[1].level, 2);
            assert_eq!(result.outlines[1].title, "Section 1.1");
            assert_eq!(result.outlines[2].level, 1);
            assert_eq!(result.outlines[2].title, "Chapter 2");
        }
    }

    #[test]
    fn sect_pr_geometry_conversion() {
        let sect = CT_SectPr::default_letter();
        let geom = sect_pr_to_geometry(&sect);
        assert!((geom.page_width - 612.0).abs() < 0.01);
        assert!((geom.page_height - 792.0).abs() < 0.01);
        assert!((geom.margin_top - 72.0).abs() < 0.01);
        assert!((geom.content_width() - 468.0).abs() < 0.01);
    }

    #[test]
    fn section_page_number_start_requires_a_direct_word_child_and_decodes_entities() {
        let mut section = CT_SectPr::default_letter();
        section.extra_xml = vec![
            br#"<x:pgNumType xmlns:x="urn:producer" x:start="2"/>"#.to_vec(),
            br#"<w:pgNumType xmlns:w="urn:producer" w:start="2"/>"#.to_vec(),
            format!(
                r#"<w:wrapper xmlns:w="{}"><w:pgNumType w:start="2"/></w:wrapper>"#,
                rdocx_oxml::namespace::W_NS
            )
            .into_bytes(),
            format!(
                r#"<q:pgNumType xmlns:q="{}" q:start="&#x31;"/>"#,
                rdocx_oxml::namespace::W_NS
            )
            .into_bytes(),
        ];

        assert_eq!(section_page_number_start(&section), Some(1));
    }

    #[test]
    fn sect_pr_a4_geometry() {
        let sect = CT_SectPr::default_a4();
        let geom = sect_pr_to_geometry(&sect);
        // A4: 210mm = 595.3pt, 297mm = 841.9pt
        assert!((geom.page_width - 595.3).abs() < 0.5);
        assert!((geom.page_height - 841.9).abs() < 0.5);
    }

    // F-X013a, footnote line advance.

    /// Build a document whose single body paragraph references footnote 1, and
    /// whose footnote 1 is one paragraph made of `note_runs` separate runs.
    fn make_input_with_footnote(note_runs: &[&str]) -> LayoutInput {
        use rdocx_oxml::footnotes::{CT_Footnote, CT_Footnotes, NoteType};
        use rdocx_oxml::text::CT_R;

        let mut doc = rdocx_oxml::document::CT_Document::new();
        let mut body = CT_P::new();
        body.add_run("Body text carrying a note");
        let mut marker_run = CT_R::new("");
        marker_run.content = vec![RunContent::FootnoteRef { id: 1 }];
        body.runs.push(marker_run);
        doc.body.add_paragraph(body);

        let mut note = CT_P::new();
        for text in note_runs {
            note.add_run(text);
        }

        LayoutInput {
            revision_view: crate::input::RevisionView::Accepted,
            document: doc,
            styles: CT_Styles::new_default(),
            numbering: None,
            headers: HashMap::new(),
            footers: HashMap::new(),
            images: HashMap::new(),
            charts: HashMap::new(),
            chart_theme: oxml_drawing::theme::CT_OfficeStyleSheet::office_default(),
            chart_color_map: oxml_drawing::color::ColorMap::default(),
            core_properties: None,
            hyperlink_urls: HashMap::new(),
            footnotes: Some(CT_Footnotes {
                footnotes: vec![CT_Footnote {
                    id: 1,
                    note_type: NoteType::Normal,
                    paragraphs: vec![note],
                }],
            }),
            endnotes: None,
            theme: None,
            fonts: Vec::new(),
        }
    }

    /// The x origin of every glyph run sitting below the footnote separator,
    /// in the order the renderer emitted them. The first is the note marker.
    fn footnote_glyph_x(page: &oxml_layout::output::PageFrame) -> Vec<f64> {
        let separator_y = page
            .elements
            .iter()
            .find_map(|element| match element {
                PositionedElement::Line { start, .. } => Some(start.y),
                _ => None,
            })
            .expect("a page with a footnote draws a separator line");

        page.elements
            .iter()
            .filter_map(|element| match element {
                PositionedElement::Text(run) if run.origin.y > separator_y => Some(run.origin.x),
                _ => None,
            })
            .collect()
    }

    #[test]
    fn a_multi_segment_footnote_does_not_stack_its_segments_at_one_x() {
        let input = make_input_with_footnote(&["Alpha", "Beta", "Gamma"]);
        let mut engine = Engine::new();
        let output = engine.layout(&input).expect("layout succeeds");
        let xs = footnote_glyph_x(&output.pages[0]);

        assert!(
            xs.len() >= 4,
            "expected a marker and three note segments, got {xs:?}"
        );
        for pair in xs.windows(2) {
            assert!(
                pair[1] > pair[0],
                "footnote segments must advance, got {xs:?}"
            );
        }
    }

    #[test]
    fn a_single_segment_footnote_keeps_its_original_position() {
        let input = make_input_with_footnote(&["Solitary"]);
        let mut engine = Engine::new();
        let output = engine.layout(&input).expect("layout succeeds");
        let xs = footnote_glyph_x(&output.pages[0]);
        let geometry = sect_pr_to_geometry(&CT_SectPr::default_letter());

        // The marker sits at the left margin, the single segment one indent in.
        assert_eq!(xs.len(), 2, "expected a marker and one segment, got {xs:?}");
        assert!(
            (xs[0] - geometry.margin_left).abs() < 0.01,
            "marker at {xs:?}"
        );
        assert!(
            (xs[1] - (geometry.margin_left + 12.0)).abs() < 0.01,
            "segment at {xs:?}"
        );
    }

    #[test]
    fn a_long_footnote_does_not_overrun_the_right_margin() {
        // Long enough to wrap, which is what exposes a break width that
        // disagrees with the indent the note is drawn at.
        let long = "In paged media, footnotes are usually displayed at the \
                    bottom of the text. However, in ebooks, a better paradigm \
                    is to make them clickable endnotes that the reader can \
                    browse at leisure, which this sentence exists to force.";
        let input = make_input_with_footnote(&[long]);
        let mut engine = Engine::new();
        let output = engine.layout(&input).expect("layout succeeds");
        let page = &output.pages[0];
        let geometry = sect_pr_to_geometry(&CT_SectPr::default_letter());
        let right_margin = geometry.page_width - geometry.margin_right;

        let separator_y = page
            .elements
            .iter()
            .find_map(|element| match element {
                PositionedElement::Line { start, .. } => Some(start.y),
                _ => None,
            })
            .expect("a page with a footnote draws a separator line");

        let mut wrapped = false;
        let mut first_y = None;
        for element in &page.elements {
            let PositionedElement::Text(run) = element else {
                continue;
            };
            if run.origin.y <= separator_y {
                continue;
            }
            let first = *first_y.get_or_insert(run.origin.y);
            if run.origin.y > first + 0.01 {
                wrapped = true;
            }
            let right_edge = run.origin.x + run.advances.iter().sum::<f64>();
            assert!(
                right_edge <= right_margin + 0.01,
                "note text reaches {right_edge}, past the right margin {right_margin}"
            );
        }
        assert!(wrapped, "the note must wrap for this test to mean anything");
    }

    #[test]
    fn a_tab_inside_a_footnote_still_advances_the_text_after_it() {
        use rdocx_oxml::footnotes::{CT_Footnote, CT_Footnotes, NoteType};
        use rdocx_oxml::text::CT_R;

        // Two notes differing only by a tab between their runs. The tab is not
        // drawn, but it occupies width, so the run after it must shift right.
        let build = |with_tab: bool| {
            let mut doc = rdocx_oxml::document::CT_Document::new();
            let mut body = CT_P::new();
            body.add_run("Body");
            let mut marker_run = CT_R::new("");
            marker_run.content = vec![RunContent::FootnoteRef { id: 1 }];
            body.runs.push(marker_run);
            doc.body.add_paragraph(body);

            let mut note = CT_P::new();
            note.add_run("Alpha");
            if with_tab {
                let mut tab_run = CT_R::new("");
                tab_run.content = vec![RunContent::Tab];
                note.runs.push(tab_run);
            }
            note.add_run("Beta");

            LayoutInput {
                revision_view: crate::input::RevisionView::Accepted,
                document: doc,
                styles: CT_Styles::new_default(),
                numbering: None,
                headers: HashMap::new(),
                footers: HashMap::new(),
                images: HashMap::new(),
                charts: HashMap::new(),
                chart_theme: oxml_drawing::theme::CT_OfficeStyleSheet::office_default(),
                chart_color_map: oxml_drawing::color::ColorMap::default(),
                core_properties: None,
                hyperlink_urls: HashMap::new(),
                footnotes: Some(CT_Footnotes {
                    footnotes: vec![CT_Footnote {
                        id: 1,
                        note_type: NoteType::Normal,
                        paragraphs: vec![note],
                    }],
                }),
                endnotes: None,
                theme: None,
                fonts: Vec::new(),
            }
        };

        let mut engine = Engine::new();
        let plain = engine.layout(&build(false)).expect("layout succeeds");
        let tabbed = engine.layout(&build(true)).expect("layout succeeds");

        let plain_x = footnote_glyph_x(&plain.pages[0]);
        let tabbed_x = footnote_glyph_x(&tabbed.pages[0]);

        // Marker and both runs are drawn in each case. The tab draws nothing.
        assert_eq!(plain_x.len(), 3, "plain note glyphs {plain_x:?}");
        assert_eq!(tabbed_x.len(), 3, "tabbed note glyphs {tabbed_x:?}");
        assert!(
            tabbed_x[2] > plain_x[2] + 1.0,
            "the run after a tab must shift right, plain {plain_x:?} tabbed {tabbed_x:?}"
        );
    }

    #[test]
    fn footnote_segment_advance_matches_body_segment_advance() {
        let input = make_input_with_footnote(&["Alpha", "Beta", "Gamma"]);
        let mut engine = Engine::new();
        let output = engine.layout(&input).expect("layout succeeds");
        let page = &output.pages[0];

        let separator_y = page
            .elements
            .iter()
            .find_map(|element| match element {
                PositionedElement::Line { start, .. } => Some(start.y),
                _ => None,
            })
            .expect("a page with a footnote draws a separator line");

        // Gaps between consecutive note segments must equal the width of the
        // segment that precedes them, which is what the body path advances by.
        let notes: Vec<&oxml_layout::GlyphRun> = page
            .elements
            .iter()
            .filter_map(|element| match element {
                PositionedElement::Text(run) if run.origin.y > separator_y => Some(run),
                _ => None,
            })
            .skip(1) // the marker, which is positioned independently
            .collect();

        assert_eq!(notes.len(), 3, "expected three note segments");
        for pair in notes.windows(2) {
            let advance: f64 = pair[0].advances.iter().sum();
            let gap = pair[1].origin.x - pair[0].origin.x;
            assert!(
                (gap - advance).abs() < 0.01,
                "gap {gap} should equal preceding segment advance {advance}"
            );
        }
    }

    // F-X013b, reservation and splitting.

    /// A document of `body_paras` paragraphs. The paragraph at
    /// `ref_positions` each carry a reference to note 1, whose content is
    /// `note_paras` paragraphs of `note_text`.
    fn make_noted_document(
        body_paras: usize,
        ref_positions: &[usize],
        note_paras: usize,
        note_text: &str,
        continuation_separator: bool,
    ) -> LayoutInput {
        use rdocx_oxml::footnotes::{CT_Footnote, CT_Footnotes, NoteType};
        use rdocx_oxml::text::CT_R;

        let mut doc = rdocx_oxml::document::CT_Document::new();
        for index in 0..body_paras {
            let mut para = CT_P::new();
            para.add_run("Body paragraph text that occupies a line of the page.");
            if ref_positions.contains(&index) {
                let mut marker = CT_R::new("");
                marker.content = vec![RunContent::FootnoteRef { id: 1 }];
                para.runs.push(marker);
            }
            doc.body.add_paragraph(para);
        }

        let mut entries = Vec::new();
        if continuation_separator {
            entries.push(CT_Footnote {
                id: 0,
                note_type: NoteType::ContinuationSeparator,
                paragraphs: vec![CT_P::new()],
            });
        }
        entries.push(CT_Footnote {
            id: 1,
            note_type: NoteType::Normal,
            paragraphs: (0..note_paras)
                .map(|_| {
                    let mut p = CT_P::new();
                    p.add_run(note_text);
                    p
                })
                .collect(),
        });

        LayoutInput {
            revision_view: crate::input::RevisionView::Accepted,
            document: doc,
            styles: CT_Styles::new_default(),
            numbering: None,
            headers: HashMap::new(),
            footers: HashMap::new(),
            images: HashMap::new(),
            charts: HashMap::new(),
            chart_theme: oxml_drawing::theme::CT_OfficeStyleSheet::office_default(),
            chart_color_map: oxml_drawing::color::ColorMap::default(),
            core_properties: None,
            hyperlink_urls: HashMap::new(),
            footnotes: Some(CT_Footnotes { footnotes: entries }),
            endnotes: None,
            theme: None,
            fonts: Vec::new(),
        }
    }

    /// Split a page into the glyphs drawn above the note separator and those
    /// drawn below it. Notes are emitted after body content, so the separator
    /// is the boundary.
    fn split_at_separator(
        page: &oxml_layout::output::PageFrame,
    ) -> Option<(f64, Vec<f64>, Vec<String>)> {
        let separator_index = page.elements.iter().position(|element| {
            matches!(element, PositionedElement::Line { width, .. } if (*width - 0.5).abs() < 0.001)
        })?;
        let PositionedElement::Line { start, end, .. } = &page.elements[separator_index] else {
            return None;
        };
        let separator_y = start.y;
        let separator_width = end.x - start.x;

        let body_ys: Vec<f64> = page.elements[..separator_index]
            .iter()
            .filter_map(|element| match element {
                PositionedElement::Text(run) => Some(run.origin.y),
                _ => None,
            })
            .collect();
        let note_text: Vec<String> = page.elements[separator_index + 1..]
            .iter()
            .filter_map(|element| match element {
                PositionedElement::Text(run) => Some(run.text.clone()),
                _ => None,
            })
            .collect();

        let _ = separator_y;
        Some((separator_width, body_ys, note_text))
    }

    fn separator_y_of(page: &oxml_layout::output::PageFrame) -> Option<f64> {
        page.elements.iter().find_map(|element| match element {
            PositionedElement::Line { start, width, .. } if (*width - 0.5).abs() < 0.001 => {
                Some(start.y)
            }
            _ => None,
        })
    }

    #[test]
    fn a_page_whose_body_fills_the_text_area_does_not_overlap_its_notes() {
        // Enough body to reach the bottom margin, with the reference early so
        // the note is owed by the first page.
        let input = make_noted_document(
            60,
            &[0],
            2,
            "A note long enough to wrap onto a second line of the note area.",
            false,
        );
        let mut engine = Engine::new();
        let output = engine.layout(&input).expect("layout succeeds");
        let page = &output.pages[0];

        let separator_y = separator_y_of(page).expect("the page draws a separator");
        let (_, body_ys, note_text) = split_at_separator(page).unwrap();

        assert!(!note_text.is_empty(), "the note must be drawn");
        let lowest_body = body_ys.iter().cloned().fold(f64::MIN, f64::max);
        assert!(
            lowest_body < separator_y,
            "body text reaches {lowest_body}, at or below the separator at {separator_y}"
        );
    }

    #[test]
    fn a_page_referencing_one_note_twice_reserves_it_once() {
        let input = make_noted_document(4, &[0, 1], 1, "Referenced twice from one page.", false);
        let mut engine = Engine::new();
        let output = engine.layout(&input).expect("layout succeeds");
        let page = &output.pages[0];

        let separators = page
            .elements
            .iter()
            .filter(|e| matches!(e, PositionedElement::Line { width, .. } if (*width - 0.5).abs() < 0.001))
            .count();
        assert_eq!(separators, 1, "one note area, so one separator");

        let (_, _, note_text) = split_at_separator(page).unwrap();
        let markers = note_text.iter().filter(|t| t.as_str() == "1").count();
        assert_eq!(markers, 1, "the note is drawn once, got {note_text:?}");
    }

    #[test]
    fn a_note_taller_than_its_remaining_space_continues_on_the_next_page() {
        // 120 note paragraphs exceed a single page, so the note has to break.
        let input = make_noted_document(30, &[25], 120, "Note paragraph line.", true);
        let mut engine = Engine::new();
        let output = engine.layout(&input).expect("layout succeeds");

        let note_pages: Vec<usize> = output
            .pages
            .iter()
            .enumerate()
            .filter(|(_, page)| separator_y_of(page).is_some())
            .map(|(index, _)| index)
            .collect();

        assert!(
            note_pages.len() >= 2,
            "a note taller than a page must span pages, got {note_pages:?}"
        );

        let first = split_at_separator(&output.pages[note_pages[0]]).unwrap().2;
        let second = split_at_separator(&output.pages[note_pages[1]]).unwrap().2;

        assert!(!first.is_empty(), "the first page draws part of the note");
        assert!(!second.is_empty(), "the next page draws the rest");
        assert_eq!(
            first.iter().filter(|t| t.as_str() == "1").count(),
            1,
            "the marker is drawn on the page the note starts on"
        );
        assert_eq!(
            second.iter().filter(|t| t.as_str() == "1").count(),
            0,
            "a continuation does not repeat the marker, got {second:?}"
        );
    }

    #[test]
    fn a_continued_note_draws_the_continuation_separator() {
        let geometry = sect_pr_to_geometry(&CT_SectPr::default_letter());

        let widths = |continuation: bool| {
            let input = make_noted_document(30, &[25], 120, "Note paragraph line.", continuation);
            let mut engine = Engine::new();
            let output = engine.layout(&input).expect("layout succeeds");
            let pages: Vec<usize> = output
                .pages
                .iter()
                .enumerate()
                .filter(|(_, page)| separator_y_of(page).is_some())
                .map(|(index, _)| index)
                .collect();
            assert!(pages.len() >= 2, "the note must span pages");
            (
                split_at_separator(&output.pages[pages[0]]).unwrap().0,
                split_at_separator(&output.pages[pages[1]]).unwrap().0,
            )
        };

        let (first, second) = widths(true);
        assert!(
            (first - geometry.content_width() * 0.33).abs() < 0.5,
            "a note starting on its page gets the short rule, got {first}"
        );
        assert!(
            (second - geometry.content_width()).abs() < 0.5,
            "a continued note gets the full-width rule, got {second}"
        );

        // A document defining no continuation separator keeps the short rule.
        let (_, second) = widths(false);
        assert!(
            (second - geometry.content_width() * 0.33).abs() < 0.5,
            "without a continuation separator the short rule is kept, got {second}"
        );
    }

    #[test]
    fn an_oversized_note_still_leaves_room_for_body_text() {
        // A note several pages tall, referenced from the first paragraph.
        let input = make_noted_document(3, &[0], 200, "A line of an enormous note.", true);
        let mut engine = Engine::new();
        let output = engine.layout(&input).expect("layout terminates");

        let (_, body_ys, _) = split_at_separator(&output.pages[0]).unwrap();
        assert!(
            !body_ys.is_empty(),
            "an oversized note must not starve the page of body text"
        );
        assert!(
            output.pages.len() > 1 && output.pages.len() < 100,
            "the note spills over a bounded number of pages, got {}",
            output.pages.len()
        );

        // The note area has to stay on the page. Placing an oversized note
        // whole would push its separator off the top of the sheet.
        for (index, page) in output.pages.iter().enumerate() {
            let Some(separator_y) = separator_y_of(page) else {
                continue;
            };
            assert!(
                separator_y >= 0.0,
                "page {} draws its separator at {separator_y}, off the sheet",
                index + 1
            );
        }
    }

    #[test]
    fn a_note_is_drawn_on_the_page_that_carries_its_reference() {
        // Sweeping the reference across the document is what catches the two
        // ways a note drifts off its own page: notes claimed for a paragraph
        // that then moves, and a note area measured from a cursor that still
        // holds the previous paragraph's trailing space.
        let mut mismatches = Vec::new();
        for position in 0..60 {
            let input = make_noted_document(60, &[position], 1, "Note text.", false);
            let mut engine = Engine::new();
            let output = engine.layout(&input).expect("layout succeeds");

            let reference_page = output.pages.iter().position(|page| {
                page.elements.iter().any(|element| {
                    matches!(element, PositionedElement::Text(run)
                    if run.note == Some(oxml_layout::NoteRef {
                        stream: oxml_layout::NoteStream::Footnote,
                        id: 1,
                    }))
                })
            });
            let note_page = output
                .pages
                .iter()
                .position(|page| separator_y_of(page).is_some());

            if reference_page != note_page {
                mismatches.push((position, reference_page, note_page));
            }
        }

        assert!(
            mismatches.is_empty(),
            "note and reference landed on different pages for (position, ref, note): {mismatches:?}"
        );
    }

    // F-X013c, endnotes at the document end.

    /// A document whose single body paragraph references footnote `id` and
    /// endnote `id`, with each stream giving that number different text.
    fn make_document_with_both_streams(id: i32, body_paras: usize) -> LayoutInput {
        use rdocx_oxml::footnotes::{CT_Footnote, CT_Footnotes, NoteType};
        use rdocx_oxml::text::CT_R;

        let mut doc = rdocx_oxml::document::CT_Document::new();
        for index in 0..body_paras {
            let mut para = CT_P::new();
            para.add_run("Body paragraph text that occupies a line of the page.");
            if index == 0 {
                let mut foot = CT_R::new("");
                foot.content = vec![RunContent::FootnoteRef { id }];
                para.runs.push(foot);
                let mut end = CT_R::new("");
                end.content = vec![RunContent::EndnoteRef { id }];
                para.runs.push(end);
            }
            doc.body.add_paragraph(para);
        }

        let note = |text: &str| {
            let mut p = CT_P::new();
            p.add_run(text);
            CT_Footnote {
                id,
                note_type: NoteType::Normal,
                paragraphs: vec![p],
            }
        };

        LayoutInput {
            revision_view: crate::input::RevisionView::Accepted,
            document: doc,
            styles: CT_Styles::new_default(),
            numbering: None,
            headers: HashMap::new(),
            footers: HashMap::new(),
            images: HashMap::new(),
            charts: HashMap::new(),
            chart_theme: oxml_drawing::theme::CT_OfficeStyleSheet::office_default(),
            chart_color_map: oxml_drawing::color::ColorMap::default(),
            core_properties: None,
            hyperlink_urls: HashMap::new(),
            footnotes: Some(CT_Footnotes {
                footnotes: vec![note("FOOTNOTETEXT")],
            }),
            endnotes: Some(CT_Footnotes {
                footnotes: vec![note("ENDNOTETEXT")],
            }),
            theme: None,
            fonts: Vec::new(),
        }
    }

    fn page_text(page: &oxml_layout::output::PageFrame) -> String {
        page.elements
            .iter()
            .filter_map(|element| match element {
                PositionedElement::Text(run) => Some(run.text.as_str()),
                _ => None,
            })
            .collect::<Vec<_>>()
            .join(" ")
    }

    #[test]
    fn a_footnote_and_an_endnote_sharing_a_number_render_their_own_text() {
        let input = make_document_with_both_streams(2, 3);
        let mut engine = Engine::new();
        let output = engine.layout(&input).expect("layout succeeds");

        let all: String = output
            .pages
            .iter()
            .map(|page| page_text(page))
            .collect::<Vec<_>>()
            .join(" | ");
        assert!(
            all.contains("FOOTNOTETEXT"),
            "the footnote must render its own text, got {all}"
        );
        assert!(
            all.contains("ENDNOTETEXT"),
            "the endnote must render its own text, got {all}"
        );
    }

    #[test]
    fn endnotes_render_after_the_last_body_page() {
        let input = make_document_with_both_streams(2, 3);
        let mut engine = Engine::new();
        let output = engine.layout(&input).expect("layout succeeds");

        let endnote_page = output
            .pages
            .iter()
            .position(|page| page_text(page).contains("ENDNOTETEXT"))
            .expect("the endnote is rendered somewhere");
        let last_body_page = output
            .pages
            .iter()
            .rposition(|page| page_text(page).contains("occupies"))
            .expect("the body is rendered somewhere");

        assert!(
            endnote_page > last_body_page,
            "endnotes come after every body page, endnote on {endnote_page} and body to {last_body_page}"
        );
        assert!(
            !page_text(&output.pages[endnote_page]).contains("occupies"),
            "an endnote page carries no body text"
        );
    }

    #[test]
    fn footnotes_and_endnotes_keep_their_own_regions() {
        let input = make_document_with_both_streams(2, 3);
        let mut engine = Engine::new();
        let output = engine.layout(&input).expect("layout succeeds");

        let footnote_page = output
            .pages
            .iter()
            .position(|page| page_text(page).contains("FOOTNOTETEXT"))
            .expect("the footnote is rendered");

        // The footnote shares the page that carries its reference.
        assert!(
            page_text(&output.pages[footnote_page]).contains("occupies"),
            "a footnote sits on the page carrying its reference"
        );
        assert!(
            separator_y_of(&output.pages[footnote_page]).is_some(),
            "the footnote page draws a separator"
        );

        // The endnote page is a different page, and draws no separator,
        // because there is no body text there to divide it from.
        let endnote_page = output
            .pages
            .iter()
            .position(|page| page_text(page).contains("ENDNOTETEXT"))
            .expect("the endnote is rendered");
        assert_ne!(footnote_page, endnote_page, "the two regions are distinct");
        assert!(
            separator_y_of(&output.pages[endnote_page]).is_none(),
            "an endnote page draws no separator rule"
        );
    }

    #[test]
    fn an_endnote_reference_does_not_reserve_space_at_the_page_foot() {
        use rdocx_oxml::footnotes::{CT_Footnote, CT_Footnotes, NoteType};
        use rdocx_oxml::text::CT_R;

        // The same document twice, once with an endnote reference and once
        // with none. An endnote costs its page nothing, so the body must
        // paginate identically.
        let build = |with_endnote: bool| {
            let mut doc = rdocx_oxml::document::CT_Document::new();
            for index in 0..60 {
                let mut para = CT_P::new();
                para.add_run("Body paragraph text that occupies a line of the page.");
                if index == 0 && with_endnote {
                    let mut end = CT_R::new("");
                    end.content = vec![RunContent::EndnoteRef { id: 1 }];
                    para.runs.push(end);
                }
                doc.body.add_paragraph(para);
            }
            let mut note = CT_P::new();
            note.add_run("An endnote that would be tall in the margin.");
            LayoutInput {
                revision_view: crate::input::RevisionView::Accepted,
                document: doc,
                styles: CT_Styles::new_default(),
                numbering: None,
                headers: HashMap::new(),
                footers: HashMap::new(),
                images: HashMap::new(),
                charts: HashMap::new(),
                chart_theme: oxml_drawing::theme::CT_OfficeStyleSheet::office_default(),
                chart_color_map: oxml_drawing::color::ColorMap::default(),
                core_properties: None,
                hyperlink_urls: HashMap::new(),
                footnotes: None,
                endnotes: Some(CT_Footnotes {
                    footnotes: vec![CT_Footnote {
                        id: 1,
                        note_type: NoteType::Normal,
                        paragraphs: vec![note],
                    }],
                }),
                theme: None,
                fonts: Vec::new(),
            }
        };

        let mut engine = Engine::new();
        let plain = engine.layout(&build(false)).expect("layout succeeds");
        let noted = engine.layout(&build(true)).expect("layout succeeds");

        // One extra page for the endnote itself, and no separator anywhere.
        assert_eq!(
            noted.pages.len(),
            plain.pages.len() + 1,
            "an endnote adds its own page and takes none from the body"
        );
        for (index, page) in noted.pages.iter().enumerate() {
            if index < plain.pages.len() {
                assert!(
                    separator_y_of(page).is_none(),
                    "page {} reserved foot space for an endnote",
                    index + 1
                );
            }
        }

        // Body pagination is untouched.
        for (index, plain_page) in plain.pages.iter().enumerate() {
            let body_lines = |page: &oxml_layout::output::PageFrame| {
                page.elements
                    .iter()
                    .filter(|element| {
                        matches!(element, PositionedElement::Text(run)
                            if run.text.starts_with("occupies"))
                    })
                    .count()
            };
            assert_eq!(
                body_lines(plain_page),
                body_lines(&noted.pages[index]),
                "page {} holds a different amount of body text",
                index + 1
            );
        }
    }

    // F-X016, text wrapping around a floating drawing.

    /// A document of one long paragraph, with a floating drawing anchored to
    /// it. `align` places the drawing, `wrap` says how text should treat it.
    fn make_wrapping_document(
        wrap: rdocx_oxml::drawing::WrapType,
        align: Option<rdocx_oxml::drawing::AnchorAlignH>,
        width_pt: f64,
        height_pt: f64,
        dist_pt: f64,
    ) -> LayoutInput {
        use rdocx_oxml::drawing::{CT_Anchor, CT_Drawing, ST_RelativeFromH, ST_RelativeFromV};
        use rdocx_oxml::text::CT_R;
        use rdocx_oxml::units::Emu;

        let emu = |pt: f64| Emu((pt * 12700.0) as i64);

        let mut doc = rdocx_oxml::document::CT_Document::new();
        let mut para = CT_P::new();
        // Long enough that many lines sit below the drawing, which is what
        // makes "returns to the margin" a meaningful assertion.
        let mut body = String::new();
        for index in 0..40 {
            body.push_str(&format!(
                "Sentence {index} of running text that fills the paragraph out. "
            ));
        }
        para.add_run(&body);

        let mut anchor = CT_Anchor::background("rId1", 0, 0);
        anchor.extent_cx = emu(width_pt);
        anchor.extent_cy = emu(height_pt);
        anchor.behind_doc = false;
        anchor.wrap = wrap;
        anchor.pos_h_relative_from = ST_RelativeFromH::Margin;
        anchor.pos_h_align = align;
        anchor.pos_v_relative_from = ST_RelativeFromV::Paragraph;
        anchor.pos_v_offset = Emu(0);
        anchor.dist_t = emu(dist_pt);
        anchor.dist_b = emu(dist_pt);
        anchor.dist_l = emu(dist_pt);
        anchor.dist_r = emu(dist_pt);

        let mut drawing_run = CT_R::new("");
        drawing_run.content = vec![RunContent::Drawing(CT_Drawing {
            inline: None,
            anchor: Some(anchor),
        })];
        para.runs.push(drawing_run);
        doc.body.add_paragraph(para);

        let mut images = HashMap::new();
        images.insert(
            "rId1".to_string(),
            ImageData {
                data: vec![0u8; 8],
                content_type: "image/png".to_string(),
            },
        );

        LayoutInput {
            revision_view: crate::input::RevisionView::Accepted,
            document: doc,
            styles: CT_Styles::new_default(),
            numbering: None,
            headers: HashMap::new(),
            footers: HashMap::new(),
            images,
            charts: HashMap::new(),
            chart_theme: oxml_drawing::theme::CT_OfficeStyleSheet::office_default(),
            chart_color_map: oxml_drawing::color::ColorMap::default(),
            core_properties: None,
            hyperlink_urls: HashMap::new(),
            footnotes: None,
            endnotes: None,
            theme: None,
            fonts: Vec::new(),
        }
    }

    /// The x origin and right edge of every body text run, by line.
    fn text_extents(page: &oxml_layout::output::PageFrame) -> Vec<(f64, f64)> {
        let mut by_line: Vec<(f64, f64, f64)> = Vec::new();
        for element in &page.elements {
            let PositionedElement::Text(run) = element else {
                continue;
            };
            let right = run.origin.x + run.advances.iter().sum::<f64>();
            if let Some(entry) = by_line
                .iter_mut()
                .find(|(y, _, _)| (*y - run.origin.y).abs() < 0.01)
            {
                entry.1 = entry.1.min(run.origin.x);
                entry.2 = entry.2.max(right);
            } else {
                by_line.push((run.origin.y, run.origin.x, right));
            }
        }
        by_line.sort_by(|a, b| a.0.partial_cmp(&b.0).unwrap());
        by_line.into_iter().map(|(_, l, r)| (l, r)).collect()
    }

    #[test]
    fn text_wraps_beside_a_left_aligned_square_drawing() {
        use rdocx_oxml::drawing::{AnchorAlignH, WrapType};

        let input =
            make_wrapping_document(WrapType::Square, Some(AnchorAlignH::Left), 100.0, 40.0, 5.0);
        let mut engine = Engine::new();
        let output = engine.layout(&input).expect("layout succeeds");
        let geometry = sect_pr_to_geometry(&CT_SectPr::default_letter());
        let extents = text_extents(&output.pages[0]);

        assert!(
            extents.len() > 2,
            "the paragraph must wrap, got {extents:?}"
        );

        // Lines beside the drawing start to its right, past width plus distR.
        let expected_left = geometry.margin_left + 100.0 + 5.0;
        assert!(
            (extents[0].0 - expected_left).abs() < 1.0,
            "first line should start at {expected_left}, got {:?}",
            extents[0]
        );

        // A line below the drawing returns to the margin.
        let last = extents.last().unwrap();
        assert!(
            (last.0 - geometry.margin_left).abs() < 1.0,
            "the last line should return to the margin, got {last:?}"
        );
    }

    #[test]
    fn text_wraps_beside_a_right_aligned_square_drawing() {
        use rdocx_oxml::drawing::{AnchorAlignH, WrapType};

        let input = make_wrapping_document(
            WrapType::Square,
            Some(AnchorAlignH::Right),
            100.0,
            40.0,
            5.0,
        );
        let mut engine = Engine::new();
        let output = engine.layout(&input).expect("layout succeeds");
        let geometry = sect_pr_to_geometry(&CT_SectPr::default_letter());
        let extents = text_extents(&output.pages[0]);

        assert!(
            extents.len() > 2,
            "the paragraph must wrap, got {extents:?}"
        );

        // Lines beside the drawing still start at the margin but end early.
        let text_right = geometry.page_width - geometry.margin_right;
        let drawing_left = text_right - 100.0;
        assert!(
            (extents[0].0 - geometry.margin_left).abs() < 1.0,
            "a right-aligned drawing does not move the line start, got {:?}",
            extents[0]
        );
        assert!(
            extents[0].1 <= drawing_left - 5.0 + 1.0,
            "the first line should stop before the drawing at {}, got {:?}",
            drawing_left - 5.0,
            extents[0]
        );

        // Some line below the drawing runs past where the drawing sat, which
        // is only possible once the reservation stops applying. The final line
        // of a paragraph is naturally short, so the widest is the fair test.
        let widest = extents
            .iter()
            .map(|(_, right)| *right)
            .fold(f64::MIN, f64::max);
        assert!(
            widest > drawing_left,
            "a line below the drawing should reach past {drawing_left}, got {extents:?}"
        );
    }

    #[test]
    fn a_top_and_bottom_drawing_pushes_text_below_it() {
        use rdocx_oxml::drawing::WrapType;

        let input = make_wrapping_document(WrapType::TopAndBottom, None, 100.0, 40.0, 5.0);
        let mut engine = Engine::new();
        let output = engine.layout(&input).expect("layout succeeds");
        let geometry = sect_pr_to_geometry(&CT_SectPr::default_letter());
        let extents = text_extents(&output.pages[0]);

        assert!(!extents.is_empty(), "the paragraph renders");

        // The drawing sits at the paragraph top, so text starts below its
        // bottom edge plus distB.
        let first_baseline = output.pages[0]
            .elements
            .iter()
            .find_map(|element| match element {
                PositionedElement::Text(run) => Some(run.origin.y),
                _ => None,
            })
            .expect("text is rendered");
        let drawing_bottom = geometry.margin_top + 40.0 + 5.0;
        assert!(
            first_baseline >= drawing_bottom,
            "the first line at {first_baseline} should sit below {drawing_bottom}"
        );
    }

    #[test]
    fn a_wrap_none_drawing_leaves_text_untouched() {
        use rdocx_oxml::drawing::WrapType;

        // The identity case. A drawing that does not wrap must not move a
        // single glyph, which is what keeps every recorded baseline still.
        let with = make_wrapping_document(WrapType::None, None, 100.0, 40.0, 5.0);
        let mut engine = Engine::new();
        let output = engine.layout(&with).expect("layout succeeds");
        let wrapped_extents = text_extents(&output.pages[0]);

        let geometry = sect_pr_to_geometry(&CT_SectPr::default_letter());
        for (left, _) in &wrapped_extents {
            assert!(
                (left - geometry.margin_left).abs() < 0.01,
                "a wrapNone drawing must not indent any line, got {wrapped_extents:?}"
            );
        }
    }

    #[test]
    fn a_drawing_anchored_to_a_later_paragraph_still_pushes_text_aside() {
        use rdocx_oxml::drawing::{
            AnchorAlignH, AnchorAlignV, CT_Anchor, CT_Drawing, ST_RelativeFromH, ST_RelativeFromV,
            WrapType,
        };
        use rdocx_oxml::text::CT_R;
        use rdocx_oxml::units::Emu;

        // Word routinely anchors the arrow beside a paragraph to the paragraph
        // after it, which is what the external contribution's own sample does.
        let emu = |pt: f64| Emu((pt * 12700.0) as i64);
        let mut doc = rdocx_oxml::document::CT_Document::new();

        let mut first = CT_P::new();
        let mut body = String::new();
        for index in 0..40 {
            body.push_str(&format!("Sentence {index} of running text to fill lines. "));
        }
        first.add_run(&body);
        doc.body.add_paragraph(first);

        let mut second = CT_P::new();
        second.add_run("A later paragraph that owns the drawing.");
        let mut anchor = CT_Anchor::background("rId1", 0, 0);
        anchor.extent_cx = emu(100.0);
        anchor.extent_cy = emu(40.0);
        anchor.behind_doc = false;
        anchor.wrap = WrapType::Square;
        anchor.pos_h_relative_from = ST_RelativeFromH::Margin;
        anchor.pos_h_align = Some(AnchorAlignH::Left);
        // Margin-relative, so its position does not depend on where the
        // paragraph that owns it lands.
        anchor.pos_v_relative_from = ST_RelativeFromV::Margin;
        anchor.pos_v_align = Some(AnchorAlignV::Top);
        let mut drawing_run = CT_R::new("");
        drawing_run.content = vec![RunContent::Drawing(CT_Drawing {
            inline: None,
            anchor: Some(anchor),
        })];
        second.runs.push(drawing_run);
        doc.body.add_paragraph(second);

        let mut images = HashMap::new();
        images.insert(
            "rId1".to_string(),
            ImageData {
                data: vec![0u8; 8],
                content_type: "image/png".to_string(),
            },
        );

        let input = LayoutInput {
            revision_view: crate::input::RevisionView::Accepted,
            document: doc,
            styles: CT_Styles::new_default(),
            numbering: None,
            headers: HashMap::new(),
            footers: HashMap::new(),
            images,
            charts: HashMap::new(),
            chart_theme: oxml_drawing::theme::CT_OfficeStyleSheet::office_default(),
            chart_color_map: oxml_drawing::color::ColorMap::default(),
            core_properties: None,
            hyperlink_urls: HashMap::new(),
            footnotes: None,
            endnotes: None,
            theme: None,
            fonts: Vec::new(),
        };

        let mut engine = Engine::new();
        let output = engine.layout(&input).expect("layout succeeds");
        let geometry = sect_pr_to_geometry(&CT_SectPr::default_letter());
        let extents = text_extents(&output.pages[0]);

        assert!(!extents.is_empty(), "text renders");
        let expected_left = geometry.margin_left + 100.0;
        assert!(
            extents[0].0 >= expected_left - 1.0,
            "the first line of the earlier paragraph should clear the drawing at \
             {expected_left}, got {:?}",
            extents[0]
        );
    }

    #[test]
    fn a_split_paragraph_clearing_a_drawing_stays_inside_the_page() {
        use rdocx_oxml::drawing::{
            AnchorAlignH, AnchorAlignV, CT_Anchor, CT_Drawing, ST_RelativeFromH, ST_RelativeFromV,
            WrapType,
        };
        use rdocx_oxml::text::CT_R;
        use rdocx_oxml::units::Emu;

        // A top-and-bottom drawing pushes the paragraph's content down, and the
        // paragraph is long enough to split. The offset has to be counted where
        // the split point is decided, or the last lines run off the page.
        let emu = |pt: f64| Emu((pt * 12700.0) as i64);
        let mut doc = rdocx_oxml::document::CT_Document::new();
        let mut para = CT_P::new();
        let mut body = String::new();
        for index in 0..300 {
            body.push_str(&format!("Sentence {index} of a very long paragraph. "));
        }
        para.add_run(&body);

        let mut anchor = CT_Anchor::background("rId1", 0, 0);
        anchor.extent_cx = emu(200.0);
        anchor.extent_cy = emu(120.0);
        anchor.behind_doc = false;
        anchor.wrap = WrapType::TopAndBottom;
        anchor.pos_h_relative_from = ST_RelativeFromH::Margin;
        anchor.pos_h_align = Some(AnchorAlignH::Center);
        anchor.pos_v_relative_from = ST_RelativeFromV::Margin;
        anchor.pos_v_align = Some(AnchorAlignV::Top);
        anchor.dist_b = emu(10.0);
        let mut drawing_run = CT_R::new("");
        drawing_run.content = vec![RunContent::Drawing(CT_Drawing {
            inline: None,
            anchor: Some(anchor),
        })];
        para.runs.push(drawing_run);
        doc.body.add_paragraph(para);

        let mut images = HashMap::new();
        images.insert(
            "rId1".to_string(),
            ImageData {
                data: vec![0u8; 8],
                content_type: "image/png".to_string(),
            },
        );

        let input = LayoutInput {
            revision_view: crate::input::RevisionView::Accepted,
            document: doc,
            styles: CT_Styles::new_default(),
            numbering: None,
            headers: HashMap::new(),
            footers: HashMap::new(),
            images,
            charts: HashMap::new(),
            chart_theme: oxml_drawing::theme::CT_OfficeStyleSheet::office_default(),
            chart_color_map: oxml_drawing::color::ColorMap::default(),
            core_properties: None,
            hyperlink_urls: HashMap::new(),
            footnotes: None,
            endnotes: None,
            theme: None,
            fonts: Vec::new(),
        };

        let mut engine = Engine::new();
        let output = engine.layout(&input).expect("layout succeeds");
        let geometry = sect_pr_to_geometry(&CT_SectPr::default_letter());
        let bottom = geometry.page_height - geometry.margin_bottom;

        assert!(output.pages.len() > 1, "the paragraph must split");
        for (index, page) in output.pages.iter().enumerate() {
            for element in &page.elements {
                let PositionedElement::Text(run) = element else {
                    continue;
                };
                assert!(
                    run.origin.y <= bottom + 0.5,
                    "page {} draws text at {}, past the bottom margin at {bottom}",
                    index + 1,
                    run.origin.y
                );
            }
        }
    }

    // F-X017, notes broken to their own section's width.

    /// Text long enough to wrap at either measure under test, so a change of
    /// measure changes the number of lines rather than nothing at all.
    const NOTE_PROSE: &str = "A note long enough that the measure it is broken \
        to decides how many lines it occupies, which is the whole point of \
        breaking it to the width of the section that references it rather than \
        to the width of whichever section happens to come last in the document.";

    /// A document whose first section is `first_page_width` twips wide and
    /// whose body-level final section is letter portrait. The first section
    /// references note 1 and the second references note 2, and both notes carry
    /// the same text, so a difference in their line counts is a difference in
    /// the measure each was broken to.
    fn make_two_section_input(first_page_width: i32, endnotes_instead: bool) -> LayoutInput {
        use rdocx_oxml::footnotes::{CT_Footnote, CT_Footnotes, NoteType};
        use rdocx_oxml::text::CT_R;
        use rdocx_oxml::units::Twips;

        let note_of = |id: i32| {
            let mut note = CT_P::new();
            note.add_run(NOTE_PROSE);
            CT_Footnote {
                id,
                note_type: NoteType::Normal,
                paragraphs: vec![note],
            }
        };
        let reference = |id: i32| {
            let mut run = CT_R::new("");
            run.content = vec![if endnotes_instead {
                RunContent::EndnoteRef { id }
            } else {
                RunContent::FootnoteRef { id }
            }];
            run
        };

        let mut first_sect = CT_SectPr::default_letter();
        first_sect.page_width = Some(Twips(first_page_width));

        let mut doc = rdocx_oxml::document::CT_Document::new();

        // The paragraph carrying a sectPr is the one that ends its section.
        let mut first = CT_P::new();
        first.add_run("Body text in the first section");
        first.runs.push(reference(1));
        first.properties = Some(rdocx_oxml::properties::CT_PPr {
            sect_pr: Some(first_sect),
            ..Default::default()
        });
        doc.body.add_paragraph(first);

        let mut second = CT_P::new();
        second.add_run("Body text in the second section");
        second.runs.push(reference(2));
        doc.body.add_paragraph(second);
        doc.body.sect_pr = Some(CT_SectPr::default_letter());

        let stream = CT_Footnotes {
            footnotes: vec![note_of(1), note_of(2)],
        };
        LayoutInput {
            revision_view: crate::input::RevisionView::Accepted,
            document: doc,
            styles: CT_Styles::new_default(),
            numbering: None,
            headers: HashMap::new(),
            footers: HashMap::new(),
            images: HashMap::new(),
            charts: HashMap::new(),
            chart_theme: oxml_drawing::theme::CT_OfficeStyleSheet::office_default(),
            chart_color_map: oxml_drawing::color::ColorMap::default(),
            core_properties: None,
            hyperlink_urls: HashMap::new(),
            footnotes: (!endnotes_instead).then(|| stream.clone()),
            endnotes: endnotes_instead.then_some(stream),
            theme: None,
            fonts: Vec::new(),
        }
    }

    /// How many distinct baselines a page drew below its separator rule. A page
    /// without notes gives zero.
    ///
    /// This is the note's line count plus one for each note drawn, because a
    /// marker sits a rise above the line it belongs to and so has a baseline of
    /// its own. Every use below compares two of these counts over documents
    /// drawing the same number of notes, where the offset cancels.
    fn note_baseline_count(page: &oxml_layout::output::PageFrame) -> usize {
        let Some(separator_y) = separator_y_of(page) else {
            return 0;
        };
        let mut baselines: Vec<f64> = page
            .elements
            .iter()
            .filter_map(|element| match element {
                PositionedElement::Text(run) if run.origin.y > separator_y => Some(run.origin.y),
                _ => None,
            })
            .collect();
        baselines.sort_by(|a, b| a.partial_cmp(b).expect("baselines are finite"));
        baselines.dedup_by(|a, b| (*a - *b).abs() < 0.01);
        baselines.len()
    }

    #[test]
    fn a_note_is_broken_to_the_width_of_its_own_section() {
        // 17 inches wide against letter's 8.5, so the wide section's measure is
        // unmistakably different rather than different by a rounding.
        let output = Engine::new()
            .layout(&make_two_section_input(24480, false))
            .expect("layout succeeds");

        let wide = note_baseline_count(&output.pages[0]);
        let narrow = note_baseline_count(&output.pages[1]);

        assert!(wide > 0 && narrow > 0, "both sections must draw their note");
        assert!(
            wide < narrow,
            "the same note took {wide} lines in the wide section and {narrow} \
             in the narrow one, so both were broken to one measure"
        );
    }

    #[test]
    fn a_single_section_document_lays_notes_out_exactly_as_before() {
        // Two sections of identical geometry are the same document as one, so
        // the width key must collapse them. Any difference here is the fix
        // moving output it had no business moving.
        let two = Engine::new()
            .layout(&make_two_section_input(12240, false))
            .expect("layout succeeds");

        let single = Engine::new()
            .layout(&make_input_with_footnote(&[NOTE_PROSE]))
            .expect("layout succeeds");

        assert_eq!(
            note_baseline_count(&two.pages[0]),
            note_baseline_count(&single.pages[0]),
            "a note in a letter section stopped matching the same note in a \
             single-section letter document"
        );

        // And the same document laid out twice is still the same document.
        let again = Engine::new()
            .layout(&make_input_with_footnote(&[NOTE_PROSE]))
            .expect("layout succeeds");
        assert_eq!(single.pages.len(), again.pages.len());
        assert_eq!(single.pages[0].elements, again.pages[0].elements);
    }

    #[test]
    fn an_endnote_is_broken_to_the_final_sections_width() {
        // Endnotes are emitted after the last body page and drawn against the
        // final section's geometry, so that is the measure they must be broken
        // to even when the reference sits in a wider section.
        let wide_first = Engine::new()
            .layout(&make_two_section_input(24480, true))
            .expect("layout succeeds");
        let all_narrow = Engine::new()
            .layout(&make_two_section_input(12240, true))
            .expect("layout succeeds");

        // Endnotes are emitted on their own pages after every body page, and
        // this document has one short paragraph per section, so everything
        // drawn after the second page is endnote content.
        let endnote_lines = |output: &LayoutResult| {
            output.pages[2..]
                .iter()
                .map(|page| {
                    page.elements
                        .iter()
                        .filter(|element| matches!(element, PositionedElement::Text(_)))
                        .count()
                })
                .sum::<usize>()
        };

        assert_eq!(wide_first.pages.len(), all_narrow.pages.len());
        assert!(
            all_narrow.pages.len() > 2,
            "the endnotes must reach pages of their own"
        );
        assert!(endnote_lines(&all_narrow) > 0, "the endnotes must be drawn");
        assert_eq!(
            endnote_lines(&wide_first),
            endnote_lines(&all_narrow),
            "an endnote whose reference sits in a wide section was broken to \
             that section rather than to the final one it is drawn in"
        );
    }

    // F-X019, paragraph-relative drawings in later blocks should wrap.

    /// Two paragraphs, the second anchoring a wrapping drawing measured from
    /// `rel_v`. The first paragraph is the earlier text that should flow around
    /// it, which is the whole question: the drawing belongs to a block that has
    /// not been placed when the first paragraph is being laid out.
    fn make_lookahead_document(
        rel_v: rdocx_oxml::drawing::ST_RelativeFromV,
        wrap: rdocx_oxml::drawing::WrapType,
        off_v_pt: f64,
    ) -> LayoutInput {
        use rdocx_oxml::drawing::{CT_Anchor, CT_Drawing, ST_RelativeFromH};
        use rdocx_oxml::text::CT_R;
        use rdocx_oxml::units::Emu;

        let emu = |pt: f64| Emu((pt * 12700.0) as i64);

        let mut doc = rdocx_oxml::document::CT_Document::new();

        let mut first = CT_P::new();
        let mut body = String::new();
        for index in 0..30 {
            body.push_str(&format!(
                "Sentence {index} of running text that fills the paragraph out. "
            ));
        }
        first.add_run(&body);
        doc.body.add_paragraph(first);

        let mut second = CT_P::new();
        second.add_run("The paragraph the drawing is anchored to.");
        let mut anchor = CT_Anchor::background("rId1", 0, 0);
        anchor.extent_cx = emu(200.0);
        anchor.extent_cy = emu(120.0);
        anchor.behind_doc = false;
        anchor.wrap = wrap;
        anchor.pos_h_relative_from = ST_RelativeFromH::Margin;
        anchor.pos_h_align = Some(rdocx_oxml::drawing::AnchorAlignH::Right);
        anchor.pos_v_relative_from = rel_v;
        // The offset is measured from `rel_v`, so the two cases need different
        // numbers to land in the same band of the page. Above its own
        // paragraph for the paragraph-relative case, and a fixed way down the
        // page for the page-relative one. A drawing that lands below every line
        // of the first paragraph pushes nothing aside and would prove nothing.
        anchor.pos_v_offset = emu(off_v_pt);

        let mut drawing_run = CT_R::new("");
        drawing_run.content = vec![RunContent::Drawing(CT_Drawing {
            inline: None,
            anchor: Some(anchor),
        })];
        second.runs.push(drawing_run);
        doc.body.add_paragraph(second);

        let mut images = HashMap::new();
        images.insert(
            "rId1".to_string(),
            ImageData {
                data: vec![0u8; 8],
                content_type: "image/png".to_string(),
            },
        );

        LayoutInput {
            revision_view: crate::input::RevisionView::Accepted,
            document: doc,
            styles: CT_Styles::new_default(),
            numbering: None,
            headers: HashMap::new(),
            footers: HashMap::new(),
            images,
            charts: HashMap::new(),
            chart_theme: oxml_drawing::theme::CT_OfficeStyleSheet::office_default(),
            chart_color_map: oxml_drawing::color::ColorMap::default(),
            core_properties: None,
            hyperlink_urls: HashMap::new(),
            footnotes: None,
            endnotes: None,
            theme: None,
            fonts: Vec::new(),
        }
    }

    /// How many lines of body text the document drew, across every page.
    fn body_line_count(output: &LayoutResult) -> usize {
        output
            .pages
            .iter()
            .map(|page| text_extents(page).len())
            .sum()
    }

    #[test]
    fn a_paragraph_relative_wrapping_drawing_pushes_earlier_text_aside() {
        use rdocx_oxml::drawing::{ST_RelativeFromV, WrapType};

        // The same document twice, differing only in whether the drawing
        // wraps. Narrowed lines hold less text, so the paragraph needs more of
        // them, and that is visible without depending on where any one line
        // broke.
        let wrapping = Engine::new()
            .layout(&make_lookahead_document(
                ST_RelativeFromV::Paragraph,
                WrapType::Square,
                -120.0,
            ))
            .expect("layout succeeds");
        let ignoring = Engine::new()
            .layout(&make_lookahead_document(
                ST_RelativeFromV::Paragraph,
                WrapType::None,
                -120.0,
            ))
            .expect("layout succeeds");

        assert!(
            body_line_count(&wrapping) > body_line_count(&ignoring),
            "the earlier paragraph took {} lines against {}, so it flowed \
             through the drawing rather than around it",
            body_line_count(&wrapping),
            body_line_count(&ignoring)
        );
    }

    #[test]
    fn a_page_relative_drawing_in_a_later_block_still_wraps() {
        use rdocx_oxml::drawing::{ST_RelativeFromV, WrapType};

        // F-X016's case, which the second pass must not disturb. This document
        // has no paragraph-relative wrap, so it paginates in one pass.
        let wrapping = Engine::new()
            .layout(&make_lookahead_document(
                ST_RelativeFromV::Page,
                WrapType::Square,
                150.0,
            ))
            .expect("layout succeeds");
        let ignoring = Engine::new()
            .layout(&make_lookahead_document(
                ST_RelativeFromV::Page,
                WrapType::None,
                150.0,
            ))
            .expect("layout succeeds");

        assert!(body_line_count(&wrapping) > body_line_count(&ignoring));
    }

    #[test]
    fn a_second_pass_is_stable_for_the_document_that_earns_it() {
        use rdocx_oxml::drawing::{ST_RelativeFromV, WrapType};

        // Two passes, not a fixed point, so the guarantee is that the answer is
        // the same answer every time rather than that it has converged.
        let build = || {
            Engine::new()
                .layout(&make_lookahead_document(
                    ST_RelativeFromV::Paragraph,
                    WrapType::Square,
                    -120.0,
                ))
                .expect("layout succeeds")
        };
        let first = build();
        let second = build();

        assert_eq!(first.pages.len(), second.pages.len());
        for (index, page) in first.pages.iter().enumerate() {
            assert_eq!(
                page.elements,
                second.pages[index].elements,
                "page {} differs between two runs",
                index + 1
            );
        }
    }

    fn cross_reference_run(instruction: &str, display: &str) -> rdocx_oxml::text::CT_R {
        let mut run = rdocx_oxml::text::CT_R::new("");
        run.content = vec![RunContent::Field(Field::new(instruction, display))];
        run
    }

    fn target_paragraph(targets: &[(i32, &str, usize, usize)], text: &str, hidden: bool) -> CT_P {
        let mut paragraph = CT_P::new();
        paragraph.properties = Some(CT_PPr {
            page_break_before: Some(true),
            ..Default::default()
        });
        let mut run = rdocx_oxml::text::CT_R::new(text);
        if hidden {
            run.properties = Some(rdocx_oxml::properties::CT_RPr {
                vanish: Some(true),
                ..Default::default()
            });
        }
        paragraph.runs.push(run);
        for (id, name, start, end) in targets {
            assert!(paragraph.insert_bookmark_start(*start, *id, name));
            assert!(paragraph.insert_bookmark_end(*end, *id));
        }
        paragraph
    }

    fn output_text(output: &LayoutResult) -> Vec<String> {
        let mut text = Vec::new();
        for page in &output.pages {
            oxml_layout::walk(&page.elements, &mut |element, _| {
                if let PositionedElement::Text(run) = element {
                    text.push(run.text.clone());
                }
            });
        }
        text
    }

    fn deterministic_layout(input: &LayoutInput) -> LayoutResult {
        Engine::new_deterministic()
            .expect("bundled fonts")
            .layout(input)
            .expect("layout succeeds")
    }

    #[test]
    fn an_unsupported_complex_field_keeps_its_cached_display() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>DATE</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>17 August 2026</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:body></w:document>"#;
        let mut input = make_input_with_text("");
        input.document = rdocx_oxml::CT_Document::from_xml(xml).expect("field document parses");

        let text = output_text(&deterministic_layout(&input));
        assert!(text.concat().contains("17 August 2026"), "{text:?}");
    }

    #[test]
    fn a_complex_field_keeps_each_cached_result_runs_formatting() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>DATE</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>italic</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:body></w:document>"#;
        let mut input = make_input_with_text("");
        input.document = rdocx_oxml::CT_Document::from_xml(xml).expect("field document parses");

        let output = deterministic_layout(&input);
        let mut displays = Vec::new();
        for page in &output.pages {
            oxml_layout::walk(&page.elements, &mut |element, _| {
                if let PositionedElement::Text(run) = element
                    && matches!(run.text.as_str(), "bold" | "italic")
                {
                    displays.push((run.text.clone(), run.bold, run.italic));
                }
            });
        }
        assert_eq!(
            displays,
            vec![
                ("bold".to_owned(), true, false),
                ("italic".to_owned(), false, true)
            ]
        );
    }

    #[test]
    fn a_computed_complex_field_keeps_its_cached_result_run_formatting() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>PAGE</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:b/><w:i/></w:rPr><w:t>99</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:body></w:document>"#;
        let mut input = make_input_with_text("");
        input.document = rdocx_oxml::CT_Document::from_xml(xml).expect("field document parses");
        let BodyContent::Paragraph(paragraph) = &mut input.document.body.content[0] else {
            panic!("expected paragraph")
        };
        let RunContent::Field(field) = &mut paragraph.runs[0].content[0] else {
            panic!("expected field")
        };
        field.cached_result = "edited stored value".to_owned();

        let output = deterministic_layout(&input);
        let mut displays = Vec::new();
        for page in &output.pages {
            oxml_layout::walk(&page.elements, &mut |element, _| {
                if let PositionedElement::Text(run) = element
                    && run.text == "1"
                {
                    displays.push((run.bold, run.italic));
                }
            });
        }
        assert_eq!(displays, vec![(true, true)]);
    }

    #[test]
    fn a_pageref_inside_a_table_uses_the_final_target_page() {
        use rdocx_oxml::table::{CT_Row, CT_Tbl, CT_Tc, CellContent};

        let mut input = make_input_with_text("");
        input.document.body.content.clear();
        let mut field = CT_P::new();
        field
            .runs
            .push(cross_reference_run("PAGEREF destination", "cached"));
        let mut cell = CT_Tc::new();
        cell.content = vec![CellContent::Paragraph(field)];
        let mut row = CT_Row::new();
        row.cells.push(cell);
        let mut table = CT_Tbl::new();
        table.rows.push(row);
        input.document.body.content.push(BodyContent::Table(table));
        input.document.body.add_paragraph(target_paragraph(
            &[(4, "destination", 0, 1)],
            "target",
            false,
        ));

        let output = deterministic_layout(&input);
        let text = output_text(&output);
        assert!(text.iter().any(|value| value == "2"), "{text:?}");
        assert!(!text.iter().any(|value| value == "cached"), "{text:?}");
    }

    #[test]
    fn a_resolved_pageref_uses_a_fixed_pagination_placeholder() {
        let build = |display: &str| {
            let mut input = make_input_with_text("");
            input.document.body.content.clear();
            let mut field = CT_P::new();
            field
                .runs
                .push(cross_reference_run("PAGEREF destination", display));
            input.document.body.add_paragraph(field);
            input.document.body.add_paragraph(target_paragraph(
                &[(4, "destination", 0, 1)],
                "target",
                false,
            ));
            deterministic_layout(&input)
        };
        let short = build("7");
        let long = build(&"stale display ".repeat(1000));

        assert_eq!(short.pages.len(), long.pages.len());
        assert_eq!(output_text(&short), output_text(&long));
    }

    #[test]
    fn every_target_at_a_paragraph_end_is_retained() {
        let mut input = make_input_with_text("");
        input.document.body.content.clear();
        let mut fields = CT_P::new();
        for name in ["first", "second"] {
            fields
                .runs
                .push(cross_reference_run(&format!("PAGEREF {name}"), "cached"));
        }
        input.document.body.add_paragraph(fields);
        input.document.body.add_paragraph(target_paragraph(
            &[(4, "first", 1, 1), (5, "second", 1, 1)],
            "target",
            false,
        ));

        let text = output_text(&deterministic_layout(&input));
        assert_eq!(
            text.iter().filter(|value| value.as_str() == "2").count(),
            2,
            "{text:?}"
        );
    }

    #[test]
    fn a_target_before_hidden_text_is_retained() {
        let mut input = make_input_with_text("");
        input.document.body.content.clear();
        let mut field = CT_P::new();
        field
            .runs
            .push(cross_reference_run("PAGEREF destination", "cached"));
        input.document.body.add_paragraph(field);
        input.document.body.add_paragraph(target_paragraph(
            &[(4, "destination", 0, 1)],
            "hidden target",
            true,
        ));

        let text = output_text(&deterministic_layout(&input));
        assert!(text.iter().any(|value| value == "2"), "{text:?}");
        assert!(!text.iter().any(|value| value == "cached"), "{text:?}");
    }
}

fn fingerprint_paragraph(
    para: &rdocx_oxml::text::CT_P,
    content_width: f64,
    revision_view: u8,
) -> u64 {
    let mut fp = Fingerprint::new();
    fp.eat_paragraph(para);
    fp.eat(&content_width.to_bits().to_le_bytes());
    fp.eat(&[revision_view]);
    fp.finish()
}

/// FNV-1a folding with a scratch buffer for the `Debug`-formatted pieces, so
/// every fingerprint walker shares one hashing discipline.
struct Fingerprint {
    h: u64,
    buf: String,
}

impl Fingerprint {
    fn new() -> Self {
        Fingerprint {
            h: 0xcbf2_9ce4_8422_2325,
            buf: String::new(),
        }
    }

    fn eat(&mut self, bytes: &[u8]) {
        for b in bytes {
            self.h ^= u64::from(*b);
            self.h = self.h.wrapping_mul(0x0000_0100_0000_01b3);
        }
    }

    fn eat_debug<T: std::fmt::Debug>(&mut self, value: &T) {
        use std::fmt::Write as _;
        let mut buf = std::mem::take(&mut self.buf);
        buf.clear();
        let _ = write!(buf, "{value:?}");
        self.eat(buf.as_bytes());
        self.buf = buf;
    }

    /// The paragraph walk shared by the block cache and the table
    /// fingerprint: run text as raw bytes (the hot path), the small or
    /// usually-empty property/projection structures via `Debug`.
    fn eat_paragraph(&mut self, para: &rdocx_oxml::text::CT_P) {
        use rdocx_oxml::text::RunContent;

        self.eat_debug(&para.properties);
        for run in &para.runs {
            self.eat(b"\x01r");
            self.eat_debug(&run.properties);
            for content in &run.content {
                match content {
                    RunContent::Text(t) => {
                        self.eat(b"\x02t");
                        self.eat(t.text.as_bytes());
                        self.eat(&[u8::from(t.preserve_space)]);
                    }
                    // A field's Debug form includes a per-parse source_id
                    // (a global counter), so hashing it wholesale would give
                    // the same XML a different fingerprint on every
                    // build_input. Hash what layout actually reads instead.
                    RunContent::Field(field) => {
                        self.eat(b"\x02f");
                        self.eat(field.instruction.raw.as_bytes());
                        self.eat(field.cached_result.as_bytes());
                        self.eat_debug(&field.dirty);
                    }
                    other => {
                        self.eat(b"\x02o");
                        self.eat_debug(other);
                    }
                }
            }
        }
        if !(para.hyperlinks.is_empty()
            && para.comment_ranges.is_empty()
            && para.bookmark_markers.is_empty()
            && para.extra_xml.is_empty()
            && para.content_controls.is_empty()
            && para.revisions.is_empty())
        {
            self.eat_debug(&para.hyperlinks);
            self.eat_debug(&para.comment_ranges);
            self.eat_debug(&para.bookmark_markers);
            self.eat_debug(&para.extra_xml);
            self.eat_debug(&para.content_controls);
            self.eat_debug(&para.revisions);
        }
    }

    /// A header/footer part, walked so fields inside it hash by content
    /// (see the `RunContent::Field` arm of `eat_paragraph`).
    fn eat_hdr_ftr(&mut self, part: &rdocx_oxml::header_footer::CT_HdrFtr) {
        for para in &part.paragraphs {
            self.eat(b"\x01P");
            self.eat_paragraph(para);
        }
        if !(part.extra_namespaces.is_empty() && part.extra_xml.is_empty()) {
            self.eat_debug(&part.extra_namespaces);
            self.eat_debug(&part.extra_xml);
        }
    }

    /// A footnotes/endnotes part, walked for the same reason.
    fn eat_notes(&mut self, part: &Option<rdocx_oxml::footnotes::CT_Footnotes>) {
        let Some(part) = part else {
            self.eat(b"\x01-");
            return;
        };
        for note in &part.footnotes {
            self.eat(b"\x01N");
            self.eat_debug(&note.id);
            self.eat_debug(&note.note_type);
            for para in &note.paragraphs {
                self.eat_paragraph(para);
            }
        }
    }

    fn finish(&self) -> u64 {
        self.h
    }
}

/// Content fingerprint of a table for pagination identity, walking rows,
/// cells and cell paragraphs the same way `fingerprint_paragraph` walks a
/// body paragraph. Nested tables and content controls are rare enough to go
/// through `Debug` wholesale.
fn fingerprint_table(
    tbl: &rdocx_oxml::table::CT_Tbl,
    content_width: f64,
    revision_view: u8,
) -> u64 {
    use rdocx_oxml::table::CellContent;

    let mut fp = Fingerprint::new();
    fp.eat_debug(&tbl.properties);
    fp.eat_debug(&tbl.grid);
    if !(tbl.extra_xml.is_empty() && tbl.content_controls.is_empty()) {
        fp.eat_debug(&tbl.extra_xml);
        fp.eat_debug(&tbl.content_controls);
    }
    for row in &tbl.rows {
        fp.eat(b"\x01R");
        fp.eat_debug(&row.properties);
        if !(row.extra_xml.is_empty() && row.content_controls.is_empty()) {
            fp.eat_debug(&row.extra_xml);
            fp.eat_debug(&row.content_controls);
        }
        for cell in &row.cells {
            fp.eat(b"\x02C");
            fp.eat_debug(&cell.properties);
            if !cell.extra_xml.is_empty() {
                fp.eat_debug(&cell.extra_xml);
            }
            for content in &cell.content {
                match content {
                    CellContent::Paragraph(p) => {
                        fp.eat(b"\x03p");
                        fp.eat_paragraph(p);
                    }
                    other => {
                        fp.eat(b"\x03x");
                        fp.eat_debug(other);
                    }
                }
            }
        }
    }
    fp.eat(&content_width.to_bits().to_le_bytes());
    fp.eat(&[revision_view]);
    fp.finish()
}

/// Whether any of these paragraph blocks renders a numbering marker or a
/// note reference — the cross-block state that makes a block unsafe to
/// cache (a cache hit would skip the NumberingState / note-order advance
/// that produced the marker text).
fn para_blocks_render_shared_state(blocks: &[ParagraphBlock]) -> bool {
    use oxml_layout::LineItem;
    blocks.iter().any(|para| {
        para.lines.iter().any(|line| {
            line.items.iter().any(|item| match item {
                LineItem::Marker(_) => true,
                LineItem::Text(seg) => seg.note.is_some(),
                _ => false,
            })
        })
    })
}

/// Whether any cell paragraph renders a numbering marker or a note
/// reference — the cross-block state that makes a table unsafe to cache.
fn table_renders_shared_state(table: &crate::table::TableBlock) -> bool {
    table.rows.iter().any(|row| {
        row.cells
            .iter()
            .any(|cell| para_blocks_render_shared_state(&cell.paragraphs))
    })
}

/// One pagination-identity value from a block's source fingerprint and its
/// rendered-marker salt.
fn combine_fp(source: u64, salt: u64) -> u64 {
    source.wrapping_mul(0x9E37_79B9_7F4A_7C15) ^ salt
}

/// What a block renders from cross-block state that its source fingerprint
/// cannot see: list numbering markers (NumberingState) and note reference
/// markers (note numbering). Folding the rendered text of both into the
/// pagination fingerprint makes "same fingerprint" mean "paginates AND
/// renders identically", so inserting a numbered paragraph or a footnote
/// reference invalidates every block whose marker text shifts.
fn pagination_salt(block: &LayoutBlock) -> u64 {
    use oxml_layout::LineItem;

    fn eat_lines(fp: &mut Fingerprint, lines: &[oxml_layout::LayoutLine]) {
        for line in lines {
            for item in &line.items {
                match item {
                    LineItem::Marker(seg) => {
                        fp.eat(b"\x01m");
                        fp.eat(seg.text.as_bytes());
                    }
                    LineItem::Text(seg) if seg.note.is_some() => {
                        fp.eat(b"\x01n");
                        fp.eat(seg.text.as_bytes());
                        fp.eat_debug(&seg.note);
                    }
                    _ => {}
                }
            }
        }
    }

    let mut fp = Fingerprint::new();
    match block {
        LayoutBlock::Paragraph(para) => eat_lines(&mut fp, &para.lines),
        LayoutBlock::Table(table) => {
            for row in &table.rows {
                for cell in &row.cells {
                    for para in &cell.paragraphs {
                        eat_lines(&mut fp, &para.lines);
                    }
                }
            }
        }
    }
    fp.finish()
}

/// Fold the result-local source-node table: reused pages carry baked
/// `SourceSpan`s, which stay valid only while ids and paths are unchanged.
fn fingerprint_source_nodes(nodes: &[WordSourcePath]) -> u64 {
    let mut fp = Fingerprint::new();
    for node in nodes {
        fp.eat_debug(node);
    }
    fp.finish()
}
