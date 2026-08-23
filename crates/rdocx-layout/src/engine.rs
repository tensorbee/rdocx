//! Layout engine orchestrator: ties all phases together.

use std::collections::{HashMap, VecDeque};
use std::fmt::{self, Write as _};
use std::sync::Arc;

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
    paragraph_cache_context: Option<ReusableEngineContext>,
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
    /// Requested-name -> family aliases forwarded to the font manager on
    /// every layout (no bytes; see FontManager::set_caller_aliases).
    caller_font_aliases: Vec<(String, String)>,
    table_cache: VecDeque<TableCacheEntry>,
    table_cache_bytes: usize,
    table_cache_hits: usize,
    table_cache_builds: usize,
    pending_table_cache: Option<VecDeque<TableCacheEntry>>,
    pending_table_cache_bytes: usize,
    #[cfg(test)]
    pending_table_cache_peak_entries: usize,
    #[cfg(test)]
    pending_table_cache_peak_bytes: usize,
    header_footer_cache: VecDeque<HeaderFooterCacheEntry>,
    header_footer_cache_bytes: usize,
    header_footer_cache_hits: usize,
    header_footer_cache_builds: usize,
    pending_header_footer_cache: Option<VecDeque<HeaderFooterCacheEntry>>,
    pending_header_footer_cache_bytes: usize,
    #[cfg(test)]
    pending_header_footer_cache_peak_entries: usize,
    #[cfg(test)]
    pending_header_footer_cache_peak_bytes: usize,
    header_footer_cache_reads_enabled: bool,
    restart_cache: Option<RestartCache>,
    #[cfg(test)]
    last_rebuilt_page_range: Option<std::ops::Range<usize>>,
}

#[derive(Clone, PartialEq)]
struct ReusableEngineContext {
    revision_view: RevisionView,
    has_wrapping_drawing: bool,
    styles: CT_Styles,
    numbering: Option<rdocx_oxml::numbering::CT_Numbering>,
    sections: Vec<CT_SectPr>,
    headers: HashMap<String, rdocx_oxml::header_footer::CT_HdrFtr>,
    footers: HashMap<String, rdocx_oxml::header_footer::CT_HdrFtr>,
    images: HashMap<String, crate::input::ImageData>,
    charts: HashMap<String, std::result::Result<Box<oxml_chart::CT_ChartSpace>, String>>,
    chart_theme: oxml_drawing::theme::CT_OfficeStyleSheet,
    chart_color_map: oxml_drawing::color::ColorMap,
    core_properties: Option<rdocx_oxml::core_properties::CoreProperties>,
    hyperlink_urls: HashMap<String, String>,
    footnotes: Option<rdocx_oxml::footnotes::CT_Footnotes>,
    endnotes: Option<rdocx_oxml::footnotes::CT_Footnotes>,
    theme: Option<rdocx_oxml::theme::Theme>,
    fonts: Vec<oxml_layout::FontFile>,
    background_xml: Option<Vec<u8>>,
}

impl ReusableEngineContext {
    fn for_input(input: &LayoutInput) -> Self {
        let mut sections = input
            .document
            .body
            .content
            .iter()
            .filter_map(|content| match content {
                BodyContent::Paragraph(paragraph) => paragraph
                    .properties
                    .as_ref()
                    .and_then(|properties| properties.sect_pr.clone()),
                _ => None,
            })
            .collect::<Vec<_>>();
        sections.extend(input.document.body.sect_pr.iter().cloned());
        Self {
            revision_view: input.revision_view,
            has_wrapping_drawing: document_has_wrapping_drawing(input),
            styles: input.styles.clone(),
            numbering: input.numbering.clone(),
            sections,
            headers: input.headers.clone(),
            footers: input.footers.clone(),
            images: input.images.clone(),
            charts: input.charts.clone(),
            chart_theme: input.chart_theme.clone(),
            chart_color_map: input.chart_color_map.clone(),
            core_properties: input.core_properties.clone(),
            hyperlink_urls: input.hyperlink_urls.clone(),
            footnotes: input.footnotes.clone(),
            endnotes: input.endnotes.clone(),
            theme: input.theme.clone(),
            fonts: input.fonts.clone(),
            background_xml: input.document.background_xml.clone(),
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
    fingerprint: u64,
    key: ParagraphCacheKey,
    block: ParagraphBlock,
    diagnostics: Vec<Diagnostic>,
    font_trace: Vec<FontId>,
    bytes: usize,
}

#[derive(Clone, PartialEq)]
struct TableCacheKey {
    table: CT_Tbl,
    content_width_bits: u64,
    revision_view: RevisionView,
    with_provenance: bool,
}

struct TableCacheEntry {
    key: TableCacheKey,
    block: table::TableBlock,
    diagnostics: Vec<Diagnostic>,
    font_trace: Vec<FontId>,
    bytes: usize,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum HeaderFooterStoryKind {
    Header,
    Footer,
}

#[derive(Clone, PartialEq)]
struct HeaderFooterCacheKey {
    story: HeaderFooterStoryKind,
    variant: HdrFtrType,
    section: CT_SectPr,
    relationship_id: String,
    part: rdocx_oxml::header_footer::CT_HdrFtr,
    resolved_part_bytes: Vec<u8>,
    with_provenance: bool,
}

#[derive(Clone)]
struct HeaderFooterVariantContent {
    blocks: Vec<ParagraphBlock>,
    watermark: Option<GroupElement>,
}

struct HeaderFooterCacheEntry {
    key: HeaderFooterCacheKey,
    content: HeaderFooterVariantContent,
    diagnostics: Vec<Diagnostic>,
    font_trace: Vec<FontId>,
    bytes: usize,
}

struct RestartCache {
    body: Vec<String>,
    with_provenance: bool,
    raw_pages: Vec<Arc<PageFrame>>,
    pages: Vec<Arc<PageFrame>>,
    substitution_inputs: Vec<Option<FieldSubstitutionInputs>>,
    outlines: Vec<oxml_layout::OutlineEntry>,
    checkpoints: Vec<paginator::PaginationCheckpoint>,
    font_trace: Vec<FontId>,
    bytes: usize,
}

#[derive(Clone, PartialEq, Eq)]
struct FieldSubstitutionInputs {
    page_index: usize,
    page_number: usize,
    total_pages: usize,
    bookmark_pages: Vec<(usize, usize)>,
    font_identity: Vec<FontId>,
    revision_view: RevisionView,
}

const CACHE_MAX_ENTRIES: usize = 4_224;
const CACHE_MAX_BYTES: usize = 64 * 1024 * 1024;
const PARAGRAPH_CACHE_MAX_ENTRIES: usize = 4_096;
const PARAGRAPH_CACHE_MAX_BYTES: usize = 56 * 1024 * 1024;
const TABLE_CACHE_MAX_ENTRIES: usize = 32;
const TABLE_CACHE_MAX_BYTES: usize = 2 * 1024 * 1024;
const HEADER_FOOTER_CACHE_MAX_ENTRIES: usize = 64;
const HEADER_FOOTER_CACHE_MAX_BYTES: usize = 4 * 1024 * 1024;
const RESTART_CACHE_MAX_ENTRIES: usize = 32;
const RESTART_CACHE_MAX_BYTES: usize = 2 * 1024 * 1024;
const _: () = assert!(PARAGRAPH_CACHE_MAX_ENTRIES == 4_096);
const _: () = assert!(PARAGRAPH_CACHE_MAX_BYTES == 56 * 1024 * 1024);
const _: () = assert!(HEADER_FOOTER_CACHE_MAX_ENTRIES == 64);
const _: () = assert!(HEADER_FOOTER_CACHE_MAX_BYTES == 4 * 1024 * 1024);
const _: () = assert!(CACHE_MAX_ENTRIES == 4_224);
const _: () = assert!(CACHE_MAX_BYTES == 64 * 1024 * 1024);
const _: () = assert!(
    PARAGRAPH_CACHE_MAX_ENTRIES
        + TABLE_CACHE_MAX_ENTRIES
        + HEADER_FOOTER_CACHE_MAX_ENTRIES
        + RESTART_CACHE_MAX_ENTRIES
        <= CACHE_MAX_ENTRIES
);
const _: () = assert!(
    PARAGRAPH_CACHE_MAX_BYTES
        + TABLE_CACHE_MAX_BYTES
        + HEADER_FOOTER_CACHE_MAX_BYTES
        + RESTART_CACHE_MAX_BYTES
        <= CACHE_MAX_BYTES
);
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
            caller_font_aliases: Vec::new(),
            table_cache: VecDeque::new(),
            table_cache_bytes: 0,
            table_cache_hits: 0,
            table_cache_builds: 0,
            pending_table_cache: None,
            pending_table_cache_bytes: 0,
            #[cfg(test)]
            pending_table_cache_peak_entries: 0,
            #[cfg(test)]
            pending_table_cache_peak_bytes: 0,
            header_footer_cache: VecDeque::new(),
            header_footer_cache_bytes: 0,
            header_footer_cache_hits: 0,
            header_footer_cache_builds: 0,
            pending_header_footer_cache: None,
            pending_header_footer_cache_bytes: 0,
            #[cfg(test)]
            pending_header_footer_cache_peak_entries: 0,
            #[cfg(test)]
            pending_header_footer_cache_peak_bytes: 0,
            header_footer_cache_reads_enabled: false,
            restart_cache: None,
            #[cfg(test)]
            last_rebuilt_page_range: None,
        }
    }

    /// Set the requested-name -> family font aliases consulted during
    /// resolution. Carries no font bytes, so passing a large mapping every
    /// layout is cheap; a change invalidates font-derived caches on the
    /// next layout.
    pub fn set_caller_font_aliases(&mut self, aliases: Vec<(String, String)>) {
        self.caller_font_aliases = aliases;
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

    /// Take a reusable engine only when its complete retained-work context
    /// matches the proposed receiver input.
    #[doc(hidden)]
    pub fn take_if_compatible(source: &mut Option<Self>, input: &LayoutInput) -> Option<Self> {
        let compatible = source.as_ref().is_some_and(|engine| {
            engine.paragraph_cache_context.as_ref()
                == Some(&ReusableEngineContext::for_input(input))
                && engine.pending_paragraph_cache.is_none()
                && engine.pending_header_footer_cache.is_none()
        });
        compatible.then(|| source.take()).flatten()
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
        let fonts_changed = self.font_manager.load_additional_fonts(&input.fonts)
            | self
                .font_manager
                .set_caller_aliases(&self.caller_font_aliases);
        self.font_manager.begin_layout();

        let paragraph_context = ReusableEngineContext::for_input(input);
        if fonts_changed {
            self.paragraph_cache.clear();
            self.paragraph_cache_bytes = 0;
            self.table_cache.clear();
            self.table_cache_bytes = 0;
            self.header_footer_cache.clear();
            self.header_footer_cache_bytes = 0;
        }
        let context_matches =
            !fonts_changed && self.paragraph_cache_context.as_ref() == Some(&paragraph_context);
        self.paragraph_cache_reads_enabled = context_matches;
        self.header_footer_cache_reads_enabled = context_matches;
        self.pending_paragraph_cache = Some(VecDeque::new());
        self.pending_paragraph_cache_bytes = 0;
        self.pending_table_cache = Some(VecDeque::new());
        self.pending_table_cache_bytes = 0;
        self.pending_header_footer_cache = Some(VecDeque::new());
        self.pending_header_footer_cache_bytes = 0;
        #[cfg(test)]
        {
            self.pending_paragraph_cache_peak_entries = 0;
            self.pending_paragraph_cache_peak_bytes = 0;
            self.pending_table_cache_peak_entries = 0;
            self.pending_table_cache_peak_bytes = 0;
            self.pending_header_footer_cache_peak_entries = 0;
            self.pending_header_footer_cache_peak_bytes = 0;
        }

        let result = self.layout_transaction(input, sources);
        let pending = self.pending_paragraph_cache.take().unwrap_or_default();
        let pending_tables = self.pending_table_cache.take().unwrap_or_default();
        let pending_header_footers = self.pending_header_footer_cache.take().unwrap_or_default();
        self.pending_paragraph_cache_bytes = 0;
        self.pending_table_cache_bytes = 0;
        self.pending_header_footer_cache_bytes = 0;
        self.paragraph_cache_reads_enabled = false;
        self.header_footer_cache_reads_enabled = false;
        if result.is_ok() {
            if !context_matches {
                self.paragraph_cache.clear();
                self.paragraph_cache_bytes = 0;
                self.table_cache.clear();
                self.table_cache_bytes = 0;
                self.header_footer_cache.clear();
                self.header_footer_cache_bytes = 0;
                self.paragraph_cache_context = Some(paragraph_context);
            }
            for entry in pending {
                self.publish_paragraph_cache_entry(entry);
            }
            for entry in pending_tables {
                self.publish_table_cache_entry(entry);
            }
            for entry in pending_header_footers {
                self.publish_header_footer_cache_entry(entry);
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
        self.table_cache.retain(|entry| {
            entry
                .font_trace
                .iter()
                .all(|font_id| current_fonts.contains(font_id))
        });
        self.table_cache_bytes = self.table_cache.iter().map(|entry| entry.bytes).sum();
        self.header_footer_cache.retain(|entry| {
            entry
                .font_trace
                .iter()
                .all(|font_id| current_fonts.contains(font_id))
        });
        self.header_footer_cache_bytes = self
            .header_footer_cache
            .iter()
            .map(|entry| entry.bytes)
            .sum();
        self.font_manager.retain_current_fonts();
        result
    }

    fn layout_transaction(
        &mut self,
        input: &LayoutInput,
        sources: Option<&SourceRegistry>,
    ) -> Result<LayoutResult> {
        let retained_context_matches = self.paragraph_cache_reads_enabled;
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
        let mut current_sect_pr: Option<CT_SectPr> = None; // Will be set from paragraph sect_pr

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

                    current_blocks.push(LayoutBlock::Paragraph(para_block));

                    // If this paragraph has sect_pr, it ends a section
                    if let Some(sect_pr) = para_sect_pr {
                        let geometry = sect_pr_to_geometry(&sect_pr);
                        let header_footer = layout_header_footer(
                            self,
                            &sect_pr,
                            input,
                            styles,
                            &media,
                            &mut num_state,
                            &mut diagnostics,
                            sources,
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

                    let table_block = self.layout_body_table(
                        tbl,
                        geometry.content_width(),
                        styles,
                        input,
                        &media,
                        &mut num_state,
                        &mut diagnostics,
                        sources,
                        &WordStory::Document,
                        &[body_index],
                    )?;
                    current_blocks.push(LayoutBlock::Table(table_block));
                }
                _ => {} // Skip RawXml elements during layout
            }
        }

        // Remaining blocks belong to the final section
        let final_geometry = sect_pr_to_geometry(&final_sect_pr);
        let final_hf = layout_header_footer(
            self,
            &final_sect_pr,
            input,
            styles,
            &media,
            &mut num_state,
            &mut diagnostics,
            sources,
        )?;
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

        let mut body = input
            .document
            .body
            .content
            .iter()
            .map(|content| format!("{content:?}"))
            .collect::<Vec<_>>();
        let mut font_trace = self.font_manager.current_layout_fonts().to_vec();
        let restart_record_eligible = sections.len() == 1
            && input.document.background_xml.is_none()
            && input.footnotes.is_none()
            && input.endnotes.is_none()
            && !document_wraps
            && sections[0].header_footer.is_none()
            && input.document.body.content.iter().all(|content| {
                matches!(content, BodyContent::Paragraph(_) | BodyContent::Table(_))
            })
            && sections[0].blocks.iter().all(restart_record_block_is_safe);
        let restart_eligible =
            restart_record_eligible && sections[0].blocks.iter().all(restart_block_is_safe);
        let reusable_restart_record = restart_record_eligible
            && retained_context_matches
            && self.restart_cache.as_ref().is_some_and(|cache| {
                cache.font_trace == font_trace
                    && cache.with_provenance == sources.is_some()
                    && (sources.is_none() || cache.body.len() == body.len())
            });
        let reusable_restart = restart_eligible && reusable_restart_record;
        let first_changed = reusable_restart.then(|| {
            let cache = self.restart_cache.as_ref().expect("restart cache exists");
            body.iter()
                .zip(&cache.body)
                .position(|(current, previous)| current != previous)
                .unwrap_or_else(|| body.len().min(cache.body.len()))
        });
        let restart_checkpoint = first_changed.and_then(|first_changed| {
            self.restart_cache
                .as_ref()
                .expect("restart cache exists")
                .checkpoints
                .iter()
                .rev()
                .find(|checkpoint| checkpoint.next_block_index <= first_changed)
                .copied()
        });
        let tail_source = restart_checkpoint.and_then(|restart| {
            let cache = self.restart_cache.as_ref().expect("restart cache exists");
            let common_suffix = body
                .iter()
                .rev()
                .zip(cache.body.iter().rev())
                .take_while(|(current, previous)| current == previous)
                .count();
            let new_tail = body.len() - common_suffix;
            let old_tail = cache.body.len() - common_suffix;
            cache
                .checkpoints
                .iter()
                .find(|checkpoint| {
                    checkpoint.next_block_index == old_tail && new_tail > restart.next_block_index
                })
                .copied()
                .map(|old| {
                    (
                        paginator::PaginationCheckpoint {
                            next_block_index: new_tail,
                            page_count: old.page_count,
                            next_header_page_number: old.next_header_page_number,
                        },
                        old,
                    )
                })
        });

        let (mut pages, mut outlines, mut checkpoints) = if restart_eligible {
            let recorded = paginator::paginate_single_section_recorded(
                &sections[0],
                &self.font_manager,
                &media,
                &notes,
                restart_checkpoint,
                tail_source.map(|(stop, _)| stop),
            );
            let mut pages = restart_checkpoint.map_or_else(Vec::new, |checkpoint| {
                self.restart_cache
                    .as_ref()
                    .expect("a restart checkpoint belongs to retained pages")
                    .raw_pages[..checkpoint.page_count]
                    .iter()
                    .map(|page| PageFrame::clone(page))
                    .collect()
            });
            pages.extend(recorded.pages);
            let mut outlines = restart_checkpoint.map_or_else(Vec::new, |checkpoint| {
                self.restart_cache
                    .as_ref()
                    .expect("a restart checkpoint belongs to retained outlines")
                    .outlines
                    .iter()
                    .filter(|outline| outline.page_index < checkpoint.page_count)
                    .cloned()
                    .collect()
            });
            outlines.extend(recorded.outlines);
            let mut checkpoints = restart_checkpoint.map_or_else(Vec::new, |checkpoint| {
                self.restart_cache
                    .as_ref()
                    .expect("a restart checkpoint belongs to retained state")
                    .checkpoints
                    .iter()
                    .copied()
                    .filter(|candidate| candidate.page_count <= checkpoint.page_count)
                    .collect()
            });
            checkpoints.extend(recorded.checkpoints);
            if let (Some(stopped), Some((_, old_tail))) = (recorded.stopped_at, tail_source) {
                debug_assert_eq!(stopped.page_count, old_tail.page_count);
                let cache = self.restart_cache.as_ref().expect("restart cache exists");
                pages.extend(
                    cache.raw_pages[old_tail.page_count..]
                        .iter()
                        .map(|page| PageFrame::clone(page)),
                );
                outlines.extend(
                    cache
                        .outlines
                        .iter()
                        .filter(|outline| outline.page_index >= old_tail.page_count)
                        .cloned(),
                );
                let block_delta =
                    stopped.next_block_index as isize - old_tail.next_block_index as isize;
                checkpoints.extend(
                    cache
                        .checkpoints
                        .iter()
                        .filter(|candidate| candidate.page_count > old_tail.page_count)
                        .map(|candidate| paginator::PaginationCheckpoint {
                            next_block_index: candidate
                                .next_block_index
                                .checked_add_signed(block_delta)
                                .expect("common suffix block index remains in range"),
                            page_count: candidate.page_count,
                            next_header_page_number: candidate.next_header_page_number,
                        }),
                );
            }
            checkpoints.sort_unstable_by_key(|checkpoint| checkpoint.next_block_index);
            checkpoints.dedup();
            (pages, outlines, checkpoints)
        } else {
            let (pages, outlines) =
                paginator::paginate_sections(&sections, &self.font_manager, &media, &notes);
            (pages, outlines, Vec::new())
        };

        // Endnotes read at the end of the document, so they follow the last
        // body page rather than sitting at the foot of their reference's page.
        paginator::append_endnote_pages(&mut pages, &notes, final_geometry);

        apply_page_background(&mut pages, input);

        let mut pages = pages.into_iter().map(Arc::new).collect::<Vec<_>>();
        if reusable_restart_record && let Some(cache) = self.restart_cache.as_ref() {
            for (page, retained) in pages.iter_mut().zip(&cache.raw_pages) {
                if page_frames_equal(page, retained) {
                    *page = Arc::clone(retained);
                }
            }
        }
        let mut raw_pages = restart_record_eligible.then(|| pages.clone());

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
        let mut bookmark_identity = bookmark_pages
            .iter()
            .map(|(&target, &page_number)| (target, page_number))
            .collect::<Vec<_>>();
        bookmark_identity.sort_unstable();
        let mut substitution_inputs = Vec::with_capacity(pages.len());
        let mut reuse_result_pages = vec![false; pages.len()];
        for (page_index, page) in pages.iter_mut().enumerate() {
            if !page_has_substitution_state(page) {
                reuse_result_pages[page_index] = self.restart_cache.as_ref().is_some_and(|cache| {
                    cache.substitution_inputs.get(page_index) == Some(&None)
                        && cache
                            .raw_pages
                            .get(page_index)
                            .is_some_and(|retained| Arc::ptr_eq(page, retained))
                });
                substitution_inputs.push(None);
                continue;
            }
            let inputs = FieldSubstitutionInputs {
                page_index,
                page_number: page.page_number,
                total_pages,
                bookmark_pages: bookmark_identity.clone(),
                font_identity: font_trace.clone(),
                revision_view: input.revision_view,
            };
            let reusable = self.restart_cache.as_ref().is_some_and(|cache| {
                cache
                    .substitution_inputs
                    .get(page_index)
                    .and_then(Option::as_ref)
                    == Some(&inputs)
                    && cache
                        .raw_pages
                        .get(page_index)
                        .is_some_and(|retained| Arc::ptr_eq(page, retained))
            });
            if reusable {
                reuse_result_pages[page_index] = true;
                substitution_inputs.push(Some(inputs));
                continue;
            }
            let page = Arc::make_mut(page);
            let page_num = page.page_number;
            substitute_fields(
                &mut page.elements,
                page_num,
                total_pages,
                &bookmark_pages,
                &mut self.font_manager,
            );
            substitution_inputs.push(Some(inputs));
        }

        #[cfg(test)]
        {
            self.last_rebuilt_page_range = Some(0..pages.len());
        }
        if restart_checkpoint.is_some()
            && let Some(cache) = self.restart_cache.as_ref()
            && cache.font_trace == font_trace
        {
            let mut rebuilt_start = pages.len();
            let mut rebuilt_end = 0;
            for (page_index, (page, retained)) in pages.iter().zip(&cache.raw_pages).enumerate() {
                if page_frames_equal(page, retained) {
                    reuse_result_pages[page_index] = true;
                } else {
                    rebuilt_start = rebuilt_start.min(page_index);
                    rebuilt_end = page_index + 1;
                }
            }
            if pages.len() != cache.pages.len() {
                rebuilt_start = rebuilt_start.min(pages.len().min(cache.pages.len()));
                rebuilt_end = pages.len();
            }
            let rebuilt_range = if rebuilt_start < rebuilt_end {
                rebuilt_start..rebuilt_end
            } else {
                0..0
            };
            #[cfg(test)]
            {
                self.last_rebuilt_page_range = Some(rebuilt_range);
            }
            #[cfg(not(test))]
            {
                let _ = rebuilt_range;
            }
        }
        // Metrics-only empty carriers must still resolve through the result,
        // but they do not get to move a glyph-bearing font earlier in the
        // deterministic result order.
        let mut carrier_fonts = Vec::new();
        fn collect_carrier_fonts(elements: &[PositionedElement], fonts: &mut Vec<FontId>) {
            for element in elements {
                match element {
                    PositionedElement::Text(run)
                        if run.text.is_empty() && run.glyph_ids.is_empty() =>
                    {
                        if !fonts.contains(&run.font_id) {
                            fonts.push(run.font_id);
                        }
                    }
                    PositionedElement::Group(group) => {
                        collect_carrier_fonts(&group.children, fonts)
                    }
                    _ => {}
                }
            }
        }
        for page in &pages {
            collect_carrier_fonts(&page.elements, &mut carrier_fonts);
        }
        self.font_manager.replay_layout_font_trace(&carrier_fonts);

        // Remap persistent manager ids to result-local ids and omit faces that
        // are no longer present in the current layout.
        let fonts = if self.font_manager.every_loaded_font_is_current() {
            self.font_manager.all_font_data()
        } else {
            let current_fonts = self.font_manager.current_layout_fonts().to_vec();
            canonicalize_layout_fonts(&mut pages, &self.font_manager, &current_fonts)?
        };
        if let Some(cache) = self.restart_cache.as_ref() {
            for (page_index, reuse) in reuse_result_pages.into_iter().enumerate() {
                if reuse && let Some(retained) = cache.pages.get(page_index) {
                    pages[page_index] = Arc::clone(retained);
                }
            }
        }
        let mut retained_pages = restart_record_eligible.then(|| pages.clone());

        // Convert core properties to document metadata
        let metadata = input.core_properties.as_ref().map(|cp| DocumentMetadata {
            title: cp.title.clone(),
            author: cp.creator.clone(),
            subject: cp.subject.clone(),
            keywords: cp.keywords.clone(),
            creator: Some("rdocx".to_string()),
        });

        if restart_record_eligible
            && pages.len().max(checkpoints.len()) <= RESTART_CACHE_MAX_ENTRIES
            && let Some(raw_pages) = raw_pages.as_mut()
            && let Some(retained_pages) = retained_pages.as_mut()
        {
            body.shrink_to_fit();
            raw_pages.shrink_to_fit();
            retained_pages.shrink_to_fit();
            substitution_inputs.shrink_to_fit();
            outlines.shrink_to_fit();
            checkpoints.shrink_to_fit();
            font_trace.shrink_to_fit();
            let mut candidate = RestartCache {
                body,
                with_provenance: sources.is_some(),
                raw_pages: std::mem::take(raw_pages),
                pages: std::mem::take(retained_pages),
                substitution_inputs,
                outlines: outlines.clone(),
                checkpoints,
                font_trace,
                bytes: 0,
            };
            candidate.outlines.shrink_to_fit();
            for inputs in candidate.substitution_inputs.iter_mut().flatten() {
                inputs.bookmark_pages.shrink_to_fit();
                inputs.font_identity.shrink_to_fit();
            }
            let bytes = restart_cache_bytes(&candidate);
            if bytes <= RESTART_CACHE_MAX_BYTES {
                candidate.bytes = bytes;
                self.restart_cache = Some(candidate);
            } else {
                self.restart_cache = None;
            }
        } else {
            self.restart_cache = None;
        }
        let mut result = LayoutResult::new(pages, fonts, metadata, outlines);
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
        if !paragraph_is_cache_safe(paragraph, styles) {
            // Traversal-sensitive content can change generated state consumed
            // by later blocks. The conservative boundary is the first such
            // block, after which no retained block is read in this layout.
            self.paragraph_cache_reads_enabled = false;
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

        let fingerprint = paragraph_fingerprint(paragraph);
        if self.paragraph_cache_reads_enabled
            && let Some(entry) = self.paragraph_cache.iter().find(|entry| {
                entry.fingerprint == fingerprint
                    && entry.key.paragraph == *paragraph
                    && entry.key.content_width_bits == content_width.to_bits()
                    && entry.key.revision_view == input.revision_view
            })
        {
            let mut block = entry.block.clone();
            rebind_paragraph_source(&mut block, source_node);
            diagnostics.extend(entry.diagnostics.iter().cloned());
            self.font_manager
                .replay_layout_font_trace(&entry.font_trace);
            self.paragraph_cache_hits += 1;
            return Ok(block);
        }

        let diagnostics_start = diagnostics.len();
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
        let mut block = block_result?;
        self.paragraph_cache_builds += 1;

        let cached_diagnostics = diagnostics[diagnostics_start..].to_vec();
        if let Some(font_trace) = font_trace {
            let bytes = paragraph_cache_entry_bytes(
                paragraph,
                &block,
                &cached_diagnostics,
                font_trace.len(),
            );
            self.stage_paragraph_cache_entry(ParagraphCacheEntry {
                fingerprint,
                key: ParagraphCacheKey {
                    paragraph: paragraph.clone(),
                    content_width_bits: content_width.to_bits(),
                    revision_view: input.revision_view,
                },
                block: block.clone(),
                diagnostics: cached_diagnostics,
                font_trace,
                bytes,
            });
        }

        rebind_paragraph_source(&mut block, source_node);
        Ok(block)
    }

    #[allow(clippy::too_many_arguments)]
    fn layout_body_table(
        &mut self,
        table: &CT_Tbl,
        content_width: f64,
        styles: &CT_Styles,
        input: &LayoutInput,
        media: &MediaRegistry,
        numbering: &mut NumberingState,
        diagnostics: &mut Vec<Diagnostic>,
        sources: Option<&SourceRegistry>,
        story: &WordStory,
        path: &[usize],
    ) -> Result<table::TableBlock> {
        if !table_is_cache_safe(table, styles) {
            self.paragraph_cache_reads_enabled = false;
            return table::layout_table_with_provenance(
                table,
                content_width,
                styles,
                input,
                media,
                &mut self.font_manager,
                numbering,
                diagnostics,
                sources,
                story,
                path,
            );
        }

        let key = TableCacheKey {
            table: table.clone(),
            content_width_bits: content_width.to_bits(),
            revision_view: input.revision_view,
            with_provenance: sources.is_some(),
        };
        if self.paragraph_cache_reads_enabled
            && let Some(index) = self.table_cache.iter().position(|entry| entry.key == key)
        {
            let entry = self
                .table_cache
                .remove(index)
                .expect("table cache index exists");
            let mut block = entry.block.clone();
            rebind_table_sources(table, &mut block, sources, story, path);
            diagnostics.extend(entry.diagnostics.iter().cloned());
            self.font_manager
                .replay_layout_font_trace(&entry.font_trace);
            self.table_cache.push_back(entry);
            self.table_cache_hits += 1;
            return Ok(block);
        }

        let diagnostics_start = diagnostics.len();
        self.font_manager.begin_paragraph_font_trace();
        let block_result = table::layout_table_with_provenance(
            table,
            content_width,
            styles,
            input,
            media,
            &mut self.font_manager,
            numbering,
            diagnostics,
            sources,
            story,
            path,
        );
        let font_trace = self.font_manager.finish_paragraph_font_trace();
        let block = block_result?;
        self.table_cache_builds += 1;

        let cached_diagnostics = diagnostics[diagnostics_start..].to_vec();
        if let Some(font_trace) = font_trace {
            let bytes =
                table_cache_entry_bytes(&key, &block, &cached_diagnostics, font_trace.len());
            self.stage_table_cache_entry(TableCacheEntry {
                key,
                block: block.clone(),
                diagnostics: cached_diagnostics,
                font_trace,
                bytes,
            });
        }
        Ok(block)
    }

    #[cfg(test)]
    fn paragraph_cache_counts(&self) -> (usize, usize) {
        (self.paragraph_cache_hits, self.paragraph_cache_builds)
    }

    #[cfg(test)]
    fn table_cache_counts(&self) -> (usize, usize) {
        (self.table_cache_hits, self.table_cache_builds)
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

    fn publish_table_cache_entry(&mut self, entry: TableCacheEntry) {
        if entry.bytes > TABLE_CACHE_MAX_BYTES {
            return;
        }
        while self.table_cache.len() >= TABLE_CACHE_MAX_ENTRIES
            || self.table_cache_bytes.saturating_add(entry.bytes) > TABLE_CACHE_MAX_BYTES
        {
            let Some(evicted) = self.table_cache.pop_front() else {
                break;
            };
            self.table_cache_bytes = self.table_cache_bytes.saturating_sub(evicted.bytes);
        }
        self.table_cache_bytes += entry.bytes;
        self.table_cache.push_back(entry);
        let restart_entries = self.restart_cache.as_ref().map_or(0, restart_cache_entries);
        let restart_bytes = self.restart_cache.as_ref().map_or(0, |cache| cache.bytes);
        debug_assert!(
            self.paragraph_cache.len() + self.table_cache.len() + restart_entries
                <= CACHE_MAX_ENTRIES
        );
        debug_assert!(
            self.paragraph_cache_bytes + self.table_cache_bytes + restart_bytes <= CACHE_MAX_BYTES
        );
    }

    fn stage_table_cache_entry(&mut self, entry: TableCacheEntry) {
        if entry.bytes > TABLE_CACHE_MAX_BYTES {
            return;
        }
        let Some(pending) = self.pending_table_cache.as_mut() else {
            return;
        };
        while pending.len() >= TABLE_CACHE_MAX_ENTRIES
            || self.pending_table_cache_bytes.saturating_add(entry.bytes) > TABLE_CACHE_MAX_BYTES
        {
            let Some(evicted) = pending.pop_front() else {
                break;
            };
            self.pending_table_cache_bytes =
                self.pending_table_cache_bytes.saturating_sub(evicted.bytes);
        }
        self.pending_table_cache_bytes += entry.bytes;
        pending.push_back(entry);
        #[cfg(test)]
        {
            self.pending_table_cache_peak_entries =
                self.pending_table_cache_peak_entries.max(pending.len());
            self.pending_table_cache_peak_bytes = self
                .pending_table_cache_peak_bytes
                .max(self.pending_table_cache_bytes);
        }
    }

    #[cfg(test)]
    fn header_footer_cache_counts(&self) -> (usize, usize) {
        (
            self.header_footer_cache_hits,
            self.header_footer_cache_builds,
        )
    }

    fn publish_header_footer_cache_entry(&mut self, entry: HeaderFooterCacheEntry) {
        if entry.bytes > HEADER_FOOTER_CACHE_MAX_BYTES {
            return;
        }
        while self.header_footer_cache.len() >= HEADER_FOOTER_CACHE_MAX_ENTRIES
            || self.header_footer_cache_bytes.saturating_add(entry.bytes)
                > HEADER_FOOTER_CACHE_MAX_BYTES
        {
            let Some(evicted) = self.header_footer_cache.pop_front() else {
                break;
            };
            self.header_footer_cache_bytes =
                self.header_footer_cache_bytes.saturating_sub(evicted.bytes);
        }
        self.header_footer_cache_bytes += entry.bytes;
        self.header_footer_cache.push_back(entry);
        let restart_entries = self.restart_cache.as_ref().map_or(0, restart_cache_entries);
        let restart_bytes = self.restart_cache.as_ref().map_or(0, |cache| cache.bytes);
        debug_assert!(
            self.paragraph_cache.len()
                + self.table_cache.len()
                + self.header_footer_cache.len()
                + restart_entries
                <= CACHE_MAX_ENTRIES
        );
        debug_assert!(
            self.paragraph_cache_bytes
                + self.table_cache_bytes
                + self.header_footer_cache_bytes
                + restart_bytes
                <= CACHE_MAX_BYTES
        );
    }

    fn stage_header_footer_cache_entry(&mut self, entry: HeaderFooterCacheEntry) {
        if entry.bytes > HEADER_FOOTER_CACHE_MAX_BYTES {
            return;
        }
        let Some(pending) = self.pending_header_footer_cache.as_mut() else {
            return;
        };
        while pending.len() >= HEADER_FOOTER_CACHE_MAX_ENTRIES
            || self
                .pending_header_footer_cache_bytes
                .saturating_add(entry.bytes)
                > HEADER_FOOTER_CACHE_MAX_BYTES
        {
            let Some(evicted) = pending.pop_front() else {
                break;
            };
            self.pending_header_footer_cache_bytes = self
                .pending_header_footer_cache_bytes
                .saturating_sub(evicted.bytes);
        }
        self.pending_header_footer_cache_bytes += entry.bytes;
        pending.push_back(entry);
        #[cfg(test)]
        {
            self.pending_header_footer_cache_peak_entries = self
                .pending_header_footer_cache_peak_entries
                .max(pending.len());
            self.pending_header_footer_cache_peak_bytes = self
                .pending_header_footer_cache_peak_bytes
                .max(self.pending_header_footer_cache_bytes);
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

fn header_footer_part_is_cache_safe(
    part: &rdocx_oxml::header_footer::CT_HdrFtr,
    styles: &CT_Styles,
) -> bool {
    if !part.extra_xml.is_empty() {
        return false;
    }
    let raw_watermark_count = part
        .paragraphs
        .iter()
        .flat_map(|paragraph| &paragraph.runs)
        .flat_map(|run| &run.extra_xml)
        .filter(|raw| raw_xml_root_is_word_pict(raw, &part.extra_namespaces))
        .count();
    if raw_watermark_count != part.watermarks().len() {
        return false;
    }
    part.paragraphs.iter().all(|paragraph| {
        if !paragraph.extra_xml.is_empty() {
            return false;
        }
        let mut projected = paragraph.clone();
        for run in &mut projected.runs {
            if run.extra_xml.len() != run.extra_xml_positions.len()
                || !run
                    .extra_xml
                    .iter()
                    .all(|raw| raw_xml_root_is_word_pict(raw, &part.extra_namespaces))
            {
                return false;
            }
            run.extra_xml.clear();
            run.extra_xml_positions.clear();
        }
        paragraph_is_cache_safe(&projected, styles)
    })
}

fn raw_xml_root_is_word_pict(raw: &[u8], namespaces: &[(String, String)]) -> bool {
    let raw = raw
        .iter()
        .position(|byte| !byte.is_ascii_whitespace())
        .map_or(raw, |start| &raw[start..]);
    let Some((name, attributes)) = raw.strip_prefix(b"<").and_then(|raw| {
        let name_end = raw
            .iter()
            .position(|byte| byte.is_ascii_whitespace() || matches!(byte, b'>' | b'/'))?;
        Some((&raw[..name_end], &raw[name_end..]))
    }) else {
        return false;
    };
    let mut components = name.rsplitn(2, |byte| *byte == b':');
    if components.next() != Some(b"pict".as_slice()) {
        return false;
    }
    let prefix = components.next();
    let declaration = prefix.map_or_else(
        || "xmlns".to_owned(),
        |prefix| format!("xmlns:{}", String::from_utf8_lossy(prefix)),
    );
    if let Some(namespace) = raw_xml_start_attribute(attributes, declaration.as_bytes()) {
        return namespace == rdocx_oxml::namespace::W_NS.as_bytes();
    }
    let Some(prefix) = prefix.and_then(|prefix| std::str::from_utf8(prefix).ok()) else {
        return false;
    };
    if prefix == "w" {
        return !namespaces.iter().any(|(name, namespace)| {
            name != "xmlns:w" && namespace == rdocx_oxml::namespace::W_NS
        });
    }
    namespaces.iter().any(|(name, namespace)| {
        name.strip_prefix("xmlns:") == Some(prefix) && namespace == rdocx_oxml::namespace::W_NS
    })
}

fn raw_xml_start_attribute<'a>(mut input: &'a [u8], expected: &[u8]) -> Option<&'a [u8]> {
    loop {
        input = input
            .iter()
            .position(|byte| !byte.is_ascii_whitespace())
            .map_or(input, |start| &input[start..]);
        if input.first().is_none_or(|byte| matches!(byte, b'>' | b'/')) {
            return None;
        }
        let name_end = input
            .iter()
            .position(|byte| byte.is_ascii_whitespace() || matches!(byte, b'=' | b'>' | b'/'))?;
        let name = &input[..name_end];
        input = &input[name_end..];
        input = input
            .iter()
            .position(|byte| !byte.is_ascii_whitespace())
            .map_or(input, |start| &input[start..]);
        if input.first() != Some(&b'=') {
            return None;
        }
        input = &input[1..];
        input = input
            .iter()
            .position(|byte| !byte.is_ascii_whitespace())
            .map_or(input, |start| &input[start..]);
        let quote = *input.first()?;
        if !matches!(quote, b'\'' | b'"') {
            return None;
        }
        input = &input[1..];
        let value_end = input.iter().position(|byte| *byte == quote)?;
        let value = &input[..value_end];
        input = &input[value_end + 1..];
        if name == expected {
            return Some(value);
        }
    }
}

fn header_footer_section_is_cache_safe(section: &CT_SectPr) -> bool {
    section.change.is_none() && section.extra_xml.is_empty()
}

fn table_is_cache_safe(table: &CT_Tbl, styles: &CT_Styles) -> bool {
    table.extra_xml.is_empty()
        && table.content_controls.is_empty()
        && table.properties.as_ref().is_none_or(|properties| {
            properties.change.is_none() && properties.revision_xml.is_empty()
        })
        && table.rows.iter().all(|row| {
            row.extra_xml.is_empty()
                && row.content_controls.is_empty()
                && row.properties.as_ref().is_none_or(|properties| {
                    properties.revision_markers.is_empty() && properties.revision_xml.is_empty()
                })
                && row.cells.iter().all(|cell| {
                    cell.extra_xml.is_empty()
                        && cell
                            .properties
                            .as_ref()
                            .is_none_or(|properties| properties.extra_xml.is_empty())
                        && cell.content.iter().all(|content| match content {
                            CellContent::Paragraph(paragraph) => {
                                paragraph_is_cache_safe(paragraph, styles)
                            }
                            CellContent::Table(_) | CellContent::ContentControl(_) => false,
                        })
                })
        })
}

fn rebind_table_sources(
    table: &CT_Tbl,
    block: &mut table::TableBlock,
    sources: Option<&SourceRegistry>,
    story: &WordStory,
    table_path: &[usize],
) {
    for (row_index, (row, block_row)) in table.rows.iter().zip(&mut block.rows).enumerate() {
        for (cell_index, (cell, block_cell)) in
            row.cells.iter().zip(&mut block_row.cells).enumerate()
        {
            let mut block_paragraphs = block_cell.paragraphs.iter_mut();
            for (content_index, content) in cell.content.iter().enumerate() {
                let CellContent::Paragraph(_) = content else {
                    continue;
                };
                let Some(paragraph) = block_paragraphs.next() else {
                    break;
                };
                let mut source_path = table_path.to_vec();
                source_path.extend([row_index, cell_index, content_index]);
                let source = sources.and_then(|sources| sources.id(story, &source_path));
                rebind_paragraph_source(paragraph, source);
            }
        }
    }
}

fn table_cache_entry_bytes(
    key: &TableCacheKey,
    block: &table::TableBlock,
    diagnostics: &[Diagnostic],
    font_trace_len: usize,
) -> usize {
    let diagnostic_bytes = diagnostics
        .len()
        .saturating_mul(std::mem::size_of::<Diagnostic>())
        .saturating_add(
            diagnostics
                .iter()
                .map(|diagnostic| diagnostic.message.capacity())
                .fold(0usize, usize::saturating_add),
        );
    std::mem::size_of::<TableCacheEntry>()
        .saturating_add(table_key_retained_bytes(&key.table))
        .saturating_add(table_block_retained_bytes(block))
        .saturating_add(font_trace_len.saturating_mul(std::mem::size_of::<FontId>()))
        .saturating_add(diagnostic_bytes)
}

fn table_key_retained_bytes(table: &CT_Tbl) -> usize {
    let grid_bytes = table.grid.as_ref().map_or(0, |grid| {
        grid.columns
            .capacity()
            .saturating_mul(std::mem::size_of::<rdocx_oxml::table::CT_TblGridCol>())
    });
    table
        .rows
        .capacity()
        .saturating_mul(std::mem::size_of::<CT_Row>())
        .saturating_add(grid_bytes)
        .saturating_add(format!("{table:?}").len())
        .saturating_add(
            table
                .rows
                .iter()
                .map(|row| {
                    row.cells
                        .capacity()
                        .saturating_mul(std::mem::size_of::<CT_Tc>())
                        .saturating_add(
                            row.cells
                                .iter()
                                .map(|cell| {
                                    cell.content
                                        .capacity()
                                        .saturating_mul(std::mem::size_of::<CellContent>())
                                        .saturating_add(
                                            cell.content
                                                .iter()
                                                .map(|content| match content {
                                                    CellContent::Paragraph(paragraph) => paragraph
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
                                                                        .saturating_mul(
                                                                            std::mem::size_of::<
                                                                                RunContent,
                                                                            >(
                                                                            ),
                                                                        )
                                                                })
                                                                .fold(
                                                                    0usize,
                                                                    usize::saturating_add,
                                                                ),
                                                        ),
                                                    CellContent::Table(_)
                                                    | CellContent::ContentControl(_) => usize::MAX,
                                                })
                                                .fold(0usize, usize::saturating_add),
                                        )
                                })
                                .fold(0usize, usize::saturating_add),
                        )
                })
                .fold(0usize, usize::saturating_add),
        )
}

fn table_block_retained_bytes(block: &table::TableBlock) -> usize {
    fn border_bytes(borders: &rdocx_oxml::table::CT_TblBorders) -> usize {
        [
            &borders.top,
            &borders.bottom,
            &borders.left,
            &borders.right,
            &borders.inside_h,
            &borders.inside_v,
        ]
        .into_iter()
        .map(|edge| {
            edge.as_ref()
                .and_then(|edge| edge.color.as_ref())
                .map_or(0, String::capacity)
        })
        .fold(0usize, usize::saturating_add)
    }

    let rows = block
        .rows
        .iter()
        .map(|row| {
            row.cells
                .capacity()
                .saturating_mul(std::mem::size_of::<table::TableCell>())
                .saturating_add(
                    row.cells
                        .iter()
                        .map(|cell| {
                            cell.paragraphs
                                .capacity()
                                .saturating_mul(std::mem::size_of::<ParagraphBlock>())
                                .saturating_add(
                                    cell.paragraphs
                                        .iter()
                                        .map(|paragraph| {
                                            paragraph_cache_entry_bytes(
                                                &CT_P::new(),
                                                paragraph,
                                                &[],
                                                0,
                                            )
                                        })
                                        .fold(0usize, usize::saturating_add),
                                )
                                .saturating_add(cell.borders.as_ref().map_or(0, border_bytes))
                        })
                        .fold(0usize, usize::saturating_add),
                )
        })
        .fold(0usize, usize::saturating_add);
    std::mem::size_of::<table::TableBlock>()
        .saturating_add(
            block
                .col_widths
                .capacity()
                .saturating_mul(std::mem::size_of::<f64>()),
        )
        .saturating_add(
            block
                .rows
                .capacity()
                .saturating_mul(std::mem::size_of::<table::TableRow>()),
        )
        .saturating_add(
            block
                .header_row_indices
                .capacity()
                .saturating_mul(std::mem::size_of::<usize>()),
        )
        .saturating_add(rows)
        .saturating_add(block.borders.as_ref().map_or(0, border_bytes))
}

fn canonicalize_layout_fonts(
    pages: &mut [Arc<PageFrame>],
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
    for page in pages {
        rewrite(&mut Arc::make_mut(page).elements, &remap);
    }
    Ok(fonts)
}

fn restart_block_is_safe(block: &LayoutBlock) -> bool {
    let LayoutBlock::Paragraph(paragraph) = block else {
        return false;
    };
    restart_record_block_is_safe(block)
        && paragraph.lines.iter().all(|line| {
            line.items.iter().all(|item| match item {
                LineItem::Text(text) | LineItem::Marker(text) => text.field_kind.is_none(),
                LineItem::Tab {
                    leader: Some(text), ..
                } => text.field_kind.is_none(),
                _ => true,
            })
        })
}

fn restart_record_block_is_safe(block: &LayoutBlock) -> bool {
    let LayoutBlock::Paragraph(paragraph) = block else {
        return false;
    };
    paragraph.anchored.is_empty()
        && paragraph.lines.len() <= 2
        && paragraph.heading_level.is_none()
        && !paragraph.keep_next
        && !paragraph.keep_lines
        && paragraph.lines.iter().all(|line| {
            line.items.iter().all(|item| match item {
                LineItem::Text(text) | LineItem::Marker(text) => text.note.is_none(),
                LineItem::Tab {
                    leader: Some(text), ..
                } => text.note.is_none(),
                LineItem::Tab { leader: None, .. } => true,
                LineItem::Image { .. } | LineItem::Group { .. } => false,
                _ => false,
            })
        })
}

fn page_has_substitution_state(page: &PageFrame) -> bool {
    page.elements.iter().any(|element| {
        matches!(
            element,
            PositionedElement::Text(run) if run.field_kind.is_some()
        )
    })
}

fn restart_cache_entries(cache: &RestartCache) -> usize {
    cache.raw_pages.len().max(cache.checkpoints.len())
}

fn page_frames_equal(left: &PageFrame, right: &PageFrame) -> bool {
    left.page_number == right.page_number
        && left.width == right.width
        && left.height == right.height
        && left.elements == right.elements
        && left.background == right.background
}

fn restart_cache_bytes(cache: &RestartCache) -> usize {
    let vector_bytes = cache
        .body
        .capacity()
        .saturating_mul(std::mem::size_of::<String>())
        .saturating_add(
            cache
                .raw_pages
                .capacity()
                .saturating_add(cache.pages.capacity())
                .saturating_mul(std::mem::size_of::<Arc<PageFrame>>()),
        )
        .saturating_add(
            cache
                .substitution_inputs
                .capacity()
                .saturating_mul(std::mem::size_of::<Option<FieldSubstitutionInputs>>()),
        )
        .saturating_add(
            cache
                .outlines
                .capacity()
                .saturating_mul(std::mem::size_of::<oxml_layout::OutlineEntry>()),
        )
        .saturating_add(
            cache
                .checkpoints
                .capacity()
                .saturating_mul(std::mem::size_of::<paginator::PaginationCheckpoint>()),
        )
        .saturating_add(
            cache
                .font_trace
                .capacity()
                .saturating_mul(std::mem::size_of::<FontId>()),
        );
    let body_bytes = cache
        .body
        .iter()
        .map(String::capacity)
        .fold(0usize, usize::saturating_add);
    debug_assert_eq!(cache.raw_pages.len(), cache.pages.len());
    let page_bytes = cache
        .raw_pages
        .iter()
        .zip(&cache.pages)
        .map(|(pristine, substituted)| {
            let substituted_bytes = if Arc::ptr_eq(pristine, substituted) {
                0
            } else {
                page_frame_retained_bytes(substituted)
            };
            page_frame_retained_bytes(pristine).saturating_add(substituted_bytes)
        })
        .fold(0usize, usize::saturating_add);
    let outline_bytes = cache
        .outlines
        .iter()
        .map(|outline| outline.title.capacity())
        .fold(0usize, usize::saturating_add);
    let substitution_bytes = cache
        .substitution_inputs
        .iter()
        .filter_map(Option::as_ref)
        .map(|inputs| {
            inputs
                .bookmark_pages
                .capacity()
                .saturating_mul(std::mem::size_of::<(usize, usize)>())
                .saturating_add(
                    inputs
                        .font_identity
                        .capacity()
                        .saturating_mul(std::mem::size_of::<FontId>()),
                )
        })
        .fold(0usize, usize::saturating_add);
    std::mem::size_of::<RestartCache>()
        .saturating_add(vector_bytes)
        .saturating_add(body_bytes)
        .saturating_add(page_bytes)
        .saturating_add(substitution_bytes)
        .saturating_add(outline_bytes)
}

fn page_frame_retained_bytes(page: &PageFrame) -> usize {
    fn glyph_bytes(run: &GlyphRun) -> usize {
        run.text
            .capacity()
            .saturating_add(
                run.glyph_ids
                    .capacity()
                    .saturating_mul(std::mem::size_of::<u16>()),
            )
            .saturating_add(
                run.advances
                    .capacity()
                    .saturating_mul(std::mem::size_of::<f64>()),
            )
    }

    fn element_bytes(element: &PositionedElement) -> usize {
        match element {
            PositionedElement::Text(run) => glyph_bytes(run),
            PositionedElement::Image {
                data, content_type, ..
            } => data.capacity().saturating_add(content_type.capacity()),
            PositionedElement::LinkAnnotation { url, .. } => url.capacity(),
            PositionedElement::Group(group) => group
                .children
                .capacity()
                .saturating_mul(std::mem::size_of::<PositionedElement>())
                .saturating_add(
                    group
                        .children
                        .iter()
                        .map(element_bytes)
                        .fold(0usize, usize::saturating_add),
                )
                .saturating_add(format!("{:?}", group.effects).len()),
            PositionedElement::Path(path) => format!("{path:?}").len(),
            _ => 0,
        }
    }

    std::mem::size_of::<PageFrame>()
        .saturating_add(
            page.elements
                .capacity()
                .saturating_mul(std::mem::size_of::<PositionedElement>()),
        )
        .saturating_add(
            page.elements
                .iter()
                .map(element_bytes)
                .fold(0usize, usize::saturating_add),
        )
        .saturating_add(format!("{:?}", page.background).len())
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

fn rebind_header_footer_sources(
    story_kind: HeaderFooterStoryKind,
    relationship_id: &str,
    part: &rdocx_oxml::header_footer::CT_HdrFtr,
    blocks: &mut [ParagraphBlock],
    sources: Option<&SourceRegistry>,
) {
    let story = match story_kind {
        HeaderFooterStoryKind::Header => WordStory::Header {
            relationship_id: relationship_id.to_owned(),
        },
        HeaderFooterStoryKind::Footer => WordStory::Footer {
            relationship_id: relationship_id.to_owned(),
        },
    };
    for (paragraph_index, block) in blocks.iter_mut().enumerate() {
        debug_assert!(paragraph_index < part.paragraphs.len());
        let source = sources.and_then(|sources| sources.id(&story, &[paragraph_index]));
        rebind_paragraph_source(block, source);
    }
}

fn header_footer_cache_entry_bytes(
    key: &HeaderFooterCacheKey,
    content: &HeaderFooterVariantContent,
    diagnostics: &Vec<Diagnostic>,
    font_trace: &Vec<FontId>,
) -> usize {
    let block_bytes = key
        .part
        .paragraphs
        .iter()
        .zip(&content.blocks)
        .map(|(paragraph, block)| paragraph_cache_entry_bytes(paragraph, block, &[], 0))
        .fold(0usize, usize::saturating_add);
    let diagnostic_bytes = diagnostics
        .capacity()
        .saturating_mul(std::mem::size_of::<Diagnostic>())
        .saturating_add(
            diagnostics
                .iter()
                .map(|diagnostic| diagnostic.message.capacity())
                .fold(0usize, usize::saturating_add),
        );
    let watermark_bytes = content.watermark.as_ref().map_or(0, |watermark| {
        page_frame_retained_bytes(&PageFrame::new(
            1,
            0.0,
            0.0,
            vec![PositionedElement::Group(watermark.clone())],
        ))
    });
    let section_capacity = key
        .section
        .header_refs
        .capacity()
        .saturating_mul(std::mem::size_of::<rdocx_oxml::header_footer::HdrFtrRef>())
        .saturating_add(
            key.section
                .footer_refs
                .capacity()
                .saturating_mul(std::mem::size_of::<rdocx_oxml::header_footer::HdrFtrRef>()),
        )
        .saturating_add(key.section.columns.as_ref().map_or(0, |columns| {
            columns
                .columns
                .capacity()
                .saturating_mul(std::mem::size_of::<rdocx_oxml::document::CT_Column>())
        }))
        .saturating_add(
            key.section
                .extra_xml
                .capacity()
                .saturating_mul(std::mem::size_of::<Vec<u8>>()),
        )
        .saturating_add(
            key.section
                .header_refs
                .iter()
                .chain(&key.section.footer_refs)
                .map(|reference| reference.rel_id.capacity())
                .fold(0usize, usize::saturating_add),
        )
        .saturating_add(
            key.section
                .extra_xml
                .iter()
                .map(Vec::capacity)
                .fold(0usize, usize::saturating_add),
        );
    let paragraph_raw_capacity = key
        .part
        .paragraphs
        .iter()
        .map(|paragraph| {
            paragraph
                .extra_xml
                .capacity()
                .saturating_mul(std::mem::size_of::<(usize, Vec<u8>)>())
                .saturating_add(
                    paragraph
                        .extra_xml
                        .iter()
                        .map(|(_, raw)| raw.capacity())
                        .fold(0usize, usize::saturating_add),
                )
                .saturating_add(
                    paragraph
                        .runs
                        .iter()
                        .map(|run| {
                            run.extra_xml
                                .capacity()
                                .saturating_mul(std::mem::size_of::<Vec<u8>>())
                                .saturating_add(
                                    run.extra_xml
                                        .iter()
                                        .map(Vec::capacity)
                                        .fold(0usize, usize::saturating_add),
                                )
                                .saturating_add(
                                    run.extra_xml_positions
                                        .capacity()
                                        .saturating_mul(std::mem::size_of::<usize>()),
                                )
                        })
                        .fold(0usize, usize::saturating_add),
                )
        })
        .fold(0usize, usize::saturating_add);
    let watermark_capacity = key
        .part
        .watermarks()
        .iter()
        .map(|watermark| {
            std::mem::size_of::<VmlWatermark>().saturating_add(match watermark {
                VmlWatermark::Text {
                    text,
                    color,
                    font_family,
                    ..
                } => text
                    .capacity()
                    .saturating_add(color.capacity())
                    .saturating_add(font_family.as_ref().map_or(0, String::capacity)),
                VmlWatermark::Image {
                    relationship_id, ..
                } => relationship_id.capacity(),
            })
        })
        .fold(0usize, usize::saturating_add);
    let part_capacity = key
        .part
        .paragraphs
        .capacity()
        .saturating_mul(std::mem::size_of::<CT_P>())
        .saturating_add(
            content
                .blocks
                .capacity()
                .saturating_mul(std::mem::size_of::<ParagraphBlock>()),
        )
        .saturating_add(
            key.part
                .extra_namespaces
                .capacity()
                .saturating_mul(std::mem::size_of::<(String, String)>()),
        )
        .saturating_add(
            key.part
                .extra_xml
                .capacity()
                .saturating_mul(std::mem::size_of::<Vec<u8>>()),
        )
        .saturating_add(
            key.part
                .extra_namespaces
                .iter()
                .map(|(prefix, namespace)| prefix.capacity().saturating_add(namespace.capacity()))
                .fold(0usize, usize::saturating_add),
        )
        .saturating_add(
            key.part
                .extra_xml
                .iter()
                .map(Vec::capacity)
                .fold(0usize, usize::saturating_add),
        );
    std::mem::size_of::<HeaderFooterCacheEntry>()
        .saturating_add(key.relationship_id.capacity())
        .saturating_add(key.resolved_part_bytes.capacity())
        .saturating_add(format!("{:?}", key.section).len())
        .saturating_add(format!("{:?}", key.part).len())
        .saturating_add(section_capacity)
        .saturating_add(part_capacity)
        .saturating_add(paragraph_raw_capacity)
        .saturating_add(watermark_capacity)
        .saturating_add(block_bytes)
        .saturating_add(watermark_bytes)
        .saturating_add(diagnostic_bytes)
        .saturating_add(
            font_trace
                .capacity()
                .saturating_mul(std::mem::size_of::<FontId>()),
        )
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

struct StableParagraphFingerprint(u64);

impl StableParagraphFingerprint {
    fn new() -> Self {
        Self(0xcbf2_9ce4_8422_2325)
    }
}

impl fmt::Write for StableParagraphFingerprint {
    fn write_str(&mut self, value: &str) -> fmt::Result {
        for byte in value.bytes() {
            self.0 ^= u64::from(byte);
            self.0 = self.0.wrapping_mul(0x0000_0100_0000_01b3);
        }
        Ok(())
    }
}

fn paragraph_fingerprint(paragraph: &CT_P) -> u64 {
    let mut fingerprint = StableParagraphFingerprint::new();
    let written = write!(&mut fingerprint, "{paragraph:?}");
    debug_assert!(written.is_ok());
    fingerprint.0
}

/// Apply page background color from `w:background` element to all pages.
fn apply_page_background(pages: &mut [PageFrame], input: &LayoutInput) {
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
    // Resolve paragraph properties
    let para_style_id = para.properties.as_ref().and_then(|p| p.style_id.as_deref());

    let resolved_ppr = style_resolver::resolve_paragraph_properties(para_style_id, styles);

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

                    inline_items.push(InlineItem::Text(TextSegment {
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
                                inline_items.push(InlineItem::Text(TextSegment {
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

    let attributed_empty_paragraph = inline_items.is_empty();
    if attributed_empty_paragraph {
        let mut caret_rpr = style_resolver::resolve_run_properties(para_style_id, None, styles);
        if let Some(paragraph_mark_rpr) = direct_ppr.and_then(|ppr| ppr.rpr.as_ref()) {
            caret_rpr.merge_from(paragraph_mark_rpr);
        }
        let font_size = caret_rpr.sz.map(|hp| hp.to_pt()).unwrap_or(11.0);
        let bold = caret_rpr.bold.unwrap_or(false);
        let italic = caret_rpr.italic.unwrap_or(false);
        let font_family = resolve_font_family(&caret_rpr, input.theme.as_ref());
        let font_id = fm.resolve_font_for_metrics(font_family.as_deref(), bold, italic)?;
        let metrics = fm.metrics(font_id, font_size)?;
        inline_items.push(InlineItem::Text(TextSegment {
            text: String::new(),
            source: source_node.map(|node| SourceSpan {
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
            color: resolve_run_color(&caret_rpr, input.theme.as_ref()),
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
    }

    // Line breaking
    let line_params = convert::line_break_params(&effective_ppr, available_width);

    let legacy_empty_line = if attributed_empty_paragraph {
        let mut lines = break_into_lines(&[], &line_params, fm)?;
        convert::restore_word_line_heights(&mut lines, &effective_ppr);
        lines.pop()
    } else {
        None
    };

    let mut lines = break_into_lines(&inline_items, &line_params, fm)?;
    convert::restore_word_line_heights(&mut lines, &effective_ppr);
    if let (Some(line), Some(legacy)) = (lines.first_mut(), legacy_empty_line) {
        line.ascent = legacy.ascent;
        line.descent = legacy.descent;
        line.line_gap = legacy.line_gap;
        line.height = legacy.height;
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
    engine: &mut Engine,
    sect_pr: &CT_SectPr,
    input: &LayoutInput,
    styles: &CT_Styles,
    media: &MediaRegistry,
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
            let content = layout_header_footer_variant(
                engine,
                HeaderFooterStoryKind::Header,
                href.hdr_ftr_type,
                sect_pr,
                &href.rel_id,
                hdr,
                input,
                styles,
                media,
                num_state,
                diagnostics,
                sources,
                width,
                geometry,
            )?;
            target_blocks.extend(content.blocks);
            if target_watermark.is_none() {
                *target_watermark = content.watermark;
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
            let content = layout_header_footer_variant(
                engine,
                HeaderFooterStoryKind::Footer,
                fref.hdr_ftr_type,
                sect_pr,
                &fref.rel_id,
                ftr,
                input,
                styles,
                media,
                num_state,
                diagnostics,
                sources,
                width,
                geometry,
            )?;
            target_blocks.extend(content.blocks);
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

#[allow(clippy::too_many_arguments)]
fn layout_header_footer_variant(
    engine: &mut Engine,
    story_kind: HeaderFooterStoryKind,
    variant: HdrFtrType,
    sect_pr: &CT_SectPr,
    relationship_id: &str,
    part: &rdocx_oxml::header_footer::CT_HdrFtr,
    input: &LayoutInput,
    styles: &CT_Styles,
    media: &MediaRegistry,
    num_state: &mut NumberingState,
    diagnostics: &mut Vec<Diagnostic>,
    sources: Option<&SourceRegistry>,
    width: f64,
    geometry: PageGeometry,
) -> Result<HeaderFooterVariantContent> {
    let cache_safe = header_footer_section_is_cache_safe(sect_pr)
        && header_footer_part_is_cache_safe(part, styles);
    let resolved_part_bytes = match story_kind {
        HeaderFooterStoryKind::Header => part.to_xml_header(),
        HeaderFooterStoryKind::Footer => part.to_xml_footer(),
    };
    if !cache_safe || resolved_part_bytes.is_err() {
        return layout_header_footer_variant_uncached(
            story_kind,
            relationship_id,
            part,
            input,
            styles,
            media,
            &mut engine.font_manager,
            num_state,
            diagnostics,
            sources,
            false,
            width,
            geometry,
        );
    }

    let key = HeaderFooterCacheKey {
        story: story_kind,
        variant,
        section: sect_pr.clone(),
        relationship_id: relationship_id.to_owned(),
        part: part.clone(),
        resolved_part_bytes: resolved_part_bytes.expect("checked resolved part bytes"),
        with_provenance: sources.is_some(),
    };
    let hit = engine
        .header_footer_cache_reads_enabled
        .then(|| {
            engine
                .header_footer_cache
                .iter()
                .find(|entry| entry.key == key)
                .map(|entry| {
                    (
                        entry.content.clone(),
                        entry.diagnostics.clone(),
                        entry.font_trace.clone(),
                    )
                })
        })
        .flatten();
    if let Some((mut content, cached_diagnostics, font_trace)) = hit {
        rebind_header_footer_sources(
            story_kind,
            relationship_id,
            part,
            &mut content.blocks,
            sources,
        );
        diagnostics.extend(cached_diagnostics);
        engine.font_manager.replay_layout_font_trace(&font_trace);
        engine.header_footer_cache_hits += 1;
        return Ok(content);
    }

    let diagnostics_start = diagnostics.len();
    engine.font_manager.begin_paragraph_font_trace();
    let content_result = layout_header_footer_variant_uncached(
        story_kind,
        relationship_id,
        part,
        input,
        styles,
        media,
        &mut engine.font_manager,
        num_state,
        diagnostics,
        None,
        true,
        width,
        geometry,
    );
    let font_trace = engine.font_manager.finish_paragraph_font_trace();
    let mut content = content_result?;
    engine.header_footer_cache_builds += 1;
    let cached_diagnostics = diagnostics[diagnostics_start..].to_vec();
    if let Some(font_trace) = font_trace {
        let bytes =
            header_footer_cache_entry_bytes(&key, &content, &cached_diagnostics, &font_trace);
        engine.stage_header_footer_cache_entry(HeaderFooterCacheEntry {
            key,
            content: content.clone(),
            diagnostics: cached_diagnostics,
            font_trace,
            bytes,
        });
    }
    rebind_header_footer_sources(
        story_kind,
        relationship_id,
        part,
        &mut content.blocks,
        sources,
    );
    Ok(content)
}

#[allow(clippy::too_many_arguments)]
fn layout_header_footer_variant_uncached(
    story_kind: HeaderFooterStoryKind,
    relationship_id: &str,
    part: &rdocx_oxml::header_footer::CT_HdrFtr,
    input: &LayoutInput,
    styles: &CT_Styles,
    media: &MediaRegistry,
    fm: &mut FontManager,
    num_state: &mut NumberingState,
    diagnostics: &mut Vec<Diagnostic>,
    sources: Option<&SourceRegistry>,
    cache_source: bool,
    width: f64,
    geometry: PageGeometry,
) -> Result<HeaderFooterVariantContent> {
    let watermark = if story_kind == HeaderFooterStoryKind::Header {
        match part.watermarks().first() {
            Some(projected) => layout_watermark(
                projected,
                relationship_id,
                input,
                media,
                fm,
                geometry,
                diagnostics,
            )?,
            None => None,
        }
    } else {
        None
    };
    let story = match story_kind {
        HeaderFooterStoryKind::Header => WordStory::Header {
            relationship_id: relationship_id.to_owned(),
        },
        HeaderFooterStoryKind::Footer => WordStory::Footer {
            relationship_id: relationship_id.to_owned(),
        },
    };
    let mut blocks = Vec::with_capacity(part.paragraphs.len());
    for (paragraph_index, paragraph) in part.paragraphs.iter().enumerate() {
        let source = if cache_source {
            Some(CACHE_SOURCE_NODE)
        } else {
            sources.and_then(|sources| sources.id(&story, &[paragraph_index]))
        };
        blocks.push(layout_paragraph_with_source(
            paragraph,
            width,
            styles,
            input,
            media,
            fm,
            num_state,
            diagnostics,
            source,
        )?);
    }
    Ok(HeaderFooterVariantContent { blocks, watermark })
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

    fn header_footer_part(text: &str) -> rdocx_oxml::header_footer::CT_HdrFtr {
        let mut part = rdocx_oxml::header_footer::CT_HdrFtr::new();
        let mut paragraph = CT_P::new();
        paragraph.add_run(text);
        part.paragraphs.push(paragraph);
        part
    }

    fn image_watermark_header(text: &str, width_pt: f64) -> rdocx_oxml::header_footer::CT_HdrFtr {
        rdocx_oxml::header_footer::CT_HdrFtr::from_xml(
            format!(
                r#"<w:hdr xmlns:w="{}" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:p><w:r><w:pict><v:shape style="width:{width_pt}pt;height:36pt"><v:fill opacity=".5"/><v:imagedata r:id="rIdWatermark"/></v:shape></w:pict><w:t>{text}</w:t></w:r></w:p></w:hdr>"#,
                rdocx_oxml::namespace::W_NS
            )
            .as_bytes(),
        )
        .expect("watermark header parses")
    }

    fn cacheable_header_footer_input(body: &str) -> LayoutInput {
        use rdocx_oxml::header_footer::HdrFtrRef;

        let mut input = make_input_with_text(body);
        let mut section = CT_SectPr::default_letter();
        section.title_pg = Some(true);
        for (variant, suffix) in [
            (HdrFtrType::Default, "default"),
            (HdrFtrType::First, "first"),
            (HdrFtrType::Even, "even"),
        ] {
            let header_id = format!("rId-{suffix}-header");
            let footer_id = format!("rId-{suffix}-footer");
            section.header_refs.push(HdrFtrRef {
                hdr_ftr_type: variant,
                rel_id: header_id.clone(),
            });
            section.footer_refs.push(HdrFtrRef {
                hdr_ftr_type: variant,
                rel_id: footer_id.clone(),
            });
            let header = if variant == HdrFtrType::Default {
                image_watermark_header(&format!("{suffix} header"), 72.0)
            } else {
                header_footer_part(&format!("{suffix} header"))
            };
            input.headers.insert(header_id, header);
            input
                .footers
                .insert(footer_id, header_footer_part(&format!("{suffix} footer")));
        }
        input.images.insert(
            "rId-default-header\0rIdWatermark".to_owned(),
            ImageData {
                data: vec![1, 2, 3, 4],
                content_type: "image/png".to_owned(),
            },
        );
        input.document.body.sect_pr = Some(section);
        input
    }

    fn header_footer_page_text(page: &PageFrame) -> String {
        page.elements
            .iter()
            .filter_map(|element| match element {
                PositionedElement::Text(text) => Some(text.text.as_str()),
                _ => None,
            })
            .collect()
    }

    fn assert_header_footer_context_miss(
        base: &LayoutInput,
        name: &str,
        mutate: impl FnOnce(&mut LayoutInput),
    ) {
        let mut engine = Engine::new_deterministic().expect("bundled fonts load");
        engine.layout(base).expect("prime exact identity");
        let mut changed = base.clone();
        mutate(&mut changed);
        engine
            .layout(&changed)
            .unwrap_or_else(|error| panic!("{name}: {error}"));
        assert_eq!(engine.header_footer_cache_counts(), (0, 12), "{name}");
    }

    #[test]
    fn safe_header_footer_variants_reuse_exactly() {
        let mut input = cacheable_header_footer_input(&"body ".repeat(4_000));
        let mut engine = Engine::new_deterministic().expect("bundled fonts load");
        let (cold, cold_sources) = engine
            .layout_with_provenance(&input)
            .expect("cold header/footer layout");
        assert!(cold.pages.len() > 2);
        assert!(header_footer_page_text(&cold.pages[0]).contains("first header"));
        assert!(header_footer_page_text(&cold.pages[0]).contains("first footer"));
        assert!(header_footer_page_text(&cold.pages[1]).contains("even header"));
        assert!(header_footer_page_text(&cold.pages[1]).contains("even footer"));
        assert!(header_footer_page_text(&cold.pages[2]).contains("default header"));
        assert!(header_footer_page_text(&cold.pages[2]).contains("default footer"));
        assert_eq!(engine.header_footer_cache_counts(), (0, 6));
        assert_eq!(engine.header_footer_cache.len(), 6);
        assert!(
            engine
                .header_footer_cache
                .iter()
                .all(|entry| !entry.font_trace.is_empty())
        );

        for (index, entry) in engine.header_footer_cache.iter_mut().enumerate() {
            entry.diagnostics = vec![Diagnostic {
                message: format!("cached header/footer diagnostic {index}"),
            }];
        }
        input.document.body.content.insert(
            0,
            BodyContent::Paragraph({
                let mut paragraph = CT_P::new();
                paragraph.add_run("inserted body source");
                paragraph
            }),
        );
        let (warm, warm_sources) = engine
            .layout_with_provenance(&input)
            .expect("warm header/footer layout");
        let (fresh, fresh_sources) = Engine::new_deterministic()
            .expect("bundled fonts load")
            .layout_with_provenance(&input)
            .expect("fresh comparison layout");
        assert_eq!(format!("{:?}", warm.pages), format!("{:?}", fresh.pages));
        assert_eq!(format!("{:?}", warm.fonts), format!("{:?}", fresh.fonts));
        assert_eq!(
            format!("{:?}", warm.outlines),
            format!("{:?}", fresh.outlines)
        );
        assert_eq!(warm_sources, fresh_sources);
        assert_ne!(cold_sources, warm_sources);
        assert_eq!(engine.header_footer_cache_counts(), (6, 6));
        assert_eq!(warm.diagnostics.len(), 6);
        assert!(warm.diagnostics.iter().all(|diagnostic| {
            diagnostic
                .message
                .starts_with("cached header/footer diagnostic")
        }));
        let header_source = warm
            .pages
            .iter()
            .flat_map(|page| &page.elements)
            .filter_map(|element| match element {
                PositionedElement::Text(text) if text.text.contains("header") => text.source,
                _ => None,
            })
            .next()
            .expect("cached header text has provenance");
        assert!(matches!(
            warm_sources[header_source.node.get() as usize - 1].story,
            WordStory::Header { .. }
        ));

        // F-X042 resolves inherited references onto each section before layout.
        // Model that exact input shape and prove both the authored first section
        // and inherited final section reuse their variants on the next layout.
        let mut inherited_input = cacheable_header_footer_input(&"second section ".repeat(2_000));
        let inherited_section = inherited_input
            .document
            .body
            .sect_pr
            .clone()
            .expect("final section");
        let mut first_section_end = CT_P::new();
        first_section_end.add_run("authored section with shared variants");
        first_section_end.properties.get_or_insert_default().sect_pr = Some(inherited_section);
        inherited_input
            .document
            .body
            .content
            .insert(0, BodyContent::Paragraph(first_section_end));
        let mut inherited_engine = Engine::new_deterministic().expect("bundled fonts load");
        inherited_engine
            .layout_with_provenance(&inherited_input)
            .expect("cold inherited layout");
        let (inherited_warm, inherited_sources) = inherited_engine
            .layout_with_provenance(&inherited_input)
            .expect("warm inherited layout");
        let (inherited_fresh, fresh_inherited_sources) = Engine::new_deterministic()
            .expect("bundled fonts load")
            .layout_with_provenance(&inherited_input)
            .expect("fresh inherited layout");
        assert_eq!(inherited_engine.header_footer_cache_counts(), (12, 12));
        assert_eq!(
            format!("{:?}", inherited_warm.pages),
            format!("{:?}", inherited_fresh.pages)
        );
        assert_eq!(inherited_sources, fresh_inherited_sources);
    }

    #[test]
    fn header_footer_media_geometry_and_context_changes_miss() {
        use rdocx_oxml::footnotes::CT_Footnotes;
        use rdocx_oxml::numbering::CT_Numbering;
        use rdocx_oxml::theme::Theme;
        use rdocx_oxml::units::Twips;

        let base = cacheable_header_footer_input("body");
        assert_header_footer_context_miss(&base, "header text", |input| {
            input.headers.insert(
                "rId-first-header".to_owned(),
                header_footer_part("changed first header"),
            );
        });
        assert_header_footer_context_miss(&base, "media bytes", |input| {
            input
                .images
                .get_mut("rId-default-header\0rIdWatermark")
                .expect("watermark image")
                .data
                .push(5);
        });
        assert_header_footer_context_miss(&base, "watermark", |input| {
            input.headers.insert(
                "rId-default-header".to_owned(),
                image_watermark_header("default header", 73.0),
            );
        });
        assert_header_footer_context_miss(&base, "same-width page height", |input| {
            input
                .document
                .body
                .sect_pr
                .as_mut()
                .expect("section")
                .page_height = Some(Twips(15_841));
        });
        assert_header_footer_context_miss(&base, "styles", |input| {
            input.styles = CT_Styles::new();
        });
        assert_header_footer_context_miss(&base, "numbering", |input| {
            input.numbering = Some(CT_Numbering::new());
        });
        assert_header_footer_context_miss(&base, "notes", |input| {
            input.footnotes = Some(CT_Footnotes::new());
        });
        assert_header_footer_context_miss(&base, "theme", |input| {
            input.theme = Some(Theme::default());
        });
        assert_header_footer_context_miss(&base, "revision", |input| {
            input.revision_view = RevisionView::Tracked;
        });
        assert_header_footer_context_miss(&base, "fonts", |input| {
            let (family, data) = oxml_layout::bundled_fonts::bundled_font_data()[0];
            input.fonts.push(oxml_layout::FontFile {
                family: family.to_owned(),
                data: data.to_vec(),
            });
        });

        let mut source_mode = Engine::new_deterministic().expect("bundled fonts load");
        source_mode.layout(&base).expect("prime unsourced cache");
        source_mode
            .layout_with_provenance(&base)
            .expect("sourced layout misses unsourced entries");
        assert_eq!(source_mode.header_footer_cache_counts(), (0, 12));

        let mut unsafe_input = base.clone();
        unsafe_input
            .headers
            .get_mut("rId-first-header")
            .expect("first header")
            .paragraphs[0]
            .properties
            .get_or_insert_default()
            .num_id = Some(1);
        let mut unsafe_engine = Engine::new_deterministic().expect("bundled fonts load");
        unsafe_engine
            .layout(&unsafe_input)
            .expect("unsafe part lays out");
        assert_eq!(unsafe_engine.header_footer_cache_counts(), (0, 5));

        let mut opaque_input = base.clone();
        let opaque_run = &mut opaque_input
            .headers
            .get_mut("rId-first-header")
            .expect("first header")
            .paragraphs[0]
            .runs[0];
        opaque_run
            .extra_xml
            .push(br#"<w:object xmlns:w="urn:unrepresented"/>"#.to_vec());
        opaque_run.extra_xml_positions.push(0);
        let mut opaque_engine = Engine::new_deterministic().expect("bundled fonts load");
        opaque_engine
            .layout(&opaque_input)
            .expect("opaque producer XML lays out without reuse");
        assert_eq!(opaque_engine.header_footer_cache_counts(), (0, 5));

        let foreign_wrapper = rdocx_oxml::header_footer::CT_HdrFtr::from_xml(
            format!(
                r#"<w:hdr xmlns:w="{}" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:x="urn:producer"><w:p><w:r><x:pict><w:pict><v:shape style="width:72pt;height:36pt"><v:textpath string="DRAFT"/></v:shape></w:pict></x:pict></w:r></w:p></w:hdr>"#,
                rdocx_oxml::namespace::W_NS
            )
            .as_bytes(),
        )
        .expect("foreign pict wrapper parses");
        assert_eq!(foreign_wrapper.watermarks().len(), 1);
        assert!(!header_footer_part_is_cache_safe(
            &foreign_wrapper,
            &base.styles
        ));
        let rebound_word_prefix = rdocx_oxml::header_footer::CT_HdrFtr::from_xml(
            format!(
                r#"<q:hdr xmlns:q="{}" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w="urn:producer"><q:p><q:r><w:pict><q:pict><v:shape style="width:72pt;height:36pt"><v:textpath string="DRAFT"/></v:shape></q:pict></w:pict></q:r></q:p></q:hdr>"#,
                rdocx_oxml::namespace::W_NS
            )
            .as_bytes(),
        )
        .expect("rebound conventional prefix parses");
        assert_eq!(rebound_word_prefix.watermarks().len(), 1);
        assert!(!header_footer_part_is_cache_safe(
            &rebound_word_prefix,
            &base.styles
        ));
    }

    #[test]
    fn header_footer_cache_publishes_transactionally_and_stays_bounded() {
        use rdocx_oxml::header_footer::HdrFtrRef;

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

        let mut failing_input = make_input_with_text("section-ending prefix");
        let mut section = CT_SectPr::default_letter();
        section.header_refs.push(HdrFtrRef {
            hdr_ftr_type: HdrFtrType::Default,
            rel_id: "rIdHeader".to_owned(),
        });
        let BodyContent::Paragraph(prefix) = &mut failing_input.document.body.content[0] else {
            panic!("prefix paragraph");
        };
        prefix.properties.get_or_insert_default().sect_pr = Some(section);
        prefix.runs[0].properties.get_or_insert_default().font_ascii =
            Some(valid_family.to_owned());
        failing_input.headers.insert(
            "rIdHeader".to_owned(),
            header_footer_part("staged header before late failure"),
        );
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
        let mut failing = Engine::with_font_manager(FontManager::new_with_fonts(vec![(
            valid_family.to_owned(),
            valid_bytes.to_vec(),
        )]));
        assert!(failing.layout(&failing_input).is_err());
        assert!(failing.header_footer_cache.is_empty());
        assert_eq!(failing.header_footer_cache_bytes, 0);
        assert_eq!(failing.header_footer_cache_counts(), (0, 1));

        let mut bounded_input = make_input_with_text("bounded body");
        let mut bounded_section = CT_SectPr::default_letter();
        for index in 0..(HEADER_FOOTER_CACHE_MAX_ENTRIES * 2) {
            let relationship_id = format!("rIdHeader{index:03}");
            bounded_section.header_refs.push(HdrFtrRef {
                hdr_ftr_type: HdrFtrType::Default,
                rel_id: relationship_id.clone(),
            });
            bounded_input.headers.insert(
                relationship_id,
                header_footer_part(&format!("bounded header {index:03}")),
            );
        }
        bounded_input.document.body.sect_pr = Some(bounded_section);
        let mut bounded = Engine::new_deterministic().expect("bundled fonts load");
        bounded
            .layout(&bounded_input)
            .expect("bounded pending layout succeeds");
        assert_eq!(
            bounded.header_footer_cache.len(),
            HEADER_FOOTER_CACHE_MAX_ENTRIES
        );
        assert!(bounded.header_footer_cache_bytes <= HEADER_FOOTER_CACHE_MAX_BYTES);
        assert!(
            bounded.pending_header_footer_cache_peak_entries <= HEADER_FOOTER_CACHE_MAX_ENTRIES
        );
        assert!(bounded.pending_header_footer_cache_peak_bytes <= HEADER_FOOTER_CACHE_MAX_BYTES);
        assert_eq!(
            bounded.header_footer_cache_bytes,
            bounded
                .header_footer_cache
                .iter()
                .map(|entry| entry.bytes)
                .sum::<usize>()
        );
        assert!(
            bounded.paragraph_cache.len()
                + bounded.table_cache.len()
                + bounded.header_footer_cache.len()
                + bounded
                    .restart_cache
                    .as_ref()
                    .map_or(0, |cache| cache.checkpoints.len())
                <= CACHE_MAX_ENTRIES
        );
        assert!(
            bounded.paragraph_cache_bytes
                + bounded.table_cache_bytes
                + bounded.header_footer_cache_bytes
                + bounded
                    .restart_cache
                    .as_ref()
                    .map_or(0, |cache| cache.bytes)
                <= CACHE_MAX_BYTES
        );

        let one = cacheable_header_footer_input("oversized entry body");
        let mut oversized = Engine::new_deterministic().expect("bundled fonts load");
        oversized.layout(&one).expect("prime oversized template");
        let mut entry = oversized
            .header_footer_cache
            .pop_front()
            .expect("header/footer template retained");
        let mut oversized_key = oversized
            .header_footer_cache
            .pop_front()
            .expect("second header/footer template retained");
        oversized.header_footer_cache.clear();
        oversized.header_footer_cache_bytes = 0;
        let mut reserved_namespace = String::with_capacity(HEADER_FOOTER_CACHE_MAX_BYTES + 1);
        reserved_namespace.push('x');
        oversized_key
            .key
            .part
            .extra_namespaces
            .push((reserved_namespace, "urn:test".to_owned()));
        oversized_key.bytes = header_footer_cache_entry_bytes(
            &oversized_key.key,
            &oversized_key.content,
            &oversized_key.diagnostics,
            &oversized_key.font_trace,
        );
        assert!(oversized_key.bytes > HEADER_FOOTER_CACHE_MAX_BYTES);
        oversized.publish_header_footer_cache_entry(oversized_key);
        assert!(oversized.header_footer_cache.is_empty());
        assert_eq!(oversized.header_footer_cache_bytes, 0);

        let text = entry.content.blocks[0]
            .lines
            .iter_mut()
            .flat_map(|line| &mut line.items)
            .find_map(|item| match item {
                LineItem::Text(text) => Some(text),
                _ => None,
            })
            .expect("template has text");
        text.advances = vec![0.0; HEADER_FOOTER_CACHE_MAX_BYTES / 8 + 1];
        entry.bytes = header_footer_cache_entry_bytes(
            &entry.key,
            &entry.content,
            &entry.diagnostics,
            &entry.font_trace,
        );
        assert!(entry.bytes > HEADER_FOOTER_CACHE_MAX_BYTES);
        oversized.publish_header_footer_cache_entry(entry);
        assert!(oversized.header_footer_cache.is_empty());
        assert_eq!(oversized.header_footer_cache_bytes, 0);
    }

    #[test]
    fn word_projection_leaves_break_segmentation_to_shared_layout() {
        let input = make_input_with_text("financial planning");
        let BodyContent::Paragraph(paragraph) = &input.document.body.content[0] else {
            panic!("expected paragraph");
        };
        let media = MediaRegistry::new(&input.images);
        let mut fonts = FontManager::new_deterministic().expect("bundled fonts load");
        let mut numbering = NumberingState::new();
        let mut diagnostics = Vec::new();
        let block = layout_paragraph_with_source(
            paragraph,
            468.0,
            &input.styles,
            &input,
            &media,
            &mut fonts,
            &mut numbering,
            &mut diagnostics,
            SourceNodeId::new(1),
        )
        .expect("paragraph lays out");
        let text_items = block
            .reflow
            .expect("line-breaking inputs retained")
            .items
            .into_iter()
            .filter_map(|item| match item {
                InlineItem::Text(segment) => Some(segment),
                _ => None,
            })
            .collect::<Vec<_>>();

        assert_eq!(text_items.len(), 1);
        assert_eq!(text_items[0].text, "financial planning");
        assert_eq!(text_items[0].source.expect("source span").char_start, 0);
        assert_eq!(text_items[0].source.expect("source span").char_end, 18);
    }

    #[test]
    fn break_opportunities_emit_every_scalar_and_glyph_once() {
        let text = "financial planning ttf-parser  double  spaces e\u{301}lan allocated \u{754c} "
            .repeat(12);
        let input = make_input_with_text(&text);
        let result = crate::layout_document_deterministic_with_provenance(&input)
            .expect("deterministic layout");
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

        assert_eq!(
            runs.iter().map(|run| run.text.as_str()).collect::<String>(),
            text
        );
        let mut expected_start = 0;
        let mut fonts = FontManager::new_deterministic().expect("bundled fonts load");
        for run in runs {
            let source = run.source.expect("filtered sourced run");
            assert_eq!(source.char_start, expected_start);
            assert_eq!(
                source.char_end - source.char_start,
                run.text.chars().count() as u32
            );
            expected_start = source.char_end;

            let family = result
                .layout
                .fonts
                .iter()
                .find(|font| font.id == run.font_id)
                .expect("run font is in result")
                .family
                .clone();
            let font_id = fonts
                .resolve_font(Some(&family), run.bold, run.italic)
                .expect("bundled run font resolves");
            let independently_shaped = fonts
                .shape_text(font_id, &run.text, run.font_size)
                .expect("emitted chunk reshapes");
            assert_eq!(
                run.glyph_ids, independently_shaped.glyph_ids,
                "{}",
                run.text
            );
            assert_eq!(run.advances, independently_shaped.advances, "{}", run.text);
        }
        assert_eq!(expected_start, text.chars().count() as u32);
    }

    #[test]
    fn reported_words_do_not_duplicate_boundary_glyphs() {
        for text in [
            "ttf-parser follows",
            "double  spaces follow",
            "financial planning",
            "allocated space",
        ] {
            let input = make_input_with_text(text);
            let result = crate::layout_document_deterministic_with_provenance(&input)
                .expect("deterministic layout");
            let mut fonts = FontManager::new_deterministic().expect("bundled fonts load");
            for run in result.layout.pages.iter().flat_map(|page| {
                page.elements.iter().filter_map(|element| match element {
                    PositionedElement::Text(run) if run.source.is_some() => Some(run),
                    _ => None,
                })
            }) {
                let family = result
                    .layout
                    .fonts
                    .iter()
                    .find(|font| font.id == run.font_id)
                    .expect("run font is in result")
                    .family
                    .clone();
                let font_id = fonts
                    .resolve_font(Some(&family), run.bold, run.italic)
                    .expect("bundled run font resolves");
                let independently_shaped = fonts
                    .shape_text(font_id, &run.text, run.font_size)
                    .expect("emitted chunk reshapes");
                assert_eq!(run.glyph_ids, independently_shaped.glyph_ids, "{text}");
                assert_eq!(run.advances, independently_shaped.advances, "{text}");
            }
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
    fn paragraph_fingerprint_collision_requires_typed_equality() {
        let first = make_input_with_text("first collision candidate");
        let second = make_input_with_text("second collision candidate");
        let BodyContent::Paragraph(second_paragraph) = &second.document.body.content[0] else {
            panic!("body item is a paragraph");
        };
        let mut engine = Engine::new_deterministic().expect("bundled fonts load");
        engine.layout(&first).expect("first layout succeeds");

        let forced_fingerprint = paragraph_fingerprint(second_paragraph);
        let retained = engine
            .paragraph_cache
            .front_mut()
            .expect("first paragraph is retained");
        assert_ne!(retained.fingerprint, forced_fingerprint);
        retained.fingerprint = forced_fingerprint;

        let output = engine.layout(&second).expect("collision layout succeeds");
        assert_eq!(output_text(&output).concat(), "second collision candidate");
        assert_eq!(engine.paragraph_cache_counts(), (0, 2));
    }

    #[test]
    fn editor_scale_paragraph_cache_avoids_warm_thrash() {
        let mut input = make_input_with_text("editor paragraph 000");
        for index in 1..700 {
            let mut paragraph = CT_P::new();
            paragraph.add_run(&format!("editor paragraph {index:03}"));
            input.document.body.add_paragraph(paragraph);
        }
        let mut engine = Engine::new_deterministic().expect("bundled fonts load");
        engine
            .layout_with_provenance(&input)
            .expect("editor cold layout succeeds");
        assert_eq!(engine.paragraph_cache.len(), 700);
        assert_eq!(engine.paragraph_cache_counts(), (0, 700));

        set_body_paragraph_text(&mut input, 350, "editor paragraph 350 changed");
        let warm = engine
            .layout_with_provenance(&input)
            .expect("editor warm layout succeeds");
        let cold = Engine::new_deterministic()
            .expect("bundled fonts load")
            .layout_with_provenance(&input)
            .expect("editor cold comparison succeeds");

        assert_layout_results_equal(&warm.0, &cold.0);
        assert_eq!(warm.1, cold.1);
        assert_eq!(engine.paragraph_cache_counts(), (699, 701));
        assert_eq!(engine.paragraph_cache.len(), 701);
        assert_eq!(
            engine
                .paragraph_cache
                .front()
                .expect("insertion order has a front")
                .key
                .paragraph
                .text(),
            "editor paragraph 000"
        );
        let rebuilt = engine
            .last_rebuilt_page_range
            .clone()
            .expect("edited layout reports a rebuilt range");
        assert!(
            rebuilt.end.saturating_sub(rebuilt.start) <= 2,
            "{rebuilt:?}"
        );
    }

    #[test]
    fn unsafe_prefix_still_disables_later_paragraph_hits() {
        let mut note = CT_P::new();
        let mut note_run = CT_R::new("");
        note_run.content = vec![RunContent::FootnoteRef { id: 1 }];
        note.runs.push(note_run);

        let mut field = CT_P::new();
        let mut field_run = CT_R::new("");
        field_run.content = vec![RunContent::Field(Field::new("PAGE", "1"))];
        field.runs.push(field_run);

        let mut numbered = CT_P::new();
        numbered.add_run("numbered prefix");
        numbered.properties.get_or_insert_default().num_id = Some(1);

        for (name, unsafe_paragraph) in [("note", note), ("field", field), ("numbering", numbered)]
        {
            let mut input = make_input_with_text("safe cached suffix");
            let mut engine = Engine::new_deterministic().expect("bundled fonts load");
            engine.layout(&input).expect("prime safe suffix");
            input
                .document
                .body
                .content
                .insert(0, BodyContent::Paragraph(unsafe_paragraph));

            let warm = engine.layout(&input).expect("warm unsafe-prefix layout");
            let cold = Engine::new_deterministic()
                .expect("bundled fonts load")
                .layout(&input)
                .expect("cold unsafe-prefix layout");
            assert_layout_results_equal(&warm, &cold);
            assert_eq!(engine.paragraph_cache_counts(), (0, 2), "{name}");
        }
    }

    #[test]
    fn scaled_paragraph_cache_warm_equals_cold() {
        let mut input = make_input_with_text("warm-cold paragraph 000");
        for index in 1..700 {
            let mut paragraph = CT_P::new();
            paragraph.add_run(&format!("warm-cold paragraph {index:03}"));
            input.document.body.add_paragraph(paragraph);
        }
        let mut warm_engine = Engine::new_deterministic().expect("bundled fonts load");
        warm_engine
            .layout_with_provenance(&input)
            .expect("prime warm state");
        set_body_paragraph_text(&mut input, 349, "warm-cold paragraph 349 changed");

        let warm = warm_engine
            .layout_with_provenance(&input)
            .expect("warm edited layout");
        let cold = Engine::new_deterministic()
            .expect("bundled fonts load")
            .layout_with_provenance(&input)
            .expect("cold edited layout");
        assert_layout_results_equal(&warm.0, &cold.0);
        assert_eq!(warm.1, cold.1);
        assert_eq!(format!("{:?}", warm.0), format!("{:?}", cold.0));
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
    fn compatible_engine_take_reuses_normal_layout_work() {
        let mut source_input = make_input_with_text("unchanged paragraph");
        let mut changed = CT_P::new();
        changed.add_run("old second paragraph");
        source_input.document.body.add_paragraph(changed);

        let mut source_engine = Engine::new_deterministic().expect("bundled fonts load");
        source_engine
            .layout(&source_input)
            .expect("prime reusable engine");
        assert_eq!(source_engine.paragraph_cache_counts(), (0, 2));

        let mut receiver_input = source_input.clone();
        let BodyContent::Paragraph(second) = &mut receiver_input.document.body.content[1] else {
            panic!("second body paragraph");
        };
        second.runs[0].content[0] =
            RunContent::Text(rdocx_oxml::text::CT_Text::new("new second paragraph"));

        let mut source = Some(source_engine);
        let mut transferred = Engine::take_if_compatible(&mut source, &receiver_input)
            .expect("matching context transfers");
        assert!(source.is_none());
        transferred
            .layout(&receiver_input)
            .expect("transferred layout succeeds");
        assert_eq!(transferred.paragraph_cache_counts(), (1, 3));
    }

    #[test]
    fn incompatible_or_failed_engine_take_preserves_the_source() {
        fn assert_rejected(label: &str, mut receiver: LayoutInput) {
            let source_input = make_input_with_text("retained paragraph");
            let mut engine = Engine::new_deterministic().expect("bundled fonts load");
            engine.layout(&source_input).expect("prime reusable engine");
            let mut source = Some(engine);
            assert!(
                Engine::take_if_compatible(&mut source, &receiver).is_none(),
                "{label} must reject transfer"
            );
            assert!(source.is_some());
            receiver.document.body.content.clear();
        }

        let base = make_input_with_text("retained paragraph");
        let mut changed = base.clone();
        changed.revision_view = RevisionView::Tracked;
        assert_rejected("revision view", changed);

        let wrapping = make_wrapping_document(
            WrapType::Square,
            Some(rdocx_oxml::drawing::AnchorAlignH::Left),
            100.0,
            40.0,
            5.0,
        );
        let mut changed = base.clone();
        changed.document = wrapping.document;
        assert_rejected("document wrapping state", changed);

        let mut changed = base.clone();
        changed.styles = CT_Styles::new();
        assert_rejected("styles", changed);

        let mut changed = base.clone();
        changed.numbering = Some(rdocx_oxml::numbering::CT_Numbering::new());
        assert_rejected("numbering", changed);

        let mut changed = base.clone();
        changed.headers.insert(
            "rIdHeader".to_owned(),
            rdocx_oxml::header_footer::CT_HdrFtr::new(),
        );
        assert_rejected("headers", changed);

        let mut changed = base.clone();
        changed.footers.insert(
            "rIdFooter".to_owned(),
            rdocx_oxml::header_footer::CT_HdrFtr::new(),
        );
        assert_rejected("footers", changed);

        let mut changed = base.clone();
        changed.images.insert(
            "rIdImage".to_owned(),
            crate::input::ImageData {
                data: vec![1, 2, 3],
                content_type: "image/png".to_owned(),
            },
        );
        assert_rejected("images", changed);

        let mut changed = base.clone();
        changed
            .charts
            .insert("rIdChart".to_owned(), Err("missing chart".to_owned()));
        assert_rejected("charts", changed);

        let mut changed = base.clone();
        changed.chart_theme.name = Some("Changed".to_owned());
        assert_rejected("chart theme", changed);

        let mut changed = base.clone();
        changed.core_properties = Some(rdocx_oxml::core_properties::CoreProperties {
            title: Some("Changed".to_owned()),
            ..Default::default()
        });
        assert_rejected("core properties", changed);

        let mut changed = base.clone();
        changed
            .hyperlink_urls
            .insert("rIdLink".to_owned(), "https://example.com".to_owned());
        assert_rejected("hyperlinks", changed);

        let mut changed = base.clone();
        changed.footnotes = Some(rdocx_oxml::footnotes::CT_Footnotes::new());
        assert_rejected("footnotes", changed);

        let mut changed = base.clone();
        changed.endnotes = Some(rdocx_oxml::footnotes::CT_Footnotes::new());
        assert_rejected("endnotes", changed);

        let mut changed = base.clone();
        changed.theme = Some(rdocx_oxml::theme::Theme::default());
        assert_rejected("theme", changed);

        let mut changed = base.clone();
        changed.fonts.push(oxml_layout::FontFile {
            family: "Changed".to_owned(),
            data: vec![1, 2, 3],
        });
        assert_rejected("fonts", changed);

        let mut changed = base;
        changed
            .document
            .body
            .sect_pr
            .get_or_insert_with(CT_SectPr::default_letter)
            .page_width = Some(rdocx_oxml::units::Twips(10_000));
        assert_rejected("sections", changed);
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

    fn safe_table(text: &str) -> CT_Tbl {
        let mut table = CT_Tbl::new();
        let mut row = CT_Row::new();
        let mut cell = CT_Tc::new();
        cell.paragraphs_mut()[0].add_run(text);
        row.cells.push(cell);
        table.rows.push(row);
        table
    }

    fn assert_layout_results_equal(left: &LayoutResult, right: &LayoutResult) {
        assert_eq!(left.pages.len(), right.pages.len());
        for (left, right) in left.pages.iter().zip(&right.pages) {
            assert_eq!(left.page_number, right.page_number);
            assert_eq!(left.width, right.width);
            assert_eq!(left.height, right.height);
            assert_eq!(left.elements, right.elements);
            assert_eq!(left.background, right.background);
        }
        assert_eq!(left.fonts.len(), right.fonts.len());
        for (left, right) in left.fonts.iter().zip(&right.fonts) {
            assert_eq!(left.id, right.id);
            assert_eq!(left.family, right.family);
            assert_eq!(left.data, right.data);
            assert_eq!(left.face_index, right.face_index);
            assert_eq!(left.bold, right.bold);
            assert_eq!(left.italic, right.italic);
        }
        assert_eq!(left.diagnostics, right.diagnostics);
        assert_eq!(left.outlines.len(), right.outlines.len());
        for (left, right) in left.outlines.iter().zip(&right.outlines) {
            assert_eq!(left.title, right.title);
            assert_eq!(left.level, right.level);
            assert_eq!(left.page_index, right.page_index);
            assert_eq!(left.y_position, right.y_position);
        }
    }

    #[test]
    fn earlier_note_insertion_invalidates_later_cached_markers() {
        let mut input = make_input_with_text("safe prefix");
        let mut later = CT_P::new();
        later.add_run("safe suffix");
        input.document.body.add_paragraph(later);
        let mut engine = Engine::new_deterministic().expect("bundled fonts load");
        let original = engine.layout(&input).expect("initial layout");

        let mut note_paragraph = CT_P::new();
        let mut note_run = CT_R::new("");
        note_run.content = vec![RunContent::FootnoteRef { id: 7 }];
        note_paragraph.runs.push(note_run);
        input
            .document
            .body
            .content
            .insert(1, BodyContent::Paragraph(note_paragraph));
        let warm_insert = engine.layout(&input).expect("warm insertion layout");
        let cold_insert = Engine::new_deterministic()
            .expect("bundled fonts load")
            .layout(&input)
            .expect("cold insertion layout");
        assert_layout_results_equal(&warm_insert, &cold_insert);
        assert_eq!(engine.paragraph_cache_counts(), (1, 3));

        input.document.body.content.remove(1);
        let warm_delete = engine.layout(&input).expect("warm deletion layout");
        let cold_delete = Engine::new_deterministic()
            .expect("bundled fonts load")
            .layout(&input)
            .expect("cold deletion layout");
        assert_layout_results_equal(&warm_delete, &cold_delete);
        assert_layout_results_equal(&original, &warm_delete);
        assert_eq!(engine.paragraph_cache_counts(), (3, 3));
    }

    #[test]
    fn safe_tables_reuse_transactionally_and_with_bounds() {
        let mut input = make_input_with_text("before table");
        input
            .document
            .body
            .content
            .push(BodyContent::Table(safe_table("cached cell")));
        let mut engine = Engine::new_deterministic().expect("bundled fonts load");
        let cold = engine.layout(&input).expect("cold table layout");
        let warm = engine.layout(&input).expect("warm table layout");
        assert_layout_results_equal(&cold, &warm);
        assert_eq!(engine.table_cache_counts(), (1, 1));

        let mut provenance_input = input.clone();
        let mut provenance_engine = Engine::new_deterministic().expect("bundled fonts load");
        provenance_engine
            .layout_with_provenance(&provenance_input)
            .expect("prime sourced table cache");
        let mut inserted = CT_P::new();
        inserted.add_run("inserted before table");
        provenance_input
            .document
            .body
            .content
            .insert(1, BodyContent::Paragraph(inserted));
        let (provenance_layout, sources) = provenance_engine
            .layout_with_provenance(&provenance_input)
            .expect("warm sourced table layout");
        let sourced_runs = provenance_layout
            .pages
            .iter()
            .flat_map(|page| &page.elements)
            .filter_map(|element| match element {
                PositionedElement::Text(run) => Some((run.text.clone(), run.source)),
                _ => None,
            })
            .collect::<Vec<_>>();
        let cached_cell = sourced_runs
            .iter()
            .find_map(|(text, source)| text.starts_with("cached").then_some(*source).flatten())
            .unwrap_or_else(|| panic!("cached cell keeps provenance: {sourced_runs:?}"));
        assert_eq!(
            sources[cached_cell.node.get() as usize - 1].children,
            [2, 0, 0, 0]
        );
        assert_eq!(provenance_engine.table_cache_counts(), (1, 1));

        let mut bounded = make_input_with_text("bounded prefix");
        for index in 0..(TABLE_CACHE_MAX_ENTRIES + 8) {
            bounded
                .document
                .body
                .content
                .push(BodyContent::Table(safe_table(&format!("table {index}"))));
        }
        let mut bounded_engine = Engine::new_deterministic().expect("bundled fonts load");
        bounded_engine
            .layout(&bounded)
            .expect("bounded table layout");
        assert!(bounded_engine.table_cache.len() <= TABLE_CACHE_MAX_ENTRIES);
        assert!(bounded_engine.table_cache_bytes <= TABLE_CACHE_MAX_BYTES);
        assert!(bounded_engine.pending_table_cache_peak_entries <= TABLE_CACHE_MAX_ENTRIES);
        assert!(bounded_engine.pending_table_cache_peak_bytes <= TABLE_CACHE_MAX_BYTES);

        let mut retained_border_block = engine
            .table_cache
            .back()
            .expect("safe table retained")
            .block
            .clone();
        let mut color = String::with_capacity(TABLE_CACHE_MAX_BYTES + 1);
        color.push_str("00");
        let mut edge =
            rdocx_oxml::borders::CT_BorderEdge::new(rdocx_oxml::shared::ST_Border::Single);
        edge.color = Some(color);
        retained_border_block.borders = Some(rdocx_oxml::table::CT_TblBorders {
            top: Some(edge),
            ..Default::default()
        });
        assert!(table_block_retained_bytes(&retained_border_block) > TABLE_CACHE_MAX_BYTES);

        let mut unsafe_table = safe_table("numbered cell");
        unsafe_table.rows[0].cells[0].paragraphs_mut()[0]
            .properties
            .get_or_insert_default()
            .num_id = Some(1);
        assert!(!table_is_cache_safe(&unsafe_table, &input.styles));

        let mut preserved_table = safe_table("preserved properties");
        preserved_table
            .properties
            .get_or_insert_default()
            .revision_xml
            .push(br#"<w:unknown/>"#.to_vec());
        assert!(!table_is_cache_safe(&preserved_table, &input.styles));

        let mut preserved_cell = safe_table("preserved cell properties");
        preserved_cell.rows[0].cells[0]
            .properties
            .get_or_insert_default()
            .extra_xml
            .push((0, br#"<w:unknown/>"#.to_vec()));
        assert!(!table_is_cache_safe(&preserved_cell, &input.styles));
    }

    fn restart_input() -> LayoutInput {
        let mut input = make_input_with_text("paragraph 000 stable line");
        for index in 1..140 {
            let mut paragraph = CT_P::new();
            paragraph.add_run(&format!("paragraph {index:03} stable line"));
            input.document.body.add_paragraph(paragraph);
        }
        input
    }

    fn set_body_paragraph_text(input: &mut LayoutInput, index: usize, text: &str) {
        let BodyContent::Paragraph(paragraph) = &mut input.document.body.content[index] else {
            panic!("body entry is a paragraph");
        };
        paragraph.runs[0].content = vec![RunContent::Text(rdocx_oxml::text::CT_Text {
            text: text.to_owned(),
            preserve_space: false,
        })];
    }

    fn substituted_restart_input() -> LayoutInput {
        let mut input = restart_input();
        let BodyContent::Paragraph(fields) = &mut input.document.body.content[0] else {
            panic!("body entry is a paragraph");
        };
        for (instruction, display) in [
            ("PAGE", "page"),
            ("NUMPAGES", "pages"),
            ("PAGEREF destination", "target"),
        ] {
            let mut run = CT_R::new("");
            run.content = vec![RunContent::Field(Field::new(instruction, display))];
            fields.runs.push(run);
        }
        let BodyContent::Paragraph(target) = &mut input.document.body.content[100] else {
            panic!("body entry is a paragraph");
        };
        assert!(target.insert_bookmark_start(0, 46, "destination"));
        assert!(target.insert_bookmark_end(1, 46));
        input
    }

    fn substituted_page_index(engine: &Engine) -> usize {
        engine
            .restart_cache
            .as_ref()
            .expect("restart record retained")
            .substitution_inputs
            .iter()
            .position(Option::is_some)
            .expect("field page retained")
    }

    #[test]
    fn unchanged_page_fields_reuse_substituted_frames() {
        let input = substituted_restart_input();
        let mut engine = Engine::new_deterministic().expect("bundled fonts load");
        let first = engine.layout(&input).expect("initial field layout");
        let field_page = substituted_page_index(&engine);
        let retained = engine
            .restart_cache
            .as_ref()
            .expect("restart record retained");
        assert!(
            retained.checkpoints.is_empty(),
            "field pages remain excluded from pagination restart"
        );
        assert!(!Arc::ptr_eq(
            &retained.raw_pages[field_page],
            &retained.pages[field_page]
        ));
        assert!(
            retained
                .raw_pages
                .iter()
                .zip(&retained.pages)
                .zip(&retained.substitution_inputs)
                .filter(|(_, inputs)| inputs.is_none())
                .all(|((pristine, substituted), _)| Arc::ptr_eq(pristine, substituted))
        );

        let warm = engine.layout(&input).expect("warm field layout");
        assert!(Arc::ptr_eq(
            &first.pages[field_page],
            &warm.pages[field_page]
        ));
    }

    #[test]
    fn changed_substitution_context_reshapes_pages() {
        fn assert_retained_key_miss(
            label: &str,
            mutate: impl FnOnce(&mut FieldSubstitutionInputs),
        ) {
            let input = substituted_restart_input();
            let mut engine = Engine::new_deterministic().expect("bundled fonts load");
            let first = engine.layout(&input).expect("initial field layout");
            let field_page = substituted_page_index(&engine);
            let retained = engine
                .restart_cache
                .as_mut()
                .expect("restart record retained");
            mutate(
                retained.substitution_inputs[field_page]
                    .as_mut()
                    .expect("field inputs retained"),
            );
            let warm = engine.layout(&input).expect("warm field layout");
            let cold = Engine::new_deterministic()
                .expect("bundled fonts load")
                .layout(&input)
                .expect("cold field layout");
            assert_layout_results_equal(&warm, &cold);
            assert!(
                !Arc::ptr_eq(&first.pages[field_page], &warm.pages[field_page]),
                "{label}"
            );
        }

        assert_retained_key_miss("page index must miss", |inputs| inputs.page_index += 1);
        assert_retained_key_miss("displayed page number must miss", |inputs| {
            inputs.page_number += 1;
        });
        assert_retained_key_miss("page count must miss", |inputs| inputs.total_pages += 1);
        assert_retained_key_miss("bookmark targets must miss", |inputs| {
            inputs.bookmark_pages.push((usize::MAX, usize::MAX));
        });
        assert_retained_key_miss("font identity must miss", |inputs| {
            inputs.font_identity.reverse();
            inputs.font_identity.push(FontId(u32::MAX));
        });
        assert_retained_key_miss("revision view must miss", |inputs| {
            inputs.revision_view = RevisionView::Tracked;
        });

        let mut input = substituted_restart_input();
        let mut engine = Engine::new_deterministic().expect("bundled fonts load");
        let first = engine.layout(&input).expect("initial field layout");
        let field_page = substituted_page_index(&engine);
        set_body_paragraph_text(&mut input, 0, "changed pristine field page");
        let warm = engine.layout(&input).expect("changed pristine layout");
        let cold = Engine::new_deterministic()
            .expect("bundled fonts load")
            .layout(&input)
            .expect("cold changed pristine layout");
        assert_layout_results_equal(&warm, &cold);
        assert!(!Arc::ptr_eq(
            &first.pages[field_page],
            &warm.pages[field_page]
        ));

        fn set_field_page_family(input: &mut LayoutInput, family: &str) {
            let BodyContent::Paragraph(fields) = &mut input.document.body.content[0] else {
                panic!("body entry is a paragraph");
            };
            for run in &mut fields.runs {
                run.properties = Some(rdocx_oxml::properties::CT_RPr {
                    font_ascii: Some(family.to_owned()),
                    font_hansi: Some(family.to_owned()),
                    ..Default::default()
                });
            }
        }

        let mut input = substituted_restart_input();
        set_field_page_family(&mut input, "Caladea");
        let mut engine = Engine::new_deterministic().expect("bundled fonts load");
        engine.layout(&input).expect("first bundled-family layout");
        set_field_page_family(&mut input, "Carlito");
        let transitioned = engine.layout(&input).expect("font-transition field layout");
        let field_page = substituted_page_index(&engine);
        let field_free_page = engine
            .restart_cache
            .as_ref()
            .expect("restart record retained")
            .substitution_inputs
            .iter()
            .position(Option::is_none)
            .expect("field-free page retained");
        let warm = engine.layout(&input).expect("post-transition warm layout");
        let cold = Engine::new_deterministic()
            .expect("bundled fonts load")
            .layout(&input)
            .expect("post-transition cold layout");
        assert_layout_results_equal(&warm, &cold);
        assert!(Arc::ptr_eq(
            &transitioned.pages[field_page],
            &warm.pages[field_page]
        ));
        assert!(Arc::ptr_eq(
            &transitioned.pages[field_free_page],
            &warm.pages[field_free_page]
        ));
    }

    #[test]
    fn substituted_page_reuse_is_bounded_and_complete_equal() {
        let input = substituted_restart_input();
        let mut engine = Engine::new_deterministic().expect("bundled fonts load");
        engine.layout(&input).expect("initial field layout");
        let warm = engine.layout(&input).expect("warm field layout");
        let cold = Engine::new_deterministic()
            .expect("bundled fonts load")
            .layout(&input)
            .expect("cold field layout");
        assert_layout_results_equal(&warm, &cold);
        let retained = engine
            .restart_cache
            .as_ref()
            .expect("restart record retained");
        assert!(
            retained.raw_pages.len().max(retained.checkpoints.len()) <= RESTART_CACHE_MAX_ENTRIES
        );
        assert!(retained.bytes <= RESTART_CACHE_MAX_BYTES);

        let mut oversized = make_input_with_text("");
        oversized.document.body.content.clear();
        for page in 0..=RESTART_CACHE_MAX_ENTRIES {
            let mut paragraph = CT_P::new();
            paragraph.properties = Some(CT_PPr {
                page_break_before: (page > 0).then_some(true),
                ..Default::default()
            });
            let mut run = CT_R::new("");
            run.content = vec![RunContent::Field(Field::new("PAGE", "1"))];
            paragraph.runs.push(run);
            oversized.document.body.add_paragraph(paragraph);
        }
        let mut bounded = Engine::new_deterministic().expect("bundled fonts load");
        let result = bounded
            .layout(&oversized)
            .expect("oversized layout succeeds");
        assert_eq!(result.pages.len(), RESTART_CACHE_MAX_ENTRIES + 1);
        assert!(
            bounded.restart_cache.is_none(),
            "an oversized pair set drops the optimization"
        );

        fn over_limit<T>() -> usize {
            RESTART_CACHE_MAX_BYTES / std::mem::size_of::<T>() + 1
        }
        fn assert_capacity_rejected(label: &str, mutate: impl FnOnce(&mut RestartCache)) {
            let mut candidate = RestartCache {
                body: Vec::new(),
                with_provenance: false,
                raw_pages: Vec::new(),
                pages: Vec::new(),
                substitution_inputs: Vec::new(),
                outlines: Vec::new(),
                checkpoints: Vec::new(),
                font_trace: Vec::new(),
                bytes: 0,
            };
            mutate(&mut candidate);
            assert!(
                restart_cache_bytes(&candidate) > RESTART_CACHE_MAX_BYTES,
                "{label}"
            );
        }

        assert_capacity_rejected("body vector capacity is charged", |cache| {
            cache.body = Vec::with_capacity(over_limit::<String>());
        });
        assert_capacity_rejected("pristine page vector capacity is charged", |cache| {
            cache.raw_pages = Vec::with_capacity(over_limit::<Arc<PageFrame>>());
        });
        assert_capacity_rejected("substituted page vector capacity is charged", |cache| {
            cache.pages = Vec::with_capacity(over_limit::<Arc<PageFrame>>());
        });
        assert_capacity_rejected("substitution vector capacity is charged", |cache| {
            cache.substitution_inputs =
                Vec::with_capacity(over_limit::<Option<FieldSubstitutionInputs>>());
        });
        assert_capacity_rejected("outline vector capacity is charged", |cache| {
            cache.outlines = Vec::with_capacity(over_limit::<oxml_layout::OutlineEntry>());
        });
        assert_capacity_rejected("checkpoint vector capacity is charged", |cache| {
            cache.checkpoints = Vec::with_capacity(over_limit::<paginator::PaginationCheckpoint>());
        });
        assert_capacity_rejected("font trace vector capacity is charged", |cache| {
            cache.font_trace = Vec::with_capacity(over_limit::<FontId>());
        });
        assert_capacity_rejected("body string capacity is charged", |cache| {
            cache
                .body
                .push(String::with_capacity(RESTART_CACHE_MAX_BYTES + 1));
        });
        assert_capacity_rejected("outline title capacity is charged", |cache| {
            cache.outlines.push(oxml_layout::OutlineEntry {
                title: String::with_capacity(RESTART_CACHE_MAX_BYTES + 1),
                level: 1,
                page_index: 0,
                y_position: 0.0,
            });
        });
        for nested in ["bookmark targets", "font identity"] {
            assert_capacity_rejected(&format!("{nested} capacity is charged"), |cache| {
                cache
                    .substitution_inputs
                    .push(Some(FieldSubstitutionInputs {
                        page_index: 0,
                        page_number: 1,
                        total_pages: 1,
                        bookmark_pages: if nested == "bookmark targets" {
                            Vec::with_capacity(over_limit::<(usize, usize)>())
                        } else {
                            Vec::new()
                        },
                        font_identity: if nested == "font identity" {
                            Vec::with_capacity(over_limit::<FontId>())
                        } else {
                            Vec::new()
                        },
                        revision_view: RevisionView::Accepted,
                    }));
            });
        }
    }

    #[test]
    fn warm_restart_rebuilds_only_the_bounded_changed_region() {
        let mut input = restart_input();
        let mut engine = Engine::new_deterministic().expect("bundled fonts load");
        let cold_initial = engine.layout(&input).expect("initial pagination");
        let initial_pages = cold_initial.pages.clone();
        let retained = engine
            .restart_cache
            .as_ref()
            .expect("restart state retained");
        assert!(retained.checkpoints.len() > 1);
        assert!(retained.checkpoints.len() <= RESTART_CACHE_MAX_ENTRIES);
        assert!(retained.bytes <= RESTART_CACHE_MAX_BYTES);

        set_body_paragraph_text(&mut input, 70, "paragraph 070 changed line");
        let warm = engine.layout(&input).expect("warm middle edit");
        let cold = Engine::new_deterministic()
            .expect("bundled fonts load")
            .layout(&input)
            .expect("cold middle edit");
        assert_layout_results_equal(&warm, &cold);
        let rebuilt = engine
            .last_rebuilt_page_range
            .clone()
            .expect("rebuilt range recorded");
        assert!(
            rebuilt.end.saturating_sub(rebuilt.start) <= 2,
            "{rebuilt:?}"
        );
        assert!(
            warm.pages
                .iter()
                .zip(&initial_pages)
                .take(rebuilt.start)
                .all(|(current, previous)| Arc::ptr_eq(current, previous))
        );
        assert!(
            warm.pages
                .iter()
                .zip(&initial_pages)
                .skip(rebuilt.end)
                .all(|(current, previous)| Arc::ptr_eq(current, previous))
        );

        for (index, label) in [(0, "start"), (139, "tail")] {
            set_body_paragraph_text(
                &mut input,
                index,
                &format!("paragraph {index:03} {label:>7} line"),
            );
            let warm = engine.layout(&input).expect("warm boundary edit");
            let cold = Engine::new_deterministic()
                .expect("bundled fonts load")
                .layout(&input)
                .expect("cold boundary edit");
            assert_layout_results_equal(&warm, &cold);
        }
    }

    #[test]
    fn unsafe_pagination_state_falls_back_to_full_layout() {
        let mut input = restart_input();
        let BodyContent::Paragraph(paragraph) = &mut input.document.body.content[20] else {
            panic!("paragraph");
        };
        paragraph.properties.get_or_insert_default().keep_next = Some(true);
        let mut engine = Engine::new_deterministic().expect("bundled fonts load");
        let warm = engine.layout(&input).expect("keep layout");
        assert!(engine.restart_cache.is_none());
        set_body_paragraph_text(&mut input, 30, "paragraph 030 changed line");
        let next = engine.layout(&input).expect("fallback relayout");
        let cold = Engine::new_deterministic()
            .expect("bundled fonts load")
            .layout(&input)
            .expect("cold fallback layout");
        assert_layout_results_equal(&next, &cold);
        assert!(!warm.pages.is_empty());

        let assert_fallback = |label: &str, input: LayoutInput| {
            let mut engine = Engine::new_deterministic().expect("bundled fonts load");
            engine
                .layout(&input)
                .unwrap_or_else(|error| panic!("{label}: {error}"));
            assert!(engine.restart_cache.is_none(), "{label}");
        };

        let mut table = make_input_with_text("before table");
        table
            .document
            .body
            .content
            .push(BodyContent::Table(safe_table("table")));
        assert_fallback("table", table);

        let split = make_input_with_text(&"split paragraph content ".repeat(200));
        assert_fallback("split paragraph", split);

        assert_fallback(
            "floating drawing",
            make_wrapping_document(WrapType::Square, None, 120.0, 60.0, 5.0),
        );
        assert_fallback("note continuation", make_input_with_footnote(&["note"]));

        let mut sections = restart_input();
        let BodyContent::Paragraph(first) = &mut sections.document.body.content[20] else {
            panic!("paragraph");
        };
        first.properties.get_or_insert_default().sect_pr = Some(CT_SectPr::default_letter());
        assert_fallback("multiple sections", sections);

        let mut background = restart_input();
        background.document.background_xml = Some(
            br#"<w:background xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:color="FFFFFF"/>"#
                .to_vec(),
        );
        assert_fallback("page background", background);

        let mut field = restart_input();
        let BodyContent::Paragraph(paragraph) = &mut field.document.body.content[20] else {
            panic!("paragraph");
        };
        let mut field_run = CT_R::new("");
        field_run.content = vec![RunContent::Field(rdocx_oxml::text::Field::new("PAGE", "1"))];
        paragraph.runs.push(field_run);
        let mut field_engine = Engine::new_deterministic().expect("bundled fonts load");
        field_engine.layout(&field).expect("field layout");
        let retained = field_engine
            .restart_cache
            .as_ref()
            .expect("field substitution pairs retained");
        assert!(
            retained.checkpoints.is_empty(),
            "fields must not become pagination restart boundaries"
        );

        let mut header_footer = restart_input();
        let mut section = CT_SectPr::default_letter();
        section
            .header_refs
            .push(rdocx_oxml::header_footer::HdrFtrRef {
                hdr_ftr_type: rdocx_oxml::header_footer::HdrFtrType::Default,
                rel_id: "rIdHeader".to_owned(),
            });
        section
            .footer_refs
            .push(rdocx_oxml::header_footer::HdrFtrRef {
                hdr_ftr_type: rdocx_oxml::header_footer::HdrFtrType::Default,
                rel_id: "rIdFooter".to_owned(),
            });
        header_footer.document.body.sect_pr = Some(section);
        for (relationship_id, is_header) in [("rIdHeader", true), ("rIdFooter", false)] {
            let mut part = rdocx_oxml::header_footer::CT_HdrFtr::new();
            let mut paragraph = CT_P::new();
            paragraph.add_run(relationship_id);
            part.paragraphs.push(paragraph);
            if is_header {
                header_footer
                    .headers
                    .insert(relationship_id.to_owned(), part);
            } else {
                header_footer
                    .footers
                    .insert(relationship_id.to_owned(), part);
            }
        }
        assert_fallback("header and footer", header_footer);

        let mut boundary = restart_input();
        let mut boundary_engine = Engine::new_deterministic().expect("bundled fonts load");
        boundary_engine
            .layout(&boundary)
            .expect("prime boundary state");
        let stale = boundary_engine
            .restart_cache
            .as_mut()
            .and_then(|cache| {
                cache
                    .checkpoints
                    .iter_mut()
                    .find(|checkpoint| checkpoint.next_block_index > 80)
            })
            .expect("later checkpoint exists");
        stale.next_header_page_number += 1;
        set_body_paragraph_text(&mut boundary, 70, "changed before stale boundary");
        let warm = boundary_engine
            .layout(&boundary)
            .expect("warm boundary mismatch");
        let cold = Engine::new_deterministic()
            .expect("bundled fonts load")
            .layout(&boundary)
            .expect("cold boundary mismatch");
        assert_layout_results_equal(&warm, &cold);
        assert!(
            boundary_engine
                .restart_cache
                .as_ref()
                .expect("correct state republished")
                .checkpoints
                .iter()
                .all(|checkpoint| {
                    checkpoint.next_header_page_number == checkpoint.page_count + 1
                })
        );
    }

    #[test]
    fn warm_and_cold_outputs_are_complete_equals() {
        let mut input = restart_input();
        let mut engine = Engine::new_deterministic().expect("bundled fonts load");
        engine.layout(&input).expect("prime restart state");

        let mut inserted = CT_P::new();
        inserted.add_run("inserted stable line");
        input
            .document
            .body
            .content
            .insert(60, BodyContent::Paragraph(inserted));
        let warm_insert = engine.layout(&input).expect("warm insertion");
        let cold_insert = Engine::new_deterministic()
            .expect("bundled fonts load")
            .layout(&input)
            .expect("cold insertion");
        assert_layout_results_equal(&warm_insert, &cold_insert);

        input.document.body.content.remove(60);
        let warm_delete = engine.layout(&input).expect("warm deletion");
        let cold_delete = Engine::new_deterministic()
            .expect("bundled fonts load")
            .layout(&input)
            .expect("cold deletion");
        assert_layout_results_equal(&warm_delete, &cold_delete);

        let mut sourced_engine = Engine::new_deterministic().expect("bundled fonts load");
        let (original, original_sources) = sourced_engine
            .layout_with_provenance(&input)
            .expect("prime sourced restart state");
        input.document.body.content.insert(
            60,
            BodyContent::Paragraph({
                let mut paragraph = CT_P::new();
                paragraph.add_run("inserted sourced line");
                paragraph
            }),
        );
        let (warm_insert, warm_insert_sources) = sourced_engine
            .layout_with_provenance(&input)
            .expect("warm sourced insertion");
        let (cold_insert, cold_insert_sources) = Engine::new_deterministic()
            .expect("bundled fonts load")
            .layout_with_provenance(&input)
            .expect("cold sourced insertion");
        assert_layout_results_equal(&warm_insert, &cold_insert);
        assert_eq!(warm_insert_sources, cold_insert_sources);

        input.document.body.content.remove(60);
        let (warm_delete, warm_delete_sources) = sourced_engine
            .layout_with_provenance(&input)
            .expect("warm sourced deletion");
        assert_layout_results_equal(&warm_delete, &original);
        assert_eq!(warm_delete_sources, original_sources);

        let mut truncate_engine = Engine::new_deterministic().expect("bundled fonts load");
        truncate_engine
            .layout(&input)
            .expect("prime whole-suffix deletion state");
        let checkpoint = truncate_engine
            .restart_cache
            .as_ref()
            .and_then(|cache| cache.checkpoints.iter().next_back().copied())
            .expect("restart cache has a final safe boundary");
        input
            .document
            .body
            .content
            .truncate(checkpoint.next_block_index);
        let warm_truncate = truncate_engine
            .layout(&input)
            .expect("warm whole-suffix deletion");
        let cold_truncate = Engine::new_deterministic()
            .expect("bundled fonts load")
            .layout(&input)
            .expect("cold whole-suffix deletion");
        assert_layout_results_equal(&warm_truncate, &cold_truncate);
    }

    #[test]
    fn paragraph_cache_failure_and_eviction_remain_bounded() {
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
        failing_input
            .document
            .body
            .content
            .push(BodyContent::Table(safe_table("staged before failure")));
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
        assert!(failing.table_cache.is_empty());
        assert_eq!(failing.paragraph_cache_counts(), (0, 1));
        assert_eq!(failing.table_cache_counts(), (0, 1));

        let template_input = make_input_with_text("eviction template");
        let mut bounded = Engine::new_deterministic().expect("bundled fonts load");
        bounded
            .layout(&template_input)
            .expect("template layout succeeds");
        let template = bounded
            .paragraph_cache
            .pop_front()
            .expect("template paragraph is retained");
        bounded.paragraph_cache_bytes = 0;
        for index in 0..=PARAGRAPH_CACHE_MAX_ENTRIES {
            let mut paragraph = CT_P::new();
            paragraph.add_run(&format!("eviction paragraph {index}"));
            let bytes = paragraph_cache_entry_bytes(
                &paragraph,
                &template.block,
                &template.diagnostics,
                template.font_trace.len(),
            );
            bounded.publish_paragraph_cache_entry(ParagraphCacheEntry {
                fingerprint: paragraph_fingerprint(&paragraph),
                key: ParagraphCacheKey {
                    paragraph,
                    content_width_bits: PageGeometry::default().content_width().to_bits(),
                    revision_view: RevisionView::Accepted,
                },
                block: template.block.clone(),
                diagnostics: template.diagnostics.clone(),
                font_trace: template.font_trace.clone(),
                bytes,
            });
        }
        assert_eq!(bounded.paragraph_cache.len(), PARAGRAPH_CACHE_MAX_ENTRIES);
        assert!(bounded.paragraph_cache_bytes <= PARAGRAPH_CACHE_MAX_BYTES);
        assert_eq!(
            bounded
                .paragraph_cache
                .front()
                .expect("FIFO cache has a front")
                .key
                .paragraph
                .text(),
            "eviction paragraph 1"
        );
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
        assert_eq!(
            engine.pending_paragraph_cache_peak_entries,
            PARAGRAPH_CACHE_MAX_ENTRIES
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
            fingerprint: paragraph_fingerprint(paragraph),
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
            fingerprint: paragraph_fingerprint(&paragraph),
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
            for element in &mut Arc::make_mut(page).elements {
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
