#![doc = include_str!("../README.md")]
#![allow(non_camel_case_types)]
#![allow(clippy::too_many_arguments)]

pub mod block;
mod convert;
pub mod engine;
pub mod input;
mod math;
pub mod notes;
pub mod paginator;
pub mod style_resolver;
pub mod table;

pub use input::{ImageData, LayoutInput, MediaRegistry, RevisionStyle, RevisionView};
use oxml_layout::Diagnostic;
pub use oxml_layout::{
    Color, DocumentMetadata, FontData, FontFile, FontId, GlyphRun, LayoutError, LayoutResult,
    PageFrame, Point, PositionedElement, Rect, Result, SourceNodeId, SourceSpan,
};
use rdocx_oxml::document::BodyContent;

/// Word story containing a source paragraph.
#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum WordStory {
    Document,
    Header {
        relationship_id: String,
    },
    Footer {
        relationship_id: String,
    },
    Footnote {
        id: i32,
    },
    Endnote {
        id: i32,
    },
    TextBox {
        part_name: String,
        owner_children: Vec<usize>,
    },
}

/// Modeled path to one paragraph in a Word story.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct WordSourcePath {
    pub story: WordStory,
    pub children: Vec<usize>,
}

/// Point-space extent occupied by one top-level Word body item on one page.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct WordBodyLayoutFragment {
    /// One-based physical page number in the produced layout.
    pub physical_page: usize,
    /// One-based displayed page number after section page-number restarts.
    pub displayed_page: usize,
    /// Left edge in points from the physical page origin.
    pub x: f64,
    /// Top edge in points from the physical page origin.
    pub y: f64,
    /// Width in points.
    pub width: f64,
    /// Height in points.
    pub height: f64,
}

/// One physical page occupied by a section in this immutable layout snapshot.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct WordPageSection {
    pub physical_page: usize,
    pub displayed_page: usize,
    pub section_index: usize,
}

/// One occurrence of a field in this immutable layout snapshot.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct WordFieldPlacement {
    pub source: oxml_layout::FieldSource,
    /// Physical page on which this field display is painted.
    pub physical_page: usize,
    /// Word's semantic PAGE value. Note fields use their unique reference page.
    pub displayed_page: usize,
    /// Owning section, following the unique reference for note fields.
    pub section_index: usize,
}

/// One successful main-document sequence counter delta.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WordSequenceEvent {
    pub source: oxml_layout::FieldSource,
    pub accepted_run: usize,
    pub identifier: String,
    pub value: i64,
}

/// Immutable source-qualified sequence decisions without shaping or pagination.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct WordSequenceSnapshot {
    source_nodes: Vec<WordSourcePath>,
    navigation_document: Option<rdocx_oxml::document::CT_Document>,
    navigation_failure: Option<String>,
    navigation_part: Option<String>,
    navigation_exclusions: std::collections::HashMap<WordSourcePath, Vec<usize>>,
    navigation_owners: std::collections::HashSet<WordSourcePath>,
    source_ids: std::collections::HashMap<WordSourcePath, SourceNodeId>,
    field_source_xml: std::collections::HashMap<oxml_layout::FieldSource, Vec<u8>>,
    field_source_contexts:
        std::collections::HashMap<oxml_layout::FieldSource, (String, bool, bool)>,
    text_box_owner_indices: std::collections::HashMap<WordStory, usize>,
    physical_story_owners: std::collections::HashMap<WordStory, (String, WordStory)>,
    field_run_positions: std::collections::HashMap<oxml_layout::FieldSource, usize>,
    position_bookmark_starts: Vec<(i32, String, SourceNodeId, usize, usize)>,
    position_bookmark_ends: Vec<(i32, SourceNodeId, usize, usize)>,
    field_values:
        std::collections::HashMap<oxml_layout::FieldSource, std::result::Result<String, String>>,
    main_events: Vec<WordSequenceEvent>,
    event_sources: std::collections::HashMap<oxml_layout::FieldSource, usize>,
    identifier_events: std::collections::HashMap<String, Vec<usize>>,
    repeat_requests: std::collections::HashMap<
        oxml_layout::FieldSource,
        (rdocx_oxml::text::FieldInstruction, String),
    >,
}

impl WordSequenceSnapshot {
    /// Attach only the facade's checked physical generated-result exclusions.
    /// Full layout regenerates all counters and validates this complete source stamp.
    #[doc(hidden)]
    pub fn set_navigation_exclusions(
        &mut self,
        input: &LayoutInput,
        exclusions: Vec<(WordSourcePath, Vec<usize>)>,
    ) -> Result<()> {
        let mut facts = std::collections::HashMap::new();
        for (path, runs) in exclusions {
            if facts.insert(path, runs).is_some() {
                return Err(LayoutError::Layout(
                    "duplicate navigation exclusion source".into(),
                ));
            }
        }
        engine::validate_navigation_exclusions(input, &facts)?;
        let mut owners = std::collections::HashSet::new();
        for path in facts.keys() {
            for end in 1..=path.children.len() {
                owners.insert(WordSourcePath {
                    story: path.story.clone(),
                    children: path.children[..end].to_vec(),
                });
            }
        }
        self.navigation_document = Some(input.document.clone());
        self.navigation_part = input.story_part_names.get(&WordStory::Document).cloned();
        self.navigation_failure = None;
        self.navigation_exclusions = facts;
        self.navigation_owners = owners;
        Ok(())
    }

    /// Retain ambiguous source discovery without changing ordinary cache paint.
    #[doc(hidden)]
    pub fn set_navigation_failure(&mut self, input: &LayoutInput, message: String) -> Result<()> {
        if message.trim().is_empty() {
            return Err(LayoutError::Layout("empty navigation failure fact".into()));
        }
        self.set_navigation_exclusions(input, Vec::new())?;
        self.navigation_failure = Some(message);
        Ok(())
    }

    /// Successful counter changes in physical accepted source order.
    pub fn main_events(&self) -> &[WordSequenceEvent] {
        &self.main_events
    }

    /// The shared physical path registered for one source node.
    pub fn source_node(&self, id: SourceNodeId) -> Option<&WordSourcePath> {
        self.source_nodes.get(id.get() as usize - 1)
    }

    /// The result-local source identity of one registered physical path.
    pub fn source_id(&self, path: &WordSourcePath) -> Option<SourceNodeId> {
        self.source_ids.get(path).copied()
    }

    /// Exact original parsed field XML in the current physical inventory.
    #[doc(hidden)]
    pub fn source_field_xml(&self, source: oxml_layout::FieldSource) -> Option<&[u8]> {
        self.field_source_xml.get(&source).map(Vec::as_slice)
    }

    /// Effective instruction, inherited lock and generated cache membership for a physical field.
    #[doc(hidden)]
    pub fn source_field_context(
        &self,
        source: oxml_layout::FieldSource,
    ) -> Option<(&str, bool, bool)> {
        self.field_source_contexts
            .get(&source)
            .map(|(instruction, locked, cached)| (instruction.as_str(), *locked, *cached))
    }

    /// Relative position within one qualified accepted physical bookmark owner.
    #[doc(hidden)]
    pub fn bookmark_relative_position(
        &self,
        target: &str,
        source: oxml_layout::FieldSource,
    ) -> std::result::Result<&'static str, String> {
        let (_, locked, cached) = self
            .source_field_context(source)
            .ok_or("REF relative source has no registered field context")?;
        if locked {
            return Err("locked field retains its stored display".into());
        }
        if cached {
            return Err("REF relative source is generated cache content".into());
        }
        let run = self
            .field_run_positions
            .get(&source)
            .ok_or("REF relative source lacks an accepted run boundary")?;
        let position = |node, run| {
            let path = self.source_node(node)?;
            let owner = self.physical_story_owners.get(&path.story)?;
            Some((owner, path.children.as_slice(), run))
        };
        let (owner, children, run) = position(source.node, *run)
            .ok_or("REF relative source physical owner is unavailable")?;
        let mut starts = Vec::new();
        for (id, name, node, run, order) in &self.position_bookmark_starts {
            if name != target {
                continue;
            }
            let Some((owner, children, run)) = position(*node, *run) else {
                return Err("REF relative target physical owner is unavailable".into());
            };
            let value = (*id, owner, children, run, *order);
            if !starts.contains(&value) {
                starts.push(value);
            }
        }
        let [(id, target_owner, start_children, start_run, start_order)] = starts.as_slice() else {
            return Err("REF relative target has no unique paired physical range".into());
        };
        let mut ends = Vec::new();
        for (end_id, node, run, order) in &self.position_bookmark_ends {
            if end_id != id {
                continue;
            }
            if let Some((end_owner, children, run)) = position(*node, *run)
                && end_owner == *target_owner
            {
                let value = (children, run, *order);
                if !ends.contains(&value) {
                    ends.push(value);
                }
            }
        }
        let [(end_children, end_run, end_order)] = ends.as_slice() else {
            return Err("REF relative target has no unique paired physical range".into());
        };
        let mut ids = Vec::new();
        for (other_id, name, node, run, order) in &self.position_bookmark_starts {
            if other_id != id {
                continue;
            }
            if let Some((other_owner, children, run)) = position(*node, *run)
                && other_owner == *target_owner
            {
                let value = (name, children, run, order);
                if !ids.contains(&value) {
                    ids.push(value);
                }
            }
        }
        let start = (*start_children, *start_run);
        let end = (*end_children, *end_run);
        if ids.len() != 1
            || start > end
            || (start_children == end_children && start_order > end_order)
        {
            return Err("REF relative target has no unique forward paired physical range".into());
        }
        if owner != *target_owner {
            return Err("REF relative position is unavailable across story owners".into());
        }
        let source = (children, run);
        if source < start {
            Ok("below")
        } else if source >= end {
            Ok("above")
        } else {
            Err("REF relative target contains the field".into())
        }
    }

    /// Physical selected text-box owner identity within its actual OPC part.
    #[doc(hidden)]
    pub fn text_box_owner_index(&self, story: &WordStory) -> Option<usize> {
        self.text_box_owner_indices.get(story).copied()
    }

    /// Resolved display or preservation diagnostic for a source-qualified SEQ.
    pub fn field_value(
        &self,
        source: oxml_layout::FieldSource,
    ) -> Option<std::result::Result<&str, &str>> {
        self.field_values
            .get(&source)
            .map(|value| value.as_deref().map_err(String::as_str))
    }

    /// Counter state immediately before an event boundary, without map replay.
    pub fn context_value(&self, before_event: usize, identifier: &str) -> Option<i64> {
        let indexes = self
            .identifier_events
            .get(&identifier.to_ascii_lowercase())?;
        let position = indexes.partition_point(|index| *index < before_event);
        position
            .checked_sub(1)
            .map(|previous| self.main_events[indexes[previous]].value)
    }
}

/// Effective accepted paragraph facts from the production style cascade.
#[derive(Debug, Clone)]
pub(crate) struct ParagraphNavigation {
    pub omitted_by_table: bool,
    pub level: Option<u32>,
    pub text: String,
    pub style: Option<String>,
    pub suppressed_numbering: bool,
    pub accepted_runs: usize,
}

/// Complete layout output plus its result-local Word source map.
#[derive(Debug)]
pub struct WordLayoutResult {
    pub layout: LayoutResult,
    pub revision_view: RevisionView,
    source_nodes: Vec<WordSourcePath>,
    body_fragments: Vec<Vec<WordBodyLayoutFragment>>,
    numbering_by_source: Vec<Option<style_resolver::ResolvedNumbering>>,
    navigation_failure: Option<String>,
    navigation_by_source: std::collections::BTreeMap<SourceNodeId, ParagraphNavigation>,
    paragraph_pages: std::collections::BTreeMap<SourceNodeId, usize>,
    bookmark_numbering: std::collections::BTreeMap<String, style_resolver::ResolvedNumbering>,
    bookmark_suppressed: std::collections::BTreeSet<String>,
    page_reference_names: Vec<String>,
    page_sections: Vec<WordPageSection>,
    field_placements: Vec<WordFieldPlacement>,
    bookmark_pages: std::collections::BTreeMap<usize, WordPageSection>,
    field_source_xml: std::collections::BTreeMap<oxml_layout::FieldSource, Vec<u8>>,
    text_box_owner_indices: std::collections::BTreeMap<WordStory, usize>,
}

impl WordLayoutResult {
    /// Physical and displayed section placement from the same pagination pass.
    pub fn page_sections(&self) -> &[WordPageSection] {
        &self.page_sections
    }

    /// Exact parsed text-box field source used to bind a physical story paragraph.
    #[doc(hidden)]
    pub fn source_field_xml(&self, source: oxml_layout::FieldSource) -> Option<&[u8]> {
        self.field_source_xml.get(&source).map(Vec::as_slice)
    }

    /// Physical selected text-box owner identity in its actual OPC part.
    #[doc(hidden)]
    pub fn text_box_owner_index(&self, story: &WordStory) -> Option<usize> {
        self.text_box_owner_indices.get(story).copied()
    }

    /// Physical paragraph preorder within one selected text-box owner.
    #[doc(hidden)]
    pub fn text_box_paragraph_index(&self, node: SourceNodeId) -> Option<usize> {
        let index = node.get() as usize - 1;
        let path = self.source_node(node)?;
        if !matches!(path.story, WordStory::TextBox { .. }) {
            return None;
        }
        Some(
            self.source_nodes[..index]
                .iter()
                .filter(|other| other.story == path.story)
                .count(),
        )
    }

    /// Every placed field occurrence, including shared story repetitions.
    pub fn field_placements(&self) -> &[WordFieldPlacement] {
        &self.field_placements
    }

    /// Displayed page of a unique, placed bookmark target.
    pub fn bookmark_page(&self, name: &str) -> Option<usize> {
        self.bookmark_page_section(name)
            .map(|record| record.displayed_page)
    }

    /// Physical page, displayed number and owning section of a unique placed target.
    pub fn bookmark_page_section(&self, name: &str) -> Option<WordPageSection> {
        let index = self
            .page_reference_names
            .iter()
            .position(|value| value == name)?;
        self.bookmark_pages.get(&index).copied()
    }

    /// Return placed fragments for one direct body item by its source index.
    ///
    /// Modeled items return one fragment per occupied page. Preserved content
    /// that does not enter layout returns an empty slice. An out-of-range body
    /// index returns `None`.
    pub fn body_layout_fragments(&self, body_index: usize) -> Option<&[WordBodyLayoutFragment]> {
        self.body_fragments.get(body_index).map(Vec::as_slice)
    }

    /// Resolve a result-local source identity.
    pub fn source_node(&self, id: SourceNodeId) -> Option<&WordSourcePath> {
        self.source_nodes.get(id.get() as usize - 1)
    }

    /// An exact-input source ambiguity that prevents qualified navigation.
    #[doc(hidden)]
    pub fn navigation_failure(&self) -> Option<&str> {
        self.navigation_failure.as_deref()
    }

    /// Effective accepted paragraph facts from this exact layout's shared cascade.
    /// Unplaced paragraphs remain present. IDs belong only to this result.
    #[doc(hidden)]
    #[allow(clippy::type_complexity)]
    pub fn paragraph_navigation(
        &self,
    ) -> impl Iterator<
        Item = (
            SourceNodeId,
            &WordSourcePath,
            Option<u32>,
            &str,
            Option<&str>,
            Option<&style_resolver::ResolvedNumbering>,
            Option<usize>,
        ),
    > {
        self.source_nodes
            .iter()
            .enumerate()
            .filter_map(|(index, path)| {
                let node = SourceNodeId::new(u32::try_from(index + 1).ok()?)?;
                let facts = self.navigation_by_source.get(&node)?;
                Some((
                    node,
                    path,
                    facts.level,
                    facts.text.as_str(),
                    facts.style.as_deref(),
                    self.paragraph_numbering(node),
                    self.paragraph_pages.get(&node).copied(),
                ))
            })
    }

    /// Look up one source in the same private navigation facts without a scan.
    #[doc(hidden)]
    #[allow(clippy::type_complexity)]
    pub fn paragraph_navigation_source(
        &self,
        node: SourceNodeId,
    ) -> Option<(
        &WordSourcePath,
        Option<u32>,
        &str,
        Option<&str>,
        Option<&style_resolver::ResolvedNumbering>,
        Option<usize>,
    )> {
        let facts = self.navigation_by_source.get(&node)?;
        Some((
            self.source_node(node)?,
            facts.level,
            facts.text.as_str(),
            facts.style.as_deref(),
            self.paragraph_numbering(node),
            self.paragraph_pages.get(&node).copied(),
        ))
    }

    /// Whether this result's actual table layout omitted a registered paragraph.
    /// `None` means that this source has no qualified navigation facts.
    #[doc(hidden)]
    pub fn paragraph_omitted_by_table(&self, node: SourceNodeId) -> Option<bool> {
        self.source_node(node)?;
        self.navigation_by_source
            .get(&node)
            .map(|facts| facts.omitted_by_table)
    }

    /// Whether the effective paragraph explicitly suppresses its numbering.
    #[doc(hidden)]
    pub fn paragraph_numbering_suppressed(&self, source: SourceNodeId) -> bool {
        self.navigation_by_source
            .get(&source)
            .is_some_and(|facts| facts.suppressed_numbering)
    }

    /// Resolve the visible numbering marker assigned to a source paragraph.
    pub fn paragraph_numbering(
        &self,
        id: SourceNodeId,
    ) -> Option<&style_resolver::ResolvedNumbering> {
        self.numbering_by_source
            .get(id.get() as usize - 1)
            .and_then(Option::as_ref)
    }

    /// Visible numbering for one uniquely paired accepted physical bookmark owner.
    #[doc(hidden)]
    pub fn bookmark_numbering(&self, name: &str) -> Option<&style_resolver::ResolvedNumbering> {
        self.bookmark_numbering.get(name)
    }

    /// Explicit numbering suppression at a uniquely paired physical bookmark.
    #[doc(hidden)]
    pub fn bookmark_numbering_suppressed(&self, name: &str) -> bool {
        self.bookmark_suppressed.contains(name)
    }

    /// Resolve numbering by flattened main-story paragraph order.
    #[doc(hidden)]
    pub fn document_paragraph_numbering(
        &self,
        paragraph_index: usize,
    ) -> Option<&style_resolver::ResolvedNumbering> {
        let source_index = self
            .source_nodes
            .iter()
            .enumerate()
            .filter(|(_, path)| path.story == WordStory::Document)
            .nth(paragraph_index)
            .map(|(index, _)| index)?;
        self.numbering_by_source
            .get(source_index)
            .and_then(Option::as_ref)
    }

    /// Resolve numbering for a top-level body paragraph by its body index.
    #[doc(hidden)]
    pub fn document_body_paragraph_numbering(
        &self,
        body_index: usize,
    ) -> Option<&style_resolver::ResolvedNumbering> {
        let source_index = self.source_nodes.iter().position(|path| {
            path.story == WordStory::Document && path.children.as_slice() == [body_index]
        })?;
        self.numbering_by_source
            .get(source_index)
            .and_then(Option::as_ref)
    }

    /// Discard the Word source map and return the backend-neutral layout.
    pub fn into_layout_result(self) -> LayoutResult {
        self.layout
    }

    /// Resolve the bookmark name assigned to a result-local page target.
    #[doc(hidden)]
    pub fn page_reference_name(&self, target: usize) -> Option<&str> {
        self.page_reference_names.get(target).map(String::as_str)
    }
}

fn ordered<K: Ord, V>(map: std::collections::HashMap<K, V>) -> std::collections::BTreeMap<K, V> {
    map.into_iter().collect()
}

fn ordered_set<T: Ord>(set: std::collections::HashSet<T>) -> std::collections::BTreeSet<T> {
    set.into_iter().collect()
}

fn paragraph_pages(layout: &LayoutResult) -> std::collections::BTreeMap<SourceNodeId, usize> {
    let mut pages = std::collections::BTreeMap::new();
    for page in &layout.pages {
        oxml_layout::walk(&page.elements, &mut |element, _| {
            let source = match element {
                PositionedElement::Text(run) => run.source,
                PositionedElement::MultilingualText(run) => run.source,
                _ => None,
            };
            if let Some(source) = source {
                pages.entry(source.node).or_insert(page.page_number);
            }
        });
    }
    pages
}

/// Lay out a complete DOCX document, producing positioned page frames.
pub fn layout_document(input: &LayoutInput) -> Result<LayoutResult> {
    engine::Engine::new().layout(input)
}

/// Lay out a complete DOCX and retain exact Word paragraph provenance.
pub fn layout_document_with_provenance(input: &LayoutInput) -> Result<WordLayoutResult> {
    let mut engine = engine::Engine::new();
    let (layout, source_nodes) = engine.layout_with_provenance(input)?;
    let body_fragments = engine.take_body_fragments();
    let (page_sections, field_placements, bookmark_pages) = engine.take_field_snapshot();
    let numbering_by_source = engine.numbering_by_source(source_nodes.len());
    let paragraph_pages = paragraph_pages(&layout);
    Ok(WordLayoutResult {
        paragraph_pages,
        layout,
        revision_view: input.revision_view,
        source_nodes,
        body_fragments,
        page_sections,
        field_placements,
        bookmark_pages: ordered(bookmark_pages),
        field_source_xml: ordered(engine.take_field_source_xml()),
        text_box_owner_indices: ordered(engine.take_text_box_owner_indices()),
        numbering_by_source,
        navigation_by_source: ordered(engine.take_navigation()),
        navigation_failure: input
            .sequence_snapshot
            .as_ref()
            .filter(|_| input.revision_view == RevisionView::Accepted)
            .and_then(|snapshot| snapshot.navigation_failure.clone()),
        bookmark_numbering: ordered(engine.take_bookmark_numbering()),
        bookmark_suppressed: ordered_set(engine.take_bookmark_suppressed()),
        page_reference_names: engine::page_reference_names(input),
    })
}

/// Lay out a DOCX with a reusable normal-font engine.
///
/// This hidden facade hook lets `rdocx::Document` retain expensive normal-font
/// work without exposing cache ownership as a second public abstraction.
#[doc(hidden)]
pub fn layout_document_with_reusable_engine(
    engine: &mut engine::Engine,
    input: &LayoutInput,
) -> Result<WordLayoutResult> {
    let (layout, source_nodes) = engine.layout_with_provenance(input)?;
    let body_fragments = engine.take_body_fragments();
    let (page_sections, field_placements, bookmark_pages) = engine.take_field_snapshot();
    let numbering_by_source = engine.numbering_by_source(source_nodes.len());
    let paragraph_pages = paragraph_pages(&layout);
    Ok(WordLayoutResult {
        paragraph_pages,
        layout,
        revision_view: input.revision_view,
        source_nodes,
        body_fragments,
        page_sections,
        field_placements,
        bookmark_pages: ordered(bookmark_pages),
        field_source_xml: ordered(engine.take_field_source_xml()),
        text_box_owner_indices: ordered(engine.take_text_box_owner_indices()),
        numbering_by_source,
        navigation_by_source: ordered(engine.take_navigation()),
        navigation_failure: input
            .sequence_snapshot
            .as_ref()
            .filter(|_| input.revision_view == RevisionView::Accepted)
            .and_then(|snapshot| snapshot.navigation_failure.clone()),
        bookmark_numbering: ordered(engine.take_bookmark_numbering()),
        bookmark_suppressed: ordered_set(engine.take_bookmark_suppressed()),
        page_reference_names: engine::page_reference_names(input),
    })
}

/// Lay out a DOCX using only caller-supplied and document-embedded fonts.
#[doc(hidden)]
pub fn layout_document_with_caller_fonts_and_provenance(
    input: &LayoutInput,
) -> Result<WordLayoutResult> {
    let mut engine = engine::Engine::new_with_caller_fonts();
    let (layout, source_nodes) = engine.layout_with_provenance(input)?;
    let body_fragments = engine.take_body_fragments();
    let (page_sections, field_placements, bookmark_pages) = engine.take_field_snapshot();
    let numbering_by_source = engine.numbering_by_source(source_nodes.len());
    let paragraph_pages = paragraph_pages(&layout);
    Ok(WordLayoutResult {
        paragraph_pages,
        layout,
        revision_view: input.revision_view,
        source_nodes,
        body_fragments,
        page_sections,
        field_placements,
        bookmark_pages: ordered(bookmark_pages),
        field_source_xml: ordered(engine.take_field_source_xml()),
        text_box_owner_indices: ordered(engine.take_text_box_owner_indices()),
        numbering_by_source,
        navigation_by_source: ordered(engine.take_navigation()),
        navigation_failure: input
            .sequence_snapshot
            .as_ref()
            .filter(|_| input.revision_view == RevisionView::Accepted)
            .and_then(|snapshot| snapshot.navigation_failure.clone()),
        bookmark_numbering: ordered(engine.take_bookmark_numbering()),
        bookmark_suppressed: ordered_set(engine.take_bookmark_suppressed()),
        page_reference_names: engine::page_reference_names(input),
    })
}

/// Lay out a DOCX using bundled fonts without system font discovery.
pub fn layout_document_deterministic(input: &LayoutInput) -> Result<LayoutResult> {
    engine::Engine::new_deterministic()?.layout(input)
}

/// Measure one supported Word block with the deterministic production engine.
///
/// This hidden facade hook keeps caller-width measurement on the same font,
/// paragraph, table, media, and numbering path as whole-document layout.
#[doc(hidden)]
pub fn measure_content_deterministic(
    input: &LayoutInput,
    content: &BodyContent,
    available_width: f64,
    related_story_scope: Option<&str>,
) -> Result<(f64, Vec<Diagnostic>)> {
    engine::Engine::new_deterministic()?.measure_content(
        input,
        content,
        available_width,
        related_story_scope,
    )
}

/// Lay out a DOCX deterministically and retain exact Word paragraph provenance.
pub fn layout_document_deterministic_with_provenance(
    input: &LayoutInput,
) -> Result<WordLayoutResult> {
    let mut engine = engine::Engine::new_deterministic()?;
    let (layout, source_nodes) = engine.layout_with_provenance(input)?;
    let body_fragments = engine.take_body_fragments();
    let (page_sections, field_placements, bookmark_pages) = engine.take_field_snapshot();
    let numbering_by_source = engine.numbering_by_source(source_nodes.len());
    let paragraph_pages = paragraph_pages(&layout);
    Ok(WordLayoutResult {
        paragraph_pages,
        layout,
        revision_view: input.revision_view,
        source_nodes,
        body_fragments,
        page_sections,
        field_placements,
        bookmark_pages: ordered(bookmark_pages),
        field_source_xml: ordered(engine.take_field_source_xml()),
        text_box_owner_indices: ordered(engine.take_text_box_owner_indices()),
        numbering_by_source,
        navigation_by_source: ordered(engine.take_navigation()),
        navigation_failure: input
            .sequence_snapshot
            .as_ref()
            .filter(|_| input.revision_view == RevisionView::Accepted)
            .and_then(|snapshot| snapshot.navigation_failure.clone()),
        bookmark_numbering: ordered(engine.take_bookmark_numbering()),
        bookmark_suppressed: ordered_set(engine.take_bookmark_suppressed()),
        page_reference_names: engine::page_reference_names(input),
    })
}
