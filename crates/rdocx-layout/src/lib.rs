#![doc = include_str!("../README.md")]
#![allow(non_camel_case_types)]
#![allow(clippy::too_many_arguments)]

pub mod block;
mod convert;
pub mod engine;
pub mod input;
pub mod notes;
pub mod paginator;
pub mod style_resolver;
pub mod table;

pub use input::{ImageData, LayoutInput, MediaRegistry, RevisionView};
pub use oxml_layout::{
    Color, DocumentMetadata, FontData, FontFile, FontId, GlyphRun, LayoutError, LayoutResult,
    PageFrame, Point, PositionedElement, Rect, Result, SourceNodeId, SourceSpan,
};

/// Word story containing a source paragraph.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum WordStory {
    Document,
    Header { relationship_id: String },
    Footer { relationship_id: String },
    Footnote { id: i32 },
    Endnote { id: i32 },
}

/// Modeled path to one paragraph in a Word story.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct WordSourcePath {
    pub story: WordStory,
    pub children: Vec<usize>,
}

/// Complete layout output plus its result-local Word source map.
#[derive(Debug)]
pub struct WordLayoutResult {
    pub layout: LayoutResult,
    pub revision_view: RevisionView,
    source_nodes: Vec<WordSourcePath>,
}

impl WordLayoutResult {
    /// Resolve a result-local source identity.
    pub fn source_node(&self, id: SourceNodeId) -> Option<&WordSourcePath> {
        self.source_nodes.get(id.get() as usize - 1)
    }

    /// Discard the Word source map and return the backend-neutral layout.
    pub fn into_layout_result(self) -> LayoutResult {
        self.layout
    }
}

/// Lay out a complete DOCX document, producing positioned page frames.
pub fn layout_document(input: &LayoutInput) -> Result<LayoutResult> {
    engine::Engine::new().layout(input)
}

/// Lay out a complete DOCX and retain exact Word paragraph provenance.
pub fn layout_document_with_provenance(input: &LayoutInput) -> Result<WordLayoutResult> {
    let (layout, source_nodes) = engine::Engine::new().layout_with_provenance(input)?;
    Ok(WordLayoutResult {
        layout,
        revision_view: input.revision_view,
        source_nodes,
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
    Ok(WordLayoutResult {
        layout,
        revision_view: input.revision_view,
        source_nodes,
    })
}

/// Lay out a DOCX using only caller-supplied and document-embedded fonts.
#[doc(hidden)]
pub fn layout_document_with_caller_fonts_and_provenance(
    input: &LayoutInput,
) -> Result<WordLayoutResult> {
    let (layout, source_nodes) =
        engine::Engine::new_with_caller_fonts().layout_with_provenance(input)?;
    Ok(WordLayoutResult {
        layout,
        revision_view: input.revision_view,
        source_nodes,
    })
}

/// SVG PoC patch: lay out with caller fonts on top of the bundled
/// metric-compatible set — for wasm editors that inject one or two fonts
/// and still need Calibri/Times to resolve. A separate entry point so the
/// strict caller-fonts isolation of
/// [`layout_document_with_caller_fonts_and_provenance`] stays intact.
pub fn layout_document_with_fallback_fonts_and_provenance(
    input: &LayoutInput,
) -> Result<WordLayoutResult> {
    let mut engine = engine::Engine::new_deterministic()?;
    layout_document_with_fallback_fonts_engine(&mut engine, input)
}

/// SVG PoC patch: the reusable-engine form of
/// [`layout_document_with_fallback_fonts_and_provenance`], so an editor can
/// keep the engine's relayout caches warm across edits (mirrors the hidden
/// `layout_document_with_reusable_engine` facade).
pub fn layout_document_with_fallback_fonts_engine(
    engine: &mut engine::Engine,
    input: &LayoutInput,
) -> Result<WordLayoutResult> {
    let (layout, source_nodes) = engine.layout_with_provenance(input)?;
    Ok(WordLayoutResult {
        layout,
        revision_view: input.revision_view,
        source_nodes,
    })
}

/// Lay out a DOCX using bundled fonts without system font discovery.
pub fn layout_document_deterministic(input: &LayoutInput) -> Result<LayoutResult> {
    engine::Engine::new_deterministic()?.layout(input)
}

/// Lay out a DOCX deterministically and retain exact Word paragraph provenance.
pub fn layout_document_deterministic_with_provenance(
    input: &LayoutInput,
) -> Result<WordLayoutResult> {
    let (layout, source_nodes) =
        engine::Engine::new_deterministic()?.layout_with_provenance(input)?;
    Ok(WordLayoutResult {
        layout,
        revision_view: input.revision_view,
        source_nodes,
    })
}
