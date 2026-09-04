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
    page_reference_names: Vec<String>,
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

    /// Resolve the bookmark name assigned to a result-local page target.
    #[doc(hidden)]
    pub fn page_reference_name(&self, target: usize) -> Option<&str> {
        self.page_reference_names.get(target).map(String::as_str)
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
    Ok(WordLayoutResult {
        layout,
        revision_view: input.revision_view,
        source_nodes,
        page_reference_names: engine::page_reference_names(input),
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
        page_reference_names: engine::page_reference_names(input),
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
        page_reference_names: engine::page_reference_names(input),
    })
}
