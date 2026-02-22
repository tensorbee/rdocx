//! Layout engine for converting DOCX flow model to positioned page frames.
//!
//! This crate implements style resolution, font loading, text shaping,
//! line breaking, block/table layout, and pagination to produce
//! [`LayoutResult`] containing [`PageFrame`]s with absolutely-positioned elements.

#![allow(non_camel_case_types)]
#![allow(clippy::too_many_arguments)]

pub mod block;
pub mod bundled_fonts;
pub mod engine;
pub mod error;
pub mod font;
pub mod input;
pub mod line;
pub mod output;
pub mod paginator;
pub mod style_resolver;
pub mod table;

pub use error::{LayoutError, Result};
pub use input::{FontFile, ImageData, LayoutInput};
pub use output::{
    Color, DocumentMetadata, FontData, FontId, GlyphRun, LayoutResult, PageFrame, Point,
    PositionedElement, Rect,
};

/// Lay out a complete DOCX document, producing positioned page frames.
pub fn layout_document(input: &LayoutInput) -> Result<LayoutResult> {
    engine::Engine::new().layout(input)
}
