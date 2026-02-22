//! Input types for the layout engine.

use std::collections::HashMap;

use rdocx_oxml::core_properties::CoreProperties;
use rdocx_oxml::document::CT_Document;
use rdocx_oxml::footnotes::CT_Footnotes;
use rdocx_oxml::header_footer::CT_HdrFtr;
use rdocx_oxml::numbering::CT_Numbering;
use rdocx_oxml::styles::CT_Styles;
use rdocx_oxml::theme::Theme;

/// Image data keyed by relationship/embed ID.
#[derive(Debug, Clone)]
pub struct ImageData {
    /// Raw image bytes (PNG, JPEG, etc.).
    pub data: Vec<u8>,
    /// MIME content type (e.g., "image/png").
    pub content_type: String,
}

/// Font data provided by the user or extracted from a DOCX file.
#[derive(Debug, Clone)]
pub struct FontFile {
    /// Font family name (e.g., "Calibri", "Arial").
    pub family: String,
    /// Raw font file bytes (TTF/OTF).
    pub data: Vec<u8>,
}

/// All inputs needed to lay out a DOCX document.
#[derive(Debug, Clone)]
pub struct LayoutInput {
    /// The parsed document content.
    pub document: CT_Document,
    /// Style definitions.
    pub styles: CT_Styles,
    /// Numbering definitions (optional).
    pub numbering: Option<CT_Numbering>,
    /// Header parts keyed by relationship ID.
    pub headers: HashMap<String, CT_HdrFtr>,
    /// Footer parts keyed by relationship ID.
    pub footers: HashMap<String, CT_HdrFtr>,
    /// Images keyed by embed ID.
    pub images: HashMap<String, ImageData>,
    /// Document core properties (metadata).
    pub core_properties: Option<CoreProperties>,
    /// Hyperlink URLs keyed by relationship ID.
    pub hyperlink_urls: HashMap<String, String>,
    /// Footnote definitions.
    pub footnotes: Option<CT_Footnotes>,
    /// Endnote definitions.
    pub endnotes: Option<CT_Footnotes>,
    /// Document theme (colors + fonts).
    pub theme: Option<Theme>,
    /// User-provided or DOCX-embedded font files.
    /// These are loaded before system fonts, so they take priority.
    pub fonts: Vec<FontFile>,
}
