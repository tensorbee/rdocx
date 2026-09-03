//! Input types for the layout engine.

use std::collections::HashMap;

use oxml_chart::CT_ChartSpace;
use oxml_drawing::color::ColorMap;
use oxml_drawing::theme::CT_OfficeStyleSheet;
pub use oxml_layout::FontFile;
use oxml_layout::MediaId;
use rdocx_oxml::core_properties::CoreProperties;
use rdocx_oxml::document::CT_Document;
use rdocx_oxml::footnotes::CT_Footnotes;
use rdocx_oxml::header_footer::CT_HdrFtr;
use rdocx_oxml::math::MathProperties;
use rdocx_oxml::numbering::CT_Numbering;
use rdocx_oxml::styles::CT_Styles;
use rdocx_oxml::theme::Theme;

/// The tracked-revision projection used for Word layout.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum RevisionView {
    /// Render the document as though all modeled revisions were accepted.
    #[default]
    Accepted,
    /// Render both sides of modeled revisions with tracked decorations.
    Tracked,
}

/// Image data keyed by relationship/embed ID.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ImageData {
    /// Raw image bytes (PNG, JPEG, etc.).
    pub data: Vec<u8>,
    /// MIME content type (e.g., "image/png").
    pub content_type: String,
}

/// Collision-safe media lookup shared by layout and pagination.
#[derive(Debug, Clone)]
pub struct MediaRegistry {
    relationship_ids: HashMap<String, MediaId>,
    media: HashMap<MediaId, ImageData>,
    missing_id: MediaId,
}

impl MediaRegistry {
    /// Resolve relationship IDs and image bytes once for a layout operation.
    pub fn new(images: &HashMap<String, ImageData>) -> Self {
        Self::with_hasher(images, MediaId::from_bytes)
    }

    /// Resolve the renderer-local ID for one package relationship.
    pub fn id_for_relationship(&self, relationship_id: &str) -> MediaId {
        self.relationship_ids
            .get(relationship_id)
            .copied()
            .unwrap_or(self.missing_id)
    }

    /// Return the image bytes and content types keyed by resolved media ID.
    pub fn media(&self) -> &HashMap<MediaId, ImageData> {
        &self.media
    }

    pub(crate) fn with_hasher<F>(images: &HashMap<String, ImageData>, media_id_for_bytes: F) -> Self
    where
        F: Fn(&[u8]) -> MediaId,
    {
        let missing_id = media_id_for_bytes(&[]);
        let mut media = HashMap::from([(
            missing_id,
            ImageData {
                data: Vec::new(),
                content_type: String::new(),
            },
        )]);
        let mut relationship_ids = HashMap::new();
        let mut images = images.iter().collect::<Vec<_>>();
        images.sort_unstable_by(|(left_id, left), (right_id, right)| {
            left.data
                .cmp(&right.data)
                .then_with(|| left.content_type.cmp(&right.content_type))
                .then_with(|| left_id.cmp(right_id))
        });

        for (relationship_id, image) in images {
            let mut media_id = media_id_for_bytes(&image.data);
            loop {
                match media.get(&media_id) {
                    Some(existing) if existing.data == image.data => break,
                    Some(_) => media_id.0 = media_id.0.wrapping_add(1),
                    None => {
                        media.insert(media_id, image.clone());
                        break;
                    }
                }
            }
            relationship_ids.insert(relationship_id.clone(), media_id);
        }

        Self {
            relationship_ids,
            media,
            missing_id,
        }
    }
}

/// All inputs needed to lay out a DOCX document.
#[derive(Debug, Clone)]
pub struct LayoutInput {
    /// The parsed document content.
    pub document: CT_Document,
    /// Whether document settings enable automatic hyphenation.
    pub automatic_hyphenation: bool,
    /// Document-wide OfficeMath defaults from the settings part.
    pub math_properties: Option<MathProperties>,
    /// The tracked-revision projection to lay out.
    pub revision_view: RevisionView,
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
    /// Parsed chart parts, or contextual relationship failures, keyed by ID.
    pub charts: HashMap<String, std::result::Result<Box<CT_ChartSpace>, String>>,
    /// DrawingML theme used by the shared chart renderer.
    pub chart_theme: CT_OfficeStyleSheet,
    /// Standard Word chart colour mapping.
    pub chart_color_map: ColorMap,
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
