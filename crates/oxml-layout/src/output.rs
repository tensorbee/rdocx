//! Output types for the layout engine: positioned page frames, glyph runs, etc.

use std::num::NonZeroU32;

use crate::paint::{Paint, Stroke};
use crate::path::Path;
use crate::transform::Transform;

/// A point in 2D space (in typographic points from the top-left corner).
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Point {
    pub x: f64,
    pub y: f64,
}

/// An axis-aligned rectangle.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Rect {
    pub x: f64,
    pub y: f64,
    pub width: f64,
    pub height: f64,
}

/// An RGBA color with components in [0.0, 1.0].
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Color {
    pub r: f64,
    pub g: f64,
    pub b: f64,
    pub a: f64,
}

impl Color {
    pub const BLACK: Color = Color {
        r: 0.0,
        g: 0.0,
        b: 0.0,
        a: 1.0,
    };
    pub const WHITE: Color = Color {
        r: 1.0,
        g: 1.0,
        b: 1.0,
        a: 1.0,
    };

    /// Parse a hex color string like "FF0000" to Color.
    pub fn from_hex(hex: &str) -> Self {
        let hex = hex.trim_start_matches('#');
        if hex.len() >= 6 {
            let r = u8::from_str_radix(&hex[0..2], 16).unwrap_or(0) as f64 / 255.0;
            let g = u8::from_str_radix(&hex[2..4], 16).unwrap_or(0) as f64 / 255.0;
            let b = u8::from_str_radix(&hex[4..6], 16).unwrap_or(0) as f64 / 255.0;
            Color { r, g, b, a: 1.0 }
        } else {
            Color::BLACK
        }
    }
}

/// Opaque font identifier assigned by FontManager.
///
/// Ordered, so a backend keying a map on it can iterate in a fixed order rather
/// than a hashed one. The PDF writer does exactly that, and the order it
/// iterates in reaches the bytes it writes.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct FontId(pub u32);

/// Stable content-addressed media key for renderer-local reuse.
///
/// This compact key is not a collision-free content guarantee.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct MediaId(pub u64);

impl MediaId {
    /// Derive a stable key from raw media bytes using 64-bit FNV-1a.
    pub fn from_bytes(bytes: &[u8]) -> Self {
        let mut hash = 0xcbf2_9ce4_8422_2325_u64;
        for byte in bytes {
            hash ^= u64::from(*byte);
            hash = hash.wrapping_mul(0x0000_0100_0000_01b3);
        }
        Self(hash)
    }
}

/// Result-local identity of one format-specific source node.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct SourceNodeId(NonZeroU32);

impl SourceNodeId {
    /// Construct an identity from its one-based side-table index.
    pub const fn new(value: u32) -> Option<Self> {
        match NonZeroU32::new(value) {
            Some(value) => Some(Self(value)),
            None => None,
        }
    }

    /// Return the one-based side-table index.
    pub const fn get(self) -> u32 {
        self.0.get()
    }
}

/// Exclusive Unicode-scalar range within one source node.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SourceSpan {
    pub node: SourceNodeId,
    pub char_start: u32,
    pub char_end: u32,
}

/// Kind of field for post-pagination substitution.
///
/// Target carriers stay format-neutral at this shared layout boundary.
///
/// ```
/// use oxml_layout::FieldKind;
///
/// assert_eq!(FieldKind::TargetPage(3), FieldKind::TargetPage(3));
/// assert_eq!(FieldKind::Target(3), FieldKind::Target(3));
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FieldKind {
    /// Current page number.
    Page,
    /// Total number of pages.
    NumPages,
    /// Page containing a target.
    TargetPage(usize),
    /// Zero-width target position retained until page locations are collected.
    Target(usize),
}

/// A positioned run of shaped glyphs.
#[derive(Debug, Clone, PartialEq)]
pub struct GlyphRun {
    /// Baseline origin of the first glyph (in points).
    pub origin: Point,
    /// Font identifier (from FontManager).
    pub font_id: FontId,
    /// Font size in points.
    pub font_size: f64,
    /// Shaped glyph IDs.
    pub glyph_ids: Vec<u16>,
    /// Per-glyph advances in points.
    pub advances: Vec<f64>,
    /// Original text (for PDF ToUnicode mapping).
    pub text: String,
    /// Exact source range for this run, when it is a direct text projection.
    pub source: Option<SourceSpan>,
    /// Text color.
    pub color: Color,
    /// Whether the font is bold.
    pub bold: bool,
    /// Whether the font is italic.
    pub italic: bool,
    /// If this glyph run is a field placeholder, the kind of field.
    pub field_kind: Option<FieldKind>,
    /// If this glyph run is a footnote/endnote reference marker, its ID.
    pub note: Option<crate::line::NoteRef>,
}

/// A positioned element on a page.
#[derive(Debug, Clone, PartialEq)]
#[non_exhaustive]
pub enum PositionedElement {
    /// A run of shaped text glyphs.
    Text(GlyphRun),
    /// A line segment (for borders, underlines, strikethrough).
    Line {
        start: Point,
        end: Point,
        width: f64,
        color: Color,
        /// Optional dash pattern (dash_on, dash_off) in points. None = solid line.
        dash_pattern: Option<(f64, f64)>,
    },
    /// A filled rectangle (for shading, highlights).
    FilledRect { rect: Rect, color: Color },
    /// An inline image.
    Image {
        rect: Rect,
        data: Vec<u8>,
        content_type: String,
        media_id: MediaId,
    },
    /// A link annotation (hyperlink).
    LinkAnnotation { rect: Rect, url: String },
    /// A backend-neutral filled or stroked path.
    Path(PathElement),
    /// A nested group with one child-local transform.
    Group(GroupElement),
}

/// One path with optional fill and stroke paints.
#[derive(Debug, Clone, PartialEq)]
pub struct PathElement {
    pub path: Path,
    pub fill: Option<Paint>,
    pub stroke: Option<Stroke>,
}

/// A rendering approximation or fallback message.
#[derive(Debug, Clone, PartialEq)]
pub struct Diagnostic {
    pub message: String,
}

/// An effect applied to a group.
#[derive(Debug, Clone, PartialEq)]
#[non_exhaustive]
pub enum Effect {
    OuterShadow {
        dx: f64,
        dy: f64,
        blur: f64,
        color: Color,
    },
}

/// A group of positioned children in one local coordinate system.
#[derive(Debug, Clone, PartialEq)]
pub struct GroupElement {
    /// Maps child-local coordinates into the parent coordinate system.
    pub transform: Transform,
    pub clip: Option<Path>,
    pub opacity: f64,
    pub effects: Vec<Effect>,
    pub children: Vec<PositionedElement>,
}

/// Visit every non-group element in depth-first document order.
pub fn walk(elements: &[PositionedElement], f: &mut impl FnMut(&PositionedElement, &Transform)) {
    fn visit(
        elements: &[PositionedElement],
        accumulated: Transform,
        f: &mut dyn FnMut(&PositionedElement, &Transform),
    ) {
        for element in elements {
            match element {
                PositionedElement::Group(group) => {
                    let child_to_page = group.transform.then(accumulated);
                    visit(&group.children, child_to_page, f);
                }
                leaf => f(leaf, &accumulated),
            }
        }
    }

    visit(elements, Transform::IDENTITY, f);
}

/// A single page of laid-out content.
#[derive(Debug, Clone)]
#[non_exhaustive]
pub struct PageFrame {
    /// 1-based page number.
    pub page_number: usize,
    /// Page width in points.
    pub width: f64,
    /// Page height in points.
    pub height: f64,
    /// All positioned elements on this page.
    pub elements: Vec<PositionedElement>,
    /// Optional paint behind every page element.
    pub background: Option<Paint>,
}

impl PageFrame {
    /// Construct a page with no background paint.
    ///
    /// ```
    /// use oxml_layout::PageFrame;
    ///
    /// let page = PageFrame::new(1, 612.0, 792.0, Vec::new());
    /// assert!(page.background.is_none());
    /// ```
    pub fn new(
        page_number: usize,
        width: f64,
        height: f64,
        elements: Vec<PositionedElement>,
    ) -> Self {
        Self {
            page_number,
            width,
            height,
            elements,
            background: None,
        }
    }
}

/// Font data for embedding in PDF output.
#[derive(Debug, Clone)]
pub struct FontData {
    /// Font identifier.
    pub id: FontId,
    /// Font family name.
    pub family: String,
    /// Raw TTF/OTF bytes for PDF embedding. Shared, so producing a
    /// `LayoutResult` does not copy every loaded face (tens of MB with CJK
    /// fallbacks resident) on each layout. Pre-1.0 type break.
    pub data: std::sync::Arc<[u8]>,
    /// Face index within a font collection.
    pub face_index: u32,
    /// Whether this is a bold variant.
    pub bold: bool,
    /// Whether this is an italic variant.
    pub italic: bool,
}

/// Document metadata to pass through to PDF output.
#[derive(Debug, Clone, Default)]
pub struct DocumentMetadata {
    /// Document title.
    pub title: Option<String>,
    /// Document author.
    pub author: Option<String>,
    /// Document subject.
    pub subject: Option<String>,
    /// Document keywords.
    pub keywords: Option<String>,
    /// Creator application.
    pub creator: Option<String>,
}

/// An outline/bookmark entry for PDF generation.
#[derive(Debug, Clone)]
pub struct OutlineEntry {
    /// The heading text.
    pub title: String,
    /// Heading level (1 for Heading1, 2 for Heading2, etc.).
    pub level: u32,
    /// 0-based page index this heading appears on.
    pub page_index: usize,
    /// Y position on the page (in points from top).
    pub y_position: f64,
}

/// The complete result of laying out a document.
#[derive(Debug, Clone)]
#[non_exhaustive]
pub struct LayoutResult {
    /// Laid-out pages, shared: an interactive caller relayouting per edit
    /// keeps unchanged pages alive across results instead of deep-copying
    /// them. Pre-1.0 type break, like `FontData.data`.
    pub pages: Vec<std::sync::Arc<PageFrame>>,
    /// Font data for all fonts used.
    pub fonts: Vec<FontData>,
    /// Optional document metadata for PDF output.
    pub metadata: Option<DocumentMetadata>,
    /// Outline/bookmark entries from headings.
    pub outlines: Vec<OutlineEntry>,
    /// Rendering approximations and fallbacks collected during layout.
    pub diagnostics: Vec<Diagnostic>,
}

impl LayoutResult {
    /// Construct a result with no diagnostics.
    ///
    /// ```
    /// use oxml_layout::LayoutResult;
    ///
    /// let result = LayoutResult::new(Vec::new(), Vec::new(), None, Vec::new());
    /// assert!(result.diagnostics.is_empty());
    /// ```
    pub fn new(
        pages: Vec<PageFrame>,
        fonts: Vec<FontData>,
        metadata: Option<DocumentMetadata>,
        outlines: Vec<OutlineEntry>,
    ) -> Self {
        Self::from_shared(
            pages.into_iter().map(std::sync::Arc::new).collect(),
            fonts,
            metadata,
            outlines,
        )
    }

    /// Construct a result from pages that are already shared.
    pub fn from_shared(
        pages: Vec<std::sync::Arc<PageFrame>>,
        fonts: Vec<FontData>,
        metadata: Option<DocumentMetadata>,
        outlines: Vec<OutlineEntry>,
    ) -> Self {
        Self {
            pages,
            fonts,
            metadata,
            outlines,
            diagnostics: Vec::new(),
        }
    }
}

#[cfg(test)]
mod media_id_tests {
    use std::collections::HashSet;

    use super::{MediaId, PositionedElement, Rect};

    #[test]
    fn the_same_image_bytes_inserted_twice_produce_one_media_id() {
        let ids = HashSet::from([
            MediaId::from_bytes(b"same image"),
            MediaId::from_bytes(b"same image"),
        ]);
        assert_eq!(ids.len(), 1);
    }

    #[test]
    fn media_id_depends_on_bytes_not_relationship_context() {
        assert_eq!(
            MediaId::from_bytes(b"image bytes"),
            MediaId::from_bytes(b"image bytes")
        );
    }

    #[test]
    fn different_image_bytes_have_different_fixture_ids() {
        assert_ne!(
            MediaId::from_bytes(b"first image"),
            MediaId::from_bytes(b"second image")
        );
    }

    #[test]
    fn staged_output_image_uses_media_id_instead_of_embed_id() {
        let media_id = MediaId::from_bytes(b"image bytes");
        let image = PositionedElement::Image {
            rect: Rect {
                x: 0.0,
                y: 0.0,
                width: 10.0,
                height: 20.0,
            },
            data: b"image bytes".to_vec(),
            content_type: "image/png".to_owned(),
            media_id,
        };
        let PositionedElement::Image {
            media_id: actual, ..
        } = image
        else {
            panic!("constructed image should remain an image");
        };
        assert_eq!(actual, media_id);
    }
}

#[cfg(test)]
mod group_output_tests {
    use super::{
        Color, Diagnostic, Effect, GroupElement, LayoutResult, PageFrame, PathElement,
        PositionedElement, Rect,
    };
    use crate::{FillRule, Paint, Path, Stroke, Transform};

    #[test]
    fn path_and_group_arms_preserve_their_payloads() {
        let path = Path::rect(Rect {
            x: 1.0,
            y: 2.0,
            width: 3.0,
            height: 4.0,
        });
        let path_element = PathElement {
            path: path.clone(),
            fill: Some(Paint::Solid(Color::BLACK)),
            stroke: Some(Stroke::new(Paint::Solid(Color::WHITE), 2.0)),
        };
        let element = PositionedElement::Path(path_element.clone());
        assert!(matches!(
            element,
            PositionedElement::Path(actual) if actual == path_element
        ));

        let transform = Transform::rotate_about(15.0, 2.0, 3.0);
        let clip = Path {
            commands: Vec::new(),
            fill_rule: FillRule::EvenOdd,
        };
        let effect = Effect::OuterShadow {
            dx: 1.0,
            dy: 2.0,
            blur: 3.0,
            color: Color::BLACK,
        };
        let child_rect = Rect {
            x: 5.0,
            y: 6.0,
            width: 7.0,
            height: 8.0,
        };
        let group = GroupElement {
            transform,
            clip: Some(clip.clone()),
            opacity: 0.5,
            effects: vec![effect.clone()],
            children: vec![PositionedElement::FilledRect {
                rect: child_rect,
                color: Color::WHITE,
            }],
        };
        let element = PositionedElement::Group(group);
        let PositionedElement::Group(actual) = element else {
            panic!("constructed group should remain a group");
        };
        assert_eq!(actual.transform, transform);
        assert_eq!(actual.clip, Some(clip));
        assert_eq!(actual.opacity, 0.5);
        assert_eq!(actual.effects, vec![effect]);
        assert!(matches!(
            actual.children.as_slice(),
            [PositionedElement::FilledRect { rect, color }]
                if *rect == child_rect && *color == Color::WHITE
        ));
    }

    #[test]
    fn page_frame_new_defaults_background_to_none() {
        let page = PageFrame::new(1, 612.0, 792.0, Vec::new());
        assert_eq!(page.page_number, 1);
        assert_eq!(page.background, None);
    }

    #[test]
    fn layout_result_new_defaults_diagnostics_to_empty() {
        let result = LayoutResult::new(Vec::new(), Vec::new(), None, Vec::new());
        assert_eq!(result.diagnostics, Vec::<Diagnostic>::new());
    }

    #[test]
    fn group_transform_maps_child_coordinates_into_parent_coordinates() {
        let child_to_parent = Transform {
            a: 1.0,
            b: 0.0,
            c: 0.0,
            d: 1.0,
            e: 10.0,
            f: 20.0,
        };
        let group = GroupElement {
            transform: child_to_parent,
            clip: None,
            opacity: 1.0,
            effects: Vec::new(),
            children: Vec::new(),
        };
        assert_eq!(
            group.transform.apply(super::Point { x: 1.0, y: 2.0 }),
            super::Point { x: 11.0, y: 22.0 }
        );
    }
}

#[cfg(test)]
mod walk_tests {
    use super::{Color, GroupElement, PositionedElement, Rect, walk};
    use crate::{Point, Transform};

    fn translate(x: f64, y: f64) -> Transform {
        Transform {
            e: x,
            f: y,
            ..Transform::IDENTITY
        }
    }

    fn scale(value: f64) -> Transform {
        Transform {
            a: value,
            d: value,
            ..Transform::IDENTITY
        }
    }

    fn leaf(id: f64) -> PositionedElement {
        PositionedElement::FilledRect {
            rect: Rect {
                x: id,
                y: 0.0,
                width: 1.0,
                height: 1.0,
            },
            color: Color::BLACK,
        }
    }

    #[test]
    fn three_deep_groups_yield_every_leaf_once_with_the_correct_accumulated_transform() {
        let elements = vec![
            leaf(1.0),
            PositionedElement::Group(GroupElement {
                transform: translate(10.0, 0.0),
                clip: None,
                opacity: 1.0,
                effects: Vec::new(),
                children: vec![PositionedElement::Group(GroupElement {
                    transform: scale(2.0),
                    clip: None,
                    opacity: 1.0,
                    effects: Vec::new(),
                    children: vec![PositionedElement::Group(GroupElement {
                        transform: translate(0.0, 5.0),
                        clip: None,
                        opacity: 1.0,
                        effects: Vec::new(),
                        children: vec![leaf(2.0)],
                    })],
                })],
            }),
            leaf(3.0),
        ];
        let mut visited = Vec::new();
        walk(&elements, &mut |element, transform| {
            let PositionedElement::FilledRect { rect, .. } = element else {
                panic!("walk should yield leaves only");
            };
            visited.push((rect.x, transform.apply(Point { x: 1.0, y: 1.0 })));
        });
        assert_eq!(
            visited,
            vec![
                (1.0, Point { x: 1.0, y: 1.0 }),
                (2.0, Point { x: 12.0, y: 12.0 }),
                (3.0, Point { x: 1.0, y: 1.0 }),
            ]
        );
    }

    #[test]
    fn nested_group_transform_order_applies_child_before_parent() {
        let group = PositionedElement::Group(GroupElement {
            transform: translate(10.0, 0.0),
            clip: None,
            opacity: 1.0,
            effects: Vec::new(),
            children: vec![PositionedElement::Group(GroupElement {
                transform: scale(2.0),
                clip: None,
                opacity: 1.0,
                effects: Vec::new(),
                children: vec![leaf(1.0)],
            })],
        });
        let mut points = Vec::new();
        walk(&[group], &mut |_, transform| {
            points.push(transform.apply(Point { x: 1.0, y: 1.0 }));
        });
        assert_eq!(points, vec![Point { x: 12.0, y: 2.0 }]);
    }

    #[test]
    fn walk_does_not_yield_group_nodes() {
        let group = PositionedElement::Group(GroupElement {
            transform: Transform::IDENTITY,
            clip: None,
            opacity: 1.0,
            effects: Vec::new(),
            children: vec![leaf(1.0)],
        });
        walk(&[group], &mut |element, _| {
            assert!(!matches!(element, PositionedElement::Group(_)));
        });
    }

    #[test]
    fn walk_passes_identity_for_root_leaves() {
        walk(&[leaf(1.0)], &mut |_, transform| {
            assert_eq!(*transform, Transform::IDENTITY);
        });
    }
}
