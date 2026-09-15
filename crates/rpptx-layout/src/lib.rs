//! PresentationML inheritance resolution and shape-tree flattening.

use std::collections::HashMap;

use oxml_chart::CT_ChartSpace;
use oxml_layout::{
    Diagnostic, Effect, GroupElement, MediaId, Paint, Path, Point, Rect, Stroke, Transform,
};

mod context;
mod font;
mod style;
mod text;
pub mod timeline;

pub use context::{
    BackgroundContent, BackgroundSource, EffectiveBackground, FlattenedItem, FlattenedSource,
    ResolveCtx,
};
pub use style::{EffectiveShapeStyle, ResolveError};
pub use text::{EffectiveListStyle, EffectiveTextProperties};

/// Content-addressed media identifiers kept separate by relationship scope.
#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub struct ScopedMediaIds {
    pub slide: HashMap<String, MediaId>,
    pub layout: HashMap<String, MediaId>,
    pub master: HashMap<String, MediaId>,
    /// Package content type for each content-addressed media item.
    pub media_content_types: HashMap<MediaId, String>,
}

impl ScopedMediaIds {
    /// Look up an embedded relationship only in its producing part's scope.
    pub fn get(&self, source: FlattenedSource, relationship_id: &str) -> Option<MediaId> {
        match source {
            FlattenedSource::Slide => &self.slide,
            FlattenedSource::Layout => &self.layout,
            FlattenedSource::Master => &self.master,
            FlattenedSource::Background => return None,
        }
        .get(relationship_id)
        .copied()
    }

    /// Return the package content type for a relationship-resolved media item.
    pub fn content_type(&self, source: FlattenedSource, relationship_id: &str) -> Option<&str> {
        let media_id = self.get(source, relationship_id)?;
        self.media_content_types.get(&media_id).map(String::as_str)
    }
}

/// One chart relationship target projected during package assembly.
#[derive(Clone, Debug, PartialEq)]
pub enum ChartResource {
    Parsed(Box<CT_ChartSpace>),
    External(String),
    MissingTarget(String),
    Invalid(String),
}

/// Parsed chart resources kept separate by their producing part's scope.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct ScopedChartResources {
    pub slide: HashMap<String, ChartResource>,
    pub layout: HashMap<String, ChartResource>,
    pub master: HashMap<String, ChartResource>,
}

impl ScopedChartResources {
    /// Look up a chart relationship only in its producing part's scope.
    pub fn get(&self, source: FlattenedSource, relationship_id: &str) -> Option<&ChartResource> {
        match source {
            FlattenedSource::Slide => &self.slide,
            FlattenedSource::Layout => &self.layout,
            FlattenedSource::Master => &self.master,
            FlattenedSource::Background => return None,
        }
        .get(relationship_id)
    }
}

/// External hyperlink targets kept separate by relationship scope.
#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub struct ScopedHyperlinkTargets {
    pub slide: HashMap<String, String>,
    pub layout: HashMap<String, String>,
    pub master: HashMap<String, String>,
}

impl ScopedHyperlinkTargets {
    /// Look up an external hyperlink only in its producing part's scope.
    pub fn get(&self, source: FlattenedSource, relationship_id: &str) -> Option<&str> {
        match source {
            FlattenedSource::Slide => &self.slide,
            FlattenedSource::Layout => &self.layout,
            FlattenedSource::Master => &self.master,
            FlattenedSource::Background => return None,
        }
        .get(relationship_id)
        .map(String::as_str)
    }
}

/// An owned slide with every renderer-visible value resolved.
#[derive(Clone, Debug, PartialEq)]
pub struct ResolvedSlide {
    /// Slide width and height in typographic points.
    pub size: (f64, f64),
    pub background: Option<ResolvedBackground>,
    /// Shapes in final draw order.
    pub shapes: Vec<ResolvedShape>,
    pub diagnostics: Vec<Diagnostic>,
}

/// Paragraph base directions parallel to one resolved slide's shape and text-body order.
pub type ResolvedSlideTextDirections = Vec<Vec<Vec<oxml_layout::TextDirection>>>;

/// One concrete slide background ready for renderer lowering.
#[derive(Clone, Debug, PartialEq)]
pub enum ResolvedBackground {
    Paint(Paint),
    Image(ResolvedImage),
}

/// One renderer-facing shape with no OOXML model types or part-tree lifetimes.
#[derive(Clone, Debug, PartialEq)]
pub struct ResolvedShape {
    /// Affine mapping from this shape's coordinate space through all parent groups.
    pub group_transform: Transform,
    pub bounds: Rect,
    pub rotation_deg: f64,
    pub flip_h: bool,
    pub flip_v: bool,
    pub geometry: ResolvedGeometry,
    pub fill: Option<Paint>,
    pub image_fill: Option<ResolvedImage>,
    pub line: Option<Stroke>,
    pub head_end: Option<ResolvedLineEnd>,
    pub tail_end: Option<ResolvedLineEnd>,
    pub shadow: Option<Effect>,
    pub content: ResolvedContent,
    /// Stable fallback category when the renderer cannot reproduce the source.
    pub unsupported: Option<&'static str>,
}

/// One source-neutral line endpoint ready for renderer geometry lowering.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct ResolvedLineEnd {
    pub kind: ResolvedLineEndKind,
    pub width: ResolvedLineEndSize,
    pub length: ResolvedLineEndSize,
}

/// A visible line endpoint kind. DrawingML `none` resolves to no endpoint.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum ResolvedLineEndKind {
    Triangle,
    Stealth,
    Diamond,
    Oval,
    Arrow,
}

/// A line endpoint width or length multiplier category.
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub enum ResolvedLineEndSize {
    Small,
    #[default]
    Medium,
    Large,
}

/// Concrete geometry ready for the renderer, or a visible bounds fallback.
#[derive(Clone, Debug, PartialEq)]
pub enum ResolvedGeometry {
    Rectangle,
    Custom {
        paths: Vec<Path>,
        text_rect: Option<Rect>,
    },
    BoundsFallback,
}

/// Owned content attached to one resolved shape.
#[derive(Clone, Debug, PartialEq)]
pub enum ResolvedContent {
    None,
    Text(ResolvedTextBody),
    Image(ResolvedImage),
    Table(ResolvedTable),
    Group(GroupElement),
}

/// One resolved image shared by shape pictures and slide backgrounds.
#[derive(Clone, Debug, PartialEq)]
pub struct ResolvedImage {
    pub media: MediaId,
    pub src_rect: Option<CropRect>,
    pub placement: ResolvedImagePlacement,
    pub dpi: Option<f64>,
    pub rotate_with_shape: bool,
    /// Picture opacity from `a:alphaModFix`, from 0.0 to 1.0 for opaque.
    pub opacity: f64,
}

/// Destination placement for one resolved picture.
#[derive(Clone, Debug, PartialEq)]
pub enum ResolvedImagePlacement {
    Stretch { fill_rect: Option<CropRect> },
    Tile(ResolvedTilePlacement),
}

impl Default for ResolvedImagePlacement {
    fn default() -> Self {
        Self::Stretch { fill_rect: None }
    }
}

/// Concrete tile placement with all DrawingML defaults applied.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct ResolvedTilePlacement {
    pub translation: Point,
    pub scale_x: f64,
    pub scale_y: f64,
    pub flip: ResolvedTileFlip,
    pub alignment: ResolvedRectAlignment,
}

/// Alternating tile reflection axes.
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub enum ResolvedTileFlip {
    #[default]
    None,
    Horizontal,
    Vertical,
    Both,
}

/// The nine DrawingML rectangle alignment positions.
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub enum ResolvedRectAlignment {
    #[default]
    TopLeft,
    Top,
    TopRight,
    Left,
    Center,
    Right,
    BottomLeft,
    Bottom,
    BottomRight,
}

/// Picture cropping as fractions of the source dimensions.
#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub struct CropRect {
    pub left: f64,
    pub top: f64,
    pub right: f64,
    pub bottom: f64,
}

/// Text-box geometry and paragraphs after inheritance.
#[derive(Clone, Debug, PartialEq)]
pub struct ResolvedTextBody {
    pub insets: TextInsets,
    pub anchor: TextAnchor,
    pub wrap: bool,
    pub vertical: TextDirection,
    pub space_first_last_paragraph: bool,
    pub autofit: ResolvedAutofit,
    pub paragraphs: Vec<ResolvedParagraph>,
}

/// Text-box edge insets in points.
#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub struct TextInsets {
    pub left: f64,
    pub top: f64,
    pub right: f64,
    pub bottom: f64,
}

#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub enum TextAnchor {
    #[default]
    Top,
    Center,
    Bottom,
    Justified,
    Distributed,
}

#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub enum TextDirection {
    #[default]
    Horizontal,
    Vertical,
    Vertical270,
    EastAsianVertical,
    MongolianVertical,
    WordArtVertical,
    WordArtVerticalRtl,
}

#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub enum ResolvedAutofit {
    #[default]
    None,
    Shape,
    Normal {
        font_scale: Option<f64>,
        line_spacing_reduction: Option<f64>,
    },
}

/// One paragraph with concrete layout and text-run values.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct ResolvedParagraph {
    pub level: u8,
    pub left_margin: f64,
    pub right_margin: f64,
    pub indent: f64,
    pub alignment: ParagraphAlignment,
    pub line_spacing: Option<ResolvedTextSpacing>,
    pub space_before: Option<ResolvedTextSpacing>,
    pub space_after: Option<ResolvedTextSpacing>,
    pub bullet: Option<ResolvedBullet>,
    pub end_style: ResolvedRunStyle,
    pub runs: Vec<ResolvedTextRun>,
}

#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub enum ParagraphAlignment {
    #[default]
    Left,
    Center,
    Right,
    Justified,
    Distributed,
}

/// Concrete paragraph spacing as a font-relative fraction or typographic points.
#[derive(Clone, Debug, PartialEq)]
pub enum ResolvedTextSpacing {
    Percent(f64),
    Points(f64),
}

/// Concrete bullet size as a font-relative fraction or typographic points.
#[derive(Clone, Debug, PartialEq)]
pub enum ResolvedBulletSize {
    Percent(f64),
    Points(f64),
}

#[derive(Clone, Debug, PartialEq)]
pub enum ResolvedBullet {
    Character {
        character: String,
        font: Option<String>,
        color: Option<oxml_layout::Color>,
        size: Option<ResolvedBulletSize>,
    },
    AutoNumber {
        scheme: String,
        start_at: u32,
        font: Option<String>,
        color: Option<oxml_layout::Color>,
        size: Option<ResolvedBulletSize>,
    },
}

/// Ordered text, explicit breaks, and fields in one paragraph.
#[derive(Clone, Debug, PartialEq)]
pub enum ResolvedTextRun {
    Text {
        text: String,
        style: ResolvedRunStyle,
    },
    Break,
    Field {
        text: String,
        field_type: Option<String>,
        style: ResolvedRunStyle,
    },
}

/// Concrete renderer-visible properties of one text run.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct ResolvedRunStyle {
    pub font_size: Option<f64>,
    pub bold: bool,
    pub italic: bool,
    pub all_caps: bool,
    pub underline: bool,
    pub strike: bool,
    pub spacing: Option<f64>,
    pub baseline: Option<f64>,
    pub fill: Option<Paint>,
    pub latin_typeface: Option<String>,
    pub east_asian_typeface: Option<String>,
    pub complex_script_typeface: Option<String>,
    pub symbol_typeface: Option<String>,
    /// Direct external URI resolved in the run's producing relationship scope.
    pub hyperlink_url: Option<String>,
}

/// One table with concrete dimensions and owned cell text.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct ResolvedTable {
    pub right_to_left: bool,
    pub column_widths: Vec<f64>,
    pub rows: Vec<ResolvedTableRow>,
}

#[derive(Clone, Debug, Default, PartialEq)]
pub struct ResolvedTableRow {
    pub height: f64,
    pub cells: Vec<ResolvedTableCell>,
}

#[derive(Clone, Debug, PartialEq)]
pub struct ResolvedTableCell {
    pub text: Option<ResolvedTextBody>,
    pub fill: Option<Paint>,
    pub margins: TextInsets,
    pub left: Option<ResolvedTableBorder>,
    pub right: Option<ResolvedTableBorder>,
    pub top: Option<ResolvedTableBorder>,
    pub bottom: Option<ResolvedTableBorder>,
    pub row_span: u32,
    pub grid_span: u32,
    pub horizontal_merge: bool,
    pub vertical_merge: bool,
}

impl Default for ResolvedTableCell {
    fn default() -> Self {
        Self {
            text: None,
            fill: None,
            margins: TextInsets::default(),
            left: None,
            right: None,
            top: None,
            bottom: None,
            row_span: 1,
            grid_span: 1,
            horizontal_merge: false,
            vertical_merge: false,
        }
    }
}

/// One table edge plus its resolved style-precedence rank.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct ResolvedTableBorder {
    pub stroke: Option<Stroke>,
    pub priority: u8,
}
