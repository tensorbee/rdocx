//! Public facade for PresentationML packages.
//!
//! The `default-template` feature embeds a 16:9 zero-slide template derived
//! from the MIT-licensed python-pptx 1.0.2 default template. The slide size was
//! changed to 12,192,000 by 6,858,000 EMU, the size kind was changed to
//! `screen16x9`, and python-pptx generated the notes-master infrastructure.

mod arrange;
mod embedded;

pub use arrange::{ArrangeReference, ConnectorEnd, DistributeDirection, ShapeAlignment};
pub use embedded::{
    EmbeddedContentInfo, EmbeddedContentKind, EmbeddedMutationPolicy, EmbeddedSignatureState,
};

use std::collections::{HashMap, HashSet};
use std::io::Cursor;
use std::path::Path;

#[cfg(feature = "render")]
use diagram::{DiagramResources, ScopedDiagramResources};
#[cfg(feature = "render")]
use oxml_chart::CT_ChartSpace;
pub use oxml_chart::{ChartData, ChartKind, RgbColor};
use oxml_core::OxmlError;
pub use oxml_core::core_properties::CoreProperties;
pub use oxml_core::units::{Angle, Emu, Percent1000};
#[cfg(feature = "render")]
use oxml_drawing::color::ColorMap;
pub use oxml_drawing::color::{ColorChoice, ColorTransform};
pub use oxml_drawing::effect::{CT_EffectList, CT_OuterShadowEffect, RectAlignment};
use oxml_drawing::fill::RelativeRect;
pub use oxml_drawing::fill::{Fill, NoFill, PatternFill, SolidFill};
use oxml_drawing::geometry::{CT_PresetGeometry2D, Guide, GuideOp, GuideOperand};
pub use oxml_drawing::line::{
    CT_LineProperties, LineDash, LineEnd, LineEndSize, LineEndType, PresetDash,
    ST_PresetLineDashVal,
};
use oxml_drawing::namespace::A_NS;
use oxml_drawing::shape_props::CT_ShapeProperties;
#[cfg(feature = "render")]
use oxml_drawing::table::CT_TableStyleList;
use oxml_drawing::table::{CT_Table, CT_TableCell, CT_TableCellProperties, CT_TableProperties};
#[cfg(feature = "render")]
use oxml_drawing::text::CT_TextListStyle;
use oxml_drawing::text::{
    CT_RegularTextRun, CT_TextBody, CT_TextParagraph, Coordinate32Value, NormalAutofit,
    TextAutofit, TextHyperlink, TextRun, TextWrap,
};
pub use oxml_drawing::text::{
    CT_TextCharacterProperties, CT_TextParagraphProperties, TextAlignment, TextAnchor, TextBullet,
    TextBulletCharacter, TextBulletChoice, TextFont, TextNoBullet, TextSpacing, TextStrike,
    TextUnderline,
};
#[cfg(feature = "render")]
use oxml_drawing::theme::CT_OfficeStyleSheet;
use oxml_drawing::xfrm::{CT_Point2D, CT_PositiveSize2D, CT_Transform2D};
use oxml_layout::MediaId;
#[cfg(feature = "render")]
use oxml_layout::{
    Color, FontManager, GlyphRun, GroupElement, LayoutResult, PageFrame, Path as LayoutPath,
    PathElement, Point, PositionedElement, Rect, Transform,
};
use oxml_media::{
    ImageFormat, MediaNamer, audio_video_signature_matches, is_safe_content_type, probe, resolve,
};
pub use oxml_opc::PackageReadLimits;
use oxml_opc::content_types;
use oxml_opc::relationship::{Relationship, rel_types};
#[cfg(feature = "digital-signatures")]
pub use oxml_opc::{
    CoveredRelationship, SignatureIssue, SignatureReport, SignerCertificateIdentity,
};
use oxml_opc::{OpcError, OpcPackage, Relationships, write_atomic_file};
use quick_xml::events::{BytesStart as XmlBytesStart, Event as XmlEvent};
use quick_xml::name::{Namespace as XmlNamespace, ResolveResult as XmlResolveResult};
use quick_xml::reader::NsReader;
use quick_xml::{Writer as XmlWriter, XmlVersion};
#[cfg(feature = "render")]
pub use rpptx_layout::timeline::{
    EvaluatedFrameState, EvaluatedMediaState, MediaPlaybackPhase, TimelinePosition,
};
#[cfg(feature = "render")]
use rpptx_layout::timeline::{ResolvedTimelineSlide, evaluate_media_playback};
#[cfg(feature = "render")]
use rpptx_layout::{
    ChartResource, FlattenedItem, FlattenedSource, ResolveCtx, ResolvedAutofit, ResolvedContent,
    ResolvedSlide, ResolvedSlideTextDirections, ScopedChartResources, ScopedHyperlinkTargets,
    ScopedMediaFailures, ScopedMediaIds, inherited_xfrm,
};
pub use rpptx_oxml::comments::{Comment, CommentAuthor, CommentReply};
use rpptx_oxml::comments::{CommentAuthorList, CommentList};
use rpptx_oxml::connector::CT_ConnectionShape;
pub use rpptx_oxml::diagram::{
    CT_DiagramColorsDefinition, CT_DiagramData, CT_DiagramDrawing, CT_DiagramLayoutDefinition,
    CT_DiagramStyleDefinition, DiagramConnection, DiagramConnectionKind, DiagramLayoutFamily,
    DiagramPoint, DiagramPointKind, DiagramRelationshipIds, DiagramShapeStyle,
};
use rpptx_oxml::graphic_frame::{CT_GraphicFrame, GraphicDataPayload};
use rpptx_oxml::namespace::{P_NS, R_NS};
use rpptx_oxml::notes_parts::{CT_HandoutMaster, CT_NotesMaster, CT_NotesSlide};
pub use rpptx_oxml::picture::MediaKind;
use rpptx_oxml::picture::{CT_Picture, MediaSource as PictureMediaSource, PictureMedia};
#[cfg(feature = "render")]
use rpptx_oxml::placeholder::PlaceholderKey;
use rpptx_oxml::placeholder::{CT_Placeholder, PhType};
pub use rpptx_oxml::presentation::Section;
use rpptx_oxml::presentation::{CT_Presentation, CT_SlideId, custom_show_relationship_ids};
use rpptx_oxml::relmap::{
    relationship_ids, release_hyperlinks, rewrite_exact_rel_ids, rewrite_rel_ids,
};
use rpptx_oxml::shape_tree::{
    CT_GroupShape, CT_Shape, CT_ShapeTree, ShapeIdAllocator, ShapeTreeChild, rewrite_shape_ids,
};
#[cfg(feature = "render")]
use rpptx_oxml::slide_parts::CT_CommonSlideData;
#[cfg(feature = "render")]
use rpptx_oxml::slide_parts::ColorMapOverrideKind;
use rpptx_oxml::slide_parts::{
    BackgroundRendering, CT_HeaderFooter, CT_Slide, CT_SlideLayout, CT_SlideMaster,
};
pub use rpptx_oxml::timing::MediaPlaybackTrigger;
use rpptx_oxml::timing::{CT_Timing, MediaCommandKind, MediaDisplayPolicy};
#[cfg(feature = "render")]
use rpptx_oxml::timing::{TimingNode, TimingTarget};
#[cfg(feature = "render")]
use rpptx_render::timeline::{
    compose_morph_transition, compose_transition, layout_timeline_slide_with_font_manager,
};
#[cfg(feature = "render")]
use rpptx_render::{
    MediaData, RenderInput, layout_presentation_with_font_manager_and_text_directions_mut,
};
use thiserror::Error;

#[cfg(feature = "render")]
mod animation;
#[cfg(feature = "render")]
mod diagram;
#[cfg(feature = "default-template")]
mod html;
#[cfg(feature = "default-template")]
mod odp;
#[cfg(feature = "render")]
mod pdf;
#[cfg(feature = "render")]
pub use animation::{
    AnimationExportOptions, AnimationFormat, AnimationSegment, AnimationTransition,
    DeterministicAnimation, GifLoopBehavior,
};
#[cfg(feature = "default-template")]
pub use html::{HtmlDiagnostic, HtmlImageResource, HtmlReadResult};
#[cfg(feature = "default-template")]
pub use odp::{OdpDiagnostic, OdpReadResult, OdpWriteResult};
#[cfg(feature = "render")]
pub use pdf::{PdfImportDiagnostic, PdfImportLimits, PdfImportMode, PdfImportResult};

#[cfg(any(feature = "default-template", feature = "render"))]
const DEFAULT_TEMPLATE: &[u8] = include_bytes!("../assets/default.pptx");
const DEFAULT_CORE_PROPERTIES_PART: &str = "/docProps/core.xml";

/// One deterministic evaluated slide frame returned by the native facade.
#[cfg(feature = "render")]
#[derive(Clone, Debug)]
pub struct DeterministicTimelineFrame {
    pub page: PageFrame,
    pub state: EvaluatedFrameState,
    pub diagnostics: Vec<oxml_layout::Diagnostic>,
}

/// How the media-aware renderer handles a media picture's poster.
#[cfg(feature = "render")]
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum MediaFallbackPolicy {
    PosterFrame,
    DeterministicPlaceholder,
    Fail,
}

/// An audience-handout slide arrangement.
#[cfg(feature = "render")]
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum HandoutLayout {
    One,
    Two,
    Three,
    Four,
    Six,
    Nine,
}

#[cfg(feature = "render")]
impl HandoutLayout {
    const fn slides_per_page(self) -> usize {
        match self {
            Self::One => 1,
            Self::Two => 2,
            Self::Three => 3,
            Self::Four => 4,
            Self::Six => 6,
            Self::Nine => 9,
        }
    }
}

/// One deterministic page and its synchronized media playback state.
#[cfg(feature = "render")]
#[derive(Clone, Debug)]
pub struct DeterministicMediaTimelineFrame {
    pub frame: DeterministicTimelineFrame,
    pub media: Vec<EvaluatedMediaState>,
}

/// One text-bearing slide shape as the deterministic renderer lays it out.
#[cfg(feature = "render")]
#[derive(Clone, Debug, PartialEq)]
pub struct TextFrameLayout {
    pub slide_index: usize,
    pub shape_id: Option<u32>,
    pub name: Option<String>,
    /// The autofit mode in effect after placeholder inheritance.
    pub autofit: AutofitMode,
    pub layout: rpptx_render::ShapeTextLayout,
}
const DEFAULT_POWERPOINT_AUTHORS_PART: &str = "/ppt/authors.xml";
const DEFAULT_POWERPOINT_COMMENTS_PART: &str = "/ppt/comments/comment1.xml";
const MAX_SMARTART_TRANSFER_DIAGRAM_PARTS: usize = 128;

/// A result returned by the `rpptx` read facade.
pub type Result<T> = std::result::Result<T, Error>;

/// One relationship-resolved diagram part.
#[derive(Clone, Debug, PartialEq)]
pub enum DiagramPart<T> {
    Parsed(T),
    External(String),
    MissingTarget(String),
    Invalid(String),
}

/// One SmartArt frame and the diagram resources owned by its relationship scope.
#[derive(Clone, Debug, PartialEq)]
pub struct SmartArtInfo {
    pub slide_index: usize,
    pub shape_id: u32,
    pub relationships: DiagramRelationshipIds,
    pub data: DiagramPart<CT_DiagramData>,
    pub layout: DiagramPart<CT_DiagramLayoutDefinition>,
    pub style: DiagramPart<CT_DiagramStyleDefinition>,
    pub colors: DiagramPart<CT_DiagramColorsDefinition>,
    pub drawing: Option<DiagramPart<CT_DiagramDrawing>>,
}

/// One inspected media source resolved through the slide relationship scope.
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum MediaLocation {
    Embedded {
        part_name: String,
        content_type: String,
    },
    Linked {
        target: String,
    },
}

/// Caller bytes and package metadata for one embedded media payload.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct EmbeddedMediaInput<'a> {
    pub bytes: &'a [u8],
    pub filename: &'a str,
    pub content_type: &'a str,
}

/// A new embedded or external media source.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum MediaSourceInput<'a> {
    Embedded(EmbeddedMediaInput<'a>),
    Linked {
        target: &'a str,
        content_type: &'a str,
    },
}

/// Required poster bytes for a newly authored media picture.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct MediaPoster<'a> {
    pub bytes: &'a [u8],
    pub filename: &'a str,
}

/// The supported media playback settings retained by inspection and mutation.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct MediaPlaybackSettings {
    pub trim_start_ms: Option<u64>,
    pub trim_end_ms: Option<u64>,
    pub volume: u32,
    pub looped: bool,
    pub show_when_stopped: bool,
    pub trigger: MediaPlaybackTrigger,
}

impl Default for MediaPlaybackSettings {
    fn default() -> Self {
        Self {
            trim_start_ms: None,
            trim_end_ms: None,
            volume: 100_000,
            looped: false,
            show_when_stopped: false,
            trigger: MediaPlaybackTrigger::OnClick,
        }
    }
}

/// A non-fatal limitation discovered while inspecting media.
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum MediaDiagnostic {
    UnsupportedContentType { content_type: String },
    MissingPlaybackTiming,
    UnsupportedPlaybackCommand { command: String },
    MissingPoster,
}

/// One media picture and its package-resolved playback state.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct MediaInfo {
    pub slide_index: usize,
    pub shape_id: u32,
    pub kind: MediaKind,
    pub source: MediaLocation,
    pub poster_relationship_id: Option<String>,
    pub settings: MediaPlaybackSettings,
    pub diagnostics: Vec<MediaDiagnostic>,
}

/// The embedded image part shown by one picture shape.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct PictureImage<'a> {
    pub part_name: String,
    pub content_type: String,
    pub bytes: &'a [u8],
}

/// An error opening, resolving, or serialising a presentation package.
#[derive(Debug, Error)]
pub enum Error {
    #[error(transparent)]
    Package(#[from] OpcError),

    #[cfg(feature = "default-template")]
    #[error("HTML conversion error at {path}: {message}")]
    Html { path: String, message: String },

    #[cfg(feature = "default-template")]
    #[error("ODP conversion error in {part:?} at byte {offset}: {message}")]
    Odp {
        part: Option<String>,
        offset: u64,
        message: String,
    },

    #[cfg(feature = "render")]
    #[error("PDF conformance error: {0}")]
    Pdf(#[from] oxml_pdf::PdfError),

    #[cfg(feature = "render")]
    #[error("PDF import error on page {page:?} at operation {offset:?}: {message}")]
    PdfImport {
        page: Option<u32>,
        offset: Option<usize>,
        message: String,
    },

    #[error("presentation package has no officeDocument relationship")]
    MissingMainDocument,

    #[error("unsupported presentation main content type: {content_type:?}")]
    UnsupportedPackageClass { content_type: Option<String> },

    #[error("{source_part}: relationship {relationship_id} is missing")]
    MissingRelationship {
        source_part: String,
        relationship_id: String,
    },

    #[error("{source_part}: relationship {relationship_id} has type {actual}, expected {expected}")]
    WrongRelationshipType {
        source_part: String,
        relationship_id: String,
        expected: &'static str,
        actual: String,
    },

    #[error("{source_part}: relationship {relationship_id} has an external target")]
    ExternalRelationship {
        source_part: String,
        relationship_id: String,
    },

    #[error("package part is missing: {part_name}")]
    MissingPart { part_name: String },

    #[error("cannot create core properties at occupied package part {part_name}")]
    CorePropertiesPartCollision { part_name: String },

    #[error("cannot create collaboration data at occupied package part {part_name}")]
    CollaborationPartCollision { part_name: String },

    #[error(
        "{source_part}: found {count} relationships of type {relationship_type}, expected at most one"
    )]
    DuplicateRelationships {
        source_part: String,
        relationship_type: &'static str,
        count: usize,
    },

    #[error("{source_part}: found {count} notes-slide relationships, expected at most one")]
    DuplicateNotesSlides { source_part: String, count: usize },

    #[error("malformed PresentationML part {part_name}: {message}")]
    MalformedPart { part_name: String, message: String },

    #[error("layout index {index} is out of range for {layout_count} layouts")]
    UnknownLayoutIndex { index: usize, layout_count: usize },

    #[error("slide index {index} is out of range for {slide_count} slides")]
    UnknownSlideIndex { index: usize, slide_count: usize },

    #[error("picture {filename} has unsupported image bytes")]
    UnsupportedPicture { filename: String },

    #[error("picture {filename} has no usable native dimensions")]
    UnavailablePictureDimensions { filename: String },

    #[error("picture {filename} has an inferred {axis} outside the EMU range")]
    PictureDimensionOutOfRange {
        filename: String,
        axis: &'static str,
    },

    #[error("PowerPoint slide ids are exhausted")]
    SlideIdExhausted,

    #[error("{operation} is not supported for {shape_kind:?}")]
    UnsupportedShapeMutation {
        operation: &'static str,
        shape_kind: ShapeKind,
    },

    #[error("{operation} failed: {message}")]
    InvalidShapeMutation {
        operation: &'static str,
        message: String,
    },

    #[error("{operation} failed: {message}")]
    InvalidTableMutation {
        operation: &'static str,
        message: String,
    },

    #[error("{operation} failed: {message}")]
    InvalidPresentationMutation {
        operation: &'static str,
        message: String,
    },

    #[error("{operation} failed: {message}")]
    InvalidSlideMutation {
        operation: &'static str,
        message: String,
    },

    #[error("{operation} failed: {message}")]
    InvalidChartMutation {
        operation: &'static str,
        message: String,
    },

    #[error("{operation} failed: {message}")]
    InvalidMediaMutation {
        operation: &'static str,
        message: String,
    },

    #[error("{operation} failed: {message}")]
    InvalidEmbeddedMutation {
        operation: &'static str,
        message: String,
    },

    #[error("adjustment {name} requires a finite value")]
    NonFiniteAdjustmentValue { name: String },

    #[cfg(feature = "render")]
    #[error("presentation render failed: {message}")]
    Render { message: String },
}

/// The startup and macro class carried by a modern PresentationML package.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum PresentationPackageClass {
    Presentation,
    MacroEnabledPresentation,
    Template,
    MacroEnabledTemplate,
    Slideshow,
    MacroEnabledSlideshow,
}

impl PresentationPackageClass {
    fn from_content_type(content_type: &str) -> Option<Self> {
        match content_type {
            content_types::PRESENTATION => Some(Self::Presentation),
            content_types::PRESENTATION_MACRO_ENABLED => Some(Self::MacroEnabledPresentation),
            content_types::PRESENTATION_TEMPLATE => Some(Self::Template),
            content_types::PRESENTATION_TEMPLATE_MACRO_ENABLED => Some(Self::MacroEnabledTemplate),
            content_types::SLIDESHOW => Some(Self::Slideshow),
            content_types::SLIDESHOW_MACRO_ENABLED => Some(Self::MacroEnabledSlideshow),
            _ => None,
        }
    }

    fn content_type(self) -> &'static str {
        match self {
            Self::Presentation => content_types::PRESENTATION,
            Self::MacroEnabledPresentation => content_types::PRESENTATION_MACRO_ENABLED,
            Self::Template => content_types::PRESENTATION_TEMPLATE,
            Self::MacroEnabledTemplate => content_types::PRESENTATION_TEMPLATE_MACRO_ENABLED,
            Self::Slideshow => content_types::SLIDESHOW,
            Self::MacroEnabledSlideshow => content_types::SLIDESHOW_MACRO_ENABLED,
        }
    }

    /// The class that a `.pptx`, `.pptm`, `.potx`, `.potm`, `.ppsx`, or `.ppsm` path names.
    fn from_path(path: &Path) -> Option<Self> {
        match path.extension()?.to_str()?.to_ascii_lowercase().as_str() {
            "pptx" => Some(Self::Presentation),
            "pptm" => Some(Self::MacroEnabledPresentation),
            "potx" => Some(Self::Template),
            "potm" => Some(Self::MacroEnabledTemplate),
            "ppsx" => Some(Self::Slideshow),
            "ppsm" => Some(Self::MacroEnabledSlideshow),
            _ => None,
        }
    }

    fn is_macro_enabled(self) -> bool {
        matches!(
            self,
            Self::MacroEnabledPresentation
                | Self::MacroEnabledTemplate
                | Self::MacroEnabledSlideshow
        )
    }
}

/// A package or presentation invariant that would make a saved deck unsafe.
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum ValidationIssue {
    DuplicateShapeId {
        slide: usize,
        id: u32,
    },
    SlideIdOutOfRange {
        slide: usize,
        id: u32,
    },
    DuplicateSlideId {
        id: u32,
    },
    MissingContentTypeOverride {
        part: String,
    },
    DanglingRelationship {
        part: String,
        r_id: String,
        target: String,
    },
    UnreachableRelationshipTarget {
        part: String,
        target: String,
    },
    EmptyTextBody {
        slide: usize,
        shape: usize,
    },
    DuplicatePlaceholderIdx {
        slide: usize,
        idx: u32,
    },
    OrphanMedia {
        part: String,
    },
    CustomShowReference {
        slide_id: u32,
    },
    MissingLayoutRel {
        slide: usize,
    },
    MissingThemeRel {
        master: usize,
    },
}

/// An opened PresentationML package and its ordered slide read model.
#[derive(Clone, Debug)]
pub struct Presentation {
    package: OpcPackage,
    media_store: MediaStore,
    presentation_part: String,
    presentation: CT_Presentation,
    core_properties: Option<CoreProperties>,
    core_properties_part_name: String,
    core_properties_dirty: bool,
    comment_authors_part: Option<String>,
    comment_authors: CommentAuthorList,
    comment_authors_dirty: bool,
    notes_master_part: Option<String>,
    notes_master: Option<Box<CT_NotesMaster>>,
    notes_master_dirty: bool,
    handout_master_part: Option<String>,
    handout_master: Option<Box<CT_HandoutMaster>>,
    handout_master_dirty: bool,
    embedded_invalidated_signatures: HashSet<(String, String)>,
    package_signatures_invalidated: bool,
    layouts: Vec<LayoutRecord>,
    slides: Vec<SlideRecord>,
}

#[derive(Clone, Debug)]
struct LayoutRecord {
    part_name: String,
    layout: CT_SlideLayout,
}

#[derive(Clone, Debug)]
struct SlideRecord {
    id: u32,
    part_name: String,
    slide: CT_Slide,
    notes: Option<NotesRecord>,
    comments: Option<CommentRecord>,
}

#[derive(Clone, Debug)]
struct CommentRecord {
    part_name: String,
    comments: CommentList,
    dirty: bool,
}

#[derive(Clone, Debug)]
struct NotesRecord {
    part_name: String,
    notes: CT_NotesSlide,
}

#[derive(Clone, Debug)]
struct MediaEntry {
    part_name: String,
    bytes: Vec<u8>,
    extension: String,
    content_type: String,
    newly_inserted: bool,
}

#[derive(Clone, Debug)]
struct MediaStore {
    by_id: HashMap<MediaId, Vec<MediaEntry>>,
    namer: MediaNamer,
    media_namer: MediaNamer,
}

impl MediaStore {
    fn scan(package: &OpcPackage) -> Self {
        let namer = MediaNamer::scan(
            "/ppt/media",
            "image",
            package.parts.keys().map(String::as_str),
        );
        let media_namer = MediaNamer::scan(
            "/ppt/media",
            "media",
            package.parts.keys().map(String::as_str),
        );
        let mut by_id: HashMap<MediaId, Vec<MediaEntry>> = HashMap::new();
        let mut parts = package
            .parts
            .iter()
            .filter(|(part_name, _)| part_name.to_ascii_lowercase().starts_with("/ppt/media/"))
            .collect::<Vec<_>>();
        parts.sort_unstable_by_key(|(part_name, _)| part_name.as_str());
        for (part_name, bytes) in parts {
            let format = resolve(bytes, part_name);
            let extension = part_name
                .rsplit_once('.')
                .map_or_else(|| format.extension(), |(_, extension)| extension)
                .to_owned();
            let content_type = package
                .content_types
                .content_type_for(part_name)
                .unwrap_or_else(|| format.content_type())
                .to_owned();
            by_id
                .entry(MediaId::from_bytes(bytes))
                .or_default()
                .push(MediaEntry {
                    part_name: part_name.clone(),
                    bytes: bytes.clone(),
                    extension,
                    content_type,
                    newly_inserted: false,
                });
        }
        Self {
            by_id,
            namer,
            media_namer,
        }
    }

    fn insert(&mut self, package: &mut OpcPackage, bytes: &[u8], filename: &str) -> String {
        self.insert_with_id(package, MediaId::from_bytes(bytes), bytes, filename)
    }

    #[allow(dead_code)]
    fn insert_with_id(
        &mut self,
        package: &mut OpcPackage,
        media_id: MediaId,
        bytes: &[u8],
        filename: &str,
    ) -> String {
        if let Some(existing) = self
            .by_id
            .get(&media_id)
            .and_then(|bucket| bucket.iter().find(|entry| entry.bytes == bytes))
        {
            return existing.part_name.clone();
        }

        let format = resolve(bytes, filename);
        let extension = format.extension();
        let content_type = format.content_type();
        let part_name = self.namer.next_part_name(extension);
        package.set_part(&part_name, bytes.to_vec());
        register_content_type(package, &part_name, extension, content_type);
        self.by_id.entry(media_id).or_default().push(MediaEntry {
            part_name: part_name.clone(),
            bytes: bytes.to_vec(),
            extension: extension.to_owned(),
            content_type: content_type.to_owned(),
            newly_inserted: true,
        });
        part_name
    }

    fn insert_explicit(
        &mut self,
        package: &mut OpcPackage,
        bytes: &[u8],
        extension: &str,
        content_type: &str,
    ) -> String {
        let media_id = MediaId::from_bytes(bytes);
        if let Some(existing) = self.by_id.get(&media_id).and_then(|bucket| {
            bucket.iter().find(|entry| {
                entry.bytes == bytes
                    && entry.extension == extension
                    && entry.content_type == content_type
            })
        }) {
            return existing.part_name.clone();
        }
        let part_name = self.media_namer.next_part_name(extension);
        package.set_part(&part_name, bytes.to_vec());
        register_content_type(package, &part_name, extension, content_type);
        self.by_id.entry(media_id).or_default().push(MediaEntry {
            part_name: part_name.clone(),
            bytes: bytes.to_vec(),
            extension: extension.to_owned(),
            content_type: content_type.to_owned(),
            newly_inserted: true,
        });
        part_name
    }

    fn write_new_parts(&self, package: &mut OpcPackage) {
        for entry in self.by_id.values().flatten() {
            if entry.newly_inserted {
                package.set_part(&entry.part_name, entry.bytes.clone());
                register_content_type(
                    package,
                    &entry.part_name,
                    &entry.extension,
                    &entry.content_type,
                );
            }
        }
    }
}

fn register_content_type(
    package: &mut OpcPackage,
    part_name: &str,
    extension: &str,
    content_type: &str,
) {
    match package.content_types.content_type_for(part_name) {
        Some(existing) if existing == content_type => {}
        Some(_) => package.content_types.add_override(part_name, content_type),
        None => package.content_types.add_default(extension, content_type),
    }
}

impl Presentation {
    /// Creates an empty 16:9 presentation from the bundled standard template.
    #[cfg(feature = "default-template")]
    pub fn new() -> Result<Self> {
        Self::from_bytes(DEFAULT_TEMPLATE)
    }

    /// Opens a presentation from a `.pptx` path.
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        Self::from_package(OpcPackage::open(path)?)
    }

    /// Opens a password-protected Agile-encrypted presentation from a path.
    #[cfg(all(feature = "agile-encryption", not(target_arch = "wasm32")))]
    pub fn open_encrypted<P: AsRef<Path>>(path: P, password: &str) -> Result<Self> {
        let file = std::fs::File::open(path).map_err(OpcError::from)?;
        Self::from_package(OpcPackage::from_encrypted_reader(file, password)?)
    }

    /// Opens a presentation from in-memory `.pptx` bytes.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_package(OpcPackage::from_reader(Cursor::new(bytes))?)
    }

    /// Opens a password-protected Agile-encrypted presentation from bytes.
    #[cfg(all(feature = "agile-encryption", not(target_arch = "wasm32")))]
    pub fn from_encrypted_bytes(bytes: &[u8], password: &str) -> Result<Self> {
        Self::from_encrypted_bytes_with_limits(bytes, password, PackageReadLimits::UNBOUNDED)
    }

    /// Opens an Agile-encrypted presentation with bounded archive expansion.
    #[cfg(all(feature = "agile-encryption", not(target_arch = "wasm32")))]
    pub fn from_encrypted_bytes_with_limits(
        bytes: &[u8],
        password: &str,
        limits: PackageReadLimits,
    ) -> Result<Self> {
        Self::from_package(OpcPackage::from_encrypted_reader_with_limits(
            Cursor::new(bytes),
            password,
            limits,
        )?)
    }

    fn from_package(package: OpcPackage) -> Result<Self> {
        let package_signatures_invalidated =
            embedded::known_invalid_package_signature_on_open(&package);
        let media_store = MediaStore::scan(&package);
        let main_relationship = package
            .package_rels
            .get_by_type(rel_types::DOCUMENT)
            .ok_or(Error::MissingMainDocument)?;
        reject_external("/", main_relationship)?;
        let presentation_part = OpcPackage::resolve_rel_target("/", &main_relationship.target);
        let main_content_type = package
            .content_types
            .content_type_for(&presentation_part)
            .map(str::to_owned);
        PresentationPackageClass::from_content_type(main_content_type.as_deref().unwrap_or(""))
            .ok_or(Error::UnsupportedPackageClass {
                content_type: main_content_type,
            })?;
        let presentation_xml = required_part(&package, &presentation_part)?;
        let presentation =
            CT_Presentation::from_xml(presentation_xml).map_err(|error| Error::MalformedPart {
                part_name: presentation_part.clone(),
                message: error.to_string(),
            })?;
        let core_properties_relationship =
            package.package_rels.get_by_type(rel_types::CORE_PROPERTIES);
        if let Some(relationship) = core_properties_relationship {
            reject_external("/", relationship)?;
        }
        let core_properties_part_name = core_properties_relationship
            .map(|relationship| OpcPackage::resolve_rel_target("/", &relationship.target));
        let core_properties = if let Some(part_name) = &core_properties_part_name {
            let xml = required_part(&package, part_name)?;
            Some(
                CoreProperties::from_xml(xml).map_err(|error| Error::MalformedPart {
                    part_name: part_name.clone(),
                    message: error.to_string(),
                })?,
            )
        } else {
            None
        };
        let layouts = resolve_layouts(&package, &presentation_part, &presentation)?;
        let (comment_authors_part, comment_authors) =
            resolve_comment_authors(&package, &presentation_part)?;
        let (notes_master_part, notes_master) = resolve_notes_master(&package, &presentation_part)?;
        let (handout_master_part, handout_master) =
            resolve_handout_master(&package, &presentation_part)?;

        let presentation_relationships = package.get_part_rels(&presentation_part);
        let mut slides = Vec::with_capacity(presentation.slide_ids.len());
        for slide_id in &presentation.slide_ids {
            let relationship = presentation_relationships
                .and_then(|relationships| relationships.get_by_id(&slide_id.relationship_id))
                .ok_or_else(|| Error::MissingRelationship {
                    source_part: presentation_part.clone(),
                    relationship_id: slide_id.relationship_id.clone(),
                })?;
            require_relationship_type(&presentation_part, relationship, rel_types::SLIDE)?;
            reject_external(&presentation_part, relationship)?;
            let slide_part =
                OpcPackage::resolve_rel_target(&presentation_part, &relationship.target);
            let slide_xml = required_part(&package, &slide_part)?;
            let slide = CT_Slide::from_xml(slide_xml).map_err(|error| Error::MalformedPart {
                part_name: slide_part.clone(),
                message: error.to_string(),
            })?;
            let notes = resolve_notes(&package, &slide_part)?;
            let comments = resolve_comments(&package, &slide_part, &slide)?;
            slides.push(SlideRecord {
                id: slide_id.id,
                part_name: slide_part,
                slide,
                notes,
                comments,
            });
        }
        validate_collaboration_model(&comment_authors, &slides).map_err(|message| {
            Error::MalformedPart {
                part_name: presentation_part.clone(),
                message,
            }
        })?;

        Ok(Self {
            package,
            media_store,
            presentation_part,
            presentation,
            core_properties,
            core_properties_part_name: core_properties_part_name
                .unwrap_or_else(|| DEFAULT_CORE_PROPERTIES_PART.to_owned()),
            core_properties_dirty: false,
            comment_authors_part,
            comment_authors,
            comment_authors_dirty: false,
            notes_master_part,
            notes_master,
            notes_master_dirty: false,
            handout_master_part,
            handout_master,
            handout_master_dirty: false,
            embedded_invalidated_signatures: HashSet::new(),
            package_signatures_invalidated,
            layouts,
            slides,
        })
    }

    /// Serialises the package through the deterministic OPC writer.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        debug_assert!(
            self.validate().is_empty(),
            "invalid presentation at byte save boundary: {:?}",
            self.validate()
        );
        let package_signatures_invalidated = self.package_signatures_invalidated
            || self.retained_package_signature_would_be_invalidated()?;
        let mut package = self.staged_package(true)?;
        embedded::persist_invalidated_package_signature(
            &mut package,
            package_signatures_invalidated,
        )?;
        let mut output = Cursor::new(Vec::new());
        package.write_to(&mut output)?;
        Ok(output.into_inner())
    }

    /// Returns the exact modern package class carried by the main part.
    pub fn package_class(&self) -> Result<PresentationPackageClass> {
        let content_type = self
            .package
            .content_types
            .content_type_for(&self.presentation_part)
            .map(str::to_owned);
        PresentationPackageClass::from_content_type(content_type.as_deref().unwrap_or(""))
            .ok_or(Error::UnsupportedPackageClass { content_type })
    }

    /// Serialises an output copy with the selected package class.
    ///
    /// Executable payloads and every unrelated part and relationship remain
    /// byte-preserved. The live presentation keeps its original class.
    pub fn to_bytes_as(&self, class: PresentationPackageClass) -> Result<Vec<u8>> {
        debug_assert!(
            self.validate().is_empty(),
            "invalid presentation at package-class save boundary: {:?}",
            self.validate()
        );
        let class_changed = self.package_class()? != class;
        let preserve_signed_parts = self
            .package
            .package_rels
            .get_by_type(rel_types::DIGITAL_SIGNATURE_ORIGIN)
            .is_some();
        let package_signatures_invalidated = self.package_signatures_invalidated
            || self.retained_package_signature_would_be_invalidated()?
            || (class_changed && preserve_signed_parts);
        let mut package = self.staged_package(true)?;
        package
            .content_types
            .add_override(&self.presentation_part, class.content_type());
        embedded::persist_invalidated_package_signature(
            &mut package,
            package_signatures_invalidated,
        )?;
        let mut output = Cursor::new(Vec::new());
        package.write_to(&mut output)?;
        Ok(output.into_inner())
    }

    /// Serialises the bytes that [`Presentation::save`] writes to `path`.
    pub fn to_bytes_for_path<P: AsRef<Path>>(&self, path: P) -> Result<Vec<u8>> {
        let Some(target) = PresentationPackageClass::from_path(path.as_ref()) else {
            return self.to_bytes();
        };
        let current = self.package_class()?;
        if target == current {
            return self.to_bytes();
        }
        let carries_vba = self
            .package
            .get_part_rels(&self.presentation_part)
            .is_some_and(|relationships| {
                relationships.get_by_type(rel_types::VBA_PROJECT).is_some()
            });
        if carries_vba && !target.is_macro_enabled() {
            return Err(invalid_presentation_mutation(
                "save",
                format!(
                    "{} cannot be written: the presentation carries a VBA project and the extension names a macro-free class, save it as .pptm, .potm, or .ppsm",
                    path.as_ref().display()
                ),
            ));
        }
        self.to_bytes_as(target)
    }

    /// Serialises the current presentation through the fixed Agile write profile.
    #[cfg(all(feature = "agile-encryption", not(target_arch = "wasm32")))]
    pub fn to_encrypted_bytes(&self, password: &str) -> Result<Vec<u8>> {
        let package_signatures_invalidated = self.package_signatures_invalidated
            || self.retained_package_signature_would_be_invalidated()?;
        let mut package = self.staged_package(true)?;
        embedded::persist_invalidated_package_signature(
            &mut package,
            package_signatures_invalidated,
        )?;
        let mut output = Vec::new();
        package.write_encrypted_to(&mut output, password)?;
        Ok(output)
    }

    /// Verifies signatures against the current staged presentation package.
    ///
    /// Cryptographic validity authenticates the embedded certificate's key. It
    /// does not establish certificate-chain trust, which remains caller policy.
    #[cfg(feature = "digital-signatures")]
    pub fn verify_signatures(&self) -> Result<Vec<SignatureReport>> {
        Ok(self.staged_package(true)?.verify_signatures()?)
    }

    /// Signs the current staged presentation with the strict RSA-SHA256 profile.
    ///
    /// The private key must be PKCS#8 DER and the certificate must be X.509 DER.
    /// The live package changes only after the shared signer verifies the cloned
    /// candidate with complete coverage. Certificate-chain trust remains caller
    /// policy.
    #[cfg(all(feature = "digital-signatures", not(target_arch = "wasm32")))]
    pub fn sign(
        &mut self,
        private_key_pkcs8_der: &[u8],
        certificate_der: &[u8],
    ) -> Result<SignatureReport> {
        let mut candidate = self.staged_package(true)?;
        let report = candidate.sign(private_key_pkcs8_der, certificate_der)?;
        if !report.cryptographically_valid || !report.coverage_complete {
            return Err(OpcError::SignatureCreationFailed(
                "staged presentation signature did not verify with complete coverage".to_owned(),
            )
            .into());
        }
        self.package = candidate;
        Ok(report)
    }

    /// Resolves and lays out the current presentation with deterministic fonts.
    #[cfg(feature = "render")]
    pub fn render_deterministic(&self) -> Result<(RenderInput, LayoutResult)> {
        assemble_render_input(&self.staged_package(false)?)
    }

    /// Evaluates and renders one slide at an explicit slide-local timestamp.
    ///
    /// `outgoing_slide_index` supplies the zero-based previous slide for a
    /// two-slide transition. Static rendering remains on its independent path.
    #[cfg(feature = "render")]
    pub fn render_timeline_deterministic(
        &self,
        slide_index: usize,
        position: TimelinePosition,
        outgoing_slide_index: Option<usize>,
    ) -> Result<DeterministicTimelineFrame> {
        self.render_timeline_deterministic_inner(slide_index, position, outgoing_slide_index, None)
            .map(|(frame, _)| frame)
    }

    /// Evaluates one deterministic page and synchronized media playback state.
    #[cfg(feature = "render")]
    pub fn render_media_timeline_deterministic(
        &self,
        slide_index: usize,
        position: TimelinePosition,
        outgoing_slide_index: Option<usize>,
        fallback_policy: MediaFallbackPolicy,
    ) -> Result<DeterministicMediaTimelineFrame> {
        let (frame, media) = self.render_timeline_deterministic_inner(
            slide_index,
            position,
            outgoing_slide_index,
            Some(fallback_policy),
        )?;
        Ok(DeterministicMediaTimelineFrame { frame, media })
    }

    #[cfg(feature = "render")]
    fn render_timeline_deterministic_inner(
        &self,
        slide_index: usize,
        position: TimelinePosition,
        outgoing_slide_index: Option<usize>,
        fallback_policy: Option<MediaFallbackPolicy>,
    ) -> Result<(DeterministicTimelineFrame, Vec<EvaluatedMediaState>)> {
        let package = self.staged_package(false)?;
        let mut assembly = prepare_render_context(&package, fallback_policy.is_some())?;
        render_prepared_timeline_request(
            &mut assembly,
            TimelineRequest {
                slide_index,
                outgoing_slide_index,
                position,
                fallback_policy,
            },
            self.slides.len(),
            false,
        )
    }

    /// Renders the current presentation to a complete deterministic PDF.
    #[cfg(feature = "render")]
    pub fn to_pdf_deterministic(&self) -> Result<Vec<u8>> {
        let (_, layout) = self.render_deterministic()?;
        Ok(oxml_pdf::render_to_pdf(&layout))
    }

    /// Renders one deterministic slide PNG at `dpi` by zero-based index.
    #[cfg(feature = "render")]
    pub fn slide_png_deterministic(&self, slide_index: usize, dpi: f64) -> Result<Option<Vec<u8>>> {
        let (_, layout) = self.render_deterministic()?;
        render_export_png(&layout, slide_index, dpi)
    }

    /// Renders one deterministic PNG per slide at `dpi`.
    #[cfg(feature = "render")]
    pub fn slide_pngs_deterministic(&self, dpi: f64) -> Result<Vec<Vec<u8>>> {
        let (_, layout) = self.render_deterministic()?;
        render_export_pngs(&layout, dpi)
    }

    /// Lays out the text of every text-bearing slide shape with deterministic fonts.
    ///
    /// Frames come in slide order, then draw order, from the same resolved
    /// slides and line breaker as [`Self::to_pdf_deterministic`] and
    /// [`Self::slide_png_deterministic`]. Only shapes in each slide's own shape
    /// tree whose text draws a visible character are reported, so layout and
    /// master shapes, SmartArt and table cells are left out. `width_factor`
    /// scales every usable width before line breaking, so `1.0` reports the
    /// rendered layout and `0.95` asks whether the text fits a narrower frame.
    #[cfg(feature = "render")]
    pub fn text_layout_deterministic(&self, width_factor: f64) -> Result<Vec<TextFrameLayout>> {
        if !width_factor.is_finite() || width_factor <= 0.0 {
            return Err(render_failure(format!(
                "text layout width factor must be finite and positive, found {width_factor}"
            )));
        }
        let package = self.staged_package(false)?;
        let mut assembly = prepare_render_context(&package, false)?;
        let mut frames = Vec::new();
        for (slide_index, ((prepared, slide), directions)) in assembly
            .slides
            .iter()
            .zip(&assembly.input.slides)
            .zip(&assembly.text_directions)
            .enumerate()
        {
            let context = ResolveCtx::new(
                &prepared.theme,
                render_effective_color_map(&prepared.master, &prepared.layout, &prepared.slide),
                &prepared.master,
                &prepared.layout,
                &prepared.slide,
                &assembly.default_text_style,
            );
            let sources = context
                .flatten()
                .into_iter()
                .filter(|item| render_source_shape_has_bounds(&context, item))
                .collect::<Vec<_>>();
            if sources.len() != slide.shapes.len() {
                return Err(render_failure(format!(
                    "{}: text layout source {}, resolved {}",
                    prepared.slide_part,
                    sources.len(),
                    slide.shapes.len()
                )));
            }
            for (shape_index, ((item, shape), smartart_clip)) in sources
                .into_iter()
                .zip(&slide.shapes)
                .zip(&prepared.smartart_clips)
                .enumerate()
            {
                let (
                    FlattenedItem::Shape {
                        source: FlattenedSource::Slide,
                        child,
                        ..
                    },
                    ResolvedContent::Text(text),
                    None,
                ) = (item, &shape.content, smartart_clip)
                else {
                    continue;
                };
                let layout = rpptx_render::layout_shape_text(
                    shape,
                    text,
                    &mut assembly.font_manager,
                    slide_index + 1,
                    directions.get(shape_index).map_or(&[], Vec::as_slice),
                    width_factor,
                )
                .map_err(|error| {
                    render_failure(format!(
                        "slide {}, shape id {}: {error}",
                        slide_index + 1,
                        child
                            .non_visual_id()
                            .map_or_else(|| "unknown".to_owned(), |id| id.to_string())
                    ))
                })?;
                if layout.lines.iter().all(|line| line.text.trim().is_empty()) {
                    continue;
                }
                frames.push(TextFrameLayout {
                    slide_index,
                    shape_id: child.non_visual_id(),
                    name: child.non_visual_name(),
                    autofit: match text.autofit {
                        ResolvedAutofit::None => AutofitMode::None,
                        ResolvedAutofit::Normal { .. } => AutofitMode::Normal,
                        ResolvedAutofit::Shape => AutofitMode::Shape,
                    },
                    layout,
                });
            }
        }
        Ok(frames)
    }

    /// Render the presentation to the selected archival PDF profile.
    #[cfg(feature = "render")]
    pub fn to_pdfa_deterministic(&self, profile: oxml_pdf::PdfConformance) -> Result<Vec<u8>> {
        let (_, layout) = self.render_deterministic()?;
        Ok(oxml_pdf::render_to_pdf_with_options(
            &layout,
            oxml_pdf::PdfOptions::new(profile),
        )?)
    }

    /// Renders one deterministic speaker-notes page per source slide.
    #[cfg(feature = "render")]
    pub fn to_notes_pdf_deterministic(&self) -> Result<Vec<u8>> {
        let package = self.staged_package(false)?;
        let layout = render_notes_pages(&package)?;
        Ok(oxml_pdf::render_to_pdf(&layout))
    }

    /// Renders one deterministic PNG per speaker-notes page at `dpi`.
    #[cfg(feature = "render")]
    pub fn notes_page_pngs_deterministic(&self, dpi: f64) -> Result<Vec<Vec<u8>>> {
        let package = self.staged_package(false)?;
        let layout = render_notes_pages(&package)?;
        render_export_pngs(&layout, dpi)
    }

    /// Renders deterministic audience handouts in the selected arrangement.
    #[cfg(feature = "render")]
    pub fn to_handout_pdf_deterministic(&self, layout: HandoutLayout) -> Result<Vec<u8>> {
        let package = self.staged_package(false)?;
        let layout = render_handout_pages(&package, layout)?;
        Ok(oxml_pdf::render_to_pdf(&layout))
    }

    /// Renders one deterministic PNG per audience-handout page at `dpi`.
    #[cfg(feature = "render")]
    pub fn handout_page_pngs_deterministic(
        &self,
        layout: HandoutLayout,
        dpi: f64,
    ) -> Result<Vec<Vec<u8>>> {
        let package = self.staged_package(false)?;
        let layout = render_handout_pages(&package, layout)?;
        render_export_pngs(&layout, dpi)
    }

    fn staged_package(&self, preserve_unchanged_modelled_parts: bool) -> Result<OpcPackage> {
        #[cfg(all(feature = "render", test))]
        STAGED_PACKAGE_COUNT.with(|count| count.set(count.get() + 1));
        if self.core_properties_dirty
            && self
                .package
                .package_rels
                .get_by_type(rel_types::CORE_PROPERTIES)
                .is_none()
            && self.package.contains_part(&self.core_properties_part_name)
        {
            return Err(Error::CorePropertiesPartCollision {
                part_name: self.core_properties_part_name.clone(),
            });
        }
        let mut package = self.package.clone();
        self.media_store.write_new_parts(&mut package);
        if let Some(properties) = self
            .core_properties
            .as_ref()
            .filter(|_| self.core_properties_dirty)
        {
            let xml = properties.to_xml().map_err(|error| Error::MalformedPart {
                part_name: self.core_properties_part_name.clone(),
                message: error.to_string(),
            })?;
            package.set_part(&self.core_properties_part_name, xml);
            package.content_types.add_override(
                &self.core_properties_part_name,
                content_types::CORE_PROPERTIES,
            );
            if package
                .package_rels
                .get_by_type(rel_types::CORE_PROPERTIES)
                .is_none()
            {
                let target = self
                    .core_properties_part_name
                    .strip_prefix('/')
                    .unwrap_or(&self.core_properties_part_name);
                package.package_rels.add(rel_types::CORE_PROPERTIES, target);
            }
        }
        let presentation_changed = self
            .package
            .get_part(&self.presentation_part)
            .and_then(|xml| CT_Presentation::from_xml(xml).ok())
            .as_ref()
            != Some(&self.presentation);
        if !preserve_unchanged_modelled_parts || presentation_changed {
            package.set_part(
                &self.presentation_part,
                self.presentation
                    .to_xml()
                    .map_err(|error| Error::MalformedPart {
                        part_name: self.presentation_part.clone(),
                        message: error.to_string(),
                    })?,
            );
        }
        if let Some(part_name) = self
            .comment_authors_part
            .as_ref()
            .filter(|_| self.comment_authors_dirty)
        {
            package.set_part(
                part_name,
                self.comment_authors
                    .to_xml()
                    .map_err(|error| Error::MalformedPart {
                        part_name: part_name.clone(),
                        message: error.to_string(),
                    })?,
            );
            package
                .content_types
                .add_override(part_name, content_types::POWERPOINT_AUTHORS);
        }
        if let (Some(part_name), Some(master)) = (
            self.notes_master_part
                .as_ref()
                .filter(|_| self.notes_master_dirty),
            &self.notes_master,
        ) {
            package.set_part(
                part_name,
                master.to_xml().map_err(|error| Error::MalformedPart {
                    part_name: part_name.clone(),
                    message: error.to_string(),
                })?,
            );
        }
        if let (Some(part_name), Some(master)) = (
            self.handout_master_part
                .as_ref()
                .filter(|_| self.handout_master_dirty),
            &self.handout_master,
        ) {
            package.set_part(
                part_name,
                master.to_xml().map_err(|error| Error::MalformedPart {
                    part_name: part_name.clone(),
                    message: error.to_string(),
                })?,
            );
        }
        for record in &self.slides {
            let slide_changed = self
                .package
                .get_part(&record.part_name)
                .and_then(|xml| CT_Slide::from_xml(xml).ok())
                .as_ref()
                != Some(&record.slide);
            if !preserve_unchanged_modelled_parts || slide_changed {
                package.set_part(
                    &record.part_name,
                    record
                        .slide
                        .to_xml()
                        .map_err(|error| Error::MalformedPart {
                            part_name: record.part_name.clone(),
                            message: error.to_string(),
                        })?,
                );
            }
            if let Some(notes) = &record.notes {
                let notes_changed = self
                    .package
                    .get_part(&notes.part_name)
                    .and_then(|xml| CT_NotesSlide::from_xml(xml).ok())
                    .as_ref()
                    != Some(&notes.notes);
                if !preserve_unchanged_modelled_parts || notes_changed {
                    package.set_part(
                        &notes.part_name,
                        notes.notes.to_xml().map_err(|error| Error::MalformedPart {
                            part_name: notes.part_name.clone(),
                            message: error.to_string(),
                        })?,
                    );
                }
            }
            if let Some(comments) = record.comments.as_ref().filter(|comments| comments.dirty) {
                package.set_part(
                    &comments.part_name,
                    comments
                        .comments
                        .to_xml()
                        .map_err(|error| Error::MalformedPart {
                            part_name: comments.part_name.clone(),
                            message: error.to_string(),
                        })?,
                );
                package
                    .content_types
                    .add_override(&comments.part_name, content_types::POWERPOINT_COMMENTS);
            }
        }
        Ok(package)
    }

    /// Saves the deterministic package bytes to a path.
    ///
    /// The bytes are staged in a synced sibling file and renamed over `path`
    /// through [`oxml_opc::write_atomic_file`], so a failed save leaves an
    /// existing file as it was. A symbolic link at `path` is kept and the file
    /// it names is replaced, and on Unix that file keeps its permission bits.
    ///
    /// A `.pptx`, `.pptm`, `.potx`, `.potm`, `.ppsx`, or `.ppsm` extension
    /// selects the main part content type, so a template saved as `.pptx`
    /// declares a presentation. Any other extension keeps the opened class,
    /// and `save_encrypted` ignores the extension. When the class changes, a
    /// presentation part that carries a VBA project
    /// cannot be saved under a macro-free extension, because the project
    /// would remain in a file that claims to carry none.
    /// [`Presentation::save_as_package_class`] performs that conversion
    /// explicitly and keeps the VBA part.
    pub fn save<P: AsRef<Path>>(&self, path: P) -> Result<()> {
        debug_assert!(
            self.validate().is_empty(),
            "invalid presentation at path save boundary: {:?}",
            self.validate()
        );
        write_atomic_file(
            path.as_ref(),
            &self.to_bytes_for_path(path.as_ref())?,
            "rpptx",
            "invalid PowerPoint package file name",
            "could not allocate PowerPoint package save staging file",
        )
        .map_err(OpcError::from)?;
        Ok(())
    }

    /// Saves a password-protected presentation through an atomic path publication.
    #[cfg(all(feature = "agile-encryption", not(target_arch = "wasm32")))]
    pub fn save_encrypted<P: AsRef<Path>>(&self, path: P, password: &str) -> Result<()> {
        let bytes = self.to_encrypted_bytes(password)?;
        write_atomic_file(
            path.as_ref(),
            &bytes,
            "rpptx",
            "invalid file name",
            "could not allocate encrypted-save staging file",
        )
        .map_err(OpcError::from)?;
        Ok(())
    }

    /// Saves a slideshow package without changing later ordinary saves.
    pub fn save_as_show<P: AsRef<Path>>(&self, path: P) -> Result<()> {
        self.save_as_package_class(path, PresentationPackageClass::Slideshow)
    }

    /// Saves an output copy with the selected modern package class.
    ///
    /// The file is replaced the way [`Presentation::save`] replaces it.
    pub fn save_as_package_class<P: AsRef<Path>>(
        &self,
        path: P,
        class: PresentationPackageClass,
    ) -> Result<()> {
        write_atomic_file(
            path.as_ref(),
            &self.to_bytes_as(class)?,
            "rpptx",
            "invalid PowerPoint package file name",
            "could not allocate PowerPoint package save staging file",
        )
        .map_err(OpcError::from)?;
        Ok(())
    }

    /// Returns every detected package and PresentationML invariant violation.
    pub fn validate(&self) -> Vec<ValidationIssue> {
        let mut duplicate_shape_ids = Vec::new();
        let mut out_of_range_slide_ids = Vec::new();
        let mut duplicate_slide_ids = Vec::new();
        let mut missing_content_types = Vec::new();
        let mut dangling_relationships = Vec::new();
        let mut unreachable_targets = Vec::new();
        let mut empty_text_bodies = Vec::new();
        let mut duplicate_placeholder_indices = Vec::new();
        let mut orphan_media = Vec::new();
        let mut custom_show_references = Vec::new();
        let mut missing_layout_relationships = Vec::new();
        let mut missing_theme_relationships = Vec::new();

        for (slide_index, record) in self.slides.iter().enumerate() {
            let mut shape_ids = record
                .slide
                .common_slide_data
                .shape_tree
                .non_visual_id()
                .into_iter()
                .collect();
            let mut placeholder_indices = HashSet::new();
            let mut shape_index = 0usize;
            validate_shape_children(
                slide_index,
                &record.slide.common_slide_data.shape_tree.children,
                &mut shape_index,
                &mut shape_ids,
                &mut placeholder_indices,
                &mut duplicate_shape_ids,
                &mut empty_text_bodies,
                &mut duplicate_placeholder_indices,
            );
        }

        let mut slide_ids = HashSet::new();
        for (slide_index, slide) in self.presentation.slide_ids.iter().enumerate() {
            if !(256..=2_147_483_647).contains(&slide.id) {
                out_of_range_slide_ids.push(ValidationIssue::SlideIdOutOfRange {
                    slide: slide_index,
                    id: slide.id,
                });
            }
            if !slide_ids.insert(slide.id) {
                duplicate_slide_ids.push(ValidationIssue::DuplicateSlideId { id: slide.id });
            }
        }

        let mut part_names = self.package.parts.keys().collect::<Vec<_>>();
        part_names.sort_unstable();
        for part_name in part_names {
            if !part_has_content_type(&self.package, part_name) {
                missing_content_types.push(ValidationIssue::MissingContentTypeOverride {
                    part: part_name.clone(),
                });
            }
        }

        let mut incoming_targets = HashSet::new();
        validate_relationship_scope(
            &self.package,
            "/",
            None,
            &self.package.package_rels,
            &mut incoming_targets,
            &mut dangling_relationships,
            &mut unreachable_targets,
        );
        let mut relationship_sources = self
            .package
            .parts
            .keys()
            .chain(self.package.part_rels.keys())
            .map(String::as_str)
            .collect::<Vec<_>>();
        relationship_sources.sort_unstable();
        relationship_sources.dedup();
        let empty_relationships = Relationships::new();
        for source_part in relationship_sources {
            let relationships = self
                .package
                .get_part_rels(source_part)
                .unwrap_or(&empty_relationships);
            let source_xml = self.current_part_xml(source_part);
            validate_relationship_scope(
                &self.package,
                source_part,
                source_xml.as_deref(),
                relationships,
                &mut incoming_targets,
                &mut dangling_relationships,
                &mut unreachable_targets,
            );
        }

        let mut media_parts = self
            .package
            .parts
            .keys()
            .filter(|part| part.to_ascii_lowercase().starts_with("/ppt/media/"))
            .collect::<Vec<_>>();
        media_parts.sort_unstable();
        for part in media_parts {
            if !incoming_targets.contains(&part.to_ascii_lowercase()) {
                orphan_media.push(ValidationIssue::OrphanMedia { part: part.clone() });
            }
        }

        let presentation_relationships = self.package.get_part_rels(&self.presentation_part);
        let slide_relationship_ids = self
            .presentation
            .slide_ids
            .iter()
            .map(|slide| slide.relationship_id.as_str())
            .collect::<HashSet<_>>();
        if let Some(xml) = self.current_part_xml(&self.presentation_part)
            && let Ok(references) = custom_show_relationship_ids(&xml)
        {
            for relationship_id in references {
                if !slide_relationship_ids.contains(relationship_id.as_str()) {
                    custom_show_references.push(ValidationIssue::CustomShowReference {
                        slide_id: numeric_relationship_id(&relationship_id),
                    });
                }
            }
        }

        for (slide_index, slide) in self.slides.iter().enumerate() {
            let count = self
                .package
                .get_part_rels(&slide.part_name)
                .map(|relationships| {
                    relationships
                        .items
                        .iter()
                        .filter(|relationship| {
                            relationship.rel_type == rel_types::SLIDE_LAYOUT
                                && !relationship_is_external(relationship)
                        })
                        .count()
                })
                .unwrap_or_default();
            if count != 1 {
                missing_layout_relationships
                    .push(ValidationIssue::MissingLayoutRel { slide: slide_index });
            }
        }

        for (master_index, master) in self.presentation.slide_master_ids.iter().enumerate() {
            let Some(relationship) = presentation_relationships
                .and_then(|relationships| relationships.get_by_id(&master.relationship_id))
            else {
                missing_theme_relationships.push(ValidationIssue::MissingThemeRel {
                    master: master_index,
                });
                continue;
            };
            let master_part =
                OpcPackage::resolve_rel_target(&self.presentation_part, &relationship.target);
            let has_theme = self
                .package
                .get_part_rels(&master_part)
                .is_some_and(|relationships| {
                    relationships.items.iter().any(|relationship| {
                        relationship.rel_type == rel_types::THEME
                            && !relationship_is_external(relationship)
                    })
                });
            if !has_theme {
                missing_theme_relationships.push(ValidationIssue::MissingThemeRel {
                    master: master_index,
                });
            }
        }

        duplicate_shape_ids
            .into_iter()
            .chain(out_of_range_slide_ids)
            .chain(duplicate_slide_ids)
            .chain(missing_content_types)
            .chain(dangling_relationships)
            .chain(unreachable_targets)
            .chain(empty_text_bodies)
            .chain(duplicate_placeholder_indices)
            .chain(orphan_media)
            .chain(custom_show_references)
            .chain(missing_layout_relationships)
            .chain(missing_theme_relationships)
            .collect()
    }

    fn current_part_xml(&self, part_name: &str) -> Option<Vec<u8>> {
        if part_name.eq_ignore_ascii_case(&self.presentation_part) {
            return self.presentation.to_xml().ok();
        }
        if let Some(slide) = self
            .slides
            .iter()
            .find(|slide| slide.part_name.eq_ignore_ascii_case(part_name))
        {
            return slide.slide.to_xml().ok();
        }
        if let Some(notes) = self
            .slides
            .iter()
            .filter_map(|slide| slide.notes.as_ref())
            .find(|notes| notes.part_name.eq_ignore_ascii_case(part_name))
        {
            return notes.notes.to_xml().ok();
        }
        self.package.get_part(part_name).map(<[u8]>::to_vec)
    }

    /// Returns the number of slides in presentation order.
    pub fn len(&self) -> usize {
        self.slides.len()
    }

    /// Returns whether the presentation has no slides.
    pub fn is_empty(&self) -> bool {
        self.slides.is_empty()
    }

    /// Returns one slide by zero-based presentation index.
    pub fn slide(&self, index: usize) -> Option<SlideRef<'_>> {
        self.slides.get(index).map(slide_ref)
    }

    /// Returns one slide for in-place mutation by zero-based index.
    pub fn slide_mut(&mut self, index: usize) -> Option<SlideMut<'_>> {
        self.slides.get_mut(index).map(slide_mut)
    }

    /// Returns the shapes of one slide, or of one group on it, for adding
    /// shapes.
    ///
    /// `group` holds one z-order index per nesting level, from the slide's
    /// own shapes down to the group, and is empty for the slide's own shapes.
    /// A missing slide or a path that does not end at a group returns `None`.
    pub fn shapes_mut(&mut self, slide_index: usize, group: &[usize]) -> Option<ShapesMut<'_>> {
        let tree = &mut self
            .slides
            .get_mut(slide_index)?
            .slide
            .common_slide_data
            .shape_tree;
        if !group.is_empty() {
            group_at_mut(&mut tree.children, group)?;
        }
        Some(ShapesMut {
            presentation: self,
            slide_index,
            group: group.to_vec(),
        })
    }

    /// Iterates slides in `p:sldIdLst` order.
    pub fn slides(&self) -> impl ExactSizeIterator<Item = SlideRef<'_>> {
        self.slides.iter().map(slide_ref)
    }

    /// Replaces literal text across regular DrawingML runs in every editable shape.
    ///
    /// Replacement text retains the formatting and opaque run content of the
    /// first matched run. Unmatched prefixes and suffixes stay in their
    /// original runs. Groups and table cells are traversed recursively, while
    /// preserved alternate-content branches remain untouched.
    pub fn replace_text(&mut self, placeholder: &str, value: &str) -> usize {
        if placeholder.is_empty() {
            return 0;
        }
        // Stored as the run text setters store it, so it can be saved.
        let value = &oxml_drawing::text::escape_invalid_xml_characters(value);
        self.slides
            .iter_mut()
            .map(|record| replace_text_in_slide_record(record, placeholder, value, true))
            .sum()
    }

    /// Replaces literal slide and speaker-note text on a staged candidate.
    ///
    /// The candidate is serialized before it replaces the live presentation,
    /// so a malformed retained subtree cannot publish a partial mutation.
    pub fn try_replace_text(&mut self, placeholder: &str, value: &str) -> Result<usize> {
        require_replacement_placeholder(placeholder)?;
        let mut candidate = self.clone();
        let count = candidate.replace_text(placeholder, value);
        candidate.staged_package(false)?;
        *self = candidate;
        Ok(count)
    }

    /// Replaces literal text on one slide, and in its speaker notes when
    /// `notes` is true, and returns the count.
    ///
    /// Matching, formatting and traversal are those of
    /// [`Presentation::try_replace_text`], restricted to that slide. Other
    /// slides and their notes are never touched. The replacement runs on a
    /// copy of the slide and its notes, which must serialize before the copy
    /// replaces them. With `expect`, a different count also leaves the slide
    /// unchanged and is returned for the caller to compare.
    pub fn try_replace_slide_text(
        &mut self,
        slide_index: usize,
        placeholder: &str,
        value: &str,
        notes: bool,
        expect: Option<usize>,
    ) -> Result<usize> {
        self.require_slide_index(slide_index)?;
        require_replacement_placeholder(placeholder)?;
        let mut candidate = self.slides[slide_index].clone();
        let value = &oxml_drawing::text::escape_invalid_xml_characters(value);
        let count = replace_text_in_slide_record(&mut candidate, placeholder, value, notes);
        if expect.is_some_and(|expected| expected != count) {
            return Ok(count);
        }
        candidate
            .slide
            .to_xml()
            .map_err(|error| Error::MalformedPart {
                part_name: candidate.part_name.clone(),
                message: error.to_string(),
            })?;
        if let Some(record) = candidate.notes.as_ref().filter(|_| notes) {
            record
                .notes
                .to_xml()
                .map_err(|error| Error::MalformedPart {
                    part_name: record.part_name.clone(),
                    message: error.to_string(),
                })?;
        }
        self.slides[slide_index] = candidate;
        Ok(count)
    }

    /// Returns the optional slide dimensions in EMUs.
    pub fn slide_size(&self) -> Option<(Emu, Emu)> {
        self.presentation
            .slide_size
            .as_ref()
            .map(|size| (size.cx, size.cy))
    }

    /// Replaces only the slide dimensions after validating both values.
    pub fn set_slide_size(&mut self, width: Emu, height: Emu) -> Result<()> {
        self.presentation
            .set_slide_size(width, height)
            .map_err(|error| Error::InvalidPresentationMutation {
                operation: "set slide size",
                message: error.to_string(),
            })
    }

    /// Returns shared package core properties without dirtying their source part.
    pub fn core_properties(&self) -> Option<&CoreProperties> {
        self.core_properties.as_ref()
    }

    /// Returns mutable shared core properties, materializing the package model.
    pub fn core_properties_mut(&mut self) -> &mut CoreProperties {
        self.core_properties_dirty = true;
        self.core_properties
            .get_or_insert_with(CoreProperties::default)
    }

    /// Returns the number of layouts reachable through presentation masters.
    pub fn layout_count(&self) -> usize {
        self.layouts.len()
    }

    /// Returns a layout's optional `p:cSld` name by zero-based index.
    pub fn layout_name(&self, index: usize) -> Option<&str> {
        self.layouts
            .get(index)
            .and_then(|record| record.layout.common_slide_data.name.as_deref())
    }

    /// Returns the zero-based index of the layout one slide uses.
    ///
    /// The slide's internal layout relationship must target a layout that the
    /// presentation masters reach.
    pub fn slide_layout_index(&self, slide_index: usize) -> Option<usize> {
        let slide_part = &self.slides.get(slide_index)?.part_name;
        let layout_part =
            related_internal_part(&self.package, slide_part, rel_types::SLIDE_LAYOUT).ok()??;
        self.layouts
            .iter()
            .position(|layout| layout.part_name.eq_ignore_ascii_case(&layout_part))
    }

    /// Returns one slide shape's effective offset and extent in EMU.
    ///
    /// `shape_path` holds the shape's index in the slide's shape tree, then
    /// its index within each enclosing group. The shape's own transform wins
    /// whole. A placeholder without one inherits the transform of its layout
    /// placeholder, then of that placeholder's master counterpart, as
    /// rendering resolves it. The result is `(left, top, width, height)`, and
    /// a missing offset reads as zero, as rendering draws it. `None` means
    /// the resolved transform has no extent, so the shape is not drawn.
    /// [`ShapeRef::position`] and [`ShapeRef::size`] keep reporting only the
    /// shape's own values. An empty `shape_path` or one that names no shape
    /// is [`Error::InvalidShapeMutation`], as a missing picture is for
    /// [`Self::picture_image`].
    #[cfg(feature = "render")]
    pub fn effective_geometry(
        &self,
        slide_index: usize,
        shape_path: &[usize],
    ) -> Result<Option<(Emu, Emu, Emu, Emu)>> {
        let child = self
            .shape_at(slide_index, shape_path, "read effective geometry")?
            .child;
        let own = shape_transform(child);
        let inherited = match (own, shape_placeholder(child)) {
            (None, Some(placeholder)) => self.inherited_transform(slide_index, placeholder)?,
            _ => None,
        };
        Ok(own.or(inherited.as_ref()).and_then(|transform| {
            let offset = transform.offset.unwrap_or_default();
            let extent = transform.extent?;
            Some((offset.x, offset.y, extent.cx, extent.cy))
        }))
    }

    /// Writes a placeholder's inherited geometry onto the shape itself.
    ///
    /// `shape_path` is read as in [`Self::effective_geometry`]. A placeholder
    /// whose own transform lacks an offset or an extent takes the missing
    /// part from the transform it inherits, and one without a transform
    /// copies the inherited transform whole, rotation and flips included.
    /// [`ShapeMut::set_position`] or [`ShapeMut::set_size`] then changes one
    /// pair while the shape keeps the other, and [`ShapeMut::set_rotation`]
    /// changes the angle while it keeps both, where an incomplete transform
    /// would make rendering skip the shape. Shapes that are not placeholders,
    /// placeholders with a complete transform and placeholders that inherit
    /// none are left unchanged.
    #[cfg(feature = "render")]
    pub fn materialize_geometry(&mut self, slide_index: usize, shape_path: &[usize]) -> Result<()> {
        const OPERATION: &str = "materialize geometry";
        let child = self.shape_at(slide_index, shape_path, OPERATION)?.child;
        let own = shape_transform(child);
        if own.is_some_and(|own| own.offset.is_some() && own.extent.is_some()) {
            return Ok(());
        }
        let Some(placeholder) = shape_placeholder(child) else {
            return Ok(());
        };
        let Some(inherited) = self.inherited_transform(slide_index, placeholder)? else {
            return Ok(());
        };
        let mut materialized = own.cloned().unwrap_or_else(|| inherited.clone());
        materialized.offset = materialized.offset.or(inherited.offset);
        materialized.extent = materialized.extent.or(inherited.extent);
        let mut indices = shape_path.iter();
        let mut shape = indices
            .next()
            .and_then(|index| slide_mut(&mut self.slides[slide_index]).into_shape_mut(*index));
        for index in indices {
            shape = shape.and_then(|shape| shape.into_child_mut(*index));
        }
        let mut shape = shape.ok_or_else(|| {
            invalid_shape_mutation(
                OPERATION,
                format!("slide {slide_index} has no editable shape at path {shape_path:?}"),
            )
        })?;
        *shape.transform_mut(OPERATION)? = materialized;
        Ok(())
    }

    /// Changes the layout one slide uses, by zero-based layout index.
    ///
    /// The slide's layout relationship is retargeted, so its placeholders
    /// inherit from the new layout and that layout's master. A placeholder
    /// without a transform of its own that the new layout does not place
    /// first receives the transform it inherited from the old layout, so it
    /// stays where it was drawn. Placeholders of the new layout that the
    /// slide lacks are not added. The layout may belong to another master,
    /// whose theme and text styles the slide then follows. The change is
    /// staged and publishes only after the staged package reopens.
    #[cfg(feature = "render")]
    pub fn set_slide_layout(&mut self, slide_index: usize, layout_index: usize) -> Result<()> {
        const OPERATION: &str = "set slide layout";
        self.require_slide_index(slide_index)?;
        if layout_index >= self.layouts.len() {
            return Err(Error::UnknownLayoutIndex {
                index: layout_index,
                layout_count: self.layouts.len(),
            });
        }
        let current = self.slide_layout_index(slide_index).ok_or_else(|| {
            invalid_slide_mutation(
                OPERATION,
                format!("slide {slide_index} has no layout reachable through the slide masters"),
            )
        })?;
        if current == layout_index {
            return Ok(());
        }
        let (old_layout, old_master) = self.layout_and_master(current)?;
        let (new_layout, new_master) = self.layout_and_master(layout_index)?;
        let mut staged = self.clone();
        keep_unplaced_placeholder_transforms(
            &mut staged.slides[slide_index]
                .slide
                .common_slide_data
                .shape_tree
                .children,
            (old_layout, &old_master),
            (new_layout, &new_master),
        );
        let slide_part = &staged.slides[slide_index].part_name;
        let target = relative_part_target(slide_part, &staged.layouts[layout_index].part_name);
        staged
            .package
            .get_part_rels_mut(slide_part)
            .and_then(|relationships| {
                relationships
                    .items
                    .iter_mut()
                    .find(|relationship| relationship.rel_type == rel_types::SLIDE_LAYOUT)
            })
            .ok_or_else(|| {
                invalid_slide_mutation(
                    OPERATION,
                    format!("{slide_part} has no layout relationship"),
                )
            })?
            .target = target;
        self.commit_candidate(staged)
    }

    #[cfg(feature = "render")]
    fn shape_at(
        &self,
        slide_index: usize,
        shape_path: &[usize],
        operation: &'static str,
    ) -> Result<ShapeRef<'_>> {
        self.require_slide_index(slide_index)?;
        let slide = slide_ref(&self.slides[slide_index]);
        let mut indices = shape_path.iter();
        let mut shape = indices.next().and_then(|index| slide.shape(*index));
        for index in indices {
            shape = shape.and_then(|shape| shape.child(*index));
        }
        shape.ok_or_else(|| {
            invalid_shape_mutation(
                operation,
                format!("slide {slide_index} has no shape at path {shape_path:?}"),
            )
        })
    }

    /// Returns the transform a slide placeholder inherits through its layout.
    #[cfg(feature = "render")]
    fn inherited_transform(
        &self,
        slide_index: usize,
        placeholder: &CT_Placeholder,
    ) -> Result<Option<CT_Transform2D>> {
        let Some(layout_index) = self.slide_layout_index(slide_index) else {
            return Ok(None);
        };
        let (layout, master) = self.layout_and_master(layout_index)?;
        Ok(inherited_xfrm(placeholder, layout, &master).cloned())
    }

    #[cfg(feature = "render")]
    fn layout_and_master(&self, layout_index: usize) -> Result<(&CT_SlideLayout, CT_SlideMaster)> {
        let record = &self.layouts[layout_index];
        let master_part =
            related_internal_part(&self.package, &record.part_name, rel_types::SLIDE_MASTER)?
                .ok_or_else(|| Error::MalformedPart {
                    part_name: record.part_name.clone(),
                    message: "layout has no slide master relationship".to_owned(),
                })?;
        let master = CT_SlideMaster::from_xml(required_part(&self.package, &master_part)?)
            .map_err(|error| Error::MalformedPart {
                part_name: master_part,
                message: error.to_string(),
            })?;
        Ok((&record.layout, master))
    }

    /// Returns modern PowerPoint comment authors in producer order.
    pub fn comment_authors(&self) -> &[CommentAuthor] {
        &self.comment_authors.authors
    }

    /// Adds a caller-identified modern comment author atomically.
    pub fn add_comment_author(&mut self, author: CommentAuthor) -> Result<()> {
        if self
            .comment_authors
            .authors
            .iter()
            .any(|existing| existing.id == author.id)
        {
            return Err(invalid_presentation_mutation(
                "add comment author",
                format!("duplicate author id {}", author.id),
            ));
        }
        let mut staged = self.clone();
        if staged.comment_authors_part.is_none() {
            if staged
                .package
                .contains_part(DEFAULT_POWERPOINT_AUTHORS_PART)
            {
                return Err(Error::CollaborationPartCollision {
                    part_name: DEFAULT_POWERPOINT_AUTHORS_PART.to_owned(),
                });
            }
            staged.comment_authors_part = Some(DEFAULT_POWERPOINT_AUTHORS_PART.to_owned());
            let target =
                relative_part_target(&staged.presentation_part, DEFAULT_POWERPOINT_AUTHORS_PART);
            staged
                .package
                .get_or_create_part_rels(&staged.presentation_part)
                .add(rel_types::POWERPOINT_AUTHORS, &target);
        }
        staged.comment_authors.authors.push(author);
        staged.comment_authors_dirty = true;
        self.commit_candidate(staged)
    }

    /// Returns modern comments on one slide in producer order.
    pub fn comments(&self, slide_index: usize) -> Option<&[Comment]> {
        self.slides
            .get(slide_index)
            .and_then(|slide| slide.comments.as_ref())
            .map(|comments| comments.comments.comments.as_slice())
    }

    /// Adds one caller-identified modern comment atomically.
    pub fn add_comment(&mut self, slide_index: usize, comment: Comment) -> Result<()> {
        self.require_slide_index(slide_index)?;
        self.validate_comment_identity(&comment.id, &comment.author_id)?;
        let mut staged = self.clone();
        if staged.slides[slide_index].comments.is_none() {
            if staged
                .package
                .contains_part(DEFAULT_POWERPOINT_COMMENTS_PART)
                && !staged.slides.iter().any(|slide| {
                    slide.comments.as_ref().is_some_and(|comments| {
                        comments
                            .part_name
                            .eq_ignore_ascii_case(DEFAULT_POWERPOINT_COMMENTS_PART)
                    })
                })
            {
                return Err(Error::CollaborationPartCollision {
                    part_name: DEFAULT_POWERPOINT_COMMENTS_PART.to_owned(),
                });
            }
            let comment_part = MediaNamer::scan(
                "/ppt/comments",
                "comment",
                staged.package.parts.keys().map(String::as_str),
            )
            .next_part_name("xml");
            let slide_part = staged.slides[slide_index].part_name.clone();
            let relationship_id = staged.package.get_or_create_part_rels(&slide_part).add(
                rel_types::POWERPOINT_COMMENTS,
                &relative_part_target(&slide_part, &comment_part),
            );
            staged.slides[slide_index]
                .slide
                .ensure_modern_comment_relationship(&relationship_id)
                .map_err(|error| Error::MalformedPart {
                    part_name: slide_part,
                    message: error.to_string(),
                })?;
            staged.slides[slide_index].comments = Some(CommentRecord {
                part_name: comment_part,
                comments: CommentList::new(),
                dirty: true,
            });
        }
        let comments = staged.slides[slide_index]
            .comments
            .as_mut()
            .expect("comment record was materialized");
        comments.dirty = true;
        comments.comments.comments.push(comment);
        self.commit_candidate(staged)
    }

    /// Adds a modern comment anchored to a shape or its UTF-16 text range.
    pub fn add_comment_at_shape(
        &mut self,
        slide_index: usize,
        mut comment: Comment,
        shape_id: u32,
        text_range: Option<(usize, usize)>,
    ) -> Result<()> {
        self.require_slide_index(slide_index)?;
        let record = &self.slides[slide_index];
        if shape_id_count(
            &record.slide.common_slide_data.shape_tree.children,
            shape_id,
        ) != 1
        {
            return Err(invalid_shape_mutation(
                "comment anchor",
                format!("shape id {shape_id} is missing or ambiguous"),
            ));
        }
        let mut monikers = Vec::new();
        let target = find_comment_anchor_shape(
            &record.slide.common_slide_data.shape_tree.children,
            shape_id,
            &mut monikers,
        )
        .expect("unique shape id was checked");
        let text_context = if let Some((start, length)) = text_range {
            let ShapeTreeChild::Shape(shape) = target else {
                return Err(invalid_shape_mutation(
                    "comment anchor",
                    "text range requires an ordinary text shape",
                ));
            };
            let text = shape
                .text_body
                .as_ref()
                .map_or_else(String::new, office_art_comment_text);
            let text_length = text.encode_utf16().count();
            if start > text_length || length > text_length - start {
                return Err(invalid_shape_mutation(
                    "comment anchor",
                    "text range exceeds shape text",
                ));
            }
            if comment_text_context_count(
                &record.slide.common_slide_data.shape_tree.children,
                &text,
            ) != 1
            {
                return Err(invalid_shape_mutation(
                    "comment anchor",
                    "text context is ambiguous on the slide",
                ));
            }
            Some(text)
        } else {
            None
        };
        comment
            .set_drawing_anchor(record.id, &monikers, text_range, text_context.as_deref())
            .map_err(|error| invalid_shape_mutation("comment anchor", error.to_string()))?;
        self.add_comment(slide_index, comment)
    }

    /// Appends one caller-identified reply to an existing comment atomically.
    pub fn reply_to_comment(
        &mut self,
        slide_index: usize,
        comment_id: &str,
        reply: CommentReply,
    ) -> Result<()> {
        self.require_slide_index(slide_index)?;
        self.validate_comment_identity(&reply.id, &reply.author_id)?;
        let mut staged = self.clone();
        let comments = staged.slides[slide_index]
            .comments
            .as_mut()
            .ok_or_else(|| {
                invalid_presentation_mutation(
                    "reply to comment",
                    format!("unknown comment id {comment_id}"),
                )
            })?;
        let comment = comments
            .comments
            .comments
            .iter_mut()
            .find(|comment| comment.id == comment_id)
            .ok_or_else(|| {
                invalid_presentation_mutation(
                    "reply to comment",
                    format!("unknown comment id {comment_id}"),
                )
            })?;
        comment.add_reply(reply).map_err(|error| {
            invalid_presentation_mutation("reply to comment", error.to_string())
        })?;
        comments.dirty = true;
        self.commit_candidate(staged)
    }

    /// Moves one modern comment to a final zero-based position.
    pub fn move_comment(&mut self, slide_index: usize, from: usize, to: usize) -> Result<()> {
        self.require_slide_index(slide_index)?;
        let mut staged = self.clone();
        let comments = staged.slides[slide_index]
            .comments
            .as_mut()
            .ok_or_else(|| {
                invalid_presentation_mutation(
                    "move comment",
                    "slide has no modern comments".to_owned(),
                )
            })?;
        comments
            .comments
            .move_comment(from, to)
            .map_err(|error| invalid_presentation_mutation("move comment", error.to_string()))?;
        comments.dirty = true;
        self.commit_candidate(staged)
    }

    /// Moves one reply within an existing comment to a final zero-based position.
    pub fn move_reply(
        &mut self,
        slide_index: usize,
        comment_id: &str,
        from: usize,
        to: usize,
    ) -> Result<()> {
        self.require_slide_index(slide_index)?;
        let mut staged = self.clone();
        let comments = staged.slides[slide_index]
            .comments
            .as_mut()
            .ok_or_else(|| {
                invalid_presentation_mutation(
                    "move comment reply",
                    format!("unknown comment id {comment_id}"),
                )
            })?;
        let comment = comments
            .comments
            .comments
            .iter_mut()
            .find(|comment| comment.id == comment_id)
            .ok_or_else(|| {
                invalid_presentation_mutation(
                    "move comment reply",
                    format!("unknown comment id {comment_id}"),
                )
            })?;
        comment.move_reply(from, to).map_err(|error| {
            invalid_presentation_mutation("move comment reply", error.to_string())
        })?;
        comments.dirty = true;
        self.commit_candidate(staged)
    }

    /// Marks one modern comment thread resolved atomically.
    ///
    /// The status is stored on the thread's top-level comment, so a reply id
    /// is rejected as unknown.
    pub fn resolve_comment(&mut self, slide_index: usize, comment_id: &str) -> Result<()> {
        self.require_slide_index(slide_index)?;
        let mut staged = self.clone();
        let unknown = || {
            invalid_presentation_mutation(
                "resolve comment",
                format!("unknown comment id {comment_id}"),
            )
        };
        let comments = staged.slides[slide_index]
            .comments
            .as_mut()
            .ok_or_else(unknown)?;
        let comment = comments
            .comments
            .comments
            .iter_mut()
            .find(|comment| comment.id == comment_id)
            .ok_or_else(unknown)?;
        comment.status = Some("resolved".to_owned());
        comments.dirty = true;
        self.commit_candidate(staged)
    }

    /// Removes one modern comment with its replies, or one reply, atomically.
    ///
    /// The slide keeps its comment part when its last comment is removed.
    pub fn remove_comment(&mut self, slide_index: usize, comment_id: &str) -> Result<()> {
        self.require_slide_index(slide_index)?;
        let mut staged = self.clone();
        let unknown = || {
            invalid_presentation_mutation(
                "remove comment",
                format!("unknown comment id {comment_id}"),
            )
        };
        let comments = staged.slides[slide_index]
            .comments
            .as_mut()
            .ok_or_else(unknown)?;
        let list = &mut comments.comments.comments;
        if let Some(index) = list.iter().position(|comment| comment.id == comment_id) {
            list.remove(index);
        } else if !list
            .iter_mut()
            .any(|comment| comment.remove_reply(comment_id))
        {
            return Err(unknown());
        }
        comments.dirty = true;
        self.commit_candidate(staged)
    }

    /// Returns presentation sections in producer order.
    pub fn sections(&self) -> &[Section] {
        self.presentation.sections()
    }

    /// Replaces all presentation sections after validating stable slide ids.
    pub fn set_sections(&mut self, sections: Vec<Section>) -> Result<()> {
        let mut staged = self.clone();
        staged
            .presentation
            .set_sections(sections)
            .map_err(|error| invalid_presentation_mutation("set sections", error.to_string()))?;
        self.commit_candidate(staged)
    }

    /// Returns mutable notes-master header/footer settings when both exist.
    pub fn notes_header_footer_mut(&mut self) -> Option<&mut CT_HeaderFooter> {
        self.notes_master_dirty = self
            .notes_master
            .as_ref()
            .is_some_and(|master| master.header_footer.is_some());
        self.notes_master
            .as_mut()
            .and_then(|master| master.header_footer.as_mut())
    }

    /// Returns mutable handout-master header/footer settings when both exist.
    pub fn handout_header_footer_mut(&mut self) -> Option<&mut CT_HeaderFooter> {
        self.handout_master_dirty = self
            .handout_master
            .as_ref()
            .is_some_and(|master| master.header_footer.is_some());
        self.handout_master
            .as_mut()
            .and_then(|master| master.header_footer.as_mut())
    }

    fn require_slide_index(&self, index: usize) -> Result<()> {
        if index < self.slides.len() {
            Ok(())
        } else {
            Err(Error::UnknownSlideIndex {
                index,
                slide_count: self.slides.len(),
            })
        }
    }

    fn validate_comment_identity(&self, id: &str, author_id: &str) -> Result<()> {
        if !self
            .comment_authors
            .authors
            .iter()
            .any(|author| author.id == author_id)
        {
            return Err(invalid_presentation_mutation(
                "mutate comments",
                format!("unknown comment author id {author_id}"),
            ));
        }
        if self.slides.iter().any(|slide| {
            slide.comments.as_ref().is_some_and(|comments| {
                comments.comments.comments.iter().any(|comment| {
                    comment.id == id || comment.replies().iter().any(|reply| reply.id == id)
                })
            })
        }) {
            return Err(invalid_presentation_mutation(
                "mutate comments",
                format!("duplicate modern comment id {id}"),
            ));
        }
        Ok(())
    }

    fn commit_candidate(&mut self, staged: Self) -> Result<()> {
        let embedded_invalidated_signatures = staged.embedded_invalidated_signatures.clone();
        let package_signatures_invalidated = staged.package_signatures_invalidated
            || staged.retained_package_signature_would_be_invalidated()?;
        let mut package = staged.staged_package(false)?;
        embedded::persist_invalidated_package_signature(
            &mut package,
            package_signatures_invalidated,
        )?;
        let mut output = Cursor::new(Vec::new());
        package.write_to(&mut output)?;
        let mut reopened = Self::from_bytes(output.get_ref())?;
        reopened.embedded_invalidated_signatures = embedded_invalidated_signatures;
        reopened.package_signatures_invalidated = package_signatures_invalidated;
        *self = reopened;
        Ok(())
    }

    /// Removes one slide and its owned package graph by zero-based index.
    pub fn remove_slide(&mut self, index: usize) -> Result<()> {
        if index >= self.slides.len() {
            return Err(Error::UnknownSlideIndex {
                index,
                slide_count: self.slides.len(),
            });
        }
        let mut staged = self.clone();
        staged.remove_slide_in_place(index)?;
        self.commit_candidate(staged)
    }

    /// Moves one slide to an existing final zero-based index.
    pub fn move_slide(&mut self, from_index: usize, to_index: usize) -> Result<()> {
        let slide_count = self.slides.len();
        if from_index >= slide_count {
            return Err(Error::UnknownSlideIndex {
                index: from_index,
                slide_count,
            });
        }
        if to_index >= slide_count {
            return Err(Error::UnknownSlideIndex {
                index: to_index,
                slide_count,
            });
        }
        if from_index == to_index {
            return Ok(());
        }
        let slide = self.slides.remove(from_index);
        self.slides.insert(to_index, slide);
        let slide_id = self.presentation.slide_ids.remove(from_index);
        self.presentation.slide_ids.insert(to_index, slide_id);
        Ok(())
    }

    /// Replaces one slide's speaker-note text, creating its notes slide when absent.
    ///
    /// A new notes slide follows python-pptx: it relates to the notes master
    /// and the slide, and clones the master's slide image, body and slide
    /// number placeholders. A presentation without a notes master first
    /// receives a copy of the bundled template's notes master and its theme,
    /// which needs the `default-template` or `render` feature.
    pub fn set_notes_text(&mut self, slide_index: usize, text: &str) -> Result<()> {
        self.require_slide_index(slide_index)?;
        if self.slides[slide_index].notes.is_some() {
            return slide_mut(&mut self.slides[slide_index]).set_notes_text(text);
        }
        let mut staged = self.clone();
        staged.add_notes_slide_in_place(slide_index)?;
        slide_mut(&mut staged.slides[slide_index]).set_notes_text(text)?;
        self.commit_candidate(staged)
    }

    fn add_notes_slide_in_place(&mut self, slide_index: usize) -> Result<()> {
        const OPERATION: &str = "add notes slide";
        if self.notes_master.is_none() {
            self.add_default_notes_master_in_place()?;
        }
        let (Some(master_part), Some(master)) = (&self.notes_master_part, &self.notes_master)
        else {
            return Err(Error::InvalidSlideMutation {
                operation: OPERATION,
                message: "presentation has no notes master".to_owned(),
            });
        };
        let mut notes = CT_NotesSlide::from_xml(
            format!(
                r#"<p:notes xmlns:a="{A_NS}" xmlns:p="{P_NS}" xmlns:r="{R_NS}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>"#
            )
            .as_bytes(),
        )
        .expect("canonical notes slide shell");
        let tree = &mut notes.common_slide_data.shape_tree;
        let mut shape_ids = ShapeIdAllocator::scan(tree);
        for child in &master.common_slide_data.shape_tree.children {
            let ShapeTreeChild::Shape(shape) = child else {
                continue;
            };
            let Some(placeholder) = &shape.placeholder else {
                continue;
            };
            let label = match placeholder.ph_type {
                Some(PhType::SlideImage) => "Slide Image Placeholder",
                Some(PhType::Body) => "Notes Placeholder",
                Some(PhType::SlideNumber) => "Slide Number Placeholder",
                _ => continue,
            };
            let id = shape_ids.allocate();
            let mut clone = CT_Shape::new_placeholder(id, placeholder.clone())
                .map_err(|error| invalid_slide_mutation(OPERATION, error.to_string()))?;
            clone
                .set_name(&format!("{label} {}", id - 1))
                .map_err(|error| invalid_slide_mutation(OPERATION, error.to_string()))?;
            if placeholder.ph_type != Some(PhType::Body) {
                clone.text_body = None;
            }
            tree.append_child(ShapeTreeChild::Shape(clone));
        }

        let slide_part = self.slides[slide_index].part_name.clone();
        let notes_part = MediaNamer::scan(
            "/ppt/notesSlides",
            "notesSlide",
            self.package.parts.keys().map(String::as_str),
        )
        .next_part_name("xml");
        let mut notes_relationships = Relationships::new();
        notes_relationships.add(
            rel_types::NOTES_MASTER,
            &relative_part_target(&notes_part, master_part),
        );
        notes_relationships.add(
            rel_types::SLIDE,
            &relative_part_target(&notes_part, &slide_part),
        );
        let mut slide_relationships = self
            .package
            .get_part_rels(&slide_part)
            .cloned()
            .unwrap_or_default();
        slide_relationships.add(
            rel_types::NOTES_SLIDE,
            &relative_part_target(&slide_part, &notes_part),
        );
        let notes_xml = notes
            .to_xml()
            .map_err(|error| invalid_slide_mutation(OPERATION, error.to_string()))?;
        self.package.set_part(&notes_part, notes_xml);
        self.package.set_part_rels(&notes_part, notes_relationships);
        self.package.set_part_rels(&slide_part, slide_relationships);
        self.package
            .content_types
            .add_override(&notes_part, content_types::NOTES_SLIDE);
        self.slides[slide_index].notes = Some(NotesRecord {
            part_name: notes_part,
            notes,
        });
        Ok(())
    }

    #[cfg(any(feature = "default-template", feature = "render"))]
    fn add_default_notes_master_in_place(&mut self) -> Result<()> {
        let template = Self::from_bytes(DEFAULT_TEMPLATE)?;
        let (Some(template_master_part), Some(master)) =
            (template.notes_master_part, template.notes_master)
        else {
            return Err(invalid_slide_mutation(
                "add notes master",
                "the bundled template has no notes master",
            ));
        };
        let template_theme_part =
            related_internal_part(&template.package, &template_master_part, rel_types::THEME)?
                .ok_or_else(|| {
                    invalid_slide_mutation(
                        "add notes master",
                        "the bundled notes master has no theme",
                    )
                })?;
        let master_part = MediaNamer::scan(
            "/ppt/notesMasters",
            "notesMaster",
            self.package.parts.keys().map(String::as_str),
        )
        .next_part_name("xml");
        let theme_part = MediaNamer::scan(
            "/ppt/theme",
            "theme",
            self.package.parts.keys().map(String::as_str),
        )
        .next_part_name("xml");
        self.package.set_part(
            &master_part,
            required_part(&template.package, &template_master_part)?.to_vec(),
        );
        self.package.set_part(
            &theme_part,
            required_part(&template.package, &template_theme_part)?.to_vec(),
        );
        let mut master_relationships = Relationships::new();
        master_relationships.add(
            rel_types::THEME,
            &relative_part_target(&master_part, &theme_part),
        );
        self.package
            .set_part_rels(&master_part, master_relationships);
        self.package
            .content_types
            .add_override(&master_part, content_types::NOTES_MASTER);
        self.package
            .content_types
            .add_override(&theme_part, content_types::THEME);
        let mut presentation_relationships = self
            .package
            .get_part_rels(&self.presentation_part)
            .cloned()
            .unwrap_or_default();
        presentation_relationships.add(
            rel_types::NOTES_MASTER,
            &relative_part_target(&self.presentation_part, &master_part),
        );
        self.package
            .set_part_rels(&self.presentation_part, presentation_relationships);
        self.notes_master_part = Some(master_part);
        self.notes_master = Some(master);
        Ok(())
    }

    #[cfg(not(any(feature = "default-template", feature = "render")))]
    fn add_default_notes_master_in_place(&mut self) -> Result<()> {
        Err(invalid_slide_mutation(
            "add notes master",
            "presentation has no notes master and the bundled template is not compiled in",
        ))
    }

    /// Duplicates one slide immediately after its source.
    pub fn duplicate_slide(&mut self, index: usize) -> Result<SlideRef<'_>> {
        if index >= self.slides.len() {
            return Err(Error::UnknownSlideIndex {
                index,
                slide_count: self.slides.len(),
            });
        }
        let mut staged = self.clone();
        let inserted = staged.duplicate_slide_in_place(index)?;
        *self = staged;
        self.slides
            .get(inserted)
            .map(slide_ref)
            .ok_or(Error::UnknownSlideIndex {
                index: inserted,
                slide_count: self.slides.len(),
            })
    }

    /// Copies one placeholder-free SmartArt slide from another presentation.
    ///
    /// `destination_layout_index` explicitly selects the destination-owned layout.
    /// Internal relationships are limited to that layout, images, and the five
    /// SmartArt diagram types. Internal images must not own relationships, and the
    /// copied diagram graph is limited to 128 parts. Slides with placeholders,
    /// external layouts, notes, comments, charts, embedded media, OLE objects, or
    /// other internal dependencies anywhere in the copied diagram graph are
    /// rejected without mutation.
    pub fn transfer_smartart_slide_from(
        &mut self,
        source: &Presentation,
        index: usize,
        destination_layout_index: usize,
    ) -> Result<SlideRef<'_>> {
        if index >= source.slides.len() {
            return Err(Error::UnknownSlideIndex {
                index,
                slide_count: source.slides.len(),
            });
        }
        let destination_layout = self
            .layouts
            .get(destination_layout_index)
            .ok_or(Error::UnknownLayoutIndex {
                index: destination_layout_index,
                layout_count: self.layouts.len(),
            })?
            .part_name
            .clone();
        preflight_cross_presentation_transfer(&source.package, &source.slides[index])?;
        let mut staged = self.clone();
        let inserted = staged.slides.len();
        staged.copy_slide_from_in_place(
            &source.package,
            source.slides[index].clone(),
            inserted,
            "transfer SmartArt slide",
            Some(&destination_layout),
        )?;
        self.commit_candidate(staged)?;
        self.slides
            .get(inserted)
            .map(slide_ref)
            .ok_or(Error::UnknownSlideIndex {
                index: inserted,
                slide_count: self.slides.len(),
            })
    }

    /// Imports one slide of another presentation, using this presentation's theme.
    ///
    /// The copy takes the layout at `layout_index` when one is given, and
    /// otherwise the first layout of this presentation named like the source
    /// slide's layout. It is inserted at `insert_index`, or appended when that
    /// is `None`. Shapes, placeholders, the slide background, speaker notes and
    /// SmartArt graphs are copied, and every relationship attribute is rebound
    /// to the new relationship scope. Pictures, audio and video reuse equal
    /// media parts of this presentation. External relationships such as
    /// hyperlinks keep their targets. A slide with comments, a chart, an
    /// embedded object, a jump to another slide, or any other internal
    /// relationship type is rejected without mutation. A notes jump to another
    /// slide is rejected too. Jumps that name no relationship, such as a
    /// `ppaction://customshow` action, are carried unchanged and may name a
    /// custom show this presentation lacks. The name match takes the first
    /// layout of that name, so a copy inside one multi-master deck should pass
    /// the source slide's own layout index.
    pub fn import_slide(
        &mut self,
        source: &Presentation,
        index: usize,
        layout_index: Option<usize>,
        insert_index: Option<usize>,
    ) -> Result<SlideRef<'_>> {
        const OPERATION: &str = "import slide";
        let record = source.slides.get(index).ok_or(Error::UnknownSlideIndex {
            index,
            slide_count: source.slides.len(),
        })?;
        let inserted = insert_index.unwrap_or(self.slides.len());
        if inserted > self.slides.len() {
            return Err(Error::UnknownSlideIndex {
                index: inserted,
                slide_count: self.slides.len(),
            });
        }
        let layout_index = match layout_index {
            Some(layout_index) if layout_index >= self.layouts.len() => {
                return Err(Error::UnknownLayoutIndex {
                    index: layout_index,
                    layout_count: self.layouts.len(),
                });
            }
            Some(layout_index) => layout_index,
            None => {
                let name = source
                    .slide_layout_index(index)
                    .and_then(|layout| source.layout_name(layout))
                    .ok_or_else(|| {
                        invalid_presentation_mutation(
                            OPERATION,
                            "the source slide's layout has no name, pass the destination layout"
                                .to_owned(),
                        )
                    })?;
                (0..self.layouts.len())
                    .find(|&layout| self.layout_name(layout) == Some(name))
                    .ok_or_else(|| {
                        invalid_presentation_mutation(
                            OPERATION,
                            format!(
                                "this presentation has no layout named {name:?}, pass the destination layout"
                            ),
                        )
                    })?
            }
        };
        let layout_part = self.layouts[layout_index].part_name.clone();
        preflight_slide_import(&source.package, record)?;
        let mut staged = self.clone();
        if record.notes.is_some() && staged.notes_master.is_none() {
            staged.add_default_notes_master_in_place()?;
        }
        staged.copy_slide_from_in_place(
            &source.package,
            record.clone(),
            inserted,
            OPERATION,
            Some(&layout_part),
        )?;
        self.commit_candidate(staged)?;
        self.slides
            .get(inserted)
            .map(slide_ref)
            .ok_or(Error::UnknownSlideIndex {
                index: inserted,
                slide_count: self.slides.len(),
            })
    }

    fn remove_slide_in_place(&mut self, index: usize) -> Result<()> {
        let record = self.slides[index].clone();
        self.presentation.remove_slide_from_sections(record.id);
        let presentation_relationship_id =
            self.presentation.slide_ids[index].relationship_id.clone();
        self.presentation
            .remove_custom_show_slide(&presentation_relationship_id)
            .map_err(|error| Error::MalformedPart {
                part_name: self.presentation_part.clone(),
                message: error.to_string(),
            })?;

        let mut media_candidates = HashSet::new();
        self.release_slide_jumps_to(&record.part_name, &mut media_candidates)?;
        collect_media_targets(&self.package, &record.part_name, &mut media_candidates);
        if let Some(notes) = &record.notes {
            collect_media_targets(&self.package, &notes.part_name, &mut media_candidates);
            self.package.remove_part(&notes.part_name);
            self.package.remove_part_rels(&notes.part_name);
            self.package.content_types.remove_override(&notes.part_name);
        }
        if let Some(comments) = &record.comments {
            self.package.remove_part(&comments.part_name);
            self.package.remove_part_rels(&comments.part_name);
            self.package
                .content_types
                .remove_override(&comments.part_name);
        }
        self.package.remove_part(&record.part_name);
        self.package.remove_part_rels(&record.part_name);
        self.package
            .content_types
            .remove_override(&record.part_name);
        if let Some(relationships) = self.package.get_part_rels_mut(&self.presentation_part) {
            relationships
                .items
                .retain(|relationship| relationship.id != presentation_relationship_id);
        }
        self.presentation.slide_ids.remove(index);
        self.slides.remove(index);
        prune_unreachable_parts(&mut self.package, &media_candidates);
        self.media_store = MediaStore::scan(&self.package);
        Ok(())
    }

    /// Makes every hyperlink of another slide that jumps to `slide_part` do
    /// nothing and removes its relationship, as PowerPoint does when the
    /// target slide is deleted.
    ///
    /// Every `a:hlinkClick`, `a:hlinkHover`, or `a:hlinkMouseOver` that
    /// named it, on a shape or a text run, keeps an empty `r:id` and the
    /// action `ppaction://noaction`. Every relationship that the slide named
    /// only inside a released element goes, a click sound included, and the
    /// internal targets of those relationships join `candidates` for pruning.
    fn release_slide_jumps_to(
        &mut self,
        slide_part: &str,
        candidates: &mut HashSet<String>,
    ) -> Result<()> {
        for record in &mut self.slides {
            let Some(relationships) = self.package.get_part_rels_mut(&record.part_name) else {
                continue;
            };
            let released: HashSet<String> = relationships
                .items
                .iter()
                .filter(|relationship| {
                    !relationship_is_external(relationship)
                        && OpcPackage::resolve_rel_target(&record.part_name, &relationship.target)
                            == slide_part
                })
                .map(|relationship| relationship.id.clone())
                .collect();
            if released.is_empty() || record.part_name == slide_part {
                continue;
            }
            let malformed = |error: OxmlError| Error::MalformedPart {
                part_name: record.part_name.clone(),
                message: error.to_string(),
            };
            let original = record.slide.to_xml().map_err(malformed)?;
            let xml = release_hyperlinks(&original, &released).map_err(malformed)?;
            let before: HashSet<String> = relationship_ids(&original)
                .map_err(malformed)?
                .into_iter()
                .collect();
            let after: HashSet<String> = relationship_ids(&xml)
                .map_err(malformed)?
                .into_iter()
                .collect();
            relationships.items.retain(|relationship| {
                let dropped = (released.contains(&relationship.id)
                    || before.contains(&relationship.id))
                    && !after.contains(&relationship.id);
                if dropped && !relationship_is_external(relationship) {
                    let target =
                        OpcPackage::resolve_rel_target(&record.part_name, &relationship.target);
                    if target != slide_part {
                        candidates.insert(target);
                    }
                }
                !dropped
            });
            record.slide = CT_Slide::from_xml(&xml).map_err(malformed)?;
        }
        Ok(())
    }

    fn duplicate_slide_in_place(&mut self, index: usize) -> Result<usize> {
        let source_package = self.package.clone();
        let source = self.slides[index].clone();
        self.copy_slide_from_in_place(&source_package, source, index + 1, "duplicate slide", None)
    }

    fn copy_slide_from_in_place(
        &mut self,
        source_package: &OpcPackage,
        source: SlideRecord,
        inserted: usize,
        operation: &'static str,
        layout_target: Option<&str>,
    ) -> Result<usize> {
        if source.comments.is_some() {
            return Err(Error::InvalidPresentationMutation {
                operation,
                message: "slides with modern comments require caller-supplied fresh comment ids"
                    .to_owned(),
            });
        }
        let slide_part = MediaNamer::scan(
            "/ppt/slides",
            "slide",
            self.package.parts.keys().map(String::as_str),
        )
        .next_part_name("xml");
        let id = self
            .presentation
            .slide_ids
            .iter()
            .map(|slide| slide.id)
            .max()
            .unwrap_or(255)
            .max(255)
            .checked_add(1)
            .ok_or(Error::SlideIdExhausted)?;

        let notes_part = source.notes.as_ref().map(|_| {
            MediaNamer::scan(
                "/ppt/notesSlides",
                "notesSlide",
                self.package.parts.keys().map(String::as_str),
            )
            .next_part_name("xml")
        });
        let notes = if let (Some(source_notes), Some(notes_part)) = (&source.notes, &notes_part) {
            let source_relationships = source_package
                .get_part_rels(&source_notes.part_name)
                .cloned()
                .unwrap_or_default();
            let (mut relationships, mut relationship_map) = duplicate_relationship_scope(
                source_package,
                &mut self.package,
                &mut self.media_store,
                &source_notes.part_name,
                notes_part,
                &source_relationships,
                RelationshipScopeTargets {
                    slide: Some(&slide_part),
                    notes: None,
                    layout: None,
                    notes_master: self.notes_master_part.as_deref(),
                },
            )?;
            relationships
                .items
                .retain(|relationship| relationship.rel_type != rel_types::SLIDE);
            let back_relationship_id = relationships.add(
                rel_types::SLIDE,
                &relative_part_target(notes_part, &slide_part),
            );
            for relationship in source_relationships
                .items
                .iter()
                .filter(|relationship| relationship.rel_type == rel_types::SLIDE)
            {
                relationship_map.insert(relationship.id.clone(), back_relationship_id.clone());
            }
            let xml = source_notes
                .notes
                .to_xml()
                .map_err(|error| Error::MalformedPart {
                    part_name: source_notes.part_name.clone(),
                    message: error.to_string(),
                })?;
            // One exact pass over the complete map: a second pass would
            // re-map ids the first one had just produced.
            let xml = rewrite_exact_rel_ids(&xml, &relationship_map).map_err(|error| {
                Error::MalformedPart {
                    part_name: source_notes.part_name.clone(),
                    message: error.to_string(),
                }
            })?;
            let xml = rewrite_shape_ids(&xml).map_err(|error| Error::MalformedPart {
                part_name: source_notes.part_name.clone(),
                message: error.to_string(),
            })?;
            let notes = CT_NotesSlide::from_xml(&xml).map_err(|error| Error::MalformedPart {
                part_name: notes_part.clone(),
                message: error.to_string(),
            })?;
            self.package.set_part(notes_part, xml);
            self.package
                .part_rels
                .insert(notes_part.clone(), relationships);
            self.package
                .content_types
                .add_override(notes_part, content_types::NOTES_SLIDE);
            Some(NotesRecord {
                part_name: notes_part.clone(),
                notes,
            })
        } else {
            None
        };

        let source_relationships = source_package
            .get_part_rels(&source.part_name)
            .cloned()
            .unwrap_or_default();
        let (slide_relationships, relationship_map) = duplicate_relationship_scope(
            source_package,
            &mut self.package,
            &mut self.media_store,
            &source.part_name,
            &slide_part,
            &source_relationships,
            RelationshipScopeTargets {
                slide: None,
                notes: notes_part.as_deref(),
                layout: layout_target,
                notes_master: None,
            },
        )?;
        let slide_xml = source
            .slide
            .to_xml()
            .map_err(|error| Error::MalformedPart {
                part_name: source.part_name.clone(),
                message: error.to_string(),
            })?;
        let slide_xml = rewrite_rel_ids(&slide_xml, &relationship_map).map_err(|error| {
            Error::MalformedPart {
                part_name: source.part_name.clone(),
                message: error.to_string(),
            }
        })?;
        let slide_xml = rewrite_shape_ids(&slide_xml).map_err(|error| Error::MalformedPart {
            part_name: source.part_name.clone(),
            message: error.to_string(),
        })?;
        let slide = CT_Slide::from_xml(&slide_xml).map_err(|error| Error::MalformedPart {
            part_name: slide_part.clone(),
            message: error.to_string(),
        })?;

        let mut presentation_relationships = self
            .package
            .get_part_rels(&self.presentation_part)
            .cloned()
            .unwrap_or_default();
        let presentation_relationship_id = presentation_relationships.add(
            rel_types::SLIDE,
            &relative_part_target(&self.presentation_part, &slide_part),
        );
        let slide_id = CT_SlideId::new(id, presentation_relationship_id).map_err(|error| {
            Error::MalformedPart {
                part_name: self.presentation_part.clone(),
                message: error.to_string(),
            }
        })?;
        self.package.set_part(&slide_part, slide_xml);
        self.package.set_part_rels(&slide_part, slide_relationships);
        self.package
            .set_part_rels(&self.presentation_part, presentation_relationships);
        self.package
            .content_types
            .add_override(&slide_part, content_types::SLIDE);

        self.presentation.slide_ids.insert(inserted, slide_id);
        self.slides.insert(
            inserted,
            SlideRecord {
                id,
                part_name: slide_part,
                slide,
                notes,
                comments: source.comments.clone(),
            },
        );
        Ok(inserted)
    }

    /// Synthesizes a new slide from one layout selected by zero-based index.
    pub fn add_slide(&mut self, layout_index: usize) -> Result<SlideRef<'_>> {
        let layout = self
            .layouts
            .get(layout_index)
            .ok_or(Error::UnknownLayoutIndex {
                index: layout_index,
                layout_count: self.layouts.len(),
            })?;
        let layout_part = layout.part_name.clone();
        let placeholders = layout
            .layout
            .common_slide_data
            .shape_tree
            .children
            .iter()
            .filter_map(|child| match child {
                ShapeTreeChild::Shape(shape) => shape.placeholder.as_ref(),
                _ => None,
            })
            .filter(|placeholder| !is_latent_placeholder(placeholder))
            .map(|placeholder| {
                (
                    placeholder.ph_type.clone(),
                    placeholder.idx,
                    placeholder.orientation() == Some("vert"),
                )
            })
            .collect::<Vec<_>>();

        let mut shape_tree = CT_ShapeTree::new();
        let mut shape_ids = ShapeIdAllocator::scan(&shape_tree);
        for (ph_type, idx, vertical) in placeholders {
            let id = shape_ids.allocate();
            let name = placeholder_name(ph_type.as_ref(), vertical, id);
            let shape = CT_Shape::new_placeholder(id, CT_Placeholder::new(ph_type, idx))
                .and_then(|mut shape| shape.set_name(&name).map(|()| shape))
                .map_err(|error| Error::MalformedPart {
                    part_name: "new slide".to_owned(),
                    message: error.to_string(),
                })?;
            shape_tree.children.push(ShapeTreeChild::Shape(shape));
        }

        let slide_part = MediaNamer::scan(
            "/ppt/slides",
            "slide",
            self.package.parts.keys().map(String::as_str),
        )
        .next_part_name("xml");
        let slide = CT_Slide::new(shape_tree);
        let slide_xml = slide.to_xml().map_err(|error| Error::MalformedPart {
            part_name: slide_part.clone(),
            message: error.to_string(),
        })?;

        let id = self
            .presentation
            .slide_ids
            .iter()
            .map(|slide| slide.id)
            .max()
            .unwrap_or(255)
            .max(255)
            .checked_add(1)
            .ok_or(Error::SlideIdExhausted)?;

        let mut slide_relationships = Relationships::new();
        slide_relationships.add(
            rel_types::SLIDE_LAYOUT,
            &relative_part_target(&slide_part, &layout_part),
        );
        let mut presentation_relationships = self
            .package
            .get_part_rels(&self.presentation_part)
            .cloned()
            .unwrap_or_default();
        let presentation_relationship_id = presentation_relationships.add(
            rel_types::SLIDE,
            &relative_part_target(&self.presentation_part, &slide_part),
        );
        let slide_id = CT_SlideId::new(id, presentation_relationship_id).map_err(|error| {
            Error::MalformedPart {
                part_name: self.presentation_part.clone(),
                message: error.to_string(),
            }
        })?;

        self.package.set_part(&slide_part, slide_xml);
        self.package.set_part_rels(&slide_part, slide_relationships);
        self.package
            .set_part_rels(&self.presentation_part, presentation_relationships);
        self.package
            .content_types
            .add_override(&slide_part, content_types::SLIDE);
        self.presentation.slide_ids.push(slide_id);
        self.slides.push(SlideRecord {
            id,
            part_name: slide_part,
            slide,
            notes: None,
            comments: None,
        });
        self.slides
            .last()
            .map(slide_ref)
            .ok_or(Error::MalformedPart {
                part_name: self.presentation_part.clone(),
                message: "new slide record was not retained".to_owned(),
            })
    }

    /// Appends a relationship-backed picture at the top of one slide's z-order.
    #[allow(clippy::too_many_arguments)]
    pub fn add_picture(
        &mut self,
        slide_index: usize,
        image_data: &[u8],
        image_filename: &str,
        left: Emu,
        top: Emu,
        width: Option<Emu>,
        height: Option<Emu>,
    ) -> Result<ShapeRef<'_>> {
        self.append_picture(
            slide_index,
            &[],
            image_data,
            image_filename,
            left,
            top,
            width,
            height,
        )
        .map(|picture| shape_ref(picture))
    }

    /// Stages the image part and slide relationship of a new picture, then
    /// appends it to the slide's own shapes or to the group that `group`
    /// indexes, which callers have checked.
    #[allow(clippy::too_many_arguments)]
    fn append_picture(
        &mut self,
        slide_index: usize,
        group: &[usize],
        image_data: &[u8],
        image_filename: &str,
        left: Emu,
        top: Emu,
        width: Option<Emu>,
        height: Option<Emu>,
    ) -> Result<&mut ShapeTreeChild> {
        let slide = self
            .slides
            .get(slide_index)
            .ok_or(Error::UnknownSlideIndex {
                index: slide_index,
                slide_count: self.slides.len(),
            })?;
        let slide_part = slide.part_name.clone();
        let (width, height) = picture_dimensions(image_data, image_filename, width, height)?;
        let id = ShapeIdAllocator::scan(&slide.slide.common_slide_data.shape_tree).allocate();

        let (package, media_store, relationship_id) =
            self.stage_slide_image(&slide_part, image_data, image_filename);
        let picture = CT_Picture::new(
            id,
            &format!("Picture {id}"),
            &relationship_id,
            positioned_transform(left, top, width, height),
        )
        .map_err(|error| invalid_shape_construction("add picture", error))?;

        self.package = package;
        self.media_store = media_store;
        let tree = &mut self.slides[slide_index].slide.common_slide_data.shape_tree;
        Ok(append_member(tree, group, ShapeTreeChild::Picture(picture)))
    }

    /// Stages one image part and a slide's relationship to it on copies of
    /// the package and media store, and returns them with the relationship
    /// id. A relationship the slide already has to the same part is reused.
    fn stage_slide_image(
        &self,
        slide_part: &str,
        image_data: &[u8],
        image_filename: &str,
    ) -> (OpcPackage, MediaStore, String) {
        let mut package = self.package.clone();
        let mut media_store = self.media_store.clone();
        let media_part = media_store.insert(&mut package, image_data, image_filename);
        let mut relationships = package
            .get_part_rels(slide_part)
            .cloned()
            .unwrap_or_default();
        let relationship_id = relationships
            .items
            .iter()
            .find(|relationship| {
                relationship.rel_type == rel_types::IMAGE
                    && !relationship_is_external(relationship)
                    && OpcPackage::resolve_rel_target(slide_part, &relationship.target)
                        == media_part
            })
            .map(|relationship| relationship.id.clone())
            .unwrap_or_else(|| {
                relationships.add(
                    rel_types::IMAGE,
                    &relative_part_target(slide_part, &media_part),
                )
            });
        package.set_part_rels(slide_part, relationships);
        (package, media_store, relationship_id)
    }

    /// Fills the placeholder at `shape_path` with a picture, as python-pptx
    /// `insert_picture` does, and returns the picture.
    ///
    /// `shape_path` is read as in [`Self::effective_geometry`]. The
    /// placeholder's `p:sp` becomes a `p:pic` with the same id, name, and
    /// `p:ph`, so it keeps the position and size it inherits from its layout,
    /// or its own when it has one. The image is cropped evenly on its two
    /// longer sides to the placeholder's aspect ratio, so it fills the frame
    /// without distortion. A picture placeholder and a content placeholder,
    /// one of type `obj` or without a type, take a picture. Any other shape,
    /// a placeholder without a size, and an image without pixel dimensions
    /// are rejected without change.
    #[cfg(feature = "render")]
    pub fn insert_picture(
        &mut self,
        slide_index: usize,
        shape_path: &[usize],
        image_data: &[u8],
        image_filename: &str,
    ) -> Result<ShapeRef<'_>> {
        const OPERATION: &str = "insert picture";
        let child = self.shape_at(slide_index, shape_path, OPERATION)?.child;
        let ShapeTreeChild::Shape(shape) = child else {
            return Err(invalid_shape_mutation(
                OPERATION,
                "only a picture or content placeholder takes a picture; use add_picture for \
                 a free-standing one",
            ));
        };
        let placeholder = match &shape.placeholder {
            Some(placeholder)
                if matches!(
                    placeholder.ph_type,
                    None | Some(PhType::Object | PhType::Picture)
                ) =>
            {
                placeholder.clone()
            }
            Some(placeholder) => {
                return Err(invalid_shape_mutation(
                    OPERATION,
                    format!(
                        "a {} placeholder does not take a picture; use a picture or content \
                         placeholder, or add_picture",
                        placeholder.effective_type().as_str()
                    ),
                ));
            }
            None => {
                return Err(invalid_shape_mutation(
                    OPERATION,
                    "the shape is not a placeholder; use add_picture",
                ));
            }
        };
        let own_transform = shape.shape_properties.transform.clone();
        let id = child.non_visual_id().ok_or_else(|| {
            invalid_shape_mutation(OPERATION, "the placeholder has no p:cNvPr id")
        })?;
        let name = child
            .non_visual_name()
            .unwrap_or_else(|| format!("Picture Placeholder {id}"));
        let (_, _, width, height) = self
            .effective_geometry(slide_index, shape_path)?
            .filter(|(_, _, width, height)| width.0 > 0 && height.0 > 0)
            .ok_or_else(|| {
                invalid_shape_mutation(OPERATION, "the placeholder has no size to fill")
            })?;
        if ImageFormat::sniff(image_data).is_none() {
            return Err(Error::UnsupportedPicture {
                filename: image_filename.to_owned(),
            });
        }
        let info = probe(image_data)
            .filter(|info| info.width_px > 0 && info.height_px > 0)
            .ok_or_else(|| Error::UnavailablePictureDimensions {
                filename: image_filename.to_owned(),
            })?;
        let crop = fill_crop(
            (f64::from(info.width_px), f64::from(info.height_px)),
            (width.0 as f64, height.0 as f64),
        );
        let slide_part = self.slides[slide_index].part_name.clone();
        let (package, media_store, relationship_id) =
            self.stage_slide_image(&slide_part, image_data, image_filename);
        let mut picture =
            CT_Picture::new_placeholder(id, &name, &relationship_id, &placeholder, crop)
                .map_err(|error| invalid_shape_construction(OPERATION, error))?;
        picture.shape_properties.transform = own_transform;
        self.package = package;
        self.media_store = media_store;
        let mut indices = shape_path.iter();
        let mut target = indices
            .next()
            .and_then(|index| slide_mut(&mut self.slides[slide_index]).into_shape_mut(*index));
        for index in indices {
            target = target.and_then(|shape| shape.into_child_mut(*index));
        }
        let target = target.expect("the placeholder was found above");
        *target.child = ShapeTreeChild::Picture(picture);
        Ok(shape_ref(target.child))
    }

    /// Returns the embedded image one picture shows, found by shape id.
    ///
    /// Pictures inside groups and alternate-content fallbacks are found too.
    /// The content type comes from the package, falling back to the sniffed
    /// image format.
    pub fn picture_image(&self, slide_index: usize, shape_id: u32) -> Result<PictureImage<'_>> {
        self.require_slide_index(slide_index)?;
        let record = &self.slides[slide_index];
        let picture = find_picture(
            &record.slide.common_slide_data.shape_tree.children,
            shape_id,
        )
        .ok_or_else(|| not_a_picture("read picture image", shape_id))?;
        let relationship_id = picture_image_relationship(picture, "read picture image", shape_id)?;
        let part_name = require_internal_related_part(
            &self.package,
            &record.part_name,
            relationship_id,
            rel_types::IMAGE,
        )?;
        let bytes = required_part(&self.package, &part_name)?;
        let content_type = self
            .package
            .content_types
            .content_type_for(&part_name)
            .map_or_else(
                || resolve(bytes, &part_name).content_type().to_owned(),
                str::to_owned,
            );
        Ok(PictureImage {
            part_name,
            content_type,
            bytes,
        })
    }

    /// Replaces the image one picture shows while keeping its geometry.
    ///
    /// Only this picture changes. A relationship other pictures share is
    /// split first, a part no other relationship targets is rewritten in
    /// place when its extension fits the new format, and any other part is
    /// left to its remaining users while the picture moves to a new or equal
    /// media part. A picture that carries a second image, such as an SVG
    /// alternate, is rejected because PowerPoint would keep showing it.
    pub fn replace_picture_image(
        &mut self,
        slide_index: usize,
        shape_id: u32,
        image_data: &[u8],
    ) -> Result<()> {
        const OPERATION: &str = "replace picture image";
        self.require_slide_index(slide_index)?;
        let format = ImageFormat::sniff(image_data).ok_or_else(|| Error::InvalidShapeMutation {
            operation: OPERATION,
            message: "replacement bytes are not a supported image".to_owned(),
        })?;
        let mut staged = self.clone();
        let slide_part = staged.slides[slide_index].part_name.clone();
        let slide = &mut staged.slides[slide_index].slide;
        let picture = find_picture_mut(&mut slide.common_slide_data.shape_tree.children, shape_id)
            .ok_or_else(|| not_a_picture(OPERATION, shape_id))?;
        let relationship_id = picture_image_relationship(picture, OPERATION, shape_id)?.to_owned();
        let old_part = require_internal_related_part(
            &staged.package,
            &slide_part,
            &relationship_id,
            rel_types::IMAGE,
        )?;
        let mut relationships = staged
            .package
            .get_part_rels(&slide_part)
            .cloned()
            .unwrap_or_default();
        let picture_xml = picture
            .to_xml()
            .map_err(|error| invalid_shape_mutation(OPERATION, error.to_string()))?;
        let second_image = relationship_ids(&picture_xml)
            .map_err(|error| invalid_shape_mutation(OPERATION, error.to_string()))?
            .into_iter()
            .filter(|id| *id != relationship_id)
            .any(|id| {
                relationships
                    .get_by_id(&id)
                    .is_some_and(|relationship| relationship.rel_type == rel_types::IMAGE)
            });
        if second_image {
            return Err(invalid_shape_mutation(
                OPERATION,
                format!("picture {shape_id} carries a second image, such as an SVG alternate"),
            ));
        }
        let slide_xml = slide
            .to_xml()
            .map_err(|error| invalid_shape_mutation(OPERATION, error.to_string()))?;
        let uses = relationship_ids(&slide_xml)
            .map_err(|error| invalid_shape_mutation(OPERATION, error.to_string()))?
            .into_iter()
            .filter(|id| *id == relationship_id)
            .count();
        let picture = find_picture_mut(&mut slide.common_slide_data.shape_tree.children, shape_id)
            .expect("picture found above");
        let relationship_id = if uses > 1 {
            let isolated = relationships.add(
                rel_types::IMAGE,
                &relative_part_target(&slide_part, &old_part),
            );
            picture
                .blip_fill
                .as_mut()
                .and_then(|fill| fill.blip.as_mut())
                .expect("picture has an embedded image")
                .embed = Some(isolated.clone());
            isolated
        } else {
            relationship_id
        };
        let part_users = std::iter::once(("/", &staged.package.package_rels))
            .chain(
                staged
                    .package
                    .part_rels
                    .iter()
                    .filter(|(source, _)| **source != slide_part)
                    .map(|(source, items)| (source.as_str(), items)),
            )
            .chain(std::iter::once((slide_part.as_str(), &relationships)))
            .flat_map(|(source, items)| items.items.iter().map(move |item| (source, item)))
            .filter(|(source, relationship)| {
                !relationship_is_external(relationship)
                    && OpcPackage::resolve_rel_target(source, &relationship.target)
                        .eq_ignore_ascii_case(&old_part)
            })
            .count();
        let extension = old_part
            .rsplit_once('.')
            .map_or("", |(_, extension)| extension)
            .to_owned();
        if part_users == 1 && ImageFormat::from_extension(&extension) == Some(format) {
            staged.package.set_part(&old_part, image_data.to_vec());
            register_content_type(
                &mut staged.package,
                &old_part,
                &extension,
                format.content_type(),
            );
        } else {
            let new_part = staged.media_store.insert(
                &mut staged.package,
                image_data,
                &format!("image.{}", format.extension()),
            );
            let relationship = relationships
                .items
                .iter_mut()
                .find(|relationship| relationship.id == relationship_id)
                .expect("picture relationship exists");
            relationship.target = relative_part_target(&slide_part, &new_part);
        }
        staged.package.set_part_rels(&slide_part, relationships);
        prune_unreachable_parts(&mut staged.package, &HashSet::from([old_part]));
        staged.media_store = MediaStore::scan(&staged.package);
        self.commit_candidate(staged)
    }

    /// Removes one immediate slide child and the package parts only it used.
    ///
    /// Connectors attached to a removed shape are detached, as PowerPoint does
    /// on delete. The media timing a removed media picture owns goes with it.
    /// A shape that other slide animations still target is rejected without
    /// any change.
    pub fn remove_shape(&mut self, slide_index: usize, shape_index: usize) -> Result<()> {
        const OPERATION: &str = "remove shape";
        self.require_slide_index(slide_index)?;
        let mut staged = self.clone();
        let slide_part = staged.slides[slide_index].part_name.clone();
        let slide = &mut staged.slides[slide_index].slide;
        let children = &slide.common_slide_data.shape_tree.children;
        let child = children
            .get(shape_index)
            .ok_or_else(|| Error::InvalidSlideMutation {
                operation: OPERATION,
                message: format!(
                    "shape index {shape_index} is out of range for {} shapes",
                    children.len()
                ),
            })?;
        let shape_id = child
            .non_visual_id()
            .ok_or_else(|| Error::UnsupportedShapeMutation {
                operation: OPERATION,
                shape_kind: shape_kind(child),
            })?;
        if children
            .iter()
            .position(|child| child.non_visual_id() == Some(shape_id))
            != Some(shape_index)
        {
            return Err(Error::InvalidSlideMutation {
                operation: OPERATION,
                message: format!("shape id {shape_id} is not unique on the slide"),
            });
        }
        let mut removed_ids = HashSet::new();
        let mut media_ids = Vec::new();
        collect_subtree_ids(
            std::slice::from_ref(child),
            &mut removed_ids,
            &mut media_ids,
        );
        let before = slide_relationship_ids(slide)?;
        if let Some(timing) = &mut slide.timing {
            for media_id in media_ids {
                timing
                    .remove_media(media_id)
                    .map_err(|error| invalid_shape_mutation(OPERATION, error.to_string()))?;
            }
            if let Some(animated) = removed_ids
                .iter()
                .copied()
                .filter(|id| timing.references_shape(*id))
                .min()
            {
                return Err(invalid_shape_mutation(
                    OPERATION,
                    format!("slide animations still target shape id {animated}"),
                ));
            }
        }
        detach_connectors(
            &mut slide.common_slide_data.shape_tree.children,
            &removed_ids,
        );
        slide
            .common_slide_data
            .shape_tree
            .remove_child_by_id(shape_id)
            .map_err(|error| invalid_shape_mutation(OPERATION, error.to_string()))?;
        let after = slide_relationship_ids(slide)?;
        let mut candidates = HashSet::new();
        if let Some(relationships) = staged.package.get_part_rels_mut(&slide_part) {
            relationships.items.retain(|relationship| {
                if !before.contains(&relationship.id) || after.contains(&relationship.id) {
                    return true;
                }
                if !relationship_is_external(relationship) {
                    candidates.insert(OpcPackage::resolve_rel_target(
                        &slide_part,
                        &relationship.target,
                    ));
                }
                false
            });
        }
        prune_unreachable_parts(&mut staged.package, &candidates);
        staged.media_store = MediaStore::scan(&staged.package);
        self.commit_candidate(staged)
    }

    /// Returns the target of one relationship of a slide, the address that a
    /// hyperlink naming this relationship id opens.
    ///
    /// Pass the `r:id` of an `a:hlinkClick`. An external hyperlink returns
    /// its URL, and an internal relationship, such as a jump to another
    /// slide, returns its stored relative target, as python-pptx `address`
    /// does. An unknown slide or id returns `None`.
    pub fn hyperlink_address(&self, slide_index: usize, relationship_id: &str) -> Option<&str> {
        let record = self.slides.get(slide_index)?;
        self.package
            .get_part_rels(&record.part_name)?
            .get_by_id(relationship_id)
            .map(|relationship| relationship.target.as_str())
    }

    /// Resolves the external click hyperlink of one slide shape by its non-visual id.
    pub fn shape_hyperlink_address(
        &self,
        slide_index: usize,
        shape_id: u32,
    ) -> Result<Option<&str>> {
        let (_, location) = self.locate_shape_click("shape hyperlink", slide_index, shape_id)?;
        Ok(location
            .relationship_id
            .as_deref()
            .and_then(|id| self.hyperlink_address(slide_index, id)))
    }

    /// Returns the zero-based index of the slide that the click action of one
    /// slide shape jumps to, as python-pptx `click_action.target_slide` does.
    ///
    /// The shape is found by its non-visual id as for
    /// [`Self::shape_hyperlink_address`]. A named slide jump,
    /// `ppaction://hlinksldjump`, returns the slide its relationship targets.
    /// A `ppaction://hlinkshowjump` to the first, last, next, or previous
    /// slide returns that slide, counted from `slide_index`. Any other click
    /// action, or none, returns `None`, and so does a jump that names no slide
    /// of the presentation, such as a next-slide jump on the last slide, where
    /// python-pptx raises `ValueError`.
    pub fn shape_target_slide(&self, slide_index: usize, shape_id: u32) -> Result<Option<usize>> {
        let (_, location) = self.locate_shape_click("shape target slide", slide_index, shape_id)?;
        Ok(self.click_target_slide(
            slide_index,
            location.action.as_deref(),
            location.relationship_id.as_deref(),
        ))
    }

    /// Returns the slide a click action of slide `slide_index` jumps to, from
    /// its `action` and `r:id`, as [`Self::shape_target_slide`] reads them.
    fn click_target_slide(
        &self,
        slide_index: usize,
        action: Option<&str>,
        relationship_id: Option<&str>,
    ) -> Option<usize> {
        let action = action?;
        let (verb, query) = action.split_once('?').unwrap_or((action, ""));
        match verb {
            SLIDE_JUMP_ACTION => {
                let part_name = &self.slides[slide_index].part_name;
                relationship_id
                    .and_then(|id| self.package.get_part_rels(part_name)?.get_by_id(id))
                    .filter(|relationship| !relationship_is_external(relationship))
                    .and_then(|relationship| {
                        let target =
                            OpcPackage::resolve_rel_target(part_name, &relationship.target);
                        self.slides
                            .iter()
                            .position(|slide| slide.part_name == target)
                    })
            }
            SHOW_JUMP_ACTION => {
                match query
                    .split('&')
                    .find_map(|field| field.strip_prefix("jump="))
                {
                    Some("firstslide") => Some(0),
                    Some("lastslide") => self.slides.len().checked_sub(1),
                    Some("nextslide") => {
                        Some(slide_index + 1).filter(|next| *next < self.slides.len())
                    }
                    Some("previousslide") => slide_index.checked_sub(1),
                    _ => None,
                }
            }
            _ => None,
        }
    }

    /// Sets or clears the external click hyperlink of one slide shape atomically.
    ///
    /// A new address replaces any other click action of the shape, such as a
    /// slide jump, and keeps the other attributes and children of an existing
    /// `a:hlinkClick`.
    pub fn set_shape_hyperlink(
        &mut self,
        slide_index: usize,
        shape_id: u32,
        address: Option<&str>,
    ) -> Result<()> {
        const OPERATION: &str = "set shape hyperlink";
        self.require_slide_index(slide_index)?;
        if address.is_some_and(|address| {
            address.is_empty()
                || address.chars().any(|character| {
                    character.is_control() || matches!(character, '\u{FFFE}' | '\u{FFFF}')
                })
        }) {
            return Err(invalid_shape_mutation(
                OPERATION,
                "a hyperlink address must be non-empty text without control characters",
            ));
        }
        self.set_shape_click_action(
            OPERATION,
            slide_index,
            shape_id,
            address.map(ShapeClickTarget::Address),
        )
    }

    /// Makes the click action of one slide shape jump to the slide at
    /// `target_slide_index`, or clears it with `None`, as python-pptx
    /// `click_action.target_slide` does.
    ///
    /// The shape is found as for [`Self::set_shape_hyperlink`]. Its
    /// `a:hlinkClick` gets the action `ppaction://hlinksldjump` and names the
    /// slide's relationship to the target slide, which is reused when the
    /// slide has one. Jumping to the current target changes nothing, and a
    /// relationship that only the old click action named is removed.
    pub fn set_shape_target_slide(
        &mut self,
        slide_index: usize,
        shape_id: u32,
        target_slide_index: Option<usize>,
    ) -> Result<()> {
        if let Some(target) = target_slide_index {
            self.require_slide_index(target)?;
        }
        self.set_shape_click_action(
            "set shape target slide",
            slide_index,
            shape_id,
            target_slide_index.map(ShapeClickTarget::Slide),
        )
    }

    /// Serializes one slide and finds the `p:cNvPr` of its single child or
    /// group member whose id is `shape_id`, with its click action.
    fn locate_shape_click(
        &self,
        operation: &'static str,
        slide_index: usize,
        shape_id: u32,
    ) -> Result<(Vec<u8>, ShapeHyperlinkLocation)> {
        self.require_slide_index(slide_index)?;
        let record = &self.slides[slide_index];
        if shape_id_count(
            &record.slide.common_slide_data.shape_tree.children,
            shape_id,
        ) != 1
        {
            return Err(invalid_shape_mutation(
                operation,
                "shape id is missing or ambiguous",
            ));
        }
        let xml = record
            .slide
            .to_xml()
            .map_err(|error| invalid_shape_mutation(operation, error.to_string()))?;
        let location = locate_shape_hyperlink(&xml, shape_id)?
            .ok_or_else(|| invalid_shape_mutation(operation, "shape has no p:cNvPr"))?;
        Ok((xml, location))
    }

    fn set_shape_click_action(
        &mut self,
        operation: &'static str,
        slide_index: usize,
        shape_id: u32,
        target: Option<ShapeClickTarget<'_>>,
    ) -> Result<()> {
        let (xml, location) = self.locate_shape_click(operation, slide_index, shape_id)?;
        let part_name = self.slides[slide_index].part_name.clone();
        let mut relationships = self
            .package
            .get_part_rels(&part_name)
            .cloned()
            .unwrap_or_default();
        let current = location
            .relationship_id
            .as_deref()
            .and_then(|id| relationships.get_by_id(id));
        let action = location.action.as_deref();
        let link = match target {
            None if location.click.is_none() => return Ok(()),
            None => None,
            Some(ShapeClickTarget::Address(address)) => {
                if action.is_none()
                    && current.is_some_and(|relationship| {
                        relationship.rel_type == rel_types::HYPERLINK
                            && relationship_is_external(relationship)
                            && relationship.target == address
                    })
                {
                    return Ok(());
                }
                let id = relationships
                    .items
                    .iter()
                    .find(|relationship| {
                        relationship.rel_type == rel_types::HYPERLINK
                            && relationship_is_external(relationship)
                            && relationship.target == address
                    })
                    .map(|relationship| relationship.id.clone())
                    .unwrap_or_else(|| relationships.add_external(rel_types::HYPERLINK, address));
                Some((id, None))
            }
            Some(ShapeClickTarget::Slide(index)) => {
                let target_part = &self.slides[index].part_name;
                let is_target = |relationship: &Relationship| {
                    relationship.rel_type == rel_types::SLIDE
                        && !relationship_is_external(relationship)
                        && OpcPackage::resolve_rel_target(&part_name, &relationship.target)
                            == *target_part
                };
                if action == Some(SLIDE_JUMP_ACTION) && current.is_some_and(is_target) {
                    return Ok(());
                }
                let id = relationships
                    .items
                    .iter()
                    .find(|relationship| is_target(relationship))
                    .map(|relationship| relationship.id.clone())
                    .unwrap_or_else(|| {
                        relationships.add(
                            rel_types::SLIDE,
                            &relative_part_target(&part_name, target_part),
                        )
                    });
                Some((id, Some(SLIDE_JUMP_ACTION)))
            }
        };
        let rewritten = rewrite_shape_hyperlink(
            &xml,
            &location,
            link.as_ref().map(|(id, action)| (id.as_str(), *action)),
        )?;
        let slide = CT_Slide::from_xml(&rewritten)
            .map_err(|error| invalid_shape_mutation(operation, error.to_string()))?;
        let before = slide_relationship_ids(&self.slides[slide_index].slide)?;
        let after = slide_relationship_ids(&slide)?;
        relationships.items.retain(|relationship| {
            !before.contains(&relationship.id) || after.contains(&relationship.id)
        });
        self.slides[slide_index].slide = slide;
        self.package.set_part_rels(&part_name, relationships);
        Ok(())
    }

    /// Points one text run's click hyperlink at an external `address`, or
    /// removes it with `None`, as python-pptx `hyperlink.address` does.
    ///
    /// The run is the zero-based regular run of a paragraph in the ordinary
    /// shape whose `p:cNvPr/@id` is `shape_id`, inside groups too. It gets an
    /// `a:hlinkClick` naming the slide's external hyperlink relationship to
    /// `address`, which is reused when the slide has one. Assigning the
    /// current address changes nothing. Every relationship that only the old
    /// hyperlink named, its target or a click sound, is removed, so no
    /// relationship is left unreferenced. An empty address, one with a
    /// control character, or a shape id that is missing, shared, or not an
    /// ordinary shape with that run is rejected without change.
    pub fn set_run_hyperlink(
        &mut self,
        slide_index: usize,
        shape_id: u32,
        paragraph_index: usize,
        run_index: usize,
        address: Option<&str>,
    ) -> Result<()> {
        const OPERATION: &str = "set run hyperlink";
        self.require_slide_index(slide_index)?;
        if address.is_some_and(|address| {
            address.is_empty()
                || address.chars().any(|character| {
                    character.is_control() || matches!(character, '\u{FFFE}' | '\u{FFFF}')
                })
        }) {
            return Err(invalid_shape_mutation(
                OPERATION,
                "a hyperlink address must be non-empty text without control characters",
            ));
        }
        self.set_run_click(
            OPERATION,
            slide_index,
            (shape_id, paragraph_index, run_index),
            address.map(ShapeClickTarget::Address),
        )
    }

    /// Makes one text run's click hyperlink jump to the slide at
    /// `target_slide_index`, or removes it with `None`, as
    /// [`Self::set_shape_target_slide`] does for a shape.
    ///
    /// The run is found as for [`Self::set_run_hyperlink`]. Its
    /// `a:hlinkClick` gets the action `ppaction://hlinksldjump` and names the
    /// slide's relationship to the target slide, which is reused when the
    /// slide has one. Jumping to the current target changes nothing, and a
    /// relationship that only the old hyperlink named is removed.
    pub fn set_run_target_slide(
        &mut self,
        slide_index: usize,
        shape_id: u32,
        paragraph_index: usize,
        run_index: usize,
        target_slide_index: Option<usize>,
    ) -> Result<()> {
        self.require_slide_index(slide_index)?;
        if let Some(target) = target_slide_index {
            self.require_slide_index(target)?;
        }
        self.set_run_click(
            "set run target slide",
            slide_index,
            (shape_id, paragraph_index, run_index),
            target_slide_index.map(ShapeClickTarget::Slide),
        )
    }

    /// Returns the zero-based index of the slide one text run's click
    /// hyperlink jumps to, read as [`Self::shape_target_slide`] reads a
    /// shape's. The run is found as for [`Self::set_run_hyperlink`], and a
    /// run without a slide jump returns `None`.
    pub fn run_target_slide(
        &self,
        slide_index: usize,
        shape_id: u32,
        paragraph_index: usize,
        run_index: usize,
    ) -> Result<Option<usize>> {
        self.require_slide_index(slide_index)?;
        let missing = || {
            invalid_shape_mutation(
                "read run target slide",
                format!(
                    "shape id {shape_id} has no ordinary shape run {run_index} in paragraph {paragraph_index}"
                ),
            )
        };
        let body = find_shape(
            &self.slides[slide_index]
                .slide
                .common_slide_data
                .shape_tree
                .children,
            shape_id,
        )
        .and_then(|shape| shape.text_body.as_ref())
        .ok_or_else(missing)?;
        let hyperlink = TextFrameRef { body }
            .paragraph(paragraph_index)
            .and_then(|paragraph| paragraph.run(run_index))
            .ok_or_else(missing)?
            .properties()
            .and_then(|properties| properties.hyperlink_click.clone());
        Ok(hyperlink.and_then(|hyperlink| {
            self.click_target_slide(
                slide_index,
                hyperlink.action.as_deref(),
                hyperlink.relationship_id.as_deref(),
            )
        }))
    }

    fn set_run_click(
        &mut self,
        operation: &'static str,
        slide_index: usize,
        (shape_id, paragraph_index, run_index): (u32, usize, usize),
        target: Option<ShapeClickTarget<'_>>,
    ) -> Result<()> {
        let record = &self.slides[slide_index];
        let part_name = record.part_name.clone();
        if shape_id_count(
            &record.slide.common_slide_data.shape_tree.children,
            shape_id,
        ) > 1
        {
            return Err(invalid_shape_mutation(
                operation,
                format!("shape id {shape_id} is not unique on the slide"),
            ));
        }
        let target_part = match target {
            Some(ShapeClickTarget::Slide(index)) => Some(self.slides[index].part_name.clone()),
            _ => None,
        };
        let mut slide = record.slide.clone();
        let mut relationships = self
            .package
            .get_part_rels(&part_name)
            .cloned()
            .unwrap_or_default();
        let (old, new) = {
            let missing = || {
                invalid_shape_mutation(
                    operation,
                    format!(
                        "shape id {shape_id} has no ordinary shape run {run_index} in paragraph {paragraph_index}"
                    ),
                )
            };
            let body = find_shape_mut(&mut slide.common_slide_data.shape_tree.children, shape_id)
                .and_then(|shape| shape.text_body.as_mut())
                .ok_or_else(missing)?;
            let mut paragraph = TextFrame { body }
                .into_paragraph_mut(paragraph_index)
                .ok_or_else(missing)?;
            let mut run = paragraph.run_mut(run_index).ok_or_else(missing)?;
            let mut properties = run.properties().cloned().unwrap_or_default();
            let old = properties
                .hyperlink_click
                .as_ref()
                .and_then(|hyperlink| hyperlink.relationship_id.clone())
                .filter(|id| !id.is_empty());
            let action = properties
                .hyperlink_click
                .as_ref()
                .and_then(|hyperlink| hyperlink.action.as_deref());
            let current = old.as_deref().and_then(|id| relationships.get_by_id(id));
            let is_address = |relationship: &Relationship, address: &str| {
                relationship.rel_type == rel_types::HYPERLINK
                    && relationship_is_external(relationship)
                    && relationship.target == address
            };
            let is_slide = |relationship: &Relationship, target_part: &str| {
                relationship.rel_type == rel_types::SLIDE
                    && !relationship_is_external(relationship)
                    && OpcPackage::resolve_rel_target(&part_name, &relationship.target)
                        == target_part
            };
            let unchanged = match (target, target_part.as_deref()) {
                (None, _) => properties.hyperlink_click.is_none(),
                (Some(ShapeClickTarget::Address(address)), _) => {
                    current.is_some_and(|relationship| is_address(relationship, address))
                }
                (Some(ShapeClickTarget::Slide(_)), Some(target_part)) => {
                    action == Some(SLIDE_JUMP_ACTION)
                        && current.is_some_and(|relationship| is_slide(relationship, target_part))
                }
                (Some(ShapeClickTarget::Slide(_)), None) => unreachable!("the part was read"),
            };
            if unchanged {
                return Ok(());
            }
            let new = match (target, target_part.as_deref()) {
                (None, _) => None,
                (Some(ShapeClickTarget::Address(address)), _) => Some((
                    relationships
                        .items
                        .iter()
                        .find(|relationship| is_address(relationship, address))
                        .map(|relationship| relationship.id.clone())
                        .unwrap_or_else(|| {
                            relationships.add_external(rel_types::HYPERLINK, address)
                        }),
                    None,
                )),
                (Some(ShapeClickTarget::Slide(_)), Some(target_part)) => Some((
                    relationships
                        .items
                        .iter()
                        .find(|relationship| is_slide(relationship, target_part))
                        .map(|relationship| relationship.id.clone())
                        .unwrap_or_else(|| {
                            relationships.add(
                                rel_types::SLIDE,
                                &relative_part_target(&part_name, target_part),
                            )
                        }),
                    Some(SLIDE_JUMP_ACTION),
                )),
                (Some(ShapeClickTarget::Slide(_)), None) => unreachable!("the part was read"),
            };
            properties.hyperlink_click = new.clone().map(|(id, action)| {
                let mut hyperlink = TextHyperlink::default();
                hyperlink.relationship_id = Some(id);
                hyperlink.action = action.map(str::to_owned);
                hyperlink
            });
            run.set_properties(properties);
            (old, new.map(|(id, _)| id))
        };
        if old.is_some_and(|old| Some(&old) != new.as_ref()) {
            // The old a:hlinkClick can carry more than its r:id, such as an
            // a:snd click sound naming an audio relationship. Remove every
            // relationship this edit left without a reference.
            let before = slide_relationship_ids(&self.slides[slide_index].slide)?;
            let after = slide_relationship_ids(&slide)?;
            relationships.items.retain(|relationship| {
                !before.contains(&relationship.id) || after.contains(&relationship.id)
            });
        }
        self.slides[slide_index].slide = slide;
        self.package.set_part_rels(&part_name, relationships);
        Ok(())
    }

    /// Moves one immediate slide child so that it ends up at z-order index
    /// `to_index`, where later children draw on top.
    ///
    /// Animations, connector glue, and relationships refer to shape ids, so
    /// they follow the moved shape. Unmodelled shape-tree members keep their
    /// place among the other children.
    pub fn move_shape(
        &mut self,
        slide_index: usize,
        from_index: usize,
        to_index: usize,
    ) -> Result<()> {
        const OPERATION: &str = "move shape";
        self.require_slide_index(slide_index)?;
        let tree = &mut self.slides[slide_index].slide.common_slide_data.shape_tree;
        let count = tree.children.len();
        if let Some(index) = [from_index, to_index]
            .into_iter()
            .find(|index| *index >= count)
        {
            return Err(invalid_slide_mutation(
                OPERATION,
                format!("shape index {index} is out of range for {count} shapes"),
            ));
        }
        tree.move_child(from_index, to_index)
            .map_err(|error| invalid_shape_mutation(OPERATION, error.to_string()))
    }

    /// Inspects slide, layout, and master SmartArt in producing-scope order.
    pub fn smart_art(&self, slide_index: usize) -> Result<Vec<SmartArtInfo>> {
        let record = self
            .slides
            .get(slide_index)
            .ok_or(Error::UnknownSlideIndex {
                index: slide_index,
                slide_count: self.slides.len(),
            })?;
        let mut infos = Vec::new();
        collect_smartart_infos(
            &self.package,
            slide_index,
            &record.part_name,
            &record.slide.common_slide_data.shape_tree.children,
            &mut infos,
        );
        let Some(layout_part) =
            related_internal_part(&self.package, &record.part_name, rel_types::SLIDE_LAYOUT)?
        else {
            return Ok(infos);
        };
        let layout = CT_SlideLayout::from_xml(required_part(&self.package, &layout_part)?)
            .map_err(|error| Error::MalformedPart {
                part_name: layout_part.clone(),
                message: error.to_string(),
            })?;
        collect_smartart_infos(
            &self.package,
            slide_index,
            &layout_part,
            &layout.common_slide_data.shape_tree.children,
            &mut infos,
        );
        let Some(master_part) =
            related_internal_part(&self.package, &layout_part, rel_types::SLIDE_MASTER)?
        else {
            return Ok(infos);
        };
        let master = CT_SlideMaster::from_xml(required_part(&self.package, &master_part)?)
            .map_err(|error| Error::MalformedPart {
                part_name: master_part.clone(),
                message: error.to_string(),
            })?;
        collect_smartart_infos(
            &self.package,
            slide_index,
            &master_part,
            &master.common_slide_data.shape_tree.children,
            &mut infos,
        );
        Ok(infos)
    }

    /// Replaces one supported SmartArt node's text atomically.
    pub fn set_smart_art_node_text(
        &mut self,
        slide_index: usize,
        shape_id: u32,
        model_id: &str,
        text: &str,
    ) -> Result<()> {
        let record = self
            .slides
            .get(slide_index)
            .ok_or(Error::UnknownSlideIndex {
                index: slide_index,
                slide_count: self.slides.len(),
            })?;
        let relationships = find_smartart_relationships(
            &record.slide.common_slide_data.shape_tree.children,
            shape_id,
        )
        .ok_or_else(|| Error::InvalidPresentationMutation {
            operation: "set SmartArt node text",
            message: format!("shape {shape_id} is not a SmartArt frame"),
        })?
        .clone();
        let slide_part = record.part_name.clone();
        let (data_part, mut data) =
            validate_smartart_relationship_roles(&self.package, &slide_part, &relationships)?;
        data.set_node_text(model_id, text)
            .map_err(|error| Error::InvalidPresentationMutation {
                operation: "set SmartArt node text",
                message: error.to_string(),
            })?;
        let xml = data.to_xml().map_err(|error| Error::MalformedPart {
            part_name: data_part.clone(),
            message: error.to_string(),
        })?;
        CT_DiagramData::from_xml(&xml).map_err(|error| Error::MalformedPart {
            part_name: data_part.clone(),
            message: error.to_string(),
        })?;

        let mut staged = self.clone();
        staged.package.set_part(&data_part, xml);
        let issues = staged.validate();
        if !issues.is_empty() {
            return Err(Error::InvalidPresentationMutation {
                operation: "set SmartArt node text",
                message: format!("staged diagram graph is invalid: {issues:?}"),
            });
        }
        let reopened = Self::from_bytes(&staged.to_bytes()?)?;
        *self = reopened;
        Ok(())
    }

    /// Inspects the media pictures on one slide in z-order.
    pub fn media(&self, slide_index: usize) -> Result<Vec<MediaInfo>> {
        let record = self
            .slides
            .get(slide_index)
            .ok_or(Error::UnknownSlideIndex {
                index: slide_index,
                slide_count: self.slides.len(),
            })?;
        media_infos_for_slide(&self.package, slide_index, &record.part_name, &record.slide)
    }

    /// Adds one media picture, poster, relationship scope, and timing node atomically.
    #[allow(clippy::too_many_arguments)]
    pub fn add_media(
        &mut self,
        slide_index: usize,
        kind: MediaKind,
        source: MediaSourceInput<'_>,
        poster: MediaPoster<'_>,
        left: Emu,
        top: Emu,
        width: Emu,
        height: Emu,
        settings: MediaPlaybackSettings,
    ) -> Result<ShapeRef<'_>> {
        self.require_slide_index(slide_index)?;
        validate_media_settings(settings)?;
        validate_media_source(source)?;
        if ImageFormat::sniff(poster.bytes).is_none() {
            return Err(invalid_media_mutation(
                "add media",
                format!("poster {} has unsupported image bytes", poster.filename),
            ));
        }
        if width.0 <= 0 || height.0 <= 0 {
            return Err(invalid_media_mutation(
                "add media",
                "media width and height must be positive",
            ));
        }

        let mut staged = self.clone();
        let slide_part = staged.slides[slide_index].part_name.clone();
        let shape_id = ShapeIdAllocator::scan(
            &staged.slides[slide_index]
                .slide
                .common_slide_data
                .shape_tree,
        )
        .allocate();
        let poster_part =
            staged
                .media_store
                .insert(&mut staged.package, poster.bytes, poster.filename);
        let mut relationships = staged
            .package
            .get_part_rels(&slide_part)
            .cloned()
            .unwrap_or_default();
        let poster_relationship_id = relationships.add(
            rel_types::IMAGE,
            &relative_part_target(&slide_part, &poster_part),
        );
        let (standard_relationship_id, microsoft_relationship_id, picture_source) =
            add_media_relationships(
                &mut staged.package,
                &mut staged.media_store,
                &slide_part,
                &mut relationships,
                kind,
                source,
                true,
            )?;
        let microsoft_relationship_id = microsoft_relationship_id
            .expect("new media pictures always include the Office media relationship");
        let picture = CT_Picture::new_media(
            shape_id,
            &format!("{} {shape_id}", media_kind_name(kind)),
            kind,
            match picture_source {
                PictureMediaSource::Embedded { .. } => PictureMediaSource::Embedded {
                    relationship_id: standard_relationship_id,
                },
                PictureMediaSource::Linked { .. } => PictureMediaSource::Linked {
                    relationship_id: standard_relationship_id,
                },
            },
            &microsoft_relationship_id,
            &poster_relationship_id,
            settings.trim_start_ms,
            settings.trim_end_ms,
            positioned_transform(left, top, width, height),
        )
        .map_err(|error| invalid_media_mutation("add media", error.to_string()))?;
        staged.slides[slide_index]
            .slide
            .common_slide_data
            .shape_tree
            .append_child(ShapeTreeChild::Picture(picture));
        let timing = staged.slides[slide_index]
            .slide
            .timing
            .get_or_insert_with(|| {
                CT_Timing::from_xml(
                    format!(r#"<p:timing xmlns:p="{P_NS}"><p:tnLst/></p:timing>"#).as_bytes(),
                )
                .expect("canonical media timing shell")
            });
        timing
            .add_media(
                kind == MediaKind::Video,
                shape_id,
                settings.volume,
                settings.looped,
                settings.show_when_stopped,
                settings.trigger,
            )
            .map_err(|error| invalid_media_mutation("add media", error.to_string()))?;
        staged.package.set_part_rels(&slide_part, relationships);
        self.commit_candidate(staged)?;
        self.slides[slide_index]
            .slide
            .common_slide_data
            .shape_tree
            .children
            .iter()
            .find(|child| child.non_visual_id() == Some(shape_id))
            .map(shape_ref)
            .ok_or_else(|| invalid_media_mutation("add media", "authored picture was not retained"))
    }

    /// Replaces one media source while preserving its shape, poster, and timing.
    pub fn replace_media(
        &mut self,
        slide_index: usize,
        shape_id: u32,
        source: MediaSourceInput<'_>,
    ) -> Result<()> {
        self.require_slide_index(slide_index)?;
        validate_media_source(source)?;
        let mut staged = self.clone();
        let slide_part = staged.slides[slide_index].part_name.clone();
        let picture = media_picture_shape(&staged.slides[slide_index].slide, shape_id)?;
        let media = picture.media.as_ref().expect("media picture").clone();
        let standard_relationship_id = picture.standard_media_relationship_id().map(str::to_owned);
        let has_office_media = !picture.office_media_relationship_ids().is_empty();
        let mut old_relationship_ids = picture.office_media_relationship_ids().to_vec();
        if !old_relationship_ids
            .iter()
            .any(|id| id == media.source.relationship_id())
        {
            old_relationship_ids.push(media.source.relationship_id().to_owned());
        }
        if let Some(standard_relationship_id) = &standard_relationship_id
            && !old_relationship_ids.contains(standard_relationship_id)
        {
            old_relationship_ids.push(standard_relationship_id.clone());
        }
        let mut relationships = staged
            .package
            .get_part_rels(&slide_part)
            .cloned()
            .unwrap_or_default();
        let old_relationships = old_relationship_ids
            .iter()
            .map(|relationship_id| {
                relationships
                    .get_by_id(relationship_id)
                    .cloned()
                    .ok_or_else(|| Error::MissingRelationship {
                        source_part: slide_part.clone(),
                        relationship_id: relationship_id.clone(),
                    })
            })
            .collect::<Result<Vec<_>>>()?;
        let (standard_id, microsoft_id, picture_source) = add_media_relationships(
            &mut staged.package,
            &mut staged.media_store,
            &slide_part,
            &mut relationships,
            media.kind,
            source,
            has_office_media,
        )?;
        let source_id = microsoft_id.unwrap_or_else(|| standard_id.clone());
        let replacement_source = match picture_source {
            PictureMediaSource::Embedded { .. } => PictureMediaSource::Embedded {
                relationship_id: source_id,
            },
            PictureMediaSource::Linked { .. } => PictureMediaSource::Linked {
                relationship_id: source_id,
            },
        };
        let target_picture = staged.slides[slide_index]
            .slide
            .common_slide_data
            .shape_tree
            .children
            .iter_mut()
            .find(|child| child.non_visual_id() == Some(shape_id))
            .and_then(|child| match child {
                ShapeTreeChild::Picture(picture) => Some(picture),
                _ => None,
            })
            .ok_or_else(|| invalid_media_mutation("replace media", "media picture is missing"))?;
        target_picture
            .replace_media_relationship_ids(&standard_id, replacement_source)
            .map_err(|error| invalid_media_mutation("replace media", error.to_string()))?;
        let referenced = slide_relationship_ids(&staged.slides[slide_index].slide)?;
        let mut candidates = HashSet::new();
        relationships.items.retain(|relationship| {
            let Some(old) = old_relationships
                .iter()
                .find(|old| old.id == relationship.id)
            else {
                return true;
            };
            if referenced.contains(&old.id) {
                return true;
            }
            if !relationship_is_external(old) {
                candidates.insert(OpcPackage::resolve_rel_target(&slide_part, &old.target));
            }
            false
        });
        staged.package.set_part_rels(&slide_part, relationships);
        prune_unreachable_parts(&mut staged.package, &candidates);
        staged.media_store = MediaStore::scan(&staged.package);
        self.commit_candidate(staged)
    }

    /// Extracts embedded media bytes, or returns `None` for linked media.
    pub fn extract_media(&self, slide_index: usize, shape_id: u32) -> Result<Option<Vec<u8>>> {
        let record = self
            .slides
            .get(slide_index)
            .ok_or(Error::UnknownSlideIndex {
                index: slide_index,
                slide_count: self.slides.len(),
            })?;
        let media = media_picture(&record.slide, shape_id)?;
        let relationship_id = media.source.relationship_id();
        let relationship = self
            .package
            .get_part_rels(&record.part_name)
            .and_then(|relationships| relationships.get_by_id(relationship_id))
            .ok_or_else(|| Error::MissingRelationship {
                source_part: record.part_name.clone(),
                relationship_id: relationship_id.to_owned(),
            })?;
        if relationship_is_external(relationship) {
            return Ok(None);
        }
        let part_name = OpcPackage::resolve_rel_target(&record.part_name, &relationship.target);
        Ok(Some(required_part(&self.package, &part_name)?.to_vec()))
    }

    /// Removes one media picture and only its relationship-owned package candidates.
    pub fn remove_media(&mut self, slide_index: usize, shape_id: u32) -> Result<()> {
        self.require_slide_index(slide_index)?;
        let mut staged = self.clone();
        let slide_part = staged.slides[slide_index].part_name.clone();
        let picture = media_picture_shape(&staged.slides[slide_index].slide, shape_id)?;
        let media = picture.media.as_ref().expect("media picture").clone();
        let standard_relationship_id = picture.standard_media_relationship_id().map(str::to_owned);
        let office_relationship_ids = picture.office_media_relationship_ids().to_vec();
        let mut relationships = staged
            .package
            .get_part_rels(&slide_part)
            .cloned()
            .unwrap_or_default();
        let mut owned_ids = office_relationship_ids.into_iter().collect::<HashSet<_>>();
        owned_ids.insert(media.source.relationship_id().to_owned());
        if let Some(standard_relationship_id) = standard_relationship_id {
            owned_ids.insert(standard_relationship_id);
        }
        if let Some(poster_relationship_id) = media.poster_relationship_id {
            owned_ids.insert(poster_relationship_id);
        }
        for relationship_id in &owned_ids {
            relationships
                .get_by_id(relationship_id)
                .ok_or_else(|| Error::MissingRelationship {
                    source_part: slide_part.clone(),
                    relationship_id: relationship_id.clone(),
                })?;
        }
        staged.slides[slide_index]
            .slide
            .common_slide_data
            .shape_tree
            .remove_child_by_id(shape_id)
            .map_err(|error| invalid_media_mutation("remove media", error.to_string()))?
            .ok_or_else(|| invalid_media_mutation("remove media", "media picture is missing"))?;
        if let Some(timing) = &mut staged.slides[slide_index].slide.timing {
            timing
                .remove_media(shape_id)
                .map_err(|error| invalid_media_mutation("remove media", error.to_string()))?;
        }
        let referenced = slide_relationship_ids(&staged.slides[slide_index].slide)?;
        let mut candidates = HashSet::new();
        relationships.items.retain(|relationship| {
            if !owned_ids.contains(&relationship.id) || referenced.contains(&relationship.id) {
                return true;
            }
            if !relationship_is_external(relationship) {
                candidates.insert(OpcPackage::resolve_rel_target(
                    &slide_part,
                    &relationship.target,
                ));
            }
            false
        });
        staged.package.set_part_rels(&slide_part, relationships);
        prune_unreachable_parts(&mut staged.package, &candidates);
        staged.media_store = MediaStore::scan(&staged.package);
        self.commit_candidate(staged)
    }

    /// Appends an editable relationship-backed chart to one slide.
    #[allow(clippy::too_many_arguments)]
    pub fn add_chart(
        &mut self,
        slide_index: usize,
        kind: ChartKind,
        left: Emu,
        top: Emu,
        width: Emu,
        height: Emu,
        data: &ChartData,
    ) -> Result<ShapeRef<'_>> {
        let slide = self
            .slides
            .get(slide_index)
            .ok_or(Error::UnknownSlideIndex {
                index: slide_index,
                slide_count: self.slides.len(),
            })?;
        validate_chart_extents(width, height)?;
        let slide_part = slide.part_name.clone();
        let id = ShapeIdAllocator::scan(&slide.slide.common_slide_data.shape_tree).allocate();
        let chart_number = next_numbered_part_number(&self.package, "/ppt/charts/chart", ".xml");
        let workbook_number =
            next_numbered_part_number(&self.package, "/ppt/embeddings/Workbook", ".xlsx");
        let chart_part = format!("/ppt/charts/chart{chart_number}.xml");
        let workbook_part = format!("/ppt/embeddings/Workbook{workbook_number}.xlsx");

        let mut package = self.package.clone();
        let slide_relationship_id = package.get_or_create_part_rels(&slide_part).add(
            rel_types::CHART,
            &relative_part_target(&slide_part, &chart_part),
        );
        let workbook_relationship_id = package.get_or_create_part_rels(&chart_part).add(
            rel_types::PACKAGE,
            &relative_part_target(&chart_part, &workbook_part),
        );
        let (chart_xml, workbook_bytes) =
            oxml_chart::authored_chart_parts(kind, data, &workbook_relationship_id)
                .map_err(|error| invalid_chart_mutation("add chart", error.to_string()))?;
        let frame = CT_GraphicFrame::new_chart(
            id,
            &format!("Chart {id}"),
            positioned_transform(left, top, width, height),
            &slide_relationship_id,
        )
        .map_err(|error| invalid_chart_mutation("add chart", error.to_string()))?;
        frame
            .to_xml()
            .map_err(|error| invalid_chart_mutation("add chart", error.to_string()))?;

        package.set_part(&chart_part, chart_xml);
        package.set_part(&workbook_part, workbook_bytes);
        package
            .content_types
            .add_override(&chart_part, content_types::CHART);
        package
            .content_types
            .add_override(&workbook_part, content_types::EMBEDDED_WORKBOOK);

        self.package = package;
        let tree = &mut self.slides[slide_index].slide.common_slide_data.shape_tree;
        Ok(shape_ref(tree.append_child(ShapeTreeChild::GraphicFrame(
            Box::new(frame),
        ))))
    }
}

#[cfg(feature = "render")]
fn render_prepared_timeline_request(
    assembly: &mut PreparedRenderAssembly,
    request: TimelineRequest,
    slide_count: usize,
    reuse_prepared_fonts: bool,
) -> Result<(DeterministicTimelineFrame, Vec<EvaluatedMediaState>)> {
    #[cfg(test)]
    let _retention = ActiveResolvedSample::enter();
    let incoming_source =
        assembly
            .slides
            .get(request.slide_index)
            .ok_or(Error::UnknownSlideIndex {
                index: request.slide_index,
                slide_count,
            })?;
    let mut legacy_font_manager;
    let font_manager = if reuse_prepared_fonts {
        &mut assembly.font_manager
    } else {
        legacy_font_manager = FontManager::new_deterministic()
            .map_err(|error| render_failure(format!("deterministic fonts: {error}")))?;
        legacy_font_manager.load_additional_fonts(&assembly.input.fonts);
        &mut legacy_font_manager
    };
    let mut incoming_context = ResolveCtx::new(
        &incoming_source.theme,
        render_effective_color_map(
            &incoming_source.master,
            &incoming_source.layout,
            &incoming_source.slide,
        ),
        &incoming_source.master,
        &incoming_source.layout,
        &incoming_source.slide,
        &assembly.default_text_style,
    )
    .with_media_failures(&incoming_source.media_failures);
    if let Some(styles) = assembly.table_styles.as_ref() {
        incoming_context = incoming_context.with_table_styles(styles);
    }
    if request.fallback_policy.is_some() {
        incoming_context = incoming_context.with_media_poster_diagnostic_identity();
    }
    let (mut incoming, incoming_directions) = incoming_context
        .resolve_slide_with_chart_resources_and_text_directions_at(
            assembly.size,
            &incoming_source.media,
            &incoming_source.hyperlinks,
            &incoming_source.charts,
            font_manager,
            request.position,
        )
        .map_err(|error| render_failure(format!("{}: {error}", incoming_source.slide_part)))?;
    let mut incoming_media = Vec::new();
    let mut incoming_media_diagnostics = Vec::new();
    if let Some(policy) = request.fallback_policy {
        apply_media_fallback_policy(&mut incoming, &incoming_source.slide, policy, font_manager)?;
        let (states, diagnostics) = evaluate_slide_media(&incoming_source.slide, request.position);
        let consumed_shape_ids = states
            .iter()
            .map(|state| state.shape_id)
            .collect::<HashSet<_>>();
        incoming_media = states;
        incoming_media_diagnostics = incoming_source.media_diagnostics.clone();
        incoming_media_diagnostics.extend(diagnostics);
        suppress_consumed_media_timing_diagnostics(
            &mut incoming.state.diagnostics,
            incoming_source.slide.timing.as_ref(),
            &consumed_shape_ids,
        );
    }
    let (outgoing, outgoing_directions) = match request.outgoing_slide_index {
        None => (None, None),
        Some(index) if index == request.slide_index => {
            (Some(incoming.clone()), Some(incoming_directions.clone()))
        }
        Some(index) => {
            let source = assembly
                .slides
                .get(index)
                .ok_or(Error::UnknownSlideIndex { index, slide_count })?;
            let mut context = ResolveCtx::new(
                &source.theme,
                render_effective_color_map(&source.master, &source.layout, &source.slide),
                &source.master,
                &source.layout,
                &source.slide,
                &assembly.default_text_style,
            )
            .with_media_failures(&source.media_failures);
            if let Some(styles) = assembly.table_styles.as_ref() {
                context = context.with_table_styles(styles);
            }
            if request.fallback_policy.is_some() {
                context = context.with_media_poster_diagnostic_identity();
            }
            let (mut timeline, directions) = context
                .resolve_slide_with_chart_resources_and_text_directions_at(
                    assembly.size,
                    &source.media,
                    &source.hyperlinks,
                    &source.charts,
                    font_manager,
                    TimelinePosition {
                        elapsed_ms: u64::MAX,
                        click_count: u32::MAX,
                    },
                )
                .map_err(|error| render_failure(format!("{}: {error}", source.slide_part)))?;
            if let Some(policy) = request.fallback_policy {
                apply_media_fallback_policy(&mut timeline, &source.slide, policy, font_manager)?;
            }
            (Some(timeline), Some(directions))
        }
    };
    let mut incoming_page = layout_timeline_slide_with_font_manager(
        &assembly.input,
        request.slide_index,
        &incoming,
        Some(&incoming_directions),
        font_manager,
    )
    .map_err(|error| render_failure(error.to_string()))?;
    apply_smartart_clips(
        &mut incoming_page,
        &incoming.slide,
        &incoming_source.smartart_clips,
    )?;
    let outgoing_page = match (request.outgoing_slide_index, outgoing.as_ref()) {
        (Some(index), Some(outgoing)) => {
            let mut page = layout_timeline_slide_with_font_manager(
                &assembly.input,
                index,
                outgoing,
                outgoing_directions.as_ref(),
                font_manager,
            )
            .map_err(|error| render_failure(error.to_string()))?;
            let source = assembly
                .slides
                .get(index)
                .ok_or(Error::UnknownSlideIndex { index, slide_count })?;
            apply_smartart_clips(&mut page, &outgoing.slide, &source.smartart_clips)?;
            Some(page)
        }
        _ => None,
    };
    let outgoing_page = outgoing_page.as_ref();
    let transition = incoming.state.transition.as_ref();
    let composed = if transition.is_some_and(|transition| {
        matches!(
            transition.effect,
            rpptx_oxml::timing::TransitionEffect::Morph
        )
    }) {
        match (outgoing_page, outgoing.as_ref(), transition) {
            (Some(page), Some(outgoing), Some(transition)) => {
                compose_morph_transition(incoming_page, &incoming, page, outgoing, transition)
            }
            _ => compose_transition(incoming_page, outgoing_page, transition),
        }
    } else {
        compose_transition(incoming_page, outgoing_page, transition)
    };
    let mut diagnostics = incoming.slide.diagnostics.clone();
    diagnostics.extend(incoming.state.diagnostics.clone());
    if let Some(outgoing) = outgoing.as_ref() {
        diagnostics.extend(outgoing.slide.diagnostics.clone());
        diagnostics.extend(outgoing.state.diagnostics.clone());
    }
    diagnostics.extend(incoming_media_diagnostics);
    diagnostics.extend(composed.diagnostics);
    Ok((
        DeterministicTimelineFrame {
            page: composed.page,
            state: incoming.state.clone(),
            diagnostics,
        },
        incoming_media,
    ))
}

fn validate_chart_extents(width: Emu, height: Emu) -> Result<()> {
    if width.0 <= 0 || height.0 <= 0 {
        return Err(invalid_chart_mutation(
            "add chart",
            "chart width and height must be positive",
        ));
    }
    Ok(())
}

fn next_numbered_part_number(package: &OpcPackage, prefix: &str, suffix: &str) -> usize {
    let occupied = package
        .parts
        .keys()
        .filter_map(|part| numbered_part_suffix(part, prefix, suffix))
        .collect::<HashSet<_>>();
    let mut candidate = occupied
        .iter()
        .copied()
        .max()
        .and_then(|value| value.checked_add(1))
        .unwrap_or(1);
    while occupied.contains(&candidate) {
        candidate = candidate.checked_add(1).unwrap_or(1);
    }
    candidate
}

fn numbered_part_suffix(part: &str, prefix: &str, suffix: &str) -> Option<usize> {
    let value = part.strip_prefix(prefix)?.strip_suffix(suffix)?;
    let parsed = value.parse::<usize>().ok()?;
    (parsed != 0).then_some(parsed)
}

fn invalid_chart_mutation(operation: &'static str, message: impl Into<String>) -> Error {
    Error::InvalidChartMutation {
        operation,
        message: message.into(),
    }
}

fn invalid_media_mutation(operation: &'static str, message: impl Into<String>) -> Error {
    Error::InvalidMediaMutation {
        operation,
        message: message.into(),
    }
}

fn media_kind_name(kind: MediaKind) -> &'static str {
    match kind {
        MediaKind::Audio => "Audio",
        MediaKind::Video => "Video",
    }
}

fn media_relationship_type(kind: MediaKind) -> &'static str {
    match kind {
        MediaKind::Audio => rel_types::AUDIO,
        MediaKind::Video => rel_types::VIDEO,
    }
}

fn media_picture(slide: &CT_Slide, shape_id: u32) -> Result<&PictureMedia> {
    media_picture_shape(slide, shape_id)?
        .media
        .as_ref()
        .ok_or_else(|| {
            invalid_media_mutation(
                "inspect media",
                format!("shape id {shape_id} is not a media picture"),
            )
        })
}

fn media_picture_shape(slide: &CT_Slide, shape_id: u32) -> Result<&CT_Picture> {
    slide
        .common_slide_data
        .shape_tree
        .children
        .iter()
        .find(|child| child.non_visual_id() == Some(shape_id))
        .and_then(|child| match child {
            ShapeTreeChild::Picture(picture) if picture.media.is_some() => Some(picture),
            _ => None,
        })
        .ok_or_else(|| {
            invalid_media_mutation(
                "inspect media",
                format!("shape id {shape_id} is not a media picture"),
            )
        })
}

fn invalid_shape_mutation(operation: &'static str, message: impl Into<String>) -> Error {
    Error::InvalidShapeMutation {
        operation,
        message: message.into(),
    }
}

fn invalid_slide_mutation(operation: &'static str, message: impl Into<String>) -> Error {
    Error::InvalidSlideMutation {
        operation,
        message: message.into(),
    }
}

fn not_a_picture(operation: &'static str, shape_id: u32) -> Error {
    invalid_shape_mutation(operation, format!("shape id {shape_id} is not a picture"))
}

fn find_picture(children: &[ShapeTreeChild], shape_id: u32) -> Option<&CT_Picture> {
    children.iter().find_map(|child| match child {
        ShapeTreeChild::Picture(picture) if child.non_visual_id() == Some(shape_id) => {
            Some(picture)
        }
        ShapeTreeChild::GroupShape(group) => find_picture(&group.children, shape_id),
        ShapeTreeChild::AlternateContent(alternate) => {
            find_picture(alternate.selected_fallback().unwrap_or_default(), shape_id)
        }
        _ => None,
    })
}

fn find_picture_mut(children: &mut [ShapeTreeChild], shape_id: u32) -> Option<&mut CT_Picture> {
    for child in children {
        let id = child.non_visual_id();
        match child {
            ShapeTreeChild::Picture(picture) if id == Some(shape_id) => return Some(picture),
            ShapeTreeChild::GroupShape(group) => {
                if let Some(picture) = find_picture_mut(&mut group.children, shape_id) {
                    return Some(picture);
                }
            }
            _ => {}
        }
    }
    None
}

/// Counts the slide children and group members whose `p:cNvPr/@id` is
/// `shape_id`.
struct ShapeHyperlinkLocation {
    root_start: usize,
    root_head_end: usize,
    root_name: String,
    empty_root: bool,
    click: Option<(usize, usize, usize, bool)>,
    relationship_id: Option<String>,
    relationship_attribute: Option<Vec<u8>>,
    action: Option<String>,
}

/// The action verb of a click action that jumps to one named slide.
const SLIDE_JUMP_ACTION: &str = "ppaction://hlinksldjump";

/// The action verb of a click action that jumps within the slide show, such
/// as to the next slide.
const SHOW_JUMP_ACTION: &str = "ppaction://hlinkshowjump";

/// What a click action set on a slide shape opens.
#[derive(Clone, Copy)]
enum ShapeClickTarget<'a> {
    /// An external hyperlink address.
    Address(&'a str),
    /// The zero-based index of a slide of the presentation.
    Slide(usize),
}

fn xml_local_name(name: &[u8]) -> &[u8] {
    name.rsplit(|byte| *byte == b':').next().unwrap_or(name)
}

fn xml_is(namespace: &XmlResolveResult<'_>, expected: &str) -> bool {
    matches!(namespace, XmlResolveResult::Bound(XmlNamespace(uri)) if *uri == expected.as_bytes())
}

fn shape_xml_attribute(
    reader: &NsReader<&[u8]>,
    start: &XmlBytesStart<'_>,
    expected_namespace: Option<&str>,
    name: &[u8],
) -> Result<Option<(String, Vec<u8>)>> {
    for attribute in start.attributes().with_checks(false) {
        let attribute = attribute
            .map_err(|error| invalid_shape_mutation("shape hyperlink", error.to_string()))?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let matches_namespace = match expected_namespace {
            Some(expected) => xml_is(&namespace, expected),
            None => matches!(namespace, XmlResolveResult::Unbound),
        };
        if matches_namespace && local.as_ref() == name {
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, start.decoder())
                .map_err(|error| invalid_shape_mutation("shape hyperlink", error.to_string()))?;
            return Ok(Some((value.into_owned(), attribute.key.as_ref().to_vec())));
        }
    }
    Ok(None)
}

fn locate_shape_hyperlink(xml: &[u8], shape_id: u32) -> Result<Option<ShapeHyperlinkLocation>> {
    let mut reader = NsReader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut target_depth = None;
    let mut click_depth = None;
    let mut location: Option<ShapeHyperlinkLocation> = None;
    loop {
        let start_offset = reader.buffer_position() as usize;
        let (is_p, is_a, event) = {
            let (namespace, event) = reader
                .read_resolved_event_into(&mut buffer)
                .map_err(|error| invalid_shape_mutation("shape hyperlink", error.to_string()))?;
            (xml_is(&namespace, P_NS), xml_is(&namespace, A_NS), event)
        };
        let end_offset = reader.buffer_position() as usize;
        match event {
            XmlEvent::Start(start) | XmlEvent::Empty(start) => {
                let empty = xml[start_offset..end_offset].ends_with(b"/>");
                if target_depth.is_none()
                    && is_p
                    && xml_local_name(start.name().as_ref()) == b"cNvPr"
                    && shape_xml_attribute(&reader, &start, None, b"id")?
                        .and_then(|(id, _)| id.parse::<u32>().ok())
                        == Some(shape_id)
                {
                    let root_name = std::str::from_utf8(start.name().as_ref())
                        .map_err(|error| {
                            invalid_shape_mutation("shape hyperlink", error.to_string())
                        })?
                        .to_owned();
                    location = Some(ShapeHyperlinkLocation {
                        root_start: start_offset,
                        root_head_end: end_offset,
                        root_name,
                        empty_root: empty,
                        click: None,
                        relationship_id: None,
                        relationship_attribute: None,
                        action: None,
                    });
                    if empty {
                        return Ok(location);
                    }
                    target_depth = Some(depth + 1);
                } else if target_depth == Some(depth)
                    && is_a
                    && xml_local_name(start.name().as_ref()) == b"hlinkClick"
                {
                    let found = location.as_mut().expect("target depth has location");
                    if found.click.is_some() {
                        return Err(invalid_shape_mutation(
                            "shape hyperlink",
                            "duplicate a:hlinkClick",
                        ));
                    }
                    if let Some((id, attribute)) =
                        shape_xml_attribute(&reader, &start, Some(R_NS), b"id")?
                    {
                        found.relationship_id = Some(id);
                        found.relationship_attribute = Some(attribute);
                    }
                    found.action = shape_xml_attribute(&reader, &start, None, b"action")?
                        .map(|(action, _)| action);
                    found.click = Some((start_offset, end_offset, end_offset, empty));
                    if !empty {
                        click_depth = Some(depth + 1);
                    }
                }
                if !empty {
                    depth += 1;
                }
            }
            XmlEvent::End(_) => {
                if click_depth == Some(depth) {
                    let found = location.as_mut().expect("click depth has location");
                    if let Some(click) = found.click.as_mut() {
                        click.2 = end_offset;
                    }
                    click_depth = None;
                }
                if target_depth == Some(depth) {
                    return Ok(location);
                }
                depth = depth.saturating_sub(1);
            }
            XmlEvent::Eof => return Ok(None),
            _ => {}
        }
        buffer.clear();
    }
}

/// Rewrites the located `a:hlinkClick` to name `link`, a relationship id
/// and an optional action, or removes it for `None`.
///
/// An existing element keeps its other attributes and children, and loses
/// its action when `link` has none.
fn rewrite_shape_hyperlink(
    xml: &[u8],
    location: &ShapeHyperlinkLocation,
    link: Option<(&str, Option<&str>)>,
) -> Result<Vec<u8>> {
    let Some((id, action)) = link else {
        let Some((start, _, end, _)) = location.click else {
            return Ok(xml.to_vec());
        };
        let mut rewritten = Vec::with_capacity(xml.len() - (end - start));
        rewritten.extend_from_slice(&xml[..start]);
        rewritten.extend_from_slice(&xml[end..]);
        return Ok(rewritten);
    };
    if let Some((start, head_end, _, empty)) = location.click {
        let mut reader = quick_xml::Reader::from_reader(&xml[start..head_end]);
        let mut buffer = Vec::new();
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| invalid_shape_mutation("shape hyperlink", error.to_string()))?;
        let original = match event {
            XmlEvent::Start(start) | XmlEvent::Empty(start) => start,
            _ => {
                return Err(invalid_shape_mutation(
                    "shape hyperlink",
                    "malformed a:hlinkClick",
                ));
            }
        };
        let name = std::str::from_utf8(original.name().as_ref())
            .map_err(|error| invalid_shape_mutation("shape hyperlink", error.to_string()))?
            .to_owned();
        let mut replacement = XmlBytesStart::new(name);
        let mut replaced = false;
        let mut action_written = false;
        for attribute in original.attributes().with_checks(false) {
            let attribute = attribute
                .map_err(|error| invalid_shape_mutation("shape hyperlink", error.to_string()))?;
            if location.relationship_attribute.as_deref() == Some(attribute.key.as_ref()) {
                let key = std::str::from_utf8(attribute.key.as_ref()).map_err(|error| {
                    invalid_shape_mutation("shape hyperlink", error.to_string())
                })?;
                replacement.push_attribute((key, id));
                replaced = true;
            } else if attribute.key.as_ref() == b"action" {
                if let Some(action) = action {
                    replacement.push_attribute(("action", action));
                    action_written = true;
                }
            } else {
                replacement.push_attribute(attribute);
            }
        }
        if !replaced {
            replacement.push_attribute(("xmlns:r", R_NS));
            replacement.push_attribute(("r:id", id));
        }
        if let Some(action) = action.filter(|_| !action_written) {
            replacement.push_attribute(("action", action));
        }
        let mut writer = XmlWriter::new(Vec::new());
        writer
            .write_event(if empty {
                XmlEvent::Empty(replacement)
            } else {
                XmlEvent::Start(replacement)
            })
            .map_err(|error| invalid_shape_mutation("shape hyperlink", error.to_string()))?;
        let head = writer.into_inner();
        let mut rewritten = Vec::with_capacity(xml.len() + head.len());
        rewritten.extend_from_slice(&xml[..start]);
        rewritten.extend_from_slice(&head);
        rewritten.extend_from_slice(&xml[head_end..]);
        return Ok(rewritten);
    }
    let action = action
        .map(|action| format!(r#" action="{action}""#))
        .unwrap_or_default();
    let click = format!(r#"<a:hlinkClick xmlns:a="{A_NS}" xmlns:r="{R_NS}" r:id="{id}"{action}/>"#);
    let mut rewritten = Vec::with_capacity(xml.len() + click.len() + location.root_name.len() + 3);
    if location.empty_root {
        let head = &xml[location.root_start..location.root_head_end];
        let open = head
            .strip_suffix(b"/>")
            .ok_or_else(|| invalid_shape_mutation("shape hyperlink", "malformed empty p:cNvPr"))?;
        rewritten.extend_from_slice(&xml[..location.root_start]);
        rewritten.extend_from_slice(open);
        rewritten.extend_from_slice(b">");
        rewritten.extend_from_slice(click.as_bytes());
        rewritten.extend_from_slice(format!("</{}>", location.root_name).as_bytes());
        rewritten.extend_from_slice(&xml[location.root_head_end..]);
    } else {
        rewritten.extend_from_slice(&xml[..location.root_head_end]);
        rewritten.extend_from_slice(click.as_bytes());
        rewritten.extend_from_slice(&xml[location.root_head_end..]);
    }
    Ok(rewritten)
}

fn shape_id_count(children: &[ShapeTreeChild], shape_id: u32) -> usize {
    children
        .iter()
        .map(|child| {
            let nested = match child {
                ShapeTreeChild::GroupShape(group) => shape_id_count(&group.children, shape_id),
                _ => 0,
            };
            usize::from(child.non_visual_id() == Some(shape_id)) + nested
        })
        .sum()
}

fn find_comment_anchor_shape<'a>(
    children: &'a [ShapeTreeChild],
    shape_id: u32,
    monikers: &mut Vec<(&'static str, u32)>,
) -> Option<&'a ShapeTreeChild> {
    for child in children {
        let id = child.non_visual_id();
        let kind = match child {
            ShapeTreeChild::Shape(_) => Some("spMk"),
            ShapeTreeChild::Picture(_) => Some("picMk"),
            ShapeTreeChild::GraphicFrame(_) => Some("graphicFrameMk"),
            ShapeTreeChild::GroupShape(_) => Some("grpSpMk"),
            ShapeTreeChild::Connector(_) => Some("cxnSpMk"),
            ShapeTreeChild::AlternateContent(_) => None,
        };
        if id == Some(shape_id) {
            monikers.push((kind?, shape_id));
            return Some(child);
        }
        if let ShapeTreeChild::GroupShape(group) = child
            && let Some(group_id) = id
        {
            monikers.push(("grpSpMk", group_id));
            if let Some(found) = find_comment_anchor_shape(&group.children, shape_id, monikers) {
                return Some(found);
            }
            monikers.pop();
        }
    }
    None
}

fn office_art_comment_text(body: &CT_TextBody) -> String {
    body.paragraphs()
        .iter()
        .map(|paragraph| {
            paragraph
                .runs
                .iter()
                .map(|run| match run {
                    TextRun::Run(run) => run.text.value.as_str(),
                    TextRun::Break(_) => "\u{000b}",
                    TextRun::Field(field) => {
                        field.text.as_ref().map_or("", |text| text.value.as_str())
                    }
                })
                .collect::<String>()
        })
        .collect::<Vec<_>>()
        .join("\r")
}

fn comment_text_context_count(children: &[ShapeTreeChild], text: &str) -> usize {
    children
        .iter()
        .map(|child| match child {
            ShapeTreeChild::Shape(shape) => usize::from(
                shape
                    .text_body
                    .as_ref()
                    .is_some_and(|body| office_art_comment_text(body) == text),
            ),
            ShapeTreeChild::GroupShape(group) => comment_text_context_count(&group.children, text),
            _ => 0,
        })
        .sum()
}

fn find_shape(children: &[ShapeTreeChild], shape_id: u32) -> Option<&CT_Shape> {
    children.iter().find_map(|child| match child {
        ShapeTreeChild::Shape(shape) if child.non_visual_id() == Some(shape_id) => Some(shape),
        ShapeTreeChild::GroupShape(group) => find_shape(&group.children, shape_id),
        _ => None,
    })
}

fn find_shape_mut(children: &mut [ShapeTreeChild], shape_id: u32) -> Option<&mut CT_Shape> {
    for child in children {
        let id = child.non_visual_id();
        match child {
            ShapeTreeChild::Shape(shape) if id == Some(shape_id) => return Some(shape),
            ShapeTreeChild::GroupShape(group) => {
                if let Some(shape) = find_shape_mut(&mut group.children, shape_id) {
                    return Some(shape);
                }
            }
            _ => {}
        }
    }
    None
}

fn picture_image_relationship<'a>(
    picture: &'a CT_Picture,
    operation: &'static str,
    shape_id: u32,
) -> Result<&'a str> {
    picture
        .blip_fill
        .as_ref()
        .and_then(|fill| fill.blip.as_ref())
        .and_then(|blip| blip.embed.as_deref())
        .ok_or_else(|| {
            invalid_shape_mutation(
                operation,
                format!("picture {shape_id} has no embedded image"),
            )
        })
}

fn collect_subtree_ids(
    children: &[ShapeTreeChild],
    ids: &mut HashSet<u32>,
    media_ids: &mut Vec<u32>,
) {
    for child in children {
        ids.extend(child.non_visual_id());
        match child {
            ShapeTreeChild::Picture(picture) if picture.media.is_some() => {
                media_ids.extend(child.non_visual_id());
            }
            ShapeTreeChild::GroupShape(group) => {
                collect_subtree_ids(&group.children, ids, media_ids);
            }
            ShapeTreeChild::AlternateContent(alternate) => {
                collect_subtree_ids(
                    alternate.selected_fallback().unwrap_or_default(),
                    ids,
                    media_ids,
                );
            }
            _ => {}
        }
    }
}

fn detach_connectors(children: &mut [ShapeTreeChild], removed_ids: &HashSet<u32>) {
    for child in children {
        match child {
            ShapeTreeChild::Connector(connector) => {
                if connector
                    .start_connection
                    .as_ref()
                    .is_some_and(|connection| removed_ids.contains(&connection.id))
                {
                    connector.start_connection = None;
                }
                if connector
                    .end_connection
                    .as_ref()
                    .is_some_and(|connection| removed_ids.contains(&connection.id))
                {
                    connector.end_connection = None;
                }
            }
            ShapeTreeChild::GroupShape(group) => {
                detach_connectors(&mut group.children, removed_ids)
            }
            _ => {}
        }
    }
}

fn slide_relationship_ids(slide: &CT_Slide) -> Result<HashSet<String>> {
    let xml = slide
        .to_xml()
        .map_err(|error| invalid_media_mutation("inspect media references", error.to_string()))?;
    relationship_ids(&xml)
        .map(|ids| ids.into_iter().collect())
        .map_err(|error| invalid_media_mutation("inspect media references", error.to_string()))
}

fn media_infos_for_slide(
    package: &OpcPackage,
    slide_index: usize,
    slide_part: &str,
    slide: &CT_Slide,
) -> Result<Vec<MediaInfo>> {
    let relationships = package
        .get_part_rels(slide_part)
        .cloned()
        .unwrap_or_default();
    slide
        .common_slide_data
        .shape_tree
        .children
        .iter()
        .filter_map(|child| match child {
            ShapeTreeChild::Picture(picture) => picture
                .media
                .as_ref()
                .zip(child.non_visual_id())
                .map(|(media, shape_id)| (picture, media, shape_id)),
            _ => None,
        })
        .map(|(picture, media, shape_id)| {
            let relationship_id = media.source.relationship_id();
            let relationship = relationships.get_by_id(relationship_id).ok_or_else(|| {
                Error::MissingRelationship {
                    source_part: slide_part.to_owned(),
                    relationship_id: relationship_id.to_owned(),
                }
            })?;
            if relationship.rel_type != rel_types::POWERPOINT_MEDIA
                && relationship.rel_type != media_relationship_type(media.kind)
            {
                return Err(Error::WrongRelationshipType {
                    source_part: slide_part.to_owned(),
                    relationship_id: relationship_id.to_owned(),
                    expected: rel_types::POWERPOINT_MEDIA,
                    actual: relationship.rel_type.clone(),
                });
            }
            let source = if relationship_is_external(relationship) {
                MediaLocation::Linked {
                    target: relationship.target.clone(),
                }
            } else {
                let part_name = OpcPackage::resolve_rel_target(slide_part, &relationship.target);
                required_part(package, &part_name)?;
                let content_type = package
                    .content_types
                    .content_type_for(&part_name)
                    .ok_or_else(|| {
                        invalid_media_mutation(
                            "inspect media",
                            format!("package part {part_name} has no content type"),
                        )
                    })?
                    .to_owned();
                MediaLocation::Embedded {
                    part_name,
                    content_type,
                }
            };
            let (settings, mut diagnostics) =
                media_playback_settings(picture, slide.timing.as_ref(), shape_id);
            let content_type = match &source {
                MediaLocation::Embedded { content_type, .. } => Some(content_type.as_str()),
                MediaLocation::Linked { .. } => None,
            };
            if let Some(content_type) = content_type
                && !is_known_media_content_type(content_type)
            {
                diagnostics.push(MediaDiagnostic::UnsupportedContentType {
                    content_type: content_type.to_owned(),
                });
            }
            if media.poster_relationship_id.is_none() {
                diagnostics.push(MediaDiagnostic::MissingPoster);
            }
            Ok(MediaInfo {
                slide_index,
                shape_id,
                kind: media.kind,
                source,
                poster_relationship_id: media.poster_relationship_id.clone(),
                settings,
                diagnostics,
            })
        })
        .collect()
}

fn media_playback_settings(
    picture: &CT_Picture,
    timing: Option<&CT_Timing>,
    shape_id: u32,
) -> (MediaPlaybackSettings, Vec<MediaDiagnostic>) {
    let mut settings = MediaPlaybackSettings::default();
    (settings.trim_start_ms, settings.trim_end_ms) = picture.media_trim_bounds();
    let mut diagnostics = Vec::new();
    let Some(timing) = timing else {
        diagnostics.push(MediaDiagnostic::MissingPlaybackTiming);
        return (settings, diagnostics);
    };
    if let Some(media) = timing.media_for_shape(shape_id) {
        settings.volume = media.common.volume.unwrap_or(100_000);
        settings.looped = media.common.looped;
        settings.show_when_stopped = matches!(
            media.common.display,
            Some(MediaDisplayPolicy::ShowWhenStopped)
        );
    } else {
        diagnostics.push(MediaDiagnostic::MissingPlaybackTiming);
    }
    if let Some(trigger) = timing.media_playback_trigger_for_shape(shape_id) {
        settings.trigger = trigger;
    }
    if let Some(command) = timing.media_commands_for_shape(shape_id).first()
        && let MediaCommandKind::Other(command) = &command.command
    {
        diagnostics.push(MediaDiagnostic::UnsupportedPlaybackCommand {
            command: command.clone(),
        });
    }
    (settings, diagnostics)
}

fn validate_media_settings(settings: MediaPlaybackSettings) -> Result<()> {
    if settings.volume > 100_000 {
        return Err(invalid_media_mutation(
            "configure media",
            "volume must be between 0 and 100000",
        ));
    }
    Ok(())
}

fn validate_media_source(source: MediaSourceInput<'_>) -> Result<()> {
    let (bytes, filename, target, content_type) = match source {
        MediaSourceInput::Embedded(input) => (
            Some(input.bytes),
            Some(input.filename),
            None,
            input.content_type,
        ),
        MediaSourceInput::Linked {
            target,
            content_type,
        } => (None, None, Some(target), content_type),
    };
    if !is_safe_content_type(content_type) {
        return Err(invalid_media_mutation(
            "configure media",
            "content type must be an explicit MIME value",
        ));
    }
    if let Some(target) = target
        && (target.is_empty() || target.bytes().any(|byte| byte.is_ascii_control()))
    {
        return Err(invalid_media_mutation(
            "configure media",
            "linked target must be non-empty and contain no control characters",
        ));
    }
    if let (Some(bytes), Some(filename)) = (bytes, filename) {
        safe_media_extension(filename)?;
        if audio_video_signature_matches(bytes, content_type) == Some(false) {
            return Err(invalid_media_mutation(
                "configure media",
                format!("{filename} bytes do not match {content_type}"),
            ));
        }
    }
    Ok(())
}

fn safe_media_extension(filename: &str) -> Result<String> {
    let extension = filename
        .rsplit_once('.')
        .map(|(_, extension)| extension)
        .filter(|extension| {
            !extension.is_empty()
                && extension.len() <= 16
                && extension.bytes().all(|byte| byte.is_ascii_alphanumeric())
        })
        .ok_or_else(|| {
            invalid_media_mutation(
                "configure media",
                format!("{filename} requires a safe filename extension"),
            )
        })?;
    Ok(extension.to_ascii_lowercase())
}

fn is_known_media_content_type(content_type: &str) -> bool {
    audio_video_signature_matches(&[], content_type).is_some()
}

fn add_media_relationships(
    package: &mut OpcPackage,
    media_store: &mut MediaStore,
    slide_part: &str,
    relationships: &mut Relationships,
    kind: MediaKind,
    source: MediaSourceInput<'_>,
    include_office_media: bool,
) -> Result<(String, Option<String>, PictureMediaSource)> {
    match source {
        MediaSourceInput::Embedded(input) => {
            let extension = safe_media_extension(input.filename)?;
            let media_part =
                media_store.insert_explicit(package, input.bytes, &extension, input.content_type);
            let target = relative_part_target(slide_part, &media_part);
            let standard = relationships.add(media_relationship_type(kind), &target);
            let microsoft = include_office_media
                .then(|| relationships.add(rel_types::POWERPOINT_MEDIA, &target));
            Ok((
                standard.clone(),
                microsoft,
                PictureMediaSource::Embedded {
                    relationship_id: standard,
                },
            ))
        }
        MediaSourceInput::Linked { target, .. } => {
            let standard = relationships.add_external(media_relationship_type(kind), target);
            let microsoft = include_office_media
                .then(|| relationships.add_external(rel_types::POWERPOINT_MEDIA, target));
            Ok((
                standard.clone(),
                microsoft,
                PictureMediaSource::Linked {
                    relationship_id: standard,
                },
            ))
        }
    }
}

struct RelationshipScopeTargets<'a> {
    slide: Option<&'a str>,
    notes: Option<&'a str>,
    layout: Option<&'a str>,
    notes_master: Option<&'a str>,
}

fn preflight_cross_presentation_transfer(
    source_package: &OpcPackage,
    source: &SlideRecord,
) -> Result<()> {
    if source
        .slide
        .contains_placeholder()
        .map_err(|error| Error::MalformedPart {
            part_name: source.part_name.clone(),
            message: error.to_string(),
        })?
    {
        return Err(invalid_presentation_mutation(
            "transfer SmartArt slide",
            "placeholders require cross-presentation index reconciliation and are outside the bounded SmartArt transfer".to_owned(),
        ));
    }
    if source.comments.is_some() {
        return Err(invalid_presentation_mutation(
            "transfer SmartArt slide",
            "modern comments are outside the bounded cross-presentation transfer graph".to_owned(),
        ));
    }
    if source.notes.is_some() {
        return Err(invalid_presentation_mutation(
            "transfer SmartArt slide",
            "notes are outside the bounded cross-presentation transfer graph".to_owned(),
        ));
    }
    let relationships = source_package
        .get_part_rels(&source.part_name)
        .cloned()
        .unwrap_or_default();
    let external_layout_count = relationships
        .items
        .iter()
        .filter(|relationship| {
            relationship.rel_type == rel_types::SLIDE_LAYOUT
                && relationship_is_external(relationship)
        })
        .count();
    if external_layout_count != 0 {
        return Err(invalid_presentation_mutation(
            "transfer SmartArt slide",
            "an external slide-layout relationship cannot be replaced by the selected destination layout".to_owned(),
        ));
    }
    let internal_layout_count = relationships
        .items
        .iter()
        .filter(|relationship| {
            relationship.rel_type == rel_types::SLIDE_LAYOUT
                && !relationship_is_external(relationship)
        })
        .count();
    if internal_layout_count != 1 {
        return Err(invalid_presentation_mutation(
            "transfer SmartArt slide",
            format!(
                "source slide must own exactly one internal slide-layout relationship, found {internal_layout_count}"
            ),
        ));
    }
    let mut smartart_frames = Vec::new();
    collect_smartart_frames(
        &source.slide.common_slide_data.shape_tree.children,
        &mut smartart_frames,
    );
    for (frame, _) in smartart_frames {
        if let GraphicDataPayload::SmartArt(relationship_ids) = frame.graphic_data.payload() {
            validate_smartart_relationship_roles(
                source_package,
                &source.part_name,
                relationship_ids,
            )?;
        }
    }
    let mut visited_diagram_parts = HashSet::new();
    for relationship in &relationships.items {
        if relationship_is_external(relationship) {
            continue;
        }
        if relationship.rel_type != rel_types::SLIDE_LAYOUT
            && relationship.rel_type != rel_types::IMAGE
            && !is_diagram_relationship(&relationship.rel_type)
        {
            return Err(invalid_presentation_mutation(
                "transfer SmartArt slide",
                format!(
                    "internal relationship type {} is outside the bounded cross-presentation transfer graph",
                    relationship.rel_type
                ),
            ));
        }
        let resolved = OpcPackage::resolve_rel_target(&source.part_name, &relationship.target);
        required_part(source_package, &resolved)?;
        if relationship.rel_type == rel_types::IMAGE {
            preflight_internal_transfer_image(
                source_package,
                &resolved,
                "transfer SmartArt slide",
            )?;
        } else if is_diagram_relationship(&relationship.rel_type) {
            preflight_diagram_part_graph(
                source_package,
                &resolved,
                &mut visited_diagram_parts,
                "transfer SmartArt slide",
            )?;
        }
    }
    Ok(())
}

fn preflight_slide_import(source_package: &OpcPackage, source: &SlideRecord) -> Result<()> {
    const OPERATION: &str = "import slide";
    if source.comments.is_some() {
        return Err(invalid_presentation_mutation(
            OPERATION,
            "the slide has comments, which are not carried".to_owned(),
        ));
    }
    let mut smartart_frames = Vec::new();
    collect_smartart_frames(
        &source.slide.common_slide_data.shape_tree.children,
        &mut smartart_frames,
    );
    for (frame, _) in smartart_frames {
        if let GraphicDataPayload::SmartArt(relationship_ids) = frame.graphic_data.payload() {
            validate_smartart_relationship_roles(
                source_package,
                &source.part_name,
                relationship_ids,
            )?;
        }
    }
    preflight_import_scope(source_package, &source.part_name, None)?;
    if let Some(notes) = &source.notes {
        preflight_import_scope(source_package, &notes.part_name, Some(&source.part_name))?;
    }
    Ok(())
}

/// Accepts the internal relationships a slide import rebinds or copies.
///
/// A slide scope, whose `notes_of` is `None`, may relate to its layout, its
/// notes, images, audio, video and SmartArt graphs. A notes scope may relate to
/// its notes master, the slide named by `notes_of`, images, audio and video. The
/// copy rebinds every notes slide relationship to the imported slide, so one
/// that targets another slide is a jump and is refused. External relationships
/// are always carried.
fn preflight_import_scope(
    source_package: &OpcPackage,
    part: &str,
    notes_of: Option<&str>,
) -> Result<()> {
    let slide = notes_of.is_none();
    const OPERATION: &str = "import slide";
    let relationships = source_package
        .get_part_rels(part)
        .cloned()
        .unwrap_or_default();
    let mut visited_diagram_parts = HashSet::new();
    for relationship in &relationships.items {
        if relationship_is_external(relationship) {
            continue;
        }
        let resolved = OpcPackage::resolve_rel_target(part, &relationship.target);
        let rel_type = relationship.rel_type.as_str();
        match rel_type {
            rel_types::SLIDE_LAYOUT | rel_types::NOTES_SLIDE if slide => {}
            rel_types::NOTES_MASTER if !slide => {}
            rel_types::SLIDE if notes_of == Some(resolved.as_str()) => {}
            rel_types::IMAGE => {
                preflight_internal_transfer_image(source_package, &resolved, OPERATION)?;
            }
            _ if is_audio_video_relationship(rel_type) => {
                required_part(source_package, &resolved)?;
            }
            _ if slide && is_diagram_relationship(rel_type) => {
                preflight_diagram_part_graph(
                    source_package,
                    &resolved,
                    &mut visited_diagram_parts,
                    OPERATION,
                )?;
            }
            _ => {
                let what = match rel_type {
                    rel_types::CHART => "a chart".to_owned(),
                    rel_types::OLE_OBJECT | rel_types::STRICT_OLE_OBJECT | rel_types::PACKAGE => {
                        "an embedded object".to_owned()
                    }
                    rel_types::COMMENTS | rel_types::POWERPOINT_COMMENTS => "comments".to_owned(),
                    rel_types::SLIDE => {
                        "a hyperlink that jumps to another slide of the source presentation"
                            .to_owned()
                    }
                    _ => format!("an internal relationship of type {rel_type}"),
                };
                return Err(invalid_presentation_mutation(
                    OPERATION,
                    format!("{part} has {what}, which is not carried"),
                ));
            }
        }
    }
    Ok(())
}

fn preflight_internal_transfer_image(
    source_package: &OpcPackage,
    image_part: &str,
    operation: &'static str,
) -> Result<()> {
    required_part(source_package, image_part)?;
    if source_package
        .get_part_rels(image_part)
        .is_some_and(|relationships| !relationships.items.is_empty())
    {
        return Err(invalid_presentation_mutation(
            operation,
            format!(
                "image part {image_part} owns relationships outside the bounded cross-presentation transfer"
            ),
        ));
    }
    Ok(())
}

fn preflight_diagram_part_graph(
    source_package: &OpcPackage,
    source_part: &str,
    visited: &mut HashSet<String>,
    operation: &'static str,
) -> Result<()> {
    if visited.contains(source_part) {
        return Ok(());
    }
    if visited.len() >= MAX_SMARTART_TRANSFER_DIAGRAM_PARTS {
        return Err(invalid_presentation_mutation(
            operation,
            format!(
                "diagram graph exceeds the {MAX_SMARTART_TRANSFER_DIAGRAM_PARTS}-part transfer bound"
            ),
        ));
    }
    visited.insert(source_part.to_owned());
    required_part(source_package, source_part)?;
    let relationships = source_package
        .get_part_rels(source_part)
        .cloned()
        .unwrap_or_default();
    for relationship in &relationships.items {
        if relationship_is_external(relationship) {
            continue;
        }
        let resolved = OpcPackage::resolve_rel_target(source_part, &relationship.target);
        required_part(source_package, &resolved)?;
        if relationship.rel_type == rel_types::IMAGE {
            preflight_internal_transfer_image(source_package, &resolved, operation)?;
        } else if is_diagram_relationship(&relationship.rel_type) {
            preflight_diagram_part_graph(source_package, &resolved, visited, operation)?;
        } else {
            return Err(invalid_presentation_mutation(
                operation,
                format!(
                    "diagram part {source_part} owns internal relationship type {} outside the bounded cross-presentation transfer",
                    relationship.rel_type
                ),
            ));
        }
    }
    Ok(())
}

fn duplicate_relationship_scope(
    source_package: &OpcPackage,
    destination_package: &mut OpcPackage,
    media_store: &mut MediaStore,
    source_part: &str,
    destination_part: &str,
    source_relationships: &Relationships,
    targets: RelationshipScopeTargets<'_>,
) -> Result<(Relationships, HashMap<String, String>)> {
    let mut destination = Relationships::new();
    let mut relationship_map = HashMap::new();
    let mut diagram_copies = HashMap::new();
    for relationship in &source_relationships.items {
        let target = if relationship_is_external(relationship) {
            relationship.target.clone()
        } else {
            let resolved = match relationship.rel_type.as_str() {
                rel_types::SLIDE => targets.slide.map(str::to_owned).unwrap_or_else(|| {
                    OpcPackage::resolve_rel_target(source_part, &relationship.target)
                }),
                rel_types::NOTES_SLIDE => targets.notes.map(str::to_owned).unwrap_or_else(|| {
                    OpcPackage::resolve_rel_target(source_part, &relationship.target)
                }),
                rel_types::SLIDE_LAYOUT => targets.layout.map(str::to_owned).unwrap_or_else(|| {
                    OpcPackage::resolve_rel_target(source_part, &relationship.target)
                }),
                rel_types::NOTES_MASTER => {
                    targets.notes_master.map(str::to_owned).unwrap_or_else(|| {
                        OpcPackage::resolve_rel_target(source_part, &relationship.target)
                    })
                }
                _ => OpcPackage::resolve_rel_target(source_part, &relationship.target),
            };
            let resolved = if relationship.rel_type == rel_types::IMAGE {
                let bytes = required_part(source_package, &resolved)?.to_vec();
                media_store.insert(destination_package, &bytes, &resolved)
            } else if is_audio_video_relationship(&relationship.rel_type) {
                let bytes = required_part(source_package, &resolved)?;
                if destination_package.get_part(&resolved) == Some(bytes) {
                    // A duplicate inside one deck keeps sharing its media
                    // part, wherever that part is stored.
                    resolved
                } else {
                    let format = resolve(bytes, &resolved);
                    let extension = resolved
                        .rsplit_once('.')
                        .map_or_else(|| format.extension(), |(_, extension)| extension);
                    let content_type = source_package
                        .content_types
                        .content_type_for(&resolved)
                        .unwrap_or_else(|| format.content_type());
                    media_store.insert_explicit(destination_package, bytes, extension, content_type)
                }
            } else if is_diagram_relationship(&relationship.rel_type) {
                duplicate_diagram_part_graph(
                    source_package,
                    destination_package,
                    media_store,
                    &resolved,
                    &mut diagram_copies,
                )?
            } else {
                resolved
            };
            relative_part_target(destination_part, &resolved)
        };
        let new_id = if relationship
            .id
            .strip_prefix("rId")
            .is_some_and(|id| !id.is_empty() && id.bytes().all(|byte| byte.is_ascii_digit()))
        {
            destination.add(&relationship.rel_type, &target)
        } else {
            destination.add_with_id(&relationship.id, &relationship.rel_type, &target);
            relationship.id.clone()
        };
        if let Some(inserted) = destination.items.last_mut() {
            inserted.target_mode = relationship.target_mode.clone();
        }
        relationship_map.insert(relationship.id.clone(), new_id);
    }
    for relationship in source_relationships.items.iter().filter(|relationship| {
        relationship.rel_type == rel_types::DIAGRAM_DATA && !relationship_is_external(relationship)
    }) {
        let source_data = OpcPackage::resolve_rel_target(source_part, &relationship.target);
        let Some(copied_data) = diagram_copies.get(&source_data) else {
            continue;
        };
        let mut data = CT_DiagramData::from_xml(required_part(destination_package, copied_data)?)
            .map_err(|error| Error::MalformedPart {
            part_name: copied_data.clone(),
            message: error.to_string(),
        })?;
        data.remap_drawing_relationship(&relationship_map)
            .map_err(|error| Error::MalformedPart {
                part_name: copied_data.clone(),
                message: error.to_string(),
            })?;
        destination_package.set_part(
            copied_data,
            data.to_xml().map_err(|error| Error::MalformedPart {
                part_name: copied_data.clone(),
                message: error.to_string(),
            })?,
        );
    }
    Ok((destination, relationship_map))
}

fn duplicate_diagram_part_graph(
    source_package: &OpcPackage,
    destination_package: &mut OpcPackage,
    media_store: &mut MediaStore,
    source_part: &str,
    copies: &mut HashMap<String, String>,
) -> Result<String> {
    if let Some(existing) = copies.get(source_part) {
        return Ok(existing.clone());
    }
    if copies.len() >= MAX_SMARTART_TRANSFER_DIAGRAM_PARTS {
        return Err(invalid_presentation_mutation(
            "copy SmartArt diagram graph",
            format!(
                "diagram graph exceeds the {MAX_SMARTART_TRANSFER_DIAGRAM_PARTS}-part copy bound"
            ),
        ));
    }
    let source_xml = required_part(source_package, source_part)?.to_vec();
    let destination_part = fresh_related_part_name(destination_package, source_part);
    copies.insert(source_part.to_owned(), destination_part.clone());
    destination_package.set_part(&destination_part, source_xml.clone());
    if let Some(content_type) = source_package
        .content_types
        .content_type_for(source_part)
        .map(str::to_owned)
    {
        destination_package
            .content_types
            .add_override(&destination_part, &content_type);
    }

    let source_relationships = source_package
        .get_part_rels(source_part)
        .cloned()
        .unwrap_or_default();
    let mut destination_relationships = Relationships::new();
    let mut relationship_map = HashMap::new();
    for relationship in &source_relationships.items {
        let target = if relationship_is_external(relationship) {
            relationship.target.clone()
        } else {
            let resolved = OpcPackage::resolve_rel_target(source_part, &relationship.target);
            let copied = if relationship.rel_type == rel_types::IMAGE {
                let bytes = required_part(source_package, &resolved)?.to_vec();
                media_store.insert(destination_package, &bytes, &resolved)
            } else if is_diagram_relationship(&relationship.rel_type) {
                duplicate_diagram_part_graph(
                    source_package,
                    destination_package,
                    media_store,
                    &resolved,
                    copies,
                )?
            } else {
                resolved
            };
            relative_part_target(&destination_part, &copied)
        };
        let id = destination_relationships.add(&relationship.rel_type, &target);
        if let Some(inserted) = destination_relationships.items.last_mut() {
            inserted.target_mode = relationship.target_mode.clone();
        }
        relationship_map.insert(relationship.id.clone(), id);
    }
    let rewritten = rewrite_exact_rel_ids(&source_xml, &relationship_map).map_err(|error| {
        Error::MalformedPart {
            part_name: source_part.to_owned(),
            message: error.to_string(),
        }
    })?;
    destination_package.set_part(&destination_part, rewritten);
    if !destination_relationships.items.is_empty() {
        destination_package
            .part_rels
            .insert(destination_part.clone(), destination_relationships);
    }
    Ok(destination_part)
}

fn fresh_related_part_name(package: &OpcPackage, source_part: &str) -> String {
    let (directory, filename) = source_part.rsplit_once('/').unwrap_or(("", source_part));
    let (stem_with_suffix, extension) = filename.rsplit_once('.').unwrap_or((filename, "xml"));
    let stem = stem_with_suffix.trim_end_matches(|character: char| character.is_ascii_digit());
    MediaNamer::scan(
        &format!("/{}", directory.trim_start_matches('/')),
        if stem.is_empty() { "part" } else { stem },
        package.parts.keys().map(String::as_str),
    )
    .next_part_name(extension)
}

fn is_audio_video_relationship(relationship_type: &str) -> bool {
    matches!(
        relationship_type,
        rel_types::AUDIO | rel_types::VIDEO | rel_types::POWERPOINT_MEDIA
    )
}

fn is_diagram_relationship(relationship_type: &str) -> bool {
    matches!(
        relationship_type,
        rel_types::DIAGRAM_DATA
            | rel_types::DIAGRAM_LAYOUT
            | rel_types::DIAGRAM_QUICK_STYLE
            | rel_types::DIAGRAM_COLORS
            | rel_types::DIAGRAM_DRAWING
    )
}

fn collect_smartart_infos(
    package: &OpcPackage,
    slide_index: usize,
    source_part: &str,
    children: &[ShapeTreeChild],
    infos: &mut Vec<SmartArtInfo>,
) {
    let mut frames = Vec::new();
    collect_smartart_frames(children, &mut frames);
    for (frame, shape_id) in frames {
        let GraphicDataPayload::SmartArt(relationship_ids) = frame.graphic_data.payload() else {
            continue;
        };
        let mut relationships = (**relationship_ids).clone();
        let data = resolve_diagram_part(
            package,
            source_part,
            &relationships.data,
            rel_types::DIAGRAM_DATA,
            CT_DiagramData::from_xml,
        );
        let layout = resolve_diagram_part(
            package,
            source_part,
            &relationships.layout,
            rel_types::DIAGRAM_LAYOUT,
            CT_DiagramLayoutDefinition::from_xml,
        );
        let style = resolve_diagram_part(
            package,
            source_part,
            &relationships.style,
            rel_types::DIAGRAM_QUICK_STYLE,
            CT_DiagramStyleDefinition::from_xml,
        );
        let colors = resolve_diagram_part(
            package,
            source_part,
            &relationships.colors,
            rel_types::DIAGRAM_COLORS,
            CT_DiagramColorsDefinition::from_xml,
        );
        let drawing_id = match &data {
            DiagramPart::Parsed(data) => data.drawing_relationship_id().map(str::to_owned),
            _ => None,
        };
        relationships.drawing = drawing_id.clone();
        let drawing = drawing_id.map(|relationship_id| {
            resolve_diagram_part(
                package,
                source_part,
                &relationship_id,
                rel_types::DIAGRAM_DRAWING,
                CT_DiagramDrawing::from_xml,
            )
        });
        infos.push(SmartArtInfo {
            slide_index,
            shape_id,
            relationships,
            data,
            layout,
            style,
            colors,
            drawing,
        });
    }
}

fn collect_smartart_frames<'a>(
    children: &'a [ShapeTreeChild],
    frames: &mut Vec<(&'a CT_GraphicFrame, u32)>,
) {
    for child in children {
        match child {
            ShapeTreeChild::GraphicFrame(frame)
                if matches!(
                    frame.graphic_data.payload(),
                    GraphicDataPayload::SmartArt(_)
                ) =>
            {
                if let Some(shape_id) = child.non_visual_id() {
                    frames.push((frame, shape_id));
                }
            }
            ShapeTreeChild::GroupShape(group) => {
                collect_smartart_frames(&group.children, frames);
            }
            ShapeTreeChild::AlternateContent(alternate) => {
                if let Some(fallback) = alternate.selected_fallback() {
                    collect_smartart_frames(fallback, frames);
                }
            }
            _ => {}
        }
    }
}

fn find_smartart_relationships(
    children: &[ShapeTreeChild],
    shape_id: u32,
) -> Option<DiagramRelationshipIds> {
    for child in children {
        if child.non_visual_id() == Some(shape_id)
            && let ShapeTreeChild::GraphicFrame(frame) = child
            && let GraphicDataPayload::SmartArt(relationships) = frame.graphic_data.payload()
        {
            return Some((**relationships).clone());
        }
        match child {
            ShapeTreeChild::GroupShape(group) => {
                if let Some(found) = find_smartart_relationships(&group.children, shape_id) {
                    return Some(found);
                }
            }
            ShapeTreeChild::AlternateContent(alternate) => {
                if let Some(found) = alternate
                    .selected_fallback()
                    .and_then(|children| find_smartart_relationships(children, shape_id))
                {
                    return Some(found);
                }
            }
            _ => {}
        }
    }
    None
}

fn validate_smartart_relationship_roles(
    package: &OpcPackage,
    source_part: &str,
    relationship_ids: &DiagramRelationshipIds,
) -> Result<(String, CT_DiagramData)> {
    let data_part = require_internal_related_part(
        package,
        source_part,
        &relationship_ids.data,
        rel_types::DIAGRAM_DATA,
    )?;
    require_internal_related_part(
        package,
        source_part,
        &relationship_ids.layout,
        rel_types::DIAGRAM_LAYOUT,
    )?;
    require_internal_related_part(
        package,
        source_part,
        &relationship_ids.style,
        rel_types::DIAGRAM_QUICK_STYLE,
    )?;
    require_internal_related_part(
        package,
        source_part,
        &relationship_ids.colors,
        rel_types::DIAGRAM_COLORS,
    )?;
    let data = CT_DiagramData::from_xml(required_part(package, &data_part)?).map_err(|error| {
        Error::MalformedPart {
            part_name: data_part.clone(),
            message: error.to_string(),
        }
    })?;
    if let Some(drawing_id) = data.drawing_relationship_id() {
        require_internal_related_part(
            package,
            source_part,
            drawing_id,
            rel_types::DIAGRAM_DRAWING,
        )?;
    }
    Ok((data_part, data))
}

fn require_internal_related_part(
    package: &OpcPackage,
    source_part: &str,
    relationship_id: &str,
    expected_type: &'static str,
) -> Result<String> {
    let relationship = package
        .get_part_rels(source_part)
        .and_then(|items| items.get_by_id(relationship_id))
        .ok_or_else(|| Error::MissingRelationship {
            source_part: source_part.to_owned(),
            relationship_id: relationship_id.to_owned(),
        })?;
    require_relationship_type(source_part, relationship, expected_type)?;
    reject_external(source_part, relationship)?;
    let target = OpcPackage::resolve_rel_target(source_part, &relationship.target);
    required_part(package, &target)?;
    Ok(target)
}

fn resolve_diagram_part<T>(
    package: &OpcPackage,
    source_part: &str,
    relationship_id: &str,
    expected_type: &'static str,
    parse: fn(&[u8]) -> std::result::Result<T, OxmlError>,
) -> DiagramPart<T> {
    let Some(relationship) = package
        .get_part_rels(source_part)
        .and_then(|items| items.get_by_id(relationship_id))
    else {
        return DiagramPart::MissingTarget(relationship_id.to_owned());
    };
    if relationship.rel_type != expected_type {
        return DiagramPart::Invalid(format!(
            "relationship {relationship_id} has type {}, expected {expected_type}",
            relationship.rel_type
        ));
    }
    if relationship_is_external(relationship) {
        return DiagramPart::External(relationship.target.clone());
    }
    let target = OpcPackage::resolve_rel_target(source_part, &relationship.target);
    let Some(xml) = package.get_part(&target) else {
        return DiagramPart::MissingTarget(target);
    };
    match parse(xml) {
        Ok(part) => DiagramPart::Parsed(part),
        Err(error) => DiagramPart::Invalid(format!("{target}: {error}")),
    }
}

fn related_internal_part(
    package: &OpcPackage,
    source_part: &str,
    relationship_type: &str,
) -> Result<Option<String>> {
    let Some(relationship) = package
        .get_part_rels(source_part)
        .and_then(|relationships| relationships.get_by_type(relationship_type))
    else {
        return Ok(None);
    };
    reject_external(source_part, relationship)?;
    Ok(Some(OpcPackage::resolve_rel_target(
        source_part,
        &relationship.target,
    )))
}

fn collect_media_targets(package: &OpcPackage, source_part: &str, targets: &mut HashSet<String>) {
    let Some(relationships) = package.get_part_rels(source_part) else {
        return;
    };
    for relationship in &relationships.items {
        if relationship_is_external(relationship) {
            continue;
        }
        let target = OpcPackage::resolve_rel_target(source_part, &relationship.target);
        if target
            .get(.."/ppt/media/".len())
            .is_some_and(|prefix| prefix.eq_ignore_ascii_case("/ppt/media/"))
        {
            targets.insert(target);
        }
    }
}

/// Removes candidate parts that no relationship reaches any longer.
///
/// A removed part's own internal targets become candidates in turn, so a
/// removed chart also drops its embedded workbook.
fn prune_unreachable_parts(package: &mut OpcPackage, candidates: &HashSet<String>) {
    let mut pending = candidates.iter().cloned().collect::<Vec<_>>();
    pending.sort_unstable();
    while let Some(candidate) = pending.pop() {
        let reachable_from_root = package.package_rels.items.iter().any(|relationship| {
            !relationship_is_external(relationship)
                && OpcPackage::resolve_rel_target("/", &relationship.target)
                    .eq_ignore_ascii_case(&candidate)
        });
        let reachable_from_part = package
            .part_rels
            .iter()
            .any(|(source_part, relationships)| {
                relationships.items.iter().any(|relationship| {
                    !relationship_is_external(relationship)
                        && OpcPackage::resolve_rel_target(source_part, &relationship.target)
                            .eq_ignore_ascii_case(&candidate)
                })
            });
        if reachable_from_root || reachable_from_part {
            continue;
        }
        if let Some(relationships) = package.get_part_rels(&candidate) {
            pending.extend(
                relationships
                    .items
                    .iter()
                    .filter(|relationship| !relationship_is_external(relationship))
                    .map(|relationship| {
                        OpcPackage::resolve_rel_target(&candidate, &relationship.target)
                    }),
            );
        }
        package.remove_part(&candidate);
        package.remove_part_rels(&candidate);
        package.content_types.remove_override(&candidate);
    }
}

/// The left, top, right, and bottom `a:srcRect` insets, in thousandths of a
/// percent, that crop an image of `image` pixels to the aspect ratio of a
/// `frame`, evenly on the image's longer sides, as python-pptx crops a
/// picture placeholder.
#[cfg(feature = "render")]
fn fill_crop(image: (f64, f64), frame: (f64, f64)) -> [i32; 4] {
    let image_ratio = image.0 / image.1;
    let frame_ratio = frame.0 / frame.1;
    let inset = |share: f64| (share / 2.0 * 100_000.0).round() as i32;
    if frame_ratio < image_ratio {
        let side = inset(1.0 - frame_ratio / image_ratio);
        [side, 0, side, 0]
    } else if frame_ratio > image_ratio {
        let side = inset(1.0 - image_ratio / frame_ratio);
        [0, side, 0, side]
    } else {
        [0; 4]
    }
}

fn picture_dimensions(
    image_data: &[u8],
    image_filename: &str,
    width: Option<Emu>,
    height: Option<Emu>,
) -> Result<(Emu, Emu)> {
    if let (Some(width), Some(height)) = (width, height) {
        return Ok((width, height));
    }
    if ImageFormat::sniff(image_data).is_none() {
        return Err(Error::UnsupportedPicture {
            filename: image_filename.to_owned(),
        });
    }
    let info = probe(image_data).ok_or_else(|| Error::UnavailablePictureDimensions {
        filename: image_filename.to_owned(),
    })?;
    match (width, height) {
        (None, None) => info
            .native_size(72.0)
            .map(|size| (Emu(size.width_emu), Emu(size.height_emu)))
            .ok_or_else(|| Error::UnavailablePictureDimensions {
                filename: image_filename.to_owned(),
            }),
        (Some(width), None) => infer_picture_dimension(
            width,
            info.height_px,
            info.width_px,
            image_filename,
            "height",
        )
        .map(|height| (width, height)),
        (None, Some(height)) => infer_picture_dimension(
            height,
            info.width_px,
            info.height_px,
            image_filename,
            "width",
        )
        .map(|width| (width, height)),
        (Some(_), Some(_)) => unreachable!(),
    }
}

fn infer_picture_dimension(
    provided: Emu,
    inferred_pixels: u32,
    provided_pixels: u32,
    image_filename: &str,
    axis: &'static str,
) -> Result<Emu> {
    let inferred =
        i128::from(provided.0) * i128::from(inferred_pixels) / i128::from(provided_pixels);
    i64::try_from(inferred)
        .map(Emu)
        .map_err(|_| Error::PictureDimensionOutOfRange {
            filename: image_filename.to_owned(),
            axis,
        })
}

fn resolve_layouts(
    package: &OpcPackage,
    presentation_part: &str,
    presentation: &CT_Presentation,
) -> Result<Vec<LayoutRecord>> {
    let presentation_relationships = package.get_part_rels(presentation_part);
    let mut layouts = Vec::new();
    let mut seen = HashSet::new();
    for master in &presentation.slide_master_ids {
        let relationship = presentation_relationships
            .and_then(|relationships| relationships.get_by_id(&master.relationship_id))
            .ok_or_else(|| Error::MissingRelationship {
                source_part: presentation_part.to_owned(),
                relationship_id: master.relationship_id.clone(),
            })?;
        require_relationship_type(presentation_part, relationship, rel_types::SLIDE_MASTER)?;
        reject_external(presentation_part, relationship)?;
        let master_part = OpcPackage::resolve_rel_target(presentation_part, &relationship.target);
        required_part(package, &master_part)?;
        let Some(master_relationships) = package.get_part_rels(&master_part) else {
            continue;
        };
        for relationship in master_relationships
            .items
            .iter()
            .filter(|relationship| relationship.rel_type == rel_types::SLIDE_LAYOUT)
        {
            reject_external(&master_part, relationship)?;
            let part_name = OpcPackage::resolve_rel_target(&master_part, &relationship.target);
            if !seen.insert(part_name.clone()) {
                continue;
            }
            let xml = required_part(package, &part_name)?;
            let layout = CT_SlideLayout::from_xml(xml).map_err(|error| Error::MalformedPart {
                part_name: part_name.clone(),
                message: error.to_string(),
            })?;
            layouts.push(LayoutRecord { part_name, layout });
        }
    }
    Ok(layouts)
}

#[allow(clippy::too_many_arguments)]
fn validate_shape_children(
    slide_index: usize,
    children: &[ShapeTreeChild],
    shape_index: &mut usize,
    shape_ids: &mut HashSet<u32>,
    placeholder_indices: &mut HashSet<u32>,
    duplicate_shape_ids: &mut Vec<ValidationIssue>,
    empty_text_bodies: &mut Vec<ValidationIssue>,
    duplicate_placeholder_indices: &mut Vec<ValidationIssue>,
) {
    let mut pending = children.iter().rev().collect::<Vec<_>>();
    while let Some(child) = pending.pop() {
        let current_shape_index = *shape_index;
        *shape_index += 1;
        if let Some(id) = child.non_visual_id()
            && !shape_ids.insert(id)
        {
            duplicate_shape_ids.push(ValidationIssue::DuplicateShapeId {
                slide: slide_index,
                id,
            });
        }
        let placeholder = match child {
            ShapeTreeChild::Shape(shape) => {
                if shape
                    .text_body
                    .as_ref()
                    .is_some_and(|body| body.paragraph_count() == 0)
                {
                    empty_text_bodies.push(ValidationIssue::EmptyTextBody {
                        slide: slide_index,
                        shape: current_shape_index,
                    });
                }
                shape.placeholder.as_ref()
            }
            ShapeTreeChild::Picture(picture) => picture.placeholder.as_ref(),
            _ => None,
        };
        if let Some(idx) = placeholder
            .and_then(|placeholder| placeholder.idx)
            .filter(|idx| *idx != u32::MAX)
            && !placeholder_indices.insert(idx)
        {
            duplicate_placeholder_indices.push(ValidationIssue::DuplicatePlaceholderIdx {
                slide: slide_index,
                idx,
            });
        }
        match child {
            ShapeTreeChild::GroupShape(group) => {
                pending.extend(group.children.iter().rev());
            }
            ShapeTreeChild::AlternateContent(alternate) => {
                if let Some(fallback) = alternate.selected_fallback() {
                    pending.extend(fallback.iter().rev());
                }
            }
            _ => {}
        }
    }
}

fn validate_relationship_scope(
    package: &OpcPackage,
    source_part: &str,
    source_xml: Option<&[u8]>,
    relationships: &Relationships,
    incoming_targets: &mut HashSet<String>,
    dangling_relationships: &mut Vec<ValidationIssue>,
    unreachable_targets: &mut Vec<ValidationIssue>,
) {
    let referenced = source_xml.and_then(|xml| relationship_ids(xml).ok());
    let has_relationship_namespace = source_xml.is_some_and(|xml| {
        xml.windows(rpptx_oxml::namespace::R_NS.len())
            .any(|window| window == rpptx_oxml::namespace::R_NS.as_bytes())
    });
    if let Some(referenced) = &referenced {
        let mut reported = HashSet::new();
        for relationship_id in referenced {
            if relationships.get_by_id(relationship_id).is_none()
                && reported.insert(relationship_id)
            {
                dangling_relationships.push(ValidationIssue::DanglingRelationship {
                    part: source_part.to_owned(),
                    r_id: relationship_id.clone(),
                    target: relationship_id.clone(),
                });
            }
        }
    }
    for relationship in &relationships.items {
        if let Some(referenced) = &referenced
            && has_relationship_namespace
            && relationship_requires_xml_reference(&relationship.rel_type)
            && !referenced.iter().any(|id| id == &relationship.id)
        {
            dangling_relationships.push(ValidationIssue::DanglingRelationship {
                part: source_part.to_owned(),
                r_id: relationship.id.clone(),
                target: relationship.target.clone(),
            });
        }
        if relationship_is_external(relationship) {
            continue;
        }
        let target = OpcPackage::resolve_rel_target(source_part, &relationship.target);
        incoming_targets.insert(target.to_ascii_lowercase());
        if !package.contains_part(&target) {
            unreachable_targets.push(ValidationIssue::UnreachableRelationshipTarget {
                part: source_part.to_owned(),
                target,
            });
        }
    }
}

fn relationship_requires_xml_reference(relationship_type: &str) -> bool {
    matches!(
        relationship_type,
        rel_types::IMAGE | rel_types::HYPERLINK | rel_types::CHART
    ) || relationship_type.ends_with("/audio")
        || relationship_type.ends_with("/video")
        || relationship_type.ends_with("/oleObject")
        || relationship_type.ends_with("/diagramData")
}

fn part_has_content_type(package: &OpcPackage, part_name: &str) -> bool {
    package.content_types.content_type_for(part_name).is_some()
        || part_name.rsplit_once('.').is_some_and(|(_, extension)| {
            package
                .content_types
                .defaults
                .contains_key(&extension.to_ascii_lowercase())
        })
}

fn relationship_is_external(relationship: &Relationship) -> bool {
    relationship
        .target_mode
        .as_deref()
        .is_some_and(|mode| mode.eq_ignore_ascii_case("external"))
}

fn numeric_relationship_id(relationship_id: &str) -> u32 {
    relationship_id
        .strip_prefix("rId")
        .and_then(|value| value.parse().ok())
        .unwrap_or_default()
}

/// Gives placeholders the new layout does not place the transform they inherited.
///
/// Only placeholders without a transform of their own are touched, and
/// only when the old layout chain supplied one.
#[cfg(feature = "render")]
fn keep_unplaced_placeholder_transforms(
    children: &mut [ShapeTreeChild],
    old: (&CT_SlideLayout, &CT_SlideMaster),
    new: (&CT_SlideLayout, &CT_SlideMaster),
) {
    for child in children {
        let (placeholder, transform) = match child {
            ShapeTreeChild::Shape(shape) => (
                shape.placeholder.as_ref(),
                &mut shape.shape_properties.transform,
            ),
            ShapeTreeChild::Picture(picture) => (
                picture.placeholder.as_ref(),
                &mut picture.shape_properties.transform,
            ),
            ShapeTreeChild::GroupShape(group) => {
                keep_unplaced_placeholder_transforms(&mut group.children, old, new);
                continue;
            }
            _ => continue,
        };
        if let Some(placeholder) = placeholder
            && transform.is_none()
            && inherited_xfrm(placeholder, new.0, new.1).is_none()
        {
            *transform = inherited_xfrm(placeholder, old.0, old.1).cloned();
        }
    }
}

/// The name PowerPoint and python-pptx give a placeholder a new slide copies
/// from its layout, such as `Title 1` or `Content Placeholder 2`: the kind's
/// base name, prefixed `Vertical` for vertical text, then the shape id less
/// one.
fn placeholder_name(ph_type: Option<&PhType>, vertical: bool, id: u32) -> String {
    let base = match ph_type.unwrap_or(&PhType::Object) {
        PhType::Title | PhType::CenteredTitle => "Title",
        PhType::Subtitle => "Subtitle",
        PhType::Body => "Text Placeholder",
        PhType::Object => "Content Placeholder",
        PhType::Picture => "Picture Placeholder",
        PhType::Chart => "Chart Placeholder",
        PhType::Table => "Table Placeholder",
        PhType::Diagram => "SmartArt Placeholder",
        PhType::Media => "Media Placeholder",
        PhType::ClipArt => "ClipArt Placeholder",
        PhType::DateTime => "Date Placeholder",
        PhType::Footer => "Footer Placeholder",
        PhType::SlideNumber => "Slide Number Placeholder",
        PhType::Header => "Header Placeholder",
        PhType::SlideImage => "Slide Image Placeholder",
        PhType::VerticalTitle => "Vertical Title",
        PhType::VerticalBody => "Vertical Text Placeholder",
        PhType::VerticalObject => "Vertical Content Placeholder",
        PhType::Other(_) => "Placeholder",
    };
    let vertical = vertical && !base.starts_with("Vertical");
    format!(
        "{}{base} {}",
        if vertical { "Vertical " } else { "" },
        id.saturating_sub(1)
    )
}

fn is_latent_placeholder(placeholder: &CT_Placeholder) -> bool {
    matches!(
        placeholder.ph_type.as_ref(),
        Some(PhType::DateTime | PhType::Footer | PhType::SlideNumber)
    )
}

fn relative_part_target(source_part: &str, target_part: &str) -> String {
    let mut source = source_part
        .trim_start_matches('/')
        .split('/')
        .collect::<Vec<_>>();
    source.pop();
    let target = target_part
        .trim_start_matches('/')
        .split('/')
        .collect::<Vec<_>>();
    let common = source
        .iter()
        .zip(&target)
        .take_while(|(left, right)| left == right)
        .count();
    std::iter::repeat_n("..", source.len() - common)
        .chain(target[common..].iter().copied())
        .collect::<Vec<_>>()
        .join("/")
}

fn required_part<'a>(package: &'a OpcPackage, part_name: &str) -> Result<&'a [u8]> {
    package
        .get_part(part_name)
        .ok_or_else(|| Error::MissingPart {
            part_name: part_name.to_owned(),
        })
}

fn require_relationship_type(
    source_part: &str,
    relationship: &Relationship,
    expected: &'static str,
) -> Result<()> {
    if relationship.rel_type == expected {
        return Ok(());
    }
    Err(Error::WrongRelationshipType {
        source_part: source_part.to_owned(),
        relationship_id: relationship.id.clone(),
        expected,
        actual: relationship.rel_type.clone(),
    })
}

fn reject_external(source_part: &str, relationship: &Relationship) -> Result<()> {
    if relationship
        .target_mode
        .as_deref()
        .is_some_and(|mode| mode.eq_ignore_ascii_case("external"))
    {
        return Err(Error::ExternalRelationship {
            source_part: source_part.to_owned(),
            relationship_id: relationship.id.clone(),
        });
    }
    Ok(())
}

fn resolve_comment_authors(
    package: &OpcPackage,
    presentation_part: &str,
) -> Result<(Option<String>, CommentAuthorList)> {
    let Some(relationships) = package.get_part_rels(presentation_part) else {
        return Ok((None, CommentAuthorList::new()));
    };
    let matches = relationships.get_all_by_type(rel_types::POWERPOINT_AUTHORS);
    if matches.len() > 1 {
        return Err(Error::DuplicateRelationships {
            source_part: presentation_part.to_owned(),
            relationship_type: rel_types::POWERPOINT_AUTHORS,
            count: matches.len(),
        });
    }
    let Some(relationship) = matches.first() else {
        return Ok((None, CommentAuthorList::new()));
    };
    reject_external(presentation_part, relationship)?;
    let part_name = OpcPackage::resolve_rel_target(presentation_part, &relationship.target);
    let authors =
        CommentAuthorList::from_xml(required_part(package, &part_name)?).map_err(|error| {
            Error::MalformedPart {
                part_name: part_name.clone(),
                message: error.to_string(),
            }
        })?;
    Ok((Some(part_name), authors))
}

fn resolve_comments(
    package: &OpcPackage,
    slide_part: &str,
    slide: &CT_Slide,
) -> Result<Option<CommentRecord>> {
    let reference =
        slide
            .modern_comment_relationship_id()
            .map_err(|error| Error::MalformedPart {
                part_name: slide_part.to_owned(),
                message: error.to_string(),
            })?;
    let relationships = package.get_part_rels(slide_part);
    let typed = relationships
        .map(|items| items.get_all_by_type(rel_types::POWERPOINT_COMMENTS))
        .unwrap_or_default();
    if typed.len() > 1 {
        return Err(Error::DuplicateRelationships {
            source_part: slide_part.to_owned(),
            relationship_type: rel_types::POWERPOINT_COMMENTS,
            count: typed.len(),
        });
    }
    let Some(reference) = reference else {
        if let Some(relationship) = typed.first() {
            reject_external(slide_part, relationship)?;
            return Err(Error::MalformedPart {
                part_name: slide_part.to_owned(),
                message: format!(
                    "modern comment relationship {} has no p188:commentRel reference",
                    relationship.id
                ),
            });
        }
        return Ok(None);
    };
    let relationship = relationships
        .and_then(|items| items.get_by_id(&reference))
        .ok_or_else(|| Error::MissingRelationship {
            source_part: slide_part.to_owned(),
            relationship_id: reference.clone(),
        })?;
    require_relationship_type(slide_part, relationship, rel_types::POWERPOINT_COMMENTS)?;
    reject_external(slide_part, relationship)?;
    let part_name = OpcPackage::resolve_rel_target(slide_part, &relationship.target);
    let comments = CommentList::from_xml(required_part(package, &part_name)?).map_err(|error| {
        Error::MalformedPart {
            part_name: part_name.clone(),
            message: error.to_string(),
        }
    })?;
    Ok(Some(CommentRecord {
        part_name,
        comments,
        dirty: false,
    }))
}

fn resolve_notes_master(
    package: &OpcPackage,
    presentation_part: &str,
) -> Result<(Option<String>, Option<Box<CT_NotesMaster>>)> {
    let Some(relationships) = package.get_part_rels(presentation_part) else {
        return Ok((None, None));
    };
    let matches = relationships.get_all_by_type(rel_types::NOTES_MASTER);
    if matches.len() > 1 {
        return Err(Error::DuplicateRelationships {
            source_part: presentation_part.to_owned(),
            relationship_type: rel_types::NOTES_MASTER,
            count: matches.len(),
        });
    }
    let Some(relationship) = matches.first() else {
        return Ok((None, None));
    };
    reject_external(presentation_part, relationship)?;
    let part_name = OpcPackage::resolve_rel_target(presentation_part, &relationship.target);
    let master =
        CT_NotesMaster::from_xml(required_part(package, &part_name)?).map_err(|error| {
            Error::MalformedPart {
                part_name: part_name.clone(),
                message: error.to_string(),
            }
        })?;
    Ok((Some(part_name), Some(Box::new(master))))
}

fn resolve_handout_master(
    package: &OpcPackage,
    presentation_part: &str,
) -> Result<(Option<String>, Option<Box<CT_HandoutMaster>>)> {
    let Some(relationships) = package.get_part_rels(presentation_part) else {
        return Ok((None, None));
    };
    let matches = relationships.get_all_by_type(rel_types::HANDOUT_MASTER);
    if matches.len() > 1 {
        return Err(Error::DuplicateRelationships {
            source_part: presentation_part.to_owned(),
            relationship_type: rel_types::HANDOUT_MASTER,
            count: matches.len(),
        });
    }
    let Some(relationship) = matches.first() else {
        return Ok((None, None));
    };
    reject_external(presentation_part, relationship)?;
    let part_name = OpcPackage::resolve_rel_target(presentation_part, &relationship.target);
    let master =
        CT_HandoutMaster::from_xml(required_part(package, &part_name)?).map_err(|error| {
            Error::MalformedPart {
                part_name: part_name.clone(),
                message: error.to_string(),
            }
        })?;
    Ok((Some(part_name), Some(Box::new(master))))
}

fn invalid_presentation_mutation(operation: &'static str, message: String) -> Error {
    Error::InvalidPresentationMutation { operation, message }
}

fn validate_collaboration_model(
    authors: &CommentAuthorList,
    slides: &[SlideRecord],
) -> std::result::Result<(), String> {
    let mut author_ids = HashSet::new();
    for author in &authors.authors {
        if !author_ids.insert(author.id.as_str()) {
            return Err(format!("duplicate modern comment author id {}", author.id));
        }
    }
    let mut comment_ids = HashSet::new();
    let mut comment_parts = HashSet::new();
    for slide in slides {
        let Some(comments) = &slide.comments else {
            continue;
        };
        if !comment_parts.insert(comments.part_name.as_str()) {
            return Err(format!(
                "modern comment part {} is referenced by more than one slide",
                comments.part_name
            ));
        }
        for comment in &comments.comments.comments {
            if !author_ids.contains(comment.author_id.as_str()) {
                return Err(format!(
                    "modern comment {} refers to unknown author {}",
                    comment.id, comment.author_id
                ));
            }
            if !comment_ids.insert(comment.id.as_str()) {
                return Err(format!("duplicate modern comment id {}", comment.id));
            }
            for reply in comment.replies() {
                if !author_ids.contains(reply.author_id.as_str()) {
                    return Err(format!(
                        "modern comment reply {} refers to unknown author {}",
                        reply.id, reply.author_id
                    ));
                }
                if !comment_ids.insert(reply.id.as_str()) {
                    return Err(format!("duplicate modern comment id {}", reply.id));
                }
            }
        }
    }
    Ok(())
}

fn resolve_notes(package: &OpcPackage, slide_part: &str) -> Result<Option<NotesRecord>> {
    let Some(relationships) = package.get_part_rels(slide_part) else {
        return Ok(None);
    };
    let notes_relationships = relationships.get_all_by_type(rel_types::NOTES_SLIDE);
    if notes_relationships.len() > 1 {
        return Err(Error::DuplicateNotesSlides {
            source_part: slide_part.to_owned(),
            count: notes_relationships.len(),
        });
    }
    let Some(relationship) = notes_relationships.first() else {
        return Ok(None);
    };
    reject_external(slide_part, relationship)?;
    let part_name = OpcPackage::resolve_rel_target(slide_part, &relationship.target);
    let xml = required_part(package, &part_name)?;
    let notes = CT_NotesSlide::from_xml(xml).map_err(|error| Error::MalformedPart {
        part_name: part_name.clone(),
        message: error.to_string(),
    })?;
    Ok(Some(NotesRecord { part_name, notes }))
}

/// A borrowed slide in presentation order.
#[derive(Clone, Copy)]
pub struct SlideRef<'a> {
    record: &'a SlideRecord,
}

fn slide_ref(record: &SlideRecord) -> SlideRef<'_> {
    SlideRef { record }
}

/// A mutably borrowed slide in presentation order.
pub struct SlideMut<'a> {
    record: &'a mut SlideRecord,
}

fn slide_mut(record: &mut SlideRecord) -> SlideMut<'_> {
    SlideMut { record }
}

impl<'a> SlideMut<'a> {
    /// Sets whether the slide is hidden from the slideshow.
    pub fn set_hidden(&mut self, hidden: bool) {
        self.record.slide.set_hidden(hidden);
    }

    /// Sets a direct DrawingML fill without replacing a theme reference.
    pub fn set_background(&mut self, fill: Fill) -> Result<()> {
        self.record
            .slide
            .set_background(fill)
            .map_err(|error| Error::InvalidSlideMutation {
                operation: "set slide background",
                message: error.to_string(),
            })
    }

    /// Removes only a direct-fill background.
    pub fn clear_background(&mut self) {
        self.record.slide.clear_background();
    }

    /// Removes any slide background, direct fill or theme reference, so the
    /// slide follows its layout and master background.
    pub fn remove_background(&mut self) {
        self.record.slide.common_slide_data.background = None;
    }

    /// Replaces speaker-note text while preserving its body placeholder.
    pub fn set_notes_text(&mut self, text: &str) -> Result<()> {
        let notes = self
            .record
            .notes
            .as_mut()
            .ok_or_else(|| Error::InvalidSlideMutation {
                operation: "set notes text",
                message: "slide has no notes part".to_owned(),
            })?;
        let mut candidate = notes.notes.clone();
        let body = first_notes_body_mut(&mut candidate.common_slide_data.shape_tree.children)
            .ok_or_else(|| Error::InvalidSlideMutation {
                operation: "set notes text",
                message: "notes slide has no body placeholder".to_owned(),
            })?;
        body.set_text(text);
        candidate
            .to_xml()
            .map_err(|error| Error::InvalidSlideMutation {
                operation: "set notes text",
                message: error.to_string(),
            })?;
        notes.notes = candidate;
        Ok(())
    }

    /// Appends a rectangular table at the top of the slide's z-order.
    pub fn add_table(
        &mut self,
        rows: usize,
        columns: usize,
        left: Emu,
        top: Emu,
        width: Emu,
        height: Emu,
    ) -> Result<ShapeMut<'_>> {
        append_new_member(self.shape_tree(), &[], |id| {
            table_member(id, rows, columns, left, top, width, height)
        })
    }

    /// Appends a no-fill textbox at the top of the slide's z-order.
    pub fn add_textbox(
        &mut self,
        left: Emu,
        top: Emu,
        width: Emu,
        height: Emu,
    ) -> Result<ShapeMut<'_>> {
        append_new_member(self.shape_tree(), &[], |id| {
            textbox_member(id, left, top, width, height)
        })
    }

    /// Appends an ordinary preset shape at the top of the slide's z-order.
    ///
    /// The shape carries the theme style python-pptx writes, so it shows the
    /// theme's accent fill, line, and effect with light text until direct
    /// properties replace them.
    pub fn add_shape(
        &mut self,
        preset: &str,
        left: Emu,
        top: Emu,
        width: Emu,
        height: Emu,
    ) -> Result<ShapeMut<'_>> {
        append_new_member(self.shape_tree(), &[], |id| {
            preset_member(id, preset, left, top, width, height)
        })
    }

    /// Appends a preset shape without a theme style for HTML and ODP imports.
    pub(crate) fn add_unstyled_shape(
        &mut self,
        preset: &str,
        left: Emu,
        top: Emu,
        width: Emu,
        height: Emu,
    ) -> Result<ShapeMut<'_>> {
        append_new_member(self.shape_tree(), &[], |id| {
            unstyled_preset_member(id, preset, left, top, width, height)
        })
    }

    /// Appends a free-standing connector at the top of the slide's z-order.
    ///
    /// The connector carries the theme style python-pptx writes, so it shows
    /// the theme's line until a direct line replaces it.
    pub fn add_connector(
        &mut self,
        connector: ConnectorType,
        begin_x: Emu,
        begin_y: Emu,
        end_x: Emu,
        end_y: Emu,
    ) -> Result<ShapeMut<'_>> {
        append_new_member(self.shape_tree(), &[], |id| {
            connector_member(id, connector, begin_x, begin_y, end_x, end_y)
        })
    }

    /// Appends an empty group at the top of the slide's z-order.
    pub fn add_group_shape(&mut self) -> Result<ShapeMut<'_>> {
        append_new_member(self.shape_tree(), &[], |id| Ok(group_member(id, false)))
    }

    fn shape_tree(&mut self) -> &mut CT_ShapeTree {
        &mut self.record.slide.common_slide_data.shape_tree
    }

    /// Returns one immediate z-order child as a read-only handle.
    pub fn shape(&self, index: usize) -> Option<ShapeRef<'_>> {
        self.record
            .slide
            .common_slide_data
            .shape_tree
            .children
            .get(index)
            .map(shape_ref)
    }

    /// Returns one immediate z-order child for in-place mutation.
    pub fn shape_mut(&mut self, index: usize) -> Option<ShapeMut<'_>> {
        self.record
            .slide
            .common_slide_data
            .shape_tree
            .children
            .get_mut(index)
            .map(shape_mut)
    }

    /// Consumes this handle and returns one immediate z-order child for mutation.
    pub fn into_shape_mut(self, index: usize) -> Option<ShapeMut<'a>> {
        self.record
            .slide
            .common_slide_data
            .shape_tree
            .children
            .get_mut(index)
            .map(shape_mut)
    }
}

/// The shapes of one slide, or of one group on it, borrowed for adding
/// shapes.
///
/// A new shape takes a `p:cNvPr` id unused across the whole slide, groups
/// included, and lands at the top of its collection's z-order. Coordinates
/// are in the collection's own space, which for a group is its member space.
/// An addition to a group refits that group and every group enclosing it to
/// the union of their members, as python-pptx does.
pub struct ShapesMut<'a> {
    presentation: &'a mut Presentation,
    slide_index: usize,
    group: Vec<usize>,
}

impl ShapesMut<'_> {
    /// Appends a no-fill textbox.
    pub fn add_textbox(
        &mut self,
        left: Emu,
        top: Emu,
        width: Emu,
        height: Emu,
    ) -> Result<ShapeMut<'_>> {
        self.append(|id| textbox_member(id, left, top, width, height))
    }

    /// Appends an ordinary preset shape.
    pub fn add_shape(
        &mut self,
        preset: &str,
        left: Emu,
        top: Emu,
        width: Emu,
        height: Emu,
    ) -> Result<ShapeMut<'_>> {
        self.append(|id| preset_member(id, preset, left, top, width, height))
    }

    /// Appends a free-standing connector.
    pub fn add_connector(
        &mut self,
        connector: ConnectorType,
        begin_x: Emu,
        begin_y: Emu,
        end_x: Emu,
        end_y: Emu,
    ) -> Result<ShapeMut<'_>> {
        self.append(|id| connector_member(id, connector, begin_x, begin_y, end_x, end_y))
    }

    /// Appends an empty group, which the first addition to it fits.
    ///
    /// A group added inside a group gets the zero `a:xfrm` python-pptx
    /// writes, because python-pptx cannot refit a group whose member group
    /// has none. One added to the slide's own shapes has no transform, as
    /// [`SlideMut::add_group_shape`] writes it.
    pub fn add_group_shape(&mut self) -> Result<ShapeMut<'_>> {
        let nested = !self.group.is_empty();
        self.append(|id| Ok(group_member(id, nested)))
    }

    /// Appends a rectangular table.
    pub fn add_table(
        &mut self,
        rows: usize,
        columns: usize,
        left: Emu,
        top: Emu,
        width: Emu,
        height: Emu,
    ) -> Result<ShapeMut<'_>> {
        self.append(|id| table_member(id, rows, columns, left, top, width, height))
    }

    /// Appends a relationship-backed picture, sized as
    /// [`Presentation::add_picture`] sizes it.
    pub fn add_picture(
        &mut self,
        image_data: &[u8],
        image_filename: &str,
        left: Emu,
        top: Emu,
        width: Option<Emu>,
        height: Option<Emu>,
    ) -> Result<ShapeMut<'_>> {
        self.presentation
            .append_picture(
                self.slide_index,
                &self.group,
                image_data,
                image_filename,
                left,
                top,
                width,
                height,
            )
            .map(shape_mut)
    }

    fn append(
        &mut self,
        build: impl FnOnce(u32) -> Result<ShapeTreeChild>,
    ) -> Result<ShapeMut<'_>> {
        let tree = &mut self.presentation.slides[self.slide_index]
            .slide
            .common_slide_data
            .shape_tree;
        append_new_member(tree, &self.group, build)
    }
}

fn textbox_member(id: u32, left: Emu, top: Emu, width: Emu, height: Emu) -> Result<ShapeTreeChild> {
    CT_Shape::new_textbox(
        id,
        &format!("TextBox {id}"),
        positioned_transform(left, top, width, height),
    )
    .map(ShapeTreeChild::Shape)
    .map_err(|error| invalid_shape_construction("add textbox", error))
}

fn preset_member(
    id: u32,
    preset: &str,
    left: Emu,
    top: Emu,
    width: Emu,
    height: Emu,
) -> Result<ShapeTreeChild> {
    let mut shape = CT_Shape::new_preset(
        id,
        &format!("Shape {id}"),
        preset,
        positioned_transform(left, top, width, height),
    )
    .map_err(|error| invalid_shape_construction("add shape", error))?;
    shape
        .set_default_style()
        .map_err(|error| invalid_shape_construction("add shape", error))?;
    Ok(ShapeTreeChild::Shape(shape))
}

fn unstyled_preset_member(
    id: u32,
    preset: &str,
    left: Emu,
    top: Emu,
    width: Emu,
    height: Emu,
) -> Result<ShapeTreeChild> {
    CT_Shape::new_preset(
        id,
        &format!("Shape {id}"),
        preset,
        positioned_transform(left, top, width, height),
    )
    .map(ShapeTreeChild::Shape)
    .map_err(|error| invalid_shape_construction("add shape", error))
}

fn connector_member(
    id: u32,
    connector: ConnectorType,
    begin_x: Emu,
    begin_y: Emu,
    end_x: Emu,
    end_y: Emu,
) -> Result<ShapeTreeChild> {
    let transform = connector_transform(begin_x, begin_y, end_x, end_y)?;
    let mut connector = CT_ConnectionShape::new_free_standing(
        id,
        &format!("Connector {id}"),
        connector.preset_name(),
        transform,
    )
    .map_err(|error| invalid_shape_construction("add connector", error))?;
    connector
        .set_default_style()
        .map_err(|error| invalid_shape_construction("add connector", error))?;
    Ok(ShapeTreeChild::Connector(connector))
}

/// Builds an empty group, with a zero `a:xfrm` when `nested` in a group.
fn group_member(id: u32, nested: bool) -> ShapeTreeChild {
    let mut group = CT_GroupShape::new_empty(id, &format!("Group {id}"));
    if nested {
        let transform = group.group_transform_mut();
        transform.offset = Some(CT_Point2D::default());
        transform.extent = Some(CT_PositiveSize2D::default());
        transform.child_offset = Some(CT_Point2D::default());
        transform.child_extent = Some(CT_PositiveSize2D::default());
    }
    ShapeTreeChild::GroupShape(Box::new(group))
}

fn table_member(
    id: u32,
    rows: usize,
    columns: usize,
    left: Emu,
    top: Emu,
    width: Emu,
    height: Emu,
) -> Result<ShapeTreeChild> {
    let table = CT_Table::new(rows, columns, width, height)
        .map_err(|error| invalid_table_mutation("add table", error.to_string()))?;
    let frame = CT_GraphicFrame::new_table(
        id,
        &format!("Table {id}"),
        positioned_transform(left, top, width, height),
        table,
    )
    .map_err(|error| invalid_table_mutation("add table", error.to_string()))?;
    Ok(ShapeTreeChild::GraphicFrame(Box::new(frame)))
}

/// Appends a new member built with an id that is unused across the slide.
fn append_new_member<'t>(
    tree: &'t mut CT_ShapeTree,
    group: &[usize],
    build: impl FnOnce(u32) -> Result<ShapeTreeChild>,
) -> Result<ShapeMut<'t>> {
    let id = ShapeIdAllocator::scan(tree).allocate();
    let member = build(id)?;
    Ok(shape_mut(append_member(tree, group, member)))
}

/// Appends `member` to the slide's own shapes, for an empty `group`, or to
/// the group that `group` indexes, then refits that group and each group
/// enclosing it, innermost first.
///
/// Callers pass a path that ends at a group.
fn append_member<'t>(
    tree: &'t mut CT_ShapeTree,
    group: &[usize],
    member: ShapeTreeChild,
) -> &'t mut ShapeTreeChild {
    const CHECKED: &str = "the caller checked the group path";
    if group.is_empty() {
        return tree.append_child(member);
    }
    // Only a group that already has a counted member has members whose drawn
    // place its refit must keep.
    let occupied = (1..=group.len())
        .map(|depth| {
            group_at_mut(&mut tree.children, &group[..depth])
                .expect(CHECKED)
                .children
                .iter()
                .any(|child| member_box(child).is_some())
        })
        .collect::<Vec<_>>();
    group_at_mut(&mut tree.children, group)
        .expect(CHECKED)
        .append_child(member);
    for depth in (1..=group.len()).rev() {
        fit_group_to_members(
            group_at_mut(&mut tree.children, &group[..depth]).expect(CHECKED),
            occupied[depth - 1],
        );
    }
    group_at_mut(&mut tree.children, group)
        .expect(CHECKED)
        .children
        .last_mut()
        .expect("the member was appended")
}

/// Returns the group that `path` indexes from `children`, one z-order index
/// per nesting level.
fn group_at_mut<'t>(
    children: &'t mut [ShapeTreeChild],
    path: &[usize],
) -> Option<&'t mut CT_GroupShape> {
    let (index, rest) = path.split_first()?;
    let ShapeTreeChild::GroupShape(group) = children.get_mut(*index)? else {
        return None;
    };
    if rest.is_empty() {
        Some(group.as_mut())
    } else {
        group_at_mut(&mut group.children, rest)
    }
}

/// Returns the offset and extent a member adds to its group's union, or
/// `None` for a member without both and for a group without members.
fn member_box(child: &ShapeTreeChild) -> Option<(CT_Point2D, CT_PositiveSize2D)> {
    if let ShapeTreeChild::GroupShape(group) = child
        && group.children.is_empty()
    {
        return None;
    }
    let transform = shape_transform(child)?;
    Some((transform.offset?, transform.extent?))
}

/// Refits a group to the union of its members' own boxes, as python-pptx
/// does after each addition.
///
/// `a:chOff` and `a:chExt` become the union in member space, and `a:off` and
/// `a:ext` follow through the group's current mapping from member space to
/// its parent, the one the renderer applies. The group's flips and rotation
/// apply about the centre of `a:off` and `a:ext`, so when the group is
/// `occupied`, meaning it had a counted member before the addition, `a:off`
/// also moves by [`pivot_shift`]. No member drawn before the addition moves
/// on the slide. A group without all four values draws its members unscaled
/// in place, so its four values end equal, as python-pptx writes them. As
/// in python-pptx, member rotation does not widen the union. A member
/// without an offset and an extent does not count, nor does a group without
/// members, and a group without a counted member keeps its transform.
fn fit_group_to_members(group: &mut CT_GroupShape, occupied: bool) {
    let mut union: Option<(i64, i64, i64, i64)> = None;
    for (offset, extent) in group.children.iter().filter_map(member_box) {
        let (left, top) = (offset.x.0, offset.y.0);
        let right = left.saturating_add(extent.cx.0);
        let bottom = top.saturating_add(extent.cy.0);
        union = Some(union.map_or((left, top, right, bottom), |(l, t, r, b)| {
            (l.min(left), t.min(top), r.max(right), b.max(bottom))
        }));
    }
    let Some((left, top, right, bottom)) = union else {
        return;
    };
    let width = right.saturating_sub(left).max(0);
    let height = bottom.saturating_sub(top).max(0);
    let transform = group.group_transform_mut();
    let mapping = transform
        .offset
        .zip(transform.extent)
        .zip(transform.child_offset.zip(transform.child_extent));
    let (mut x, cx) = member_span_to_parent(
        left,
        width,
        mapping.map(|((offset, extent), (child_offset, child_extent))| {
            (offset.x.0, extent.cx.0, child_offset.x.0, child_extent.cx.0)
        }),
    );
    let (mut y, cy) = member_span_to_parent(
        top,
        height,
        mapping.map(|((offset, extent), (child_offset, child_extent))| {
            (offset.y.0, extent.cy.0, child_offset.y.0, child_extent.cy.0)
        }),
    );
    if occupied && let Some(((offset, extent), _)) = mapping {
        let centre = |start: i64, length: i64| start as f64 + length as f64 / 2.0;
        let (shift_x, shift_y) = pivot_shift(
            transform,
            centre(x, cx) - centre(offset.x.0, extent.cx.0),
            centre(y, cy) - centre(offset.y.0, extent.cy.0),
        );
        x = x.saturating_add(shift_x);
        y = y.saturating_add(shift_y);
    }
    transform.offset = Some(CT_Point2D {
        x: Emu(x),
        y: Emu(y),
    });
    transform.extent = Some(CT_PositiveSize2D {
        cx: Emu(cx),
        cy: Emu(cy),
    });
    transform.child_offset = Some(CT_Point2D {
        x: Emu(left),
        y: Emu(top),
    });
    transform.child_extent = Some(CT_PositiveSize2D {
        cx: Emu(width),
        cy: Emu(height),
    });
}

/// Returns how far a group's `a:off` must move so that its flips and
/// rotation, which apply about the centre of `a:off` and `a:ext`, place its
/// members as before after that centre moved by `(dx, dy)`.
///
/// PowerPoint flips, then rotates, as the rpptx renderer does, so for that
/// linear part `L` the shift is `(L - I)(dx, dy)`, which is zero without
/// rotation and flips.
fn pivot_shift(transform: &CT_Transform2D, dx: f64, dy: f64) -> (i64, i64) {
    let (sin, cos) = (f64::from(transform.rotation.0) / 60_000.0)
        .to_radians()
        .sin_cos();
    let flip_x = if transform.flip_horizontal { -1.0 } else { 1.0 };
    let flip_y = if transform.flip_vertical { -1.0 } else { 1.0 };
    (
        ((cos * flip_x - 1.0) * dx - sin * flip_y * dy).round() as i64,
        (sin * flip_x * dx + (cos * flip_y - 1.0) * dy).round() as i64,
    )
}

/// Maps one axis of a member-space span onto a group's parent, given the
/// group's offset, extent, child offset, and child extent on that axis, as
/// `offset + (start - child_offset) * extent / child_extent`. A zero child
/// extent scales by one, as in the renderer, and `None` maps unchanged.
fn member_span_to_parent(
    start: i64,
    length: i64,
    mapping: Option<(i64, i64, i64, i64)>,
) -> (i64, i64) {
    let Some((offset, extent, child_offset, child_extent)) = mapping else {
        return (start, length);
    };
    let scale = |value: i64| {
        if child_extent == 0 || extent == child_extent {
            value
        } else {
            (value as f64 * extent as f64 / child_extent as f64).round() as i64
        }
    };
    (
        offset.saturating_add(scale(start.saturating_sub(child_offset))),
        scale(length),
    )
}

/// The geometry family of a free-standing connector.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum ConnectorType {
    Straight,
    Elbow,
    Curve,
}

impl ConnectorType {
    const fn preset_name(self) -> &'static str {
        match self {
            Self::Straight => "line",
            Self::Elbow => "bentConnector3",
            Self::Curve => "curvedConnector3",
        }
    }
}

fn positioned_transform(left: Emu, top: Emu, width: Emu, height: Emu) -> CT_Transform2D {
    let mut transform = CT_Transform2D::default();
    transform.offset = Some(CT_Point2D { x: left, y: top });
    transform.extent = Some(CT_PositiveSize2D {
        cx: width,
        cy: height,
    });
    transform
}

fn connector_transform(
    begin_x: Emu,
    begin_y: Emu,
    end_x: Emu,
    end_y: Emu,
) -> Result<CT_Transform2D> {
    let width =
        i64::try_from(begin_x.0.abs_diff(end_x.0)).map_err(|_| Error::InvalidShapeMutation {
            operation: "add connector",
            message: "horizontal endpoint span exceeds the EMU range".to_owned(),
        })?;
    let height =
        i64::try_from(begin_y.0.abs_diff(end_y.0)).map_err(|_| Error::InvalidShapeMutation {
            operation: "add connector",
            message: "vertical endpoint span exceeds the EMU range".to_owned(),
        })?;
    let mut transform = CT_Transform2D::default();
    transform.offset = Some(CT_Point2D {
        x: Emu(begin_x.0.min(end_x.0)),
        y: Emu(begin_y.0.min(end_y.0)),
    });
    transform.extent = Some(CT_PositiveSize2D {
        cx: Emu(width),
        cy: Emu(height),
    });
    transform.flip_horizontal = begin_x.0 > end_x.0;
    transform.flip_vertical = begin_y.0 > end_y.0;
    Ok(transform)
}

fn invalid_shape_construction(operation: &'static str, error: OxmlError) -> Error {
    Error::InvalidShapeMutation {
        operation,
        message: error.to_string(),
    }
}

fn invalid_table_mutation(operation: &'static str, message: String) -> Error {
    Error::InvalidTableMutation { operation, message }
}

impl<'a> SlideRef<'a> {
    /// Returns the producer-assigned slide id.
    pub fn id(&self) -> u32 {
        self.record.id
    }

    /// Returns the optional `p:cSld` name.
    pub fn name(&self) -> Option<&str> {
        self.record.slide.common_slide_data.name.as_deref()
    }

    /// Returns whether the slide is hidden from the slideshow.
    pub fn hidden(&self) -> bool {
        self.record.slide.hidden()
    }

    /// Returns whether the slide carries an explicit `p:bg` element.
    pub fn has_explicit_background(&self) -> bool {
        self.record.slide.has_explicit_background()
    }

    /// Returns the direct `p:bgPr` fill when the slide background has one.
    pub fn background_fill(&self) -> Option<&'a Fill> {
        match self
            .record
            .slide
            .common_slide_data
            .background
            .as_ref()?
            .rendering()
        {
            BackgroundRendering::Properties(fill) => fill.as_deref(),
            BackgroundRendering::Reference { .. } | BackgroundRendering::Unsupported(_) => None,
        }
    }

    /// Returns one immediate z-order child by zero-based index.
    pub fn shape(&self, index: usize) -> Option<ShapeRef<'a>> {
        self.record
            .slide
            .common_slide_data
            .shape_tree
            .children
            .get(index)
            .map(shape_ref)
    }

    /// Returns the first title or centered-title placeholder on the slide.
    pub fn title(&self) -> Option<ShapeRef<'_>> {
        self.shapes().find(|shape| {
            matches!(
                shape_placeholder(shape.child).and_then(|placeholder| placeholder.ph_type.as_ref()),
                Some(PhType::Title | PhType::CenteredTitle)
            )
        })
    }

    /// Returns the placeholder whose OOXML placeholder index equals `index`.
    pub fn placeholder(&self, index: u32) -> Option<ShapeRef<'_>> {
        self.shapes().find(|shape| {
            shape_placeholder(shape.child).map(|placeholder| placeholder.idx.unwrap_or(0))
                == Some(index)
        })
    }

    /// Iterates the immediate shape tree in z-order.
    pub fn shapes(&self) -> impl ExactSizeIterator<Item = ShapeRef<'a>> + 'a {
        self.record
            .slide
            .common_slide_data
            .shape_tree
            .children
            .iter()
            .map(shape_ref)
    }

    /// Returns visible shape and table text recursively in z-order.
    pub fn text(&self) -> String {
        let mut text = Vec::new();
        collect_text(
            &self.record.slide.common_slide_data.shape_tree.children,
            &mut text,
        );
        text.join("\n")
    }

    /// Returns speaker-note text, distinguishing no notes part from empty notes.
    pub fn notes_text(&self) -> Option<String> {
        self.record
            .notes
            .as_ref()
            .map(|notes| notes.notes.notes_text())
    }
}

/// The structural kind of one PresentationML shape-tree child.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum ShapeKind {
    Shape,
    Picture,
    GraphicFrame,
    Group,
    Connector,
    AlternateContent,
}

/// One DrawingML text body autofit choice.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum AutofitMode {
    None,
    Normal,
    Shape,
}

/// The python-pptx `MSO_SHAPE_TYPE` classification of one shape-tree child.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum ShapeType {
    AutoShape,
    Chart,
    EmbeddedOleObject,
    Freeform,
    Group,
    Line,
    LinkedOleObject,
    Media,
    Picture,
    Placeholder,
    Table,
    TextBox,
}

/// A borrowed shape-tree child.
#[derive(Clone, Copy)]
pub struct ShapeRef<'a> {
    child: &'a ShapeTreeChild,
}

impl PartialEq for ShapeRef<'_> {
    fn eq(&self, other: &Self) -> bool {
        std::ptr::eq(self.child, other.child)
    }
}

impl Eq for ShapeRef<'_> {}

fn shape_ref(child: &ShapeTreeChild) -> ShapeRef<'_> {
    ShapeRef { child }
}

/// A mutably borrowed shape-tree child.
pub struct ShapeMut<'a> {
    child: &'a mut ShapeTreeChild,
}

fn shape_mut(child: &mut ShapeTreeChild) -> ShapeMut<'_> {
    ShapeMut { child }
}

impl<'a> ShapeMut<'a> {
    /// Returns the child's normalized structural kind.
    pub fn kind(&self) -> ShapeKind {
        shape_kind(self.child)
    }

    /// Returns one immediate group child for in-place mutation.
    ///
    /// `mc:AlternateContent` fallback children remain a read-only projection.
    pub fn child_mut(&mut self, index: usize) -> Option<ShapeMut<'_>> {
        match self.child {
            ShapeTreeChild::GroupShape(group) => group.children.get_mut(index).map(shape_mut),
            _ => None,
        }
    }

    /// Consumes this handle and returns one immediate group child for mutation.
    pub fn into_child_mut(self, index: usize) -> Option<ShapeMut<'a>> {
        match self.child {
            ShapeTreeChild::GroupShape(group) => group.children.get_mut(index).map(shape_mut),
            _ => None,
        }
    }

    /// Changes the shape offset in EMU.
    pub fn set_position(&mut self, left: Emu, top: Emu) -> Result<()> {
        self.transform_mut("set position")?.offset = Some(CT_Point2D { x: left, y: top });
        Ok(())
    }

    /// Changes the shape extent in EMU.
    pub fn set_size(&mut self, width: Emu, height: Emu) -> Result<()> {
        self.transform_mut("set size")?.extent = Some(CT_PositiveSize2D {
            cx: width,
            cy: height,
        });
        Ok(())
    }

    /// Changes the clockwise DrawingML rotation.
    pub fn set_rotation(&mut self, rotation: Angle) -> Result<()> {
        self.transform_mut("set rotation")?.rotation = rotation;
        Ok(())
    }

    /// Changes the producer-facing non-visual name.
    pub fn set_name(&mut self, name: &str) -> Result<()> {
        let result = match self.child {
            ShapeTreeChild::Shape(shape) => shape.set_name(name),
            ShapeTreeChild::Picture(picture) => picture.set_name(name),
            ShapeTreeChild::GraphicFrame(frame) => frame.set_name(name),
            ShapeTreeChild::GroupShape(group) => group.set_name(name),
            ShapeTreeChild::Connector(connector) => connector.set_name(name),
            ShapeTreeChild::AlternateContent(_) => {
                return Err(self.unsupported("set name"));
            }
        };
        result.map_err(|error| Error::InvalidShapeMutation {
            operation: "set name",
            message: error.to_string(),
        })
    }

    /// Sets the alternative text screen readers announce, `p:cNvPr/@descr`,
    /// or removes it for `None`.
    pub fn set_alt_text(&mut self, text: Option<&str>) -> Result<()> {
        self.child
            .set_alt_text(text)
            .map_err(|error| invalid_shape_mutation("set alt text", error.to_string()))
    }

    /// Sets the alternative-text title, `p:cNvPr/@title`, or removes it for
    /// `None`.
    pub fn set_alt_title(&mut self, title: Option<&str>) -> Result<()> {
        self.child
            .set_alt_title(title)
            .map_err(|error| invalid_shape_mutation("set alt title", error.to_string()))
    }

    /// Marks the shape decorative, so that screen readers skip it, or removes
    /// the mark. PowerPoint writes the flag as an `adec:decorative` extension
    /// of `p:cNvPr`.
    pub fn set_decorative(&mut self, decorative: bool) -> Result<()> {
        self.child
            .set_decorative(decorative)
            .map_err(|error| invalid_shape_mutation("set decorative", error.to_string()))
    }

    /// Mirrors the shape left to right about its centre, `a:xfrm/@flipH`.
    ///
    /// A connector flipped this way swaps the horizontal place of its two
    /// ends. A table, chart, or other graphic frame does not flip in
    /// PowerPoint and is rejected.
    pub fn set_flip_horizontal(&mut self, flipped: bool) -> Result<()> {
        self.flip_transform_mut("set flip horizontal")?
            .flip_horizontal = flipped;
        Ok(())
    }

    /// Mirrors the shape top to bottom about its centre, `a:xfrm/@flipV`, as
    /// [`Self::set_flip_horizontal`] does left to right.
    pub fn set_flip_vertical(&mut self, flipped: bool) -> Result<()> {
        self.flip_transform_mut("set flip vertical")?.flip_vertical = flipped;
        Ok(())
    }

    fn flip_transform_mut(&mut self, operation: &'static str) -> Result<&mut CT_Transform2D> {
        if matches!(self.child, ShapeTreeChild::GraphicFrame(_)) {
            return Err(self.unsupported(operation));
        }
        self.transform_mut(operation)
    }

    /// Replaces the direct fill on a shape with typed shape properties.
    pub fn set_fill(&mut self, fill: Fill) -> Result<()> {
        self.shape_properties_mut("set fill")?.fill = Some(fill);
        Ok(())
    }

    /// Replaces the direct line on a shape with typed shape properties.
    pub fn set_line(&mut self, line: CT_LineProperties) -> Result<()> {
        self.shape_properties_mut("set line")?.line = Some(line);
        Ok(())
    }

    /// Crops a picture by left, top, right, and bottom insets of its image.
    ///
    /// Each inset is a share of the image, where `Percent1000(25_000)` is a
    /// quarter, and a negative inset extends the image. An edge keeps its
    /// stored attribute when its value is unchanged, a changed edge of zero
    /// drops its attribute, and four zero insets add no `a:srcRect` to a
    /// picture without one. Other shape kinds are rejected.
    pub fn set_crop(
        &mut self,
        left: Percent1000,
        top: Percent1000,
        right: Percent1000,
        bottom: Percent1000,
    ) -> Result<()> {
        const OPERATION: &str = "set crop";
        let shape_kind = shape_kind(self.child);
        let ShapeTreeChild::Picture(picture) = self.child else {
            return Err(Error::UnsupportedShapeMutation {
                operation: OPERATION,
                shape_kind,
            });
        };
        let fill = picture
            .blip_fill
            .as_mut()
            .ok_or_else(|| invalid_shape_mutation(OPERATION, "picture has no p:blipFill"))?;
        let zero = Percent1000::default();
        if fill.source_rect.is_none() && [left, top, right, bottom] == [zero; 4] {
            return Ok(());
        }
        let edge = |stored: Option<Percent1000>, value: Percent1000| {
            if stored.unwrap_or_default() == value {
                stored
            } else {
                (value != zero).then_some(value)
            }
        };
        let rect = fill.source_rect.get_or_insert_with(RelativeRect::default);
        rect.left = edge(rect.left, left);
        rect.top = edge(rect.top, top);
        rect.right = edge(rect.right, right);
        rect.bottom = edge(rect.bottom, bottom);
        Ok(())
    }

    /// Replaces or removes the direct `a:effectLst` of a shape, picture,
    /// connector, or group.
    ///
    /// `None` removes the list so the shape inherits its theme effect again,
    /// and an empty list suppresses that effect. The list is written in its
    /// schema place and keeps the unmodelled effects it carries. A shape with
    /// an `a:effectDag` refuses a list, which the schema excludes beside it.
    pub fn set_effects(&mut self, effects: Option<CT_EffectList>) -> Result<()> {
        const OPERATION: &str = "set effects";
        if let Some(list) = &effects {
            list.to_xml()
                .map_err(|error| invalid_shape_mutation(OPERATION, error.to_string()))?;
        }
        if let ShapeTreeChild::GroupShape(group) = self.child {
            return group
                .set_effects(effects)
                .map_err(|error| invalid_shape_mutation(OPERATION, error.to_string()));
        }
        let properties = self.shape_properties_mut(OPERATION)?;
        if effects.is_some() && properties.has_unmodelled_effect() {
            return Err(invalid_shape_mutation(
                OPERATION,
                "the shape carries an a:effectDag, which excludes an a:effectLst",
            ));
        }
        properties.effects = effects;
        Ok(())
    }

    /// Changes the theme effect style that the shape's `p:style` references.
    ///
    /// Index 0 references no theme effect, so a connector from
    /// [`SlideMut::add_connector`] loses the theme's shadow in PowerPoint,
    /// LibreOffice, and the renderer alike. An empty direct `a:effectLst`
    /// does not remove it in LibreOffice. Only an ordinary shape or a
    /// connector with a typed style has this reference.
    pub fn set_theme_effect_index(&mut self, index: u32) -> Result<()> {
        let changed = match self.child {
            ShapeTreeChild::Shape(shape) => Ok(shape.set_effect_reference_index(index)),
            ShapeTreeChild::Connector(connector) => connector.set_effect_reference_index(index),
            _ => return Err(self.unsupported("set theme effect index")),
        };
        match changed {
            Ok(true) => Ok(()),
            Ok(false) => Err(Error::InvalidShapeMutation {
                operation: "set theme effect index",
                message: "the shape has no p:style with typed theme references".to_owned(),
            }),
            Err(error) => Err(Error::InvalidShapeMutation {
                operation: "set theme effect index",
                message: error.to_string(),
            }),
        }
    }

    /// Inserts or replaces one finite preset-geometry adjustment.
    pub fn set_adjust_value(&mut self, name: &str, value: f64) -> Result<()> {
        if !value.is_finite() {
            return Err(Error::NonFiniteAdjustmentValue {
                name: name.to_owned(),
            });
        }
        let shape_kind = shape_kind(self.child);
        let properties = self.shape_properties_mut("set adjustment")?;
        let geometry =
            properties
                .preset_geometry
                .as_mut()
                .ok_or(Error::UnsupportedShapeMutation {
                    operation: "set adjustment",
                    shape_kind,
                })?;
        geometry
            .set_adjust_value(name, value)
            .map_err(|error| Error::InvalidShapeMutation {
                operation: "set adjustment",
                message: error.to_string(),
            })
    }

    /// Replaces the geometry of an auto shape or a picture mask with a preset.
    ///
    /// `preset` is a DrawingML preset name such as `roundRect`. A custom
    /// geometry is replaced in the same schema slot, and the adjustments
    /// restart from the preset's defaults with an empty `a:avLst`. Text,
    /// fill, line, effects, identity and z-order are kept. A text box, a
    /// connector, a graphic frame, a group and alternate content are refused.
    pub fn set_auto_shape_type(&mut self, preset: &str) -> Result<()> {
        const OPERATION: &str = "set auto shape type";
        let properties = match self.child {
            ShapeTreeChild::Shape(shape) if shape.is_textbox() => {
                return Err(Error::InvalidShapeMutation {
                    operation: OPERATION,
                    message: "a text box is not an auto shape".to_owned(),
                });
            }
            ShapeTreeChild::Shape(shape) => &mut shape.shape_properties,
            ShapeTreeChild::Picture(picture) => &mut picture.shape_properties,
            _ => return Err(self.unsupported(OPERATION)),
        };
        let geometry =
            CT_PresetGeometry2D::new(preset).map_err(|error| Error::InvalidShapeMutation {
                operation: OPERATION,
                message: error.to_string(),
            })?;
        properties.custom_geometry = None;
        properties.preset_geometry = Some(geometry);
        Ok(())
    }

    /// Replaces ordinary shape text without changing placeholder identity,
    /// one paragraph per line as [`CT_TextBody::set_text`] describes.
    pub fn set_text(&mut self, text: &str) -> Result<()> {
        let shape_kind = shape_kind(self.child);
        let ShapeTreeChild::Shape(shape) = self.child else {
            return Err(Error::UnsupportedShapeMutation {
                operation: "set text",
                shape_kind,
            });
        };
        shape
            .text_body
            .get_or_insert_with(CT_TextBody::new)
            .set_text(text);
        Ok(())
    }

    /// Returns the ordinary shape text body for in-place mutation.
    pub fn text_frame(&mut self) -> Option<TextFrame<'_>> {
        match self.child {
            ShapeTreeChild::Shape(shape) => shape.text_body.as_mut().map(|body| TextFrame { body }),
            _ => None,
        }
    }

    /// Consumes this handle and returns its ordinary-shape text body for mutation.
    pub fn into_text_frame(self) -> Option<TextFrame<'a>> {
        match self.child {
            ShapeTreeChild::Shape(shape) => shape.text_body.as_mut().map(|body| TextFrame { body }),
            _ => None,
        }
    }

    /// Returns a behavior-bearing mutable handle for a table graphic frame.
    pub fn table_mut(&mut self) -> Option<TableMut<'_>> {
        let ShapeTreeChild::GraphicFrame(frame) = self.child else {
            return None;
        };
        let table = frame.graphic_data.table_mut()?;
        Some(TableMut {
            table,
            transform: &mut frame.transform,
        })
    }

    /// Consumes this handle and returns its table for mutation.
    pub fn into_table_mut(self) -> Option<TableMut<'a>> {
        let ShapeTreeChild::GraphicFrame(frame) = self.child else {
            return None;
        };
        let table = frame.graphic_data.table_mut()?;
        Some(TableMut {
            table,
            transform: &mut frame.transform,
        })
    }

    fn transform_mut(&mut self, operation: &'static str) -> Result<&mut CT_Transform2D> {
        match self.child {
            ShapeTreeChild::Shape(shape) => Ok(shape
                .shape_properties
                .transform
                .get_or_insert_with(CT_Transform2D::default)),
            ShapeTreeChild::Picture(picture) => Ok(picture
                .shape_properties
                .transform
                .get_or_insert_with(CT_Transform2D::default)),
            ShapeTreeChild::GraphicFrame(frame) => Ok(&mut frame.transform),
            ShapeTreeChild::GroupShape(group) => Ok(group.group_transform_mut()),
            ShapeTreeChild::Connector(connector) => Ok(connector
                .shape_properties
                .transform
                .get_or_insert_with(CT_Transform2D::default)),
            ShapeTreeChild::AlternateContent(_) => Err(self.unsupported(operation)),
        }
    }

    fn shape_properties_mut(&mut self, operation: &'static str) -> Result<&mut CT_ShapeProperties> {
        match self.child {
            ShapeTreeChild::Shape(shape) => Ok(&mut shape.shape_properties),
            ShapeTreeChild::Picture(picture) => Ok(&mut picture.shape_properties),
            ShapeTreeChild::Connector(connector) => Ok(&mut connector.shape_properties),
            _ => Err(self.unsupported(operation)),
        }
    }

    fn unsupported(&self, operation: &'static str) -> Error {
        Error::UnsupportedShapeMutation {
            operation,
            shape_kind: shape_kind(self.child),
        }
    }
}

/// A mutably borrowed ordinary-shape text body.
pub struct TextFrame<'a> {
    body: &'a mut CT_TextBody,
}

/// A borrowed ordinary-shape text body.
#[derive(Clone, Copy)]
pub struct TextFrameRef<'a> {
    body: &'a CT_TextBody,
}

/// A borrowed DrawingML text paragraph.
#[derive(Clone, Copy)]
pub struct TextParagraphRef<'a> {
    paragraph: &'a CT_TextParagraph,
}

/// A borrowed regular DrawingML text run.
#[derive(Clone, Copy)]
pub struct TextRunRef<'a> {
    run: &'a CT_RegularTextRun,
}

/// A borrowed DrawingML table.
#[derive(Clone, Copy)]
pub struct TableRef<'a> {
    table: &'a CT_Table,
}

impl<'a> TableRef<'a> {
    /// Returns the number of explicit table rows.
    pub fn row_count(&self) -> usize {
        self.table.rows.len()
    }

    /// Returns the number of grid columns.
    pub fn column_count(&self) -> usize {
        self.table.grid.columns.len()
    }

    /// Returns one explicit grid cell by zero-based row and column.
    pub fn cell(&self, row: usize, column: usize) -> Option<TableCellRef<'a>> {
        table_cell(self.table, row, column).map(|cell| TableCellRef { cell })
    }

    /// Returns one grid-column width in EMU.
    pub fn column_width(&self, column: usize) -> Option<Emu> {
        self.table.grid.columns.get(column).copied()
    }

    /// Returns one stored row height in EMU.
    pub fn row_height(&self, row: usize) -> Option<Emu> {
        self.table.rows.get(row).map(|row| row.height)
    }

    /// Returns whether first-row table styling is enabled.
    pub fn first_row(&self) -> bool {
        self.table
            .properties
            .as_ref()
            .is_some_and(|properties| properties.first_row)
    }

    /// Returns whether last-row table styling is enabled.
    pub fn last_row(&self) -> bool {
        self.table
            .properties
            .as_ref()
            .is_some_and(|properties| properties.last_row)
    }

    /// Returns whether first-column table styling is enabled.
    pub fn first_column(&self) -> bool {
        self.table
            .properties
            .as_ref()
            .is_some_and(|properties| properties.first_column)
    }

    /// Returns whether last-column table styling is enabled.
    pub fn last_column(&self) -> bool {
        self.table
            .properties
            .as_ref()
            .is_some_and(|properties| properties.last_column)
    }

    /// Returns whether horizontal row banding is enabled.
    pub fn horizontal_banding(&self) -> bool {
        self.table
            .properties
            .as_ref()
            .is_some_and(|properties| properties.band_rows)
    }

    /// Returns whether vertical column banding is enabled.
    pub fn vertical_banding(&self) -> bool {
        self.table
            .properties
            .as_ref()
            .is_some_and(|properties| properties.band_columns)
    }
}

/// A behavior-bearing mutable DrawingML table and its containing frame extent.
pub struct TableMut<'a> {
    table: &'a mut CT_Table,
    transform: &'a mut CT_Transform2D,
}

impl<'a> TableMut<'a> {
    /// Returns the number of explicit table rows.
    pub fn row_count(&self) -> usize {
        self.table.rows.len()
    }

    /// Returns the number of grid columns.
    pub fn column_count(&self) -> usize {
        self.table.grid.columns.len()
    }

    /// Returns one explicit grid cell by zero-based row and column.
    pub fn cell(&self, row: usize, column: usize) -> Option<TableCellRef<'_>> {
        table_cell(self.table, row, column).map(|cell| TableCellRef { cell })
    }

    /// Returns one explicit grid cell for in-place mutation.
    pub fn cell_mut(&mut self, row: usize, column: usize) -> Option<TableCellMut<'_>> {
        table_cell(self.table, row, column)?;
        Some(TableCellMut {
            table: self.table,
            row,
            column,
        })
    }

    /// Consumes this handle and returns one explicit cell for mutation.
    pub fn into_cell_mut(self, row: usize, column: usize) -> Option<TableCellMut<'a>> {
        table_cell(self.table, row, column)?;
        Some(TableCellMut {
            table: self.table,
            row,
            column,
        })
    }

    /// Returns one grid-column width in EMU.
    pub fn column_width(&self, column: usize) -> Option<Emu> {
        self.table.grid.columns.get(column).copied()
    }

    /// Changes one grid width and synchronizes the containing frame width.
    pub fn set_column_width(&mut self, column: usize, width: Emu) -> Result<()> {
        const OPERATION: &str = "set column width";
        if width.0 <= 0 {
            return Err(invalid_table_mutation(
                OPERATION,
                "column width must be positive".to_owned(),
            ));
        }
        if column >= self.table.grid.columns.len() {
            return Err(invalid_table_mutation(
                OPERATION,
                format!("column index {column} is out of range"),
            ));
        }
        let mut staged = self.table.clone();
        staged
            .set_column_width(column, width)
            .map_err(|error| invalid_table_mutation(OPERATION, error.to_string()))?;
        let (total_width, total_height) = table_extent(&staged, OPERATION)?;
        staged
            .to_xml()
            .map_err(|error| invalid_table_mutation(OPERATION, error.to_string()))?;
        *self.table = staged;
        let height = self
            .transform
            .extent
            .map_or(total_height, |extent| extent.cy);
        self.transform.extent = Some(CT_PositiveSize2D {
            cx: total_width,
            cy: height,
        });
        Ok(())
    }

    /// Returns one stored row height in EMU.
    pub fn row_height(&self, row: usize) -> Option<Emu> {
        self.table.rows.get(row).map(|row| row.height)
    }

    /// Changes one row height and synchronizes the containing frame height.
    ///
    /// The stored height is a minimum, as in PowerPoint, which grows a row to
    /// fit its text when it lays the table out.
    pub fn set_row_height(&mut self, row: usize, height: Emu) -> Result<()> {
        const OPERATION: &str = "set row height";
        if height.0 <= 0 {
            return Err(invalid_table_mutation(
                OPERATION,
                "row height must be positive".to_owned(),
            ));
        }
        if row >= self.table.rows.len() {
            return Err(invalid_table_mutation(
                OPERATION,
                format!("row index {row} is out of range"),
            ));
        }
        let mut staged = self.table.clone();
        staged.rows[row].height = height;
        let (total_width, total_height) = table_extent(&staged, OPERATION)?;
        staged
            .to_xml()
            .map_err(|error| invalid_table_mutation(OPERATION, error.to_string()))?;
        self.table.rows[row].height = height;
        let width = self
            .transform
            .extent
            .map_or(total_width, |extent| extent.cx);
        self.transform.extent = Some(CT_PositiveSize2D {
            cx: width,
            cy: total_height,
        });
        Ok(())
    }

    /// Inserts a row before `index`, or appends one when `index` equals the
    /// row count, and grows the frame height by the new row's height.
    ///
    /// The new row copies the height of the row above it, or of the first row
    /// when it becomes the first, and each new cell copies that row's cell
    /// properties and first-paragraph formatting without its text. A row
    /// inserted inside a merged cell extends the merge.
    pub fn insert_row(&mut self, index: usize) -> Result<()> {
        self.restructure("insert row", |table| table.insert_row(index))
    }

    /// Removes one row, which must not be the only one, and shrinks the frame
    /// height by its height.
    ///
    /// A merged cell across the row loses one row, and when the row holds the
    /// merge origin the row below takes it over with the origin's text.
    pub fn remove_row(&mut self, index: usize) -> Result<()> {
        self.restructure("remove row", |table| table.remove_row(index))
    }

    /// Inserts a grid column before `index`, or appends one when `index`
    /// equals the column count, and grows the frame width by its width.
    ///
    /// The new column copies the width of the column left of it, or of the
    /// first column when it becomes the first, and its cells copy that
    /// column's cells as `insert_row` does. A column inserted inside a merged
    /// cell extends the merge.
    pub fn insert_column(&mut self, index: usize) -> Result<()> {
        self.restructure("insert column", |table| table.insert_column(index))
    }

    /// Removes one grid column, which must not be the only one, and shrinks
    /// the frame width by its width.
    ///
    /// A merged cell across the column loses one column, and when the column
    /// holds the merge origin the column to its right takes it over with the
    /// origin's text.
    pub fn remove_column(&mut self, index: usize) -> Result<()> {
        self.restructure("remove column", |table| table.remove_column(index))
    }

    /// Applies one row or column edit to a staged table and commits it.
    ///
    /// Each frame edge moves by the change of its grid sum, so the frame grows
    /// or shrinks by the inserted or removed row or column. A frame that
    /// PowerPoint measured taller than its stored rows, because it grew rows
    /// to fit their text, keeps that excess.
    fn restructure(
        &mut self,
        operation: &'static str,
        edit: impl FnOnce(&mut CT_Table) -> std::result::Result<(), OxmlError>,
    ) -> Result<()> {
        let (old_width, old_height) = table_extent(self.table, operation)?;
        let mut staged = self.table.clone();
        edit(&mut staged).map_err(|error| invalid_table_mutation(operation, error.to_string()))?;
        let (new_width, new_height) = table_extent(&staged, operation)?;
        staged
            .to_xml()
            .map_err(|error| invalid_table_mutation(operation, error.to_string()))?;
        let frame = self.transform.extent.unwrap_or(CT_PositiveSize2D {
            cx: old_width,
            cy: old_height,
        });
        let extent = CT_PositiveSize2D {
            cx: moved_frame_edge(frame.cx, old_width, new_width, operation)?,
            cy: moved_frame_edge(frame.cy, old_height, new_height, operation)?,
        };
        *self.table = staged;
        self.transform.extent = Some(extent);
        Ok(())
    }

    /// Returns whether first-row table styling is enabled.
    pub fn first_row(&self) -> bool {
        self.table
            .properties
            .as_ref()
            .is_some_and(|properties| properties.first_row)
    }

    /// Enables or disables first-row table styling.
    pub fn set_first_row(&mut self, value: bool) {
        table_properties_mut(self.table).first_row = value;
    }

    /// Returns whether last-row table styling is enabled.
    pub fn last_row(&self) -> bool {
        self.table
            .properties
            .as_ref()
            .is_some_and(|properties| properties.last_row)
    }

    /// Enables or disables last-row table styling.
    pub fn set_last_row(&mut self, value: bool) {
        table_properties_mut(self.table).last_row = value;
    }

    /// Returns whether first-column table styling is enabled.
    pub fn first_column(&self) -> bool {
        self.table
            .properties
            .as_ref()
            .is_some_and(|properties| properties.first_column)
    }

    /// Enables or disables first-column table styling.
    pub fn set_first_column(&mut self, value: bool) {
        table_properties_mut(self.table).first_column = value;
    }

    /// Returns whether last-column table styling is enabled.
    pub fn last_column(&self) -> bool {
        self.table
            .properties
            .as_ref()
            .is_some_and(|properties| properties.last_column)
    }

    /// Enables or disables last-column table styling.
    pub fn set_last_column(&mut self, value: bool) {
        table_properties_mut(self.table).last_column = value;
    }

    /// Returns whether horizontal row banding is enabled.
    pub fn horizontal_banding(&self) -> bool {
        self.table
            .properties
            .as_ref()
            .is_some_and(|properties| properties.band_rows)
    }

    /// Enables or disables horizontal row banding.
    pub fn set_horizontal_banding(&mut self, value: bool) {
        table_properties_mut(self.table).band_rows = value;
    }

    /// Returns whether vertical column banding is enabled.
    pub fn vertical_banding(&self) -> bool {
        self.table
            .properties
            .as_ref()
            .is_some_and(|properties| properties.band_columns)
    }

    /// Enables or disables vertical column banding.
    pub fn set_vertical_banding(&mut self, value: bool) {
        table_properties_mut(self.table).band_columns = value;
    }
}

/// One edge of a table cell, drawn by the `a:lnL`, `a:lnR`, `a:lnT`, or `a:lnB`
/// line of its `a:tcPr`.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum CellBorder {
    Left,
    Right,
    Top,
    Bottom,
}

/// A borrowed explicit table-grid cell.
#[derive(Clone, Copy)]
pub struct TableCellRef<'a> {
    cell: &'a CT_TableCell,
}

impl TableCellRef<'_> {
    /// Returns the cell's visible plain text.
    pub fn text(&self) -> String {
        self.cell
            .text_body
            .as_ref()
            .map_or_else(String::new, CT_TextBody::plain_text)
    }

    /// Returns the cell text body when present.
    pub fn text_frame(&self) -> Option<TextFrameRef<'_>> {
        self.cell
            .text_body
            .as_ref()
            .map(|body| TextFrameRef { body })
    }

    /// Returns whether this cell is the top-left origin of a merge.
    pub fn is_merge_origin(&self) -> bool {
        !self.is_spanned() && (self.cell.row_span > 1 || self.cell.grid_span > 1)
    }

    /// Returns whether this cell continues a merge from another grid cell.
    pub fn is_spanned(&self) -> bool {
        self.cell.horizontal_merge || self.cell.vertical_merge
    }

    /// Returns the origin's vertical span, or one for an unmerged cell.
    pub fn span_height(&self) -> u32 {
        self.cell.row_span
    }

    /// Returns the origin's horizontal span, or one for an unmerged cell.
    pub fn span_width(&self) -> u32 {
        self.cell.grid_span
    }

    /// Returns the direct cell fill, when present.
    pub fn fill(&self) -> Option<&Fill> {
        self.cell
            .properties
            .as_ref()
            .and_then(|properties| properties.fill.as_ref())
    }

    /// Returns the optional left, right, top, and bottom cell margins.
    pub fn margins(&self) -> (Option<Emu>, Option<Emu>, Option<Emu>, Option<Emu>) {
        self.cell
            .properties
            .as_ref()
            .map_or((None, None, None, None), |properties| {
                (
                    properties.margin_left,
                    properties.margin_right,
                    properties.margin_top,
                    properties.margin_bottom,
                )
            })
    }

    /// Returns the direct line of one cell edge, when present.
    pub fn border(&self, edge: CellBorder) -> Option<&CT_LineProperties> {
        let properties = self.cell.properties.as_ref()?;
        match edge {
            CellBorder::Left => properties.left.as_ref(),
            CellBorder::Right => properties.right.as_ref(),
            CellBorder::Top => properties.top.as_ref(),
            CellBorder::Bottom => properties.bottom.as_ref(),
        }
    }
}

/// A behavior-bearing mutable table-grid cell addressed within its table.
pub struct TableCellMut<'a> {
    table: &'a mut CT_Table,
    row: usize,
    column: usize,
}

impl TableCellMut<'_> {
    /// Returns the cell's visible plain text.
    pub fn text(&self) -> String {
        self.cell_ref().text()
    }

    /// Replaces the cell text with one paragraph per line, as
    /// [`CT_TextBody::set_text`] describes.
    pub fn set_text(&mut self, text: &str) {
        self.cell_mut()
            .text_body
            .get_or_insert_with(CT_TextBody::new)
            .set_text(text);
    }

    /// Returns the cell text body for in-place mutation.
    pub fn text_frame(&mut self) -> TextFrame<'_> {
        let body = self
            .cell_mut()
            .text_body
            .get_or_insert_with(CT_TextBody::new);
        TextFrame { body }
    }

    /// Replaces or clears the direct cell fill.
    pub fn set_fill(&mut self, fill: Option<Fill>) {
        self.properties_mut().fill = fill;
    }

    /// Returns the direct cell fill, when present.
    pub fn fill(&self) -> Option<&Fill> {
        self.cell_ref().cell.properties.as_ref()?.fill.as_ref()
    }

    /// Replaces all four optional cell margins.
    pub fn set_margins(
        &mut self,
        left: Option<Emu>,
        right: Option<Emu>,
        top: Option<Emu>,
        bottom: Option<Emu>,
    ) {
        let properties = self.properties_mut();
        properties.margin_left = left;
        properties.margin_right = right;
        properties.margin_top = top;
        properties.margin_bottom = bottom;
    }

    /// Returns the optional left, right, top, and bottom cell margins.
    pub fn margins(&self) -> (Option<Emu>, Option<Emu>, Option<Emu>, Option<Emu>) {
        let Some(properties) = self.cell_ref().cell.properties.as_ref() else {
            return (None, None, None, None);
        };
        (
            properties.margin_left,
            properties.margin_right,
            properties.margin_top,
            properties.margin_bottom,
        )
    }

    /// Returns the direct line of one cell edge, when present.
    pub fn border(&self, edge: CellBorder) -> Option<&CT_LineProperties> {
        let properties = self.cell_ref().cell.properties.as_ref()?;
        match edge {
            CellBorder::Left => properties.left.as_ref(),
            CellBorder::Right => properties.right.as_ref(),
            CellBorder::Top => properties.top.as_ref(),
            CellBorder::Bottom => properties.bottom.as_ref(),
        }
    }

    /// Replaces or clears the direct line of one cell edge.
    pub fn set_border(&mut self, edge: CellBorder, line: Option<CT_LineProperties>) {
        let properties = self.properties_mut();
        let slot = match edge {
            CellBorder::Left => &mut properties.left,
            CellBorder::Right => &mut properties.right,
            CellBorder::Top => &mut properties.top,
            CellBorder::Bottom => &mut properties.bottom,
        };
        *slot = line;
    }

    /// Returns whether this cell is the top-left origin of a merge.
    pub fn is_merge_origin(&self) -> bool {
        self.cell_ref().is_merge_origin()
    }

    /// Returns whether this cell continues a merge from another grid cell.
    pub fn is_spanned(&self) -> bool {
        self.cell_ref().is_spanned()
    }

    /// Returns the origin's vertical span, or one for an unmerged cell.
    pub fn span_height(&self) -> u32 {
        self.cell_ref().span_height()
    }

    /// Returns the origin's horizontal span, or one for an unmerged cell.
    pub fn span_width(&self) -> u32 {
        self.cell_ref().span_width()
    }

    /// Merges the rectangle between this cell and the other grid coordinate.
    pub fn merge_to(&mut self, row: usize, column: usize) -> Result<()> {
        let mut staged = self.table.clone();
        merge_cells(&mut staged, self.row, self.column, row, column)
            .map_err(|message| invalid_table_mutation("merge cells", message))?;
        staged
            .to_xml()
            .map_err(|error| invalid_table_mutation("merge cells", error.to_string()))?;
        *self.table = staged;
        Ok(())
    }

    /// Splits this merge origin back into its explicit grid cells.
    pub fn split(&mut self) -> Result<()> {
        let mut staged = self.table.clone();
        split_cell(&mut staged, self.row, self.column)
            .map_err(|message| invalid_table_mutation("split cell", message))?;
        staged
            .to_xml()
            .map_err(|error| invalid_table_mutation("split cell", error.to_string()))?;
        *self.table = staged;
        Ok(())
    }

    fn cell_ref(&self) -> TableCellRef<'_> {
        TableCellRef {
            cell: table_cell(self.table, self.row, self.column)
                .expect("TableCellMut coordinates were validated at construction"),
        }
    }

    fn cell_mut(&mut self) -> &mut CT_TableCell {
        &mut self.table.rows[self.row].cells[self.column]
    }

    fn properties_mut(&mut self) -> &mut CT_TableCellProperties {
        self.cell_mut()
            .properties
            .get_or_insert_with(CT_TableCellProperties::default)
    }
}

fn table_cell(table: &CT_Table, row: usize, column: usize) -> Option<&CT_TableCell> {
    table.rows.get(row)?.cells.get(column)
}

fn table_properties_mut(table: &mut CT_Table) -> &mut CT_TableProperties {
    table
        .properties
        .get_or_insert_with(CT_TableProperties::default)
}

/// Returns the checked sums of a table's column widths and row heights.
fn table_extent(table: &CT_Table, operation: &'static str) -> Result<(Emu, Emu)> {
    let total_width = table
        .grid
        .columns
        .iter()
        .try_fold(0i64, |total, width| total.checked_add(width.0))
        .ok_or_else(|| {
            invalid_table_mutation(operation, "table width exceeds the EMU range".to_owned())
        })?;
    let total_height = table
        .rows
        .iter()
        .try_fold(0i64, |total, row| total.checked_add(row.height.0))
        .ok_or_else(|| {
            invalid_table_mutation(operation, "table height exceeds the EMU range".to_owned())
        })?;
    if total_width <= 0 || total_height <= 0 {
        return Err(invalid_table_mutation(
            operation,
            "table width and height must remain positive".to_owned(),
        ));
    }
    Ok((Emu(total_width), Emu(total_height)))
}

/// Moves one frame edge by the change of its grid sum and requires the edge
/// to stay positive.
fn moved_frame_edge(
    frame: Emu,
    old_total: Emu,
    new_total: Emu,
    operation: &'static str,
) -> Result<Emu> {
    new_total
        .0
        .checked_sub(old_total.0)
        .and_then(|change| frame.0.checked_add(change))
        .filter(|size| *size > 0)
        .map(Emu)
        .ok_or_else(|| {
            invalid_table_mutation(
                operation,
                "table frame size must remain positive and within the EMU range".to_owned(),
            )
        })
}

fn rectangular_dimensions(table: &CT_Table) -> std::result::Result<(usize, usize), String> {
    let rows = table.rows.len();
    let columns = table.grid.columns.len();
    if rows == 0 || columns == 0 || table.rows.iter().any(|row| row.cells.len() != columns) {
        return Err("table does not have a rectangular explicit cell grid".to_owned());
    }
    Ok((rows, columns))
}

fn merge_cells(
    table: &mut CT_Table,
    first_row: usize,
    first_column: usize,
    second_row: usize,
    second_column: usize,
) -> std::result::Result<(), String> {
    let (rows, columns) = rectangular_dimensions(table)?;
    if first_row >= rows
        || second_row >= rows
        || first_column >= columns
        || second_column >= columns
    {
        return Err("merge coordinate is out of range".to_owned());
    }
    let top = first_row.min(second_row);
    let bottom = first_row.max(second_row);
    let left = first_column.min(second_column);
    let right = first_column.max(second_column);
    if top == bottom && left == right {
        return Err("merge requires at least two cells".to_owned());
    }
    if table.rows[top..=bottom].iter().any(|row| {
        row.cells[left..=right].iter().any(|cell| {
            cell.row_span != 1
                || cell.grid_span != 1
                || cell.horizontal_merge
                || cell.vertical_merge
        })
    }) {
        return Err("merge range contains one or more merged cells".to_owned());
    }
    let row_span = u32::try_from(bottom - top + 1)
        .map_err(|_| "merge row span exceeds the DrawingML range".to_owned())?;
    let grid_span = u32::try_from(right - left + 1)
        .map_err(|_| "merge column span exceeds the DrawingML range".to_owned())?;

    let mut moved_bodies = Vec::new();
    for row in top..=bottom {
        for column in left..=right {
            if row == top && column == left {
                continue;
            }
            let body = table.rows[row].cells[column]
                .text_body
                .take()
                .unwrap_or_default();
            moved_bodies.push((row, column, body));
        }
    }
    for (row, column, mut body) in moved_bodies {
        let origin_body = table.rows[top].cells[left]
            .text_body
            .get_or_insert_with(CT_TextBody::new);
        body.move_content_to(origin_body);
        table.rows[row].cells[column].text_body = Some(body);
    }

    for row in top..=bottom {
        for column in left..=right {
            let cell = &mut table.rows[row].cells[column];
            cell.row_span = if row == top { row_span } else { 1 };
            cell.grid_span = if column == left { grid_span } else { 1 };
            cell.horizontal_merge = column != left;
            cell.vertical_merge = row != top;
        }
    }
    Ok(())
}

fn split_cell(table: &mut CT_Table, row: usize, column: usize) -> std::result::Result<(), String> {
    let (rows, columns) = rectangular_dimensions(table)?;
    let origin = table_cell(table, row, column)
        .ok_or_else(|| "split coordinate is out of range".to_owned())?;
    if origin.horizontal_merge
        || origin.vertical_merge
        || (origin.row_span == 1 && origin.grid_span == 1)
    {
        return Err("only a merge-origin cell can be split".to_owned());
    }
    let bottom = row
        .checked_add(origin.row_span as usize - 1)
        .ok_or_else(|| "merge row span exceeds the table grid".to_owned())?;
    let right = column
        .checked_add(origin.grid_span as usize - 1)
        .ok_or_else(|| "merge column span exceeds the table grid".to_owned())?;
    if bottom >= rows || right >= columns {
        return Err("merge origin span exceeds the table grid".to_owned());
    }
    let row_span = origin.row_span;
    let grid_span = origin.grid_span;
    for current_row in row..=bottom {
        for current_column in column..=right {
            let cell = &table.rows[current_row].cells[current_column];
            let expected = (
                if current_row == row { row_span } else { 1 },
                if current_column == column {
                    grid_span
                } else {
                    1
                },
                current_column != column,
                current_row != row,
            );
            if (
                cell.row_span,
                cell.grid_span,
                cell.horizontal_merge,
                cell.vertical_merge,
            ) != expected
            {
                return Err("merge origin does not have a valid continuation rectangle".to_owned());
            }
        }
    }
    for current_row in row..=bottom {
        for current_column in column..=right {
            let cell = &mut table.rows[current_row].cells[current_column];
            cell.row_span = 1;
            cell.grid_span = 1;
            cell.horizontal_merge = false;
            cell.vertical_merge = false;
        }
    }
    Ok(())
}

impl<'a> TextFrame<'a> {
    /// Returns plain text in paragraph and ordered text-choice order.
    pub fn text(&self) -> String {
        self.body.plain_text()
    }

    /// Replaces the frame content with one paragraph per line, as
    /// [`CT_TextBody::set_text`] describes.
    pub fn set_text(&mut self, text: &str) {
        self.body.set_text(text);
    }

    /// Replaces literal text in this text body only and returns the count.
    ///
    /// Matching and formatting are those of
    /// [`Presentation::try_replace_text`]. An empty placeholder is rejected
    /// before anything changes. With `expect`, the replacement runs on a copy
    /// of the body, and a different count leaves the body unchanged and is
    /// returned for the caller to compare. Nothing is serialized, because
    /// replacing run text cannot make a body fail to serialize.
    pub fn try_replace_text(
        &mut self,
        placeholder: &str,
        value: &str,
        expect: Option<usize>,
    ) -> Result<usize> {
        require_replacement_placeholder(placeholder)?;
        let value = &oxml_drawing::text::escape_invalid_xml_characters(value);
        let Some(expected) = expect else {
            return Ok(replace_text_in_body(self.body, placeholder, value));
        };
        let mut candidate = self.body.clone();
        let count = replace_text_in_body(&mut candidate, placeholder, value);
        if count == expected {
            *self.body = candidate;
        }
        Ok(count)
    }

    /// Returns the number of paragraphs in the text body.
    pub fn paragraph_count(&self) -> usize {
        self.body.paragraph_count()
    }

    /// Returns one paragraph without mutating the text body.
    pub fn paragraph(&self, index: usize) -> Option<TextParagraphRef<'_>> {
        self.body
            .paragraphs()
            .get(index)
            .map(|paragraph| TextParagraphRef { paragraph })
    }

    /// Returns one paragraph for in-place mutation.
    pub fn paragraph_mut(&mut self, index: usize) -> Option<TextParagraphMut<'_>> {
        self.body
            .paragraph_mut(index)
            .map(|paragraph| TextParagraphMut { paragraph })
    }

    /// Consumes this handle and returns one paragraph for mutation.
    pub fn into_paragraph_mut(self, index: usize) -> Option<TextParagraphMut<'a>> {
        self.body
            .paragraph_mut(index)
            .map(|paragraph| TextParagraphMut { paragraph })
    }

    /// Appends one empty paragraph.
    pub fn add_paragraph(&mut self) -> TextParagraphMut<'_> {
        TextParagraphMut {
            paragraph: self.body.add_paragraph(),
        }
    }

    /// Replaces all four optional text insets.
    ///
    /// An inset outside the 32-bit `ST_Coordinate32` range is rejected before
    /// any inset changes.
    pub fn set_insets(
        &mut self,
        left: Option<Emu>,
        right: Option<Emu>,
        top: Option<Emu>,
        bottom: Option<Emu>,
    ) -> Result<()> {
        let coordinate = |value: Option<Emu>| {
            value
                .map(|value| {
                    i32::try_from(value.0)
                        .map(Coordinate32Value::Emu)
                        .map_err(|_| Error::InvalidShapeMutation {
                            operation: "set text insets",
                            message: format!("{} EMU is outside the coordinate range", value.0),
                        })
                })
                .transpose()
        };
        let (left, right, top, bottom) = (
            coordinate(left)?,
            coordinate(right)?,
            coordinate(top)?,
            coordinate(bottom)?,
        );
        let properties = &mut self.body.body_properties;
        properties.left_inset = left;
        properties.right_inset = right;
        properties.top_inset = top;
        properties.bottom_inset = bottom;
        Ok(())
    }

    /// Sets or clears the direct vertical anchor.
    pub fn set_vertical_anchor(&mut self, anchor: Option<TextAnchor>) {
        self.body.body_properties.anchor = anchor;
    }

    /// Sets or clears the direct wrap choice, `true` wrapping at the frame width.
    pub fn set_word_wrap(&mut self, wrap: Option<bool>) {
        self.body.body_properties.wrap = wrap.map(|wrap| {
            if wrap {
                TextWrap::Square
            } else {
                TextWrap::None
            }
        });
    }

    /// Sets or clears the direct autofit choice.
    ///
    /// Choosing normal autofit again keeps a stored font scale and line
    /// spacing reduction.
    pub fn set_autofit_mode(&mut self, mode: Option<AutofitMode>) {
        let properties = &mut self.body.body_properties;
        properties.autofit = match mode {
            None => None,
            Some(AutofitMode::None) => Some(TextAutofit::NoAutofit),
            Some(AutofitMode::Shape) => Some(TextAutofit::ShapeAutofit),
            Some(AutofitMode::Normal) => match properties.autofit.take() {
                Some(TextAutofit::Normal(normal)) => Some(TextAutofit::Normal(normal)),
                _ => Some(TextAutofit::Normal(NormalAutofit::default())),
            },
        };
    }
}

impl<'a> TextFrameRef<'a> {
    /// Returns plain text in paragraph and ordered text-choice order.
    pub fn text(&self) -> String {
        self.body.plain_text()
    }

    /// Returns the number of paragraphs in the text body.
    pub fn paragraph_count(&self) -> usize {
        self.body.paragraph_count()
    }

    /// Returns one paragraph by zero-based index.
    pub fn paragraph(&self, index: usize) -> Option<TextParagraphRef<'a>> {
        self.body
            .paragraphs()
            .get(index)
            .map(|paragraph| TextParagraphRef { paragraph })
    }

    /// Returns the direct autofit choice on the text body.
    pub fn autofit_mode(&self) -> Option<AutofitMode> {
        self.body
            .body_properties
            .autofit
            .as_ref()
            .map(|autofit| match autofit {
                TextAutofit::NoAutofit => AutofitMode::None,
                TextAutofit::Normal(_) => AutofitMode::Normal,
                TextAutofit::ShapeAutofit => AutofitMode::Shape,
            })
    }

    /// Returns the optional left, right, top, and bottom text insets in EMU.
    ///
    /// A universal measure such as `0.1in` is converted to the nearest EMU.
    pub fn insets(&self) -> (Option<Emu>, Option<Emu>, Option<Emu>, Option<Emu>) {
        let properties = &self.body.body_properties;
        (
            properties.left_inset.as_ref().and_then(coordinate_emu),
            properties.right_inset.as_ref().and_then(coordinate_emu),
            properties.top_inset.as_ref().and_then(coordinate_emu),
            properties.bottom_inset.as_ref().and_then(coordinate_emu),
        )
    }

    /// Returns the direct vertical anchor.
    pub fn vertical_anchor(&self) -> Option<TextAnchor> {
        self.body.body_properties.anchor
    }

    /// Returns the direct wrap choice, `true` when lines wrap at the frame width.
    pub fn word_wrap(&self) -> Option<bool> {
        self.body
            .body_properties
            .wrap
            .map(|wrap| wrap == TextWrap::Square)
    }
}

fn coordinate_emu(value: &Coordinate32Value) -> Option<Emu> {
    match value {
        Coordinate32Value::Emu(value) => Some(Emu(i64::from(*value))),
        Coordinate32Value::UniversalMeasure(value) => {
            let (number, emu_per_unit) = [
                ("mm", 36_000.0),
                ("cm", 360_000.0),
                ("in", 914_400.0),
                ("pt", 12_700.0),
                ("pc", 152_400.0),
                ("pi", 152_400.0),
            ]
            .into_iter()
            .find_map(|(unit, emu_per_unit)| {
                value
                    .strip_suffix(unit)
                    .map(|number| (number, emu_per_unit))
            })?;
            let number = number.parse::<f64>().ok()?;
            Some(Emu((number * emu_per_unit).round() as i64))
        }
    }
}

impl<'a> TextParagraphRef<'a> {
    /// Returns visible text in ordered run, break, and field order.
    pub fn text(&self) -> String {
        self.paragraph
            .runs
            .iter()
            .map(|run| match run {
                TextRun::Run(run) => run.text.value.as_str(),
                TextRun::Break(_) => "\u{000b}",
                TextRun::Field(field) => field.text.as_ref().map_or("", |text| text.value.as_str()),
            })
            .collect()
    }

    /// Returns the direct paragraph level, defaulting to zero.
    pub fn level(&self) -> u8 {
        self.paragraph
            .properties
            .as_ref()
            .and_then(|properties| properties.level)
            .unwrap_or(0)
    }

    /// Returns direct default-run character properties when present.
    pub fn default_run_properties(&self) -> Option<&'a CT_TextCharacterProperties> {
        self.paragraph
            .properties
            .as_ref()
            .and_then(|properties| properties.default_run_properties.as_ref())
    }

    /// Returns direct paragraph properties when present.
    pub fn properties(&self) -> Option<&'a CT_TextParagraphProperties> {
        self.paragraph.properties.as_ref()
    }

    /// Returns the number of regular runs, excluding breaks and fields.
    pub fn run_count(&self) -> usize {
        self.paragraph
            .runs
            .iter()
            .filter(|run| matches!(run, TextRun::Run(_)))
            .count()
    }

    /// Returns one regular run by zero-based regular-run index.
    pub fn run(&self, index: usize) -> Option<TextRunRef<'a>> {
        self.paragraph
            .runs
            .iter()
            .filter_map(|run| match run {
                TextRun::Run(run) => Some(TextRunRef { run }),
                TextRun::Break(_) | TextRun::Field(_) => None,
            })
            .nth(index)
    }
}

impl TextRunRef<'_> {
    /// Returns the run text.
    pub fn text(&self) -> &str {
        &self.run.text.value
    }

    /// Returns direct character properties when present.
    pub fn properties(&self) -> Option<&CT_TextCharacterProperties> {
        self.run.properties.as_ref()
    }

    /// Returns the direct Latin typeface on this run.
    pub fn font_name(&self) -> Option<&str> {
        self.properties()?
            .latin
            .as_ref()
            .map(|font| font.typeface.as_str())
    }

    /// Returns the direct run size in hundredths of a point.
    pub fn font_size(&self) -> Option<i32> {
        self.properties()?.font_size
    }

    /// Returns the direct sRGB run colour as `RRGGBB`.
    pub fn font_color(&self) -> Option<String> {
        let Fill::Solid(fill) = self.properties()?.fill.as_ref()? else {
            return None;
        };
        let ColorChoice::Srgb { value, .. } = fill.color.as_ref()? else {
            return None;
        };
        Some(value.to_string())
    }
}

/// A mutably borrowed DrawingML text paragraph.
pub struct TextParagraphMut<'a> {
    paragraph: &'a mut CT_TextParagraph,
}

impl TextParagraphMut<'_> {
    /// Replaces fields, breaks, and runs with one regular run, split by an
    /// `a:br` at each line feed, CRLF pair or vertical tab, as python-pptx
    /// assigns paragraph text.
    pub fn set_text(&mut self, text: &str) {
        self.paragraph
            .set_text(&text.replace("\r\n", "\u{b}").replace('\n', "\u{b}"));
    }

    /// Appends a regular run after the existing ordered text choices.
    pub fn add_run(&mut self, text: &str) -> TextRunMut<'_> {
        TextRunMut {
            run: self.paragraph.add_run(text),
        }
    }

    /// Returns one regular run for in-place mutation.
    pub fn run_mut(&mut self, index: usize) -> Option<TextRunMut<'_>> {
        self.paragraph
            .runs
            .iter_mut()
            .filter_map(|run| match run {
                TextRun::Run(run) => Some(TextRunMut { run }),
                TextRun::Break(_) | TextRun::Field(_) => None,
            })
            .nth(index)
    }

    /// Sets the direct paragraph level.
    pub fn set_level(&mut self, level: u8) -> bool {
        if level > 8 {
            return false;
        }
        self.paragraph.properties_mut().level = Some(level);
        true
    }

    /// Returns mutable direct default-run character properties.
    pub fn default_run_properties_mut(&mut self) -> &mut CT_TextCharacterProperties {
        self.paragraph
            .properties_mut()
            .default_run_properties
            .get_or_insert_with(CT_TextCharacterProperties::default)
    }

    /// Returns direct paragraph properties when present.
    pub fn properties(&self) -> Option<&CT_TextParagraphProperties> {
        self.paragraph.properties.as_ref()
    }

    /// Replaces the paragraph's direct typed properties.
    pub fn set_properties(&mut self, properties: CT_TextParagraphProperties) {
        *self.paragraph.properties_mut() = properties;
    }

    /// Sets or clears the direct paragraph bullet.
    ///
    /// A preserved picture or follow-text bullet part that the new value
    /// replaces is removed, so the paragraph keeps one bullet choice.
    pub fn set_bullet(&mut self, bullet: Option<TextBullet>) {
        if let Some(properties) = self.paragraph.properties.as_mut() {
            properties.set_bullet(bullet);
        } else if bullet.is_some() {
            self.paragraph.properties_mut().set_bullet(bullet);
        }
    }
}

/// A mutably borrowed regular DrawingML text run.
pub struct TextRunMut<'a> {
    run: &'a mut CT_RegularTextRun,
}

impl TextRunMut<'_> {
    /// Returns direct character properties when present.
    pub fn properties(&self) -> Option<&CT_TextCharacterProperties> {
        self.run.properties.as_ref()
    }

    /// Replaces text while retaining the run's typed and unmodelled state.
    pub fn set_text(&mut self, text: &str) {
        self.run.set_text(text);
    }

    /// Replaces the run's direct typed character properties.
    pub fn set_properties(&mut self, properties: CT_TextCharacterProperties) {
        self.run.properties = Some(properties);
    }

    /// Sets or clears the direct Latin typeface.
    pub fn set_font(&mut self, font: Option<TextFont>) {
        if let Some(properties) = self.run.properties.as_mut() {
            properties.latin = font;
        } else if font.is_some() {
            let mut properties = CT_TextCharacterProperties::default();
            properties.latin = font;
            self.run.properties = Some(properties);
        }
    }
}

fn shape_kind(child: &ShapeTreeChild) -> ShapeKind {
    match child {
        ShapeTreeChild::Shape(_) => ShapeKind::Shape,
        ShapeTreeChild::Picture(_) => ShapeKind::Picture,
        ShapeTreeChild::GraphicFrame(_) => ShapeKind::GraphicFrame,
        ShapeTreeChild::GroupShape(_) => ShapeKind::Group,
        ShapeTreeChild::Connector(_) => ShapeKind::Connector,
        ShapeTreeChild::AlternateContent(_) => ShapeKind::AlternateContent,
    }
}

fn shape_transform(child: &ShapeTreeChild) -> Option<&CT_Transform2D> {
    match child {
        ShapeTreeChild::Shape(shape) => shape.shape_properties.transform.as_ref(),
        ShapeTreeChild::Picture(picture) => picture.shape_properties.transform.as_ref(),
        ShapeTreeChild::GraphicFrame(frame) => Some(&frame.transform),
        ShapeTreeChild::GroupShape(group) => group.group_transform(),
        ShapeTreeChild::Connector(connector) => connector.shape_properties.transform.as_ref(),
        ShapeTreeChild::AlternateContent(alternate) => Some(&alternate.chart_choice()?.transform),
    }
}

fn literal_guide_value(guide: &Guide) -> Option<f64> {
    match (&guide.op, guide.args.as_slice()) {
        (GuideOp::Val, [GuideOperand::Literal(value)]) => Some(*value),
        _ => None,
    }
}

fn shape_properties(child: &ShapeTreeChild) -> Option<&CT_ShapeProperties> {
    match child {
        ShapeTreeChild::Shape(shape) => Some(&shape.shape_properties),
        ShapeTreeChild::Picture(picture) => Some(&picture.shape_properties),
        ShapeTreeChild::Connector(connector) => Some(&connector.shape_properties),
        ShapeTreeChild::GraphicFrame(_)
        | ShapeTreeChild::GroupShape(_)
        | ShapeTreeChild::AlternateContent(_) => None,
    }
}

impl<'a> ShapeRef<'a> {
    /// Returns the child's normalized structural kind.
    pub fn kind(&self) -> ShapeKind {
        shape_kind(self.child)
    }

    /// Classifies the child the way python-pptx `shape_type` does.
    ///
    /// Only an ordinary shape reports `Placeholder`: python-pptx classifies a
    /// picture placeholder as a picture and a graphic-frame placeholder by its
    /// payload. Only video counts as `Media`, so an audio picture is a
    /// `Picture`. `None` covers a shape without geometry, SmartArt and other
    /// graphic payloads, and alternate content without a chart choice.
    pub fn shape_type(&self) -> Option<ShapeType> {
        match self.child {
            ShapeTreeChild::Shape(shape) => {
                if shape.placeholder.is_some() {
                    Some(ShapeType::Placeholder)
                } else if shape.shape_properties.custom_geometry.is_some() {
                    Some(ShapeType::Freeform)
                } else if shape.is_textbox() {
                    Some(ShapeType::TextBox)
                } else if shape.shape_properties.preset_geometry.is_some() {
                    Some(ShapeType::AutoShape)
                } else {
                    None
                }
            }
            ShapeTreeChild::Picture(picture) => Some(
                if picture
                    .media
                    .as_ref()
                    .is_some_and(|media| media.kind == MediaKind::Video)
                {
                    ShapeType::Media
                } else {
                    ShapeType::Picture
                },
            ),
            ShapeTreeChild::GraphicFrame(frame) => match frame.graphic_data.payload() {
                GraphicDataPayload::Table(_) => Some(ShapeType::Table),
                GraphicDataPayload::Chart(_) => Some(ShapeType::Chart),
                GraphicDataPayload::Ole { .. } => {
                    if frame.graphic_data.is_embedded_ole_object() == Some(true) {
                        Some(ShapeType::EmbeddedOleObject)
                    } else {
                        Some(ShapeType::LinkedOleObject)
                    }
                }
                GraphicDataPayload::SmartArt(_) | GraphicDataPayload::Other(_) => None,
            },
            ShapeTreeChild::GroupShape(_) => Some(ShapeType::Group),
            ShapeTreeChild::Connector(_) => Some(ShapeType::Line),
            ShapeTreeChild::AlternateContent(alternate) => {
                alternate.chart_choice().map(|_| ShapeType::Chart)
            }
        }
    }

    /// Returns the effective preset adjustments of an ordinary shape.
    ///
    /// Like python-pptx, the preset definition's defaults come in definition
    /// order and a literal `val` guide in the shape's own `a:avLst` replaces
    /// the default of the same name. Values are raw guide values, where
    /// 100000 means 1.0. Other shape kinds and custom geometry have none.
    pub fn adjustments(&self) -> Result<Vec<(String, f64)>> {
        let ShapeTreeChild::Shape(shape) = self.child else {
            return Ok(Vec::new());
        };
        let Some(geometry) = &shape.shape_properties.preset_geometry else {
            return Ok(Vec::new());
        };
        let defaults = geometry
            .default_adjust_values()
            .map_err(|error| invalid_shape_mutation("read adjustments", error.to_string()))?;
        Ok(defaults
            .into_iter()
            .filter_map(|default| {
                let value = literal_guide_value(
                    geometry
                        .adjust_values()
                        .iter()
                        .find(|guide| guide.name == default.name)
                        .unwrap_or(&default),
                )
                .or_else(|| literal_guide_value(&default))?;
                Some((default.name, value))
            })
            .collect())
    }

    /// Returns the preset geometry name of an auto shape or a picture mask.
    ///
    /// Like python-pptx, an ordinary shape is an auto shape when it carries
    /// `a:prstGeom` and is not a text box, and a picture reports the preset
    /// that masks it. Custom geometry, inherited geometry and every other
    /// shape kind report `None`.
    pub fn auto_shape_type(&self) -> Option<&'a str> {
        let properties = match self.child {
            ShapeTreeChild::Shape(shape) if !shape.is_textbox() => &shape.shape_properties,
            ShapeTreeChild::Picture(picture) => &picture.shape_properties,
            _ => return None,
        };
        properties
            .preset_geometry
            .as_ref()
            .map(|geometry| geometry.preset.as_str())
    }

    /// Returns the direct clockwise rotation, or `None` without a transform.
    pub fn rotation(&self) -> Option<Angle> {
        shape_transform(self.child).map(|transform| transform.rotation)
    }

    /// Returns the direct fill of a shape, picture, or connector.
    pub fn fill(&self) -> Option<&'a Fill> {
        shape_properties(self.child)?.fill.as_ref()
    }

    /// Returns the direct line of a shape, picture, or connector.
    pub fn line(&self) -> Option<&'a CT_LineProperties> {
        shape_properties(self.child)?.line.as_ref()
    }

    /// Returns a picture's `a:srcRect` crop as left, top, right, and bottom
    /// insets of its image, reading an absent edge as zero.
    ///
    /// Other shape kinds return `None`.
    pub fn crop(&self) -> Option<(Percent1000, Percent1000, Percent1000, Percent1000)> {
        let ShapeTreeChild::Picture(picture) = self.child else {
            return None;
        };
        let rect = picture
            .blip_fill
            .as_ref()
            .and_then(|fill| fill.source_rect.as_ref());
        Some(rect.map_or_else(Default::default, |rect| {
            (
                rect.left.unwrap_or_default(),
                rect.top.unwrap_or_default(),
                rect.right.unwrap_or_default(),
                rect.bottom.unwrap_or_default(),
            )
        }))
    }

    /// Returns the direct `a:effectLst` of a shape, picture, connector, or
    /// group. `None` means the shape inherits its theme effect.
    pub fn effects(&self) -> Option<&'a CT_EffectList> {
        match self.child {
            ShapeTreeChild::GroupShape(group) => group.effects(),
            child => shape_properties(child)?.effects.as_ref(),
        }
    }

    /// Returns the theme effect style index of an ordinary shape's or a
    /// connector's `p:style`, or `None` without a typed style.
    ///
    /// Index 0 references no theme effect.
    pub fn theme_effect_index(&self) -> Option<u32> {
        let style = match self.child {
            ShapeTreeChild::Shape(shape) => shape.style(),
            ShapeTreeChild::Connector(connector) => connector.style(),
            _ => None,
        }?;
        Some(style.effect_reference.index)
    }

    /// Serialises the child as a self-contained element.
    ///
    /// Typed children declare the `p`, `a`, `r` and `mc` prefixes they use.
    /// Alternate content is returned verbatim and may rely on prefixes that
    /// only the slide root declares.
    pub fn xml(&self) -> Result<Vec<u8>> {
        match self.child {
            ShapeTreeChild::Shape(shape) => shape.to_xml(),
            ShapeTreeChild::Picture(picture) => picture.to_xml(),
            ShapeTreeChild::GraphicFrame(frame) => frame.to_xml(),
            ShapeTreeChild::GroupShape(group) => group.to_xml(),
            ShapeTreeChild::Connector(connector) => connector.to_xml(),
            ShapeTreeChild::AlternateContent(alternate) => Ok(alternate.raw_xml().to_vec()),
        }
        .map_err(|error| Error::InvalidShapeMutation {
            operation: "serialize shape",
            message: error.to_string(),
        })
    }

    /// Returns the direct shape offset in EMU.
    pub fn position(&self) -> Option<(Emu, Emu)> {
        let offset = shape_transform(self.child)?.offset?;
        Some((offset.x, offset.y))
    }

    /// Returns the direct shape extent in EMU.
    pub fn size(&self) -> Option<(Emu, Emu)> {
        let extent = shape_transform(self.child)?.extent?;
        Some((extent.cx, extent.cy))
    }

    /// Returns the alternative text, `p:cNvPr/@descr`.
    pub fn alt_text(&self) -> Option<String> {
        self.child.alt_text()
    }

    /// Returns the alternative-text title, `p:cNvPr/@title`.
    pub fn alt_title(&self) -> Option<String> {
        self.child.alt_title()
    }

    /// Returns whether the shape is marked decorative.
    pub fn decorative(&self) -> bool {
        self.child.decorative()
    }

    /// Returns whether the shape is mirrored left to right.
    pub fn flip_horizontal(&self) -> bool {
        shape_transform(self.child).is_some_and(|transform| transform.flip_horizontal)
    }

    /// Returns whether the shape is mirrored top to bottom.
    pub fn flip_vertical(&self) -> bool {
        shape_transform(self.child).is_some_and(|transform| transform.flip_vertical)
    }

    /// Returns a connector's begin and end points in its parent's
    /// coordinates, flips and rotation applied, as python-pptx `begin_x`,
    /// `begin_y`, `end_x`, and `end_y` read them. Other shapes and a
    /// connector without a position and size return `None`.
    pub fn connector_endpoints(&self) -> Option<[(Emu, Emu); 2]> {
        let ShapeTreeChild::Connector(connector) = self.child else {
            return None;
        };
        let points = arrange::connector_endpoints(connector.shape_properties.transform.as_ref()?)?;
        Some(points.map(|(x, y)| (Emu(x.round() as i64), Emu(y.round() as i64))))
    }

    /// Returns the shape id and connection site one end of a connector is
    /// glued to, `a:stCxn` or `a:endCxn`.
    pub fn connector_connection(&self, end: ConnectorEnd) -> Option<(u32, u32)> {
        let ShapeTreeChild::Connector(connector) = self.child else {
            return None;
        };
        match end {
            ConnectorEnd::Begin => connector.start_connection.as_ref(),
            ConnectorEnd::End => connector.end_connection.as_ref(),
        }
        .map(|connection| (connection.id, connection.idx))
    }

    /// Returns the producer-facing non-visual shape id.
    pub fn non_visual_id(&self) -> Option<u32> {
        self.child.non_visual_id()
    }

    /// Returns the producer-facing non-visual shape name.
    pub fn non_visual_name(&self) -> Option<String> {
        self.child.non_visual_name()
    }

    /// Returns ordinary shape text or row-major table text when modelled.
    pub fn text(&self) -> Option<String> {
        match self.child {
            ShapeTreeChild::Shape(shape) => shape.text_body.as_ref().map(|text| text.plain_text()),
            ShapeTreeChild::GraphicFrame(frame) => match frame.graphic_data.payload() {
                GraphicDataPayload::Table(table) => Some(
                    table
                        .rows
                        .iter()
                        .map(|row| {
                            row.cells
                                .iter()
                                .map(|cell| {
                                    cell.text_body
                                        .as_ref()
                                        .map_or_else(String::new, |text| text.plain_text())
                                })
                                .collect::<Vec<_>>()
                                .join("\t")
                        })
                        .collect::<Vec<_>>()
                        .join("\n"),
                ),
                _ => None,
            },
            _ => None,
        }
    }

    /// Returns the ordinary shape text body without mutating it.
    pub fn text_frame(&self) -> Option<TextFrameRef<'a>> {
        match self.child {
            ShapeTreeChild::Shape(shape) => {
                shape.text_body.as_ref().map(|body| TextFrameRef { body })
            }
            _ => None,
        }
    }

    /// Returns the OOXML placeholder index when this child is a placeholder.
    pub fn placeholder_idx(&self) -> Option<u32> {
        shape_placeholder(self.child).map(|placeholder| placeholder.idx.unwrap_or(0))
    }

    /// Returns the explicit placeholder type token, such as `title`.
    ///
    /// A placeholder that omits its type returns `None`, as does a child that
    /// is not a placeholder.
    pub fn placeholder_type(&self) -> Option<&'a str> {
        shape_placeholder(self.child)?
            .ph_type
            .as_ref()
            .map(PhType::as_str)
    }

    /// Returns the typed table carried by this graphic frame, when present.
    pub fn table(&self) -> Option<TableRef<'a>> {
        let ShapeTreeChild::GraphicFrame(frame) = self.child else {
            return None;
        };
        let GraphicDataPayload::Table(table) = frame.graphic_data.payload() else {
            return None;
        };
        Some(TableRef { table })
    }

    /// Returns the number of immediate group or selected fallback children.
    pub fn child_count(&self) -> usize {
        self.child_slice().len()
    }

    /// Returns one immediate group or selected fallback child.
    pub fn child(&self, index: usize) -> Option<ShapeRef<'a>> {
        self.child_slice().get(index).map(shape_ref)
    }

    /// Iterates immediate group or selected fallback children.
    pub fn children(&self) -> impl ExactSizeIterator<Item = ShapeRef<'a>> + 'a {
        self.child_slice().iter().map(shape_ref)
    }

    fn child_slice(&self) -> &'a [ShapeTreeChild] {
        match self.child {
            ShapeTreeChild::GroupShape(group) => &group.children,
            ShapeTreeChild::AlternateContent(alternate) => {
                alternate.selected_fallback().unwrap_or_default()
            }
            _ => &[],
        }
    }
}

fn shape_placeholder(child: &ShapeTreeChild) -> Option<&CT_Placeholder> {
    match child {
        ShapeTreeChild::Shape(shape) => shape.placeholder.as_ref(),
        ShapeTreeChild::Picture(picture) => picture.placeholder.as_ref(),
        _ => None,
    }
}

fn collect_text(children: &[ShapeTreeChild], output: &mut Vec<String>) {
    for child in children {
        let shape = shape_ref(child);
        if let Some(text) = shape.text()
            && !text.is_empty()
        {
            output.push(text);
        }
        collect_text(shape.child_slice(), output);
    }
}

fn first_notes_body_mut(children: &mut [ShapeTreeChild]) -> Option<&mut CT_TextBody> {
    for child in children {
        match child {
            ShapeTreeChild::Shape(shape)
                if shape
                    .placeholder
                    .as_ref()
                    .is_some_and(|placeholder| placeholder.effective_type() == PhType::Body) =>
            {
                return Some(shape.text_body.get_or_insert_with(CT_TextBody::new));
            }
            ShapeTreeChild::GroupShape(group) => {
                if let Some(body) = first_notes_body_mut(&mut group.children) {
                    return Some(body);
                }
            }
            ShapeTreeChild::Shape(_)
            | ShapeTreeChild::Picture(_)
            | ShapeTreeChild::GraphicFrame(_)
            | ShapeTreeChild::Connector(_)
            | ShapeTreeChild::AlternateContent(_) => {}
        }
    }
    None
}

fn require_replacement_placeholder(placeholder: &str) -> Result<()> {
    if placeholder.is_empty() {
        return Err(Error::InvalidSlideMutation {
            operation: "replace text",
            message: "placeholder must not be empty".to_owned(),
        });
    }
    Ok(())
}

fn replace_text_in_slide_record(
    record: &mut SlideRecord,
    placeholder: &str,
    value: &str,
    notes: bool,
) -> usize {
    let slide_count = replace_text_in_children(
        &mut record.slide.common_slide_data.shape_tree.children,
        placeholder,
        value,
    );
    let notes_count = record.notes.as_mut().filter(|_| notes).map_or(0, |notes| {
        replace_text_in_children(
            &mut notes.notes.common_slide_data.shape_tree.children,
            placeholder,
            value,
        )
    });
    slide_count + notes_count
}

fn replace_text_in_children(
    children: &mut [ShapeTreeChild],
    placeholder: &str,
    value: &str,
) -> usize {
    let mut count = 0;
    for child in children {
        match child {
            ShapeTreeChild::Shape(shape) => {
                if let Some(body) = &mut shape.text_body {
                    count += replace_text_in_body(body, placeholder, value);
                }
            }
            ShapeTreeChild::GraphicFrame(frame) => {
                if let Some(table) = frame.graphic_data.table_mut() {
                    for row in &mut table.rows {
                        for cell in &mut row.cells {
                            if let Some(body) = &mut cell.text_body {
                                count += replace_text_in_body(body, placeholder, value);
                            }
                        }
                    }
                }
            }
            ShapeTreeChild::GroupShape(group) => {
                count += replace_text_in_children(&mut group.children, placeholder, value);
            }
            ShapeTreeChild::Picture(_)
            | ShapeTreeChild::Connector(_)
            | ShapeTreeChild::AlternateContent(_) => {}
        }
    }
    count
}

fn replace_text_in_body(body: &mut CT_TextBody, placeholder: &str, value: &str) -> usize {
    let mut count = 0;
    for index in 0..body.paragraph_count() {
        if let Some(paragraph) = body.paragraph_mut(index) {
            count += replace_text_in_paragraph(paragraph, placeholder, value);
        }
    }
    count
}

fn replace_text_in_paragraph(
    paragraph: &mut CT_TextParagraph,
    placeholder: &str,
    value: &str,
) -> usize {
    let mut count = 0;
    let mut segment_start = 0;
    while segment_start < paragraph.runs.len() {
        while segment_start < paragraph.runs.len()
            && !matches!(paragraph.runs[segment_start], TextRun::Run(_))
        {
            segment_start += 1;
        }
        let mut segment_end = segment_start;
        while segment_end < paragraph.runs.len()
            && matches!(paragraph.runs[segment_end], TextRun::Run(_))
        {
            segment_end += 1;
        }
        count += replace_text_in_run_segment(
            &mut paragraph.runs[segment_start..segment_end],
            placeholder,
            value,
        );
        segment_start = segment_end.saturating_add(1);
    }
    count
}

fn replace_text_in_run_segment(runs: &mut [TextRun], placeholder: &str, value: &str) -> usize {
    if placeholder.is_empty() {
        return 0;
    }
    let mut text = String::new();
    let mut ranges = Vec::with_capacity(runs.len());
    for run in runs.iter() {
        let TextRun::Run(run) = run else {
            unreachable!("run segments contain only regular runs")
        };
        let start = text.len();
        text.push_str(&run.text.value);
        ranges.push((start, text.len()));
    }
    let matches = text
        .match_indices(placeholder)
        .map(|(start, _)| (start, start + placeholder.len()))
        .collect::<Vec<_>>();

    for &(match_start, match_end) in matches.iter().rev() {
        let start_run = ranges
            .iter()
            .position(|&(_, end)| match_start < end)
            .expect("non-empty match has a starting run");
        let end_run = ranges
            .iter()
            .position(|&(start, end)| match_end > start && match_end <= end)
            .expect("non-empty match has an ending run");
        let start_offset = match_start - ranges[start_run].0;
        let end_offset = match_end - ranges[end_run].0;

        if start_run == end_run {
            let TextRun::Run(run) = &mut runs[start_run] else {
                unreachable!("run segments contain only regular runs")
            };
            run.text
                .value
                .replace_range(start_offset..end_offset, value);
            continue;
        }

        let TextRun::Run(first) = &mut runs[start_run] else {
            unreachable!("run segments contain only regular runs")
        };
        let first_len = first.text.value.len();
        first
            .text
            .value
            .replace_range(start_offset..first_len, value);
        for run in &mut runs[start_run + 1..end_run] {
            let TextRun::Run(run) = run else {
                unreachable!("run segments contain only regular runs")
            };
            run.set_text("");
        }
        let TextRun::Run(last) = &mut runs[end_run] else {
            unreachable!("run segments contain only regular runs")
        };
        last.text.value.replace_range(..end_offset, "");
    }
    matches.len()
}

#[cfg(feature = "render")]
fn assemble_render_input(package: &OpcPackage) -> Result<(RenderInput, LayoutResult)> {
    let assembly = prepare_render_context(package, false)?;
    Ok((assembly.input, assembly.layout))
}

#[cfg(feature = "render")]
const MAX_PRESENTATION_EXPORT_DPI: f64 = 600.0;
#[cfg(feature = "render")]
const MAX_PRESENTATION_EXPORT_RASTER_BYTES: u64 = 256 * 1024 * 1024;

#[cfg(feature = "render")]
fn render_export_pngs(layout: &LayoutResult, dpi: f64) -> Result<Vec<Vec<u8>>> {
    validate_export_png_pages(&layout.pages, dpi)?;
    (0..layout.pages.len())
        .map(|index| {
            oxml_pdf::render_page_to_png(layout, index, dpi).ok_or_else(|| {
                render_failure(format!("could not rasterize export page {}", index + 1))
            })
        })
        .collect()
}

#[cfg(feature = "render")]
fn render_export_png(layout: &LayoutResult, index: usize, dpi: f64) -> Result<Option<Vec<u8>>> {
    let Some(page) = layout.pages.get(index) else {
        validate_export_png_pages(&[], dpi)?;
        return Ok(None);
    };
    validate_export_png_pages(std::slice::from_ref(page), dpi)?;
    oxml_pdf::render_page_to_png(layout, index, dpi)
        .map(Some)
        .ok_or_else(|| render_failure(format!("could not rasterize export page {}", index + 1)))
}

#[cfg(feature = "render")]
fn validate_export_png_pages(pages: &[std::sync::Arc<PageFrame>], dpi: f64) -> Result<()> {
    if !dpi.is_finite() || dpi <= 0.0 || dpi > MAX_PRESENTATION_EXPORT_DPI {
        return Err(render_failure(format!(
            "raster DPI must be finite and in 0..={MAX_PRESENTATION_EXPORT_DPI}, found {dpi}"
        )));
    }
    let scale = dpi / 72.0;
    let estimated = pages.iter().try_fold(0u64, |total, page| {
        let width = (page.width * scale).ceil();
        let height = (page.height * scale).ceil();
        if !width.is_finite() || !height.is_finite() || width < 1.0 || height < 1.0 {
            return None;
        }
        let page_bytes = (width as u64).checked_mul(height as u64)?.checked_mul(4)?;
        total.checked_add(page_bytes)
    });
    if estimated.is_none_or(|bytes| bytes > MAX_PRESENTATION_EXPORT_RASTER_BYTES) {
        return Err(render_failure(format!(
            "raster export exceeds the {} byte decoded-output limit",
            MAX_PRESENTATION_EXPORT_RASTER_BYTES
        )));
    }
    Ok(())
}

#[cfg(feature = "render")]
fn render_notes_pages(package: &OpcPackage) -> Result<LayoutResult> {
    let mut assembly = prepare_render_context(package, false)?;
    if assembly.slides.is_empty() {
        return Ok(LayoutResult::new(
            Vec::new(),
            assembly.font_manager.all_font_data(),
            assembly.input.metadata.clone(),
            Vec::new(),
        ));
    }
    let presentation_part = package
        .main_document_part()
        .ok_or(Error::MissingMainDocument)?;
    let presentation = CT_Presentation::from_xml(required_part(package, &presentation_part)?)
        .map_err(|error| Error::MalformedPart {
            part_name: presentation_part.clone(),
            message: error.to_string(),
        })?;
    let notes_size = export_page_size(&presentation);
    let notes_master_part = render_exact_related_part(
        package,
        &presentation_part,
        rel_types::NOTES_MASTER,
        content_types::NOTES_MASTER,
    )?;
    let notes_master = CT_NotesMaster::from_xml(required_part(package, &notes_master_part)?)
        .map_err(|error| Error::MalformedPart {
            part_name: notes_master_part.clone(),
            message: error.to_string(),
        })?;
    let theme_part = render_exact_related_part(
        package,
        &notes_master_part,
        rel_types::THEME,
        content_types::THEME,
    )?;
    let theme =
        CT_OfficeStyleSheet::from_xml(required_part(package, &theme_part)?).map_err(|error| {
            Error::MalformedPart {
                part_name: theme_part,
                message: error.to_string(),
            }
        })?;
    let default_text_style = notes_master
        .notes_style
        .clone()
        .or(presentation.default_text_style.clone())
        .unwrap_or_default();
    let shell_layout = assembly.slides[0].layout.clone();
    let shell_master = assembly.slides[0].master.clone();
    let mut export_package = package.clone();
    let mut pages = Vec::with_capacity(assembly.slides.len());
    let mut diagnostics = assembly.layout.diagnostics.clone();
    for (slide_index, prepared) in assembly.slides.iter().enumerate() {
        let notes = resolve_notes(package, &prepared.slide_part)?;
        if let Some(notes) = &notes {
            validate_notes_export_graph(
                package,
                &prepared.slide_part,
                &notes.part_name,
                &notes_master_part,
            )?;
        }
        let (page_source, page_master, page_notes) = remap_notes_export_scope(
            &mut export_package,
            &notes_master_part,
            &notes_master,
            notes.as_ref(),
        )?;
        let common = compose_notes_common(
            &page_master,
            page_notes.as_ref(),
            slide_index + 1,
            &mut diagnostics,
        )?;
        let preview = assembly.layout.pages[slide_index].as_ref();
        let mut page = render_export_surface(
            &export_package,
            &page_source,
            common,
            page_master.color_map.clone(),
            &default_text_style,
            &theme,
            notes_size,
            slide_index + 1,
            Some(preview),
            &shell_layout,
            &shell_master,
            &mut assembly.input.media,
            &mut assembly.font_manager,
            assembly.table_styles.as_ref(),
        )?;
        diagnostics.append(&mut page.1);
        pages.push(std::sync::Arc::new(page.0));
    }
    let mut result = LayoutResult::new(
        pages,
        assembly.font_manager.all_font_data(),
        assembly.input.metadata.clone(),
        Vec::new(),
    );
    result.diagnostics = diagnostics;
    Ok(result)
}

#[cfg(feature = "render")]
fn remap_notes_export_scope(
    package: &mut OpcPackage,
    master_part: &str,
    master: &CT_NotesMaster,
    notes: Option<&NotesRecord>,
) -> Result<(String, CT_NotesMaster, Option<CT_NotesSlide>)> {
    const SCOPE_PART: &str = "/ppt/rdocx-notes-export.xml";

    let mut relationships = Vec::new();
    let master_map =
        copy_export_relationship_scope(package, master_part, "rdocxMaster", &mut relationships)?;
    let master_xml = master.to_xml().map_err(|error| Error::MalformedPart {
        part_name: master_part.to_owned(),
        message: error.to_string(),
    })?;
    let master_xml =
        rewrite_exact_rel_ids(&master_xml, &master_map).map_err(|error| Error::MalformedPart {
            part_name: master_part.to_owned(),
            message: error.to_string(),
        })?;
    let remapped_master =
        CT_NotesMaster::from_xml(&master_xml).map_err(|error| Error::MalformedPart {
            part_name: master_part.to_owned(),
            message: error.to_string(),
        })?;

    let remapped_notes = notes
        .map(|record| {
            let notes_map = copy_export_relationship_scope(
                package,
                &record.part_name,
                "rdocxNotes",
                &mut relationships,
            )?;
            let notes_xml = record
                .notes
                .to_xml()
                .map_err(|error| Error::MalformedPart {
                    part_name: record.part_name.clone(),
                    message: error.to_string(),
                })?;
            let notes_xml = rewrite_exact_rel_ids(&notes_xml, &notes_map).map_err(|error| {
                Error::MalformedPart {
                    part_name: record.part_name.clone(),
                    message: error.to_string(),
                }
            })?;
            CT_NotesSlide::from_xml(&notes_xml).map_err(|error| Error::MalformedPart {
                part_name: record.part_name.clone(),
                message: error.to_string(),
            })
        })
        .transpose()?;

    package.get_or_create_part_rels(SCOPE_PART).items = relationships;
    Ok((SCOPE_PART.to_owned(), remapped_master, remapped_notes))
}

#[cfg(feature = "render")]
fn copy_export_relationship_scope(
    package: &OpcPackage,
    source_part: &str,
    prefix: &str,
    output: &mut Vec<Relationship>,
) -> Result<HashMap<String, String>> {
    let mut map = HashMap::new();
    let Some(source) = package.get_part_rels(source_part) else {
        return Ok(map);
    };
    for (index, relationship) in source.items.iter().enumerate() {
        let id = format!("{prefix}{}", index + 1);
        if map.insert(relationship.id.clone(), id.clone()).is_some() {
            return Err(render_failure(format!(
                "{source_part}: duplicate relationship id {}",
                relationship.id
            )));
        }
        let external = relationship
            .target_mode
            .as_deref()
            .is_some_and(|mode| mode.eq_ignore_ascii_case("external"));
        output.push(Relationship {
            id,
            rel_type: relationship.rel_type.clone(),
            target: if external {
                relationship.target.clone()
            } else {
                OpcPackage::resolve_rel_target(source_part, &relationship.target)
            },
            target_mode: relationship.target_mode.clone(),
        });
    }
    Ok(map)
}

#[cfg(feature = "render")]
fn render_handout_pages(
    package: &OpcPackage,
    handout_layout: HandoutLayout,
) -> Result<LayoutResult> {
    let mut assembly = prepare_render_context(package, false)?;
    if assembly.slides.is_empty() {
        return Ok(LayoutResult::new(
            Vec::new(),
            assembly.font_manager.all_font_data(),
            assembly.input.metadata.clone(),
            Vec::new(),
        ));
    }
    let presentation_part = package
        .main_document_part()
        .ok_or(Error::MissingMainDocument)?;
    let presentation = CT_Presentation::from_xml(required_part(package, &presentation_part)?)
        .map_err(|error| Error::MalformedPart {
            part_name: presentation_part.clone(),
            message: error.to_string(),
        })?;
    let page_size = export_page_size(&presentation);
    let handout_master_part = render_exact_related_part(
        package,
        &presentation_part,
        rel_types::HANDOUT_MASTER,
        content_types::HANDOUT_MASTER,
    )?;
    let handout_master = CT_HandoutMaster::from_xml(required_part(package, &handout_master_part)?)
        .map_err(|error| Error::MalformedPart {
            part_name: handout_master_part.clone(),
            message: error.to_string(),
        })?;
    let theme_part = render_exact_related_part(
        package,
        &handout_master_part,
        rel_types::THEME,
        content_types::THEME,
    )?;
    let theme =
        CT_OfficeStyleSheet::from_xml(required_part(package, &theme_part)?).map_err(|error| {
            Error::MalformedPart {
                part_name: theme_part,
                message: error.to_string(),
            }
        })?;
    let default_text_style = presentation.default_text_style.clone().unwrap_or_default();
    let shell_layout = assembly.slides[0].layout.clone();
    let shell_master = assembly.slides[0].master.clone();
    let capacity = handout_layout.slides_per_page();
    let page_count = assembly.layout.pages.len().div_ceil(capacity);
    let mut pages = Vec::with_capacity(page_count);
    let mut diagnostics = assembly.layout.diagnostics.clone();
    for page_index in 0..page_count {
        let common = compose_handout_common(&handout_master, page_index + 1)?;
        let (mut base, mut page_diagnostics) = render_export_surface(
            package,
            &handout_master_part,
            common,
            handout_master.color_map.clone(),
            &default_text_style,
            &theme,
            page_size,
            page_index + 1,
            None,
            &shell_layout,
            &shell_master,
            &mut assembly.input.media,
            &mut assembly.font_manager,
            assembly.table_styles.as_ref(),
        )?;
        diagnostics.append(&mut page_diagnostics);
        let start = page_index * capacity;
        let end = (start + capacity).min(assembly.layout.pages.len());
        let mut elements = std::mem::take(&mut base.elements);
        let mut thumbnails = handout_thumbnail_elements(
            &assembly.layout.pages[start..end],
            start,
            handout_layout,
            page_size,
            &mut assembly.font_manager,
        )?;
        elements.append(&mut thumbnails);
        base.elements = elements;
        pages.push(std::sync::Arc::new(base));
    }
    let mut result = LayoutResult::new(
        pages,
        assembly.font_manager.all_font_data(),
        assembly.input.metadata.clone(),
        Vec::new(),
    );
    result.diagnostics = diagnostics;
    Ok(result)
}

#[cfg(feature = "render")]
fn export_page_size(presentation: &CT_Presentation) -> (f64, f64) {
    (
        presentation.notes_size.cx.0 as f64 / 12_700.0,
        presentation.notes_size.cy.0 as f64 / 12_700.0,
    )
}

#[cfg(feature = "render")]
fn render_exact_related_part(
    package: &OpcPackage,
    source_part: &str,
    relationship_type: &str,
    content_type: &str,
) -> Result<String> {
    let matches = package
        .get_part_rels(source_part)
        .map(|relationships| relationships.get_all_by_type(relationship_type))
        .unwrap_or_default();
    if matches.len() != 1 {
        return Err(render_failure(format!(
            "{source_part}: found {} relationships of type {relationship_type}, expected exactly one",
            matches.len()
        )));
    }
    let relationship = matches[0];
    reject_external(source_part, relationship)?;
    let target = OpcPackage::resolve_rel_target(source_part, &relationship.target);
    required_part(package, &target)?;
    let actual = package.content_types.content_type_for(&target);
    if actual != Some(content_type) {
        return Err(render_failure(format!(
            "{source_part}: relationship {} targets {target} with content type {:?}, expected {content_type}",
            relationship.id, actual
        )));
    }
    Ok(target)
}

#[cfg(feature = "render")]
fn validate_notes_export_graph(
    package: &OpcPackage,
    slide_part: &str,
    notes_part: &str,
    notes_master_part: &str,
) -> Result<()> {
    if package.content_types.content_type_for(notes_part) != Some(content_types::NOTES_SLIDE) {
        return Err(render_failure(format!(
            "{notes_part}: expected notes-slide content type"
        )));
    }
    let master = render_exact_related_part(
        package,
        notes_part,
        rel_types::NOTES_MASTER,
        content_types::NOTES_MASTER,
    )?;
    if master != notes_master_part {
        return Err(render_failure(format!(
            "{notes_part}: notes master {master} differs from presentation master {notes_master_part}"
        )));
    }
    let slide_relationships = package
        .get_part_rels(notes_part)
        .map(|relationships| relationships.get_all_by_type(rel_types::SLIDE))
        .unwrap_or_default();
    if slide_relationships.len() > 1 {
        return Err(render_failure(format!(
            "{notes_part}: found {} relationships of type {}, expected at most one",
            slide_relationships.len(),
            rel_types::SLIDE
        )));
    }
    if let Some(relationship) = slide_relationships.first() {
        reject_external(notes_part, relationship)?;
        let slide = OpcPackage::resolve_rel_target(notes_part, &relationship.target);
        required_part(package, &slide)?;
        let actual = package.content_types.content_type_for(&slide);
        if actual != Some(content_types::SLIDE) {
            return Err(render_failure(format!(
                "{notes_part}: relationship {} targets {slide} with content type {:?}, expected {}",
                relationship.id,
                actual,
                content_types::SLIDE
            )));
        }
        if slide != slide_part {
            return Err(render_failure(format!(
                "{notes_part}: source slide {slide} differs from owner {slide_part}"
            )));
        }
    }
    Ok(())
}

#[cfg(feature = "render")]
fn compose_notes_common(
    master: &CT_NotesMaster,
    notes: Option<&CT_NotesSlide>,
    slide_number: usize,
    diagnostics: &mut Vec<oxml_layout::Diagnostic>,
) -> Result<CT_CommonSlideData> {
    let mut common = master.common_slide_data.clone();
    if let Some(notes) = notes {
        let mut overlays = Vec::new();
        collect_export_placeholder_shapes(
            &notes.common_slide_data.shape_tree.children,
            &mut overlays,
        )?;
        let mut consumed = HashSet::new();
        apply_export_placeholder_overlays(
            &mut common.shape_tree.children,
            &overlays,
            &mut consumed,
        )?;
        for (index, (unmatched, _)) in overlays.iter().enumerate() {
            if !consumed.contains(&index) {
                diagnostics.push(oxml_layout::Diagnostic {
                    message: format!(
                        "notes slide placeholder {unmatched:?} has no matching notes-master placeholder and was skipped"
                    ),
                });
            }
        }
        for child in &notes.common_slide_data.shape_tree.children {
            if child_export_placeholder(child).is_none() {
                common.shape_tree.children.push(child.clone());
            }
        }
    }
    filter_export_master_shapes(
        &mut common.shape_tree.children,
        master.header_footer.as_ref(),
        notes.is_some(),
        slide_number,
    );
    let slide_images = count_export_placeholders(&common.shape_tree.children, |kind| {
        kind == &PhType::SlideImage
    });
    if slide_images != 1 {
        return Err(render_failure(format!(
            "notes master has {slide_images} visible slide-image placeholders, expected exactly one"
        )));
    }
    Ok(common)
}

#[cfg(feature = "render")]
fn compose_handout_common(
    master: &CT_HandoutMaster,
    page_number: usize,
) -> Result<CT_CommonSlideData> {
    let mut common = master.common_slide_data.clone();
    filter_export_master_shapes(
        &mut common.shape_tree.children,
        master.header_footer.as_ref(),
        false,
        page_number,
    );
    Ok(common)
}

#[cfg(feature = "render")]
fn collect_export_placeholder_shapes(
    children: &[ShapeTreeChild],
    output: &mut Vec<(PlaceholderKey, CT_Shape)>,
) -> Result<()> {
    for child in children {
        match child {
            ShapeTreeChild::Shape(shape) => {
                if let Some(placeholder) = shape.placeholder.as_ref() {
                    let key = placeholder.key();
                    if output.iter().any(|(existing, _)| existing.matches(&key)) {
                        return Err(render_failure(format!(
                            "notes slide has ambiguous placeholder {key:?}"
                        )));
                    }
                    output.push((key, shape.clone()));
                }
            }
            ShapeTreeChild::GroupShape(group) => {
                collect_export_placeholder_shapes(&group.children, output)?;
            }
            _ => {}
        }
    }
    Ok(())
}

#[cfg(feature = "render")]
fn apply_export_placeholder_overlays(
    children: &mut [ShapeTreeChild],
    overlays: &[(PlaceholderKey, CT_Shape)],
    consumed: &mut HashSet<usize>,
) -> Result<()> {
    for child in children {
        match child {
            ShapeTreeChild::Shape(shape) => {
                let Some(placeholder) = shape.placeholder.as_ref() else {
                    continue;
                };
                let key = placeholder.key();
                let mut matches = overlays
                    .iter()
                    .enumerate()
                    .filter(|(_, (overlay_key, _))| key.matches(overlay_key));
                let Some((overlay_index, (_, overlay))) = matches.next() else {
                    continue;
                };
                if matches.next().is_some() {
                    return Err(render_failure(format!(
                        "notes master placeholder {key:?} matches multiple notes-slide placeholders"
                    )));
                }
                if !consumed.insert(overlay_index) {
                    return Err(render_failure(format!(
                        "notes master has duplicate placeholder {key:?}"
                    )));
                }
                if let Some(transform) = overlay.shape_properties.transform.clone() {
                    shape.shape_properties.transform = Some(transform);
                }
                if let Some(text) = overlay.text_body.clone() {
                    shape.text_body = Some(text);
                }
            }
            ShapeTreeChild::GroupShape(group) => {
                apply_export_placeholder_overlays(&mut group.children, overlays, consumed)?;
            }
            _ => {}
        }
    }
    Ok(())
}

#[cfg(feature = "render")]
fn child_export_placeholder(child: &ShapeTreeChild) -> Option<PhType> {
    match child {
        ShapeTreeChild::Shape(shape) => shape
            .placeholder
            .as_ref()
            .map(CT_Placeholder::effective_type),
        ShapeTreeChild::Picture(picture) => picture
            .placeholder
            .as_ref()
            .map(CT_Placeholder::effective_type),
        _ => None,
    }
}

#[cfg(feature = "render")]
fn filter_export_master_shapes(
    children: &mut Vec<ShapeTreeChild>,
    header_footer: Option<&CT_HeaderFooter>,
    keep_body: bool,
    page_number: usize,
) {
    children.retain_mut(|child| {
        if let ShapeTreeChild::GroupShape(group) = child {
            filter_export_master_shapes(&mut group.children, header_footer, keep_body, page_number);
            return true;
        }
        let Some(kind) = child_export_placeholder(child) else {
            return true;
        };
        let enabled = match kind {
            PhType::DateTime => header_footer.is_none_or(CT_HeaderFooter::date_time_enabled),
            PhType::Header => header_footer.is_none_or(CT_HeaderFooter::header_enabled),
            PhType::Footer => header_footer.is_none_or(CT_HeaderFooter::footer_enabled),
            PhType::SlideNumber => header_footer.is_none_or(CT_HeaderFooter::slide_number_enabled),
            PhType::Body => keep_body,
            _ => true,
        };
        if enabled
            && kind == PhType::SlideNumber
            && let ShapeTreeChild::Shape(shape) = child
        {
            if let Some(placeholder) = shape.placeholder.as_mut() {
                placeholder.ph_type = Some(PhType::Other("exportPageNumber".to_owned()));
            }
            if let Some(body) = shape.text_body.as_mut() {
                body.set_text(&page_number.to_string());
            }
        }
        if enabled
            && kind == PhType::Body
            && let ShapeTreeChild::Shape(shape) = child
            && shape
                .text_body
                .as_ref()
                .is_none_or(|body| body.plain_text().trim().is_empty())
        {
            return false;
        }
        enabled
    });
}

#[cfg(feature = "render")]
fn count_export_placeholders(
    children: &[ShapeTreeChild],
    predicate: impl Copy + Fn(&PhType) -> bool,
) -> usize {
    children
        .iter()
        .map(|child| match child {
            ShapeTreeChild::GroupShape(group) => {
                count_export_placeholders(&group.children, predicate)
            }
            _ => child_export_placeholder(child)
                .as_ref()
                .is_some_and(predicate) as usize,
        })
        .sum()
}

#[cfg(feature = "render")]
#[allow(clippy::too_many_arguments)]
fn render_export_surface(
    package: &OpcPackage,
    source_part: &str,
    common: CT_CommonSlideData,
    color_map: ColorMap,
    default_text_style: &CT_TextListStyle,
    theme: &CT_OfficeStyleSheet,
    size: (f64, f64),
    page_number: usize,
    preview: Option<&PageFrame>,
    shell_layout: &CT_SlideLayout,
    shell_master: &CT_SlideMaster,
    deck_media: &mut HashMap<MediaId, MediaData>,
    font_manager: &mut FontManager,
    table_styles: Option<&CT_TableStyleList>,
) -> Result<(PageFrame, Vec<oxml_layout::Diagnostic>)> {
    let mut slide = CT_Slide::new(CT_ShapeTree::new());
    slide.common_slide_data = common;
    let mut layout = shell_layout.clone();
    layout.common_slide_data.background = None;
    layout.common_slide_data.shape_tree = CT_ShapeTree::new();
    layout.header_footer = None;
    let mut master = shell_master.clone();
    master.common_slide_data.background = None;
    master.common_slide_data.shape_tree = CT_ShapeTree::new();
    master.color_map = color_map.clone();
    master.text_styles = None;
    master.header_footer = None;
    let empty: &[ShapeTreeChild] = &[];
    let (media, media_failures, hyperlinks, charts, diagrams) = render_scoped_resources(
        package,
        [source_part, source_part, source_part],
        [&slide.common_slide_data.shape_tree.children, empty, empty],
        deck_media,
    )?;
    let expansion = diagram::expand_tree(&mut slide.common_slide_data.shape_tree, &diagrams.slide);
    let mut context = ResolveCtx::new(
        theme,
        color_map,
        &master,
        &layout,
        &slide,
        default_text_style,
    )
    .with_media_failures(&media_failures);
    if let Some(styles) = table_styles {
        context = context.with_table_styles(styles);
    }
    let mut preview_index = None;
    let mut source_index = 0usize;
    for item in context.flatten() {
        if !render_source_shape_has_bounds(&context, &item) {
            continue;
        }
        if let FlattenedItem::Shape { child, .. } = &item
            && child_export_placeholder(child) == Some(PhType::SlideImage)
            && preview_index.replace(source_index).is_some()
        {
            return Err(render_failure(
                "notes surface has multiple bounded slide-image placeholders",
            ));
        }
        source_index += 1;
    }
    if preview.is_some() != preview_index.is_some() {
        return Err(render_failure(
            "notes surface and slide-image placeholder do not correspond",
        ));
    }
    let (mut resolved, directions) = context
        .resolve_slide_with_chart_resources_and_text_directions(
            size,
            &media,
            &hyperlinks,
            &charts,
            font_manager,
        )
        .map_err(|error| render_failure(format!("{source_part}: {error}")))?;
    resolved.diagnostics.extend(expansion.diagnostics);
    if let (Some(preview), Some(index)) = (preview, preview_index) {
        let shape = resolved.shapes.get_mut(index).ok_or_else(|| {
            render_failure("notes slide-image placeholder was dropped during resolution")
        })?;
        let target = Rect {
            x: 0.0,
            y: 0.0,
            width: shape.bounds.width,
            height: shape.bounds.height,
        };
        shape.content = ResolvedContent::Group(page_thumbnail_group(preview, target).0);
    }
    let diagnostics = resolved.diagnostics.clone();
    let input = RenderInput {
        slides: vec![resolved],
        media: deck_media.clone(),
        fonts: Vec::new(),
        metadata: None,
    };
    let layout_result = layout_presentation_with_font_manager_and_text_directions_mut(
        &input,
        font_manager,
        &[directions],
    )
    .map_err(|error| render_failure(error.to_string()))?;
    let mut page = layout_result
        .pages
        .first()
        .map(|page| page.as_ref().clone())
        .ok_or_else(|| render_failure("export surface produced no page"))?;
    page.page_number = page_number;
    Ok((page, diagnostics))
}

#[cfg(feature = "render")]
fn page_thumbnail_group(page: &PageFrame, target: Rect) -> (GroupElement, Rect) {
    let scale = (target.width / page.width).min(target.height / page.height);
    let fitted = Rect {
        x: target.x + (target.width - page.width * scale) / 2.0,
        y: target.y + (target.height - page.height * scale) / 2.0,
        width: page.width * scale,
        height: page.height * scale,
    };
    let mut children = vec![PositionedElement::FilledRect {
        rect: Rect {
            x: 0.0,
            y: 0.0,
            width: page.width,
            height: page.height,
        },
        color: Color::WHITE,
    }];
    if let Some(background) = &page.background {
        children.push(PositionedElement::Path(PathElement {
            path: LayoutPath::rect(Rect {
                x: 0.0,
                y: 0.0,
                width: page.width,
                height: page.height,
            }),
            fill: Some(background.clone()),
            stroke: None,
        }));
    }
    children.extend(page.elements.clone());
    (
        GroupElement {
            transform: Transform {
                a: scale,
                b: 0.0,
                c: 0.0,
                d: scale,
                e: fitted.x,
                f: fitted.y,
            },
            clip: Some(LayoutPath::rect(Rect {
                x: 0.0,
                y: 0.0,
                width: page.width,
                height: page.height,
            })),
            opacity: 1.0,
            effects: Vec::new(),
            children,
        },
        fitted,
    )
}

#[cfg(feature = "render")]
fn handout_thumbnail_elements(
    pages: &[std::sync::Arc<PageFrame>],
    source_start: usize,
    layout: HandoutLayout,
    page_size: (f64, f64),
    font_manager: &mut FontManager,
) -> Result<Vec<PositionedElement>> {
    let targets = handout_targets(layout, page_size);
    let mut elements = Vec::new();
    for (offset, page) in pages.iter().enumerate() {
        let target = targets[offset];
        let (thumbnail, fitted) = page_thumbnail_group(page, target);
        elements.push(PositionedElement::Group(thumbnail));
        push_thumbnail_border(&mut elements, fitted);
        elements.push(handout_slide_number(
            source_start + offset + 1,
            fitted,
            font_manager,
        )?);
        if layout == HandoutLayout::Three {
            push_three_up_note_rules(&mut elements, target, page_size.0);
        }
    }
    Ok(elements)
}

#[cfg(feature = "render")]
fn handout_targets(layout: HandoutLayout, page_size: (f64, f64)) -> Vec<Rect> {
    let margin_x = 36.0;
    let margin_y = 54.0;
    let gap = 18.0;
    let label_height = 14.0;
    let usable_width = (page_size.0 - 2.0 * margin_x).max(1.0);
    let usable_height = (page_size.1 - 2.0 * margin_y).max(1.0);
    let (columns, rows, width_fraction) = match layout {
        HandoutLayout::One => (1, 1, 1.0),
        HandoutLayout::Two => (1, 2, 1.0),
        HandoutLayout::Three => (1, 3, 0.46),
        HandoutLayout::Four => (2, 2, 1.0),
        HandoutLayout::Six => (2, 3, 1.0),
        HandoutLayout::Nine => (3, 3, 1.0),
    };
    let cell_width = (usable_width - gap * (columns - 1) as f64) / columns as f64;
    let cell_height = (usable_height - gap * (rows - 1) as f64) / rows as f64;
    (0..layout.slides_per_page())
        .map(|index| {
            let row = index / columns;
            let column = index % columns;
            Rect {
                x: margin_x + column as f64 * (cell_width + gap),
                y: margin_y + row as f64 * (cell_height + gap),
                width: cell_width * width_fraction,
                height: (cell_height - label_height).max(1.0),
            }
        })
        .collect()
}

#[cfg(feature = "render")]
fn push_thumbnail_border(elements: &mut Vec<PositionedElement>, rect: Rect) {
    for (start, end) in [
        (
            Point {
                x: rect.x,
                y: rect.y,
            },
            Point {
                x: rect.x + rect.width,
                y: rect.y,
            },
        ),
        (
            Point {
                x: rect.x + rect.width,
                y: rect.y,
            },
            Point {
                x: rect.x + rect.width,
                y: rect.y + rect.height,
            },
        ),
        (
            Point {
                x: rect.x + rect.width,
                y: rect.y + rect.height,
            },
            Point {
                x: rect.x,
                y: rect.y + rect.height,
            },
        ),
        (
            Point {
                x: rect.x,
                y: rect.y + rect.height,
            },
            Point {
                x: rect.x,
                y: rect.y,
            },
        ),
    ] {
        elements.push(PositionedElement::Line {
            start,
            end,
            width: 0.75,
            color: Color::BLACK,
            dash_pattern: None,
        });
    }
}

#[cfg(feature = "render")]
fn handout_slide_number(
    number: usize,
    thumbnail: Rect,
    font_manager: &mut FontManager,
) -> Result<PositionedElement> {
    let text = number.to_string();
    let font_id = font_manager
        .resolve_font_for_text(Some("Carlito"), false, false, &text)
        .map_err(|error| render_failure(format!("handout slide number font: {error}")))?;
    let shaped = font_manager
        .shape_text(font_id, &text, 9.0)
        .map_err(|error| render_failure(format!("handout slide number shaping: {error}")))?;
    Ok(PositionedElement::Text(GlyphRun {
        origin: Point {
            x: thumbnail.x,
            y: thumbnail.y + thumbnail.height + 11.0,
        },
        font_id,
        font_size: 9.0,
        glyph_ids: shaped.glyph_ids,
        advances: shaped.advances,
        text,
        source: None,
        color: Color::BLACK,
        bold: false,
        italic: false,
        field_kind: None,
        field_source: None,
        note: None,
        note_reference_source: None,
        tab_aligned: None,
    }))
}

#[cfg(feature = "render")]
fn push_three_up_note_rules(
    elements: &mut Vec<PositionedElement>,
    thumbnail: Rect,
    page_width: f64,
) {
    let start_x = thumbnail.x + thumbnail.width + 18.0;
    let end_x = page_width - 36.0;
    if end_x <= start_x {
        return;
    }
    for line in 1..=5 {
        let y = thumbnail.y + thumbnail.height * line as f64 / 6.0;
        elements.push(PositionedElement::Line {
            start: Point { x: start_x, y },
            end: Point { x: end_x, y },
            width: 0.5,
            color: Color::BLACK,
            dash_pattern: None,
        });
    }
}

#[cfg(feature = "render")]
#[derive(Clone, Copy)]
struct TimelineRequest {
    slide_index: usize,
    outgoing_slide_index: Option<usize>,
    position: TimelinePosition,
    fallback_policy: Option<MediaFallbackPolicy>,
}

#[cfg(feature = "render")]
struct PreparedSlideAssembly {
    slide_part: String,
    slide: CT_Slide,
    layout: CT_SlideLayout,
    master: CT_SlideMaster,
    theme: CT_OfficeStyleSheet,
    media: ScopedMediaIds,
    media_failures: ScopedMediaFailures,
    hyperlinks: ScopedHyperlinkTargets,
    charts: ScopedChartResources,
    media_diagnostics: Vec<oxml_layout::Diagnostic>,
    smartart_clips: Vec<Option<Rect>>,
}

#[cfg(feature = "render")]
struct PreparedRenderAssembly {
    input: RenderInput,
    layout: LayoutResult,
    font_manager: FontManager,
    slides: Vec<PreparedSlideAssembly>,
    /// Paragraph directions parallel to `input.slides`, as lowering used them.
    text_directions: Vec<ResolvedSlideTextDirections>,
    size: (f64, f64),
    default_text_style: CT_TextListStyle,
    table_styles: Option<CT_TableStyleList>,
}

#[cfg(all(feature = "render", test))]
std::thread_local! {
    static STAGED_PACKAGE_COUNT: std::cell::Cell<usize> = const {
        std::cell::Cell::new(0)
    };
    static PREPARED_RENDER_ASSEMBLY_COUNT: std::cell::Cell<usize> = const {
        std::cell::Cell::new(0)
    };
    static PREPARED_MEDIA_ASSEMBLY_COUNT: std::cell::Cell<usize> = const {
        std::cell::Cell::new(0)
    };
    static PREPARED_LAYOUT_COUNT: std::cell::Cell<usize> = const {
        std::cell::Cell::new(0)
    };
    static PREPARED_FONT_MANAGER_COUNT: std::cell::Cell<usize> = const {
        std::cell::Cell::new(0)
    };
}

#[cfg(all(feature = "render", test))]
fn animation_preparation_counts() -> (usize, usize, usize, usize, usize) {
    (
        STAGED_PACKAGE_COUNT.get(),
        PREPARED_RENDER_ASSEMBLY_COUNT.get(),
        PREPARED_MEDIA_ASSEMBLY_COUNT.get(),
        PREPARED_LAYOUT_COUNT.get(),
        PREPARED_FONT_MANAGER_COUNT.get(),
    )
}

#[cfg(all(feature = "render", test))]
std::thread_local! {
    static ACTIVE_RESOLVED_SAMPLE_COUNT: std::cell::Cell<usize> = const {
        std::cell::Cell::new(0)
    };
    static MAX_RESOLVED_SAMPLE_COUNT: std::cell::Cell<usize> = const {
        std::cell::Cell::new(0)
    };
}

#[cfg(all(feature = "render", test))]
struct ActiveResolvedSample;

#[cfg(all(feature = "render", test))]
impl ActiveResolvedSample {
    fn enter() -> Self {
        ACTIVE_RESOLVED_SAMPLE_COUNT.with(|active| {
            let current = active.get() + 1;
            active.set(current);
            MAX_RESOLVED_SAMPLE_COUNT.with(|maximum| maximum.set(maximum.get().max(current)));
        });
        Self
    }
}

#[cfg(all(feature = "render", test))]
impl Drop for ActiveResolvedSample {
    fn drop(&mut self) {
        ACTIVE_RESOLVED_SAMPLE_COUNT.with(|active| active.set(active.get() - 1));
    }
}

#[cfg(all(feature = "render", test))]
fn reset_resolved_sample_retention_count() {
    ACTIVE_RESOLVED_SAMPLE_COUNT.set(0);
    MAX_RESOLVED_SAMPLE_COUNT.set(0);
}

#[cfg(all(feature = "render", test))]
fn max_resolved_sample_retention_count() -> usize {
    MAX_RESOLVED_SAMPLE_COUNT.get()
}

#[cfg(feature = "render")]
fn prepare_render_context(
    package: &OpcPackage,
    collect_media_diagnostics: bool,
) -> Result<PreparedRenderAssembly> {
    #[cfg(test)]
    PREPARED_RENDER_ASSEMBLY_COUNT.with(|count| count.set(count.get() + 1));
    let presentation_part = package
        .main_document_part()
        .ok_or(Error::MissingMainDocument)?;
    let presentation = CT_Presentation::from_xml(required_part(package, &presentation_part)?)
        .map_err(|error| Error::MalformedPart {
            part_name: presentation_part.clone(),
            message: error.to_string(),
        })?;
    let default_text_style = presentation.default_text_style.clone().unwrap_or_default();
    let size = presentation
        .slide_size
        .as_ref()
        .map_or((720.0, 540.0), |size| {
            (size.cx.0 as f64 / 12_700.0, size.cy.0 as f64 / 12_700.0)
        });
    let table_styles = render_table_styles(package, &presentation_part)?;
    let presentation_relationships = package
        .get_part_rels(&presentation_part)
        .ok_or_else(|| render_failure("presentation relationships are missing"))?;

    let mut media = HashMap::new();
    #[cfg(test)]
    PREPARED_MEDIA_ASSEMBLY_COUNT.with(|count| count.set(count.get() + 1));
    #[cfg(test)]
    PREPARED_FONT_MANAGER_COUNT.with(|count| count.set(count.get() + 1));
    let mut font_manager = FontManager::new_deterministic()
        .map_err(|error| render_failure(format!("deterministic fonts: {error}")))?;
    let mut resolved_slides = Vec::with_capacity(presentation.slide_ids.len());
    let mut text_directions = Vec::with_capacity(presentation.slide_ids.len());
    let mut prepared_slides = Vec::with_capacity(presentation.slide_ids.len());
    for (slide_index, slide_id) in presentation.slide_ids.iter().enumerate() {
        let slide_relationship = presentation_relationships
            .get_by_id(&slide_id.relationship_id)
            .ok_or_else(|| Error::MissingRelationship {
                source_part: presentation_part.clone(),
                relationship_id: slide_id.relationship_id.clone(),
            })?;
        require_relationship_type(&presentation_part, slide_relationship, rel_types::SLIDE)?;
        reject_external(&presentation_part, slide_relationship)?;
        let slide_part =
            OpcPackage::resolve_rel_target(&presentation_part, &slide_relationship.target);
        let layout_part = render_related_part(package, &slide_part, rel_types::SLIDE_LAYOUT)?;
        let master_part = render_related_part(package, &layout_part, rel_types::SLIDE_MASTER)?;
        let theme_part = render_related_part(package, &master_part, rel_types::THEME)?;
        let mut slide =
            CT_Slide::from_xml(required_part(package, &slide_part)?).map_err(|error| {
                Error::MalformedPart {
                    part_name: slide_part.clone(),
                    message: error.to_string(),
                }
            })?;
        let mut layout =
            CT_SlideLayout::from_xml(required_part(package, &layout_part)?).map_err(|error| {
                Error::MalformedPart {
                    part_name: layout_part.clone(),
                    message: error.to_string(),
                }
            })?;
        let mut master =
            CT_SlideMaster::from_xml(required_part(package, &master_part)?).map_err(|error| {
                Error::MalformedPart {
                    part_name: master_part.clone(),
                    message: error.to_string(),
                }
            })?;
        let theme = CT_OfficeStyleSheet::from_xml(required_part(package, &theme_part)?).map_err(
            |error| Error::MalformedPart {
                part_name: theme_part.clone(),
                message: error.to_string(),
            },
        )?;
        let (slide_media, slide_media_failures, slide_hyperlinks, slide_charts, slide_diagrams) =
            render_scoped_resources(
                package,
                [&slide_part, &layout_part, &master_part],
                [
                    &slide.common_slide_data.shape_tree.children,
                    &layout.common_slide_data.shape_tree.children,
                    &master.common_slide_data.shape_tree.children,
                ],
                &mut media,
            )?;
        let mut slide_expansion = diagram::expand_tree(
            &mut slide.common_slide_data.shape_tree,
            &slide_diagrams.slide,
        );
        let mut layout_expansion = diagram::expand_tree(
            &mut layout.common_slide_data.shape_tree,
            &slide_diagrams.layout,
        );
        let mut master_expansion = diagram::expand_tree(
            &mut master.common_slide_data.shape_tree,
            &slide_diagrams.master,
        );
        let mut smartart_diagnostics = std::mem::take(&mut slide_expansion.diagnostics);
        smartart_diagnostics.append(&mut layout_expansion.diagnostics);
        smartart_diagnostics.append(&mut master_expansion.diagnostics);
        let color_map = render_effective_color_map(&master, &layout, &slide);
        let mut context = ResolveCtx::new(
            &theme,
            color_map,
            &master,
            &layout,
            &slide,
            &default_text_style,
        )
        .with_media_failures(&slide_media_failures);
        if let Some(styles) = table_styles.as_ref() {
            context = context.with_table_styles(styles);
        }
        let smartart_source_clips = context
            .flatten()
            .into_iter()
            .filter(|item| render_source_shape_has_bounds(&context, item))
            .map(|item| {
                render_source_smartart_clip(
                    &item,
                    &slide_expansion.clips,
                    &layout_expansion.clips,
                    &master_expansion.clips,
                )
            })
            .collect::<Result<Vec<_>>>()?;
        let source_count = smartart_source_clips.len();
        let (mut resolved, slide_text_directions) = context
            .resolve_slide_with_chart_resources_and_text_directions(
                size,
                &slide_media,
                &slide_hyperlinks,
                &slide_charts,
                &mut font_manager,
            )
            .map_err(|error| render_failure(format!("{slide_part}: {error}")))?;
        resolved.diagnostics.append(&mut smartart_diagnostics);

        let prepared_media_diagnostics = if collect_media_diagnostics {
            let media_infos = media_infos_for_slide(package, slide_index, &slide_part, &slide)?;
            package_media_diagnostics(&media_infos)
        } else {
            Vec::new()
        };
        if source_count != resolved.shapes.len() {
            return Err(render_failure(format!(
                "{slide_part}: dropped bounded source shape, source {source_count}, resolved {}",
                resolved.shapes.len()
            )));
        }
        for (clip, shape) in smartart_source_clips.iter().zip(&resolved.shapes) {
            if let Some(clip) = clip {
                smartart_local_clip(*clip, shape.bounds)?;
            }
        }
        resolved_slides.push(resolved);
        text_directions.push(slide_text_directions);
        prepared_slides.push(PreparedSlideAssembly {
            slide_part,
            slide,
            layout,
            master,
            theme,
            media: slide_media,
            media_failures: slide_media_failures,
            hyperlinks: slide_hyperlinks,
            charts: slide_charts,
            media_diagnostics: prepared_media_diagnostics,
            smartart_clips: smartart_source_clips,
        });
    }

    let input = RenderInput {
        slides: resolved_slides,
        media,
        fonts: Vec::new(),
        metadata: None,
    };
    #[cfg(test)]
    PREPARED_LAYOUT_COUNT.with(|count| count.set(count.get() + 1));
    let mut layout = layout_presentation_with_font_manager_and_text_directions_mut(
        &input,
        &mut font_manager,
        &text_directions,
    )
    .map_err(|error| render_failure(error.to_string()))?;
    if layout.pages.len() != input.slides.len() {
        return Err(render_failure(format!(
            "rendered {} pages for {} slides",
            layout.pages.len(),
            input.slides.len()
        )));
    }
    for ((page, slide), prepared) in layout
        .pages
        .iter_mut()
        .zip(&input.slides)
        .zip(&prepared_slides)
    {
        apply_smartart_clips(
            std::sync::Arc::make_mut(page),
            slide,
            &prepared.smartart_clips,
        )?;
    }
    Ok(PreparedRenderAssembly {
        input,
        layout,
        font_manager,
        slides: prepared_slides,
        text_directions,
        size,
        default_text_style,
        table_styles,
    })
}

#[cfg(feature = "render")]
fn evaluate_slide_media(
    slide: &CT_Slide,
    position: TimelinePosition,
) -> (Vec<EvaluatedMediaState>, Vec<oxml_layout::Diagnostic>) {
    let mut specs = Vec::new();
    collect_slide_media_specs(&slide.common_slide_data.shape_tree.children, &mut specs);
    let mut states = Vec::with_capacity(specs.len());
    let mut diagnostics = Vec::new();
    for (shape_id, _, trim_start_ms, trim_end_ms) in specs {
        let (state, mut media_diagnostics) = evaluate_media_playback(
            slide.timing.as_ref(),
            shape_id,
            trim_start_ms,
            trim_end_ms,
            position,
        );
        states.push(state);
        diagnostics.append(&mut media_diagnostics);
    }
    (states, diagnostics)
}

#[cfg(feature = "render")]
fn package_media_diagnostics(media_infos: &[MediaInfo]) -> Vec<oxml_layout::Diagnostic> {
    media_infos
        .iter()
        .flat_map(|media| {
            media
                .diagnostics
                .iter()
                .map(move |diagnostic| oxml_layout::Diagnostic {
                    message: match diagnostic {
                        MediaDiagnostic::UnsupportedContentType { content_type } => format!(
                            "media shape {} has unsupported content type `{content_type}`",
                            media.shape_id
                        ),
                        MediaDiagnostic::MissingPlaybackTiming => {
                            format!("media shape {} has no playback timing", media.shape_id)
                        }
                        MediaDiagnostic::UnsupportedPlaybackCommand { command } => format!(
                            "media shape {} has unsupported playback command `{command}`",
                            media.shape_id
                        ),
                        MediaDiagnostic::MissingPoster => {
                            format!("media shape {} has no poster relationship", media.shape_id)
                        }
                    },
                })
        })
        .collect()
}

#[cfg(feature = "render")]
fn suppress_consumed_media_timing_diagnostics(
    diagnostics: &mut Vec<oxml_layout::Diagnostic>,
    timing: Option<&CT_Timing>,
    consumed_shape_ids: &HashSet<u32>,
) {
    const RETAINED_MEDIA_TIMING: &str = "media timing is retained for synchronized playback";

    let mut suppressions = Vec::new();
    if let Some(timing) = timing {
        collect_media_timing_suppressions(timing.nodes(), consumed_shape_ids, &mut suppressions);
    }
    let mut marker_index = 0;
    diagnostics.retain(|diagnostic| {
        if diagnostic.message != RETAINED_MEDIA_TIMING {
            return true;
        }
        let suppress = suppressions.get(marker_index).copied().unwrap_or(false);
        marker_index += 1;
        !suppress
    });
}

#[cfg(feature = "render")]
fn collect_media_timing_suppressions(
    nodes: &[TimingNode],
    consumed_shape_ids: &HashSet<u32>,
    suppressions: &mut Vec<bool>,
) {
    for node in nodes {
        match node {
            TimingNode::Audio(media) | TimingNode::Video(media) => suppressions.push(
                matches!(media.common.target, TimingTarget::Shape(shape_id) if consumed_shape_ids.contains(&shape_id)),
            ),
            TimingNode::MediaCommand(command) => suppressions.push(
                matches!(command.target, TimingTarget::Shape(shape_id) if consumed_shape_ids.contains(&shape_id)),
            ),
            TimingNode::Parallel(container) => collect_media_timing_suppressions(
                &container.common.children,
                consumed_shape_ids,
                suppressions,
            ),
            TimingNode::Sequence(sequence) => collect_media_timing_suppressions(
                &sequence.common.children,
                consumed_shape_ids,
                suppressions,
            ),
            TimingNode::Set(_)
            | TimingNode::Animate(_)
            | TimingNode::Effect(_)
            | TimingNode::Motion(_)
            | TimingNode::Unsupported(_) => {}
        }
    }
}

#[cfg(feature = "render")]
fn collect_slide_media_specs(
    children: &[ShapeTreeChild],
    specs: &mut Vec<(u32, MediaKind, Option<u64>, Option<u64>)>,
) {
    for child in children {
        match child {
            ShapeTreeChild::Picture(picture) => {
                if let Some(media) = picture.media.as_ref()
                    && let Some(shape_id) = child.non_visual_id()
                {
                    let (trim_start_ms, trim_end_ms) = picture.media_trim_bounds();
                    specs.push((shape_id, media.kind, trim_start_ms, trim_end_ms));
                }
            }
            ShapeTreeChild::GroupShape(group) => {
                collect_slide_media_specs(&group.children, specs);
            }
            _ => {}
        }
    }
}

#[cfg(feature = "render")]
fn apply_media_fallback_policy(
    timeline: &mut ResolvedTimelineSlide,
    slide: &CT_Slide,
    policy: MediaFallbackPolicy,
    fonts: &mut FontManager,
) -> Result<()> {
    let mut specs = Vec::new();
    collect_slide_media_specs(&slide.common_slide_data.shape_tree.children, &mut specs);
    let kinds = specs
        .into_iter()
        .map(|(shape_id, kind, _, _)| (shape_id, kind))
        .collect::<HashMap<_, _>>();
    for (shape, identity) in timeline.slide.shapes.iter_mut().zip(&timeline.identities) {
        if identity.source != FlattenedSource::Slide {
            continue;
        }
        let Some(shape_id) = identity.shape_id else {
            continue;
        };
        let Some(kind) = kinds.get(&shape_id).copied() else {
            continue;
        };
        let poster_resolved = matches!(shape.content, ResolvedContent::Image(_));
        if poster_resolved && policy != MediaFallbackPolicy::DeterministicPlaceholder {
            continue;
        }
        if policy == MediaFallbackPolicy::Fail {
            return Err(render_failure(format!(
                "media shape {shape_id} has no renderer-compatible poster"
            )));
        }
        let label = match kind {
            MediaKind::Audio => "Audio",
            MediaKind::Video => "Video",
        };
        shape.content = ResolvedContent::Group(render_media_placeholder(
            label,
            shape.bounds.width,
            shape.bounds.height,
            fonts,
        )?);
        shape.unsupported = Some("media poster");
        let fallback_message =
            format!("media shape {shape_id} rendered as deterministic {label} placeholder");
        let poster_diagnostic_prefix =
            format!("unresolved slide media poster for shape {shape_id}:");
        if let Some(diagnostic) = timeline
            .slide
            .diagnostics
            .iter_mut()
            .find(|diagnostic| diagnostic.message.starts_with(&poster_diagnostic_prefix))
        {
            diagnostic.message = fallback_message;
        } else {
            timeline.slide.diagnostics.push(oxml_layout::Diagnostic {
                message: fallback_message,
            });
        }
    }
    restore_media_poster_diagnostic_messages(&mut timeline.slide.diagnostics);
    Ok(())
}

#[cfg(feature = "render")]
fn restore_media_poster_diagnostic_messages(diagnostics: &mut [oxml_layout::Diagnostic]) {
    for diagnostic in diagnostics {
        let Some(message) = legacy_media_poster_diagnostic_message(&diagnostic.message) else {
            continue;
        };
        diagnostic.message = message.to_owned();
    }
}

#[cfg(feature = "render")]
fn legacy_media_poster_diagnostic_message(message: &str) -> Option<&str> {
    ["slide", "layout", "master"]
        .into_iter()
        .find_map(|source| {
            let identity =
                message.strip_prefix(&format!("unresolved {source} media poster for shape "))?;
            let (shape_id, legacy) = identity.split_once(": ")?;
            (!shape_id.is_empty() && shape_id.chars().all(|character| character.is_ascii_digit()))
                .then_some(legacy)
        })
}

#[cfg(feature = "render")]
fn render_media_placeholder(
    label: &str,
    width: f64,
    height: f64,
    fonts: &mut FontManager,
) -> Result<GroupElement> {
    let font_id = fonts
        .resolve_font_for_text(Some("Carlito"), false, false, label)
        .map_err(|error| render_failure(format!("media fallback font: {error}")))?;
    let shaped = fonts
        .shape_text(font_id, label, 9.0)
        .map_err(|error| render_failure(format!("media fallback shaping: {error}")))?;
    Ok(GroupElement {
        transform: Transform::IDENTITY,
        clip: Some(LayoutPath::rect(Rect {
            x: 0.0,
            y: 0.0,
            width,
            height,
        })),
        opacity: 1.0,
        effects: Vec::new(),
        children: vec![PositionedElement::Text(GlyphRun {
            origin: Point {
                x: 4.0,
                y: (height / 2.0).max(9.0),
            },
            font_id,
            font_size: 9.0,
            glyph_ids: shaped.glyph_ids,
            advances: shaped.advances,
            text: label.to_owned(),
            source: None,
            color: Color::BLACK,
            bold: false,
            italic: false,
            field_kind: None,
            field_source: None,
            note: None,
            note_reference_source: None,
            tab_aligned: None,
        })],
    })
}

#[cfg(feature = "render")]
fn render_failure(message: impl Into<String>) -> Error {
    Error::Render {
        message: message.into(),
    }
}

#[cfg(feature = "render")]
fn render_table_styles(
    package: &OpcPackage,
    presentation_part: &str,
) -> Result<Option<CT_TableStyleList>> {
    let Some(relationship) = package
        .get_part_rels(presentation_part)
        .and_then(|relationships| relationships.get_by_type(rel_types::TABLE_STYLES))
    else {
        return Ok(None);
    };
    let target = OpcPackage::resolve_rel_target(presentation_part, &relationship.target);
    CT_TableStyleList::from_xml(required_part(package, &target)?)
        .map(Some)
        .map_err(|error| Error::MalformedPart {
            part_name: target,
            message: error.to_string(),
        })
}

#[cfg(feature = "render")]
fn render_scoped_resources(
    package: &OpcPackage,
    parts: [&str; 3],
    shape_trees: [&[ShapeTreeChild]; 3],
    deck_media: &mut HashMap<MediaId, MediaData>,
) -> Result<(
    ScopedMediaIds,
    ScopedMediaFailures,
    ScopedHyperlinkTargets,
    ScopedChartResources,
    ScopedDiagramResources,
)> {
    let slide = render_part_resources(package, parts[0], shape_trees[0], deck_media)?;
    let layout = render_part_resources(package, parts[1], shape_trees[1], deck_media)?;
    let master = render_part_resources(package, parts[2], shape_trees[2], deck_media)?;
    Ok((
        ScopedMediaIds {
            slide: slide.media_ids,
            layout: layout.media_ids,
            master: master.media_ids,
            media_content_types: deck_media
                .iter()
                .map(|(media_id, media)| (*media_id, media.content_type.clone()))
                .collect(),
        },
        ScopedMediaFailures {
            slide: slide.media_failures,
            layout: layout.media_failures,
            master: master.media_failures,
        },
        ScopedHyperlinkTargets {
            slide: slide.hyperlinks,
            layout: layout.hyperlinks,
            master: master.hyperlinks,
        },
        ScopedChartResources {
            slide: slide.charts,
            layout: layout.charts,
            master: master.charts,
        },
        ScopedDiagramResources {
            slide: slide.diagrams,
            layout: layout.diagrams,
            master: master.diagrams,
        },
    ))
}

#[cfg(feature = "render")]
struct RenderPartResources {
    media_ids: HashMap<String, MediaId>,
    media_failures: HashMap<String, String>,
    hyperlinks: HashMap<String, String>,
    charts: HashMap<String, ChartResource>,
    diagrams: HashMap<String, DiagramResources>,
}

#[cfg(feature = "render")]
fn render_part_resources(
    package: &OpcPackage,
    source_part: &str,
    shape_tree: &[ShapeTreeChild],
    deck_media: &mut HashMap<MediaId, MediaData>,
) -> Result<RenderPartResources> {
    let mut media_ids = HashMap::new();
    let mut media_failures = HashMap::new();
    let mut hyperlinks = HashMap::new();
    let mut charts = HashMap::new();
    let diagrams = render_part_diagram_resources(package, source_part, shape_tree);
    let Some(relationships) = package.get_part_rels(source_part) else {
        return Ok(RenderPartResources {
            media_ids,
            media_failures,
            hyperlinks,
            charts,
            diagrams,
        });
    };
    for relationship in &relationships.items {
        let external = relationship
            .target_mode
            .as_deref()
            .is_some_and(|mode| mode.eq_ignore_ascii_case("external"));
        if relationship.rel_type == rel_types::HYPERLINK && external {
            hyperlinks.insert(relationship.id.clone(), relationship.target.clone());
        }
        if relationship.rel_type == rel_types::CHART {
            let resource = if external {
                ChartResource::External(relationship.target.clone())
            } else {
                let target = OpcPackage::resolve_rel_target(source_part, &relationship.target);
                match package.get_part(&target) {
                    Some(bytes) => CT_ChartSpace::from_xml(bytes)
                        .map(|chart| ChartResource::Parsed(Box::new(chart)))
                        .unwrap_or_else(|error| ChartResource::Invalid(error.to_string())),
                    None => ChartResource::MissingTarget(target),
                }
            };
            charts.insert(relationship.id.clone(), resource);
        }
        if relationship.rel_type != rel_types::IMAGE || external {
            continue;
        }
        let target = OpcPackage::resolve_rel_target(source_part, &relationship.target);
        let Some(bytes) = package.get_part(&target) else {
            continue;
        };
        let content_type = package
            .content_types
            .content_type_for(&target)
            .unwrap_or("application/octet-stream")
            .to_owned();
        match render_compatible_media(bytes, &content_type) {
            Ok(true) => {}
            Ok(false) => continue,
            Err(message) => {
                media_failures.insert(relationship.id.clone(), message.to_owned());
                continue;
            }
        }
        let media_id = MediaId::from_bytes(bytes);
        deck_media.entry(media_id).or_insert_with(|| MediaData {
            bytes: bytes.to_vec(),
            content_type,
        });
        media_ids.insert(relationship.id.clone(), media_id);
    }
    Ok(RenderPartResources {
        media_ids,
        media_failures,
        hyperlinks,
        charts,
        diagrams,
    })
}

#[cfg(feature = "render")]
fn render_part_diagram_resources(
    package: &OpcPackage,
    source_part: &str,
    shape_tree: &[ShapeTreeChild],
) -> HashMap<String, DiagramResources> {
    let mut frames = Vec::new();
    collect_smartart_frames(shape_tree, &mut frames);
    let mut resources = HashMap::new();
    for (frame, _) in frames {
        let GraphicDataPayload::SmartArt(relationship_ids) = frame.graphic_data.payload() else {
            continue;
        };
        let mut relationships = (**relationship_ids).clone();
        let data_part = resolve_diagram_part(
            package,
            source_part,
            &relationships.data,
            rel_types::DIAGRAM_DATA,
            CT_DiagramData::from_xml,
        );
        let drawing_id = match &data_part {
            DiagramPart::Parsed(data) => data.drawing_relationship_id().map(str::to_owned),
            _ => None,
        };
        relationships.drawing = drawing_id.clone();
        let diagram = DiagramResources {
            relationships: relationships.clone(),
            data: diagram_part_result(data_part),
            layout: diagram_part_result(resolve_diagram_part(
                package,
                source_part,
                &relationships.layout,
                rel_types::DIAGRAM_LAYOUT,
                CT_DiagramLayoutDefinition::from_xml,
            )),
            style: diagram_part_result(resolve_diagram_part(
                package,
                source_part,
                &relationships.style,
                rel_types::DIAGRAM_QUICK_STYLE,
                CT_DiagramStyleDefinition::from_xml,
            )),
            colors: diagram_part_result(resolve_diagram_part(
                package,
                source_part,
                &relationships.colors,
                rel_types::DIAGRAM_COLORS,
                CT_DiagramColorsDefinition::from_xml,
            )),
        };
        resources
            .entry(relationships.data.clone())
            .and_modify(|existing: &mut DiagramResources| {
                if existing.relationships != relationships {
                    existing.data = Err(format!(
                        "conflicting SmartArt relationship sets share data id `{}`",
                        relationships.data
                    ));
                }
            })
            .or_insert(diagram);
    }
    resources
}

#[cfg(feature = "render")]
fn diagram_part_result<T>(part: DiagramPart<T>) -> std::result::Result<Box<T>, String> {
    match part {
        DiagramPart::Parsed(value) => Ok(Box::new(value)),
        DiagramPart::External(target) => Err(format!("external SmartArt target `{target}`")),
        DiagramPart::MissingTarget(target) => Err(format!("missing SmartArt target `{target}`")),
        DiagramPart::Invalid(detail) => Err(format!("invalid SmartArt part: {detail}")),
    }
}

#[cfg(feature = "render")]
fn render_compatible_media(
    bytes: &[u8],
    content_type: &str,
) -> std::result::Result<bool, &'static str> {
    let format = match ImageFormat::sniff(bytes) {
        Some(ImageFormat::Png) if content_type.eq_ignore_ascii_case("image/png") => {
            ImageFormat::Png
        }
        Some(ImageFormat::Jpeg)
            if content_type.eq_ignore_ascii_case("image/jpeg")
                || content_type.eq_ignore_ascii_case("image/jpg") =>
        {
            ImageFormat::Jpeg
        }
        _ => return Ok(false),
    };
    if !render_media_within_decode_bounds(bytes, format)? {
        return Ok(false);
    }
    let Some(info) = probe(bytes).filter(|info| info.width_px > 0 && info.height_px > 0) else {
        return Ok(false);
    };
    if format == ImageFormat::Jpeg && (info.bit_depth != 8 || info.channels != 3) {
        return Ok(false);
    }
    let width = f64::from(info.width_px);
    let height = f64::from(info.height_px);
    let media_id = MediaId::from_bytes(bytes);
    Ok([Color::BLACK, Color::WHITE].into_iter().any(|background| {
        let background_element = PositionedElement::FilledRect {
            rect: Rect {
                x: 0.0,
                y: 0.0,
                width,
                height,
            },
            color: background,
        };
        let baseline = LayoutResult::new(
            vec![PageFrame::new(1, width, height, vec![background_element.clone()]).into()],
            Vec::new(),
            None,
            Vec::new(),
        );
        let candidate = LayoutResult::new(
            vec![
                PageFrame::new(
                    1,
                    width,
                    height,
                    vec![
                        background_element,
                        PositionedElement::Image {
                            rect: Rect {
                                x: 0.0,
                                y: 0.0,
                                width,
                                height,
                            },
                            data: bytes.to_vec(),
                            content_type: content_type.to_owned(),
                            media_id,
                        },
                    ],
                )
                .into(),
            ],
            Vec::new(),
            None,
            Vec::new(),
        );
        oxml_pdf::render_page_to_png(&baseline, 0, 72.0)
            .zip(oxml_pdf::render_page_to_png(&candidate, 0, 72.0))
            .is_some_and(|(baseline, candidate)| baseline != candidate)
    }))
}

#[cfg(feature = "render")]
const MAX_RENDER_PREVIEW_BYTES: usize = 16 * 1024 * 1024;
#[cfg(feature = "render")]
const MAX_RENDER_PREVIEW_DECODED_BYTES: usize = 64 * 1024 * 1024;

#[cfg(feature = "render")]
fn render_media_within_decode_bounds(
    bytes: &[u8],
    format: ImageFormat,
) -> std::result::Result<bool, &'static str> {
    if bytes.len() > MAX_RENDER_PREVIEW_BYTES {
        return Err("exceeds the 16 MiB encoded-image render limit");
    }
    let Some(info) = probe(bytes) else {
        return Ok(false);
    };
    let Some(pixel_bytes) = usize::try_from(info.width_px)
        .ok()
        .and_then(|width| width.checked_mul(info.height_px as usize))
        .and_then(|pixels| pixels.checked_mul(4))
    else {
        return Err("has dimensions that overflow the decoded-image size");
    };
    if pixel_bytes > MAX_RENDER_PREVIEW_DECODED_BYTES {
        return Err("exceeds the 64 MiB decoded-image render limit");
    }
    if format != ImageFormat::Png {
        return Ok(true);
    }
    let Some(expected) = usize::try_from(info.width_px)
        .ok()
        .and_then(|width| width.checked_mul(info.channels as usize))
        .and_then(|stride| stride.checked_add(1))
        .and_then(|row| row.checked_mul(info.height_px as usize))
    else {
        return Err("has dimensions that overflow the decoded-image size");
    };
    if expected > MAX_RENDER_PREVIEW_DECODED_BYTES {
        return Err("exceeds the 64 MiB decoded-image render limit");
    }
    let mut idat = Vec::new();
    let mut offset = 33usize;
    while offset < bytes.len() {
        let Some(length_bytes) = bytes.get(offset..offset.saturating_add(4)) else {
            return Ok(false);
        };
        let length = u32::from_be_bytes(length_bytes.try_into().unwrap()) as usize;
        let Some(kind_start) = offset.checked_add(4) else {
            return Ok(false);
        };
        let Some(payload_start) = kind_start.checked_add(4) else {
            return Ok(false);
        };
        let Some(payload_end) = payload_start.checked_add(length) else {
            return Ok(false);
        };
        let Some(chunk_end) = payload_end.checked_add(4) else {
            return Ok(false);
        };
        let (Some(kind), Some(payload), Some(_)) = (
            bytes.get(kind_start..payload_start),
            bytes.get(payload_start..payload_end),
            bytes.get(payload_end..chunk_end),
        ) else {
            return Ok(false);
        };
        if kind == b"IDAT" {
            if idat.len().saturating_add(payload.len()) > MAX_RENDER_PREVIEW_BYTES {
                return Err("exceeds the 16 MiB encoded-image render limit");
            }
            idat.extend_from_slice(payload);
        }
        if kind == b"IEND" {
            break;
        }
        offset = chunk_end;
    }
    Ok(!idat.is_empty()
        && miniz_oxide::inflate::decompress_to_vec_zlib_with_limit(&idat, expected)
            .is_ok_and(|decoded| decoded.len() >= expected))
}

#[cfg(feature = "render")]
fn render_source_shape_has_bounds(context: &ResolveCtx<'_>, item: &FlattenedItem<'_>) -> bool {
    let FlattenedItem::Shape { child, .. } = item else {
        return false;
    };
    match child {
        ShapeTreeChild::Shape(shape) => context
            .effective_xfrm(shape)
            .as_ref()
            .is_some_and(render_transform_has_bounds),
        ShapeTreeChild::Picture(picture) => context
            .effective_picture_xfrm(picture)
            .as_ref()
            .is_some_and(render_transform_has_bounds),
        ShapeTreeChild::GraphicFrame(frame) => render_transform_has_bounds(&frame.transform),
        ShapeTreeChild::Connector(connector) => connector
            .shape_properties
            .transform
            .as_ref()
            .is_some_and(render_connector_transform_has_bounds),
        ShapeTreeChild::AlternateContent(alternate) => alternate
            .chart_choice()
            .is_some_and(|frame| render_transform_has_bounds(&frame.transform)),
        ShapeTreeChild::GroupShape(_) => false,
    }
}

#[cfg(feature = "render")]
fn render_source_smartart_clip(
    item: &FlattenedItem<'_>,
    slide_clips: &HashMap<u32, Rect>,
    layout_clips: &HashMap<u32, Rect>,
    master_clips: &HashMap<u32, Rect>,
) -> Result<Option<Rect>> {
    let FlattenedItem::Shape {
        source,
        child,
        group_scale,
        ..
    } = item
    else {
        return Ok(None);
    };
    let Some(id) = child.non_visual_id() else {
        return Ok(None);
    };
    let frame = match source {
        FlattenedSource::Slide => slide_clips.get(&id).copied(),
        FlattenedSource::Layout => layout_clips.get(&id).copied(),
        FlattenedSource::Master => master_clips.get(&id).copied(),
        FlattenedSource::Background => None,
    };
    let Some(frame) = frame else {
        return Ok(None);
    };
    let frame = Rect {
        x: frame.x * group_scale.0,
        y: frame.y * group_scale.1,
        width: frame.width * group_scale.0,
        height: frame.height * group_scale.1,
    };
    if ![
        frame.x,
        frame.y,
        frame.width,
        frame.height,
        group_scale.0,
        group_scale.1,
    ]
    .into_iter()
    .all(f64::is_finite)
        || frame.width <= 0.0
        || frame.height <= 0.0
    {
        return Err(render_failure(
            "SmartArt parent group has invalid clip scale",
        ));
    }
    Ok(Some(frame))
}

#[cfg(feature = "render")]
fn smartart_local_clip(frame: Rect, shape: Rect) -> Result<Rect> {
    let clip = Rect {
        x: frame.x - shape.x,
        y: frame.y - shape.y,
        width: frame.width,
        height: frame.height,
    };
    if ![
        clip.x,
        clip.y,
        clip.width,
        clip.height,
        shape.x,
        shape.y,
        shape.width,
        shape.height,
    ]
    .into_iter()
    .all(f64::is_finite)
        || clip.width <= 0.0
        || clip.height <= 0.0
        || shape.width <= 0.0
        || shape.height <= 0.0
    {
        return Err(render_failure(
            "SmartArt clip has invalid resolved geometry",
        ));
    }
    Ok(clip)
}

#[cfg(feature = "render")]
fn apply_smartart_clips(
    page: &mut PageFrame,
    slide: &ResolvedSlide,
    clips: &[Option<Rect>],
) -> Result<()> {
    if clips.len() != slide.shapes.len() {
        return Err(render_failure(format!(
            "SmartArt clip count {}, resolved shape count {}",
            clips.len(),
            slide.shapes.len()
        )));
    }
    let shape_offset = page
        .elements
        .len()
        .checked_sub(slide.shapes.len())
        .ok_or_else(|| render_failure("rendered page dropped resolved shapes"))?;
    for ((element, shape), frame) in page.elements[shape_offset..]
        .iter_mut()
        .zip(&slide.shapes)
        .zip(clips)
    {
        let Some(frame) = frame else {
            continue;
        };
        let clip = LayoutPath::rect(smartart_local_clip(*frame, shape.bounds)?);
        let PositionedElement::Group(group) = element else {
            return Err(render_failure(
                "SmartArt resolved shape did not lower to a render group",
            ));
        };
        if group.clip.as_ref() == Some(&clip) {
            continue;
        }
        if let Some(existing) = group.clip.replace(clip) {
            group.children = vec![PositionedElement::Group(GroupElement {
                transform: Transform::IDENTITY,
                clip: Some(existing),
                opacity: 1.0,
                effects: Vec::new(),
                children: std::mem::take(&mut group.children),
            })];
        }
    }
    Ok(())
}

#[cfg(feature = "render")]
fn render_transform_has_bounds(transform: &CT_Transform2D) -> bool {
    transform
        .extent
        .is_some_and(|extent| extent.cx.0 > 0 && extent.cy.0 > 0)
}

#[cfg(feature = "render")]
fn render_connector_transform_has_bounds(transform: &CT_Transform2D) -> bool {
    transform.extent.is_some_and(|extent| {
        extent.cx.0 >= 0 && extent.cy.0 >= 0 && (extent.cx.0 > 0 || extent.cy.0 > 0)
    })
}

#[cfg(feature = "render")]
fn render_effective_color_map(
    master: &CT_SlideMaster,
    layout: &CT_SlideLayout,
    slide: &CT_Slide,
) -> ColorMap {
    match slide
        .color_map_override
        .as_ref()
        .or(layout.color_map_override.as_ref())
    {
        Some(override_value) => match &override_value.kind {
            ColorMapOverrideKind::Master => master.color_map.clone(),
            ColorMapOverrideKind::Override(map) => map.clone(),
        },
        None => master.color_map.clone(),
    }
}

#[cfg(feature = "render")]
fn render_related_part(
    package: &OpcPackage,
    source_part: &str,
    relationship_type: &str,
) -> Result<String> {
    let relationship = package
        .get_part_rels(source_part)
        .and_then(|relationships| relationships.get_by_type(relationship_type))
        .ok_or_else(|| {
            render_failure(format!(
                "{source_part}: missing relationship {relationship_type}"
            ))
        })?;
    if relationship
        .target_mode
        .as_deref()
        .is_some_and(|mode| mode.eq_ignore_ascii_case("external"))
    {
        return Err(Error::ExternalRelationship {
            source_part: source_part.to_owned(),
            relationship_id: relationship.id.clone(),
        });
    }
    Ok(OpcPackage::resolve_rel_target(
        source_part,
        &relationship.target,
    ))
}

#[cfg(test)]
mod write_tests {
    use std::collections::HashSet;

    use rpptx_oxml::placeholder::PhType;

    use super::*;

    #[cfg(feature = "render")]
    fn render_bound_png_header(width: u32, height: u32) -> Vec<u8> {
        let mut bytes = b"\x89PNG\r\n\x1a\n\0\0\0\rIHDR".to_vec();
        bytes.extend_from_slice(&width.to_be_bytes());
        bytes.extend_from_slice(&height.to_be_bytes());
        bytes.extend_from_slice(&[8, 6, 0, 0, 0]);
        bytes.extend_from_slice(&[0; 4]);
        bytes
    }

    #[cfg(feature = "render")]
    #[test]
    fn notes_placeholder_skip_diagnostics_are_ordered_and_hard_failures_remain() {
        const P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
        const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
        let scaffold = r#"<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>"#;
        let shape = |id: u32, kind: &str, index: u32| {
            format!(
                r#"<p:sp><p:nvSpPr><p:cNvPr id="{id}" name="placeholder"/><p:cNvSpPr/><p:nvPr><p:ph type="{kind}" idx="{index}"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>"#
            )
        };
        let build_master = |include_slide_image: bool| {
            let slide_image = if include_slide_image {
                shape(2, "sldImg", 2)
            } else {
                String::new()
            };
            CT_NotesMaster::from_xml(
                format!(
                    r#"<p:notesMaster xmlns:p="{P}" xmlns:a="{A}"><p:cSld><p:spTree>{scaffold}{slide_image}{}</p:spTree></p:cSld><p:clrMap accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" bg1="lt1" bg2="lt2" folHlink="folHlink" hlink="hlink" tx1="dk1" tx2="dk2"/></p:notesMaster>"#,
                    shape(3, "sldNum", 5),
                )
                .as_bytes(),
            )
            .unwrap()
        };
        let notes = |overlays: &[(u32, &str, u32)]| {
            let overlays = overlays
                .iter()
                .map(|(id, kind, index)| shape(*id, kind, *index))
                .collect::<String>();
            CT_NotesSlide::from_xml(
                format!(
                    r#"<p:notes xmlns:p="{P}" xmlns:a="{A}"><p:cSld><p:spTree>{scaffold}{overlays}</p:spTree></p:cSld></p:notes>"#
                )
                .as_bytes(),
            )
            .unwrap()
        };

        let master = build_master(true);
        let unmatched = notes(&[(4, "sldNum", 12), (5, "ftr", 13)]);
        let mut diagnostics = Vec::new();
        compose_notes_common(&master, Some(&unmatched), 1, &mut diagnostics).unwrap();
        assert_eq!(
            diagnostics
                .iter()
                .map(|diagnostic| diagnostic.message.as_str())
                .collect::<Vec<_>>(),
            [
                "notes slide placeholder PlaceholderKey { ph_type: SlideNumber, idx: Some(12) } has no matching notes-master placeholder and was skipped",
                "notes slide placeholder PlaceholderKey { ph_type: Footer, idx: Some(13) } has no matching notes-master placeholder and was skipped",
            ]
        );

        let ambiguous = notes(&[(4, "sldNum", 12), (5, "sldNum", 12)]);
        assert!(
            compose_notes_common(&master, Some(&ambiguous), 1, &mut Vec::new())
                .unwrap_err()
                .to_string()
                .contains("ambiguous placeholder")
        );
        assert!(
            compose_notes_common(&build_master(false), None, 1, &mut Vec::new())
                .unwrap_err()
                .to_string()
                .contains("expected exactly one")
        );
    }

    #[cfg(feature = "render")]
    #[test]
    fn render_media_bounds_reject_malformed_and_overflowing_dimensions() {
        assert_eq!(
            render_media_within_decode_bounds(b"not an image", ImageFormat::Png),
            Ok(false)
        );
        assert_eq!(
            render_media_within_decode_bounds(
                &render_bound_png_header(u32::MAX, u32::MAX),
                ImageFormat::Png,
            ),
            Err("has dimensions that overflow the decoded-image size")
        );
        assert_eq!(
            render_media_within_decode_bounds(
                &render_bound_png_header(5_000, 4_000),
                ImageFormat::Png,
            ),
            Err("exceeds the 64 MiB decoded-image render limit")
        );
    }

    #[cfg(feature = "render")]
    #[test]
    fn handout_geometry_is_exact_clipped_and_one_point_sensitive() {
        let page_size = (540.0, 720.0);
        let cases = [
            (
                HandoutLayout::One,
                1,
                Rect {
                    x: 36.0,
                    y: 54.0,
                    width: 468.0,
                    height: 598.0,
                },
            ),
            (
                HandoutLayout::Two,
                2,
                Rect {
                    x: 36.0,
                    y: 369.0,
                    width: 468.0,
                    height: 283.0,
                },
            ),
            (
                HandoutLayout::Three,
                3,
                Rect {
                    x: 36.0,
                    y: 474.0,
                    width: 215.28,
                    height: 178.0,
                },
            ),
            (
                HandoutLayout::Four,
                4,
                Rect {
                    x: 279.0,
                    y: 369.0,
                    width: 225.0,
                    height: 283.0,
                },
            ),
            (
                HandoutLayout::Six,
                6,
                Rect {
                    x: 279.0,
                    y: 474.0,
                    width: 225.0,
                    height: 178.0,
                },
            ),
            (
                HandoutLayout::Nine,
                9,
                Rect {
                    x: 360.0,
                    y: 474.0,
                    width: 144.0,
                    height: 178.0,
                },
            ),
        ];
        for (layout, count, expected_last) in cases {
            let targets = handout_targets(layout, page_size);
            assert_eq!(targets.len(), count);
            assert_eq!(targets[0].x, 36.0);
            assert_eq!(targets[0].y, 54.0);
            assert_eq!(*targets.last().unwrap(), expected_last);
            assert!(targets.iter().all(|target| {
                target.x >= 0.0
                    && target.y >= 0.0
                    && target.x + target.width <= page_size.0
                    && target.y + target.height <= page_size.1
            }));
        }

        let page = PageFrame::new(1, 720.0, 540.0, Vec::new());
        let target = handout_targets(HandoutLayout::One, page_size)[0];
        let (thumbnail, fitted) = page_thumbnail_group(&page, target);
        assert_eq!(
            fitted,
            Rect {
                x: 36.0,
                y: 177.5,
                width: 468.0,
                height: 351.0,
            }
        );
        assert!(thumbnail.clip.is_some());

        let mut rules = Vec::new();
        let three_up = handout_targets(HandoutLayout::Three, page_size)[0];
        push_three_up_note_rules(&mut rules, three_up, page_size.0);
        assert_eq!(rules.len(), 5);
        assert!(
            rules
                .iter()
                .all(|element| matches!(element, PositionedElement::Line { .. }))
        );

        let mut shifted = fitted;
        shifted.x += 1.01;
        let maximum_edge_error = [
            (shifted.x - fitted.x).abs(),
            (shifted.y - fitted.y).abs(),
            ((shifted.x + shifted.width) - (fitted.x + fitted.width)).abs(),
            ((shifted.y + shifted.height) - (fitted.y + fitted.height)).abs(),
        ]
        .into_iter()
        .fold(0.0, f64::max);
        assert!(maximum_edge_error > 1.0);
    }

    #[test]
    fn two_dimensional_merge_encodes_origins_and_continuations() {
        let mut presentation = Presentation::new().expect("open bundled template");
        presentation.add_slide(0).expect("add slide");
        let mut slide = presentation.slide_mut(0).expect("borrow slide");
        let mut shape = slide
            .add_table(3, 3, Emu(0), Emu(0), Emu(300), Emu(300))
            .expect("add table");
        let mut table = shape.table_mut().expect("table handle");
        table
            .cell_mut(0, 0)
            .expect("merge origin")
            .merge_to(2, 1)
            .expect("merge cells");

        let expected = [
            [(3, 2, false, false), (3, 1, true, false)],
            [(1, 2, false, true), (1, 1, true, true)],
            [(1, 2, false, true), (1, 1, true, true)],
        ];
        for (row_index, row) in expected.into_iter().enumerate() {
            for (column_index, expected_cell) in row.into_iter().enumerate() {
                let cell = &table.table.rows[row_index].cells[column_index];
                assert_eq!(
                    (
                        cell.row_span,
                        cell.grid_span,
                        cell.horizontal_merge,
                        cell.vertical_merge,
                    ),
                    expected_cell,
                    "cell {row_index},{column_index}"
                );
            }
        }
        for row in 0..3 {
            let cell = &table.table.rows[row].cells[2];
            assert_eq!(
                (
                    cell.row_span,
                    cell.grid_span,
                    cell.horizontal_merge,
                    cell.vertical_merge,
                ),
                (1, 1, false, false),
                "outside cell {row},2"
            );
        }
    }

    #[test]
    fn merge_moves_formatted_content_to_origin_in_row_major_order() {
        let mut presentation = Presentation::new().expect("open bundled template");
        presentation.add_slide(0).expect("add slide");
        let mut slide = presentation.slide_mut(0).expect("borrow slide");
        let mut shape = slide
            .add_table(2, 2, Emu(0), Emu(0), Emu(200), Emu(200))
            .expect("add table");
        let mut table = shape.table_mut().expect("table handle");
        for (index, text) in ["one", "two", "three", "four"].into_iter().enumerate() {
            table.cell_mut(index / 2, index % 2).unwrap().set_text(text);
        }
        {
            let mut source = table.cell_mut(0, 1).unwrap();
            let mut frame = source.text_frame();
            let mut paragraph = frame.paragraph_mut(0).unwrap();
            let mut properties = CT_TextParagraphProperties::default();
            properties.level = Some(2);
            paragraph.set_properties(properties);
            let mut run = paragraph.add_run(" formatted");
            let mut properties = CT_TextCharacterProperties::default();
            properties.bold = Some(true);
            run.set_properties(properties);
        }
        table.cell_mut(1, 1).unwrap().merge_to(0, 0).unwrap();
        assert_eq!(
            table.cell(0, 0).unwrap().text(),
            "one\ntwo formatted\nthree\nfour"
        );
        let origin = &table.table.rows[0].cells[0];
        let paragraphs = origin.text_body.as_ref().unwrap().paragraphs();
        assert_eq!(paragraphs[1].properties.as_ref().unwrap().level, Some(2));
        let oxml_drawing::text::TextRun::Run(run) = &paragraphs[1].runs[1] else {
            panic!("expected regular run");
        };
        assert_eq!(run.properties.as_ref().unwrap().bold, Some(true));
        table.cell_mut(0, 0).unwrap().split().unwrap();
        assert_eq!(
            table.cell(0, 0).unwrap().text(),
            "one\ntwo formatted\nthree\nfour"
        );
        assert_eq!(table.cell(0, 1).unwrap().text(), "");
    }

    #[test]
    fn merge_materializes_minimal_text_bodies_for_absent_sources() {
        let mut presentation = Presentation::new().expect("open bundled template");
        presentation.add_slide(0).expect("add slide");
        let mut slide = presentation.slide_mut(0).expect("borrow slide");
        let mut shape = slide
            .add_table(2, 2, Emu(0), Emu(0), Emu(200), Emu(200))
            .expect("add table");
        let mut table = shape.table_mut().expect("table handle");
        for (row, column) in [(0, 1), (1, 0), (1, 1)] {
            table.table.rows[row].cells[column].text_body = None;
        }

        table.cell_mut(0, 0).unwrap().merge_to(1, 1).unwrap();

        for (row, column) in [(0, 1), (1, 0), (1, 1)] {
            let body = table.table.rows[row].cells[column]
                .text_body
                .as_ref()
                .expect("merge source receives a text body");
            assert_eq!(body.paragraph_count(), 1);
            assert_eq!(body.plain_text(), "");
        }
    }

    #[test]
    fn connector_constructor_normalizes_every_direction() {
        assert_eq!(ConnectorType::Straight.preset_name(), "line");
        assert_eq!(ConnectorType::Elbow.preset_name(), "bentConnector3");
        assert_eq!(ConnectorType::Curve.preset_name(), "curvedConnector3");
        let cases = [
            ((10, 20), (30, 40), (10, 20, 20, 20, false, false)),
            ((30, 20), (10, 40), (10, 20, 20, 20, true, false)),
            ((10, 40), (30, 20), (10, 20, 20, 20, false, true)),
            ((30, 40), (10, 20), (10, 20, 20, 20, true, true)),
            ((10, 20), (10, 40), (10, 20, 0, 20, false, false)),
            ((30, 20), (10, 20), (10, 20, 20, 0, true, false)),
        ];
        for (begin, end, expected) in cases {
            let transform =
                connector_transform(Emu(begin.0), Emu(begin.1), Emu(end.0), Emu(end.1)).unwrap();
            let offset = transform.offset.unwrap();
            let extent = transform.extent.unwrap();
            assert_eq!(
                (
                    offset.x.0,
                    offset.y.0,
                    extent.cx.0,
                    extent.cy.0,
                    transform.flip_horizontal,
                    transform.flip_vertical,
                ),
                expected
            );
        }

        assert!(connector_transform(Emu(i64::MIN), Emu(0), Emu(i64::MAX), Emu(0)).is_err());
    }

    #[test]
    fn every_validation_issue_variant_detects_its_corrupted_deck() {
        let mut duplicate_shape_package = package_after_save(presentation_with_slide());
        let slide_part = "/ppt/slides/slide1.xml";
        let slide_xml = String::from_utf8(
            duplicate_shape_package
                .get_part(slide_part)
                .unwrap()
                .to_vec(),
        )
        .unwrap();
        let duplicate_shape_xml = slide_xml.replacen(r#"id="1""#, r#"id="2""#, 1);
        assert_ne!(duplicate_shape_xml, slide_xml);
        duplicate_shape_package.set_part(slide_part, duplicate_shape_xml.into_bytes());
        let duplicate_shape = presentation_from_package(duplicate_shape_package);
        assert_has_issue(&duplicate_shape, |issue| {
            matches!(issue, ValidationIssue::DuplicateShapeId { slide: 0, id: 2 })
        });

        let mut out_of_range = presentation_with_slide();
        out_of_range.presentation.slide_ids[0].id = 1;
        assert_has_issue(&out_of_range, |issue| {
            matches!(
                issue,
                ValidationIssue::SlideIdOutOfRange { slide: 0, id: 1 }
            )
        });

        let mut duplicate_slide = presentation_with_slide();
        duplicate_slide.add_slide(0).unwrap();
        let duplicate_id = duplicate_slide.presentation.slide_ids[0].id;
        duplicate_slide.presentation.slide_ids[1].id = duplicate_id;
        assert_has_issue(
            &duplicate_slide,
            |issue| matches!(issue, ValidationIssue::DuplicateSlideId { id } if *id == duplicate_id),
        );

        let mut missing_content_type = Presentation::new().unwrap();
        missing_content_type
            .package
            .set_part("/ppt/untyped/payload.unknown", vec![1]);
        assert_has_issue(&missing_content_type, |issue| {
            matches!(
                issue,
                ValidationIssue::MissingContentTypeOverride { part }
                    if part == "/ppt/untyped/payload.unknown"
            )
        });

        let mut dangling_relationship_package = package_after_save(presentation_with_slide());
        let slide_part = "/ppt/slides/slide1.xml";
        let slide_xml = String::from_utf8(
            dangling_relationship_package
                .get_part(slide_part)
                .unwrap()
                .to_vec(),
        )
        .unwrap();
        let dangling_relationship_xml = slide_xml.replacen(
            "</p:sld>",
            r#"<x:probe xmlns:x="urn:validation" r:embed="rId999"/></p:sld>"#,
            1,
        );
        assert_ne!(dangling_relationship_xml, slide_xml);
        dangling_relationship_package.set_part(slide_part, dangling_relationship_xml.into_bytes());
        dangling_relationship_package.part_rels.remove(slide_part);
        let dangling_relationship = presentation_from_package(dangling_relationship_package);
        assert_has_issue(&dangling_relationship, |issue| {
            matches!(
                issue,
                ValidationIssue::DanglingRelationship { r_id, .. } if r_id == "rId999"
            )
        });

        let mut unreachable = presentation_with_slide();
        let slide_part = unreachable.slides[0].part_name.clone();
        unreachable
            .package
            .get_or_create_part_rels(&slide_part)
            .add(rel_types::IMAGE, "../media/missing.png");
        assert_has_issue(&unreachable, |issue| {
            matches!(
                issue,
                ValidationIssue::UnreachableRelationshipTarget { target, .. }
                    if target == "/ppt/media/missing.png"
            )
        });

        let mut empty_text_package = package_after_save(presentation_with_slide());
        let slide_part = "/ppt/slides/slide1.xml";
        let slide_xml =
            String::from_utf8(empty_text_package.get_part(slide_part).unwrap().to_vec()).unwrap();
        let empty_text_xml = slide_xml.replacen("<a:p/>", "", 1);
        assert_ne!(empty_text_xml, slide_xml);
        empty_text_package.set_part(slide_part, empty_text_xml.into_bytes());
        let empty_text = presentation_from_package(empty_text_package);
        assert_has_issue(&empty_text, |issue| {
            matches!(issue, ValidationIssue::EmptyTextBody { slide: 0, shape: 0 })
        });

        let mut duplicate_placeholder = presentation_with_slide();
        duplicate_placeholder.slides[0]
            .slide
            .common_slide_data
            .shape_tree
            .children
            .push(ShapeTreeChild::Shape(
                CT_Shape::new_placeholder(100, CT_Placeholder::new(Some(PhType::Body), Some(999)))
                    .unwrap(),
            ));
        duplicate_placeholder.slides[0]
            .slide
            .common_slide_data
            .shape_tree
            .children
            .push(ShapeTreeChild::Shape(
                CT_Shape::new_placeholder(
                    101,
                    CT_Placeholder::new(Some(PhType::Object), Some(999)),
                )
                .unwrap(),
            ));
        assert_has_issue(&duplicate_placeholder, |issue| {
            matches!(
                issue,
                ValidationIssue::DuplicatePlaceholderIdx { slide: 0, idx: 999 }
            )
        });

        let mut orphan_media = Presentation::new().unwrap();
        orphan_media
            .package
            .set_part("/ppt/media/orphan.png", b"orphan".to_vec());
        orphan_media
            .package
            .content_types
            .add_default("png", "image/png");
        assert_has_issue(&orphan_media, |issue| {
            matches!(
                issue,
                ValidationIssue::OrphanMedia { part } if part == "/ppt/media/orphan.png"
            )
        });

        let custom_show = presentation_with_custom_show("rId999");
        assert_has_issue(&custom_show, |issue| {
            matches!(
                issue,
                ValidationIssue::CustomShowReference { slide_id: 999 }
            )
        });

        let mut missing_layout = presentation_with_slide();
        let slide_part = missing_layout.slides[0].part_name.clone();
        missing_layout.package.part_rels.remove(&slide_part);
        assert_has_issue(&missing_layout, |issue| {
            matches!(issue, ValidationIssue::MissingLayoutRel { slide: 0 })
        });

        let mut missing_theme = Presentation::new().unwrap();
        let presentation_relationships = missing_theme
            .package
            .get_part_rels(&missing_theme.presentation_part)
            .unwrap();
        let master_relationship = presentation_relationships
            .get_by_id(&missing_theme.presentation.slide_master_ids[0].relationship_id)
            .unwrap();
        let master_part = OpcPackage::resolve_rel_target(
            &missing_theme.presentation_part,
            &master_relationship.target,
        );
        missing_theme
            .package
            .get_or_create_part_rels(&master_part)
            .items
            .retain(|relationship| relationship.rel_type != rel_types::THEME);
        assert_has_issue(&missing_theme, |issue| {
            matches!(issue, ValidationIssue::MissingThemeRel { master: 0 })
        });
    }

    #[test]
    fn validate_collects_all_issues_in_deterministic_order() {
        let mut presentation = presentation_with_slide();
        presentation.presentation.slide_ids[0].id = 1;
        presentation
            .package
            .set_part("/ppt/media/orphan.png", b"orphan".to_vec());
        presentation
            .package
            .content_types
            .add_default("png", "image/png");
        let first = presentation.validate();
        let second = presentation.validate();

        assert_eq!(first, second);
        assert!(matches!(
            first.first(),
            Some(ValidationIssue::SlideIdOutOfRange { .. })
        ));
        assert!(matches!(
            first.last(),
            Some(ValidationIssue::OrphanMedia { .. })
        ));
    }

    #[test]
    #[cfg(debug_assertions)]
    fn debug_save_boundaries_assert_on_invalid_presentations() {
        let mut presentation = presentation_with_slide();
        presentation.presentation.slide_ids[0].id = 1;
        let output = std::env::temp_dir().join(format!(
            "rpptx-f108-invalid-save-{}.pptx",
            std::process::id()
        ));

        assert!(std::panic::catch_unwind(|| presentation.to_bytes()).is_err());
        assert!(std::panic::catch_unwind(|| presentation.save(&output)).is_err());
        assert!(!output.exists());
    }

    #[test]
    fn validation_xml_scan_is_prefix_tolerant_and_non_mutating() {
        let presentation = presentation_with_slide();
        let relationship_id = presentation.presentation.slide_ids[0]
            .relationship_id
            .clone();
        let presentation = presentation_with_custom_show(&relationship_id);
        let before = presentation
            .package
            .get_part(&presentation.presentation_part)
            .unwrap()
            .to_vec();

        assert!(presentation.validate().is_empty());
        assert_eq!(
            presentation
                .package
                .get_part(&presentation.presentation_part)
                .unwrap(),
            before
        );
    }

    fn presentation_with_slide() -> Presentation {
        let mut presentation = Presentation::new().unwrap();
        presentation.add_slide(0).unwrap();
        presentation
    }

    fn package_after_save(presentation: Presentation) -> OpcPackage {
        OpcPackage::from_reader(Cursor::new(presentation.to_bytes().unwrap())).unwrap()
    }

    #[test]
    fn presentation_main_part_and_relationship_owner_resolve_case_equivalent_spelling() {
        let mut package = package_after_save(presentation_with_slide());
        let presentation_xml = package.parts.remove("/ppt/presentation.xml").unwrap();
        package
            .parts
            .insert("/PPT/PRESENTATION.XML".to_owned(), presentation_xml);
        let relationships = package.part_rels.remove("/ppt/presentation.xml").unwrap();
        package
            .part_rels
            .insert("/PPT/PRESENTATION.XML".to_owned(), relationships);

        let presentation = Presentation::from_package(package).unwrap();
        let saved = package_after_save(presentation);
        assert!(saved.parts.contains_key("/PPT/PRESENTATION.XML"));
        assert!(saved.part_rels.contains_key("/PPT/PRESENTATION.XML"));
    }

    fn presentation_from_package(package: OpcPackage) -> Presentation {
        let mut bytes = Cursor::new(Vec::new());
        package.write_to(&mut bytes).unwrap();
        Presentation::from_bytes(bytes.get_ref()).unwrap()
    }

    fn presentation_with_custom_show(relationship_id: &str) -> Presentation {
        let mut package = package_after_save(presentation_with_slide());
        let presentation_part = "/ppt/presentation.xml";
        let xml = String::from_utf8(package.get_part(presentation_part).unwrap().to_vec()).unwrap();
        let custom_show = format!(
            r#"<q:custShowLst xmlns:q="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><q:custShow name="validation"><q:sldLst><q:sld rel:id="{relationship_id}"/></q:sldLst></q:custShow></q:custShowLst>"#
        );
        package.set_part(
            presentation_part,
            xml.replacen(
                "</p:presentation>",
                &format!("{custom_show}</p:presentation>"),
                1,
            )
            .into_bytes(),
        );
        presentation_from_package(package)
    }

    fn assert_has_issue(presentation: &Presentation, predicate: impl Fn(&ValidationIssue) -> bool) {
        let issues = presentation.validate();
        assert!(issues.iter().any(predicate), "issues: {issues:?}");
    }

    fn layout_index(presentation: &Presentation, name: &str) -> usize {
        (0..presentation.layout_count())
            .find(|index| presentation.layout_name(*index) == Some(name))
            .unwrap_or_else(|| panic!("missing layout {name}"))
    }

    #[test]
    fn three_added_slides_have_unique_ids_and_reopen() {
        let mut presentation = Presentation::new().unwrap();
        for _ in 0..3 {
            presentation.add_slide(0).unwrap();
        }

        let ids = presentation
            .slides()
            .map(|slide| slide.id())
            .collect::<HashSet<_>>();
        assert_eq!(ids.len(), 3);
        assert!(ids.iter().all(|id| *id >= 256));
        assert_eq!(
            Presentation::from_bytes(&presentation.to_bytes().unwrap())
                .unwrap()
                .len(),
            3
        );
    }

    #[test]
    fn add_slide_allocates_after_the_highest_existing_part_suffix() {
        let mut presentation = Presentation::new().unwrap();
        let sentinel = b"existing sparse slide".to_vec();
        presentation
            .package
            .set_part("/ppt/slides/slide7.xml", sentinel.clone());

        presentation.add_slide(0).unwrap();

        assert_eq!(
            presentation.slides.last().unwrap().part_name,
            "/ppt/slides/slide8.xml"
        );
        assert_eq!(
            presentation.package.get_part("/ppt/slides/slide7.xml"),
            Some(sentinel.as_slice())
        );
    }

    #[test]
    fn add_slide_synthesizes_only_non_latent_layout_placeholders() {
        let mut presentation = Presentation::new().unwrap();
        let layout_index = layout_index(&presentation, "Title and Content");
        let expected = presentation.layouts[layout_index]
            .layout
            .common_slide_data
            .shape_tree
            .children
            .iter()
            .filter_map(|child| match child {
                ShapeTreeChild::Shape(shape) => shape.placeholder.as_ref(),
                _ => None,
            })
            .filter(|placeholder| {
                !matches!(
                    placeholder.ph_type,
                    Some(PhType::DateTime | PhType::Footer | PhType::SlideNumber)
                )
            })
            .map(|placeholder| (placeholder.ph_type.clone(), placeholder.idx))
            .collect::<Vec<_>>();

        presentation.add_slide(layout_index).unwrap();

        let actual = presentation.slides[0]
            .slide
            .common_slide_data
            .shape_tree
            .children
            .iter()
            .filter_map(|child| match child {
                ShapeTreeChild::Shape(shape) => shape.placeholder.as_ref(),
                _ => None,
            })
            .map(|placeholder| (placeholder.ph_type.clone(), placeholder.idx))
            .collect::<Vec<_>>();
        assert_eq!(actual, expected);
    }

    #[test]
    fn synthesized_slide_uses_schema_order_and_one_relative_layout_relationship() {
        let mut presentation = Presentation::new().unwrap();
        let relocated_layout = "/custom/layouts/relocated.xml";
        let original_layout = presentation.layouts[0].part_name.clone();
        let layout_xml = presentation
            .package
            .get_part(&original_layout)
            .unwrap()
            .to_vec();
        presentation.package.set_part(relocated_layout, layout_xml);
        presentation.layouts[0].part_name = relocated_layout.to_owned();
        presentation.add_slide(0).unwrap();
        let record = &presentation.slides[0];
        let xml = String::from_utf8(record.slide.to_xml().unwrap()).unwrap();
        let relationships = presentation
            .package
            .get_part_rels(&record.part_name)
            .unwrap();

        assert!(xml.find("<p:cSld").unwrap() < xml.find("<p:clrMapOvr").unwrap());
        assert_eq!(relationships.items.len(), 1);
        assert_eq!(relationships.items[0].rel_type, rel_types::SLIDE_LAYOUT);
        assert!(!relationships.items[0].target.starts_with('/'));
        assert_eq!(
            OpcPackage::resolve_rel_target(&record.part_name, &relationships.items[0].target),
            relocated_layout
        );
        assert_eq!(CT_Slide::from_xml(xml.as_bytes()).unwrap(), record.slide);
    }

    #[test]
    fn synthesized_text_bodies_always_contain_a_paragraph() {
        let mut presentation = Presentation::new().unwrap();
        let layout_index = layout_index(&presentation, "Title and Content");
        presentation.add_slide(layout_index).unwrap();

        for child in &presentation.slides[0]
            .slide
            .common_slide_data
            .shape_tree
            .children
        {
            if let ShapeTreeChild::Shape(shape) = child
                && shape.text_body.is_some()
            {
                assert!(
                    String::from_utf8(shape.to_xml().unwrap())
                        .unwrap()
                        .contains("<a:p")
                );
            }
        }
    }

    #[test]
    fn add_slide_rejects_an_unknown_layout_index_without_mutation() {
        let mut presentation = Presentation::new().unwrap();
        let before = presentation.to_bytes().unwrap();

        let error = presentation
            .add_slide(presentation.layout_count())
            .err()
            .expect("unknown layout must fail");

        assert!(matches!(error, Error::UnknownLayoutIndex { .. }));
        assert_eq!(presentation.to_bytes().unwrap(), before);
    }

    #[test]
    fn equal_media_bytes_inserted_twice_reuse_one_part() {
        let mut package = OpcPackage::new();
        let mut media = MediaStore::scan(&package);
        let png = b"\x89PNG\r\n\x1a\nfixture";

        let first = media.insert(&mut package, png, "first.png");
        let second = media.insert(&mut package, png, "second.png");

        assert_eq!(first, second);
        assert_eq!(
            package
                .parts
                .keys()
                .filter(|part| part.starts_with("/ppt/media/"))
                .count(),
            1
        );
    }

    #[test]
    fn media_store_compares_bytes_inside_a_hash_bucket() {
        let mut package = OpcPackage::new();
        let mut media = MediaStore::scan(&package);
        let first = b"\x89PNG\r\n\x1a\nfirst";
        let second = b"\x89PNG\r\n\x1a\nsecond";

        let first_part = media.insert_with_id(&mut package, MediaId(7), first, "first.png");
        let second_part = media.insert_with_id(&mut package, MediaId(7), second, "second.png");

        assert_ne!(first_part, second_part);
        assert_eq!(package.get_part(&first_part), Some(first.as_slice()));
        assert_eq!(package.get_part(&second_part), Some(second.as_slice()));
    }

    #[test]
    fn media_store_allocates_after_the_highest_existing_suffix() {
        let mut package = OpcPackage::new();
        package.set_part("/ppt/media/image1.png", b"\x89PNG\r\n\x1a\none".to_vec());
        package.set_part("/ppt/media/image3.png", b"\x89PNG\r\n\x1a\nthree".to_vec());
        package.content_types.add_default("png", "image/png");
        let mut media = MediaStore::scan(&package);

        let part = media.insert(&mut package, b"\x89PNG\r\n\x1a\nfour", "misleading.jpeg");

        assert_eq!(part, "/ppt/media/image4.png");
        assert_eq!(
            package.content_types.content_type_for(&part),
            Some("image/png")
        );
    }

    #[test]
    fn media_store_overrides_a_conflicting_existing_default() {
        let mut package = OpcPackage::new();
        package
            .content_types
            .add_default("png", "application/octet-stream");
        let mut media = MediaStore::scan(&package);

        let part = media.insert(&mut package, b"\x89PNG\r\n\x1a\nnew", "new.png");

        assert_eq!(
            package.content_types.content_type_for(&part),
            Some("image/png")
        );
        assert_eq!(
            package
                .content_types
                .defaults
                .get("png")
                .map(String::as_str),
            Some("application/octet-stream")
        );
    }

    #[test]
    fn media_store_reuses_the_first_sorted_duplicate_part() {
        let mut package = OpcPackage::new();
        let png = b"\x89PNG\r\n\x1a\nduplicate";
        package.set_part("/ppt/media/image2.png", png.to_vec());
        package.set_part("/ppt/media/image1.png", png.to_vec());
        let mut media = MediaStore::scan(&package);

        let part = media.insert(&mut package, png, "duplicate.png");

        assert_eq!(part, "/ppt/media/image1.png");
    }
}
