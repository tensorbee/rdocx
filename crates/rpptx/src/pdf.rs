//! Bounded PDF page import into the PresentationML facade.

use std::collections::{BTreeMap, HashMap, HashSet};
#[cfg(not(target_arch = "wasm32"))]
use std::io::{Read, Take};
#[cfg(not(target_arch = "wasm32"))]
use std::path::Path;
use std::sync::Arc;

use lopdf::content::{Content, Operation};
use lopdf::{Dictionary, Document, LoadOptions, Object, ObjectId, Stream};
use oxml_core::units::{Angle, Emu};
use oxml_drawing::fill::Fill;
use oxml_drawing::geometry::CT_CustomGeometry2D;
use oxml_drawing::line::CT_LineProperties;
use oxml_drawing::text::{CT_TextCharacterProperties, Coordinate32Value, TextFont};
use oxml_layout::{
    Color, FillRule, FontFile, FontManager, GroupElement, LayoutResult, LineCap, LineJoin, MediaId,
    PageFrame, Paint, Path as LayoutPath, PathCommand, PathElement, Point, PositionedElement, Rect,
    Stroke, Transform,
};

use super::*;

const EMU_PER_POINT: f64 = 12_700.0;
const MAX_PDF_GRAPH_DEPTH: usize = 128;

/// Selects the bounded PDF projection.
#[derive(Clone, Copy, Debug, PartialEq)]
pub enum PdfImportMode {
    /// Render supported page content through the shared rasterizer and insert one picture.
    PreservedGraphic { dpi: f64 },
    /// Project supported text, images, paths, and links as editable slide content.
    Editable,
}

/// Hard resource limits for one PDF import.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct PdfImportLimits {
    pub max_input_bytes: u64,
    pub max_pages: usize,
    pub max_objects: usize,
    pub max_decompressed_bytes: usize,
    pub max_operations_per_page: usize,
    pub max_pixels_per_page: u64,
    pub max_shapes_per_page: usize,
    pub max_diagnostics: usize,
}

impl Default for PdfImportLimits {
    fn default() -> Self {
        Self {
            max_input_bytes: 64 * 1024 * 1024,
            max_pages: 256,
            max_objects: 250_000,
            max_decompressed_bytes: 128 * 1024 * 1024,
            max_operations_per_page: 1_000_000,
            max_pixels_per_page: 100_000_000,
            max_shapes_per_page: 100_000,
            max_diagnostics: 10_000,
        }
    }
}

/// One stable source-path diagnostic for unsupported but safely skippable content.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct PdfImportDiagnostic {
    pub path: String,
    pub message: String,
}

/// A complete transactional PDF import.
pub struct PdfImportResult {
    pub presentation: Presentation,
    pub diagnostics: Vec<PdfImportDiagnostic>,
}

#[derive(Clone, Copy, Debug, PartialEq)]
struct PageGeometry {
    left: f64,
    bottom: f64,
    right: f64,
    top: f64,
    rotation: i32,
    width: f64,
    height: f64,
}

#[derive(Clone, Copy, Debug, PartialEq)]
struct Matrix {
    a: f64,
    b: f64,
    c: f64,
    d: f64,
    e: f64,
    f: f64,
}

impl Matrix {
    const IDENTITY: Self = Self {
        a: 1.0,
        b: 0.0,
        c: 0.0,
        d: 1.0,
        e: 0.0,
        f: 0.0,
    };

    fn then(self, next: Self) -> Self {
        Self {
            a: next.a * self.a + next.c * self.b,
            b: next.b * self.a + next.d * self.b,
            c: next.a * self.c + next.c * self.d,
            d: next.b * self.c + next.d * self.d,
            e: next.a * self.e + next.c * self.f + next.e,
            f: next.b * self.e + next.d * self.f + next.f,
        }
    }

    fn point(self, x: f64, y: f64) -> Point {
        Point {
            x: self.a * x + self.c * y + self.e,
            y: self.b * x + self.d * y + self.f,
        }
    }

    fn translated(self, x: f64, y: f64) -> Self {
        Self {
            e: self.e + self.a * x + self.c * y,
            f: self.f + self.b * x + self.d * y,
            ..self
        }
    }
}

impl PageGeometry {
    fn normalize(self, point: Point) -> Point {
        match self.rotation {
            0 => Point {
                x: point.x - self.left,
                y: self.top - point.y,
            },
            90 => Point {
                x: point.y - self.bottom,
                y: point.x - self.left,
            },
            180 => Point {
                x: self.right - point.x,
                y: point.y - self.bottom,
            },
            270 => Point {
                x: self.top - point.y,
                y: self.right - point.x,
            },
            _ => unreachable!("rotation is normalized"),
        }
    }

    fn rect(self, values: [f64; 4]) -> Rect {
        let first = self.normalize(Point {
            x: values[0],
            y: values[1],
        });
        let second = self.normalize(Point {
            x: values[2],
            y: values[3],
        });
        Rect {
            x: first.x.min(second.x),
            y: first.y.min(second.y),
            width: (first.x - second.x).abs(),
            height: (first.y - second.y).abs(),
        }
    }
}

#[derive(Clone)]
enum EditableElement {
    Text {
        rect: Rect,
        transform: Transform,
        text: String,
        family: String,
        size: f64,
        color: Color,
    },
    Path(PathElement),
    Image {
        rect: Rect,
        transform: Transform,
        data: Vec<u8>,
        filename: String,
    },
}

struct NormalizedPage {
    number: u32,
    geometry: PageGeometry,
    layout: Vec<PositionedElement>,
    editable: Vec<EditableElement>,
    links: Vec<(Rect, String)>,
    image_pixels: u64,
    retained_elements: usize,
}

#[derive(Clone)]
struct PageFont {
    display_family: String,
    shaping_family: String,
    encoding: String,
    encoding_supported: bool,
    decoded: BTreeMap<Vec<u8>, String>,
    widths: Option<SimpleFontWidths>,
    composite_widths_supported: bool,
}

#[derive(Clone)]
struct SimpleFontWidths {
    first_char: i64,
    widths: Vec<f64>,
    missing_width: f64,
}

struct EmbeddedFonts {
    files: Vec<FontFile>,
    retained_bytes: usize,
    widths: HashMap<ObjectId, Option<SimpleFontWidths>>,
}

enum CachedFontEncoding {
    Unicode(lopdf::Encoding<'static>),
    SingleByte(Vec<String>),
}

#[derive(Default)]
struct FontEncodingCaches {
    to_unicode_by_stream: HashMap<ObjectId, Arc<CachedFontEncoding>>,
    ordinary_by_font: HashMap<ObjectId, Arc<CachedFontEncoding>>,
    to_unicode_parses: usize,
}

impl CachedFontEncoding {
    fn decode(&self, bytes: &[u8]) -> String {
        match self {
            Self::SingleByte(values) => bytes
                .iter()
                .filter_map(|byte| values.get(*byte as usize))
                .cloned()
                .collect(),
            Self::Unicode(encoding) => Document::decode_text(encoding, bytes)
                .unwrap_or_else(|_| String::from_utf8_lossy(bytes).into_owned()),
        }
    }
}

#[derive(Clone, Debug, PartialEq)]
struct DashState {
    pattern: Vec<f64>,
    phase: f64,
}

#[derive(Clone)]
struct GraphicsState {
    ctm: Matrix,
    fill: Color,
    stroke: Color,
    line_width: f64,
    line_cap: LineCap,
    line_join: LineJoin,
    dash: Option<DashState>,
    text_matrix: Matrix,
    line_matrix: Matrix,
    font_resource: Option<Vec<u8>>,
    font_size: f64,
    character_spacing: f64,
    word_spacing: f64,
    horizontal_scale: f64,
    leading: f64,
    rise: f64,
    in_text: bool,
    text_visible: bool,
    text_metrics_supported: bool,
    text_position_supported: bool,
    fill_supported: bool,
    stroke_supported: bool,
    dash_supported: bool,
    graphics_supported: bool,
}

impl Default for GraphicsState {
    fn default() -> Self {
        Self {
            ctm: Matrix::IDENTITY,
            fill: Color::BLACK,
            stroke: Color::BLACK,
            line_width: 1.0,
            line_cap: LineCap::Butt,
            line_join: LineJoin::Miter,
            dash: None,
            text_matrix: Matrix::IDENTITY,
            line_matrix: Matrix::IDENTITY,
            font_resource: None,
            font_size: 12.0,
            character_spacing: 0.0,
            word_spacing: 0.0,
            horizontal_scale: 1.0,
            leading: 0.0,
            rise: 0.0,
            in_text: false,
            text_visible: true,
            text_metrics_supported: true,
            text_position_supported: true,
            fill_supported: true,
            stroke_supported: true,
            dash_supported: true,
            graphics_supported: true,
        }
    }
}

struct Importer<'a> {
    document: Document,
    limits: PdfImportLimits,
    mode: PdfImportMode,
    diagnostics: Vec<PdfImportDiagnostic>,
    decompressed: usize,
    fonts: FontManager,
    embedded_font_aliases: HashSet<String>,
    encoding_caches: FontEncodingCaches,
    font_widths: HashMap<ObjectId, Option<SimpleFontWidths>>,
    _bytes: &'a [u8],
}

impl Presentation {
    /// Imports one complete PDF byte slice with default hard limits.
    pub fn from_pdf_bytes(bytes: &[u8], mode: PdfImportMode) -> Result<PdfImportResult> {
        Self::from_pdf_bytes_with_limits(bytes, mode, PdfImportLimits::default())
    }

    /// Imports one complete PDF byte slice with caller-selected hard limits.
    pub fn from_pdf_bytes_with_limits(
        bytes: &[u8],
        mode: PdfImportMode,
        limits: PdfImportLimits,
    ) -> Result<PdfImportResult> {
        validate_limits(bytes, mode, limits)?;
        let options = LoadOptions {
            strict: true,
            max_decompressed_size: Some(limits.max_decompressed_bytes),
            ..LoadOptions::default()
        };
        let document = Document::load_mem_with_options(bytes, options).map_err(|error| {
            import_error(None, None, format!("strict PDF parse failed: {error}"))
        })?;
        if document.trailer.has(b"Encrypt") {
            return Err(import_error(
                None,
                None,
                "encrypted PDF input is unsupported",
            ));
        }
        if document.objects.len() > limits.max_objects {
            return Err(import_error(
                None,
                None,
                format!(
                    "PDF object count {} exceeds limit {}",
                    document.objects.len(),
                    limits.max_objects
                ),
            ));
        }
        reject_active_content(&document, limits.max_objects)?;
        let fonts = FontManager::new_deterministic().map_err(|error| {
            import_error(
                None,
                None,
                format!("deterministic font manager failed: {error}"),
            )
        })?;
        Importer {
            document,
            limits,
            mode,
            diagnostics: Vec::new(),
            decompressed: 0,
            fonts,
            embedded_font_aliases: HashSet::new(),
            encoding_caches: FontEncodingCaches::default(),
            font_widths: HashMap::new(),
            _bytes: bytes,
        }
        .import()
    }

    /// Opens and imports one PDF path without reading beyond the input limit.
    #[cfg(not(target_arch = "wasm32"))]
    pub fn open_pdf<P: AsRef<Path>>(path: P, mode: PdfImportMode) -> Result<PdfImportResult> {
        let limits = PdfImportLimits::default();
        let file = std::fs::File::open(path).map_err(OpcError::from)?;
        let length = file.metadata().map_err(OpcError::from)?.len();
        if length > limits.max_input_bytes {
            return Err(import_error(
                None,
                None,
                format!(
                    "PDF input length {length} exceeds limit {}",
                    limits.max_input_bytes
                ),
            ));
        }
        let mut reader: Take<_> = file.take(limits.max_input_bytes.saturating_add(1));
        let mut bytes = Vec::with_capacity(usize::try_from(length).unwrap_or(0));
        reader.read_to_end(&mut bytes).map_err(OpcError::from)?;
        Self::from_pdf_bytes_with_limits(&bytes, mode, limits)
    }
}

impl Importer<'_> {
    fn import(mut self) -> Result<PdfImportResult> {
        let pages = strict_pages(
            &self.document,
            self.limits.max_pages,
            self.limits.max_objects,
        )?;
        for &(number, id) in &pages {
            page_geometry(&self.document, id, number)?;
        }
        let page_resources = pages
            .iter()
            .map(|&(number, id)| {
                resolve_page_resources(&self.document, id, number, self.limits.max_objects)
            })
            .collect::<Result<Vec<_>>>()?;

        let embedded = embedded_fonts(
            &mut self.document,
            page_resources.iter().map(|resources| &resources.fonts),
            self.limits.max_decompressed_bytes,
            self.limits.max_objects,
        )?;
        self.decompressed = embedded.retained_bytes;
        self.embedded_font_aliases = embedded
            .files
            .iter()
            .map(|font| font.family.clone())
            .collect();
        self.font_widths = embedded.widths;
        self.fonts.load_additional_fonts(&embedded.files);
        let mut normalized = Vec::with_capacity(pages.len());
        for ((number, id), resources) in pages.into_iter().zip(&page_resources) {
            normalized.push(self.parse_page(number, id, resources)?);
        }
        let first = normalized[0].geometry;
        for page in normalized.iter().skip(1) {
            if (page.geometry.width - first.width).abs() > 0.000_001
                || (page.geometry.height - first.height).abs() > 0.000_001
            {
                return Err(import_error(
                    Some(page.number),
                    None,
                    format!(
                        "effective page size {} by {} differs from first page {} by {}",
                        page.geometry.width, page.geometry.height, first.width, first.height
                    ),
                ));
            }
        }
        let presentation = self.project(&normalized)?;
        Ok(PdfImportResult {
            presentation,
            diagnostics: self.diagnostics,
        })
    }

    fn parse_page(
        &mut self,
        number: u32,
        id: ObjectId,
        resources: &ResolvedPageResources,
    ) -> Result<NormalizedPage> {
        let geometry = page_geometry(&self.document, id, number)?;
        let remaining = self.remaining_decompressed(Some(number), None)?;
        let data = strict_page_content(
            &self.document,
            id,
            number,
            self.limits.max_objects,
            remaining,
        )?;
        self.charge_decompressed(data.len(), Some(number), None, "page content")?;
        let sanitized = sanitize_content_comments(&data);
        let content = Content::decode_strict(&sanitized).map_err(|error| {
            import_error(
                Some(number),
                None,
                format!("strict content decode failed: {error}"),
            )
        })?;
        if content.operations.len() > self.limits.max_operations_per_page {
            return Err(import_error(
                Some(number),
                None,
                format!(
                    "page operation count {} exceeds limit {}",
                    content.operations.len(),
                    self.limits.max_operations_per_page
                ),
            ));
        }
        let fonts = page_fonts(
            &self.document,
            &mut self.encoding_caches,
            &resources.fonts,
            &content.operations,
            &self.font_widths,
        )?;
        let mut page = NormalizedPage {
            number,
            geometry,
            layout: Vec::new(),
            editable: Vec::new(),
            links: Vec::new(),
            image_pixels: 0,
            retained_elements: 0,
        };
        self.interpret(&mut page, &content.operations, &fonts, &resources.xobjects)?;
        self.parse_annotations(&mut page, id)?;
        if page.retained_elements > self.limits.max_shapes_per_page {
            return Err(import_error(
                Some(number),
                None,
                format!(
                    "page shape count exceeds limit {}",
                    self.limits.max_shapes_per_page
                ),
            ));
        }
        Ok(page)
    }

    fn diagnostic(&mut self, page: u32, offset: usize, message: impl Into<String>) -> Result<()> {
        if self.diagnostics.len() >= self.limits.max_diagnostics {
            return Err(import_error(
                Some(page),
                Some(offset),
                format!(
                    "diagnostic count exceeds limit {}",
                    self.limits.max_diagnostics
                ),
            ));
        }
        self.diagnostics.push(PdfImportDiagnostic {
            path: format!("pdf/page[{page}]/content/op[{offset}]"),
            message: message.into(),
        });
        Ok(())
    }

    fn remaining_decompressed(&self, page: Option<u32>, offset: Option<usize>) -> Result<usize> {
        let remaining = self
            .limits
            .max_decompressed_bytes
            .checked_sub(self.decompressed)
            .ok_or_else(|| {
                import_error(
                    page,
                    offset,
                    "aggregate decompressed byte budget is exhausted",
                )
            })?;
        if remaining == 0 {
            return Err(import_error(
                page,
                offset,
                "aggregate decompressed byte budget is exhausted",
            ));
        }
        Ok(remaining)
    }

    fn charge_decompressed(
        &mut self,
        bytes: usize,
        page: Option<u32>,
        offset: Option<usize>,
        source: &str,
    ) -> Result<()> {
        self.decompressed = self.decompressed.checked_add(bytes).ok_or_else(|| {
            import_error(page, offset, "aggregate decompressed byte count overflow")
        })?;
        if self.decompressed > self.limits.max_decompressed_bytes {
            return Err(import_error(
                page,
                offset,
                format!(
                    "aggregate decompressed bytes after {source} total {}, exceeding limit {}",
                    self.decompressed, self.limits.max_decompressed_bytes
                ),
            ));
        }
        Ok(())
    }

    fn interpret(
        &mut self,
        page: &mut NormalizedPage,
        operations: &[Operation],
        fonts: &BTreeMap<Vec<u8>, PageFont>,
        xobjects: &BTreeMap<Vec<u8>, ObjectId>,
    ) -> Result<()> {
        let mut state = GraphicsState::default();
        let mut stack = Vec::new();
        let mut path = Vec::<PathCommand>::new();
        let mut current: Option<Point> = None;
        let mut subpath: Option<Point> = None;

        for (offset, operation) in operations.iter().enumerate() {
            validate_operator_arity(operation, page.number, offset)?;
            match operation.operator.as_str() {
                "q" => stack.push(state.clone()),
                "Q" => {
                    state = stack.pop().ok_or_else(|| {
                        import_error(Some(page.number), Some(offset), "unbalanced Q operator")
                    })?;
                }
                "cm" => {
                    let values = numbers(operation, 6, page.number, offset)?;
                    let next = Matrix {
                        a: values[0],
                        b: values[1],
                        c: values[2],
                        d: values[3],
                        e: values[4],
                        f: values[5],
                    };
                    state.ctm = checked_matrix(
                        next.then(state.ctm),
                        page.number,
                        offset,
                        "concatenated CTM",
                    )?;
                }
                "m" => {
                    let p = transformed_point(
                        operation,
                        &state,
                        page.geometry,
                        page.number,
                        offset,
                        0,
                    )?;
                    path.push(PathCommand::MoveTo(p));
                    current = Some(p);
                    subpath = Some(p);
                }
                "l" => {
                    require_current_point(current, page.number, offset, "l")?;
                    let p = transformed_point(
                        operation,
                        &state,
                        page.geometry,
                        page.number,
                        offset,
                        0,
                    )?;
                    path.push(PathCommand::LineTo(p));
                    current = Some(p);
                }
                "c" => {
                    require_current_point(current, page.number, offset, "c")?;
                    let values = numbers(operation, 6, page.number, offset)?;
                    let c1 = checked_point(
                        page.geometry
                            .normalize(state.ctm.point(values[0], values[1])),
                        page.number,
                        offset,
                        "curve control point",
                    )?;
                    let c2 = checked_point(
                        page.geometry
                            .normalize(state.ctm.point(values[2], values[3])),
                        page.number,
                        offset,
                        "curve control point",
                    )?;
                    let to = checked_point(
                        page.geometry
                            .normalize(state.ctm.point(values[4], values[5])),
                        page.number,
                        offset,
                        "curve endpoint",
                    )?;
                    path.push(PathCommand::CurveTo { c1, c2, to });
                    current = Some(to);
                }
                "v" => {
                    let values = numbers(operation, 4, page.number, offset)?;
                    let c1 = current.ok_or_else(|| {
                        import_error(Some(page.number), Some(offset), "v has no current point")
                    })?;
                    let c2 = checked_point(
                        page.geometry
                            .normalize(state.ctm.point(values[0], values[1])),
                        page.number,
                        offset,
                        "curve control point",
                    )?;
                    let to = checked_point(
                        page.geometry
                            .normalize(state.ctm.point(values[2], values[3])),
                        page.number,
                        offset,
                        "curve endpoint",
                    )?;
                    path.push(PathCommand::CurveTo { c1, c2, to });
                    current = Some(to);
                }
                "y" => {
                    require_current_point(current, page.number, offset, "y")?;
                    let values = numbers(operation, 4, page.number, offset)?;
                    let c1 = checked_point(
                        page.geometry
                            .normalize(state.ctm.point(values[0], values[1])),
                        page.number,
                        offset,
                        "curve control point",
                    )?;
                    let to = checked_point(
                        page.geometry
                            .normalize(state.ctm.point(values[2], values[3])),
                        page.number,
                        offset,
                        "curve endpoint",
                    )?;
                    path.push(PathCommand::CurveTo { c1, c2: to, to });
                    current = Some(to);
                }
                "h" => {
                    require_current_point(current, page.number, offset, "h")?;
                    let start = subpath.ok_or_else(|| {
                        import_error(Some(page.number), Some(offset), "h has no current subpath")
                    })?;
                    path.push(PathCommand::Close);
                    current = Some(start);
                }
                "re" => {
                    let values = numbers(operation, 4, page.number, offset)?;
                    let right = checked_sum(
                        values[0],
                        values[2],
                        page.number,
                        offset,
                        "rectangle x extent",
                    )?;
                    let top = checked_sum(
                        values[1],
                        values[3],
                        page.number,
                        offset,
                        "rectangle y extent",
                    )?;
                    let points = [
                        (values[0], values[1]),
                        (right, values[1]),
                        (right, top),
                        (values[0], top),
                    ];
                    for (index, (x, y)) in points.into_iter().enumerate() {
                        let p = checked_point(
                            page.geometry.normalize(state.ctm.point(x, y)),
                            page.number,
                            offset,
                            "rectangle corner",
                        )?;
                        path.push(if index == 0 {
                            PathCommand::MoveTo(p)
                        } else {
                            PathCommand::LineTo(p)
                        });
                    }
                    path.push(PathCommand::Close);
                    current = path.iter().rev().find_map(|command| match command {
                        PathCommand::MoveTo(point) => Some(*point),
                        _ => None,
                    });
                    subpath = current;
                }
                "S" | "s" | "f" | "F" | "B" | "b" => {
                    if matches!(operation.operator.as_str(), "s" | "b") {
                        require_current_point(current, page.number, offset, &operation.operator)?;
                        path.push(PathCommand::Close);
                    }
                    let fill = (state.graphics_supported
                        && state.fill_supported
                        && matches!(operation.operator.as_str(), "f" | "F" | "B" | "b"))
                    .then_some(Paint::Solid(state.fill));
                    let stroke = if state.graphics_supported
                        && state.stroke_supported
                        && state.dash_supported
                        && matches!(operation.operator.as_str(), "S" | "s" | "B" | "b")
                    {
                        self.transformed_stroke(page.number, offset, &state)?
                    } else {
                        None
                    };
                    self.paint_path(page, &mut path, fill, stroke)?;
                    current = None;
                    subpath = None;
                }
                "f*" | "B*" | "b*" => {
                    self.diagnostic(page.number, offset, "even-odd path painting is unsupported")?;
                    path.clear();
                    current = None;
                    subpath = None;
                }
                "n" => {
                    path.clear();
                    current = None;
                    subpath = None;
                }
                "w" => {
                    state.line_width = finite_nonnegative(
                        number_at(operation, 0, page.number, offset)?,
                        page.number,
                        offset,
                        "line width",
                    )?
                }
                "J" => {
                    state.line_cap = match integer_at(operation, 0, page.number, offset)? {
                        0 => LineCap::Butt,
                        1 => LineCap::Round,
                        2 => LineCap::Square,
                        value => {
                            return Err(import_error(
                                Some(page.number),
                                Some(offset),
                                format!("invalid line cap {value}"),
                            ));
                        }
                    }
                }
                "j" => {
                    state.line_join = match integer_at(operation, 0, page.number, offset)? {
                        0 => LineJoin::Miter,
                        1 => LineJoin::Round,
                        2 => LineJoin::Bevel,
                        value => {
                            return Err(import_error(
                                Some(page.number),
                                Some(offset),
                                format!("invalid line join {value}"),
                            ));
                        }
                    }
                }
                "d" => {
                    let dash = parse_dash(operation, page.number, offset)?;
                    if dash
                        .as_ref()
                        .is_some_and(|dash| dash.pattern.contains(&0.0))
                    {
                        self.diagnostic(
                            page.number,
                            offset,
                            "zero-length dash members are unsupported; dependent strokes are isolated until dash state reset",
                        )?;
                        state.dash = None;
                        state.dash_supported = false;
                    } else if let Some(dash) = dash {
                        if let Some(pattern) = normalized_dash(&dash.pattern, dash.phase) {
                            state.dash = Some(DashState {
                                pattern,
                                phase: 0.0,
                            });
                            state.dash_supported = true;
                        } else {
                            self.diagnostic(
                                page.number,
                                offset,
                                "dash phase requires a zero DrawingML dash stop; dependent strokes are isolated until dash state reset",
                            )?;
                            state.dash = None;
                            state.dash_supported = false;
                        }
                    } else {
                        state.dash = None;
                        state.dash_supported = true;
                    }
                }
                "g" => {
                    state.fill = gray(
                        number_at(operation, 0, page.number, offset)?,
                        page.number,
                        offset,
                    )?;
                    state.fill_supported = true;
                }
                "G" => {
                    state.stroke = gray(
                        number_at(operation, 0, page.number, offset)?,
                        page.number,
                        offset,
                    )?;
                    state.stroke_supported = true;
                }
                "rg" => {
                    state.fill = rgb(operation, page.number, offset)?;
                    state.fill_supported = true;
                }
                "RG" => {
                    state.stroke = rgb(operation, page.number, offset)?;
                    state.stroke_supported = true;
                }
                "BT" => {
                    if state.in_text {
                        return Err(import_error(
                            Some(page.number),
                            Some(offset),
                            "nested BT operator",
                        ));
                    }
                    state.in_text = true;
                    state.text_matrix = Matrix::IDENTITY;
                    state.line_matrix = Matrix::IDENTITY;
                    state.text_position_supported = true;
                }
                "ET" => {
                    if !state.in_text {
                        return Err(import_error(
                            Some(page.number),
                            Some(offset),
                            "ET without BT",
                        ));
                    }
                    state.in_text = false;
                }
                "Tf" => {
                    require_text(&state, page.number, offset)?;
                    let name = name_at(operation, 0, page.number, offset)?.to_vec();
                    state.font_resource = Some(name.clone());
                    state.font_size = finite_positive(
                        number_at(operation, 1, page.number, offset)?,
                        page.number,
                        offset,
                        "font size",
                    )?;
                    state.text_metrics_supported = fonts
                        .get(&name)
                        .is_none_or(|font| font.composite_widths_supported);
                    if !state.text_metrics_supported {
                        self.diagnostic(
                            page.number,
                            offset,
                            format!(
                                "composite font /{} widths are unsupported; dependent text is isolated",
                                String::from_utf8_lossy(&name)
                            ),
                        )?;
                    }
                }
                "Tm" => {
                    require_text(&state, page.number, offset)?;
                    let v = numbers(operation, 6, page.number, offset)?;
                    state.text_matrix = Matrix {
                        a: v[0],
                        b: v[1],
                        c: v[2],
                        d: v[3],
                        e: v[4],
                        f: v[5],
                    };
                    state.line_matrix = state.text_matrix;
                    state.text_position_supported = true;
                }
                "Td" | "TD" => {
                    require_text(&state, page.number, offset)?;
                    let v = numbers(operation, 2, page.number, offset)?;
                    if operation.operator == "TD" {
                        state.leading = -v[1];
                    }
                    state.line_matrix = checked_matrix(
                        state.line_matrix.translated(v[0], v[1]),
                        page.number,
                        offset,
                        "text line translation",
                    )?;
                    state.text_matrix = state.line_matrix;
                    state.text_position_supported = true;
                }
                "T*" => {
                    require_text(&state, page.number, offset)?;
                    state.line_matrix = checked_matrix(
                        state.line_matrix.translated(0.0, -state.leading),
                        page.number,
                        offset,
                        "text leading translation",
                    )?;
                    state.text_matrix = state.line_matrix;
                    state.text_position_supported = true;
                }
                "Tc" => {
                    require_text(&state, page.number, offset)?;
                    state.character_spacing = number_at(operation, 0, page.number, offset)?;
                }
                "Tw" => {
                    require_text(&state, page.number, offset)?;
                    state.word_spacing = number_at(operation, 0, page.number, offset)?;
                }
                "Tz" => {
                    require_text(&state, page.number, offset)?;
                    state.horizontal_scale = number_at(operation, 0, page.number, offset)? / 100.0;
                    if !state.horizontal_scale.is_finite() || state.horizontal_scale <= 0.0 {
                        return Err(import_error(
                            Some(page.number),
                            Some(offset),
                            "horizontal text scale must be finite and positive",
                        ));
                    }
                }
                "TL" => {
                    require_text(&state, page.number, offset)?;
                    state.leading = number_at(operation, 0, page.number, offset)?;
                }
                "Ts" => {
                    require_text(&state, page.number, offset)?;
                    state.rise = number_at(operation, 0, page.number, offset)?;
                }
                "Tr" => {
                    require_text(&state, page.number, offset)?;
                    let mode = integer_at(operation, 0, page.number, offset)?;
                    if !(0..=7).contains(&mode) {
                        return Err(import_error(
                            Some(page.number),
                            Some(offset),
                            format!("invalid text rendering mode {mode}"),
                        ));
                    }
                    state.text_visible = mode == 0;
                    if (4..=7).contains(&mode) {
                        state.graphics_supported = false;
                    }
                    if mode != 0 && mode != 3 {
                        self.diagnostic(
                            page.number,
                            offset,
                            format!("text rendering mode {mode} is unsupported"),
                        )?;
                    }
                }
                "Tj" => {
                    let text = string_bytes_at(operation, 0, page.number, offset)?;
                    self.show_text(page, offset, &mut state, fonts, &text)?;
                }
                "TJ" => {
                    let array = operation
                        .operands
                        .first()
                        .and_then(|value| value.as_array().ok())
                        .ok_or_else(|| {
                            import_error(Some(page.number), Some(offset), "TJ requires an array")
                        })?;
                    for item in array {
                        match item {
                            Object::String(text, _) => {
                                self.show_text(page, offset, &mut state, fonts, text)?;
                            }
                            Object::Integer(value) => {
                                let adjustment = -(*value as f64) / 1000.0
                                    * state.font_size
                                    * state.horizontal_scale;
                                advance_text_matrix(&mut state, adjustment, page.number, offset)?;
                            }
                            Object::Real(value) => {
                                let adjustment = -f64::from(*value) / 1000.0
                                    * state.font_size
                                    * state.horizontal_scale;
                                advance_text_matrix(&mut state, adjustment, page.number, offset)?;
                            }
                            _ => {
                                return Err(import_error(
                                    Some(page.number),
                                    Some(offset),
                                    "TJ array contains a non-string, non-number operand",
                                ));
                            }
                        }
                    }
                }
                "'" | "\"" => {
                    require_text(&state, page.number, offset)?;
                    if operation.operator == "\"" {
                        state.word_spacing = number_at(operation, 0, page.number, offset)?;
                        state.character_spacing = number_at(operation, 1, page.number, offset)?;
                    }
                    state.line_matrix = checked_matrix(
                        state.line_matrix.translated(0.0, -state.leading),
                        page.number,
                        offset,
                        "text leading translation",
                    )?;
                    state.text_matrix = state.line_matrix;
                    state.text_position_supported = true;
                    let index = if operation.operator == "\"" { 2 } else { 0 };
                    let text = string_bytes_at(operation, index, page.number, offset)?;
                    self.show_text(page, offset, &mut state, fonts, &text)?;
                }
                "Do" => {
                    let name = name_at(operation, 0, page.number, offset)?;
                    if !state.graphics_supported {
                        self.diagnostic(
                            page.number,
                            offset,
                            "XObject omitted while unsupported graphics state is active",
                        )?;
                    } else if let Some(id) = xobjects.get(name) {
                        self.show_xobject(page, offset, &state, *id)?;
                    } else {
                        self.diagnostic(
                            page.number,
                            offset,
                            format!("missing XObject /{}", String::from_utf8_lossy(name)),
                        )?;
                    }
                }
                "W" | "W*" => {
                    self.diagnostic(
                        page.number,
                        offset,
                        "content clipping path is unsupported and dependent content is isolated",
                    )?;
                    state.graphics_supported = false;
                }
                "CS" | "SC" | "SCN" => {
                    self.diagnostic(
                        page.number,
                        offset,
                        format!("colour operator {} is unsupported", operation.operator),
                    )?;
                    state.stroke_supported = false;
                }
                "cs" | "sc" | "scn" => {
                    self.diagnostic(
                        page.number,
                        offset,
                        format!("colour operator {} is unsupported", operation.operator),
                    )?;
                    state.fill_supported = false;
                }
                "gs" => {
                    self.diagnostic(
                        page.number,
                        offset,
                        "extended graphics state is unsupported and dependent content is isolated",
                    )?;
                    state.graphics_supported = false;
                }
                "sh" => self.diagnostic(page.number, offset, "shading is unsupported")?,
                "BI" | "ID" | "EI" => {
                    self.diagnostic(page.number, offset, "inline image is unsupported")?
                }
                other => self.diagnostic(
                    page.number,
                    offset,
                    format!("unsupported PDF operator {other}"),
                )?,
            }
        }
        if !stack.is_empty() {
            return Err(import_error(
                Some(page.number),
                None,
                "unbalanced q operator",
            ));
        }
        if state.in_text {
            return Err(import_error(
                Some(page.number),
                None,
                "unbalanced BT operator",
            ));
        }
        Ok(())
    }

    fn transformed_stroke(
        &mut self,
        page: u32,
        offset: usize,
        state: &GraphicsState,
    ) -> Result<Option<Stroke>> {
        let x_scale = state.ctm.a.hypot(state.ctm.b);
        let y_scale = state.ctm.c.hypot(state.ctm.d);
        let dot = state.ctm.a * state.ctm.c + state.ctm.b * state.ctm.d;
        checked_finite(x_scale, page, offset, "stroke x scale")?;
        checked_finite(y_scale, page, offset, "stroke y scale")?;
        checked_finite(dot, page, offset, "stroke axis product")?;
        let tolerance = x_scale.max(y_scale).max(1.0) * 1.0e-9;
        if x_scale <= f64::EPSILON
            || y_scale <= f64::EPSILON
            || dot.abs() > tolerance
            || (x_scale - y_scale).abs() > tolerance
        {
            self.diagnostic(
                page,
                offset,
                "stroked path omitted because a nonuniform or sheared CTM cannot preserve its outline",
            )?;
            return Ok(None);
        }
        let width = checked_product(state.line_width, x_scale, page, offset, "stroke width")?;
        let dash = if let Some(dash) = &state.dash {
            let mut transformed = Vec::new();
            for value in &dash.pattern {
                transformed.push(checked_product(
                    *value,
                    x_scale,
                    page,
                    offset,
                    "stroke dash length",
                )?);
            }
            if transformed
                .iter()
                .any(|value| drawingml_dash_units(*value, width).is_none())
            {
                self.diagnostic(
                    page,
                    offset,
                    "dash member is too small for a positive DrawingML dash stop; affected stroke is omitted",
                )?;
                return Ok(None);
            }
            Some(transformed)
        } else {
            None
        };
        Ok(Some(Stroke {
            paint: Paint::Solid(state.stroke),
            width,
            cap: state.line_cap,
            join: state.line_join,
            dash,
        }))
    }

    fn paint_path(
        &mut self,
        page: &mut NormalizedPage,
        commands: &mut Vec<PathCommand>,
        fill: Option<Paint>,
        stroke: Option<Stroke>,
    ) -> Result<()> {
        if commands.is_empty() {
            return Ok(());
        }
        if fill.is_none() && stroke.is_none() {
            commands.clear();
            return Ok(());
        }
        self.ensure_shape_room(page, 1)?;
        let element = PathElement {
            path: LayoutPath {
                commands: std::mem::take(commands),
                fill_rule: FillRule::NonZero,
            },
            fill,
            stroke,
        };
        let bounds = element
            .path
            .bounds()
            .ok_or_else(|| import_error(Some(page.number), None, "painted path has no bounds"))?;
        checked_rect(bounds, page.number, 0, "painted path bounds")?;
        page.layout.push(PositionedElement::Path(element.clone()));
        page.editable.push(EditableElement::Path(element));
        Ok(())
    }

    fn show_text(
        &mut self,
        page: &mut NormalizedPage,
        offset: usize,
        state: &mut GraphicsState,
        fonts: &BTreeMap<Vec<u8>, PageFont>,
        text_bytes: &[u8],
    ) -> Result<()> {
        require_text(state, page.number, offset)?;
        if !state.text_metrics_supported {
            state.text_position_supported = false;
            return Ok(());
        }
        if !state.text_position_supported {
            return Ok(());
        }
        let font = state
            .font_resource
            .as_ref()
            .and_then(|name| fonts.get(name));
        let text = if let Some(font) = font {
            if !font.encoding_supported {
                self.diagnostic(
                    page.number,
                    offset,
                    format!(
                        "font `{}` uses unsupported encoding `{}`; bytes use deterministic direct decoding",
                        font.display_family, font.encoding
                    ),
                )?;
                decode_pdf_string(text_bytes)
            } else {
                font.decoded
                    .get(text_bytes)
                    .ok_or_else(|| {
                        import_error(
                            Some(page.number),
                            Some(offset),
                            format!("font encoding `{}` was not cached", font.encoding),
                        )
                    })?
                    .clone()
            }
        } else {
            self.diagnostic(
                page.number,
                offset,
                "missing font resource; bytes use deterministic direct decoding",
            )?;
            decode_pdf_string(text_bytes)
        };
        if text.is_empty() {
            return Ok(());
        }
        let requested = font.map(|font| font.shaping_family.clone());
        let family = requested.clone().unwrap_or_else(|| "Carlito".to_owned());
        let (font_id, mut substituted) = match self.fonts.resolve_font(Some(&family), false, false)
        {
            Ok(id) => (id, false),
            Err(_) => (
                self.fonts
                    .resolve_font(Some("Carlito"), false, false)
                    .map_err(|error| {
                        import_error(
                            Some(page.number),
                            Some(offset),
                            format!("font resolution failed: {error}"),
                        )
                    })?,
                true,
            ),
        };
        let actual_family = self
            .fonts
            .font_data(font_id)
            .map_err(|error| {
                import_error(
                    Some(page.number),
                    Some(offset),
                    format!("resolved font data failed: {error}"),
                )
            })?
            .family;
        substituted |= !self.embedded_font_aliases.contains(&family);
        if requested.is_none() || substituted {
            self.diagnostic(
                page.number,
                offset,
                format!("font `{family}` substituted with deterministic `{actual_family}`"),
            )?;
        }
        let combined = checked_matrix(
            state.text_matrix.then(state.ctm),
            page.number,
            offset,
            "combined text transform",
        )?;
        let origin = checked_point(
            page.geometry.normalize(combined.point(0.0, state.rise)),
            page.number,
            offset,
            "text origin",
        )?;
        let x_axis = checked_point(
            page.geometry.normalize(combined.point(1.0, state.rise)),
            page.number,
            offset,
            "text x axis",
        )?;
        let y_axis = checked_point(
            page.geometry
                .normalize(combined.point(0.0, state.rise + 1.0)),
            page.number,
            offset,
            "text y axis",
        )?;
        let x_vector = Point {
            x: (x_axis.x - origin.x) * state.horizontal_scale,
            y: (x_axis.y - origin.y) * state.horizontal_scale,
        };
        let y_vector = Point {
            x: origin.x - y_axis.x,
            y: origin.y - y_axis.y,
        };
        let x_scale = point_distance(Point { x: 0.0, y: 0.0 }, x_vector);
        let y_scale = point_distance(Point { x: 0.0, y: 0.0 }, y_vector);
        checked_finite(x_scale, page.number, offset, "text x scale")?;
        checked_finite(y_scale, page.number, offset, "text y scale")?;
        if x_scale <= f64::EPSILON || y_scale <= f64::EPSILON {
            return Err(import_error(
                Some(page.number),
                Some(offset),
                "text transform is singular",
            ));
        }
        let mut shaped = self
            .fonts
            .shape_text(font_id, &text, state.font_size)
            .map_err(|error| {
                import_error(
                    Some(page.number),
                    Some(offset),
                    format!("text shaping failed: {error}"),
                )
            })?;
        let characters = text.chars().collect::<Vec<_>>();
        let source_advances = font
            .and_then(|font| font.widths.as_ref())
            .map(|widths| {
                text_bytes
                    .iter()
                    .map(|byte| {
                        let index = i64::from(*byte) - widths.first_char;
                        let width = usize::try_from(index)
                            .ok()
                            .and_then(|index| widths.widths.get(index))
                            .copied()
                            .unwrap_or(widths.missing_width);
                        let mut advance =
                            width / 1000.0 * state.font_size + state.character_spacing;
                        if *byte == b' ' {
                            advance += state.word_spacing;
                        }
                        checked_finite(advance, page.number, offset, "source font advance")
                    })
                    .collect::<Result<Vec<_>>>()
            })
            .transpose()?;
        let unscaled_advance = if let Some(source_advances) = source_advances {
            if source_advances.len() == shaped.advances.len() {
                shaped.advances.clone_from(&source_advances);
            }
            source_advances.iter().sum()
        } else {
            for (index, advance) in shaped.advances.iter_mut().enumerate() {
                *advance += state.character_spacing;
                if characters.get(index) == Some(&' ') {
                    *advance += state.word_spacing;
                }
            }
            let word_count = text.chars().filter(|character| *character == ' ').count() as f64;
            shaped.width
                + state.character_spacing * characters.len() as f64
                + state.word_spacing * word_count
        };
        let text_advance = checked_product(
            unscaled_advance,
            state.horizontal_scale,
            page.number,
            offset,
            "text advance",
        )?;
        let run = GlyphRun {
            origin: Point { x: 0.0, y: 0.0 },
            font_id,
            font_size: state.font_size,
            glyph_ids: shaped.glyph_ids,
            advances: shaped.advances,
            text: text.clone(),
            source: None,
            color: state.fill,
            bold: false,
            italic: false,
            field_kind: None,
            field_source: None,
            note: None,
        };
        if state.graphics_supported && state.fill_supported && state.text_visible {
            self.ensure_shape_room(page, 1)?;
            let transform = Transform {
                a: x_vector.x,
                b: x_vector.y,
                c: y_vector.x,
                d: y_vector.y,
                e: origin.x,
                f: origin.y,
            };
            page.layout.push(PositionedElement::Group(GroupElement {
                transform,
                clip: None,
                opacity: 1.0,
                effects: Vec::new(),
                children: vec![PositionedElement::Text(run)],
            }));
            let rect = Rect {
                x: 0.0,
                y: -state.font_size * 0.875,
                width: unscaled_advance.abs().max(state.font_size * 0.25),
                height: state.font_size * 1.4,
            };
            checked_rect(rect, page.number, offset, "editable text local rectangle")?;
            page.editable.push(EditableElement::Text {
                rect,
                transform,
                text,
                family: if substituted {
                    actual_family
                } else {
                    font.map_or(family, |font| font.display_family.clone())
                },
                size: state.font_size,
                color: state.fill,
            });
        }
        advance_text_matrix(state, text_advance, page.number, offset)?;
        Ok(())
    }

    fn show_xobject(
        &mut self,
        page: &mut NormalizedPage,
        offset: usize,
        state: &GraphicsState,
        id: ObjectId,
    ) -> Result<()> {
        let stream = self
            .document
            .get_object(id)
            .and_then(Object::as_stream)
            .map_err(|error| {
                import_error(
                    Some(page.number),
                    Some(offset),
                    format!("XObject failed: {error}"),
                )
            })?
            .clone();
        let subtype = stream
            .dict
            .get(b"Subtype")
            .and_then(Object::as_name)
            .unwrap_or_default();
        if subtype != b"Image" {
            self.diagnostic(page.number, offset, "form XObject recursion is unsupported")?;
            return Ok(());
        }
        let width = positive_i64(&stream.dict, b"Width", page.number, offset)?;
        let height = positive_i64(&stream.dict, b"Height", page.number, offset)?;
        let width = u32::try_from(width).map_err(|_| {
            import_error(
                Some(page.number),
                Some(offset),
                "image Width exceeds the supported integer range",
            )
        })?;
        let height = u32::try_from(height).map_err(|_| {
            import_error(
                Some(page.number),
                Some(offset),
                "image Height exceeds the supported integer range",
            )
        })?;
        let corners = image_corners(page.geometry, state.ctm, page.number, offset)?;
        let bounds = points_bounds(&corners);
        checked_rect(bounds, page.number, offset, "image bounds")?;
        if bounds.width <= f64::EPSILON || bounds.height <= f64::EPSILON {
            return Err(import_error(
                Some(page.number),
                Some(offset),
                "image transform has empty page bounds",
            ));
        }
        let transform = Transform {
            a: corners[1].x - corners[0].x,
            b: corners[1].y - corners[0].y,
            c: corners[2].x - corners[0].x,
            d: corners[2].y - corners[0].y,
            e: corners[0].x,
            f: corners[0].y,
        };
        let determinant = transform.a * transform.d - transform.b * transform.c;
        if !determinant.is_finite() || determinant.abs() <= f64::EPSILON {
            return Err(import_error(
                Some(page.number),
                Some(offset),
                "image transform is singular",
            ));
        }
        let filters = filter_names(&stream.dict)
            .map_err(|message| import_error(Some(page.number), Some(offset), message))?;
        if let Some(message) = unsupported_image_semantics(&stream.dict, &filters)
            .map_err(|message| import_error(Some(page.number), Some(offset), message))?
        {
            self.diagnostic(page.number, offset, message)?;
            return Ok(());
        }
        let intrinsic = if filters.as_slice() == ["DCTDecode"] {
            let intrinsic = jpeg_dimensions(&stream.content)
                .map_err(|message| import_error(Some(page.number), Some(offset), message))?;
            if intrinsic != (width, height) {
                return Err(import_error(
                    Some(page.number),
                    Some(offset),
                    format!(
                        "JPEG intrinsic dimensions {}x{} do not match declared image dimensions {width}x{height}",
                        intrinsic.0, intrinsic.1
                    ),
                ));
            }
            intrinsic
        } else {
            (width, height)
        };
        let pixels = u64::from(intrinsic.0)
            .checked_mul(u64::from(intrinsic.1))
            .ok_or_else(|| {
                import_error(
                    Some(page.number),
                    Some(offset),
                    "image pixel count overflow",
                )
            })?;
        page.image_pixels = page.image_pixels.checked_add(pixels).ok_or_else(|| {
            import_error(
                Some(page.number),
                Some(offset),
                "aggregate page image pixel count overflow",
            )
        })?;
        if page.image_pixels > self.limits.max_pixels_per_page {
            return Err(import_error(
                Some(page.number),
                Some(offset),
                format!(
                    "aggregate page image pixel count {} exceeds limit {}",
                    page.image_pixels, self.limits.max_pixels_per_page
                ),
            ));
        }
        let (data, filename, content_type) = if filters.as_slice() == ["DCTDecode"] {
            (
                stream.content.clone(),
                "image.jpg".to_owned(),
                "image/jpeg".to_owned(),
            )
        } else if filters.is_empty() || filters.as_slice() == ["FlateDecode"] {
            let remaining = self.remaining_decompressed(Some(page.number), Some(offset))?;
            let expected = decoded_image_len(&stream.dict, width, height)
                .map_err(|message| import_error(Some(page.number), Some(offset), message))?;
            if expected > remaining {
                return Err(import_error(
                    Some(page.number),
                    Some(offset),
                    format!(
                        "aggregate decompressed bytes would exceed limit {} while decoding image",
                        self.limits.max_decompressed_bytes
                    ),
                ));
            }
            match decode_png(&stream, width, height, remaining) {
                Ok((data, decoded)) => {
                    self.charge_decompressed(decoded, Some(page.number), Some(offset), "image")?;
                    (data, "image.png".to_owned(), "image/png".to_owned())
                }
                Err(message) => {
                    self.diagnostic(page.number, offset, message)?;
                    return Ok(());
                }
            }
        } else {
            self.diagnostic(
                page.number,
                offset,
                format!("image filters {} are unsupported", filters.join(",")),
            )?;
            return Ok(());
        };
        self.ensure_shape_room(page, 1)?;
        let image = PositionedElement::Image {
            rect: Rect {
                x: 0.0,
                y: 0.0,
                width: 1.0,
                height: 1.0,
            },
            media_id: MediaId::from_bytes(&data),
            data: data.clone(),
            content_type,
        };
        page.layout.push(PositionedElement::Group(GroupElement {
            transform,
            clip: None,
            opacity: 1.0,
            effects: Vec::new(),
            children: vec![image],
        }));
        page.editable.push(EditableElement::Image {
            rect: Rect {
                x: 0.0,
                y: 0.0,
                width: 1.0,
                height: 1.0,
            },
            transform,
            data,
            filename,
        });
        Ok(())
    }

    fn parse_annotations(&mut self, page: &mut NormalizedPage, id: ObjectId) -> Result<()> {
        let mut resolver = ResourceReferenceResolver::new(self.limits.max_objects);
        let ids = strict_annotation_ids(
            &self.document,
            id,
            page.number,
            page.retained_elements,
            self.limits,
            &mut resolver,
        )?;
        self.ensure_shape_room(page, ids.len())?;
        let mut seen = HashSet::new();
        for (index, id) in ids.into_iter().enumerate() {
            match strict_annotation(&self.document, page, index, id, &mut resolver, &mut seen)? {
                AnnotationDecision::Skip => {}
                AnnotationDecision::Diagnostic(message) => {
                    self.annotation_diagnostic(page.number, index, message)?;
                }
                AnnotationDecision::Link { rect, uri } => {
                    page.layout.push(PositionedElement::LinkAnnotation {
                        rect,
                        url: uri.clone(),
                    });
                    page.links.push((rect, uri));
                }
            }
        }
        Ok(())
    }

    fn ensure_shape_room(&self, page: &mut NormalizedPage, additional: usize) -> Result<()> {
        let retained = page.retained_elements.saturating_add(additional);
        if retained > self.limits.max_shapes_per_page {
            Err(import_error(
                Some(page.number),
                None,
                format!(
                    "page shape count exceeds limit {}",
                    self.limits.max_shapes_per_page
                ),
            ))
        } else {
            page.retained_elements = retained;
            Ok(())
        }
    }

    fn annotation_diagnostic(
        &mut self,
        page: u32,
        index: usize,
        message: impl Into<String>,
    ) -> Result<()> {
        if self.diagnostics.len() >= self.limits.max_diagnostics {
            return Err(import_error(
                Some(page),
                None,
                format!(
                    "diagnostic count exceeds limit {}",
                    self.limits.max_diagnostics
                ),
            ));
        }
        self.diagnostics.push(PdfImportDiagnostic {
            path: format!("pdf/page[{page}]/annotation[{index}]"),
            message: message.into(),
        });
        Ok(())
    }

    fn project(&mut self, pages: &[NormalizedPage]) -> Result<Presentation> {
        let mut presentation = Presentation::from_bytes(DEFAULT_TEMPLATE)?;
        while !presentation.slides.is_empty() {
            presentation.remove_slide(0)?;
        }
        presentation.set_slide_size(
            points_to_emu(pages[0].geometry.width)?,
            points_to_emu(pages[0].geometry.height)?,
        )?;
        let font_data = self.fonts.all_font_data();
        for (index, page) in pages.iter().enumerate() {
            presentation.add_slide(0)?;
            presentation.slides[index]
                .slide
                .common_slide_data
                .shape_tree
                .children
                .clear();
            match self.mode {
                PdfImportMode::PreservedGraphic { dpi } => {
                    let pixels = raster_pixels(page.geometry.width, page.geometry.height, dpi)?;
                    if pixels > self.limits.max_pixels_per_page {
                        return Err(import_error(
                            Some(page.number),
                            None,
                            format!(
                                "raster pixel count {pixels} exceeds limit {}",
                                self.limits.max_pixels_per_page
                            ),
                        ));
                    }
                    let mut frame = PageFrame::new(
                        index + 1,
                        page.geometry.width,
                        page.geometry.height,
                        page.layout.clone(),
                    );
                    frame.background = Some(Paint::Solid(Color::WHITE));
                    let layout = LayoutResult::new(
                        vec![Arc::new(frame)],
                        font_data.clone(),
                        None,
                        Vec::new(),
                    );
                    let png = oxml_pdf::render_page_to_png(&layout, 0, dpi).ok_or_else(|| {
                        import_error(Some(page.number), None, "shared PDF rasterizer failed")
                    })?;
                    presentation.add_picture(
                        index,
                        &png,
                        "page.png",
                        Emu(0),
                        Emu(0),
                        Some(points_to_emu(page.geometry.width)?),
                        Some(points_to_emu(page.geometry.height)?),
                    )?;
                }
                PdfImportMode::Editable => self.project_editable(&mut presentation, index, page)?,
            }
        }
        let bytes = presentation.to_bytes()?;
        let reopened = Presentation::from_bytes(&bytes)?;
        let issues = reopened.validate();
        if !issues.is_empty() {
            return Err(import_error(
                None,
                None,
                format!(
                    "generated presentation failed validation with {} issue(s)",
                    issues.len()
                ),
            ));
        }
        Ok(reopened)
    }

    fn project_editable(
        &self,
        presentation: &mut Presentation,
        slide_index: usize,
        page: &NormalizedPage,
    ) -> Result<()> {
        for element in &page.editable {
            match element {
                EditableElement::Text {
                    rect,
                    transform,
                    text,
                    family,
                    size,
                    color,
                } => {
                    let affine = editable_affine(*rect, *transform)?;
                    let slide_count = presentation.slides.len();
                    let mut slide =
                        presentation
                            .slide_mut(slide_index)
                            .ok_or(Error::UnknownSlideIndex {
                                index: slide_index,
                                slide_count,
                            })?;
                    let mut shape = slide.add_textbox(
                        rect_emu(*rect)?.0,
                        rect_emu(*rect)?.1,
                        rect_emu(*rect)?.2,
                        rect_emu(*rect)?.3,
                    )?;
                    shape.set_text(text)?;
                    shape.set_rotation(affine.child_rotation)?;
                    if let Some(frame) = shape.text_frame() {
                        frame.body.body_properties.left_inset = Some(Coordinate32Value::Emu(0));
                        frame.body.body_properties.top_inset = Some(Coordinate32Value::Emu(0));
                        frame.body.body_properties.right_inset = Some(Coordinate32Value::Emu(0));
                        frame.body.body_properties.bottom_inset = Some(Coordinate32Value::Emu(0));
                    }
                    if let Some(frame) = shape.text_frame()
                        && let Some(mut paragraph) = frame.into_paragraph_mut(0)
                        && let Some(mut run) = paragraph.run_mut(0)
                    {
                        let mut properties = CT_TextCharacterProperties::default();
                        properties.font_size = Some((size * 100.0).trunc() as i32);
                        properties.latin =
                            Some(TextFont::new(family.clone()).map_err(|error| {
                                import_error(
                                    None,
                                    None,
                                    format!("text font construction failed: {error}"),
                                )
                            })?);
                        properties.fill = Some(solid_fill(*color)?);
                        run.set_properties(properties);
                    }
                    let _ = shape;
                    wrap_last_shape_in_affine_group(
                        presentation,
                        slide_index,
                        affine.group_transform,
                    )?;
                }
                EditableElement::Path(path) => append_custom_path(presentation, slide_index, path)?,
                EditableElement::Image {
                    rect,
                    transform,
                    data,
                    filename,
                } => {
                    let affine = editable_affine(*rect, *transform)?;
                    let (left, top, width, height) = rect_emu(*rect)?;
                    presentation.add_picture(
                        slide_index,
                        data,
                        filename,
                        left,
                        top,
                        Some(width),
                        Some(height),
                    )?;
                    let shape_count = presentation
                        .slide(slide_index)
                        .map(|slide| slide.shapes().len())
                        .ok_or(Error::UnknownSlideIndex {
                            index: slide_index,
                            slide_count: presentation.slides.len(),
                        })?;
                    let mut shape = presentation
                        .slide_mut(slide_index)
                        .and_then(|slide| slide.into_shape_mut(shape_count - 1))
                        .ok_or_else(|| {
                            import_error(None, None, "inserted PDF image was not retained")
                        })?;
                    shape.set_rotation(affine.child_rotation)?;
                    let _ = shape;
                    wrap_last_shape_in_affine_group(
                        presentation,
                        slide_index,
                        affine.group_transform,
                    )?;
                }
            }
        }
        for (rect, uri) in &page.links {
            append_link_overlay(presentation, slide_index, *rect, uri)?;
        }
        Ok(())
    }
}

struct EditableAffine {
    group_transform: CT_Transform2D,
    child_rotation: Angle,
}

fn editable_affine(rect: Rect, transform: Transform) -> Result<EditableAffine> {
    let p = transform.a * transform.a + transform.b * transform.b;
    let q = transform.a * transform.c + transform.b * transform.d;
    let r = transform.c * transform.c + transform.d * transform.d;
    let discriminant = (p - r).hypot(2.0 * q);
    let first = ((p + r + discriminant) / 2.0).sqrt();
    let second = ((p + r - discriminant) / 2.0).max(0.0).sqrt();
    if !first.is_finite() || !second.is_finite() || first <= f64::EPSILON || second <= f64::EPSILON
    {
        return Err(import_error(
            None,
            None,
            "editable affine transform is singular",
        ));
    }

    let theta = 0.5 * (2.0 * q).atan2(p - r);
    let (sin, cos) = theta.sin_cos();
    let u1_x = (transform.a * cos + transform.c * sin) / first;
    let u1_y = (transform.b * cos + transform.d * sin) / first;
    let u2_x = (-transform.a * sin + transform.c * cos) / second;
    let u2_y = (-transform.b * sin + transform.d * cos) / second;
    let determinant = u1_x * u2_y - u1_y * u2_x;
    let (rotation, flip_vertical) = if determinant < 0.0 {
        ((-u1_y).atan2(u1_x).to_degrees(), true)
    } else {
        (u1_y.atan2(u1_x).to_degrees(), false)
    };

    let local_center = Point {
        x: rect.x + rect.width / 2.0,
        y: rect.y + rect.height / 2.0,
    };
    let target_center = transform.apply(local_center);
    let extent_width = first * rect.width;
    let extent_height = second * rect.height;
    let mut group_transform = CT_Transform2D::default();
    group_transform.offset = Some(CT_Point2D {
        x: points_to_emu(target_center.x - extent_width / 2.0)?,
        y: points_to_emu(target_center.y - extent_height / 2.0)?,
    });
    group_transform.extent = Some(CT_PositiveSize2D {
        cx: points_to_emu(extent_width)?,
        cy: points_to_emu(extent_height)?,
    });
    group_transform.child_offset = Some(CT_Point2D {
        x: points_to_emu(rect.x)?,
        y: points_to_emu(rect.y)?,
    });
    group_transform.child_extent = Some(CT_PositiveSize2D {
        cx: points_to_emu(rect.width)?,
        cy: points_to_emu(rect.height)?,
    });
    group_transform.rotation = Angle::from_degrees(rotation);
    group_transform.flip_vertical = flip_vertical;
    Ok(EditableAffine {
        group_transform,
        child_rotation: Angle::from_degrees(-theta.to_degrees()),
    })
}

fn wrap_last_shape_in_affine_group(
    presentation: &mut Presentation,
    slide_index: usize,
    transform: CT_Transform2D,
) -> Result<()> {
    let slide_count = presentation.slides.len();
    let slide = presentation
        .slides
        .get_mut(slide_index)
        .ok_or(Error::UnknownSlideIndex {
            index: slide_index,
            slide_count,
        })?;
    let tree = &mut slide.slide.common_slide_data.shape_tree;
    let id = ShapeIdAllocator::scan(tree).allocate();
    let child = tree
        .children
        .pop()
        .ok_or_else(|| import_error(None, None, "affine PDF child shape is missing"))?;
    let mut group = CT_GroupShape::new_empty(id, &format!("PDF affine group {id}"));
    *group.group_transform_mut() = transform;
    group.children.push(child);
    tree.append_child(ShapeTreeChild::GroupShape(Box::new(group)));
    Ok(())
}

fn validate_limits(bytes: &[u8], mode: PdfImportMode, limits: PdfImportLimits) -> Result<()> {
    if u64::try_from(bytes.len()).unwrap_or(u64::MAX) > limits.max_input_bytes {
        return Err(import_error(
            None,
            None,
            format!(
                "PDF input length {} exceeds limit {}",
                bytes.len(),
                limits.max_input_bytes
            ),
        ));
    }
    if limits.max_pages == 0
        || limits.max_objects == 0
        || limits.max_decompressed_bytes == 0
        || limits.max_operations_per_page == 0
        || limits.max_pixels_per_page == 0
        || limits.max_shapes_per_page == 0
        || limits.max_diagnostics == 0
    {
        return Err(import_error(
            None,
            None,
            "PDF import limits must all be positive",
        ));
    }
    if let PdfImportMode::PreservedGraphic { dpi } = mode {
        finite_positive(dpi, 0, 0, "DPI")?;
    }
    Ok(())
}

fn import_error(page: Option<u32>, offset: Option<usize>, message: impl Into<String>) -> Error {
    Error::PdfImport {
        page,
        offset,
        message: message.into(),
    }
}

fn reject_active_content(document: &Document, object_limit: usize) -> Result<usize> {
    let work_limit = object_limit.saturating_mul(256);
    let mut scanner = ActiveContentScanner {
        document,
        seen: HashSet::new(),
        active_actions: HashSet::new(),
        work: 0,
        work_limit,
    };
    for id in document.objects.keys() {
        scanner.scan(
            &Object::Reference(*id),
            &format!("object {} {}", id.0, id.1),
            false,
            0,
        )?;
    }
    Ok(scanner.work)
}

struct ActiveContentScanner<'a> {
    document: &'a Document,
    seen: HashSet<(ObjectId, bool)>,
    active_actions: HashSet<ObjectId>,
    work: usize,
    work_limit: usize,
}

impl ActiveContentScanner<'_> {
    fn scan(
        &mut self,
        object: &Object,
        path: &str,
        action_context: bool,
        depth: usize,
    ) -> Result<()> {
        self.work = self.work.saturating_add(1);
        if self.work > self.work_limit {
            return Err(import_error(
                None,
                None,
                format!(
                    "active-content graph work exceeds limit {} at {path}",
                    self.work_limit
                ),
            ));
        }
        if depth > MAX_PDF_GRAPH_DEPTH {
            return Err(import_error(
                None,
                None,
                format!("active-content graph exceeds depth limit at {path}"),
            ));
        }
        match object {
            Object::Reference(id) => {
                if action_context {
                    if self.seen.contains(&(*id, true)) {
                        return Ok(());
                    }
                    if !self.active_actions.insert(*id) {
                        return Err(import_error(
                            None,
                            None,
                            format!("cyclic action reference at {path}"),
                        ));
                    }
                    let target = self.document.objects.get(id).ok_or_else(|| {
                        import_error(
                            None,
                            None,
                            format!(
                                "active-content reference failed at {path}: missing object {} {}",
                                id.0, id.1
                            ),
                        )
                    })?;
                    let result = self.scan(
                        target,
                        &format!("{path}/{} {}", id.0, id.1),
                        true,
                        depth + 1,
                    );
                    self.active_actions.remove(id);
                    result?;
                    self.seen.insert((*id, true));
                } else if self.seen.insert((*id, false)) {
                    let target = self.document.get_object(*id).map_err(|error| {
                        import_error(
                            None,
                            None,
                            format!("active-content reference failed at {path}: {error}"),
                        )
                    })?;
                    self.scan(
                        target,
                        &format!("{path}/{} {}", id.0, id.1),
                        false,
                        depth + 1,
                    )?;
                }
            }
            Object::Array(values) => {
                if action_context {
                    return Err(import_error(
                        None,
                        None,
                        format!("action at {path} must resolve to a dictionary"),
                    ));
                }
                for (index, value) in values.iter().enumerate() {
                    self.scan(value, &format!("{path}[{index}]"), false, depth + 1)?;
                }
            }
            Object::Dictionary(dictionary) => {
                self.scan_dictionary(dictionary, path, action_context, depth)?;
            }
            Object::Stream(stream) => {
                if action_context {
                    return Err(import_error(
                        None,
                        None,
                        format!("action at {path} must resolve to a dictionary"),
                    ));
                }
                self.scan_dictionary(&stream.dict, path, false, depth)?;
            }
            _ if action_context => {
                return Err(import_error(
                    None,
                    None,
                    format!("action at {path} must resolve to a dictionary"),
                ));
            }
            _ => {}
        }
        Ok(())
    }

    fn scan_dictionary(
        &mut self,
        dictionary: &Dictionary,
        path: &str,
        action_context: bool,
        depth: usize,
    ) -> Result<()> {
        if dictionary.has(b"AA") {
            return Err(import_error(
                None,
                None,
                format!("additional actions are unsupported at {path}/AA"),
            ));
        }
        let action_kind = dictionary.get(b"S").and_then(Object::as_name).ok();
        if dictionary.has(b"JS") || (action_context && action_kind == Some(b"JavaScript")) {
            return Err(import_error(
                None,
                None,
                format!("JavaScript active content is unsupported at {path}"),
            ));
        }
        if action_context {
            strict_uri_action_dictionary(dictionary, None, path)?;
        }
        let object_type = dictionary.get(b"Type").and_then(Object::as_name).ok();
        for (key, value) in dictionary.iter() {
            let nested_action = match key.as_slice() {
                b"OpenAction" => object_type == Some(b"Catalog"),
                b"A" => object_type == Some(b"Annot"),
                b"Next" => action_context,
                _ => false,
            };
            self.scan(
                value,
                &format!("{path}/{}", String::from_utf8_lossy(key)),
                nested_action,
                depth + 1,
            )?;
        }
        Ok(())
    }
}

fn strict_pages(
    document: &Document,
    max_pages: usize,
    max_objects: usize,
) -> Result<Vec<(u32, ObjectId)>> {
    enum Visit {
        Enter {
            id: ObjectId,
            parent: Option<ObjectId>,
            depth: usize,
        },
        Exit {
            id: ObjectId,
            declared_count: usize,
            first_page: usize,
        },
    }

    let root = document
        .trailer
        .get(b"Root")
        .and_then(Object::as_reference)
        .map_err(|_| import_error(None, None, "PDF trailer Root must be an indirect reference"))?;
    let catalog = document
        .get_dictionary(root)
        .map_err(|error| import_error(None, None, format!("PDF catalog object failed: {error}")))?;
    if catalog.get(b"Type").and_then(Object::as_name).ok() != Some(b"Catalog") {
        return Err(import_error(
            None,
            None,
            "PDF catalog must declare /Type /Catalog",
        ));
    }
    let pages_root = catalog
        .get(b"Pages")
        .and_then(Object::as_reference)
        .map_err(|_| {
            import_error(
                None,
                None,
                "PDF catalog Pages must be an indirect reference",
            )
        })?;

    let mut pages = Vec::new();
    let mut active = HashSet::new();
    let mut visited = HashSet::new();
    let mut work = 1usize;
    let mut stack = vec![Visit::Enter {
        id: pages_root,
        parent: None,
        depth: 0,
    }];
    while let Some(visit) = stack.pop() {
        match visit {
            Visit::Exit {
                id,
                declared_count,
                first_page,
            } => {
                let actual_count = pages.len().checked_sub(first_page).ok_or_else(|| {
                    import_error(None, None, "PDF page-tree count arithmetic underflow")
                })?;
                if actual_count != declared_count {
                    return Err(import_error(
                        None,
                        None,
                        format!(
                            "page-tree node {} {} declares Count {declared_count} but contains {actual_count} pages",
                            id.0, id.1
                        ),
                    ));
                }
                active.remove(&id);
                visited.insert(id);
            }
            Visit::Enter { id, parent, depth } => {
                if active.contains(&id) {
                    return Err(import_error(
                        None,
                        None,
                        format!("cyclic page tree at object {} {}", id.0, id.1),
                    ));
                }
                if visited.contains(&id) {
                    return Err(import_error(
                        None,
                        None,
                        format!("duplicate page-tree node at object {} {}", id.0, id.1),
                    ));
                }
                if depth > MAX_PDF_GRAPH_DEPTH {
                    return Err(import_error(
                        None,
                        None,
                        format!(
                            "page tree exceeds depth limit {} at object {} {}",
                            MAX_PDF_GRAPH_DEPTH, id.0, id.1
                        ),
                    ));
                }
                let dictionary = document.get_dictionary(id).map_err(|error| {
                    import_error(
                        None,
                        None,
                        format!("page-tree object {} {} failed: {error}", id.0, id.1),
                    )
                })?;
                match parent {
                    Some(expected) => {
                        let actual = dictionary
                            .get(b"Parent")
                            .and_then(Object::as_reference)
                            .map_err(|_| {
                                import_error(
                                    None,
                                    None,
                                    format!(
                                        "page-tree object {} {} Parent must be an indirect reference",
                                        id.0, id.1
                                    ),
                                )
                            })?;
                        if actual != expected {
                            return Err(import_error(
                                None,
                                None,
                                format!(
                                    "page-tree object {} {} Parent does not match its containing Pages node",
                                    id.0, id.1
                                ),
                            ));
                        }
                    }
                    None if dictionary.has(b"Parent") => {
                        return Err(import_error(
                            None,
                            None,
                            "root Pages node must not declare Parent",
                        ));
                    }
                    None => {}
                }
                let kind = dictionary
                    .get(b"Type")
                    .and_then(Object::as_name)
                    .map_err(|_| {
                        import_error(
                            None,
                            None,
                            format!(
                                "page-tree object {} {} must declare a name Type",
                                id.0, id.1
                            ),
                        )
                    })?;
                match kind {
                    b"Page" => {
                        if pages.len() >= max_pages {
                            return Err(import_error(
                                None,
                                None,
                                format!("PDF page count exceeds limit {max_pages}"),
                            ));
                        }
                        let number = u32::try_from(pages.len() + 1).map_err(|_| {
                            import_error(None, None, "PDF page number exceeds u32 range")
                        })?;
                        pages.push((number, id));
                        visited.insert(id);
                    }
                    b"Pages" => {
                        let declared_count = dictionary
                            .get(b"Count")
                            .and_then(Object::as_i64)
                            .ok()
                            .and_then(|value| usize::try_from(value).ok())
                            .ok_or_else(|| {
                                import_error(
                                    None,
                                    None,
                                    format!(
                                        "Pages node {} {} Count must be a nonnegative integer",
                                        id.0, id.1
                                    ),
                                )
                            })?;
                        let kids =
                            dictionary
                                .get(b"Kids")
                                .and_then(Object::as_array)
                                .map_err(|_| {
                                    import_error(
                                        None,
                                        None,
                                        format!(
                                            "Pages node {} {} Kids must be an array",
                                            id.0, id.1
                                        ),
                                    )
                                })?;
                        work = work.checked_add(kids.len()).ok_or_else(|| {
                            import_error(None, None, "page-tree work count overflow")
                        })?;
                        if work > max_objects {
                            return Err(import_error(
                                None,
                                None,
                                format!("page-tree work exceeds object limit {max_objects}"),
                            ));
                        }
                        active.insert(id);
                        stack.push(Visit::Exit {
                            id,
                            declared_count,
                            first_page: pages.len(),
                        });
                        for kid in kids.iter().rev() {
                            let kid = kid.as_reference().map_err(|_| {
                                import_error(
                                    None,
                                    None,
                                    format!(
                                        "Pages node {} {} Kids members must be indirect references",
                                        id.0, id.1
                                    ),
                                )
                            })?;
                            stack.push(Visit::Enter {
                                id: kid,
                                parent: Some(id),
                                depth: depth + 1,
                            });
                        }
                    }
                    other => {
                        return Err(import_error(
                            None,
                            None,
                            format!(
                                "page-tree object {} {} has unsupported Type /{}",
                                id.0,
                                id.1,
                                String::from_utf8_lossy(other)
                            ),
                        ));
                    }
                }
            }
        }
    }
    if pages.is_empty() {
        return Err(import_error(None, None, "PDF contains no pages"));
    }
    Ok(pages)
}

fn strict_page_content(
    document: &Document,
    page_id: ObjectId,
    page: u32,
    max_objects: usize,
    max_decompressed: usize,
) -> Result<Vec<u8>> {
    let dictionary = document.get_dictionary(page_id).map_err(|error| {
        import_error(
            Some(page),
            None,
            format!("page dictionary failed while resolving Contents: {error}"),
        )
    })?;
    let Ok(contents) = dictionary.get(b"Contents") else {
        return Ok(Vec::new());
    };
    let mut resolver = ResourceReferenceResolver::new(max_objects);
    let mut ids = Vec::new();
    match contents {
        Object::Reference(id) => ids.push(*id),
        Object::Array(values) => {
            resolver.charge_many(page, values.len())?;
            if values.len() > max_objects {
                return Err(import_error(
                    Some(page),
                    None,
                    format!("page Contents count exceeds object limit {max_objects}"),
                ));
            }
            ids.reserve(values.len());
            for (index, value) in values.iter().enumerate() {
                let id = value.as_reference().map_err(|_| {
                    import_error(
                        Some(page),
                        None,
                        format!("page Contents array member {index} must be an indirect reference"),
                    )
                })?;
                ids.push(id);
            }
        }
        _ => {
            return Err(import_error(
                Some(page),
                None,
                "page Contents must be an indirect stream reference or an array of indirect stream references",
            ));
        }
    }

    let mut content = Vec::new();
    for (index, id) in ids.into_iter().enumerate() {
        let label = format!("page Contents stream {index}");
        let (_, object) = resolver.resolve(document, id, page, &label)?;
        let stream = object.as_stream().map_err(|_| {
            import_error(Some(page), None, format!("{label} must reference a stream"))
        })?;
        let remaining = max_decompressed.saturating_sub(content.len());
        let decoded = strict_content_stream_data(stream, page, index, remaining)?;
        let retained = content
            .len()
            .checked_add(decoded.len())
            .and_then(|length| length.checked_add(1))
            .ok_or_else(|| import_error(Some(page), None, "page content byte count overflow"))?;
        if retained > max_decompressed {
            return Err(import_error(
                Some(page),
                None,
                format!("page content bytes exceed limit {max_decompressed}"),
            ));
        }
        content.extend_from_slice(&decoded);
        content.push(b'\n');
    }
    Ok(content)
}

fn strict_content_stream_data(
    stream: &Stream,
    page: u32,
    index: usize,
    limit: usize,
) -> Result<Vec<u8>> {
    strict_flate_stream_data(
        stream,
        Some(page),
        &format!("page Contents stream {index}"),
        limit,
    )
}

fn strict_flate_stream_data(
    stream: &Stream,
    page: Option<u32>,
    label: &str,
    limit: usize,
) -> Result<Vec<u8>> {
    if stream.dict.has(b"DecodeParms") {
        return Err(import_error(
            page,
            None,
            format!("{label} DecodeParms are unsupported"),
        ));
    }
    let filters = match stream.dict.get(b"Filter") {
        Err(_) => Vec::new(),
        Ok(Object::Name(name)) => vec![name.as_slice()],
        Ok(Object::Array(values)) => values
            .iter()
            .map(|value| {
                value.as_name().map_err(|_| {
                    import_error(page, None, format!("{label} Filter members must be names"))
                })
            })
            .collect::<Result<Vec<_>>>()?,
        Ok(_) => {
            return Err(import_error(
                page,
                None,
                format!("{label} Filter must be a name or array of names"),
            ));
        }
    };
    if let Some(filter) = filters.iter().find(|filter| **filter != b"FlateDecode") {
        return Err(import_error(
            page,
            None,
            format!(
                "{label} uses unsupported Filter /{}",
                String::from_utf8_lossy(filter)
            ),
        ));
    }
    if filters.is_empty() {
        if stream.content.len() > limit {
            return Err(import_error(
                page,
                None,
                format!("{label} bytes exceed limit {limit}"),
            ));
        }
        return Ok(stream.content.clone());
    }
    let mut decoded =
        miniz_oxide::inflate::decompress_to_vec_zlib_with_limit(&stream.content, limit).map_err(
            |error| import_error(page, None, format!("{label} FlateDecode failed: {error:?}")),
        )?;
    for _ in filters.iter().skip(1) {
        decoded = miniz_oxide::inflate::decompress_to_vec_zlib_with_limit(&decoded, limit)
            .map_err(|error| {
                import_error(page, None, format!("{label} FlateDecode failed: {error:?}"))
            })?;
    }
    Ok(decoded)
}

enum AnnotationDecision {
    Skip,
    Diagnostic(String),
    Link { rect: Rect, uri: String },
}

fn strict_annotation_ids(
    document: &Document,
    page_id: ObjectId,
    page: u32,
    retained_elements: usize,
    limits: PdfImportLimits,
    resolver: &mut ResourceReferenceResolver,
) -> Result<Vec<ObjectId>> {
    let dictionary = document.get_dictionary(page_id).map_err(|error| {
        import_error(
            Some(page),
            None,
            format!("page dictionary failed while resolving Annots: {error}"),
        )
    })?;
    let Ok(mut annotations) = dictionary.get(b"Annots") else {
        return Ok(Vec::new());
    };
    if let Object::Reference(id) = annotations {
        annotations = resolver.resolve(document, *id, page, "page Annots")?.1;
    }
    let annotations = annotations
        .as_array()
        .map_err(|_| import_error(Some(page), None, "page Annots must resolve to an array"))?;
    resolver.charge_many(page, annotations.len())?;
    let retained = retained_elements
        .checked_add(annotations.len())
        .ok_or_else(|| import_error(Some(page), None, "annotation shape count overflow"))?;
    if retained > limits.max_shapes_per_page {
        return Err(import_error(
            Some(page),
            None,
            format!(
                "page annotation count {} exceeds remaining shape limit {}",
                annotations.len(),
                limits.max_shapes_per_page.saturating_sub(retained_elements)
            ),
        ));
    }
    let mut ids = Vec::with_capacity(annotations.len());
    for (index, annotation) in annotations.iter().enumerate() {
        ids.push(annotation.as_reference().map_err(|_| {
            import_error(
                Some(page),
                None,
                format!("page Annots member {index} must be an indirect reference"),
            )
        })?);
    }
    Ok(ids)
}

fn strict_annotation(
    document: &Document,
    page: &NormalizedPage,
    index: usize,
    id: ObjectId,
    resolver: &mut ResourceReferenceResolver,
    seen: &mut HashSet<ObjectId>,
) -> Result<AnnotationDecision> {
    let label = format!("annotation {index}");
    let (terminal, object) = resolver.resolve(document, id, page.number, &label)?;
    if !seen.insert(terminal) {
        return Err(import_error(
            Some(page.number),
            None,
            format!(
                "page Annots repeats annotation target {} {}",
                terminal.0, terminal.1
            ),
        ));
    }
    let annotation = object.as_dict().map_err(|_| {
        import_error(
            Some(page.number),
            None,
            format!("{label} must reference a dictionary"),
        )
    })?;
    if annotation.get(b"Type").and_then(Object::as_name).ok() != Some(b"Annot") {
        return Err(import_error(
            Some(page.number),
            None,
            format!("{label} dictionary must declare /Type /Annot"),
        ));
    }
    let subtype = annotation
        .get(b"Subtype")
        .and_then(Object::as_name)
        .map_err(|_| {
            import_error(
                Some(page.number),
                None,
                format!("{label} dictionary must declare a name Subtype"),
            )
        })?;
    if subtype != b"Link" {
        return Ok(AnnotationDecision::Skip);
    }
    let values = annotation
        .get(b"Rect")
        .and_then(Object::as_array)
        .map_err(|error| {
            import_error(
                Some(page.number),
                None,
                format!("link annotation {index} rectangle failed: {error}"),
            )
        })?;
    let rect = page.geometry.rect(
        array_numbers(values, 4, page.number, index)?
            .try_into()
            .map_err(|_| {
                import_error(
                    Some(page.number),
                    None,
                    "link rectangle requires four values",
                )
            })?,
    );
    checked_rect(rect, page.number, index, "link rectangle")?;
    let Ok(mut action) = annotation.get(b"A") else {
        return Ok(AnnotationDecision::Diagnostic(
            "link annotation without a URI action was dropped".to_owned(),
        ));
    };
    if let Object::Reference(id) = action {
        action = resolver
            .resolve(
                document,
                *id,
                page.number,
                &format!("link annotation {index} action"),
            )?
            .1;
    }
    let action = action.as_dict().map_err(|_| {
        import_error(
            Some(page.number),
            None,
            format!("link annotation {index} action must resolve to a dictionary"),
        )
    })?;
    let uri = strict_uri_action_dictionary(
        action,
        Some(page.number),
        &format!("link annotation {index} action"),
    )?;
    Ok(AnnotationDecision::Link { rect, uri })
}

fn page_geometry(document: &Document, page_id: ObjectId, page: u32) -> Result<PageGeometry> {
    let media = inherited(document, page_id, b"MediaBox")?
        .ok_or_else(|| import_error(Some(page), None, "page has no inherited MediaBox"))?;
    let crop = inherited(document, page_id, b"CropBox")?.unwrap_or(media);
    let values = crop
        .as_array()
        .map_err(|error| import_error(Some(page), None, format!("page box failed: {error}")))?;
    let values: [f64; 4] = array_numbers(values, 4, page, 0)?
        .try_into()
        .map_err(|_| import_error(Some(page), None, "page box requires four numbers"))?;
    if values[2] <= values[0] || values[3] <= values[1] {
        return Err(import_error(
            Some(page),
            None,
            "page box has non-positive dimensions",
        ));
    }
    let rotation = inherited(document, page_id, b"Rotate")?
        .map(object_number)
        .transpose()
        .map_err(|error| import_error(Some(page), None, error))?
        .unwrap_or(0.0);
    if rotation.fract() != 0.0 {
        return Err(import_error(
            Some(page),
            None,
            "page rotation must be an integer",
        ));
    }
    let rotation = (rotation as i32).rem_euclid(360);
    if !matches!(rotation, 0 | 90 | 180 | 270) {
        return Err(import_error(
            Some(page),
            None,
            format!("unsupported page rotation {rotation}"),
        ));
    }
    let raw_width = values[2] - values[0];
    let raw_height = values[3] - values[1];
    if !raw_width.is_finite() || !raw_height.is_finite() {
        return Err(import_error(
            Some(page),
            None,
            "page dimensions produce non-finite geometry",
        ));
    }
    let (width, height) = if matches!(rotation, 90 | 270) {
        (raw_height, raw_width)
    } else {
        (raw_width, raw_height)
    };
    Ok(PageGeometry {
        left: values[0],
        bottom: values[1],
        right: values[2],
        top: values[3],
        rotation,
        width,
        height,
    })
}

fn inherited<'a>(
    document: &'a Document,
    mut id: ObjectId,
    key: &[u8],
) -> Result<Option<&'a Object>> {
    let mut seen = HashSet::new();
    loop {
        if !seen.insert(id) {
            return Err(import_error(
                None,
                None,
                format!("cyclic page tree at object {} {}", id.0, id.1),
            ));
        }
        let dictionary = document.get_dictionary(id).map_err(|error| {
            import_error(None, None, format!("page tree object failed: {error}"))
        })?;
        if let Ok(value) = dictionary.get(key) {
            let (_, value) = document.dereference(value).map_err(|error| {
                import_error(None, None, format!("inherited value failed: {error}"))
            })?;
            return Ok(Some(value));
        }
        match dictionary.get(b"Parent").and_then(Object::as_reference) {
            Ok(parent) => id = parent,
            Err(_) => return Ok(None),
        }
    }
}

fn page_fonts(
    document: &Document,
    caches: &mut FontEncodingCaches,
    fonts: &BTreeMap<Vec<u8>, ObjectId>,
    operations: &[Operation],
    widths: &HashMap<ObjectId, Option<SimpleFontWidths>>,
) -> Result<BTreeMap<Vec<u8>, PageFont>> {
    let requested = text_strings_by_font(operations);
    fonts
        .iter()
        .map(|(name, id)| {
            let dictionary = document.get_dictionary(*id).map_err(|error| {
                import_error(
                    None,
                    None,
                    format!(
                        "resolved font /{} failed: {error}",
                        String::from_utf8_lossy(name)
                    ),
                )
            })?;
            let shaping_family = dictionary
                .get(b"BaseFont")
                .and_then(Object::as_name)
                .map(|name| String::from_utf8_lossy(name).into_owned())
                .unwrap_or_else(|_| String::from_utf8_lossy(name).into_owned());
            let display_family = shaping_family
                .split_once('+')
                .map_or(shaping_family.as_str(), |(_, family)| family)
                .to_owned();
            let has_to_unicode = dictionary.has(b"ToUnicode");
            let (encoding, encoding_supported) = if has_to_unicode {
                ("ToUnicode".to_owned(), true)
            } else {
                match dictionary.get(b"Encoding") {
                    Ok(Object::Name(name)) => {
                        let label = String::from_utf8_lossy(name).into_owned();
                        let identity = matches!(name.as_slice(), b"Identity-H" | b"Identity-V");
                        let supported = matches!(
                            name.as_slice(),
                            b"StandardEncoding"
                                | b"WinAnsiEncoding"
                                | b"MacRomanEncoding"
                                | b"MacExpertEncoding"
                        );
                        let supported = supported && !identity;
                        (label, supported)
                    }
                    Ok(Object::Dictionary(_)) | Ok(Object::Reference(_)) => {
                        ("custom encoding dictionary".to_owned(), false)
                    }
                    Ok(other) => (format!("{} encoding object", other.enum_variant()), false),
                    Err(_) => ("implicit StandardEncoding".to_owned(), true),
                }
            };
            let mut decoded = BTreeMap::new();
            if encoding_supported {
                let decoder = if has_to_unicode {
                    let stream_id = dictionary
                        .get(b"ToUnicode")
                        .and_then(Object::as_reference)
                        .map_err(|_| {
                            import_error(
                                None,
                                None,
                                "prevalidated font ToUnicode must reference its terminal stream",
                            )
                        })?;
                    if let Some(decoder) = caches.to_unicode_by_stream.get(&stream_id) {
                        Arc::clone(decoder)
                    } else {
                        let limit = document
                            .get_object(stream_id)
                            .and_then(Object::as_stream)
                            .map_err(|error| {
                                import_error(
                                    None,
                                    None,
                                    format!("prevalidated font ToUnicode failed: {error}"),
                                )
                            })?
                            .content
                            .len()
                            .max(1);
                        let mut to_unicode_dictionary = dictionary.clone();
                        to_unicode_dictionary.remove(b"Encoding");
                        let parsed = to_unicode_dictionary
                            .get_font_encoding_with_limit(document, limit)
                            .map_err(|error| {
                                import_error(
                                    None,
                                    None,
                                    format!("font encoding `{encoding}` failed: {error}"),
                                )
                            })?;
                        let lopdf::Encoding::UnicodeMapEncoding(map) = parsed else {
                            return Err(import_error(
                                None,
                                None,
                                "font ToUnicode stream does not contain a valid Unicode CMap",
                            ));
                        };
                        let decoder = Arc::new(CachedFontEncoding::Unicode(
                            lopdf::Encoding::UnicodeMapEncoding(map),
                        ));
                        caches.to_unicode_parses =
                            caches.to_unicode_parses.checked_add(1).ok_or_else(|| {
                                import_error(None, None, "ToUnicode parse count overflow")
                            })?;
                        caches
                            .to_unicode_by_stream
                            .insert(stream_id, Arc::clone(&decoder));
                        decoder
                    }
                } else if let Some(decoder) = caches.ordinary_by_font.get(id) {
                    Arc::clone(decoder)
                } else {
                    let parsed = dictionary
                        .get_font_encoding_with_limit(document, 1)
                        .map_err(|error| {
                            import_error(
                                None,
                                None,
                                format!("font encoding `{encoding}` failed: {error}"),
                            )
                        })?;
                    let decoder = match parsed {
                        lopdf::Encoding::UnicodeMapEncoding(map) => {
                            CachedFontEncoding::Unicode(lopdf::Encoding::UnicodeMapEncoding(map))
                        }
                        other => {
                            let mut values = Vec::with_capacity(256);
                            for byte in 0_u8..=u8::MAX {
                                values.push(Document::decode_text(&other, &[byte]).map_err(
                                    |error| {
                                        import_error(
                                            None,
                                            None,
                                            format!("font encoding `{encoding}` failed: {error}"),
                                        )
                                    },
                                )?);
                            }
                            CachedFontEncoding::SingleByte(values)
                        }
                    };
                    let decoder = Arc::new(decoder);
                    caches.ordinary_by_font.insert(*id, Arc::clone(&decoder));
                    decoder
                };
                if encoding_supported && let Some(values) = requested.get(name) {
                    for value in values {
                        decoded.insert(value.clone(), decoder.decode(value));
                    }
                }
            }
            let composite_widths_supported =
                dictionary.get(b"Subtype").and_then(Object::as_name).ok() != Some(b"Type0");
            let widths = widths.get(id).cloned().ok_or_else(|| {
                import_error(
                    None,
                    None,
                    format!(
                        "font /{} was not strictly prevalidated",
                        String::from_utf8_lossy(name)
                    ),
                )
            })?;
            Ok((
                name.clone(),
                PageFont {
                    display_family,
                    shaping_family,
                    encoding,
                    encoding_supported,
                    decoded,
                    widths,
                    composite_widths_supported,
                },
            ))
        })
        .collect()
}

fn strict_font_widths(
    document: &Document,
    dictionary: &Dictionary,
    resolver: &mut ResourceReferenceResolver,
    missing_width: f64,
) -> Result<Option<SimpleFontWidths>> {
    let Ok(mut values) = dictionary.get(b"Widths") else {
        return Ok(None);
    };
    if let Object::Reference(id) = values {
        values = resolver.resolve(document, *id, 0, "font Widths")?.1;
    }
    let first_char = dictionary
        .get(b"FirstChar")
        .and_then(Object::as_i64)
        .map_err(|error| import_error(None, None, format!("font FirstChar failed: {error}")))?;
    let values = values
        .as_array()
        .map_err(|_| import_error(None, None, "font Widths must resolve to an array"))?;
    resolver.charge_many(0, values.len())?;
    if let Ok(last_char) = dictionary.get(b"LastChar") {
        let last_char = last_char
            .as_i64()
            .map_err(|_| import_error(None, None, "font LastChar must be an integer"))?;
        let expected = last_char
            .checked_sub(first_char)
            .and_then(|span| span.checked_add(1))
            .and_then(|count| usize::try_from(count).ok())
            .ok_or_else(|| import_error(None, None, "font character range is invalid"))?;
        if expected != values.len() {
            return Err(import_error(
                None,
                None,
                format!(
                    "font Widths length {} does not match FirstChar and LastChar range {expected}",
                    values.len()
                ),
            ));
        }
    }
    let mut widths = Vec::with_capacity(values.len());
    for value in values {
        let width = object_number(value)
            .map_err(|message| import_error(None, None, format!("font Widths {message}")))?;
        if !width.is_finite() || width < 0.0 {
            return Err(import_error(
                None,
                None,
                "font Widths entries must be finite and nonnegative",
            ));
        }
        widths.push(width);
    }
    Ok(Some(SimpleFontWidths {
        first_char,
        widths,
        missing_width,
    }))
}

fn text_strings_by_font(operations: &[Operation]) -> BTreeMap<Vec<u8>, HashSet<Vec<u8>>> {
    let mut result = BTreeMap::<Vec<u8>, HashSet<Vec<u8>>>::new();
    let mut font = None::<Vec<u8>>;
    let mut stack = Vec::new();
    for operation in operations {
        match operation.operator.as_str() {
            "q" => stack.push(font.clone()),
            "Q" => font = stack.pop().unwrap_or(None),
            "Tf" => {
                font = operation
                    .operands
                    .first()
                    .and_then(|value| value.as_name().ok())
                    .map(Vec::from);
            }
            "Tj" | "'" => {
                if let (Some(font), Some(value)) = (
                    font.as_ref(),
                    operation
                        .operands
                        .first()
                        .and_then(|value| value.as_str().ok()),
                ) {
                    result
                        .entry(font.clone())
                        .or_default()
                        .insert(value.to_vec());
                }
            }
            "\"" => {
                if let (Some(font), Some(value)) = (
                    font.as_ref(),
                    operation
                        .operands
                        .get(2)
                        .and_then(|value| value.as_str().ok()),
                ) {
                    result
                        .entry(font.clone())
                        .or_default()
                        .insert(value.to_vec());
                }
            }
            "TJ" => {
                if let (Some(font), Some(array)) = (
                    font.as_ref(),
                    operation
                        .operands
                        .first()
                        .and_then(|value| value.as_array().ok()),
                ) {
                    let values = result.entry(font.clone()).or_default();
                    for value in array.iter().filter_map(|value| value.as_str().ok()) {
                        values.insert(value.to_vec());
                    }
                }
            }
            _ => {}
        }
    }
    result
}

fn embedded_fonts<'a>(
    document: &mut Document,
    pages: impl Iterator<Item = &'a BTreeMap<Vec<u8>, ObjectId>>,
    limit: usize,
    max_objects: usize,
) -> Result<EmbeddedFonts> {
    let mut result = Vec::new();
    let mut retained = 0usize;
    let mut retained_streams = HashSet::new();
    let mut retained_font_streams = HashSet::new();
    let mut visited_fonts = HashSet::new();
    let mut decoded_maps = Vec::<(ObjectId, Vec<u8>)>::new();
    let mut normalized_maps = Vec::<(ObjectId, ObjectId)>::new();
    let mut normalized_encodings = Vec::<(ObjectId, Object)>::new();
    let mut font_widths = HashMap::new();
    let mut resolver = ResourceReferenceResolver::new(max_objects);
    for fonts in pages {
        for id in fonts.values().copied() {
            if !visited_fonts.insert(id) {
                continue;
            }
            let dictionary = document.get_dictionary(id).map_err(|error| {
                import_error(
                    None,
                    None,
                    format!("resolved font dictionary failed: {error}"),
                )
            })?;
            let family = match dictionary.get(b"BaseFont") {
                Ok(Object::Name(name)) => String::from_utf8_lossy(name).into_owned(),
                Ok(_) => {
                    return Err(import_error(
                        None,
                        None,
                        "font BaseFont must be a name when present",
                    ));
                }
                Err(_) => "Embedded PDF font".to_owned(),
            };
            if let Ok(value) = dictionary.get(b"Encoding") {
                match value {
                    Object::Name(_) => {}
                    Object::Dictionary(dictionary) => {
                        resolver.charge_many(0, dictionary.len())?;
                    }
                    Object::Reference(reference) => {
                        let resolved = resolver
                            .resolve(document, *reference, 0, "font Encoding")?
                            .1;
                        match resolved {
                            Object::Name(_) => {}
                            Object::Dictionary(dictionary) => {
                                resolver.charge_many(0, dictionary.len())?;
                            }
                            _ => {
                                return Err(import_error(
                                    None,
                                    None,
                                    "font Encoding reference must resolve to a name or dictionary",
                                ));
                            }
                        }
                        normalized_encodings.push((id, resolved.clone()));
                    }
                    _ => {
                        return Err(import_error(
                            None,
                            None,
                            "font Encoding must be a name or dictionary",
                        ));
                    }
                }
            }
            if let Ok(value) = dictionary.get(b"ToUnicode") {
                let reference = value.as_reference().map_err(|_| {
                    import_error(
                        None,
                        None,
                        "font ToUnicode must be an indirect stream reference",
                    )
                })?;
                let (stream_id, object) =
                    resolver.resolve(document, reference, 0, "font ToUnicode")?;
                let stream = object.as_stream().map_err(|_| {
                    import_error(None, None, "font ToUnicode must reference a stream")
                })?;
                normalized_maps.push((id, stream_id));
                if retained_streams.insert(stream_id) {
                    let remaining = limit.saturating_sub(retained);
                    let data =
                        strict_flate_stream_data(stream, None, "font ToUnicode stream", remaining)?;
                    retained = retained.checked_add(data.len()).ok_or_else(|| {
                        import_error(None, None, "font encoding byte count overflow")
                    })?;
                    decoded_maps.push((stream_id, data));
                }
            }
            let mut missing_width = 0.0;
            let mut font_file = None;
            if let Ok(mut descriptor) = dictionary.get(b"FontDescriptor") {
                if let Object::Reference(reference) = descriptor {
                    descriptor = resolver
                        .resolve(document, *reference, 0, "font descriptor")?
                        .1;
                }
                let descriptor = descriptor.as_dict().map_err(|_| {
                    import_error(None, None, "font descriptor must resolve to a dictionary")
                })?;
                if let Ok(value) = descriptor.get(b"Type")
                    && value.as_name().ok() != Some(b"FontDescriptor")
                {
                    return Err(import_error(
                        None,
                        None,
                        "font descriptor Type must be /FontDescriptor",
                    ));
                }
                if let Ok(value) = descriptor.get(b"MissingWidth") {
                    missing_width = object_number(value).map_err(|message| {
                        import_error(None, None, format!("font MissingWidth {message}"))
                    })?;
                    if missing_width < 0.0 {
                        return Err(import_error(
                            None,
                            None,
                            "font MissingWidth must be finite and nonnegative",
                        ));
                    }
                }
                if let Ok(value) = descriptor.get(b"FontFile2") {
                    let reference = value.as_reference().map_err(|_| {
                        import_error(
                            None,
                            None,
                            "font FontFile2 must be an indirect stream reference",
                        )
                    })?;
                    let (stream_id, object) =
                        resolver.resolve(document, reference, 0, "font FontFile2")?;
                    let stream = object.as_stream().map_err(|_| {
                        import_error(None, None, "font FontFile2 must reference a stream")
                    })?;
                    font_file = Some((stream_id, stream));
                }
            }
            font_widths.insert(
                id,
                strict_font_widths(document, dictionary, &mut resolver, missing_width)?,
            );
            if dictionary.get(b"Subtype").and_then(Object::as_name).ok() != Some(b"TrueType") {
                continue;
            }
            let Some((stream_id, stream)) = font_file else {
                continue;
            };
            if !retained_font_streams.insert(stream_id) {
                continue;
            }
            let remaining = limit.saturating_sub(retained);
            let data =
                strict_flate_stream_data(stream, None, "embedded TrueType font stream", remaining)?;
            retained = retained
                .checked_add(data.len())
                .ok_or_else(|| import_error(None, None, "embedded font byte count overflow"))?;
            if retained > limit {
                return Err(import_error(
                    None,
                    None,
                    format!("embedded font bytes exceed limit {limit}"),
                ));
            }
            result.push(FontFile { family, data });
        }
    }
    for (id, data) in decoded_maps {
        let stream = document
            .get_object_mut(id)
            .and_then(Object::as_stream_mut)
            .map_err(|error| {
                import_error(None, None, format!("ToUnicode cache failed: {error}"))
            })?;
        stream.content = data;
        stream.dict.remove(b"Filter");
        stream.dict.remove(b"DecodeParms");
        stream.dict.set("Length", stream.content.len() as i64);
    }
    for (font_id, stream_id) in normalized_maps {
        document
            .get_dictionary_mut(font_id)
            .map_err(|error| {
                import_error(None, None, format!("ToUnicode font cache failed: {error}"))
            })?
            .set("ToUnicode", stream_id);
    }
    for (font_id, encoding) in normalized_encodings {
        document
            .get_dictionary_mut(font_id)
            .map_err(|error| {
                import_error(None, None, format!("font Encoding cache failed: {error}"))
            })?
            .set("Encoding", encoding);
    }
    Ok(EmbeddedFonts {
        files: result,
        retained_bytes: retained,
        widths: font_widths,
    })
}

#[derive(Debug)]
struct ResolvedPageResources {
    fonts: BTreeMap<Vec<u8>, ObjectId>,
    xobjects: BTreeMap<Vec<u8>, ObjectId>,
}

fn resolve_page_resources(
    document: &Document,
    page_id: ObjectId,
    page: u32,
    max_objects: usize,
) -> Result<ResolvedPageResources> {
    let mut resolver = ResourceReferenceResolver::new(max_objects);
    let Some(resources) = inherited_page_resources(document, page_id, page, &mut resolver)? else {
        return Ok(ResolvedPageResources {
            fonts: BTreeMap::new(),
            xobjects: BTreeMap::new(),
        });
    };
    let fonts = font_entries(document, resources, page, &mut resolver)?;
    let xobjects =
        xobjects_from_resources(document, std::iter::once(resources), page, &mut resolver)?;
    Ok(ResolvedPageResources { fonts, xobjects })
}

#[cfg(test)]
fn page_xobjects(
    document: &Document,
    page_id: ObjectId,
    page: u32,
    max_objects: usize,
) -> Result<BTreeMap<Vec<u8>, ObjectId>> {
    let mut resolver = ResourceReferenceResolver::new(max_objects);
    let resources = page_resource_dictionaries(document, page_id, page, &mut resolver)?;
    xobjects_from_resources(document, resources.into_iter(), page, &mut resolver)
}

fn xobjects_from_resources<'a>(
    document: &'a Document,
    resources: impl Iterator<Item = &'a Dictionary>,
    page: u32,
    resolver: &mut ResourceReferenceResolver,
) -> Result<BTreeMap<Vec<u8>, ObjectId>> {
    enum Visit {
        Enter { id: ObjectId, depth: usize },
        Exit(ObjectId),
    }

    let mut values = BTreeMap::new();
    let mut visited = HashSet::new();
    let mut active = HashSet::new();
    let mut stack = Vec::new();
    for resources in resources {
        let entries = xobject_entries(document, resources, page, resolver)?;
        for (name, id) in entries {
            values.entry(name).or_insert(id);
            stack.push(Visit::Enter { id, depth: 0 });
        }
    }
    while let Some(visit) = stack.pop() {
        match visit {
            Visit::Exit(id) => {
                active.remove(&id);
                visited.insert(id);
            }
            Visit::Enter { id, depth } => {
                if active.contains(&id) {
                    return Err(import_error(
                        Some(page),
                        None,
                        format!("cyclic XObject resource graph at object {} {}", id.0, id.1),
                    ));
                }
                if visited.contains(&id) {
                    continue;
                }
                if depth > MAX_PDF_GRAPH_DEPTH {
                    return Err(import_error(
                        Some(page),
                        None,
                        format!(
                            "XObject resource graph exceeds depth limit {} at object {} {}",
                            MAX_PDF_GRAPH_DEPTH, id.0, id.1
                        ),
                    ));
                }
                let stream = document
                    .objects
                    .get(&id)
                    .and_then(|object| object.as_stream().ok())
                    .ok_or_else(|| {
                        import_error(
                            Some(page),
                            None,
                            format!("validated XObject {} {} is not a stream", id.0, id.1),
                        )
                    })?;
                let subtype = stream
                    .dict
                    .get(b"Subtype")
                    .and_then(Object::as_name)
                    .map_err(|_| {
                        import_error(
                            Some(page),
                            None,
                            format!("validated XObject {} {} has no name Subtype", id.0, id.1),
                        )
                    })?;
                active.insert(id);
                stack.push(Visit::Exit(id));
                if subtype == b"Form"
                    && let Ok(nested) = stream.dict.get(b"Resources")
                {
                    let nested = strict_resource_dictionary(
                        document,
                        nested,
                        page,
                        "Form Resources",
                        resolver,
                    )?;
                    let entries = xobject_entries(document, nested, page, resolver)?;
                    for (_, child) in entries.into_iter().rev() {
                        stack.push(Visit::Enter {
                            id: child,
                            depth: depth + 1,
                        });
                    }
                }
            }
        }
    }
    Ok(values)
}

struct ResourceReferenceResolver {
    resolved: HashMap<ObjectId, ObjectId>,
    active: HashSet<ObjectId>,
    work: usize,
    max_work: usize,
}

impl ResourceReferenceResolver {
    fn new(max_work: usize) -> Self {
        Self {
            resolved: HashMap::new(),
            active: HashSet::new(),
            work: 0,
            max_work,
        }
    }

    fn charge(&mut self, page: u32) -> Result<()> {
        self.charge_many(page, 1)
    }

    fn charge_many(&mut self, page: u32, count: usize) -> Result<()> {
        self.work = self
            .work
            .checked_add(count)
            .ok_or_else(|| import_error(Some(page), None, "resource graph work overflow"))?;
        if self.work > self.max_work {
            return Err(import_error(
                Some(page),
                None,
                format!("resource graph work exceeds object limit {}", self.max_work),
            ));
        }
        Ok(())
    }

    fn resolve<'a>(
        &mut self,
        document: &'a Document,
        start: ObjectId,
        page: u32,
        label: &str,
    ) -> Result<(ObjectId, &'a Object)> {
        if let Some(terminal) = self.resolved.get(&start).copied() {
            let object = document.objects.get(&terminal).ok_or_else(|| {
                import_error(
                    Some(page),
                    None,
                    format!(
                        "cached {label} target is missing object {} {}",
                        terminal.0, terminal.1
                    ),
                )
            })?;
            return Ok((terminal, object));
        }

        let mut chain = Vec::new();
        let mut id = start;
        let terminal = loop {
            if let Some(terminal) = self.resolved.get(&id).copied() {
                break terminal;
            }
            if !self.active.insert(id) {
                for member in &chain {
                    self.active.remove(member);
                }
                return Err(import_error(
                    Some(page),
                    None,
                    format!("cyclic {label} reference at object {} {}", id.0, id.1),
                ));
            }
            chain.push(id);
            if chain.len() > MAX_PDF_GRAPH_DEPTH {
                for member in &chain {
                    self.active.remove(member);
                }
                return Err(import_error(
                    Some(page),
                    None,
                    format!(
                        "{label} reference exceeds depth limit {}",
                        MAX_PDF_GRAPH_DEPTH
                    ),
                ));
            }
            if let Err(error) = self.charge(page) {
                for member in &chain {
                    self.active.remove(member);
                }
                return Err(error);
            }
            let Some(object) = document.objects.get(&id) else {
                for member in &chain {
                    self.active.remove(member);
                }
                return Err(import_error(
                    Some(page),
                    None,
                    format!("{label} reference targets missing object {} {}", id.0, id.1),
                ));
            };
            match object {
                Object::Reference(next) => id = *next,
                _ => break id,
            }
        };
        for member in chain {
            self.active.remove(&member);
            self.resolved.insert(member, terminal);
        }
        let object = document.objects.get(&terminal).ok_or_else(|| {
            import_error(
                Some(page),
                None,
                format!(
                    "resolved {label} target is missing object {} {}",
                    terminal.0, terminal.1
                ),
            )
        })?;
        Ok((terminal, object))
    }
}

fn inherited_page_resources<'a>(
    document: &'a Document,
    page_id: ObjectId,
    page: u32,
    resolver: &mut ResourceReferenceResolver,
) -> Result<Option<&'a Dictionary>> {
    let mut seen = HashSet::new();
    let mut current = page_id;
    for depth in 0..=MAX_PDF_GRAPH_DEPTH {
        if !seen.insert(current) {
            return Err(import_error(
                Some(page),
                None,
                format!(
                    "cyclic page resource graph at object {} {}",
                    current.0, current.1
                ),
            ));
        }
        resolver.charge(page)?;
        let dictionary = document.get_dictionary(current).map_err(|error| {
            import_error(
                Some(page),
                None,
                format!("page resource node failed: {error}"),
            )
        })?;
        if let Ok(resources) = dictionary.get(b"Resources") {
            return strict_resource_dictionary(document, resources, page, "Resources", resolver)
                .map(Some);
        }
        match dictionary.get(b"Parent") {
            Ok(Object::Reference(parent)) => current = *parent,
            Ok(_) => {
                return Err(import_error(
                    Some(page),
                    None,
                    "page Parent must be an indirect reference",
                ));
            }
            Err(_) => return Ok(None),
        }
        if depth == MAX_PDF_GRAPH_DEPTH {
            return Err(import_error(
                Some(page),
                None,
                format!(
                    "page resource graph exceeds depth limit {}",
                    MAX_PDF_GRAPH_DEPTH
                ),
            ));
        }
    }
    unreachable!("bounded resource traversal returns inside the loop")
}

fn font_entries(
    document: &Document,
    resources: &Dictionary,
    page: u32,
    resolver: &mut ResourceReferenceResolver,
) -> Result<BTreeMap<Vec<u8>, ObjectId>> {
    let Ok(value) = resources.get(b"Font") else {
        return Ok(BTreeMap::new());
    };
    let table = strict_resource_dictionary(document, value, page, "Font table", resolver)?;
    resolver.charge_many(page, table.len())?;
    let mut fonts = BTreeMap::new();
    for (name, value) in table.iter() {
        let id = value.as_reference().map_err(|_| {
            import_error(
                Some(page),
                None,
                format!(
                    "Font /{} member must be an indirect reference",
                    String::from_utf8_lossy(name)
                ),
            )
        })?;
        let label = format!("Font /{}", String::from_utf8_lossy(name));
        let (id, object) = strict_referenced_object(document, id, page, &label, resolver)?;
        let dictionary = object.as_dict().map_err(|_| {
            import_error(
                Some(page),
                None,
                format!("{label} must reference a dictionary"),
            )
        })?;
        if dictionary.get(b"Type").and_then(Object::as_name).ok() != Some(b"Font") {
            return Err(import_error(
                Some(page),
                None,
                format!("{label} dictionary must declare /Type /Font"),
            ));
        }
        dictionary
            .get(b"Subtype")
            .and_then(Object::as_name)
            .map_err(|_| {
                import_error(
                    Some(page),
                    None,
                    format!("{label} dictionary must declare a name Subtype"),
                )
            })?;
        fonts.insert(name.clone(), id);
    }
    Ok(fonts)
}

#[cfg(test)]
fn page_resource_dictionaries<'a>(
    document: &'a Document,
    page_id: ObjectId,
    page: u32,
    resolver: &mut ResourceReferenceResolver,
) -> Result<Vec<&'a lopdf::Dictionary>> {
    let mut result = Vec::new();
    let mut seen = HashSet::new();
    let mut current = page_id;
    loop {
        if !seen.insert(current) {
            return Err(import_error(
                Some(page),
                None,
                format!(
                    "cyclic page resource graph at object {} {}",
                    current.0, current.1
                ),
            ));
        }
        let dictionary = document.get_dictionary(current).map_err(|error| {
            import_error(
                Some(page),
                None,
                format!("page resource node failed: {error}"),
            )
        })?;
        if let Ok(resources) = dictionary.get(b"Resources") {
            result.push(strict_resource_dictionary(
                document,
                resources,
                page,
                "Resources",
                resolver,
            )?);
        }
        match dictionary.get(b"Parent") {
            Ok(Object::Reference(parent)) => current = *parent,
            Ok(_) => {
                return Err(import_error(
                    Some(page),
                    None,
                    "page Parent must be an indirect reference",
                ));
            }
            Err(_) => break,
        }
    }
    Ok(result)
}

fn strict_resource_dictionary<'a>(
    document: &'a Document,
    mut value: &'a Object,
    page: u32,
    label: &str,
    resolver: &mut ResourceReferenceResolver,
) -> Result<&'a lopdf::Dictionary> {
    if let Object::Reference(id) = value {
        value = resolver.resolve(document, *id, page, label)?.1;
    }
    value.as_dict().map_err(|_| {
        import_error(
            Some(page),
            None,
            format!("{label} must resolve to a dictionary"),
        )
    })
}

fn strict_referenced_object<'a>(
    document: &'a Document,
    id: ObjectId,
    page: u32,
    label: &str,
    resolver: &mut ResourceReferenceResolver,
) -> Result<(ObjectId, &'a Object)> {
    resolver.resolve(document, id, page, label)
}

fn xobject_entries(
    document: &Document,
    resources: &lopdf::Dictionary,
    page: u32,
    resolver: &mut ResourceReferenceResolver,
) -> Result<Vec<(Vec<u8>, ObjectId)>> {
    let Ok(value) = resources.get(b"XObject") else {
        return Ok(Vec::new());
    };
    let table = strict_resource_dictionary(document, value, page, "XObject table", resolver)?;
    resolver.charge_many(page, table.len())?;
    let mut entries = Vec::with_capacity(table.len());
    for (name, value) in table.iter() {
        let id = value.as_reference().map_err(|_| {
            import_error(
                Some(page),
                None,
                format!(
                    "XObject /{} member must be an indirect reference",
                    String::from_utf8_lossy(name)
                ),
            )
        })?;
        let label = format!("XObject /{}", String::from_utf8_lossy(name));
        let (id, object) = strict_referenced_object(document, id, page, &label, resolver)?;
        let stream = object.as_stream().map_err(|_| {
            import_error(Some(page), None, format!("{label} must reference a stream"))
        })?;
        if stream.dict.get(b"Type").and_then(Object::as_name).ok() != Some(b"XObject") {
            return Err(import_error(
                Some(page),
                None,
                format!("{label} stream must declare /Type /XObject"),
            ));
        }
        let subtype = stream
            .dict
            .get(b"Subtype")
            .and_then(Object::as_name)
            .map_err(|_| {
                import_error(
                    Some(page),
                    None,
                    format!("{label} stream must declare a name Subtype"),
                )
            })?;
        let _ = subtype;
        entries.push((name.clone(), id));
    }
    Ok(entries)
}

fn numbers(operation: &Operation, count: usize, page: u32, offset: usize) -> Result<Vec<f64>> {
    if operation.operands.len() != count {
        return Err(import_error(
            Some(page),
            Some(offset),
            format!("{} requires {count} operands", operation.operator),
        ));
    }
    operation
        .operands
        .iter()
        .take(count)
        .map(|value| {
            object_number(value).map_err(|message| import_error(Some(page), Some(offset), message))
        })
        .collect()
}

fn validate_operator_arity(operation: &Operation, page: u32, offset: usize) -> Result<()> {
    let expected = match operation.operator.as_str() {
        "q" | "Q" | "h" | "S" | "s" | "f" | "F" | "f*" | "B" | "B*" | "b" | "b*" | "n" | "BT"
        | "ET" | "T*" | "W" | "W*" => Some(0),
        "m" | "l" | "Td" | "TD" | "d" | "Tf" => Some(2),
        "c" | "cm" | "Tm" => Some(6),
        "v" | "y" | "re" => Some(4),
        "rg" | "RG" | "\"" => Some(3),
        "w" | "J" | "j" | "g" | "G" | "Tc" | "Tw" | "Tz" | "TL" | "Ts" | "Tr" | "Tj" | "TJ"
        | "'" | "Do" | "gs" | "sh" | "CS" | "cs" => Some(1),
        _ => None,
    };
    if let Some(expected) = expected
        && operation.operands.len() != expected
    {
        return Err(import_error(
            Some(page),
            Some(offset),
            format!(
                "{} requires exactly {expected} operands, got {}",
                operation.operator,
                operation.operands.len()
            ),
        ));
    }
    Ok(())
}

fn require_current_point(
    current: Option<Point>,
    page: u32,
    offset: usize,
    operator: &str,
) -> Result<Point> {
    current.ok_or_else(|| {
        import_error(
            Some(page),
            Some(offset),
            format!("{operator} has no current point"),
        )
    })
}

fn checked_finite(value: f64, page: u32, offset: usize, label: &str) -> Result<f64> {
    if value.is_finite() {
        Ok(value)
    } else {
        Err(import_error(
            Some(page),
            Some(offset),
            format!("{label} produced non-finite geometry"),
        ))
    }
}

fn checked_sum(left: f64, right: f64, page: u32, offset: usize, label: &str) -> Result<f64> {
    checked_finite(left + right, page, offset, label)
}

fn checked_product(left: f64, right: f64, page: u32, offset: usize, label: &str) -> Result<f64> {
    checked_finite(left * right, page, offset, label)
}

fn checked_matrix(matrix: Matrix, page: u32, offset: usize, label: &str) -> Result<Matrix> {
    if [matrix.a, matrix.b, matrix.c, matrix.d, matrix.e, matrix.f]
        .into_iter()
        .all(f64::is_finite)
    {
        Ok(matrix)
    } else {
        Err(import_error(
            Some(page),
            Some(offset),
            format!("{label} produced non-finite geometry"),
        ))
    }
}

fn checked_point(point: Point, page: u32, offset: usize, label: &str) -> Result<Point> {
    if point.x.is_finite() && point.y.is_finite() {
        Ok(point)
    } else {
        Err(import_error(
            Some(page),
            Some(offset),
            format!("{label} produced a non-finite point"),
        ))
    }
}

fn checked_rect(rect: Rect, page: u32, offset: usize, label: &str) -> Result<Rect> {
    if [rect.x, rect.y, rect.width, rect.height]
        .into_iter()
        .all(f64::is_finite)
    {
        Ok(rect)
    } else {
        Err(import_error(
            Some(page),
            Some(offset),
            format!("{label} produced non-finite geometry"),
        ))
    }
}

fn array_numbers(values: &[Object], count: usize, page: u32, offset: usize) -> Result<Vec<f64>> {
    if values.len() != count {
        return Err(import_error(
            Some(page),
            Some(offset),
            format!("array requires exactly {count} values"),
        ));
    }
    values
        .iter()
        .map(|value| {
            object_number(value).map_err(|message| import_error(Some(page), Some(offset), message))
        })
        .collect()
}

fn object_number(value: &Object) -> std::result::Result<f64, String> {
    let number = match value {
        Object::Integer(value) => *value as f64,
        Object::Real(value) => f64::from(*value),
        _ => return Err("expected finite PDF number".to_owned()),
    };
    if !number.is_finite() {
        return Err("non-finite PDF number".to_owned());
    }
    Ok(number)
}

fn number_at(operation: &Operation, index: usize, page: u32, offset: usize) -> Result<f64> {
    operation
        .operands
        .get(index)
        .ok_or_else(|| {
            import_error(
                Some(page),
                Some(offset),
                format!("{} is missing operand {index}", operation.operator),
            )
        })
        .and_then(|value| {
            object_number(value).map_err(|message| import_error(Some(page), Some(offset), message))
        })
}

fn integer_at(operation: &Operation, index: usize, page: u32, offset: usize) -> Result<i64> {
    operation
        .operands
        .get(index)
        .and_then(|value| value.as_i64().ok())
        .ok_or_else(|| {
            import_error(
                Some(page),
                Some(offset),
                format!("{} requires integer operand {index}", operation.operator),
            )
        })
}

fn name_at(operation: &Operation, index: usize, page: u32, offset: usize) -> Result<&[u8]> {
    operation
        .operands
        .get(index)
        .and_then(|value| value.as_name().ok())
        .ok_or_else(|| {
            import_error(
                Some(page),
                Some(offset),
                format!("{} requires name operand {index}", operation.operator),
            )
        })
}

fn string_bytes_at(
    operation: &Operation,
    index: usize,
    page: u32,
    offset: usize,
) -> Result<Vec<u8>> {
    operation
        .operands
        .get(index)
        .and_then(|value| value.as_str().ok())
        .map(Vec::from)
        .ok_or_else(|| {
            import_error(
                Some(page),
                Some(offset),
                format!("{} requires string operand {index}", operation.operator),
            )
        })
}

fn decode_pdf_string(bytes: &[u8]) -> String {
    if bytes.starts_with(&[0xfe, 0xff]) {
        String::from_utf16_lossy(
            &bytes[2..]
                .chunks_exact(2)
                .map(|chunk| u16::from_be_bytes([chunk[0], chunk[1]]))
                .collect::<Vec<_>>(),
        )
    } else {
        bytes.iter().map(|byte| char::from(*byte)).collect()
    }
}

fn decode_pdf_uri_string(bytes: &[u8]) -> Option<String> {
    let (payload, little_endian) = if let Some(payload) = bytes.strip_prefix(&[0xfe, 0xff]) {
        (payload, false)
    } else if let Some(payload) = bytes.strip_prefix(&[0xff, 0xfe]) {
        (payload, true)
    } else {
        return Some(bytes.iter().map(|byte| char::from(*byte)).collect());
    };
    if payload.len() % 2 != 0 {
        return None;
    }
    char::decode_utf16(payload.chunks_exact(2).map(|chunk| {
        if little_endian {
            u16::from_le_bytes([chunk[0], chunk[1]])
        } else {
            u16::from_be_bytes([chunk[0], chunk[1]])
        }
    }))
    .collect::<std::result::Result<String, _>>()
    .ok()
}

fn transformed_point(
    operation: &Operation,
    state: &GraphicsState,
    geometry: PageGeometry,
    page: u32,
    offset: usize,
    start: usize,
) -> Result<Point> {
    let x = number_at(operation, start, page, offset)?;
    let y = number_at(operation, start + 1, page, offset)?;
    checked_point(
        geometry.normalize(state.ctm.point(x, y)),
        page,
        offset,
        "transformed path",
    )
}

fn point_distance(first: Point, second: Point) -> f64 {
    (second.x - first.x).hypot(second.y - first.y)
}

fn advance_text_matrix(
    state: &mut GraphicsState,
    advance: f64,
    page: u32,
    offset: usize,
) -> Result<()> {
    state.text_matrix = checked_matrix(
        state.text_matrix.translated(advance, 0.0),
        page,
        offset,
        "text advance",
    )?;
    Ok(())
}

fn finite_positive(value: f64, page: u32, offset: usize, label: &str) -> Result<f64> {
    if !value.is_finite() || value <= 0.0 {
        Err(import_error(
            (page != 0).then_some(page),
            (page != 0).then_some(offset),
            format!("{label} must be finite and positive"),
        ))
    } else {
        Ok(value)
    }
}

fn finite_nonnegative(value: f64, page: u32, offset: usize, label: &str) -> Result<f64> {
    if !value.is_finite() || value < 0.0 {
        Err(import_error(
            (page != 0).then_some(page),
            (page != 0).then_some(offset),
            format!("{label} must be finite and nonnegative"),
        ))
    } else {
        Ok(value)
    }
}

fn gray(value: f64, page: u32, offset: usize) -> Result<Color> {
    component(value, page, offset).map(|value| Color {
        r: value,
        g: value,
        b: value,
        a: 1.0,
    })
}
fn rgb(operation: &Operation, page: u32, offset: usize) -> Result<Color> {
    let v = numbers(operation, 3, page, offset)?;
    Ok(Color {
        r: component(v[0], page, offset)?,
        g: component(v[1], page, offset)?,
        b: component(v[2], page, offset)?,
        a: 1.0,
    })
}
fn component(value: f64, page: u32, offset: usize) -> Result<f64> {
    if (0.0..=1.0).contains(&value) {
        Ok(value)
    } else {
        Err(import_error(
            Some(page),
            Some(offset),
            "colour component must be between zero and one",
        ))
    }
}

fn parse_dash(operation: &Operation, page: u32, offset: usize) -> Result<Option<DashState>> {
    let values = operation
        .operands
        .first()
        .and_then(|value| value.as_array().ok())
        .ok_or_else(|| import_error(Some(page), Some(offset), "dash operator requires an array"))?;
    let mut dash = Vec::with_capacity(values.len());
    for value in values {
        dash.push(finite_nonnegative(
            object_number(value)
                .map_err(|message| import_error(Some(page), Some(offset), message))?,
            page,
            offset,
            "dash length",
        )?);
    }
    let phase = number_at(operation, 1, page, offset)?;
    if !phase.is_finite() || phase < 0.0 {
        return Err(import_error(
            Some(page),
            Some(offset),
            "dash phase must be finite and nonnegative",
        ));
    }
    if dash.is_empty() {
        if phase != 0.0 {
            return Err(import_error(
                Some(page),
                Some(offset),
                "an empty dash array requires zero phase",
            ));
        }
        Ok(None)
    } else if dash.iter().all(|value| *value == 0.0) {
        Err(import_error(
            Some(page),
            Some(offset),
            "dash array must contain at least one nonzero element",
        ))
    } else {
        Ok(Some(DashState {
            pattern: dash,
            phase,
        }))
    }
}

fn normalized_dash(pattern: &[f64], phase: f64) -> Option<Vec<f64>> {
    let mut cycle = pattern.to_vec();
    if cycle.len() % 2 == 1 {
        cycle.extend_from_slice(pattern);
    }
    let total = cycle.iter().sum::<f64>();
    if phase == 0.0 {
        return Some(cycle);
    }
    let mut remaining_phase = phase % total;
    if remaining_phase == 0.0 {
        return Some(cycle);
    }
    let mut index = 0usize;
    while remaining_phase > cycle[index] {
        remaining_phase -= cycle[index];
        index = (index + 1) % cycle.len();
    }
    if remaining_phase != cycle[index] {
        return None;
    }
    index = (index + 1) % cycle.len();
    if index % 2 == 1 {
        return None;
    }
    cycle.rotate_left(index);
    Some(cycle)
}

fn require_text(state: &GraphicsState, page: u32, offset: usize) -> Result<()> {
    if state.in_text {
        Ok(())
    } else {
        Err(import_error(
            Some(page),
            Some(offset),
            "text operator outside BT/ET",
        ))
    }
}

fn positive_i64(dictionary: &Dictionary, key: &[u8], page: u32, offset: usize) -> Result<i64> {
    let value = dictionary
        .get(key)
        .and_then(Object::as_i64)
        .map_err(|error| {
            import_error(
                Some(page),
                Some(offset),
                format!("image {} failed: {error}", String::from_utf8_lossy(key)),
            )
        })?;
    if value <= 0 {
        Err(import_error(
            Some(page),
            Some(offset),
            "image dimension must be positive",
        ))
    } else {
        Ok(value)
    }
}

fn filter_names(dictionary: &Dictionary) -> std::result::Result<Vec<String>, &'static str> {
    match dictionary.get(b"Filter") {
        Ok(Object::Name(name)) => Ok(vec![String::from_utf8_lossy(name).into_owned()]),
        Ok(Object::Array(values)) => values
            .iter()
            .map(|value| {
                value
                    .as_name()
                    .map(|name| String::from_utf8_lossy(name).into_owned())
                    .map_err(|_| "image Filter array members must be names")
            })
            .collect(),
        Err(_) => Ok(Vec::new()),
        Ok(_) => Err("image Filter must be a name or array of names"),
    }
}

fn unsupported_image_semantics(
    dictionary: &Dictionary,
    filters: &[String],
) -> std::result::Result<Option<String>, String> {
    for key in [b"Mask".as_slice(), b"SMask", b"Decode"] {
        if dictionary.has(key) {
            return Ok(Some(format!(
                "image {} semantics are unsupported",
                String::from_utf8_lossy(key)
            )));
        }
    }
    let image_mask = match dictionary.get(b"ImageMask") {
        Ok(Object::Boolean(value)) => *value,
        Ok(_) => return Err("image ImageMask must be a boolean".to_owned()),
        Err(_) => false,
    };
    if image_mask {
        return Ok(Some("image masks are unsupported".to_owned()));
    }
    if filters.is_empty() || filters == ["FlateDecode"] || filters == ["DCTDecode"] {
        let bits = image_bits(dictionary)?;
        if bits != 8 {
            return Ok(Some(format!("image bit depth {bits} is unsupported")));
        }
        let space = dictionary
            .get(b"ColorSpace")
            .and_then(Object::as_name)
            .unwrap_or_default();
        if !matches!(space, b"DeviceGray" | b"DeviceRGB") {
            return Ok(Some(format!(
                "image colour space {} is unsupported",
                String::from_utf8_lossy(space)
            )));
        }
    }
    Ok(None)
}

fn image_bits(dictionary: &Dictionary) -> std::result::Result<i64, String> {
    match dictionary.get(b"BitsPerComponent") {
        Ok(Object::Integer(value)) => Ok(*value),
        Ok(_) => Err("image BitsPerComponent must be an integer".to_owned()),
        Err(_) => Ok(8),
    }
}

fn jpeg_dimensions(data: &[u8]) -> std::result::Result<(u32, u32), String> {
    if !data.starts_with(&[0xff, 0xd8]) {
        return Err("DCT image is missing the JPEG SOI marker".to_owned());
    }
    let mut cursor = 2usize;
    while cursor < data.len() {
        while data.get(cursor) == Some(&0xff) {
            cursor += 1;
        }
        let marker = *data
            .get(cursor)
            .ok_or_else(|| "DCT image has a truncated JPEG marker".to_owned())?;
        cursor += 1;
        if marker == 0xd9 || marker == 0xda {
            break;
        }
        if marker == 0x01 || (0xd0..=0xd8).contains(&marker) {
            continue;
        }
        let length = data
            .get(cursor..cursor + 2)
            .and_then(|bytes| <[u8; 2]>::try_from(bytes).ok())
            .map(u16::from_be_bytes)
            .map(usize::from)
            .ok_or_else(|| "DCT image has a truncated JPEG segment length".to_owned())?;
        if length < 2 {
            return Err("DCT image has an invalid JPEG segment length".to_owned());
        }
        let end = cursor
            .checked_add(length)
            .ok_or_else(|| "DCT image JPEG segment length overflow".to_owned())?;
        if end > data.len() {
            return Err("DCT image has a truncated JPEG segment".to_owned());
        }
        if matches!(
            marker,
            0xc0 | 0xc1
                | 0xc2
                | 0xc3
                | 0xc5
                | 0xc6
                | 0xc7
                | 0xc9
                | 0xca
                | 0xcb
                | 0xcd
                | 0xce
                | 0xcf
        ) {
            if length < 8 {
                return Err("DCT image has a truncated JPEG SOF segment".to_owned());
            }
            let height = u16::from_be_bytes([data[cursor + 3], data[cursor + 4]]);
            let width = u16::from_be_bytes([data[cursor + 5], data[cursor + 6]]);
            if width == 0 || height == 0 {
                return Err("DCT image has zero intrinsic JPEG dimensions".to_owned());
            }
            return Ok((u32::from(width), u32::from(height)));
        }
        cursor = end;
    }
    Err("DCT image has no supported JPEG SOF marker".to_owned())
}

fn decode_png(
    stream: &Stream,
    width: u32,
    height: u32,
    limit: usize,
) -> std::result::Result<(Vec<u8>, usize), String> {
    let expected = decoded_image_len(&stream.dict, width, height)?;
    if expected > limit {
        return Err(format!(
            "decoded image bytes {expected} exceed limit {limit}"
        ));
    }
    let data = stream
        .decompressed_content_with_limit(limit)
        .map_err(|error| format!("image decompression failed: {error}"))?;
    if data.len() != expected {
        return Err(format!(
            "decoded image length {} does not match expected {expected}",
            data.len()
        ));
    }
    let channels = match stream
        .dict
        .get(b"ColorSpace")
        .and_then(Object::as_name)
        .unwrap_or_default()
    {
        b"DeviceGray" => 1,
        _ => 3,
    };
    let mut rgba = Vec::with_capacity(expected / channels * 4);
    for pixel in data.chunks_exact(channels) {
        let (r, g, b) = if channels == 1 {
            (pixel[0], pixel[0], pixel[0])
        } else {
            (pixel[0], pixel[1], pixel[2])
        };
        rgba.extend_from_slice(&[r, g, b, 255]);
    }
    let size = tiny_skia::IntSize::from_wh(width, height)
        .ok_or_else(|| "invalid image dimensions".to_owned())?;
    let png = tiny_skia::Pixmap::from_vec(rgba, size)
        .ok_or_else(|| "invalid image buffer".to_owned())?
        .encode_png()
        .map_err(|error| format!("PNG encode failed: {error}"))?;
    Ok((png, data.len()))
}

fn decoded_image_len(
    dictionary: &Dictionary,
    width: u32,
    height: u32,
) -> std::result::Result<usize, String> {
    if image_bits(dictionary)? != 8 {
        return Err("only 8-bit Flate images are supported".to_owned());
    }
    let space = dictionary
        .get(b"ColorSpace")
        .and_then(Object::as_name)
        .map_err(|_| "image colour space is missing or unsupported".to_owned())?;
    let channels = match space {
        b"DeviceGray" => 1usize,
        b"DeviceRGB" => 3usize,
        _ => {
            return Err(format!(
                "image colour space {} is unsupported",
                String::from_utf8_lossy(space)
            ));
        }
    };
    usize::try_from(width)
        .ok()
        .and_then(|width| {
            usize::try_from(height)
                .ok()
                .and_then(|height| width.checked_mul(height))
        })
        .and_then(|pixels| pixels.checked_mul(channels))
        .ok_or_else(|| "image byte count overflow".to_owned())
}

fn image_corners(
    geometry: PageGeometry,
    matrix: Matrix,
    page: u32,
    offset: usize,
) -> Result<[Point; 4]> {
    let mut corners = [
        geometry.normalize(matrix.point(0.0, 0.0)),
        geometry.normalize(matrix.point(1.0, 0.0)),
        geometry.normalize(matrix.point(0.0, 1.0)),
        geometry.normalize(matrix.point(1.0, 1.0)),
    ];
    for corner in &mut corners {
        *corner = checked_point(*corner, page, offset, "image corner")?;
    }
    Ok(corners)
}

fn points_bounds(points: &[Point; 4]) -> Rect {
    let min_x = points
        .iter()
        .map(|point| point.x)
        .fold(f64::INFINITY, f64::min);
    let max_x = points
        .iter()
        .map(|point| point.x)
        .fold(f64::NEG_INFINITY, f64::max);
    let min_y = points
        .iter()
        .map(|point| point.y)
        .fold(f64::INFINITY, f64::min);
    let max_y = points
        .iter()
        .map(|point| point.y)
        .fold(f64::NEG_INFINITY, f64::max);
    Rect {
        x: min_x,
        y: min_y,
        width: max_x - min_x,
        height: max_y - min_y,
    }
}

fn strict_uri_action_dictionary(
    dictionary: &Dictionary,
    page: Option<u32>,
    path: &str,
) -> Result<String> {
    let kind = dictionary
        .get(b"S")
        .and_then(Object::as_name)
        .map_err(|_| {
            import_error(
                page,
                None,
                format!("action at {path} must declare a name S"),
            )
        })?;
    if kind != b"URI" {
        return Err(import_error(
            page,
            None,
            format!(
                "non-URI action {} is unsupported at {path}",
                String::from_utf8_lossy(kind)
            ),
        ));
    }
    let uri = dictionary
        .get(b"URI")
        .and_then(Object::as_str)
        .map_err(|_| import_error(page, None, format!("URI action at {path} is malformed")))
        .and_then(|bytes| {
            decode_pdf_uri_string(bytes).ok_or_else(|| {
                import_error(page, None, format!("URI action at {path} is malformed"))
            })
        })?;
    if !safe_uri(&uri) {
        return Err(import_error(
            page,
            None,
            format!("URI action at {path} has an unsupported or malformed URI"),
        ));
    }
    Ok(uri)
}

fn safe_uri(uri: &str) -> bool {
    if uri.is_empty()
        || uri.trim() != uri
        || uri
            .chars()
            .any(|character| character.is_whitespace() || character.is_control())
    {
        return false;
    }
    let lower = uri.to_ascii_lowercase();
    ["https://", "http://", "mailto:"]
        .into_iter()
        .any(|prefix| {
            lower
                .strip_prefix(prefix)
                .is_some_and(|value| !value.is_empty())
        })
}

fn sanitize_content_comments(data: &[u8]) -> Vec<u8> {
    let mut output = Vec::with_capacity(data.len());
    let mut literal_depth = 0usize;
    let mut escaped = false;
    let mut hex_string = false;
    let mut index = 0usize;
    while index < data.len() {
        let byte = data[index];
        if literal_depth > 0 {
            output.push(byte);
            if escaped {
                escaped = false;
            } else if byte == b'\\' {
                escaped = true;
            } else if byte == b'(' {
                literal_depth += 1;
            } else if byte == b')' {
                literal_depth -= 1;
            }
            index += 1;
            continue;
        }
        if hex_string {
            output.push(byte);
            if byte == b'>' {
                hex_string = false;
            }
            index += 1;
            continue;
        }
        match byte {
            b'(' => {
                literal_depth = 1;
                output.push(byte);
                index += 1;
            }
            b'<' if data.get(index + 1) != Some(&b'<') => {
                hex_string = true;
                output.push(byte);
                index += 1;
            }
            b'%' => {
                while index < data.len() && !matches!(data[index], b'\r' | b'\n') {
                    index += 1;
                }
                output.push(b'\n');
                if data.get(index) == Some(&b'\r') {
                    index += 1;
                }
                if data.get(index) == Some(&b'\n') {
                    index += 1;
                }
            }
            _ => {
                output.push(byte);
                index += 1;
            }
        }
    }
    output
}

fn raster_pixels(width: f64, height: f64, dpi: f64) -> Result<u64> {
    finite_positive(dpi, 0, 0, "DPI")?;
    let width = (width * dpi / 72.0).ceil();
    let height = (height * dpi / 72.0).ceil();
    if width > u32::MAX as f64 || height > u32::MAX as f64 {
        return Err(import_error(None, None, "raster dimensions overflow"));
    }
    (width as u64)
        .checked_mul(height as u64)
        .ok_or_else(|| import_error(None, None, "raster pixel count overflow"))
}

fn points_to_emu(points: f64) -> Result<Emu> {
    if !points.is_finite() {
        return Err(import_error(None, None, "non-finite point coordinate"));
    }
    let value = points * EMU_PER_POINT;
    if value < i64::MIN as f64 || value > i64::MAX as f64 {
        return Err(import_error(
            None,
            None,
            "point coordinate exceeds EMU range",
        ));
    }
    Ok(Emu(value.trunc() as i64))
}

fn rect_emu(rect: Rect) -> Result<(Emu, Emu, Emu, Emu)> {
    Ok((
        points_to_emu(rect.x)?,
        points_to_emu(rect.y)?,
        points_to_emu(rect.width.max(1.0 / EMU_PER_POINT))?,
        points_to_emu(rect.height.max(1.0 / EMU_PER_POINT))?,
    ))
}

fn solid_fill(color: Color) -> Result<Fill> {
    let xml = format!(
        r#"<a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:srgbClr val="{}"/></a:solidFill>"#,
        color_hex(color)
    );
    Fill::from_xml(xml.as_bytes()).map_err(|error| {
        import_error(
            None,
            None,
            format!("solid fill construction failed: {error}"),
        )
    })
}

fn color_hex(color: Color) -> String {
    format!(
        "{:02X}{:02X}{:02X}",
        (color.r.clamp(0.0, 1.0) * 255.0).round() as u8,
        (color.g.clamp(0.0, 1.0) * 255.0).round() as u8,
        (color.b.clamp(0.0, 1.0) * 255.0).round() as u8
    )
}

fn append_custom_path(
    presentation: &mut Presentation,
    slide_index: usize,
    element: &PathElement,
) -> Result<()> {
    let path_bounds = element
        .path
        .bounds()
        .ok_or_else(|| import_error(None, None, "painted path has no bounds"))?;
    let stroke_margin = element
        .stroke
        .as_ref()
        .map_or(0.0, |stroke| stroke.width / 2.0);
    let bounds = Rect {
        x: path_bounds.x - stroke_margin,
        y: path_bounds.y - stroke_margin,
        width: path_bounds.width + stroke_margin * 2.0,
        height: path_bounds.height + stroke_margin * 2.0,
    };
    let (left, top, width, height) = rect_emu(bounds)?;
    let mut commands = String::new();
    for command in &element.path.commands {
        match command {
            PathCommand::MoveTo(point) => commands.push_str(&format!(r#"<a:moveTo><a:pt x="{}" y="{}"/></a:moveTo>"#, relative_coord(point.x, bounds.x, bounds.width), relative_coord(point.y, bounds.y, bounds.height))),
            PathCommand::LineTo(point) => commands.push_str(&format!(r#"<a:lnTo><a:pt x="{}" y="{}"/></a:lnTo>"#, relative_coord(point.x, bounds.x, bounds.width), relative_coord(point.y, bounds.y, bounds.height))),
            PathCommand::CurveTo { c1, c2, to } => commands.push_str(&format!(r#"<a:cubicBezTo><a:pt x="{}" y="{}"/><a:pt x="{}" y="{}"/><a:pt x="{}" y="{}"/></a:cubicBezTo>"#, relative_coord(c1.x, bounds.x, bounds.width), relative_coord(c1.y, bounds.y, bounds.height), relative_coord(c2.x, bounds.x, bounds.width), relative_coord(c2.y, bounds.y, bounds.height), relative_coord(to.x, bounds.x, bounds.width), relative_coord(to.y, bounds.y, bounds.height))),
            PathCommand::Close => commands.push_str("<a:close/>"),
        }
    }
    let geometry_xml = format!(
        r#"<a:custGeom xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="100000" h="100000" fill="{}" stroke="{}">{commands}</a:path></a:pathLst></a:custGeom>"#,
        if element.fill.is_some() {
            "norm"
        } else {
            "none"
        },
        element.stroke.is_some()
    );
    let geometry = CT_CustomGeometry2D::from_xml(geometry_xml.as_bytes()).map_err(|error| {
        import_error(
            None,
            None,
            format!("custom path construction failed: {error}"),
        )
    })?;
    let slide_count = presentation.slides.len();
    let slide = presentation
        .slides
        .get_mut(slide_index)
        .ok_or(Error::UnknownSlideIndex {
            index: slide_index,
            slide_count,
        })?;
    let tree = &mut slide.slide.common_slide_data.shape_tree;
    let id = ShapeIdAllocator::scan(tree).allocate();
    let mut shape = CT_Shape::new_preset(
        id,
        &format!("PDF path {id}"),
        "rect",
        positioned_transform(left, top, width, height),
    )
    .map_err(|error| import_error(None, None, format!("path shape failed: {error}")))?;
    shape.shape_properties.preset_geometry = None;
    shape.shape_properties.custom_geometry = Some(geometry);
    shape.shape_properties.fill = element.fill.as_ref().and_then(|paint| match paint {
        Paint::Solid(color) => solid_fill(*color).ok(),
        _ => None,
    });
    shape.shape_properties.line = element.stroke.as_ref().map(stroke_line).transpose()?;
    tree.append_child(ShapeTreeChild::Shape(shape));
    Ok(())
}

fn stroke_line(stroke: &Stroke) -> Result<CT_LineProperties> {
    let color = match &stroke.paint {
        Paint::Solid(color) => *color,
        _ => Color::BLACK,
    };
    let cap = match stroke.cap {
        LineCap::Butt => "flat",
        LineCap::Round => "rnd",
        LineCap::Square => "sq",
    };
    let join = match stroke.join {
        LineJoin::Miter => "<a:miter/>",
        LineJoin::Round => "<a:round/>",
        LineJoin::Bevel => "<a:bevel/>",
    };
    let dash = if let Some(values) = &stroke.dash {
        let mut values = values.clone();
        if values.len() % 2 == 1 {
            let repeated = values.clone();
            values.extend(repeated);
        }
        let width = stroke.width.max(1.0 / EMU_PER_POINT);
        let mut xml = String::from("<a:custDash>");
        for pair in values.chunks_exact(2) {
            let dash = drawingml_dash_units(pair[0], width).ok_or_else(|| {
                import_error(
                    None,
                    None,
                    "dash member cannot produce a positive DrawingML dash stop",
                )
            })?;
            let space = drawingml_dash_units(pair[1], width).ok_or_else(|| {
                import_error(
                    None,
                    None,
                    "dash member cannot produce a positive DrawingML dash stop",
                )
            })?;
            xml.push_str(&format!(r#"<a:ds d="{dash}" sp="{space}"/>"#));
        }
        xml.push_str("</a:custDash>");
        xml
    } else {
        String::new()
    };
    let xml = format!(
        r#"<a:ln xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" w="{}" cap="{cap}"><a:solidFill><a:srgbClr val="{}"/></a:solidFill>{dash}{join}</a:ln>"#,
        (stroke.width * EMU_PER_POINT).trunc() as i64,
        color_hex(color)
    );
    CT_LineProperties::from_xml(xml.as_bytes()).map_err(|error| {
        import_error(
            None,
            None,
            format!("path line construction failed: {error}"),
        )
    })
}

fn drawingml_dash_units(value: f64, width: f64) -> Option<i32> {
    let units = value / width.max(1.0 / EMU_PER_POINT) * 100_000.0;
    (units.is_finite() && units >= 1.0).then(|| units.min(i32::MAX as f64).trunc() as i32)
}

fn relative_coord(value: f64, origin: f64, span: f64) -> i64 {
    if span.abs() < f64::EPSILON {
        0
    } else {
        ((value - origin) * 100_000.0 / span).trunc() as i64
    }
}

fn append_link_overlay(
    presentation: &mut Presentation,
    slide_index: usize,
    rect: Rect,
    uri: &str,
) -> Result<()> {
    let slide_count = presentation.slides.len();
    let slide = presentation
        .slides
        .get_mut(slide_index)
        .ok_or(Error::UnknownSlideIndex {
            index: slide_index,
            slide_count,
        })?;
    let relationships = presentation
        .package
        .part_rels
        .entry(slide.part_name.clone())
        .or_default();
    let relationship_id = relationships.add_external(rel_types::HYPERLINK, uri);
    let (left, top, width, height) = rect_emu(rect)?;
    let tree = &mut slide.slide.common_slide_data.shape_tree;
    let id = ShapeIdAllocator::scan(tree).allocate();
    let xml = format!(
        r#"<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:nvSpPr><p:cNvPr id="{id}" name="PDF link {id}"><a:hlinkClick r:id="{relationship_id}"/></p:cNvPr><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="{}" y="{}"/><a:ext cx="{}" cy="{}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr></p:sp>"#,
        left.0, top.0, width.0, height.0
    );
    let shape = CT_Shape::from_xml(xml.as_bytes()).map_err(|error| {
        import_error(
            None,
            None,
            format!("link overlay construction failed: {error}"),
        )
    })?;
    tree.append_child(ShapeTreeChild::Shape(shape));
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use lopdf::{Object, Stream, dictionary};

    fn source_pdf() -> Vec<u8> {
        let mut document = Document::with_version("1.7");
        let pages_id = document.new_object_id();
        let font_id = document.add_object(dictionary! {
            "Type" => "Font",
            "Subtype" => "Type1",
            "BaseFont" => "Helvetica",
            "Encoding" => "WinAnsiEncoding",
            "FirstChar" => 0,
            "LastChar" => 255,
            "Widths" => (0..=255).map(|_| Object::Integer(500)).collect::<Vec<_>>(),
        });
        let image_id = document.add_object(Stream::new(
            dictionary! {
                "Type" => "XObject",
                "Subtype" => "Image",
                "Width" => 2,
                "Height" => 1,
                "ColorSpace" => "DeviceRGB",
                "BitsPerComponent" => 8,
            },
            vec![255, 0, 0, 0, 255, 0],
        ));
        let content = Content {
            operations: vec![
                Operation::new("rg", vec![1.into(), 0.into(), 0.into()]),
                Operation::new("RG", vec![0.into(), 0.into(), 1.into()]),
                Operation::new("w", vec![2.into()]),
                Operation::new("J", vec![1.into()]),
                Operation::new("j", vec![2.into()]),
                Operation::new("d", vec![Object::Array(vec![6.into(), 2.into()]), 0.into()]),
                Operation::new("re", vec![20.into(), 20.into(), 80.into(), 40.into()]),
                Operation::new("B", vec![]),
                Operation::new("BT", vec![]),
                Operation::new("Tf", vec![Object::Name(b"F1".to_vec()), 18.into()]),
                Operation::new(
                    "Tm",
                    vec![
                        1.into(),
                        0.into(),
                        0.into(),
                        1.into(),
                        30.into(),
                        100.into(),
                    ],
                ),
                Operation::new("Tj", vec![Object::string_literal("Hello PDF")]),
                Operation::new("ET", vec![]),
                Operation::new("q", vec![]),
                Operation::new(
                    "cm",
                    vec![
                        20.into(),
                        0.into(),
                        0.into(),
                        10.into(),
                        130.into(),
                        20.into(),
                    ],
                ),
                Operation::new("Do", vec![Object::Name(b"Im1".to_vec())]),
                Operation::new("Q", vec![]),
            ],
        }
        .encode()
        .unwrap();
        let content_id = document.add_object(Stream::new(dictionary! {}, content));
        let action_id = document.add_object(dictionary! {
            "S" => "URI",
            "URI" => Object::string_literal("https://example.com/pdf"),
        });
        let annotation_id = document.add_object(dictionary! {
            "Type" => "Annot",
            "Subtype" => "Link",
            "Rect" => vec![20.into(), 20.into(), 100.into(), 60.into()],
            "A" => action_id,
        });
        let page_id = document.add_object(dictionary! {
            "Type" => "Page",
            "Parent" => pages_id,
            "MediaBox" => vec![0.into(), 0.into(), 200.into(), 120.into()],
            "Resources" => dictionary! {
                "Font" => dictionary! { "F1" => font_id },
                "XObject" => dictionary! { "Im1" => image_id },
            },
            "Contents" => content_id,
            "Annots" => vec![annotation_id.into()],
        });
        document.objects.insert(
            pages_id,
            Object::Dictionary(dictionary! {
                "Type" => "Pages",
                "Kids" => vec![page_id.into()],
                "Count" => 1,
            }),
        );
        let catalog_id = document.add_object(dictionary! {
            "Type" => "Catalog",
            "Pages" => pages_id,
        });
        document.trailer.set("Root", catalog_id);
        let mut bytes = Vec::new();
        document.save_to(&mut bytes).unwrap();
        bytes
    }

    fn with_extra_content(bytes: &[u8], content: &[u8]) -> Vec<u8> {
        let mut document = Document::load_mem(bytes).unwrap();
        let page_id = page_id(&document);
        document
            .add_page_contents(page_id, content.to_vec())
            .unwrap();
        let mut output = Vec::new();
        document.save_to(&mut output).unwrap();
        output
    }

    fn with_replaced_content(bytes: &[u8], content: &[u8]) -> Vec<u8> {
        mutate_pdf(bytes, |document| {
            let page = page_id(document);
            let content_id = document.add_object(Stream::new(dictionary! {}, content.to_vec()));
            document
                .get_dictionary_mut(page)
                .unwrap()
                .set("Contents", content_id);
        })
    }

    fn imported_render_png(bytes: &[u8], mode: PdfImportMode) -> Vec<u8> {
        let imported = Presentation::from_pdf_bytes(bytes, mode).unwrap();
        let saved = imported.presentation.to_bytes().unwrap();
        let reopened = Presentation::from_bytes(&saved).unwrap();
        match mode {
            PdfImportMode::Editable => {
                let (_, layout) = reopened.render_deterministic().unwrap();
                oxml_pdf::render_page_to_png(&layout, 0, 144.0).unwrap()
            }
            PdfImportMode::PreservedGraphic { .. } => {
                let package = OpcPackage::from_reader(std::io::Cursor::new(saved)).unwrap();
                let relationship = package
                    .get_part_rels("/ppt/slides/slide1.xml")
                    .unwrap()
                    .items
                    .iter()
                    .find(|relationship| relationship.rel_type == rel_types::IMAGE)
                    .unwrap();
                let part =
                    OpcPackage::resolve_rel_target("/ppt/slides/slide1.xml", &relationship.target);
                package.get_part(&part).unwrap().to_vec()
            }
        }
    }

    fn painted_pixel_bounds(png: &[u8]) -> (u32, u32, u32, u32) {
        let pixmap = tiny_skia::Pixmap::decode_png(png).unwrap();
        let mut left = pixmap.width();
        let mut top = pixmap.height();
        let mut right = 0;
        let mut bottom = 0;
        for (index, pixel) in pixmap.data().chunks_exact(4).enumerate() {
            if pixel[3] == 0 || (pixel[0] > 240 && pixel[1] > 240 && pixel[2] > 240) {
                continue;
            }
            let x = index as u32 % pixmap.width();
            let y = index as u32 / pixmap.width();
            left = left.min(x);
            top = top.min(y);
            right = right.max(x);
            bottom = bottom.max(y);
        }
        assert!(left <= right && top <= bottom);
        (left, top, right, bottom)
    }

    fn with_second_page(bytes: &[u8]) -> Vec<u8> {
        let mut document = Document::load_mem(bytes).unwrap();
        let page_id = page_id(&document);
        let page = document.get_object(page_id).unwrap().clone();
        let second_id = document.add_object(page);
        let pages_id = document
            .get_dictionary(page_id)
            .unwrap()
            .get(b"Parent")
            .unwrap()
            .as_reference()
            .unwrap();
        let pages = document.get_dictionary_mut(pages_id).unwrap();
        pages.set("Kids", vec![page_id.into(), second_id.into()]);
        pages.set("Count", 2);
        let mut output = Vec::new();
        document.save_to(&mut output).unwrap();
        output
    }

    fn with_base_font(bytes: &[u8], family: &str) -> Vec<u8> {
        let mut document = Document::load_mem(bytes).unwrap();
        for object in document.objects.values_mut() {
            let Object::Dictionary(dictionary) = object else {
                continue;
            };
            if dictionary.get(b"Type").and_then(Object::as_name).ok() == Some(b"Font") {
                dictionary.set("BaseFont", Object::Name(family.as_bytes().to_vec()));
            }
        }
        let mut output = Vec::new();
        document.save_to(&mut output).unwrap();
        output
    }

    fn mutate_pdf(bytes: &[u8], mutate: impl FnOnce(&mut Document)) -> Vec<u8> {
        let mut document = Document::load_mem(bytes).unwrap();
        mutate(&mut document);
        let mut output = Vec::new();
        document.save_to(&mut output).unwrap();
        output
    }

    fn page_id(document: &Document) -> ObjectId {
        strict_pages(document, usize::MAX, usize::MAX).unwrap()[0].1
    }

    fn first_annotation_id(document: &Document) -> ObjectId {
        document
            .get_dictionary(page_id(document))
            .unwrap()
            .get(b"Annots")
            .unwrap()
            .as_array()
            .unwrap()[0]
            .as_reference()
            .unwrap()
    }

    fn first_action_id(document: &Document) -> ObjectId {
        document
            .get_dictionary(first_annotation_id(document))
            .unwrap()
            .get(b"A")
            .unwrap()
            .as_reference()
            .unwrap()
    }

    fn utf16_uri(value: &str, little_endian: bool) -> Vec<u8> {
        let mut bytes = if little_endian {
            vec![0xff, 0xfe]
        } else {
            vec![0xfe, 0xff]
        };
        for unit in value.encode_utf16() {
            bytes.extend(if little_endian {
                unit.to_le_bytes()
            } else {
                unit.to_be_bytes()
            });
        }
        bytes
    }

    fn strict_font_preload(bytes: &[u8]) -> Result<()> {
        let mut document = Document::load_mem(bytes).unwrap();
        let page = page_id(&document);
        let resources =
            resolve_page_resources(&document, page, 1, PdfImportLimits::default().max_objects)?;
        embedded_fonts(
            &mut document,
            std::iter::once(&resources.fonts),
            usize::MAX,
            PdfImportLimits::default().max_objects,
        )
        .map(|_| ())
    }

    fn first_font_id(document: &Document) -> ObjectId {
        document
            .objects
            .iter()
            .find_map(|(id, object)| {
                object.as_dict().ok().and_then(|dictionary| {
                    (dictionary.get(b"Type").and_then(Object::as_name).ok() == Some(b"Font"))
                        .then_some(*id)
                })
            })
            .unwrap()
    }

    fn with_form_resource_chain(bytes: &[u8], length: usize, cycle: bool) -> Vec<u8> {
        assert!(length > 0);
        mutate_pdf(bytes, |document| {
            let ids = (0..length)
                .map(|_| document.new_object_id())
                .collect::<Vec<_>>();
            for (index, id) in ids.iter().copied().enumerate() {
                let child = ids.get(index + 1).copied().or(cycle.then_some(ids[0]));
                let resources = child.map_or_else(Dictionary::new, |child| {
                    dictionary! {
                        "XObject" => dictionary! { "Next" => child },
                    }
                });
                document.objects.insert(
                    id,
                    Object::Stream(Stream::new(
                        dictionary! {
                            "Type" => "XObject",
                            "Subtype" => "Form",
                            "BBox" => vec![0.into(), 0.into(), 1.into(), 1.into()],
                            "Resources" => resources,
                        },
                        Vec::new(),
                    )),
                );
            }
            let page = page_id(document);
            document
                .get_dictionary_mut(page)
                .unwrap()
                .get_mut(b"Resources")
                .unwrap()
                .as_dict_mut()
                .unwrap()
                .set("XObject", dictionary! { "Root" => ids[0] });
        })
    }

    fn add_indirect_reference_chain(
        document: &mut Document,
        terminal: ObjectId,
        aliases: usize,
    ) -> ObjectId {
        let mut current = terminal;
        for _ in 0..aliases {
            current = document.add_object(Object::Reference(current));
        }
        current
    }

    fn page_content_len(bytes: &[u8]) -> usize {
        let document = Document::load_mem(bytes).unwrap();
        strict_page_content(&document, page_id(&document), 1, usize::MAX, usize::MAX)
            .unwrap()
            .len()
    }

    fn with_second_image_draw(bytes: &[u8]) -> Vec<u8> {
        with_extra_content(bytes, b"q 20 0 0 10 100 60 cm /Im1 Do Q\n")
    }

    fn with_operations(bytes: &[u8], operations: Vec<Operation>) -> Vec<u8> {
        let content = Content { operations }.encode().unwrap();
        with_extra_content(bytes, &content)
    }

    fn with_to_unicode(bytes: &[u8], encoding: &str, include_map: bool) -> (Vec<u8>, usize) {
        let cmap = b"/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n/CMapName /Adobe-Identity-UCS def\n/CMapType 2 def\n1 begincodespacerange\n<00> <FF>\nendcodespacerange\n1 beginbfrange\n<00> <FF> <0000>\nendbfrange\nendcmap\nCMapName currentdict /CMap defineresource pop\nend\nend\n".to_vec();
        let output = mutate_pdf(bytes, |document| {
            let map_id =
                include_map.then(|| document.add_object(Stream::new(dictionary! {}, cmap.clone())));
            for object in document.objects.values_mut() {
                let Object::Dictionary(dictionary) = object else {
                    continue;
                };
                if dictionary.get(b"Type").and_then(Object::as_name).ok() == Some(b"Font") {
                    dictionary.set("Encoding", Object::Name(encoding.as_bytes().to_vec()));
                    if let Some(map_id) = map_id {
                        dictionary.set("ToUnicode", map_id);
                    } else {
                        dictionary.remove(b"ToUnicode");
                    }
                }
            }
        });
        (output, cmap.len())
    }

    fn cmap_for_a(target: &str, padding: usize) -> Vec<u8> {
        let mut cmap = format!(
            "/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n/CMapName /Adobe-Identity-UCS def\n/CMapType 2 def\n1 begincodespacerange\n<00> <FF>\nendcodespacerange\n1 beginbfrange\n<41> <41> <{target}>\nendbfrange\nendcmap\nCMapName currentdict /CMap defineresource pop\nend\nend\n"
        )
        .into_bytes();
        cmap.resize(cmap.len() + padding, b' ');
        cmap
    }

    fn with_font_encoding(bytes: &[u8], encoding: &str) -> Vec<u8> {
        mutate_pdf(bytes, |document| {
            for object in document.objects.values_mut() {
                let Object::Dictionary(dictionary) = object else {
                    continue;
                };
                if dictionary.get(b"Type").and_then(Object::as_name).ok() == Some(b"Font") {
                    dictionary.set("Encoding", Object::Name(encoding.as_bytes().to_vec()));
                }
            }
        })
    }

    #[test]
    fn pdf_points_crop_rotation_and_fractional_coordinates_map_to_truncating_emu() {
        assert_eq!(points_to_emu(1.0).unwrap(), Emu(12_700));
        assert_eq!(points_to_emu(1.25).unwrap(), Emu(15_875));
        assert_eq!(points_to_emu(-1.25).unwrap(), Emu(-15_875));
        let geometry = PageGeometry {
            left: 10.0,
            bottom: 20.0,
            right: 210.0,
            top: 120.0,
            rotation: 90,
            width: 100.0,
            height: 200.0,
        };
        assert_eq!(
            geometry.normalize(Point { x: 10.0, y: 20.0 }),
            Point { x: 0.0, y: 0.0 }
        );
        assert_eq!(
            geometry.normalize(Point { x: 210.0, y: 120.0 }),
            Point { x: 100.0, y: 200.0 }
        );
    }

    #[test]
    fn pdf_matrix_concatenation_text_translation_and_rotated_image_bounds_are_affine() {
        let scale = Matrix {
            a: 2.0,
            b: 0.0,
            c: 0.0,
            d: 3.0,
            e: 0.0,
            f: 0.0,
        };
        let translate = Matrix {
            e: 10.0,
            f: 20.0,
            ..Matrix::IDENTITY
        };
        assert_eq!(
            translate.then(scale).point(1.0, 1.0),
            Point { x: 22.0, y: 63.0 }
        );

        let rotated_text = Matrix {
            a: 0.0,
            b: 1.0,
            c: -1.0,
            d: 0.0,
            e: 40.0,
            f: 50.0,
        };
        assert_eq!(
            rotated_text.translated(5.0, 7.0).point(0.0, 0.0),
            Point { x: 33.0, y: 55.0 }
        );

        let geometry = PageGeometry {
            left: 0.0,
            bottom: 0.0,
            right: 200.0,
            top: 200.0,
            rotation: 0,
            width: 200.0,
            height: 200.0,
        };
        let image = Matrix {
            a: 20.0,
            b: 20.0,
            c: -10.0,
            d: 10.0,
            e: 80.0,
            f: 40.0,
        };
        let corners = image_corners(geometry, image, 1, 0).unwrap();
        let bounds = points_bounds(&corners);
        assert_eq!(
            bounds,
            Rect {
                x: 70.0,
                y: 130.0,
                width: 30.0,
                height: 30.0
            }
        );
        assert!((point_distance(corners[0], corners[1]) - 800.0_f64.sqrt()).abs() < 1e-9);
    }

    #[test]
    fn pdf_dash_phase_cap_join_and_image_semantics_are_preserved_or_diagnosed() {
        assert_eq!(normalized_dash(&[6.0, 2.0], 1.0), None);
        assert_eq!(normalized_dash(&[6.0, 2.0], 6.0), None);
        assert_eq!(normalized_dash(&[6.0, 2.0], 7.0), None);
        assert_eq!(normalized_dash(&[6.0, 2.0], 8.0), Some(vec![6.0, 2.0]));
        assert_eq!(
            normalized_dash(&[6.0, 2.0, 4.0, 2.0], 8.0),
            Some(vec![4.0, 2.0, 6.0, 2.0])
        );
        assert_eq!(normalized_dash(&[3.0], 6.0), Some(vec![3.0, 3.0]));
        let line = stroke_line(&Stroke {
            paint: Paint::Solid(Color::BLACK),
            width: 2.0,
            cap: LineCap::Round,
            join: LineJoin::Bevel,
            dash: normalized_dash(&[6.0, 2.0], 0.0),
        })
        .unwrap();
        assert_eq!(format!("{:?}", line.cap), "Some(Round)");
        assert!(format!("{:?}", line.join).contains("Bevel"));
        assert!(format!("{:?}", line.dash).contains("Custom"));
        assert!(
            stroke_line(&Stroke {
                paint: Paint::Solid(Color::BLACK),
                width: 2.0,
                cap: LineCap::Butt,
                join: LineJoin::Miter,
                dash: Some(vec![0.000_000_001, 2.0]),
            })
            .is_err()
        );

        for key in [b"Mask".as_slice(), b"SMask", b"Decode"] {
            let mut dictionary = dictionary! {
                "Width" => 1,
                "Height" => 1,
                "ColorSpace" => "DeviceRGB",
                "BitsPerComponent" => 8,
            };
            dictionary.set(key, true);
            assert!(
                unsupported_image_semantics(&dictionary, &[])
                    .unwrap()
                    .is_some()
            );
        }
        let dictionary = dictionary! {
            "Width" => 1,
            "Height" => 1,
            "ColorSpace" => "DeviceCMYK",
            "BitsPerComponent" => 8,
        };
        assert!(
            unsupported_image_semantics(&dictionary, &[])
                .unwrap()
                .unwrap()
                .contains("colour space")
        );
    }

    #[test]
    fn nonzero_dash_phase_and_tiny_positive_stops_fail_closed_without_tainting_siblings() {
        let imported = Presentation::from_pdf_bytes(
            &with_replaced_content(
                &source_pdf(),
                b"2 w [6 2] 1 d 10 10 m 30 10 l S [6 2] 8 d 10 20 m 30 20 l S [.000000001 2] 0 d 10 30 m 30 30 l S [6 2] 0 d 10 40 m 30 40 l S\n",
            ),
            PdfImportMode::Editable,
        )
        .unwrap();
        assert_eq!(imported.presentation.slide(0).unwrap().shapes().len(), 3);
        assert!(imported.diagnostics.iter().any(|diagnostic| {
            diagnostic
                .message
                .contains("dash phase requires a zero DrawingML dash stop")
        }));
        assert!(imported.diagnostics.iter().any(|diagnostic| {
            diagnostic
                .message
                .contains("dash member is too small for a positive DrawingML dash stop")
        }));
        let xml = String::from_utf8(
            imported
                .presentation
                .current_part_xml("/ppt/slides/slide1.xml")
                .unwrap(),
        )
        .unwrap();
        assert_eq!(xml.matches("<a:custDash>").count(), 2, "{xml}");
        assert!(!xml.contains(r#" d="0""#), "{xml}");
        assert!(!xml.contains(r#" sp="0""#), "{xml}");
    }

    #[test]
    fn editable_affine_text_and_images_preserve_rotation_reflection_and_shear() {
        let supported = with_extra_content(
            &source_pdf(),
            b"q 0 20 10 0 150 20 cm /Im1 Do Q\nBT /F1 12 Tf 0 1 1 0 120 80 Tm (R) Tj ET\n",
        );
        let imported = Presentation::from_pdf_bytes(&supported, PdfImportMode::Editable).unwrap();
        let xml = String::from_utf8(
            imported
                .presentation
                .current_part_xml("/ppt/slides/slide1.xml")
                .unwrap(),
        )
        .unwrap();
        assert!(xml.contains("flipV=\"1\""), "{xml}");
        assert!(xml.matches("rot=\"").count() >= 2, "{xml}");

        let sheared = with_extra_content(
            &source_pdf(),
            b"q 20 0 10 10 130 20 cm /Im1 Do Q\nBT /F1 12 Tf 1 0 .5 1 120 80 Tm (S) Tj ET\n",
        );
        let imported = Presentation::from_pdf_bytes(&sheared, PdfImportMode::Editable).unwrap();
        assert!(!imported.diagnostics.iter().any(|diagnostic| {
            diagnostic.message.contains("editable image omitted")
                || diagnostic.message.contains("editable text omitted")
        }));
        let saved = imported.presentation.to_bytes().unwrap();
        let reopened = Presentation::from_bytes(&saved).unwrap();
        assert!(reopened.slide(0).unwrap().text().contains('S'));
        assert_eq!(reopened.slide(0).unwrap().shapes().len(), 6);
        let package = OpcPackage::from_reader(std::io::Cursor::new(saved)).unwrap();
        let xml = String::from_utf8(package.get_part("/ppt/slides/slide1.xml").unwrap().to_vec())
            .unwrap();
        assert!(xml.matches("<p:grpSp>").count() >= 4, "{xml}");
        assert!(xml.matches("rot=\"").count() >= 2, "{xml}");
    }

    #[test]
    fn editable_affine_text_render_matches_preserved_uniform_nonuniform_and_shear_geometry() {
        for matrix in ["2 0 0 2", "2 0 0 1", "1 0 .5 1"] {
            let content = format!("BT /F1 20 Tf {matrix} 50 80 Tm (H) Tj ET\n");
            let source = with_replaced_content(&source_pdf(), content.as_bytes());
            let preserved = painted_pixel_bounds(&imported_render_png(
                &source,
                PdfImportMode::PreservedGraphic { dpi: 144.0 },
            ));
            let editable =
                painted_pixel_bounds(&imported_render_png(&source, PdfImportMode::Editable));
            for (expected, actual) in [
                (preserved.0, editable.0),
                (preserved.1, editable.1),
                (preserved.2, editable.2),
                (preserved.3, editable.3),
            ] {
                assert!(
                    expected.abs_diff(actual) <= 3,
                    "affine {matrix} preserved {preserved:?}, editable {editable:?}"
                );
            }
        }
    }

    #[test]
    fn stroke_metrics_follow_uniform_ctm_and_invalid_affine_strokes_are_isolated() {
        let imported = Presentation::from_pdf_bytes(
            &with_replaced_content(
                &source_pdf(),
                b"q 2 0 0 2 0 0 cm 2 w 1 J 2 j [0 2] 0 d 10 10 m 20 10 l S [6 2] 0 d 10 20 m 20 20 l S Q\nq 1 0 .5 1 0 0 cm 2 w 10 30 m 20 30 l S Q\n0 w 10 80 m 20 80 l S\n",
            ),
            PdfImportMode::Editable,
        )
        .unwrap();
        let xml = String::from_utf8(
            imported
                .presentation
                .current_part_xml("/ppt/slides/slide1.xml")
                .unwrap(),
        )
        .unwrap();
        assert!(xml.contains(r#"<a:ln w="50800" cap="rnd""#), "{xml}");
        assert!(xml.contains("<a:custDash>"));
        assert!(xml.contains(r#"<a:ln w="0""#), "{xml}");
        assert_eq!(imported.presentation.slide(0).unwrap().shapes().len(), 3);
        assert!(imported.diagnostics.iter().any(|diagnostic| {
            diagnostic
                .message
                .contains("zero-length dash members are unsupported")
        }));
        assert!(
            imported
                .diagnostics
                .iter()
                .any(|diagnostic| diagnostic.message.contains("nonuniform or sheared CTM"))
        );
    }

    #[test]
    fn malformed_operator_arity_subpaths_and_derived_overflow_fail_before_layout() {
        for (content, expected) in [
            (b"1 q\n".as_slice(), "requires exactly 0 operands"),
            (b"10 10 l S\n".as_slice(), "l has no current point"),
            (b"h\n".as_slice(), "h has no current point"),
        ] {
            let error = Presentation::from_pdf_bytes(
                &with_extra_content(&source_pdf(), content),
                PdfImportMode::Editable,
            )
            .err()
            .expect("malformed content rejects");
            assert!(error.to_string().contains(expected), "{error}");
        }

        let huge = Object::Integer(i64::MAX);
        let operations = (0..40)
            .map(|_| {
                Operation::new(
                    "cm",
                    vec![
                        huge.clone(),
                        0.into(),
                        0.into(),
                        huge.clone(),
                        0.into(),
                        0.into(),
                    ],
                )
            })
            .collect();
        let error = Presentation::from_pdf_bytes(
            &with_operations(&source_pdf(), operations),
            PdfImportMode::Editable,
        )
        .err()
        .expect("derived matrix overflow rejects");
        assert!(error.to_string().contains("non-finite geometry"), "{error}");
    }

    #[test]
    fn tounicode_is_charged_once_cached_and_identity_without_a_map_is_diagnosed() {
        let repeated = with_extra_content(&source_pdf(), b"BT /F1 10 Tf (A) Tj (B) Tj (A) Tj ET\n");
        let (mapped, cmap_len) = with_to_unicode(&repeated, "Identity-H", true);
        let limits = PdfImportLimits {
            max_decompressed_bytes: page_content_len(&mapped) + cmap_len + 6,
            ..PdfImportLimits::default()
        };
        let imported =
            Presentation::from_pdf_bytes_with_limits(&mapped, PdfImportMode::Editable, limits)
                .expect("one charged cached ToUnicode map serves repeated text shows");
        assert!(
            imported
                .presentation
                .slide(0)
                .unwrap()
                .text()
                .contains("A\nB\nA")
        );

        let (missing, _) = with_to_unicode(&source_pdf(), "Identity-H", false);
        let imported = Presentation::from_pdf_bytes(&missing, PdfImportMode::Editable).unwrap();
        assert!(imported.diagnostics.iter().any(|diagnostic| {
            diagnostic.message.contains("Identity-H")
                && diagnostic.message.contains("unsupported encoding")
        }));

        let ordinary = mutate_pdf(
            &with_replaced_content(&source_pdf(), b"BT /F1 10 Tf (A) Tj ET\n"),
            |document| {
                let map = document.add_object(Stream::new(
                    dictionary! {},
                    b"/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n/CMapName /Adobe-Identity-UCS def\n/CMapType 2 def\n1 begincodespacerange\n<00> <FF>\nendcodespacerange\n1 beginbfrange\n<41> <41> <03A9>\nendbfrange\nendcmap\nCMapName currentdict /CMap defineresource pop\nend\nend\n".to_vec(),
                ));
                for object in document.objects.values_mut() {
                    let Object::Dictionary(dictionary) = object else {
                        continue;
                    };
                    if dictionary.get(b"Type").and_then(Object::as_name).ok() == Some(b"Font") {
                        dictionary.set("Encoding", Object::Name(b"WinAnsiEncoding".to_vec()));
                        dictionary.set("ToUnicode", map);
                    }
                }
            },
        );
        let mut ordinary_document = Document::load_mem(&ordinary).unwrap();
        let ordinary_page = page_id(&ordinary_document);
        let ordinary_resources = resolve_page_resources(
            &ordinary_document,
            ordinary_page,
            1,
            PdfImportLimits::default().max_objects,
        )
        .unwrap();
        let ordinary_embedded = embedded_fonts(
            &mut ordinary_document,
            std::iter::once(&ordinary_resources.fonts),
            usize::MAX,
            usize::MAX,
        )
        .unwrap();
        let operations = Content::decode_strict(
            &strict_page_content(&ordinary_document, ordinary_page, 1, usize::MAX, usize::MAX)
                .unwrap(),
        )
        .unwrap()
        .operations;
        let ordinary_fonts = page_fonts(
            &ordinary_document,
            &mut FontEncodingCaches::default(),
            &ordinary_resources.fonts,
            &operations,
            &ordinary_embedded.widths,
        )
        .unwrap();
        let ordinary_font = ordinary_fonts.get(b"F1".as_slice()).unwrap();
        assert!(
            ordinary_font.encoding_supported,
            "encoding {} decoded {:?}",
            ordinary_font.encoding, ordinary_font.decoded
        );
        assert_eq!(ordinary_font.decoded.get(b"A".as_slice()).unwrap(), "Ω");
        let imported = Presentation::from_pdf_bytes(&ordinary, PdfImportMode::Editable).unwrap();
        assert_eq!(imported.presentation.slide(0).unwrap().text(), "Ω");
    }

    #[test]
    fn tounicode_cache_is_owned_by_terminal_stream_while_ordinary_encodings_remain_per_font() {
        const FONT_COUNT: usize = 16;
        let shared_map = cmap_for_a("03A9", 32 * 1024);
        let shared = mutate_pdf(&source_pdf(), |document| {
            let base_id = first_font_id(document);
            let base = document.get_object(base_id).unwrap().clone();
            let map_id = document.add_object(Stream::new(dictionary! {}, shared_map.clone()));
            let mut table = Dictionary::new();
            let mut operations = vec![Operation::new("BT", vec![])];
            for index in 0..FONT_COUNT {
                let id = if index == 0 {
                    base_id
                } else {
                    document.add_object(base.clone())
                };
                let alias = add_indirect_reference_chain(document, map_id, 1);
                document
                    .get_dictionary_mut(id)
                    .unwrap()
                    .set("ToUnicode", alias);
                let name = format!("F{index}");
                table.set(name.clone(), id);
                operations.push(Operation::new(
                    "Tf",
                    vec![Object::Name(name.into_bytes()), 10.into()],
                ));
                operations.push(Operation::new("Tj", vec![Object::string_literal("A")]));
            }
            operations.push(Operation::new("ET", vec![]));
            let content = Content { operations }.encode().unwrap();
            let content = document.add_object(Stream::new(dictionary! {}, content));
            let page = page_id(document);
            let resources = document
                .get_dictionary_mut(page)
                .unwrap()
                .get_mut(b"Resources")
                .unwrap()
                .as_dict_mut()
                .unwrap();
            resources.set("Font", table);
            document
                .get_dictionary_mut(page)
                .unwrap()
                .set("Contents", content);
        });
        let mut document = Document::load_mem(&shared).unwrap();
        let page = page_id(&document);
        let resources =
            resolve_page_resources(&document, page, 1, PdfImportLimits::default().max_objects)
                .unwrap();
        let embedded = embedded_fonts(
            &mut document,
            std::iter::once(&resources.fonts),
            usize::MAX,
            PdfImportLimits::default().max_objects,
        )
        .unwrap();
        assert_eq!(embedded.retained_bytes, shared_map.len());
        let operations = Content::decode_strict(
            &strict_page_content(&document, page, 1, usize::MAX, usize::MAX).unwrap(),
        )
        .unwrap()
        .operations;
        let mut caches = FontEncodingCaches::default();
        let fonts = page_fonts(
            &document,
            &mut caches,
            &resources.fonts,
            &operations,
            &embedded.widths,
        )
        .unwrap();
        assert_eq!(caches.to_unicode_by_stream.len(), 1);
        assert_eq!(caches.to_unicode_parses, 1);
        assert!(caches.ordinary_by_font.is_empty());
        assert_eq!(fonts.len(), FONT_COUNT);
        for font in fonts.values() {
            assert_eq!(font.decoded.get(b"A".as_slice()).unwrap(), "Ω");
        }
        let imported = Presentation::from_pdf_bytes(&shared, PdfImportMode::Editable).unwrap();
        assert_eq!(
            imported
                .presentation
                .slide(0)
                .unwrap()
                .text()
                .lines()
                .filter(|text| *text == "Ω")
                .count(),
            FONT_COUNT
        );

        let distinct = mutate_pdf(&source_pdf(), |document| {
            let first_id = first_font_id(document);
            let second_id = document.add_object(document.get_object(first_id).unwrap().clone());
            let first_map = document.add_object(Stream::new(dictionary! {}, cmap_for_a("03A9", 0)));
            let second_map =
                document.add_object(Stream::new(dictionary! {}, cmap_for_a("03A3", 0)));
            document
                .get_dictionary_mut(first_id)
                .unwrap()
                .set("ToUnicode", first_map);
            document
                .get_dictionary_mut(second_id)
                .unwrap()
                .set("ToUnicode", second_map);
            let page = page_id(document);
            document
                .get_dictionary_mut(page)
                .unwrap()
                .get_mut(b"Resources")
                .unwrap()
                .as_dict_mut()
                .unwrap()
                .set("Font", dictionary! { "F1" => first_id, "F2" => second_id });
            let content = Content {
                operations: vec![
                    Operation::new("BT", vec![]),
                    Operation::new("Tf", vec![Object::Name(b"F1".to_vec()), 10.into()]),
                    Operation::new("Tj", vec![Object::string_literal("A")]),
                    Operation::new("Tf", vec![Object::Name(b"F2".to_vec()), 10.into()]),
                    Operation::new("Tj", vec![Object::string_literal("A")]),
                    Operation::new("ET", vec![]),
                ],
            }
            .encode()
            .unwrap();
            let content = document.add_object(Stream::new(dictionary! {}, content));
            document
                .get_dictionary_mut(page)
                .unwrap()
                .set("Contents", content);
        });
        let mut document = Document::load_mem(&distinct).unwrap();
        let page = page_id(&document);
        let resources = resolve_page_resources(&document, page, 1, usize::MAX).unwrap();
        let embedded = embedded_fonts(
            &mut document,
            std::iter::once(&resources.fonts),
            usize::MAX,
            usize::MAX,
        )
        .unwrap();
        let operations = Content::decode_strict(
            &strict_page_content(&document, page, 1, usize::MAX, usize::MAX).unwrap(),
        )
        .unwrap()
        .operations;
        let mut caches = FontEncodingCaches::default();
        let fonts = page_fonts(
            &document,
            &mut caches,
            &resources.fonts,
            &operations,
            &embedded.widths,
        )
        .unwrap();
        assert_eq!(caches.to_unicode_by_stream.len(), 2);
        assert_eq!(caches.to_unicode_parses, 2);
        assert_eq!(fonts[b"F1".as_slice()].decoded[b"A".as_slice()], "Ω");
        assert_eq!(fonts[b"F2".as_slice()].decoded[b"A".as_slice()], "Σ");

        let no_map = mutate_pdf(&source_pdf(), |document| {
            let first_id = first_font_id(document);
            let second_id = document.add_object(document.get_object(first_id).unwrap().clone());
            let page = page_id(document);
            document
                .get_dictionary_mut(page)
                .unwrap()
                .get_mut(b"Resources")
                .unwrap()
                .as_dict_mut()
                .unwrap()
                .set("Font", dictionary! { "F1" => first_id, "F2" => second_id });
        });
        let mut document = Document::load_mem(&no_map).unwrap();
        let page = page_id(&document);
        let resources = resolve_page_resources(&document, page, 1, usize::MAX).unwrap();
        let embedded = embedded_fonts(
            &mut document,
            std::iter::once(&resources.fonts),
            usize::MAX,
            usize::MAX,
        )
        .unwrap();
        let operations = vec![
            Operation::new("Tf", vec![Object::Name(b"F1".to_vec()), 10.into()]),
            Operation::new("Tj", vec![Object::string_literal("A")]),
            Operation::new("Tf", vec![Object::Name(b"F2".to_vec()), 10.into()]),
            Operation::new("Tj", vec![Object::string_literal("A")]),
        ];
        let mut caches = FontEncodingCaches::default();
        page_fonts(
            &document,
            &mut caches,
            &resources.fonts,
            &operations,
            &embedded.widths,
        )
        .unwrap();
        assert!(caches.to_unicode_by_stream.is_empty());
        assert_eq!(caches.to_unicode_parses, 0);
        assert_eq!(caches.ordinary_by_font.len(), 2);
    }

    #[test]
    fn optional_font_objects_are_absent_or_strictly_resolved_without_silent_carlito_fallback() {
        let absent = mutate_pdf(&source_pdf(), |document| {
            let font = document
                .get_dictionary_mut(first_font_id(document))
                .unwrap();
            font.set("Subtype", Object::Name(b"TrueType".to_vec()));
            font.set("BaseFont", Object::Name(b"Carlito".to_vec()));
            for key in [
                b"ToUnicode".as_slice(),
                b"Encoding",
                b"Widths",
                b"FirstChar",
                b"LastChar",
                b"FontDescriptor",
            ] {
                font.remove(key);
            }
        });
        strict_font_preload(&absent).expect("absent optional font objects allow explicit fallback");
        let imported = Presentation::from_pdf_bytes(&absent, PdfImportMode::Editable).unwrap();
        assert_eq!(imported.presentation.slide(0).unwrap().text(), "Hello PDF");
        assert!(imported.diagnostics.iter().any(|diagnostic| {
            diagnostic
                .message
                .contains("font `Carlito` substituted with deterministic `Carlito`")
        }));

        let descriptor_scalar = mutate_pdf(&source_pdf(), |document| {
            let font = document
                .get_dictionary_mut(first_font_id(document))
                .unwrap();
            font.set("Subtype", Object::Name(b"TrueType".to_vec()));
            font.set("BaseFont", Object::Name(b"Carlito".to_vec()));
            font.remove(b"Widths");
            font.set("FontDescriptor", 7);
        });
        let descriptor_target = mutate_pdf(&source_pdf(), |document| {
            let target = document.add_object(Object::Integer(7));
            document
                .get_dictionary_mut(first_font_id(document))
                .unwrap()
                .set("FontDescriptor", target);
        });
        let descriptor_cycle = mutate_pdf(&source_pdf(), |document| {
            let first = document.new_object_id();
            let second = document.new_object_id();
            document.objects.insert(first, Object::Reference(second));
            document.objects.insert(second, Object::Reference(first));
            document
                .get_dictionary_mut(first_font_id(document))
                .unwrap()
                .set("FontDescriptor", first);
        });
        for (source, expected) in [
            (descriptor_scalar, "descriptor must resolve to a dictionary"),
            (descriptor_target, "descriptor must resolve to a dictionary"),
            (descriptor_cycle, "cyclic font descriptor reference"),
        ] {
            let error = strict_font_preload(&source).unwrap_err();
            assert!(error.to_string().contains(expected), "{error}");
            Presentation::from_pdf_bytes(&source, PdfImportMode::Editable)
                .err()
                .expect("malformed font descriptors reject before publication");
        }

        let font_file_scalar = mutate_pdf(&source_pdf(), |document| {
            let descriptor = document.add_object(dictionary! { "FontFile2" => 7 });
            document
                .get_dictionary_mut(first_font_id(document))
                .unwrap()
                .set("FontDescriptor", descriptor);
        });
        let font_file_target = mutate_pdf(&source_pdf(), |document| {
            let target = document.add_object(Object::Integer(7));
            let descriptor = document.add_object(dictionary! { "FontFile2" => target });
            document
                .get_dictionary_mut(first_font_id(document))
                .unwrap()
                .set("FontDescriptor", descriptor);
        });
        let font_file_cycle = mutate_pdf(&source_pdf(), |document| {
            let first = document.new_object_id();
            let second = document.new_object_id();
            document.objects.insert(first, Object::Reference(second));
            document.objects.insert(second, Object::Reference(first));
            let descriptor = document.add_object(dictionary! { "FontFile2" => first });
            document
                .get_dictionary_mut(first_font_id(document))
                .unwrap()
                .set("FontDescriptor", descriptor);
        });
        for (source, expected) in [
            (
                font_file_scalar,
                "FontFile2 must be an indirect stream reference",
            ),
            (font_file_target, "FontFile2 must reference a stream"),
            (font_file_cycle, "cyclic font FontFile2 reference"),
        ] {
            let error = strict_font_preload(&source).unwrap_err();
            assert!(error.to_string().contains(expected), "{error}");
        }

        let tounicode_scalar = mutate_pdf(&source_pdf(), |document| {
            document
                .get_dictionary_mut(first_font_id(document))
                .unwrap()
                .set("ToUnicode", 7);
        });
        let tounicode_target = mutate_pdf(&source_pdf(), |document| {
            let target = document.add_object(Object::Integer(7));
            document
                .get_dictionary_mut(first_font_id(document))
                .unwrap()
                .set("ToUnicode", target);
        });
        let tounicode_cycle = mutate_pdf(&source_pdf(), |document| {
            let first = document.new_object_id();
            let second = document.new_object_id();
            document.objects.insert(first, Object::Reference(second));
            document.objects.insert(second, Object::Reference(first));
            document
                .get_dictionary_mut(first_font_id(document))
                .unwrap()
                .set("ToUnicode", first);
        });
        for (source, expected) in [
            (
                tounicode_scalar,
                "ToUnicode must be an indirect stream reference",
            ),
            (tounicode_target, "ToUnicode must reference a stream"),
            (tounicode_cycle, "cyclic font ToUnicode reference"),
        ] {
            let error = strict_font_preload(&source).unwrap_err();
            assert!(error.to_string().contains(expected), "{error}");
        }
        let malformed_cmap = mutate_pdf(&source_pdf(), |document| {
            let stream = document.add_object(Stream::new(dictionary! {}, b"not a CMap".to_vec()));
            document
                .get_dictionary_mut(first_font_id(document))
                .unwrap()
                .set("ToUnicode", stream);
        });
        Presentation::from_pdf_bytes(&malformed_cmap, PdfImportMode::Editable)
            .err()
            .expect("a malformed present ToUnicode CMap rejects instead of falling back");

        let malformed_encoding = mutate_pdf(&source_pdf(), |document| {
            document
                .get_dictionary_mut(first_font_id(document))
                .unwrap()
                .set("Encoding", 7);
        });
        let malformed_widths = mutate_pdf(&source_pdf(), |document| {
            document
                .get_dictionary_mut(first_font_id(document))
                .unwrap()
                .set("Widths", 7);
        });
        let malformed_missing_width = mutate_pdf(&source_pdf(), |document| {
            let descriptor = document.add_object(dictionary! {
                "MissingWidth" => Object::Name(b"wide".to_vec()),
            });
            document
                .get_dictionary_mut(first_font_id(document))
                .unwrap()
                .set("FontDescriptor", descriptor);
        });
        for (source, expected) in [
            (malformed_encoding, "Encoding must be a name or dictionary"),
            (malformed_widths, "Widths must resolve to an array"),
            (malformed_missing_width, "MissingWidth"),
        ] {
            let error = strict_font_preload(&source).unwrap_err();
            assert!(error.to_string().contains(expected), "{error}");
        }
    }

    #[test]
    fn distinct_same_family_subset_fonts_keep_concrete_shaping_identity() {
        let source = mutate_pdf(&source_pdf(), |document| {
            let regular = oxml_layout::bundled_fonts::bundled_font_data()[0].1;
            let bold = oxml_layout::bundled_fonts::bundled_font_data()[1].1;
            let regular_stream = document.add_object(Stream::new(
                dictionary! { "Length1" => regular.len() as i64 },
                regular.to_vec(),
            ));
            let bold_stream = document.add_object(Stream::new(
                dictionary! { "Length1" => bold.len() as i64 },
                bold.to_vec(),
            ));
            let regular_descriptor =
                document.add_object(dictionary! { "FontFile2" => regular_stream });
            let bold_descriptor = document.add_object(dictionary! { "FontFile2" => bold_stream });
            let first_id = document
                .objects
                .iter()
                .find_map(|(id, object)| {
                    object.as_dict().ok().and_then(|dictionary| {
                        (dictionary.get(b"Type").and_then(Object::as_name).ok() == Some(b"Font"))
                            .then_some(*id)
                    })
                })
                .unwrap();
            let first = document.get_dictionary_mut(first_id).unwrap();
            first.set("Subtype", Object::Name(b"TrueType".to_vec()));
            first.set("BaseFont", Object::Name(b"AAAAAA+Carlito".to_vec()));
            first.set("FontDescriptor", regular_descriptor);
            let mut second = document.get_object(first_id).unwrap().clone();
            let second_dictionary = second.as_dict_mut().unwrap();
            second_dictionary.set("BaseFont", Object::Name(b"BBBBBB+Carlito".to_vec()));
            second_dictionary.set("FontDescriptor", bold_descriptor);
            let second_id = document.add_object(second);
            let page = page_id(document);
            document
                .get_dictionary_mut(page)
                .unwrap()
                .get_mut(b"Resources")
                .unwrap()
                .as_dict_mut()
                .unwrap()
                .get_mut(b"Font")
                .unwrap()
                .as_dict_mut()
                .unwrap()
                .set("F2", second_id);
        });
        let source = with_replaced_content(
            &source,
            b"BT /F1 12 Tf 1 0 0 1 20 80 Tm (A) Tj /F2 12 Tf 1 0 0 1 60 80 Tm (B) Tj ET\n",
        );
        let mut document = Document::load_mem(&source).unwrap();
        let page = page_id(&document);
        let content = strict_page_content(&document, page, 1, usize::MAX, usize::MAX).unwrap();
        let operations = Content::decode_strict(&content).unwrap().operations;
        let resources =
            resolve_page_resources(&document, page, 1, PdfImportLimits::default().max_objects)
                .unwrap();
        let embedded = embedded_fonts(
            &mut document,
            std::iter::once(&resources.fonts),
            usize::MAX,
            usize::MAX,
        )
        .unwrap();
        assert_eq!(
            embedded
                .files
                .iter()
                .map(|font| font.family.as_str())
                .collect::<Vec<_>>(),
            ["AAAAAA+Carlito", "BBBBBB+Carlito"]
        );
        let page_fonts = page_fonts(
            &document,
            &mut FontEncodingCaches::default(),
            &resources.fonts,
            &operations,
            &embedded.widths,
        )
        .unwrap();
        let first_font = page_fonts.get(b"F1".as_slice()).unwrap();
        let second_font = page_fonts.get(b"F2".as_slice()).unwrap();
        assert_eq!(first_font.display_family, "Carlito");
        assert_eq!(second_font.display_family, "Carlito");
        let mut manager = FontManager::new_deterministic().unwrap();
        manager.load_additional_fonts(&embedded.files);
        let first = manager
            .resolve_font(Some(&first_font.shaping_family), false, false)
            .unwrap();
        let second = manager
            .resolve_font(Some(&second_font.shaping_family), false, false)
            .unwrap();
        assert_ne!(first, second);

        let imported = Presentation::from_pdf_bytes(&source, PdfImportMode::Editable).unwrap();
        assert_eq!(imported.presentation.slide(0).unwrap().text(), "A\nB");
        assert!(
            !imported
                .diagnostics
                .iter()
                .any(|diagnostic| diagnostic.message.contains("substituted"))
        );
    }

    #[test]
    fn unsupported_graphics_and_text_state_isolates_dependent_content_until_reset() {
        let imported = Presentation::from_pdf_bytes(
            &with_extra_content(
                &source_pdf(),
                b"BT /F1 10 Tf 3 Tr (HIDDEN) Tj 0 Tr (VISIBLE) Tj ET\nq /Bad cs 10 70 10 10 re f Q\nq /Bad cs 1 0 0 rg 30 70 10 10 re f Q\nq /GS1 gs 50 70 10 10 re f Q\n",
            ),
            PdfImportMode::Editable,
        )
        .unwrap();
        let text = imported.presentation.slide(0).unwrap().text();
        assert!(!text.contains("HIDDEN"));
        assert!(text.contains("VISIBLE"));
        assert!(
            imported
                .diagnostics
                .iter()
                .any(|diagnostic| diagnostic.message.contains("dependent content is isolated"))
        );

        let clipped = Presentation::from_pdf_bytes(
            &with_extra_content(
                &source_pdf(),
                b"q BT /F1 10 Tf 4 Tr (CLIP) Tj ET 10 70 10 10 re f Q 30 70 10 10 re f\n",
            ),
            PdfImportMode::Editable,
        )
        .unwrap();
        assert!(
            !clipped
                .presentation
                .slide(0)
                .unwrap()
                .text()
                .contains("CLIP")
        );
        assert_eq!(clipped.presentation.slide(0).unwrap().shapes().len(), 5);
        assert!(clipped.diagnostics.iter().any(|diagnostic| {
            diagnostic
                .message
                .contains("text rendering mode 4 is unsupported")
        }));
    }

    #[test]
    fn retained_preserved_elements_and_malformed_image_filters_fail_closed() {
        let sheared = with_extra_content(
            &source_pdf(),
            b"BT /F1 10 Tf 1 0 .5 1 120 80 Tm (S) Tj 1 0 .5 1 140 80 Tm (T) Tj ET\n",
        );
        let error = Presentation::from_pdf_bytes_with_limits(
            &sheared,
            PdfImportMode::Editable,
            PdfImportLimits {
                max_shapes_per_page: 4,
                ..PdfImportLimits::default()
            },
        )
        .err()
        .expect("all retained affine elements count against the shape limit");
        assert!(error.to_string().contains("shape count"), "{error}");

        for malformed in [
            Object::Integer(7),
            Object::Array(vec![
                Object::Name(b"FlateDecode".to_vec()),
                Object::Integer(7),
            ]),
        ] {
            let source = mutate_pdf(&source_pdf(), |document| {
                for object in document.objects.values_mut() {
                    let Object::Stream(stream) = object else {
                        continue;
                    };
                    if stream.dict.get(b"Subtype").and_then(Object::as_name).ok() == Some(b"Image")
                    {
                        stream.dict.set("Filter", malformed.clone());
                    }
                }
            });
            let error = Presentation::from_pdf_bytes(&source, PdfImportMode::Editable)
                .err()
                .expect("malformed image filter rejects");
            assert!(error.to_string().contains("image Filter"), "{error}");
        }
    }

    #[test]
    fn malformed_xobject_tables_members_targets_and_resource_cycles_fail_before_publication() {
        let malformed_type = mutate_pdf(&source_pdf(), |document| {
            document
                .get_dictionary_mut(page_id(document))
                .unwrap()
                .get_mut(b"Resources")
                .unwrap()
                .as_dict_mut()
                .unwrap()
                .set("XObject", 7);
        });
        let missing_table = mutate_pdf(&source_pdf(), |document| {
            document
                .get_dictionary_mut(page_id(document))
                .unwrap()
                .get_mut(b"Resources")
                .unwrap()
                .as_dict_mut()
                .unwrap()
                .set("XObject", Object::Reference((9_999, 0)));
        });
        let nondictionary_table = mutate_pdf(&source_pdf(), |document| {
            let table = document.add_object(Object::Integer(7));
            document
                .get_dictionary_mut(page_id(document))
                .unwrap()
                .get_mut(b"Resources")
                .unwrap()
                .as_dict_mut()
                .unwrap()
                .set("XObject", table);
        });
        let direct_member = mutate_pdf(&source_pdf(), |document| {
            document
                .get_dictionary_mut(page_id(document))
                .unwrap()
                .get_mut(b"Resources")
                .unwrap()
                .as_dict_mut()
                .unwrap()
                .set("XObject", dictionary! { "Bad" => 7 });
        });
        let malformed_target = mutate_pdf(&source_pdf(), |document| {
            let target = document.add_object(dictionary! { "Type" => "XObject" });
            document
                .get_dictionary_mut(page_id(document))
                .unwrap()
                .get_mut(b"Resources")
                .unwrap()
                .as_dict_mut()
                .unwrap()
                .set("XObject", dictionary! { "Bad" => target });
        });
        let cyclic = with_form_resource_chain(&source_pdf(), 2, true);

        for (source, expected) in [
            (malformed_type, "XObject table must resolve to a dictionary"),
            (missing_table, "object ID 9999 0 not found"),
            (
                nondictionary_table,
                "XObject table must resolve to a dictionary",
            ),
            (direct_member, "member must be an indirect reference"),
            (malformed_target, "must reference a stream"),
            (cyclic, "cyclic XObject resource graph"),
        ] {
            let error = Presentation::from_pdf_bytes(&source, PdfImportMode::Editable)
                .err()
                .expect("malformed XObject resources reject transactionally");
            assert!(error.to_string().contains(expected), "{error}");
        }
    }

    #[test]
    fn nested_form_resource_traversal_is_iterative_depth_and_work_bounded() {
        let ordinary = with_form_resource_chain(&source_pdf(), 3, false);
        let ordinary_document = Document::load_mem(&ordinary).unwrap();
        let xobjects = page_xobjects(
            &ordinary_document,
            page_id(&ordinary_document),
            1,
            PdfImportLimits::default().max_objects,
        )
        .unwrap();
        assert!(xobjects.contains_key(b"Root".as_slice()));
        Presentation::from_pdf_bytes(&ordinary, PdfImportMode::Editable)
            .expect("ordinary nested Form resources remain a valid bounded graph");

        let work_error =
            page_xobjects(&ordinary_document, page_id(&ordinary_document), 1, 2).unwrap_err();
        assert!(
            work_error
                .to_string()
                .contains("work exceeds object limit 2"),
            "{work_error}"
        );

        let deep = with_form_resource_chain(&source_pdf(), MAX_PDF_GRAPH_DEPTH + 2, false);
        let deep_document = Document::load_mem(&deep).unwrap();
        let depth_error = page_xobjects(
            &deep_document,
            page_id(&deep_document),
            1,
            PdfImportLimits::default().max_objects,
        )
        .unwrap_err();
        assert!(
            depth_error
                .to_string()
                .contains("XObject resource graph exceeds depth limit 128"),
            "{depth_error}"
        );
        let public_error = Presentation::from_pdf_bytes(&deep, PdfImportMode::Editable)
            .err()
            .expect("deep acyclic Form resources fail closed");
        assert!(
            public_error.to_string().contains("exceeds depth limit"),
            "{public_error}"
        );

        let cyclic = with_form_resource_chain(&source_pdf(), 3, true);
        let cyclic_document = Document::load_mem(&cyclic).unwrap();
        let cycle_error = page_xobjects(
            &cyclic_document,
            page_id(&cyclic_document),
            1,
            PdfImportLimits::default().max_objects,
        )
        .unwrap_err();
        assert!(
            cycle_error
                .to_string()
                .contains("cyclic XObject resource graph"),
            "{cycle_error}"
        );
    }

    #[test]
    fn indirect_xobject_reference_hops_are_shared_cached_cycle_checked_and_work_bounded() {
        const NAMES: usize = 64;
        const ALIASES: usize = 31;
        const REQUIRED_WORK: usize = NAMES + ALIASES + 1;
        let shared = mutate_pdf(&source_pdf(), |document| {
            let page = page_id(document);
            let image = document
                .get_dictionary(page)
                .unwrap()
                .get(b"Resources")
                .unwrap()
                .as_dict()
                .unwrap()
                .get(b"XObject")
                .unwrap()
                .as_dict()
                .unwrap()
                .get(b"Im1")
                .unwrap()
                .as_reference()
                .unwrap();
            let shared_start = add_indirect_reference_chain(document, image, ALIASES);
            let mut table = Dictionary::new();
            for index in 0..NAMES {
                table.set(format!("Shared{index}"), shared_start);
            }
            document
                .get_dictionary_mut(page)
                .unwrap()
                .get_mut(b"Resources")
                .unwrap()
                .as_dict_mut()
                .unwrap()
                .set("XObject", table);
        });
        let document = Document::load_mem(&shared).unwrap();
        let table_error = page_xobjects(&document, page_id(&document), 1, NAMES - 1)
            .expect_err("the complete XObject table is charged before entry allocation");
        assert!(
            table_error
                .to_string()
                .contains("work exceeds object limit 63"),
            "{table_error}"
        );
        let xobjects = page_xobjects(&document, page_id(&document), 1, REQUIRED_WORK).unwrap();
        assert_eq!(xobjects.len(), NAMES);
        assert_eq!(xobjects.values().copied().collect::<HashSet<_>>().len(), 1);
        let error = page_xobjects(&document, page_id(&document), 1, REQUIRED_WORK - 1)
            .expect_err("every unique reference hop and table member is charged");
        assert!(
            error.to_string().contains("work exceeds object limit 95"),
            "{error}"
        );

        const FORM_REQUIRED_WORK: usize = 8;
        let form_resources = mutate_pdf(&source_pdf(), |document| {
            let page = page_id(document);
            let image = document
                .get_dictionary(page)
                .unwrap()
                .get(b"Resources")
                .unwrap()
                .as_dict()
                .unwrap()
                .get(b"XObject")
                .unwrap()
                .as_dict()
                .unwrap()
                .get(b"Im1")
                .unwrap()
                .as_reference()
                .unwrap();
            let nested = document.add_object(dictionary! {
                "XObject" => dictionary! { "Nested" => image },
            });
            let nested_start = add_indirect_reference_chain(document, nested, 3);
            let form = document.add_object(Stream::new(
                dictionary! {
                    "Type" => "XObject",
                    "Subtype" => "Form",
                    "BBox" => vec![0.into(), 0.into(), 1.into(), 1.into()],
                    "Resources" => nested_start,
                },
                Vec::new(),
            ));
            document
                .get_dictionary_mut(page)
                .unwrap()
                .get_mut(b"Resources")
                .unwrap()
                .as_dict_mut()
                .unwrap()
                .set("XObject", dictionary! { "Root" => form });
        });
        let document = Document::load_mem(&form_resources).unwrap();
        page_xobjects(&document, page_id(&document), 1, FORM_REQUIRED_WORK)
            .expect("Form Resources reference hops fit the exact linear budget");
        let error = page_xobjects(&document, page_id(&document), 1, FORM_REQUIRED_WORK - 1)
            .expect_err("Form Resources reference hops consume the shared budget");
        assert!(
            error.to_string().contains("work exceeds object limit 7"),
            "{error}"
        );

        let cyclic = mutate_pdf(&source_pdf(), |document| {
            let first = document.new_object_id();
            let second = document.new_object_id();
            document.objects.insert(first, Object::Reference(second));
            document.objects.insert(second, Object::Reference(first));
            let page = page_id(document);
            document
                .get_dictionary_mut(page)
                .unwrap()
                .get_mut(b"Resources")
                .unwrap()
                .as_dict_mut()
                .unwrap()
                .set("XObject", dictionary! { "Cycle" => first });
        });
        let document = Document::load_mem(&cyclic).unwrap();
        let error = page_xobjects(
            &document,
            page_id(&document),
            1,
            PdfImportLimits::default().max_objects,
        )
        .expect_err("indirect XObject reference cycles fail closed");
        assert!(
            error
                .to_string()
                .contains("cyclic XObject /Cycle reference"),
            "{error}"
        );
    }

    #[test]
    fn jpeg_intrinsic_pixels_and_typed_image_state_fail_closed() {
        let jpeg = vec![
            0xff, 0xd8, 0xff, 0xc0, 0x00, 0x08, 0x08, 0x00, 0x03, 0x00, 0x02, 0x01, 0xff, 0xd9,
        ];
        assert_eq!(jpeg_dimensions(&jpeg).unwrap(), (2, 3));

        let jpeg_source = |declared_width: i64, declared_height: i64| {
            mutate_pdf(&source_pdf(), |document| {
                for object in document.objects.values_mut() {
                    let Object::Stream(stream) = object else {
                        continue;
                    };
                    if stream.dict.get(b"Subtype").and_then(Object::as_name).ok() == Some(b"Image")
                    {
                        stream.dict.set("Width", declared_width);
                        stream.dict.set("Height", declared_height);
                        stream
                            .dict
                            .set("Filter", Object::Name(b"DCTDecode".to_vec()));
                        stream.set_content(jpeg.clone());
                    }
                }
            })
        };
        let mismatched = jpeg_source(1, 1);
        let mismatch_document = Document::load_mem(&mismatched).unwrap();
        let image_id = page_xobjects(
            &mismatch_document,
            page_id(&mismatch_document),
            1,
            PdfImportLimits::default().max_objects,
        )
        .unwrap()[b"Im1".as_slice()];
        let image_object = mismatch_document.get_object(image_id).unwrap();
        assert!(
            image_object.as_stream().is_ok(),
            "image id {image_id:?} resolved to {}",
            image_object.enum_variant()
        );
        let error = Presentation::from_pdf_bytes(&mismatched, PdfImportMode::Editable)
            .err()
            .expect("JPEG declaration mismatch rejects");
        assert!(
            error.to_string().contains("intrinsic dimensions"),
            "{error}"
        );
        let error = Presentation::from_pdf_bytes_with_limits(
            &jpeg_source(2, 3),
            PdfImportMode::Editable,
            PdfImportLimits {
                max_pixels_per_page: 5,
                ..PdfImportLimits::default()
            },
        )
        .err()
        .expect("intrinsic JPEG pixels are charged before publication");
        assert!(error.to_string().contains("pixel count"), "{error}");

        for (key, malformed, expected) in [
            (
                "ImageMask",
                Object::Integer(7),
                "ImageMask must be a boolean",
            ),
            (
                "BitsPerComponent",
                Object::Name(b"Eight".to_vec()),
                "BitsPerComponent must be an integer",
            ),
        ] {
            let malformed = mutate_pdf(&source_pdf(), |document| {
                for object in document.objects.values_mut() {
                    let Object::Stream(stream) = object else {
                        continue;
                    };
                    if stream.dict.get(b"Subtype").and_then(Object::as_name).ok() == Some(b"Image")
                    {
                        stream.dict.set(key, malformed.clone());
                    }
                }
            });
            let error = Presentation::from_pdf_bytes(&malformed, PdfImportMode::Editable)
                .err()
                .expect("wrong-typed image state rejects");
            assert!(error.to_string().contains(expected), "{error}");
        }
    }

    #[test]
    fn simple_font_widths_advance_following_text_and_composite_widths_are_isolated() {
        let source = mutate_pdf(&source_pdf(), |document| {
            let descriptor = document.add_object(dictionary! { "MissingWidth" => 300 });
            for object in document.objects.values_mut() {
                let Object::Dictionary(dictionary) = object else {
                    continue;
                };
                if dictionary.get(b"Type").and_then(Object::as_name).ok() == Some(b"Font") {
                    dictionary.set("FirstChar", 65);
                    dictionary.set("LastChar", 65);
                    dictionary.set("Widths", vec![Object::Integer(700)]);
                    dictionary.set("FontDescriptor", descriptor);
                }
            }
        });
        let imported = Presentation::from_pdf_bytes(
            &with_extra_content(
                &source,
                b"BT /F1 10 Tf 1 0 0 1 10 80 Tm (A) Tj [(B) -100] TJ (C) Tj ET\n",
            ),
            PdfImportMode::Editable,
        )
        .unwrap();
        assert!(
            imported
                .presentation
                .slide(0)
                .unwrap()
                .text()
                .contains("A\nB\nC")
        );
        let xml = String::from_utf8(
            imported
                .presentation
                .current_part_xml("/ppt/slides/slide1.xml")
                .unwrap(),
        )
        .unwrap();
        let c = xml.split("<a:t>C</a:t>").next().unwrap();
        let group = &c[c.rfind("<p:grpSp>").unwrap()..];
        let offset = group.split("<a:off x=\"").nth(1).unwrap();
        assert!(offset.starts_with("266700\""), "{offset}");

        let invalid_missing_width = mutate_pdf(&source, |document| {
            for object in document.objects.values_mut() {
                let Object::Dictionary(dictionary) = object else {
                    continue;
                };
                if dictionary.has(b"MissingWidth") {
                    dictionary.set("MissingWidth", -1);
                }
            }
        });
        let error = Presentation::from_pdf_bytes(&invalid_missing_width, PdfImportMode::Editable)
            .err()
            .expect("negative MissingWidth rejects");
        assert!(error.to_string().contains("MissingWidth"), "{error}");

        let composite = mutate_pdf(&source_pdf(), |document| {
            let font_id = document
                .objects
                .iter()
                .find_map(|(id, object)| {
                    object.as_dict().ok().and_then(|dictionary| {
                        (dictionary.get(b"Type").and_then(Object::as_name).ok() == Some(b"Font"))
                            .then_some(*id)
                    })
                })
                .unwrap();
            let simple_id = document.add_object(document.get_object(font_id).unwrap().clone());
            document
                .get_dictionary_mut(font_id)
                .unwrap()
                .set("Subtype", Object::Name(b"Type0".to_vec()));
            let page = page_id(document);
            let resources = document
                .get_dictionary_mut(page)
                .unwrap()
                .get_mut(b"Resources")
                .unwrap()
                .as_dict_mut()
                .unwrap();
            resources
                .get_mut(b"Font")
                .unwrap()
                .as_dict_mut()
                .unwrap()
                .set("F2", simple_id);
        });
        let composite = with_extra_content(
            &composite,
            b"BT /F1 10 Tf 1 0 0 1 10 100 Tm (BAD) Tj /F2 10 Tf (STALE) Tj 5 0 Td (TD_OK) Tj /F1 10 Tf (BAD_TD) Tj /F2 10 Tf 5 -10 TD (BIG_TD_OK) Tj /F1 10 Tf (BAD_STAR) Tj /F2 10 Tf 10 TL T* (STAR_OK) Tj /F1 10 Tf (BAD_QUOTE) Tj /F2 10 Tf (QUOTE_OK) ' /F1 10 Tf (BAD_DQUOTE) Tj /F2 10 Tf 0 0 (DQUOTE_OK) \" ET\n",
        );
        let imported = Presentation::from_pdf_bytes(&composite, PdfImportMode::Editable).unwrap();
        let text = imported.presentation.slide(0).unwrap().text();
        assert!(!text.contains("BAD"), "{text}");
        assert!(!text.contains("STALE"), "{text}");
        for supported in ["TD_OK", "BIG_TD_OK", "STAR_OK", "QUOTE_OK", "DQUOTE_OK"] {
            assert!(text.contains(supported), "{supported} missing from {text}");
        }
        for isolated in [
            "BAD",
            "STALE",
            "BAD_TD",
            "BAD_STAR",
            "BAD_QUOTE",
            "BAD_DQUOTE",
        ] {
            assert!(!text.contains(isolated), "{isolated} leaked into {text}");
        }
        assert!(imported.diagnostics.iter().any(|diagnostic| {
            diagnostic
                .message
                .contains("composite font /F1 widths are unsupported")
        }));
    }

    #[test]
    fn strict_page_tree_preserves_order_and_rejects_cycles_depth_malformed_nodes_and_work() {
        let mut document = Document::load_mem(&source_pdf()).unwrap();
        let first = page_id(&document);
        let root = document
            .get_dictionary(first)
            .unwrap()
            .get(b"Parent")
            .unwrap()
            .as_reference()
            .unwrap();
        let second = document.add_object(document.get_object(first).unwrap().clone());
        let nested = document.add_object(dictionary! {
            "Type" => "Pages",
            "Parent" => root,
            "Kids" => vec![Object::Reference(second)],
            "Count" => 1,
        });
        document
            .get_dictionary_mut(second)
            .unwrap()
            .set("Parent", nested);
        let root_dictionary = document.get_dictionary_mut(root).unwrap();
        root_dictionary.set(
            "Kids",
            vec![Object::Reference(first), Object::Reference(nested)],
        );
        root_dictionary.set("Count", 2);
        assert_eq!(
            strict_pages(&document, 2, 4).unwrap(),
            vec![(1, first), (2, second)]
        );
        let work_error = strict_pages(&document, 2, 3).unwrap_err();
        assert!(
            work_error
                .to_string()
                .contains("page-tree work exceeds object limit 3"),
            "{work_error}"
        );

        let mut cycle = document;
        cycle
            .get_dictionary_mut(nested)
            .unwrap()
            .set("Kids", vec![Object::Reference(nested)]);
        let cycle_error = strict_pages(&cycle, 2, usize::MAX).unwrap_err();
        assert!(
            cycle_error.to_string().contains("cyclic page tree"),
            "{cycle_error}"
        );

        let mut deep = Document::load_mem(&source_pdf()).unwrap();
        let page = page_id(&deep);
        let root = deep
            .get_dictionary(page)
            .unwrap()
            .get(b"Parent")
            .unwrap()
            .as_reference()
            .unwrap();
        let nodes = (0..=MAX_PDF_GRAPH_DEPTH)
            .map(|_| deep.new_object_id())
            .collect::<Vec<_>>();
        for (index, id) in nodes.iter().copied().enumerate() {
            let parent = index.checked_sub(1).map_or(root, |index| nodes[index]);
            let child = nodes.get(index + 1).copied().unwrap_or(page);
            deep.objects.insert(
                id,
                Object::Dictionary(dictionary! {
                    "Type" => "Pages",
                    "Parent" => parent,
                    "Kids" => vec![Object::Reference(child)],
                    "Count" => 1,
                }),
            );
        }
        deep.get_dictionary_mut(page)
            .unwrap()
            .set("Parent", *nodes.last().unwrap());
        deep.get_dictionary_mut(root)
            .unwrap()
            .set("Kids", vec![Object::Reference(nodes[0])]);
        let depth_error = strict_pages(&deep, 1, usize::MAX).unwrap_err();
        assert!(
            depth_error
                .to_string()
                .contains("page tree exceeds depth limit 128"),
            "{depth_error}"
        );

        let mut malformed = Document::load_mem(&source_pdf()).unwrap();
        let page = page_id(&malformed);
        let root = malformed
            .get_dictionary(page)
            .unwrap()
            .get(b"Parent")
            .unwrap()
            .as_reference()
            .unwrap();
        malformed
            .get_dictionary_mut(root)
            .unwrap()
            .set("Kids", vec![Object::Integer(7)]);
        let malformed_error = strict_pages(&malformed, 1, usize::MAX).unwrap_err();
        assert!(
            malformed_error
                .to_string()
                .contains("Kids members must be indirect references"),
            "{malformed_error}"
        );

        let mut wrong_count = Document::load_mem(&source_pdf()).unwrap();
        let page = page_id(&wrong_count);
        let root = wrong_count
            .get_dictionary(page)
            .unwrap()
            .get(b"Parent")
            .unwrap()
            .as_reference()
            .unwrap();
        wrong_count
            .get_dictionary_mut(root)
            .unwrap()
            .set("Count", 2);
        let count_error = strict_pages(&wrong_count, 2, usize::MAX).unwrap_err();
        assert!(
            count_error.to_string().contains("declares Count 2"),
            "{count_error}"
        );
    }

    #[test]
    fn page_contents_are_strict_ordered_filtered_and_aggregate_bounded() {
        let first = Content {
            operations: vec![
                Operation::new("BT", vec![]),
                Operation::new("Tf", vec![Object::Name(b"F1".to_vec()), 10.into()]),
                Operation::new(
                    "Tm",
                    vec![1.into(), 0.into(), 0.into(), 1.into(), 10.into(), 80.into()],
                ),
                Operation::new("Tj", vec![Object::string_literal("FIRST")]),
                Operation::new("ET", vec![]),
            ],
        }
        .encode()
        .unwrap();
        let second = Content {
            operations: vec![
                Operation::new("BT", vec![]),
                Operation::new("Tf", vec![Object::Name(b"F1".to_vec()), 10.into()]),
                Operation::new(
                    "Tm",
                    vec![1.into(), 0.into(), 0.into(), 1.into(), 10.into(), 60.into()],
                ),
                Operation::new("Tj", vec![Object::string_literal("SECOND")]),
                Operation::new("ET", vec![]),
            ],
        }
        .encode()
        .unwrap();
        let ordered = mutate_pdf(&source_pdf(), |document| {
            let first_id = document.add_object(Stream::new(dictionary! {}, first.clone()));
            let compressed = miniz_oxide::deflate::compress_to_vec_zlib(&second, 6);
            let second_id = document.add_object(Stream::new(
                dictionary! { "Filter" => "FlateDecode" },
                compressed,
            ));
            document.get_dictionary_mut(page_id(document)).unwrap().set(
                "Contents",
                vec![Object::Reference(first_id), Object::Reference(second_id)],
            );
        });
        let imported = Presentation::from_pdf_bytes(&ordered, PdfImportMode::Editable).unwrap();
        assert_eq!(
            imported.presentation.slide(0).unwrap().text(),
            "FIRST\nSECOND"
        );
        let document = Document::load_mem(&ordered).unwrap();
        let decoded = strict_page_content(
            &document,
            page_id(&document),
            1,
            PdfImportLimits::default().max_objects,
            first.len() + second.len() + 2,
        )
        .unwrap();
        assert_eq!(decoded.len(), first.len() + second.len() + 2);
        let aggregate_error = strict_page_content(
            &document,
            page_id(&document),
            1,
            PdfImportLimits::default().max_objects,
            decoded.len() - 1,
        )
        .unwrap_err();
        assert!(
            aggregate_error.to_string().contains("page content"),
            "{aggregate_error}"
        );
        let aggregate_error = Presentation::from_pdf_bytes_with_limits(
            &ordered,
            PdfImportMode::Editable,
            PdfImportLimits {
                max_decompressed_bytes: decoded.len() - 1,
                ..PdfImportLimits::default()
            },
        )
        .err()
        .expect("multi-stream page content shares the public aggregate limit");
        assert!(
            aggregate_error.to_string().contains("page content"),
            "{aggregate_error}"
        );

        let scalar = mutate_pdf(&source_pdf(), |document| {
            document
                .get_dictionary_mut(page_id(document))
                .unwrap()
                .set("Contents", 7);
        });
        let direct_member = mutate_pdf(&source_pdf(), |document| {
            document
                .get_dictionary_mut(page_id(document))
                .unwrap()
                .set("Contents", vec![Object::Integer(7)]);
        });
        let malformed_target = mutate_pdf(&source_pdf(), |document| {
            let target = document.add_object(Object::Integer(7));
            document
                .get_dictionary_mut(page_id(document))
                .unwrap()
                .set("Contents", target);
        });
        let missing_target = mutate_pdf(&source_pdf(), |document| {
            document
                .get_dictionary_mut(page_id(document))
                .unwrap()
                .set("Contents", Object::Reference((9_999, 0)));
        });
        let cyclic_target = mutate_pdf(&source_pdf(), |document| {
            let first = document.new_object_id();
            let second = document.new_object_id();
            document.objects.insert(first, Object::Reference(second));
            document.objects.insert(second, Object::Reference(first));
            document
                .get_dictionary_mut(page_id(document))
                .unwrap()
                .set("Contents", first);
        });
        for source in [
            scalar,
            direct_member,
            malformed_target,
            missing_target,
            cyclic_target,
        ] {
            Presentation::from_pdf_bytes(&source, PdfImportMode::Editable)
                .err()
                .expect("malformed page Contents reject before publication");
        }

        for filter in [
            Object::Name(b"FlateDecode".to_vec()),
            Object::Name(b"ASCIIHexDecode".to_vec()),
            Object::Integer(7),
            Object::Array(vec![
                Object::Name(b"FlateDecode".to_vec()),
                Object::Integer(7),
            ]),
        ] {
            let malformed = mutate_pdf(&source_pdf(), |document| {
                let stream = document.add_object(Stream::new(
                    dictionary! { "Filter" => filter.clone() },
                    b"BT /F1 10 Tf (RAW_FALLBACK_MUST_NOT_RUN) Tj ET\n".to_vec(),
                ));
                document
                    .get_dictionary_mut(page_id(document))
                    .unwrap()
                    .set("Contents", stream);
            });
            let error = Presentation::from_pdf_bytes(&malformed, PdfImportMode::Editable)
                .err()
                .expect("malformed or unsupported content Filter rejects");
            assert!(error.to_string().contains("Contents stream"), "{error}");
        }
    }

    #[test]
    fn page_annotations_are_strict_ordered_and_amplification_bounded() {
        let ordered = mutate_pdf(&source_pdf(), |document| {
            let first_action = document.add_object(dictionary! {
                "S" => "URI",
                "URI" => Object::string_literal("https://example.com/first"),
            });
            let second_action = document.add_object(dictionary! {
                "S" => "URI",
                "URI" => Object::string_literal("https://example.com/second"),
            });
            let first = document.add_object(dictionary! {
                "Type" => "Annot",
                "Subtype" => "Link",
                "Rect" => vec![10.into(), 10.into(), 20.into(), 20.into()],
                "A" => first_action,
            });
            let second = document.add_object(dictionary! {
                "Type" => "Annot",
                "Subtype" => "Link",
                "Rect" => vec![30.into(), 30.into(), 40.into(), 40.into()],
                "A" => second_action,
            });
            let annotations = document.add_object(Object::Array(vec![
                Object::Reference(first),
                Object::Reference(second),
            ]));
            document
                .get_dictionary_mut(page_id(document))
                .unwrap()
                .set("Annots", annotations);
        });
        let imported = Presentation::from_pdf_bytes(&ordered, PdfImportMode::Editable).unwrap();
        let saved = imported.presentation.to_bytes().unwrap();
        let package = OpcPackage::from_reader(std::io::Cursor::new(saved)).unwrap();
        let links = package
            .get_part_rels("/ppt/slides/slide1.xml")
            .unwrap()
            .items
            .iter()
            .filter(|relationship| relationship.rel_type == rel_types::HYPERLINK)
            .map(|relationship| relationship.target.as_str())
            .collect::<Vec<_>>();
        assert_eq!(
            links,
            ["https://example.com/first", "https://example.com/second"]
        );

        let scalar = mutate_pdf(&source_pdf(), |document| {
            document
                .get_dictionary_mut(page_id(document))
                .unwrap()
                .set("Annots", 7);
        });
        let direct_member = mutate_pdf(&source_pdf(), |document| {
            document
                .get_dictionary_mut(page_id(document))
                .unwrap()
                .set("Annots", vec![Object::Integer(7)]);
        });
        let malformed_target = mutate_pdf(&source_pdf(), |document| {
            let target = document.add_object(Object::Integer(7));
            document
                .get_dictionary_mut(page_id(document))
                .unwrap()
                .set("Annots", vec![Object::Reference(target)]);
        });
        let missing_target = mutate_pdf(&source_pdf(), |document| {
            document
                .get_dictionary_mut(page_id(document))
                .unwrap()
                .set("Annots", vec![Object::Reference((9_999, 0))]);
        });
        let malformed_subtype = mutate_pdf(&source_pdf(), |document| {
            let target = document.add_object(dictionary! {
                "Type" => "Annot",
                "Subtype" => 7,
            });
            document
                .get_dictionary_mut(page_id(document))
                .unwrap()
                .set("Annots", vec![Object::Reference(target)]);
        });
        let cyclic_target = mutate_pdf(&source_pdf(), |document| {
            let first = document.new_object_id();
            let second = document.new_object_id();
            document.objects.insert(first, Object::Reference(second));
            document.objects.insert(second, Object::Reference(first));
            document
                .get_dictionary_mut(page_id(document))
                .unwrap()
                .set("Annots", vec![Object::Reference(first)]);
        });
        for source in [
            scalar,
            direct_member,
            malformed_target,
            missing_target,
            malformed_subtype,
            cyclic_target,
        ] {
            Presentation::from_pdf_bytes(&source, PdfImportMode::Editable)
                .err()
                .expect("malformed Annots reject before publication");
        }

        let repeated = mutate_pdf(&source_pdf(), |document| {
            let page = page_id(document);
            let annotation = document
                .get_dictionary(page)
                .unwrap()
                .get(b"Annots")
                .unwrap()
                .as_array()
                .unwrap()[0]
                .as_reference()
                .unwrap();
            document
                .get_dictionary_mut(page)
                .unwrap()
                .set("Annots", vec![Object::Reference(annotation); 2]);
        });
        let repeated_error = Presentation::from_pdf_bytes(&repeated, PdfImportMode::Editable)
            .err()
            .expect("repeated annotation targets reject amplification");
        assert!(
            repeated_error
                .to_string()
                .contains("repeats annotation target"),
            "{repeated_error}"
        );

        let oversized = mutate_pdf(&source_pdf(), |document| {
            let page = page_id(document);
            let annotation = document
                .get_dictionary(page)
                .unwrap()
                .get(b"Annots")
                .unwrap()
                .as_array()
                .unwrap()[0]
                .as_reference()
                .unwrap();
            document
                .get_dictionary_mut(page)
                .unwrap()
                .set("Annots", vec![Object::Reference(annotation); 64]);
        });
        let object_count = Document::load_mem(&oversized).unwrap().objects.len();
        let work_error = Presentation::from_pdf_bytes_with_limits(
            &oversized,
            PdfImportMode::Editable,
            PdfImportLimits {
                max_objects: object_count,
                ..PdfImportLimits::default()
            },
        )
        .err()
        .expect("annotation array work rejects before bounded ID allocation");
        assert!(
            work_error.to_string().contains("work exceeds object limit"),
            "{work_error}"
        );
        let shape_error = Presentation::from_pdf_bytes_with_limits(
            &ordered,
            PdfImportMode::Editable,
            PdfImportLimits {
                max_shapes_per_page: 4,
                ..PdfImportLimits::default()
            },
        )
        .err()
        .expect("annotation entries consume the remaining shape budget before allocation");
        assert!(
            shape_error.to_string().contains("annotation count"),
            "{shape_error}"
        );
    }

    #[test]
    fn page_font_resources_are_strict_inherited_cached_and_work_bounded() {
        let inherited = mutate_pdf(&source_pdf(), |document| {
            let page = page_id(document);
            let root = document
                .get_dictionary(page)
                .unwrap()
                .get(b"Parent")
                .unwrap()
                .as_reference()
                .unwrap();
            let resources = document
                .get_dictionary_mut(page)
                .unwrap()
                .remove(b"Resources")
                .unwrap();
            document
                .get_dictionary_mut(root)
                .unwrap()
                .set("Resources", resources);
        });
        let document = Document::load_mem(&inherited).unwrap();
        let page = page_id(&document);
        let resources = resolve_page_resources(&document, page, 1, 6).unwrap();
        assert!(resources.fonts.contains_key(b"F1".as_slice()));
        assert!(resources.xobjects.contains_key(b"Im1".as_slice()));
        let work_error = resolve_page_resources(&document, page, 1, 5).unwrap_err();
        assert!(
            work_error
                .to_string()
                .contains("work exceeds object limit 5"),
            "{work_error}"
        );
        let imported = Presentation::from_pdf_bytes(&inherited, PdfImportMode::Editable)
            .expect("valid inherited Font and XObject tables import");
        assert_eq!(imported.presentation.slide(0).unwrap().text(), "Hello PDF");

        let malformed_table = mutate_pdf(&source_pdf(), |document| {
            let page = page_id(document);
            document
                .get_dictionary_mut(page)
                .unwrap()
                .get_mut(b"Resources")
                .unwrap()
                .as_dict_mut()
                .unwrap()
                .set("Font", 7);
        });
        let direct_member = mutate_pdf(&source_pdf(), |document| {
            let page = page_id(document);
            document
                .get_dictionary_mut(page)
                .unwrap()
                .get_mut(b"Resources")
                .unwrap()
                .as_dict_mut()
                .unwrap()
                .set("Font", dictionary! { "Bad" => 7 });
        });
        let malformed_target = mutate_pdf(&source_pdf(), |document| {
            let target = document.add_object(Object::Integer(7));
            let page = page_id(document);
            document
                .get_dictionary_mut(page)
                .unwrap()
                .get_mut(b"Resources")
                .unwrap()
                .as_dict_mut()
                .unwrap()
                .set("Font", dictionary! { "Bad" => target });
        });
        let cyclic_target = mutate_pdf(&source_pdf(), |document| {
            let first = document.new_object_id();
            let second = document.new_object_id();
            document.objects.insert(first, Object::Reference(second));
            document.objects.insert(second, Object::Reference(first));
            let page = page_id(document);
            document
                .get_dictionary_mut(page)
                .unwrap()
                .get_mut(b"Resources")
                .unwrap()
                .as_dict_mut()
                .unwrap()
                .set("Font", dictionary! { "Bad" => first });
        });
        for (source, expected) in [
            (malformed_table, "Font table must resolve to a dictionary"),
            (
                direct_member,
                "Font /Bad member must be an indirect reference",
            ),
            (malformed_target, "Font /Bad must reference a dictionary"),
        ] {
            let error = Presentation::from_pdf_bytes(&source, PdfImportMode::Editable)
                .err()
                .expect("malformed font resources reject before publication");
            assert!(error.to_string().contains(expected), "{error}");
        }
        let cyclic_document = Document::load_mem(&cyclic_target).unwrap();
        let cyclic_page = page_id(&cyclic_document);
        let error = resolve_page_resources(
            &cyclic_document,
            cyclic_page,
            1,
            PdfImportLimits::default().max_objects,
        )
        .unwrap_err();
        assert!(
            error.to_string().contains("cyclic Font /Bad reference"),
            "{error}"
        );
        Presentation::from_pdf_bytes(&cyclic_target, PdfImportMode::Editable)
            .err()
            .expect("cyclic font resources reject before publication");

        let deep_resources = mutate_pdf(&source_pdf(), |document| {
            let page = page_id(document);
            let resources = document
                .get_dictionary_mut(page)
                .unwrap()
                .remove(b"Resources")
                .unwrap();
            let terminal = document.add_object(resources);
            let start = add_indirect_reference_chain(document, terminal, MAX_PDF_GRAPH_DEPTH + 1);
            document
                .get_dictionary_mut(page)
                .unwrap()
                .set("Resources", start);
        });
        let deep_document = Document::load_mem(&deep_resources).unwrap();
        let deep_page = page_id(&deep_document);
        let error = resolve_page_resources(
            &deep_document,
            deep_page,
            1,
            PdfImportLimits::default().max_objects,
        )
        .unwrap_err();
        assert!(
            error
                .to_string()
                .contains("Resources reference exceeds depth limit 128"),
            "{error}"
        );
        Presentation::from_pdf_bytes(&deep_resources, PdfImportMode::Editable)
            .err()
            .expect("deep Resources reference chain rejects before publication");
    }

    #[test]
    fn pdf_import_rejects_malformed_encrypted_cyclic_or_unbounded_input_before_publication() {
        let limits = PdfImportLimits {
            max_input_bytes: 3,
            ..PdfImportLimits::default()
        };
        let error =
            Presentation::from_pdf_bytes_with_limits(b"%PDF", PdfImportMode::Editable, limits)
                .err()
                .expect("limit rejects");
        assert!(error.to_string().contains("exceeds limit"));
        let error = Presentation::from_pdf_bytes(b"not a PDF", PdfImportMode::Editable)
            .err()
            .expect("strict parse rejects");
        assert!(error.to_string().contains("strict PDF parse failed"));

        let active_beyond_cap = mutate_pdf(&source_pdf(), |document| {
            document.add_object(dictionary! {
                "S" => "JavaScript",
                "JS" => Object::string_literal("app.alert('x')"),
            });
        });
        let error = Presentation::from_pdf_bytes_with_limits(
            &active_beyond_cap,
            PdfImportMode::Editable,
            PdfImportLimits {
                max_objects: 1,
                ..PdfImportLimits::default()
            },
        )
        .err()
        .expect("object cap runs before active-content traversal");
        assert!(error.to_string().contains("object count"), "{error}");

        let mut shared = Document::with_version("1.7");
        let shared_id = shared.add_object(dictionary! {
            "Safe" => Object::Array((0..64).map(Object::Integer).collect()),
        });
        for _ in 0..64 {
            shared.add_object(dictionary! { "Shared" => shared_id });
        }
        let work = reject_active_content(&shared, shared.objects.len()).unwrap();
        assert!(work < 512, "shared graph required {work} scan steps");

        let source = source_pdf();
        for (limits, expected) in [
            (
                PdfImportLimits {
                    max_objects: 1,
                    ..PdfImportLimits::default()
                },
                "object count",
            ),
            (
                PdfImportLimits {
                    max_operations_per_page: 1,
                    ..PdfImportLimits::default()
                },
                "operation count",
            ),
            (
                PdfImportLimits {
                    max_shapes_per_page: 1,
                    ..PdfImportLimits::default()
                },
                "shape count",
            ),
            (
                PdfImportLimits {
                    max_pixels_per_page: 2,
                    ..PdfImportLimits::default()
                },
                "raster pixel count",
            ),
        ] {
            let mode = if expected == "raster pixel count" {
                PdfImportMode::PreservedGraphic { dpi: 96.0 }
            } else {
                PdfImportMode::Editable
            };
            let error = Presentation::from_pdf_bytes_with_limits(&source, mode, limits)
                .err()
                .expect("limit rejects");
            assert!(error.to_string().contains(expected), "{error}");
        }

        let page_limited = PdfImportLimits {
            max_pages: 1,
            ..PdfImportLimits::default()
        };
        let error = Presentation::from_pdf_bytes_with_limits(
            &with_second_page(&source),
            PdfImportMode::Editable,
            page_limited,
        )
        .err()
        .expect("page cap rejects");
        assert!(error.to_string().contains("page count"));

        let diagnostic_limited = PdfImportLimits {
            max_diagnostics: 1,
            ..PdfImportLimits::default()
        };
        let error = Presentation::from_pdf_bytes_with_limits(
            &with_extra_content(&source, b"zz\nzz\n"),
            PdfImportMode::Editable,
            diagnostic_limited,
        )
        .err()
        .expect("diagnostic cap rejects");
        assert!(error.to_string().contains("diagnostic count"));

        let encrypted = mutate_pdf(&source, |document| {
            document.trailer.set(
                "Encrypt",
                dictionary! { "Filter" => "Standard", "V" => 1, "R" => 2, "O" => Object::string_literal("owner"), "U" => Object::string_literal("user"), "P" => -4 },
            );
        });
        let error = Presentation::from_pdf_bytes(&encrypted, PdfImportMode::Editable)
            .err()
            .expect("encrypted PDF rejects");
        assert!(error.to_string().contains("encrypted PDF"), "{error}");

        let cyclic = mutate_pdf(&source, |document| {
            let page = page_id(document);
            let dictionary = document.get_dictionary_mut(page).unwrap();
            dictionary.remove(b"MediaBox");
            dictionary.set("Parent", page);
        });
        let error = Presentation::from_pdf_bytes(&cyclic, PdfImportMode::Editable)
            .err()
            .expect("cyclic inheritance rejects");
        assert!(
            error.to_string().contains("Parent does not match"),
            "{error}"
        );
    }

    #[test]
    fn aggregate_decompressed_bytes_and_image_pixels_are_hard_limits() {
        let source = source_pdf();
        let content_len = page_content_len(&source);
        let limits = PdfImportLimits {
            max_decompressed_bytes: content_len + 5,
            ..PdfImportLimits::default()
        };
        let error =
            Presentation::from_pdf_bytes_with_limits(&source, PdfImportMode::Editable, limits)
                .err()
                .expect("content plus image exceeds aggregate byte cap");
        assert!(
            error.to_string().contains("aggregate decompressed bytes"),
            "{error}"
        );

        let limits = PdfImportLimits {
            max_pixels_per_page: 3,
            ..PdfImportLimits::default()
        };
        let error = Presentation::from_pdf_bytes_with_limits(
            &with_second_image_draw(&source),
            PdfImportMode::Editable,
            limits,
        )
        .err()
        .expect("two individually small images exceed aggregate pixel cap");
        assert!(
            error
                .to_string()
                .contains("aggregate page image pixel count"),
            "{error}"
        );
    }

    #[test]
    fn recursive_active_content_and_non_uri_actions_fail_closed() {
        let nested = mutate_pdf(&source_pdf(), |document| {
            let catalog = document
                .trailer
                .get(b"Root")
                .unwrap()
                .as_reference()
                .unwrap();
            document.get_dictionary_mut(catalog).unwrap().set(
                "Names",
                dictionary! {
                    "Nested" => vec![Object::Dictionary(dictionary! {
                        "AA" => dictionary! {
                            "S" => "JavaScript",
                            "JS" => Object::string_literal("app.alert('no')"),
                        },
                    })],
                },
            );
        });
        let error = Presentation::from_pdf_bytes(&nested, PdfImportMode::Editable)
            .err()
            .expect("nested additional actions reject");
        assert!(
            error
                .to_string()
                .contains("additional actions are unsupported"),
            "{error}"
        );

        let non_uri = mutate_pdf(&source_pdf(), |document| {
            for object in document.objects.values_mut() {
                let Object::Dictionary(dictionary) = object else {
                    continue;
                };
                if dictionary.get(b"S").and_then(Object::as_name).ok() == Some(b"URI") {
                    dictionary.set("S", Object::Name(b"GoTo".to_vec()));
                    dictionary.remove(b"URI");
                    dictionary.set("D", Object::Name(b"Destination".to_vec()));
                }
            }
        });
        let error = Presentation::from_pdf_bytes(&non_uri, PdfImportMode::Editable)
            .err()
            .expect("non-URI action rejects");
        assert!(error.to_string().contains("non-URI action"), "{error}");
    }

    #[test]
    fn action_contexts_are_strict_and_do_not_reclassify_incidental_s_entries() {
        let malformed_actions = [
            ("scalar", Object::Integer(7), "resolve to a dictionary"),
            (
                "destination array",
                Object::Array(vec![Object::Name(b"Destination".to_vec())]),
                "resolve to a dictionary",
            ),
            (
                "named destination",
                Object::Name(b"Destination".to_vec()),
                "resolve to a dictionary",
            ),
            (
                "empty dictionary",
                Object::Dictionary(dictionary! {}),
                "declare a name S",
            ),
            (
                "wrong S type",
                Object::Dictionary(dictionary! { "S" => 7 }),
                "declare a name S",
            ),
            (
                "non-URI action",
                Object::Dictionary(dictionary! {
                    "S" => "GoTo",
                    "D" => Object::Name(b"Destination".to_vec()),
                }),
                "non-URI action",
            ),
            (
                "missing URI",
                Object::Dictionary(dictionary! { "S" => "URI" }),
                "is malformed",
            ),
            (
                "wrong URI type",
                Object::Dictionary(dictionary! { "S" => "URI", "URI" => 7 }),
                "is malformed",
            ),
            (
                "unsafe URI scheme",
                Object::Dictionary(dictionary! {
                    "S" => "URI",
                    "URI" => Object::string_literal("javascript:alert(1)"),
                }),
                "unsupported or malformed URI",
            ),
            (
                "malformed URI",
                Object::Dictionary(dictionary! {
                    "S" => "URI",
                    "URI" => Object::string_literal("https://example.com/a b"),
                }),
                "unsupported or malformed URI",
            ),
        ];

        for (label, action, expected) in &malformed_actions {
            let open_action = mutate_pdf(&source_pdf(), |document| {
                let catalog = document
                    .trailer
                    .get(b"Root")
                    .unwrap()
                    .as_reference()
                    .unwrap();
                document
                    .get_dictionary_mut(catalog)
                    .unwrap()
                    .set("OpenAction", action.clone());
            });
            let error = Presentation::from_pdf_bytes(&open_action, PdfImportMode::Editable)
                .err()
                .unwrap_or_else(|| panic!("catalog OpenAction {label} must reject"));
            assert!(error.to_string().contains(expected), "{label}: {error}");

            let annotation_action = mutate_pdf(&source_pdf(), |document| {
                let annotation = first_annotation_id(document);
                document
                    .get_dictionary_mut(annotation)
                    .unwrap()
                    .set("A", action.clone());
            });
            let error = Presentation::from_pdf_bytes(&annotation_action, PdfImportMode::Editable)
                .err()
                .unwrap_or_else(|| panic!("annotation action {label} must reject"));
            assert!(error.to_string().contains(expected), "{label}: {error}");
        }

        for annotation in [false, true] {
            let missing = mutate_pdf(&source_pdf(), |document| {
                if annotation {
                    let id = first_annotation_id(document);
                    document
                        .get_dictionary_mut(id)
                        .unwrap()
                        .set("A", Object::Reference((9_999, 0)));
                } else {
                    let catalog = document
                        .trailer
                        .get(b"Root")
                        .unwrap()
                        .as_reference()
                        .unwrap();
                    document
                        .get_dictionary_mut(catalog)
                        .unwrap()
                        .set("OpenAction", Object::Reference((9_999, 0)));
                }
            });
            let error = Presentation::from_pdf_bytes(&missing, PdfImportMode::Editable)
                .err()
                .expect("missing action target must reject");
            assert!(
                error
                    .to_string()
                    .contains("active-content reference failed"),
                "{error}"
            );

            let non_dictionary = mutate_pdf(&source_pdf(), |document| {
                let target = document.add_object(Object::Integer(7));
                if annotation {
                    let id = first_annotation_id(document);
                    document.get_dictionary_mut(id).unwrap().set("A", target);
                } else {
                    let catalog = document
                        .trailer
                        .get(b"Root")
                        .unwrap()
                        .as_reference()
                        .unwrap();
                    document
                        .get_dictionary_mut(catalog)
                        .unwrap()
                        .set("OpenAction", target);
                }
            });
            let error = Presentation::from_pdf_bytes(&non_dictionary, PdfImportMode::Editable)
                .err()
                .expect("non-dictionary action target must reject");
            assert!(
                error.to_string().contains("resolve to a dictionary"),
                "{error}"
            );

            let cyclic = mutate_pdf(&source_pdf(), |document| {
                let first = document.new_object_id();
                let second = document.new_object_id();
                document.objects.insert(first, Object::Reference(second));
                document.objects.insert(second, Object::Reference(first));
                if annotation {
                    let id = first_annotation_id(document);
                    document.get_dictionary_mut(id).unwrap().set("A", first);
                } else {
                    let catalog = document
                        .trailer
                        .get(b"Root")
                        .unwrap()
                        .as_reference()
                        .unwrap();
                    document
                        .get_dictionary_mut(catalog)
                        .unwrap()
                        .set("OpenAction", first);
                }
            });
            let error = Presentation::from_pdf_bytes(&cyclic, PdfImportMode::Editable)
                .err()
                .expect("cyclic action references must reject");
            assert!(
                error.to_string().contains("cyclic action reference"),
                "{error}"
            );
        }

        let valid = mutate_pdf(&source_pdf(), |document| {
            let action = document.add_object(dictionary! {
                "S" => "URI",
                "URI" => Object::string_literal("mailto:pdf@example.com"),
            });
            let catalog = document
                .trailer
                .get(b"Root")
                .unwrap()
                .as_reference()
                .unwrap();
            let catalog = document.get_dictionary_mut(catalog).unwrap();
            catalog.set("OpenAction", action);
            catalog.set(
                "Unrelated",
                dictionary! { "S" => Object::Name(b"JavaScript".to_vec()) },
            );
        });
        let imported = Presentation::from_pdf_bytes(&valid, PdfImportMode::Editable)
            .expect("valid URI actions and incidental S entries import");
        assert_eq!(imported.presentation.slide(0).unwrap().text(), "Hello PDF");
    }

    #[test]
    fn additional_actions_are_always_rejected_before_shape_interpretation() {
        let conforming = mutate_pdf(&source_pdf(), |document| {
            let action = document.add_object(dictionary! {
                "S" => "URI",
                "URI" => Object::string_literal("https://example.com/additional"),
            });
            document
                .get_dictionary_mut(page_id(document))
                .unwrap()
                .set("AA", dictionary! { "O" => action });
        });
        let action_shaped = mutate_pdf(&source_pdf(), |document| {
            document.get_dictionary_mut(page_id(document)).unwrap().set(
                "AA",
                dictionary! {
                    "S" => "URI",
                    "URI" => Object::string_literal("https://example.com/malformed-aa"),
                },
            );
        });
        let scalar = mutate_pdf(&source_pdf(), |document| {
            document
                .get_dictionary_mut(page_id(document))
                .unwrap()
                .set("AA", 7);
        });
        let indirect = mutate_pdf(&source_pdf(), |document| {
            let action = document.add_object(dictionary! {
                "S" => "URI",
                "URI" => Object::string_literal("https://example.com/indirect-aa"),
            });
            let target = document.add_object(dictionary! { "O" => action });
            document
                .get_dictionary_mut(page_id(document))
                .unwrap()
                .set("AA", target);
        });
        let cyclic = mutate_pdf(&source_pdf(), |document| {
            let first = document.new_object_id();
            let second = document.new_object_id();
            document.objects.insert(first, Object::Reference(second));
            document.objects.insert(second, Object::Reference(first));
            document
                .get_dictionary_mut(page_id(document))
                .unwrap()
                .set("AA", first);
        });

        for (label, source) in [
            ("conforming trigger dictionary", conforming),
            ("malformed action-shaped dictionary", action_shaped),
            ("scalar", scalar),
            ("indirect dictionary", indirect),
            ("indirect cycle", cyclic),
        ] {
            let error = Presentation::from_pdf_bytes(&source, PdfImportMode::Editable)
                .err()
                .unwrap_or_else(|| panic!("{label} additional actions must reject"));
            assert!(
                error
                    .to_string()
                    .contains("additional actions are unsupported"),
                "{label}: {error}"
            );
        }

        let imported = Presentation::from_pdf_bytes(&source_pdf(), PdfImportMode::Editable)
            .expect("an ordinary URI link without AA remains valid");
        let saved = imported.presentation.to_bytes().unwrap();
        let package = OpcPackage::from_reader(std::io::Cursor::new(saved)).unwrap();
        assert!(
            package
                .get_part_rels("/ppt/slides/slide1.xml")
                .unwrap()
                .items
                .iter()
                .any(|relationship| relationship.target == "https://example.com/pdf")
        );
    }

    #[test]
    fn uri_text_strings_reject_malformed_utf16_and_preserve_valid_unicode() {
        let mut odd_be = utf16_uri("https://example.com/be", false);
        odd_be.push(0);
        let mut odd_le = utf16_uri("https://example.com/le", true);
        odd_le.push(0);
        let mut lone_high = utf16_uri("https://example.com/high/", false);
        lone_high.extend(0xd800_u16.to_be_bytes());
        let mut lone_low = utf16_uri("https://example.com/low/", true);
        lone_low.extend(0xdc00_u16.to_le_bytes());

        for (label, bytes) in [
            ("odd UTF-16BE payload", odd_be),
            ("odd UTF-16LE payload", odd_le),
            ("lone high surrogate", lone_high),
            ("lone low surrogate", lone_low),
            ("empty URI", Vec::new()),
            ("NUL URI", b"https://example.com/\0path".to_vec()),
            ("control URI", b"https://example.com/\x1fpath".to_vec()),
        ] {
            let malformed = mutate_pdf(&source_pdf(), |document| {
                document
                    .get_dictionary_mut(first_action_id(document))
                    .unwrap()
                    .set("URI", Object::string_literal(bytes));
            });
            let error = Presentation::from_pdf_bytes(&malformed, PdfImportMode::Editable)
                .err()
                .unwrap_or_else(|| panic!("{label} must reject"));
            assert!(error.to_string().contains("malformed"), "{label}: {error}");
        }

        for (expected, bytes) in [
            (
                "https://example.com/\u{03a9}",
                utf16_uri("https://example.com/\u{03a9}", false),
            ),
            (
                "https://example.com/\u{1f600}",
                utf16_uri("https://example.com/\u{1f600}", true),
            ),
        ] {
            let valid = mutate_pdf(&source_pdf(), |document| {
                document
                    .get_dictionary_mut(first_action_id(document))
                    .unwrap()
                    .set("URI", Object::string_literal(bytes));
            });
            let imported = Presentation::from_pdf_bytes(&valid, PdfImportMode::Editable)
                .unwrap_or_else(|error| panic!("valid UTF-16 URI rejected: {error}"));
            let saved = imported.presentation.to_bytes().unwrap();
            let package = OpcPackage::from_reader(std::io::Cursor::new(saved)).unwrap();
            assert!(
                package
                    .get_part_rels("/ppt/slides/slide1.xml")
                    .unwrap()
                    .items
                    .iter()
                    .any(|relationship| relationship.target == expected),
                "decoded URI relationship {expected} was not preserved"
            );
        }
    }

    #[test]
    fn unsupported_pdf_font_encoding_has_a_stable_source_diagnostic() {
        let imported = Presentation::from_pdf_bytes(
            &with_font_encoding(&source_pdf(), "MadeUpEncoding"),
            PdfImportMode::Editable,
        )
        .expect("unsupported encoding is safely decoded with a diagnostic");
        assert!(imported.diagnostics.iter().any(|diagnostic| {
            diagnostic.path.contains("content/op[")
                && diagnostic.message.contains("MadeUpEncoding")
                && diagnostic.message.contains("direct decoding")
        }));
    }

    #[test]
    fn imported_ctm_and_text_state_project_affine_geometry_and_orientation() {
        let source = with_extra_content(
            &source_pdf(),
            br#"
                q 2 0 0 2 0 0 cm 1 0 0 1 50 0 cm 0 0 10 10 re f Q
                BT /F1 10 Tf 0 1 -1 0 100 60 Tm 12 TL [(A) 250 (B)] TJ T* (C) Tj ET
            "#,
        );
        let imported = Presentation::from_pdf_bytes(&source, PdfImportMode::Editable)
            .expect("affine content imports");
        let slide = imported.presentation.slide(0).unwrap();
        assert_eq!(slide.text(), "Hello PDF\nA\nB\nC");
        let slide_xml = String::from_utf8(
            imported
                .presentation
                .current_part_xml("/ppt/slides/slide1.xml")
                .unwrap(),
        )
        .unwrap();
        assert!(slide_xml.contains(r#"<a:off x="1270000""#), "{slide_xml}");
        assert!(slide_xml.contains("rot=\""), "{slide_xml}");
        for value in ["Hello PDF", "A", "B", "C"] {
            assert!(slide_xml.contains(&format!("<a:t>{value}</a:t>")));
        }
        assert!(slide_xml.matches("rot=\"").count() >= 3, "{slide_xml}");
    }

    #[test]
    fn both_pdf_import_modes_publish_valid_reopenable_presentations() {
        let bytes = source_pdf();
        for mode in [
            PdfImportMode::Editable,
            PdfImportMode::PreservedGraphic { dpi: 96.0 },
        ] {
            let imported = Presentation::from_pdf_bytes(&bytes, mode).expect("PDF imports");
            assert_eq!(imported.presentation.len(), 1);
            assert_eq!(
                imported.presentation.slide_size(),
                Some((Emu(2_540_000), Emu(1_524_000)))
            );
            assert!(imported.presentation.validate().is_empty());
            let saved = imported.presentation.to_bytes().expect("saves");
            let reopened = Presentation::from_bytes(&saved).expect("reopens");
            assert_eq!(reopened.len(), 1);
        }
    }

    #[test]
    fn editable_pdf_text_images_paths_and_links_survive_save_and_reopen() {
        let imported = Presentation::from_pdf_bytes(&source_pdf(), PdfImportMode::Editable)
            .expect("editable PDF imports");
        let saved = imported.presentation.to_bytes().unwrap();
        let reopened = Presentation::from_bytes(&saved).unwrap();
        assert_eq!(reopened.slide(0).unwrap().text(), "Hello PDF");
        assert_eq!(reopened.slide(0).unwrap().shapes().len(), 4);
        let package = OpcPackage::from_reader(std::io::Cursor::new(saved)).unwrap();
        let slide_xml =
            String::from_utf8(package.get_part("/ppt/slides/slide1.xml").unwrap().to_vec())
                .unwrap();
        let commands = r#"<a:moveTo><a:pt x="1219" y="97619"/></a:moveTo><a:lnTo><a:pt x="98780" y="97619"/></a:lnTo><a:lnTo><a:pt x="98780" y="2380"/></a:lnTo><a:lnTo><a:pt x="1219" y="2380"/></a:lnTo><a:close/>"#;
        assert!(slide_xml.contains(commands), "{slide_xml}");
        assert!(slide_xml.contains(r#"<a:srgbClr val="FF0000""#));
        assert!(slide_xml.contains(r#"<a:srgbClr val="0000FF""#));

        let relationships = package.get_part_rels("/ppt/slides/slide1.xml").unwrap();
        let image_relationship = relationships
            .items
            .iter()
            .find(|relationship| relationship.rel_type == rel_types::IMAGE)
            .unwrap();
        let image_part =
            OpcPackage::resolve_rel_target("/ppt/slides/slide1.xml", &image_relationship.target);
        let expected_image = tiny_skia::Pixmap::from_vec(
            vec![255, 0, 0, 255, 0, 255, 0, 255],
            tiny_skia::IntSize::from_wh(2, 1).unwrap(),
        )
        .unwrap()
        .encode_png()
        .unwrap();
        assert_eq!(package.get_part(&image_part).unwrap(), expected_image);

        let hyperlink = relationships
            .items
            .iter()
            .find(|relationship| relationship.target == "https://example.com/pdf")
            .unwrap();
        assert_eq!(hyperlink.target_mode.as_deref(), Some("External"));
        let marker = format!(r#"r:id="{}""#, hyperlink.id);
        let marker_offset = slide_xml.find(&marker).unwrap();
        let shape_start = slide_xml[..marker_offset].rfind("<p:sp>").unwrap();
        let shape_end = slide_xml[marker_offset..].find("</p:sp>").unwrap() + marker_offset;
        let overlay = &slide_xml[shape_start..shape_end];
        assert!(
            overlay.contains(r#"<a:off x="254000" y="762000""#),
            "{overlay}"
        );
        assert!(
            overlay.contains(r#"<a:ext cx="1016000" cy="508000""#),
            "{overlay}"
        );
    }

    #[test]
    fn unsupported_pdf_operators_are_diagnosed_without_losing_supported_siblings() {
        let content =
            Content::decode_strict(b"0 0 10 10 re f 1 0 0 1 0 0 cm /GS1 gs 20 20 10 10 re f")
                .unwrap();
        assert_eq!(content.operations.len(), 6);
        assert_eq!(content.operations[2].operator, "cm");
        assert_eq!(content.operations[3].operator, "gs");

        let imported = Presentation::from_pdf_bytes(
            &with_extra_content(&source_pdf(), b"zz\n10 10 20 20 re f\n"),
            PdfImportMode::Editable,
        )
        .expect("unsupported safe operator is diagnostic");
        assert_eq!(imported.diagnostics.len(), 2);
        assert_eq!(
            imported.diagnostics[1].message,
            "unsupported PDF operator zz"
        );
        assert!(imported.diagnostics[1].path.contains("content/op["));
        assert_eq!(imported.presentation.slide(0).unwrap().text(), "Hello PDF");
        assert_eq!(imported.presentation.slide(0).unwrap().shapes().len(), 5);
    }

    #[test]
    fn pdf_font_substitution_is_explicit_and_deterministic() {
        let mut manager = FontManager::new_deterministic().unwrap();
        let first = manager.resolve_font(Some("Carlito"), false, false).unwrap();
        let second = manager.resolve_font(Some("Carlito"), false, false).unwrap();
        assert_eq!(first, second);

        let source = with_base_font(&source_pdf(), "DefinitelyMissingPDFTypeface");
        let first = Presentation::from_pdf_bytes(&source, PdfImportMode::Editable).unwrap();
        let second = Presentation::from_pdf_bytes(&source, PdfImportMode::Editable).unwrap();
        assert_eq!(first.diagnostics, second.diagnostics);
        assert!(first.diagnostics.iter().any(|diagnostic| {
            diagnostic.message.contains("DefinitelyMissingPDFTypeface")
                && diagnostic.message.contains("deterministic")
        }));
    }

    #[test]
    fn pdf_content_comments_cannot_silently_truncate_operator_decoding() {
        let sanitized = sanitize_content_comments(b"0 0 m % comment\n 10 10 l\nS\n");
        let content = Content::decode_strict(&sanitized).unwrap();
        assert_eq!(
            content
                .operations
                .iter()
                .map(|operation| operation.operator.as_str())
                .collect::<Vec<_>>(),
            ["m", "l", "S"]
        );
    }
}
