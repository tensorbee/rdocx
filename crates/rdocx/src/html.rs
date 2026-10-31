//! Bounded HTML and CSS import into the native Word document model.

use std::cmp::Ordering;
use std::collections::{HashMap, HashSet};
use std::fs::File;
use std::io::Read;
use std::ops::Range;
use std::path::Path;

use base64::Engine as _;
use base64::engine::general_purpose::STANDARD as BASE64;
use rdocx_oxml::content_control::{CT_Sdt, SdtContent};
use rdocx_oxml::document::BodyContent;
use rdocx_oxml::table::{
    CT_Row, CT_Tbl, CT_TblGrid, CT_TblGridCol, CT_TblPr, CT_TblWidth, CT_Tc, VMerge,
};
use rdocx_oxml::text::{CT_P, RunContent};
use rdocx_oxml::units::Twips;
use scraper::{ElementRef, Html, Node, Selector};
use sha2::{Digest, Sha256};

use crate::document::visit_accepted_drawings;
use crate::paragraph::{Alignment, Paragraph};
use crate::revision::{accepted_revision_content, revision_removes_text};
use crate::run::{DrawingKind, DrawingRelationshipKind, Run};
use crate::table::Cell;
use crate::{
    BodyItemRef, CellItemRef, CellRef, Document, Error, HyperlinkItemRef, HyperlinkRef, Length,
    ListLevel, ParagraphItemRef, ParagraphRef, Result, RowRef, RunItemRef, RunRef, StoryId,
    TableRef,
};
use rdocx_oxml::CT_Revision;

const MAX_INPUT_BYTES: usize = 64 * 1024 * 1024;
const MAX_MHTML_PARTS: usize = 1_024;
const MAX_MHTML_HEADER_BYTES: usize = 64 * 1024;
const MAX_MHTML_OUTPUT_BYTES: usize = 128 * 1024 * 1024;

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct HtmlDiagnostic {
    pub location: String,
    pub property: Option<String>,
    pub message: String,
}

pub struct HtmlReadResult {
    pub document: Document,
    pub diagnostics: Vec<HtmlDiagnostic>,
}

/// One caller-owned image available to an HTML fragment insertion.
#[derive(Clone, Copy, Debug)]
pub struct HtmlImageResource<'a> {
    /// Exact `src` attribute value used as the lookup key.
    pub source: &'a str,
    /// Caller-owned raster image bytes.
    pub bytes: &'a [u8],
    /// Filename used to select the package media extension after byte sniffing.
    pub filename: &'a str,
}

/// The direct story range inserted by one HTML fragment operation.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct HtmlFragmentInsertResult {
    /// Refreshed destination identity valid for the mutated document.
    pub story: StoryId,
    /// Half-open range of inserted direct children in the destination story.
    pub direct_range: Range<usize>,
    /// Ordered location-aware diagnostics for dropped or approximated input.
    pub diagnostics: Vec<HtmlDiagnostic>,
}

#[derive(Clone, Debug, Eq, Hash, PartialEq)]
pub struct MhtmlDiagnostic {
    pub location: String,
    pub property: Option<String>,
    pub message: String,
}

pub struct MhtmlReadResult {
    pub document: Document,
    pub diagnostics: Vec<MhtmlDiagnostic>,
}

pub struct MhtmlWriteResult {
    pub bytes: Vec<u8>,
    pub diagnostics: Vec<MhtmlDiagnostic>,
}

#[derive(Clone, Copy)]
struct MhtmlLimits {
    input_bytes: usize,
    header_bytes: usize,
    parts: usize,
    part_bytes: usize,
    total_decoded_bytes: usize,
    output_bytes: usize,
}

impl Default for MhtmlLimits {
    fn default() -> Self {
        Self {
            input_bytes: MAX_INPUT_BYTES,
            header_bytes: MAX_MHTML_HEADER_BYTES,
            parts: MAX_MHTML_PARTS,
            part_bytes: MAX_INPUT_BYTES,
            total_decoded_bytes: MAX_INPUT_BYTES,
            output_bytes: MAX_MHTML_OUTPUT_BYTES,
        }
    }
}

#[derive(Clone)]
struct MhtmlResource {
    bytes: Vec<u8>,
    content_type: String,
    filename: String,
}

struct HtmlProjection {
    resources: HashMap<String, MhtmlResource>,
    hyperlinks: HashMap<String, String>,
    strict_references: bool,
}

#[derive(Clone, Debug)]
struct EmbeddedHtmlImage {
    rel_id: String,
    drawing_id: Option<u32>,
    description: Option<String>,
    width: Length,
    height: Length,
}

#[derive(Debug, Clone, Copy)]
struct Limits {
    input_bytes: usize,
    retained_text: usize,
    depth: usize,
    nodes: usize,
    blocks: usize,
    runs: usize,
    rows: usize,
    columns: usize,
    cells: usize,
    diagnostics: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            input_bytes: MAX_INPUT_BYTES,
            retained_text: MAX_INPUT_BYTES,
            depth: 256,
            nodes: 100_000,
            blocks: 100_000,
            runs: 100_000,
            rows: 10_000,
            columns: 256,
            cells: 50_000,
            diagnostics: 10_000,
        }
    }
}

#[derive(Debug, Clone, Default)]
struct ComputedStyle {
    font: Option<String>,
    code: bool,
    character_style: Option<String>,
    size: Option<f64>,
    bold: Option<bool>,
    italic: Option<bool>,
    underline: Option<bool>,
    strike: Option<bool>,
    color: Option<String>,
    background: Option<String>,
    mark_highlight: Option<String>,
    alignment: Option<Alignment>,
    space_before: Option<Length>,
    space_after: Option<Length>,
    indent_left: Option<Length>,
    indent_right: Option<Length>,
    first_line_indent: Option<Length>,
    vertical: Option<VerticalText>,
    hyperlink: Option<String>,
}

impl ComputedStyle {
    fn inherited(&self) -> Self {
        Self {
            font: self.font.clone(),
            code: self.code,
            character_style: self.character_style.clone(),
            size: self.size,
            bold: self.bold,
            italic: self.italic,
            underline: self.underline,
            strike: self.strike,
            color: self.color.clone(),
            mark_highlight: self.mark_highlight.clone(),
            alignment: self.alignment,
            vertical: self.vertical,
            hyperlink: self.hyperlink.clone(),
            ..Self::default()
        }
    }
}

#[derive(Debug, Clone, Copy)]
enum VerticalText {
    Superscript,
    Subscript,
}

#[derive(Debug, Clone)]
enum StyleChange {
    Font(String),
    Size(f64),
    Bold(bool),
    Italic(bool),
    Underline(bool),
    Strike(bool),
    Color(String),
    Background(Option<String>),
    Alignment(Alignment),
    SpaceBefore(Length),
    SpaceAfter(Length),
    IndentLeft(Length),
    IndentRight(Length),
    FirstLineIndent(Length),
}

impl StyleChange {
    fn apply(&self, style: &mut ComputedStyle) {
        match self {
            Self::Font(value) => style.font = Some(value.clone()),
            Self::Size(value) => style.size = Some(*value),
            Self::Bold(value) => style.bold = Some(*value),
            Self::Italic(value) => style.italic = Some(*value),
            Self::Underline(value) => style.underline = Some(*value),
            Self::Strike(value) => style.strike = Some(*value),
            Self::Color(value) => style.color = Some(value.clone()),
            Self::Background(value) => style.background = value.clone(),
            Self::Alignment(value) => style.alignment = Some(*value),
            Self::SpaceBefore(value) => style.space_before = Some(*value),
            Self::SpaceAfter(value) => style.space_after = Some(*value),
            Self::IndentLeft(value) => style.indent_left = Some(*value),
            Self::IndentRight(value) => style.indent_right = Some(*value),
            Self::FirstLineIndent(value) => style.first_line_indent = Some(*value),
        }
    }
}

#[derive(Debug)]
struct CssRule {
    selector: Selector,
    specificity: (u32, u32, u32),
    order: usize,
    changes: Vec<StyleChange>,
}

#[derive(Debug, Clone)]
enum InlinePiece {
    Text(String, Box<ComputedStyle>, bool),
    Break,
    Image(EmbeddedHtmlImage),
}

#[derive(Debug)]
struct ParagraphModel {
    pieces: Vec<InlinePiece>,
    style: ComputedStyle,
    paragraph_style: Option<String>,
    numbering: Option<(u32, u32)>,
}

#[derive(Debug)]
struct TableCellModel {
    start: usize,
    span: usize,
    v_merge: Option<VMerge>,
    paragraphs: Vec<ParagraphModel>,
    header: bool,
}

#[derive(Debug)]
struct TableRowModel {
    cells: Vec<TableCellModel>,
    header: bool,
}

#[derive(Debug, Clone, Copy)]
struct ActiveSpan {
    remaining: usize,
    span: usize,
}

fn from_html_with_limits(html: &str, limits: Limits) -> Result<HtmlReadResult> {
    if html.len() > limits.input_bytes {
        return Err(html_error("input", "HTML input exceeds the 64 MiB limit"));
    }
    preflight_markup(html, limits)?;
    let dom = Html::parse_fragment(html);
    validate_dom(&dom, limits)?;
    let projection = fragment_projection(&dom, &[], limits)?;
    from_html_with_resources(html, limits, Some(&projection))
}

fn from_html_with_resources(
    html: &str,
    limits: Limits,
    projection: Option<&HtmlProjection>,
) -> Result<HtmlReadResult> {
    from_html_in_document(html, limits, projection, Document::new())
}

fn from_html_in_document(
    html: &str,
    limits: Limits,
    projection: Option<&HtmlProjection>,
    mut document: Document,
) -> Result<HtmlReadResult> {
    if html.len() > limits.input_bytes {
        return Err(html_error("input", "HTML input exceeds the 64 MiB limit"));
    }
    preflight_markup(html, limits)?;

    let is_document = contains_ascii_case_insensitive(html, b"<html")
        || contains_ascii_case_insensitive(html, b"<!doctype");
    let dom = if is_document {
        Html::parse_document(html)
    } else {
        Html::parse_fragment(html)
    };
    validate_dom(&dom, limits)?;

    let importer = Importer {
        dom: &dom,
        document: &mut document,
        story: None,
        output: Vec::new(),
        container_styles: HashSet::new(),
        quote_owners: HashMap::new(),
        diagnostics: Vec::new(),
        diagnostic_keys: HashSet::new(),
        rules: Vec::new(),
        limits,
        blocks: 0,
        runs: 0,
        rows: 0,
        cells: 0,
        projection,
        embedded_images: HashMap::new(),
    };
    let mut importer = importer;
    importer.record_parser_repairs()?;
    importer.record_head_resources()?;
    importer.collect_styles()?;
    importer.project()?;
    importer.finish()
}

pub(crate) struct HtmlFragmentProjection {
    pub content: Vec<BodyContent>,
    pub diagnostics: Vec<HtmlDiagnostic>,
}

pub(crate) fn project_html_fragment(
    document: &mut Document,
    story: &StoryId,
    html: &str,
    images: &[HtmlImageResource<'_>],
) -> Result<HtmlFragmentProjection> {
    let limits = Limits::default();
    if html.len() > limits.input_bytes {
        return Err(html_error("input", "HTML input exceeds the 64 MiB limit"));
    }
    preflight_markup(html, limits)?;
    let is_document = contains_ascii_case_insensitive(html, b"<html")
        || contains_ascii_case_insensitive(html, b"<!doctype");
    let dom = if is_document {
        Html::parse_document(html)
    } else {
        Html::parse_fragment(html)
    };
    validate_dom(&dom, limits)?;
    let projection = fragment_projection(&dom, images, limits)?;
    let mut importer = Importer {
        dom: &dom,
        document,
        story: Some(story),
        output: Vec::new(),
        container_styles: HashSet::new(),
        quote_owners: HashMap::new(),
        diagnostics: Vec::new(),
        diagnostic_keys: HashSet::new(),
        rules: Vec::new(),
        limits,
        blocks: 0,
        runs: 0,
        rows: 0,
        cells: 0,
        projection: Some(&projection),
        embedded_images: HashMap::new(),
    };
    importer.record_parser_repairs()?;
    importer.record_head_resources()?;
    importer.collect_styles()?;
    importer.project()?;
    Ok(HtmlFragmentProjection {
        content: importer.output,
        diagnostics: importer.diagnostics,
    })
}

fn fragment_projection(
    dom: &Html,
    images: &[HtmlImageResource<'_>],
    limits: Limits,
) -> Result<HtmlProjection> {
    if images.len() > MAX_MHTML_PARTS {
        return Err(html_error("images", "HTML image count exceeds the limit"));
    }
    let mut resources = HashMap::with_capacity(images.len());
    let mut total_bytes = 0_usize;
    for resource in images {
        if resource.source.is_empty() {
            return Err(html_error("images", "image source names must not be empty"));
        }
        if resource.filename.is_empty() {
            return Err(html_error(
                format!("images/{}", resource.source),
                "image filenames must not be empty",
            ));
        }
        let info = oxml_media::probe(resource.bytes).ok_or_else(|| {
            html_error(
                format!("images/{}", resource.source),
                "image bytes are unsupported or malformed",
            )
        })?;
        total_bytes = total_bytes
            .checked_add(resource.bytes.len())
            .ok_or_else(|| html_error("images", "image byte count overflowed"))?;
        if total_bytes > limits.input_bytes {
            return Err(html_error("images", "HTML images exceed the 64 MiB limit"));
        }
        let projected = MhtmlResource {
            bytes: resource.bytes.to_vec(),
            content_type: info.format.content_type().to_owned(),
            filename: resource.filename.to_owned(),
        };
        if resources
            .insert(resource.source.to_owned(), projected)
            .is_some()
        {
            return Err(html_error(
                format!("images/{}", resource.source),
                "duplicate HTML image source",
            ));
        }
    }

    let image_selector = Selector::parse("img[src]").expect("static selector");
    let mut inline_index = 0_usize;
    for image in dom.select(&image_selector) {
        let source = image.attr("src").expect("image source selected");
        if !source
            .get(..5)
            .is_some_and(|prefix| prefix.eq_ignore_ascii_case("data:"))
            || resources.contains_key(source)
        {
            continue;
        }
        let (metadata, encoded) = source[5..].split_once(',').ok_or_else(|| {
            html_error(
                element_path(image),
                "image data URI has no payload separator",
            )
        })?;
        let (content_type, encoding) = metadata.split_once(';').ok_or_else(|| {
            html_error(element_path(image), "image data URI has no base64 encoding")
        })?;
        if !encoding.eq_ignore_ascii_case("base64") {
            return Err(html_error(
                element_path(image),
                "image data URI must use base64 encoding",
            ));
        }
        let bytes = BASE64.decode(encoded).map_err(|_| {
            html_error(
                element_path(image),
                "image data URI has invalid base64 data",
            )
        })?;
        total_bytes = total_bytes
            .checked_add(bytes.len())
            .ok_or_else(|| html_error("images", "image byte count overflowed"))?;
        if total_bytes > limits.input_bytes {
            return Err(html_error("images", "HTML images exceed the 64 MiB limit"));
        }
        let info = oxml_media::probe(&bytes).ok_or_else(|| {
            html_error(
                element_path(image),
                "image data URI bytes are unsupported or malformed",
            )
        })?;
        if !content_type.eq_ignore_ascii_case(info.format.content_type()) {
            return Err(html_error(
                element_path(image),
                "image data URI MIME type does not match its bytes",
            ));
        }
        let filename = format!("inline-{inline_index}.{}", info.format.extension());
        inline_index = inline_index
            .checked_add(1)
            .ok_or_else(|| html_error("images", "inline image count overflowed"))?;
        resources.insert(
            source.to_owned(),
            MhtmlResource {
                bytes,
                content_type: info.format.content_type().to_owned(),
                filename,
            },
        );
    }
    if resources.len() > MAX_MHTML_PARTS {
        return Err(html_error("images", "HTML image count exceeds the limit"));
    }

    let anchor_selector = Selector::parse("a[href]").expect("static selector");
    let hyperlinks = dom
        .select(&anchor_selector)
        .filter_map(|anchor| {
            let href = anchor.attr("href").expect("anchor href selected");
            safe_hyperlink(href, None).map(|target| (href.to_owned(), target))
        })
        .collect();
    Ok(HtmlProjection {
        resources,
        hyperlinks,
        strict_references: false,
    })
}

/// Bound the escaped CommonMark projection before retaining any output bytes.
struct MarkdownHtmlBuffer {
    bytes: Vec<u8>,
    limit: usize,
}

impl std::io::Write for MarkdownHtmlBuffer {
    fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
        if bytes.len() > self.limit.saturating_sub(self.bytes.len()) {
            return Err(std::io::Error::other(
                "Markdown HTML projection exceeds its byte limit",
            ));
        }
        self.bytes.extend_from_slice(bytes);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> std::io::Result<()> {
        Ok(())
    }
}

struct MarkdownHtmlProjection {
    html: String,
    diagnostics: Vec<HtmlDiagnostic>,
    source_elements: Vec<(&'static str, Range<usize>)>,
}

fn project_markdown(markdown: &str, limits: Limits) -> Result<MarkdownHtmlProjection> {
    use pulldown_cmark::{Event, Options, Parser, Tag, TagEnd};

    if markdown.len() > limits.input_bytes {
        return Err(html_error("input", "Markdown input exceeds its byte limit"));
    }
    let mut events = Vec::new();
    let mut diagnostics = Vec::new();
    let mut depth = 0_usize;
    let mut retained_text = 0_usize;
    let mut event_count = 0_usize;
    let mut image_depth = 0_usize;
    let mut source_elements = Vec::new();
    let options = Options::ENABLE_TABLES | Options::ENABLE_STRIKETHROUGH;
    for (event, range) in Parser::new_ext(markdown, options).into_offset_iter() {
        event_count = event_count
            .checked_add(1)
            .ok_or_else(|| html_error("input", "Markdown event count overflowed"))?;
        if event_count > limits.nodes {
            return Err(html_error("input", "Markdown exceeds its event limit"));
        }
        match &event {
            Event::Start(_) => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| html_error("input", "Markdown nesting overflowed"))?;
                if depth > limits.depth {
                    return Err(html_error("input", "Markdown exceeds its nesting limit"));
                }
            }
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| html_error("input", "Markdown event nesting is unbalanced"))?;
            }
            Event::Text(text) | Event::Code(text) | Event::Html(text) | Event::InlineHtml(text) => {
                retained_text = retained_text
                    .checked_add(text.len())
                    .ok_or_else(|| html_error("input", "Markdown retained text overflowed"))?;
                if retained_text > limits.retained_text {
                    return Err(html_error(
                        "input",
                        "Markdown exceeds its retained text limit",
                    ));
                }
            }
            _ => {}
        }
        match &event {
            Event::Start(Tag::Image { .. }) => {
                if image_depth == 0 {
                    source_elements.push(("img", range.clone()));
                }
                image_depth += 1;
            }
            Event::End(TagEnd::Image) => image_depth -= 1,
            Event::Start(Tag::Link { .. }) if image_depth == 0 => {
                source_elements.push(("a", range.clone()));
            }
            _ => {}
        }
        if matches!(event, Event::Html(_) | Event::InlineHtml(_)) {
            if diagnostics.len() >= limits.diagnostics {
                return Err(html_error("input", "Markdown exceeds its diagnostic limit"));
            }
            diagnostics.push(HtmlDiagnostic {
                location: format!("bytes {}..{}", range.start, range.end),
                property: None,
                message: "dropped raw Markdown HTML".to_owned(),
            });
        } else {
            events.push(event);
        }
    }
    if depth != 0 {
        return Err(html_error("input", "Markdown event nesting is unbalanced"));
    }
    let mut output = MarkdownHtmlBuffer {
        bytes: Vec::new(),
        limit: limits.input_bytes,
    };
    pulldown_cmark::html::write_html_io(&mut output, events.into_iter())
        .map_err(|error| html_error("input", error.to_string()))?;
    let html = String::from_utf8(output.bytes)
        .map_err(|_| html_error("input", "Markdown projection is not UTF-8"))?;
    Ok(MarkdownHtmlProjection {
        html,
        diagnostics,
        source_elements,
    })
}

fn html_style_id(document: &Document, id: &str) -> String {
    if document.styles.get_by_id(id).is_some() {
        return id.to_owned();
    }
    let name = if let Some(level) = id.strip_prefix("Heading") {
        format!("Heading {level}")
    } else {
        match id {
            "HTMLPreformatted" => "HTML Preformatted".to_owned(),
            "HTMLQuotedPreformatted" => "HTML Quoted Preformatted".to_owned(),
            "HTMLCode" => "HTML Code".to_owned(),
            _ => id.to_owned(),
        }
    };
    document
        .styles
        .styles
        .iter()
        .find(|style| {
            style
                .name
                .as_deref()
                .is_some_and(|existing| existing.eq_ignore_ascii_case(&name))
        })
        .map_or_else(|| id.to_owned(), |style| style.style_id.clone())
}

fn ensure_markdown_styles(document: &mut Document) -> Result<()> {
    use crate::StyleBuilder;

    for (id, name) in [
        ("Normal", "Normal"),
        ("Quote", "Quote"),
        ("HTMLPreformatted", "HTML Preformatted"),
        ("HTMLQuotedPreformatted", "HTML Quoted Preformatted"),
    ] {
        if document
            .styles
            .get_by_id(&html_style_id(document, id))
            .is_none()
        {
            let mut builder = StyleBuilder::paragraph(id, name);
            if matches!(id, "HTMLPreformatted" | "HTMLQuotedPreformatted") {
                builder = builder.font("Courier New").size(10.0);
            }
            document.add_style(builder)?;
        }
    }
    for level in 1..=6 {
        let id = format!("Heading{level}");
        if document
            .styles
            .get_by_id(&html_style_id(document, &id))
            .is_none()
        {
            let properties = rdocx_oxml::properties::CT_PPr {
                outline_lvl: Some(level - 1),
                ..Default::default()
            };
            document.add_style(
                StyleBuilder::paragraph(&id, &format!("Heading {level}"))
                    .paragraph_properties(properties)
                    .bold(true)
                    .size(24.0 - level as f64 * 2.0),
            )?;
        }
    }
    if document.styles.get_by_id("HTMLThematicBreak").is_none() {
        document.add_style(StyleBuilder::paragraph(
            "HTMLThematicBreak",
            "HTML Thematic Break",
        ))?;
    }
    if document
        .styles
        .get_by_id(&html_style_id(document, "HTMLCode"))
        .is_none()
    {
        document.add_style(StyleBuilder::character("HTMLCode", "HTML Code").font("Courier New"))?;
    }
    Ok(())
}

fn open_local_import(
    path: &Path,
    template: Option<&Document>,
    markdown: bool,
) -> Result<HtmlReadResult> {
    let mut input = File::open(path)?;
    let size = usize::try_from(input.metadata()?.len())
        .map_err(|_| html_error("input", "text input size is not representable"))?;
    let bytes = read_bounded(&mut input, size, MAX_INPUT_BYTES)?;
    let text =
        std::str::from_utf8(&bytes).map_err(|_| html_error("input", "text input must be UTF-8"))?;
    let projection;
    let html = if markdown {
        projection = project_markdown(text, Limits::default())?;
        projection.html.as_str()
    } else {
        preflight_markup(text, Limits::default())?;
        text
    };
    let dom = Html::parse_fragment(html);
    validate_dom(&dom, Limits::default())?;
    let selector = Selector::parse("img[src]").expect("constant image selector");
    let root = path
        .parent()
        .filter(|p| !p.as_os_str().is_empty())
        .unwrap_or(Path::new("."))
        .canonicalize()?;
    let mut images = Vec::new();
    let mut seen = HashSet::new();
    let mut total = 0_usize;
    for element in dom.select(&selector) {
        let source = element.value().attr("src").expect("selected source");
        if source.starts_with("data:") || !seen.insert(source.to_owned()) {
            continue;
        }
        if seen.len() > MAX_MHTML_PARTS {
            return Err(html_error("images", "local image count exceeds the limit"));
        }
        let raw = source.split(['?', '#']).next().unwrap_or(source).as_bytes();
        let mut decoded = Vec::with_capacity(raw.len());
        let mut index = 0;
        while index < raw.len() {
            if raw[index] == b'%' {
                let pair = raw.get(index + 1..index + 3).ok_or_else(|| {
                    html_error("images", "image path has incomplete percent encoding")
                })?;
                let hex = std::str::from_utf8(pair)
                    .map_err(|_| html_error("images", "invalid image path encoding"))?;
                decoded.push(
                    u8::from_str_radix(hex, 16)
                        .map_err(|_| html_error("images", "invalid image path encoding"))?,
                );
                index += 3;
            } else {
                decoded.push(raw[index]);
                index += 1;
            }
        }
        let name = String::from_utf8(decoded)
            .map_err(|_| html_error("images", "image path must be UTF-8"))?;
        let relative = Path::new(&name);
        if name.is_empty()
            || name.contains(':')
            || relative.components().any(|c| {
                !matches!(
                    c,
                    std::path::Component::Normal(_) | std::path::Component::CurDir
                )
            })
        {
            return Err(html_error(
                format!("images/{source}"),
                "only relative image files within the input directory are allowed",
            ));
        }
        let resolved = match root.join(relative).canonicalize() {
            Ok(value) => value,
            Err(error) if error.kind() == std::io::ErrorKind::NotFound => continue,
            Err(error) => return Err(error.into()),
        };
        if !resolved.starts_with(&root) {
            return Err(html_error(
                format!("images/{source}"),
                "image path leaves the input directory",
            ));
        }
        let mut file = File::open(&resolved)?;
        let size = usize::try_from(file.metadata()?.len())
            .map_err(|_| html_error("images", "image size is not representable"))?;
        let remaining = MAX_INPUT_BYTES
            .checked_sub(total)
            .ok_or_else(|| html_error("images", "image byte count overflowed"))?;
        let data = read_bounded(&mut file, size, remaining)?;
        total = total
            .checked_add(data.len())
            .ok_or_else(|| html_error("images", "image byte count overflowed"))?;
        let filename = resolved
            .file_name()
            .and_then(|v| v.to_str())
            .ok_or_else(|| html_error("images", "image filename must be UTF-8"))?
            .to_owned();
        images.push((source.to_owned(), data, filename));
    }
    let resources = images
        .iter()
        .map(|(source, bytes, filename)| HtmlImageResource {
            source,
            bytes,
            filename,
        })
        .collect::<Vec<_>>();
    if markdown {
        Document::from_markdown_with_options(text, template, &resources)
    } else {
        Document::from_html_with_options(text, template, &resources)
    }
}

impl Document {
    /// Read HTML and explicit local image files beneath the input directory.
    pub fn open_html_with_local_images(
        path: impl AsRef<Path>,
        template: Option<&Document>,
    ) -> Result<HtmlReadResult> {
        open_local_import(path.as_ref(), template, false)
    }

    /// Read Markdown and explicit local image files beneath the input directory.
    pub fn open_markdown_with_local_images(
        path: impl AsRef<Path>,
        template: Option<&Document>,
    ) -> Result<HtmlReadResult> {
        open_local_import(path.as_ref(), template, true)
    }

    /// Import bounded CommonMark with tables and strikethrough into native Word content.
    pub fn from_markdown(markdown: &str) -> Result<HtmlReadResult> {
        Self::from_markdown_with_options(markdown, None, &[])
    }

    /// Import CommonMark while preserving the supplied template's non-body package graph.
    pub fn from_markdown_with_options(
        markdown: &str,
        template: Option<&Document>,
        images: &[HtmlImageResource<'_>],
    ) -> Result<HtmlReadResult> {
        let projection = project_markdown(markdown, Limits::default())?;
        let mut result = Self::from_html_with_options(&projection.html, template, images)?;
        let dom = Html::parse_fragment(&projection.html);
        let selector = Selector::parse("a, img").expect("constant Markdown element selector");
        let elements = dom.select(&selector).collect::<Vec<_>>();
        if elements.len() != projection.source_elements.len() {
            return Err(html_error(
                "input",
                "Markdown source element projection is inconsistent",
            ));
        }
        let mut source_locations = HashMap::new();
        for (element, (tag, range)) in elements.into_iter().zip(projection.source_elements) {
            if element.value().name() != tag {
                return Err(html_error(
                    "input",
                    "Markdown source element order is inconsistent",
                ));
            }
            source_locations.insert(element_path(element), range);
        }
        for diagnostic in &mut result.diagnostics {
            if let Some(range) = source_locations.get(&diagnostic.location) {
                diagnostic.location = format!("bytes {}..{}", range.start, range.end);
            }
        }
        let mut diagnostics = projection.diagnostics;
        if diagnostics
            .len()
            .checked_add(result.diagnostics.len())
            .is_none_or(|count| count > Limits::default().diagnostics)
        {
            return Err(html_error(
                "input",
                "Markdown exceeds its combined diagnostic limit",
            ));
        }
        diagnostics.append(&mut result.diagnostics);
        diagnostics.sort_by_key(|diagnostic| {
            diagnostic
                .location
                .strip_prefix("bytes ")
                .and_then(|range| range.split_once(".."))
                .and_then(|(start, _)| start.parse::<usize>().ok())
                .unwrap_or(usize::MAX)
        });
        result.diagnostics = diagnostics;
        Ok(result)
    }

    /// Read bounded UTF-8 Markdown without resolving filesystem or network resources.
    pub fn open_markdown<P: AsRef<Path>>(path: P) -> Result<HtmlReadResult> {
        let path = path.as_ref();
        let expected_size = usize::try_from(std::fs::metadata(path)?.len())
            .map_err(|_| html_error("input", "Markdown input size is not representable"))?;
        let bytes = read_bounded(&mut File::open(path)?, expected_size, MAX_INPUT_BYTES)?;
        let markdown = std::str::from_utf8(&bytes)
            .map_err(|_| html_error("input", "Markdown input is not UTF-8"))?;
        Self::from_markdown(markdown)
    }

    /// Import HTML with explicit images and an optional preserved template.
    pub fn from_html_with_options(
        html: &str,
        template: Option<&Document>,
        images: &[HtmlImageResource<'_>],
    ) -> Result<HtmlReadResult> {
        let limits = Limits::default();
        if html.len() > limits.input_bytes {
            return Err(html_error("input", "HTML input exceeds the 64 MiB limit"));
        }
        preflight_markup(html, limits)?;
        let dom = Html::parse_fragment(html);
        validate_dom(&dom, limits)?;
        let projection = fragment_projection(&dom, images, limits)?;
        let mut document = if let Some(template) = template {
            template.clone_for_staging().prepare_and_reopen_staged()?
        } else {
            Document::new()
        };
        document.document.body.content.clear();
        ensure_markdown_styles(&mut document)?;
        from_html_in_document(html, limits, Some(&projection), document)
    }

    pub fn from_html(html: &str) -> Result<HtmlReadResult> {
        from_html_with_limits(html, Limits::default())
    }

    pub fn open_html<P: AsRef<Path>>(path: P) -> Result<HtmlReadResult> {
        let mut file = File::open(path)?;
        let size = usize::try_from(file.metadata()?.len())
            .map_err(|_| html_error("input", "HTML input size is not representable"))?;
        if size > MAX_INPUT_BYTES {
            return Err(html_error("input", "HTML input exceeds the 64 MiB limit"));
        }
        let bytes = read_bounded(&mut file, size, MAX_INPUT_BYTES)?;
        let html = std::str::from_utf8(&bytes)
            .map_err(|_| html_error("input", "HTML input is not valid UTF-8"))?;
        Self::from_html(html)
    }

    pub fn from_mhtml_bytes(bytes: &[u8]) -> Result<MhtmlReadResult> {
        from_mhtml_with_limits(bytes, MhtmlLimits::default())
    }

    pub fn open_mhtml<P: AsRef<Path>>(path: P) -> Result<MhtmlReadResult> {
        let mut file = File::open(path)?;
        let size = usize::try_from(file.metadata()?.len())
            .map_err(|_| mhtml_error(None, 0, "MHTML input size is not representable"))?;
        if size > MAX_INPUT_BYTES {
            return Err(mhtml_error(None, 0, "MHTML input exceeds the 64 MiB limit"));
        }
        let bytes = read_mhtml_bounded(&mut file, size, MAX_INPUT_BYTES)?;
        Self::from_mhtml_bytes(&bytes)
    }

    pub fn to_mhtml_bytes(&self) -> Result<MhtmlWriteResult> {
        to_mhtml_with_limits(self, MhtmlLimits::default())
    }

    pub fn save_mhtml<P: AsRef<Path>>(&self, path: P) -> Result<Vec<MhtmlDiagnostic>> {
        let result = self.to_mhtml_bytes()?;
        oxml_opc::write_atomic_file(
            path.as_ref(),
            &result.bytes,
            "rdocx",
            "MHTML output path has no file name",
            "could not allocate an MHTML temporary file",
        )?;
        Ok(result.diagnostics)
    }
}

fn from_mhtml_with_limits(bytes: &[u8], limits: MhtmlLimits) -> Result<MhtmlReadResult> {
    let (html, projection) = parse_mhtml(bytes, limits)?;
    let parsed =
        from_html_with_resources(&html, Limits::default(), Some(&projection)).map_err(|error| {
            match error {
                Error::Html { location, message } => mhtml_error(
                    Some("HTML root"),
                    0,
                    format!("HTML projection failed at {location}: {message}"),
                ),
                other => other,
            }
        })?;
    Ok(MhtmlReadResult {
        document: parsed.document,
        diagnostics: parsed
            .diagnostics
            .into_iter()
            .map(|diagnostic| MhtmlDiagnostic {
                location: diagnostic.location,
                property: diagnostic.property,
                message: diagnostic.message,
            })
            .collect(),
    })
}

fn read_mhtml_bounded<R: Read>(
    reader: &mut R,
    expected_size: usize,
    limit: usize,
) -> Result<Vec<u8>> {
    let read_limit = u64::try_from(limit)
        .ok()
        .and_then(|limit| limit.checked_add(1))
        .ok_or_else(|| mhtml_error(None, 0, "MHTML input limit is not representable"))?;
    let mut bytes = Vec::with_capacity(expected_size.min(limit));
    reader.take(read_limit).read_to_end(&mut bytes)?;
    if bytes.len() > limit {
        return Err(mhtml_error(None, 0, "MHTML input exceeds the 64 MiB limit"));
    }
    Ok(bytes)
}

fn base64_lines(bytes: &[u8]) -> String {
    let encoded = BASE64.encode(bytes);
    let mut output = String::with_capacity(encoded.len() + encoded.len() / 76 * 2 + 2);
    for chunk in encoded.as_bytes().chunks(76) {
        output.push_str(std::str::from_utf8(chunk).expect("base64 is ASCII"));
        output.push_str("\r\n");
    }
    output
}

fn mhtml_export_html(document: &Document) -> Result<(String, Vec<MhtmlResource>)> {
    let html = document.to_html();
    // The HTML emitter writes the pictures in the order this visitor reaches
    // them, those of content controls, tracked insertions, smart tags and
    // custom XML included.
    let mut image_sizes = Vec::new();
    visit_accepted_drawings(&document.document.body.content, &mut |drawing| {
        let inline = drawing
            .inline
            .as_ref()
            .map(|image| (&image.embed_id, image.extent_cx.0, image.extent_cy.0));
        let anchor = drawing
            .anchor
            .as_ref()
            .map(|image| (&image.embed_id, image.extent_cx.0, image.extent_cy.0));
        for (embed_id, width_emu, height_emu) in inline.into_iter().chain(anchor) {
            if !embed_id.is_empty() && document.image_data(embed_id).is_some() {
                image_sizes.push((width_emu, height_emu));
            }
        }
    });
    let mut image_size_index = 0_usize;
    let mut output = String::with_capacity(html.len());
    let mut remainder = html.as_str();
    let prefix = "<img src=\"data:";
    let mut resources = Vec::<MhtmlResource>::new();
    let mut by_digest = HashMap::<[u8; 32], usize>::new();
    while let Some(position) = remainder.find(prefix) {
        output.push_str(&remainder[..position]);
        let data_start = position + prefix.len();
        let tail = &remainder[data_start..];
        let quote = tail.find('"').ok_or_else(|| {
            mhtml_error(
                None,
                0,
                "HTML emitter produced an unterminated image source",
            )
        })?;
        let data_uri = &tail[..quote];
        let (content_type, encoded) = data_uri.split_once(";base64,").ok_or_else(|| {
            mhtml_error(None, 0, "HTML emitter produced an unsupported image source")
        })?;
        let bytes = BASE64
            .decode(encoded)
            .map_err(|_| mhtml_error(None, 0, "HTML emitter produced invalid base64 image data"))?;
        let format = oxml_media::ImageFormat::sniff(&bytes)
            .ok_or_else(|| mhtml_error(None, 0, "document image has an unsupported format"))?;
        if !matches!(
            format,
            oxml_media::ImageFormat::Png | oxml_media::ImageFormat::Jpeg
        ) {
            return Err(mhtml_error(
                None,
                0,
                "MHTML export supports only PNG and JPEG images",
            ));
        }
        if format.content_type() != content_type {
            return Err(mhtml_error(
                None,
                0,
                "document image MIME type does not match its bytes",
            ));
        }
        let digest: [u8; 32] = Sha256::digest(&bytes).into();
        let index = if let Some(index) = by_digest.get(&digest) {
            *index
        } else {
            let index = resources.len();
            resources.push(MhtmlResource {
                bytes,
                content_type: content_type.to_owned(),
                filename: format!("image-{index}.{}", format.extension()),
            });
            by_digest.insert(digest, index);
            index
        };
        let (width_emu, height_emu) = *image_sizes.get(image_size_index).ok_or_else(|| {
            mhtml_error(
                None,
                0,
                "HTML image order does not match document image order",
            )
        })?;
        image_size_index += 1;
        output.push_str(&format!(
            "<img src=\"cid:image-{index}@rdocx\" width=\"{:.12}\" height=\"{:.12}\"",
            width_emu as f64 / 9_525.0,
            height_emu as f64 / 9_525.0,
        ));
        remainder = &tail[quote + 1..];
    }
    output.push_str(remainder);
    if image_size_index != image_sizes.len() {
        return Err(mhtml_error(
            None,
            0,
            "document contains images the HTML emitter did not serialize",
        ));
    }
    Ok((output, resources))
}

fn choose_boundary(html: &[u8], resources: &[MhtmlResource]) -> String {
    let mut digest = Sha256::new();
    digest.update(html);
    for resource in resources {
        digest.update(&resource.bytes);
    }
    let hash = digest.finalize();
    for suffix in 0_u32.. {
        let candidate = format!(
            "----=_rdocx_{:02x}{:02x}{:02x}{:02x}_{suffix}",
            hash[0], hash[1], hash[2], hash[3]
        );
        if !html
            .windows(candidate.len())
            .any(|window| window == candidate.as_bytes())
            && resources.iter().all(|resource| {
                !resource
                    .bytes
                    .windows(candidate.len())
                    .any(|window| window == candidate.as_bytes())
            })
        {
            return candidate;
        }
    }
    unreachable!("u32 boundary suffix space cannot be exhausted")
}

fn push_mhtml_loss(
    diagnostics: &mut Vec<MhtmlDiagnostic>,
    location: String,
    message: &'static str,
) -> Result<()> {
    if diagnostics.len() >= Limits::default().diagnostics {
        return Err(mhtml_error(
            None,
            0,
            "MHTML export exceeds the diagnostic limit",
        ));
    }
    diagnostics.push(MhtmlDiagnostic {
        location,
        property: None,
        message: message.to_owned(),
    });
    Ok(())
}

fn paragraph_mhtml_losses(
    document: &Document,
    paragraph: ParagraphRef<'_>,
    location: &str,
    diagnostics: &mut Vec<MhtmlDiagnostic>,
) -> Result<()> {
    for (index, item) in paragraph.items().enumerate() {
        let message = match item {
            ParagraphItemRef::Run(run) => {
                run_mhtml_losses(
                    document,
                    run,
                    &format!("{location}/item[{index}]/run"),
                    diagnostics,
                )?;
                continue;
            }
            ParagraphItemRef::Hyperlink(hyperlink) => {
                hyperlink_mhtml_losses(
                    document,
                    hyperlink,
                    &format!("{location}/item[{index}]/hyperlink"),
                    diagnostics,
                )?;
                continue;
            }
            ParagraphItemRef::Equation(_) => "dropped Word equation",
            ParagraphItemRef::ContentControl(control) => {
                let location = format!("{location}/item[{index}]");
                push_mhtml_loss(
                    diagnostics,
                    location.clone(),
                    "flattened Word paragraph content control",
                )?;
                control_mhtml_losses(document, control.inner, &location, diagnostics)?;
                continue;
            }
            ParagraphItemRef::Revision(revision) => {
                let location = format!("{location}/item[{index}]");
                revision_mhtml_losses(
                    document,
                    revision.inner,
                    &location,
                    "flattened Word revision",
                    "dropped Word revision",
                    diagnostics,
                )?;
                continue;
            }
            ParagraphItemRef::CommentRangeStart { .. } => "dropped Word comment range start",
            ParagraphItemRef::CommentRangeEnd { .. } => "dropped Word comment range end",
            ParagraphItemRef::BookmarkStart { .. } => "dropped Word bookmark start",
            ParagraphItemRef::BookmarkEnd { .. } => "dropped Word bookmark end",
            ParagraphItemRef::UnsupportedXml(raw) => match CT_P::raw_run_wrapper_content(raw) {
                Some(content) => {
                    let location = format!("{location}/item[{index}]");
                    push_mhtml_loss(
                        diagnostics,
                        location.clone(),
                        "flattened Word smart tag or custom XML element",
                    )?;
                    paragraph_mhtml_losses(
                        document,
                        ParagraphRef { inner: &content },
                        &format!("{location}/paragraph"),
                        diagnostics,
                    )?;
                    continue;
                }
                None => "dropped unsupported Word paragraph XML",
            },
        };
        push_mhtml_loss(diagnostics, format!("{location}/item[{index}]"), message)?;
    }
    Ok(())
}

fn run_mhtml_losses(
    document: &Document,
    run: RunRef<'_>,
    location: &str,
    diagnostics: &mut Vec<MhtmlDiagnostic>,
) -> Result<()> {
    for (index, item) in run.items().enumerate() {
        let message = match item {
            RunItemRef::Text(_)
            | RunItemRef::Tab
            | RunItemRef::Break(_)
            | RunItemRef::SpecialCharacter(_) => continue,
            // A producer hint, not content, so nothing is lost by omitting it.
            RunItemRef::LastRenderedPageBreak(_) => continue,
            RunItemRef::Symbol { .. } => "dropped Word symbol character",
            RunItemRef::Drawing(drawing) => match drawing.kind() {
                DrawingKind::Shape => "dropped Word DrawingML shape",
                DrawingKind::Other => "dropped unsupported Word drawing",
                DrawingKind::Image => match drawing.relationship_kind() {
                    Some(DrawingRelationshipKind::Linked) => "dropped linked Word image",
                    Some(DrawingRelationshipKind::Embedded)
                        if drawing
                            .relationship_id()
                            .is_some_and(|id| document.image_data(id).is_some()) =>
                    {
                        continue;
                    }
                    Some(DrawingRelationshipKind::Embedded) | None => {
                        "dropped unresolved Word image"
                    }
                },
            },
            RunItemRef::DeletedText(_) => "dropped Word deleted-text semantics",
            RunItemRef::Field(_) => "dropped Word field semantics",
            RunItemRef::FootnoteReference(_) => "dropped Word footnote reference",
            RunItemRef::EndnoteReference(_) => "dropped Word endnote reference",
            RunItemRef::CommentReference(_) => "dropped Word comment reference",
            RunItemRef::LegacyHorizontalRule(_) => "dropped Word legacy horizontal rule",
            RunItemRef::UnsupportedXml(_) => "dropped unsupported Word run XML",
        };
        push_mhtml_loss(diagnostics, format!("{location}/item[{index}]"), message)?;
    }
    Ok(())
}

fn hyperlink_mhtml_losses(
    document: &Document,
    hyperlink: HyperlinkRef<'_>,
    location: &str,
    diagnostics: &mut Vec<MhtmlDiagnostic>,
) -> Result<()> {
    if hyperlink.anchor().is_some() {
        push_mhtml_loss(
            diagnostics,
            format!("{location}/anchor"),
            "dropped Word internal hyperlink anchor",
        )?;
    }
    if hyperlink.tooltip().is_some() {
        push_mhtml_loss(
            diagnostics,
            format!("{location}/tooltip"),
            "dropped Word hyperlink tooltip",
        )?;
    }
    if hyperlink.doc_location().is_some() {
        push_mhtml_loss(
            diagnostics,
            format!("{location}/doc-location"),
            "dropped Word hyperlink document location",
        )?;
    }
    if hyperlink.has_unmodeled_semantic_attributes() {
        push_mhtml_loss(
            diagnostics,
            format!("{location}/attributes"),
            "dropped unsupported Word hyperlink attributes",
        )?;
    }
    for (index, item) in hyperlink.items().enumerate() {
        match item {
            HyperlinkItemRef::Run(run) => run_mhtml_losses(
                document,
                run,
                &format!("{location}/item[{index}]/run"),
                diagnostics,
            )?,
            HyperlinkItemRef::Revision(revision) => {
                let location = format!("{location}/item[{index}]");
                revision_mhtml_losses(
                    document,
                    revision.inner,
                    &location,
                    "flattened Word hyperlink revision",
                    "dropped Word hyperlink revision",
                    diagnostics,
                )?;
            }
            HyperlinkItemRef::UnsupportedXml(_) => push_mhtml_loss(
                diagnostics,
                format!("{location}/item[{index}]"),
                "dropped unsupported Word hyperlink XML",
            )?,
        }
    }
    Ok(())
}

fn table_mhtml_losses(
    document: &Document,
    table: TableRef<'_>,
    location: &str,
    diagnostics: &mut Vec<MhtmlDiagnostic>,
) -> Result<()> {
    if !table.inner.extra_xml.is_empty() {
        push_mhtml_loss(
            diagnostics,
            format!("{location}/content"),
            "dropped unsupported Word table content",
        )?;
    }
    // Rows that content controls wrap are reported in document order.
    for row_index in 0..=table.row_count() {
        for (control_index, (_, _, control)) in table
            .inner
            .content_controls
            .iter()
            .enumerate()
            .filter(|(_, (at, _, _))| *at == row_index)
        {
            let control_location = format!("{location}/content-control[{control_index}]");
            push_mhtml_loss(
                diagnostics,
                control_location.clone(),
                "flattened Word table row content control",
            )?;
            control_mhtml_losses(document, control, &control_location, diagnostics)?;
        }
        if let Some(row) = table.row(row_index) {
            row_mhtml_losses(
                document,
                row,
                &format!("{location}/row[{row_index}]"),
                diagnostics,
            )?;
        }
    }
    Ok(())
}

fn row_mhtml_losses(
    document: &Document,
    row: RowRef<'_>,
    location: &str,
    diagnostics: &mut Vec<MhtmlDiagnostic>,
) -> Result<()> {
    if row
        .inner
        .extra_xml
        .iter()
        .any(|(position, raw)| !CT_Row::raw_is_root_attributes(*position, raw))
    {
        push_mhtml_loss(
            diagnostics,
            format!("{location}/content"),
            "dropped unsupported Word table row content",
        )?;
    }
    // Cells that content controls wrap are reported in document order.
    for cell_index in 0..=row.cell_count() {
        for (control_index, (_, _, control)) in row
            .inner
            .content_controls
            .iter()
            .enumerate()
            .filter(|(_, (at, _, _))| *at == cell_index)
        {
            let control_location = format!("{location}/content-control[{control_index}]");
            push_mhtml_loss(
                diagnostics,
                control_location.clone(),
                "flattened Word table cell content control",
            )?;
            control_mhtml_losses(document, control, &control_location, diagnostics)?;
        }
        if let Some(cell) = row.cell(cell_index) {
            cell_mhtml_losses(
                document,
                cell,
                &format!("{location}/cell[{cell_index}]"),
                diagnostics,
            )?;
        }
    }
    Ok(())
}

fn cell_mhtml_losses(
    document: &Document,
    cell: CellRef<'_>,
    location: &str,
    diagnostics: &mut Vec<MhtmlDiagnostic>,
) -> Result<()> {
    for (item_index, item) in cell.items().enumerate() {
        match item {
            CellItemRef::Paragraph(paragraph) => paragraph_mhtml_losses(
                document,
                paragraph,
                &format!("{location}/paragraph[{item_index}]"),
                diagnostics,
            )?,
            CellItemRef::Table(table) => table_mhtml_losses(
                document,
                table,
                &format!("{location}/table[{item_index}]"),
                diagnostics,
            )?,
            CellItemRef::ContentControl(control) => {
                let control_location = format!("{location}/item[{item_index}]");
                push_mhtml_loss(
                    diagnostics,
                    control_location.clone(),
                    "flattened Word table cell content control",
                )?;
                control_mhtml_losses(document, control.inner, &control_location, diagnostics)?;
            }
            CellItemRef::UnsupportedXml(_) => push_mhtml_loss(
                diagnostics,
                format!("{location}/item[{item_index}]"),
                "dropped unsupported Word table cell XML",
            )?,
        }
    }
    Ok(())
}

/// Record the losses of what a content control wraps. The HTML emitter
/// writes that content in place, as if the control were not there.
fn control_mhtml_losses(
    document: &Document,
    control: &CT_Sdt,
    location: &str,
    diagnostics: &mut Vec<MhtmlDiagnostic>,
) -> Result<()> {
    for (index, item) in control.content.iter().enumerate() {
        let item_location = format!("{location}/item[{index}]");
        match item {
            SdtContent::Paragraph(paragraph) => paragraph_mhtml_losses(
                document,
                ParagraphRef { inner: paragraph },
                &format!("{item_location}/paragraph"),
                diagnostics,
            )?,
            SdtContent::Table(table) => table_mhtml_losses(
                document,
                TableRef { inner: table },
                &format!("{item_location}/table"),
                diagnostics,
            )?,
            SdtContent::Row(row) => row_mhtml_losses(
                document,
                RowRef { inner: row },
                &format!("{item_location}/row"),
                diagnostics,
            )?,
            SdtContent::Cell(cell) => cell_mhtml_losses(
                document,
                CellRef { inner: cell },
                &format!("{item_location}/cell"),
                diagnostics,
            )?,
            SdtContent::Run(run) => run_mhtml_losses(
                document,
                RunRef { inner: run },
                &format!("{item_location}/run"),
                diagnostics,
            )?,
            SdtContent::ContentControl(nested) => {
                push_mhtml_loss(
                    diagnostics,
                    item_location.clone(),
                    "flattened nested Word content control",
                )?;
                control_mhtml_losses(document, nested, &item_location, diagnostics)?;
            }
            SdtContent::RawXml(_) => push_mhtml_loss(
                diagnostics,
                item_location,
                "dropped unsupported Word content control XML",
            )?,
        }
    }
    for (index, (_, revision)) in control.revisions().iter().enumerate() {
        revision_mhtml_losses(
            document,
            revision,
            &format!("{location}/revision[{index}]"),
            "flattened Word revision",
            "dropped Word revision",
            diagnostics,
        )?;
    }
    Ok(())
}

/// Record a tracked revision: an insertion or a move in is written in place
/// and its losses are recorded, a deletion or a move away is left out.
fn revision_mhtml_losses(
    document: &Document,
    revision: &CT_Revision,
    location: &str,
    flattened: &'static str,
    dropped: &'static str,
    diagnostics: &mut Vec<MhtmlDiagnostic>,
) -> Result<()> {
    if revision_removes_text(revision) {
        return push_mhtml_loss(diagnostics, location.to_owned(), dropped);
    }
    push_mhtml_loss(diagnostics, location.to_owned(), flattened)?;
    if let Some(content) = accepted_revision_content(revision) {
        paragraph_mhtml_losses(
            document,
            ParagraphRef { inner: content },
            &format!("{location}/paragraph"),
            diagnostics,
        )?;
    }
    Ok(())
}

fn mhtml_write_diagnostics(document: &Document) -> Result<Vec<MhtmlDiagnostic>> {
    let mut diagnostics = Vec::new();
    for (index, item) in document.body_items().enumerate() {
        match item {
            BodyItemRef::Paragraph(paragraph) => paragraph_mhtml_losses(
                document,
                paragraph,
                &format!("body[{index}]/paragraph"),
                &mut diagnostics,
            )?,
            BodyItemRef::Table(table) => table_mhtml_losses(
                document,
                table,
                &format!("body[{index}]/table"),
                &mut diagnostics,
            )?,
            BodyItemRef::ContentControl(control) => {
                let location = format!("body[{index}]");
                push_mhtml_loss(
                    &mut diagnostics,
                    location.clone(),
                    "flattened Word body content control",
                )?;
                control_mhtml_losses(document, control.inner, &location, &mut diagnostics)?;
            }
            BodyItemRef::UnsupportedXml(_) => push_mhtml_loss(
                &mut diagnostics,
                format!("body[{index}]"),
                "dropped unsupported Word body XML",
            )?,
        }
    }
    Ok(diagnostics)
}

fn to_mhtml_with_limits(document: &Document, limits: MhtmlLimits) -> Result<MhtmlWriteResult> {
    let diagnostics = mhtml_write_diagnostics(document)?;
    let (html, resources) = mhtml_export_html(document)?;
    let boundary = choose_boundary(html.as_bytes(), &resources);
    let mut output = Vec::new();
    let mut append = |value: &[u8]| -> Result<()> {
        let length = output
            .len()
            .checked_add(value.len())
            .ok_or_else(|| mhtml_error(None, 0, "MHTML output size overflowed"))?;
        if length > limits.output_bytes {
            return Err(mhtml_error(None, 0, "MHTML output exceeds the limit"));
        }
        output.extend_from_slice(value);
        Ok(())
    };
    append(format!(
        "MIME-Version: 1.0\r\nContent-Type: multipart/related; type=\"text/html\"; boundary=\"{boundary}\"; start=\"<document@rdocx>\"\r\n\r\n"
    ).as_bytes())?;
    append(format!(
        "--{boundary}\r\nContent-Type: text/html; charset=utf-8\r\nContent-Transfer-Encoding: base64\r\nContent-ID: <document@rdocx>\r\nContent-Location: https://rdocx.invalid/document.html\r\n\r\n"
    ).as_bytes())?;
    append(base64_lines(html.as_bytes()).as_bytes())?;
    for (index, resource) in resources.iter().enumerate() {
        append(format!(
            "--{boundary}\r\nContent-Type: {}\r\nContent-Transfer-Encoding: base64\r\nContent-ID: <image-{index}@rdocx>\r\nContent-Location: https://rdocx.invalid/{}\r\n\r\n",
            resource.content_type, resource.filename
        ).as_bytes())?;
        append(base64_lines(&resource.bytes).as_bytes())?;
    }
    append(format!("--{boundary}--\r\n").as_bytes())?;
    let MhtmlReadResult {
        mut document,
        diagnostics: _,
    } = from_mhtml_with_limits(&output, limits)?;
    let bytes = document.to_bytes()?;
    Document::from_bytes(&bytes)?;
    Ok(MhtmlWriteResult {
        bytes: output,
        diagnostics,
    })
}

fn read_bounded<R: Read>(reader: &mut R, expected_size: usize, limit: usize) -> Result<Vec<u8>> {
    let read_limit = u64::try_from(limit)
        .ok()
        .and_then(|limit| limit.checked_add(1))
        .ok_or_else(|| html_error("input", "HTML input limit is not representable"))?;
    let mut bytes = Vec::with_capacity(expected_size.min(limit));
    reader.take(read_limit).read_to_end(&mut bytes)?;
    if bytes.len() > limit {
        return Err(html_error("input", "HTML input exceeds the 64 MiB limit"));
    }
    Ok(bytes)
}

fn html_error(location: impl Into<String>, message: impl Into<String>) -> Error {
    Error::Html {
        location: location.into(),
        message: message.into(),
    }
}

fn validate_dom(dom: &Html, limits: Limits) -> Result<()> {
    let mut nodes = 0_usize;
    let mut retained_text = 0_usize;
    for node in dom.tree.nodes() {
        nodes = nodes
            .checked_add(1)
            .ok_or_else(|| html_error("html", "DOM node count overflowed"))?;
        if nodes > limits.nodes {
            return Err(html_error("html", "HTML DOM exceeds the node limit"));
        }
        let mut depth = 0_usize;
        let mut ancestor = Some(node);
        while let Some(current) = ancestor {
            depth = depth
                .checked_add(1)
                .ok_or_else(|| html_error("html", "DOM depth overflowed"))?;
            if depth > limits.depth {
                return Err(html_error("html", "HTML DOM exceeds the depth limit"));
            }
            ancestor = current.parent();
        }
        if let Node::Text(text) = node.value() {
            retained_text = retained_text
                .checked_add(text.len())
                .ok_or_else(|| html_error("html", "retained text length overflowed"))?;
            if retained_text > limits.retained_text {
                return Err(html_error("html", "HTML retained text exceeds the limit"));
            }
        }
    }
    Ok(())
}

fn preflight_markup(html: &str, limits: Limits) -> Result<()> {
    let bytes = html.as_bytes();
    let mut estimated_nodes = 4_usize;
    let mut position = 0_usize;
    while position < bytes.len() {
        let Some(relative_open) = bytes[position..].iter().position(|byte| *byte == b'<') else {
            if bytes[position..]
                .iter()
                .any(|byte| !byte.is_ascii_whitespace())
            {
                estimated_nodes = estimated_nodes
                    .checked_add(1)
                    .ok_or_else(|| html_error("html", "DOM node estimate overflowed"))?;
                if estimated_nodes > limits.nodes {
                    return Err(html_error("html", "HTML DOM exceeds the node limit"));
                }
            }
            break;
        };
        let open = position + relative_open;
        if bytes[position..open]
            .iter()
            .any(|byte| !byte.is_ascii_whitespace())
        {
            estimated_nodes = estimated_nodes
                .checked_add(1)
                .ok_or_else(|| html_error("html", "DOM node estimate overflowed"))?;
        }
        if bytes.get(open + 1).is_some_and(|byte| *byte != b'/') {
            estimated_nodes = estimated_nodes
                .checked_add(1)
                .ok_or_else(|| html_error("html", "DOM node estimate overflowed"))?;
        }
        if estimated_nodes > limits.nodes {
            return Err(html_error("html", "HTML DOM exceeds the node limit"));
        }
        if bytes[open + 1..].starts_with(b"!--") {
            position = bytes[open + 4..]
                .windows(3)
                .position(|window| window == b"-->")
                .map_or(bytes.len(), |relative_close| open + relative_close + 7);
            continue;
        }
        position = bytes[open + 1..]
            .iter()
            .position(|byte| *byte == b'>')
            .map_or(bytes.len(), |relative_close| open + relative_close + 2);
    }
    Ok(())
}

fn contains_ascii_case_insensitive(haystack: &str, needle: &[u8]) -> bool {
    haystack
        .as_bytes()
        .windows(needle.len())
        .any(|window| window.eq_ignore_ascii_case(needle))
}

#[derive(Default)]
struct MimeHeaders {
    fields: Vec<(String, String)>,
}

impl MimeHeaders {
    fn one(&self, name: &str, part: Option<&str>, offset: usize) -> Result<Option<&str>> {
        let mut values = self
            .fields
            .iter()
            .filter(|(field, _)| field == name)
            .map(|(_, value)| value.as_str());
        let first = values.next();
        if values.next().is_some() {
            return Err(mhtml_error(
                part,
                offset,
                format!("duplicate {name} header"),
            ));
        }
        Ok(first)
    }
}

struct MimePart {
    content_type: String,
    charset: Option<String>,
    content_id: Option<String>,
    content_location: Option<String>,
    bytes: Vec<u8>,
    offset: usize,
}

fn mhtml_error(part: Option<&str>, offset: usize, message: impl Into<String>) -> Error {
    Error::Mhtml {
        part: part.map(str::to_owned),
        offset: offset as u64,
        message: message.into(),
    }
}

fn header_end(
    bytes: &[u8],
    limit: usize,
    part: Option<&str>,
    offset: usize,
) -> Result<(usize, usize)> {
    let search = bytes.len().min(limit.saturating_add(4));
    if let Some(position) = bytes[..search]
        .windows(4)
        .position(|value| value == b"\r\n\r\n")
    {
        return Ok((position, 4));
    }
    if let Some(position) = bytes[..search]
        .windows(2)
        .position(|value| value == b"\n\n")
    {
        return Ok((position, 2));
    }
    Err(mhtml_error(part, offset, "missing MIME header terminator"))
}

fn parse_headers(
    bytes: &[u8],
    limit: usize,
    part: Option<&str>,
    offset: usize,
) -> Result<(MimeHeaders, usize)> {
    let (end, separator) = header_end(bytes, limit, part, offset)?;
    if end > limit {
        return Err(mhtml_error(part, offset, "MIME headers exceed the limit"));
    }
    let text = std::str::from_utf8(&bytes[..end])
        .map_err(|_| mhtml_error(part, offset, "MIME headers are not valid UTF-8"))?;
    if !text.is_ascii() {
        return Err(mhtml_error(part, offset, "MIME headers are not ASCII"));
    }
    let mut unfolded: Vec<String> = Vec::new();
    for line in text.replace("\r\n", "\n").split('\n') {
        if line.starts_with([' ', '\t']) {
            let previous = unfolded
                .last_mut()
                .ok_or_else(|| mhtml_error(part, offset, "orphan folded MIME header"))?;
            previous.push(' ');
            previous.push_str(line.trim());
        } else if !line.is_empty() {
            unfolded.push(line.to_owned());
        }
    }
    let mut headers = MimeHeaders::default();
    for line in unfolded {
        let (name, value) = line
            .split_once(':')
            .ok_or_else(|| mhtml_error(part, offset, "malformed MIME header"))?;
        if name.is_empty()
            || !name
                .bytes()
                .all(|byte| byte.is_ascii_alphanumeric() || byte == b'-')
        {
            return Err(mhtml_error(part, offset, "invalid MIME header name"));
        }
        if value
            .bytes()
            .any(|byte| byte == 0 || byte == b'\r' || byte == b'\n')
        {
            return Err(mhtml_error(part, offset, "invalid MIME header value"));
        }
        headers
            .fields
            .push((name.to_ascii_lowercase(), value.trim().to_owned()));
    }
    Ok((headers, end + separator))
}

fn parse_content_type(value: &str) -> Result<(String, HashMap<String, String>)> {
    let mut sections = Vec::new();
    let mut start = 0_usize;
    let mut quoted = false;
    let mut escaped = false;
    for (index, character) in value.char_indices() {
        if escaped {
            escaped = false;
        } else if quoted && character == '\\' {
            escaped = true;
        } else if character == '"' {
            quoted = !quoted;
        } else if character == ';' && !quoted {
            sections.push(&value[start..index]);
            start = index + 1;
        }
    }
    if quoted || escaped {
        return Err(mhtml_error(None, 0, "unterminated MIME parameter"));
    }
    sections.push(&value[start..]);
    let media_type = sections
        .first()
        .copied()
        .unwrap_or_default()
        .trim()
        .to_ascii_lowercase();
    if !media_type.contains('/') {
        return Err(mhtml_error(None, 0, "invalid MIME content type"));
    }
    let mut parameters = HashMap::new();
    for section in sections.into_iter().skip(1) {
        let Some((name, raw_value)) = section.split_once('=') else {
            return Err(mhtml_error(None, 0, "invalid MIME content type parameter"));
        };
        let name = name.trim().to_ascii_lowercase();
        let raw_value = raw_value.trim();
        let parameter = if raw_value.starts_with('"') {
            if raw_value.len() < 2 || !raw_value.ends_with('"') {
                return Err(mhtml_error(None, 0, "unterminated MIME parameter"));
            }
            let mut output = String::new();
            let mut escaped = false;
            for character in raw_value[1..raw_value.len() - 1].chars() {
                if escaped {
                    output.push(character);
                    escaped = false;
                } else if character == '\\' {
                    escaped = true;
                } else {
                    output.push(character);
                }
            }
            if escaped {
                return Err(mhtml_error(None, 0, "invalid MIME parameter escape"));
            }
            output
        } else {
            raw_value.to_owned()
        };
        if name.is_empty() || parameter.is_empty() || parameters.insert(name, parameter).is_some() {
            return Err(mhtml_error(None, 0, "invalid or duplicate MIME parameter"));
        }
    }
    Ok((media_type, parameters))
}

fn trim_part_body(mut bytes: &[u8]) -> &[u8] {
    if bytes.ends_with(b"\r\n") {
        bytes = &bytes[..bytes.len() - 2];
    } else if bytes.ends_with(b"\n") {
        bytes = &bytes[..bytes.len() - 1];
    }
    bytes
}

fn multipart_segments<'a>(
    body: &'a [u8],
    boundary: &str,
    offset: usize,
) -> Result<Vec<(usize, &'a [u8])>> {
    if boundary.is_empty()
        || boundary.len() > 70
        || boundary.bytes().any(|byte| byte <= b' ' || byte >= 0x7f)
    {
        return Err(mhtml_error(None, offset, "invalid multipart boundary"));
    }
    let marker = format!("--{boundary}").into_bytes();
    let mut boundaries = Vec::new();
    let mut position = 0;
    while position <= body.len() {
        let line_end = body[position..]
            .iter()
            .position(|byte| *byte == b'\n')
            .map_or(body.len(), |length| position + length + 1);
        let mut line = &body[position..line_end];
        if line.ends_with(b"\n") {
            line = &line[..line.len() - 1];
        }
        if line.ends_with(b"\r") {
            line = &line[..line.len() - 1];
        }
        let line = line.trim_ascii_end();
        let closing = line == [marker.as_slice(), b"--"].concat();
        if line == marker || closing {
            boundaries.push((position, line_end, closing));
        }
        if line_end == body.len() {
            break;
        }
        position = line_end;
    }
    let first = boundaries
        .first()
        .ok_or_else(|| mhtml_error(None, offset, "multipart boundary was not found"))?;
    if body[..first.0]
        .iter()
        .any(|byte| !byte.is_ascii_whitespace())
    {
        return Err(mhtml_error(
            None,
            offset,
            "non-whitespace multipart preamble",
        ));
    }
    let closing: Vec<_> = boundaries
        .iter()
        .enumerate()
        .filter(|(_, entry)| entry.2)
        .collect();
    if closing.len() != 1 || closing[0].0 + 1 != boundaries.len() {
        return Err(mhtml_error(
            None,
            offset,
            "multipart must contain one final closing boundary",
        ));
    }
    let closing_end = boundaries.last().map(|entry| entry.1).unwrap_or(body.len());
    if body[closing_end..]
        .iter()
        .any(|byte| !byte.is_ascii_whitespace())
    {
        return Err(mhtml_error(
            None,
            offset + closing_end,
            "non-whitespace multipart epilogue",
        ));
    }
    let mut result = Vec::new();
    for pair in boundaries.windows(2) {
        let start = pair[0].1;
        if pair[0].2 {
            break;
        }
        result.push((offset + start, trim_part_body(&body[start..pair[1].0])));
    }
    Ok(result)
}

fn decode_quoted_printable(bytes: &[u8], part: Option<&str>, offset: usize) -> Result<Vec<u8>> {
    let mut output = Vec::with_capacity(bytes.len());
    let mut position = 0;
    while position < bytes.len() {
        if bytes[position] != b'=' {
            output.push(bytes[position]);
            position += 1;
            continue;
        }
        if bytes.get(position + 1..position + 3) == Some(b"\r\n") {
            position += 3;
            continue;
        }
        if bytes.get(position + 1) == Some(&b'\n') {
            position += 2;
            continue;
        }
        let hex = bytes.get(position + 1..position + 3).ok_or_else(|| {
            mhtml_error(part, offset + position, "truncated quoted-printable escape")
        })?;
        let digit = |byte: u8| match byte {
            b'0'..=b'9' => Some(byte - b'0'),
            b'a'..=b'f' => Some(byte - b'a' + 10),
            b'A'..=b'F' => Some(byte - b'A' + 10),
            _ => None,
        };
        let high = digit(hex[0]).ok_or_else(|| {
            mhtml_error(part, offset + position, "invalid quoted-printable escape")
        })?;
        let low = digit(hex[1]).ok_or_else(|| {
            mhtml_error(part, offset + position, "invalid quoted-printable escape")
        })?;
        output.push((high << 4) | low);
        position += 3;
    }
    Ok(output)
}

fn decode_transfer(
    encoding: Option<&str>,
    bytes: &[u8],
    part: Option<&str>,
    offset: usize,
) -> Result<Vec<u8>> {
    match encoding
        .unwrap_or("7bit")
        .trim()
        .to_ascii_lowercase()
        .as_str()
    {
        "7bit" => {
            if bytes.iter().any(|byte| *byte > 0x7f) {
                return Err(mhtml_error(
                    part,
                    offset,
                    "7bit MIME part contains non-ASCII bytes",
                ));
            }
            Ok(bytes.to_vec())
        }
        "8bit" => Ok(bytes.to_vec()),
        "base64" => {
            let compact: Vec<_> = bytes
                .iter()
                .copied()
                .filter(|byte| !byte.is_ascii_whitespace())
                .collect();
            BASE64
                .decode(compact)
                .map_err(|_| mhtml_error(part, offset, "invalid base64 MIME body"))
        }
        "quoted-printable" => decode_quoted_printable(bytes, part, offset),
        other => Err(mhtml_error(
            part,
            offset,
            format!("unsupported content-transfer-encoding `{other}`"),
        )),
    }
}

fn normalize_content_id(value: &str) -> Result<String> {
    let value = value.trim();
    let value = value
        .strip_prefix('<')
        .and_then(|value| value.strip_suffix('>'))
        .unwrap_or(value);
    if value.is_empty()
        || value.bytes().any(|byte| {
            byte.is_ascii_whitespace() || byte.is_ascii_control() || matches!(byte, b'<' | b'>')
        })
    {
        return Err(mhtml_error(None, 0, "invalid Content-ID"));
    }
    Ok(value.to_owned())
}

fn normalized_absolute_location(value: &str) -> Option<String> {
    let value = value.trim();
    if value
        .bytes()
        .any(|byte| byte.is_ascii_control() || byte.is_ascii_whitespace() || byte == b'\\')
    {
        return None;
    }
    let (scheme, remainder) = value.split_once("://")?;
    if !matches!(scheme.to_ascii_lowercase().as_str(), "http" | "https") {
        return None;
    }
    let end = remainder.find(['/', '?', '#']).unwrap_or(remainder.len());
    let authority = &remainder[..end];
    if authority.is_empty() || authority.contains('@') {
        return None;
    }
    let tail = &remainder[end..];
    let tail = tail.split('#').next().unwrap_or_default();
    let (path, query) = tail
        .split_once('?')
        .map_or((tail, None), |(path, query)| (path, Some(query)));
    let path = normalize_path(if path.is_empty() { "/" } else { path })?;
    let query = query.map_or(String::new(), |query| format!("?{query}"));
    Some(format!(
        "{}://{}{}{}",
        scheme.to_ascii_lowercase(),
        authority.to_ascii_lowercase(),
        path,
        query
    ))
}

fn normalize_path(path: &str) -> Option<String> {
    let absolute = path.starts_with('/');
    let mut segments = Vec::new();
    for segment in path.split('/') {
        match segment {
            "" | "." => {}
            ".." => {
                segments.pop()?;
            }
            _ if segment
                .bytes()
                .any(|byte| byte.is_ascii_control() || byte == b'\\') =>
            {
                return None;
            }
            _ => segments.push(segment),
        }
    }
    let mut output = if absolute {
        "/".to_owned()
    } else {
        String::new()
    };
    output.push_str(&segments.join("/"));
    if path.ends_with('/') && !output.ends_with('/') {
        output.push('/');
    }
    Some(output)
}

fn resolve_location(base: Option<&str>, value: &str) -> Option<String> {
    let value = value.trim();
    if let Some(absolute) = normalized_absolute_location(value) {
        return Some(absolute);
    }
    if value.starts_with("//") || value.starts_with('/') || value.contains(':') || value.is_empty()
    {
        return None;
    }
    let base = base?;
    let base = normalized_absolute_location(base)?;
    let scheme_end = base.find("://")? + 3;
    let path_start = base[scheme_end..]
        .find('/')
        .map_or(base.len(), |index| scheme_end + index);
    let origin = &base[..path_start];
    let base_path = base[path_start..].split('?').next().unwrap_or("/");
    let directory = base_path
        .rsplit_once('/')
        .map_or("/", |(directory, _)| directory);
    let joined = format!("{directory}/{value}");
    Some(format!("{origin}{}", normalize_path(&joined)?))
}

fn resolve_resource_reference(
    reference: &str,
    root_location: Option<&str>,
    ids: &HashMap<String, usize>,
    locations: &HashMap<String, usize>,
) -> Result<Option<usize>> {
    if reference.to_ascii_lowercase().starts_with("cid:") {
        let id = normalize_content_id(&reference[4..])?;
        Ok(ids.get(&id).copied())
    } else {
        Ok(resolve_location(root_location, reference)
            .and_then(|location| locations.get(&location).copied()))
    }
}

fn decode_css_escapes(css: &str) -> Result<String> {
    let mut decoded = String::with_capacity(css.len());
    let mut characters = css.chars().peekable();
    while let Some(character) = characters.next() {
        if character != '\\' {
            decoded.push(character);
            continue;
        }

        let Some(&next) = characters.peek() else {
            return Err(mhtml_error(None, 0, "unterminated CSS escape"));
        };
        if next.is_ascii_hexdigit() {
            let mut value = 0_u32;
            for _ in 0..6 {
                let Some(&digit) = characters.peek() else {
                    break;
                };
                let Some(nibble) = digit.to_digit(16) else {
                    break;
                };
                characters.next();
                value = value * 16 + nibble;
            }
            if characters.peek().is_some_and(|next| next.is_whitespace()) {
                let whitespace = characters.next();
                if whitespace == Some('\r') && characters.peek() == Some(&'\n') {
                    characters.next();
                }
            }
            decoded.push(
                char::from_u32(value)
                    .filter(|value| *value != '\0')
                    .unwrap_or('\u{fffd}'),
            );
        } else if matches!(next, '\n' | '\r' | '\u{000c}') {
            characters.next();
            if next == '\r' && characters.peek() == Some(&'\n') {
                characters.next();
            }
        } else {
            decoded.push(characters.next().expect("peeked CSS escape"));
        }
    }
    Ok(decoded)
}

fn css_url_reference(css: &str, mut position: usize) -> Result<(String, usize)> {
    let bytes = css.as_bytes();
    while bytes.get(position).is_some_and(u8::is_ascii_whitespace) {
        position += 1;
    }
    let (reference, end) = match bytes.get(position) {
        Some(&quote @ (b'\'' | b'"')) => {
            let start = position + 1;
            let end = css[start..]
                .find(char::from(quote))
                .map(|relative| start + relative)
                .ok_or_else(|| mhtml_error(None, 0, "unterminated quoted CSS resource URL"))?;
            position = end + 1;
            while bytes.get(position).is_some_and(u8::is_ascii_whitespace) {
                position += 1;
            }
            if bytes.get(position) != Some(&b')') {
                return Err(mhtml_error(
                    None,
                    0,
                    "quoted CSS resource URL has trailing syntax",
                ));
            }
            (&css[start..end], position + 1)
        }
        _ => {
            let end = css[position..]
                .find(')')
                .map(|relative| position + relative)
                .ok_or_else(|| mhtml_error(None, 0, "unterminated CSS resource URL"))?;
            let reference = css[position..end].trim();
            if reference.contains(['\'', '"']) {
                return Err(mhtml_error(None, 0, "invalid unquoted CSS resource URL"));
            }
            (reference, end + 1)
        }
    };
    let reference = reference.trim();
    if reference.is_empty() {
        return Err(mhtml_error(None, 0, "empty CSS resource URL"));
    }
    Ok((reference.to_owned(), end))
}

fn css_resource_references(css: &str) -> Result<Vec<String>> {
    if css.contains('\\') {
        let decoded = decode_css_escapes(css)?;
        let decoded_lower = decoded.to_ascii_lowercase();
        if decoded_lower.contains("url(") || decoded_lower.contains("@import") {
            return Err(mhtml_error(None, 0, "escaped CSS resource syntax"));
        }
    }
    let lower = css.to_ascii_lowercase();
    let mut references = Vec::new();
    let mut position = 0_usize;
    while let Some(relative) = lower[position..].find("url(") {
        let start = position + relative + 4;
        let (reference, end) = css_url_reference(css, start)?;
        references.push(reference);
        position = end;
    }

    position = 0;
    while let Some(relative) = lower[position..].find("@import") {
        let mut start = position + relative + "@import".len();
        while css
            .as_bytes()
            .get(start)
            .is_some_and(u8::is_ascii_whitespace)
        {
            start += 1;
        }
        if lower[start..].starts_with("url(") {
            position = start + 4;
            continue;
        }
        let Some(&quote @ (b'\'' | b'"')) = css.as_bytes().get(start) else {
            return Err(mhtml_error(None, 0, "unsupported CSS import resource"));
        };
        start += 1;
        let mut end = start;
        while let Some(&byte) = css.as_bytes().get(end) {
            if byte == b'\\' {
                return Err(mhtml_error(None, 0, "escaped CSS import resource"));
            }
            if byte == quote {
                break;
            }
            end += 1;
        }
        if css.as_bytes().get(end) != Some(&quote) {
            return Err(mhtml_error(None, 0, "unterminated CSS import resource"));
        }
        let reference = css[start..end].trim();
        if reference.is_empty() {
            return Err(mhtml_error(None, 0, "empty CSS import resource"));
        }
        references.push(reference.to_owned());
        position = end + 1;
    }
    Ok(references)
}

fn safe_hyperlink(value: &str, base: Option<&str>) -> Option<String> {
    let value = value.trim();
    if value.is_empty()
        || value
            .bytes()
            .any(|byte| byte.is_ascii_control() || byte.is_ascii_whitespace())
    {
        return None;
    }
    if value.starts_with('#') {
        return Some(value.to_owned());
    }
    if value.to_ascii_lowercase().starts_with("mailto:") {
        return Some(value.to_owned());
    }
    normalized_absolute_location(value).or_else(|| resolve_location(base, value))
}

fn image_dimension(element: ElementRef<'_>, name: &str) -> Result<Option<Length>> {
    let Some(value) = element.attr(name) else {
        return Ok(None);
    };
    let pixels = value
        .trim()
        .trim_end_matches("px")
        .trim()
        .parse::<f64>()
        .map_err(|_| mhtml_error(None, 0, format!("invalid MHTML image {name} `{value}`")))?;
    if !pixels.is_finite() || pixels <= 0.0 || pixels > 1_000_000.0 {
        return Err(mhtml_error(
            None,
            0,
            format!("MHTML image {name} is outside the supported range"),
        ));
    }
    let emu = pixels * 9_525.0;
    if !emu.is_finite() || emu >= i64::MAX as f64 {
        return Err(mhtml_error(None, 0, "MHTML image dimension overflowed"));
    }
    Ok(Some(Length::emu(emu as i64)))
}

fn parse_mhtml(bytes: &[u8], limits: MhtmlLimits) -> Result<(String, HtmlProjection)> {
    if bytes.len() > limits.input_bytes {
        return Err(mhtml_error(None, 0, "MHTML input exceeds the 64 MiB limit"));
    }
    let (headers, body_start) = parse_headers(bytes, limits.header_bytes, None, 0)?;
    let mime_version = headers
        .one("mime-version", None, 0)?
        .ok_or_else(|| mhtml_error(None, 0, "MIME-Version header is required"))?;
    if mime_version != "1.0" {
        return Err(mhtml_error(None, 0, "unsupported MIME-Version"));
    }
    let content_type = headers
        .one("content-type", None, 0)?
        .ok_or_else(|| mhtml_error(None, 0, "missing MHTML Content-Type header"))?;
    let (media_type, parameters) = parse_content_type(content_type)?;
    if media_type != "multipart/related" {
        return Err(mhtml_error(None, 0, "MHTML root is not multipart/related"));
    }
    if parameters
        .get("type")
        .is_some_and(|value| !value.eq_ignore_ascii_case("text/html"))
    {
        return Err(mhtml_error(
            None,
            0,
            "multipart/related type does not declare text/html",
        ));
    }
    let boundary = parameters
        .get("boundary")
        .ok_or_else(|| mhtml_error(None, 0, "multipart/related boundary is required"))?;
    let segments = multipart_segments(&bytes[body_start..], boundary, body_start)?;
    if segments.is_empty() || segments.len() > limits.parts {
        return Err(mhtml_error(
            None,
            body_start,
            "MHTML part count is outside the supported limit",
        ));
    }
    let mut parts = Vec::with_capacity(segments.len());
    let mut total_decoded = 0_usize;
    for (index, (part_offset, segment)) in segments.into_iter().enumerate() {
        let identity = format!("part[{index}]");
        let (part_headers, content_start) =
            parse_headers(segment, limits.header_bytes, Some(&identity), part_offset)?;
        let content_type = part_headers
            .one("content-type", Some(&identity), part_offset)?
            .ok_or_else(|| {
                mhtml_error(
                    Some(&identity),
                    part_offset,
                    "part Content-Type is required",
                )
            })?;
        let (content_type, content_parameters) = parse_content_type(content_type)?;
        let content_id = part_headers
            .one("content-id", Some(&identity), part_offset)?
            .map(normalize_content_id)
            .transpose()?;
        let content_location = part_headers
            .one("content-location", Some(&identity), part_offset)?
            .map(str::trim)
            .map(str::to_owned);
        let transfer =
            part_headers.one("content-transfer-encoding", Some(&identity), part_offset)?;
        let decoded = decode_transfer(
            transfer,
            &segment[content_start..],
            content_id.as_deref().or(Some(&identity)),
            part_offset + content_start,
        )?;
        if decoded.len() > limits.part_bytes {
            return Err(mhtml_error(
                content_id.as_deref(),
                part_offset,
                "decoded MIME part exceeds the limit",
            ));
        }
        total_decoded = total_decoded
            .checked_add(decoded.len())
            .ok_or_else(|| mhtml_error(None, part_offset, "decoded MIME size overflowed"))?;
        if total_decoded > limits.total_decoded_bytes {
            return Err(mhtml_error(
                None,
                part_offset,
                "decoded MHTML content exceeds the total limit",
            ));
        }
        parts.push(MimePart {
            content_type,
            charset: content_parameters.get("charset").cloned(),
            content_id,
            content_location,
            bytes: decoded,
            offset: part_offset,
        });
    }

    let root_index = if let Some(start) = parameters.get("start") {
        let start = normalize_content_id(start)?;
        let matches: Vec<_> = parts
            .iter()
            .enumerate()
            .filter(|(_, part)| part.content_id.as_deref() == Some(start.as_str()))
            .map(|(index, _)| index)
            .collect();
        if matches.len() != 1 {
            return Err(mhtml_error(
                Some(&start),
                0,
                "MHTML start does not select exactly one part",
            ));
        }
        matches[0]
    } else {
        let roots: Vec<_> = parts
            .iter()
            .enumerate()
            .filter(|(_, part)| part.content_type == "text/html")
            .map(|(index, _)| index)
            .collect();
        if roots.len() != 1 {
            return Err(mhtml_error(
                None,
                0,
                "MHTML must contain exactly one HTML root",
            ));
        }
        roots[0]
    };
    if parts[root_index].content_type != "text/html" {
        return Err(mhtml_error(
            parts[root_index].content_id.as_deref(),
            parts[root_index].offset,
            "selected MHTML root is not text/html",
        ));
    }
    if parts[root_index]
        .charset
        .as_deref()
        .is_some_and(|charset| !charset.eq_ignore_ascii_case("utf-8"))
    {
        return Err(mhtml_error(
            parts[root_index].content_id.as_deref(),
            parts[root_index].offset,
            "HTML root charset is not UTF-8",
        ));
    }
    let root_location = match parts[root_index].content_location.as_deref() {
        Some(location) => Some(normalized_absolute_location(location).ok_or_else(|| {
            mhtml_error(
                None,
                parts[root_index].offset,
                "invalid HTML root Content-Location",
            )
        })?),
        None => None,
    };

    let mut ids = HashMap::new();
    let mut locations = HashMap::new();
    for (index, part) in parts.iter().enumerate() {
        if let Some(id) = &part.content_id
            && ids.insert(id.clone(), index).is_some()
        {
            return Err(mhtml_error(Some(id), part.offset, "duplicate Content-ID"));
        }
        if let Some(location) = &part.content_location {
            let normalized = normalized_absolute_location(location)
                .or_else(|| resolve_location(root_location.as_deref(), location))
                .ok_or_else(|| {
                    mhtml_error(
                        part.content_id.as_deref(),
                        part.offset,
                        "unsafe Content-Location",
                    )
                })?;
            if locations.insert(normalized, index).is_some() {
                return Err(mhtml_error(
                    part.content_id.as_deref(),
                    part.offset,
                    "duplicate normalized Content-Location",
                ));
            }
        }
    }
    let html = std::str::from_utf8(&parts[root_index].bytes)
        .map_err(|_| {
            mhtml_error(
                parts[root_index].content_id.as_deref(),
                parts[root_index].offset,
                "HTML root is not UTF-8",
            )
        })?
        .to_owned();
    let dom = Html::parse_document(&html);
    let mut resources = HashMap::new();
    let resource_selector = Selector::parse(
        "img[src], link[href], script[src], iframe[src], frame[src], embed[src], object[data], video[src], audio[src], track[src]",
    )
    .expect("static selector");
    for element in dom.select(&resource_selector) {
        let attribute = if element.value().name() == "object" {
            "data"
        } else if element.value().name() == "link" {
            "href"
        } else {
            "src"
        };
        let reference = element.attr(attribute).unwrap_or_default().trim();
        let resource_index =
            resolve_resource_reference(reference, root_location.as_deref(), &ids, &locations)?
                .ok_or_else(|| {
                    mhtml_error(
                        None,
                        0,
                        format!("unresolved or external subresource `{reference}`"),
                    )
                })?;
        if resource_index == root_index {
            return Err(mhtml_error(
                None,
                0,
                "HTML root cannot be used as a subresource",
            ));
        }
        if element.value().name() == "img" {
            let part = &parts[resource_index];
            let format = oxml_media::ImageFormat::sniff(&part.bytes).ok_or_else(|| {
                mhtml_error(
                    part.content_id.as_deref(),
                    part.offset,
                    "image bytes have an unsupported format",
                )
            })?;
            if !matches!(
                format,
                oxml_media::ImageFormat::Png | oxml_media::ImageFormat::Jpeg
            ) {
                return Err(mhtml_error(
                    part.content_id.as_deref(),
                    part.offset,
                    "MHTML import supports only PNG and JPEG images",
                ));
            }
            if format.content_type() != part.content_type {
                return Err(mhtml_error(
                    part.content_id.as_deref(),
                    part.offset,
                    "declared image MIME type does not match its bytes",
                ));
            }
            resources.insert(
                reference.to_owned(),
                MhtmlResource {
                    bytes: part.bytes.clone(),
                    content_type: part.content_type.clone(),
                    filename: format!("resource.{}", format.extension()),
                },
            );
        }
    }
    let responsive_selector =
        Selector::parse("img[srcset], source[src], source[srcset], video[poster], input[src]")
            .expect("static selector");
    for element in dom.select(&responsive_selector) {
        for attribute in ["src", "srcset", "poster"] {
            let Some(value) = element.attr(attribute) else {
                continue;
            };
            let candidates: Vec<_> = if attribute == "srcset" {
                value
                    .split(',')
                    .filter_map(|candidate| candidate.split_whitespace().next())
                    .collect()
            } else {
                vec![value.trim()]
            };
            for reference in candidates {
                if resolve_resource_reference(
                    reference,
                    root_location.as_deref(),
                    &ids,
                    &locations,
                )?
                .is_none()
                {
                    return Err(mhtml_error(
                        None,
                        0,
                        format!("unresolved or external subresource `{reference}`"),
                    ));
                }
            }
        }
    }
    let background_selector = Selector::parse("[background]").expect("static selector");
    for element in dom.select(&background_selector) {
        let reference = element.attr("background").unwrap_or_default().trim();
        if resolve_resource_reference(reference, root_location.as_deref(), &ids, &locations)?
            .is_none()
        {
            return Err(mhtml_error(
                None,
                0,
                format!("unresolved or external subresource `{reference}`"),
            ));
        }
    }
    let styled_selector = Selector::parse("[style], style").expect("static selector");
    for element in dom.select(&styled_selector) {
        let css = if element.value().name() == "style" {
            element.text().collect::<String>()
        } else {
            element.attr("style").unwrap_or_default().to_owned()
        };
        for reference in css_resource_references(&css)? {
            if resolve_resource_reference(&reference, root_location.as_deref(), &ids, &locations)?
                .is_none()
            {
                return Err(mhtml_error(
                    None,
                    0,
                    format!("unresolved or external subresource `{reference}`"),
                ));
            }
        }
    }
    let anchor_selector = Selector::parse("a[href]").expect("static selector");
    let mut hyperlinks = HashMap::new();
    for anchor in dom.select(&anchor_selector) {
        let href = anchor.attr("href").unwrap_or_default();
        let normalized = safe_hyperlink(href, root_location.as_deref())
            .ok_or_else(|| mhtml_error(None, 0, format!("unsafe hyperlink target `{href}`")))?;
        hyperlinks.insert(href.to_owned(), normalized);
    }
    Ok((
        html,
        HtmlProjection {
            resources,
            hyperlinks,
            strict_references: true,
        },
    ))
}

struct Importer<'a> {
    dom: &'a Html,
    document: &'a mut Document,
    story: Option<&'a StoryId>,
    output: Vec<BodyContent>,
    container_styles: HashSet<String>,
    quote_owners: HashMap<String, u32>,
    diagnostics: Vec<HtmlDiagnostic>,
    diagnostic_keys: HashSet<HtmlDiagnostic>,
    rules: Vec<CssRule>,
    limits: Limits,
    blocks: usize,
    runs: usize,
    rows: usize,
    cells: usize,
    projection: Option<&'a HtmlProjection>,
    embedded_images: HashMap<String, String>,
}

impl Importer<'_> {
    fn record_parser_repairs(&mut self) -> Result<()> {
        for error in &self.dom.errors {
            self.diagnostic("html", None, format!("HTML parser repair: {error}"))?;
        }
        Ok(())
    }

    fn record_head_resources(&mut self) -> Result<()> {
        let selector = Selector::parse("head link, head script").expect("static selector");
        let resources: Vec<_> = self.dom.select(&selector).collect();
        for element in resources {
            match element.value().name() {
                "link"
                    if element.attr("rel").is_some_and(|value| {
                        value
                            .split_whitespace()
                            .any(|item| item.eq_ignore_ascii_case("stylesheet"))
                    }) =>
                {
                    self.diagnostic(
                        &element_path(element),
                        None,
                        "dropped external HTML stylesheet".to_string(),
                    )?;
                }
                "script" => {
                    self.diagnostic(
                        &element_path(element),
                        None,
                        "dropped HTML script".to_string(),
                    )?;
                }
                _ => {}
            }
        }
        Ok(())
    }

    fn collect_styles(&mut self) -> Result<()> {
        let selector = Selector::parse("style").expect("static selector");
        let styles: Vec<_> = self
            .dom
            .select(&selector)
            .map(|element| {
                let css = element.text().collect::<String>();
                (element, css)
            })
            .collect();
        let mut order = 0_usize;
        for (element, css) in styles {
            let location = element_path(element);
            let css = strip_css_comments(&css);
            let mut remainder = css.as_str();
            while let Some(open) = remainder.find('{') {
                let selector_text = remainder[..open].trim();
                let after_open = &remainder[open + 1..];
                let Some(close) = after_open.find('}') else {
                    self.diagnostic(
                        &location,
                        None,
                        "unsupported unterminated CSS rule".to_string(),
                    )?;
                    break;
                };
                let declarations = &after_open[..close];
                remainder = &after_open[close + 1..];
                if selector_text.starts_with('@') {
                    self.diagnostic(
                        &location,
                        None,
                        format!("unsupported CSS at-rule `{selector_text}`"),
                    )?;
                    continue;
                }
                for selector_part in selector_text.split(',').map(str::trim) {
                    if !supported_selector(selector_part) {
                        self.diagnostic(
                            &location,
                            None,
                            format!("unsupported CSS selector `{selector_part}`"),
                        )?;
                        continue;
                    }
                    let selector = match Selector::parse(selector_part) {
                        Ok(selector) => selector,
                        Err(_) => {
                            self.diagnostic(
                                &location,
                                None,
                                format!("invalid CSS selector `{selector_part}`"),
                            )?;
                            continue;
                        }
                    };
                    let changes = self.parse_declarations(declarations, &location)?;
                    if !changes.is_empty() {
                        if self.rules.len() >= self.limits.nodes {
                            return Err(html_error(&location, "HTML exceeds the CSS rule limit"));
                        }
                        self.rules.push(CssRule {
                            selector,
                            specificity: specificity(selector_part),
                            order,
                            changes,
                        });
                        order = order
                            .checked_add(1)
                            .ok_or_else(|| html_error(&location, "CSS source order overflowed"))?;
                    }
                }
            }
            if !remainder.trim().is_empty() {
                self.diagnostic(
                    &location,
                    None,
                    "unsupported CSS text outside a rule".to_string(),
                )?;
            }
        }
        Ok(())
    }

    fn parse_declarations(
        &mut self,
        declarations: &str,
        location: &str,
    ) -> Result<Vec<StyleChange>> {
        let mut changes = Vec::new();
        for declaration in declarations.split(';').map(str::trim) {
            if declaration.is_empty() {
                continue;
            }
            let Some((property, value)) = declaration.split_once(':') else {
                self.diagnostic(
                    location,
                    None,
                    format!("invalid CSS declaration `{declaration}`"),
                )?;
                continue;
            };
            let property = property.trim().to_ascii_lowercase();
            let value = value.trim();
            match parse_style_change(&property, value) {
                Ok(mut parsed) => changes.append(&mut parsed),
                Err(message) => {
                    self.diagnostic(location, Some(property), message)?;
                }
            }
        }
        Ok(changes)
    }

    fn project(&mut self) -> Result<()> {
        let root = self.dom.root_element();
        let body_selector = Selector::parse("body").expect("static selector");
        let container = if root.value().name() == "html" {
            root.select(&body_selector).next().unwrap_or(root)
        } else {
            root
        };
        let base = ComputedStyle::default();
        self.project_container(container, &base)
    }

    fn project_container(
        &mut self,
        container: ElementRef<'_>,
        inherited: &ComputedStyle,
    ) -> Result<()> {
        let container_style = self.computed_style(container, inherited)?;
        let mut inline_group = Vec::new();
        for child in container.children() {
            match child.value() {
                Node::Text(text) => {
                    inline_group.push(InlinePiece::Text(
                        text.to_string(),
                        Box::new(container_style.clone()),
                        false,
                    ));
                }
                Node::Element(_) => {
                    let element = ElementRef::wrap(child).expect("element checked");
                    if is_block(element.value().name()) {
                        self.flush_inline_group(&mut inline_group, &container_style)?;
                        self.project_block(element, &container_style)?;
                    } else if !matches!(element.value().name(), "head" | "style" | "title") {
                        self.collect_inline(element, &container_style, false, &mut inline_group)?;
                    }
                }
                _ => {}
            }
        }
        self.flush_inline_group(&mut inline_group, &container_style)
    }

    fn flush_inline_group(
        &mut self,
        pieces: &mut Vec<InlinePiece>,
        style: &ComputedStyle,
    ) -> Result<()> {
        if pieces.iter().any(inline_piece_is_visible) {
            let model = ParagraphModel {
                pieces: std::mem::take(pieces),
                style: style.clone(),
                paragraph_style: None,
                numbering: None,
            };
            let paragraph = self.build_paragraph(model)?;
            self.output
                .push(BodyContent::Paragraph(Box::new(paragraph)));
        } else {
            pieces.clear();
        }
        Ok(())
    }

    fn project_block(&mut self, element: ElementRef<'_>, inherited: &ComputedStyle) -> Result<()> {
        let name = element.value().name();
        match name {
            "ul" | "ol" => {
                let level = element
                    .ancestors()
                    .filter_map(ElementRef::wrap)
                    .filter(|ancestor| ancestor.value().name() == "li")
                    .count() as u32;
                self.project_list(element, inherited, level, None)
            }
            "table" => self.project_table(element, inherited),
            "hr" => {
                let style = self.computed_style(element, inherited)?;
                let mut paragraph = self.build_paragraph(ParagraphModel {
                    pieces: Vec::new(),
                    style,
                    paragraph_style: Some("HTMLThematicBreak".to_owned()),
                    numbering: None,
                })?;
                Paragraph {
                    inner: &mut paragraph,
                }
                .set_border_bottom(crate::BorderStyle::Single, 4, "000000");
                self.output
                    .push(BodyContent::Paragraph(Box::new(paragraph)));
                Ok(())
            }
            "div" => self.project_container(element, inherited),
            "blockquote" => {
                let start = self.output.len();
                self.project_container(element, inherited)?;
                let depth = element
                    .ancestors()
                    .filter_map(ElementRef::wrap)
                    .filter(|ancestor| ancestor.value().name() == "blockquote")
                    .count()
                    + 1;
                let container_path = self.html_container_path(element);
                let quote_id = html_style_id(self.document, "Quote");
                let pre_id = html_style_id(self.document, "HTMLPreformatted");
                let quoted_pre_id = html_style_id(self.document, "HTMLQuotedPreformatted");
                for content in &mut self.output[start..] {
                    if let BodyContent::Paragraph(inner) = content {
                        let mut paragraph = Paragraph { inner };
                        let style_id = paragraph
                            .inner
                            .properties
                            .as_ref()
                            .and_then(|properties| properties.style_id.as_deref());
                        if style_id != Some(quote_id.as_str())
                            && style_id != Some(quoted_pre_id.as_str())
                            && !style_id.is_some_and(|id| self.container_styles.contains(id))
                        {
                            if style_id == Some(pre_id.as_str()) {
                                paragraph.set_style(&quoted_pre_id);
                            } else if style_id != Some("HTMLThematicBreak") {
                                paragraph.set_style(&quote_id);
                            }
                            paragraph.set_indent_left(Length::inches(depth as f64 * 0.25));
                        }
                    }
                }
                self.retain_container_path(start, &container_path)?;
                Ok(())
            }
            "p" | "pre" | "h1" | "h2" | "h3" | "h4" | "h5" | "h6" => {
                let style = self.computed_style(element, inherited)?;
                let mut pieces = Vec::new();
                let pre = name == "pre";
                self.collect_inline_children(element, &style, pre, &mut pieces)?;
                let paragraph_style = match name {
                    "h1" => Some("Heading1".to_string()),
                    "h2" => Some("Heading2".to_string()),
                    "h3" => Some("Heading3".to_string()),
                    "h4" => Some("Heading4".to_string()),
                    "h5" => Some("Heading5".to_string()),
                    "h6" => Some("Heading6".to_string()),
                    "pre" => Some("HTMLPreformatted".to_string()),
                    _ => None,
                };
                let model = ParagraphModel {
                    pieces,
                    style,
                    paragraph_style,
                    numbering: None,
                };
                let paragraph = self.build_paragraph(model)?;
                self.output
                    .push(BodyContent::Paragraph(Box::new(paragraph)));
                Ok(())
            }
            _ => {
                self.diagnostic(
                    &element_path(element),
                    None,
                    format!("unsupported visible HTML element `{name}` retained as content"),
                )?;
                self.project_container(element, inherited)
            }
        }
    }

    fn html_container_path(&mut self, element: ElementRef<'_>) -> String {
        let mut ancestors: Vec<_> = element.ancestors().filter_map(ElementRef::wrap).collect();
        ancestors.reverse();
        ancestors.push(element);
        let path: String = ancestors
            .iter()
            .filter_map(|ancestor| match ancestor.value().name() {
                "blockquote" => Some('q'),
                "li" => Some('l'),
                _ => None,
            })
            .collect();
        let quote = ancestors
            .iter()
            .find(|ancestor| ancestor.value().name() == "blockquote");
        let compound = quote.is_some_and(|quote| {
            self.quote_owners.contains_key(&format!("{:?}", quote.id()))
                || quote
                    .descendants()
                    .filter_map(ElementRef::wrap)
                    .any(|child| matches!(child.value().name(), "ul" | "ol" | "table" | "hr"))
        });
        if compound {
            let mut owners = Vec::new();
            for quote in ancestors
                .iter()
                .filter(|ancestor| ancestor.value().name() == "blockquote")
            {
                let next = self.quote_owners.len() as u32 + 1;
                let owner = *self
                    .quote_owners
                    .entry(format!("{:?}", quote.id()))
                    .or_insert(next);
                owners.push(owner.to_string());
            }
            format!("{path}|{}", owners.join("."))
        } else {
            path
        }
    }

    fn retain_container_path(&mut self, start: usize, path: &str) -> Result<()> {
        if !path.contains('q') {
            return Ok(());
        }
        for index in start..self.output.len() {
            let (base, table) = match &self.output[index] {
                BodyContent::Paragraph(paragraph) => (
                    paragraph
                        .properties
                        .as_ref()
                        .and_then(|p| p.style_id.clone()),
                    false,
                ),
                BodyContent::Table(table) => (
                    table.properties.as_ref().and_then(|p| p.style_id.clone()),
                    true,
                ),
                _ => continue,
            };
            if base
                .as_deref()
                .is_some_and(|id| self.container_styles.contains(id))
            {
                continue;
            }
            if !path.contains('|')
                && path.bytes().all(|b| b == b'q')
                && !table
                && base.as_deref() != Some("HTMLThematicBreak")
            {
                continue;
            }
            let name = format!("RDOCX HTML containers {path}");
            let stem = format!(
                "RDOCXHTMLContainers{}{}",
                path.replace('|', "Owner"),
                if table { "Table" } else { "Paragraph" }
            );
            let mut suffix = 0u32;
            let id = loop {
                let id = format!("{stem}{suffix}");
                let mut builder = if table {
                    crate::StyleBuilder::table(&id, &name)
                } else {
                    crate::StyleBuilder::paragraph(&id, &name)
                };
                if let Some(base) = &base {
                    builder = builder.based_on(base);
                }
                let expected = builder.build().0;
                match self.document.styles.get_by_id(&id) {
                    Some(style) if *style == expected => break id,
                    Some(_) => suffix += 1,
                    None => {
                        let mut builder = if table {
                            crate::StyleBuilder::table(&id, &name)
                        } else {
                            crate::StyleBuilder::paragraph(&id, &name)
                        };
                        if let Some(base) = &base {
                            builder = builder.based_on(base);
                        }
                        self.document.add_style(builder)?;
                        break id;
                    }
                }
            };
            self.container_styles.insert(id.clone());
            match &mut self.output[index] {
                BodyContent::Paragraph(paragraph) => {
                    paragraph
                        .properties
                        .get_or_insert_with(Default::default)
                        .style_id = Some(id)
                }
                BodyContent::Table(table) => {
                    table
                        .properties
                        .get_or_insert_with(Default::default)
                        .style_id = Some(id)
                }
                _ => {}
            }
        }
        Ok(())
    }

    fn list_definition(
        &mut self,
        list: ElementRef<'_>,
        level: u32,
        list_id: Option<u32>,
    ) -> Result<u32> {
        if level > 8 {
            return Err(html_error(
                element_path(list),
                "HTML list nesting exceeds Word's nine-level limit",
            ));
        }
        let ordered = list.value().name() == "ol";
        let list_id = list_id.unwrap_or_else(|| {
            let levels: Vec<_> = (0..9)
                .map(|_| {
                    if ordered {
                        ListLevel::decimal()
                    } else {
                        ListLevel::bullet()
                    }
                })
                .collect();
            self.document.add_list_definition(&levels)
        });
        let mut spec = if ordered {
            ListLevel::decimal()
        } else {
            ListLevel::bullet()
        };
        if ordered {
            if let Some(start) = list.attr("start") {
                match start.parse::<u32>() {
                    Ok(start) if start > 0 => spec = spec.start(start),
                    _ => {
                        self.diagnostic(
                            &element_path(list),
                            Some("start".to_string()),
                            format!("unsupported ordered-list start value `{start}`"),
                        )?;
                    }
                }
            }
            if list.attr("reversed").is_some() {
                self.diagnostic(
                    &element_path(list),
                    Some("reversed".to_string()),
                    "unsupported reversed ordered list".to_string(),
                )?;
            }
        }
        if !self.document.set_list_level(list_id, level, spec) {
            return Err(html_error(
                element_path(list),
                "could not define list level",
            ));
        }
        Ok(list_id)
    }

    fn attach_list_paragraph(paragraph: &mut CT_P, list_id: u32, level: u32, first: &mut bool) {
        let properties = paragraph.properties.get_or_insert_with(Default::default);
        properties.num_id = Some(if *first { list_id } else { 0 });
        properties.num_ilvl = Some(level);
        if !*first {
            let mut paragraph = Paragraph { inner: paragraph };
            paragraph.set_indent_left(Length::inches((level + 1) as f64 * 0.25));
        }
        *first = false;
    }

    fn flush_list_inline(
        &mut self,
        pieces: &mut Vec<InlinePiece>,
        style: &ComputedStyle,
        list_id: u32,
        level: u32,
        first: &mut bool,
    ) -> Result<()> {
        if pieces.iter().any(inline_piece_is_visible) {
            let mut paragraph = self.build_paragraph(ParagraphModel {
                pieces: std::mem::take(pieces),
                style: style.clone(),
                paragraph_style: None,
                numbering: None,
            })?;
            Self::attach_list_paragraph(&mut paragraph, list_id, level, first);
            self.output
                .push(BodyContent::Paragraph(Box::new(paragraph)));
        } else {
            pieces.clear();
        }
        Ok(())
    }

    fn project_list(
        &mut self,
        list: ElementRef<'_>,
        inherited: &ComputedStyle,
        level: u32,
        list_id: Option<u32>,
    ) -> Result<()> {
        let list_id = self.list_definition(list, level, list_id)?;
        let style = self.computed_style(list, inherited)?;
        for item in list
            .child_elements()
            .filter(|child| child.value().name() == "li")
        {
            let item_style = self.computed_style(item, &style)?;
            let start = self.output.len();
            let mut first = true;
            let mut pieces = Vec::new();
            for child in item.children() {
                match child.value() {
                    Node::Text(text) => pieces.push(InlinePiece::Text(
                        text.to_string(),
                        Box::new(item_style.clone()),
                        false,
                    )),
                    Node::Element(_) => {
                        let element = ElementRef::wrap(child).expect("element checked");
                        let name = element.value().name();
                        if is_block(name) {
                            self.flush_list_inline(
                                &mut pieces,
                                &item_style,
                                list_id,
                                level,
                                &mut first,
                            )?;
                            if matches!(name, "ul" | "ol" | "table") && first {
                                let mut paragraph = self.build_paragraph(ParagraphModel {
                                    pieces: Vec::new(),
                                    style: item_style.clone(),
                                    paragraph_style: None,
                                    numbering: None,
                                })?;
                                Self::attach_list_paragraph(
                                    &mut paragraph,
                                    list_id,
                                    level,
                                    &mut first,
                                );
                                self.output
                                    .push(BodyContent::Paragraph(Box::new(paragraph)));
                            }
                            if matches!(name, "ul" | "ol") {
                                // Separate nested definitions preserve mixed ordered/unordered siblings.
                                self.project_list(element, &item_style, level + 1, None)?;
                            } else {
                                let start = self.output.len();
                                self.project_block(element, &item_style)?;
                                for content in &mut self.output[start..] {
                                    if let BodyContent::Paragraph(paragraph) = content {
                                        if !paragraph
                                            .properties
                                            .as_ref()
                                            .is_some_and(|p| p.num_id.is_some_and(|id| id != 0))
                                        {
                                            Self::attach_list_paragraph(
                                                paragraph, list_id, level, &mut first,
                                            );
                                        }
                                    } else if let BodyContent::Table(table) = content {
                                        table
                                            .properties
                                            .get_or_insert_with(Default::default)
                                            .indent =
                                            Some(CT_TblWidth::dxa((level + 1) as i32 * 360));
                                    }
                                }
                            }
                        } else {
                            self.collect_inline(element, &item_style, false, &mut pieces)?;
                        }
                    }
                    _ => {}
                }
            }
            self.flush_list_inline(&mut pieces, &item_style, list_id, level, &mut first)?;
            if first {
                let mut paragraph = self.build_paragraph(ParagraphModel {
                    pieces: Vec::new(),
                    style: item_style,
                    paragraph_style: None,
                    numbering: None,
                })?;
                Self::attach_list_paragraph(&mut paragraph, list_id, level, &mut first);
                self.output
                    .push(BodyContent::Paragraph(Box::new(paragraph)));
            }
            let path = self.html_container_path(item);
            self.retain_container_path(start, &path)?;
        }
        Ok(())
    }

    fn list_models(
        &mut self,
        list: ElementRef<'_>,
        inherited: &ComputedStyle,
        level: u32,
        list_id: Option<u32>,
    ) -> Result<Vec<ParagraphModel>> {
        let list_id = self.list_definition(list, level, list_id)?;
        let style = self.computed_style(list, inherited)?;
        let mut models = Vec::new();
        for item in list
            .child_elements()
            .filter(|child| child.value().name() == "li")
        {
            let item_style = self.computed_style(item, &style)?;
            let mut pieces = Vec::new();
            self.collect_inline_children(item, &item_style, false, &mut pieces)?;
            models.push(ParagraphModel {
                pieces,
                style: item_style.clone(),
                paragraph_style: None,
                numbering: Some((list_id, level)),
            });
            for nested in item
                .child_elements()
                .filter(|child| matches!(child.value().name(), "ul" | "ol"))
            {
                models.extend(self.list_models(nested, &item_style, level + 1, Some(list_id))?);
            }
        }
        Ok(models)
    }

    fn project_table(&mut self, table: ElementRef<'_>, inherited: &ComputedStyle) -> Result<()> {
        let style = self.computed_style(table, inherited)?;
        for caption in table
            .child_elements()
            .filter(|child| child.value().name() == "caption")
        {
            self.diagnostic(
                &element_path(caption),
                None,
                "unsupported HTML table caption retained before the table".to_string(),
            )?;
            let caption_style = self.computed_style(caption, &style)?;
            let mut pieces = Vec::new();
            self.collect_inline_children(caption, &caption_style, false, &mut pieces)?;
            let paragraph = self.build_paragraph(ParagraphModel {
                pieces,
                style: caption_style,
                paragraph_style: None,
                numbering: None,
            })?;
            self.output
                .push(BodyContent::Paragraph(Box::new(paragraph)));
        }
        let row_selector = Selector::parse("tr").expect("static selector");
        let rows: Vec<_> = table
            .select(&row_selector)
            .filter(|row| nearest_ancestor_named(*row, "table") == Some(table))
            .collect();
        self.rows = self
            .rows
            .checked_add(rows.len())
            .ok_or_else(|| html_error(element_path(table), "table row count overflowed"))?;
        if self.rows > self.limits.rows {
            return Err(html_error(
                element_path(table),
                "HTML tables exceed the row limit",
            ));
        }

        let mut models = Vec::with_capacity(rows.len());
        let mut active: Vec<Option<ActiveSpan>> = Vec::new();
        let mut column_count = 0_usize;
        for row in rows {
            let previous = std::mem::take(&mut active);
            let mut occupied = vec![false; previous.len()];
            let mut cells = Vec::new();
            for (start, span) in previous.iter().enumerate() {
                if let Some(span) = span {
                    let end = start
                        .checked_add(span.span)
                        .ok_or_else(|| html_error(element_path(row), "table span overflowed"))?;
                    if occupied.len() < end {
                        occupied.resize(end, false);
                    }
                    occupied[start..end].fill(true);
                    cells.push(TableCellModel {
                        start,
                        span: span.span,
                        v_merge: Some(VMerge::Continue),
                        paragraphs: Vec::new(),
                        header: false,
                    });
                    if span.remaining > 1 {
                        if active.len() <= start {
                            active.resize(start + 1, None);
                        }
                        active[start] = Some(ActiveSpan {
                            remaining: span.remaining - 1,
                            span: span.span,
                        });
                    }
                }
            }

            let mut cursor = 0_usize;
            for cell in row
                .child_elements()
                .filter(|cell| matches!(cell.value().name(), "td" | "th"))
            {
                let colspan =
                    parse_span(cell.attr("colspan"), "colspan", cell, self.limits.columns)?;
                let rowspan = parse_span(cell.attr("rowspan"), "rowspan", cell, self.limits.rows)?;
                loop {
                    while occupied.get(cursor).copied().unwrap_or(false) {
                        cursor = cursor.checked_add(1).ok_or_else(|| {
                            html_error(element_path(cell), "table column overflowed")
                        })?;
                    }
                    let end = cursor.checked_add(colspan).ok_or_else(|| {
                        html_error(element_path(cell), "table column span overflowed")
                    })?;
                    if end > self.limits.columns {
                        return Err(html_error(
                            element_path(cell),
                            "HTML table exceeds the column limit",
                        ));
                    }
                    if occupied.len() < end {
                        occupied.resize(end, false);
                    }
                    if occupied[cursor..end].iter().any(|value| *value) {
                        cursor = cursor.checked_add(1).ok_or_else(|| {
                            html_error(element_path(cell), "table column overflowed")
                        })?;
                        continue;
                    }
                    occupied[cursor..end].fill(true);
                    let cell_style = self.computed_style(cell, &style)?;
                    let paragraphs = self.cell_paragraphs(cell, &cell_style)?;
                    cells.push(TableCellModel {
                        start: cursor,
                        span: colspan,
                        v_merge: (rowspan > 1).then_some(VMerge::Restart),
                        paragraphs,
                        header: cell.value().name() == "th",
                    });
                    if rowspan > 1 {
                        if active.len() < end {
                            active.resize(end, None);
                        }
                        active[cursor] = Some(ActiveSpan {
                            remaining: rowspan - 1,
                            span: colspan,
                        });
                    }
                    cursor = end;
                    break;
                }
            }
            column_count = column_count.max(occupied.len());
            cells.sort_by_key(|cell| cell.start);
            self.cells = self
                .cells
                .checked_add(cells.len())
                .ok_or_else(|| html_error(element_path(row), "table cell count overflowed"))?;
            if self.cells > self.limits.cells {
                return Err(html_error(
                    element_path(table),
                    "HTML tables exceed the cell limit",
                ));
            }
            models.push(TableRowModel {
                header: cells.iter().any(|cell| cell.header),
                cells,
            });
        }
        if column_count == 0 {
            self.diagnostic(
                &element_path(table),
                None,
                "dropped empty HTML table".to_string(),
            )?;
            return Ok(());
        }
        if column_count > self.limits.columns {
            return Err(html_error(
                element_path(table),
                "HTML table exceeds the column limit",
            ));
        }

        let width = Twips(9360 / column_count.max(1) as i32);
        let mut output = CT_Tbl::new();
        output.properties = Some(CT_TblPr {
            width: Some(CT_TblWidth::dxa(width.0 * column_count as i32)),
            ..CT_TblPr::default()
        });
        output.grid = Some(CT_TblGrid {
            columns: (0..column_count).map(|_| CT_TblGridCol { width }).collect(),
            ..Default::default()
        });
        for model in models {
            let mut row = CT_Row::new();
            if model.header {
                row.properties.get_or_insert_with(Default::default).header = Some(true);
            }
            for model_cell in model.cells {
                let mut cell = CT_Tc::new();
                cell.content.clear();
                {
                    let mut facade = Cell { inner: &mut cell };
                    if model_cell.span > 1 {
                        facade.set_grid_span(model_cell.span as u32);
                    }
                    match model_cell.v_merge {
                        Some(VMerge::Restart) => facade.set_v_merge_restart(),
                        Some(VMerge::Continue) => facade.set_v_merge_continue(),
                        None => {}
                    }
                    if model_cell.header {
                        facade.set_shading("D9EAF7");
                    }
                }
                for paragraph in model_cell.paragraphs {
                    cell.content.push(rdocx_oxml::table::CellContent::Paragraph(
                        self.build_paragraph(paragraph)?,
                    ));
                }
                if cell.content.is_empty() {
                    cell.content
                        .push(rdocx_oxml::table::CellContent::Paragraph(CT_P::new()));
                }
                row.cells.push(cell);
            }
            output.rows.push(row);
        }
        self.bump_blocks(&element_path(table))?;
        self.output.push(BodyContent::Table(Box::new(output)));
        Ok(())
    }

    fn cell_paragraphs(
        &mut self,
        cell: ElementRef<'_>,
        inherited: &ComputedStyle,
    ) -> Result<Vec<ParagraphModel>> {
        let mut paragraphs = Vec::new();
        let mut inline_group = Vec::new();
        for child in cell.children() {
            match child.value() {
                Node::Text(text) => inline_group.push(InlinePiece::Text(
                    text.to_string(),
                    Box::new(inherited.clone()),
                    false,
                )),
                Node::Element(_) => {
                    let element = ElementRef::wrap(child).expect("element checked");
                    if matches!(element.value().name(), "p" | "div" | "blockquote" | "pre") {
                        if inline_group.iter().any(inline_piece_is_visible) {
                            paragraphs.push(ParagraphModel {
                                pieces: std::mem::take(&mut inline_group),
                                style: inherited.clone(),
                                paragraph_style: None,
                                numbering: None,
                            });
                        }
                        let style = self.computed_style(element, inherited)?;
                        let mut pieces = Vec::new();
                        self.collect_inline_children(
                            element,
                            &style,
                            element.value().name() == "pre",
                            &mut pieces,
                        )?;
                        paragraphs.push(ParagraphModel {
                            pieces,
                            style,
                            paragraph_style: None,
                            numbering: None,
                        });
                    } else if matches!(element.value().name(), "ul" | "ol") {
                        if inline_group.iter().any(inline_piece_is_visible) {
                            paragraphs.push(ParagraphModel {
                                pieces: std::mem::take(&mut inline_group),
                                style: inherited.clone(),
                                paragraph_style: None,
                                numbering: None,
                            });
                        }
                        paragraphs.extend(self.list_models(element, inherited, 0, None)?);
                    } else if element.value().name() != "table" {
                        self.collect_inline(element, inherited, false, &mut inline_group)?;
                    } else {
                        self.diagnostic(
                            &element_path(element),
                            None,
                            "dropped nested HTML table".to_string(),
                        )?;
                    }
                }
                _ => {}
            }
        }
        if inline_group.iter().any(inline_piece_is_visible) || paragraphs.is_empty() {
            paragraphs.push(ParagraphModel {
                pieces: inline_group,
                style: inherited.clone(),
                paragraph_style: None,
                numbering: None,
            });
        }
        Ok(paragraphs)
    }

    fn collect_inline_children(
        &mut self,
        element: ElementRef<'_>,
        style: &ComputedStyle,
        pre: bool,
        pieces: &mut Vec<InlinePiece>,
    ) -> Result<()> {
        for child in element.children() {
            match child.value() {
                Node::Text(text) => {
                    pieces.push(InlinePiece::Text(
                        text.to_string(),
                        Box::new(style.clone()),
                        pre,
                    ));
                }
                Node::Element(_) => {
                    let child = ElementRef::wrap(child).expect("element checked");
                    if matches!(child.value().name(), "ul" | "ol") {
                        continue;
                    }
                    if child.value().name() == "table" {
                        self.diagnostic(
                            &element_path(child),
                            None,
                            "dropped nested HTML table".to_string(),
                        )?;
                        continue;
                    }
                    self.collect_inline(child, style, pre, pieces)?;
                }
                _ => {}
            }
        }
        Ok(())
    }

    fn collect_inline(
        &mut self,
        element: ElementRef<'_>,
        inherited: &ComputedStyle,
        pre: bool,
        pieces: &mut Vec<InlinePiece>,
    ) -> Result<()> {
        let name = element.value().name();
        let location = element_path(element);
        match name {
            "br" => {
                pieces.push(InlinePiece::Break);
                return Ok(());
            }
            "script" => {
                self.diagnostic(&location, None, "dropped HTML script".to_string())?;
                return Ok(());
            }
            "style" | "head" | "title" => return Ok(()),
            "img" => {
                if let Some(projection) = self.projection {
                    let Some(source) = element.attr("src") else {
                        if projection.strict_references {
                            return Err(mhtml_error(None, 0, "MHTML image has no src"));
                        }
                        self.diagnostic(
                            &location,
                            None,
                            "dropped HTML image without a source and retained alternate text"
                                .to_string(),
                        )?;
                        if let Some(alt) = element.attr("alt")
                            && !alt.is_empty()
                        {
                            let style = self.computed_style(element, inherited)?;
                            pieces.push(InlinePiece::Text(alt.to_string(), Box::new(style), pre));
                        }
                        return Ok(());
                    };
                    let Some(resource) = projection.resources.get(source) else {
                        if projection.strict_references {
                            return Err(mhtml_error(
                                None,
                                0,
                                format!("unresolved image resource `{source}`"),
                            ));
                        }
                        self.diagnostic(
                            &location,
                            None,
                            format!(
                                "dropped unresolved HTML image `{source}` and retained alternate text"
                            ),
                        )?;
                        if let Some(alt) = element.attr("alt")
                            && !alt.is_empty()
                        {
                            let style = self.computed_style(element, inherited)?;
                            pieces.push(InlinePiece::Text(alt.to_string(), Box::new(style), pre));
                        }
                        return Ok(());
                    };
                    let declared_width = image_dimension(element, "width")?;
                    let declared_height = image_dimension(element, "height")?;
                    let (width, height) = match (declared_width, declared_height) {
                        (Some(width), Some(height)) => (width, height),
                        dimensions => {
                            let size = oxml_media::probe(&resource.bytes)
                                .and_then(|info| info.native_size(96.0))
                                .ok_or_else(|| {
                                    mhtml_error(
                                        None,
                                        0,
                                        format!("image dimensions are unavailable for `{source}`"),
                                    )
                                })?;
                            match dimensions {
                                (Some(width), None) => (
                                    width,
                                    Length::emu(
                                        (width.to_emu() as f64 * size.height_emu as f64
                                            / size.width_emu as f64)
                                            as i64,
                                    ),
                                ),
                                (None, Some(height)) => (
                                    Length::emu(
                                        (height.to_emu() as f64 * size.width_emu as f64
                                            / size.height_emu as f64)
                                            as i64,
                                    ),
                                    height,
                                ),
                                (None, None) => {
                                    (Length::emu(size.width_emu), Length::emu(size.height_emu))
                                }
                                (Some(_), Some(_)) => unreachable!(),
                            }
                        }
                    };
                    let rel_id = if let Some(rel_id) = self.embedded_images.get(source) {
                        rel_id.clone()
                    } else {
                        let rel_id = if let Some(story) = self.story {
                            self.document.add_html_image_relationship_to_story(
                                story,
                                &resource.bytes,
                                &resource.filename,
                            )?
                        } else {
                            self.document
                                .embed_image(&resource.bytes, &resource.filename)
                        };
                        self.embedded_images
                            .insert(source.to_owned(), rel_id.clone());
                        rel_id
                    };
                    let image = EmbeddedHtmlImage {
                        rel_id,
                        drawing_id: self
                            .story
                            .map(|story| self.document.reserve_html_drawing_id(story))
                            .transpose()?,
                        description: element.attr("alt").map(str::to_owned),
                        width,
                        height,
                    };
                    pieces.push(InlinePiece::Image(image));
                    return Ok(());
                }
                self.diagnostic(
                    &location,
                    None,
                    "dropped HTML image and retained alternate text".to_string(),
                )?;
                if let Some(alt) = element.attr("alt")
                    && !alt.is_empty()
                {
                    let style = self.computed_style(element, inherited)?;
                    pieces.push(InlinePiece::Text(alt.to_string(), Box::new(style), pre));
                }
                return Ok(());
            }
            "link" => {
                self.diagnostic(
                    &location,
                    None,
                    "dropped external HTML stylesheet".to_string(),
                )?;
                return Ok(());
            }
            "iframe" | "frame" | "embed" | "object" => {
                self.diagnostic(&location, None, format!("dropped HTML {name} content"))?;
                return Ok(());
            }
            "input" | "button" | "select" | "textarea" => {
                self.diagnostic(
                    &location,
                    None,
                    format!("dropped HTML form control `{name}`"),
                )?;
                if let Some(value) = element.attr("value")
                    && !value.is_empty()
                {
                    let style = self.computed_style(element, inherited)?;
                    pieces.push(InlinePiece::Text(value.to_string(), Box::new(style), pre));
                }
                return Ok(());
            }
            "a" if element.attr("href").is_some() => {
                let href = element.attr("href").expect("anchor href checked");
                if self.projection.is_none()
                    || self.projection.is_some_and(|projection| {
                        !projection.strict_references && !projection.hyperlinks.contains_key(href)
                    })
                {
                    self.diagnostic(
                        &location,
                        None,
                        "dropped HTML link target and retained anchor text".to_string(),
                    )?;
                }
            }
            "form" => {
                self.diagnostic(
                    &location,
                    None,
                    "dropped HTML form semantics and retained supported content".to_string(),
                )?;
            }
            _ if !is_supported_inline(name) => {
                self.diagnostic(
                    &location,
                    None,
                    format!("unsupported visible HTML element `{name}` retained as text"),
                )?;
            }
            _ => {}
        }

        let mut style = self.computed_style(element, inherited)?;
        if name == "a"
            && let Some(href) = element.attr("href")
            && let Some(projection) = self.projection
        {
            if let Some(target) = projection.hyperlinks.get(href) {
                style.hyperlink = Some(target.clone());
            } else if projection.strict_references {
                return Err(mhtml_error(
                    None,
                    0,
                    format!("unsafe hyperlink target `{href}`"),
                ));
            }
        }
        self.collect_inline_children(element, &style, pre || name == "code", pieces)
    }

    fn computed_style(
        &mut self,
        element: ElementRef<'_>,
        inherited: &ComputedStyle,
    ) -> Result<ComputedStyle> {
        let mut style = inherited.inherited();
        match element.value().name() {
            "b" | "strong" => style.bold = Some(true),
            "i" | "em" | "cite" => style.italic = Some(true),
            "u" | "ins" => style.underline = Some(true),
            "s" | "strike" | "del" => style.strike = Some(true),
            "sup" => style.vertical = Some(VerticalText::Superscript),
            "sub" => style.vertical = Some(VerticalText::Subscript),
            "code" | "kbd" | "samp" => {
                let id = html_style_id(self.document, "HTMLCode");
                if self.document.styles.get_by_id(&id).is_some() {
                    style.character_style = Some(id);
                } else {
                    style.font = Some("Courier New".to_string());
                }
                style.code = true;
            }
            "mark" => style.background = Some("FFFF00".to_string()),
            _ => {}
        }
        let mut matching: Vec<_> = self
            .rules
            .iter()
            .filter(|rule| rule.selector.matches(&element))
            .collect();
        matching.sort_by(|left, right| {
            let specificity = compare_specificity(left.specificity, right.specificity);
            if specificity == Ordering::Equal {
                left.order.cmp(&right.order)
            } else {
                specificity
            }
        });
        for rule in matching {
            for change in &rule.changes {
                change.apply(&mut style);
            }
        }
        if let Some(inline) = element.attr("style") {
            let location = element_path(element);
            for change in self.parse_declarations(inline, &location)? {
                change.apply(&mut style);
            }
        }
        if element.value().name() == "mark" {
            style.mark_highlight = style.background.clone();
        }
        Ok(style)
    }

    fn build_paragraph(&mut self, model: ParagraphModel) -> Result<CT_P> {
        self.bump_blocks("html")?;
        let mut hyperlink_ids = HashMap::new();
        for piece in &model.pieces {
            if let InlinePiece::Text(_, style, _) = piece
                && let Some(url) = &style.hyperlink
                && !hyperlink_ids.contains_key(url)
            {
                let relationship_id = if let Some(story) = self.story {
                    self.document
                        .add_html_hyperlink_relationship_to_story(story, url)?
                } else {
                    self.document.add_hyperlink_relationship(url)
                };
                hyperlink_ids.insert(url.clone(), relationship_id);
            }
        }
        let mut output = CT_P::new();
        {
            let mut paragraph = Paragraph { inner: &mut output };
            apply_paragraph_style(&mut paragraph, &model.style);
            if let Some(style) = model.paragraph_style {
                paragraph.set_style(&html_style_id(self.document, &style));
                if let Some(level) = style
                    .strip_prefix("Heading")
                    .and_then(|level| level.parse::<u32>().ok())
                    && (1..=6).contains(&level)
                {
                    paragraph
                        .inner
                        .properties
                        .get_or_insert_with(Default::default)
                        .outline_lvl = Some(level - 1);
                }
            }
            if let Some((list_id, level)) = model.numbering
                && !paragraph.set_numbering(list_id, level)
            {
                return Err(html_error("html", "could not attach paragraph numbering"));
            }
            let mut emitted = false;
            let mut pending_space = false;
            let mut space_style: Option<Box<ComputedStyle>> = None;
            let mut space_ends_last_run = false;
            for piece in model.pieces {
                match piece {
                    InlinePiece::Break => {
                        self.bump_runs()?;
                        paragraph.add_line_break();
                        emitted = false;
                        pending_space = false;
                        space_style = None;
                    }
                    InlinePiece::Image(image) => {
                        self.bump_runs()?;
                        let run = paragraph.add_picture(&image.rel_id, image.width, image.height);
                        if let Some(RunContent::Drawing(drawing)) = run.content.last_mut()
                            && let Some(inline) = drawing.inline.as_mut()
                        {
                            inline.description = image.description;
                        }
                        if let Some(drawing_id) = image.drawing_id
                            && let Some(RunContent::Drawing(drawing)) = run.content.last_mut()
                            && let Some(inline) = drawing.inline.as_mut()
                        {
                            inline.doc_pr_id = drawing_id;
                        }
                        emitted = true;
                        pending_space = false;
                    }
                    InlinePiece::Text(text, style, true) => {
                        if pending_space
                            && emitted
                            && let Some(space) = space_style.take()
                        {
                            if space_ends_last_run
                                && let Some(last) = paragraph.inner.runs.last_mut()
                            {
                                let mut run = Run { inner: last };
                                run.set_text(&format!("{} ", run.text()));
                            } else {
                                self.bump_runs()?;
                                let mut run = if let Some(url) = &space.hyperlink {
                                    paragraph.add_hyperlink(" ", &hyperlink_ids[url])
                                } else {
                                    paragraph.add_run(" ")
                                };
                                apply_run_style(&mut run, &space);
                            }
                        }
                        let mut lines = text.split('\n').peekable();
                        while let Some(line) = lines.next() {
                            if !line.is_empty() {
                                self.bump_runs()?;
                                let mut run = if let Some(url) = &style.hyperlink {
                                    paragraph.add_hyperlink(line, &hyperlink_ids[url])
                                } else {
                                    paragraph.add_run(line)
                                };
                                apply_run_style(&mut run, &style);
                                emitted = true;
                            }
                            if lines.peek().is_some() {
                                self.bump_runs()?;
                                paragraph.add_line_break();
                                emitted = true;
                            }
                        }
                        pending_space = false;
                    }
                    InlinePiece::Text(text, style, false) => {
                        let mut text = text.as_str();
                        if pending_space
                            && emitted
                            && let Some(space) = space_style.take_if(|space| {
                                !same_run_format(space, &style)
                                    && text
                                        .chars()
                                        .any(|character| !character.is_ascii_whitespace())
                            })
                        {
                            if space_ends_last_run
                                && let Some(last) = paragraph.inner.runs.last_mut()
                            {
                                let mut run = Run { inner: last };
                                run.set_text(&format!("{} ", run.text()));
                            } else {
                                self.bump_runs()?;
                                let mut run = if let Some(url) = &space.hyperlink {
                                    paragraph.add_hyperlink(" ", &hyperlink_ids[url])
                                } else {
                                    paragraph.add_run(" ")
                                };
                                apply_run_style(&mut run, &space);
                            }
                            pending_space = false;
                            text = text.trim_start_matches(|character: char| {
                                character.is_ascii_whitespace()
                            });
                        }
                        let was_pending = pending_space;
                        let normalized = collapse_text(text, &mut pending_space, emitted);
                        if !normalized.is_empty() {
                            self.bump_runs()?;
                            let mut run = if let Some(url) = &style.hyperlink {
                                paragraph.add_hyperlink(&normalized, &hyperlink_ids[url])
                            } else {
                                paragraph.add_run(&normalized)
                            };
                            apply_run_style(&mut run, &style);
                            emitted = true;
                        }
                        if pending_space && (!was_pending || !normalized.is_empty()) {
                            space_ends_last_run = !normalized.is_empty();
                            space_style = Some(style);
                        }
                    }
                }
            }
        }
        Ok(output)
    }

    fn bump_blocks(&mut self, location: &str) -> Result<()> {
        self.blocks = self
            .blocks
            .checked_add(1)
            .ok_or_else(|| html_error(location, "projected block count overflowed"))?;
        if self.blocks > self.limits.blocks {
            return Err(html_error(
                location,
                "HTML exceeds the projected block limit",
            ));
        }
        Ok(())
    }

    fn bump_runs(&mut self) -> Result<()> {
        self.runs = self
            .runs
            .checked_add(1)
            .ok_or_else(|| html_error("html", "projected run count overflowed"))?;
        if self.runs > self.limits.runs {
            return Err(html_error("html", "HTML exceeds the run limit"));
        }
        Ok(())
    }

    fn diagnostic(
        &mut self,
        location: &str,
        property: Option<String>,
        message: String,
    ) -> Result<()> {
        let diagnostic = HtmlDiagnostic {
            location: location.to_string(),
            property,
            message,
        };
        if self.diagnostic_keys.contains(&diagnostic) {
            return Ok(());
        }
        if self.diagnostics.len() >= self.limits.diagnostics {
            return Err(html_error(location, "HTML exceeds the diagnostic limit"));
        }
        self.diagnostic_keys.insert(diagnostic.clone());
        self.diagnostics.push(diagnostic);
        Ok(())
    }

    fn finish(mut self) -> Result<HtmlReadResult> {
        self.document.document.body.content.append(&mut self.output);
        let bytes = self.document.to_bytes()?;
        let document = Document::from_bytes(&bytes)?;
        Ok(HtmlReadResult {
            document,
            diagnostics: self.diagnostics,
        })
    }
}

fn apply_paragraph_style(paragraph: &mut Paragraph<'_>, style: &ComputedStyle) {
    if let Some(alignment) = style.alignment {
        paragraph.set_alignment(alignment);
    }
    if let Some(value) = style.space_before {
        paragraph.set_space_before(value);
    }
    if let Some(value) = style.space_after {
        paragraph.set_space_after(value);
    }
    if let Some(value) = style.indent_left {
        paragraph.set_indent_left(value);
    }
    if let Some(value) = style.indent_right {
        paragraph.set_indent_right(value);
    }
    if let Some(value) = style.first_line_indent {
        paragraph.set_signed_first_line_indent_value(Some(value));
    }
    if let Some(background) = &style.background {
        paragraph.set_shading(background);
    }
}

fn apply_run_style(run: &mut Run<'_>, style: &ComputedStyle) {
    if let Some(id) = &style.character_style {
        run.set_style(id);
    }
    if let Some(font) = &style.font {
        run.set_font(font);
    }
    if let Some(size) = style.size {
        run.set_size(size);
    }
    if let Some(value) = style.bold {
        run.set_bold(value);
    }
    if let Some(value) = style.italic {
        run.set_italic(value);
    }
    if let Some(value) = style.underline {
        run.set_underline(value);
    }
    if let Some(value) = style.strike {
        run.set_strike(value);
    }
    if let Some(color) = &style.color {
        run.set_color(color);
    }
    if let Some(background) = style.background.as_ref().or(style.mark_highlight.as_ref()) {
        run.set_highlight(background);
    }
    match style.vertical {
        Some(VerticalText::Superscript) => run.set_superscript(),
        Some(VerticalText::Subscript) => run.set_subscript(),
        None => {}
    }
}

fn same_run_format(left: &ComputedStyle, right: &ComputedStyle) -> bool {
    left.font == right.font
        && left.code == right.code
        && left.character_style == right.character_style
        && left.size == right.size
        && left.bold == right.bold
        && left.italic == right.italic
        && left.underline == right.underline
        && left.strike == right.strike
        && left.color == right.color
        && left.background == right.background
        && left.mark_highlight == right.mark_highlight
        && matches!(
            (left.vertical, right.vertical),
            (None, None)
                | (
                    Some(VerticalText::Superscript),
                    Some(VerticalText::Superscript)
                )
                | (Some(VerticalText::Subscript), Some(VerticalText::Subscript))
        )
        && left.hyperlink == right.hyperlink
}

fn collapse_text(text: &str, pending_space: &mut bool, already_emitted: bool) -> String {
    let mut output = String::new();
    let mut emitted = already_emitted;
    for character in text.chars() {
        if character.is_ascii_whitespace() {
            *pending_space = true;
        } else {
            if *pending_space && emitted {
                output.push(' ');
            }
            output.push(character);
            *pending_space = false;
            emitted = true;
        }
    }
    output
}

fn inline_piece_is_visible(piece: &InlinePiece) -> bool {
    match piece {
        InlinePiece::Text(text, _, pre) => {
            *pre || !text
                .trim_matches(|c: char| c.is_ascii_whitespace())
                .is_empty()
        }
        InlinePiece::Break => true,
        InlinePiece::Image(_) => true,
    }
}

fn is_block(name: &str) -> bool {
    matches!(
        name,
        "address"
            | "article"
            | "aside"
            | "blockquote"
            | "div"
            | "footer"
            | "h1"
            | "h2"
            | "h3"
            | "h4"
            | "h5"
            | "h6"
            | "header"
            | "hr"
            | "main"
            | "nav"
            | "ol"
            | "p"
            | "pre"
            | "section"
            | "table"
            | "ul"
    )
}

fn is_supported_inline(name: &str) -> bool {
    matches!(
        name,
        "a" | "abbr"
            | "b"
            | "br"
            | "cite"
            | "code"
            | "del"
            | "em"
            | "form"
            | "i"
            | "ins"
            | "kbd"
            | "mark"
            | "s"
            | "samp"
            | "span"
            | "strike"
            | "strong"
            | "sub"
            | "sup"
            | "u"
    )
}

fn parse_span(
    value: Option<&str>,
    attribute: &str,
    element: ElementRef<'_>,
    maximum: usize,
) -> Result<usize> {
    let Some(value) = value else {
        return Ok(1);
    };
    let parsed = value.parse::<usize>().map_err(|_| {
        html_error(
            element_path(element),
            format!("invalid HTML {attribute} value `{value}`"),
        )
    })?;
    if parsed == 0 || parsed > maximum {
        return Err(html_error(
            element_path(element),
            format!("HTML {attribute} value is outside the supported bound"),
        ));
    }
    Ok(parsed)
}

fn nearest_ancestor_named<'a>(element: ElementRef<'a>, name: &str) -> Option<ElementRef<'a>> {
    let mut ancestor = element.parent();
    while let Some(node) = ancestor {
        if let Some(element) = ElementRef::wrap(node)
            && element.value().name() == name
        {
            return Some(element);
        }
        ancestor = node.parent();
    }
    None
}

fn element_path(element: ElementRef<'_>) -> String {
    let mut segments = Vec::new();
    let mut current = Some(*element);
    while let Some(node) = current {
        if let Some(element) = ElementRef::wrap(node) {
            let name = element.value().name();
            let mut index = 1_usize;
            let mut sibling = element.prev_sibling();
            while let Some(node) = sibling {
                if let Some(sibling_element) = ElementRef::wrap(node)
                    && sibling_element.value().name() == name
                {
                    index += 1;
                }
                sibling = node.prev_sibling();
            }
            if matches!(name, "html" | "head" | "body") {
                segments.push(name.to_string());
            } else {
                segments.push(format!("{name}[{index}]"));
            }
        }
        current = node.parent();
    }
    segments.reverse();
    if segments.first().is_none_or(|segment| segment != "html") {
        segments.insert(0, "html".to_string());
    }
    if segments
        .get(1)
        .is_none_or(|segment| !segment.starts_with("body"))
    {
        segments.insert(1, "body".to_string());
    }
    segments.join("/")
}

fn strip_css_comments(css: &str) -> String {
    let mut output = String::with_capacity(css.len());
    let mut remainder = css;
    while let Some(start) = remainder.find("/*") {
        output.push_str(&remainder[..start]);
        let after_start = &remainder[start + 2..];
        let Some(end) = after_start.find("*/") else {
            break;
        };
        remainder = &after_start[end + 2..];
    }
    output.push_str(remainder);
    output
}

fn supported_selector(selector: &str) -> bool {
    !selector.is_empty()
        && !selector
            .chars()
            .any(|character| matches!(character, '[' | ']' | ':' | '+' | '~' | '*' | '|'))
        && selector.chars().all(|character| {
            character.is_ascii_alphanumeric()
                || character.is_ascii_whitespace()
                || matches!(character, '-' | '_' | '.' | '#' | '>')
        })
        && !selector.contains(">>")
}

fn specificity(selector: &str) -> (u32, u32, u32) {
    let ids = selector.bytes().filter(|byte| *byte == b'#').count() as u32;
    let classes = selector.bytes().filter(|byte| *byte == b'.').count() as u32;
    let types = selector
        .split(|character: char| character.is_ascii_whitespace() || character == '>')
        .filter(|part| !part.is_empty())
        .filter(|part| {
            part.as_bytes()
                .first()
                .is_some_and(|byte| byte.is_ascii_alphabetic())
        })
        .count() as u32;
    (ids, classes, types)
}

fn compare_specificity(left: (u32, u32, u32), right: (u32, u32, u32)) -> Ordering {
    left.0
        .cmp(&right.0)
        .then_with(|| left.1.cmp(&right.1))
        .then_with(|| left.2.cmp(&right.2))
}

fn parse_style_change(
    property: &str,
    value: &str,
) -> std::result::Result<Vec<StyleChange>, String> {
    let lower = value.trim().to_ascii_lowercase();
    let changes = match property {
        "font-family" => {
            let family = value
                .split(',')
                .next()
                .unwrap_or_default()
                .trim()
                .trim_matches(['\'', '"']);
            if family.is_empty() {
                return Err("unsupported empty font-family value".to_string());
            }
            vec![StyleChange::Font(family.to_string())]
        }
        "font-size" => vec![StyleChange::Size(parse_points(value)?)],
        "font-weight" => vec![StyleChange::Bold(match lower.as_str() {
            "normal" | "400" | "500" => false,
            "bold" | "bolder" | "600" | "700" | "800" | "900" => true,
            _ => return Err(format!("unsupported font-weight value `{value}`")),
        })],
        "font-style" => vec![StyleChange::Italic(match lower.as_str() {
            "normal" => false,
            "italic" | "oblique" => true,
            _ => return Err(format!("unsupported font-style value `{value}`")),
        })],
        "text-decoration" | "text-decoration-line" => {
            if lower == "none" {
                vec![StyleChange::Underline(false), StyleChange::Strike(false)]
            } else {
                let mut parsed = Vec::new();
                for token in lower.split_whitespace() {
                    match token {
                        "underline" => parsed.push(StyleChange::Underline(true)),
                        "line-through" => parsed.push(StyleChange::Strike(true)),
                        _ => {
                            return Err(format!("unsupported text-decoration value `{value}`"));
                        }
                    }
                }
                parsed
            }
        }
        "color" => vec![StyleChange::Color(parse_color(value)?)],
        "background" | "background-color" => {
            vec![StyleChange::Background(if lower == "transparent" {
                None
            } else {
                Some(parse_color(value)?)
            })]
        }
        "text-align" => vec![StyleChange::Alignment(match lower.as_str() {
            "left" | "start" => Alignment::Left,
            "center" => Alignment::Center,
            "right" | "end" => Alignment::Right,
            "justify" => Alignment::Justify,
            _ => return Err(format!("unsupported text-align value `{value}`")),
        })],
        "margin-top" => vec![StyleChange::SpaceBefore(parse_length(value)?)],
        "margin-bottom" => vec![StyleChange::SpaceAfter(parse_length(value)?)],
        "margin-left" => vec![StyleChange::IndentLeft(parse_length(value)?)],
        "margin-right" => vec![StyleChange::IndentRight(parse_length(value)?)],
        "text-indent" => vec![StyleChange::FirstLineIndent(parse_signed_length(value)?)],
        _ => return Err(format!("unsupported CSS property `{property}`")),
    };
    Ok(changes)
}

fn parse_points(value: &str) -> std::result::Result<f64, String> {
    let value = value.trim().to_ascii_lowercase();
    let points = if let Some(number) = value.strip_suffix("pt") {
        number.trim().parse::<f64>()
    } else if let Some(number) = value.strip_suffix("px") {
        number.trim().parse::<f64>().map(|number| number * 0.75)
    } else {
        return Err(format!("unsupported CSS length `{value}`"));
    }
    .map_err(|_| format!("invalid CSS length `{value}`"))?;
    if !points.is_finite() || points <= 0.0 || points > 1000.0 {
        return Err(format!(
            "CSS length is outside the supported range `{value}`"
        ));
    }
    Ok(points)
}

fn parse_length(value: &str) -> std::result::Result<Length, String> {
    if value.trim() == "0" {
        return Ok(Length::pt(0.0));
    }
    parse_points(value).map(Length::pt)
}

fn parse_signed_length(value: &str) -> std::result::Result<Length, String> {
    let value = value.trim().to_ascii_lowercase();
    if value == "0" {
        return Ok(Length::pt(0.0));
    }
    let points = if let Some(number) = value.strip_suffix("pt") {
        number.trim().parse::<f64>()
    } else if let Some(number) = value.strip_suffix("px") {
        number.trim().parse::<f64>().map(|number| number * 0.75)
    } else {
        return Err(format!("unsupported CSS length `{value}`"));
    }
    .map_err(|_| format!("invalid CSS length `{value}`"))?;
    if !points.is_finite() || points.abs() > 1000.0 {
        return Err(format!(
            "CSS length is outside the supported range `{value}`"
        ));
    }
    Ok(Length::pt(points))
}

fn parse_color(value: &str) -> std::result::Result<String, String> {
    let value = value.trim().to_ascii_lowercase();
    let color = match value.as_str() {
        "black" => "000000".to_string(),
        "white" => "FFFFFF".to_string(),
        "red" => "FF0000".to_string(),
        "green" => "008000".to_string(),
        "blue" => "0000FF".to_string(),
        "yellow" => "FFFF00".to_string(),
        "gray" | "grey" => "808080".to_string(),
        value if value.len() == 7 && value.starts_with('#') => value[1..].to_ascii_uppercase(),
        value if value.len() == 4 && value.starts_with('#') => {
            let mut expanded = String::with_capacity(6);
            for character in value[1..].chars() {
                expanded.push(character.to_ascii_uppercase());
                expanded.push(character.to_ascii_uppercase());
            }
            expanded
        }
        _ => return Err(format!("unsupported CSS color `{value}`")),
    };
    if !color.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return Err(format!("invalid CSS color `{value}`"));
    }
    Ok(color)
}

#[cfg(test)]
mod tests {
    use std::io::Cursor;

    use base64::Engine as _;
    use oxml_opc::OpcPackage;
    use rdocx_oxml::document::BodyContent;

    use super::{
        Document, Length, Limits, MhtmlLimits, from_html_with_limits, from_mhtml_with_limits,
        parse_mhtml, preflight_markup, read_bounded, to_mhtml_with_limits,
    };

    /// The text and the underline, strike, bold and italic of each run.
    fn fx231_word_marked_runs(
        paragraph: &crate::paragraph::ParagraphRef<'_>,
    ) -> Vec<(String, [bool; 4])> {
        (0..paragraph.run_count())
            .map(|index| paragraph.run(index).unwrap())
            .map(|run| {
                (
                    run.text(),
                    [
                        run.is_underline(),
                        run.is_strike(),
                        run.is_bold(),
                        run.is_italic(),
                    ],
                )
            })
            .collect()
    }

    #[test]
    fn fx231_word_advanced_whitespace_equivalence() {
        const PLAIN: [bool; 4] = [false; 4];
        const U: [bool; 4] = [true, false, false, false];
        const S: [bool; 4] = [false, true, false, false];
        const B: [bool; 4] = [false, false, true, false];
        const I: [bool; 4] = [false, false, false, true];
        let html = Document::from_html(
            "<p>see <u>u</u> then <s>s</s> then <b>b</b> then <i>i</i> end</p>\
         <p><u>u</u> <s>s</s></p><p><u>u </u>tail</p>",
        )
        .unwrap()
        .document;
        // "see " then each span followed by " then ", the last by " end".
        let expected = |spans: &[(&str, [bool; 4])]| {
            let mut runs = vec![("see ".to_owned(), PLAIN)];
            for (index, (text, mark)) in spans.iter().enumerate() {
                runs.push(((*text).to_owned(), *mark));
                let after = if index + 1 == spans.len() {
                    " end"
                } else {
                    " then "
                };
                runs.push((after.to_owned(), PLAIN));
            }
            runs
        };
        let all = [("u", U), ("s", S), ("b", B), ("i", I)];
        assert_eq!(
            fx231_word_marked_runs(&html.paragraph(0).unwrap()),
            expected(&all)
        );
        // A space between two spans is in neither of them.
        assert_eq!(
            fx231_word_marked_runs(&html.paragraph(1).unwrap()),
            [
                ("u".to_owned(), U),
                (" ".to_owned(), PLAIN),
                ("s".to_owned(), S)
            ]
        );
        // A space written inside the span stays in it, as a browser draws it.
        assert_eq!(
            fx231_word_marked_runs(&html.paragraph(2).unwrap()),
            [("u ".to_owned(), U), ("tail".to_owned(), PLAIN)]
        );

        let root = std::env::temp_dir().join(format!("fx231-word-{}", std::process::id()));
        std::fs::create_dir_all(&root).unwrap();
        let path = root.join("source.html");
        std::fs::write(
            &path,
            "<p>see <u>u</u> then <s>s</s> then <b>b</b> then <i>i</i> end</p>",
        )
        .unwrap();
        let mut opened = Document::open_html(&path).unwrap().document;
        let reopened = Document::from_bytes(&opened.to_bytes().unwrap()).unwrap();
        assert_eq!(
            fx231_word_marked_runs(&reopened.paragraph(0).unwrap()),
            expected(&all)
        );
        std::fs::remove_dir_all(root).unwrap();
        let mut html = html;
        let reopened = Document::from_bytes(&html.to_bytes().unwrap()).unwrap();
        for index in 0..html.paragraph_count() {
            assert_eq!(
                fx231_word_marked_runs(&reopened.paragraph(index).unwrap()),
                fx231_word_marked_runs(&html.paragraph(index).unwrap())
            );
        }

        // Markdown has no underline: strike, bold and italic.
        let md = Document::from_markdown("see ~~s~~ then **b** then *i* end\n")
            .unwrap()
            .document;
        assert_eq!(
            fx231_word_marked_runs(&md.paragraph(0).unwrap()),
            expected(&all[1..])
        );

        // Inserting a fragment builds its paragraphs the same way.
        let mut inserted = Document::new();
        let body = inserted
            .stories()
            .unwrap()
            .into_iter()
            .find(|story| story.kind() == crate::StoryKind::Body)
            .unwrap();
        inserted
            .insert_html_fragment(
                &crate::ContentLocation::end(body),
                "<p>see <u>u</u> then <s>s</s> then <b>b</b> then <i>i</i> end</p>",
                &[],
            )
            .unwrap();
        let inserted = Document::from_bytes(&inserted.to_bytes().unwrap()).unwrap();
        let last = inserted.paragraph_count() - 1;
        assert_eq!(
            fx231_word_marked_runs(&inserted.paragraph(last).unwrap()),
            expected(&all)
        );
    }

    #[test]
    fn fx231_word_repeated_edges_nbsp_and_links_are_equivalent() {
        const PLAIN: [bool; 4] = [false; 4];
        const U: [bool; 4] = [true, false, false, false];
        const S: [bool; 4] = [false, true, false, false];
        const B: [bool; 4] = [false, false, true, false];
        const I: [bool; 4] = [false, false, false, true];
        const BI: [bool; 4] = [false, false, true, true];
        let runs = |items: &[(&str, [bool; 4])]| {
            items
                .iter()
                .map(|(text, marks)| ((*text).to_owned(), *marks))
                .collect::<Vec<_>>()
        };
        let html = |source: &str| Document::from_html(source).unwrap().document;
        // The one space stays where the whitespace sequence starts: at the end
        // of the span when it ends one, before the span when it starts one.
        for (tag, mark) in [("u", U), ("s", S), ("b", B), ("i", I)] {
            let ending = html(&format!("<p>a <{tag}>b </{tag}> c</p>"));
            assert_eq!(
                fx231_word_marked_runs(&ending.paragraph(0).unwrap()),
                runs(&[("a ", PLAIN), ("b ", mark), ("c", PLAIN)]),
                "{tag} end"
            );
            let starting = html(&format!("<p>and <{tag}> that</{tag}></p>"));
            assert_eq!(
                fx231_word_marked_runs(&starting.paragraph(0).unwrap()),
                runs(&[("and ", PLAIN), ("that", mark)]),
                "{tag} start"
            );
        }
        // Nested spans, and a span that holds only whitespace between two edges.
        let nested = html(
            "<p>a <b>b <i> c</i></b> d</p><p><b>x <i>y </i></b> z</p><p>a <s>b </s><i> </i> c</p>",
        );
        assert_eq!(
            fx231_word_marked_runs(&nested.paragraph(0).unwrap()),
            runs(&[("a ", PLAIN), ("b ", B), ("c", BI), (" d", PLAIN)])
        );
        assert_eq!(
            fx231_word_marked_runs(&nested.paragraph(1).unwrap()),
            runs(&[("x ", B), ("y ", BI), ("z", PLAIN)])
        );
        assert_eq!(
            fx231_word_marked_runs(&nested.paragraph(2).unwrap()),
            runs(&[("a ", PLAIN), ("b ", S), ("c", PLAIN)])
        );
        // A non-breaking space is not collapsible and stays as written.
        let kept = html("<p>a&nbsp;&nbsp;<s>b</s></p><p>a <s>b </s>&nbsp;c</p>");
        assert_eq!(kept.paragraph(0).unwrap().text(), "a\u{a0}\u{a0}b");
        assert_eq!(kept.paragraph(1).unwrap().text(), "a b \u{a0}c");

        // A link keeps the space it ends with, and the text after it none.
        let link_texts = |document: &Document| {
            document
                .links()
                .into_iter()
                .map(|link| link.text)
                .collect::<Vec<_>>()
        };
        let linked = html(
            "<p>a <a href=\"https://example.test/\">b </a> c</p>\
         <p>a <a href=\"https://example.test/\"> b</a></p>",
        );
        assert_eq!(
            fx231_word_marked_runs(&linked.paragraph(0).unwrap()),
            runs(&[("a ", PLAIN), ("b ", PLAIN), ("c", PLAIN)])
        );
        assert_eq!(
            fx231_word_marked_runs(&linked.paragraph(1).unwrap()),
            runs(&[("a ", PLAIN), ("b", PLAIN)])
        );
        assert_eq!(link_texts(&linked), ["b ", "b"]);

        // Markdown writes the span edges with links, also inside a strong span.
        let md = Document::from_markdown(
            "a [b ](https://example.test/) c\n\na [ b](https://example.test/)\n\n\
         **b [c ](https://example.test/)** d\n",
        )
        .unwrap()
        .document;
        assert_eq!(
            fx231_word_marked_runs(&md.paragraph(0).unwrap()),
            runs(&[("a ", PLAIN), ("b ", PLAIN), ("c", PLAIN)])
        );
        assert_eq!(
            fx231_word_marked_runs(&md.paragraph(1).unwrap()),
            runs(&[("a ", PLAIN), ("b", PLAIN)])
        );
        assert_eq!(
            fx231_word_marked_runs(&md.paragraph(2).unwrap()),
            runs(&[("b ", B), ("c ", B), ("d", PLAIN)])
        );
        assert_eq!(link_texts(&md), ["b ", "b", "c "]);

        // Inserting a fragment builds its paragraphs the same way.
        let mut inserted = Document::new();
        let body = inserted
            .stories()
            .unwrap()
            .into_iter()
            .find(|story| story.kind() == crate::StoryKind::Body)
            .unwrap();
        inserted
            .insert_html_fragment(
                &crate::ContentLocation::end(body),
                "<p>a <s>b </s> c</p>",
                &[],
            )
            .unwrap();
        let inserted = Document::from_bytes(&inserted.to_bytes().unwrap()).unwrap();
        let last = inserted.paragraph_count() - 1;
        assert_eq!(
            fx231_word_marked_runs(&inserted.paragraph(last).unwrap()),
            runs(&[("a ", PLAIN), ("b ", S), ("c", PLAIN)])
        );
    }

    #[test]
    fn fx217_markdown_all_supported_constructs() {
        let markdown = "# Heading\n\n**bold** *em* ~~strike~~ `a  b`\n\n> First\n>\n> Second\n\n3. Third\n4. Fourth\n\n| Head | Right |\n| --- | ---: |\n| A | B |\n\n[link](https://example.com)\n";
        let mut document = Document::from_markdown(markdown).unwrap().document;
        let bytes = document.to_bytes().unwrap();
        let reopened = Document::from_bytes(&bytes).unwrap();
        assert_eq!(reopened.paragraph(0).unwrap().text(), "Heading");
        assert_eq!(reopened.paragraph(1).unwrap().text(), "bold em strike a  b");
        assert_eq!(reopened.paragraph(2).unwrap().text(), "First");
        assert_eq!(reopened.paragraph(3).unwrap().text(), "Second");
        assert_eq!(reopened.table(0).unwrap().cell(1, 0).unwrap().text(), "A");
        let paragraph = reopened.paragraph(1).unwrap();
        let runs = paragraph.runs().collect::<Vec<_>>();
        assert!(runs.iter().any(|run| run.bold_value() == Some(true)));
        assert!(runs.iter().any(|run| run.italic_value() == Some(true)));
        assert!(runs.iter().any(|run| run.strike_value() == Some(true)));
        assert!(runs.iter().any(|run| run.style_id() == Some("HTMLCode")));
    }

    #[test]
    fn fx217_import_diagnostics_and_limits_are_atomic() {
        let projection = super::project_markdown(
            "<b>raw</b>\n\n<script>ignored</script>\n",
            Limits::default(),
        )
        .unwrap();
        assert!(!projection.diagnostics.is_empty());
        assert!(
            projection
                .diagnostics
                .iter()
                .all(|item| item.property.is_none() && item.location.starts_with("bytes "))
        );
        for limits in [
            Limits {
                input_bytes: 2,
                ..Limits::default()
            },
            Limits {
                nodes: 1,
                ..Limits::default()
            },
            Limits {
                retained_text: 1,
                ..Limits::default()
            },
            Limits {
                depth: 0,
                ..Limits::default()
            },
        ] {
            assert!(super::project_markdown("ordinary", limits).is_err());
        }
        assert!(
            super::project_markdown(
                "<b>raw</b>",
                Limits {
                    diagnostics: 0,
                    ..Limits::default()
                }
            )
            .is_err()
        );
        let mut template = Document::new();
        template.add_paragraph("Template content");
        let before = template.to_bytes().unwrap();
        assert!(
            Document::from_markdown_with_options(
                "# Heading",
                Some(&template),
                &[super::HtmlImageResource {
                    source: "x",
                    bytes: b"invalid",
                    filename: "x.png",
                }]
            )
            .is_err()
        );
        assert_eq!(before, template.to_bytes().unwrap());
    }

    #[test]
    fn fx217_template_body_is_replaced_without_mutating_template() {
        let mut template = Document::new();
        template.add_paragraph("Template sample content");
        let before = template.to_bytes().unwrap();
        let mut imported =
            Document::from_markdown_with_options("# Replacement\n\nBody", Some(&template), &[])
                .unwrap()
                .document;
        let bytes = imported.to_bytes().unwrap();
        let reopened = Document::from_bytes(&bytes).unwrap();
        assert_eq!(reopened.text(), "Replacement\nBody\n");
        assert_eq!(before, template.to_bytes().unwrap());
        assert_eq!(
            reopened.document.body.sect_pr,
            Document::from_bytes(&before).unwrap().document.body.sect_pr
        );
    }

    #[test]
    fn fx217_markdown_diagnostics_use_original_source_order_and_byte_spans() {
        let parsed = Document::from_markdown("![alt](missing.png)\n\n<br>\n").unwrap();
        assert_eq!(parsed.diagnostics.len(), 2);
        assert!(
            parsed.diagnostics[0]
                .message
                .contains("unresolved HTML image")
        );
        assert_eq!(parsed.diagnostics[0].location, "bytes 0..19");
        assert!(parsed.diagnostics[1].message.contains("raw Markdown HTML"));
        assert_eq!(parsed.diagnostics[1].location, "bytes 21..26");
    }

    #[test]
    fn fx217_markdown_image_alt_does_not_create_phantom_resource_locations() {
        let parsed = Document::from_markdown(
            "![outer [link](https://example.com) ![inner](inner.png)](outer.png)",
        )
        .unwrap();
        assert_eq!(parsed.diagnostics.len(), 1);
        assert!(
            parsed.diagnostics[0]
                .message
                .contains("unresolved HTML image")
        );
        assert!(parsed.diagnostics[0].location.starts_with("bytes 0.."));
    }

    fn review_roundtrip(markdown: &str) -> String {
        let mut document = Document::from_markdown(markdown).unwrap().document;
        Document::from_bytes(&document.to_bytes().unwrap())
            .unwrap()
            .to_markdown()
    }

    #[test]
    fn fx217_review_escaped_block_prefixes_and_continuations() {
        for source in [
            "\\# literal heading\n\n",
            "\\> literal quote\n\n",
            "1\\. literal item\n\n",
            "text  \n\\# continued\n\n",
        ] {
            assert_eq!(review_roundtrip(source), source);
        }
    }

    #[test]
    fn fx217_review_headings_stay_in_list_items() {
        assert_eq!(review_roundtrip("- # Heading\n"), "- # Heading\n");
    }

    #[test]
    fn fx217_review_quotes_keep_outer_list_order() {
        assert_eq!(
            review_roundtrip("> - First\n> - Second\n"),
            "> - First\n> - Second\n"
        );
    }

    #[test]
    fn fx217_review_quote_rule_and_table_owners_survive() {
        let source = "> ---\n\n> | Head |\n> | --- |\n> | Cell |\n\n";
        assert_eq!(review_roundtrip(source), source);
    }

    #[test]
    fn fx217_review_nonbreaking_space_only_content_survives() {
        for source in ["<ul><li>&nbsp;</li></ul>", "&nbsp;"] {
            let document = Document::from_html(source).unwrap().document;
            assert_eq!(document.paragraph(0).unwrap().text(), "\u{a0}");
        }
        let document = Document::from_markdown("- \u{a0}\n").unwrap().document;
        assert_eq!(document.paragraph(0).unwrap().text(), "\u{a0}");
        assert_eq!(review_roundtrip("- \u{a0}\n"), "- \u{a0}\n");
    }

    #[test]
    fn fx217_review_quote_keeps_list_and_following_paragraph_in_one_owner() {
        for source in [
            "> - First\n> - Second\n>\n> Following\n",
            "> - First\n>\n> Following\n>\n> ---\n>\n> ```\n> code\n> ```\n>\n> | Head |\n> | --- |\n> | Cell |\n",
            "> - First\n\n> - Second\n",
            "> - First\n>\n> | Head |\n> | --- |\n> | Cell |\n",
            "- First\n\n  > - Inner\n  >\n  > Following\n  >\n  > Last\n",
        ] {
            let exported = review_roundtrip(source);
            let options = pulldown_cmark::Options::ENABLE_TABLES;
            let original: Vec<_> = pulldown_cmark::Parser::new_ext(source, options).collect();
            let actual: Vec<_> = pulldown_cmark::Parser::new_ext(&exported, options).collect();
            assert_eq!(actual, original, "{source:?} exported as {exported:?}");
            assert_eq!(review_roundtrip(&exported), exported);
        }
    }

    #[test]
    fn fx217_review_nested_quote_owners_survive_paragraph_joins() {
        for source in [
            "> - First\n>\n> > Inner A\n> >\n> > Inner B\n",
            "> - First\n>\n> > Inner A\n>\n> > Inner B\n",
            "> - First\n>\n> > > Deep A\n> > >\n> > > Deep B\n> >\n> > Outer Inner\n>\n> Last\n",
        ] {
            let exported = review_roundtrip(source);
            let original: Vec<_> = pulldown_cmark::Parser::new(source).collect();
            let actual: Vec<_> = pulldown_cmark::Parser::new(&exported).collect();
            assert_eq!(actual, original, "{source:?} exported as {exported:?}");
            assert_eq!(review_roundtrip(&exported), exported);
        }
    }

    #[test]
    fn fx217_review_nested_container_order_and_collision_roundtrip() {
        for source in [
            "> - # Heading\n",
            "> - First\n>\n>   > Inner\n\n",
            "> \\# literal\n\n",
            "> - First\n>\n>   ```\n>   code\n>   ```\n\n",
            "- First\n\n  ## Next\n\n",
            "> - First\n>\n>   | Head |\n>   | --- |\n>   | Cell |\n\n",
        ] {
            let exported = review_roundtrip(source);
            let options = pulldown_cmark::Options::ENABLE_TABLES
                | pulldown_cmark::Options::ENABLE_STRIKETHROUGH;
            let original: Vec<_> = pulldown_cmark::Parser::new_ext(source, options).collect();
            let actual: Vec<_> = pulldown_cmark::Parser::new_ext(&exported, options).collect();
            assert_eq!(actual, original, "{source:?} exported as {exported:?}");
            assert_eq!(review_roundtrip(&exported), exported);
        }
        let mut template = Document::new();
        template
            .add_style(
                crate::StyleBuilder::paragraph(
                    "RDOCXHTMLContainersqlOwner1Paragraph0",
                    "RDOCX HTML containers ql",
                )
                .bold(true),
            )
            .unwrap();
        let before = template.to_bytes().unwrap();
        let mut imported =
            Document::from_markdown_with_options("> - First\n", Some(&template), &[])
                .unwrap()
                .document;
        assert_eq!(template.to_bytes().unwrap(), before);
        let reopened = Document::from_bytes(&imported.to_bytes().unwrap()).unwrap();
        assert_eq!(reopened.to_markdown(), "> - First\n");
        assert!(
            reopened
                .styles
                .get_by_id("RDOCXHTMLContainersqlOwner1Paragraph1")
                .is_some()
        );
        assert_eq!(
            reopened
                .styles
                .get_by_id("RDOCXHTMLContainersqlOwner1Paragraph0")
                .unwrap()
                .rpr
                .as_ref()
                .unwrap()
                .bold,
            Some(true)
        );
    }

    #[test]
    fn fx217_canonical_markdown_keeps_heading_and_code_inside_quotes() {
        const CANONICAL: &str = "> # Heading\n\n> ```\n> a  b\n> ```\n\n";
        let mut imported = Document::from_markdown(CANONICAL).unwrap().document;
        let bytes = imported.to_bytes().unwrap();
        let reopened = Document::from_bytes(&bytes).unwrap();
        assert_eq!(reopened.to_markdown(), CANONICAL);
    }

    #[test]
    fn fx217_canonical_markdown_keeps_code_and_quotes_inside_list_items() {
        const CANONICAL: &str = "- First\n\n  ```\n  a  b\n  ```\n\n  > quoted\n\n- Last\n";
        let mut imported = Document::from_markdown(CANONICAL).unwrap().document;
        let bytes = imported.to_bytes().unwrap();
        let reopened = Document::from_bytes(&bytes).unwrap();
        assert_eq!(reopened.to_markdown(), CANONICAL);
    }

    #[test]
    fn fx217_canonical_markdown_keeps_ordered_starts_and_table_alignment() {
        const CANONICAL: &str =
            "3. Third\n1. Fourth\n\n| Head | Right |\n| --- | ---: |\n| A\\|B | `x y` |\n\n";
        let mut imported = Document::from_markdown(CANONICAL).unwrap().document;
        let bytes = imported.to_bytes().unwrap();
        let reopened = Document::from_bytes(&bytes).unwrap();
        assert_eq!(reopened.to_markdown(), CANONICAL);
    }

    #[test]
    fn fx217_template_styles_headers_footers_and_opaque_parts() {
        let mut template = Document::new();
        template.add_paragraph("Discarded sample body");
        let list = template.add_list_definition(&[crate::ListLevel::decimal().start(7)]);
        template
            .add_paragraph("Old numbered sample")
            .set_numbering(list, 0);
        template.set_header_image(
            &one_pixel_png(),
            "header.png",
            Length::inches(0.2),
            Length::inches(0.2),
        );
        template.set_footer("Preserved footer");
        let header = template
            .stories()
            .unwrap()
            .into_iter()
            .find(|story| story.kind() == crate::StoryKind::Header)
            .unwrap();
        template
            .add_hyperlink_to_story(&header, "Header link", "https://example.com")
            .unwrap();
        let mut header_xml = String::from_utf8(
            template
                .package
                .get_part(header.part_name())
                .unwrap()
                .to_vec(),
        )
        .unwrap();
        header_xml = header_xml.replace("</w:hdr>", "<w:p><w:fldSimple w:instr=\"PAGE\"><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p></w:hdr>");
        template
            .package
            .set_part(header.part_name(), header_xml.into_bytes());
        template
            .package
            .set_part("/custom/opaque.bin", vec![0, 1, 2, 255]);
        template
            .package
            .content_types
            .add_override("/custom/opaque.bin", "application/octet-stream");
        let before = template.to_bytes().unwrap();
        let source = OpcPackage::from_reader(Cursor::new(&before)).unwrap();
        let loaded = Document::from_bytes(&before).unwrap();
        let old_numbering = loaded.numbering.clone().unwrap();
        let mut imported = Document::from_markdown_with_options(
            "# Replacement\n\n3. New item\n4. Next item",
            Some(&loaded),
            &[],
        )
        .unwrap()
        .document;
        let bytes = imported.to_bytes().unwrap();
        let output = OpcPackage::from_reader(Cursor::new(&bytes)).unwrap();
        assert_eq!(template.to_bytes().unwrap(), before);
        let reopened = Document::from_bytes(&bytes).unwrap();
        assert!(!reopened.text().contains("Discarded"));
        assert!(!reopened.text().contains("Old numbered"));
        assert_eq!(reopened.document.body.sect_pr, loaded.document.body.sect_pr);
        for part in source.parts.keys() {
            if part.as_str() != loaded.doc_part_name
                && !part.ends_with("styles.xml")
                && !part.ends_with("numbering.xml")
            {
                assert_eq!(
                    output.get_part(part),
                    source.get_part(part),
                    "template part {part}"
                );
            }
        }
        let numbering = reopened.numbering.as_ref().unwrap();
        for definition in &old_numbering.abstract_nums {
            assert!(numbering.abstract_nums.contains(definition));
        }
        for instance in &old_numbering.nums {
            assert!(numbering.nums.contains(instance));
        }
        assert!(numbering.nums.len() > old_numbering.nums.len());
    }

    #[test]
    fn fx217_supported_markdown_docx_markdown_canonical_round_trip() {
        const CANONICAL: &str = "# Heading\n\nBody **bold** *em* ~~strike~~ `a  b`.\n\n> First\n\n> Second\n\n```\nx  y\n```\n\n---\n\n";
        let mut imported = Document::from_markdown(CANONICAL).unwrap().document;
        let bytes = imported.to_bytes().unwrap();
        let reopened = Document::from_bytes(&bytes).unwrap();
        assert_eq!(reopened.to_markdown(), CANONICAL);
    }

    #[test]
    fn fx217_markdown_images_keep_bytes_and_per_occurrence_alternate_text() {
        let png = one_pixel_png();
        let images = [super::HtmlImageResource {
            source: "pixel.png",
            bytes: &png,
            filename: "pixel.png",
        }];
        let mut imported = Document::from_markdown_with_options(
            "![first](pixel.png) ![second](pixel.png)",
            None,
            &images,
        )
        .unwrap()
        .document;
        let bytes = imported.to_bytes().unwrap();
        let reopened = Document::from_bytes(&bytes).unwrap();
        let markdown = reopened.to_markdown();
        assert!(markdown.contains("![first](data:image/png;base64,"));
        assert!(markdown.contains("![second](data:image/png;base64,"));
        let again = Document::from_markdown(&markdown).unwrap().document;
        assert_eq!(again.document.body.content.len(), 1);
        assert_eq!(again.to_markdown(), markdown);
    }

    #[test]
    fn fx217_multiblock_list_items_retain_ordered_blocks_and_nested_tables() {
        let markdown =
            "- First\n\n  Second\n\n  | Head |\n  | --- |\n  | Cell |\n\n  - Nested\n\n- Last\n";
        let imported = Document::from_markdown(markdown).unwrap().document;
        let text = imported
            .paragraphs()
            .into_iter()
            .map(|paragraph| paragraph.text())
            .collect::<Vec<_>>();
        assert_eq!(text, ["First", "Second", "Nested", "Last"]);
        assert_eq!(
            imported.table(0).unwrap().cell(0, 0).unwrap().text(),
            "Head"
        );
        assert_eq!(
            imported.table(0).unwrap().cell(1, 0).unwrap().text(),
            "Cell"
        );
        assert!(matches!(
            imported.document.body.content[2],
            BodyContent::Table(_)
        ));
        assert_eq!(imported.tables().len(), 1);
        let BodyContent::Paragraph(first) = &imported.document.body.content[0] else {
            panic!("first item")
        };
        let BodyContent::Paragraph(second) = &imported.document.body.content[1] else {
            panic!("continuation")
        };
        assert!(
            first
                .properties
                .as_ref()
                .unwrap()
                .num_id
                .is_some_and(|id| id > 0)
        );
        assert_eq!(second.properties.as_ref().unwrap().num_id, Some(0));
    }

    #[test]
    fn fx217_named_template_styles_keep_their_existing_identities() {
        use crate::StyleBuilder;
        let mut template = Document::new();
        std::sync::Arc::make_mut(&mut template.styles)
            .styles
            .retain(|style| style.style_id != "Quote" && style.style_id != "Heading1");
        template
            .add_style(
                StyleBuilder::paragraph("CustomHeading", "Heading 1")
                    .font("Gelasio")
                    .size(18.0),
            )
            .unwrap();
        template
            .add_style(StyleBuilder::paragraph("CustomQuote", "Quote").font("Gelasio"))
            .unwrap();
        template
            .add_style(StyleBuilder::character("CustomCode", "HTML Code").font("Selawik"))
            .unwrap();
        let imported = Document::from_markdown_with_options(
            "# Heading\n\n> Quote\n\n`code`",
            Some(&template),
            &[],
        )
        .unwrap()
        .document;
        assert_eq!(imported.document.body.content.len(), 3);
        for (index, expected) in [(0, "CustomHeading"), (1, "CustomQuote")] {
            let BodyContent::Paragraph(paragraph) = &imported.document.body.content[index] else {
                panic!("paragraph expected")
            };
            assert_eq!(
                paragraph.properties.as_ref().unwrap().style_id.as_deref(),
                Some(expected)
            );
        }
        assert_eq!(
            imported.paragraph(2).unwrap().run(0).unwrap().style_id(),
            Some("CustomCode")
        );
    }

    #[test]
    fn fx217_html_nested_block_containers_keep_distinct_paragraphs() {
        let mut imported = Document::from_html(
            "<blockquote><p>First</p><p>Second</p><blockquote><p>Nested</p></blockquote></blockquote>",
        ).unwrap().document;
        let bytes = imported.to_bytes().unwrap();
        let reopened = Document::from_bytes(&bytes).unwrap();
        let texts = reopened
            .paragraphs()
            .into_iter()
            .map(|paragraph| paragraph.text())
            .collect::<Vec<_>>();
        assert_eq!(texts, ["First", "Second", "Nested"]);
    }

    #[test]
    fn mark_highlight_reaches_formatted_children_without_leaking_to_siblings() {
        let parsed = Document::from_html(
            "<p>Before <mark>key <strong>fact</strong> <em>here</em></mark> after</p>",
        )
        .unwrap();
        let paragraph = parsed.document.paragraph(0).unwrap();
        let runs = paragraph.runs().collect::<Vec<_>>();
        assert_eq!(
            runs.iter().map(|run| run.text()).collect::<Vec<_>>(),
            ["Before ", "key ", "fact", " ", "here", " after"]
        );
        assert_eq!(runs[0].highlight(), None);
        for run in &runs[1..5] {
            assert_eq!(run.highlight().as_deref(), Some("FFFF00"));
        }
        assert_eq!(runs[5].highlight(), None);
        assert_eq!(runs[2].bold_value(), Some(true));
        assert_eq!(runs[3].bold_value(), None);
        assert_eq!(runs[3].italic_value(), None);
        assert_eq!(runs[4].italic_value(), Some(true));
    }

    #[test]
    fn semantic_marks_preserve_css_precedence_without_inheriting_ordinary_backgrounds() {
        let parsed = Document::from_html(concat!(
            r#"<p><mark style="background-color:#112233">A<strong>B</strong>"#,
            r#"<mark style="background-color:#445566"><em>C</em></mark>"#,
            r#"<span style="background-color:#0000FF">D<b>E</b></span>F</mark>"#,
            r#"<span style="background-color:#ABCDEF">G<em>H</em></span>I</p>"#,
        ))
        .unwrap();
        let paragraph = parsed.document.paragraph(0).unwrap();
        let actual = paragraph
            .runs()
            .map(|run| (run.text(), run.highlight()))
            .collect::<Vec<_>>();
        assert_eq!(
            actual,
            vec![
                ("A".into(), Some("112233".into())),
                ("B".into(), Some("112233".into())),
                ("C".into(), Some("445566".into())),
                ("D".into(), Some("0000FF".into())),
                ("E".into(), Some("112233".into())),
                ("F".into(), Some("112233".into())),
                ("G".into(), Some("ABCDEF".into())),
                ("H".into(), None),
                ("I".into(), None),
            ]
        );
    }

    #[test]
    fn semantic_mark_highlights_reopen_with_nested_run_formatting() {
        let mut imported = Document::from_html(
            "<p>before<mark>key<strong>bold</strong><em>italic</em></mark>after</p>",
        )
        .unwrap()
        .document;
        let formats = |document: &Document| {
            document
                .paragraph(0)
                .unwrap()
                .runs()
                .map(|run| {
                    (
                        run.text(),
                        run.highlight(),
                        run.bold_value(),
                        run.italic_value(),
                    )
                })
                .collect::<Vec<_>>()
        };
        let expected = formats(&imported);
        assert_eq!(expected[2].1.as_deref(), Some("FFFF00"));
        assert_eq!(expected[2].2, Some(true));
        assert_eq!(expected[3].1.as_deref(), Some("FFFF00"));
        assert_eq!(expected[3].3, Some(true));
        let mut package =
            OpcPackage::from_reader(Cursor::new(imported.to_bytes().unwrap())).unwrap();
        let opaque =
            r#"<x:opaque xmlns:x="urn:mark-regression" x:keep="exact"><x:child /></x:opaque>"#;
        let document_xml =
            String::from_utf8(package.get_part("/word/document.xml").unwrap().to_vec())
                .unwrap()
                .replace("</w:body>", &format!("{opaque}</w:body>"));
        package.set_part("/word/document.xml", document_xml.into_bytes());
        let opaque_part = b"<producer xmlns='urn:mark-regression' keep='exact' />";
        package.set_part("/customXml/mark-control.xml", opaque_part.to_vec());
        let mut source = Cursor::new(Vec::new());
        package.write_to(&mut source).unwrap();
        let mut preserved = Document::from_bytes(source.get_ref()).unwrap();
        let saved = preserved.to_bytes().unwrap();
        let saved_package = OpcPackage::from_reader(Cursor::new(&saved)).unwrap();
        assert_eq!(
            saved_package.get_part("/customXml/mark-control.xml"),
            Some(opaque_part.as_slice())
        );
        assert!(
            String::from_utf8_lossy(saved_package.get_part("/word/document.xml").unwrap())
                .contains(opaque)
        );
        let reopened = Document::from_bytes(&saved).unwrap();
        assert_eq!(formats(&reopened), expected);
    }

    fn one_pixel_png() -> Vec<u8> {
        vec![
            0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48,
            0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00,
            0x00, 0x90, 0x77, 0x53, 0xde, 0x00, 0x00, 0x00, 0x0c, 0x49, 0x44, 0x41, 0x54, 0x08,
            0xd7, 0x63, 0xf8, 0xcf, 0xc0, 0x00, 0x00, 0x00, 0x02, 0x00, 0x01, 0xe2, 0x21, 0xbc,
            0x33, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
        ]
    }

    fn one_pixel_jpeg() -> Vec<u8> {
        vec![
            0xff, 0xd8, 0xff, 0xe0, 0x00, 0x10, 0x4a, 0x46, 0x49, 0x46, 0x00, 0x01, 0x01, 0x00,
            0x00, 0x01, 0x00, 0x01, 0x00, 0x00, 0xff, 0xdb, 0x00, 0x43, 0x00, 0x08, 0x06, 0x06,
            0x07, 0x06, 0x05, 0x08, 0x07, 0x07, 0x07, 0x09, 0x09, 0x08, 0x0a, 0x0c, 0x14, 0x0d,
            0x0c, 0x0b, 0x0b, 0x0c, 0x19, 0x12, 0x13, 0x0f, 0x14, 0x1d, 0x1a, 0x1f, 0x1e, 0x1d,
            0x1a, 0x1c, 0x1c, 0x20, 0x24, 0x2e, 0x27, 0x20, 0x22, 0x2c, 0x23, 0x1c, 0x1c, 0x28,
            0x37, 0x29, 0x2c, 0x30, 0x31, 0x34, 0x34, 0x34, 0x1f, 0x27, 0x39, 0x3d, 0x38, 0x32,
            0x3c, 0x2e, 0x33, 0x34, 0x32, 0xff, 0xc0, 0x00, 0x0b, 0x08, 0x00, 0x01, 0x00, 0x01,
            0x01, 0x01, 0x11, 0x00, 0xff, 0xc4, 0x00, 0x14, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x07, 0xff, 0xc4,
            0x00, 0x14, 0x10, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xff, 0xda, 0x00, 0x08, 0x01, 0x01, 0x00, 0x00,
            0x3f, 0x00, 0x7f, 0x7f, 0xff, 0xd9,
        ]
    }

    fn mhtml_fixture(html: &str) -> Vec<u8> {
        format!(
            "MIME-Version: 1.0\r\nContent-Type: multipart/related; boundary=fixture\r\n\r\n--fixture\r\nContent-Type: text/html; charset=utf-8\r\nContent-Location: https://example.test/index.html\r\n\r\n{html}\r\n--fixture--\r\n"
        )
        .into_bytes()
    }

    fn mhtml_image_fixture(content_type: &str, bytes: &[u8]) -> Vec<u8> {
        let encoded = super::BASE64.encode(bytes);
        format!(
            "MIME-Version: 1.0\r\nContent-Type: multipart/related; boundary=fixture; start=\"<root@rdocx>\"\r\n\r\n--fixture\r\nContent-Type: text/html; charset=utf-8\r\nContent-ID: <root@rdocx>\r\nContent-Location: https://example.test/index.html\r\n\r\n<p>before<img src='cid:image@rdocx' width='1' height='1'>after</p>\r\n--fixture\r\nContent-Type: {content_type}\r\nContent-ID: <image@rdocx>\r\nContent-Transfer-Encoding: base64\r\n\r\n{encoded}\r\n--fixture--\r\n"
        )
        .into_bytes()
    }

    #[test]
    fn mhtml_parser_rejects_ambiguous_unsafe_and_over_limit_resources_before_projection() {
        let malformed = [
            "MIME-Version: 1.0\r\nContent-Type: multipart/related\r\n\r\npartial".to_owned(),
            mhtml_fixture("<p>one</p>")
                .into_iter()
                .map(char::from)
                .collect::<String>()
                .replace("--fixture--", "--fixture\r\nContent-Type: text/html\r\n\r\n<p>two</p>\r\n--fixture--"),
            mhtml_fixture("<img src='https://outside.test/missing.png'>")
                .into_iter()
                .map(char::from)
                .collect(),
            mhtml_fixture("<a href='javascript:alert(1)'>unsafe</a>")
                .into_iter()
                .map(char::from)
                .collect(),
            mhtml_fixture("<img srcset='https://outside.test/a.png 1x'>")
                .into_iter()
                .map(char::from)
                .collect(),
            mhtml_fixture("<video src='https://outside.test/video.mp4'></video>")
                .into_iter()
                .map(char::from)
                .collect(),
            mhtml_fixture("<video src='cid:missing@rdocx'></video>")
                .into_iter()
                .map(char::from)
                .collect(),
            mhtml_fixture("<audio src='https://outside.test/audio.mp3'></audio>")
                .into_iter()
                .map(char::from)
                .collect(),
            mhtml_fixture("<audio src='cid:missing@rdocx'></audio>")
                .into_iter()
                .map(char::from)
                .collect(),
            mhtml_fixture("<track src='https://outside.test/captions.vtt'>")
                .into_iter()
                .map(char::from)
                .collect(),
            mhtml_fixture("<track src='cid:missing@rdocx'>")
                .into_iter()
                .map(char::from)
                .collect(),
            mhtml_fixture("<body background='https://outside.test/background.png'><p>x</p></body>")
                .into_iter()
                .map(char::from)
                .collect(),
            mhtml_fixture("<body background='cid:missing@rdocx'><p>x</p></body>")
                .into_iter()
                .map(char::from)
                .collect(),
            mhtml_fixture("<p style=\"background-image:url('https://outside.test/a.png')\">x</p>")
                .into_iter()
                .map(char::from)
                .collect(),
            mhtml_fixture("<style>@import \"https://outside.test/a.css\";</style><p>x</p>")
                .into_iter()
                .map(char::from)
                .collect(),
            mhtml_fixture("<style>@import 'cid:missing@rdocx';</style><p>x</p>")
                .into_iter()
                .map(char::from)
                .collect(),
            mhtml_fixture(
                r#"<p style="background-image:u\72l(https://outside.test/a.png)">x</p>"#,
            )
            .into_iter()
            .map(char::from)
            .collect(),
            mhtml_fixture(r#"<style>@\69mport 'cid:missing@rdocx';</style><p>x</p>"#)
                .into_iter()
                .map(char::from)
                .collect(),
            mhtml_fixture(
                r#"<p style="background-image:url(https://example.test/index.html\)outside)">x</p>"#,
            )
            .into_iter()
            .map(char::from)
            .collect(),
            mhtml_fixture(
                r#"<style>@import "https://example.test/index.html\"outside";</style><p>x</p>"#,
            )
            .into_iter()
            .map(char::from)
            .collect(),
            mhtml_fixture(
                r#"<style>@import url("https://example.test/index.html)outside");</style><p>x</p>"#,
            )
            .into_iter()
            .map(char::from)
            .collect(),
            mhtml_fixture("<p>x</p>")
                .into_iter()
                .map(char::from)
                .collect::<String>()
                .replace("boundary=fixture", "type=image/png; boundary=fixture"),
            mhtml_fixture("<p>x</p>")
                .into_iter()
                .map(char::from)
                .collect::<String>()
                .replace("charset=utf-8", "charset=iso-8859-1"),
            mhtml_fixture("<p>x</p>")
                .into_iter()
                .map(char::from)
                .collect::<String>()
                .replace(
                    "https://example.test/index.html",
                    "https://example.test/bad path.html",
                ),
            mhtml_fixture("<p>x</p>")
                .into_iter()
                .map(char::from)
                .collect::<String>()
                .replace("--fixture--\r\n", "--fixture--\r\n--fixture--\r\n"),
            "MIME-Version: 1.0\r\nContent-Type: multipart/related; boundary=x\r\n\r\n--x\r\nContent-Type: text/html\r\nContent-ID: <same>\r\n\r\n<p>x</p>\r\n--x\r\nContent-Type: image/png\r\nContent-ID: <same>\r\n\r\nbytes\r\n--x--\r\n".to_owned(),
            "MIME-Version: 1.0\r\nContent-Type: multipart/related; boundary=x\r\n\r\n--x\r\nContent-Type: text/html\r\nContent-Transfer-Encoding: binary\r\n\r\n<p>x</p>\r\n--x--\r\n".to_owned(),
        ];
        for input in malformed {
            assert!(
                Document::from_mhtml_bytes(input.as_bytes()).is_err(),
                "accepted {input:?}"
            );
        }
        assert_eq!(
            super::css_resource_references(
                r#"background-image: url("https://example.test/image)one.png")"#,
            )
            .unwrap(),
            vec!["https://example.test/image)one.png"]
        );

        let two_parts = "MIME-Version: 1.0\r\nContent-Type: multipart/related; boundary=x\r\n\r\n--x\r\nContent-Type: text/html\r\n\r\n<p>x</p>\r\n--x\r\nContent-Type: image/png\r\nContent-ID: <image>\r\n\r\nbytes\r\n--x--\r\n";
        for limits in [
            MhtmlLimits {
                input_bytes: 1,
                ..MhtmlLimits::default()
            },
            MhtmlLimits {
                header_bytes: 1,
                ..MhtmlLimits::default()
            },
            MhtmlLimits {
                parts: 1,
                ..MhtmlLimits::default()
            },
            MhtmlLimits {
                part_bytes: 1,
                ..MhtmlLimits::default()
            },
            MhtmlLimits {
                total_decoded_bytes: 1,
                ..MhtmlLimits::default()
            },
        ] {
            assert!(from_mhtml_with_limits(two_parts.as_bytes(), limits).is_err());
        }
    }

    #[test]
    fn mhtml_images_are_limited_to_png_and_jpeg_in_both_directions() {
        for (content_type, bytes) in [
            ("image/png", one_pixel_png()),
            ("image/jpeg", one_pixel_jpeg()),
        ] {
            let imported = Document::from_mhtml_bytes(&mhtml_image_fixture(content_type, &bytes))
                .expect("supported MHTML image");
            assert_eq!(imported.document.images().len(), 1);
            let exported = imported
                .document
                .to_mhtml_bytes()
                .expect("supported export");
            assert!(
                exported
                    .bytes
                    .windows(content_type.len())
                    .any(|window| window == content_type.as_bytes())
            );
        }

        let unsupported: [(&str, &[u8], &str); 7] = [
            ("image/gif", b"GIF89a", "image.gif"),
            ("image/bmp", b"BM", "image.bmp"),
            ("image/tiff", b"II*\0", "image.tiff"),
            ("image/webp", b"RIFF\0\0\0\0WEBP", "image.webp"),
            (
                "image/svg+xml",
                b"<svg xmlns=\"http://www.w3.org/2000/svg\"></svg>",
                "image.svg",
            ),
            (
                "image/emf",
                b"\x01\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0 EMF",
                "image.emf",
            ),
            ("image/wmf", b"\xd7\xcd\xc6\x9a", "image.wmf"),
        ];
        for (content_type, bytes, filename) in unsupported {
            assert!(
                Document::from_mhtml_bytes(&mhtml_image_fixture(content_type, bytes)).is_err(),
                "accepted {content_type} import"
            );
            let mut document = Document::new();
            document.add_picture(bytes, filename, Length::emu(9_525), Length::emu(9_525));
            assert!(
                document.to_mhtml_bytes().is_err(),
                "accepted {content_type} export"
            );
        }
    }

    #[test]
    fn mhtml_transfer_decoding_and_resource_resolution_are_exact() {
        let png = super::BASE64.encode(one_pixel_png());
        let input = format!(
            "MIME-Version: 1.0\r\nContent-Type: multipart/related;\r\n boundary=folded; start=\"<root@rdocx>\"\r\n\r\n--folded\r\nContent-Type: text/html; charset=utf-8\r\nContent-ID: <root@rdocx>\r\nContent-Location: https://Example.Test/folder/index.html\r\nContent-Transfer-Encoding: quoted-printable\r\n\r\n<p>before<img src=3D\"images/pixel.png\" width=3D\"1\"><a href=3D\"next.html\">link</a>after</p>\r\n--folded\r\nContent-Type: image/png; name=pixel.png\r\nContent-ID: <pixel@rdocx>\r\nContent-Location: images/./pixel.png\r\nContent-Transfer-Encoding: base64\r\n\r\n{png}\r\n--folded--\r\n"
        );
        let parsed =
            Document::from_mhtml_bytes(input.as_bytes()).expect("bounded multipart fixture");
        assert_eq!(parsed.document.text(), "beforelinkafter\n");
        assert!(parsed.diagnostics.is_empty());
        let image = &parsed.document.images()[0];
        assert_eq!((image.width_emu, image.height_emu), (9_525, 9_525));
        assert_eq!(
            parsed.document.image_data(&image.embed_id),
            Some(one_pixel_png())
        );
        assert_eq!(
            parsed.document.links()[0].url.as_deref(),
            Some("https://example.test/folder/next.html")
        );

        let cid_input = input.replace("images/pixel.png\" width", "cid:pixel@rdocx\" width");
        let cid = Document::from_mhtml_bytes(cid_input.as_bytes()).expect("exact cid lookup");
        assert_eq!(cid.document.images().len(), 1);
    }

    #[test]
    fn mhtml_writer_is_deterministic_bounded_and_collision_safe() {
        let mut document = Document::new();
        document.add_paragraph("deterministic");
        document.add_picture(
            &one_pixel_png(),
            "same.png",
            Length::emu(19_050),
            Length::emu(28_575),
        );
        document.add_picture(
            &one_pixel_png(),
            "same-again.png",
            Length::emu(19_050),
            Length::emu(28_575),
        );
        let default_html = document.to_html();
        let first = document.to_mhtml_bytes().expect("first MHTML write");
        let second = document.to_mhtml_bytes().expect("second MHTML write");
        assert_eq!(first.bytes, second.bytes);
        assert_eq!(document.to_html(), default_html);
        assert!(first.bytes.windows(2).all(|window| window != b"\n\n"));
        for (index, byte) in first.bytes.iter().enumerate() {
            if *byte == b'\n' {
                assert!(index > 0 && first.bytes[index - 1] == b'\r');
            }
        }
        let text = std::str::from_utf8(&first.bytes).unwrap();
        assert_eq!(text.matches("Content-ID: <image-0@rdocx>").count(), 1);
        let (html, _) = parse_mhtml(&first.bytes, MhtmlLimits::default()).unwrap();
        assert_eq!(html.matches("cid:image-0@rdocx").count(), 2);
        for section in text.split("Content-Transfer-Encoding: base64\r\n").skip(1) {
            let body = section.split("\r\n\r\n").nth(1).unwrap();
            for line in body.lines().take_while(|line| !line.starts_with("--")) {
                assert!(line.len() <= 76);
            }
        }
        let reopened = Document::from_mhtml_bytes(&first.bytes).unwrap();
        assert_eq!(reopened.document.images()[0].width_emu, 19_050);
        assert_eq!(reopened.document.images()[0].height_emu, 28_575);
        assert!(
            to_mhtml_with_limits(
                &document,
                MhtmlLimits {
                    output_bytes: 16,
                    ..MhtmlLimits::default()
                }
            )
            .is_err()
        );
    }

    #[test]
    fn mhtml_writer_sizes_the_pictures_content_controls_wrap() {
        let mut document = Document::new();
        for (name, width, height) in [
            ("body-control.png", 9_525, 9_525),
            ("outside.png", 19_050, 28_575),
            ("inline-control.png", 9_525, 9_525),
        ] {
            document.add_picture(
                &one_pixel_png(),
                name,
                Length::emu(width),
                Length::emu(height),
            );
        }
        let mut package =
            OpcPackage::from_reader(Cursor::new(document.to_bytes().unwrap())).unwrap();
        let mut xml =
            String::from_utf8(package.get_part("/word/document.xml").unwrap().to_vec()).unwrap();
        // The last run holds the third picture. Google Docs wraps it in an
        // inline control, and Word wraps the first one in a picture control.
        let run_start = xml.rfind("<w:r>").unwrap();
        let run_end = xml.rfind("</w:r>").unwrap() + "</w:r>".len();
        xml.insert_str(run_end, "</w:sdtContent></w:sdt>");
        xml.insert_str(
            run_start,
            r#"<w:sdt><w:sdtPr><w:tag w:val="goog_rdk_0"/></w:sdtPr><w:sdtContent>"#,
        );
        let xml = xml
            .replacen(
                "<w:p>",
                "<w:sdt><w:sdtPr><w:picture/></w:sdtPr><w:sdtContent><w:p>",
                1,
            )
            .replacen("</w:p>", "</w:p></w:sdtContent></w:sdt>", 1);
        package.set_part("/word/document.xml", xml.into_bytes());
        let mut saved = Cursor::new(Vec::new());
        package.write_to(&mut saved).unwrap();
        let document = Document::from_bytes(saved.get_ref()).unwrap();
        assert_eq!(document.images().len(), 3);

        let written = document
            .to_mhtml_bytes()
            .expect("pictures in content controls are exported in place");
        assert_eq!(
            written
                .diagnostics
                .iter()
                .map(|diagnostic| (diagnostic.location.as_str(), diagnostic.message.as_str()))
                .collect::<Vec<_>>(),
            vec![
                ("body[0]", "flattened Word body content control"),
                (
                    "body[2]/paragraph/item[0]",
                    "flattened Word paragraph content control"
                ),
            ]
        );
        let reopened = Document::from_mhtml_bytes(&written.bytes).unwrap();
        assert_eq!(
            reopened
                .document
                .images()
                .iter()
                .map(|image| (image.width_emu, image.height_emu))
                .collect::<Vec<_>>(),
            [(9_525, 9_525), (19_050, 28_575), (9_525, 9_525)]
        );
    }

    #[test]
    fn mhtml_writer_leaves_out_deleted_row_pictures() {
        let mut document = Document::new();
        for (name, width, height) in [("deleted.png", 9_525, 9_525), ("kept.png", 19_050, 28_575)] {
            document.add_picture(
                &one_pixel_png(),
                name,
                Length::emu(width),
                Length::emu(height),
            );
        }
        let mut package =
            OpcPackage::from_reader(Cursor::new(document.to_bytes().unwrap())).unwrap();
        let xml =
            String::from_utf8(package.get_part("/word/document.xml").unwrap().to_vec()).unwrap();
        let start = xml.find("<w:p>").unwrap();
        let end = xml.find("</w:p>").unwrap() + "</w:p>".len();
        let xml = format!(
            r#"{}<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:trPr><w:del w:id="1" w:author="Ada"/></w:trPr><w:tc>{}</w:tc></w:tr></w:tbl>{}"#,
            &xml[..start],
            &xml[start..end],
            &xml[end..]
        );
        package.set_part("/word/document.xml", xml.into_bytes());
        let mut saved = Cursor::new(Vec::new());
        package.write_to(&mut saved).unwrap();
        let document = Document::from_bytes(saved.get_ref()).unwrap();
        assert!(!document.to_html().contains("<table"));

        let written = document.to_mhtml_bytes().unwrap();
        let reopened = Document::from_mhtml_bytes(&written.bytes).unwrap();
        assert_eq!(
            reopened
                .document
                .images()
                .iter()
                .map(|image| (image.width_emu, image.height_emu))
                .collect::<Vec<_>>(),
            [(19_050, 28_575)]
        );
    }

    #[test]
    fn mhtml_loss_records_do_not_hide_supported_siblings() {
        let parsed =
            Document::from_mhtml_bytes(&mhtml_fixture("<p>before<object>loss</object>after</p>"))
                .expect("safe lossy MHTML import");
        assert_eq!(parsed.document.text(), "beforeafter\n");
        assert_eq!(parsed.diagnostics.len(), 1);

        let mut seed = Document::new();
        let bytes = seed.to_bytes().unwrap();
        let mut package = OpcPackage::from_reader(Cursor::new(bytes)).unwrap();
        package.set_part(
            "/word/document.xml",
            br#"<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:producer" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t>before</w:t></w:r><w:r><w:delText>deleted</w:delText></w:r><w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple><w:r><w:footnoteReference w:id="1"/></w:r><w:r><w:endnoteReference w:id="2"/></w:r><w:r><w:commentReference w:id="3"/></w:r><w:r><w:t>kept</w:t><x:run/></w:r><w:r><w:pict><v:rect o:hr="t"/></w:pict></w:r><w:r><w:drawing><wp:inline><wp:extent cx="10" cy="20"/><wp:docPr id="1" name="Linked"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:link="rIdLinked"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r><w:r><w:drawing><wp:inline><wp:extent cx="10" cy="20"/><wp:docPr id="2" name="Missing"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rIdMissing"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r><w:r><w:drawing><wp:anchor><wp:docPr id="3" name="Shape"/><wps:wsp><wps:spPr><a:blipFill><a:blip r:embed="rIdFill"/></a:blipFill></wps:spPr></wps:wsp></wp:anchor></w:drawing></w:r><w:r><w:drawing><wp:inline><wp:extent cx="10" cy="20"/><wp:docPr id="4" name="Chart"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rIdChart"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r><w:hyperlink w:anchor="bookmark" w:tooltip="tip" w:docLocation="there" x:flag="yes"><x:link/><w:ins w:id="4" w:author="Ada"><w:r><w:t>revision</w:t></w:r></w:ins><w:r><w:t>linked</w:t><x:nested/></w:r></w:hyperlink><w:r><w:t>after</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"#.to_vec(),
        );
        let mut saved = Cursor::new(Vec::new());
        package.write_to(&mut saved).unwrap();
        let document = Document::from_bytes(saved.get_ref()).unwrap();
        let written = document.to_mhtml_bytes().unwrap();
        assert_eq!(
            written
                .diagnostics
                .iter()
                .map(|diagnostic| (diagnostic.location.as_str(), diagnostic.message.as_str()))
                .collect::<Vec<_>>(),
            vec![
                (
                    "body[0]/paragraph/item[1]/run/item[0]",
                    "dropped Word deleted-text semantics",
                ),
                (
                    "body[0]/paragraph/item[2]/run/item[0]",
                    "dropped Word field semantics",
                ),
                (
                    "body[0]/paragraph/item[3]/run/item[0]",
                    "dropped Word footnote reference",
                ),
                (
                    "body[0]/paragraph/item[4]/run/item[0]",
                    "dropped Word endnote reference",
                ),
                (
                    "body[0]/paragraph/item[5]/run/item[0]",
                    "dropped Word comment reference",
                ),
                (
                    "body[0]/paragraph/item[6]/run/item[1]",
                    "dropped unsupported Word run XML",
                ),
                (
                    "body[0]/paragraph/item[7]/run/item[0]",
                    "dropped Word legacy horizontal rule",
                ),
                (
                    "body[0]/paragraph/item[8]/run/item[0]",
                    "dropped linked Word image",
                ),
                (
                    "body[0]/paragraph/item[9]/run/item[0]",
                    "dropped unresolved Word image",
                ),
                (
                    "body[0]/paragraph/item[10]/run/item[0]",
                    "dropped Word DrawingML shape",
                ),
                (
                    "body[0]/paragraph/item[11]/run/item[0]",
                    "dropped unsupported Word drawing",
                ),
                (
                    "body[0]/paragraph/item[12]/hyperlink/anchor",
                    "dropped Word internal hyperlink anchor",
                ),
                (
                    "body[0]/paragraph/item[12]/hyperlink/tooltip",
                    "dropped Word hyperlink tooltip",
                ),
                (
                    "body[0]/paragraph/item[12]/hyperlink/doc-location",
                    "dropped Word hyperlink document location",
                ),
                (
                    "body[0]/paragraph/item[12]/hyperlink/attributes",
                    "dropped unsupported Word hyperlink attributes",
                ),
                (
                    "body[0]/paragraph/item[12]/hyperlink/item[0]",
                    "dropped unsupported Word hyperlink XML",
                ),
                (
                    "body[0]/paragraph/item[12]/hyperlink/item[1]",
                    "flattened Word hyperlink revision",
                ),
                (
                    "body[0]/paragraph/item[12]/hyperlink/item[2]/run/item[1]",
                    "dropped unsupported Word run XML",
                ),
            ]
        );
        let reopened = Document::from_mhtml_bytes(&written.bytes).unwrap();
        let text = reopened.document.text();
        for supported in ["before", "1", "kept", "linked", "after"] {
            assert!(
                text.contains(supported),
                "missing supported sibling {supported:?}"
            );
        }
    }

    #[test]
    fn html_import_rejects_each_declared_resource_limit() {
        let cases = [
            (
                "<p>x</p>",
                Limits {
                    input_bytes: 1,
                    ..Limits::default()
                },
            ),
            (
                "<p>x</p>",
                Limits {
                    retained_text: 0,
                    ..Limits::default()
                },
            ),
            (
                "<div><p>x</p></div>",
                Limits {
                    depth: 2,
                    ..Limits::default()
                },
            ),
            (
                "<p>x</p>",
                Limits {
                    nodes: 1,
                    ..Limits::default()
                },
            ),
            (
                "<p>x</p>",
                Limits {
                    blocks: 0,
                    ..Limits::default()
                },
            ),
            (
                "<p>x</p>",
                Limits {
                    runs: 0,
                    ..Limits::default()
                },
            ),
            (
                "<pre>\n\n</pre>",
                Limits {
                    runs: 0,
                    ..Limits::default()
                },
            ),
            (
                "<table><tr><td>x</td></tr></table>",
                Limits {
                    rows: 0,
                    ..Limits::default()
                },
            ),
            (
                "<table><tr><td colspan='2'>x</td></tr></table>",
                Limits {
                    columns: 1,
                    ..Limits::default()
                },
            ),
            (
                "<table><tr><td>x</td></tr></table>",
                Limits {
                    cells: 0,
                    ..Limits::default()
                },
            ),
            (
                "<p style='unknown:x'>x</p>",
                Limits {
                    diagnostics: 0,
                    ..Limits::default()
                },
            ),
        ];
        for (html, limits) in cases {
            assert!(
                from_html_with_limits(html, limits).is_err(),
                "limit accepted {html}"
            );
        }

        let oversized = "x".repeat(64 * 1024 * 1024 + 1);
        assert!(Document::from_html(&oversized).is_err());

        let mut reader = Cursor::new(b"12345".to_vec());
        assert!(read_bounded(&mut reader, 0, 4).is_err());
        assert!(
            preflight_markup(
                "<!-- one --><!-- two -->",
                Limits {
                    nodes: 4,
                    ..Limits::default()
                }
            )
            .is_err()
        );
        assert!(
            preflight_markup(
                "<!-- <x><x><x> -->",
                Limits {
                    nodes: 5,
                    ..Limits::default()
                }
            )
            .is_ok()
        );
    }
}
