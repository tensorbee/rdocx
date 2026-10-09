//! CLI command implementations.

use std::collections::HashMap;
use std::io::{self, Write};
use std::ops::Range;
use std::path::{Path, PathBuf};

use oxml_cli_support::{
    StagedOutputSet, default_output_path, ensure_output_paths_allowed,
    ensure_output_paths_available, json_envelope, parse_range,
};
use oxml_opc::relationship::rel_types;
use quick_xml::XmlVersion;
use quick_xml::escape::escape;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use rdocx::{
    BodyItemRef, ComparisonGranularity, ComparisonOptions, ComparisonStoryKind, Document,
    HdrFtrType, HeaderFooterKind, RasterFormat, RasterOptions, RasterOutput, RevisionKind,
    RevisionView, RunRange, StoryId, StoryItemKind, StoryKind,
};
use rdocx_oxml::content_control::{CT_Sdt, SdtContent};
use rdocx_oxml::document::{BodyContent, CT_Document};
use rdocx_oxml::namespace::{MC_NS, W_NS};
use rdocx_oxml::table::{CT_Row, CT_Tbl, CT_Tc, CellContent};
use rdocx_oxml::text::{CT_P, CT_R};
use serde_json::{Value, json};

type Result<T> = std::result::Result<T, Box<dyn std::error::Error>>;

pub struct ImageOptions<'a> {
    pub pages: Option<&'a str>,
    pub quality: u8,
    pub transparent: bool,
}

pub struct RenderOptions<'a> {
    pub page: Option<usize>,
    pub pages: Option<&'a str>,
    pub format: &'a str,
    pub quality: u8,
    pub transparent: bool,
    pub revision_view: RevisionView,
}

#[derive(Clone, Copy)]
pub enum RevisionAction {
    Accept,
    Reject,
}

pub struct RevisionSelector<'a> {
    pub id: Option<i32>,
    pub author: Option<&'a str>,
    pub start_date: Option<&'a str>,
    pub end_date: Option<&'a str>,
}

/// Inspect a DOCX file and print structure information.
pub fn inspect(file: &Path, json: bool) -> Result<()> {
    let doc = Document::open(file)?;

    let paragraph_count = doc.paragraph_count();
    let table_count = doc.table_count();
    let content_count = doc.content_count();

    let title = doc.title().map(|s| s.to_string());
    let author = doc.author().map(|s| s.to_string());
    let subject = doc.subject().map(|s| s.to_string());
    let keywords = doc.keywords().map(|s| s.to_string());

    // Collect style IDs used in paragraphs
    let mut style_ids: Vec<String> = Vec::new();
    for para in doc.paragraphs() {
        if let Some(style) = para.style_id() {
            let s = style.to_string();
            if !style_ids.contains(&s) {
                style_ids.push(s);
            }
        }
    }

    let mut stdout = io::stdout().lock();
    if json {
        let obj = inspect_json(file, &doc, style_ids)?;
        writeln!(stdout, "{}", serde_json::to_string_pretty(&obj)?)?;
    } else {
        writeln!(stdout, "File: {}", file.display())?;
        writeln!(stdout, "Paragraphs: {paragraph_count}")?;
        writeln!(stdout, "Tables: {table_count}")?;
        writeln!(stdout, "Content elements: {content_count}")?;
        writeln!(stdout)?;
        writeln!(stdout, "Metadata:")?;
        if let Some(t) = &title {
            writeln!(stdout, "  Title: {t}")?;
        }
        if let Some(a) = &author {
            writeln!(stdout, "  Author: {a}")?;
        }
        if let Some(s) = &subject {
            writeln!(stdout, "  Subject: {s}")?;
        }
        if let Some(k) = &keywords {
            writeln!(stdout, "  Keywords: {k}")?;
        }
        if title.is_none() && author.is_none() && subject.is_none() && keywords.is_none() {
            writeln!(stdout, "  (none)")?;
        }
        writeln!(stdout)?;
        writeln!(stdout, "Styles used:")?;
        if style_ids.is_empty() {
            writeln!(stdout, "  (none)")?;
        } else {
            for sid in &style_ids {
                writeln!(stdout, "  - {sid}")?;
            }
        }
    }

    Ok(())
}

fn inspect_json(file: &Path, doc: &Document, style_ids: Vec<String>) -> Result<Value> {
    Ok(json_envelope(json!({
        "file": file.display().to_string(),
        "paragraphs": doc.paragraph_count(),
        "tables": doc.table_count(),
        "content_elements": doc.content_count(),
        "metadata": {
            "title": doc.title(),
            "author": doc.author(),
            "subject": doc.subject(),
            "keywords": doc.keywords(),
        },
        "styles_used": style_ids,
    }))?)
}

/// Extract plain text from a DOCX file.
///
/// Body paragraphs and table cell text are both emitted, in document order —
/// printing only `paragraphs()` would silently drop everything inside tables.
/// Every other story follows the body: text boxes, headers, footers,
/// footnotes, endnotes, and comments.
pub fn text(file: &Path, json_output: bool) -> Result<()> {
    let doc = Document::open(file)?;
    if json_output {
        let document = parsed_main_document(file)?;
        let mut paragraphs = Vec::new();
        let mut joining = Vec::new();
        for (body_index, content) in document.body.content.iter().enumerate() {
            match content {
                BodyContent::Paragraph(paragraph) => {
                    joining.push(paragraph);
                    if !document.body.accepted_paragraph_joins_next(body_index) {
                        paragraphs.push(joined_paragraph_json(body_index, &joining));
                        joining.clear();
                    }
                }
                BodyContent::Table(table) => {
                    collect_table_paragraphs(body_index, &[], table, &mut paragraphs);
                }
                BodyContent::ContentControl(control) => {
                    collect_control_paragraphs(body_index, &[], control, &mut paragraphs);
                }
                BodyContent::RawXml(_) => {}
            }
        }
        let stories = readable_stories(file, &doc);
        // Without the other stories the record covers the main story only.
        let scope = if stories.is_some() {
            "all-supported-stories"
        } else {
            "main"
        };
        let stories = stories
            .unwrap_or_default()
            .iter()
            .map(|story| {
                let items = story
                    .items
                    .iter()
                    .map(|(index_path, kind, text)| {
                        json!({ "index_path": index_path, "kind": kind, "text": text })
                    })
                    .collect::<Vec<_>>();
                json!({
                    "kind": story_kind_name(story.story.kind()),
                    "part_name": story.story.part_name(),
                    "owner_index": story.story.owner_index(),
                    "items": items,
                })
            })
            .collect::<Vec<_>>();
        print_json(json!({
            "scope": scope,
            "revision_view": "accepted",
            "paragraphs": paragraphs,
            "stories": stories,
        }))?;
    } else {
        let mut stdout = io::stdout().lock();
        write!(stdout, "{}", doc.text())?;
        let stories = readable_stories(file, &doc).unwrap_or_default();
        for (kind, part_name, texts) in story_parts(&stories) {
            writeln!(stdout, "--- {} ({part_name}) ---", story_kind_name(kind))?;
            for text in texts {
                writeln!(stdout, "{text}")?;
            }
        }
    }
    Ok(())
}

/// One story that the body view leaves out, with the index path, kind, and
/// accepted-view text of each of its direct paragraphs and block content
/// controls.
struct StoryText {
    story: StoryId,
    items: Vec<(Vec<usize>, &'static str, String)>,
}

/// Read the other stories for a text view that already has its body.
///
/// A story part that cannot be read, such as a truncated header or a header
/// relationship to a missing part, does not cost the body. The view prints
/// one warning on stderr, naming the part when it can, and lists no other
/// story. `validate` reports the same part as an error.
fn readable_stories(file: &Path, doc: &Document) -> Option<Vec<StoryText>> {
    let error = match other_stories(doc) {
        Ok(stories) => return Some(stories),
        Err(error) => error,
    };
    let malformed = oxml_opc::OpcPackage::open(file)
        .ok()
        .and_then(|package| malformed_story_part(&package));
    eprintln!(
        "Warning: other stories left out: {}",
        malformed.unwrap_or_else(|| error.to_string())
    );
    None
}

/// Name the first story part of the main document that is not well-formed
/// XML, with the reason.
fn malformed_story_part(package: &oxml_opc::OpcPackage) -> Option<String> {
    let doc_part = package.main_document_part()?;
    let rels = package.get_part_rels(&doc_part)?;
    rels.items
        .iter()
        .filter(|rel| {
            rel.target_mode.as_deref() != Some("External")
                && STORY_RELATIONSHIPS.contains(&rel.rel_type.as_str())
        })
        .find_map(|rel| {
            let part_name = oxml_opc::OpcPackage::resolve_rel_target(&doc_part, &rel.target);
            let detail = xml_style_references(package.get_part(&part_name)?).err()?;
            Some(format!("part {part_name} is not well-formed XML: {detail}"))
        })
}

/// The relationships from the main document to its other story parts.
const STORY_RELATIONSHIPS: [&str; 5] = [
    rel_types::HEADER,
    rel_types::FOOTER,
    rel_types::FOOTNOTES,
    rel_types::ENDNOTES,
    rel_types::COMMENTS,
];

/// Read every story but the main body and its table cells, which the body
/// view already covers, in [`Document::stories`] order.
///
/// Only the direct items of a story are read, so the text of an inline
/// control or a field is not repeated after its paragraph. A table is left
/// out because each of its cells is a table-cell story of its own.
fn other_stories(doc: &Document) -> Result<Vec<StoryText>> {
    let mut main_part = None;
    let mut stories: Vec<StoryText> = Vec::new();
    for item in doc.story_item_snapshots()? {
        let story = item.location().story();
        if story.kind() == StoryKind::Body {
            main_part = Some(story.part_name().to_owned());
            continue;
        }
        if story.kind() == StoryKind::TableCell && main_part.as_deref() == Some(story.part_name()) {
            continue;
        }
        if stories.last().is_none_or(|last| &last.story != story) {
            stories.push(StoryText {
                story: story.clone(),
                items: Vec::new(),
            });
        }
        let kind = match item.location().item_kind() {
            StoryItemKind::Paragraph => "paragraph",
            StoryItemKind::ContentControl => "content_control",
            _ => continue,
        };
        if item.is_direct_child()
            && let Some(last) = stories.last_mut()
        {
            last.items.push((
                item.location().index_path().to_vec(),
                kind,
                item.text().unwrap_or_default().to_owned(),
            ));
        }
    }
    Ok(stories)
}

/// Group the texts of the stories by package part, each part labelled with
/// the kind of its first story, such as a header part holding table cells.
/// A part without any text, such as an empty header variant, is left out.
fn story_parts(stories: &[StoryText]) -> Vec<(StoryKind, &str, Vec<&str>)> {
    let mut parts: Vec<(StoryKind, &str, Vec<&str>)> = Vec::new();
    for story in stories {
        let part_name = story.story.part_name();
        if parts.last().is_none_or(|(_, last, _)| *last != part_name) {
            parts.push((story.story.kind(), part_name, Vec::new()));
        }
        if let Some((_, _, texts)) = parts.last_mut() {
            texts.extend(story.items.iter().map(|(_, _, text)| text.as_str()));
        }
    }
    parts.retain(|(_, _, texts)| texts.iter().any(|text| !text.trim().is_empty()));
    parts
}

/// Name a story kind as Python `Story.kind` does.
fn story_kind_name(kind: StoryKind) -> &'static str {
    match kind {
        StoryKind::Body => "body",
        StoryKind::TableCell => "table_cell",
        StoryKind::Header => "header",
        StoryKind::Footer => "footer",
        StoryKind::Footnote => "footnote",
        StoryKind::Endnote => "endnote",
        StoryKind::Comment => "comment",
        StoryKind::TextBox => "text_box",
        _ => "unknown",
    }
}

/// The story parts a converted document carries after its body. Comments are
/// review annotations rather than document content, so they are left out.
fn converted_story_parts(stories: &[StoryText]) -> Vec<(StoryKind, &str, Vec<&str>)> {
    story_parts(stories)
        .into_iter()
        .filter(|(kind, _, _)| *kind != StoryKind::Comment)
        .collect()
}

/// Append the stories after the Markdown body, each part under a bold label.
fn push_markdown_stories(markdown: &mut String, stories: &[StoryText]) {
    let parts = converted_story_parts(stories);
    if parts.is_empty() {
        return;
    }
    if !markdown.is_empty() {
        if !markdown.ends_with("\n\n") {
            markdown.push_str(if markdown.ends_with('\n') {
                "\n"
            } else {
                "\n\n"
            });
        }
        markdown.push_str("---\n\n");
    }
    for (kind, part_name, texts) in parts {
        markdown.push_str(&format!("**{}** ({part_name})\n\n", story_kind_name(kind)));
        for text in texts.iter().map(|text| text.trim()) {
            if !text.is_empty() {
                markdown.push_str(text);
                markdown.push_str("\n\n");
            }
        }
    }
}

/// Insert the stories before the end of the HTML body, one section per part.
fn push_html_stories(html: &mut String, stories: &[StoryText]) {
    let parts = converted_story_parts(stories);
    if parts.is_empty() {
        return;
    }
    let mut sections = String::from("<hr>\n");
    for (kind, part_name, texts) in parts {
        sections.push_str(&format!(
            "<section>\n<p><strong>{}</strong> ({})</p>\n",
            story_kind_name(kind),
            escape(part_name)
        ));
        for text in texts.iter().map(|text| text.trim()) {
            if !text.is_empty() {
                sections.push_str(&format!("<p>{}</p>\n", escape(text)));
            }
        }
        sections.push_str("</section>\n");
    }
    match html.rfind("</body>") {
        Some(end) => html.insert_str(end, &sections),
        None => html.push_str(&sections),
    }
}

/// Emit deterministic point-space extents for every direct body item.
pub fn layout(file: &Path, json_output: bool) -> Result<()> {
    let doc = Document::open(file)?;
    let layout = doc.layout_deterministic()?;
    let body_items = doc
        .body_items()
        .enumerate()
        .map(|(body_index, item)| {
            let kind = match item {
                BodyItemRef::Paragraph(_) => "paragraph",
                BodyItemRef::Table(_) => "table",
                BodyItemRef::ContentControl(_) => "content-control",
                BodyItemRef::UnsupportedXml(_) => "unsupported-xml",
            };
            let fragments = layout
                .body_layout_fragments(body_index)
                .unwrap_or_default()
                .iter()
                .map(|fragment| {
                    json!({
                        "physical_page": fragment.physical_page,
                        "displayed_page": fragment.displayed_page,
                        "x": fragment.x,
                        "y": fragment.y,
                        "width": fragment.width,
                        "height": fragment.height,
                    })
                })
                .collect::<Vec<_>>();
            json!({
                "body_index": body_index,
                "kind": kind,
                "fragments": fragments,
            })
        })
        .collect::<Vec<_>>();

    if json_output {
        print_json(json!({
            "scope": "main",
            "revision_view": "accepted",
            "units": "points",
            "page_count": layout.layout.pages.len(),
            "body_items": body_items,
        }))?;
    } else {
        let mut stdout = io::stdout().lock();
        writeln!(stdout, "Pages: {}", layout.layout.pages.len())?;
        for item in body_items {
            writeln!(
                stdout,
                "Body item {} ({}): {} fragment(s)",
                item["body_index"],
                item["kind"].as_str().unwrap_or("unknown"),
                item["fragments"].as_array().map_or(0, Vec::len)
            )?;
        }
    }
    Ok(())
}

fn parsed_main_document(file: &Path) -> Result<CT_Document> {
    let package = oxml_opc::OpcPackage::open(file)?;
    let part = package
        .main_document_part()
        .ok_or("package declares no main document relationship")?;
    let bytes = package
        .get_part(&part)
        .ok_or_else(|| format!("main document part {part} is missing"))?;
    Ok(CT_Document::from_xml(bytes)?)
}

fn path_segment(kind: &str, index: usize) -> Value {
    json!({ "kind": kind, "index": index })
}

fn paragraph_json(body_index: usize, path: &[Value], paragraph: &CT_P) -> Value {
    let runs = paragraph
        .accepted_bookmark_runs()
        .into_iter()
        .enumerate()
        .map(|(index, run)| {
            json!({
                "index": index,
                "text": run.text(),
                "formatting": run_formatting_json(run),
            })
        })
        .collect::<Vec<_>>();
    let style = paragraph
        .properties
        .as_ref()
        .and_then(|properties| properties.style_id.as_deref());
    let numbering = paragraph.properties.as_ref().and_then(|properties| {
        properties.num_id.map(|num_id| {
            json!({
                "num_id": num_id,
                "level": properties.num_ilvl.unwrap_or(0),
            })
        })
    });
    json!({
        "body_index": body_index,
        "path": path,
        "style": style,
        "numbering": numbering,
        "text": paragraph.accepted_text(),
        "runs": runs,
    })
}

fn joined_paragraph_json(body_index: usize, paragraphs: &[&CT_P]) -> Value {
    let mut result = paragraph_json(body_index, &[], paragraphs[paragraphs.len() - 1]);
    if paragraphs.len() == 1 {
        return result;
    }
    let mut text = String::new();
    let mut runs = Vec::new();
    for paragraph in paragraphs {
        let projected = paragraph_json(body_index, &[], paragraph);
        text.push_str(projected["text"].as_str().unwrap_or_default());
        if let Some(projected_runs) = projected["runs"].as_array() {
            runs.extend(projected_runs.iter().cloned());
        }
    }
    for (index, run) in runs.iter_mut().enumerate() {
        run["index"] = json!(index);
    }
    result["text"] = json!(text);
    result["runs"] = json!(runs);
    result
}

fn run_formatting_json(run: &CT_R) -> Value {
    let Some(properties) = run.properties.as_ref() else {
        return Value::Null;
    };
    json!({
        "bold": properties.bold,
        "italic": properties.italic,
        "strike": properties.strike,
        "underline": properties.underline.map(|value| value.to_str()),
        "font": properties.font_ascii,
        "size_points": properties.sz.map(|value| value.to_pt()),
        "color": properties.color,
        "highlight": properties.highlight.map(|value| value.to_str()),
        "language": properties.language,
        "style": properties.style_id,
    })
}

fn collect_table_paragraphs(
    body_index: usize,
    path: &[Value],
    table: &CT_Tbl,
    output: &mut Vec<Value>,
) {
    // Row controls and direct rows retain their model paths and source order.
    for boundary in 0..=table.rows.len() {
        for (control_index, (at, _, control)) in table.content_controls.iter().enumerate() {
            if *at == boundary {
                let mut control_path = path.to_vec();
                control_path.push(path_segment("content-control", control_index));
                collect_control_paragraphs(body_index, &control_path, control, output);
            }
        }
        if let Some(row) = table.rows.get(boundary)
            && !row.accepted_view_removes()
        {
            let mut row_path = path.to_vec();
            row_path.push(path_segment("row", boundary));
            collect_row_paragraphs(body_index, &row_path, row, output);
        }
    }
}

fn collect_row_paragraphs(
    body_index: usize,
    path: &[Value],
    row: &CT_Row,
    output: &mut Vec<Value>,
) {
    for boundary in 0..=row.cells.len() {
        for (control_index, (at, _, control)) in row.content_controls.iter().enumerate() {
            if *at == boundary {
                let mut control_path = path.to_vec();
                control_path.push(path_segment("content-control", control_index));
                collect_control_paragraphs(body_index, &control_path, control, output);
            }
        }
        if let Some(cell) = row.cells.get(boundary) {
            let mut cell_path = path.to_vec();
            cell_path.push(path_segment("cell", boundary));
            collect_cell_paragraphs(body_index, &cell_path, cell, output);
        }
    }
}

fn collect_cell_paragraphs(
    body_index: usize,
    path: &[Value],
    cell: &CT_Tc,
    output: &mut Vec<Value>,
) {
    for (content_index, content) in cell.content.iter().enumerate() {
        let mut content_path = path.to_vec();
        match content {
            CellContent::Paragraph(paragraph) => {
                content_path.push(path_segment("paragraph", content_index));
                output.push(paragraph_json(body_index, &content_path, paragraph));
            }
            CellContent::Table(table) => {
                content_path.push(path_segment("table", content_index));
                collect_table_paragraphs(body_index, &content_path, table, output);
            }
            CellContent::ContentControl(control) => {
                content_path.push(path_segment("content-control", content_index));
                collect_control_paragraphs(body_index, &content_path, control, output);
            }
        }
    }
}

fn collect_control_paragraphs(
    body_index: usize,
    path: &[Value],
    control: &CT_Sdt,
    output: &mut Vec<Value>,
) {
    for (content_index, content) in control.content.iter().enumerate() {
        let mut content_path = path.to_vec();
        match content {
            SdtContent::Paragraph(paragraph) => {
                content_path.push(path_segment("paragraph", content_index));
                output.push(paragraph_json(body_index, &content_path, paragraph));
            }
            SdtContent::Table(table) => {
                content_path.push(path_segment("table", content_index));
                collect_table_paragraphs(body_index, &content_path, table, output);
            }
            SdtContent::Row(row) if row.accepted_view_removes() => {}
            SdtContent::Row(row) => {
                content_path.push(path_segment("row", content_index));
                collect_row_paragraphs(body_index, &content_path, row, output);
            }
            SdtContent::Cell(cell) => {
                content_path.push(path_segment("cell", content_index));
                collect_cell_paragraphs(body_index, &content_path, cell, output);
            }
            SdtContent::ContentControl(nested) => {
                content_path.push(path_segment("content-control", content_index));
                collect_control_paragraphs(body_index, &content_path, nested, output);
            }
            SdtContent::Run(_) | SdtContent::RawXml(_) => {}
        }
    }
}

/// Convert a DOCX file to another format.
#[allow(clippy::too_many_arguments)]
pub fn convert(
    file: &Path,
    to: &str,
    output: Option<&Path>,
    force: bool,
    dpi: u32,
    font_dir: Option<&Path>,
    revision_view: RevisionView,
    image: ImageOptions<'_>,
) -> Result<()> {
    let doc = Document::open(file)?;
    let render_options = rdocx::RenderOptions { revision_view };

    let default_ext = match to {
        "pdf" => "pdf",
        "html" => "html",
        "md" | "markdown" => "md",
        "png" => "png",
        "jpg" | "jpeg" => "jpg",
        "tif" | "tiff" => "tiff",
        other => {
            return Err(format!(
                "Unknown format: {other}. Supported: pdf, html, md, png, jpeg, tiff"
            )
            .into());
        }
    };

    if revision_view == RevisionView::Tracked && matches!(default_ext, "html" | "md") {
        return Err("--revision-view tracked applies only to PDF and image output".into());
    }

    let output_path = match output {
        Some(p) => p.to_path_buf(),
        None => default_output_path(file, default_ext),
    };
    // Several PNG or JPEG pages are written under numbered names, so those
    // outputs are checked once the pages are selected.
    if !matches!(to, "png" | "jpg" | "jpeg") {
        ensure_output_paths_allowed(std::slice::from_ref(&output_path), file, force)?;
    }

    let mut stdout = io::stdout().lock();
    match to {
        "pdf" => {
            let bytes = if let Some(dir) = font_dir {
                let font_files = Document::load_fonts_from_dir(dir);
                let font_refs: Vec<(&str, &[u8])> = font_files
                    .iter()
                    .map(|f| (f.family.as_str(), f.data.as_slice()))
                    .collect();
                doc.to_pdf_with_fonts_and_options(&font_refs, render_options)?
            } else {
                doc.to_pdf_with_options(render_options)?
            };
            stage_and_publish(&[(output_path.clone(), bytes)], force)?;
        }
        "html" => {
            let mut html = doc.to_html();
            if let Some(stories) = readable_stories(file, &doc) {
                push_html_stories(&mut html, &stories);
            }
            stage_and_publish(&[(output_path.clone(), html.into_bytes())], force)?;
        }
        "md" | "markdown" => {
            let mut md = doc.to_markdown();
            if let Some(stories) = readable_stories(file, &doc) {
                push_markdown_stories(&mut md, &stories);
            }
            stage_and_publish(&[(output_path.clone(), md.into_bytes())], force)?;
        }
        "png" | "jpg" | "jpeg" | "tif" | "tiff" => {
            let (format, extension) = parse_image_format(to, image.quality, image.transparent)?;
            let layout = doc.layout_deterministic_with_options(render_options)?;
            let selected = selected_zero_based_pages(layout.layout.pages.len(), image.pages)?;
            match format {
                RasterFormat::Tiff => {
                    let output = oxml_pdf::render_pages(
                        &layout.layout,
                        &selected,
                        RasterOptions {
                            dpi: dpi as f64,
                            format,
                        },
                    )?;
                    let RasterOutput::MultiPageTiff(tiff) = output else {
                        return Err("TIFF render did not produce one stream".into());
                    };
                    stage_and_publish(&[(output_path.clone(), tiff)], force)?;
                    writeln!(stdout, "Written to {}", output_path.display())?;
                }
                RasterFormat::Png { .. } | RasterFormat::Jpeg { .. } => {
                    let output_paths =
                        convert_separate_output_paths(&output_path, extension, selected.len());
                    ensure_output_paths_allowed(&output_paths, file, force)?;
                    let mut staged = StagedOutputSet::with_replace_existing(force);
                    for (page_index, path) in selected.iter().zip(output_paths.iter()) {
                        let image = render_one_raster_page(
                            &layout.layout,
                            *page_index,
                            RasterOptions {
                                dpi: dpi as f64,
                                format,
                            },
                        )?;
                        staged.stage_bytes(path, &image)?;
                    }
                    staged.publish()?;
                    if output_paths.len() == 1 {
                        writeln!(stdout, "Written to {}", output_path.display())?;
                    } else {
                        let stem = output_path
                            .file_stem()
                            .unwrap_or_default()
                            .to_string_lossy();
                        let parent = output_path.parent().unwrap_or(Path::new("."));
                        writeln!(
                            stdout,
                            "Written {} pages to {}/{stem}_NNN.{extension}",
                            output_paths.len(),
                            parent.display()
                        )?;
                    }
                }
            }
            return Ok(());
        }
        _ => unreachable!(),
    }

    writeln!(stdout, "Written to {}", output_path.display())?;
    Ok(())
}

/// Compare the paragraph text of every story of two DOCX files.
///
/// The body paragraphs keep their historical `[i]` locations, every other
/// paragraph is located by its story. `diff_status` is set before anything is
/// printed, so the verdict survives a reader that closes standard output.
pub fn diff(file_a: &Path, file_b: &Path, json_output: bool, diff_status: &mut i32) -> Result<()> {
    let doc_a = Document::open(file_a)?;
    let doc_b = Document::open(file_b)?;
    let (streams_a, mut not_compared) = diff_streams(file_a, &doc_a)?;
    let (streams_b, not_compared_b) = diff_streams(file_b, &doc_b)?;
    for entry in not_compared_b {
        if !not_compared.contains(&entry) {
            not_compared.push(entry);
        }
    }

    // Streams pair by name, in the order of the first file and then of the
    // streams that only the second file has.
    let mut names: Vec<&str> = streams_a
        .iter()
        .map(|stream| stream.name.as_str())
        .collect();
    for stream in &streams_b {
        if !names.contains(&stream.name.as_str()) {
            names.push(&stream.name);
        }
    }
    let mut changes = Vec::new();
    for name in names {
        changes.extend(diff_items(
            stream_items(&streams_a, name),
            stream_items(&streams_b, name),
        ));
    }
    let count = |a: bool, b: bool| {
        changes
            .iter()
            .filter(|(old, new)| old.is_some() == a && new.is_some() == b)
            .count()
    };
    let (changed, added, removed) = (count(true, true), count(false, true), count(true, false));
    *diff_status = if !not_compared.is_empty() {
        2
    } else if changes.is_empty() {
        0
    } else {
        1
    };

    if json_output {
        let side = |item: &Option<DiffItem>, field: fn(&DiffItem) -> &str| {
            item.as_ref().map(|item| field(item).to_owned())
        };
        let differences = changes
            .iter()
            .map(|(old, new)| {
                let (change, story) = match (old, new) {
                    (Some(item), Some(_)) => ("changed", item.story),
                    (None, Some(item)) => ("added", item.story),
                    (Some(item), None) => ("removed", item.story),
                    (None, None) => unreachable!("a change has at least one side"),
                };
                json!({
                    "change": change,
                    "story": story,
                    "location_a": side(old, |item| &item.location),
                    "text_a": side(old, |item| &item.text),
                    "location_b": side(new, |item| &item.location),
                    "text_b": side(new, |item| &item.text),
                })
            })
            .collect::<Vec<_>>();
        print_json(json!({
            "file_a": file_a.display().to_string(),
            "file_b": file_b.display().to_string(),
            "scope": "all-supported-stories",
            "revision_view": "accepted",
            "changed": changed,
            "added": added,
            "removed": removed,
            "differences": differences,
            "not_compared": not_compared,
        }))?;
        return Ok(());
    }

    let mut stdout = io::stdout().lock();
    writeln!(
        stdout,
        "--- {} ({} paragraphs, {} tables)",
        file_a.display(),
        doc_a.paragraph_count(),
        doc_a.table_count()
    )?;
    writeln!(
        stdout,
        "+++ {} ({} paragraphs, {} tables)",
        file_b.display(),
        doc_b.paragraph_count(),
        doc_b.table_count()
    )?;
    writeln!(stdout)?;
    for (old, new) in &changes {
        if let Some(item) = old {
            writeln!(stdout, "- [{}] {}", item.location, item.text)?;
        }
        if let Some(item) = new {
            writeln!(stdout, "+ [{}] {}", item.location, item.text)?;
        }
    }
    if !changes.is_empty() {
        writeln!(stdout)?;
    }
    for entry in &not_compared {
        writeln!(stdout, "(not compared: {entry})")?;
    }
    if changes.is_empty() && not_compared.is_empty() {
        writeln!(stdout, "(no differences in paragraph text)")?;
    } else if !changes.is_empty() {
        writeln!(
            stdout,
            "{changed} paragraph(s) changed, {added} added, {removed} removed."
        )?;
    }
    Ok(())
}

/// One compared paragraph or block content control: its story kind, its
/// location as `diff` prints it between brackets, and its accepted-view text.
#[derive(Clone)]
struct DiffItem {
    story: &'static str,
    location: String,
    text: String,
}

/// The items that `diff` compares as one sequence, such as the body, the
/// tables of the body, or one header. Streams of two files pair by name.
struct DiffStream {
    name: String,
    items: Vec<DiffItem>,
}

/// Read every story of a document into comparable streams, and name the
/// stories that could not be read instead of dropping them silently.
fn diff_streams(file: &Path, doc: &Document) -> Result<(Vec<DiffStream>, Vec<String>)> {
    let body = doc
        .paragraphs()
        .iter()
        .enumerate()
        .map(|(index, paragraph)| DiffItem {
            story: "body",
            location: (index + 1).to_string(),
            text: paragraph.text(),
        })
        .collect();
    let mut streams = vec![DiffStream {
        name: "body".to_owned(),
        items: body,
    }];

    // The table cells of the body use the `text --json` traversal. A table
    // inside a block content control is left to the body paragraphs, which
    // already hold its text.
    let document = parsed_main_document(file)?;
    let mut cells = Vec::new();
    let mut table_number = 0;
    for (body_index, content) in document.body.content.iter().enumerate() {
        if let BodyContent::Table(table) = content {
            table_number += 1;
            let mut paragraphs = Vec::new();
            collect_table_paragraphs(body_index, &[], table, &mut paragraphs);
            cells.extend(paragraphs.iter().map(|paragraph| DiffItem {
                story: "table_cell",
                location: format!("table {table_number}, {}", json_path_label(paragraph)),
                text: paragraph["text"].as_str().unwrap_or_default().to_owned(),
            }));
        }
    }
    streams.push(DiffStream {
        name: "tables".to_owned(),
        items: cells,
    });

    let mut not_compared = Vec::new();
    match other_story_streams(doc) {
        Ok((other, unsupported)) => {
            streams.extend(other);
            not_compared.extend(unsupported);
        }
        Err(error) => not_compared.push(format!(
            "headers, footers, footnotes, endnotes, comments and text boxes of {} ({error})",
            file.display()
        )),
    }
    Ok((streams, not_compared))
}

/// Label a `text --json` paragraph path one-based, such as `row 1, cell 2,
/// paragraph 1`.
fn json_path_label(paragraph: &Value) -> String {
    paragraph["path"]
        .as_array()
        .map(|segments| {
            segments
                .iter()
                .map(|segment| {
                    format!(
                        "{} {}",
                        segment["kind"]
                            .as_str()
                            .unwrap_or_default()
                            .replace('-', " "),
                        segment["index"].as_u64().unwrap_or_default() + 1
                    )
                })
                .collect::<Vec<_>>()
                .join(", ")
        })
        .unwrap_or_default()
}

/// Read the stories beyond the main body and its tables: text boxes of the
/// body, headers and footers named by section and type, footnotes, endnotes,
/// and comments. Each package part is one stream. Only the direct paragraphs
/// and block content controls of a story are read, so the text of an inline
/// control or a field is not compared twice.
fn other_story_streams(doc: &Document) -> Result<(Vec<DiffStream>, Vec<String>)> {
    // A header or footer part is named by the first section that references
    // it explicitly, so that a renamed part still pairs with its section.
    let mut section_names: Vec<(String, String)> = Vec::new();
    for section in 0..doc.section_count() {
        for (kind, kind_name) in [
            (HeaderFooterKind::Header, "header"),
            (HeaderFooterKind::Footer, "footer"),
        ] {
            for (hdr_type, type_name) in [
                (HdrFtrType::Default, "default"),
                (HdrFtrType::First, "first"),
                (HdrFtrType::Even, "even"),
            ] {
                if let Some(story) = doc.section_story(section, kind, hdr_type)?
                    && !story.is_inherited()
                    && !section_names
                        .iter()
                        .any(|(part, _)| part == story.story().part_name())
                {
                    section_names.push((
                        story.story().part_name().to_owned(),
                        format!("{kind_name} {type_name}, section {}", section + 1),
                    ));
                }
            }
        }
    }

    let mut main_part = None;
    let mut streams: Vec<(String, DiffStream)> = Vec::new();
    let mut unsupported = Vec::new();
    // The current owner and its paragraph and content control counts so far.
    let mut owner: Option<(StoryId, [usize; 2])> = None;
    for item in doc.story_item_snapshots()? {
        let story = item.location().story();
        let part = story.part_name();
        if story.kind() == StoryKind::Body {
            main_part = Some(part.to_owned());
            continue;
        }
        if story.kind() == StoryKind::TableCell && main_part.as_deref() == Some(part) {
            continue;
        }
        if !item.is_direct_child() {
            continue;
        }
        // Paragraphs and block content controls count apart within their
        // owner, so `paragraph 2` is the second paragraph of its story.
        let (item_name, is_paragraph) = match item.location().item_kind() {
            StoryItemKind::Paragraph => ("paragraph", true),
            StoryItemKind::ContentControl => ("content control", false),
            // A table's text is read from its cells, which are stories of
            // their own, and the other direct items carry no text.
            _ => continue,
        };
        if owner.as_ref().is_none_or(|(current, _)| current != story) {
            owner = Some((story.clone(), [0, 0]));
        }
        let ordinal = match &mut owner {
            Some((_, counts)) => {
                let count = &mut counts[usize::from(!is_paragraph)];
                *count += 1;
                *count
            }
            None => unreachable!("the owner was just set"),
        };
        let (story_name, owner_name) = match story.kind() {
            StoryKind::Header => ("header", None),
            StoryKind::Footer => ("footer", None),
            StoryKind::Footnote => ("footnote", Some("footnote")),
            StoryKind::Endnote => ("endnote", Some("endnote")),
            StoryKind::Comment => ("comment", Some("comment")),
            StoryKind::TableCell => ("table_cell", Some("table cell")),
            StoryKind::TextBox => ("text_box", Some("text box")),
            _ => {
                let entry = format!("a story of an unsupported kind in {part}");
                if !unsupported.contains(&entry) {
                    unsupported.push(entry);
                }
                continue;
            }
        };
        let position = match streams
            .iter()
            .position(|(stream_part, _)| stream_part == part)
        {
            Some(position) => position,
            None => {
                let name = section_names
                    .iter()
                    .find(|(named_part, _)| named_part == part)
                    .map(|(_, name)| name.clone())
                    .unwrap_or_else(|| match story.kind() {
                        StoryKind::Header | StoryKind::Footer => format!("{story_name} {part}"),
                        StoryKind::Footnote => "footnotes".to_owned(),
                        StoryKind::Endnote => "endnotes".to_owned(),
                        StoryKind::Comment => "comments".to_owned(),
                        _ if main_part.as_deref() == Some(part) => String::new(),
                        _ => part.to_owned(),
                    });
                streams.push((
                    part.to_owned(),
                    DiffStream {
                        name,
                        items: Vec::new(),
                    },
                ));
                streams.len() - 1
            }
        };
        let stream = &mut streams[position].1;
        // A note or a comment names itself. A header or a footer takes the
        // name of its stream, and a cell or a text box is placed in it.
        let mut location = Vec::new();
        let names_itself = matches!(
            story.kind(),
            StoryKind::Footnote | StoryKind::Endnote | StoryKind::Comment
        );
        if !names_itself && !stream.name.is_empty() {
            location.push(stream.name.clone());
        }
        if let Some(owner_name) = owner_name {
            location.push(format!("{owner_name} {}", story.owner_index() + 1));
        }
        location.push(format!("{item_name} {ordinal}"));
        stream.items.push(DiffItem {
            story: story_name,
            location: location.join(", "),
            text: item.text().unwrap_or_default().to_owned(),
        });
    }
    Ok((
        streams.into_iter().map(|(_, stream)| stream).collect(),
        unsupported,
    ))
}

/// The items of the stream with this name, or none when the file lacks it.
fn stream_items<'a>(streams: &'a [DiffStream], name: &str) -> &'a [DiffItem] {
    streams
        .iter()
        .find(|stream| stream.name == name)
        .map_or(&[], |stream| stream.items.as_slice())
}

/// A removed item, an added item, or both for a changed item.
type ItemChange = (Option<DiffItem>, Option<DiffItem>);

/// Pair the items of two streams that differ. Matched items come from a
/// shortest edit script between their texts. Between two matches, removed
/// and added items pair in order as changed items, and the rest stay removed
/// or added.
fn diff_items(a: &[DiffItem], b: &[DiffItem]) -> Vec<ItemChange> {
    // Equal texts share one number, so the edit search compares integers.
    let mut numbers: HashMap<&str, u32> = HashMap::new();
    let mut numbered = Vec::with_capacity(a.len() + b.len());
    for item in a.iter().chain(b) {
        let next = numbers.len() as u32;
        numbered.push(*numbers.entry(item.text.as_str()).or_insert(next));
    }
    let (numbers_a, numbers_b) = numbered.split_at(a.len());
    let mut matches = Vec::new();
    let mut frontiers = Frontiers::new(a.len() + b.len());
    matching_indexes(
        numbers_a,
        0..a.len(),
        numbers_b,
        0..b.len(),
        &mut frontiers,
        &mut matches,
    );
    matches.push((a.len(), b.len()));
    let mut changes = Vec::new();
    let (mut i, mut j) = (0, 0);
    for (next_i, next_j) in matches {
        let removed = &a[i..next_i];
        let added = &b[j..next_j];
        for index in 0..removed.len().max(added.len()) {
            changes.push((removed.get(index).cloned(), added.get(index).cloned()));
        }
        i = next_i + 1;
        j = next_j + 1;
    }
    changes
}

/// The forward and backward furthest-reaching `x` of each diagonal in
/// Myers' search, indexed by the diagonal `k` offset to stay non-negative.
struct Frontiers {
    forward: Vec<usize>,
    backward: Vec<usize>,
    offset: isize,
}

impl Frontiers {
    fn new(total: usize) -> Self {
        let size = total + 3;
        Self {
            forward: vec![0; 2 * size + 1],
            backward: vec![0; 2 * size + 1],
            offset: size as isize,
        }
    }

    fn slot(&self, k: isize) -> usize {
        (k + self.offset) as usize
    }
}

/// Collect the matched index pairs of a shortest edit script between two
/// ranges in order, by Myers' linear-space divide and conquer: O((N+M)D)
/// time and O(N+M) memory, where D is the number of edits. A few edits in a
/// long story therefore stay cheap.
fn matching_indexes(
    a: &[u32],
    mut a_range: Range<usize>,
    b: &[u32],
    mut b_range: Range<usize>,
    frontiers: &mut Frontiers,
    matches: &mut Vec<(usize, usize)>,
) {
    while !a_range.is_empty() && !b_range.is_empty() && a[a_range.start] == b[b_range.start] {
        matches.push((a_range.start, b_range.start));
        a_range.start += 1;
        b_range.start += 1;
    }
    let mut suffix = 0;
    while !a_range.is_empty() && !b_range.is_empty() && a[a_range.end - 1] == b[b_range.end - 1] {
        a_range.end -= 1;
        b_range.end -= 1;
        suffix += 1;
    }
    if !a_range.is_empty() && !b_range.is_empty() {
        let (x, y) = middle_snake(a, a_range.clone(), b, b_range.clone(), frontiers);
        matching_indexes(a, a_range.start..x, b, b_range.start..y, frontiers, matches);
        matching_indexes(a, x..a_range.end, b, y..b_range.end, frontiers, matches);
    }
    matches.extend((0..suffix).map(|offset| (a_range.end + offset, b_range.end + offset)));
}

/// Find a point on a shortest edit path that splits it into two halves of
/// about half its edits each. Both ranges are non-empty and differ at both
/// ends, so each half has fewer edits than the whole.
fn middle_snake(
    a: &[u32],
    a_range: Range<usize>,
    b: &[u32],
    b_range: Range<usize>,
    frontiers: &mut Frontiers,
) -> (usize, usize) {
    let (n, m) = (a_range.len(), b_range.len());
    let delta = n as isize - m as isize;
    let odd = delta & 1 == 1;
    let one = frontiers.slot(1);
    frontiers.forward[one] = 0;
    frontiers.backward[one] = 0;
    let common_prefix = |x: usize, y: usize| {
        a[a_range.start + x..a_range.end]
            .iter()
            .zip(&b[b_range.start + y..b_range.end])
            .take_while(|(left, right)| left == right)
            .count()
    };
    let common_suffix = |x: usize, y: usize| {
        a[a_range.start..a_range.end - x]
            .iter()
            .rev()
            .zip(b[b_range.start..b_range.end - y].iter().rev())
            .take_while(|(left, right)| left == right)
            .count()
    };
    for d in 0..=(n + m).div_ceil(2) as isize {
        for k in (-d..=d).rev().step_by(2) {
            let (below, above) = (frontiers.slot(k - 1), frontiers.slot(k + 1));
            let mut x =
                if k == -d || (k != d && frontiers.forward[below] < frontiers.forward[above]) {
                    frontiers.forward[above]
                } else {
                    frontiers.forward[below] + 1
                };
            let y = (x as isize - k) as usize;
            let start = (x, y);
            if x < n && y < m {
                x += common_prefix(x, y);
            }
            let slot = frontiers.slot(k);
            frontiers.forward[slot] = x;
            if odd
                && (k - delta).abs() < d
                && x + frontiers.backward[frontiers.slot(delta - k)] >= n
            {
                return (a_range.start + start.0, b_range.start + start.1);
            }
        }
        for k in (-d..=d).rev().step_by(2) {
            let (below, above) = (frontiers.slot(k - 1), frontiers.slot(k + 1));
            let mut x =
                if k == -d || (k != d && frontiers.backward[below] < frontiers.backward[above]) {
                    frontiers.backward[above]
                } else {
                    frontiers.backward[below] + 1
                };
            let mut y = (x as isize - k) as usize;
            if x < n && y < m {
                let common = common_suffix(x, y);
                x += common;
                y += common;
            }
            let slot = frontiers.slot(k);
            frontiers.backward[slot] = x;
            if !odd
                && (k - delta).abs() <= d
                && x + frontiers.forward[frontiers.slot(delta - k)] >= n
            {
                return (a_range.end - x, b_range.end - y);
            }
        }
    }
    unreachable!("two non-empty ranges always meet within half their total length of edits")
}

fn comment_position_json(
    document: &Document,
    position: &rdocx::StoryRunPosition,
    items: &mut HashMap<rdocx::ContentLocation, rdocx::StoryItemSnapshot>,
) -> Result<Value> {
    if !items.contains_key(&position.location) {
        items.insert(
            position.location.clone(),
            document.story_range_paragraph_snapshot(&position.location)?,
        );
    }
    let snapshot = &items[&position.location];
    let location = snapshot.location();
    Ok(json!({
        "story_kind": story_kind_name(location.story().kind()),
        "part_name": location.story().part_name(),
        "owner_index": location.story().owner_index(),
        "item_kind": "paragraph",
        "index_path": location.index_path(),
        "run_index": position.run_index,
        "direct_body_index": snapshot.direct_body_index(),
    }))
}

/// List Word comments in their package order.
pub fn comment_list(file: &Path, json_output: bool) -> Result<()> {
    let doc = Document::open(file)?;
    let comments = doc.comments();
    let mut anchors = if json_output {
        doc.comment_anchor_snapshots()?
    } else {
        Default::default()
    };
    let mut items = if json_output {
        doc.story_item_snapshots()?
            .into_iter()
            .map(|item| (item.location().clone(), item))
            .collect::<HashMap<_, _>>()
    } else {
        HashMap::new()
    };
    let records = comments
        .iter()
        .map(|comment| -> Result<Value> {
            let (range, anchor_text) = if json_output {
                anchors
                    .remove(&comment.id())
                    .ok_or_else(|| io::Error::other("checked comment ownership record missing"))?
            } else {
                (None, None)
            };
            let anchor = range
                .map(|range| -> Result<Value> {
                    Ok(json!({
                        "start": comment_position_json(&doc, &range.start, &mut items)?,
                        "end": comment_position_json(&doc, &range.end, &mut items)?,
                    }))
                })
                .transpose()?;
            Ok(json!({
                "id": comment.id(), "author": comment.author(), "initials": comment.initials(),
                "date": comment.date(), "text": comment.text(), "parent_id": comment.parent_id(),
                "resolved": comment.resolved(), "anchor_text": anchor_text, "anchor": anchor,
            }))
        })
        .collect::<Result<Vec<_>>>()?;
    let mut stdout = io::stdout().lock();
    if json_output {
        print_json(json!({
            "scope": "all_stories",
            "comments": records,
        }))?;
    } else if comments.is_empty() {
        writeln!(stdout, "(no comments)")?;
    } else {
        for comment in comments {
            writeln!(
                stdout,
                "{}\t{}\t{}\t{}",
                comment.id(),
                comment.author().unwrap_or(""),
                if comment.resolved() {
                    "resolved"
                } else {
                    "open"
                },
                comment.text().replace('\n', " ")
            )?;
        }
    }
    Ok(())
}

/// Where `comment add` anchors its comment.
pub enum CommentAnchor<'a> {
    /// A zero-based half-open body run range.
    Range(RunRange),
    /// The zero-based occurrence of a literal text in the main story.
    Text { text: &'a str, occurrence: usize },
}

/// Add one Word comment and publish the complete mutated document atomically.
#[allow(clippy::too_many_arguments)]
pub fn comment_add(
    file: &Path,
    anchor: CommentAnchor<'_>,
    author: &str,
    initials: Option<&str>,
    text: &str,
    date: Option<&str>,
    output: &Path,
    json_output: bool,
) -> Result<()> {
    let mut doc = Document::open(file)?;
    let id = match anchor {
        CommentAnchor::Range(range) => {
            doc.add_comment_with_date(range, author, initials, text, date)?
        }
        CommentAnchor::Text {
            text: anchor,
            occurrence,
        } => doc.add_comment_on_text(anchor, occurrence, author, initials, text, date)?,
    };
    publish_document(&mut doc, output)?;
    mutation_record(
        json_output,
        "main",
        "add",
        json!({ "comment_id": id }),
        output,
    )
}

/// Move one existing thread and publish only the reopened complete package.
pub fn comment_move(
    file: &Path,
    id: i32,
    text: &str,
    occurrence: usize,
    output: &Path,
    json_output: bool,
) -> Result<()> {
    let mut doc = Document::open(file)?;
    doc.move_comment_to_text(id, text, occurrence)?;
    publish_document(&mut doc, output)?;
    mutation_record(
        json_output,
        "main",
        "move",
        json!({"comment_id": id}),
        output,
    )
}

/// Add one reply and publish the complete mutated document atomically.
pub fn comment_reply(
    file: &Path,
    parent_id: i32,
    author: &str,
    text: &str,
    date: Option<&str>,
    output: &Path,
    json_output: bool,
) -> Result<()> {
    let mut doc = Document::open(file)?;
    let id = doc.reply_to_with_date(parent_id, author, text, date)?;
    publish_document(&mut doc, output)?;
    mutation_record(
        json_output,
        "main",
        "reply",
        json!({ "comment_id": id, "parent_id": parent_id }),
        output,
    )
}

/// Resolve one comment thread and publish the complete document atomically.
pub fn comment_resolve(file: &Path, id: i32, output: &Path, json_output: bool) -> Result<()> {
    let mut doc = Document::open(file)?;
    if !doc.resolve_comment(id, true)? {
        return Err(format!("comment id {id} does not exist").into());
    }
    publish_document(&mut doc, output)?;
    mutation_record(
        json_output,
        "main",
        "resolve",
        json!({ "comment_id": id }),
        output,
    )
}

/// Remove one comment thread and publish the complete document atomically.
pub fn comment_remove(file: &Path, id: i32, output: &Path, json_output: bool) -> Result<()> {
    let mut doc = Document::open(file)?;
    if !doc.remove_comment(id)? {
        return Err(format!("comment id {id} does not exist").into());
    }
    publish_document(&mut doc, output)?;
    mutation_record(
        json_output,
        "main",
        "remove",
        json!({ "comment_id": id }),
        output,
    )
}

/// List modeled revisions from every supported story.
pub fn revision_list(file: &Path, json_output: bool) -> Result<()> {
    let doc = Document::open(file)?;
    let revisions = doc.story_revisions()?;
    let records = revisions
        .iter()
        .map(|revision| {
            json!({
                "id": revision.id(),
                "author": revision.author(),
                "timestamp": revision.timestamp(),
                "kind": revision_kind_label(revision.kind()),
                "story": story_json(revision.story()),
            })
        })
        .collect::<Vec<_>>();
    let mut stdout = io::stdout().lock();
    if json_output {
        print_json(json!({
            "scope": "all-supported-stories",
            "revisions": records,
        }))?;
    } else if revisions.is_empty() {
        writeln!(stdout, "(no revisions)")?;
    } else {
        for revision in revisions {
            writeln!(
                stdout,
                "{}\t{}\t{}\t{}\t{}\t{}\t{}",
                revision.id(),
                revision.author(),
                revision.timestamp().unwrap_or(""),
                revision_kind_label(revision.kind()),
                story_kind_label(revision.story().kind()),
                revision.story().part_name(),
                revision.story().owner_index()
            )?;
        }
    }
    Ok(())
}

/// Accept or reject a validated revision selection across every supported story.
pub fn resolve_revisions(
    file: &Path,
    action: RevisionAction,
    selector: RevisionSelector<'_>,
    output: &Path,
    json_output: bool,
) -> Result<()> {
    validate_revision_selector(&selector)?;
    let mut doc = Document::open(file)?;
    let count = match (
        action,
        selector.id,
        selector.author,
        selector.start_date,
        selector.end_date,
    ) {
        (RevisionAction::Accept, Some(id), None, None, None) => doc.accept_revision_id(id)?,
        (RevisionAction::Reject, Some(id), None, None, None) => doc.reject_revision_id(id)?,
        (RevisionAction::Accept, None, Some(author), None, None) => {
            doc.accept_revisions_by_author(author)?
        }
        (RevisionAction::Reject, None, Some(author), None, None) => {
            doc.reject_revisions_by_author(author)?
        }
        (RevisionAction::Accept, None, None, Some(start), Some(end)) => {
            doc.accept_revisions_in_date_range(start, end)?
        }
        (RevisionAction::Reject, None, None, Some(start), Some(end)) => {
            doc.reject_revisions_in_date_range(start, end)?
        }
        (RevisionAction::Accept, None, None, None, None) => doc.accept_all()?,
        (RevisionAction::Reject, None, None, None, None) => doc.reject_all()?,
        _ => return Err("revision selector is invalid".into()),
    };
    publish_document(&mut doc, output)?;
    let action_label = action.label();
    if json_output {
        print_json(json!({
            "scope": "all-supported-stories",
            "action": action_label,
            "selector": selector_json(&selector),
            "resolved": count,
            "output": output.display().to_string(),
        }))?;
    } else {
        let mut stdout = io::stdout().lock();
        writeln!(stdout, "{action_label}: {count} revision element(s)")?;
        writeln!(stdout, "Written to {}", output.display())?;
    }
    Ok(())
}

/// Granularity names, shared by argument parsing and the JSON record.
const COMPARISON_GRANULARITIES: [(&str, ComparisonGranularity); 3] = [
    ("run", ComparisonGranularity::Run),
    ("word", ComparisonGranularity::Word),
    ("character", ComparisonGranularity::Character),
];

/// Ignorable story names. They follow the Python `Story.kind` names, so
/// `body` selects the main story.
const COMPARISON_STORIES: [(&str, ComparisonStoryKind); 7] = [
    ("body", ComparisonStoryKind::Main),
    ("header", ComparisonStoryKind::Header),
    ("footer", ComparisonStoryKind::Footer),
    ("comment", ComparisonStoryKind::Comment),
    ("text_box", ComparisonStoryKind::TextBox),
    ("footnote", ComparisonStoryKind::Footnote),
    ("endnote", ComparisonStoryKind::Endnote),
];

/// Parse a `--granularity` value.
pub fn parse_comparison_granularity(
    name: &str,
) -> std::result::Result<ComparisonGranularity, String> {
    COMPARISON_GRANULARITIES
        .iter()
        .find(|(known, _)| *known == name)
        .map(|(_, granularity)| *granularity)
        .ok_or_else(|| {
            format!("unknown comparison granularity {name:?}, expected run, word, or character")
        })
}

pub fn parse_revision_view(name: &str) -> std::result::Result<RevisionView, String> {
    match name {
        "accepted" => Ok(RevisionView::Accepted),
        "tracked" => Ok(RevisionView::Tracked),
        _ => Err(format!(
            "unknown revision view {name:?}, expected accepted or tracked"
        )),
    }
}

/// Parse one `--ignore-story` value.
pub fn parse_comparison_story(name: &str) -> std::result::Result<ComparisonStoryKind, String> {
    COMPARISON_STORIES
        .iter()
        .find(|(known, _)| *known == name)
        .map(|(_, kind)| *kind)
        .ok_or_else(|| {
            format!(
                "unknown comparison story {name:?}, expected body, header, footer, comment, \
                 text_box, footnote, or endnote"
            )
        })
}

fn comparison_options_json(options: &ComparisonOptions) -> Value {
    let granularity = COMPARISON_GRANULARITIES
        .iter()
        .find(|(_, known)| *known == options.granularity)
        .map(|(name, _)| *name);
    let stories = options
        .ignored_stories
        .iter()
        .map(|kind| {
            COMPARISON_STORIES
                .iter()
                .find(|(_, known)| known == kind)
                .map(|(name, _)| *name)
        })
        .collect::<Vec<_>>();
    json!({
        "granularity": granularity,
        "ignore_formatting": options.ignore_formatting,
        "ignore_whitespace": options.ignore_whitespace,
        "ignore_fields": options.ignore_fields,
        "ignore_comments": options.ignore_comments,
        "ignored_stories": stories,
    })
}

/// Create a tracked-changes document from an original and edited input.
pub fn compare(
    original: &Path,
    edited: &Path,
    author: &str,
    timestamp: &str,
    options: &ComparisonOptions,
    output: &Path,
    json_output: bool,
) -> Result<()> {
    let mut original_doc = Document::open(original)?;
    let edited_doc = Document::open(edited)?;
    let diagnostics = original_doc.compare_with_options(&edited_doc, author, timestamp, options)?;
    let main_story_revisions = original_doc.revisions().len();
    let revisions = original_doc.story_revisions()?;
    let mut stories = Vec::<(&StoryId, usize)>::new();
    for revision in &revisions {
        match stories
            .iter_mut()
            .find(|(story, _)| *story == revision.story())
        {
            Some((_, count)) => *count += 1,
            None => stories.push((revision.story(), 1)),
        }
    }
    let records = diagnostics
        .iter()
        .map(|diagnostic| {
            json!({
                "location": diagnostic.location,
                "message": diagnostic.message,
            })
        })
        .collect::<Vec<_>>();
    publish_document(&mut original_doc, output)?;
    if json_output {
        let story_records = stories
            .iter()
            .map(|(story, count)| {
                let mut record = story_json(story);
                record["revisions"] = json!(count);
                record
            })
            .collect::<Vec<_>>();
        print_json(json!({
            "scope": "all-supported-stories",
            "options": comparison_options_json(options),
            "revisions": revisions.len(),
            "stories": story_records,
            "main_story_revisions": main_story_revisions,
            "diagnostics": records,
            "output": output.display().to_string(),
        }))?;
    } else {
        let mut stdout = io::stdout().lock();
        writeln!(
            stdout,
            "Created {} revision element(s) in {} story(ies)",
            revisions.len(),
            stories.len()
        )?;
        for (story, count) in &stories {
            writeln!(
                stdout,
                "  {}\t{}\t{}\t{count}",
                story_kind_label(story.kind()),
                story.part_name(),
                story.owner_index()
            )?;
        }
        writeln!(stdout, "Diagnostics: {}", diagnostics.len())?;
        writeln!(stdout, "Written to {}", output.display())?;
    }
    Ok(())
}

/// Rebuild supported existing table-of-contents fields.
pub fn toc_rebuild(file: &Path, output: &Path, json_output: bool) -> Result<()> {
    let mut doc = Document::open(file)?;
    let report = doc.rebuild_toc()?;
    publish_document(&mut doc, output)?;
    if json_output {
        print_json(json!({
            "scope": "main",
            "entry_count": report.entry_count,
            "bookmark_count": report.bookmark_count,
            "diagnostic_count": report.diagnostic_count(),
            "output": output.display().to_string(),
        }))?;
    } else {
        let mut stdout = io::stdout().lock();
        writeln!(stdout, "Entries: {}", report.entry_count)?;
        writeln!(stdout, "Bookmarks: {}", report.bookmark_count)?;
        writeln!(stdout, "Diagnostics: {}", report.diagnostic_count())?;
        writeln!(stdout, "Written to {}", output.display())?;
    }
    Ok(())
}

impl RevisionAction {
    fn label(self) -> &'static str {
        match self {
            Self::Accept => "accept",
            Self::Reject => "reject",
        }
    }
}

fn validate_revision_selector(selector: &RevisionSelector<'_>) -> Result<()> {
    let date_count =
        usize::from(selector.start_date.is_some()) + usize::from(selector.end_date.is_some());
    if date_count == 1 {
        return Err("--start-date and --end-date must be provided together".into());
    }
    let selector_count = usize::from(selector.id.is_some())
        + usize::from(selector.author.is_some())
        + usize::from(date_count == 2);
    if selector_count > 1 {
        return Err("--id, --author, and the date range are mutually exclusive".into());
    }
    Ok(())
}

fn selector_json(selector: &RevisionSelector<'_>) -> Value {
    if let Some(id) = selector.id {
        json!({ "kind": "id", "id": id })
    } else if let Some(author) = selector.author {
        json!({ "kind": "author", "author": author })
    } else if let (Some(start), Some(end)) = (selector.start_date, selector.end_date) {
        json!({ "kind": "date-range", "start": start, "end": end })
    } else {
        json!({ "kind": "all" })
    }
}

fn revision_kind_label(kind: RevisionKind) -> &'static str {
    match kind {
        RevisionKind::Insertion => "insertion",
        RevisionKind::Deletion => "deletion",
        RevisionKind::MoveFrom => "move-from",
        RevisionKind::MoveTo => "move-to",
        RevisionKind::RunPropertyChange => "run-property-change",
        RevisionKind::ParagraphPropertyChange => "paragraph-property-change",
        RevisionKind::TablePropertyChange => "table-property-change",
        RevisionKind::SectionPropertyChange => "section-property-change",
    }
}

fn story_kind_label(kind: StoryKind) -> &'static str {
    match kind {
        StoryKind::Body => "body",
        StoryKind::TableCell => "table-cell",
        StoryKind::Header => "header",
        StoryKind::Footer => "footer",
        StoryKind::Footnote => "footnote",
        StoryKind::Endnote => "endnote",
        StoryKind::Comment => "comment",
        StoryKind::TextBox => "text-box",
        _ => "unknown",
    }
}

fn story_json(story: &StoryId) -> Value {
    json!({
        "kind": story_kind_label(story.kind()),
        "part_name": story.part_name(),
        "owner_index": story.owner_index(),
    })
}

fn publish_document(doc: &mut Document, output: &Path) -> Result<()> {
    let bytes = doc.to_bytes_for_path(output)?;
    stage_and_publish(&[(output.to_path_buf(), bytes)], false)
}

fn mutation_record(
    json_output: bool,
    scope: &str,
    action: &str,
    detail: Value,
    output: &Path,
) -> Result<()> {
    if json_output {
        let mut payload = json!({
            "scope": scope,
            "action": action,
            "output": output.display().to_string(),
        });
        let object = payload.as_object_mut().expect("literal is an object");
        let Value::Object(detail) = detail else {
            return Err("mutation detail must be an object".into());
        };
        object.extend(detail);
        print_json(payload)?;
    } else {
        let mut stdout = io::stdout().lock();
        writeln!(stdout, "{action}")?;
        writeln!(stdout, "Written to {}", output.display())?;
    }
    Ok(())
}

fn print_json(payload: Value) -> Result<()> {
    writeln!(
        io::stdout(),
        "{}",
        serde_json::to_string_pretty(&json_envelope(payload)?)?
    )?;
    Ok(())
}

/// Replace a placeholder in a DOCX file and save to output.
pub fn replace(
    file: &Path,
    placeholder: &str,
    value: &str,
    expect: Option<usize>,
    output: &Path,
) -> Result<()> {
    let mut doc = Document::open(file)?;
    let count = doc.try_replace_text(placeholder, value)?;
    if let Some(expected) = expect
        && count != expected
    {
        return Err(format!(
            "expected {expected} replacement(s) of \"{placeholder}\", found {count}"
        )
        .into());
    }
    publish_document(&mut doc, output)?;
    let mut stdout = io::stdout().lock();
    writeln!(
        stdout,
        "Replaced {count} occurrence(s) of \"{placeholder}\" -> \"{value}\""
    )?;
    writeln!(stdout, "Written to {}", output.display())?;
    Ok(())
}

/// Render document pages to image files.
pub fn render(
    file: &Path,
    output_dir: Option<&Path>,
    force: bool,
    dpi: f64,
    options: RenderOptions<'_>,
) -> Result<()> {
    validate_dpi(dpi)?;
    let doc = Document::open(file)?;
    let out_dir = output_dir.unwrap_or_else(|| Path::new("."));
    let (format, extension) =
        parse_image_format(options.format, options.quality, options.transparent)?;
    let layout = doc.layout_deterministic_with_options(rdocx::RenderOptions {
        revision_view: options.revision_view,
    })?;
    let selected = selected_render_pages(layout.layout.pages.len(), options.page, options.pages)?;
    let stem = file.file_stem().unwrap_or_default().to_string_lossy();
    let legacy_single_page = options.page.is_some();

    let mut stdout = io::stdout().lock();
    // Writing into a directory the user named but has not created should not
    // fail with a bare "No such file or directory". Do it after validation and
    // encoding so invalid options leave no partial output.
    match format {
        RasterFormat::Tiff => {
            let out_path = out_dir.join(format!("{stem}.tiff"));
            ensure_output_paths_allowed(std::slice::from_ref(&out_path), file, force)?;
            let output =
                oxml_pdf::render_pages(&layout.layout, &selected, RasterOptions { dpi, format })?;
            let RasterOutput::MultiPageTiff(tiff) = output else {
                return Err("TIFF render did not produce one stream".into());
            };
            std::fs::create_dir_all(out_dir)?;
            let tiff_len = tiff.len();
            stage_and_publish(&[(out_path.clone(), tiff)], force)?;
            writeln!(
                stdout,
                "Pages {} -> {} ({} bytes)",
                selected
                    .iter()
                    .map(|page| (page + 1).to_string())
                    .collect::<Vec<_>>()
                    .join(","),
                out_path.display(),
                tiff_len
            )?;
        }
        RasterFormat::Png { .. } | RasterFormat::Jpeg { .. } => {
            let output_paths = selected
                .iter()
                .map(|page_index| {
                    let one_based = page_index + 1;
                    out_dir.join(format!("{stem}_page{one_based}.{extension}"))
                })
                .collect::<Vec<_>>();
            std::fs::create_dir_all(out_dir)?;
            ensure_output_paths_allowed(&output_paths, file, force)?;
            let mut staged = StagedOutputSet::with_replace_existing(force);
            let mut rendered = Vec::with_capacity(selected.len());
            for (page_index, out_path) in selected.iter().zip(output_paths.iter()) {
                let one_based = page_index + 1;
                let image = render_one_raster_page(
                    &layout.layout,
                    *page_index,
                    RasterOptions { dpi, format },
                )?;
                staged.stage_bytes(out_path, &image)?;
                rendered.push((one_based, out_path.clone(), image.len()));
            }
            staged.publish()?;
            for (one_based, out_path, len) in rendered {
                writeln!(
                    stdout,
                    "Page {one_based} -> {} ({len} bytes)",
                    out_path.display()
                )?;
            }
            if !legacy_single_page {
                writeln!(stdout, "Rendered {} page(s) at {dpi} DPI", selected.len())?;
            }
        }
    }

    Ok(())
}

fn parse_image_format(
    format: &str,
    quality: u8,
    transparent: bool,
) -> Result<(RasterFormat, &'static str)> {
    match format {
        "png" => Ok((
            RasterFormat::Png {
                transparent_background: transparent,
            },
            "png",
        )),
        "jpg" | "jpeg" => {
            if !(1..=100).contains(&quality) {
                return Err(format!("JPEG quality must be 1 through 100, got {quality}").into());
            }
            Ok((RasterFormat::Jpeg { quality }, "jpg"))
        }
        "tif" | "tiff" => Ok((RasterFormat::Tiff, "tiff")),
        other => Err(format!("Unknown image format: {other}. Supported: png, jpeg, tiff").into()),
    }
}

fn selected_zero_based_pages(page_count: usize, pages: Option<&str>) -> Result<Vec<usize>> {
    let one_based = match pages {
        Some(pages) => parse_range(pages)?,
        None => (1..=page_count).collect(),
    };
    if let Some(page) = one_based.iter().copied().find(|page| *page > page_count) {
        return Err(format!("page {page} is out of range for {page_count} pages").into());
    }
    Ok(one_based.into_iter().map(|page| page - 1).collect())
}

fn selected_render_pages(
    page_count: usize,
    page: Option<usize>,
    pages: Option<&str>,
) -> Result<Vec<usize>> {
    match (page, pages) {
        (Some(_), Some(_)) => Err("--page and --pages cannot be used together".into()),
        (Some(page), None) => {
            if page >= page_count {
                return Err(format!(
                    "zero-based page {page} is out of range for {page_count} pages"
                )
                .into());
            }
            Ok(vec![page])
        }
        (None, pages) => selected_zero_based_pages(page_count, pages),
    }
}

fn convert_separate_output_paths(
    output_path: &Path,
    extension: &str,
    page_count: usize,
) -> Vec<PathBuf> {
    if page_count == 1 {
        return vec![output_path.to_path_buf()];
    }
    let stem = output_path
        .file_stem()
        .unwrap_or_default()
        .to_string_lossy();
    let parent = output_path.parent().unwrap_or(Path::new("."));
    (1..=page_count)
        .map(|index| parent.join(format!("{stem}_{index:03}.{extension}")))
        .collect()
}

fn render_one_raster_page(
    layout: &oxml_layout::LayoutResult,
    page_index: usize,
    options: RasterOptions,
) -> Result<Vec<u8>> {
    let output = oxml_pdf::render_pages(layout, &[page_index], options)?;
    let RasterOutput::SeparatePages(mut pages) = output else {
        return Err("single-page render did not produce a separate image".into());
    };
    if pages.len() != 1 {
        return Err("single-page render returned the wrong page count".into());
    }
    Ok(pages.remove(0))
}

/// Publishes complete outputs, replacing existing files only with `force`.
fn stage_and_publish(outputs: &[(PathBuf, Vec<u8>)], force: bool) -> Result<()> {
    if !force {
        let paths = outputs
            .iter()
            .map(|(path, _)| path.clone())
            .collect::<Vec<_>>();
        ensure_output_paths_available(&paths)?;
    }
    let mut staged = StagedOutputSet::with_replace_existing(force);
    for (path, bytes) in outputs {
        staged.stage_bytes(path, bytes)?;
    }
    staged.publish()?;
    Ok(())
}

fn validate_dpi(dpi: f64) -> Result<()> {
    if !dpi.is_finite() || dpi <= 0.0 {
        return Err("DPI must be a positive finite number".into());
    }
    Ok(())
}

/// Validate a DOCX file's structure.
///
/// Returns `Ok(false)` when a structural error was found, so the caller can
/// exit non-zero — a validator that always succeeds cannot gate anything in CI.
/// A related part that is not well-formed XML and a paragraph, character, or
/// table style id that no style defines are structural errors too.
/// Advisory findings (empty paragraphs, missing metadata) are reported but do
/// not affect the exit status.
pub fn validate(file: &Path) -> Result<bool> {
    let mut errors: Vec<String> = Vec::new();
    let mut warnings: Vec<String> = Vec::new();

    // --- Structural errors ---

    // Every relationship must point at a part that is actually in the package.
    let package = oxml_opc::OpcPackage::open(file)?;
    if let Some(doc_part) = package.main_document_part() {
        if package.get_part(&doc_part).is_none() {
            errors.push(format!("main document part {doc_part} is missing"));
        }
        if let Some(rels) = package.get_part_rels(&doc_part) {
            for rel in &rels.items {
                // External targets live outside the package by definition.
                if rel.target_mode.as_deref() == Some("External") {
                    continue;
                }
                let target = oxml_opc::OpcPackage::resolve_rel_target(&doc_part, &rel.target);
                if package.get_part(&target).is_none() {
                    errors.push(format!(
                        "relationship {} points at missing part {target}",
                        rel.id
                    ));
                }
            }
        }
    } else {
        errors.push("package declares no main document relationship".to_string());
    }

    // Every part needs a declared content type.
    for part_name in package.parts.keys() {
        if package.content_types.content_type_for(part_name).is_none() {
            errors.push(format!("part {part_name} has no declared content type"));
        }
    }

    // Markup Compatibility requires every prefix that `mc:Ignorable` or
    // `mc:MustUnderstand` lists to be declared, and a consumer may reject a
    // part that breaks the rule. A part that does not parse as XML is
    // outside this check.
    let mut xml_parts = package
        .parts
        .iter()
        .filter(|(part_name, _)| {
            package
                .content_types
                .content_type_for(part_name)
                .is_some_and(|content_type| content_type.ends_with("xml"))
        })
        .collect::<Vec<_>>();
    xml_parts.sort();
    for (part_name, xml) in xml_parts {
        let Ok(findings) = rdocx_oxml::namespace::undeclared_compatibility_prefixes(xml) else {
            continue;
        };
        for (attribute, prefix) in findings {
            errors.push(format!(
                "part {part_name} lists undeclared prefix `{prefix}` in mc:{attribute}"
            ));
        }
    }

    // Every XML part the main document relates to must be well formed. The
    // style ids that its story parts name are checked once the document opens.
    let mut style_references = Vec::new();
    if let Some(doc_part) = package.main_document_part() {
        let mut parts = vec![(doc_part.clone(), true)];
        for rel in package
            .get_part_rels(&doc_part)
            .iter()
            .flat_map(|rels| &rels.items)
        {
            if rel.target_mode.as_deref() == Some("External") {
                continue;
            }
            let target = oxml_opc::OpcPackage::resolve_rel_target(&doc_part, &rel.target);
            let story = STORY_RELATIONSHIPS.contains(&rel.rel_type.as_str());
            if !parts.iter().any(|(known, _)| *known == target) {
                parts.push((target, story));
            }
        }
        for (part_name, story) in parts {
            let is_xml = package
                .content_types
                .content_type_for(&part_name)
                .is_some_and(|content_type| {
                    content_type.ends_with("+xml") || content_type.ends_with("/xml")
                });
            let Some(xml) = package.get_part(&part_name).filter(|_| is_xml) else {
                continue;
            };
            match xml_style_references(xml) {
                Err(detail) => {
                    errors.push(format!("part {part_name} is not well-formed XML: {detail}"));
                }
                Ok(references) if story => style_references.extend(
                    references
                        .into_iter()
                        .map(|(kind, style_id)| (kind, style_id, part_name.clone())),
                ),
                Ok(_) => {}
            }
        }
    }

    // A malformed part can keep the document from opening. The errors found
    // so far explain why, so they are reported together with the failure.
    let doc = match Document::open(file) {
        Ok(doc) => doc,
        Err(error) if !errors.is_empty() => {
            errors.push(format!("the document does not open: {error}"));
            return report_validation(file, &errors, &warnings);
        }
        Err(error) => return Err(error.into()),
    };

    // Word falls back to the default style for a style id that no style
    // defines, so the style that the document names is silently lost.
    for (kind, style_id, part_name) in style_references {
        let issue = format!("{kind} style {style_id:?} used in {part_name} is not defined");
        if doc.style(&style_id).is_none() && !errors.contains(&issue) {
            errors.push(issue);
        }
    }

    if let Err(error) = doc.validate_comment_ownership() {
        errors.push(format!("invalid comment ownership: {error}"));
    }

    // --- Advisory findings ---

    if doc.content_count() == 0 {
        warnings.push("Document has no content (no paragraphs or tables)".to_string());
    }

    let empty_count = doc
        .paragraphs()
        .iter()
        .filter(|p| !p.has_visible_content())
        .count();
    if empty_count > 0 {
        warnings.push(format!("{empty_count} empty paragraph(s) found"));
    }

    let mut prev_level: Option<u32> = None;
    for (level, _) in doc.headings() {
        if let Some(prev) = prev_level
            && level > prev + 1
        {
            warnings.push(format!(
                "Heading level gap: Heading{prev} -> Heading{level} (skipped level(s))"
            ));
        }
        prev_level = Some(level);
    }

    if doc.title().is_none() {
        warnings.push("Missing document title".to_string());
    }
    if doc.author().is_none() {
        warnings.push("Missing document author".to_string());
    }

    report_validation(file, &errors, &warnings)
}

fn report_validation(file: &Path, errors: &[String], warnings: &[String]) -> Result<bool> {
    match print_validation_report(file, errors, warnings) {
        Err(error) if error.kind() != io::ErrorKind::BrokenPipe => Err(error.into()),
        _ => Ok(errors.is_empty()),
    }
}

fn print_validation_report(file: &Path, errors: &[String], warnings: &[String]) -> io::Result<()> {
    let mut stdout = io::stdout().lock();
    if errors.is_empty() && warnings.is_empty() {
        return writeln!(stdout, "OK — no issues found in {}", file.display());
    }

    if !errors.is_empty() {
        writeln!(stdout, "{} error(s) in {}:", errors.len(), file.display())?;
        for (i, issue) in errors.iter().enumerate() {
            writeln!(stdout, "  {}. {issue}", i + 1)?;
        }
    }
    if !warnings.is_empty() {
        writeln!(
            stdout,
            "{} warning(s) in {}:",
            warnings.len(),
            file.display()
        )?;
        for (i, issue) in warnings.iter().enumerate() {
            writeln!(stdout, "  {}. {issue}", i + 1)?;
        }
    }
    Ok(())
}

/// Read `xml` as one well-formed element tree and return the paragraph,
/// character, and table style ids it names. The ids inside a tracked property
/// change are left out, because they record the formatting before the change,
/// and so are those inside `mc:Fallback`, which Word does not read. An empty
/// id names no style.
fn xml_style_references(xml: &[u8]) -> std::result::Result<Vec<(&'static str, String)>, String> {
    let mut reader = NsReader::from_reader(xml);
    // One entry per open element: whether its style ids are left out.
    let mut open = Vec::new();
    let mut roots = 0usize;
    let mut references = Vec::new();
    let mut buffer = Vec::new();
    loop {
        let (namespace, event) = match reader.read_resolved_event_into(&mut buffer) {
            Ok(read) => read,
            Err(error) => return Err(format!("{error} at byte {}", reader.error_position())),
        };
        let is_word =
            matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == W_NS.as_bytes());
        let is_compatibility =
            matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == MC_NS.as_bytes());
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                if open.is_empty() {
                    roots += 1;
                    if roots > 1 {
                        return Err("the part has more than one root element".to_owned());
                    }
                }
                let local_name = element.local_name();
                let kind = match local_name.as_ref() {
                    b"pStyle" => Some("paragraph"),
                    b"rStyle" => Some("character"),
                    b"tblStyle" => Some("table"),
                    _ => None,
                };
                if let Some(kind) = kind.filter(|_| is_word && !open.contains(&true))
                    && let Some(style_id) = word_value(&reader, element)?
                    && !style_id.is_empty()
                {
                    references.push((kind, style_id));
                }
                if matches!(event, Event::Start(_)) {
                    open.push(
                        (is_word && local_name.as_ref().ends_with(b"PrChange"))
                            || (is_compatibility && local_name.as_ref() == b"Fallback"),
                    );
                }
            }
            Event::End(_) => {
                open.pop();
            }
            Event::Text(text) if open.is_empty() && !text.iter().all(u8::is_ascii_whitespace) => {
                return Err("the part has text outside its root element".to_owned());
            }
            Event::CData(_) if open.is_empty() => {
                return Err("the part has CDATA outside its root element".to_owned());
            }
            Event::Eof if roots == 0 => return Err("the part has no root element".to_owned()),
            Event::Eof if !open.is_empty() => {
                return Err(format!(
                    "the part ends inside {} unclosed element(s)",
                    open.len()
                ));
            }
            Event::Eof => return Ok(references),
            _ => {}
        }
        buffer.clear();
    }
}

/// Return the `w:val` attribute of a WordprocessingML element.
fn word_value(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> std::result::Result<Option<String>, String> {
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| error.to_string())?;
        let (namespace, local_name) = reader.resolver().resolve_attribute(attribute.key);
        if matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == W_NS.as_bytes())
            && local_name.as_ref() == b"val"
        {
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map_err(|error| error.to_string())?;
            return Ok(Some(value.into_owned()));
        }
    }
    Ok(None)
}

#[cfg(test)]
mod tests {
    use super::*;

    /// The Myers matches of `diff` are a longest common subsequence: equal,
    /// strictly increasing pairs as many as the quadratic table finds.
    #[test]
    fn diff_matches_are_a_longest_common_subsequence() {
        let mut seed = 0x2545_f491_u64;
        let mut next = |bound: u64| {
            seed = seed.wrapping_mul(6_364_136_223_846_793_005).wrapping_add(1);
            (seed >> 33) % bound
        };
        for _ in 0..2_000 {
            let alphabet = next(4) + 1;
            let a: Vec<u32> = (0..next(30)).map(|_| next(alphabet) as u32).collect();
            let b: Vec<u32> = (0..next(30)).map(|_| next(alphabet) as u32).collect();
            let mut matches = Vec::new();
            let mut frontiers = Frontiers::new(a.len() + b.len());
            matching_indexes(&a, 0..a.len(), &b, 0..b.len(), &mut frontiers, &mut matches);
            assert!(matches.iter().all(|&(i, j)| a[i] == b[j]), "{a:?} {b:?}");
            assert!(
                matches
                    .windows(2)
                    .all(|pair| pair[0].0 < pair[1].0 && pair[0].1 < pair[1].1),
                "{a:?} {b:?} {matches:?}"
            );
            let mut lengths = vec![vec![0; b.len() + 1]; a.len() + 1];
            for i in 1..=a.len() {
                for j in 1..=b.len() {
                    lengths[i][j] = if a[i - 1] == b[j - 1] {
                        lengths[i - 1][j - 1] + 1
                    } else {
                        lengths[i - 1][j].max(lengths[i][j - 1])
                    };
                }
            }
            assert_eq!(matches.len(), lengths[a.len()][b.len()], "{a:?} {b:?}");
        }
    }

    #[test]
    fn inspect_json_uses_the_shared_schema_one_envelope() {
        let document = Document::new();
        let styles = vec!["Heading1".to_owned(), "Normal".to_owned()];
        let value = inspect_json(Path::new("input.docx"), &document, styles.clone()).unwrap();

        assert_eq!(
            value,
            json!({
                "schema": 1,
                "file": "input.docx",
                "paragraphs": document.paragraph_count(),
                "tables": document.table_count(),
                "content_elements": document.content_count(),
                "metadata": {
                    "title": document.title(),
                    "author": document.author(),
                    "subject": document.subject(),
                    "keywords": document.keywords(),
                },
                "styles_used": styles,
            })
        );
    }

    #[test]
    fn convert_without_output_uses_the_default_extension_path() {
        let unique = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let input = std::env::temp_dir().join(format!(
            "rdocx_cli_default_convert_{}_{unique}.docx",
            std::process::id()
        ));
        let expected_output = input.with_extension("md");
        let mut document = Document::new();
        document.add_paragraph("Default path regression");
        document.save(&input).unwrap();

        convert(
            &input,
            "md",
            None,
            false,
            96,
            None,
            RevisionView::Accepted,
            ImageOptions {
                pages: None,
                quality: 90,
                transparent: false,
            },
        )
        .unwrap();

        let converted = std::fs::read_to_string(&expected_output);
        std::fs::remove_file(input).unwrap();
        std::fs::remove_file(expected_output).ok();
        assert_eq!(converted.unwrap(), "Default path regression\n\n");
    }
}
