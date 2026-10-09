//! CLI command implementations.

use std::collections::HashSet;
use std::io::{self, Write};
use std::path::{Path, PathBuf};

use oxml_cli_support::{
    ReplacementPair, StagedOutputSet, default_output_path, ensure_output_paths_allowed,
    ensure_output_paths_available, json_envelope, parse_range, parse_replacement_map,
};
use oxml_pdf::{RasterFormat, RasterOptions, RasterOutput};
use rpptx::{
    AutofitMode, Comment, CommentAuthor, CommentReply, CoreProperties, Presentation, ShapeKind,
    ShapeRef, SlideRef, TextFrameRef, TextParagraphRef, TextRunRef,
};
use serde_json::{Value, json};

type Result<T> = std::result::Result<T, Box<dyn std::error::Error>>;

const MAX_RASTER_PIXELS: u64 = 8_000_000;
const MAX_LCS_CELLS: usize = 1_000_000;
const THUMBNAIL_WIDTH: f64 = 320.0;

pub struct ImageOptions<'a> {
    pub slides: Option<&'a str>,
    pub quality: u8,
    pub transparent: bool,
}

pub fn inspect(file: &Path, as_json: bool) -> Result<()> {
    let presentation = Presentation::open(file)?;
    let size = presentation.slide_size();
    let core = presentation.core_properties();

    if as_json {
        let slides = presentation
            .slides()
            .map(|slide| {
                json!({
                    "id": slide.id(),
                    "name": slide.name(),
                    "hidden": slide.hidden(),
                    "shapes": slide.shapes().len(),
                    "shape_details": slide
                        .shapes()
                        .enumerate()
                        .map(|(index, shape)| shape_json(index, shape))
                        .collect::<Vec<_>>(),
                })
            })
            .collect::<Vec<_>>();
        let value = json_envelope(json!({
            "file": file.display().to_string(),
            "slides": presentation.len(),
            "layouts": presentation.layout_count(),
            "slide_size": size.map(|(width, height)| json!({
                "width_emu": width.0,
                "height_emu": height.0,
            })),
            "metadata": {
                "title": core.and_then(|value| value.title.as_deref()),
                "creator": core.and_then(|value| value.creator.as_deref()),
                "subject": core.and_then(|value| value.subject.as_deref()),
                "description": core.and_then(|value| value.description.as_deref()),
                "keywords": core.and_then(|value| value.keywords.as_deref()),
                "last_modified_by": core.and_then(|value| value.last_modified_by.as_deref()),
                "created": core.and_then(|value| value.created.as_deref()),
                "modified": core.and_then(|value| value.modified.as_deref()),
                "category": core.and_then(|value| value.category.as_deref()),
                "content_status": core.and_then(|value| value.content_status.as_deref()),
                "identifier": core.and_then(|value| value.identifier.as_deref()),
                "language": core.and_then(|value| value.language.as_deref()),
                "last_printed": core.and_then(|value| value.last_printed.as_deref()),
                "revision": core.and_then(|value| value.revision.as_deref()),
                "version": core.and_then(|value| value.version.as_deref()),
            },
            "slide_details": slides,
        }))?;
        writeln!(io::stdout(), "{}", serde_json::to_string_pretty(&value)?)?;
        return Ok(());
    }

    let mut stdout = io::stdout().lock();
    writeln!(stdout, "File: {}", file.display())?;
    writeln!(stdout, "Slides: {}", presentation.len())?;
    writeln!(stdout, "Layouts: {}", presentation.layout_count())?;
    if let Some((width, height)) = size {
        writeln!(stdout, "Slide size: {} x {} EMU", width.0, height.0)?;
    }
    writeln!(stdout)?;
    writeln!(stdout, "Metadata:")?;
    if let Some(core) = core {
        if let Some(value) = core.title.as_deref() {
            writeln!(stdout, "  Title: {value}")?;
        }
        if let Some(value) = core.creator.as_deref() {
            writeln!(stdout, "  Creator: {value}")?;
        }
        if let Some(value) = core.subject.as_deref() {
            writeln!(stdout, "  Subject: {value}")?;
        }
        if let Some(value) = core.description.as_deref() {
            writeln!(stdout, "  Description: {value}")?;
        }
        if let Some(value) = core.keywords.as_deref() {
            writeln!(stdout, "  Keywords: {value}")?;
        }
        if let Some(value) = core.last_modified_by.as_deref() {
            writeln!(stdout, "  Last modified by: {value}")?;
        }
        if let Some(value) = core.created.as_deref() {
            writeln!(stdout, "  Created: {value}")?;
        }
        if let Some(value) = core.modified.as_deref() {
            writeln!(stdout, "  Modified: {value}")?;
        }
        if let Some(value) = core.category.as_deref() {
            println!("  Category: {value}");
        }
        if let Some(value) = core.content_status.as_deref() {
            println!("  Content status: {value}");
        }
        if let Some(value) = core.identifier.as_deref() {
            println!("  Identifier: {value}");
        }
        if let Some(value) = core.language.as_deref() {
            println!("  Language: {value}");
        }
        if let Some(value) = core.last_printed.as_deref() {
            println!("  Last printed: {value}");
        }
        if let Some(value) = core.revision.as_deref() {
            println!("  Revision: {value}");
        }
        if let Some(value) = core.version.as_deref() {
            println!("  Version: {value}");
        }
        if core == &rpptx::CoreProperties::default() {
            writeln!(stdout, "  (none)")?;
        }
    } else {
        writeln!(stdout, "  (none)")?;
    }
    writeln!(stdout)?;
    for (index, slide) in presentation.slides().enumerate() {
        writeln!(
            stdout,
            "Slide {}: id={}, name={}, hidden={}, shapes={}",
            index + 1,
            slide.id(),
            slide.name().unwrap_or(""),
            slide.hidden(),
            slide.shapes().len()
        )?;
    }
    Ok(())
}

pub fn text(file: &Path, as_json: bool, notes: bool) -> Result<()> {
    let presentation = Presentation::open(file)?;
    if as_json {
        let slides = presentation
            .slides()
            .enumerate()
            .map(|(index, slide)| {
                let mut paragraphs = Vec::new();
                for (shape_index, shape) in slide.shapes().enumerate() {
                    let path = [path_segment("shape", shape_index)];
                    collect_shape_paragraphs(shape, &path, &mut paragraphs);
                }
                json!({
                    "slide": index + 1,
                    "id": slide.id(),
                    "paragraphs": paragraphs,
                    "notes": slide.notes_text(),
                })
            })
            .collect::<Vec<_>>();
        return print_json(json!({ "slides": slides }));
    }
    let mut stdout = io::stdout().lock();
    for slide in presentation.slides() {
        writeln!(stdout, "{}", slide.text())?;
        if notes {
            print_notes(slide)?;
        }
    }
    Ok(())
}

fn path_segment(kind: &str, index: usize) -> Value {
    json!({ "kind": kind, "index": index })
}

fn collect_shape_paragraphs(shape: ShapeRef<'_>, path: &[Value], output: &mut Vec<Value>) {
    if let Some(frame) = shape.text_frame() {
        collect_frame_paragraphs(shape, frame, path, output);
    }
    if let Some(table) = shape.table() {
        for row in 0..table.row_count() {
            for column in 0..table.column_count() {
                let Some(cell) = table.cell(row, column) else {
                    continue;
                };
                if let Some(frame) = cell.text_frame() {
                    let mut cell_path = path.to_vec();
                    cell_path.push(path_segment("row", row));
                    cell_path.push(path_segment("cell", column));
                    collect_frame_paragraphs(shape, frame, &cell_path, output);
                }
            }
        }
    }
    for (index, child) in shape.children().enumerate() {
        let mut child_path = path.to_vec();
        child_path.push(path_segment("shape", index));
        collect_shape_paragraphs(child, &child_path, output);
    }
}

fn collect_frame_paragraphs(
    shape: ShapeRef<'_>,
    frame: TextFrameRef<'_>,
    path: &[Value],
    output: &mut Vec<Value>,
) {
    for index in 0..frame.paragraph_count() {
        let Some(paragraph) = frame.paragraph(index) else {
            continue;
        };
        let mut paragraph_path = path.to_vec();
        paragraph_path.push(path_segment("paragraph", index));
        let mut value = paragraph_json(paragraph);
        value["path"] = Value::from(paragraph_path);
        value["shape_id"] = json!(shape.non_visual_id());
        output.push(value);
    }
}

fn shape_json(index: usize, shape: ShapeRef<'_>) -> Value {
    let frame = shape.text_frame();
    json!({
        "index": index,
        "id": shape.non_visual_id(),
        "name": shape.non_visual_name(),
        "kind": shape_kind_label(shape.kind()),
        "placeholder": shape.placeholder_idx().map(|idx| json!({
            "type": shape.placeholder_type(),
            "idx": idx,
        })),
        "position": shape.position().map(|(left, top)| json!({
            "left_emu": left.0,
            "top_emu": top.0,
        })),
        "size": shape.size().map(|(width, height)| json!({
            "width_emu": width.0,
            "height_emu": height.0,
        })),
        "rotation_degrees": shape.rotation().map(|rotation| rotation.to_degrees()),
        "autofit": frame.and_then(|frame| frame.autofit_mode()).map(autofit_label),
        "paragraphs": frame.map(|frame| {
            (0..frame.paragraph_count())
                .filter_map(|index| {
                    let mut value = paragraph_json(frame.paragraph(index)?);
                    value["index"] = json!(index);
                    Some(value)
                })
                .collect::<Vec<_>>()
        }),
        "table": shape.table().map(|table| json!({
            "rows": table.row_count(),
            "columns": table.column_count(),
        })),
        "children": shape
            .children()
            .enumerate()
            .map(|(index, child)| shape_json(index, child))
            .collect::<Vec<_>>(),
    })
}

fn paragraph_json(paragraph: TextParagraphRef<'_>) -> Value {
    let runs = (0..paragraph.run_count())
        .filter_map(|index| {
            let run = paragraph.run(index)?;
            Some(json!({
                "index": index,
                "text": run.text(),
                "formatting": run_formatting_json(run),
            }))
        })
        .collect::<Vec<_>>();
    json!({
        "level": paragraph.level(),
        "text": paragraph.text(),
        "runs": runs,
    })
}

fn run_formatting_json(run: TextRunRef<'_>) -> Value {
    let Some(properties) = run.properties() else {
        return Value::Null;
    };
    json!({
        "bold": properties.bold,
        "italic": properties.italic,
        "underline": properties.underline.map(|underline| underline.as_str()),
        "font": run.font_name(),
        "size_points": run.font_size().map(|size| f64::from(size) / 100.0),
        "color": run.font_color(),
    })
}

fn shape_kind_label(kind: ShapeKind) -> &'static str {
    match kind {
        ShapeKind::Shape => "shape",
        ShapeKind::Picture => "picture",
        ShapeKind::GraphicFrame => "graphic-frame",
        ShapeKind::Group => "group",
        ShapeKind::Connector => "connector",
        ShapeKind::AlternateContent => "alternate-content",
    }
}

fn autofit_label(mode: AutofitMode) -> &'static str {
    match mode {
        AutofitMode::None => "none",
        AutofitMode::Normal => "normal",
        AutofitMode::Shape => "shape",
    }
}

fn print_notes(slide: SlideRef<'_>) -> io::Result<()> {
    let mut stdout = io::stdout().lock();
    for line in slide.notes_text().unwrap_or_default().lines() {
        let line = normalize_outline_text(line);
        if !line.is_empty() {
            writeln!(stdout, "Notes: {line}")?;
        }
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

pub fn convert(
    file: &Path,
    format: &str,
    output: Option<&Path>,
    force: bool,
    dpi: f64,
    image: ImageOptions<'_>,
) -> Result<()> {
    validate_dpi(dpi)?;
    let presentation = Presentation::open(file)?;
    let image_format = parse_image_format(format, image.quality, image.transparent);
    let extension = match (format, image_format.as_ref()) {
        ("pdf", _) => "pdf",
        (_, Ok((_, extension))) => extension,
        (other, Err(_)) => {
            return Err(format!("Unknown format: {other}. Supported: pdf, png, jpeg, tiff").into());
        }
    };
    let output = output
        .map(Path::to_path_buf)
        .unwrap_or_else(|| default_output_path(file, extension));
    // Several PNG or JPEG slides are written under numbered names, so those
    // outputs are checked once the slides are selected.
    if !matches!(format, "png" | "jpg" | "jpeg") {
        ensure_output_paths_allowed(std::slice::from_ref(&output), file, force)?;
    }
    let mut stdout = io::stdout().lock();
    match format {
        "pdf" => {
            stage_and_publish(
                &[(output.clone(), presentation.to_pdf_deterministic()?)],
                force,
            )?;
            writeln!(stdout, "Written to {}", output.display())?;
        }
        "png" | "jpg" | "jpeg" | "tif" | "tiff" => {
            if presentation.is_empty() {
                return Err("cannot convert a presentation with no slides to an image".into());
            }
            let (_, layout) = presentation.render_deterministic()?;
            if layout.pages.len() != presentation.len() {
                return Err(format!(
                    "rendered {} image pages for {} slides",
                    layout.pages.len(),
                    presentation.len()
                )
                .into());
            }
            let selected = selected_zero_based_slides(presentation.len(), image.slides)?;
            for index in &selected {
                let page = &layout.pages[*index];
                validate_raster_dimensions(page.width, page.height, dpi)?;
            }
            let (format, extension) = image_format?;
            match format {
                RasterFormat::Tiff => {
                    let output_bytes =
                        oxml_pdf::render_pages(&layout, &selected, RasterOptions { dpi, format })?;
                    let RasterOutput::MultiPageTiff(tiff) = output_bytes else {
                        return Err("TIFF render did not produce one stream".into());
                    };
                    stage_and_publish(&[(output.clone(), tiff)], force)?;
                    writeln!(stdout, "Written to {}", output.display())?;
                }
                RasterFormat::Png { .. } | RasterFormat::Jpeg { .. } => {
                    let output_paths =
                        convert_separate_output_paths(&output, extension, selected.len());
                    ensure_output_paths_allowed(&output_paths, file, force)?;
                    let mut staged = StagedOutputSet::with_replace_existing(force);
                    let mut rendered = Vec::with_capacity(selected.len());
                    for (one_based, (index, path)) in
                        selected.iter().zip(output_paths.iter()).enumerate()
                    {
                        let image =
                            render_one_raster_page(&layout, *index, RasterOptions { dpi, format })?;
                        staged.stage_bytes(path, &image)?;
                        rendered.push((one_based + 1, path.clone()));
                    }
                    staged.publish()?;
                    for (one_based, path) in rendered {
                        writeln!(stdout, "Slide {one_based} -> {}", path.display())?;
                    }
                }
            }
        }
        _ => unreachable!(),
    }
    Ok(())
}

pub fn diff(file_a: &Path, file_b: &Path) -> Result<()> {
    let presentation_a = Presentation::open(file_a)?;
    let presentation_b = Presentation::open(file_b)?;
    let text_a = presentation_a
        .slides()
        .map(|slide| slide.text())
        .collect::<Vec<_>>();
    let text_b = presentation_b
        .slides()
        .map(|slide| slide.text())
        .collect::<Vec<_>>();
    let common = longest_common_subsequence(&text_a, &text_b)?;
    let mut a = 0;
    let mut b = 0;
    let mut stdout = io::stdout().lock();
    for value in common {
        while text_a.get(a) != Some(&value) {
            writeln!(stdout, "- [{}] {}", a + 1, text_a[a])?;
            a += 1;
        }
        while text_b.get(b) != Some(&value) {
            writeln!(stdout, "+ [{}] {}", b + 1, text_b[b])?;
            b += 1;
        }
        a += 1;
        b += 1;
    }
    while a < text_a.len() {
        writeln!(stdout, "- [{}] {}", a + 1, text_a[a])?;
        a += 1;
    }
    while b < text_b.len() {
        writeln!(stdout, "+ [{}] {}", b + 1, text_b[b])?;
        b += 1;
    }
    Ok(())
}

fn longest_common_subsequence(a: &[String], b: &[String]) -> Result<Vec<String>> {
    let rows = a
        .len()
        .checked_add(1)
        .ok_or("first slide count is too large to diff")?;
    let columns = b
        .len()
        .checked_add(1)
        .ok_or("second slide count is too large to diff")?;
    let cells = rows
        .checked_mul(columns)
        .ok_or("slide diff LCS matrix size overflowed")?;
    if cells > MAX_LCS_CELLS {
        return Err(format!(
            "slide diff requires {cells} LCS cells, which exceeds the 1,000,000 LCS cell limit"
        )
        .into());
    }
    let mut lengths = vec![vec![0usize; b.len() + 1]; a.len() + 1];
    for a_index in 1..=a.len() {
        for b_index in 1..=b.len() {
            lengths[a_index][b_index] = if a[a_index - 1] == b[b_index - 1] {
                lengths[a_index - 1][b_index - 1] + 1
            } else {
                lengths[a_index - 1][b_index].max(lengths[a_index][b_index - 1])
            };
        }
    }
    let mut result = Vec::new();
    let mut a_index = a.len();
    let mut b_index = b.len();
    while a_index > 0 && b_index > 0 {
        if a[a_index - 1] == b[b_index - 1] {
            result.push(a[a_index - 1].clone());
            a_index -= 1;
            b_index -= 1;
        } else if lengths[a_index - 1][b_index] >= lengths[a_index][b_index - 1] {
            a_index -= 1;
        } else {
            b_index -= 1;
        }
    }
    result.reverse();
    Ok(result)
}

/// Replaces one pair, or the pairs of a map in order, and publishes only
/// when every pair found its expected count, or at least one occurrence
/// when it gives none.
pub fn replace(
    file: &Path,
    pair: Option<(String, String)>,
    map: Option<&Path>,
    expect: Option<usize>,
    output: &Path,
    as_json: bool,
) -> Result<()> {
    ensure_output_paths_available(&[output.to_path_buf()])?;
    let pairs = match (pair, map) {
        (Some((placeholder, value)), None) => vec![ReplacementPair {
            placeholder,
            value,
            expect,
        }],
        (None, Some(map)) => {
            let text = std::fs::read_to_string(map)
                .map_err(|error| format!("cannot read --map {}: {error}", map.display()))?;
            parse_replacement_map(&text)?
        }
        _ => return Err("give -p and -v, or --map".into()),
    };
    let mut presentation = Presentation::open(file)?;
    let mut counts = Vec::with_capacity(pairs.len());
    for (index, pair) in pairs.iter().enumerate() {
        let count = presentation.try_replace_text(&pair.placeholder, &pair.value)?;
        let prefix = if pairs.len() == 1 {
            String::new()
        } else {
            format!("pair {index}: ")
        };
        if let Some(expected) = pair.expect
            && count != expected
        {
            return Err(format!(
                "{prefix}expected {expected} replacement(s) of \"{}\", found {count}, \
                 nothing written",
                pair.placeholder
            )
            .into());
        }
        if count == 0 && pair.expect.is_none() {
            return Err(format!(
                "{prefix}no replacements found for \"{}\", nothing written",
                pair.placeholder
            )
            .into());
        }
        counts.push(count);
    }
    publish_presentation(&presentation, output)?;
    if as_json {
        let entries = pairs
            .iter()
            .zip(&counts)
            .map(|(pair, count)| {
                json!({
                    "placeholder": pair.placeholder,
                    "value": pair.value,
                    "expect": pair.expect,
                    "count": count,
                })
            })
            .collect::<Vec<_>>();
        return mutation_record(
            true,
            "replace",
            json!({ "pairs": entries, "total": counts.iter().sum::<usize>() }),
            output,
        );
    }
    let mut stdout = io::stdout().lock();
    for (pair, count) in pairs.iter().zip(&counts) {
        writeln!(
            stdout,
            "Replaced {count} occurrence(s) of \"{}\" -> \"{}\"",
            pair.placeholder, pair.value
        )?;
    }
    writeln!(stdout, "Written to {}", output.display())?;
    Ok(())
}

pub fn validate(file: &Path) -> Result<bool> {
    let presentation = Presentation::open(file)?;
    let issues = presentation.validate();
    for issue in &issues {
        eprintln!("{issue:?}");
    }
    if issues.is_empty() {
        // A reader that closes standard output early, as `| head` does, does
        // not change the verdict.
        if let Err(error) = writeln!(io::stdout(), "Validation passed: {}", file.display())
            && error.kind() != io::ErrorKind::BrokenPipe
        {
            return Err(error.into());
        }
    } else {
        eprintln!("Validation failed with {} issue(s)", issues.len());
    }
    Ok(issues.is_empty())
}

pub fn render(
    file: &Path,
    output: Option<&Path>,
    force: bool,
    dpi: f64,
    format: &str,
    image: ImageOptions<'_>,
) -> Result<()> {
    validate_dpi(dpi)?;
    let presentation = Presentation::open(file)?;
    let (format, extension) = parse_image_format(format, image.quality, image.transparent)?;
    let selected = selected_zero_based_slides(presentation.len(), image.slides)?;
    let output = output.unwrap_or_else(|| Path::new("."));
    let stem = file.file_stem().unwrap_or_default().to_string_lossy();
    let (_, layout) = presentation.render_deterministic()?;
    for index in &selected {
        let page = layout
            .pages
            .get(*index)
            .ok_or_else(|| format!("slide {} has no rendered page", index + 1))?;
        validate_raster_dimensions(page.width, page.height, dpi)?;
    }
    let mut stdout = io::stdout().lock();
    match format {
        RasterFormat::Tiff => {
            let path = output.join(format!("{stem}.tiff"));
            ensure_output_paths_allowed(std::slice::from_ref(&path), file, force)?;
            let output_bytes =
                oxml_pdf::render_pages(&layout, &selected, RasterOptions { dpi, format })?;
            let RasterOutput::MultiPageTiff(tiff) = output_bytes else {
                return Err("TIFF render did not produce one stream".into());
            };
            std::fs::create_dir_all(output)?;
            stage_and_publish(&[(path.clone(), tiff)], force)?;
            writeln!(stdout, "Written to {}", path.display())?;
        }
        RasterFormat::Png { .. } | RasterFormat::Jpeg { .. } => {
            let output_paths = selected
                .iter()
                .map(|index| {
                    let one_based = index + 1;
                    output.join(format!("{stem}_slide{one_based}.{extension}"))
                })
                .collect::<Vec<_>>();
            std::fs::create_dir_all(output)?;
            ensure_output_paths_allowed(&output_paths, file, force)?;
            let mut staged = StagedOutputSet::with_replace_existing(force);
            let mut rendered = Vec::with_capacity(selected.len());
            for (index, path) in selected.iter().zip(output_paths.iter()) {
                let one_based = index + 1;
                let image = render_one_raster_page(&layout, *index, RasterOptions { dpi, format })?;
                staged.stage_bytes(path, &image)?;
                rendered.push((one_based, path.clone()));
            }
            staged.publish()?;
            for (one_based, path) in rendered {
                writeln!(stdout, "Slide {one_based} -> {}", path.display())?;
            }
        }
    }
    Ok(())
}

pub fn thumbnail(file: &Path, output: Option<&Path>, force: bool) -> Result<()> {
    let presentation = Presentation::open(file)?;
    if presentation.is_empty() {
        return Err("cannot thumbnail a presentation with no slides".into());
    }
    let output = output
        .map(Path::to_path_buf)
        .unwrap_or_else(|| default_output_path(file, "png"));
    ensure_output_paths_allowed(std::slice::from_ref(&output), file, force)?;
    let (_, layout) = presentation.render_deterministic()?;
    let page = layout
        .pages
        .first()
        .ok_or("slide one has no rendered page")?;
    if !page.width.is_finite() || page.width <= 0.0 {
        return Err("slide one has an invalid rendered width".into());
    }
    let dpi = (THUMBNAIL_WIDTH - 1e-9) * 72.0 / page.width;
    validate_raster_dimensions(page.width, page.height, dpi)?;
    let png = oxml_pdf::render_page_to_png(&layout, 0, dpi)
        .ok_or("slide one did not rasterize for thumbnail")?;
    stage_and_publish(&[(output.clone(), png)], force)?;
    writeln!(io::stdout(), "Written to {}", output.display())?;
    Ok(())
}

pub fn outline(file: &Path, as_json: bool, notes: bool) -> Result<()> {
    let presentation = Presentation::open(file)?;
    let mut slides = Vec::new();
    let mut stdout = io::stdout().lock();
    for (index, slide) in presentation.slides().enumerate() {
        let title_shape = slide.title();
        let title = normalize_outline_text(
            &title_shape
                .and_then(|shape| shape.text())
                .unwrap_or_default(),
        );
        let mut items = Vec::new();
        for shape in slide.shapes() {
            if title_shape == Some(shape) {
                continue;
            }
            collect_shape_outline(shape, &mut items);
        }
        if as_json {
            slides.push(json!({
                "slide": index + 1,
                "id": slide.id(),
                "title": (!title.is_empty()).then_some(title),
                "items": items
                    .into_iter()
                    .map(|(level, text)| json!({ "level": level, "text": text }))
                    .collect::<Vec<_>>(),
                "notes": slide.notes_text(),
            }));
            continue;
        }
        if title.is_empty() {
            writeln!(stdout, "Slide {}", index + 1)?;
        } else {
            writeln!(stdout, "Slide {}: {title}", index + 1)?;
        }
        for (level, text) in items {
            writeln!(stdout, "{}- {text}", "  ".repeat(level as usize))?;
        }
        if notes {
            print_notes(slide)?;
        }
    }
    if as_json {
        print_json(json!({ "slides": slides }))?;
    }
    Ok(())
}

fn collect_shape_outline(shape: ShapeRef<'_>, items: &mut Vec<(u8, String)>) {
    if let Some(frame) = shape.text_frame() {
        collect_frame_outline(frame, items);
    }
    if let Some(table) = shape.table() {
        for row in 0..table.row_count() {
            for column in 0..table.column_count() {
                let Some(cell) = table.cell(row, column) else {
                    continue;
                };
                if cell.is_spanned() {
                    continue;
                }
                if let Some(frame) = cell.text_frame() {
                    collect_frame_outline(frame, items);
                }
            }
        }
    }
    for child in shape.children() {
        collect_shape_outline(child, items);
    }
}

fn collect_frame_outline(frame: TextFrameRef<'_>, items: &mut Vec<(u8, String)>) {
    for index in 0..frame.paragraph_count() {
        let Some(paragraph) = frame.paragraph(index) else {
            continue;
        };
        let text = normalize_outline_text(&paragraph.text());
        if !text.is_empty() {
            items.push((paragraph.level(), text));
        }
    }
}

fn normalize_outline_text(text: &str) -> String {
    text.trim().replace(['\r', '\n', '\u{000b}'], " ")
}

/// Author, text, and creation time of a new comment or reply.
pub struct CommentInput<'a> {
    pub author: &'a str,
    pub initials: Option<&'a str>,
    pub text: &'a str,
    pub date: &'a str,
}

/// One modern comment or reply, flattened in slide and thread order.
struct CommentEntry<'a> {
    slide: usize,
    id: &'a str,
    parent_id: Option<&'a str>,
    author_id: &'a str,
    date: &'a str,
    status: Option<&'a str>,
    text: String,
}

/// Lists modern comments, each followed by its replies, in slide order.
pub fn comment_list(file: &Path, as_json: bool) -> Result<()> {
    let presentation = Presentation::open(file)?;
    let entries = comment_entries(&presentation);
    if as_json {
        let records = entries
            .iter()
            .map(|entry| {
                let author = comment_author(&presentation, entry.author_id);
                json!({
                    "slide": entry.slide,
                    "id": entry.id,
                    "author": author.map(|author| author.name.as_str()),
                    "initials": author.and_then(|author| author.initials.as_deref()),
                    "date": entry.date,
                    "text": entry.text,
                    "parent_id": entry.parent_id,
                    "resolved": entry.status == Some("resolved"),
                    "status": entry.status,
                })
            })
            .collect::<Vec<_>>();
        return print_json(json!({ "comments": records }));
    }
    let mut stdout = io::stdout().lock();
    if entries.is_empty() {
        writeln!(stdout, "(no comments)")?;
    }
    for entry in &entries {
        writeln!(
            stdout,
            "{}\t{}\t{}\t{}\t{}",
            entry.slide,
            entry.id,
            comment_author(&presentation, entry.author_id).map_or("", |author| &author.name),
            entry
                .status
                .filter(|status| *status != "active")
                .unwrap_or("open"),
            entry.text.replace('\n', " ")
        )?;
    }
    Ok(())
}

/// Adds one modern comment and publishes the complete presentation atomically.
pub fn comment_add(
    file: &Path,
    slide: usize,
    input: CommentInput<'_>,
    output: &Path,
    as_json: bool,
) -> Result<()> {
    ensure_output_paths_available(&[output.to_path_buf()])?;
    let mut presentation = Presentation::open(file)?;
    if slide == 0 || slide > presentation.len() {
        return Err(format!(
            "slide {slide} is out of range for {} slides",
            presentation.len()
        )
        .into());
    }
    let author_id = comment_author_id(&mut presentation, input.author, input.initials)?;
    let id = next_comment_guid(&presentation);
    let comment = Comment::new(id.as_str(), author_id, input.date, input.text)?;
    presentation.add_comment(slide - 1, comment)?;
    publish_presentation(&presentation, output)?;
    mutation_record(
        as_json,
        "add",
        json!({ "comment_id": id, "slide": slide }),
        output,
    )
}

/// Adds one reply to a thread and publishes the complete presentation atomically.
pub fn comment_reply(
    file: &Path,
    parent_id: &str,
    input: CommentInput<'_>,
    output: &Path,
    as_json: bool,
) -> Result<()> {
    ensure_output_paths_available(&[output.to_path_buf()])?;
    let mut presentation = Presentation::open(file)?;
    let slide = thread_slide(&presentation, parent_id)?;
    let author_id = comment_author_id(&mut presentation, input.author, input.initials)?;
    let id = next_comment_guid(&presentation);
    let reply = CommentReply::new(id.as_str(), author_id, input.date, input.text)?;
    presentation.reply_to_comment(slide - 1, parent_id, reply)?;
    publish_presentation(&presentation, output)?;
    mutation_record(
        as_json,
        "reply",
        json!({ "comment_id": id, "parent_id": parent_id, "slide": slide }),
        output,
    )
}

/// Resolves one comment thread and publishes the complete presentation atomically.
pub fn comment_resolve(file: &Path, id: &str, output: &Path, as_json: bool) -> Result<()> {
    ensure_output_paths_available(&[output.to_path_buf()])?;
    let mut presentation = Presentation::open(file)?;
    let slide = thread_slide(&presentation, id)?;
    presentation.resolve_comment(slide - 1, id)?;
    publish_presentation(&presentation, output)?;
    mutation_record(
        as_json,
        "resolve",
        json!({ "comment_id": id, "slide": slide }),
        output,
    )
}

/// Removes one thread or reply and publishes the complete presentation atomically.
pub fn comment_remove(file: &Path, id: &str, output: &Path, as_json: bool) -> Result<()> {
    ensure_output_paths_available(&[output.to_path_buf()])?;
    let mut presentation = Presentation::open(file)?;
    let (slide, _) = locate_comment(&presentation, id)?;
    presentation.remove_comment(slide - 1, id)?;
    publish_presentation(&presentation, output)?;
    mutation_record(
        as_json,
        "remove",
        json!({ "comment_id": id, "slide": slide }),
        output,
    )
}

fn comment_entries(presentation: &Presentation) -> Vec<CommentEntry<'_>> {
    let mut entries = Vec::new();
    for slide_index in 0..presentation.len() {
        for comment in presentation.comments(slide_index).unwrap_or_default() {
            entries.push(CommentEntry {
                slide: slide_index + 1,
                id: &comment.id,
                parent_id: None,
                author_id: &comment.author_id,
                date: &comment.created,
                status: comment.status.as_deref(),
                text: comment.text(),
            });
            for reply in comment.replies() {
                entries.push(CommentEntry {
                    slide: slide_index + 1,
                    id: &reply.id,
                    parent_id: Some(&comment.id),
                    author_id: &reply.author_id,
                    date: &reply.created,
                    status: reply.status.as_deref(),
                    text: reply.text(),
                });
            }
        }
    }
    entries
}

fn comment_author<'a>(
    presentation: &'a Presentation,
    author_id: &str,
) -> Option<&'a CommentAuthor> {
    presentation
        .comment_authors()
        .iter()
        .find(|author| author.id == author_id)
}

/// Returns the one-based slide and the thread id of the comment or reply `id`.
fn locate_comment<'a>(
    presentation: &'a Presentation,
    id: &str,
) -> Result<(usize, Option<&'a str>)> {
    comment_entries(presentation)
        .into_iter()
        .find(|entry| entry.id == id)
        .map(|entry| (entry.slide, entry.parent_id))
        .ok_or_else(|| format!("comment id {id} does not exist").into())
}

/// Returns the one-based slide of the thread whose top-level comment is `id`.
fn thread_slide(presentation: &Presentation, id: &str) -> Result<usize> {
    match locate_comment(presentation, id)? {
        (slide, None) => Ok(slide),
        (_, Some(thread)) => Err(format!("comment id {id} is a reply in thread {thread}").into()),
    }
}

/// Returns the id of the author named `name`, adding that author when absent.
///
/// A new author records the display name as `userId` and `None` as
/// `providerId`, as PowerPoint does for an author without an online account.
fn comment_author_id(
    presentation: &mut Presentation,
    name: &str,
    initials: Option<&str>,
) -> Result<String> {
    if let Some(author) = presentation
        .comment_authors()
        .iter()
        .find(|author| author.name == name)
    {
        return Ok(author.id.clone());
    }
    let id = next_comment_guid(presentation);
    presentation.add_comment_author(CommentAuthor::new(
        id.as_str(),
        name,
        initials,
        name,
        "None",
    )?)?;
    Ok(id)
}

/// Returns the first sequential GUID no comment author, comment, or reply uses.
///
/// Sequential ids keep command output reproducible without a clock or a
/// random source.
fn next_comment_guid(presentation: &Presentation) -> String {
    let used = presentation
        .comment_authors()
        .iter()
        .map(|author| author.id.to_ascii_uppercase())
        .chain(
            comment_entries(presentation)
                .into_iter()
                .map(|entry| entry.id.to_ascii_uppercase()),
        )
        .collect::<HashSet<_>>();
    (1_u64..)
        .map(|index| format!("{{00000000-0000-4000-8000-{index:012X}}}"))
        .find(|id| !used.contains(id))
        .expect("a finite set of used ids leaves a sequential id free")
}

fn publish_presentation(presentation: &Presentation, output: &Path) -> Result<()> {
    stage_and_publish(
        &[(
            output.to_path_buf(),
            presentation.to_bytes_for_path(output)?,
        )],
        false,
    )
}

/// Prints a schema-1 operation record, or the action and the output path.
fn mutation_record(as_json: bool, action: &str, mut record: Value, output: &Path) -> Result<()> {
    if !as_json {
        let mut stdout = io::stdout().lock();
        writeln!(stdout, "{action}")?;
        writeln!(stdout, "Written to {}", output.display())?;
        return Ok(());
    }
    record["action"] = json!(action);
    record["output"] = json!(output.display().to_string());
    print_json(record)
}

fn validate_dpi(dpi: f64) -> Result<()> {
    if !dpi.is_finite() || dpi <= 0.0 {
        return Err("DPI must be a positive finite number".into());
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

fn selected_zero_based_slides(slide_count: usize, range: Option<&str>) -> Result<Vec<usize>> {
    let selected = match range {
        Some(range) => parse_range(range)?,
        None => (1..=slide_count).collect(),
    };
    if selected.is_empty() {
        return Err("no slides selected".into());
    }
    if let Some(index) = selected.iter().copied().find(|index| *index > slide_count) {
        return Err(format!("slide {index} is out of range for {slide_count} slides").into());
    }
    Ok(selected.into_iter().map(|index| index - 1).collect())
}

fn convert_separate_output_paths(
    output: &Path,
    extension: &str,
    page_count: usize,
) -> Vec<PathBuf> {
    if page_count == 1 {
        return vec![output.to_path_buf()];
    }
    let parent = output.parent().unwrap_or_else(|| Path::new("."));
    let stem = output.file_stem().unwrap_or_default().to_string_lossy();
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
        return Err("single-slide render did not produce a separate image".into());
    };
    if pages.len() != 1 {
        return Err("single-slide render returned the wrong page count".into());
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

fn validate_raster_dimensions(width_points: f64, height_points: f64, dpi: f64) -> Result<()> {
    let scale = dpi / 72.0;
    let width = (width_points * scale).ceil();
    let height = (height_points * scale).ceil();
    if !width.is_finite()
        || !height.is_finite()
        || width < 1.0
        || height < 1.0
        || width > u32::MAX as f64
        || height > u32::MAX as f64
    {
        return Err(format!("DPI {dpi} produces invalid raster dimensions").into());
    }
    let width = width as u64;
    let height = height as u64;
    let pixels = width
        .checked_mul(height)
        .ok_or("raster pixel count overflowed")?;
    if pixels > MAX_RASTER_PIXELS {
        return Err(format!(
            "DPI {dpi} produces a {width} x {height} raster, which exceeds the 8,000,000 pixel limit"
        )
        .into());
    }
    Ok(())
}

// ---- One-shot editing commands: slides, notes, fit, metadata ----

/// Returns the zero-based index of one-based slide `slide`, or an error that
/// names the slide count.
fn slide_index(presentation: &Presentation, slide: usize) -> Result<usize> {
    if slide == 0 || slide > presentation.len() {
        return Err(format!(
            "slide {slide} is out of range for {} slides",
            presentation.len()
        )
        .into());
    }
    Ok(slide - 1)
}

/// Lists the layouts as `N "name"` for an error message.
fn layout_list(presentation: &Presentation) -> String {
    (0..presentation.layout_count())
        .map(|index| {
            format!(
                "{} \"{}\"",
                index + 1,
                presentation.layout_name(index).unwrap_or_default()
            )
        })
        .collect::<Vec<_>>()
        .join(", ")
}

/// Resolves a layout by exact name, then by case-insensitive name, then by
/// one-based number, refusing a name that several masters share.
fn resolve_layout(presentation: &Presentation, layout: &str) -> Result<usize> {
    let names = (0..presentation.layout_count())
        .map(|index| presentation.layout_name(index).unwrap_or_default())
        .collect::<Vec<_>>();
    for exact in [true, false] {
        let found = names
            .iter()
            .enumerate()
            .filter(|(_, name)| {
                if exact {
                    **name == layout
                } else {
                    name.eq_ignore_ascii_case(layout)
                }
            })
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        match found.as_slice() {
            [index] => return Ok(*index),
            [] => {}
            several => {
                return Err(format!(
                    "layout name \"{layout}\" is shared by layouts {}, pass its number",
                    several
                        .iter()
                        .map(|index| (index + 1).to_string())
                        .collect::<Vec<_>>()
                        .join(", ")
                )
                .into());
            }
        }
    }
    match layout.parse::<usize>() {
        Ok(number) if (1..=names.len()).contains(&number) => Ok(number - 1),
        _ => Err(format!(
            "no layout is named or numbered \"{layout}\" (layouts: {})",
            layout_list(presentation)
        )
        .into()),
    }
}

/// Adds one slide from a layout, then moves it to `at` when given.
pub fn slide_add(
    file: &Path,
    layout: &str,
    at: Option<usize>,
    output: &Path,
    as_json: bool,
) -> Result<()> {
    ensure_output_paths_available(&[output.to_path_buf()])?;
    let mut presentation = Presentation::open(file)?;
    let layout_index = resolve_layout(&presentation, layout)?;
    let final_count = presentation.len() + 1;
    let position = at.unwrap_or(final_count);
    if position == 0 || position > final_count {
        return Err(format!(
            "--at {position} is out of range, the new slide can go at 1 to {final_count}"
        )
        .into());
    }
    presentation.add_slide(layout_index)?;
    presentation.move_slide(final_count - 1, position - 1)?;
    publish_presentation(&presentation, output)?;
    mutation_record(
        as_json,
        "add",
        json!({
            "slide": position,
            "slides": presentation.len(),
            "layout": layout_index + 1,
            "layout_name": presentation.layout_name(layout_index),
        }),
        output,
    )
}

/// One slide edit, with one-based slide numbers.
pub enum SlideEdit {
    Duplicate(usize),
    Remove(usize),
    Move(usize, usize),
    Hidden(usize, bool),
}

/// Applies one slide edit and publishes the complete presentation.
pub fn slide_edit(file: &Path, edit: SlideEdit, output: &Path, as_json: bool) -> Result<()> {
    ensure_output_paths_available(&[output.to_path_buf()])?;
    let mut presentation = Presentation::open(file)?;
    // `changed` is false for an edit that leaves the deck as it was, such as
    // moving a slide to its own position or hiding a hidden slide.
    let (action, slide, changed) = match edit {
        SlideEdit::Duplicate(slide) => {
            let index = slide_index(&presentation, slide)?;
            presentation.duplicate_slide(index)?;
            ("duplicate", slide + 1, true)
        }
        SlideEdit::Remove(slide) => {
            let index = slide_index(&presentation, slide)?;
            presentation.remove_slide(index)?;
            ("remove", slide, true)
        }
        SlideEdit::Move(slide, to) => {
            let index = slide_index(&presentation, slide)?;
            let to_index = slide_index(&presentation, to).map_err(|_| {
                format!(
                    "--to {to} is out of range for {} slides",
                    presentation.len()
                )
            })?;
            presentation.move_slide(index, to_index)?;
            ("move", to, index != to_index)
        }
        SlideEdit::Hidden(slide, hidden) => {
            let index = slide_index(&presentation, slide)?;
            let was_hidden = presentation
                .slide(index)
                .ok_or("slide index was checked")?
                .hidden();
            presentation
                .slide_mut(index)
                .ok_or("slide index was checked")?
                .set_hidden(hidden);
            (
                if hidden { "hide" } else { "show" },
                slide,
                was_hidden != hidden,
            )
        }
    };
    if !changed && !as_json {
        eprintln!("Note: {action} left slide {slide} as it was");
    }
    publish_presentation(&presentation, output)?;
    mutation_record(
        as_json,
        action,
        json!({ "slide": slide, "slides": presentation.len(), "changed": changed }),
        output,
    )
}

/// Replaces one slide's speaker notes.
pub fn notes_set(
    file: &Path,
    slide: usize,
    text: Option<&str>,
    from_file: Option<&Path>,
    output: &Path,
    as_json: bool,
) -> Result<()> {
    ensure_output_paths_available(&[output.to_path_buf()])?;
    let text = match (text, from_file) {
        (Some(text), None) => text.to_owned(),
        (None, Some(path)) => std::fs::read_to_string(path)
            .map_err(|error| format!("cannot read --from-file {}: {error}", path.display()))?,
        _ => return Err("give --text or --from-file".into()),
    };
    let mut presentation = Presentation::open(file)?;
    let index = slide_index(&presentation, slide)?;
    presentation.set_notes_text(index, text.trim_end_matches(['\r', '\n']))?;
    publish_presentation(&presentation, output)?;
    mutation_record(as_json, "set", json!({ "slide": slide }), output)
}

/// Reports every overflowing text frame and returns whether none overflows.
pub fn fit(file: &Path, as_json: bool) -> Result<bool> {
    let report = Presentation::open(file)?.text_fit_report()?;
    let entries = report
        .overflowing
        .iter()
        .map(|overflow| {
            let frame = &overflow.frame;
            json!({
                "slide": frame.slide_index + 1,
                "shape_id": frame.shape_id,
                "name": frame.name,
                "autofit": autofit_label(frame.autofit),
                "font_scale": round4(frame.layout.font_scale),
                "text_height_points": round4(frame.layout.height),
                "frame_height_points": round4(frame.layout.usable.height),
                "needed_font_scale": overflow.needed_font_scale.map(round4),
            })
        })
        .collect::<Vec<_>>();
    // A reader that closes standard output early, as `| head` does, does not
    // change the verdict.
    if let Err(error) = print_fit_report(&entries, report.frames_checked, as_json)
        && !error
            .downcast_ref::<io::Error>()
            .is_some_and(|error| error.kind() == io::ErrorKind::BrokenPipe)
    {
        return Err(error);
    }
    Ok(entries.is_empty())
}

/// Rounds a layout measure to four decimals for the report.
fn round4(value: f64) -> f64 {
    (value * 10_000.0).round() / 10_000.0
}

fn print_fit_report(entries: &[Value], frames_checked: usize, as_json: bool) -> Result<()> {
    if as_json {
        print_json(json!({
            "fits": entries.is_empty(),
            "frames_checked": frames_checked,
            "overflowing": entries,
        }))?;
    } else {
        let mut stdout = io::stdout().lock();
        for entry in entries {
            writeln!(
                stdout,
                "Slide {}, shape {} ({}): overflows, needed font scale {}",
                entry["slide"],
                entry["shape_id"],
                entry["name"].as_str().unwrap_or_default(),
                entry["needed_font_scale"]
                    .as_f64()
                    .map_or_else(|| "none (below 0.25)".to_owned(), |scale| scale.to_string())
            )?;
        }
        writeln!(
            stdout,
            "{} of {} text frame(s) overflow",
            entries.len(),
            frames_checked
        )?;
    }
    Ok(())
}

fn core_json(core: Option<&CoreProperties>) -> Value {
    let field = |value: fn(&CoreProperties) -> &Option<String>| {
        core.and_then(|core| value(core).as_deref())
    };
    json!({
        "title": field(|core| &core.title),
        "author": field(|core| &core.creator),
        "subject": field(|core| &core.subject),
        "keywords": field(|core| &core.keywords),
        "description": field(|core| &core.description),
        "category": field(|core| &core.category),
        "last_modified_by": field(|core| &core.last_modified_by),
        "created": field(|core| &core.created),
        "modified": field(|core| &core.modified),
        "content_status": field(|core| &core.content_status),
        "identifier": field(|core| &core.identifier),
        "language": field(|core| &core.language),
        "last_printed": field(|core| &core.last_printed),
        "revision": field(|core| &core.revision),
        "version": field(|core| &core.version),
    })
}

/// Prints the core properties.
pub fn meta_get(file: &Path, as_json: bool) -> Result<()> {
    let presentation = Presentation::open(file)?;
    let core = core_json(presentation.core_properties());
    if as_json {
        return print_json(json!({ "core": core }));
    }
    let mut stdout = io::stdout().lock();
    if let Value::Object(fields) = &core {
        for (name, value) in fields {
            if let Some(value) = value.as_str() {
                writeln!(stdout, "{name}: {value}")?;
            }
        }
    }
    Ok(())
}

/// Sets the given core properties: title, author, subject, keywords,
/// description and category, in that order.
pub fn meta_set(
    file: &Path,
    values: [Option<String>; 6],
    output: &Path,
    as_json: bool,
) -> Result<()> {
    ensure_output_paths_available(&[output.to_path_buf()])?;
    if values.iter().all(Option::is_none) {
        return Err(
            "meta set needs at least one of --title, --author, --subject, --keywords, \
                    --description or --category"
                .into(),
        );
    }
    let mut presentation = Presentation::open(file)?;
    let core = presentation.core_properties_mut();
    let slots = [
        &mut core.title,
        &mut core.creator,
        &mut core.subject,
        &mut core.keywords,
        &mut core.description,
        &mut core.category,
    ];
    for (value, slot) in values.into_iter().zip(slots) {
        if value.is_some() {
            *slot = value;
        }
    }
    publish_presentation(&presentation, output)?;
    mutation_record(
        as_json,
        "set",
        json!({ "core": core_json(presentation.core_properties()) }),
        output,
    )
}
