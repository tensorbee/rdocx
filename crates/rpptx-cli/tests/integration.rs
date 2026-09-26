use std::fs;
use std::io::Cursor;
use std::path::{Path, PathBuf};
use std::process::{Command, Output};
use std::sync::atomic::{AtomicUsize, Ordering};

use oxml_opc::relationship::rel_types;
use oxml_opc::{OpcPackage, content_types};
use rpptx::{Angle, CT_TextCharacterProperties, Comment, CommentAuthor, Emu, Presentation};
use serde_json::json;

static TEMP_COUNTER: AtomicUsize = AtomicUsize::new(0);

struct TempWorkspace {
    path: PathBuf,
}

impl TempWorkspace {
    fn new(label: &str) -> Self {
        let id = TEMP_COUNTER.fetch_add(1, Ordering::Relaxed);
        let path =
            std::env::temp_dir().join(format!("rpptx-cli-{label}-{}-{id}", std::process::id()));
        fs::create_dir_all(&path).expect("create temporary workspace");
        Self { path }
    }
}

impl Drop for TempWorkspace {
    fn drop(&mut self) {
        let _ = fs::remove_dir_all(&self.path);
    }
}

fn cli(args: &[&str]) -> Output {
    Command::new(env!("CARGO_BIN_EXE_rpptx"))
        .args(args)
        .output()
        .expect("run rpptx CLI")
}

fn write_deck(path: &Path, texts: &[&str]) {
    let mut presentation = Presentation::new().expect("open bundled template");
    for text in texts {
        presentation.add_slide(6).expect("add blank slide");
        presentation
            .slide_mut(presentation.len() - 1)
            .unwrap()
            .add_textbox(Emu(100_000), Emu(100_000), Emu(3_000_000), Emu(800_000))
            .expect("add text box")
            .set_text(text)
            .expect("set slide text");
    }
    presentation.save(path).expect("write fixture deck");
}

fn add_speaker_notes(path: &Path, text: &str) {
    let mut package = OpcPackage::open(path).expect("open notes fixture package");
    let notes_part = "/ppt/notesSlides/notesSlide1.xml";
    let xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes Placeholder"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr b="1"><a:extLst><a:ext uri="{{6D487B31-4C56-4F45-AED3-4C741AC43E77}}"><x:payload xmlns:x="urn:rdocx:test"/></a:ext></a:extLst></a:rPr><a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>"#
    );
    package.set_part(notes_part, xml.into_bytes());
    package
        .content_types
        .add_override(notes_part, content_types::NOTES_SLIDE);
    package
        .get_or_create_part_rels("/ppt/slides/slide1.xml")
        .add_with_id(
            "notes",
            rel_types::NOTES_SLIDE,
            "../notesSlides/notesSlide1.xml",
        );
    package.get_or_create_part_rels(notes_part).add_with_id(
        "master",
        rel_types::NOTES_MASTER,
        "../notesMasters/notesMaster1.xml",
    );
    package.get_or_create_part_rels(notes_part).add_with_id(
        "slide",
        rel_types::SLIDE,
        "../slides/slide1.xml",
    );
    package.save(path).expect("write notes fixture package");
}

fn png_dimensions(bytes: &[u8]) -> (u32, u32) {
    assert!(bytes.starts_with(b"\x89PNG\r\n\x1a\n"));
    assert_eq!(&bytes[12..16], b"IHDR");
    (
        u32::from_be_bytes(bytes[16..20].try_into().unwrap()),
        u32::from_be_bytes(bytes[20..24].try_into().unwrap()),
    )
}

fn write_outline_deck(path: &Path) {
    let mut presentation = Presentation::new().expect("open bundled template");
    presentation.add_slide(0).expect("add title slide");
    let title_placeholder = presentation
        .slide(0)
        .unwrap()
        .title()
        .expect("title layout supplies title")
        .placeholder_idx();
    let title_index = presentation
        .slide(0)
        .unwrap()
        .shapes()
        .position(|shape| shape.placeholder_idx() == title_placeholder)
        .expect("locate title shape");
    presentation
        .slide_mut(0)
        .unwrap()
        .shape_mut(title_index)
        .unwrap()
        .set_text("Roadmap")
        .unwrap();

    {
        let mut slide = presentation.slide_mut(0).unwrap();
        let mut shape = slide
            .add_textbox(Emu(100_000), Emu(1_000_000), Emu(4_000_000), Emu(1_500_000))
            .unwrap();
        shape.set_text("First item").unwrap();
        let mut frame = shape.text_frame().unwrap();
        let mut nested = frame.add_paragraph();
        nested.set_text("Nested item");
        assert!(nested.set_level(2));
        frame.add_paragraph().set_text("");
        slide.add_group_shape().unwrap();
        slide
            .add_textbox(Emu(200_000), Emu(200_000), Emu(2_000_000), Emu(500_000))
            .unwrap()
            .set_text("Grouped item")
            .unwrap();
    }
    presentation.save(path).expect("write outline fixture");

    let mut package = OpcPackage::open(path).expect("open outline package");
    let slide_part = package
        .content_types
        .overrides
        .iter()
        .find_map(|(part, content_type)| {
            (content_type
                == "application/vnd.openxmlformats-officedocument.presentationml.slide+xml")
                .then_some(part.clone())
        })
        .expect("outline slide part");
    let xml = String::from_utf8(package.get_part(&slide_part).unwrap().to_vec()).unwrap();
    let marker = xml.find("Grouped item").expect("grouped marker");
    let shape_start = xml[..marker].rfind("<p:sp>").expect("grouped shape start");
    let shape_end = marker + xml[marker..].find("</p:sp>").expect("grouped shape end") + 7;
    let shape = xml[shape_start..shape_end].to_owned();
    let without_shape = format!("{}{}", &xml[..shape_start], &xml[shape_end..]);
    let group_end = without_shape.rfind("</p:grpSp>").expect("empty group end");
    let grouped = format!(
        "{}{}{}",
        &without_shape[..group_end],
        shape,
        &without_shape[group_end..]
    );
    let first_marker = grouped.find("First item").expect("body marker");
    let first_start = grouped[..first_marker]
        .rfind("<p:sp>")
        .expect("body shape start");
    let first_end = first_marker
        + grouped[first_marker..]
            .find("</p:sp>")
            .expect("body shape end")
        + 7;
    let mut body_shape = grouped[first_start..first_end].to_owned();
    body_shape = body_shape.replacen("<p:nvPr/>", "<p:nvPr><p:ph type=\"body\"/></p:nvPr>", 1);
    assert!(body_shape.contains("<p:ph type=\"body\"/>"));
    let without_body = format!("{}{}", &grouped[..first_start], &grouped[first_end..]);
    let title_marker = without_body.find("Roadmap").expect("title marker");
    let title_start = without_body[..title_marker]
        .rfind("<p:sp>")
        .expect("title shape start");
    let reordered = format!(
        "{}{}{}",
        &without_body[..title_start],
        body_shape,
        &without_body[title_start..]
    );
    let broken = reordered.replacen(
        "<a:t>Grouped item</a:t>",
        "<a:t>Grouped</a:t></a:r><a:br/><a:r><a:t>item</a:t>",
        1,
    );
    assert_ne!(broken, reordered);
    package.set_part(&slide_part, broken.into_bytes());
    package.save(path).expect("write grouped outline fixture");
}

fn make_title_field_only(path: &Path) {
    let mut package = OpcPackage::open(path).expect("open outline package");
    let slide_part = package
        .content_types
        .overrides
        .iter()
        .find_map(|(part, content_type)| {
            (content_type
                == "application/vnd.openxmlformats-officedocument.presentationml.slide+xml")
                .then_some(part.clone())
        })
        .expect("outline slide part");
    let xml = String::from_utf8(package.get_part(&slide_part).unwrap().to_vec()).unwrap();
    let field = r#"<a:fld id="{00000000-0000-0000-0000-000000000145}" type="title"><a:t>Roadmap</a:t></a:fld>"#;
    let xml = xml.replacen("<a:r><a:t>Roadmap</a:t></a:r>", field, 1);
    assert!(xml.contains(field));
    package.set_part(&slide_part, xml.into_bytes());
    package.save(path).expect("write field-only title fixture");
}

/// Writes a title-and-content slide with a rotated, formatted text box, a
/// table, and an autofit text body, followed by a blank slide.
fn write_structured_deck(path: &Path) {
    let mut presentation = Presentation::new().expect("open bundled template");
    presentation
        .add_slide(1)
        .expect("add title and content slide");
    let title_index = presentation
        .slide(0)
        .unwrap()
        .shapes()
        .position(|shape| shape.placeholder_type() == Some("title"))
        .expect("layout supplies a title placeholder");
    presentation
        .slide_mut(0)
        .unwrap()
        .shape_mut(title_index)
        .unwrap()
        .set_text("Roadmap")
        .unwrap();
    {
        let mut slide = presentation.slide_mut(0).unwrap();
        let mut shape = slide
            .add_textbox(Emu(100_000), Emu(200_000), Emu(3_000_000), Emu(800_000))
            .unwrap();
        shape.set_rotation(Angle::from_degrees(45.0)).unwrap();
        shape.set_text("plain ").unwrap();
        let mut frame = shape.text_frame().unwrap();
        frame.paragraph_mut(0).unwrap().add_run("rich").set_properties(
            CT_TextCharacterProperties::from_xml(
                br#"<a:rPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" sz="1800" b="1" i="0" u="sng"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:latin typeface="Carlito"/></a:rPr>"#,
            )
            .unwrap(),
        );
        let mut nested = frame.add_paragraph();
        nested.set_text("nested");
        assert!(nested.set_level(1));
    }
    {
        let mut slide = presentation.slide_mut(0).unwrap();
        let mut frame = slide
            .add_table(1, 2, Emu(0), Emu(1_000_000), Emu(2_000_000), Emu(400_000))
            .unwrap();
        let mut table = frame.table_mut().unwrap();
        table.cell_mut(0, 0).unwrap().set_text("A1");
        table.cell_mut(0, 1).unwrap().set_text("B1");
    }
    presentation.add_slide(6).expect("add blank slide");
    presentation.save(path).expect("write structured deck");

    let mut package = OpcPackage::open(path).expect("open structured deck");
    let slide_part = "/ppt/slides/slide1.xml";
    let xml = String::from_utf8(package.get_part(slide_part).unwrap().to_vec()).unwrap();
    let marker = xml.find(">plain </a:t>").expect("text box marker");
    let body = xml[..marker].rfind("<a:bodyPr/>").expect("text box body");
    let xml = format!(
        "{}<a:bodyPr><a:normAutofit/></a:bodyPr>{}",
        &xml[..body],
        &xml[body + "<a:bodyPr/>".len()..]
    );
    package.set_part(slide_part, xml.into_bytes());
    package.save(path).expect("write autofit text box");
}

fn corpus_dir() -> PathBuf {
    std::env::var_os("RDOCX_PPTX_CORPUS_DIR")
        .map(PathBuf::from)
        .unwrap_or_else(|| Path::new(env!("CARGO_MANIFEST_DIR")).join("../../corpus/pptx"))
}

#[test]
fn thumbnail_and_outline_match_the_presentation_contract() {
    let temp = TempWorkspace::new("thumbnail-outline");
    let deck = temp.path.join("roadmap.pptx");
    write_outline_deck(&deck);

    let thumbnail = temp.path.join("thumb.png");
    let output = cli(&[
        "thumbnail",
        deck.to_str().unwrap(),
        "--output",
        thumbnail.to_str().unwrap(),
    ]);
    assert!(
        output.status.success(),
        "thumbnail failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let dimensions = png_dimensions(&fs::read(&thumbnail).unwrap());
    assert_eq!(dimensions.0, 320);

    let output = cli(&["outline", deck.to_str().unwrap()]);
    assert!(
        output.status.success(),
        "outline failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    assert_eq!(
        String::from_utf8(output.stdout).unwrap(),
        "Slide 1: Roadmap\n- First item\n    - Nested item\n- Grouped item\n"
    );
}

#[test]
fn outline_emits_a_field_only_title_once() {
    let temp = TempWorkspace::new("outline-field-title");
    let deck = temp.path.join("field-title.pptx");
    write_outline_deck(&deck);
    make_title_field_only(&deck);

    let output = cli(&["outline", deck.to_str().unwrap()]);
    assert!(output.status.success());
    assert_eq!(
        String::from_utf8(output.stdout).unwrap(),
        "Slide 1: Roadmap\n- First item\n    - Nested item\n- Grouped item\n"
    );
}

#[test]
fn thumbnail_preserves_a_nonstandard_slide_aspect_ratio() {
    let temp = TempWorkspace::new("thumbnail-aspect");
    let deck = temp.path.join("portrait.pptx");
    let thumbnail = temp.path.join("portrait.png");
    let mut presentation = Presentation::new().unwrap();
    presentation
        .set_slide_size(Emu(4_000_000), Emu(8_000_000))
        .unwrap();
    presentation.add_slide(6).unwrap();
    presentation.save(&deck).unwrap();

    let output = cli(&[
        "thumbnail",
        deck.to_str().unwrap(),
        "--output",
        thumbnail.to_str().unwrap(),
    ]);
    assert!(output.status.success());
    assert_eq!(png_dimensions(&fs::read(thumbnail).unwrap()), (320, 640));
}

#[test]
fn thumbnail_uses_shared_default_output_and_explicit_output_wins() {
    let temp = TempWorkspace::new("thumbnail-output");
    let deck = temp.path.join("deck.pptx");
    write_deck(&deck, &["thumbnail"]);

    let defaulted = cli(&["thumbnail", deck.to_str().unwrap()]);
    assert!(defaulted.status.success());
    let default_path = temp.path.join("deck.png");
    assert_eq!(png_dimensions(&fs::read(&default_path).unwrap()).0, 320);

    fs::remove_file(&default_path).unwrap();
    let explicit = temp.path.join("chosen.png");
    let selected = cli(&[
        "thumbnail",
        deck.to_str().unwrap(),
        "--output",
        explicit.to_str().unwrap(),
    ]);
    assert!(selected.status.success());
    assert_eq!(png_dimensions(&fs::read(explicit).unwrap()).0, 320);
    assert!(!default_path.exists());
}

#[test]
fn validate_rejects_corruption_and_accepts_the_pinned_corpus() {
    let temp = TempWorkspace::new("validate");
    let valid = temp.path.join("valid.pptx");
    write_deck(&valid, &["valid"]);
    let mut package = OpcPackage::open(&valid).expect("open fixture package");
    let slide = package
        .content_types
        .overrides
        .iter()
        .find_map(|(part, content_type)| {
            (content_type
                == "application/vnd.openxmlformats-officedocument.presentationml.slide+xml")
                .then_some(part.clone())
        })
        .expect("fixture slide part");
    package
        .get_or_create_part_rels(&slide)
        .add(rel_types::IMAGE, "../media/missing.png");
    let corrupt = temp.path.join("corrupt.pptx");
    package.save(&corrupt).expect("write corrupted deck");
    assert!(
        !cli(&["validate", corrupt.to_str().unwrap()])
            .status
            .success()
    );

    let corpus = corpus_dir();
    assert!(
        corpus.is_dir(),
        "required corpus missing at {}",
        corpus.display()
    );
    let entries = include_str!("../../../scripts/pptx-corpus-manifest.tsv")
        .lines()
        .skip(1)
        .map(|line| line.split('\t').next().unwrap())
        .collect::<Vec<_>>();
    assert_eq!(entries.len(), 50);
    for entry in entries {
        let deck = corpus.join(entry);
        assert!(
            deck.is_file(),
            "missing pinned corpus deck {}",
            deck.display()
        );
        let output = cli(&["validate", deck.to_str().unwrap()]);
        assert!(
            output.status.success(),
            "validate failed for {}: {}",
            deck.display(),
            String::from_utf8_lossy(&output.stderr)
        );
    }
}

#[test]
fn inspect_and_text_report_presentation_order() {
    let temp = TempWorkspace::new("inspect-text");
    let deck = temp.path.join("ordered.pptx");
    write_deck(&deck, &["first slide", "second slide"]);
    let mut presentation = Presentation::open(&deck).unwrap();
    let core = presentation.core_properties_mut();
    core.title = Some("Quarterly deck".to_owned());
    core.creator = Some("A. Presenter".to_owned());
    core.subject = Some("Plain inspect metadata".to_owned());
    core.description = Some("Metadata command regression".to_owned());
    core.keywords = Some("slides, metadata".to_owned());
    core.last_modified_by = Some("F-144".to_owned());
    core.created = Some("2026-08-13T10:00:00Z".to_owned());
    core.modified = Some("2026-08-13T11:00:00Z".to_owned());
    presentation.save(&deck).unwrap();
    let inspected = cli(&["inspect", deck.to_str().unwrap(), "--json"]);
    assert!(inspected.status.success());
    let value: serde_json::Value = serde_json::from_slice(&inspected.stdout).unwrap();
    assert_eq!(value["schema"], 1);
    assert_eq!(value["slides"], 2);
    assert_eq!(value["slide_details"][0]["shapes"], 1);

    let inspected = cli(&["inspect", deck.to_str().unwrap()]);
    assert!(inspected.status.success());
    let inspected = String::from_utf8(inspected.stdout).unwrap();
    for expected in [
        "Title: Quarterly deck",
        "Creator: A. Presenter",
        "Subject: Plain inspect metadata",
        "Description: Metadata command regression",
        "Keywords: slides, metadata",
        "Last modified by: F-144",
        "Created: 2026-08-13T10:00:00Z",
        "Modified: 2026-08-13T11:00:00Z",
    ] {
        assert!(inspected.contains(expected), "missing {expected:?}");
    }

    let text = cli(&["text", deck.to_str().unwrap()]);
    assert!(text.status.success());
    assert_eq!(
        String::from_utf8(text.stdout).unwrap(),
        "first slide\nsecond slide\n"
    );

    let help = cli(&["--help"]);
    assert!(help.status.success());
    let help = String::from_utf8(help.stdout).unwrap();
    for command in [
        "inspect",
        "text",
        "convert",
        "diff",
        "replace",
        "validate",
        "render",
        "thumbnail",
        "outline",
        "comment",
    ] {
        assert!(help.contains(command), "missing command {command}");
    }
}

#[test]
fn convert_rejects_png_rasters_above_the_pixel_budget() {
    let temp = TempWorkspace::new("convert-dpi-bound");
    let deck = temp.path.join("bounded.pptx");
    let output = temp.path.join("bounded.png");
    write_deck(&deck, &["bounded"]);

    let result = cli(&[
        "convert",
        deck.to_str().unwrap(),
        "--to",
        "png",
        "--output",
        output.to_str().unwrap(),
        "--dpi",
        "330",
    ]);

    assert!(!result.status.success());
    assert!(String::from_utf8_lossy(&result.stderr).contains("8,000,000 pixel limit"));
    assert!(!output.exists());
}

#[test]
fn render_rejects_png_rasters_above_the_pixel_budget() {
    let temp = TempWorkspace::new("render-dpi-bound");
    let deck = temp.path.join("bounded.pptx");
    let output = temp.path.join("rendered");
    write_deck(&deck, &["bounded"]);

    let result = cli(&[
        "render",
        deck.to_str().unwrap(),
        "--output",
        output.to_str().unwrap(),
        "--dpi",
        "330",
    ]);

    assert!(!result.status.success());
    assert!(String::from_utf8_lossy(&result.stderr).contains("8,000,000 pixel limit"));
    assert!(!output.join("bounded_slide1.png").exists());
}

#[test]
fn convert_rejects_zero_slide_png_without_creating_output() {
    let temp = TempWorkspace::new("convert-empty");
    let deck = temp.path.join("empty.pptx");
    let output = temp.path.join("empty.png");
    let presentation = Presentation::new().unwrap();
    assert!(presentation.is_empty());
    presentation.save(&deck).unwrap();

    let result = cli(&[
        "convert",
        deck.to_str().unwrap(),
        "--to",
        "png",
        "--output",
        output.to_str().unwrap(),
    ]);

    assert!(!result.status.success());
    assert!(String::from_utf8_lossy(&result.stderr).contains("no slides"));
    assert!(result.stdout.is_empty());
    assert!(!output.exists());
}

#[test]
fn render_rejects_zero_slide_image_formats_without_creating_output() {
    let temp = TempWorkspace::new("render-empty");
    let deck = temp.path.join("empty.pptx");
    let presentation = Presentation::new().unwrap();
    assert!(presentation.is_empty());
    presentation.save(&deck).unwrap();

    for format in ["png", "jpeg", "tiff"] {
        let output = temp.path.join(format!("empty-{format}"));
        let result = cli(&[
            "render",
            deck.to_str().unwrap(),
            "--output",
            output.to_str().unwrap(),
            "--format",
            format,
        ]);

        assert!(!result.status.success(), "{format} unexpectedly succeeded");
        assert!(
            String::from_utf8_lossy(&result.stderr).contains("no slides selected"),
            "unexpected {format} error: {}",
            String::from_utf8_lossy(&result.stderr)
        );
        assert!(result.stdout.is_empty());
        assert!(!output.exists());
    }
}

#[test]
fn convert_and_render_write_deterministic_pdf_and_png_outputs() {
    let temp = TempWorkspace::new("render");
    let deck = temp.path.join("rendered.pptx");
    write_deck(&deck, &["one", "two"]);
    let pdf = temp.path.join("rendered.pdf");
    let converted = cli(&["convert", deck.to_str().unwrap(), "--to", "pdf"]);
    assert!(converted.status.success());
    assert!(fs::read(&pdf).unwrap().starts_with(b"%PDF-"));

    let converted = cli(&["convert", deck.to_str().unwrap(), "--to", "png"]);
    assert!(converted.status.success());
    assert!(
        fs::read(temp.path.join("rendered_001.png"))
            .unwrap()
            .starts_with(b"\x89PNG")
    );
    assert!(
        fs::read(temp.path.join("rendered_002.png"))
            .unwrap()
            .starts_with(b"\x89PNG")
    );

    let output_dir = temp.path.join("selected");
    let rendered = cli(&[
        "render",
        deck.to_str().unwrap(),
        "--output",
        output_dir.to_str().unwrap(),
        "--slide",
        "2",
    ]);
    assert!(rendered.status.success());
    assert!(
        fs::read(output_dir.join("rendered_slide2.png"))
            .unwrap()
            .starts_with(b"\x89PNG")
    );
    assert!(!output_dir.join("rendered_slide1.png").exists());
}

#[test]
fn image_export_options_write_declared_formats_and_ranges() {
    let temp = TempWorkspace::new("image-options");
    let deck = temp.path.join("images.pptx");
    write_deck(&deck, &["one", "two", "three"]);

    let converted = cli(&[
        "convert",
        deck.to_str().unwrap(),
        "--to",
        "jpeg",
        "--dpi",
        "72",
        "--quality",
        "80",
        "--slides",
        "2",
    ]);
    assert!(converted.status.success());
    assert!(
        fs::read(temp.path.join("images.jpg"))
            .unwrap()
            .starts_with(b"\xff\xd8")
    );
    assert!(!temp.path.join("images_001.jpg").exists());

    let output_dir = temp.path.join("tiff");
    let rendered = cli(&[
        "render",
        deck.to_str().unwrap(),
        "--output",
        output_dir.to_str().unwrap(),
        "--format",
        "tiff",
        "--dpi",
        "72",
        "--slide",
        "1,3",
    ]);
    assert!(rendered.status.success());
    assert!(
        fs::read(output_dir.join("images.tiff"))
            .unwrap()
            .starts_with(b"II*\0")
    );
    assert!(!output_dir.join("images_slide2.tiff").exists());

    let bad_output = temp.path.join("bad.jpg");
    let rejected = cli(&[
        "convert",
        deck.to_str().unwrap(),
        "--to",
        "jpeg",
        "--output",
        bad_output.to_str().unwrap(),
        "--quality",
        "0",
        "--slides",
        "1",
    ]);
    assert!(!rejected.status.success());
    assert!(!bad_output.exists());

    let bad_dir = temp.path.join("bad-range");
    let rejected = cli(&[
        "render",
        deck.to_str().unwrap(),
        "--output",
        bad_dir.to_str().unwrap(),
        "--slide",
        "4",
    ]);
    assert!(!rejected.status.success());
    assert!(!bad_dir.exists());
}

#[test]
fn convert_streams_crafted_multi_slide_pngs_with_one_based_names() {
    let temp = TempWorkspace::new("convert-stream");
    let deck = temp.path.join("streamed.pptx");
    let texts = vec!["streamed"; 24];
    write_deck(&deck, &texts);

    let converted = cli(&[
        "convert",
        deck.to_str().unwrap(),
        "--to",
        "png",
        "--dpi",
        "72",
    ]);

    assert!(converted.status.success());
    let stdout = String::from_utf8(converted.stdout).unwrap();
    for one_based in 1..=24 {
        let output = temp.path.join(format!("streamed_{one_based:03}.png"));
        assert!(fs::read(output).unwrap().starts_with(b"\x89PNG"));
        assert!(stdout.contains(&format!("Slide {one_based} ->")));
    }
    let commands = include_str!("../src/commands.rs");
    assert!(commands.contains("render_one_raster_page"));
    assert!(
        !commands.contains("render_all_pages"),
        "convert must not retain every encoded PNG"
    );
    assert!(
        !commands.contains("RasterOutput::SeparatePages(images)"),
        "separate PNG and JPEG export must not branch on an all-pages image Vec"
    );
    assert!(
        !commands.contains("zip(images.iter())"),
        "separate PNG and JPEG export must not retain every encoded page"
    );
}

#[test]
fn multi_file_image_export_preserves_existing_outputs_before_streaming() {
    let temp = TempWorkspace::new("convert-existing-output");
    let deck = temp.path.join("existing.pptx");
    let output = temp.path.join("export.png");
    write_deck(&deck, &["one", "two"]);
    let preexisting = temp.path.join("export_002.png");
    fs::write(&preexisting, b"keep me").unwrap();

    let converted = cli(&[
        "convert",
        deck.to_str().unwrap(),
        "--to",
        "png",
        "--output",
        output.to_str().unwrap(),
        "--dpi",
        "72",
    ]);

    assert!(!converted.status.success());
    assert!(!temp.path.join("export_001.png").exists());
    assert_eq!(fs::read(preexisting).unwrap(), b"keep me");
}

#[test]
fn diff_reports_slide_text_lcs_changes() {
    let temp = TempWorkspace::new("diff");
    let before = temp.path.join("before.pptx");
    let after = temp.path.join("after.pptx");
    write_deck(&before, &["same", "removed"]);
    write_deck(&after, &["same", "added"]);
    let output = cli(&["diff", before.to_str().unwrap(), after.to_str().unwrap()]);
    assert!(output.status.success());
    let stdout = String::from_utf8(output.stdout).unwrap();
    assert!(stdout.contains("- [2] removed"));
    assert!(stdout.contains("+ [2] added"));
}

#[test]
fn diff_rejects_large_repeated_slide_matrix_above_cell_budget() {
    let temp = TempWorkspace::new("diff-bound");
    let deck = temp.path.join("repeated.pptx");
    let repeated = vec!["repeat"; 1_000];
    write_deck(&deck, &repeated);

    let output = cli(&["diff", deck.to_str().unwrap(), deck.to_str().unwrap()]);

    assert!(!output.status.success());
    assert!(output.stdout.is_empty());
    assert!(String::from_utf8_lossy(&output.stderr).contains("1,000,000 LCS cell limit"));
}

#[test]
fn replacement_preserves_formatting_and_opaque_parts() {
    let temp = TempWorkspace::new("replace");
    let source = temp.path.join("source.pptx");
    let mut presentation = Presentation::new().unwrap();
    presentation.add_slide(6).unwrap();
    let mut slide = presentation.slide_mut(0).unwrap();
    let mut shape = slide
        .add_textbox(Emu(1), Emu(2), Emu(3_000_000), Emu(800_000))
        .unwrap();
    shape.set_text("pre old").unwrap();
    let mut frame = shape.text_frame().unwrap();
    let mut paragraph = frame.paragraph_mut(0).unwrap();
    let mut first_properties = CT_TextCharacterProperties::default();
    first_properties.bold = Some(true);
    paragraph
        .run_mut(0)
        .unwrap()
        .set_properties(first_properties.clone());
    let mut second = paragraph.add_run(" needle post");
    let mut second_properties = CT_TextCharacterProperties::default();
    second_properties.italic = Some(true);
    second.set_properties(second_properties.clone());
    let mut package =
        OpcPackage::from_reader(Cursor::new(presentation.to_bytes().unwrap())).unwrap();
    package.set_part("/custom/opaque.bin", b"opaque bytes".to_vec());
    package
        .content_types
        .add_default("bin", "application/octet-stream");
    package.save(&source).unwrap();

    let output = temp.path.join("replaced.pptx");
    let result = cli(&[
        "replace",
        source.to_str().unwrap(),
        "--placeholder",
        "old needle",
        "--value",
        "new",
        "--output",
        output.to_str().unwrap(),
    ]);
    assert!(result.status.success());
    let reopened = Presentation::open(&output).unwrap();
    let paragraph = reopened
        .slide(0)
        .unwrap()
        .shape(0)
        .unwrap()
        .text_frame()
        .unwrap()
        .paragraph(0)
        .unwrap();
    assert_eq!(paragraph.text(), "pre new post");
    assert_eq!(
        paragraph.run(0).unwrap().properties(),
        Some(&first_properties)
    );
    assert_eq!(
        paragraph.run(1).unwrap().properties(),
        Some(&second_properties)
    );
    let package = OpcPackage::open(&output).unwrap();
    assert_eq!(
        package.get_part("/custom/opaque.bin"),
        Some(b"opaque bytes".as_slice())
    );
}

#[test]
fn rpptx_replace_is_guarded_counted_and_includes_notes() {
    let temp = TempWorkspace::new("guarded-replace");
    let source = temp.path.join("source.pptx");
    write_deck(&source, &["Draft title Draft"]);
    add_speaker_notes(&source, "Draft speaker note");
    let source_bytes = fs::read(&source).unwrap();

    let result = cli(&[
        "replace",
        source.to_str().unwrap(),
        "--placeholder",
        "Draft",
        "--value",
        "Final",
        "--expect",
        "3",
        "--output",
        source.to_str().unwrap(),
    ]);
    assert!(!result.status.success());
    assert!(String::from_utf8_lossy(&result.stderr).contains("output already exists"));
    assert_eq!(fs::read(&source).unwrap(), source_bytes);

    let existing = temp.path.join("existing.pptx");
    fs::write(&existing, b"keep me").unwrap();
    let result = cli(&[
        "replace",
        source.to_str().unwrap(),
        "--placeholder",
        "Draft",
        "--value",
        "Final",
        "--expect",
        "3",
        "--output",
        existing.to_str().unwrap(),
    ]);
    assert!(!result.status.success());
    assert!(String::from_utf8_lossy(&result.stderr).contains("output already exists"));
    assert_eq!(fs::read(&existing).unwrap(), b"keep me");

    let zero_output = temp.path.join("zero.pptx");
    let result = cli(&[
        "replace",
        source.to_str().unwrap(),
        "--placeholder",
        "Missing",
        "--value",
        "Final",
        "--output",
        zero_output.to_str().unwrap(),
    ]);
    assert!(!result.status.success());
    assert!(!zero_output.exists());

    let mismatch = temp.path.join("mismatch.pptx");
    let result = cli(&[
        "replace",
        source.to_str().unwrap(),
        "--placeholder",
        "Draft",
        "--value",
        "Final",
        "--expect",
        "2",
        "--output",
        mismatch.to_str().unwrap(),
    ]);
    assert!(!result.status.success());
    assert!(
        String::from_utf8_lossy(&result.stderr)
            .contains("expected 2 replacement(s) of \"Draft\", found 3")
    );
    assert!(!mismatch.exists());

    let explicit_zero = temp.path.join("explicit-zero.pptx");
    let result = cli(&[
        "replace",
        source.to_str().unwrap(),
        "--placeholder",
        "Missing",
        "--value",
        "Final",
        "--expect",
        "0",
        "--output",
        explicit_zero.to_str().unwrap(),
    ]);
    assert!(result.status.success());
    assert_eq!(
        String::from_utf8(result.stdout).unwrap(),
        format!(
            "Replaced 0 occurrence(s) of \"Missing\" -> \"Final\"\nWritten to {}\n",
            explicit_zero.display()
        )
    );
    assert_eq!(
        Presentation::open(&explicit_zero)
            .unwrap()
            .slide(0)
            .unwrap()
            .notes_text()
            .as_deref(),
        Some("Draft speaker note")
    );

    let output = temp.path.join("replaced.pptx");
    let result = cli(&[
        "replace",
        source.to_str().unwrap(),
        "--placeholder",
        "Draft",
        "--value",
        "Final",
        "--expect",
        "3",
        "--output",
        output.to_str().unwrap(),
    ]);
    assert!(
        result.status.success(),
        "replace failed: {}",
        String::from_utf8_lossy(&result.stderr)
    );
    assert_eq!(
        String::from_utf8(result.stdout).unwrap(),
        format!(
            "Replaced 3 occurrence(s) of \"Draft\" -> \"Final\"\nWritten to {}\n",
            output.display()
        )
    );
    let reopened = Presentation::open(&output).unwrap();
    assert_eq!(reopened.slide(0).unwrap().text(), "Final title Final");
    assert_eq!(
        reopened.slide(0).unwrap().notes_text().as_deref(),
        Some("Final speaker note")
    );
    let package = OpcPackage::open(&output).unwrap();
    let notes_xml = String::from_utf8(
        package
            .get_part("/ppt/notesSlides/notesSlide1.xml")
            .unwrap()
            .to_vec(),
    )
    .unwrap();
    assert!(notes_xml.contains(r#"b="1""#));
    assert!(notes_xml.contains("x:payload"));
    assert!(fs::read_dir(&temp.path).unwrap().all(|entry| {
        !entry
            .unwrap()
            .file_name()
            .to_string_lossy()
            .ends_with(".tmp")
    }));
}

#[test]
fn text_json_anchors_paragraphs_by_typed_shape_paths_with_direct_run_formatting() {
    let temp = TempWorkspace::new("text-json");
    let deck = temp.path.join("structured.pptx");
    write_structured_deck(&deck);
    add_speaker_notes(&deck, "Presenter reminder");
    let shape = |index: usize| json!({ "kind": "shape", "index": index });
    let paragraph = |index: usize| json!({ "kind": "paragraph", "index": index });
    let cell = |index: usize| json!({ "kind": "cell", "index": index });
    let row = json!({ "kind": "row", "index": 0 });
    let plain_run = |text: &str| json!([{ "index": 0, "text": text, "formatting": null }]);

    let output = cli(&["text", deck.to_str().unwrap(), "--json"]);
    assert!(
        output.status.success(),
        "text --json failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let value: serde_json::Value = serde_json::from_slice(&output.stdout).unwrap();
    assert_eq!(
        value,
        json!({
            "schema": 1,
            "slides": [
                {
                    "slide": 1,
                    "id": 256,
                    "paragraphs": [
                        {
                            "path": [shape(0), paragraph(0)],
                            "shape_id": 2,
                            "level": 0,
                            "text": "Roadmap",
                            "runs": plain_run("Roadmap"),
                        },
                        {
                            "path": [shape(1), paragraph(0)],
                            "shape_id": 3,
                            "level": 0,
                            "text": "",
                            "runs": [],
                        },
                        {
                            "path": [shape(2), paragraph(0)],
                            "shape_id": 4,
                            "level": 0,
                            "text": "plain rich",
                            "runs": [
                                { "index": 0, "text": "plain ", "formatting": null },
                                {
                                    "index": 1,
                                    "text": "rich",
                                    "formatting": {
                                        "bold": true,
                                        "italic": false,
                                        "underline": "sng",
                                        "font": "Carlito",
                                        "size_points": 18.0,
                                        "color": "FF0000",
                                    },
                                },
                            ],
                        },
                        {
                            "path": [shape(2), paragraph(1)],
                            "shape_id": 4,
                            "level": 1,
                            "text": "nested",
                            "runs": plain_run("nested"),
                        },
                        {
                            "path": [shape(3), row.clone(), cell(0), paragraph(0)],
                            "shape_id": 5,
                            "level": 0,
                            "text": "A1",
                            "runs": plain_run("A1"),
                        },
                        {
                            "path": [shape(3), row.clone(), cell(1), paragraph(0)],
                            "shape_id": 5,
                            "level": 0,
                            "text": "B1",
                            "runs": plain_run("B1"),
                        },
                    ],
                    "notes": "Presenter reminder",
                },
                { "slide": 2, "id": 257, "paragraphs": [], "notes": null },
            ],
        })
    );

    let grouped = temp.path.join("grouped.pptx");
    write_outline_deck(&grouped);
    let output = cli(&["text", grouped.to_str().unwrap(), "--json"]);
    assert!(output.status.success());
    let value: serde_json::Value = serde_json::from_slice(&output.stdout).unwrap();
    assert_eq!(
        value["slides"][0]["paragraphs"].as_array().unwrap().last(),
        Some(&json!({
            "path": [shape(3), shape(0), paragraph(0)],
            "shape_id": 6,
            "level": 0,
            "text": "Grouped\u{b}item",
            "runs": [
                { "index": 0, "text": "Grouped", "formatting": null },
                { "index": 1, "text": "item", "formatting": null },
            ],
        }))
    );
}

#[test]
fn inspect_json_adds_shape_details_without_changing_existing_slide_keys() {
    let temp = TempWorkspace::new("inspect-shapes");
    let deck = temp.path.join("structured.pptx");
    write_structured_deck(&deck);
    let paragraph = |index: usize, level: u8, text: &str| {
        json!({
            "index": index,
            "level": level,
            "text": text,
            "runs": [{ "index": 0, "text": text, "formatting": null }],
        })
    };

    let output = cli(&["inspect", deck.to_str().unwrap(), "--json"]);
    assert!(
        output.status.success(),
        "inspect --json failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let value: serde_json::Value = serde_json::from_slice(&output.stdout).unwrap();
    assert_eq!(value["slides"], 2);
    assert_eq!(
        value["slide_details"],
        json!([
            {
                "id": 256,
                "name": null,
                "hidden": false,
                "shapes": 4,
                "shape_details": [
                    {
                        "index": 0,
                        "id": 2,
                        "name": "Placeholder 2",
                        "kind": "shape",
                        "placeholder": { "type": "title", "idx": 0 },
                        "position": null,
                        "size": null,
                        "rotation_degrees": null,
                        "autofit": null,
                        "paragraphs": [paragraph(0, 0, "Roadmap")],
                        "table": null,
                        "children": [],
                    },
                    {
                        "index": 1,
                        "id": 3,
                        "name": "Placeholder 3",
                        "kind": "shape",
                        "placeholder": { "type": null, "idx": 1 },
                        "position": null,
                        "size": null,
                        "rotation_degrees": null,
                        "autofit": null,
                        "paragraphs": [{ "index": 0, "level": 0, "text": "", "runs": [] }],
                        "table": null,
                        "children": [],
                    },
                    {
                        "index": 2,
                        "id": 4,
                        "name": "TextBox 4",
                        "kind": "shape",
                        "placeholder": null,
                        "position": { "left_emu": 100_000, "top_emu": 200_000 },
                        "size": { "width_emu": 3_000_000, "height_emu": 800_000 },
                        "rotation_degrees": 45.0,
                        "autofit": "normal",
                        "paragraphs": [
                            {
                                "index": 0,
                                "level": 0,
                                "text": "plain rich",
                                "runs": [
                                    { "index": 0, "text": "plain ", "formatting": null },
                                    {
                                        "index": 1,
                                        "text": "rich",
                                        "formatting": {
                                            "bold": true,
                                            "italic": false,
                                            "underline": "sng",
                                            "font": "Carlito",
                                            "size_points": 18.0,
                                            "color": "FF0000",
                                        },
                                    },
                                ],
                            },
                            paragraph(1, 1, "nested"),
                        ],
                        "table": null,
                        "children": [],
                    },
                    {
                        "index": 3,
                        "id": 5,
                        "name": "Table 5",
                        "kind": "graphic-frame",
                        "placeholder": null,
                        "position": { "left_emu": 0, "top_emu": 1_000_000 },
                        "size": { "width_emu": 2_000_000, "height_emu": 400_000 },
                        "rotation_degrees": 0.0,
                        "autofit": null,
                        "paragraphs": null,
                        "table": { "rows": 1, "columns": 2 },
                        "children": [],
                    },
                ],
            },
            { "id": 257, "name": null, "hidden": false, "shapes": 0, "shape_details": [] },
        ])
    );

    let grouped = temp.path.join("grouped.pptx");
    write_outline_deck(&grouped);
    let output = cli(&["inspect", grouped.to_str().unwrap(), "--json"]);
    assert!(output.status.success());
    let value: serde_json::Value = serde_json::from_slice(&output.stdout).unwrap();
    let shapes = &value["slide_details"][0]["shape_details"];
    assert_eq!(
        shapes[1]["placeholder"],
        json!({ "type": "ctrTitle", "idx": 0 })
    );
    assert_eq!(shapes[3]["kind"], "group");
    assert_eq!(shapes[3]["paragraphs"], serde_json::Value::Null);
    let children = shapes[3]["children"].as_array().unwrap();
    assert_eq!(children.len(), 1);
    assert_eq!(children[0]["index"], 0);
    assert_eq!(children[0]["id"], 6);
    assert_eq!(children[0]["paragraphs"][0]["text"], "Grouped\u{b}item");
}

#[test]
fn outline_json_reports_the_plain_outline_titles_levels_and_notes() {
    let temp = TempWorkspace::new("outline-json");
    let deck = temp.path.join("structured.pptx");
    write_structured_deck(&deck);
    add_speaker_notes(&deck, "Presenter reminder");

    let plain = cli(&["outline", deck.to_str().unwrap()]);
    assert!(plain.status.success());
    assert_eq!(
        String::from_utf8(plain.stdout).unwrap(),
        "Slide 1: Roadmap\n- plain rich\n  - nested\n- A1\n- B1\nSlide 2\n"
    );

    let output = cli(&["outline", deck.to_str().unwrap(), "--json"]);
    assert!(
        output.status.success(),
        "outline --json failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let with_notes_flag = cli(&["outline", deck.to_str().unwrap(), "--json", "--notes"]);
    assert_eq!(with_notes_flag.stdout, output.stdout);
    let value: serde_json::Value = serde_json::from_slice(&output.stdout).unwrap();
    assert_eq!(
        value,
        json!({
            "schema": 1,
            "slides": [
                {
                    "slide": 1,
                    "id": 256,
                    "title": "Roadmap",
                    "items": [
                        { "level": 0, "text": "plain rich" },
                        { "level": 1, "text": "nested" },
                        { "level": 0, "text": "A1" },
                        { "level": 0, "text": "B1" },
                    ],
                    "notes": "Presenter reminder",
                },
                { "slide": 2, "id": 257, "title": null, "items": [], "notes": null },
            ],
        })
    );
}

#[test]
fn plain_text_and_outline_print_speaker_notes_only_with_the_notes_flag() {
    let temp = TempWorkspace::new("notes-flag");
    let deck = temp.path.join("notes.pptx");
    write_deck(&deck, &["first slide", "second slide"]);
    // The helper writes its text into one run, so closing that run opens an
    // empty notes paragraph and a second non-empty one.
    add_speaker_notes(
        &deck,
        "Presenter reminder</a:t></a:r></a:p><a:p/><a:p><a:r><a:t>Second line",
    );
    let path = deck.to_str().unwrap();

    for (args, expected) in [
        (vec!["text", path], "first slide\nsecond slide\n"),
        (
            vec!["text", path, "--notes"],
            "first slide\nNotes: Presenter reminder\nNotes: Second line\nsecond slide\n",
        ),
        (
            vec!["outline", path],
            "Slide 1\n- first slide\nSlide 2\n- second slide\n",
        ),
        (
            vec!["outline", path, "--notes"],
            "Slide 1\n- first slide\nNotes: Presenter reminder\nNotes: Second line\nSlide 2\n- second slide\n",
        ),
    ] {
        let output = cli(&args);
        assert!(output.status.success(), "{args:?} failed");
        assert_eq!(
            String::from_utf8(output.stdout).unwrap(),
            expected,
            "{args:?}"
        );
    }
}

#[test]
fn comment_commands_round_trip_one_resolved_thread() {
    let temp = TempWorkspace::new("comment-round-trip");
    let input = temp.path.join("input.pptx");
    let added = temp.path.join("added.pptx");
    let replied = temp.path.join("replied.pptx");
    let answered = temp.path.join("answered.pptx");
    let resolved = temp.path.join("resolved.pptx");
    let reply_removed = temp.path.join("reply-removed.pptx");
    let removed = temp.path.join("removed.pptx");
    write_deck(&input, &["first", "second"]);
    let input_bytes = fs::read(&input).unwrap();
    let arg = |path: &Path| path.to_str().unwrap().to_owned();
    let thread = "{00000000-0000-4000-8000-000000000002}";
    let bob_reply = "{00000000-0000-4000-8000-000000000004}";
    let ada_reply = "{00000000-0000-4000-8000-000000000005}";
    let json_of = |output: Output| -> serde_json::Value {
        assert!(
            output.status.success(),
            "comment command failed: {}",
            String::from_utf8_lossy(&output.stderr)
        );
        serde_json::from_slice(&output.stdout).unwrap()
    };

    let listed = cli(&["comment", "list", &arg(&input)]);
    assert_eq!(String::from_utf8(listed.stdout).unwrap(), "(no comments)\n");
    assert_eq!(
        json_of(cli(&["comment", "list", &arg(&input), "--json"])),
        json!({ "schema": 1, "comments": [] })
    );

    let record = json_of(cli(&[
        "comment",
        "add",
        &arg(&input),
        "--slide",
        "2",
        "--author",
        "Ada",
        "--initials",
        "AL",
        "--text",
        "Check this",
        "--date",
        "2026-09-25T10:00:00Z",
        "--output",
        &arg(&added),
        "--json",
    ]));
    assert_eq!(
        record,
        json!({
            "schema": 1,
            "action": "add",
            "comment_id": thread,
            "slide": 2,
            "output": arg(&added),
        })
    );

    let record = json_of(cli(&[
        "comment",
        "reply",
        &arg(&added),
        "--id",
        thread,
        "--author",
        "Bob",
        "--text",
        "Agreed",
        "--date",
        "2026-09-25T11:00:00Z",
        "--output",
        &arg(&replied),
        "--json",
    ]));
    assert_eq!(
        record,
        json!({
            "schema": 1,
            "action": "reply",
            "comment_id": bob_reply,
            "parent_id": thread,
            "slide": 2,
            "output": arg(&replied),
        })
    );
    let answered_output = cli(&[
        "comment",
        "reply",
        &arg(&replied),
        "--id",
        thread,
        "--author",
        "Ada",
        "--text",
        "Thanks",
        "--date",
        "2026-09-25T12:00:00Z",
        "--output",
        &arg(&answered),
    ]);
    assert!(answered_output.status.success());
    assert_eq!(
        String::from_utf8(answered_output.stdout).unwrap(),
        format!("reply\nWritten to {}\n", answered.display())
    );

    let resolved_output = cli(&[
        "comment",
        "resolve",
        &arg(&answered),
        "--id",
        thread,
        "--output",
        &arg(&resolved),
    ]);
    assert!(resolved_output.status.success());
    assert_eq!(
        String::from_utf8(resolved_output.stdout).unwrap(),
        format!("resolve\nWritten to {}\n", resolved.display())
    );
    let presentation = Presentation::open(&resolved).unwrap();
    assert_eq!(
        presentation
            .comment_authors()
            .iter()
            .map(|author| (author.name.as_str(), author.initials.as_deref()))
            .collect::<Vec<_>>(),
        [("Ada", Some("AL")), ("Bob", None)]
    );
    let entry = |id: &str, author: &str, initials: Option<&str>, date: &str, text: &str| {
        json!({
            "slide": 2,
            "id": id,
            "author": author,
            "initials": initials,
            "date": date,
            "text": text,
            "parent_id": (id != thread).then_some(thread),
            "resolved": id == thread,
            "status": (id == thread).then_some("resolved"),
        })
    };
    assert_eq!(
        json_of(cli(&["comment", "list", &arg(&resolved), "--json"])),
        json!({
            "schema": 1,
            "comments": [
                entry(thread, "Ada", Some("AL"), "2026-09-25T10:00:00Z", "Check this"),
                entry(bob_reply, "Bob", None, "2026-09-25T11:00:00Z", "Agreed"),
                entry(ada_reply, "Ada", Some("AL"), "2026-09-25T12:00:00Z", "Thanks"),
            ],
        })
    );

    let record = json_of(cli(&[
        "comment",
        "remove",
        &arg(&resolved),
        "--id",
        bob_reply,
        "--output",
        &arg(&reply_removed),
        "--json",
    ]));
    assert_eq!(
        record,
        json!({
            "schema": 1,
            "action": "remove",
            "comment_id": bob_reply,
            "slide": 2,
            "output": arg(&reply_removed),
        })
    );
    let listed = cli(&["comment", "list", &arg(&reply_removed)]);
    assert_eq!(
        String::from_utf8(listed.stdout).unwrap(),
        format!("2\t{thread}\tAda\tresolved\tCheck this\n2\t{ada_reply}\tAda\topen\tThanks\n")
    );

    let removed_output = cli(&[
        "comment",
        "remove",
        &arg(&reply_removed),
        "--id",
        thread,
        "--output",
        &arg(&removed),
    ]);
    assert!(removed_output.status.success());
    let listed = cli(&["comment", "list", &arg(&removed)]);
    assert_eq!(String::from_utf8(listed.stdout).unwrap(), "(no comments)\n");

    for output in [
        &added,
        &replied,
        &answered,
        &resolved,
        &reply_removed,
        &removed,
    ] {
        let validated = cli(&["validate", &arg(output)]);
        assert!(
            validated.status.success(),
            "validate failed for {}: {}",
            output.display(),
            String::from_utf8_lossy(&validated.stderr)
        );
    }
    assert_eq!(fs::read(&input).unwrap(), input_bytes);
}

#[test]
fn comment_list_reports_a_closed_thread_as_closed_rather_than_open() {
    let temp = TempWorkspace::new("comment-closed");
    let deck = temp.path.join("closed.pptx");
    let author = "{11111111-1111-1111-1111-111111111111}";
    let thread = "{22222222-2222-2222-2222-222222222222}";
    let mut presentation = Presentation::new().expect("open bundled template");
    presentation.add_slide(6).expect("add blank slide");
    presentation
        .add_comment_author(CommentAuthor::new(author, "Ada", None, "Ada", "None").unwrap())
        .unwrap();
    let mut comment = Comment::new(thread, author, "2026-09-25T10:00:00Z", "Done").unwrap();
    comment.status = Some("closed".to_owned());
    presentation.add_comment(0, comment).unwrap();
    presentation.save(&deck).expect("write closed thread deck");

    let listed = cli(&["comment", "list", deck.to_str().unwrap()]);
    assert!(listed.status.success());
    assert_eq!(
        String::from_utf8(listed.stdout).unwrap(),
        format!("1\t{thread}\tAda\tclosed\tDone\n")
    );
    let listed = cli(&["comment", "list", deck.to_str().unwrap(), "--json"]);
    assert!(listed.status.success());
    let value: serde_json::Value = serde_json::from_slice(&listed.stdout).unwrap();
    assert_eq!(value["comments"][0]["status"], "closed");
    assert_eq!(value["comments"][0]["resolved"], false);
}

#[test]
fn comment_mutations_fail_without_creating_output() {
    let temp = TempWorkspace::new("comment-failures");
    let input = temp.path.join("input.pptx");
    let commented = temp.path.join("commented.pptx");
    let replied = temp.path.join("replied.pptx");
    let output = temp.path.join("output.pptx");
    write_deck(&input, &["only slide"]);
    let thread = "{00000000-0000-4000-8000-000000000002}";
    let reply = "{00000000-0000-4000-8000-000000000004}";
    let unknown = "{99999999-9999-9999-9999-999999999999}";
    let date = "2026-09-25T10:00:00Z";
    let (input, commented, replied, output_path) = (
        input.to_str().unwrap(),
        commented.to_str().unwrap(),
        replied.to_str().unwrap(),
        output.to_str().unwrap(),
    );
    for args in [
        vec![
            "comment", "add", input, "--slide", "1", "--author", "Ada", "--text", "Thread",
            "--date", date, "--output", commented,
        ],
        vec![
            "comment", "reply", commented, "--id", thread, "--author", "Bob", "--text", "Reply",
            "--date", date, "--output", replied,
        ],
    ] {
        assert!(cli(&args).status.success(), "{args:?} failed");
    }

    for (args, message) in [
        (
            vec![
                "add", input, "--slide", "0", "--author", "Ada", "--text", "x", "--date", date,
            ],
            "slide 0 is out of range for 1 slides".to_owned(),
        ),
        (
            vec![
                "add", input, "--slide", "2", "--author", "Ada", "--text", "x", "--date", date,
            ],
            "slide 2 is out of range for 1 slides".to_owned(),
        ),
        (
            vec![
                "add",
                input,
                "--slide",
                "1",
                "--author",
                "Ada",
                "--text",
                "x",
                "--date",
                "yesterday",
            ],
            "comment timestamp is not RFC 3339: yesterday".to_owned(),
        ),
        (
            vec![
                "reply", replied, "--id", unknown, "--author", "Ada", "--text", "x", "--date", date,
            ],
            format!("comment id {unknown} does not exist"),
        ),
        (
            vec![
                "reply", replied, "--id", reply, "--author", "Ada", "--text", "x", "--date", date,
            ],
            format!("comment id {reply} is a reply in thread {thread}"),
        ),
        (
            vec!["resolve", replied, "--id", reply],
            format!("comment id {reply} is a reply in thread {thread}"),
        ),
        (
            vec!["resolve", replied, "--id", unknown],
            format!("comment id {unknown} does not exist"),
        ),
        (
            vec!["remove", replied, "--id", unknown],
            format!("comment id {unknown} does not exist"),
        ),
        (
            vec![
                "add",
                input,
                "--slide",
                "1",
                "--author",
                "Ada",
                "--text",
                "one\u{b}two",
                "--date",
                date,
            ],
            "comment text contains U+000B, which XML 1.0 cannot carry".to_owned(),
        ),
        (
            vec![
                "add", input, "--slide", "1", "--author", "A\u{1}da", "--text", "x", "--date", date,
            ],
            "comment author contains U+0001, which XML 1.0 cannot carry".to_owned(),
        ),
        (
            vec![
                "add",
                input,
                "--slide",
                "1",
                "--author",
                "Ada",
                "--initials",
                "A\u{1f}",
                "--text",
                "x",
                "--date",
                date,
            ],
            "comment initials contains U+001F, which XML 1.0 cannot carry".to_owned(),
        ),
        (
            vec![
                "reply", replied, "--id", thread, "--author", "Ada", "--text", "x\u{c}", "--date",
                date,
            ],
            "comment text contains U+000C, which XML 1.0 cannot carry".to_owned(),
        ),
    ] {
        let mut command = vec!["comment"];
        command.extend(&args);
        command.extend(["--output", output_path, "--json"]);
        let result = cli(&command);
        assert_eq!(result.status.code(), Some(1), "{args:?} succeeded");
        assert!(result.stdout.is_empty(), "{args:?} printed a record");
        assert!(
            String::from_utf8_lossy(&result.stderr).contains(&message),
            "{args:?} reported {}",
            String::from_utf8_lossy(&result.stderr)
        );
        assert!(!output.exists(), "{args:?} created output");
    }

    fs::write(&output, b"keep me").unwrap();
    let result = cli(&[
        "comment",
        "resolve",
        replied,
        "--id",
        thread,
        "--output",
        output_path,
    ]);
    assert_eq!(result.status.code(), Some(1));
    assert!(String::from_utf8_lossy(&result.stderr).contains("output already exists"));
    assert_eq!(fs::read(&output).unwrap(), b"keep me");
}
