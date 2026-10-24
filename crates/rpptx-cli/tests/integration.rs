use std::fs;
use std::io::{Cursor, Read};
use std::path::{Path, PathBuf};
use std::process::{Command, Output, Stdio};
use std::sync::atomic::{AtomicUsize, Ordering};

use oxml_opc::relationship::rel_types;
use oxml_opc::{OpcPackage, content_types};
use rpptx::{
    Angle, CT_TextCharacterProperties, Comment, CommentAuthor, Emu, Presentation,
    PresentationPackageClass,
};
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

/// Runs the CLI while its reader takes a short prefix of standard output and
/// then closes it, as `| head -1` does.
fn cli_with_closed_stdout(args: &[&str]) -> Output {
    let mut child = Command::new(env!("CARGO_BIN_EXE_rpptx"))
        .args(args)
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .expect("spawn rpptx CLI");
    let mut stdout = child.stdout.take().expect("piped stdout");
    stdout
        .read_exact(&mut [0; 64])
        .expect("read an output prefix");
    drop(stdout);
    child.wait_with_output().expect("wait for rpptx CLI")
}

/// Asserts that `args` refuses to replace an existing `output` and leaves it
/// byte-identical, and that the same command with `--force` replaces it.
fn assert_existing_output_needs_force(args: &[&str], output: &Path) {
    fs::write(output, b"keep me").unwrap();
    let refused = cli(args);
    assert_eq!(refused.status.code(), Some(1), "{args:?} was not refused");
    assert_eq!(
        String::from_utf8_lossy(&refused.stderr),
        format!(
            "Error: output already exists: {} (pass --force to replace it)\n",
            output.display()
        )
    );
    assert_eq!(fs::read(output).unwrap(), b"keep me");

    let forced = cli(&[args, &["--force"]].concat());
    assert!(
        forced.status.success(),
        "{args:?} --force failed: {}",
        String::from_utf8_lossy(&forced.stderr)
    );
    assert_ne!(fs::read(output).unwrap(), b"keep me");
    assert!(
        fs::read_dir(output.parent().unwrap())
            .unwrap()
            .all(|entry| {
                !entry
                    .unwrap()
                    .file_name()
                    .to_string_lossy()
                    .ends_with(".tmp")
            })
    );
}

/// Asserts that `args`, whose output is `input` under some spelling, leaves
/// the input byte-identical with and without `--force`.
fn assert_input_is_never_replaced(args: &[&str], input: &Path, output: &str) {
    let before = fs::read(input).unwrap();
    for force in [&[][..], &["--force"]] {
        let refused = cli(&[args, force].concat());
        assert_eq!(refused.status.code(), Some(1), "{args:?} {force:?}");
        assert_eq!(
            String::from_utf8_lossy(&refused.stderr),
            format!("Error: output is the input file: {output}\n")
        );
        assert_eq!(fs::read(input).unwrap(), before, "{args:?} {force:?}");
    }
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
fn validate_reports_duplicate_paragraph_properties_while_text_remains_readable() {
    let temp = TempWorkspace::new("second-ppr");
    let deck = temp.path.join("second-ppr.pptx");
    write_deck(&deck, &["One. Two."]);
    let mut package = OpcPackage::open(&deck).expect("open second pPr package");
    let slide_part = package
        .content_types
        .overrides
        .iter()
        .find_map(|(part, content_type)| {
            (content_type
                == "application/vnd.openxmlformats-officedocument.presentationml.slide+xml")
                .then_some(part.clone())
        })
        .expect("second pPr slide part");
    let xml = String::from_utf8(package.get_part(&slide_part).unwrap().to_vec()).unwrap();
    let text = xml.find("One. Two.").expect("text box paragraph");
    let start = xml[..text].rfind("<a:p>").expect("paragraph start");
    let end = text + xml[text..].find("</a:p>").expect("paragraph end") + "</a:p>".len();
    let paragraph = r#"<a:p><a:pPr algn="l"/><a:r><a:t>One. </a:t></a:r><a:pPr algn="l"/><a:r><a:t>Two.</a:t></a:r></a:p>"#;
    let xml = format!("{}{paragraph}{}", &xml[..start], &xml[end..]);
    package.set_part(&slide_part, xml.into_bytes());
    package.save(&deck).expect("write second pPr fixture");

    let validated = cli(&["validate", deck.to_str().unwrap()]);
    assert!(!validated.status.success());
    let diagnostic = String::from_utf8(validated.stderr).unwrap();
    assert!(diagnostic.contains("InvalidXmlChild"), "{diagnostic}");
    assert!(diagnostic.contains(&slide_part), "{diagnostic}");
    assert!(
        diagnostic.contains("finite XML child order or multiplicity is invalid"),
        "{diagnostic}"
    );
    let text = cli(&["text", deck.to_str().unwrap()]);
    assert!(
        text.status.success(),
        "{}",
        String::from_utf8_lossy(&text.stderr)
    );
    assert_eq!(String::from_utf8(text.stdout).unwrap(), "One. Two.\n");
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
fn printing_commands_end_cleanly_when_the_reader_closes_stdout() {
    let temp = TempWorkspace::new("closed-stdout");
    let deck = temp.path.join("long.pptx");
    // Several hundred kilobytes outgrow a default pipe buffer, so the CLI
    // is still writing when its reader goes away.
    let text = "lorem ipsum dolor sit amet ".repeat(40);
    write_deck(&deck, &vec![text.as_str(); 300]);
    let deck = deck.to_str().unwrap();

    for args in [vec!["text", deck], vec!["outline", deck, "--json"]] {
        let output = cli_with_closed_stdout(&args);
        let stderr = String::from_utf8_lossy(&output.stderr);
        assert!(output.status.success(), "{args:?} failed: {stderr}");
        assert!(stderr.is_empty(), "{args:?} wrote stderr: {stderr}");
    }
}

#[test]
fn validate_keeps_its_verdict_when_the_reader_closes_stdout() {
    let temp = TempWorkspace::new("validate-closed-stdout");
    let deck = temp.path.join("valid.pptx");
    write_deck(&deck, &["valid"]);

    // The reader goes away before the CLI has opened the deck, so the pass
    // line meets a closed pipe.
    let mut child = Command::new(env!("CARGO_BIN_EXE_rpptx"))
        .args(["validate", deck.to_str().unwrap()])
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .expect("spawn rpptx CLI");
    drop(child.stdout.take());
    let output = child.wait_with_output().expect("wait for rpptx CLI");
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(output.status.success(), "validate failed: {stderr}");
    assert!(stderr.is_empty(), "validate wrote stderr: {stderr}");
}

#[cfg(target_os = "linux")]
#[test]
fn a_failing_stdout_other_than_a_closed_pipe_is_an_error() {
    let temp = TempWorkspace::new("full-stdout");
    let deck = temp.path.join("short.pptx");
    write_deck(&deck, &["No space left"]);
    let full = fs::OpenOptions::new()
        .write(true)
        .open("/dev/full")
        .expect("open /dev/full");

    let output = Command::new(env!("CARGO_BIN_EXE_rpptx"))
        .args(["text", deck.to_str().unwrap()])
        .stdout(full)
        .output()
        .expect("run rpptx CLI");
    assert_eq!(output.status.code(), Some(1));
    let stderr = String::from_utf8(output.stderr).unwrap();
    assert!(stderr.starts_with("Error: "), "{stderr}");
    assert!(!stderr.contains("panicked"), "{stderr}");
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

    let mut presentation = Presentation::open(&deck).unwrap();
    *presentation.core_properties_mut() = rpptx::CoreProperties {
        category: Some("Reports".to_owned()),
        revision: Some("7".to_owned()),
        ..Default::default()
    };
    presentation.save(&deck).unwrap();
    let inspected = cli(&["inspect", deck.to_str().unwrap()]);
    assert!(inspected.status.success());
    let inspected = String::from_utf8(inspected.stdout).unwrap();
    assert!(
        inspected.contains("Metadata:\n  Category: Reports\n  Revision: 7\n\n"),
        "{inspected}"
    );
    let inspected = cli(&["inspect", deck.to_str().unwrap(), "--json"]);
    let value: serde_json::Value = serde_json::from_slice(&inspected.stdout).unwrap();
    assert_eq!(value["metadata"]["category"], "Reports");
    assert_eq!(value["metadata"]["revision"], "7");
    assert!(value["metadata"]["title"].is_null());

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
fn convert_replaces_an_existing_output_only_with_force_and_never_its_input() {
    let temp = TempWorkspace::new("convert-output-policy");
    let input = temp.path.join("deck.pptx");
    write_deck(&input, &["Output policy"]);
    let deck = input.to_str().unwrap();
    let spelled = temp.path.join("slides/../deck.pptx");
    fs::create_dir(temp.path.join("slides")).unwrap();

    let default_pdf = temp.path.join("deck.pdf");
    assert_existing_output_needs_force(&["convert", deck, "--to", "pdf"], &default_pdf);
    for to in ["pdf", "png", "tiff"] {
        let output = temp.path.join(format!("converted.{to}"));
        assert_existing_output_needs_force(
            &[
                "convert",
                deck,
                "--to",
                to,
                "--dpi",
                "24",
                "-o",
                output.to_str().unwrap(),
            ],
            &output,
        );
        for spelling in [deck, spelled.to_str().unwrap()] {
            assert_input_is_never_replaced(
                &["convert", deck, "--to", to, "--dpi", "24", "-o", spelling],
                &input,
                spelling,
            );
        }
    }
}

#[test]
fn render_replaces_existing_slides_only_with_force_and_never_its_input() {
    let temp = TempWorkspace::new("render-output-policy");
    let input = temp.path.join("deck.pptx");
    let slides = temp.path.join("slides");
    write_deck(&input, &["Output policy"]);
    fs::create_dir(&slides).unwrap();
    let render = [
        "render",
        input.to_str().unwrap(),
        "-o",
        slides.to_str().unwrap(),
        "--dpi",
        "24",
    ];

    assert_existing_output_needs_force(&render, &slides.join("deck_slide1.png"));
    assert_existing_output_needs_force(
        &[render.as_slice(), &["--format", "tiff"]].concat(),
        &slides.join("deck.tiff"),
    );

    // A deck named like its own TIFF output, rendered into its folder.
    let named = temp.path.join("named.tiff");
    fs::copy(&input, &named).unwrap();
    assert_input_is_never_replaced(
        &[
            "render",
            named.to_str().unwrap(),
            "-o",
            temp.path.to_str().unwrap(),
            "--format",
            "tiff",
            "--dpi",
            "24",
        ],
        &named,
        named.to_str().unwrap(),
    );
}

#[test]
fn thumbnail_replaces_an_existing_output_only_with_force_and_never_its_input() {
    let temp = TempWorkspace::new("thumbnail-output-policy");
    let input = temp.path.join("deck.pptx");
    write_deck(&input, &["Output policy"]);
    let deck = input.to_str().unwrap();
    let chosen = temp.path.join("chosen.png");
    let spelled = temp.path.join("slides/../deck.pptx");
    fs::create_dir(temp.path.join("slides")).unwrap();

    assert_existing_output_needs_force(&["thumbnail", deck], &temp.path.join("deck.png"));
    assert_existing_output_needs_force(
        &["thumbnail", deck, "-o", chosen.to_str().unwrap()],
        &chosen,
    );
    for spelling in [deck, spelled.to_str().unwrap()] {
        assert_input_is_never_replaced(&["thumbnail", deck, "-o", spelling], &input, spelling);
    }
    #[cfg(unix)]
    {
        let refused = cli(&["thumbnail", deck, "-o", "/dev/null", "--force"]);
        assert_eq!(refused.status.code(), Some(1));
        assert_eq!(
            String::from_utf8_lossy(&refused.stderr),
            "Error: output is not a regular file: /dev/null\n"
        );
    }
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
fn replace_output_extension_selects_the_package_class() {
    let temp = TempWorkspace::new("replace-class");
    let source = temp.path.join("template.potx");
    write_deck(&source, &["Hello {{name}}"]);
    assert_eq!(
        Presentation::open(&source)
            .unwrap()
            .package_class()
            .unwrap(),
        PresentationPackageClass::Template
    );

    let output = temp.path.join("deck.pptx");
    let result = cli(&[
        "replace",
        source.to_str().unwrap(),
        "--placeholder",
        "{{name}}",
        "--value",
        "Reader",
        "--output",
        output.to_str().unwrap(),
    ]);
    assert!(
        result.status.success(),
        "{}",
        String::from_utf8_lossy(&result.stderr)
    );
    assert_eq!(
        Presentation::open(&output)
            .unwrap()
            .package_class()
            .unwrap(),
        PresentationPackageClass::Presentation
    );
}

#[test]
fn replace_refuses_a_macro_free_output_for_a_vba_input() {
    let temp = TempWorkspace::new("replace-vba");
    let fixture = temp.path.join("fixture.pptx");
    let source = temp.path.join("macros.pptm");
    write_deck(&fixture, &["Hello {{name}}"]);
    let mut package = OpcPackage::open(&fixture).unwrap();
    fs::remove_file(&fixture).unwrap();
    package
        .get_or_create_part_rels("/ppt/presentation.xml")
        .add(rel_types::VBA_PROJECT, "vbaProject.bin");
    package.set_part("/ppt/vbaProject.bin", b"vba-project".to_vec());
    package.content_types.add_override(
        "/ppt/vbaProject.bin",
        "application/vnd.ms-office.vbaProject",
    );
    package.content_types.add_override(
        "/ppt/presentation.xml",
        content_types::PRESENTATION_MACRO_ENABLED,
    );
    package.save(&source).unwrap();

    let output = temp.path.join("deck.pptx");
    let result = cli(&[
        "replace",
        source.to_str().unwrap(),
        "--placeholder",
        "{{name}}",
        "--value",
        "Reader",
        "--output",
        output.to_str().unwrap(),
    ]);
    assert_eq!(result.status.code(), Some(1));
    assert!(String::from_utf8_lossy(&result.stderr).contains("VBA project"));
    let entries: Vec<_> = fs::read_dir(&temp.path)
        .unwrap()
        .map(|entry| entry.unwrap().file_name())
        .collect();
    assert_eq!(entries, [source.file_name().unwrap()]);
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
                        "name": "Title 1",
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
                        "name": "Content Placeholder 2",
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

#[test]
fn fx195_font_directory_validation_precedes_publication() {
    let workspace = TempWorkspace::new("fx195-font-dir");
    let input = workspace.path.join("input.pptx");
    write_deck(&input, &["caller directory atomicity"]);
    let original = fs::read(&input).unwrap();
    let fonts = workspace.path.join("fonts");
    fs::create_dir(&fonts).unwrap();
    fs::write(fonts.join("bad.ttf"), b"not a font").unwrap();
    let output = workspace.path.join("out.png");
    for bad in [fonts.clone(), workspace.path.join("missing"), input.clone()] {
        for mode in ["convert", "render"] {
            fs::write(&output, b"sentinel").unwrap();
            let result = cli(&[
                mode,
                input.to_str().unwrap(),
                "--font-dir",
                bad.to_str().unwrap(),
                if mode == "convert" {
                    "--to"
                } else {
                    "--format"
                },
                "png",
                "--output",
                output.to_str().unwrap(),
                "--force",
            ]);
            assert!(
                !result.status.success(),
                "{mode} accepted invalid font input"
            );
            assert!(
                String::from_utf8_lossy(&result.stderr).contains("font"),
                "font validation, not argument parsing: {}",
                String::from_utf8_lossy(&result.stderr)
            );
            assert_eq!(fs::read(&output).unwrap(), b"sentinel");
            assert_eq!(fs::read(&input).unwrap(), original);
        }
    }
    fs::remove_file(fonts.join("bad.ttf")).unwrap();
    fs::write(
        fonts.join("Carlito.ttf"),
        include_bytes!("../../oxml-layout/fonts/Carlito-Regular.ttf"),
    )
    .unwrap();
    let result = cli(&[
        "convert",
        input.to_str().unwrap(),
        "--font-dir",
        fonts.to_str().unwrap(),
        "--to",
        "png",
        "--output",
        output.to_str().unwrap(),
        "--force",
    ]);
    assert!(
        result.status.success(),
        "{}",
        String::from_utf8_lossy(&result.stderr)
    );
    assert!(fs::read(&output).unwrap().starts_with(b"\x89PNG\r\n\x1a\n"));
    assert_eq!(fs::read(&input).unwrap(), original);
}

fn fx195_caller_only_monospace() -> Vec<u8> {
    let mut bytes = include_bytes!("../../oxml-layout/fonts/LiberationMono-Regular.ttf").to_vec();
    for (from, to) in [
        (b"Liberation Mono".to_vec(), b"FX195CallerOnly".to_vec()),
        (
            "Liberation Mono"
                .encode_utf16()
                .flat_map(u16::to_be_bytes)
                .collect(),
            "FX195CallerOnly"
                .encode_utf16()
                .flat_map(u16::to_be_bytes)
                .collect(),
        ),
    ] {
        assert_eq!(from.len(), to.len());
        let mut count = 0;
        let mut position = 0;
        while let Some(relative) = bytes[position..]
            .windows(from.len())
            .position(|value| value == from)
        {
            let offset = position + relative;
            bytes[offset..offset + from.len()].copy_from_slice(&to);
            position = offset + from.len();
            count += 1;
        }
        assert!(count > 0);
    }
    assert!(
        oxml_layout::bundled_fonts::bundled_font_data()
            .iter()
            .all(|(name, data)| *name != "FX195CallerOnly" && *data != bytes)
    );
    bytes
}

#[test]
fn fx195_every_ppt_cli_entrypoint_uses_caller_only_face() {
    let workspace = TempWorkspace::new("fx195-caller-matrix");
    let input = workspace.path.join("caller.pptx");
    let fonts = workspace.path.join("fonts");
    fs::create_dir(&fonts).unwrap();
    let bytes = fx195_caller_only_monospace();
    fs::write(fonts.join("FilenameAlias.ttf"), &bytes).unwrap();
    let mut presentation = Presentation::new().unwrap();
    presentation.add_slide(6).unwrap();
    {
        let mut slide = presentation.slide_mut(0).unwrap();
        let mut shape = slide
            .add_textbox(Emu(100_000), Emu(100_000), Emu(3_000_000), Emu(800_000))
            .unwrap();
        shape.set_text("iiiiWWWW caller matrix").unwrap();
        shape
            .text_frame()
            .unwrap()
            .paragraph_mut(0)
            .unwrap()
            .run_mut(0)
            .unwrap()
            .set_font(Some(rpptx::TextFont::new("FX195CallerOnly").unwrap()));
    }
    presentation.save(&input).unwrap();
    let before = fs::read(&input).unwrap();
    let refs = [("FilenameAlias", bytes.as_slice())];
    let layout = presentation
        .render_deterministic_with_fonts(&refs)
        .unwrap()
        .1;
    assert!(
        layout
            .fonts
            .iter()
            .any(|font| font.data.as_ref() == bytes.as_slice() && font.face_index == 0)
    );
    let ordinary = presentation.render_deterministic().unwrap().1;
    assert_ne!(
        format!("{:?}", ordinary.pages),
        format!("{:?}", layout.pages)
    );
    for format in ["pdf", "png", "jpeg", "tiff"] {
        let out = workspace.path.join(format!("converted.{format}"));
        let converted = cli(&[
            "convert",
            input.to_str().unwrap(),
            "--to",
            format,
            "--font-dir",
            fonts.to_str().unwrap(),
            "--dpi",
            "72",
            "--output",
            out.to_str().unwrap(),
        ]);
        assert!(
            converted.status.success(),
            "{}",
            String::from_utf8_lossy(&converted.stderr)
        );
        let expected = if format == "pdf" {
            oxml_pdf::render_to_pdf_checked(&layout).unwrap()
        } else {
            let raster_format = match format {
                "png" => oxml_pdf::RasterFormat::Png {
                    transparent_background: false,
                },
                "jpeg" => oxml_pdf::RasterFormat::Jpeg { quality: 90 },
                _ => oxml_pdf::RasterFormat::Tiff,
            };
            match oxml_pdf::render_pages(
                &layout,
                &[0],
                oxml_pdf::RasterOptions {
                    dpi: 72.0,
                    format: raster_format,
                },
            )
            .unwrap()
            {
                oxml_pdf::RasterOutput::SeparatePages(mut pages) => pages.remove(0),
                oxml_pdf::RasterOutput::MultiPageTiff(data) => data,
            }
        };
        assert_eq!(
            fs::read(&out).unwrap(),
            expected,
            "{format} selected caller paint"
        );
        if format != "pdf" {
            let directory = workspace.path.join(format!("render-{format}"));
            let rendered = cli(&[
                "render",
                input.to_str().unwrap(),
                "--format",
                format,
                "--font-dir",
                fonts.to_str().unwrap(),
                "--dpi",
                "72",
                "--slide",
                "1",
                "--output",
                directory.to_str().unwrap(),
            ]);
            assert!(
                rendered.status.success(),
                "{}",
                String::from_utf8_lossy(&rendered.stderr)
            );
            let filename = match format {
                "tiff" => "caller.tiff",
                "jpeg" => "caller_slide1.jpg",
                _ => "caller_slide1.png",
            };
            assert_eq!(fs::read(directory.join(filename)).unwrap(), expected);
        }
    }
    let thumbnail = workspace.path.join("thumb.png");
    let result = cli(&[
        "thumbnail",
        input.to_str().unwrap(),
        "--font-dir",
        fonts.to_str().unwrap(),
        "--output",
        thumbnail.to_str().unwrap(),
    ]);
    assert!(
        result.status.success(),
        "{}",
        String::from_utf8_lossy(&result.stderr)
    );
    let dpi = (320.0 - 1e-9) * 72.0 / layout.pages[0].width;
    assert_eq!(
        fs::read(&thumbnail).unwrap(),
        oxml_pdf::render_page_to_png(&layout, 0, dpi).unwrap()
    );
    assert_eq!(fs::read(&input).unwrap(), before);
}

fn fx206_member_scene(placed: bool, reason: usize) -> String {
    let member = |id: u32, x: i64, width: i64, properties: &str, style: &str, positioned: bool| {
        let transform = if positioned {
            format!(
                r#"<a:xfrm><a:off x="{x}" y="127000"/><a:ext cx="{width}" cy="254000"/></a:xfrm>"#
            )
        } else {
            String::new()
        };
        format!(
            r#"<p:sp><p:nvSpPr><p:cNvPr id="{id}" name="owner{id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>{transform}<a:prstGeom prst="rect"/>{properties}<f:opaque xmlns:f="urn:f206:member" literal="&#x41;"/></p:spPr>{style}</p:sp>"#
        )
    };
    let (properties, style) = match reason {
        0 => (
            "",
            r#"<p:style><a:lnRef idx="0"/><a:fillRef idx="999"/><a:effectRef idx="0"/><a:fontRef idx="minor"/></p:style>"#,
        ),
        1 => (
            r#"<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>"#,
            "",
        ),
        _ => (
            r#"<a:solidFill><a:prstClr val="futureColour"/></a:solidFill>"#,
            "",
        ),
    };
    let children = [
        member(
            1,
            127000,
            254000,
            r#"<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>"#,
            "",
            true,
        ),
        member(2, 127000, 254000, properties, style, placed),
        member(
            3,
            127000,
            50800,
            r#"<a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>"#,
            "",
            true,
        ),
    ]
    .join("");
    format!(
        r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" showMasterSp="0"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="10" name="root"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:grpSp><p:nvGrpSpPr><p:cNvPr id="11" name="parent"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="63500" y="63500"/><a:ext cx="2540000" cy="2540000"/><a:chOff x="0" y="0"/><a:chExt cx="1270000" cy="1270000"/></a:xfrm></p:grpSpPr>{children}</p:grpSp></p:spTree></p:cSld></p:sld>"#
    )
}

#[test]
fn fx206_cli_recovery_reports_fallback_omission_and_keeps_fatal_output_atomic() {
    let temp = TempWorkspace::new("member-recovery");
    let path = temp.path.join("input.pptx");
    for placed in [false, true] {
        write_deck(&path, &["seed"]);
        let mut package = OpcPackage::open(&path).unwrap();
        package.set_part(
            "/ppt/slides/slide1.xml",
            fx206_member_scene(placed, 0).into_bytes(),
        );
        package.save(&path).unwrap();
        let before = fs::read(&path).unwrap();
        let pdf = temp.path.join(if placed {
            "fallback.pdf"
        } else {
            "omission.pdf"
        });
        let output = cli(&[
            "convert",
            path.to_str().unwrap(),
            "--to",
            "pdf",
            "--output",
            pdf.to_str().unwrap(),
        ]);
        assert!(
            output.status.success(),
            "{}",
            String::from_utf8_lossy(&output.stderr)
        );
        assert!(fs::read(&pdf).unwrap().starts_with(b"%PDF-"));
        let stderr = String::from_utf8(output.stderr).unwrap();
        assert_eq!(stderr.matches("Warning: unresolved member ").count(), 1);
        assert!(stderr.contains("slide 1 incomplete: 1 unresolved members"));
        assert!(stderr.contains(if placed {
            "retained as bounds"
        } else {
            "omitted without trustworthy finite bounds"
        }));
        assert_eq!(fs::read(&path).unwrap(), before);
    }
    let bad = temp.path.join("fatal.pptx");
    fs::write(&bad, b"not a package").unwrap();
    let target = temp.path.join("preserved.pdf");
    fs::write(&target, b"original").unwrap();
    let output = cli(&[
        "convert",
        bad.to_str().unwrap(),
        "--to",
        "pdf",
        "--output",
        target.to_str().unwrap(),
        "--force",
    ]);
    assert!(!output.status.success());
    assert_eq!(fs::read(&target).unwrap(), b"original");
}

#[test]
fn fx208_footer_cli_safe_publication_and_inspection() {
    let temp = TempWorkspace::new("fx208-footer");
    let input = temp.path.join("input.pptx");
    let output = temp.path.join("footer.pptx");
    let mut presentation = Presentation::new().unwrap();
    presentation.add_slide(1).unwrap();
    std::fs::write(&input, presentation.to_bytes().unwrap()).unwrap();
    let before = std::fs::read(&input).unwrap();
    let help = cli(&["footer", "--help"]);
    assert!(help.status.success());
    let help = String::from_utf8(help.stdout).unwrap();
    for option in [
        "--date-cache",
        "--date-format",
        "--skip-title",
        "--slide-number",
        "--footer",
    ] {
        assert!(help.contains(option), "missing {option}");
    }
    let result = cli(&[
        "footer",
        input.to_str().unwrap(),
        "--footer",
        "CLI footer",
        "--date",
        "auto",
        "--date-cache",
        "cached",
        "--slide-number",
        "--output",
        output.to_str().unwrap(),
        "--json",
    ]);
    assert!(
        result.status.success(),
        "{}",
        String::from_utf8_lossy(&result.stderr)
    );
    let reopened = Presentation::open(&output).unwrap();
    assert_eq!(
        reopened.slide_header_footer(0).unwrap().footer.as_deref(),
        Some("CLI footer")
    );
    assert_eq!(std::fs::read(&input).unwrap(), before);
    let inspect = cli(&["inspect", output.to_str().unwrap(), "--json"]);
    assert!(inspect.status.success());
    let value: serde_json::Value = serde_json::from_slice(&inspect.stdout).unwrap();
    assert!(value.get("masters").is_some());
    assert!(value["masters"][0].get("header_footer").is_some());
    let invalid = cli(&[
        "footer",
        input.to_str().unwrap(),
        "--date",
        "off",
        "--date-cache",
        "cached",
        "--output",
        input.to_str().unwrap(),
        "--force",
    ]);
    assert!(!invalid.status.success());
    assert_eq!(std::fs::read(&input).unwrap(), before);
}

#[test]
fn powerpoint_replace_map_is_ordered_and_includes_notes() {
    let temp = TempWorkspace::new("fx204-base-map");
    let input = temp.path.join("input.pptx");
    let output = temp.path.join("output.pptx");
    let map = temp.path.join("map.json");
    write_deck(&input, &["A A"]);
    add_speaker_notes(&input, "A");
    fs::write(&map, r#"[{"placeholder":"A","value":"B","expect":3},{"placeholder":"B","value":"C","expect":3}]"#).unwrap();
    let result = cli(&[
        "replace",
        input.to_str().unwrap(),
        "--map",
        map.to_str().unwrap(),
        "--json",
        "--output",
        output.to_str().unwrap(),
    ]);
    assert!(
        result.status.success(),
        "{}",
        String::from_utf8_lossy(&result.stderr)
    );
    let report: serde_json::Value = serde_json::from_slice(&result.stdout).unwrap();
    assert_eq!(report["pairs"][0]["count"], 3);
    assert_eq!(report["pairs"][1]["count"], 3);
    assert_eq!(
        Presentation::open(&output)
            .unwrap()
            .slide(0)
            .unwrap()
            .text(),
        "C C"
    );
    let reopened = Presentation::open(&output).unwrap();
    assert_eq!(
        reopened.slide(0).unwrap().notes_text().as_deref(),
        Some("C")
    );
    assert!(
        OpcPackage::open(&output).unwrap().parts["/ppt/notesSlides/notesSlide1.xml"]
            .windows(b"urn:rdocx:test".len())
            .any(|w| w == b"urn:rdocx:test")
    );
    let before = fs::read(&input).unwrap();
    let rejected = temp.path.join("rejected.pptx");
    fs::write(&map, r#"[{"placeholder":"A","value":"B","expect":3},{"placeholder":"B","value":"C","expect":4}]"#).unwrap();
    let refused = cli(&[
        "replace",
        input.to_str().unwrap(),
        "--map",
        map.to_str().unwrap(),
        "--output",
        rejected.to_str().unwrap(),
    ]);
    assert!(!refused.status.success());
    assert!(!rejected.exists());
    assert_eq!(fs::read(&input).unwrap(), before);
}

#[test]
fn slide_add_resolves_layout_and_insertion_atomically() {
    let temp = TempWorkspace::new("fx204-add");
    let input = temp.path.join("input.pptx");
    write_deck(&input, &["first", "last"]);
    let before = fs::read(&input).unwrap();
    let layout_name = Presentation::open(&input)
        .unwrap()
        .layout_name(6)
        .unwrap()
        .to_owned();
    for (n, layout) in [(1, "7"), (2, layout_name.as_str()), (3, "7")] {
        let output = temp.path.join(format!("add-{n}.pptx"));
        let result = cli(&[
            "slide",
            "add",
            input.to_str().unwrap(),
            "--layout",
            layout,
            "--at",
            &n.to_string(),
            "--json",
            "--output",
            output.to_str().unwrap(),
        ]);
        assert!(
            result.status.success(),
            "{}",
            String::from_utf8_lossy(&result.stderr)
        );
        let report: serde_json::Value = serde_json::from_slice(&result.stdout).unwrap();
        assert_eq!(report["slide"], n);
        let deck = Presentation::open(&output).unwrap();
        assert_eq!(deck.len(), 3);
        assert_eq!(deck.slide(n as usize - 1).unwrap().text(), "");
    }
    for layout in [
        "0",
        "999",
        "blank",
        "unknown",
        "999999999999999999999999999999999999",
    ] {
        let output = temp.path.join("refused.pptx");
        let result = cli(&[
            "slide",
            "add",
            input.to_str().unwrap(),
            "--layout",
            layout,
            "--output",
            output.to_str().unwrap(),
        ]);
        assert!(!result.status.success());
        assert!(!output.exists());
    }
    assert_eq!(fs::read(&input).unwrap(), before);
}

#[test]
fn slide_duplicate_preserves_native_slide_and_notes_ownership() {
    let temp = TempWorkspace::new("fx204-duplicate");
    let input = temp.path.join("input.pptx");
    let output = temp.path.join("output.pptx");
    write_deck(&input, &["first", "last"]);
    add_speaker_notes(&input, "speaker");
    let result = cli(&[
        "slide",
        "duplicate",
        input.to_str().unwrap(),
        "1",
        "--output",
        output.to_str().unwrap(),
    ]);
    assert!(
        result.status.success(),
        "{}",
        String::from_utf8_lossy(&result.stderr)
    );
    let deck = Presentation::open(&output).unwrap();
    assert_eq!(deck.len(), 3);
    assert_eq!(deck.slide(0).unwrap().text(), "first");
    assert_eq!(deck.slide(1).unwrap().text(), "first");
    assert_eq!(deck.slide(2).unwrap().text(), "last");
    assert_eq!(
        deck.slide(0).unwrap().notes_text().as_deref(),
        Some("speaker")
    );
    assert_eq!(
        deck.slide(1).unwrap().notes_text().as_deref(),
        Some("speaker")
    );
    let package = OpcPackage::open(&output).unwrap();
    let notes = package
        .parts
        .iter()
        .filter(|(name, _)| {
            name.starts_with("/ppt/notesSlides/notesSlide") && name.ends_with(".xml")
        })
        .collect::<Vec<_>>();
    assert_eq!(notes.len(), 2);
    assert!(notes.iter().all(|(_, bytes)| {
        bytes
            .windows(b"urn:rdocx:test".len())
            .any(|w| w == b"urn:rdocx:test")
    }));
}

#[test]
fn slide_remove_preserves_native_relationship_lifecycle() {
    let temp = TempWorkspace::new("fx204-remove");
    let input = temp.path.join("input.pptx");
    let output = temp.path.join("output.pptx");
    write_deck(&input, &["first", "last"]);
    add_speaker_notes(&input, "speaker");
    let result = cli(&[
        "slide",
        "remove",
        input.to_str().unwrap(),
        "1",
        "--output",
        output.to_str().unwrap(),
    ]);
    assert!(
        result.status.success(),
        "{}",
        String::from_utf8_lossy(&result.stderr)
    );
    let deck = Presentation::open(&output).unwrap();
    assert_eq!(deck.len(), 1);
    assert_eq!(deck.slide(0).unwrap().text(), "last");
    assert!(
        !OpcPackage::open(&output)
            .unwrap()
            .parts
            .contains_key("/ppt/notesSlides/notesSlide1.xml")
    );
    let rejected = temp.path.join("rejected.pptx");
    let result = cli(&[
        "slide",
        "remove",
        input.to_str().unwrap(),
        "0",
        "--output",
        rejected.to_str().unwrap(),
    ]);
    assert!(!result.status.success());
    assert!(!rejected.exists());
}

#[test]
fn slide_move_uses_one_based_cli_positions() {
    let temp = TempWorkspace::new("fx204-move");
    let input = temp.path.join("input.pptx");
    write_deck(&input, &["first", "middle", "last"]);
    for (from, to, texts) in [
        (1, 3, ["middle", "last", "first"]),
        (3, 1, ["last", "first", "middle"]),
        (2, 2, ["first", "middle", "last"]),
    ] {
        let output = temp.path.join(format!("move-{from}-{to}.pptx"));
        let result = cli(&[
            "slide",
            "move",
            input.to_str().unwrap(),
            &from.to_string(),
            "--to",
            &to.to_string(),
            "--output",
            output.to_str().unwrap(),
        ]);
        assert!(
            result.status.success(),
            "{}",
            String::from_utf8_lossy(&result.stderr)
        );
        let deck = Presentation::open(&output).unwrap();
        for (i, text) in texts.iter().enumerate() {
            assert_eq!(deck.slide(i).unwrap().text(), *text);
        }
    }
    for (from, to) in [(0, 1), (1, 0), (4, 1), (1, 4)] {
        let output = temp.path.join("refused.pptx");
        let result = cli(&[
            "slide",
            "move",
            input.to_str().unwrap(),
            &from.to_string(),
            "--to",
            &to.to_string(),
            "--output",
            output.to_str().unwrap(),
        ]);
        assert!(!result.status.success());
        assert!(!output.exists());
    }
}

#[test]
fn slide_hide_and_show_round_trip() {
    let temp = TempWorkspace::new("fx204-hidden");
    let mut input = temp.path.join("input.pptx");
    write_deck(&input, &["text"]);
    for (i, action, hidden) in [
        (0, "hide", true),
        (1, "hide", true),
        (2, "show", false),
        (3, "show", false),
    ] {
        let output = temp.path.join(format!("hidden-{i}.pptx"));
        let result = cli(&[
            "slide",
            action,
            input.to_str().unwrap(),
            "1",
            "--output",
            output.to_str().unwrap(),
        ]);
        assert!(
            result.status.success(),
            "{}",
            String::from_utf8_lossy(&result.stderr)
        );
        let deck = Presentation::open(&output).unwrap();
        assert_eq!(deck.slide(0).unwrap().hidden(), hidden);
        assert_eq!(deck.slide(0).unwrap().text(), "text");
        input = output;
    }
}

#[test]
fn notes_set_text_or_file_preserves_exact_text() {
    let temp = TempWorkspace::new("fx204-notes");
    let input = temp.path.join("input.pptx");
    write_deck(&input, &["text"]);
    add_speaker_notes(&input, "old");
    let source = temp.path.join("notes.txt");
    for (case, text) in [
        "  leading\r\nnext\n\n",
        "first\nsecond\n",
        "first\rsecond",
        "  π 日本語\n尾  ",
        "",
    ]
    .iter()
    .enumerate()
    {
        fs::write(&source, text).unwrap();
        for (mode, args) in [
            (0, vec!["--text", *text]),
            (1, vec!["--from-file", source.to_str().unwrap()]),
        ] {
            let output = temp.path.join(format!("notes-{case}-{mode}.pptx"));
            let mut args2 = vec!["notes", "set", input.to_str().unwrap(), "1"];
            args2.extend(args);
            args2.extend(["--output", output.to_str().unwrap()]);
            let result = cli(&args2);
            assert!(
                result.status.success(),
                "{}",
                String::from_utf8_lossy(&result.stderr)
            );
            assert_eq!(
                Presentation::open(&output)
                    .unwrap()
                    .slide(0)
                    .unwrap()
                    .notes_text()
                    .as_deref(),
                Some(*text)
            );
            let package = OpcPackage::open(&output).unwrap();
            let xml = String::from_utf8(package.parts["/ppt/notesSlides/notesSlide1.xml"].to_vec())
                .unwrap();
            assert!(xml.contains("urn:rdocx:test"));
            assert!(xml.contains("b=\"1\""));
        }
    }
    let rejected = temp.path.join("refused.pptx");
    let result = cli(&[
        "notes",
        "set",
        input.to_str().unwrap(),
        "1",
        "--text",
        "x",
        "--from-file",
        source.to_str().unwrap(),
        "--output",
        rejected.to_str().unwrap(),
    ]);
    assert!(!result.status.success());
    assert!(!rejected.exists());
    fs::write(&source, [0xff]).unwrap();
    let result = cli(&[
        "notes",
        "set",
        input.to_str().unwrap(),
        "1",
        "--from-file",
        source.to_str().unwrap(),
        "--output",
        rejected.to_str().unwrap(),
    ]);
    assert!(!result.status.success());
    assert!(!rejected.exists());
}

#[test]
fn powerpoint_meta_get_set_matches_python_core_properties() {
    let temp = TempWorkspace::new("fx204-meta");
    let input = temp.path.join("input.pptx");
    let output = temp.path.join("output.pptx");
    write_deck(&input, &["text"]);
    let mut original = Presentation::open(&input).unwrap();
    let mut core = original.core_properties().cloned().unwrap_or_default();
    core.language = Some("en-GB".into());
    core.identifier = Some("preserved".into());
    original.set_core_properties(core).unwrap();
    original.save(&input).unwrap();
    let result = cli(&[
        "meta",
        "set",
        input.to_str().unwrap(),
        "--title",
        "title",
        "--author",
        "author",
        "--subject",
        "subject",
        "--keywords",
        "keywords",
        "--description",
        "description",
        "--category",
        "category",
        "--json",
        "--output",
        output.to_str().unwrap(),
    ]);
    assert!(
        result.status.success(),
        "{}",
        String::from_utf8_lossy(&result.stderr)
    );
    let deck = Presentation::open(&output).unwrap();
    let core = deck.core_properties().unwrap();
    assert_eq!(core.title.as_deref(), Some("title"));
    assert_eq!(core.creator.as_deref(), Some("author"));
    assert_eq!(core.subject.as_deref(), Some("subject"));
    assert_eq!(core.keywords.as_deref(), Some("keywords"));
    assert_eq!(core.description.as_deref(), Some("description"));
    assert_eq!(core.category.as_deref(), Some("category"));
    assert_eq!(core.language.as_deref(), Some("en-GB"));
    assert_eq!(core.identifier.as_deref(), Some("preserved"));
    let result = cli(&["meta", "get", output.to_str().unwrap(), "--json"]);
    assert!(result.status.success());
    let report: serde_json::Value = serde_json::from_slice(&result.stdout).unwrap();
    assert_eq!(report["core"]["language"], "en-GB");
}

#[test]
fn fit_reports_every_supported_overflow_and_needed_scale() {
    let temp = TempWorkspace::new("fx204-fit");
    for (name, text, expected) in [
        ("fit", "short".to_owned(), 0),
        ("overflow", "long overflow text ".repeat(100), 1),
    ] {
        let input = temp.path.join(format!("{name}.pptx"));
        write_deck(&input, &[&text]);
        let before = fs::read(&input).unwrap();
        let native = Presentation::open(&input).unwrap().fit_report(&[]).unwrap();
        let result = cli(&["fit", input.to_str().unwrap(), "--json"]);
        assert_eq!(
            result.status.code(),
            Some(expected),
            "{}",
            String::from_utf8_lossy(&result.stderr)
        );
        let report: serde_json::Value = serde_json::from_slice(&result.stdout).unwrap();
        assert_eq!(report["schema"], 1);
        assert_eq!(
            report["frames"].as_array().unwrap().len(),
            native.frames.len()
        );
        assert!(!native.frames.is_empty());
        for (json, frame) in report["frames"]
            .as_array()
            .unwrap()
            .iter()
            .zip(&native.frames)
        {
            assert_eq!(json["owner_part"], frame.update.location.owner_part);
            assert_eq!(json["overflow"], frame.overflow);
            assert_eq!(json["needed_scale"], serde_json::json!(frame.needed_scale));
            assert_eq!(
                json["needed_reduction"],
                serde_json::json!(frame.needed_reduction)
            );
            assert_eq!(json["status"], frame.update.status);
            assert_eq!(
                json["quantized_trial"].is_null(),
                frame.quantized_trial.is_none()
            );
            assert_eq!(
                json["persisted_trial"].is_null(),
                frame.persisted_trial.is_none()
            );
        }
        assert_eq!(fs::read(&input).unwrap(), before);
    }
}

#[test]
fn every_new_command_has_help_and_reference_contract() {
    for command in [
        vec!["replace"],
        vec!["slide", "add"],
        vec!["slide", "duplicate"],
        vec!["slide", "remove"],
        vec!["slide", "move"],
        vec!["slide", "hide"],
        vec!["slide", "show"],
        vec!["notes", "set"],
        vec!["fit"],
        vec!["meta", "get"],
        vec!["meta", "set"],
    ] {
        let mut args = command.clone();
        args.push("--help");
        let result = cli(&args);
        assert!(
            result.status.success(),
            "{:?}: {}",
            command,
            String::from_utf8_lossy(&result.stderr)
        );
        assert!(!result.stdout.is_empty());
    }
}

#[test]
fn new_mutations_reject_input_aliases_and_output_collisions() {
    let temp = TempWorkspace::new("fx204-identity");
    let input = temp.path.join("input.pptx");
    write_deck(&input, &["text"]);
    let original = fs::read(&input).unwrap();
    let existing = temp.path.join("existing.pptx");
    fs::write(&existing, b"unrelated").unwrap();
    let commands = vec![
        vec![
            "replace",
            input.to_str().unwrap(),
            "-p",
            "text",
            "-v",
            "new",
        ],
        vec!["slide", "add", input.to_str().unwrap(), "--layout", "7"],
        vec!["slide", "duplicate", input.to_str().unwrap(), "1"],
        vec!["slide", "remove", input.to_str().unwrap(), "1"],
        vec!["slide", "move", input.to_str().unwrap(), "1", "--to", "1"],
        vec!["slide", "hide", input.to_str().unwrap(), "1"],
        vec!["slide", "show", input.to_str().unwrap(), "1"],
        vec![
            "notes",
            "set",
            input.to_str().unwrap(),
            "1",
            "--text",
            "notes",
        ],
        vec!["meta", "set", input.to_str().unwrap(), "--title", "title"],
    ];
    let hardlink = temp.path.join("hardlink.pptx");
    fs::hard_link(&input, &hardlink).unwrap();
    let mut aliases = vec![input.clone(), hardlink];
    #[cfg(unix)]
    {
        let symlink = temp.path.join("symlink.pptx");
        std::os::unix::fs::symlink(&input, &symlink).unwrap();
        aliases.push(symlink);
    }
    aliases.push(existing.clone());
    for command in commands {
        for alias in &aliases {
            let mut args = command.clone();
            args.extend(["--output", alias.to_str().unwrap()]);
            let result = cli(&args);
            assert!(!result.status.success(), "{:?}", args);
            assert_eq!(fs::read(&input).unwrap(), original);
            assert_eq!(fs::read(&existing).unwrap(), b"unrelated");
        }
    }
}

#[test]
fn fx220_cli_fonts_embed_all_faces_list_actual_parts_and_remove() {
    let workspace = TempWorkspace::new("fx220-faces");
    let source = workspace.path.join("source.pptx");
    fs::write(&source, include_bytes!("../../rpptx/assets/default.pptx")).unwrap();
    let original = fs::read(&source).unwrap();
    let font = workspace.path.join("font.ttf");
    let data = include_bytes!("../../oxml-layout/fonts/Gelasio-Regular.ttf");
    fs::write(&font, data).unwrap();
    let mut current = source.clone();
    for (index, kind) in ["regular", "bold", "italic", "bold-italic"]
        .iter()
        .enumerate()
    {
        let output = workspace.path.join(format!("face{index}.pptx"));
        let result = cli(&[
            "font",
            "embed",
            current.to_str().unwrap(),
            "--name",
            "FX220 CLI",
            "--font",
            font.to_str().unwrap(),
            "--kind",
            kind,
            "--license-authorized",
            "--license-identity",
            "OFL-1.1",
            "--output",
            output.to_str().unwrap(),
            "--json",
        ]);
        assert!(
            result.status.success(),
            "{}",
            String::from_utf8_lossy(&result.stderr)
        );
        let record: serde_json::Value = serde_json::from_slice(&result.stdout).unwrap();
        assert_eq!(record["schema"], 1);
        current = output;
    }
    let result = cli(&["font", "list", current.to_str().unwrap(), "--json"]);
    assert!(
        result.status.success(),
        "{}",
        String::from_utf8_lossy(&result.stderr)
    );
    let record: serde_json::Value = serde_json::from_slice(&result.stdout).unwrap();
    assert_eq!(record["schema"], 1);
    let entry = record["fonts"]
        .as_array()
        .unwrap()
        .iter()
        .find(|font| font["name"] == "FX220 CLI")
        .unwrap();
    let faces = entry["faces"].as_array().unwrap();
    assert_eq!(faces.len(), 4);
    let package = oxml_opc::OpcPackage::open(&current).unwrap();
    for (face, kind) in faces
        .iter()
        .zip(["regular", "bold", "italic", "bold-italic"])
    {
        assert_eq!(face["kind"], kind);
        assert_eq!(face["format"], "eot");
        let body = package
            .get_part(face["part_name"].as_str().unwrap())
            .unwrap();
        assert_eq!(face["stored_len"].as_u64().unwrap() as usize, body.len());
        assert_eq!(face["decoded_len"].as_u64().unwrap() as usize, data.len());
    }
    let removed = workspace.path.join("removed.pptx");
    let result = cli(&[
        "font",
        "remove",
        current.to_str().unwrap(),
        "--name",
        "FX220 CLI",
        "--kind",
        "bold",
        "--output",
        removed.to_str().unwrap(),
        "--json",
    ]);
    assert!(
        result.status.success(),
        "{}",
        String::from_utf8_lossy(&result.stderr)
    );
    let record: serde_json::Value = serde_json::from_slice(&result.stdout).unwrap();
    assert_eq!(record["removed"], true);
    let result = cli(&["font", "list", removed.to_str().unwrap(), "--json"]);
    let record: serde_json::Value = serde_json::from_slice(&result.stdout).unwrap();
    let entry = record["fonts"]
        .as_array()
        .unwrap()
        .iter()
        .find(|font| font["name"] == "FX220 CLI")
        .unwrap();
    assert_eq!(entry["faces"].as_array().unwrap().len(), 3);
    let all_removed = workspace.path.join("all-removed.pptx");
    let result = cli(&[
        "font",
        "remove",
        removed.to_str().unwrap(),
        "--name",
        "FX220 CLI",
        "--output",
        all_removed.to_str().unwrap(),
        "--json",
    ]);
    assert!(result.status.success());
    let result = cli(&["font", "list", all_removed.to_str().unwrap(), "--json"]);
    let record: serde_json::Value = serde_json::from_slice(&result.stdout).unwrap();
    assert!(
        record["fonts"]
            .as_array()
            .unwrap()
            .iter()
            .all(|font| font["name"] != "FX220 CLI")
    );
    assert_eq!(fs::read(&source).unwrap(), original);
    assert_eq!(fs::read(&font).unwrap(), data);
}

#[test]
fn fx220_cli_font_settings_keep_absence_and_false_distinct() {
    let workspace = TempWorkspace::new("fx220-settings");
    let source = workspace.path.join("source.pptx");
    fs::write(&source, include_bytes!("../../rpptx/assets/default.pptx")).unwrap();
    let original = fs::read(&source).unwrap();
    let mut current = source.clone();
    for (index, setting) in ["true", "false", "none"].iter().enumerate() {
        let output = workspace.path.join(format!("settings{index}.pptx"));
        let result = cli(&[
            "font",
            "settings",
            current.to_str().unwrap(),
            "--embed-true-type-fonts",
            setting,
            "--save-subset-fonts",
            setting,
            "--output",
            output.to_str().unwrap(),
            "--json",
        ]);
        assert!(
            result.status.success(),
            "{}",
            String::from_utf8_lossy(&result.stderr)
        );
        let record: serde_json::Value = serde_json::from_slice(&result.stdout).unwrap();
        let expected = match *setting {
            "true" => json!(true),
            "false" => json!(false),
            _ => serde_json::Value::Null,
        };
        assert_eq!(record["settings"]["embed_true_type_fonts"], expected);
        assert_eq!(record["settings"]["save_subset_fonts"], expected);
        let reread = cli(&["font", "settings", output.to_str().unwrap(), "--json"]);
        assert!(reread.status.success());
        let reread: serde_json::Value = serde_json::from_slice(&reread.stdout).unwrap();
        assert_eq!(reread["settings"], record["settings"]);
        current = output;
    }
    for args in [
        vec!["--embed-true-type-fonts", "false"],
        vec!["--output", current.to_str().unwrap()],
    ] {
        let result = cli(&[&["font", "settings", source.to_str().unwrap()][..], &args].concat());
        assert!(!result.status.success());
    }
    assert_eq!(fs::read(&source).unwrap(), original);
}

#[test]
fn fx220_cli_font_refusals_preserve_inputs_and_existing_outputs() {
    let workspace = TempWorkspace::new("fx220-refusals");
    let source = workspace.path.join("source.pptx");
    fs::write(&source, include_bytes!("../../rpptx/assets/default.pptx")).unwrap();
    let original = fs::read(&source).unwrap();
    let font = workspace.path.join("font.ttf");
    let data = include_bytes!("../../oxml-layout/fonts/Gelasio-Regular.ttf");
    fs::write(&font, data).unwrap();
    let output = workspace.path.join("output.pptx");
    let common = [
        "font",
        "embed",
        source.to_str().unwrap(),
        "--name",
        "FX220 CLI",
        "--font",
        font.to_str().unwrap(),
    ];
    for extra in [
        vec!["--license-identity", "OFL-1.1"],
        vec!["--license-authorized", "--license-identity", " "],
        vec![
            "--license-authorized",
            "--license-identity",
            "OFL-1.1",
            "--kind",
            "invalid",
        ],
    ] {
        let result = cli(&[&common[..], &extra, &["--output", output.to_str().unwrap()]].concat());
        assert!(!result.status.success());
        assert!(!output.exists());
    }
    let license = ["--license-authorized", "--license-identity", "OFL-1.1"];
    for collision in [&source, &font, &output] {
        if collision == &output {
            fs::write(collision, b"keep existing output").unwrap();
        }
        let before = fs::read(collision).unwrap();
        let result = cli(&[
            &common[..],
            &license,
            &["--output", collision.to_str().unwrap()],
        ]
        .concat());
        assert!(!result.status.success());
        assert_eq!(fs::read(collision).unwrap(), before);
    }
    let invalid = workspace.path.join("invalid.ttf");
    fs::write(&invalid, b"not a font").unwrap();
    let oversized = workspace.path.join("oversized.ttf");
    fs::File::create(&oversized)
        .unwrap()
        .set_len(64 * 1024 * 1024 + 1)
        .unwrap();
    let absent_output = workspace.path.join("absent.pptx");
    for bad_font in [&invalid, &oversized, &workspace.path] {
        let result = cli(&[
            "font",
            "embed",
            source.to_str().unwrap(),
            "--name",
            "FX220 CLI",
            "--font",
            bad_font.to_str().unwrap(),
            "--license-authorized",
            "--license-identity",
            "OFL-1.1",
            "--output",
            absent_output.to_str().unwrap(),
        ]);
        assert!(!result.status.success());
        assert!(!absent_output.exists());
    }
    #[cfg(unix)]
    {
        let alias = workspace.path.join("source-alias.pptx");
        std::os::unix::fs::symlink(&source, &alias).unwrap();
        let hard_alias = workspace.path.join("font-hard-alias.pptx");
        fs::hard_link(&font, &hard_alias).unwrap();
        for collision in [&alias, &hard_alias] {
            let result = cli(&[
                &common[..],
                &license,
                &["--output", collision.to_str().unwrap()],
            ]
            .concat());
            assert!(!result.status.success());
        }
    }
    assert_eq!(fs::read(&source).unwrap(), original);
    assert_eq!(fs::read(&font).unwrap(), data);
    assert!(fs::read_dir(&workspace.path).unwrap().all(|entry| {
        !entry
            .unwrap()
            .file_name()
            .to_string_lossy()
            .ends_with(".tmp")
    }));
    let help = cli(&["font", "embed", "--help"]);
    assert!(help.status.success());
    let help = String::from_utf8(help.stdout).unwrap();
    for flag in [
        "--license-authorized",
        "--license-identity",
        "--font",
        "--kind",
        "--output",
    ] {
        assert!(help.contains(flag));
    }
}

fn fx218_cli_png() -> Vec<u8> {
    vec![
        0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44,
        0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00, 0x90,
        0x77, 0x53, 0xde, 0x00, 0x00, 0x00, 0x0c, 0x49, 0x44, 0x41, 0x54, 0x78, 0xda, 0x63, 0xf8,
        0xcf, 0xc0, 0x00, 0x00, 0x03, 0x01, 0x01, 0x00, 0xf7, 0x03, 0x41, 0x43, 0x00, 0x00, 0x00,
        0x00, 0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
    ]
}

#[test]
fn fx218_installed_and_cli_semantic_exports() {
    let workspace = TempWorkspace::new("fx218-semantic");
    let input = workspace.path.join("source.pptx");
    write_deck(&input, &["Visible & literal <text>", "Hidden content"]);
    let mut deck = Presentation::open(&input).unwrap();
    deck.set_notes_text(0, "Private speaker note").unwrap();
    deck.slide_mut(1).unwrap().set_hidden(true);
    let png = fx218_cli_png();
    deck.add_picture(0, &png, "pixel.png", Emu(1), Emu(1), None, None)
        .unwrap();
    deck.save(&input).unwrap();
    let source = fs::read(&input).unwrap();
    let input_arg = input.to_str().unwrap();
    for format in ["html", "md"] {
        let output = workspace.path.join(format!("result.{format}"));
        let output_arg = output.to_str().unwrap();
        let args = ["convert", input_arg, "--to", format, "--output", output_arg];
        for rendering in [
            vec!["--dpi", "150"],
            vec!["--quality", "90"],
            vec!["--slides", "1"],
            vec!["--font-dir", "missing"],
            vec!["--transparent"],
        ] {
            let refusal = cli(&[args.as_slice(), rendering.as_slice()].concat());
            assert!(!refusal.status.success());
            assert!(
                String::from_utf8_lossy(&refusal.stderr)
                    .contains("do not accept rendering options")
            );
            assert!(!output.exists());
        }
        let success = cli(&[args.as_slice(), &["--notes", "--skip-hidden"]].concat());
        assert!(
            success.status.success(),
            "{}",
            String::from_utf8_lossy(&success.stderr)
        );
        let text = fs::read_to_string(&output).unwrap().replace("&#32;", " ");
        assert!(
            text.contains(if format == "html" {
                "Visible &amp; literal &lt;text&gt;"
            } else {
                r"Visible \& literal \<text\>"
            }),
            "{format}: {text}"
        );
        assert!(text.contains("Private speaker note"));
        assert!(!text.contains("Hidden content"));
        assert!(text.contains("data:image/png;base64,"));
        fs::remove_file(&output).unwrap();
        let options = rpptx::PresentationTextExportOptions {
            inline_images: false,
            ..Default::default()
        };
        let detached = deck.to_html_with_options(&options).unwrap();
        let asset = workspace.path.join(&detached.assets[0].0);
        fs::write(&asset, b"keep asset").unwrap();
        let failure = cli(&[args.as_slice(), &["--embed-images=false"]].concat());
        assert!(!failure.status.success());
        assert!(!output.exists());
        assert_eq!(fs::read(&asset).unwrap(), b"keep asset");
        fs::remove_file(&asset).unwrap();
        let success = cli(&[args.as_slice(), &["--embed-images=false"]].concat());
        assert!(
            success.status.success(),
            "{}",
            String::from_utf8_lossy(&success.stderr)
        );
        assert_eq!(fs::read(&asset).unwrap(), png);
        assert!(
            fs::read_to_string(&output)
                .unwrap()
                .contains(&detached.assets[0].0)
        );
        fs::remove_file(&asset).unwrap();
        fs::remove_file(&output).unwrap();
        let refusal = cli(&[
            "convert", input_arg, "--to", format, "--output", input_arg, "--force",
        ]);
        assert!(!refusal.status.success());
        assert_eq!(fs::read(&input).unwrap(), source);
    }
    let refusal = cli(&["convert", input_arg, "--to", "pdf", "--notes"]);
    assert!(!refusal.status.success());
    assert!(String::from_utf8_lossy(&refusal.stderr).contains("require --to html or --to md"));
    assert_eq!(fs::read(&input).unwrap(), source);
}

#[test]
fn fx216_import_output_extension_selects_package_class_without_touching_inputs() {
    let temp = TempWorkspace::new("import-class");
    let destination = temp.path.join("destination.potx");
    let source = temp.path.join("source.pptx");
    let output = temp.path.join("output.pptx");
    write_deck(&destination, &["destination"]);
    write_deck(&source, &["source"]);
    let before_destination = fs::read(&destination).unwrap();
    let before_source = fs::read(&source).unwrap();
    assert_eq!(
        Presentation::open(&destination)
            .unwrap()
            .package_class()
            .unwrap(),
        PresentationPackageClass::Template
    );
    let result = cli(&[
        "slide",
        "import",
        destination.to_str().unwrap(),
        source.to_str().unwrap(),
        "--slides",
        "1",
        "--output",
        output.to_str().unwrap(),
    ]);
    assert!(
        result.status.success(),
        "{}",
        String::from_utf8_lossy(&result.stderr)
    );
    let reopened = Presentation::open(&output).unwrap();
    assert_eq!(reopened.len(), 2);
    assert_eq!(
        reopened.package_class().unwrap(),
        PresentationPackageClass::Presentation
    );
    assert_eq!(fs::read(&destination).unwrap(), before_destination);
    assert_eq!(fs::read(&source).unwrap(), before_source);
}

#[test]
fn fx216_import_refuses_macro_free_vba_output_before_publication() {
    let temp = TempWorkspace::new("import-vba");
    let fixture = temp.path.join("fixture.pptx");
    let destination = temp.path.join("destination.pptm");
    let source = temp.path.join("source.pptx");
    let output = temp.path.join("output.pptx");
    write_deck(&fixture, &["destination"]);
    write_deck(&source, &["source"]);
    let mut package = OpcPackage::open(&fixture).unwrap();
    fs::remove_file(&fixture).unwrap();
    package
        .get_or_create_part_rels("/ppt/presentation.xml")
        .add(rel_types::VBA_PROJECT, "vbaProject.bin");
    package.set_part("/ppt/vbaProject.bin", b"vba-project".to_vec());
    package.content_types.add_override(
        "/ppt/vbaProject.bin",
        "application/vnd.ms-office.vbaProject",
    );
    package.content_types.add_override(
        "/ppt/presentation.xml",
        content_types::PRESENTATION_MACRO_ENABLED,
    );
    package.save(&destination).unwrap();
    let before_destination = fs::read(&destination).unwrap();
    let before_source = fs::read(&source).unwrap();
    let result = cli(&[
        "slide",
        "import",
        destination.to_str().unwrap(),
        source.to_str().unwrap(),
        "--slides",
        "1",
        "--output",
        output.to_str().unwrap(),
    ]);
    assert_eq!(result.status.code(), Some(1));
    assert!(
        String::from_utf8_lossy(&result.stderr).contains("VBA project"),
        "{}",
        String::from_utf8_lossy(&result.stderr)
    );
    assert!(!output.exists());
    assert_eq!(fs::read(&destination).unwrap(), before_destination);
    assert_eq!(fs::read(&source).unwrap(), before_source);
    assert_eq!(fs::read_dir(&temp.path).unwrap().count(), 2);
}

#[test]
fn fx216_cli_ranges_insertion_and_atomic_errors_preserve_inputs() {
    let temp = TempWorkspace::new("import-range");
    let destination = temp.path.join("destination.pptx");
    let source = temp.path.join("source.pptx");
    let output = temp.path.join("merged.pptx");
    write_deck(&destination, &["destination one", "destination two"]);
    write_deck(
        &source,
        &[
            "source one",
            "source two",
            "source three",
            "source four",
            "source five",
        ],
    );
    let inputs = (fs::read(&destination).unwrap(), fs::read(&source).unwrap());
    let result = cli(&[
        "slide",
        "import",
        destination.to_str().unwrap(),
        source.to_str().unwrap(),
        "--slides",
        "1,3-5",
        "--at",
        "2",
        "--keep-source-formatting",
        "--output",
        output.to_str().unwrap(),
        "--json",
    ]);
    assert!(
        result.status.success(),
        "{}",
        String::from_utf8_lossy(&result.stderr)
    );
    let merged = Presentation::open(&output).unwrap();
    let texts: Vec<_> = (0..merged.len())
        .map(|index| merged.slide(index).unwrap().text())
        .collect();
    assert_eq!(
        texts,
        [
            "destination one",
            "source one",
            "source three",
            "source four",
            "source five",
            "destination two"
        ]
    );
    assert!(merged.validate().is_empty());
    for selection in ["0", "3-2", "1,1", "6", "18446744073709551616"] {
        let invalid = temp.path.join("invalid.pptx");
        let result = cli(&[
            "slide",
            "import",
            destination.to_str().unwrap(),
            source.to_str().unwrap(),
            "--slides",
            selection,
            "--output",
            invalid.to_str().unwrap(),
        ]);
        assert!(
            !result.status.success(),
            "selection {selection} must refuse"
        );
        assert!(!invalid.exists());
        assert_eq!(
            (fs::read(&destination).unwrap(), fs::read(&source).unwrap()),
            inputs
        );
    }
    for alias in [&destination, &source] {
        let result = cli(&[
            "slide",
            "import",
            destination.to_str().unwrap(),
            source.to_str().unwrap(),
            "--output",
            alias.to_str().unwrap(),
        ]);
        assert!(!result.status.success());
        assert_eq!(
            (fs::read(&destination).unwrap(), fs::read(&source).unwrap()),
            inputs
        );
    }
}

#[test]
fn fx240_unsigned_status_material_bound_and_output_refusal() {
    let workspace = TempWorkspace::new("fx240-signing");
    let input = workspace.path.join("unsigned.pptx");
    Presentation::new().unwrap().save(&input).unwrap();
    let result = cli(&["verify-signatures", input.to_str().unwrap(), "--json"]);
    assert_eq!(result.status.code(), Some(1));
    let json: serde_json::Value = serde_json::from_slice(&result.stdout).unwrap();
    assert_eq!(json["schema"], 1);
    assert_eq!(json["status"], "unsigned");
    assert_eq!(json["signatures"], serde_json::json!([]));
    let before = fs::read(&input).unwrap();
    let key = workspace.path.join("oversized.key");
    let cert = workspace.path.join("invalid.cert");
    fs::write(&key, vec![b'x'; 1024 * 1024 + 1]).unwrap();
    fs::write(&cert, b"invalid").unwrap();
    let output = workspace.path.join("refused.pptx");
    let result = cli(&[
        "sign",
        input.to_str().unwrap(),
        "--key",
        key.to_str().unwrap(),
        "--cert",
        cert.to_str().unwrap(),
        "-o",
        output.to_str().unwrap(),
        "--json",
    ]);
    assert_eq!(result.status.code(), Some(1));
    assert!(result.stdout.is_empty());
    assert!(!output.exists());
    assert_eq!(fs::read(&input).unwrap(), before);
    fs::write(&output, b"keep").unwrap();
    let result = cli(&[
        "sign",
        input.to_str().unwrap(),
        "--key",
        key.to_str().unwrap(),
        "--cert",
        cert.to_str().unwrap(),
        "-o",
        output.to_str().unwrap(),
    ]);
    assert_eq!(result.status.code(), Some(1));
    assert_eq!(fs::read(&output).unwrap(), b"keep");
    assert_eq!(
        cli(&["sign", input.to_str().unwrap()]).status.code(),
        Some(2)
    );
    assert!(!workspace.path.read_dir().unwrap().any(|entry| {
        entry
            .unwrap()
            .file_name()
            .to_string_lossy()
            .contains(".tmp")
    }));
}

#[test]
fn fx240_text_and_json_io_failures_exit_one_without_panicking() {
    let workspace = TempWorkspace::new("fx240-stdout");
    let input = workspace.path.join("unsigned.pptx");
    Presentation::new().unwrap().save(&input).unwrap();
    let before = fs::read(&input).unwrap();
    let material = |name: &str, label: &str| {
        let source = include_str!("../../oxml-opc/src/signature.rs");
        let prefix = format!("const {name}: &str = \"");
        let encoded = source
            .split_once(prefix.as_str())
            .unwrap()
            .1
            .split_once('"')
            .unwrap()
            .0;
        format!("-----BEGIN {label}-----\n{encoded}\n-----END {label}-----\n")
    };
    let key = workspace.path.join("key.pem");
    let cert = workspace.path.join("cert.pem");
    fs::write(
        &key,
        material("TEST_PRIVATE_KEY_PKCS8_BASE64", "PRIVATE KEY"),
    )
    .unwrap();
    fs::write(
        &cert,
        material("TEST_CERTIFICATE_DER_BASE64", "CERTIFICATE"),
    )
    .unwrap();
    for signing in [false, true] {
        for as_json in [false, true] {
            let output = workspace.path.join(format!("signed-{as_json}.pptx"));
            let mut command = Command::new(env!("CARGO_BIN_EXE_rpptx"));
            if signing {
                command
                    .arg("sign")
                    .arg(&input)
                    .arg("--key")
                    .arg(&key)
                    .arg("--cert")
                    .arg(&cert)
                    .arg("-o")
                    .arg(&output);
            } else {
                command.arg("verify-signatures").arg(&input);
            }
            if as_json {
                command.arg("--json");
            }
            let mut child = command
                .stdout(Stdio::piped())
                .stderr(Stdio::piped())
                .spawn()
                .unwrap();
            drop(child.stdout.take());
            let result = child.wait_with_output().unwrap();
            assert_eq!(
                result.status.code(),
                Some(1),
                "signing={signing}, json={as_json}: {}",
                String::from_utf8_lossy(&result.stderr)
            );
            assert!(String::from_utf8_lossy(&result.stderr).starts_with("Error:"));
            assert!(!String::from_utf8_lossy(&result.stderr).contains("panicked"));
            assert_eq!(fs::read(&input).unwrap(), before);
            if signing {
                // Publication precedes reporting. A report I/O error keeps the valid new package.
                let reports = Presentation::open(&output)
                    .unwrap()
                    .verify_signatures()
                    .unwrap();
                assert!(!reports.is_empty());
                assert!(
                    reports
                        .iter()
                        .all(|report| report.cryptographically_valid && report.coverage_complete)
                );
            } else {
                assert!(!output.exists());
            }
        }
    }
}

#[test]
fn fx253_pdf_png_convert_render_hide_native_master_logo() {
    let temp = TempWorkspace::new("master-visibility");
    let input = temp.path.join("visibility.pptx");
    write_deck(&input, &["direct slide content"]);
    let mut deck = Presentation::open(&input).unwrap();
    let plain = deck.to_bytes().unwrap();
    let master = String::from_utf8(deck.master_xml(0).unwrap()).unwrap();
    let logo = r#"<p:sp><p:nvSpPr><p:cNvPr id="999" name="Native visibility logo"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="E02020"/></a:solidFill></p:spPr></p:sp>"#;
    deck.replace_master_xml(
        0,
        &master.replace("</p:spTree>", &format!("{logo}</p:spTree>")),
    )
    .unwrap();
    let visible = deck.to_bytes().unwrap();
    deck.set_show_master_shapes_at(rpptx::ShapeOwner::Slide(0), Some(false))
        .unwrap();
    let hidden = deck.to_bytes().unwrap();
    let mut results = Vec::new();
    for (name, bytes) in [("plain", plain), ("visible", visible), ("hidden", hidden)] {
        fs::write(&input, bytes).unwrap();
        let pdf = temp.path.join(format!("{name}.pdf"));
        let out = cli(&[
            "convert",
            input.to_str().unwrap(),
            "--to",
            "pdf",
            "--output",
            pdf.to_str().unwrap(),
        ]);
        assert!(out.status.success(), "{:?}", out);
        let directory = temp.path.join(name);
        let out = cli(&[
            "render",
            input.to_str().unwrap(),
            "--output",
            directory.to_str().unwrap(),
            "--dpi",
            "72",
        ]);
        assert!(out.status.success(), "{:?}", out);
        results.push((
            fs::read(pdf).unwrap(),
            fs::read(directory.join("visibility_slide1.png")).unwrap(),
        ));
    }
    assert_ne!(
        results[0].0, results[1].0,
        "positive master logo PDF must differ"
    );
    assert_ne!(
        results[0].1, results[1].1,
        "positive master logo PNG must differ"
    );
    assert_eq!(
        results[0], results[2],
        "hidden master output must equal the original direct-slide scene"
    );
}

#[test]
fn fx254_map_implicit_zero_is_pairwise_atomic() {
    let workspace = TempWorkspace::new("fx254-map");
    let input = workspace.path.join("input.pptx");
    let output = workspace.path.join("output.pptx");
    let map = workspace.path.join("map.json");
    let mut deck = Presentation::new().unwrap();
    deck.add_slide(6).unwrap();
    deck.slide_mut(0)
        .unwrap()
        .add_textbox(Emu(0), Emu(0), Emu(1000000), Emu(500000))
        .unwrap()
        .set_text("{{first}}")
        .unwrap();
    deck.save(&input).unwrap();
    let before = fs::read(&input).unwrap();
    fs::write(
        &map,
        r#"[{"placeholder":"{{first}}","value":"one"},{"placeholder":"absent","value":"two"}]"#,
    )
    .unwrap();
    let result = cli(&[
        "replace",
        input.to_str().unwrap(),
        "--map",
        map.to_str().unwrap(),
        "-o",
        output.to_str().unwrap(),
    ]);
    assert!(!result.status.success());
    assert!(!output.exists());
    assert_eq!(fs::read(&input).unwrap(), before);
    fs::write(&map, r#"[{"placeholder":"{{first}}","value":"one"},{"placeholder":"absent","value":"two","expect":0}]"#).unwrap();
    let result = cli(&[
        "replace",
        input.to_str().unwrap(),
        "--map",
        map.to_str().unwrap(),
        "-o",
        output.to_str().unwrap(),
    ]);
    assert!(
        result.status.success(),
        "{}",
        String::from_utf8_lossy(&result.stderr)
    );
    assert!(output.exists());
}

#[test]
fn fx254_missing_font_and_output_messages() {
    let workspace = TempWorkspace::new("fx254-font");
    let input = workspace.path.join("input.pptx");
    let output = workspace.path.join("output.pptx");
    Presentation::new().unwrap().save(&input).unwrap();
    let before = fs::read(&input).unwrap();
    for kind in [None, Some("regular")] {
        let mut args = vec![
            "font",
            "remove",
            input.to_str().unwrap(),
            "--name",
            "Missing font",
            "-o",
            output.to_str().unwrap(),
        ];
        if let Some(kind) = kind {
            args.extend(["--kind", kind]);
        }
        let result = cli(&args);
        assert!(!result.status.success());
        assert!(!output.exists());
        assert_eq!(fs::read(&input).unwrap(), before);
    }
    fs::write(&output, b"sentinel").unwrap();
    let result = cli(&[
        "font",
        "remove",
        input.to_str().unwrap(),
        "--name",
        "Missing font",
        "-o",
        output.to_str().unwrap(),
    ]);
    assert!(!result.status.success());
    assert!(!String::from_utf8_lossy(&result.stderr).contains("--force"));
    assert_eq!(fs::read(&output).unwrap(), b"sentinel");
}
