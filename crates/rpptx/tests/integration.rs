use std::collections::{BTreeMap, HashSet};
use std::fmt::Write as _;
use std::fs;
use std::io::Cursor;
use std::path::{Path, PathBuf};
use std::process::Command;
use std::sync::atomic::AtomicU64;

const F224_BROWSER_SSIM_FLOOR: f64 = 0.95;
const F225_PDF_SSIM_FLOOR: f64 = 0.995;

fn f224_browser_gate_accepts(
    structure_matches: bool,
    text_matches: bool,
    maximum_geometry_error_px: u32,
    ssim: f64,
) -> bool {
    structure_matches
        && text_matches
        && maximum_geometry_error_px <= 1
        && ssim >= F224_BROWSER_SSIM_FLOOR
}

fn f225_pdf_gate_accepts(
    geometry_matches: bool,
    text_matches: bool,
    link_matches: bool,
    point_error: f64,
    ssim: f64,
) -> bool {
    geometry_matches
        && text_matches
        && link_matches
        && point_error <= 1.0
        && ssim >= F225_PDF_SSIM_FLOOR
}

#[test]
#[cfg(feature = "default-template")]
fn html_public_import_saves_reopens_and_remains_editable() {
    let html = "<section data-slide><div style='position:absolute;left:10px;top:20px;width:240px;height:80px;background:#336699'><a href='https://example.com'><strong>editable HTML</strong></a></div><table style='position:absolute;left:10px;top:120px;width:240px;height:100px'><tr><td>one</td><td>two</td></tr></table></section>";
    let imported = rpptx::Presentation::from_html(html, &[]).expect("HTML imports");
    assert!(
        imported.diagnostics.is_empty(),
        "{:?}",
        imported.diagnostics
    );
    assert_eq!(imported.presentation.len(), 1);
    assert_eq!(imported.presentation.slide(0).unwrap().shapes().len(), 2);
    assert_eq!(
        imported.presentation.slide(0).unwrap().text(),
        "editable HTML\none\ttwo"
    );

    let bytes = imported.presentation.to_bytes().expect("HTML deck saves");
    let mut reopened = rpptx::Presentation::from_bytes(&bytes).expect("HTML deck reopens");
    reopened
        .slide_mut(0)
        .unwrap()
        .shape_mut(0)
        .unwrap()
        .set_text("edited after reopen")
        .expect("HTML shape remains editable");
    assert_eq!(
        reopened
            .slide(0)
            .unwrap()
            .shape(0)
            .unwrap()
            .text()
            .as_deref(),
        Some("edited after reopen")
    );
}

#[cfg(feature = "render")]
fn source_built_pdf_for_import() -> Vec<u8> {
    use lopdf::content::{Content, Operation};
    use lopdf::{Document, Object, Stream, dictionary};

    let mut document = Document::with_version("1.7");
    let pages_id = document.new_object_id();
    let carlito = oxml_layout::bundled_fonts::bundled_font_data()[0].1;
    let font_file_id = document.add_object(Stream::new(
        dictionary! { "Length1" => carlito.len() as i64 },
        carlito.to_vec(),
    ));
    let descriptor_id = document.add_object(dictionary! {
        "Type" => "FontDescriptor",
        "FontName" => "Carlito",
        "Flags" => 32,
        "FontBBox" => vec![(-500).into(), (-300).into(), 1500.into(), 1200.into()],
        "ItalicAngle" => 0,
        "Ascent" => 1000,
        "Descent" => -300,
        "CapHeight" => 700,
        "StemV" => 80,
        "FontFile2" => font_file_id,
    });
    let font_id = document.add_object(dictionary! {
        "Type" => "Font",
        "Subtype" => "TrueType",
        "BaseFont" => "Carlito",
        "FirstChar" => 0,
        "LastChar" => 255,
        "Widths" => (0..=255).map(|_| Object::Integer(500)).collect::<Vec<_>>(),
        "FontDescriptor" => descriptor_id,
        "Encoding" => "WinAnsiEncoding",
    });
    let image_id = document.add_object(Stream::new(
        dictionary! {
            "Type" => "XObject",
            "Subtype" => "Image",
            "Width" => 1,
            "Height" => 1,
            "ColorSpace" => "DeviceRGB",
            "BitsPerComponent" => 8,
        },
        vec![255, 0, 0],
    ));
    let content = Content {
        operations: vec![
            Operation::new("rg", vec![0.into(), 0.into(), 1.into()]),
            Operation::new("re", vec![0.into(), 0.into(), 720.into(), 720.into()]),
            Operation::new("f", vec![]),
            Operation::new("RG", vec![1.into(), 1.into(), 1.into()]),
            Operation::new("w", vec![Object::Real(1.92)]),
            Operation::new("J", vec![1.into()]),
            Operation::new("j", vec![2.into()]),
            Operation::new(
                "d",
                vec![
                    Object::Array(vec![Object::Real(4.8), Object::Real(1.92)]),
                    0.into(),
                ],
            ),
            Operation::new(
                "re",
                vec![96.into(), 96.into(), Object::Real(38.4), Object::Real(38.4)],
            ),
            Operation::new("S", vec![]),
            Operation::new("rg", vec![1.into(), 1.into(), 1.into()]),
            Operation::new("BT", vec![]),
            Operation::new("Tf", vec![Object::Name(b"F1".to_vec()), 12.into()]),
            Operation::new(
                "Tm",
                vec![
                    1.into(),
                    0.into(),
                    0.into(),
                    1.into(),
                    24.into(),
                    684.into(),
                ],
            ),
            Operation::new("Tj", vec![Object::string_literal("I")]),
            Operation::new("ET", vec![]),
            Operation::new("BT", vec![]),
            Operation::new("Tf", vec![Object::Name(b"F1".to_vec()), 12.into()]),
            Operation::new(
                "Tm",
                vec![
                    0.into(),
                    1.into(),
                    (-1).into(),
                    0.into(),
                    360.into(),
                    360.into(),
                ],
            ),
            Operation::new("Tj", vec![Object::string_literal("R")]),
            Operation::new("ET", vec![]),
            Operation::new("q", vec![]),
            Operation::new(
                "cm",
                vec![
                    (-24).into(),
                    0.into(),
                    0.into(),
                    (-12).into(),
                    696.into(),
                    36.into(),
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
        "Rect" => vec![12.into(), 12.into(), 72.into(), 42.into()],
        "Border" => vec![0.into(), 0.into(), 0.into()],
        "A" => action_id,
    });
    let page_id = document.add_object(dictionary! {
        "Type" => "Page",
        "Parent" => pages_id,
        "MediaBox" => vec![0.into(), 0.into(), 720.into(), 720.into()],
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
    let catalog_id = document.add_object(dictionary! { "Type" => "Catalog", "Pages" => pages_id });
    document.trailer.set("Root", catalog_id);
    let mut bytes = Vec::new();
    document.save_to(&mut bytes).unwrap();
    bytes
}

#[cfg(feature = "render")]
fn pdf_source_with_uri(source: &[u8], uri: &str) -> Vec<u8> {
    use lopdf::{Document, Object};

    let mut document = Document::load_mem(source).unwrap();
    for object in document.objects.values_mut() {
        let Object::Dictionary(dictionary) = object else {
            continue;
        };
        if dictionary.get(b"S").and_then(Object::as_name).ok() == Some(b"URI") {
            dictionary.set("URI", Object::string_literal(uri));
        }
    }
    let mut bytes = Vec::new();
    document.save_to(&mut bytes).unwrap();
    bytes
}

#[cfg(feature = "render")]
fn f225_actual_link_matches(presentation: &rpptx::Presentation, expected: &str) -> bool {
    let bytes = presentation.to_bytes().unwrap();
    let package = OpcPackage::from_reader(Cursor::new(bytes)).unwrap();
    package
        .get_part_rels("/ppt/slides/slide1.xml")
        .unwrap()
        .items
        .iter()
        .any(|relationship| {
            relationship.target == expected
                && relationship.target_mode.as_deref() == Some("External")
        })
}

#[cfg(feature = "render")]
fn f225_first_shape_x_points(presentation: &rpptx::Presentation) -> f64 {
    let bytes = presentation.to_bytes().unwrap();
    let package = OpcPackage::from_reader(Cursor::new(bytes)).unwrap();
    let xml =
        String::from_utf8(package.get_part("/ppt/slides/slide1.xml").unwrap().to_vec()).unwrap();
    let shape = xml.split("<p:sp>").nth(1).unwrap();
    let offset = shape.split("<a:off x=\"").nth(1).unwrap();
    offset.split('"').next().unwrap().parse::<f64>().unwrap() / 12_700.0
}

#[cfg(feature = "render")]
fn f225_first_image_x_points(presentation: &rpptx::Presentation) -> f64 {
    let bytes = presentation.to_bytes().unwrap();
    let package = OpcPackage::from_reader(Cursor::new(bytes)).unwrap();
    let xml =
        String::from_utf8(package.get_part("/ppt/slides/slide1.xml").unwrap().to_vec()).unwrap();
    let picture_start = xml.find("<p:pic>").unwrap();
    let group_start = xml[..picture_start].rfind("<p:grpSp>").unwrap();
    let group = &xml[group_start..picture_start];
    let offset = group.split("<a:off x=\"").nth(1).unwrap();
    offset.split('"').next().unwrap().parse::<f64>().unwrap() / 12_700.0
}

#[cfg(feature = "render")]
fn f225_render_png(presentation: &rpptx::Presentation) -> Vec<u8> {
    let (_, layout) = presentation.render_deterministic().unwrap();
    oxml_pdf::render_page_to_png(&layout, 0, 150.0).unwrap()
}

#[test]
#[cfg(feature = "render")]
fn both_pdf_import_modes_publish_valid_reopenable_presentations() {
    let source = source_built_pdf_for_import();
    for mode in [
        rpptx::PdfImportMode::Editable,
        rpptx::PdfImportMode::PreservedGraphic { dpi: 96.0 },
    ] {
        let imported = rpptx::Presentation::from_pdf_bytes(&source, mode).expect("PDF imports");
        assert_eq!(imported.presentation.len(), 1);
        assert_eq!(
            imported.presentation.slide_size(),
            Some((rpptx::Emu(9_144_000), rpptx::Emu(9_144_000)))
        );
        assert!(imported.presentation.validate().is_empty());
        let saved = imported.presentation.to_bytes().expect("PDF deck saves");
        let reopened = rpptx::Presentation::from_bytes(&saved).expect("PDF deck reopens");
        assert_eq!(reopened.len(), 1);
        if mode == rpptx::PdfImportMode::Editable {
            assert_eq!(reopened.slide(0).unwrap().text(), "I\nR");
            assert_eq!(reopened.slide(0).unwrap().shapes().len(), 6);
            let package = OpcPackage::from_reader(Cursor::new(&saved)).unwrap();
            let slide_xml =
                String::from_utf8(package.get_part("/ppt/slides/slide1.xml").unwrap().to_vec())
                    .unwrap();
            assert!(slide_xml.contains("<a:custGeom>"));
            assert!(slide_xml.contains(r#"<a:srgbClr val="0000FF""#));
            assert!(slide_xml.contains(r#"<a:ln w="24383" cap="rnd""#));
            assert!(slide_xml.contains("<a:custDash>"));
            assert!(slide_xml.contains("<a:bevel"));
            assert!(slide_xml.matches("rot=\"").count() >= 2, "{slide_xml}");
            assert!(slide_xml.contains("flipV=\"1\""), "{slide_xml}");
            assert!(slide_xml.contains(r#"<a:off x="152400" y="8610600""#));
            assert!(slide_xml.contains(r#"<a:ext cx="762000" cy="381000""#));
            let relationships = package.get_part_rels("/ppt/slides/slide1.xml").unwrap();
            assert!(relationships.items.iter().any(|relationship| {
                relationship.target == "https://example.com/pdf"
                    && relationship.target_mode.as_deref() == Some("External")
            }));
            let image_relationship = relationships
                .items
                .iter()
                .find(|relationship| relationship.rel_type == rel_types::IMAGE)
                .unwrap();
            let image_part = OpcPackage::resolve_rel_target(
                "/ppt/slides/slide1.xml",
                &image_relationship.target,
            );
            let image = tiny_skia::Pixmap::decode_png(package.get_part(&image_part).unwrap())
                .expect("editable imported image stays a valid PNG");
            assert_eq!((image.width(), image.height()), (1, 1));
        } else {
            assert_eq!(reopened.slide(0).unwrap().shapes().len(), 1);
        }
    }
}

#[test]
#[cfg(feature = "render")]
#[ignore = "requires Poppler 26.01.0"]
fn pdf_page_import_matches_pinned_poppler_geometry_pixels_text_and_links() {
    let version = Command::new("pdfinfo").arg("-v").output().unwrap();
    let version_text = format!(
        "{}{}",
        String::from_utf8_lossy(&version.stdout),
        String::from_utf8_lossy(&version.stderr)
    );
    assert!(version_text.contains("pdfinfo version 26.01.0"));
    let version = Command::new("pdftoppm").arg("-v").output().unwrap();
    let version_text = format!(
        "{}{}",
        String::from_utf8_lossy(&version.stdout),
        String::from_utf8_lossy(&version.stderr)
    );
    assert!(version_text.contains("pdftoppm version 26.01.0"));
    let root = std::env::temp_dir().join(format!("rpptx-f225-poppler-{}", std::process::id()));
    if root.exists() {
        fs::remove_dir_all(&root).unwrap();
    }
    fs::create_dir_all(&root).unwrap();
    let source_path = root.join("source.pdf");
    let oracle_prefix = root.join("oracle");
    let source = source_built_pdf_for_import();
    fs::write(&source_path, &source).unwrap();

    let info = Command::new("pdfinfo").arg(&source_path).output().unwrap();
    assert!(info.status.success());
    let info = String::from_utf8(info.stdout).unwrap();
    assert!(info.contains("Pages:           1"));
    assert!(info.contains("Page size:       720 x 720 pts"));
    let render = Command::new("pdftoppm")
        .args(["-png", "-singlefile", "-r", "150"])
        .arg(&source_path)
        .arg(&oracle_prefix)
        .output()
        .unwrap();
    assert!(
        render.status.success(),
        "pdftoppm failed: {}",
        String::from_utf8_lossy(&render.stderr)
    );
    let oracle_png = fs::read(root.join("oracle.png")).unwrap();

    let preserved = rpptx::Presentation::from_pdf_bytes(
        &source,
        rpptx::PdfImportMode::PreservedGraphic { dpi: 150.0 },
    )
    .unwrap();
    let preserved_bytes = preserved.presentation.to_bytes().unwrap();
    let preserved_package = OpcPackage::from_reader(Cursor::new(&preserved_bytes)).unwrap();
    let image_relationship = preserved_package
        .get_part_rels("/ppt/slides/slide1.xml")
        .unwrap()
        .items
        .iter()
        .find(|relationship| relationship.rel_type == rel_types::IMAGE)
        .unwrap();
    let image_part =
        OpcPackage::resolve_rel_target("/ppt/slides/slide1.xml", &image_relationship.target);
    let rust_png = preserved_package.get_part(&image_part).unwrap().to_vec();
    let rust = tiny_skia::Pixmap::decode_png(&rust_png).unwrap();
    let oracle = tiny_skia::Pixmap::decode_png(&oracle_png).unwrap();
    assert_eq!(
        (rust.width(), rust.height()),
        (oracle.width(), oracle.height())
    );
    let ssim = smartart_png_ssim(&rust_png, &oracle_png);
    eprintln!("F-225 Poppler raw full-image luminance SSIM: {ssim:.9}");

    let editable =
        rpptx::Presentation::from_pdf_bytes(&source, rpptx::PdfImportMode::Editable).unwrap();
    let (_, editable_layout) = editable.presentation.render_deterministic().unwrap();
    let editable_png = oxml_pdf::render_page_to_png(&editable_layout, 0, 150.0).unwrap();
    let editable_ssim = smartart_png_ssim(&editable_png, &oracle_png);
    eprintln!("F-225 Poppler editable raw luminance SSIM: {editable_ssim:.9}");
    let saved = editable.presentation.to_bytes().unwrap();
    let package = OpcPackage::from_reader(Cursor::new(&saved)).unwrap();
    let link_matches = package
        .get_part_rels("/ppt/slides/slide1.xml")
        .unwrap()
        .items
        .iter()
        .any(|relationship| {
            relationship.target == "https://example.com/pdf"
                && relationship.target_mode.as_deref() == Some("External")
        });
    let slide_xml =
        String::from_utf8(package.get_part("/ppt/slides/slide1.xml").unwrap().to_vec()).unwrap();
    let geometry_matches = editable.presentation.slide_size()
        == Some((rpptx::Emu(9_144_000), rpptx::Emu(9_144_000)))
        && slide_xml.contains(r#"<a:off x="152400" y="8610600""#)
        && slide_xml.contains(r#"<a:ext cx="762000" cy="381000""#);
    assert!(f225_pdf_gate_accepts(
        geometry_matches,
        editable.presentation.slide(0).unwrap().text() == "I\nR",
        link_matches,
        0.0,
        ssim,
    ));
    assert!(f225_pdf_gate_accepts(
        geometry_matches,
        editable.presentation.slide(0).unwrap().text() == "I\nR",
        link_matches,
        0.0,
        editable_ssim,
    ));
    fs::remove_dir_all(root).unwrap();
}

#[test]
#[cfg(feature = "render")]
fn pdf_import_differential_rejects_geometry_text_link_and_pixel_perturbations() {
    let source = source_built_pdf_for_import();
    let imported =
        rpptx::Presentation::from_pdf_bytes(&source, rpptx::PdfImportMode::Editable).unwrap();
    let baseline_png = f225_render_png(&imported.presentation);
    assert!(f225_pdf_gate_accepts(
        imported.presentation.slide_size() == Some((rpptx::Emu(9_144_000), rpptx::Emu(9_144_000)))
            && f225_first_image_x_points(&imported.presentation) == 672.0,
        imported.presentation.slide(0).unwrap().text() == "I\nR",
        f225_actual_link_matches(&imported.presentation, "https://example.com/pdf"),
        f225_first_shape_x_points(&imported.presentation),
        smartart_png_ssim(&baseline_png, &baseline_png),
    ));

    let mut geometry = rpptx::Presentation::from_pdf_bytes(&source, rpptx::PdfImportMode::Editable)
        .unwrap()
        .presentation;
    geometry
        .slide_mut(0)
        .unwrap()
        .shape_mut(0)
        .unwrap()
        .set_position(rpptx::Emu(12_827), rpptx::Emu(0))
        .unwrap();
    assert!(!f225_pdf_gate_accepts(
        geometry.slide_size() == Some((rpptx::Emu(9_144_000), rpptx::Emu(9_144_000))),
        geometry.slide(0).unwrap().text() == "I\nR",
        f225_actual_link_matches(&geometry, "https://example.com/pdf"),
        f225_first_shape_x_points(&geometry),
        smartart_png_ssim(&f225_render_png(&geometry), &baseline_png),
    ));

    let mut one_pixel_shift =
        rpptx::Presentation::from_pdf_bytes(&source, rpptx::PdfImportMode::Editable)
            .unwrap()
            .presentation;
    one_pixel_shift
        .slide_mut(0)
        .unwrap()
        .shape_mut(0)
        .unwrap()
        .set_position(rpptx::Emu(6_096), rpptx::Emu(0))
        .unwrap();
    let one_pixel_at_150_dpi = 72.0 / 150.0;
    let point_error = f225_first_shape_x_points(&one_pixel_shift).abs();
    assert!((point_error - one_pixel_at_150_dpi).abs() < 1.0 / 12_700.0);
    let shifted_ssim = smartart_png_ssim(&f225_render_png(&one_pixel_shift), &baseline_png);
    eprintln!("F-225 one-pixel imported geometry raw SSIM: {shifted_ssim:.9}");
    assert!(
        shifted_ssim < F225_PDF_SSIM_FLOOR,
        "raw SSIM {shifted_ssim}"
    );
    let geometry_matches = one_pixel_shift.slide_size()
        == Some((rpptx::Emu(9_144_000), rpptx::Emu(9_144_000)))
        && f225_first_image_x_points(&one_pixel_shift) == 672.0;
    let text_matches = one_pixel_shift.slide(0).unwrap().text() == "I\nR";
    let link_matches = f225_actual_link_matches(&one_pixel_shift, "https://example.com/pdf");
    assert!(geometry_matches && text_matches && link_matches && point_error <= 1.0);
    assert!(!f225_pdf_gate_accepts(
        geometry_matches,
        text_matches,
        link_matches,
        point_error,
        shifted_ssim,
    ));

    let mut text = rpptx::Presentation::from_pdf_bytes(&source, rpptx::PdfImportMode::Editable)
        .unwrap()
        .presentation;
    text.slide_mut(0)
        .unwrap()
        .shape_mut(2)
        .unwrap()
        .child_mut(0)
        .unwrap()
        .set_text("changed")
        .unwrap();
    assert!(!f225_pdf_gate_accepts(
        text.slide_size() == Some((rpptx::Emu(9_144_000), rpptx::Emu(9_144_000))),
        text.slide(0).unwrap().text() == "I\nR",
        f225_actual_link_matches(&text, "https://example.com/pdf"),
        f225_first_shape_x_points(&text),
        smartart_png_ssim(&f225_render_png(&text), &baseline_png),
    ));

    let wrong_link = rpptx::Presentation::from_pdf_bytes(
        &pdf_source_with_uri(&source, "https://example.com/changed"),
        rpptx::PdfImportMode::Editable,
    )
    .unwrap()
    .presentation;
    assert!(!f225_pdf_gate_accepts(
        wrong_link.slide_size() == Some((rpptx::Emu(9_144_000), rpptx::Emu(9_144_000))),
        wrong_link.slide(0).unwrap().text() == "I\nR",
        f225_actual_link_matches(&wrong_link, "https://example.com/pdf"),
        f225_first_shape_x_points(&wrong_link),
        smartart_png_ssim(&f225_render_png(&wrong_link), &baseline_png),
    ));

    let mut calibrated = None;
    for blue in [
        250_u8, 245, 240, 235, 230, 220, 210, 200, 180, 160, 128, 64, 0,
    ] {
        let mut pixels =
            rpptx::Presentation::from_pdf_bytes(&source, rpptx::PdfImportMode::Editable)
                .unwrap()
                .presentation;
        let fill = rpptx::Fill::from_xml(
            format!(
                r#"<a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:srgbClr val="0000{blue:02X}"/></a:solidFill>"#
            )
            .as_bytes(),
        )
        .unwrap();
        pixels
            .slide_mut(0)
            .unwrap()
            .shape_mut(0)
            .unwrap()
            .set_fill(fill)
            .unwrap();
        let ssim = smartart_png_ssim(&f225_render_png(&pixels), &baseline_png);
        if ssim < F225_PDF_SSIM_FLOOR {
            calibrated = Some((pixels, ssim));
            break;
        }
    }
    let (pixels, ssim) = calibrated.expect("a real editable fill perturbation crosses the floor");
    assert!(F225_PDF_SSIM_FLOOR - ssim < 0.05, "calibrated SSIM {ssim}");
    assert!(!f225_pdf_gate_accepts(
        pixels.slide_size() == Some((rpptx::Emu(9_144_000), rpptx::Emu(9_144_000))),
        pixels.slide(0).unwrap().text() == "I\nR",
        f225_actual_link_matches(&pixels, "https://example.com/pdf"),
        f225_first_shape_x_points(&pixels),
        ssim,
    ));
}

#[test]
#[cfg(feature = "default-template")]
#[ignore = "requires Google Chrome 152.0.7977.65"]
fn source_built_html_matches_pinned_chrome_after_save_and_reopen() {
    const CHROME: &str = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
    const VERSION: &str = "Google Chrome 152.0.7977.65";
    let version = Command::new(CHROME).arg("--version").output().unwrap();
    assert!(version.status.success());
    assert_eq!(String::from_utf8(version.stdout).unwrap().trim(), VERSION);

    let root = std::env::temp_dir().join(format!("rpptx-f224-chrome-{}", std::process::id()));
    if root.exists() {
        fs::remove_dir_all(&root).unwrap();
    }
    fs::create_dir_all(&root).unwrap();
    let profile = root.join("profile");
    let source_path = root.join("source.html");
    let screenshot_path = root.join("chrome.png");
    let image_path = root.join("oracle.png");
    let oracle_png = valid_one_pixel_png();
    fs::write(&image_path, &oracle_png).unwrap();
    let image_url = format!("file://{}", image_path.display());
    let font_url = format!(
        "file://{}",
        workspace_root()
            .join("crates/oxml-layout/fonts/Carlito-Regular.ttf")
            .display()
    );
    let html = format!(
        "<!doctype html><html><head><style>@font-face{{font-family:Carlito;src:url('{font_url}')}}html,body{{width:1280px;height:720px;margin:0;overflow:hidden;background:#fff}}</style></head><body><section data-slide><div style='position:absolute;left:96px;top:72px;width:320px;height:180px;background:#336699;color:#ffffff;font-family:Carlito;font-size:24pt'><a href='https://example.com'>Chrome oracle</a></div><img src='{image_url}' style='position:absolute;left:448px;top:72px;width:64px;height:64px'></section></body></html>"
    );
    fs::write(&source_path, &html).unwrap();
    let source_url = format!("file://{}", source_path.display());
    let screenshot_arg = format!("--screenshot={}", screenshot_path.display());
    let profile_arg = format!("--user-data-dir={}", profile.display());
    let mut chrome = Command::new(CHROME)
        .args([
            "--headless=new",
            "--disable-gpu",
            "--disable-background-networking",
            "--disable-component-update",
            "--disable-default-apps",
            "--disable-sync",
            "--hide-scrollbars",
            "--metrics-recording-only",
            "--no-first-run",
            "--window-size=1280,720",
            "--host-resolver-rules=MAP * ~NOTFOUND",
            &profile_arg,
            &screenshot_arg,
            &source_url,
        ])
        .stdout(std::process::Stdio::null())
        .stderr(std::process::Stdio::null())
        .spawn()
        .unwrap();
    let mut exited = None;
    for _ in 0..200 {
        exited = chrome.try_wait().unwrap();
        if exited.is_some() || screenshot_path.exists() {
            break;
        }
        std::thread::sleep(std::time::Duration::from_millis(25));
    }
    match exited {
        Some(status) => assert!(status.success()),
        None => {
            chrome.kill().unwrap();
            let _ = chrome.wait();
        }
    }
    assert!(
        screenshot_path.exists(),
        "Chrome did not produce a screenshot"
    );
    let chrome_png = fs::read(&screenshot_path).unwrap();

    let imported = rpptx::Presentation::from_html(
        &html,
        &[rpptx::HtmlImageResource {
            source: &image_url,
            bytes: &oracle_png,
            filename: "oracle.png",
        }],
    )
    .unwrap();
    let bytes = imported.presentation.to_bytes().unwrap();
    let reopened = rpptx::Presentation::from_bytes(&bytes).unwrap();
    assert_eq!(reopened.len(), 1);
    assert_eq!(reopened.slide(0).unwrap().shapes().len(), 2);
    assert_eq!(
        reopened.slide(0).unwrap().shape(0).unwrap().kind(),
        rpptx::ShapeKind::Shape
    );
    assert_eq!(
        reopened.slide(0).unwrap().shape(1).unwrap().kind(),
        rpptx::ShapeKind::Picture
    );
    assert_eq!(reopened.slide(0).unwrap().text(), "Chrome oracle");
    let text_shape = reopened.slide(0).unwrap().shape(0).unwrap();
    let run = text_shape
        .text_frame()
        .unwrap()
        .paragraph(0)
        .unwrap()
        .run(0)
        .unwrap();
    let properties = run.properties().unwrap();
    assert_eq!(properties.font_size, Some(2_400));
    assert_eq!(properties.latin.as_ref().unwrap().typeface, "Carlito");
    assert_eq!(
        properties.fill,
        Some(
            Fill::from_xml(
                br#"<a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:srgbClr val="FFFFFF"/></a:solidFill>"#,
            )
            .unwrap()
        )
    );
    let package = OpcPackage::from_reader(Cursor::new(&bytes)).unwrap();
    assert!(
        package
            .get_part_rels("/ppt/slides/slide1.xml")
            .unwrap()
            .items
            .iter()
            .any(|relationship| {
                relationship.rel_type == rel_types::HYPERLINK
                    && relationship.target == "https://example.com"
                    && relationship.target_mode.as_deref() == Some("External")
            })
    );
    let (_, layout) = reopened.render_deterministic().unwrap();
    let rust_png = oxml_pdf::render_page_to_png(&layout, 0, 96.0).unwrap();
    let chrome = tiny_skia::Pixmap::decode_png(&chrome_png).unwrap();
    let rust = tiny_skia::Pixmap::decode_png(&rust_png).unwrap();
    assert_eq!((rust.width(), rust.height()), (1_280, 720));
    assert_eq!((chrome.width(), chrome.height()), (1_280, 720));
    fn color_bounds(pixmap: &tiny_skia::Pixmap, color: [u8; 4]) -> (u32, u32, u32, u32) {
        let mut left = pixmap.width();
        let mut top = pixmap.height();
        let mut right = 0_u32;
        let mut bottom = 0_u32;
        for (index, pixel) in pixmap.data().chunks_exact(4).enumerate() {
            if pixel != color {
                continue;
            }
            let x = index as u32 % pixmap.width();
            let y = index as u32 / pixmap.width();
            left = left.min(x);
            top = top.min(y);
            right = right.max(x + 1);
            bottom = bottom.max(y + 1);
        }
        assert!(left < right && top < bottom, "target color is absent");
        (left, top, right - left, bottom - top)
    }
    let expected = [96_u32, 72_u32, 320_u32, 180_u32];
    let chrome_bounds = color_bounds(&chrome, [0x33, 0x66, 0x99, 0xff]);
    let rust_bounds = color_bounds(&rust, [0x33, 0x66, 0x99, 0xff]);
    for (actual, expected) in [
        chrome_bounds.0,
        chrome_bounds.1,
        chrome_bounds.2,
        chrome_bounds.3,
    ]
    .into_iter()
    .zip(expected)
    {
        assert!(
            actual.abs_diff(expected) <= 1,
            "Chrome geometry {chrome_bounds:?}"
        );
    }
    for (actual, expected) in [rust_bounds.0, rust_bounds.1, rust_bounds.2, rust_bounds.3]
        .into_iter()
        .zip(expected)
    {
        assert!(
            actual.abs_diff(expected) <= 1,
            "Rust geometry {rust_bounds:?}"
        );
    }
    let ssim = smartart_png_ssim(&rust_png, &chrome_png);
    eprintln!("F-224 Chrome full-image luminance SSIM: {ssim:.9}");
    assert!(
        f224_browser_gate_accepts(true, true, 1, ssim),
        "Chrome full-image luminance SSIM {ssim} is below {F224_BROWSER_SSIM_FLOOR}"
    );
    fs::remove_dir_all(root).unwrap();
}

#[test]
#[cfg(feature = "default-template")]
fn html_browser_differential_rejects_geometry_text_and_pixel_perturbations() {
    let mut reference = tiny_skia::Pixmap::new(32, 32).unwrap();
    reference.fill(tiny_skia::Color::from_rgba8(0x33, 0x66, 0x99, 0xff));
    let mut perturbed = tiny_skia::Pixmap::new(32, 32).unwrap();
    perturbed.fill(tiny_skia::Color::from_rgba8(0xff, 0xff, 0xff, 0xff));
    let reference = reference.encode_png().unwrap();
    let perturbed = perturbed.encode_png().unwrap();
    let matching_ssim = smartart_png_ssim(&reference, &reference);
    let perturbed_ssim = smartart_png_ssim(&reference, &perturbed);

    assert!(f224_browser_gate_accepts(true, true, 1, matching_ssim));
    assert!(!f224_browser_gate_accepts(false, true, 1, matching_ssim));
    assert!(!f224_browser_gate_accepts(true, false, 1, matching_ssim));
    assert!(!f224_browser_gate_accepts(true, true, 2, matching_ssim));
    assert!(perturbed_ssim < F224_BROWSER_SSIM_FLOOR);
    assert!(!f224_browser_gate_accepts(true, true, 1, perturbed_ssim));
}

#[test]
fn embedded_inventory_reports_exact_hashes_relationships_and_signature_state() {
    let mut package = embedded_fixture_package(true);
    let slide = String::from_utf8(package.get_part(SLIDE_TWO_PART).unwrap().to_vec()).unwrap();
    package.set_part(
        SLIDE_TWO_PART,
        slide
            .replacen(
                "</a:graphicData>",
                r#"<x:opaque xmlns:x="urn:f218:opaque"><p:oleObj r:id="lookalike-rel"/></x:opaque></a:graphicData>"#,
                1,
            )
            .into_bytes(),
    );
    package.set_part(
        "/custom/embeddings/lookalike.bin",
        b"same-namespace-lookalike".to_vec(),
    );
    package.content_types.add_override(
        "/custom/embeddings/lookalike.bin",
        "application/vnd.openxmlformats-officedocument.oleObject",
    );
    package.get_or_create_part_rels(SLIDE_TWO_PART).add_with_id(
        "lookalike-rel",
        rel_types::OLE_OBJECT,
        "../embeddings/lookalike.bin",
    );
    let presentation = Presentation::from_bytes(&package_bytes(package)).unwrap();
    let inventory = presentation.embedded_content().unwrap();
    assert_eq!(inventory.len(), 3);
    assert_eq!(
        inventory
            .iter()
            .map(|info| (
                info.kind,
                info.source_part.as_str(),
                info.relationship_id.as_str(),
                info.target_part.as_str(),
                info.content_type.as_str(),
                info.byte_len,
                hex32(info.sha256),
                info.signature_state,
            ))
            .collect::<Vec<_>>(),
        vec![
            (
                EmbeddedContentKind::ActiveXControl,
                "/custom/activeX/activeX1.xml",
                "binary-rel",
                "/custom/activeX/activeX1.bin",
                "application/vnd.ms-office.activeX",
                18,
                "15061d052ad086ca4077fa026fd53c13e8481df7671727c37d4d30566219996a".to_owned(),
                EmbeddedSignatureState::Present,
            ),
            (
                EmbeddedContentKind::VbaProject,
                PRESENTATION_PART,
                "vba-rel",
                "/custom/vbaProject.bin",
                "application/vnd.ms-office.vbaProject",
                14,
                "7b20cff9b07d570e5f0513f47ddc73236cdfa8951c0229fe16f6eff1086ac0b1".to_owned(),
                EmbeddedSignatureState::Present,
            ),
            (
                EmbeddedContentKind::OleObject,
                SLIDE_TWO_PART,
                "ole-rel",
                "/custom/embeddings/object1.bin",
                "application/vnd.openxmlformats-officedocument.oleObject",
                14,
                "bdd4fcc46a30a68034bba88743518dc0e815c348d33ea4024bfa1c5e8134581f".to_owned(),
                EmbeddedSignatureState::Present,
            ),
        ]
    );
}

#[test]
#[ignore = "requires pinned LibreOffice 26.2.5.2"]
fn odp_reader_and_writer_match_pinned_libreoffice_both_directions() {
    let source = f222_source_presentation();
    let source_odp = source.to_odp_bytes().unwrap();
    let root = f222_temp_directory("oracle");
    fs::create_dir_all(&root).unwrap();
    let odp_path = root.join("rust-source.odp");
    fs::write(&odp_path, &source_odp.bytes).unwrap();
    f222_libreoffice_convert(&odp_path, "pptx", &root);
    let opened = Presentation::open(root.join("rust-source.pptx")).unwrap();
    assert_eq!(opened.len(), 1);
    assert!(opened.slide(0).unwrap().text().contains("ODP oracle title"));
    f222_assert_libreoffice_pdf(&odp_path, &root, "ODP oracle title");

    let pptx_path = root.join("pptx-source.pptx");
    source.save(&pptx_path).unwrap();
    f222_libreoffice_convert(&pptx_path, "odp", &root);
    let imported = Presentation::open_odp(root.join("pptx-source.odp")).unwrap();
    assert_eq!(imported.presentation.len(), 1);
    assert!(
        imported
            .presentation
            .slide(0)
            .unwrap()
            .text()
            .contains("ODP oracle title")
    );
    f222_assert_libreoffice_pdf(&pptx_path, &root, "ODP oracle title");
    fs::remove_dir_all(root).unwrap();
}

#[test]
fn odp_round_trip_preserves_supported_presentation_content() {
    let source = Presentation::from_odp_bytes(&f222_odp_fixture()).unwrap();
    assert!(source.diagnostics.is_empty());
    let slide = source.presentation.slide(0).unwrap();
    assert!(slide.text().contains("F-222 text box"));
    assert!(slide.text().contains("F-222 rectangle"));
    assert_eq!(slide.notes_text().as_deref(), Some("F-222 speaker notes"));
    assert!(
        slide
            .shapes()
            .any(|shape| shape.kind() == ShapeKind::Picture)
    );
    assert!(slide.shapes().any(|shape| shape.table().is_some()));

    let exported = source.presentation.to_odp_bytes().unwrap();
    assert!(exported.diagnostics.is_empty());
    let reopened = Presentation::from_odp_bytes(&exported.bytes).unwrap();
    let issues = reopened.presentation.validate();
    assert!(issues.is_empty(), "{issues:?}");
    Presentation::from_bytes(&reopened.presentation.to_bytes().unwrap()).unwrap();
    assert_eq!(reopened.presentation.len(), 1);
    let slide = reopened.presentation.slide(0).unwrap();
    assert!(slide.text().contains("F-222 text box"));
    assert!(slide.text().contains("F-222 rectangle"));
    assert!(slide.text().contains("row 2, column 2"));
    assert_eq!(slide.notes_text().as_deref(), Some("F-222 speaker notes"));
    assert!(
        slide
            .shapes()
            .any(|shape| shape.kind() == ShapeKind::Picture)
    );
}

#[test]
fn odp_packages_are_bounded_deterministic_and_manifest_complete() {
    let presentation = Presentation::from_odp_bytes(&f222_odp_fixture())
        .unwrap()
        .presentation;
    let first = presentation.to_odp_bytes().unwrap().bytes;
    let second = presentation.to_odp_bytes().unwrap().bytes;
    assert_eq!(first, second);
    let mut archive = zip::ZipArchive::new(Cursor::new(&first)).unwrap();
    let mimetype = archive.by_index(0).unwrap();
    assert_eq!(mimetype.name(), "mimetype");
    assert_eq!(mimetype.compression(), zip::CompressionMethod::Stored);
    drop(mimetype);
    let manifest = {
        let mut entry = archive.by_name("META-INF/manifest.xml").unwrap();
        let mut bytes = Vec::new();
        std::io::Read::read_to_end(&mut entry, &mut bytes).unwrap();
        String::from_utf8(bytes).unwrap()
    };
    assert!(manifest.contains("Pictures/image-1-4.png"));
    assert!(
        Presentation::from_odp_bytes_with_limits(
            &first,
            rpptx::PackageReadLimits {
                max_entries: 2,
                max_part_uncompressed_bytes: u64::MAX,
                max_total_uncompressed_bytes: u64::MAX,
            }
        )
        .is_err()
    );

    let aliased = f222_minimal_odp(
        r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"><o:body><o:presentation><d:page d:name="aliased"/></o:presentation></o:body></o:document-content>"#,
        None,
        false,
    );
    assert_eq!(
        Presentation::from_odp_bytes(&aliased)
            .unwrap()
            .presentation
            .len(),
        1
    );

    let nested_lookalike = f222_minimal_odp(
        r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:x="urn:opaque"><o:body><x:opaque><o:presentation><d:page/></o:presentation></x:opaque></o:body></o:document-content>"#,
        None,
        false,
    );
    assert!(Presentation::from_odp_bytes(&nested_lookalike).is_err());

    let unsafe_path = f222_minimal_odp(
        r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><o:body><o:presentation/></o:body></o:document-content>"#,
        Some(("../unsafe.bin", b"unsafe")),
        false,
    );
    assert!(Presentation::from_odp_bytes(&unsafe_path).is_err());

    let duplicate = f222_minimal_odp(
        r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:x="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"><o:body><o:presentation><d:page d:name="first" x:name="second"/></o:presentation></o:body></o:document-content>"#,
        None,
        false,
    );
    assert!(Presentation::from_odp_bytes(&duplicate).is_err());

    let mut oversized = Presentation::new().unwrap();
    oversized.add_slide(0).unwrap();
    oversized
        .slide_mut(0)
        .unwrap()
        .add_textbox(
            Emu::from_cm(1.0),
            Emu::from_cm(1.0),
            Emu::from_cm(5.0),
            Emu::from_cm(2.0),
        )
        .unwrap()
        .set_text(&"'".repeat(3_000_000))
        .unwrap();
    assert!(oversized.to_odp_bytes().is_err());
}

#[test]
fn odp_lossy_diagnostics_are_stable_bounded_and_location_aware() {
    let mut presentation = Presentation::from_odp_bytes(&f222_odp_fixture())
        .unwrap()
        .presentation;
    presentation
        .slide_mut(0)
        .unwrap()
        .add_connector(
            ConnectorType::Straight,
            Emu::from_cm(1.0),
            Emu::from_cm(1.0),
            Emu::from_cm(4.0),
            Emu::from_cm(4.0),
        )
        .unwrap();
    let first = presentation.to_odp_bytes().unwrap();
    let second = presentation.to_odp_bytes().unwrap();
    assert_eq!(first.diagnostics, second.diagnostics);
    assert_eq!(first.diagnostics.len(), 1);
    assert_eq!(first.diagnostics[0].path, "slides/0/shapes/4");
    assert_eq!(
        first.diagnostics[0].message,
        "unsupported shape was dropped"
    );

    let unsupported = f222_minimal_odp(
        r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"><o:body><o:presentation><d:page><d:line/></d:page></o:presentation></o:body></o:document-content>"#,
        None,
        false,
    );
    let diagnostics = Presentation::from_odp_bytes(&unsupported)
        .unwrap()
        .diagnostics;
    assert_eq!(diagnostics.len(), 1);
    assert_eq!(diagnostics[0].path, "slides/0/objects/0");
    assert_eq!(
        diagnostics[0].message,
        "unsupported ODP line content was dropped"
    );

    let lines = "<d:line/>".repeat(5_001);
    let bounded_content = format!(
        r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"><o:body><o:presentation><d:page>{lines}</d:page><d:page>{lines}</d:page></o:presentation></o:body></o:document-content>"#
    );
    let bounded = f222_minimal_odp(&bounded_content, None, false);
    let error = match Presentation::from_odp_bytes(&bounded) {
        Ok(_) => panic!("diagnostic ceiling accepted unbounded ODP losses"),
        Err(error) => error,
    };
    assert!(error.to_string().contains("diagnostic limit"));
}

#[test]
fn failed_odp_read_and_save_never_publish_partial_state() {
    assert!(Presentation::from_odp_bytes(b"not an ODP package").is_err());
    let presentation = f222_source_presentation();
    let before = presentation.to_bytes().unwrap();
    let directory = f222_temp_directory("save-failure");
    fs::create_dir_all(&directory).unwrap();
    let destination = directory.join("existing.odp");
    fs::write(&destination, b"sentinel ODP destination").unwrap();
    #[cfg(unix)]
    {
        use std::os::unix::fs::PermissionsExt as _;

        fs::set_permissions(&directory, fs::Permissions::from_mode(0o500)).unwrap();
        assert!(presentation.save_odp(&destination).is_err());
        fs::set_permissions(&directory, fs::Permissions::from_mode(0o700)).unwrap();
        assert_eq!(fs::read(&destination).unwrap(), b"sentinel ODP destination");
    }
    #[cfg(not(unix))]
    {
        let invalid_destination = directory.join("missing").join("output.odp");
        assert!(presentation.save_odp(invalid_destination).is_err());
        assert_eq!(fs::read(&destination).unwrap(), b"sentinel ODP destination");
    }
    assert_eq!(presentation.to_bytes().unwrap(), before);
    fs::remove_dir_all(directory).unwrap();
}

fn f222_odp_fixture() -> Vec<u8> {
    let content = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" office:version="1.3"><office:body><office:presentation><draw:page draw:name="F-222 slide"><draw:frame svg:x="1cm" svg:y="1cm" svg:width="8cm" svg:height="2cm"><draw:text-box><text:p>F-222 text box</text:p></draw:text-box></draw:frame><draw:custom-shape svg:x="1cm" svg:y="4cm" svg:width="4cm" svg:height="2cm"><text:p>F-222 rectangle</text:p><draw:enhanced-geometry draw:type="rectangle"/></draw:custom-shape><draw:frame svg:x="6cm" svg:y="4cm" svg:width="8cm" svg:height="4cm"><table:table><table:table-row><table:table-cell office:value-type="string"><text:p>row 1, column 1</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>row 1, column 2</text:p></table:table-cell></table:table-row><table:table-row><table:table-cell office:value-type="string"><text:p>row 2, column 1</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>row 2, column 2</text:p></table:table-cell></table:table-row></table:table></draw:frame><draw:frame svg:x="15cm" svg:y="1cm" svg:width="2cm" svg:height="2cm"><draw:image xlink:href="Pictures/source.png" xlink:type="simple"/></draw:frame><presentation:notes><draw:frame svg:x="0cm" svg:y="0cm" svg:width="20cm" svg:height="5cm"><draw:text-box><text:p>F-222 speaker notes</text:p></draw:text-box></draw:frame></presentation:notes></draw:page></office:presentation></office:body></office:document-content>"#;
    let manifest = r#"<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.presentation"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="Pictures/source.png" manifest:media-type="image/png"/></manifest:manifest>"#;
    let mut cursor = Cursor::new(Vec::new());
    {
        let mut archive = zip::ZipWriter::new(&mut cursor);
        let stored = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);
        let deflated = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated);
        archive.start_file("mimetype", stored).unwrap();
        std::io::Write::write_all(
            &mut archive,
            b"application/vnd.oasis.opendocument.presentation",
        )
        .unwrap();
        archive.start_file("content.xml", deflated).unwrap();
        std::io::Write::write_all(&mut archive, content.as_bytes()).unwrap();
        archive.start_file("Pictures/source.png", deflated).unwrap();
        std::io::Write::write_all(&mut archive, &valid_one_pixel_png()).unwrap();
        archive
            .start_file("META-INF/manifest.xml", deflated)
            .unwrap();
        std::io::Write::write_all(&mut archive, manifest.as_bytes()).unwrap();
        archive.finish().unwrap();
    }
    cursor.into_inner()
}

fn f222_minimal_odp(
    content: &str,
    extra: Option<(&str, &[u8])>,
    duplicate_content: bool,
) -> Vec<u8> {
    let manifest = r#"<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.presentation"/></manifest:manifest>"#;
    let mut cursor = Cursor::new(Vec::new());
    {
        let mut archive = zip::ZipWriter::new(&mut cursor);
        let stored = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);
        let deflated = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated);
        archive.start_file("mimetype", stored).unwrap();
        std::io::Write::write_all(
            &mut archive,
            b"application/vnd.oasis.opendocument.presentation",
        )
        .unwrap();
        for _ in 0..if duplicate_content { 2 } else { 1 } {
            archive.start_file("content.xml", deflated).unwrap();
            std::io::Write::write_all(&mut archive, content.as_bytes()).unwrap();
        }
        if let Some((name, bytes)) = extra {
            archive.start_file(name, deflated).unwrap();
            std::io::Write::write_all(&mut archive, bytes).unwrap();
        }
        archive
            .start_file("META-INF/manifest.xml", deflated)
            .unwrap();
        std::io::Write::write_all(&mut archive, manifest.as_bytes()).unwrap();
        archive.finish().unwrap();
    }
    cursor.into_inner()
}

fn f222_source_presentation() -> Presentation {
    let mut presentation = Presentation::new().unwrap();
    presentation.add_slide(0).unwrap();
    presentation
        .slide_mut(0)
        .unwrap()
        .add_textbox(
            Emu::from_cm(1.0),
            Emu::from_cm(1.0),
            Emu::from_cm(8.0),
            Emu::from_cm(2.0),
        )
        .unwrap()
        .set_text("ODP oracle title")
        .unwrap();
    presentation
        .slide_mut(0)
        .unwrap()
        .add_shape(
            "rect",
            Emu::from_cm(1.0),
            Emu::from_cm(4.0),
            Emu::from_cm(4.0),
            Emu::from_cm(2.0),
        )
        .unwrap()
        .set_text("ODP oracle rectangle")
        .unwrap();
    presentation
        .slide_mut(0)
        .unwrap()
        .add_table(
            2,
            2,
            Emu::from_cm(6.0),
            Emu::from_cm(4.0),
            Emu::from_cm(8.0),
            Emu::from_cm(4.0),
        )
        .unwrap();
    presentation
}

fn f222_temp_directory(label: &str) -> PathBuf {
    static COUNTER: AtomicU64 = AtomicU64::new(0);
    let ordinal = COUNTER.fetch_add(1, std::sync::atomic::Ordering::Relaxed);
    std::env::temp_dir().join(format!(
        "rpptx-f222-{label}-{}-{ordinal}",
        std::process::id()
    ))
}

fn f222_libreoffice_convert(source: &Path, extension: &str, output: &Path) {
    let version = Command::new("soffice").arg("--version").output().unwrap();
    assert_eq!(
        String::from_utf8(version.stdout).unwrap().trim(),
        LIBREOFFICE_VERSION
    );
    let profile = f222_temp_directory("libreoffice-profile");
    let profile_arg = format!("-env:UserInstallation=file://{}", profile.display());
    let result = Command::new("soffice")
        .args([
            "--headless",
            &profile_arg,
            "--convert-to",
            extension,
            "--outdir",
        ])
        .arg(output)
        .arg(source)
        .output()
        .unwrap();
    assert!(
        result.status.success(),
        "LibreOffice conversion failed: {}",
        String::from_utf8_lossy(&result.stderr)
    );
    if profile.exists() {
        fs::remove_dir_all(profile).unwrap();
    }
}

fn f222_assert_libreoffice_pdf(source: &Path, output: &Path, expected_text: &str) {
    f222_libreoffice_convert(source, "pdf", output);
    let pdf = output.join(format!(
        "{}.pdf",
        source.file_stem().unwrap().to_string_lossy()
    ));
    let info = Command::new("pdfinfo").arg(&pdf).output().unwrap();
    assert!(info.status.success());
    assert!(
        String::from_utf8(info.stdout)
            .unwrap()
            .contains("Pages:           1")
    );
    let text = Command::new("pdftotext")
        .args([pdf.to_str().unwrap(), "-"])
        .output()
        .unwrap();
    assert!(text.status.success());
    assert!(
        String::from_utf8(text.stdout)
            .unwrap()
            .contains(expected_text)
    );
}

#[test]
fn distinct_ole_ids_in_one_compatibility_frame_fail_closed_atomically() {
    let mut package = embedded_fixture_package(false);
    let slide = String::from_utf8(package.get_part(SLIDE_TWO_PART).unwrap().to_vec()).unwrap();
    let first = r#"<p:oleObj r:id="ole-rel" name="opaque"><p:embed/><x:producer xmlns:x="urn:f218" marker="kept"/></p:oleObj>"#;
    let second = r#"<p:oleObj r:id="alternate-ole-rel" name="opaque"><p:embed/><x:producer xmlns:x="urn:f218" marker="alternate"/></p:oleObj>"#;
    let compatibility = format!(
        r#"<mc:AlternateContent xmlns:mc="{MC_NS}"><mc:Choice Requires="p">{first}</mc:Choice><mc:Fallback>{second}</mc:Fallback></mc:AlternateContent>"#
    );
    assert!(slide.contains(first));
    package.set_part(
        SLIDE_TWO_PART,
        slide.replacen(first, &compatibility, 1).into_bytes(),
    );
    package.set_part(
        "/custom/embeddings/alternate.bin",
        b"alternate-ole-executable".to_vec(),
    );
    package.content_types.add_override(
        "/custom/embeddings/alternate.bin",
        "application/vnd.openxmlformats-officedocument.oleObject",
    );
    package.get_or_create_part_rels(SLIDE_TWO_PART).add_with_id(
        "alternate-ole-rel",
        rel_types::OLE_OBJECT,
        "../embeddings/alternate.bin",
    );

    let mut presentation = Presentation::from_bytes(&package_bytes(package)).unwrap();
    let before = presentation.to_bytes().unwrap();
    assert!(presentation.embedded_content().is_err());
    for relationship_id in ["ole-rel", "alternate-ole-rel"] {
        assert!(
            presentation
                .remove_embedded_content(
                    SLIDE_TWO_PART,
                    relationship_id,
                    EmbeddedMutationPolicy::RemoveInvalidatedSignatures,
                )
                .is_err(),
            "{relationship_id}"
        );
        assert_eq!(
            presentation.to_bytes().unwrap(),
            before,
            "{relationship_id}"
        );
    }
}

#[test]
fn ole_inventory_and_mutation_use_layout_and_master_producing_scopes() {
    let mut presentation =
        Presentation::from_bytes(&package_bytes(embedded_layout_and_master_fixture_package()))
            .unwrap();
    let inventory = presentation.embedded_content().unwrap();
    assert_eq!(
        inventory
            .iter()
            .filter(|info| info.kind == EmbeddedContentKind::OleObject)
            .map(|info| (
                info.source_part.as_str(),
                info.relationship_id.as_str(),
                info.target_part.as_str(),
            ))
            .collect::<Vec<_>>(),
        vec![
            (
                "/custom/layouts/validation.xml",
                "layout-ole",
                "/custom/embeddings/layout.bin",
            ),
            (
                "/custom/masters/embedded.xml",
                "master-ole",
                "/custom/embeddings/master.bin",
            ),
            (SLIDE_TWO_PART, "ole-rel", "/custom/embeddings/object1.bin",),
        ]
    );
    assert_eq!(
        presentation
            .extract_embedded_content("/custom/layouts/validation.xml", "layout-ole")
            .unwrap(),
        b"layout-owned-ole"
    );
    assert_eq!(
        presentation
            .extract_embedded_content("/custom/masters/embedded.xml", "master-ole")
            .unwrap(),
        b"master-owned-ole"
    );

    let replaced = presentation
        .replace_embedded_content(
            "/custom/layouts/validation.xml",
            "layout-ole",
            b"layout replacement",
            EmbeddedMutationPolicy::PreserveInvalidatedSignatures,
        )
        .unwrap();
    assert_eq!(replaced.source_part, "/custom/layouts/validation.xml");
    assert_eq!(replaced.relationship_id, "layout-ole");
    presentation
        .remove_embedded_content(
            "/custom/layouts/validation.xml",
            "layout-ole",
            EmbeddedMutationPolicy::PreserveInvalidatedSignatures,
        )
        .unwrap();
    presentation
        .remove_embedded_content(
            "/custom/masters/embedded.xml",
            "master-ole",
            EmbeddedMutationPolicy::PreserveInvalidatedSignatures,
        )
        .unwrap();
    let bytes = presentation.to_bytes().unwrap();
    let saved = open_opc(&bytes, "F-218 layout and master OLE removal");
    for part in [
        "/custom/layouts/validation.xml",
        "/custom/masters/embedded.xml",
    ] {
        let xml = String::from_utf8(saved.get_part(part).unwrap().to_vec()).unwrap();
        assert!(xml.contains("producer-boundary"), "{part}");
        assert!(!xml.contains("oleObj"), "{part}");
    }
    assert!(!saved.parts.contains_key("/custom/embeddings/layout.bin"));
    assert!(!saved.parts.contains_key("/custom/embeddings/master.bin"));
    assert!(
        Presentation::from_bytes(&bytes)
            .unwrap()
            .validate()
            .is_empty()
    );
}

#[test]
fn tracked_ole_corpus_exposes_layout_and_master_producing_scopes() {
    let path = corpus_dir().join("alterman_security.pptx");
    if !path.is_file() {
        assert!(
            std::env::var_os("RDOCX_PPTX_CORPUS_REQUIRED").is_none(),
            "required F-218 OLE corpus deck is missing: {}",
            path.display()
        );
        return;
    }
    let presentation = Presentation::open(&path).unwrap();
    let inventory = presentation.embedded_content().unwrap();
    let layout_ole = inventory
        .iter()
        .filter(|info| {
            info.kind == EmbeddedContentKind::OleObject
                && info.source_part.contains("slideLayouts/")
        })
        .count();
    let master_ole = inventory
        .iter()
        .filter(|info| {
            info.kind == EmbeddedContentKind::OleObject
                && info.source_part.contains("slideMasters/")
        })
        .count();
    assert_eq!(layout_ole, 2);
    assert_eq!(master_ole, 1);
    for info in inventory.iter().filter(|info| {
        info.kind == EmbeddedContentKind::OleObject
            && (info.source_part.contains("slideLayouts/")
                || info.source_part.contains("slideMasters/"))
    }) {
        assert_eq!(
            presentation
                .extract_embedded_content(&info.source_part, &info.relationship_id)
                .unwrap()
                .len(),
            info.byte_len
        );
    }
}

#[test]
fn embedded_extract_replace_and_remove_are_atomic_and_never_execute_payloads() {
    let mut presentation =
        Presentation::from_bytes(&package_bytes(embedded_fixture_package(false))).unwrap();
    let before_ole_xml = open_opc(&presentation.to_bytes().unwrap(), "F-218 OLE source")
        .get_part(SLIDE_TWO_PART)
        .unwrap()
        .to_vec();
    let before_ole_frame = capture_exact_subtree(
        &before_ole_xml,
        br#"<p:oleObj r:id="ole-rel""#,
        b"</p:oleObj>",
    );
    assert_eq!(
        presentation
            .extract_embedded_content(SLIDE_TWO_PART, "ole-rel")
            .unwrap(),
        b"ole-executable"
    );
    presentation
        .replace_embedded_content(
            SLIDE_TWO_PART,
            "ole-rel",
            b"opaque replacement, not a compound file",
            EmbeddedMutationPolicy::PreserveInvalidatedSignatures,
        )
        .unwrap();
    let after_ole_xml = open_opc(&presentation.to_bytes().unwrap(), "F-218 OLE replacement")
        .get_part(SLIDE_TWO_PART)
        .unwrap()
        .to_vec();
    assert_eq!(
        capture_exact_subtree(
            &after_ole_xml,
            br#"<p:oleObj r:id="ole-rel""#,
            b"</p:oleObj>",
        ),
        before_ole_frame
    );
    let replacement = presentation
        .replace_embedded_content(
            "/custom/activeX/activeX1.xml",
            "binary-rel",
            b"replacement-bytes",
            EmbeddedMutationPolicy::PreserveInvalidatedSignatures,
        )
        .unwrap();
    assert_eq!(replacement.relationship_id, "binary-rel");
    assert_eq!(replacement.target_part, "/custom/activeX/activeX1.bin");
    assert_eq!(
        hex32(replacement.sha256),
        "70bf69c13743b7193ffd7a3718caab18522b61d4643fe13ac80caa5301e2345a"
    );
    assert_eq!(
        presentation
            .extract_embedded_content("/custom/activeX/activeX1.xml", "binary-rel")
            .unwrap(),
        b"replacement-bytes"
    );

    presentation
        .remove_embedded_content(
            "/custom/activeX/activeX1.xml",
            "binary-rel",
            EmbeddedMutationPolicy::PreserveInvalidatedSignatures,
        )
        .unwrap();
    let saved = presentation.to_bytes().unwrap();
    let package = open_opc(&saved, "F-218 ActiveX removal");
    assert!(!package.parts.contains_key("/custom/activeX/activeX1.xml"));
    assert!(!package.parts.contains_key("/custom/activeX/activeX1.bin"));
    assert!(
        !package
            .content_types
            .overrides
            .contains_key("/custom/activeX/activeX1.xml")
    );
    assert!(
        !String::from_utf8_lossy(package.get_part(SLIDE_TWO_PART).unwrap()).contains("control-rel")
    );
    assert!(
        Presentation::from_bytes(&saved)
            .unwrap()
            .validate()
            .is_empty()
    );

    presentation
        .remove_embedded_content(
            SLIDE_TWO_PART,
            "ole-rel",
            EmbeddedMutationPolicy::PreserveInvalidatedSignatures,
        )
        .unwrap();
    presentation
        .remove_embedded_content(
            PRESENTATION_PART,
            "vba-rel",
            EmbeddedMutationPolicy::PreserveInvalidatedSignatures,
        )
        .unwrap();
    assert!(presentation.embedded_content().unwrap().is_empty());
    let fully_removed = open_opc(&presentation.to_bytes().unwrap(), "F-218 all removals");
    assert!(
        !fully_removed
            .parts
            .contains_key("/custom/embeddings/object1.bin")
    );
    assert!(!fully_removed.parts.contains_key("/custom/vbaProject.bin"));
    assert!(fully_removed.parts.contains_key("/custom/vbaSignature.bin"));

    let before_failure = presentation.to_bytes().unwrap();
    assert!(
        presentation
            .replace_embedded_content(
                SLIDE_TWO_PART,
                "missing",
                b"must-not-land",
                EmbeddedMutationPolicy::RemoveInvalidatedSignatures,
            )
            .is_err()
    );
    assert_eq!(presentation.to_bytes().unwrap(), before_failure);
}

#[test]
fn ordinary_presentation_edits_preserve_every_retained_executable_payload_byte() {
    let package = embedded_fixture_package(true);
    let retained = [
        "/custom/embeddings/object1.bin",
        "/custom/activeX/activeX1.bin",
        "/custom/vbaProject.bin",
        "/custom/vbaSignature.bin",
        "/_xmlsignatures/origin.sigs",
        "/_xmlsignatures/sig1.xml",
    ]
    .map(|part| (part, package.get_part(part).unwrap().to_vec()));
    let mut presentation = Presentation::from_bytes(&package_bytes(package)).unwrap();
    presentation
        .slide_mut(0)
        .unwrap()
        .shape_mut(0)
        .unwrap()
        .set_name("ordinary mutation")
        .unwrap();
    let saved = open_opc(&presentation.to_bytes().unwrap(), "F-218 ordinary mutation");
    for (part, expected) in retained {
        assert_eq!(saved.get_part(part).unwrap(), expected, "retained {part}");
    }
    assert!(
        presentation
            .embedded_content()
            .unwrap()
            .iter()
            .all(|info| info.signature_state == EmbeddedSignatureState::Invalidated)
    );
}

#[cfg(feature = "digital-signatures")]
#[test]
fn ordinary_transactional_edit_and_fresh_reopen_report_retained_signature_invalidated() {
    let (private_key, certificate) = f221_signing_material();
    let mut package = embedded_fixture_package(false);
    let report = package.sign(&private_key, &certificate).unwrap();
    assert!(report.cryptographically_valid);
    assert!(report.coverage_complete);
    let mut presentation = Presentation::from_bytes(&package_bytes(package)).unwrap();
    assert!(
        presentation
            .embedded_content()
            .unwrap()
            .iter()
            .all(|info| info.signature_state == EmbeddedSignatureState::Present)
    );

    presentation.remove_slide(0).unwrap();
    assert!(
        presentation
            .embedded_content()
            .unwrap()
            .iter()
            .all(|info| info.signature_state == EmbeddedSignatureState::Invalidated)
    );
    let reopened = Presentation::from_bytes(&presentation.to_bytes().unwrap()).unwrap();
    assert!(
        reopened
            .embedded_content()
            .unwrap()
            .iter()
            .all(|info| info.signature_state == EmbeddedSignatureState::Invalidated)
    );
}

#[test]
fn fresh_reopen_detects_a_retained_signature_reference_to_a_removed_part() {
    let mut package = embedded_fixture_package(true);
    let model = CT_Presentation::from_xml(package.get_part(PRESENTATION_PART).unwrap()).unwrap();
    let removed_slide_relationship = &model.slide_ids[0].relationship_id;
    package.set_part(
        "/_xmlsignatures/sig1.xml",
        format!(
            r#"<ds:Signature xmlns:ds="http://www.w3.org/2000/09/xmldsig#" xmlns:mdssi="http://schemas.openxmlformats.org/package/2006/digital-signature"><ds:Object><ds:Manifest><ds:Reference URI="/custom/_rels/presentation-main.xml.rels?ContentType=application/vnd.openxmlformats-package.relationships+xml"><ds:Transforms><ds:Transform><mdssi:RelationshipReference SourceId="{removed_slide_relationship}"/></ds:Transform></ds:Transforms></ds:Reference></ds:Manifest></ds:Object></ds:Signature>"#
        )
        .into_bytes(),
    );
    let mut presentation = Presentation::from_bytes(&package_bytes(package)).unwrap();
    assert!(
        presentation
            .embedded_content()
            .unwrap()
            .iter()
            .all(|info| info.signature_state == EmbeddedSignatureState::Present)
    );

    presentation.remove_slide(0).unwrap();
    let bytes = presentation.to_bytes().unwrap();
    let reopened = Presentation::from_bytes(&bytes).unwrap();
    assert!(
        reopened
            .embedded_content()
            .unwrap()
            .iter()
            .all(|info| info.signature_state == EmbeddedSignatureState::Invalidated)
    );
}

#[test]
fn fresh_reopen_marks_a_same_topology_signed_shape_edit_invalidated() {
    let bytes = package_bytes(embedded_fixture_package(true));
    let mut presentation = Presentation::from_bytes(&bytes).unwrap();
    assert!(
        presentation
            .embedded_content()
            .unwrap()
            .iter()
            .all(|info| info.signature_state == EmbeddedSignatureState::Present)
    );
    presentation
        .slide_mut(0)
        .unwrap()
        .shape_mut(0)
        .unwrap()
        .set_name("signed topology unchanged")
        .unwrap();
    let saved_bytes = presentation.to_bytes().unwrap();
    let saved = open_opc(&saved_bytes, "F-218 default signature invalidation marker");
    assert!(saved.parts.contains_key("/_xmlsignatures/origin.sigs"));
    assert!(saved.parts.contains_key("/_xmlsignatures/sig1.xml"));
    assert_eq!(
        saved
            .package_rels
            .items
            .iter()
            .filter(|relationship| { relationship.rel_type == rel_types::DIGITAL_SIGNATURE_ORIGIN })
            .count(),
        1
    );
    assert_eq!(
        saved
            .get_part_rels("/_xmlsignatures/origin.sigs")
            .unwrap()
            .items
            .iter()
            .filter(|relationship| relationship.rel_type == rel_types::DIGITAL_SIGNATURE)
            .count(),
        1
    );
    let reopened = Presentation::from_bytes(&saved_bytes).unwrap();
    assert!(
        reopened
            .embedded_content()
            .unwrap()
            .iter()
            .all(|info| info.signature_state == EmbeddedSignatureState::Invalidated)
    );
}

#[test]
fn embedded_removal_deletes_only_relationship_owned_unreachable_candidates() {
    let mut package = embedded_fixture_package(false);
    let xml = String::from_utf8(package.get_part(SLIDE_TWO_PART).unwrap().to_vec()).unwrap();
    package.set_part(
        SLIDE_TWO_PART,
        xml.replacen(
            "</p:spTree>",
            &format!("{}</p:spTree>", ole_frame_xml(902, "ole-shared")),
            1,
        )
        .into_bytes(),
    );
    package.get_or_create_part_rels(SLIDE_TWO_PART).add_with_id(
        "ole-shared",
        rel_types::OLE_OBJECT,
        "../embeddings/object1.bin",
    );
    let orphan = package
        .get_part("/custom/embeddings/orphan.bin")
        .unwrap()
        .to_vec();
    let mut presentation = Presentation::from_bytes(&package_bytes(package)).unwrap();
    presentation
        .remove_embedded_content(
            SLIDE_TWO_PART,
            "ole-rel",
            EmbeddedMutationPolicy::PreserveInvalidatedSignatures,
        )
        .unwrap();
    let saved = open_opc(
        &presentation.to_bytes().unwrap(),
        "F-218 shared target removal",
    );
    assert_eq!(
        saved.get_part("/custom/embeddings/object1.bin").unwrap(),
        b"ole-executable"
    );
    assert_eq!(
        saved.get_part("/custom/embeddings/orphan.bin").unwrap(),
        orphan
    );
    assert!(
        presentation
            .embedded_content()
            .unwrap()
            .iter()
            .any(|info| info.relationship_id == "ole-shared")
    );
    assert!(
        !presentation
            .embedded_content()
            .unwrap()
            .iter()
            .any(|info| info.relationship_id == "ole-rel")
    );
}

#[test]
fn embedded_mutation_policy_preserves_or_removes_invalidated_signature_evidence() {
    let bytes = package_bytes(embedded_fixture_package(true));
    let mut preserving = Presentation::from_bytes(&bytes).unwrap();
    let preserved = preserving
        .replace_embedded_content(
            PRESENTATION_PART,
            "vba-rel",
            b"changed vba",
            EmbeddedMutationPolicy::PreserveInvalidatedSignatures,
        )
        .unwrap();
    assert_eq!(
        preserved.signature_state,
        EmbeddedSignatureState::Invalidated
    );
    let preserved_package = open_opc(&preserving.to_bytes().unwrap(), "F-218 preserve signatures");
    assert!(
        preserved_package
            .parts
            .contains_key("/custom/vbaSignature.bin")
    );
    assert!(
        preserved_package
            .parts
            .contains_key("/_xmlsignatures/sig1.xml")
    );

    let mut removing = Presentation::from_bytes(&bytes).unwrap();
    let unsigned = removing
        .replace_embedded_content(
            PRESENTATION_PART,
            "vba-rel",
            b"changed vba",
            EmbeddedMutationPolicy::RemoveInvalidatedSignatures,
        )
        .unwrap();
    assert_eq!(unsigned.signature_state, EmbeddedSignatureState::Absent);
    let unsigned_package = open_opc(&removing.to_bytes().unwrap(), "F-218 remove signatures");
    assert!(
        !unsigned_package
            .parts
            .contains_key("/custom/vbaSignature.bin")
    );
    assert!(
        !unsigned_package
            .parts
            .contains_key("/_xmlsignatures/origin.sigs")
    );
    assert!(
        !unsigned_package
            .parts
            .contains_key("/_xmlsignatures/sig1.xml")
    );
    assert!(
        unsigned_package
            .package_rels
            .get_by_type(rel_types::DIGITAL_SIGNATURE_ORIGIN)
            .is_none()
    );
}

#[test]
fn preserved_vba_signature_invalidation_survives_a_fresh_reopen() {
    let bytes = package_bytes(embedded_fixture_package(false));
    let original = open_opc(&bytes, "F-218 original VBA signature");
    let signature_bytes = original
        .get_part("/custom/vbaSignature.bin")
        .unwrap()
        .to_vec();
    let mut presentation = Presentation::from_bytes(&bytes).unwrap();
    let replaced = presentation
        .replace_embedded_content(
            PRESENTATION_PART,
            "vba-rel",
            b"changed VBA without a package signature",
            EmbeddedMutationPolicy::PreserveInvalidatedSignatures,
        )
        .unwrap();
    assert_eq!(
        replaced.signature_state,
        EmbeddedSignatureState::Invalidated
    );
    let saved_bytes = presentation.to_bytes().unwrap();
    let saved = open_opc(&saved_bytes, "F-218 persisted VBA invalidation");
    assert_eq!(
        saved.get_part("/custom/vbaSignature.bin").unwrap(),
        signature_bytes
    );
    assert!(
        saved
            .get_part_rels("/custom/vbaProject.bin")
            .unwrap()
            .items
            .iter()
            .any(|relationship| relationship.rel_type == rel_types::VBA_PROJECT_SIGNATURE)
    );
    let mut reopened = Presentation::from_bytes(&saved_bytes).unwrap();
    let reopened_vba = reopened
        .embedded_content()
        .unwrap()
        .into_iter()
        .find(|info| info.kind == EmbeddedContentKind::VbaProject)
        .unwrap();
    assert_eq!(
        reopened_vba.signature_state,
        EmbeddedSignatureState::Invalidated
    );
    reopened
        .replace_embedded_content(
            PRESENTATION_PART,
            "vba-rel",
            b"changed again",
            EmbeddedMutationPolicy::RemoveInvalidatedSignatures,
        )
        .unwrap();
    let unsigned = open_opc(
        &reopened.to_bytes().unwrap(),
        "F-218 remove persisted VBA invalidation",
    );
    assert!(!unsigned.parts.contains_key("/custom/vbaSignature.bin"));
    assert!(
        unsigned
            .get_part_rels("/custom/vbaProject.bin")
            .unwrap()
            .items
            .iter()
            .all(|relationship| {
                relationship.rel_type != rel_types::VBA_PROJECT_SIGNATURE
                    && !relationship
                        .rel_type
                        .contains("invalidated-vba-project-signature")
            })
    );
}

#[test]
fn external_and_traversal_embedded_targets_fail_closed() {
    for (label, target, external) in [
        ("external", "https://example.com/object.bin", true),
        ("traversal", "../../../../outside.bin", false),
        ("absolute traversal", "/../../outside.bin", false),
    ] {
        let mut package = embedded_fixture_package(false);
        let relationship = package
            .get_or_create_part_rels(SLIDE_TWO_PART)
            .items
            .iter_mut()
            .find(|relationship| relationship.id == "ole-rel")
            .unwrap();
        relationship.target = target.to_owned();
        relationship.target_mode = external.then(|| "External".to_owned());
        let mut presentation = Presentation::from_bytes(&package_bytes(package)).unwrap();
        let before = format!("{presentation:?}");
        assert!(
            presentation.embedded_content().is_err(),
            "{label} inventory"
        );
        assert!(
            presentation
                .extract_embedded_content(SLIDE_TWO_PART, "ole-rel")
                .is_err(),
            "{label} extraction"
        );
        assert!(
            presentation
                .replace_embedded_content(
                    SLIDE_TWO_PART,
                    "ole-rel",
                    b"must-not-land",
                    EmbeddedMutationPolicy::PreserveInvalidatedSignatures,
                )
                .is_err(),
            "{label} replacement"
        );
        assert!(
            presentation
                .remove_embedded_content(
                    SLIDE_TWO_PART,
                    "ole-rel",
                    EmbeddedMutationPolicy::RemoveInvalidatedSignatures,
                )
                .is_err(),
            "{label} removal"
        );
        assert_eq!(format!("{presentation:?}"), before, "{label} atomicity");
    }

    let mut wrong_type = embedded_fixture_package(false);
    wrong_type
        .get_or_create_part_rels(SLIDE_TWO_PART)
        .items
        .iter_mut()
        .find(|relationship| relationship.id == "ole-rel")
        .unwrap()
        .rel_type = rel_types::IMAGE.to_owned();
    let mut wrong_type = Presentation::from_bytes(&package_bytes(wrong_type)).unwrap();
    let before = format!("{wrong_type:?}");
    assert!(wrong_type.embedded_content().is_err());
    assert!(
        wrong_type
            .remove_embedded_content(
                SLIDE_TWO_PART,
                "ole-rel",
                EmbeddedMutationPolicy::RemoveInvalidatedSignatures,
            )
            .is_err()
    );
    assert_eq!(format!("{wrong_type:?}"), before);

    let mut malformed_control = embedded_fixture_package(false);
    malformed_control.set_part(
        "/custom/activeX/activeX1.xml",
        format!(
            r#"<ax:ocx xmlns:ax="http://schemas.microsoft.com/office/2006/activeX" xmlns:r="{R_NS}" r:id="binary-rel"></ax:ocx><ax:ocx xmlns:ax="http://schemas.microsoft.com/office/2006/activeX"/>"#
        )
        .into_bytes(),
    );
    assert!(
        Presentation::from_bytes(&package_bytes(malformed_control))
            .unwrap()
            .embedded_content()
            .is_err()
    );

    let mut missing_signature = embedded_fixture_package(false);
    missing_signature.parts.remove("/custom/vbaSignature.bin");
    assert!(
        Presentation::from_bytes(&package_bytes(missing_signature))
            .unwrap()
            .embedded_content()
            .is_err()
    );

    let mut strict = embedded_fixture_package(false);
    let strict_relationship_namespace = "http://purl.oclc.org/ooxml/officeDocument/relationships";
    let slide = String::from_utf8(strict.get_part(SLIDE_TWO_PART).unwrap().to_vec())
        .unwrap()
        .replace(R_NS, strict_relationship_namespace);
    strict.set_part(SLIDE_TWO_PART, slide.into_bytes());
    let control = String::from_utf8(
        strict
            .get_part("/custom/activeX/activeX1.xml")
            .unwrap()
            .to_vec(),
    )
    .unwrap()
    .replace(R_NS, strict_relationship_namespace);
    strict.set_part("/custom/activeX/activeX1.xml", control.into_bytes());
    for relationship in &mut strict.get_or_create_part_rels(SLIDE_TWO_PART).items {
        relationship.rel_type = match relationship.rel_type.as_str() {
            rel_types::OLE_OBJECT => rel_types::STRICT_OLE_OBJECT.to_owned(),
            rel_types::CONTROL => rel_types::STRICT_CONTROL.to_owned(),
            _ => relationship.rel_type.clone(),
        };
    }
    assert_eq!(
        Presentation::from_bytes(&package_bytes(strict))
            .unwrap()
            .embedded_content()
            .unwrap()
            .len(),
        3
    );
}

#[test]
fn agile_vba_signatures_are_classified_and_follow_both_mutation_policies() {
    let agile_fixture = || {
        let mut package = embedded_fixture_package(true);
        let signature = package
            .get_or_create_part_rels("/custom/vbaProject.bin")
            .items
            .iter_mut()
            .find(|relationship| relationship.id == "vba-signature-rel")
            .unwrap();
        signature.rel_type = rel_types::VBA_PROJECT_SIGNATURE_AGILE.to_owned();
        package.content_types.add_override(
            "/custom/vbaSignature.bin",
            "application/vnd.ms-office.vbaProjectSignatureAgile",
        );
        package_bytes(package)
    };

    let mut preserving = Presentation::from_bytes(&agile_fixture()).unwrap();
    let vba = preserving
        .embedded_content()
        .unwrap()
        .into_iter()
        .find(|info| info.kind == EmbeddedContentKind::VbaProject)
        .unwrap();
    assert_eq!(vba.signature_state, EmbeddedSignatureState::Present);
    let replaced = preserving
        .replace_embedded_content(
            PRESENTATION_PART,
            "vba-rel",
            b"changed agile vba",
            EmbeddedMutationPolicy::PreserveInvalidatedSignatures,
        )
        .unwrap();
    assert_eq!(
        replaced.signature_state,
        EmbeddedSignatureState::Invalidated
    );
    let preserved = open_opc(&preserving.to_bytes().unwrap(), "F-218 agile preserve");
    assert!(preserved.parts.contains_key("/custom/vbaSignature.bin"));
    assert!(
        preserved
            .get_part_rels("/custom/vbaProject.bin")
            .unwrap()
            .items
            .iter()
            .any(|relationship| relationship.rel_type == rel_types::VBA_PROJECT_SIGNATURE_AGILE)
    );

    let mut removing = Presentation::from_bytes(&agile_fixture()).unwrap();
    let replaced = removing
        .replace_embedded_content(
            PRESENTATION_PART,
            "vba-rel",
            b"changed agile vba",
            EmbeddedMutationPolicy::RemoveInvalidatedSignatures,
        )
        .unwrap();
    assert_eq!(replaced.signature_state, EmbeddedSignatureState::Absent);
    assert!(
        !open_opc(&removing.to_bytes().unwrap(), "F-218 agile remove")
            .parts
            .contains_key("/custom/vbaSignature.bin")
    );
}

#[test]
fn malformed_package_signature_graphs_fail_closed_before_mutation() {
    let mut cases = Vec::new();

    let mut external_origin = embedded_fixture_package(true);
    external_origin
        .package_rels
        .items
        .iter_mut()
        .find(|relationship| relationship.id == "package-signature-origin")
        .unwrap()
        .target_mode = Some("External".to_owned());
    cases.push(("external origin", external_origin));

    let mut traversal_origin = embedded_fixture_package(true);
    traversal_origin
        .package_rels
        .items
        .iter_mut()
        .find(|relationship| relationship.id == "package-signature-origin")
        .unwrap()
        .target = "../../outside.sigs".to_owned();
    cases.push(("traversal origin", traversal_origin));

    let mut missing_origin = embedded_fixture_package(true);
    missing_origin.parts.remove("/_xmlsignatures/origin.sigs");
    cases.push(("missing origin part", missing_origin));

    let mut missing_origin_relationships = embedded_fixture_package(true);
    missing_origin_relationships
        .part_rels
        .remove("/_xmlsignatures/origin.sigs");
    cases.push((
        "missing origin relationship set",
        missing_origin_relationships,
    ));

    let mut missing_signature = embedded_fixture_package(true);
    missing_signature.parts.remove("/_xmlsignatures/sig1.xml");
    cases.push(("missing signature part", missing_signature));

    let mut external_signature = embedded_fixture_package(true);
    external_signature
        .get_or_create_part_rels("/_xmlsignatures/origin.sigs")
        .items[0]
        .target_mode = Some("External".to_owned());
    cases.push(("external signature", external_signature));

    let mut traversal_signature = embedded_fixture_package(true);
    traversal_signature
        .get_or_create_part_rels("/_xmlsignatures/origin.sigs")
        .items[0]
        .target = "../../../outside.xml".to_owned();
    cases.push(("traversal signature", traversal_signature));

    let mut unrelated_internal_origin_relationship = embedded_fixture_package(true);
    unrelated_internal_origin_relationship
        .get_or_create_part_rels("/_xmlsignatures/origin.sigs")
        .add_with_id(
            "unrelated-internal",
            rel_types::IMAGE,
            "../custom/embeddings/orphan.bin",
        );
    cases.push((
        "unrelated internal origin relationship",
        unrelated_internal_origin_relationship,
    ));

    let mut unrelated_external_origin_relationship = embedded_fixture_package(true);
    unrelated_external_origin_relationship
        .get_or_create_part_rels("/_xmlsignatures/origin.sigs")
        .add_external(rel_types::HYPERLINK, "https://example.com/unrelated");
    cases.push((
        "unrelated external origin relationship",
        unrelated_external_origin_relationship,
    ));

    let mut duplicate_origin = embedded_fixture_package(true);
    duplicate_origin.set_part("/custom/second-origin.sigs", Vec::new());
    duplicate_origin.set_part(
        "/custom/second-signature.xml",
        b"second signature evidence".to_vec(),
    );
    duplicate_origin.content_types.add_override(
        "/custom/second-origin.sigs",
        "application/vnd.openxmlformats-package.digital-signature-origin",
    );
    duplicate_origin.content_types.add_override(
        "/custom/second-signature.xml",
        "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml",
    );
    duplicate_origin.package_rels.add_with_id(
        "second-origin",
        rel_types::DIGITAL_SIGNATURE_ORIGIN,
        "custom/second-origin.sigs",
    );
    duplicate_origin
        .get_or_create_part_rels("/custom/second-origin.sigs")
        .add_with_id(
            "second-signature",
            rel_types::DIGITAL_SIGNATURE,
            "second-signature.xml",
        );
    cases.push(("duplicate origin", duplicate_origin));

    let mut root_signature = embedded_fixture_package(true);
    root_signature.package_rels.add_with_id(
        "misplaced-root-signature",
        rel_types::DIGITAL_SIGNATURE,
        "_xmlsignatures/sig1.xml",
    );
    cases.push(("root signature relationship", root_signature));

    let mut misplaced_signature = embedded_fixture_package(true);
    misplaced_signature
        .get_or_create_part_rels(SLIDE_TWO_PART)
        .add_with_id(
            "misplaced-signature",
            rel_types::DIGITAL_SIGNATURE,
            "../../_xmlsignatures/sig1.xml",
        );
    cases.push(("ordinary-part signature relationship", misplaced_signature));

    let mut misplaced_origin = embedded_fixture_package(true);
    misplaced_origin
        .get_or_create_part_rels(SLIDE_TWO_PART)
        .add_with_id(
            "misplaced-origin",
            rel_types::DIGITAL_SIGNATURE_ORIGIN,
            "../../_xmlsignatures/origin.sigs",
        );
    cases.push(("ordinary-part origin relationship", misplaced_origin));

    for (label, package) in cases {
        let mut presentation = Presentation::from_bytes(&package_bytes(package)).unwrap();
        let before = format!("{presentation:?}");
        assert!(
            presentation.embedded_content().is_err(),
            "{label} inventory"
        );
        assert!(
            presentation
                .replace_embedded_content(
                    SLIDE_TWO_PART,
                    "ole-rel",
                    b"must-not-land",
                    EmbeddedMutationPolicy::RemoveInvalidatedSignatures,
                )
                .is_err(),
            "{label} replacement"
        );
        assert_eq!(format!("{presentation:?}"), before, "{label} atomicity");
    }
}

#[test]
fn duplicate_relationship_ids_and_semantic_attributes_fail_closed() {
    let mut duplicate_relationship = embedded_fixture_package(false);
    duplicate_relationship
        .get_or_create_part_rels(SLIDE_TWO_PART)
        .items
        .push(oxml_opc::relationship::Relationship {
            id: "ole-rel".to_owned(),
            rel_type: rel_types::IMAGE.to_owned(),
            target: "https://example.com/ambiguous".to_owned(),
            target_mode: Some("External".to_owned()),
        });
    let mut presentation =
        Presentation::from_bytes(&package_bytes(duplicate_relationship)).unwrap();
    let before = presentation.to_bytes().unwrap();
    assert!(presentation.embedded_content().is_err());
    assert!(
        presentation
            .remove_embedded_content(
                SLIDE_TWO_PART,
                "ole-rel",
                EmbeddedMutationPolicy::RemoveInvalidatedSignatures,
            )
            .is_err()
    );
    assert_eq!(presentation.to_bytes().unwrap(), before);

    let mut duplicate_vba_signature = embedded_fixture_package(false);
    duplicate_vba_signature
        .get_or_create_part_rels("/custom/vbaProject.bin")
        .items
        .push(oxml_opc::relationship::Relationship {
            id: "vba-signature-rel".to_owned(),
            rel_type: rel_types::IMAGE.to_owned(),
            target: "https://example.com/ambiguous-signature-id".to_owned(),
            target_mode: Some("External".to_owned()),
        });
    let mut presentation =
        Presentation::from_bytes(&package_bytes(duplicate_vba_signature)).unwrap();
    let before = presentation.to_bytes().unwrap();
    assert!(presentation.embedded_content().is_err());
    assert!(
        presentation
            .replace_embedded_content(
                PRESENTATION_PART,
                "vba-rel",
                b"must-not-land",
                EmbeddedMutationPolicy::RemoveInvalidatedSignatures,
            )
            .is_err()
    );
    assert_eq!(presentation.to_bytes().unwrap(), before);

    for (label, part, needle, duplicate_namespace) in [
        (
            "OLE",
            SLIDE_TWO_PART,
            r#"r:id="ole-rel""#,
            "http://purl.oclc.org/ooxml/officeDocument/relationships",
        ),
        ("control", SLIDE_TWO_PART, r#"q:id="control-rel""#, R_NS),
        (
            "ActiveX",
            "/custom/activeX/activeX1.xml",
            r#"q:id="binary-rel""#,
            R_NS,
        ),
    ] {
        let mut package = embedded_fixture_package(false);
        let xml = String::from_utf8(package.get_part(part).unwrap().to_vec()).unwrap();
        let duplicate = format!(r#"{needle} xmlns:dup="{duplicate_namespace}" dup:id="ambiguous""#);
        assert!(xml.contains(needle), "{label} fixture needle");
        package.set_part(part, xml.replacen(needle, &duplicate, 1).into_bytes());
        let presentation = Presentation::from_bytes(&package_bytes(package)).unwrap();
        assert!(
            presentation.embedded_content().is_err(),
            "{label} duplicate attribute"
        );
    }
}

#[test]
fn nested_same_namespace_ole_and_control_lookalikes_never_acquire_ownership() {
    let mut package = embedded_fixture_package(false);
    let slide = String::from_utf8(package.get_part(SLIDE_TWO_PART).unwrap().to_vec()).unwrap();
    let ole = ole_frame_xml(901, "ole-rel");
    let nested_ole = ole
        .replacen(
            r#"<a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole">"#,
            r#"<a:graphicData uri="urn:f218:unrelated"><x:opaque xmlns:x="urn:f218:nested"><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole">"#,
            1,
        )
        .replacen(
            "</a:graphicData></a:graphic>",
            "</a:graphicData></x:opaque></a:graphicData></a:graphic>",
            1,
        );
    let slide = slide
        .replacen(&ole, &nested_ole, 1)
        .replacen(
            r#"<p:controls xmlns:q="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#,
            r#"<p:extLst><p:ext uri="{F218-NESTED-CONTROLS}"><p:controls xmlns:q="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#,
            1,
        )
        .replacen("</p:control></p:controls>", "</p:control></p:controls></p:ext></p:extLst>", 1);
    package.set_part(SLIDE_TWO_PART, slide.into_bytes());
    let mut presentation = Presentation::from_bytes(&package_bytes(package)).unwrap();
    let inventory = presentation.embedded_content().unwrap();
    assert_eq!(inventory.len(), 1);
    assert_eq!(inventory[0].kind, EmbeddedContentKind::VbaProject);
    let before = presentation.to_bytes().unwrap();
    for (source, relationship) in [
        (SLIDE_TWO_PART, "ole-rel"),
        ("/custom/activeX/activeX1.xml", "binary-rel"),
    ] {
        assert!(
            presentation
                .remove_embedded_content(
                    source,
                    relationship,
                    EmbeddedMutationPolicy::RemoveInvalidatedSignatures,
                )
                .is_err()
        );
        assert_eq!(presentation.to_bytes().unwrap(), before);
    }
}

fn embedded_fixture_package(with_signatures: bool) -> OpcPackage {
    let mut package = fixture_package();
    let slide = String::from_utf8(package.get_part(SLIDE_TWO_PART).unwrap().to_vec()).unwrap();
    let slide = slide
        .replacen(
            "</p:spTree>",
            &format!("{}</p:spTree>", ole_frame_xml(901, "ole-rel")),
            1,
        )
        .replacen(
            "</p:cSld>",
            &format!(
                r#"<p:controls xmlns:q="{R_NS}" xmlns:x="urn:f218"><p:control q:id="control-rel" name="opaque"><p:extLst><x:producer marker="kept"/></p:extLst></p:control></p:controls></p:cSld>"#
            ),
            1,
        );
    package.set_part(SLIDE_TWO_PART, slide.into_bytes());
    package.set_part("/custom/embeddings/object1.bin", b"ole-executable".to_vec());
    package.set_part("/custom/embeddings/orphan.bin", b"producer-orphan".to_vec());
    package.set_part(
        "/custom/activeX/activeX1.xml",
        format!(
            r#"<?xml version="1.0"?><ax:ocx xmlns:ax="http://schemas.microsoft.com/office/2006/activeX" xmlns:q="{R_NS}" q:id="binary-rel" classid="opaque" persistence="persistStream"><ax:ocxPr ax:name="producer" ax:value="kept"/></ax:ocx>"#
        )
        .into_bytes(),
    );
    package.set_part(
        "/custom/activeX/activeX1.bin",
        b"activex-executable".to_vec(),
    );
    package.set_part("/custom/vbaProject.bin", b"vba-executable".to_vec());
    package.set_part(
        "/custom/vbaSignature.bin",
        b"vba-signature-evidence".to_vec(),
    );
    for (part, content_type) in [
        (
            "/custom/embeddings/object1.bin",
            "application/vnd.openxmlformats-officedocument.oleObject",
        ),
        (
            "/custom/embeddings/orphan.bin",
            "application/vnd.openxmlformats-officedocument.oleObject",
        ),
        (
            "/custom/activeX/activeX1.xml",
            "application/vnd.ms-office.activeX+xml",
        ),
        (
            "/custom/activeX/activeX1.bin",
            "application/vnd.ms-office.activeX",
        ),
        (
            "/custom/vbaProject.bin",
            "application/vnd.ms-office.vbaProject",
        ),
        (
            "/custom/vbaSignature.bin",
            "application/vnd.ms-office.vbaProjectSignature",
        ),
    ] {
        package.content_types.add_override(part, content_type);
    }
    package.get_or_create_part_rels(SLIDE_TWO_PART).add_with_id(
        "ole-rel",
        rel_types::OLE_OBJECT,
        "../embeddings/object1.bin",
    );
    package.get_or_create_part_rels(SLIDE_TWO_PART).add_with_id(
        "control-rel",
        rel_types::CONTROL,
        "../activeX/activeX1.xml",
    );
    package
        .get_or_create_part_rels("/custom/activeX/activeX1.xml")
        .add_with_id(
            "binary-rel",
            rel_types::ACTIVEX_CONTROL_BINARY,
            "activeX1.bin",
        );
    package
        .get_or_create_part_rels(PRESENTATION_PART)
        .add_with_id(
            "vba-rel",
            rel_types::VBA_PROJECT,
            "../custom/vbaProject.bin",
        );
    package
        .get_or_create_part_rels("/custom/vbaProject.bin")
        .add_with_id(
            "vba-signature-rel",
            rel_types::VBA_PROJECT_SIGNATURE,
            "vbaSignature.bin",
        );
    if with_signatures {
        package.set_part("/_xmlsignatures/origin.sigs", Vec::new());
        package.set_part(
            "/_xmlsignatures/sig1.xml",
            b"package-signature-evidence".to_vec(),
        );
        package.content_types.add_override(
            "/_xmlsignatures/origin.sigs",
            "application/vnd.openxmlformats-package.digital-signature-origin",
        );
        package.content_types.add_override(
            "/_xmlsignatures/sig1.xml",
            "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml",
        );
        package.package_rels.add_with_id(
            "package-signature-origin",
            rel_types::DIGITAL_SIGNATURE_ORIGIN,
            "_xmlsignatures/origin.sigs",
        );
        package
            .get_or_create_part_rels("/_xmlsignatures/origin.sigs")
            .add_with_id(
                "package-signature",
                rel_types::DIGITAL_SIGNATURE,
                "sig1.xml",
            );
    }
    package
}

fn embedded_layout_and_master_fixture_package() -> OpcPackage {
    let mut package = embedded_fixture_package(false);
    let master_part = "/custom/masters/embedded.xml";
    let presentation =
        String::from_utf8(package.get_part(PRESENTATION_PART).unwrap().to_vec()).unwrap();
    package.set_part(
        PRESENTATION_PART,
        presentation
            .replacen(
                "<p:sldIdLst>",
                r#"<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="embedded-master"/></p:sldMasterIdLst><p:sldIdLst>"#,
                1,
            )
            .into_bytes(),
    );
    package.set_part(
        LAYOUT_PART,
        embedded_ole_scope_xml("sldLayout", 911, "layout-ole", ""),
    );
    package.set_part(
        master_part,
        embedded_ole_scope_xml(
            "sldMaster",
            912,
            "master-ole",
            r#"<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst/><p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles>"#,
        ),
    );
    package.set_part(
        "/custom/embeddings/layout.bin",
        b"layout-owned-ole".to_vec(),
    );
    package.set_part(
        "/custom/embeddings/master.bin",
        b"master-owned-ole".to_vec(),
    );
    package.set_part(
        "/custom/theme/theme1.xml",
        format!(
            r#"<a:theme xmlns:a="{A_NS}" name="F-218"><a:themeElements><a:clrScheme name="F-218"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1></a:clrScheme><a:fontScheme name="F-218"><a:majorFont/><a:minorFont/></a:fontScheme><a:fmtScheme name="F-218"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme></a:themeElements></a:theme>"#
        )
        .into_bytes(),
    );
    package
        .content_types
        .add_override(master_part, content_types::SLIDE_MASTER);
    package
        .content_types
        .add_override("/custom/theme/theme1.xml", content_types::THEME);
    for part in [
        "/custom/embeddings/layout.bin",
        "/custom/embeddings/master.bin",
    ] {
        package.content_types.add_override(
            part,
            "application/vnd.openxmlformats-officedocument.oleObject",
        );
    }
    package
        .get_or_create_part_rels(PRESENTATION_PART)
        .add_with_id(
            "embedded-master",
            rel_types::SLIDE_MASTER,
            "masters/embedded.xml",
        );
    package
        .get_or_create_part_rels(master_part)
        .add(rel_types::SLIDE_LAYOUT, "../layouts/validation.xml");
    package
        .get_or_create_part_rels(master_part)
        .add(rel_types::THEME, "../theme/theme1.xml");
    package.get_or_create_part_rels(master_part).add_with_id(
        "master-ole",
        rel_types::OLE_OBJECT,
        "../embeddings/master.bin",
    );
    package.get_or_create_part_rels(LAYOUT_PART).add_with_id(
        "layout-ole",
        rel_types::OLE_OBJECT,
        "../embeddings/layout.bin",
    );
    package
}

fn embedded_ole_scope_xml(root: &str, shape_id: u32, relationship_id: &str, tail: &str) -> Vec<u8> {
    format!(
        r#"<p:{root} xmlns:p="{P_NS}" xmlns:a="{A_NS}" xmlns:r="{R_NS}" xmlns:x="urn:f218"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/><x:producer-boundary scope="{root}"/>{}</p:spTree></p:cSld>{tail}</p:{root}>"#,
        ole_frame_xml(shape_id, relationship_id)
    )
    .into_bytes()
}

fn ole_frame_xml(shape_id: u32, relationship_id: &str) -> String {
    format!(
        r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="{shape_id}" name="OLE {shape_id}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole"><p:oleObj r:id="{relationship_id}" name="opaque"><p:embed/><x:producer xmlns:x="urn:f218" marker="kept"/></p:oleObj></a:graphicData></a:graphic></p:graphicFrame>"#
    )
}

fn hex32(bytes: [u8; 32]) -> String {
    bytes.iter().map(|byte| format!("{byte:02x}")).collect()
}

#[test]
fn slide_transitions_and_morph_metadata_survive_mutation_save_and_reopen() {
    use rpptx_oxml::slide_parts::CT_Slide;
    use rpptx_oxml::timing::{TransitionEffect, TransitionSpeed};

    let slide_xml = format!(
        r#"<p:sld xmlns:p="{P_NS}" xmlns:a="{A_NS}" xmlns:mc="{MC_NS}" xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:f213" mc:Ignorable="p14 p159"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld><x:before/><mc:AlternateContent x:wrapper='A&#x20;B'><mc:Choice Requires="p159"><p:transition x:owner='C&#x20;D' spd='slow' p14:dur="725" advClick="1" advTm="4000"><x:morph option="foreign"/><p159:morph r:id='rId&#x37;' option='byObject'><x:keep/></p159:morph></p:transition></mc:Choice><mc:Fallback><p:transition spd="slow"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent><p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite"/></p:par></p:tnLst></p:timing><x:after/></p:sld>"#
    );
    let mut slide = CT_Slide::from_xml(slide_xml.as_bytes()).unwrap();
    let transition = slide.transition.as_mut().unwrap();
    assert_eq!(transition.effect, Some(TransitionEffect::Morph));
    assert_eq!(transition.duration_ms, Some(725));
    transition.set_speed(TransitionSpeed::Fast).unwrap();
    transition.set_morph_option("byWord").unwrap();
    let expected_slide = slide.to_xml().unwrap();

    let mut package = fixture_package();
    package.set_part(SLIDE_ONE_PART, expected_slide.clone());
    let reopened_package = OpcPackage::from_reader(Cursor::new(package_bytes(package))).unwrap();
    let reopened = CT_Slide::from_xml(reopened_package.get_part(SLIDE_ONE_PART).unwrap()).unwrap();
    let transition = reopened.transition.unwrap();
    assert_eq!(transition.speed, Some(TransitionSpeed::Fast));
    assert_eq!(transition.effect, Some(TransitionEffect::Morph));
    assert_eq!(transition.duration_ms, Some(725));
    assert_eq!(transition.morph.unwrap().option.as_deref(), Some("byWord"));
    let written = String::from_utf8(expected_slide).unwrap();
    assert!(written.contains(r#"<x:morph option="foreign"/>"#));
    assert!(written.contains("<mc:AlternateContent x:wrapper='A&#x20;B'>"));
    assert!(written.contains("<p:transition x:owner='C&#x20;D' spd='fast'"));
    assert!(written.contains("<p159:morph r:id='rId&#x37;' option='byWord'>"));
    assert!(written.contains(
        r#"<mc:Fallback><p:transition spd="slow"><p:fade/></p:transition></mc:Fallback>"#
    ));
    assert!(written.contains("<x:keep/>"));
    assert!(written.contains("<x:before/>"));
    assert!(written.contains("<x:after/>"));
}

#[test]
fn supported_smartart_nodes_remain_editable_after_save_and_reopen() {
    let package = smartart_fixture_package();
    let unsupported_parts = [
        "/custom/diagrams/two/layout1.xml",
        "/custom/diagrams/two/style1.xml",
        "/custom/diagrams/two/colors1.xml",
        "/custom/diagrams/two/drawing1.xml",
    ]
    .map(|part| (part, package.get_part(part).unwrap().to_vec()));
    let mut presentation = Presentation::from_bytes(&package_bytes(package)).unwrap();
    presentation
        .set_smart_art_node_text(0, 42, "n1", "Edited through facade")
        .unwrap();
    let bytes = presentation.to_bytes().unwrap();
    let reopened = Presentation::from_bytes(&bytes).unwrap();
    let info = reopened
        .smart_art(0)
        .unwrap()
        .into_iter()
        .find(|info| info.shape_id == 42)
        .unwrap();
    let rpptx::DiagramPart::Parsed(data) = info.data else {
        panic!("edited data part did not reparse");
    };
    assert_eq!(
        data.points()[0].text.as_ref().unwrap().plain_text(),
        "Edited through facade"
    );
    assert_eq!(data.connections().len(), 0);
    assert!(matches!(info.layout, rpptx::DiagramPart::Parsed(_)));
    assert!(matches!(info.style, rpptx::DiagramPart::Parsed(_)));
    assert!(matches!(info.colors, rpptx::DiagramPart::Parsed(_)));
    assert!(matches!(info.drawing, Some(rpptx::DiagramPart::Parsed(_))));
    let saved = open_opc(&bytes, "F-219 edited SmartArt");
    for (part, expected) in unsupported_parts {
        assert_eq!(saved.get_part(part).unwrap(), expected);
    }
}

#[test]
fn unsupported_smartart_algorithms_and_parts_remain_byte_preserved_after_unrelated_mutation() {
    let package = smartart_fixture_package();
    let before = [
        "/custom/diagrams/two/data1.xml",
        "/custom/diagrams/two/layout1.xml",
        "/custom/diagrams/two/style1.xml",
        "/custom/diagrams/two/colors1.xml",
        "/custom/diagrams/two/drawing1.xml",
    ]
    .map(|part| (part, package.get_part(part).unwrap().to_vec()));
    let mut presentation = Presentation::from_bytes(&package_bytes(package)).unwrap();
    presentation
        .slide_mut(0)
        .unwrap()
        .shape_mut(0)
        .unwrap()
        .set_name("ordinary edit")
        .unwrap();
    let saved = open_opc(
        &presentation.to_bytes().unwrap(),
        "F-219 unrelated mutation",
    );
    for (part, expected) in before {
        assert_eq!(saved.get_part(part).unwrap(), expected);
    }
    let info = presentation
        .smart_art(0)
        .unwrap()
        .into_iter()
        .find(|info| info.shape_id == 42)
        .unwrap();
    let rpptx::DiagramPart::Parsed(layout) = info.layout else {
        panic!("layout did not parse");
    };
    assert!(matches!(
        layout.family,
        rpptx::DiagramLayoutFamily::Unsupported(_)
    ));
}

#[test]
fn smartart_rendering_uses_producing_scope_and_updates_after_node_edit() {
    if !pinned_smartart_resources_available() {
        eprintln!(
            "authentic SmartArt rendering skipped because pinned PowerPoint resources are absent"
        );
        return;
    }
    let mut package = smartart_fixture_package();
    make_smartart_renderable(&mut package);
    wrap_slide_smartart_in_scaled_group(&mut package);
    package.set_part(
        "/custom/diagrams/two/layout1.xml",
        read_pinned_smartart_resource(
            smartart_layout_resource("list").0,
            smartart_layout_resource("list").1,
        ),
    );
    package.set_part(
        "/custom/diagrams/two/colors1.xml",
        read_pinned_smartart_resource(SMARTART_COLOR_RESOURCE.0, SMARTART_COLOR_RESOURCE.1),
    );
    let mut presentation = Presentation::from_bytes(&package_bytes(package)).unwrap();
    let (before, before_layout) = presentation.render_deterministic().unwrap();
    let smartart_shape_index = before.slides[0]
        .shapes
        .iter()
        .position(|shape| resolved_content_text(&shape.content) == "Two")
        .expect("supported SmartArt ordinary shape");
    let smartart_shape = &before.slides[0].shapes[smartart_shape_index];
    let ResolvedContent::Text(body) = &smartart_shape.content else {
        panic!("SmartArt node did not enter the shared text engine")
    };
    assert_eq!(body.anchor, rpptx_layout::TextAnchor::Center);
    assert_eq!(body.insets.left, 15.3);
    let ResolvedTextRun::Text { style, .. } = &body.paragraphs[0].runs[0] else {
        panic!("SmartArt run was not preserved")
    };
    assert!(style.bold && style.italic);
    assert_eq!(style.font_size, Some(34.0));
    let Some(Paint::Solid(fill)) = smartart_shape.fill.as_ref() else {
        panic!("SmartArt colour did not use the shared DrawingML fill resolver")
    };
    assert!((fill.r - 1.0).abs() < 1.0e-9, "{fill:?}");
    assert!((fill.g - 1.0).abs() < 1.0e-9, "{fill:?}");
    assert!((fill.b - 1.0).abs() < 1.0e-9, "{fill:?}");
    assert!((fill.a - 1.0).abs() < 1.0e-9, "{fill:?}");
    assert_smartart_frame_clip(
        &before_layout.pages[0],
        &before.slides[0],
        smartart_shape_index,
        Rect {
            x: 20.0,
            y: 40.0,
            width: 400.0,
            height: 200.0,
        },
    );
    assert!(
        before.slides[0]
            .shapes
            .iter()
            .any(|shape| resolved_content_text(&shape.content) == "Two"),
        "supported producing-scope SmartArt still resolves to a fallback: {:?}",
        before.slides[0].diagnostics
    );
    let animation_segment = [rpptx::AnimationSegment {
        slide_index: 0,
        duration_ms: 100,
        click_count: 0,
        transition: rpptx::AnimationTransition::None,
    }];
    let animation_options = rpptx::AnimationExportOptions {
        frame_rate: 10,
        width_px: 96,
        height_px: 54,
        format: rpptx::AnimationFormat::Gif {
            loop_behavior: rpptx::GifLoopBehavior::Once,
        },
        media_fallback: MediaFallbackPolicy::PosterFrame,
    };
    let animation_before = presentation
        .export_animation_deterministic(&animation_segment, animation_options)
        .unwrap();

    presentation
        .set_smart_art_node_text(0, 42, "n1", "Edited through render path")
        .unwrap();
    let (after, _) = presentation.render_deterministic().unwrap();
    assert!(
        after.slides[0]
            .shapes
            .iter()
            .any(|shape| { resolved_content_text(&shape.content) == "Edited through render path" })
    );
    let timeline = presentation
        .render_timeline_deterministic(0, TimelinePosition::default(), None)
        .unwrap();
    let edited_shape_index = after.slides[0]
        .shapes
        .iter()
        .position(|shape| resolved_content_text(&shape.content) == "Edited through render path")
        .unwrap();
    let scaled_frame = Rect {
        x: 20.0,
        y: 40.0,
        width: 400.0,
        height: 200.0,
    };
    assert_smartart_frame_clip(
        &timeline.page,
        &after.slides[0],
        edited_shape_index,
        scaled_frame,
    );
    assert!(
        page_text(&timeline.page)
            .concat()
            .contains("Edited through render path"),
        "timeline text {:?}, diagnostics {:?}",
        page_text(&timeline.page),
        timeline.diagnostics
    );
    let media = presentation
        .render_media_timeline_deterministic(
            0,
            TimelinePosition::default(),
            None,
            MediaFallbackPolicy::PosterFrame,
        )
        .unwrap();
    assert!(
        page_text(&media.frame.page)
            .concat()
            .contains("Edited through render path")
    );
    assert_smartart_frame_clip(
        &media.frame.page,
        &after.slides[0],
        edited_shape_index,
        scaled_frame,
    );
    let animation_after = presentation
        .export_animation_deterministic(&animation_segment, animation_options)
        .unwrap();
    assert!(animation_after.bytes.starts_with(b"GIF89a"));
    assert_ne!(
        animation_before.bytes, animation_after.bytes,
        "the animation path must consume the same edited SmartArt expansion"
    );
    assert!(
        animation_after
            .diagnostics
            .iter()
            .all(|diagnostic| !diagnostic.message.contains("unsupported SmartArt")),
        "animation diagnostics {:?}",
        animation_after.diagnostics
    );
}

fn assert_smartart_frame_clip(
    page: &PageFrame,
    slide: &ResolvedSlide,
    shape_index: usize,
    frame: Rect,
) {
    let shape_offset = page.elements.len() - slide.shapes.len();
    let PositionedElement::Group(group) = &page.elements[shape_offset + shape_index] else {
        panic!("SmartArt resolved shape did not lower to a group")
    };
    let clip = group.clip.as_ref().expect("authoritative SmartArt clip");
    let actual = clip.bounds().expect("rectangular SmartArt clip bounds");
    let shape = &slide.shapes[shape_index];
    let expected = Rect {
        x: frame.x - shape.bounds.x,
        y: frame.y - shape.bounds.y,
        width: frame.width,
        height: frame.height,
    };
    for (actual, expected) in [
        (actual.x, expected.x),
        (actual.y, expected.y),
        (actual.width, expected.width),
        (actual.height, expected.height),
    ] {
        assert!((actual - expected).abs() < 1.0e-9, "{actual} != {expected}");
    }
}

#[test]
fn supported_native_smartart_ignores_a_malformed_optional_cached_drawing() {
    if !pinned_smartart_resources_available() {
        eprintln!(
            "authentic SmartArt cached-drawing regression skipped because pinned PowerPoint resources are absent"
        );
        return;
    }
    let mut package = smartart_fixture_package();
    make_smartart_renderable(&mut package);
    package.set_part(
        "/custom/diagrams/two/layout1.xml",
        read_pinned_smartart_resource(
            smartart_layout_resource("list").0,
            smartart_layout_resource("list").1,
        ),
    );
    package.set_part(
        "/custom/diagrams/two/drawing1.xml",
        b"<malformed optional cache".to_vec(),
    );
    let presentation = Presentation::from_bytes(&package_bytes(package)).unwrap();
    let (resolved, _) = presentation.render_deterministic().unwrap();
    assert!(
        resolved.slides[0]
            .shapes
            .iter()
            .any(|shape| resolved_content_text(&shape.content) == "Two"),
        "native SmartArt layout must not depend on its optional cached drawing: {:?}",
        resolved.slides[0].diagnostics
    );
}

#[test]
fn supported_smartart_corpus_matches_pinned_powerpoint_geometry_and_ssim() {
    let oracle_dir = std::env::var_os("RDOCX_PPTX_SMARTART_ORACLE_DIR")
        .map(PathBuf::from)
        .unwrap_or_else(|| workspace_root().join("corpus/pptx-smartart-oracle"));
    let expected_families = smartart_cases()
        .iter()
        .map(|case| case.family)
        .collect::<HashSet<_>>();
    assert_eq!(expected_families.len(), 6);
    let manifest_path = oracle_dir.join("f220-smartart-oracle.tsv");
    let manifest = match fs::read_to_string(&manifest_path) {
        Ok(manifest) => manifest,
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => {
            assert!(
                std::env::var_os("RDOCX_PPTX_CORPUS_REQUIRED").is_none(),
                "required PowerPoint SmartArt oracle manifest is missing at {}",
                manifest_path.display()
            );
            eprintln!(
                "PowerPoint SmartArt oracle skipped because {} is absent",
                manifest_path.display()
            );
            return;
        }
        Err(error) => panic!("read {}: {error}", manifest_path.display()),
    };
    let rows = smartart_oracle_rows(&manifest);
    assert_eq!(rows.len(), 6);
    let artifacts_present = rows.iter().all(|row| {
        [&row.source, &row.oracle_pdf, &row.oracle_png]
            .into_iter()
            .all(|artifact| oracle_dir.join(artifact).is_file())
    });
    if !artifacts_present {
        assert!(
            std::env::var_os("RDOCX_PPTX_CORPUS_REQUIRED").is_none(),
            "required PowerPoint SmartArt oracle artifacts are missing at {}",
            oracle_dir.display()
        );
    }
    let mut families = HashSet::new();
    let mut divergences = Vec::new();
    for (index, row) in rows.iter().enumerate() {
        assert!(families.insert(row.family.as_str()));
        let case = smartart_cases()
            .iter()
            .find(|case| case.family == row.family)
            .expect("known SmartArt oracle family");
        assert_eq!(row.case, case.case);
        assert_eq!(row.source, case.source);
        assert_eq!(row.slide, case.slide);
        assert_eq!(row.node_texts, case.node_texts);
        assert_eq!(row.oracle_pdf, format!("{}-powerpoint.pdf", case.case));
        assert_eq!(row.oracle_png, format!("{}-powerpoint.png", case.case));
        if !artifacts_present {
            continue;
        }
        let source = oracle_dir.join(&row.source);
        let pdf = oracle_dir.join(&row.oracle_pdf);
        let oracle = oracle_dir.join(&row.oracle_png);
        assert_eq!(sha256(&source), row.source_sha256);
        assert_eq!(sha256(&pdf), row.oracle_pdf_sha256);
        assert_eq!(sha256(&oracle), row.oracle_sha256);
        let generated = fs::read(&source).unwrap();
        let presentation = Presentation::from_bytes(&generated).unwrap();
        let (input, layout) = presentation.render_deterministic().unwrap();
        let slide_index = row.slide.checked_sub(1).expect("one-based SmartArt slide");
        let slide = input
            .slides
            .get(slide_index)
            .expect("manifest SmartArt slide");
        let actual_ownership = smartart_shape_ownership(slide);
        assert_eq!(
            actual_ownership, row.shape_ownership,
            "{} ownership",
            row.case
        );
        let actual_bounds = slide
            .shapes
            .iter()
            .map(|shape| shape.bounds)
            .collect::<Vec<_>>();
        let actual_diagnostics = if slide.diagnostics.is_empty() {
            "-".to_owned()
        } else {
            slide
                .diagnostics
                .iter()
                .map(|diagnostic| diagnostic.message.as_str())
                .collect::<Vec<_>>()
                .join("|")
        };
        if actual_diagnostics != row.expected_diagnostics {
            divergences.push(format!(
                "{} diagnostics: actual `{actual_diagnostics}`, expected `{}`",
                row.case, row.expected_diagnostics
            ));
        }
        let rust_png = oxml_pdf::render_page_to_png(&layout, slide_index, 150.0).unwrap();
        let rust_pixmap = tiny_skia::Pixmap::decode_png(&rust_png).unwrap();
        assert_eq!(
            (rust_pixmap.width(), rust_pixmap.height()),
            (row.width, row.height),
            "{} Rust dimensions",
            row.case
        );
        let normalized_oracle =
            normalize_smartart_oracle_png(&oracle, row.width, row.height, index);
        let oracle_pixmap = tiny_skia::Pixmap::decode_png(&normalized_oracle).unwrap();
        assert_eq!(
            (oracle_pixmap.width(), oracle_pixmap.height()),
            (row.width, row.height),
            "{} normalized PowerPoint dimensions",
            row.case
        );
        let expected_bounds = parse_smartart_bounds(&row.shape_bounds_pt);
        let comparison = compare_smartart_geometry_and_png(
            &actual_bounds,
            &expected_bounds,
            &rust_png,
            &normalized_oracle,
        );
        let text_owner_bounds = expected_bounds
            .iter()
            .zip(row.shape_ownership.split('|'))
            .filter_map(|(bounds, owner)| owner.starts_with("text:").then_some(*bounds))
            .collect::<Vec<_>>();
        let raster = smartart_raster_comparison(&rust_png, &normalized_oracle, &text_owner_bounds);
        eprintln!(
            "{}: geometry_error_pt={:.6}, non_text_ssim={:.9}, full_ssim={:.9}, text_edge_error_pt={:.6}, text_width_error_pt={:.6}",
            row.case,
            comparison.geometry_error,
            raster.non_text_ssim,
            raster.full_ssim,
            raster.max_ink_edge_error_pt,
            raster.max_line_width_error_pt
        );
        if comparison.geometry_error > 1.0 {
            divergences.push(format!(
                "{} geometry error {} pt exceeds 1 pt",
                row.case, comparison.geometry_error
            ));
        }
        if raster.non_text_ssim < 0.90 {
            divergences.push(format!(
                "{} symmetric text-masked non-text SSIM {} is below 0.90",
                row.case, raster.non_text_ssim
            ));
        }
        if raster.actual_line_counts != row.text_line_counts
            || raster.expected_line_counts != row.text_line_counts
        {
            divergences.push(format!(
                "{} ordered text line counts differ: Rust {:?}, PowerPoint {:?}, expected {:?}",
                row.case,
                raster.actual_line_counts,
                raster.expected_line_counts,
                row.text_line_counts
            ));
        }
        if raster.max_ink_edge_error_pt > 3.0 {
            divergences.push(format!(
                "{} owner-centered text ink edge error {} pt exceeds 3 pt",
                row.case, raster.max_ink_edge_error_pt
            ));
        }
        if raster.max_line_width_error_pt > 3.0 {
            divergences.push(format!(
                "{} text line ink width error {} pt exceeds 3 pt",
                row.case, raster.max_line_width_error_pt
            ));
        }
    }
    assert_eq!(
        families,
        [
            "list",
            "hierarchy",
            "cycle",
            "relationship",
            "matrix",
            "pyramid"
        ]
        .into_iter()
        .collect()
    );
    if !artifacts_present {
        eprintln!(
            "PowerPoint SmartArt oracle skipped because an artifact is absent at {}",
            oracle_dir.display()
        );
    }
    assert!(
        divergences.is_empty(),
        "PowerPoint SmartArt divergences:\n{}",
        divergences.join("\n")
    );
}

#[derive(Clone, Debug)]
struct SmartartOracleRow {
    case: String,
    family: String,
    source: String,
    source_sha256: String,
    slide: usize,
    oracle_pdf: String,
    oracle_pdf_sha256: String,
    oracle_png: String,
    oracle_sha256: String,
    width: u32,
    height: u32,
    node_texts: String,
    shape_ownership: String,
    shape_bounds_pt: String,
    text_line_counts: Vec<usize>,
    expected_diagnostics: String,
}

fn smartart_oracle_rows(manifest: &str) -> Vec<SmartartOracleRow> {
    let lines = manifest.lines().collect::<Vec<_>>();
    for expected in [
        "application\tMicrosoft PowerPoint",
        &format!("version\t{POWERPOINT_VERSION}"),
        &format!("bundle_build\t{POWERPOINT_BUILD}"),
        &format!("app_build\t{POWERPOINT_APP_BUILD}"),
        "dpi\t150",
        "geometry_tolerance_pt\t1",
        "non_text_ssim_threshold\t0.90",
        "text_ink_tolerance_pt\t3",
        "full_image_ssim\tdiagnostic_only",
    ] {
        assert!(
            lines.contains(&expected),
            "missing external SmartArt pin `{expected}`"
        );
    }
    let header = "case\tfamily\tsource\tsource_sha256\tslide\toracle_pdf\toracle_pdf_sha256\toracle_png\toracle_sha256\twidth\theight\tnode_texts\tshape_ownership\tshape_bounds_pt\ttext_line_counts\texpected_diagnostics\tclassification";
    let start = lines
        .iter()
        .position(|line| *line == header)
        .expect("external SmartArt header");
    lines[start + 1..]
        .iter()
        .filter(|line| !line.is_empty())
        .map(|line| {
            let fields = line.split('\t').collect::<Vec<_>>();
            assert_eq!(fields.len(), 17, "external SmartArt row `{line}`");
            assert_eq!(fields[16], "PowerPoint reference, no divergence allowed");
            assert!(
                fields[12]
                    .split('|')
                    .any(|owner| owner.starts_with("decorative:")),
                "external SmartArt row omits decorative ownership `{line}`"
            );
            assert_eq!(
                fields[12].split('|').count(),
                fields[13].split('|').count(),
                "external SmartArt ownership and geometry count `{line}`"
            );
            let text_owner_count = fields[12]
                .split('|')
                .filter(|owner| owner.starts_with("text:"))
                .count();
            assert_eq!(
                fields[14].split('|').count(),
                text_owner_count,
                "external SmartArt text owner and line-count cardinality `{line}`"
            );
            assert!(
                fields[14]
                    .split('|')
                    .all(|count| count.parse::<usize>().is_ok_and(|count| count > 0)),
                "external SmartArt line counts must be positive in `{line}`"
            );
            for index in [3usize, 6, 8] {
                assert_eq!(fields[index].len(), 64);
                assert!(
                    fields[index]
                        .bytes()
                        .all(|byte| byte.is_ascii_hexdigit() && !byte.is_ascii_uppercase())
                );
                assert!(
                    fields[index].bytes().any(|byte| byte != b'0'),
                    "external SmartArt hash must not be a placeholder in `{line}`"
                );
            }
            SmartartOracleRow {
                case: fields[0].to_owned(),
                family: fields[1].to_owned(),
                source: fields[2].to_owned(),
                source_sha256: fields[3].to_owned(),
                slide: fields[4].parse().unwrap(),
                oracle_pdf: fields[5].to_owned(),
                oracle_pdf_sha256: fields[6].to_owned(),
                oracle_png: fields[7].to_owned(),
                oracle_sha256: fields[8].to_owned(),
                width: fields[9].parse().unwrap(),
                height: fields[10].parse().unwrap(),
                node_texts: fields[11].to_owned(),
                shape_ownership: fields[12].to_owned(),
                shape_bounds_pt: fields[13].to_owned(),
                text_line_counts: fields[14]
                    .split('|')
                    .map(|count| count.parse().unwrap())
                    .collect(),
                expected_diagnostics: fields[15].to_owned(),
            }
        })
        .collect()
}

#[derive(Clone, Copy, Debug)]
struct SmartartCase {
    case: &'static str,
    family: &'static str,
    source: &'static str,
    slide: usize,
    node_texts: &'static str,
}

fn smartart_cases() -> &'static [SmartartCase; 6] {
    &[
        SmartartCase {
            case: "f220-list",
            family: "list",
            source: "f220-list.pptx",
            slide: 1,
            node_texts: "F220 list 1|F220 list 2|F220 list 3",
        },
        SmartartCase {
            case: "f220-hierarchy",
            family: "hierarchy",
            source: "f220-hierarchy.pptx",
            slide: 1,
            node_texts: "F220 hierarchy 1|F220 hierarchy 2|F220 hierarchy 3",
        },
        SmartartCase {
            case: "f220-cycle",
            family: "cycle",
            source: "f220-cycle.pptx",
            slide: 1,
            node_texts: "F220 cycle 1|F220 cycle 2|F220 cycle 3",
        },
        SmartartCase {
            case: "f220-relationship",
            family: "relationship",
            source: "f220-relationship.pptx",
            slide: 1,
            node_texts: "F220 relationship 1",
        },
        SmartartCase {
            case: "f220-matrix",
            family: "matrix",
            source: "f220-matrix.pptx",
            slide: 1,
            node_texts: "F220 matrix 1",
        },
        SmartartCase {
            case: "f220-pyramid",
            family: "pyramid",
            source: "f220-pyramid.pptx",
            slide: 1,
            node_texts: "F220 pyramid 1|F220 pyramid 2|F220 pyramid 3",
        },
    ]
}

const SMARTART_RESOURCE_ROOT: &str = "/Applications/Microsoft PowerPoint.app/Contents/Frameworks/SmartArt.framework/Versions/A/Resources";
const SMARTART_STYLE_RESOURCE: (&str, &str) = (
    "qs/simple1.gqs",
    "213edbaedea282b8d13dc85a874c9494e66b90546118e40e32614d939a704bd1",
);
const SMARTART_COLOR_RESOURCE: (&str, &str) = (
    "cs/accent1_1.gcs",
    "dc0a610ca9a665158d3afaff8612e29320e009f151990f74a265197a4e96f9a4",
);

fn pinned_smartart_resources_available() -> bool {
    [
        SMARTART_STYLE_RESOURCE.0,
        SMARTART_COLOR_RESOURCE.0,
        "lo/list1.glo",
        "lo/hierarchy1.glo",
        "lo/cycle1.glo",
        "lo/circlerelationship.glo",
        "lo/matrix1.glo",
        "lo/pyramid1.glo",
    ]
    .iter()
    .all(|relative| Path::new(SMARTART_RESOURCE_ROOT).join(relative).is_file())
}

fn smartart_layout_resource(family: &str) -> (&'static str, &'static str) {
    match family {
        "list" => (
            "lo/list1.glo",
            "e0e143896ab59ea7bf1e5bcf88151467271e0baad2f998c2e3e663384b985dee",
        ),
        "hierarchy" => (
            "lo/hierarchy1.glo",
            "92d8fd7ff62d662670cbe4aa723f0565b98d577e18d3e9a26c3901a64f68bf31",
        ),
        "cycle" => (
            "lo/cycle1.glo",
            "8a7e35b9099cff9fd646490ab9b36f8349d82e2568c92354f871f90315301461",
        ),
        "relationship" => (
            "lo/circlerelationship.glo",
            "0de1977f023d64027ebac0bb5ea27522690423b504a9ff577e381a727585bcd6",
        ),
        "matrix" => (
            "lo/matrix1.glo",
            "19a09a8304e33a88048fe66ba1d9e0488da1c4d6d0ea694f18279879612e1255",
        ),
        "pyramid" => (
            "lo/pyramid1.glo",
            "5cb535131a8ee862690c56455f7c56d75a68ccac7c2f7ee81da638de1af6eba2",
        ),
        other => panic!("unsupported authentic SmartArt family `{other}`"),
    }
}

fn read_pinned_smartart_resource(relative: &str, expected_sha256: &str) -> Vec<u8> {
    let path = Path::new(SMARTART_RESOURCE_ROOT).join(relative);
    let bytes = fs::read(&path).unwrap_or_else(|error| panic!("read {}: {error}", path.display()));
    assert_eq!(
        sha256_bytes(&bytes),
        expected_sha256,
        "installed SmartArt resource drift at {}",
        path.display()
    );
    bytes
}

fn authentic_smartart_data_model(family: &str) -> Vec<u8> {
    let points = (1..=3)
        .map(|ordinal| {
            format!(
                r#"<dgm:pt modelId="{ordinal}"><dgm:prSet/><dgm:t><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:r><a:rPr sz="1600"/><a:t>F220 {family} {ordinal}</a:t></a:r></a:p></dgm:t></dgm:pt>"#
            )
        })
        .collect::<String>();
    let connections = if family == "hierarchy" {
        r#"<dgm:cxn modelId="4" srcId="0" destId="1" srcOrd="0" destOrd="0" type="parOf"/><dgm:cxn modelId="5" srcId="1" destId="2" srcOrd="0" destOrd="0" type="parOf"/><dgm:cxn modelId="6" srcId="1" destId="3" srcOrd="1" destOrd="0" type="parOf"/>"#
    } else {
        r#"<dgm:cxn modelId="4" srcId="0" destId="1" srcOrd="0" destOrd="0" type="parOf"/><dgm:cxn modelId="5" srcId="0" destId="2" srcOrd="1" destOrd="0" type="parOf"/><dgm:cxn modelId="6" srcId="0" destId="3" srcOrd="2" destOrd="0" type="parOf"/>"#
    };
    format!(
        r#"<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="{A_NS}"><dgm:ptLst><dgm:pt modelId="0" type="doc"><dgm:prSet/></dgm:pt>{points}</dgm:ptLst><dgm:cxnLst>{connections}</dgm:cxnLst><dgm:bg/><dgm:whole/></dgm:dataModel>"#
    )
    .into_bytes()
}

fn authentic_smartart_oracle_source_bytes(family: &str) -> Vec<u8> {
    let (layout_path, layout_sha256) = smartart_layout_resource(family);
    let mut package =
        OpcPackage::from_reader(Cursor::new(smartart_rust_source_bytes(family))).unwrap();
    package.set_part(
        "/ppt/diagrams/data1.xml",
        authentic_smartart_data_model(family),
    );
    package.set_part(
        "/ppt/diagrams/layout1.xml",
        read_pinned_smartart_resource(layout_path, layout_sha256),
    );
    package.set_part(
        "/ppt/diagrams/quickStyle1.xml",
        read_pinned_smartart_resource(SMARTART_STYLE_RESOURCE.0, SMARTART_STYLE_RESOURCE.1),
    );
    package.set_part(
        "/ppt/diagrams/colors1.xml",
        read_pinned_smartart_resource(SMARTART_COLOR_RESOURCE.0, SMARTART_COLOR_RESOURCE.1),
    );
    package_bytes(package)
}

fn assert_authentic_smartart_oracle_source(family: &str, bytes: &[u8]) {
    let mut baseline = Presentation::new().unwrap();
    baseline.add_slide(0).unwrap();
    let baseline_package =
        OpcPackage::from_reader(Cursor::new(baseline.to_bytes().unwrap())).unwrap();
    let package = OpcPackage::from_reader(Cursor::new(bytes)).unwrap();
    assert_eq!(
        package.parts.len(),
        baseline_package.parts.len() + 4,
        "{family} must add exactly four diagram parts to the standard template"
    );
    assert_eq!(
        package
            .parts
            .keys()
            .filter(|part| part.starts_with("/ppt/diagrams/"))
            .map(String::as_str)
            .collect::<HashSet<_>>(),
        HashSet::from([
            "/ppt/diagrams/data1.xml",
            "/ppt/diagrams/layout1.xml",
            "/ppt/diagrams/quickStyle1.xml",
            "/ppt/diagrams/colors1.xml",
        ])
    );
    let (layout_path, layout_sha256) = smartart_layout_resource(family);
    assert_eq!(
        package.get_part("/ppt/diagrams/layout1.xml").unwrap(),
        read_pinned_smartart_resource(layout_path, layout_sha256)
    );
    assert_eq!(
        package.get_part("/ppt/diagrams/quickStyle1.xml").unwrap(),
        read_pinned_smartart_resource(SMARTART_STYLE_RESOURCE.0, SMARTART_STYLE_RESOURCE.1)
    );
    assert_eq!(
        package.get_part("/ppt/diagrams/colors1.xml").unwrap(),
        read_pinned_smartart_resource(SMARTART_COLOR_RESOURCE.0, SMARTART_COLOR_RESOURCE.1)
    );
    let data_xml = package.get_part("/ppt/diagrams/data1.xml").unwrap();
    let data = rpptx::CT_DiagramData::from_xml(data_xml).unwrap();
    assert_eq!(data.points().len(), 4, "{family} data-only point count");
    assert_eq!(
        data.connections().len(),
        3,
        "{family} data-only connection count"
    );
    assert!(
        !String::from_utf8_lossy(data_xml).contains(r#"type="pres""#),
        "{family} authentic source must not pre-author a presentation graph"
    );
    let presentation_part = package.main_document_part().unwrap();
    let presentation_model =
        CT_Presentation::from_xml(package.get_part(&presentation_part).unwrap()).unwrap();
    assert_eq!(presentation_model.slide_ids.len(), 1);
    assert_eq!(presentation_model.slide_master_ids.len(), 1);
    let presentation_relationships = package.get_part_rels(&presentation_part).unwrap();
    let slide_relationship = presentation_relationships
        .get_by_id(&presentation_model.slide_ids[0].relationship_id)
        .unwrap();
    let slide_part = OpcPackage::resolve_rel_target(&presentation_part, &slide_relationship.target);
    let slide_relationships = package.get_part_rels(&slide_part).unwrap();
    for relationship_type in [
        rel_types::DIAGRAM_DATA,
        rel_types::DIAGRAM_LAYOUT,
        rel_types::DIAGRAM_QUICK_STYLE,
        rel_types::DIAGRAM_COLORS,
    ] {
        assert_eq!(
            slide_relationships
                .items
                .iter()
                .filter(|relationship| relationship.rel_type == relationship_type)
                .count(),
            1,
            "{family} role {relationship_type}"
        );
    }
    assert!(
        slide_relationships
            .items
            .iter()
            .all(|relationship| relationship.rel_type != rel_types::DIAGRAM_DRAWING),
        "{family} authentic source must not contain a cached drawing relationship"
    );
    let facade = Presentation::from_bytes(bytes).unwrap();
    let infos = facade.smart_art(0).unwrap();
    assert_eq!(infos.len(), 1, "{family} SmartArt frame count");
    let info = &infos[0];
    assert_eq!(info.relationships.data, "smart-data");
    assert_eq!(info.relationships.layout, "smart-layout");
    assert_eq!(info.relationships.style, "smart-style");
    assert_eq!(info.relationships.colors, "smart-colors");
    assert!(info.relationships.drawing.is_none());
    assert!(matches!(info.data, rpptx::DiagramPart::Parsed(_)));
    assert!(matches!(info.layout, rpptx::DiagramPart::Parsed(_)));
    assert!(matches!(info.style, rpptx::DiagramPart::Parsed(_)));
    assert!(matches!(info.colors, rpptx::DiagramPart::Parsed(_)));
    assert!(info.drawing.is_none());
}

#[derive(Clone, Copy, Debug)]
struct SmartartComparison {
    geometry_error: f64,
    ssim: f64,
}

fn compare_smartart_geometry_and_png(
    actual_bounds: &[Rect],
    expected_bounds: &[Rect],
    actual_png: &[u8],
    expected_png: &[u8],
) -> SmartartComparison {
    assert_eq!(actual_bounds.len(), expected_bounds.len());
    let geometry_error = actual_bounds
        .iter()
        .zip(expected_bounds)
        .flat_map(|(actual, expected)| {
            [
                (actual.x - expected.x).abs(),
                (actual.y - expected.y).abs(),
                (actual.width - expected.width).abs(),
                (actual.height - expected.height).abs(),
            ]
        })
        .fold(0.0_f64, f64::max);
    SmartartComparison {
        geometry_error,
        ssim: smartart_png_ssim(actual_png, expected_png),
    }
}

fn smartart_shape_ownership(slide: &ResolvedSlide) -> String {
    let mut decorative = 0usize;
    slide
        .shapes
        .iter()
        .map(|shape| {
            let text = resolved_content_text(&shape.content);
            if text.is_empty() {
                decorative += 1;
                format!("decorative:{decorative}")
            } else {
                format!("text:{text}")
            }
        })
        .collect::<Vec<_>>()
        .join("|")
}

fn smartart_png_ssim(left: &[u8], right: &[u8]) -> f64 {
    static NEXT_SSIM_PATH: AtomicU64 = AtomicU64::new(0);
    let prefix = format!(
        "rpptx-f220-{}-{}",
        std::process::id(),
        NEXT_SSIM_PATH.fetch_add(1, std::sync::atomic::Ordering::Relaxed)
    );
    let left_path = std::env::temp_dir().join(format!("{prefix}-left.png"));
    let right_path = std::env::temp_dir().join(format!("{prefix}-right.png"));
    fs::write(&left_path, left).unwrap();
    fs::write(&right_path, right).unwrap();
    let output = Command::new("python3")
        .args(["-c", "from pathlib import Path; import sys; sys.path.insert(0, sys.argv[1]); from pptx_ssim_harness import decode_png, structural_similarity; print(structural_similarity(decode_png(Path(sys.argv[2])), decode_png(Path(sys.argv[3]))))"])
        .arg(workspace_root().join("scripts"))
        .arg(&left_path)
        .arg(&right_path)
        .output()
        .unwrap();
    fs::remove_file(&left_path).unwrap();
    fs::remove_file(&right_path).unwrap();
    assert!(
        output.status.success(),
        "SmartArt SSIM helper: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    String::from_utf8(output.stdout)
        .unwrap()
        .trim()
        .parse()
        .unwrap()
}

#[test]
fn smartart_ssim_temp_paths_are_unique_under_concurrency() {
    const WORKERS: usize = 8;
    let png = valid_one_pixel_png();
    let barrier = std::sync::Arc::new(std::sync::Barrier::new(WORKERS));
    std::thread::scope(|scope| {
        let mut handles = Vec::with_capacity(WORKERS);
        for _ in 0..WORKERS {
            let barrier = barrier.clone();
            let png = &png;
            handles.push(scope.spawn(move || {
                barrier.wait();
                smartart_png_ssim(png, png)
            }));
        }
        for handle in handles {
            assert_eq!(handle.join().unwrap(), 1.0);
        }
    });
}

#[derive(Clone, Debug)]
struct SmartartRasterComparison {
    full_ssim: f64,
    non_text_ssim: f64,
    actual_line_counts: Vec<usize>,
    expected_line_counts: Vec<usize>,
    max_ink_edge_error_pt: f64,
    max_line_width_error_pt: f64,
}

fn smartart_raster_comparison(
    actual_png: &[u8],
    expected_png: &[u8],
    text_owner_bounds: &[Rect],
) -> SmartartRasterComparison {
    static NEXT_COMPARISON_PATH: AtomicU64 = AtomicU64::new(0);
    let sequence = NEXT_COMPARISON_PATH.fetch_add(1, Ordering::Relaxed);
    let prefix = format!("rpptx-f220-raster-{}-{sequence}", std::process::id());
    let actual_path = std::env::temp_dir().join(format!("{prefix}-actual.png"));
    let expected_path = std::env::temp_dir().join(format!("{prefix}-expected.png"));
    fs::write(&actual_path, actual_png).unwrap();
    fs::write(&expected_path, expected_png).unwrap();
    let bounds = text_owner_bounds
        .iter()
        .map(|bounds| {
            format!(
                "{},{},{},{}",
                bounds.x, bounds.y, bounds.width, bounds.height
            )
        })
        .collect::<Vec<_>>()
        .join("|");
    let script = r#"
from pathlib import Path
import sys
sys.path.insert(0, sys.argv[1])
from pptx_ssim_harness import decode_png, structural_similarity

DPI = 150
POINTS_PER_INCH = 72
actual = decode_png(Path(sys.argv[2]))
expected = decode_png(Path(sys.argv[3]))
if actual[:2] != expected[:2]:
    raise ValueError(f'image dimensions differ: {actual[:2]} != {expected[:2]}')
width, height = actual[:2]

def owner_rect(value):
    x, y, w, h = (float(field) for field in value.split(','))
    left = max(0, int(x * DPI / POINTS_PER_INCH) - 2)
    top = max(0, int(y * DPI / POINTS_PER_INCH) - 2)
    right = min(width - 1, int((x + w) * DPI / POINTS_PER_INCH + 0.999999) + 2)
    bottom = min(height - 1, int((y + h) * DPI / POINTS_PER_INCH + 0.999999) + 2)
    return left, top, right, bottom

owners = [owner_rect(value) for value in sys.argv[4].split('|') if value]
if not owners:
    raise ValueError('SmartArt comparison has no text owners')

def is_ink(rgba):
    red, green, blue, alpha = rgba
    return alpha > 200 and max(red, green, blue) < 80

def line_boxes(image, owner):
    left, top, right, bottom = owner
    rgba = image[2]
    rows = []
    for y in range(top, bottom + 1):
        count = 0
        for x in range(left, right + 1):
            offset = (y * width + x) * 4
            if is_ink(rgba[offset:offset + 4]):
                count += 1
        if count >= 3:
            rows.append(y)
    groups = []
    for y in rows:
        if not groups or y > groups[-1][-1] + 1:
            groups.append([y])
        else:
            groups[-1].append(y)
    boxes = []
    for group in groups:
        columns = []
        for y in group:
            for x in range(left, right + 1):
                offset = (y * width + x) * 4
                if is_ink(rgba[offset:offset + 4]):
                    columns.append(x)
        if columns:
            boxes.append((min(columns), group[0], max(columns), group[-1]))
    if not boxes:
        raise ValueError(f'text owner {owner} has no detected ink lines')
    return boxes

actual_owners = [line_boxes(actual, owner) for owner in owners]
expected_owners = [line_boxes(expected, owner) for owner in owners]
actual_counts = [len(lines) for lines in actual_owners]
expected_counts = [len(lines) for lines in expected_owners]
if actual_counts != expected_counts:
    raise ValueError(f'text line counts differ: {actual_counts} != {expected_counts}')

edge_error_pixels = 0.0
width_error_pixels = 0.0
mask_boxes = []
for actual_lines, expected_lines in zip(actual_owners, expected_owners):
    for actual_box, expected_box in zip(actual_lines, expected_lines):
        actual_width = actual_box[2] - actual_box[0] + 1
        expected_width = expected_box[2] - expected_box[0] + 1
        width_error_pixels = max(width_error_pixels, abs(actual_width - expected_width))
        edge_error_pixels = max(
            edge_error_pixels,
            abs(actual_box[1] - expected_box[1]),
            abs(actual_box[3] - expected_box[3]),
            abs(actual_width - expected_width) / 2,
        )
        mask_boxes.append((
            max(0, min(actual_box[0], expected_box[0]) - 4),
            max(0, min(actual_box[1], expected_box[1]) - 4),
            min(width - 1, max(actual_box[2], expected_box[2]) + 4),
            min(height - 1, max(actual_box[3], expected_box[3]) + 4),
        ))

def without_text(image):
    rgba = bytearray(image[2])
    for left, top, right, bottom in mask_boxes:
        for y in range(top, bottom + 1):
            for x in range(left, right + 1):
                offset = (y * width + x) * 4
                rgba[offset:offset + 4] = b'\xff\xff\xff\xff'
    return width, height, bytes(rgba)

print(
    structural_similarity(actual, expected),
    structural_similarity(without_text(actual), without_text(expected)),
    '|'.join(str(value) for value in actual_counts),
    '|'.join(str(value) for value in expected_counts),
    edge_error_pixels * POINTS_PER_INCH / DPI,
    width_error_pixels * POINTS_PER_INCH / DPI,
    sep='\t',
)
"#;
    let output = Command::new("python3")
        .args(["-c", script])
        .arg(workspace_root().join("scripts"))
        .arg(&actual_path)
        .arg(&expected_path)
        .arg(bounds)
        .output()
        .unwrap();
    fs::remove_file(actual_path).unwrap();
    fs::remove_file(expected_path).unwrap();
    assert!(
        output.status.success(),
        "SmartArt raster comparator: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8(output.stdout).unwrap();
    let fields = stdout.trim().split('\t').collect::<Vec<_>>();
    assert_eq!(fields.len(), 6, "SmartArt raster metrics `{stdout}`");
    SmartartRasterComparison {
        full_ssim: fields[0].parse().unwrap(),
        non_text_ssim: fields[1].parse().unwrap(),
        actual_line_counts: fields[2]
            .split('|')
            .map(|value| value.parse().unwrap())
            .collect(),
        expected_line_counts: fields[3]
            .split('|')
            .map(|value| value.parse().unwrap())
            .collect(),
        max_ink_edge_error_pt: fields[4].parse().unwrap(),
        max_line_width_error_pt: fields[5].parse().unwrap(),
    }
}

fn normalize_smartart_oracle_png(path: &Path, width: u32, height: u32, index: usize) -> Vec<u8> {
    let output = std::env::temp_dir().join(format!(
        "rpptx-f220-normalized-{}-{index}.png",
        std::process::id()
    ));
    let result = Command::new("sips")
        .args(["-z", &height.to_string(), &width.to_string()])
        .arg(path)
        .args(["--out"])
        .arg(&output)
        .output()
        .expect("normalize PowerPoint SmartArt PNG");
    assert!(
        result.status.success(),
        "sips: {}",
        String::from_utf8_lossy(&result.stderr)
    );
    let bytes = fs::read(&output).unwrap();
    fs::remove_file(output).unwrap();
    bytes
}

fn sha256_bytes(bytes: &[u8]) -> String {
    static NEXT_SHA_PATH: AtomicU64 = AtomicU64::new(0);
    let path = std::env::temp_dir().join(format!(
        "rpptx-f220-sha-{}-{}",
        std::process::id(),
        NEXT_SHA_PATH.fetch_add(1, std::sync::atomic::Ordering::Relaxed)
    ));
    fs::write(&path, bytes).unwrap();
    let digest = sha256(&path);
    fs::remove_file(path).unwrap();
    digest
}

fn smartart_rust_source_bytes(family: &str) -> Vec<u8> {
    let (algorithm, parameters) = match family {
        "list" => ("lin", r#"<dgm:param type="linDir" val="fromL"/>"#),
        "hierarchy" => ("hierRoot", ""),
        "cycle" => (
            "cycle",
            r#"<dgm:param type="stAng" val="-90"/><dgm:param type="spanAng" val="240"/>"#,
        ),
        "relationship" => ("cycle", ""),
        "matrix" => (
            "composite",
            r#"<dgm:param type="stElem" val="node"/><dgm:param type="bkpt" val="2"/>"#,
        ),
        "pyramid" => ("pyra", r#"<dgm:param type="pyraLvlNode" val="1,2,3"/>"#),
        other => panic!("unsupported SmartArt oracle family `{other}`"),
    };
    let kinds = if family == "hierarchy" {
        ["node", "node", "asst"]
    } else {
        ["node", "node", "node"]
    };
    let points = (0..3)
        .map(|index| {
            let point_type = if kinds[index] == "node" {
                String::new()
            } else {
                format!(r#" type="{}""#, kinds[index])
            };
            let ordinal = index + 1;
            let presentation_id = 100 + ordinal;
            format!(
                r#"<dgm:pt modelId="{ordinal}"{point_type}><dgm:prSet/><dgm:t><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:r><a:rPr sz="1600"/><a:t>F220 {family} {ordinal}</a:t></a:r></a:p></dgm:t></dgm:pt><dgm:pt modelId="{presentation_id}" type="pres"><dgm:prSet/></dgm:pt>"#
            )
        })
        .collect::<String>();
    let ownership = (1..=3)
        .map(|index| {
            let presentation_id = 100 + index;
            let connection_id = 201 + index;
            format!(
                r#"<dgm:cxn modelId="{connection_id}" srcId="{index}" destId="{presentation_id}" srcOrd="0" destOrd="0" type="presOf"/>"#
            )
        })
        .collect::<String>();
    let topology = r#"<dgm:cxn modelId="205" srcId="1" destId="2" srcOrd="0" destOrd="0" type="parOf"/><dgm:cxn modelId="206" srcId="1" destId="3" srcOrd="1" destOrd="0" type="parOf"/>"#;
    let mut presentation = Presentation::new().unwrap();
    presentation.add_slide(0).unwrap();
    let mut package =
        OpcPackage::from_reader(Cursor::new(presentation.to_bytes().unwrap())).unwrap();
    let presentation_part = package.main_document_part().unwrap();
    let presentation_model =
        CT_Presentation::from_xml(package.get_part(&presentation_part).unwrap()).unwrap();
    assert_eq!(presentation_model.slide_ids.len(), 1);
    let slide_relationship = package
        .get_part_rels(&presentation_part)
        .unwrap()
        .get_by_id(&presentation_model.slide_ids[0].relationship_id)
        .unwrap();
    let slide_part = OpcPackage::resolve_rel_target(&presentation_part, &slide_relationship.target);
    package.set_part(
        &slide_part,
        format!(r#"<p:sld xmlns:p="{P_NS}" xmlns:a="{A_NS}" xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:r="{R_NS}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2" name="F220 {family} SmartArt"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="914400" y="914400"/><a:ext cx="7315200" cy="4572000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds r:dm="smart-data" r:lo="smart-layout" r:qs="smart-style" r:cs="smart-colors"/></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld></p:sld>"#).into_bytes(),
    );
    package.set_part(
        "/ppt/diagrams/data1.xml",
        format!(r#"<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="{A_NS}"><dgm:ptLst><dgm:pt modelId="0" type="doc"><dgm:prSet/></dgm:pt>{points}</dgm:ptLst><dgm:cxnLst><dgm:cxn modelId="201" srcId="0" destId="1" srcOrd="0" destOrd="0" type="parOf"/>{ownership}{topology}</dgm:cxnLst><dgm:bg/><dgm:whole/></dgm:dataModel>"#).into_bytes(),
    );
    package.set_part(
        "/ppt/diagrams/layout1.xml",
        format!(r#"<dgm:layoutDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" uniqueId="urn:f220:{family}"><dgm:layoutNode styleLbl="node0"><dgm:alg type="{algorithm}">{parameters}</dgm:alg></dgm:layoutNode></dgm:layoutDef>"#).into_bytes(),
    );
    package.set_part(
        "/ppt/diagrams/quickStyle1.xml",
        br#"<dgm:styleDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" uniqueId="urn:f220:style"><dgm:styleLbl name="node0"><dgm:style><a:lnRef idx="1"/><a:fillRef idx="1"/><a:effectRef idx="1"/><a:fontRef idx="minor"/></dgm:style></dgm:styleLbl></dgm:styleDef>"#.to_vec(),
    );
    package.set_part(
        "/ppt/diagrams/colors1.xml",
        br#"<dgm:colorsDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" uniqueId="urn:f220:colors"><dgm:styleLbl name="node0"><dgm:fillClrLst><a:schemeClr val="accent1"/></dgm:fillClrLst><dgm:linClrLst><a:schemeClr val="accent2"/></dgm:linClrLst><dgm:effectClrLst><a:schemeClr val="accent3"/></dgm:effectClrLst><dgm:txFillClrLst><a:schemeClr val="tx1"/></dgm:txFillClrLst></dgm:styleLbl></dgm:colorsDef>"#.to_vec(),
    );
    for (part, content_type) in [
        (
            "/ppt/diagrams/data1.xml",
            "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml",
        ),
        (
            "/ppt/diagrams/layout1.xml",
            "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml",
        ),
        (
            "/ppt/diagrams/quickStyle1.xml",
            "application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml",
        ),
        (
            "/ppt/diagrams/colors1.xml",
            "application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml",
        ),
    ] {
        package.content_types.add_override(part, content_type);
    }
    let relationships = package.get_or_create_part_rels(&slide_part);
    relationships.add_with_id(
        "smart-data",
        rel_types::DIAGRAM_DATA,
        "../diagrams/data1.xml",
    );
    relationships.add_with_id(
        "smart-layout",
        rel_types::DIAGRAM_LAYOUT,
        "../diagrams/layout1.xml",
    );
    relationships.add_with_id(
        "smart-style",
        rel_types::DIAGRAM_QUICK_STYLE,
        "../diagrams/quickStyle1.xml",
    );
    relationships.add_with_id(
        "smart-colors",
        rel_types::DIAGRAM_COLORS,
        "../diagrams/colors1.xml",
    );
    package_bytes(package)
}

#[test]
#[ignore = "writes deterministic SmartArt sources and Rust geometry outside the repository"]
fn generate_smartart_oracle_sources_and_rust_geometry() {
    let oracle_dir = PathBuf::from(
        std::env::var_os("RDOCX_PPTX_SMARTART_ORACLE_DIR")
            .expect("RDOCX_PPTX_SMARTART_ORACLE_DIR is required"),
    );
    fs::create_dir_all(&oracle_dir).unwrap();
    for row in smartart_cases() {
        let source = authentic_smartart_oracle_source_bytes(row.family);
        assert_authentic_smartart_oracle_source(row.family, &source);
        fs::write(oracle_dir.join(row.source), &source).unwrap();
        let presentation = Presentation::from_bytes(&source).unwrap();
        let (input, layout) = presentation.render_deterministic().unwrap();
        let slide_index = row.slide - 1;
        let slide = &input.slides[slide_index];
        for text in row.node_texts.split('|') {
            assert!(
                slide
                    .shapes
                    .iter()
                    .any(|shape| resolved_content_text(&shape.content) == text),
                "{} missing owned text `{text}`",
                row.case
            );
        }
        let ownership = smartart_shape_ownership(slide);
        assert!(
            ownership.contains("decorative:"),
            "{} decorative ownership",
            row.case
        );
        let bounds = slide
            .shapes
            .iter()
            .map(|shape| {
                let bounds = shape.bounds;
                format!(
                    "{},{},{},{}",
                    bounds.x, bounds.y, bounds.width, bounds.height
                )
            })
            .collect::<Vec<_>>()
            .join("|");
        let diagnostics = if slide.diagnostics.is_empty() {
            "-".to_owned()
        } else {
            slide
                .diagnostics
                .iter()
                .map(|diagnostic| diagnostic.message.as_str())
                .collect::<Vec<_>>()
                .join("|")
        };
        let rust_png = oxml_pdf::render_page_to_png(&layout, slide_index, 150.0).unwrap();
        fs::write(
            oracle_dir.join(format!("{}-rust.png", row.family)),
            &rust_png,
        )
        .unwrap();
        let rust_pixmap = tiny_skia::Pixmap::decode_png(&rust_png).unwrap();
        let (width, height) = (rust_pixmap.width(), rust_pixmap.height());
        eprintln!(
            "{}\t{}\t{}\t{}\t{}\t{}\t{}",
            row.case,
            sha256_bytes(&source),
            width,
            height,
            ownership,
            bounds,
            diagnostics
        );
    }
}

fn smartart_renderer_ceiling_source(with_text: bool) -> Vec<u8> {
    fn emu(points: f64) -> i64 {
        (points * 12_700.0).round() as i64
    }
    fn shape(
        id: u32,
        rect: (f64, f64, f64, f64),
        preset: &str,
        fill: &str,
        line: &str,
        text: Option<&str>,
    ) -> String {
        let (x, y, width, height) = rect;
        let text = text.map_or_else(String::new, |value| {
            format!(
                r#"<p:txBody><a:bodyPr anchor="ctr" lIns="194310" tIns="38100" rIns="194310" bIns="12700"/><a:lstStyle/><a:p><a:r><a:rPr sz="3400"><a:latin typeface="Calibri"/></a:rPr><a:t>{value}</a:t></a:r><a:endParaRPr sz="3400"><a:latin typeface="Calibri"/></a:endParaRPr></a:p></p:txBody>"#
            )
        });
        format!(
            r#"<p:sp><p:nvSpPr><p:cNvPr id="{id}" name="Ceiling {id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="{}" y="{}"/><a:ext cx="{}" cy="{}"/></a:xfrm><a:prstGeom prst="{preset}"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="{fill}"/></a:solidFill><a:ln w="25400"><a:solidFill><a:srgbClr val="{line}"/></a:solidFill><a:prstDash val="solid"/><a:round/></a:ln></p:spPr>{text}</p:sp>"#,
            emu(x),
            emu(y),
            emu(width),
            emu(height),
        )
    }

    let mut shapes = String::new();
    for index in 0..3 {
        let y = 116.590 + index as f64 * 121.433;
        shapes.push_str(&shape(
            2 + index as u32 * 2,
            (72.0, y, 576.0, 67.465),
            "rect",
            "D4DBEA",
            "4F81BD",
            None,
        ));
        shapes.push_str(&shape(
            3 + index as u32 * 2,
            (100.801, y - 39.516, 403.199, 79.032),
            "roundRect",
            "FFFFFF",
            "4774AB",
            with_text.then(|| ["F220 list 1", "F220 list 2", "F220 list 3"][index]),
        ));
    }
    let mut presentation = Presentation::new().unwrap();
    presentation.add_slide(0).unwrap();
    let mut package =
        OpcPackage::from_reader(Cursor::new(presentation.to_bytes().unwrap())).unwrap();
    let presentation_part = package.main_document_part().unwrap();
    let model = CT_Presentation::from_xml(package.get_part(&presentation_part).unwrap()).unwrap();
    let relationships = package.get_part_rels(&presentation_part).unwrap();
    let slide_relationship = relationships
        .get_by_id(&model.slide_ids[0].relationship_id)
        .unwrap();
    let slide_part = OpcPackage::resolve_rel_target(&presentation_part, &slide_relationship.target);
    package.set_part(
        &slide_part,
        format!(
            r#"<p:sld xmlns:p="{P_NS}" xmlns:a="{A_NS}" xmlns:r="{R_NS}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{shapes}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>"#
        )
        .into_bytes(),
    );
    package_bytes(package)
}

#[test]
#[ignore = "writes ordinary-shape renderer ceiling controls outside the repository"]
fn generate_smartart_renderer_ceiling_controls() {
    let oracle_dir = PathBuf::from(
        std::env::var_os("RDOCX_PPTX_SMARTART_ORACLE_DIR")
            .expect("RDOCX_PPTX_SMARTART_ORACLE_DIR is required"),
    );
    fs::create_dir_all(&oracle_dir).unwrap();
    for (name, with_text) in [("shapes", false), ("text", true)] {
        let source = smartart_renderer_ceiling_source(with_text);
        let source_path = oracle_dir.join(format!("f220-render-ceiling-{name}.pptx"));
        fs::write(&source_path, &source).unwrap();
        let presentation = Presentation::from_bytes(&source).unwrap();
        let (input, layout) = presentation.render_deterministic().unwrap();
        assert_eq!(input.slides[0].shapes.len(), 6);
        assert!(input.slides[0].diagnostics.is_empty());
        let png = oxml_pdf::render_page_to_png(&layout, 0, 150.0).unwrap();
        let png_path = oracle_dir.join(format!("f220-render-ceiling-{name}-rust.png"));
        fs::write(&png_path, &png).unwrap();
        let pixmap = tiny_skia::Pixmap::decode_png(&png).unwrap();
        eprintln!(
            "{name}\t{}\t{}\t{}\t{}",
            sha256_bytes(&source),
            pixmap.width(),
            pixmap.height(),
            png_path.display()
        );
    }
}

#[test]
fn smartart_differential_rejects_geometry_and_pixel_perturbations() {
    if !pinned_smartart_resources_available() {
        eprintln!(
            "SmartArt differential sensitivity skipped because pinned PowerPoint resources are absent"
        );
        return;
    }
    let source = authentic_smartart_oracle_source_bytes("list");
    let presentation = Presentation::from_bytes(&source).unwrap();
    let (mut input, baseline_layout) = presentation.render_deterministic().unwrap();

    let smartart_index = input.slides[0]
        .shapes
        .iter()
        .position(|shape| resolved_content_text(&shape.content) == "F220 list 1")
        .expect("expected supported SmartArt ordinary shape");
    let text_owner_bounds = input.slides[0]
        .shapes
        .iter()
        .filter_map(|shape| {
            (!resolved_content_text(&shape.content).is_empty()).then_some(shape.bounds)
        })
        .collect::<Vec<_>>();
    let original = input.slides[0].shapes[smartart_index].bounds;
    let baseline_png = oxml_pdf::render_page_to_png(&baseline_layout, 0, 150.0).unwrap();
    let perturbed = Rect {
        x: original.x + 1.01,
        ..original
    };
    let geometry_comparison =
        compare_smartart_geometry_and_png(&[perturbed], &[original], &baseline_png, &baseline_png);
    assert!(
        geometry_comparison.geometry_error > 1.0,
        "one-point threshold must reject a 1.01-point displacement"
    );
    let decorative_index = input.slides[0]
        .shapes
        .iter()
        .position(|shape| resolved_content_text(&shape.content).is_empty())
        .expect("expected supported SmartArt decorative shape");
    input.slides[0].shapes[decorative_index].fill =
        Some(oxml_layout::Paint::Solid(oxml_layout::Color {
            r: 1.0,
            g: 0.0,
            b: 0.0,
            a: 1.0,
        }));
    input.slides[0].shapes[decorative_index].bounds.width *= 3.0;
    input.slides[0].shapes[decorative_index].bounds.height *= 3.0;
    let mutated_layout = rpptx_render::layout_presentation_deterministic(&input).unwrap();
    let mutated_png = oxml_pdf::render_page_to_png(&mutated_layout, 0, 150.0).unwrap();
    let pixel_comparison =
        compare_smartart_geometry_and_png(&[original], &[original], &mutated_png, &baseline_png);
    let raster_comparison =
        smartart_raster_comparison(&mutated_png, &baseline_png, &text_owner_bounds);
    assert_eq!(pixel_comparison.geometry_error, 0.0);
    assert!(
        raster_comparison.non_text_ssim < 0.90,
        "calibrated rendered SmartArt mutation must fail the non-text PNG gate: {}",
        raster_comparison.non_text_ssim
    );
    eprintln!(
        "SmartArt sensitivity: geometry_error_pt={:.6}, mutated_non_text_ssim={:.9}, mutated_full_ssim={:.9}",
        geometry_comparison.geometry_error, raster_comparison.non_text_ssim, pixel_comparison.ssim
    );
}

fn parse_smartart_bounds(value: &str) -> Vec<Rect> {
    value
        .split('|')
        .filter(|value| !value.is_empty())
        .map(|value| {
            let fields = value
                .split(',')
                .map(|field| field.parse::<f64>().unwrap())
                .collect::<Vec<_>>();
            assert_eq!(fields.len(), 4);
            Rect {
                x: fields[0],
                y: fields[1],
                width: fields[2],
                height: fields[3],
            }
        })
        .collect()
}

#[test]
fn smartart_relationships_resolve_only_in_their_producing_scope() {
    let presentation =
        Presentation::from_bytes(&package_bytes(smartart_fixture_package())).unwrap();
    let first = presentation.smart_art(0).unwrap();
    let second = presentation.smart_art(1).unwrap();
    assert_eq!(
        first.iter().map(|info| info.shape_id).collect::<Vec<_>>(),
        [42, 52, 62]
    );
    assert_eq!(
        second.iter().map(|info| info.shape_id).collect::<Vec<_>>(),
        [42, 52, 62]
    );
    for (infos, slide_text) in [(&first, "Two"), (&second, "One")] {
        for (shape_id, expected_text) in
            [(42, slide_text), (52, "Layout scope"), (62, "Master scope")]
        {
            let info = infos.iter().find(|info| info.shape_id == shape_id).unwrap();
            assert_eq!(info.relationships.data, "smart-data");
            assert_eq!(info.relationships.drawing.as_deref(), Some("drawing-link"));
            let rpptx::DiagramPart::Parsed(data) = &info.data else {
                panic!("shape {shape_id} data part did not parse");
            };
            assert_eq!(
                data.points()[0].text.as_ref().unwrap().plain_text(),
                expected_text
            );
            assert!(matches!(info.drawing, Some(rpptx::DiagramPart::Parsed(_))));
        }
    }

    let mut missing = smartart_fixture_package();
    missing.parts.remove("/custom/diagrams/two/data1.xml");
    let missing = Presentation::from_bytes(&package_bytes(missing)).unwrap();
    assert!(matches!(
        missing
            .smart_art(0)
            .unwrap()
            .iter()
            .find(|info| info.shape_id == 42)
            .unwrap()
            .data,
        rpptx::DiagramPart::MissingTarget(_)
    ));

    let mut external = smartart_fixture_package();
    let relationship = external
        .get_or_create_part_rels(SLIDE_TWO_PART)
        .items
        .iter_mut()
        .find(|relationship| relationship.id == "smart-layout")
        .unwrap();
    relationship.target = "https://example.com/layout.xml".to_owned();
    relationship.target_mode = Some("External".to_owned());
    let external = Presentation::from_bytes(&package_bytes(external)).unwrap();
    assert!(matches!(
        external
            .smart_art(0)
            .unwrap()
            .iter()
            .find(|info| info.shape_id == 42)
            .unwrap()
            .layout,
        rpptx::DiagramPart::External(ref target) if target == "https://example.com/layout.xml"
    ));

    let mut wrong_type = smartart_fixture_package();
    wrong_type
        .get_or_create_part_rels(SLIDE_TWO_PART)
        .items
        .iter_mut()
        .find(|relationship| relationship.id == "smart-style")
        .unwrap()
        .rel_type = rel_types::IMAGE.to_owned();
    let wrong_type = Presentation::from_bytes(&package_bytes(wrong_type)).unwrap();
    assert!(matches!(
        wrong_type
            .smart_art(0)
            .unwrap()
            .iter()
            .find(|info| info.shape_id == 42)
            .unwrap()
            .style,
        rpptx::DiagramPart::Invalid(_)
    ));
}

#[test]
fn corpus_smartart_drawings_resolve_from_their_producing_scope() {
    let Some(paths) = corpus_paths() else {
        return;
    };
    let mut resolved_drawings = 0usize;
    for path in paths {
        let presentation =
            Presentation::open(&path).unwrap_or_else(|error| panic!("{}: {error}", path.display()));
        for slide_index in 0..presentation.len() {
            for info in presentation.smart_art(slide_index).unwrap() {
                let rpptx::DiagramPart::Parsed(data) = &info.data else {
                    continue;
                };
                let Some(drawing_id) = data.drawing_relationship_id() else {
                    continue;
                };
                resolved_drawings += 1;
                assert_eq!(info.relationships.drawing.as_deref(), Some(drawing_id));
                assert!(matches!(info.drawing, Some(rpptx::DiagramPart::Parsed(_))));
            }
        }
    }
    assert!(resolved_drawings > 0);
}

#[test]
fn duplicated_smartart_remaps_the_complete_diagram_relationship_graph() {
    let mut presentation =
        Presentation::from_bytes(&package_bytes(smartart_fixture_package())).unwrap();
    presentation.duplicate_slide(0).unwrap();
    let bytes = presentation.to_bytes().unwrap();
    let package = open_opc(&bytes, "F-219 duplicated SmartArt");
    let root = CT_Presentation::from_xml(package.get_part(PRESENTATION_PART).unwrap()).unwrap();
    let presentation_rels = package.get_part_rels(PRESENTATION_PART).unwrap();
    let slide_parts = root
        .slide_ids
        .iter()
        .take(2)
        .map(|slide| {
            let relationship = presentation_rels.get_by_id(&slide.relationship_id).unwrap();
            OpcPackage::resolve_rel_target(PRESENTATION_PART, &relationship.target)
        })
        .collect::<Vec<_>>();
    assert_ne!(slide_parts[0], slide_parts[1]);
    let source_rels = package.get_part_rels(&slide_parts[0]).unwrap();
    let copied_rels = package.get_part_rels(&slide_parts[1]).unwrap();
    for relationship_type in [
        rel_types::DIAGRAM_DATA,
        rel_types::DIAGRAM_LAYOUT,
        rel_types::DIAGRAM_QUICK_STYLE,
        rel_types::DIAGRAM_COLORS,
        rel_types::DIAGRAM_DRAWING,
    ] {
        let source = source_rels.get_by_type(relationship_type).unwrap();
        let copied = copied_rels.get_by_type(relationship_type).unwrap();
        assert_ne!(
            OpcPackage::resolve_rel_target(&slide_parts[0], &source.target),
            OpcPackage::resolve_rel_target(&slide_parts[1], &copied.target)
        );
    }
    let source_data = OpcPackage::resolve_rel_target(
        &slide_parts[0],
        &source_rels
            .get_by_type(rel_types::DIAGRAM_DATA)
            .unwrap()
            .target,
    );
    let copied_data = OpcPackage::resolve_rel_target(
        &slide_parts[1],
        &copied_rels
            .get_by_type(rel_types::DIAGRAM_DATA)
            .unwrap()
            .target,
    );
    let source_drawing = OpcPackage::resolve_rel_target(
        &slide_parts[0],
        &source_rels
            .get_by_type(rel_types::DIAGRAM_DRAWING)
            .unwrap()
            .target,
    );
    let copied_drawing = OpcPackage::resolve_rel_target(
        &slide_parts[1],
        &copied_rels
            .get_by_type(rel_types::DIAGRAM_DRAWING)
            .unwrap()
            .target,
    );
    assert_ne!(source_drawing, copied_drawing);
    let source_data =
        rpptx::CT_DiagramData::from_xml(package.get_part(&source_data).unwrap()).unwrap();
    assert_eq!(
        source_data.drawing_relationship_id(),
        Some(
            source_rels
                .get_by_type(rel_types::DIAGRAM_DRAWING)
                .unwrap()
                .id
                .as_str()
        )
    );
    let copied_data =
        rpptx::CT_DiagramData::from_xml(package.get_part(&copied_data).unwrap()).unwrap();
    assert_eq!(
        copied_data.drawing_relationship_id(),
        Some(
            copied_rels
                .get_by_type(rel_types::DIAGRAM_DRAWING)
                .unwrap()
                .id
                .as_str()
        )
    );
    let source_drawing_rels = package.get_part_rels(&source_drawing).unwrap();
    let copied_drawing_rels = package.get_part_rels(&copied_drawing).unwrap();
    let source_image = OpcPackage::resolve_rel_target(
        &source_drawing,
        &source_drawing_rels
            .get_by_type(rel_types::IMAGE)
            .unwrap()
            .target,
    );
    let copied_image = OpcPackage::resolve_rel_target(
        &copied_drawing,
        &copied_drawing_rels
            .get_by_type(rel_types::IMAGE)
            .unwrap()
            .target,
    );
    assert_eq!(source_image, copied_image);
    assert!(
        Presentation::from_bytes(&bytes)
            .unwrap()
            .validate()
            .is_empty()
    );
}

#[test]
fn transferred_smartart_copies_the_complete_graph_without_cross_presentation_aliasing() {
    let mut source_package = smartart_transfer_source_package();
    source_package
        .get_or_create_part_rels("/custom/diagrams/two/data1.xml")
        .add_with_id("nested-layout", rel_types::DIAGRAM_LAYOUT, "layout1.xml");
    source_package
        .get_or_create_part_rels("/custom/diagrams/two/layout1.xml")
        .add_with_id("nested-data", rel_types::DIAGRAM_DATA, "data1.xml");
    let source_bytes = package_bytes(source_package);
    let source = Presentation::from_bytes(&source_bytes).unwrap();
    let source_before = source.to_bytes().unwrap();
    let mut destination_package = smartart_transfer_destination_package();
    destination_package.set_part(
        "/custom/diagrams/two/data1.xml",
        smartart_data_xml("Destination decoy"),
    );
    let mut destination = Presentation::from_bytes(&package_bytes(destination_package)).unwrap();

    destination
        .transfer_smartart_slide_from(&source, 0, 0)
        .expect("transfer SmartArt slide");

    assert_eq!(source.to_bytes().unwrap(), source_before);
    let transferred = destination
        .smart_art(0)
        .unwrap()
        .into_iter()
        .find(|info| {
            matches!(
                &info.data,
                rpptx::DiagramPart::Parsed(data)
                    if data.points()[0].text.as_ref().unwrap().plain_text() == "Two"
            )
        })
        .unwrap();
    let rpptx::DiagramPart::Parsed(data) = transferred.data else {
        panic!("transferred data part did not parse");
    };
    assert_eq!(data.points()[0].text.as_ref().unwrap().plain_text(), "Two");

    let bytes = destination.to_bytes().unwrap();
    let package = open_opc(&bytes, "F-219 transferred SmartArt");
    let presentation_part = package.main_document_part().unwrap();
    let root = CT_Presentation::from_xml(package.get_part(&presentation_part).unwrap()).unwrap();
    let presentation_rels = package.get_part_rels(&presentation_part).unwrap();
    let slide_parts = root
        .slide_ids
        .iter()
        .map(|slide| {
            let relationship = presentation_rels.get_by_id(&slide.relationship_id).unwrap();
            OpcPackage::resolve_rel_target(&presentation_part, &relationship.target)
        })
        .collect::<Vec<_>>();
    let transferred_part = &slide_parts[0];
    let transferred_rels = package.get_part_rels(transferred_part).unwrap();
    for (relationship_type, colliding_source_part) in [
        (rel_types::DIAGRAM_DATA, "/custom/diagrams/two/data1.xml"),
        (
            rel_types::DIAGRAM_LAYOUT,
            "/custom/diagrams/two/layout1.xml",
        ),
        (
            rel_types::DIAGRAM_QUICK_STYLE,
            "/custom/diagrams/two/style1.xml",
        ),
        (
            rel_types::DIAGRAM_COLORS,
            "/custom/diagrams/two/colors1.xml",
        ),
        (
            rel_types::DIAGRAM_DRAWING,
            "/custom/diagrams/two/drawing1.xml",
        ),
    ] {
        let transferred = transferred_rels.get_by_type(relationship_type).unwrap();
        let transferred_target =
            OpcPackage::resolve_rel_target(transferred_part, &transferred.target);
        assert_ne!(transferred_target, colliding_source_part);
        assert!(package.get_part(&transferred_target).is_some());
    }
    let transferred_data = transferred_rels
        .get_by_type(rel_types::DIAGRAM_DATA)
        .unwrap();
    let transferred_data_part =
        OpcPackage::resolve_rel_target(transferred_part, &transferred_data.target);
    let transferred_data =
        rpptx::CT_DiagramData::from_xml(package.get_part(&transferred_data_part).unwrap()).unwrap();
    assert_eq!(
        transferred_data.drawing_relationship_id(),
        Some(
            transferred_rels
                .get_by_type(rel_types::DIAGRAM_DRAWING)
                .unwrap()
                .id
                .as_str()
        )
    );
    let copied_nested_layout = package
        .get_part_rels(&transferred_data_part)
        .unwrap()
        .get_by_type(rel_types::DIAGRAM_LAYOUT)
        .unwrap();
    let copied_nested_layout =
        OpcPackage::resolve_rel_target(&transferred_data_part, &copied_nested_layout.target);
    let direct_layout = transferred_rels
        .get_by_type(rel_types::DIAGRAM_LAYOUT)
        .unwrap();
    assert_eq!(
        copied_nested_layout,
        OpcPackage::resolve_rel_target(transferred_part, &direct_layout.target)
    );
    let transferred_drawing = transferred_rels
        .get_by_type(rel_types::DIAGRAM_DRAWING)
        .unwrap();
    let transferred_drawing_part =
        OpcPackage::resolve_rel_target(transferred_part, &transferred_drawing.target);
    let image_relationship = package
        .get_part_rels(&transferred_drawing_part)
        .unwrap()
        .get_by_type(rel_types::IMAGE)
        .unwrap();
    assert_eq!(
        OpcPackage::resolve_rel_target(&transferred_drawing_part, &image_relationship.target),
        "/ppt/media/image1.png"
    );
    assert!(
        Presentation::from_bytes(&bytes)
            .unwrap()
            .validate()
            .is_empty()
    );
}

#[test]
fn cross_presentation_transfer_rejects_unsupported_owned_graphs_without_aliasing() {
    let source_with_notes =
        Presentation::from_bytes(&package_bytes(smartart_fixture_package())).unwrap();
    let mut destination =
        Presentation::from_bytes(&package_bytes(smartart_transfer_destination_package())).unwrap();
    let before = destination.to_bytes().unwrap();
    assert!(
        destination
            .transfer_smartart_slide_from(&source_with_notes, 0, 0)
            .is_err()
    );
    assert_eq!(destination.to_bytes().unwrap(), before);

    let collision_part = "/custom/charts/shared.xml";
    let mut source_package = smartart_transfer_source_package();
    source_package.set_part(collision_part, b"source-only chart".to_vec());
    source_package
        .content_types
        .add_override(collision_part, content_types::CHART);
    source_package
        .get_or_create_part_rels("/custom/diagrams/two/drawing1.xml")
        .add_with_id("source-chart", rel_types::CHART, "../../charts/shared.xml");
    let source = Presentation::from_bytes(&package_bytes(source_package)).unwrap();

    let mut destination_package = smartart_transfer_destination_package();
    destination_package.set_part(collision_part, b"destination chart".to_vec());
    destination_package
        .content_types
        .add_override(collision_part, content_types::CHART);
    let mut destination = Presentation::from_bytes(&package_bytes(destination_package)).unwrap();
    let before = destination.to_bytes().unwrap();
    assert!(
        destination
            .transfer_smartart_slide_from(&source, 0, 0)
            .is_err()
    );
    assert_eq!(destination.to_bytes().unwrap(), before);
    let saved = open_opc(&before, "bounded transfer destination");
    assert_eq!(
        saved.get_part(collision_part).unwrap(),
        b"destination chart"
    );

    assert!(
        destination
            .transfer_smartart_slide_from(&source, 0, 999)
            .is_err()
    );
    assert_eq!(destination.to_bytes().unwrap(), before);
}

#[test]
fn cross_presentation_smartart_transfer_rejects_placeholders_atomically() {
    let original = String::from_utf8(
        smartart_transfer_source_package()
            .get_part(SLIDE_TWO_PART)
            .unwrap()
            .to_vec(),
    )
    .unwrap();
    let graphic_frame_placeholder = original.replacen(
        "<p:nvPr/>",
        r#"<p:nvPr><p:ph type="body" idx="7"/></p:nvPr>"#,
        1,
    );
    let preserved_choice_placeholder = original.replace(
        "</p:spTree>",
        r#"<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="p"><p:sp><p:nvSpPr><p:cNvPr id="43" name="Preserved Choice Placeholder"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="9"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp></mc:Choice><mc:Fallback/></mc:AlternateContent></p:spTree>"#,
    );

    for slide in [graphic_frame_placeholder, preserved_choice_placeholder] {
        let mut source_package = smartart_transfer_source_package();
        source_package.set_part(SLIDE_TWO_PART, slide.into_bytes());
        let source = Presentation::from_bytes(&package_bytes(source_package)).unwrap();
        let mut destination =
            Presentation::from_bytes(&package_bytes(smartart_transfer_destination_package()))
                .unwrap();
        let before = destination.to_bytes().unwrap();

        assert!(
            destination
                .transfer_smartart_slide_from(&source, 0, 0)
                .is_err()
        );
        assert_eq!(destination.to_bytes().unwrap(), before);
    }
}

#[test]
fn cross_presentation_smartart_transfer_rejects_slide_images_with_owned_graphs_atomically() {
    let image_part = "/ppt/media/slide-owned.png";
    let dependency_part = "/custom/charts/image-owner.xml";
    let mut source_package = smartart_transfer_source_package();
    source_package.set_part(image_part, b"source slide image".to_vec());
    source_package.set_part(dependency_part, b"source nested chart".to_vec());
    source_package.content_types.add_default("png", "image/png");
    source_package
        .content_types
        .add_override(dependency_part, content_types::CHART);
    source_package
        .get_or_create_part_rels(SLIDE_TWO_PART)
        .add_with_id(
            "slide-owned-image",
            rel_types::IMAGE,
            "../../ppt/media/slide-owned.png",
        );
    source_package
        .get_or_create_part_rels(image_part)
        .add_with_id(
            "image-owned-chart",
            rel_types::CHART,
            "../../custom/charts/image-owner.xml",
        );
    let source = Presentation::from_bytes(&package_bytes(source_package)).unwrap();

    let mut destination =
        Presentation::from_bytes(&package_bytes(smartart_transfer_destination_package())).unwrap();
    let before = destination.to_bytes().unwrap();

    assert!(
        destination
            .transfer_smartart_slide_from(&source, 0, 0)
            .is_err()
    );
    assert_eq!(destination.to_bytes().unwrap(), before);
}

#[test]
fn cross_presentation_smartart_transfer_rejects_external_slide_layout_atomically() {
    let mut source_package = smartart_transfer_source_package();
    let source_layout = source_package
        .get_or_create_part_rels(SLIDE_TWO_PART)
        .items
        .iter_mut()
        .find(|relationship| relationship.rel_type == rel_types::SLIDE_LAYOUT)
        .unwrap();
    source_layout.target = "https://example.com/source-layout.xml".to_owned();
    source_layout.target_mode = Some("External".to_owned());
    let source_bytes = package_bytes(source_package);
    let source = Presentation::from_bytes(&source_bytes).unwrap();
    let mut destination =
        Presentation::from_bytes(&package_bytes(smartart_transfer_destination_package())).unwrap();
    let before = destination.to_bytes().unwrap();

    assert!(
        destination
            .transfer_smartart_slide_from(&source, 0, 0)
            .is_err()
    );
    assert_eq!(destination.to_bytes().unwrap(), before);
}

#[test]
fn cross_presentation_smartart_transfer_rejects_deep_diagram_graph_atomically() {
    let mut source_package = smartart_transfer_source_package();
    let drawing_part = "/custom/diagrams/two/drawing1.xml";
    source_package
        .get_or_create_part_rels(drawing_part)
        .add_with_id("deep-chain", rel_types::DIAGRAM_LAYOUT, "deep/part0.xml");
    for index in 0..130 {
        let part = format!("/custom/diagrams/two/deep/part{index}.xml");
        source_package.set_part(
            &part,
            br#"<dgm:layoutDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"/>"#
                .to_vec(),
        );
        source_package.content_types.add_override(
            &part,
            "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml",
        );
        if index < 129 {
            source_package
                .get_or_create_part_rels(&part)
                .add(rel_types::DIAGRAM_LAYOUT, &format!("part{}.xml", index + 1));
        }
    }
    let source_bytes = package_bytes(source_package);
    let source = Presentation::from_bytes(&source_bytes).unwrap();
    let mut destination =
        Presentation::from_bytes(&package_bytes(smartart_transfer_destination_package())).unwrap();
    let before = destination.to_bytes().unwrap();

    let error = match destination.transfer_smartart_slide_from(&source, 0, 0) {
        Ok(_) => panic!("deep graph transfer unexpectedly succeeded"),
        Err(error) => error.to_string(),
    };
    assert!(error.contains("128-part transfer bound"), "{error}");
    assert_eq!(destination.to_bytes().unwrap(), before);

    let mut duplicate = Presentation::from_bytes(&source_bytes).unwrap();
    let before = duplicate.to_bytes().unwrap();
    let error = match duplicate.duplicate_slide(0) {
        Ok(_) => panic!("deep graph duplication unexpectedly succeeded"),
        Err(error) => error.to_string(),
    };
    assert!(error.contains("128-part copy bound"), "{error}");
    assert_eq!(duplicate.to_bytes().unwrap(), before);
}

#[test]
fn failed_smartart_node_mutation_leaves_the_package_unchanged() {
    let mut presentation =
        Presentation::from_bytes(&package_bytes(smartart_fixture_package())).unwrap();
    let before = presentation.to_bytes().unwrap();
    assert!(
        presentation
            .set_smart_art_node_text(0, 999, "n1", "fail")
            .is_err()
    );
    assert_eq!(presentation.to_bytes().unwrap(), before);
    assert!(
        presentation
            .set_smart_art_node_text(0, 42, "missing", "fail")
            .is_err()
    );
    assert_eq!(presentation.to_bytes().unwrap(), before);

    let mut missing_target = smartart_fixture_package();
    missing_target
        .parts
        .remove("/custom/diagrams/two/data1.xml");
    let mut missing_target = Presentation::from_bytes(&package_bytes(missing_target)).unwrap();
    let before = format!("{missing_target:?}");
    assert!(
        missing_target
            .set_smart_art_node_text(0, 42, "n1", "fail")
            .is_err()
    );
    assert_eq!(format!("{missing_target:?}"), before);

    let mut malformed = smartart_fixture_package();
    malformed.set_part("/custom/diagrams/two/data1.xml", b"<not-diagram/>".to_vec());
    let mut malformed = Presentation::from_bytes(&package_bytes(malformed)).unwrap();
    let before = malformed.to_bytes().unwrap();
    assert!(
        malformed
            .set_smart_art_node_text(0, 42, "n1", "fail")
            .is_err()
    );
    assert_eq!(malformed.to_bytes().unwrap(), before);

    let mut invalid_graph = smartart_fixture_package();
    invalid_graph
        .parts
        .remove("/custom/diagrams/two/drawing1.xml");
    let mut invalid_graph = Presentation::from_bytes(&package_bytes(invalid_graph)).unwrap();
    let before = format!("{invalid_graph:?}");
    assert!(
        invalid_graph
            .set_smart_art_node_text(0, 42, "n1", "fail")
            .is_err()
    );
    assert_eq!(format!("{invalid_graph:?}"), before);
}

#[test]
fn smartart_node_mutation_validates_every_relationship_role_atomically() {
    let mut variants = Vec::new();

    let mut missing_style = smartart_fixture_package();
    missing_style
        .get_or_create_part_rels(SLIDE_TWO_PART)
        .items
        .retain(|relationship| relationship.id != "smart-style");
    variants.push(("missing style id", missing_style));

    let mut wrong_layout_type = smartart_fixture_package();
    wrong_layout_type
        .get_or_create_part_rels(SLIDE_TWO_PART)
        .items
        .iter_mut()
        .find(|relationship| relationship.id == "smart-layout")
        .unwrap()
        .rel_type = rel_types::IMAGE.to_owned();
    variants.push(("wrong layout type", wrong_layout_type));

    let mut missing_drawing = smartart_fixture_package();
    missing_drawing
        .get_or_create_part_rels(SLIDE_TWO_PART)
        .items
        .retain(|relationship| relationship.id != "drawing-link");
    variants.push(("missing cached drawing id", missing_drawing));

    let mut wrong_drawing_type = smartart_fixture_package();
    wrong_drawing_type
        .get_or_create_part_rels(SLIDE_TWO_PART)
        .items
        .iter_mut()
        .find(|relationship| relationship.id == "drawing-link")
        .unwrap()
        .rel_type = rel_types::IMAGE.to_owned();
    variants.push(("wrong cached drawing type", wrong_drawing_type));

    for (case, package) in variants {
        let mut presentation = Presentation::from_bytes(&package_bytes(package)).unwrap();
        let before = format!("{presentation:?}");
        let error = presentation
            .set_smart_art_node_text(0, 42, "n1", "must reject")
            .unwrap_err()
            .to_string();
        assert!(
            error.contains("relationship") || error.contains("Relationship"),
            "{case}: {error}"
        );
        assert_eq!(format!("{presentation:?}"), before, "{case}");
    }
}

#[test]
fn smartart_transfer_validates_relationship_roles_and_exactly_one_layout_atomically() {
    let mut variants = Vec::new();

    let mut missing_colors = smartart_transfer_source_package();
    missing_colors
        .get_or_create_part_rels(SLIDE_TWO_PART)
        .items
        .retain(|relationship| relationship.id != "smart-colors");
    variants.push(("missing colors id", missing_colors));

    let mut wrong_drawing_type = smartart_transfer_source_package();
    wrong_drawing_type
        .get_or_create_part_rels(SLIDE_TWO_PART)
        .items
        .iter_mut()
        .find(|relationship| relationship.id == "drawing-link")
        .unwrap()
        .rel_type = rel_types::IMAGE.to_owned();
    variants.push(("wrong cached drawing type", wrong_drawing_type));

    let mut no_layout = smartart_transfer_source_package();
    no_layout
        .get_or_create_part_rels(SLIDE_TWO_PART)
        .items
        .retain(|relationship| relationship.rel_type != rel_types::SLIDE_LAYOUT);
    variants.push(("no source slide layout", no_layout));

    let mut two_layouts = smartart_transfer_source_package();
    two_layouts
        .get_or_create_part_rels(SLIDE_TWO_PART)
        .add_with_id(
            "second-slide-layout",
            rel_types::SLIDE_LAYOUT,
            "../slideLayouts/slideLayout1.xml",
        );
    variants.push(("two source slide layouts", two_layouts));

    for (case, source_package) in variants {
        let source = Presentation::from_bytes(&package_bytes(source_package)).unwrap();
        let mut destination =
            Presentation::from_bytes(&package_bytes(smartart_transfer_destination_package()))
                .unwrap();
        let before = destination.to_bytes().unwrap();
        let error = match destination.transfer_smartart_slide_from(&source, 0, 0) {
            Ok(_) => panic!("{case}: transfer unexpectedly succeeded"),
            Err(error) => error.to_string(),
        };
        assert!(
            error.contains("relationship") || error.contains("slide-layout"),
            "{case}: {error}"
        );
        assert_eq!(destination.to_bytes().unwrap(), before, "{case}");
    }
}

fn smartart_fixture_package() -> OpcPackage {
    let mut package = fixture_package();
    let master_part = "/custom/masters/scope.xml";
    package.set_part(LAYOUT_PART, smartart_layout_xml());
    package.set_part(master_part, smartart_master_xml());
    package
        .content_types
        .add_override(master_part, content_types::SLIDE_MASTER);
    package
        .get_or_create_part_rels(LAYOUT_PART)
        .add(rel_types::SLIDE_MASTER, "../masters/scope.xml");
    for (slide_part, scope, text) in [
        (SLIDE_ONE_PART, "one", "One"),
        (SLIDE_TWO_PART, "two", "Two"),
    ] {
        package.set_part(slide_part, smartart_slide_xml());
        add_smartart_scope(&mut package, slide_part, scope, text);
    }
    add_smartart_scope(&mut package, LAYOUT_PART, "layout-scope", "Layout scope");
    add_smartart_scope(&mut package, master_part, "master-scope", "Master scope");
    package.set_part("/ppt/media/image1.png", b"shared image bytes".to_vec());
    package.content_types.add_default("png", "image/png");
    package
}

fn make_smartart_renderable(package: &mut OpcPackage) {
    let master_part = "/custom/masters/scope.xml";
    let theme_part = "/custom/theme/theme1.xml";
    package.set_part(
        theme_part,
        CT_OfficeStyleSheet::office_default().to_xml().unwrap(),
    );
    package
        .content_types
        .add_override(theme_part, content_types::THEME);
    package
        .get_or_create_part_rels(master_part)
        .add(rel_types::THEME, "../theme/theme1.xml");
    for scope in ["one", "two", "layout-scope", "master-scope"] {
        package.set_part(
            &format!("/custom/diagrams/{scope}/style1.xml"),
            read_pinned_smartart_resource(SMARTART_STYLE_RESOURCE.0, SMARTART_STYLE_RESOURCE.1),
        );
        package.set_part(
            &format!("/custom/diagrams/{scope}/colors1.xml"),
            read_pinned_smartart_resource(SMARTART_COLOR_RESOURCE.0, SMARTART_COLOR_RESOURCE.1),
        );
    }
    package.set_part(
        "/custom/diagrams/two/data1.xml",
        format!(r#"<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="{A_NS}" xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"><dgm:ptLst><dgm:pt modelId="0" type="doc"><dgm:prSet/></dgm:pt><dgm:pt modelId="n1"><dgm:prSet/><dgm:t><a:bodyPr anchor="ctr" lIns="50800"/><a:lstStyle/><a:p><a:r><a:rPr b="1" i="1" sz="1600"><a:solidFill><a:schemeClr val="accent4"/></a:solidFill><a:latin typeface="+mn-lt"/></a:rPr><a:t>Two</a:t></a:r></a:p></dgm:t></dgm:pt><dgm:pt modelId="n2"><dgm:prSet/><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1600"/><a:t>Two child 2</a:t></a:r></a:p></dgm:t></dgm:pt><dgm:pt modelId="n3"><dgm:prSet/><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1600"/><a:t>Two child 3</a:t></a:r></a:p></dgm:t></dgm:pt></dgm:ptLst><dgm:cxnLst><dgm:cxn modelId="m1" srcId="0" destId="n1" srcOrd="0" destOrd="0" type="parOf"/><dgm:cxn modelId="m2" srcId="0" destId="n2" srcOrd="1" destOrd="0" type="parOf"/><dgm:cxn modelId="m3" srcId="0" destId="n3" srcOrd="2" destOrd="0" type="parOf"/></dgm:cxnLst><dgm:extLst><a:ext uri="http://schemas.microsoft.com/office/drawing/2008/diagram"><dsp:dataModelExt relId="drawing-link" minVer="http://schemas.openxmlformats.org/drawingml/2006/diagram"/></a:ext></dgm:extLst></dgm:dataModel>"#).into_bytes(),
    );
}

fn wrap_slide_smartart_in_scaled_group(package: &mut OpcPackage) {
    let xml = String::from_utf8(package.get_part(SLIDE_TWO_PART).unwrap().to_vec()).unwrap();
    let frame_start = xml.find("<p:graphicFrame>").unwrap();
    let frame_end = frame_start
        + xml[frame_start..].find("</p:graphicFrame>").unwrap()
        + "</p:graphicFrame>".len();
    let frame = xml[frame_start..frame_end].replacen(
        r#"<a:off x="0" y="0"/><a:ext cx="5080000" cy="2540000"/>"#,
        r#"<a:off x="127000" y="254000"/><a:ext cx="2540000" cy="1270000"/>"#,
        1,
    );
    assert!(frame.contains(r#"<a:off x="127000" y="254000"/>"#));
    let group = format!(
        r#"<p:grpSp><p:nvGrpSpPr><p:cNvPr id="420" name="Scaled SmartArt owner"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="10160000" cy="5080000"/><a:chOff x="0" y="0"/><a:chExt cx="5080000" cy="2540000"/></a:xfrm></p:grpSpPr>{frame}</p:grpSp>"#
    );
    let mut grouped = xml;
    grouped.replace_range(frame_start..frame_end, &group);
    package.set_part(SLIDE_TWO_PART, grouped.into_bytes());
}

fn smartart_transfer_source_package() -> OpcPackage {
    let mut package = smartart_fixture_package();
    package
        .get_or_create_part_rels(SLIDE_TWO_PART)
        .items
        .retain(|relationship| relationship.rel_type != rel_types::NOTES_SLIDE);
    package.parts.remove(NOTES_PART);
    package.part_rels.remove(NOTES_PART);
    package.content_types.overrides.remove(NOTES_PART);
    package
}

fn smartart_transfer_destination_package() -> OpcPackage {
    let bytes = Presentation::new().unwrap().to_bytes().unwrap();
    let mut package = OpcPackage::from_reader(Cursor::new(bytes)).unwrap();
    for (part, content_type) in [
        (
            "/custom/diagrams/two/data1.xml",
            "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml",
        ),
        (
            "/custom/diagrams/two/layout1.xml",
            "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml",
        ),
        (
            "/custom/diagrams/two/style1.xml",
            "application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml",
        ),
        (
            "/custom/diagrams/two/colors1.xml",
            "application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml",
        ),
        (
            "/custom/diagrams/two/drawing1.xml",
            "application/vnd.ms-office.drawingml.diagramDrawing+xml",
        ),
    ] {
        package.set_part(part, b"destination diagram decoy".to_vec());
        package.content_types.add_override(part, content_type);
    }
    package
}

fn add_smartart_scope(package: &mut OpcPackage, owner_part: &str, scope: &str, text: &str) {
    let base = format!("/custom/diagrams/{scope}");
    let data = format!("{base}/data1.xml");
    let layout = format!("{base}/layout1.xml");
    let style = format!("{base}/style1.xml");
    let colors = format!("{base}/colors1.xml");
    let drawing = format!("{base}/drawing1.xml");
    package.set_part(&data, smartart_data_xml(text));
    package.set_part(&layout, br#"<dgm:layoutDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" uniqueId="urn:unsupported"><dgm:layoutNode><dgm:alg type="snake"/><dgm:extLst><x:keep xmlns:x="urn:f219"/></dgm:extLst></dgm:layoutNode></dgm:layoutDef>"#.to_vec());
    package.set_part(&style, br#"<dgm:styleDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" uniqueId="urn:style"><dgm:styleLbl name="node0"><dgm:style val="moderate"/></dgm:styleLbl></dgm:styleDef>"#.to_vec());
    package.set_part(&colors, br#"<dgm:colorsDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" uniqueId="urn:colors"><dgm:styleLbl name="node0"><dgm:fillClrLst><a:srgbClr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" val="112233"/></dgm:fillClrLst></dgm:styleLbl></dgm:colorsDef>"#.to_vec());
    package.set_part(&drawing, br#"<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dsp:spTree><dsp:sp/><a:blip r:embed="image-link"/></dsp:spTree></dsp:drawing>"#.to_vec());
    for (part, content_type) in [
        (
            &data,
            "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml",
        ),
        (
            &layout,
            "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml",
        ),
        (
            &style,
            "application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml",
        ),
        (
            &colors,
            "application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml",
        ),
        (
            &drawing,
            "application/vnd.ms-office.drawingml.diagramDrawing+xml",
        ),
    ] {
        package.content_types.add_override(part, content_type);
    }
    let relative_base = format!("../diagrams/{scope}");
    let relationships = package.get_or_create_part_rels(owner_part);
    relationships.add_with_id(
        "smart-data",
        rel_types::DIAGRAM_DATA,
        &format!("{relative_base}/data1.xml"),
    );
    relationships.add_with_id(
        "smart-layout",
        rel_types::DIAGRAM_LAYOUT,
        &format!("{relative_base}/layout1.xml"),
    );
    relationships.add_with_id(
        "smart-style",
        rel_types::DIAGRAM_QUICK_STYLE,
        &format!("{relative_base}/style1.xml"),
    );
    relationships.add_with_id(
        "smart-colors",
        rel_types::DIAGRAM_COLORS,
        &format!("{relative_base}/colors1.xml"),
    );
    relationships.add_with_id(
        "drawing-link",
        rel_types::DIAGRAM_DRAWING,
        &format!("{relative_base}/drawing1.xml"),
    );
    package.get_or_create_part_rels(&drawing).add_with_id(
        "image-link",
        rel_types::IMAGE,
        "../../../ppt/media/image1.png",
    );
}

fn smartart_slide_xml() -> Vec<u8> {
    format!(r#"<p:sld xmlns:p="{P_NS}" xmlns:a="{A_NS}" xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:r="{R_NS}"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="42" name="SmartArt"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="5080000" cy="2540000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds r:dm="smart-data" r:lo="smart-layout" r:qs="smart-style" r:cs="smart-colors"/></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld></p:sld>"#).into_bytes()
}

fn smartart_layout_xml() -> Vec<u8> {
    smartart_owned_root_xml("sldLayout", 52, "Layout SmartArt", "")
}

fn smartart_master_xml() -> Vec<u8> {
    smartart_owned_root_xml(
        "sldMaster",
        62,
        "Master SmartArt",
        r#"<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst/><p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles>"#,
    )
}

fn smartart_owned_root_xml(root: &str, shape_id: u32, name: &str, tail: &str) -> Vec<u8> {
    format!(r#"<p:{root} xmlns:p="{P_NS}" xmlns:a="{A_NS}" xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:r="{R_NS}"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="{shape_id}" name="{name}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="5080000" cy="2540000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds r:dm="smart-data" r:lo="smart-layout" r:qs="smart-style" r:cs="smart-colors"/></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld>{tail}</p:{root}>"#).into_bytes()
}

fn smartart_data_xml(text: &str) -> Vec<u8> {
    format!(r#"<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="{A_NS}" xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"><dgm:ptLst><dgm:pt modelId="n1"><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>{text}</a:t></a:r></a:p></dgm:t></dgm:pt></dgm:ptLst><dgm:cxnLst/><dgm:extLst><a:ext uri="http://schemas.microsoft.com/office/drawing/2008/diagram"><dsp:dataModelExt relId="drawing-link" minVer="http://schemas.openxmlformats.org/drawingml/2006/diagram"/></a:ext><x:keep xmlns:x="urn:f219"/></dgm:extLst></dgm:dataModel>"#).into_bytes()
}
use std::sync::atomic::{AtomicUsize, Ordering};

#[test]
fn animated_export_samples_transitions_clicks_and_media_fallbacks_in_order() {
    use rpptx::{AnimationExportOptions, AnimationFormat, GifLoopBehavior, MediaFallbackPolicy};

    let (presentation, click_shape_id, media_shape_id) = animation_test_presentation();
    let before_click = presentation
        .render_media_timeline_deterministic(
            1,
            TimelinePosition {
                elapsed_ms: 150,
                click_count: 0,
            },
            Some(0),
            MediaFallbackPolicy::DeterministicPlaceholder,
        )
        .unwrap();
    let after_click = presentation
        .render_media_timeline_deterministic(
            1,
            TimelinePosition {
                elapsed_ms: 150,
                click_count: 1,
            },
            Some(0),
            MediaFallbackPolicy::DeterministicPlaceholder,
        )
        .unwrap();
    assert_eq!(before_click.media[0].phase, MediaPlaybackPhase::Stopped);
    assert_eq!(after_click.media[0].phase, MediaPlaybackPhase::Playing);
    assert!(
        !before_click
            .frame
            .state
            .shapes
            .contains_key(&click_shape_id)
    );
    assert!(!after_click.frame.state.shapes[&click_shape_id].visible);
    assert_eq!(before_click.media[0].shape_id, media_shape_id);
    let before_layout = oxml_layout::LayoutResult::new(
        vec![std::sync::Arc::new(before_click.frame.page)],
        Vec::new(),
        None,
        Vec::new(),
    );
    let after_layout = oxml_layout::LayoutResult::new(
        vec![std::sync::Arc::new(after_click.frame.page)],
        Vec::new(),
        None,
        Vec::new(),
    );
    assert_ne!(
        oxml_pdf::render_page_to_png(&before_layout, 0, 72.0),
        oxml_pdf::render_page_to_png(&after_layout, 0, 72.0)
    );
    let exported = presentation
        .export_animation_deterministic(
            &animation_test_segments(),
            AnimationExportOptions {
                frame_rate: 10,
                width_px: 96,
                height_px: 54,
                format: AnimationFormat::Gif {
                    loop_behavior: GifLoopBehavior::Once,
                },
                media_fallback: MediaFallbackPolicy::DeterministicPlaceholder,
            },
        )
        .expect("the approved animated facade must render a bounded source-built deck");

    assert_eq!(
        exported.frame_timestamps_ms,
        vec![0, 100, 200, 300, 400, 500]
    );
    assert!(exported.bytes.starts_with(b"GIF89a"));
    let messages = exported
        .diagnostics
        .iter()
        .map(|diagnostic| diagnostic.message.as_str())
        .collect::<Vec<_>>();
    assert_eq!(messages.len(), 12);
    for frame in messages.chunks_exact(2) {
        assert_eq!(
            frame,
            [
                format!("media shape {media_shape_id} rendered as deterministic Video placeholder"),
                format!(
                    "media shape {media_shape_id} has unsupported content type `video/x-f227-opaque`"
                ),
            ]
        );
    }
}

#[test]
fn public_facade_keeps_mixed_font_ids_and_dynamic_media_labels_coherent() {
    use rpptx::{AnimationExportOptions, AnimationFormat, AnimationSegment, AnimationTransition};

    let (mut presentation, _, _) = animation_test_presentation();
    for (slide_index, text, family) in [
        (0, "CaladeaWAVE", "Caladea"),
        (1, "Liberationmmmm", "Liberation Sans"),
    ] {
        let mut slide = presentation.slide_mut(slide_index).unwrap();
        let mut textbox = slide
            .add_textbox(Emu(4_000_000), Emu(500_000), Emu(4_500_000), Emu(900_000))
            .unwrap();
        textbox.set_text(text).unwrap();
        let mut frame = textbox.text_frame().unwrap();
        let mut paragraph = frame.paragraph_mut(0).unwrap();
        let mut run = paragraph.run_mut(0).unwrap();
        let mut properties = CT_TextCharacterProperties::default();
        properties.font_size = Some(3_200);
        run.set_properties(properties);
        run.set_font(Some(TextFont::new(family).unwrap()));
    }

    let segments = [
        AnimationSegment {
            slide_index: 0,
            duration_ms: 100,
            click_count: 0,
            transition: AnimationTransition::None,
        },
        AnimationSegment {
            slide_index: 1,
            duration_ms: 100,
            click_count: 1,
            transition: AnimationTransition::None,
        },
        AnimationSegment {
            slide_index: 0,
            duration_ms: 100,
            click_count: 0,
            transition: AnimationTransition::None,
        },
        AnimationSegment {
            slide_index: 1,
            duration_ms: 100,
            click_count: 1,
            transition: AnimationTransition::None,
        },
    ];
    let exported = presentation
        .export_animation_deterministic(
            &segments,
            AnimationExportOptions {
                frame_rate: 10,
                width_px: 320,
                height_px: 180,
                format: AnimationFormat::Gif {
                    loop_behavior: rpptx::GifLoopBehavior::Once,
                },
                media_fallback: MediaFallbackPolicy::DeterministicPlaceholder,
            },
        )
        .unwrap();
    let frames = decoded_gif_frames(&exported.bytes);
    assert_eq!(frames.len(), 4);
    assert_eq!(frames[0], frames[2], "Caladea pixels must remain exact");
    assert_eq!(
        frames[1], frames[3],
        "Liberation Sans and Carlito label pixels must remain exact"
    );
    assert_ne!(animation_hash(&frames[0]), animation_hash(&frames[1]));

    let poster = presentation
        .export_animation_deterministic(
            &segments,
            AnimationExportOptions {
                frame_rate: 10,
                width_px: 320,
                height_px: 180,
                format: AnimationFormat::Gif {
                    loop_behavior: rpptx::GifLoopBehavior::Once,
                },
                media_fallback: MediaFallbackPolicy::PosterFrame,
            },
        )
        .unwrap();
    let poster_frames = decoded_gif_frames(&poster.bytes);
    let media_region = (32, 48, 96, 64);
    assert_ne!(
        rgba_region(&frames[1], 320, media_region),
        rgba_region(&poster_frames[1], 320, media_region),
        "the on-demand Video label must replace the static poster pixels"
    );
    let avi = presentation
        .export_animation_deterministic(
            &segments,
            AnimationExportOptions {
                frame_rate: 10,
                width_px: 320,
                height_px: 180,
                format: AnimationFormat::MotionJpegAvi { quality: 80 },
                media_fallback: MediaFallbackPolicy::DeterministicPlaceholder,
            },
        )
        .unwrap();
    let avi_frames = parse_avi_manifest(&avi.bytes).decoded_frame_hashes;
    assert_eq!(avi_frames[0], avi_frames[2]);
    assert_eq!(avi_frames[1], avi_frames[3]);
    assert_ne!(avi_frames[0], avi_frames[1]);
}

fn decoded_gif_frames(bytes: &[u8]) -> Vec<Vec<u8>> {
    let mut options = gif::DecodeOptions::new();
    options.set_color_output(gif::ColorOutput::RGBA);
    let mut decoder = options.read_info(Cursor::new(bytes)).unwrap();
    let mut frames = Vec::new();
    while let Some(frame) = decoder.read_next_frame().unwrap() {
        frames.push(frame.buffer.to_vec());
    }
    frames
}

fn rgba_region(
    rgba: &[u8],
    image_width: usize,
    (x, y, width, height): (usize, usize, usize, usize),
) -> Vec<u8> {
    let mut region = Vec::with_capacity(width * height * 4);
    for row in y..y + height {
        let start = (row * image_width + x) * 4;
        region.extend_from_slice(&rgba[start..start + width * 4]);
    }
    region
}

fn animation_test_presentation() -> (Presentation, u32, u32) {
    let mut presentation = Presentation::new().unwrap();
    presentation.add_slide(0).unwrap();
    presentation.add_slide(0).unwrap();
    {
        let mut slide = presentation.slide_mut(0).unwrap();
        let mut outgoing = slide
            .add_shape("rect", Emu(0), Emu(0), Emu(9_144_000), Emu(5_143_500))
            .unwrap();
        outgoing.set_name("Outgoing red field").unwrap();
        outgoing
        .set_fill(
            Fill::from_xml(
                br#"<a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:srgbClr val="D63C32"/></a:solidFill>"#,
            )
            .unwrap(),
        )
        .unwrap();
    }
    {
        let mut slide = presentation.slide_mut(1).unwrap();
        let mut incoming = slide
            .add_shape("rect", Emu(0), Emu(0), Emu(9_144_000), Emu(5_143_500))
            .unwrap();
        incoming.set_name("Incoming blue field").unwrap();
        incoming
        .set_fill(
            Fill::from_xml(
                br#"<a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:srgbClr val="3278D6"/></a:solidFill>"#,
            )
            .unwrap(),
        )
        .unwrap();
    }
    {
        let mut slide = presentation.slide_mut(1).unwrap();
        let mut click_shape = slide
            .add_shape(
                "rect",
                Emu(5_943_600),
                Emu(3_200_400),
                Emu(1_828_800),
                Emu(914_400),
            )
            .unwrap();
        click_shape.set_name("Click-visible yellow marker").unwrap();
        click_shape
            .set_fill(
                Fill::from_xml(
                    br#"<a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:srgbClr val="F0C541"/></a:solidFill>"#,
                )
                .unwrap(),
            )
            .unwrap();
    }
    let poster = valid_one_pixel_png();
    presentation
        .add_media(
            1,
            MediaKind::Video,
            MediaSourceInput::Embedded(EmbeddedMediaInput {
                bytes: b"\0\0\0\x18ftypisom-f227-video",
                filename: "f227.mp4",
                content_type: "video/x-f227-opaque",
            }),
            MediaPoster {
                bytes: &poster,
                filename: "poster.png",
            },
            Emu(914_400),
            Emu(1_371_600),
            Emu(2_743_200),
            Emu(1_828_800),
            MediaPlaybackSettings {
                trigger: rpptx::MediaPlaybackTrigger::OnClick,
                ..MediaPlaybackSettings::default()
            },
        )
        .unwrap();
    let media_shape_id = presentation.media(1).unwrap()[0].shape_id;
    let bytes = presentation.to_bytes().unwrap();
    let mut package = open_opc(&bytes, "F-227 animated source");
    let presentation_part = package.main_document_part().unwrap();
    let model = CT_Presentation::from_xml(package.get_part(&presentation_part).unwrap()).unwrap();
    let slide_relationship = package
        .get_part_rels(&presentation_part)
        .unwrap()
        .get_by_id(&model.slide_ids[1].relationship_id)
        .unwrap();
    let slide_part = OpcPackage::resolve_rel_target(&presentation_part, &slide_relationship.target);
    let mut slide = CT_Slide::from_xml(package.get_part(&slide_part).unwrap()).unwrap();
    let click_shape_id = slide
        .common_slide_data
        .shape_tree
        .children
        .iter()
        .find(|child| child.non_visual_name().as_deref() == Some("Click-visible yellow marker"))
        .and_then(ShapeTreeChild::non_visual_id)
        .unwrap();
    slide.transition = Some(
        rpptx_oxml::timing::CT_SlideTransition::from_xml(
            format!(
                r#"<p:transition xmlns:p="{P_NS}" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" p14:dur="300"><p:fade/></p:transition>"#
            )
            .as_bytes(),
        )
        .unwrap(),
    );
    let timing = format!(
        r#"<p:timing xmlns:p="{P_NS}"><p:tnLst>
        <p:video><p:cMediaNode vol="100000" showWhenStopped="0"><p:cTn id="1" dur="indefinite"/><p:tgtEl><p:spTgt spid="{media_shape_id}"/></p:tgtEl></p:cMediaNode></p:video>
        <p:par><p:cTn id="2" dur="indefinite" nodeType="clickEffect"><p:stCondLst><p:cond evt="onClick" delay="0"/></p:stCondLst><p:childTnLst>
        <p:cmd type="call" cmd="playFrom(0.0)"><p:cBhvr><p:cTn id="3" dur="1" fill="hold"/><p:tgtEl><p:spTgt spid="{media_shape_id}"/></p:tgtEl></p:cBhvr></p:cmd>
        <p:set><p:cBhvr><p:cTn id="4" dur="1" fill="hold"/><p:tgtEl><p:spTgt spid="{click_shape_id}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="hidden"/></p:to></p:set>
        </p:childTnLst></p:cTn></p:par>
        </p:tnLst></p:timing>"#
    );
    slide.timing = Some(rpptx_oxml::timing::CT_Timing::from_xml(timing.as_bytes()).unwrap());
    package.set_part(&slide_part, slide.to_xml().unwrap());
    let presentation = Presentation::from_bytes(&package_bytes(package)).unwrap();
    assert!(presentation.validate().is_empty());
    (presentation, click_shape_id, media_shape_id)
}

fn animation_test_options(format: rpptx::AnimationFormat) -> rpptx::AnimationExportOptions {
    rpptx::AnimationExportOptions {
        frame_rate: 10,
        width_px: 96,
        height_px: 54,
        format,
        media_fallback: rpptx::MediaFallbackPolicy::PosterFrame,
    }
}

fn animation_test_avi_options(quality: u8) -> rpptx::AnimationExportOptions {
    rpptx::AnimationExportOptions {
        media_fallback: rpptx::MediaFallbackPolicy::DeterministicPlaceholder,
        ..animation_test_options(rpptx::AnimationFormat::MotionJpegAvi { quality })
    }
}

fn animation_test_segments() -> [rpptx::AnimationSegment; 2] {
    [
        rpptx::AnimationSegment {
            slide_index: 1,
            duration_ms: 300,
            click_count: 0,
            transition: rpptx::AnimationTransition::FromSlide(0),
        },
        rpptx::AnimationSegment {
            slide_index: 1,
            duration_ms: 300,
            click_count: 1,
            transition: rpptx::AnimationTransition::FromSlide(0),
        },
    ]
}

fn animation_hash(bytes: &[u8]) -> u64 {
    bytes.iter().fold(0xcbf2_9ce4_8422_2325, |hash, byte| {
        (hash ^ u64::from(*byte)).wrapping_mul(0x100_0000_01b3)
    })
}

#[derive(Clone, Debug, Eq, PartialEq)]
struct AnimationGoldenManifest {
    timestamps: Vec<u64>,
    frame_hashes: Vec<u64>,
    loop_repetitions: gif::Repeat,
    width: u16,
    height: u16,
    container_hash: u64,
}

fn decoded_gif_manifest(bytes: &[u8], timestamps: Vec<u64>) -> AnimationGoldenManifest {
    let mut options = gif::DecodeOptions::new();
    options.set_color_output(gif::ColorOutput::RGBA);
    let mut decoder = options.read_info(Cursor::new(bytes)).unwrap();
    let width = decoder.width();
    let height = decoder.height();
    let loop_repetitions = decoder.repeat();
    let mut frame_hashes = Vec::new();
    while let Some(frame) = decoder.read_next_frame().unwrap() {
        frame_hashes.push(animation_hash(&frame.buffer));
    }
    AnimationGoldenManifest {
        timestamps,
        frame_hashes,
        loop_repetitions,
        width,
        height,
        container_hash: animation_hash(bytes),
    }
}

#[derive(Clone, Copy, Debug)]
struct AviChunk {
    id: [u8; 4],
    size: usize,
    header_start: usize,
    data_start: usize,
    data_end: usize,
}

#[derive(Clone, Debug, Eq, PartialEq)]
struct AviManifest {
    frame_rate: u32,
    frame_count: u32,
    width: u32,
    height: u32,
    duration_ms: u64,
    payload_sizes: Vec<u32>,
    payload_hashes: Vec<u64>,
    decoded_frame_hashes: Vec<u64>,
}

#[derive(Clone, Debug, Eq, PartialEq)]
struct AviGoldenManifest {
    timestamps: Vec<u64>,
    avi: AviManifest,
    container_hash: u64,
    diagnostics: Vec<String>,
}

fn avi_u16(bytes: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes(bytes[offset..offset + 2].try_into().unwrap())
}

fn avi_u32(bytes: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes(bytes[offset..offset + 4].try_into().unwrap())
}

fn avi_i32(bytes: &[u8], offset: usize) -> i32 {
    i32::from_le_bytes(bytes[offset..offset + 4].try_into().unwrap())
}

fn avi_chunks(bytes: &[u8], start: usize, end: usize) -> Vec<AviChunk> {
    let mut cursor = start;
    let mut chunks = Vec::new();
    while cursor < end {
        assert!(cursor + 8 <= end, "truncated AVI chunk header");
        let id = bytes[cursor..cursor + 4].try_into().unwrap();
        let size = avi_u32(bytes, cursor + 4) as usize;
        let data_start = cursor + 8;
        let data_end = data_start.checked_add(size).unwrap();
        assert!(data_end <= end, "AVI chunk exceeds its containing list");
        let padded_end = data_end + size % 2;
        assert!(
            padded_end <= end,
            "AVI chunk padding exceeds its containing list"
        );
        if size % 2 == 1 {
            assert_eq!(bytes[data_end], 0, "AVI padding must be zero");
        }
        chunks.push(AviChunk {
            id,
            size,
            header_start: cursor,
            data_start,
            data_end,
        });
        cursor = padded_end;
    }
    assert_eq!(cursor, end, "AVI list must end on an exact chunk boundary");
    chunks
}

fn decoded_jpeg_pixel_hash(payload: &[u8], width: u32, height: u32) -> u64 {
    let info = oxml_media::probe(payload).expect("AVI frame must contain a valid JPEG");
    assert_eq!((info.width_px, info.height_px), (width, height));
    let page = PageFrame::new(
        1,
        f64::from(width),
        f64::from(height),
        vec![PositionedElement::Image {
            rect: Rect {
                x: 0.0,
                y: 0.0,
                width: f64::from(width),
                height: f64::from(height),
            },
            data: payload.to_vec(),
            content_type: "image/jpeg".to_owned(),
            media_id: MediaId::from_bytes(payload),
        }],
    );
    let layout = oxml_layout::LayoutResult::new(
        vec![std::sync::Arc::new(page)],
        Vec::new(),
        None,
        Vec::new(),
    );
    let png = oxml_pdf::render_page_to_png(&layout, 0, 72.0).unwrap();
    let rgba = tiny_skia::Pixmap::decode_png(&png).unwrap();
    assert_eq!((rgba.width(), rgba.height()), (width, height));
    animation_hash(rgba.data())
}

fn parse_avi_manifest(bytes: &[u8]) -> AviManifest {
    assert_eq!(&bytes[..4], b"RIFF");
    assert_eq!(avi_u32(bytes, 4) as usize, bytes.len() - 8);
    assert_eq!(&bytes[8..12], b"AVI ");
    let top = avi_chunks(bytes, 12, bytes.len());
    assert_eq!(top.len(), 3);
    assert_eq!(top[0].id, *b"LIST");
    assert_eq!(top[1].id, *b"LIST");
    assert_eq!(top[2].id, *b"idx1");
    assert_eq!(&bytes[top[0].data_start..top[0].data_start + 4], b"hdrl");
    assert_eq!(&bytes[top[1].data_start..top[1].data_start + 4], b"movi");
    assert_eq!(top[0].size, top[0].data_end - top[0].data_start);
    assert_eq!(top[1].size, top[1].data_end - top[1].data_start);

    let hdrl = avi_chunks(bytes, top[0].data_start + 4, top[0].data_end);
    assert_eq!(hdrl.len(), 2);
    assert_eq!(hdrl[0].id, *b"avih");
    assert_eq!(hdrl[0].size, 56);
    assert_eq!(hdrl[1].id, *b"LIST");
    assert_eq!(&bytes[hdrl[1].data_start..hdrl[1].data_start + 4], b"strl");
    let avih = hdrl[0].data_start;
    let microseconds_per_frame = avi_u32(bytes, avih);
    assert_eq!(avi_u32(bytes, avih + 4), 0);
    assert_eq!(avi_u32(bytes, avih + 8), 0);
    assert_eq!(avi_u32(bytes, avih + 12), 0x10);
    let frame_count = avi_u32(bytes, avih + 16);
    assert_eq!(avi_u32(bytes, avih + 20), 0);
    assert_eq!(avi_u32(bytes, avih + 24), 1);
    let avih_max_payload = avi_u32(bytes, avih + 28);
    let width = avi_u32(bytes, avih + 32);
    let height = avi_u32(bytes, avih + 36);
    assert_eq!(&bytes[avih + 40..avih + 56], &[0; 16]);

    let strl = avi_chunks(bytes, hdrl[1].data_start + 4, hdrl[1].data_end);
    assert_eq!(strl.len(), 2);
    assert_eq!(strl[0].id, *b"strh");
    assert_eq!(strl[0].size, 56);
    assert_eq!(strl[1].id, *b"strf");
    assert_eq!(strl[1].size, 40);
    let strh = strl[0].data_start;
    assert_eq!(&bytes[strh..strh + 4], b"vids");
    assert_eq!(&bytes[strh + 4..strh + 8], b"MJPG");
    assert_eq!(avi_u32(bytes, strh + 8), 0);
    assert_eq!(avi_u16(bytes, strh + 12), 0);
    assert_eq!(avi_u16(bytes, strh + 14), 0);
    assert_eq!(avi_u32(bytes, strh + 16), 0);
    let scale = avi_u32(bytes, strh + 20);
    let frame_rate = avi_u32(bytes, strh + 24);
    assert_eq!(avi_u32(bytes, strh + 28), 0);
    assert_eq!(avi_u32(bytes, strh + 32), frame_count);
    let strh_max_payload = avi_u32(bytes, strh + 36);
    assert_eq!(avi_u32(bytes, strh + 40), u32::MAX);
    assert_eq!(avi_u32(bytes, strh + 44), 0);
    assert_eq!(avi_u16(bytes, strh + 48), 0);
    assert_eq!(avi_u16(bytes, strh + 50), 0);
    assert_eq!(avi_u16(bytes, strh + 52), width as u16);
    assert_eq!(avi_u16(bytes, strh + 54), height as u16);
    assert_eq!(microseconds_per_frame, 1_000_000 / frame_rate);

    let strf = strl[1].data_start;
    assert_eq!(avi_u32(bytes, strf), 40);
    assert_eq!(avi_i32(bytes, strf + 4), width as i32);
    assert_eq!(avi_i32(bytes, strf + 8), height as i32);
    assert_eq!(avi_u16(bytes, strf + 12), 1);
    assert_eq!(avi_u16(bytes, strf + 14), 24);
    assert_eq!(&bytes[strf + 16..strf + 20], b"MJPG");
    assert_eq!(avi_u32(bytes, strf + 20), width * height * 3);
    assert_eq!(&bytes[strf + 24..strf + 40], &[0; 16]);

    let movi = avi_chunks(bytes, top[1].data_start + 4, top[1].data_end);
    assert_eq!(movi.len(), frame_count as usize);
    let index = top[2];
    assert_eq!(index.size % 16, 0);
    assert_eq!(index.size / 16, movi.len());
    let movi_type_position = top[1].data_start;
    let mut payload_sizes = Vec::new();
    let mut payload_hashes = Vec::new();
    let mut decoded_frame_hashes = Vec::new();
    for (frame_index, frame) in movi.iter().enumerate() {
        assert_eq!(frame.id, *b"00dc");
        let payload = &bytes[frame.data_start..frame.data_end];
        let entry = index.data_start + frame_index * 16;
        assert_eq!(&bytes[entry..entry + 4], b"00dc");
        assert_eq!(avi_u32(bytes, entry + 4), 0x10);
        assert_eq!(
            avi_u32(bytes, entry + 8) as usize,
            frame.header_start - movi_type_position
        );
        assert_eq!(avi_u32(bytes, entry + 12) as usize, frame.size);
        payload_sizes.push(frame.size as u32);
        payload_hashes.push(animation_hash(payload));
        decoded_frame_hashes.push(decoded_jpeg_pixel_hash(payload, width, height));
    }
    let max_payload = payload_sizes.iter().copied().max().unwrap();
    assert_eq!(avih_max_payload, max_payload);
    assert_eq!(strh_max_payload, max_payload);
    let duration_ms = u64::from(frame_count) * u64::from(scale) * 1_000 / u64::from(frame_rate);
    AviManifest {
        frame_rate,
        frame_count,
        width,
        height,
        duration_ms,
        payload_sizes,
        payload_hashes,
        decoded_frame_hashes,
    }
}

fn decoded_avi_golden_manifest(animation: &rpptx::DeterministicAnimation) -> AviGoldenManifest {
    AviGoldenManifest {
        timestamps: animation.frame_timestamps_ms.clone(),
        avi: parse_avi_manifest(&animation.bytes),
        container_hash: animation_hash(&animation.bytes),
        diagnostics: animation
            .diagnostics
            .iter()
            .map(|diagnostic| diagnostic.message.clone())
            .collect(),
    }
}

#[test]
fn animated_export_does_not_change_static_rendering() {
    let (presentation, _, _) = animation_test_presentation();
    let before = presentation.render_deterministic().unwrap().1;
    let before_pdf = oxml_pdf::render_to_pdf(&before);
    let before_png = oxml_pdf::render_page_to_png(&before, 0, 72.0).unwrap();
    presentation
        .export_animation_deterministic(
            &animation_test_segments(),
            animation_test_options(rpptx::AnimationFormat::Gif {
                loop_behavior: rpptx::GifLoopBehavior::Once,
            }),
        )
        .unwrap();
    let after = presentation.render_deterministic().unwrap().1;
    assert_eq!(oxml_pdf::render_to_pdf(&after), before_pdf);
    assert_eq!(
        oxml_pdf::render_page_to_png(&after, 0, 72.0).unwrap(),
        before_png
    );
}

#[test]
fn animated_gif_and_motion_jpeg_avi_match_the_reviewed_two_machine_manifest() {
    let (presentation, _, _) = animation_test_presentation();
    let gif = presentation
        .export_animation_deterministic(
            &animation_test_segments(),
            animation_test_options(rpptx::AnimationFormat::Gif {
                loop_behavior: rpptx::GifLoopBehavior::TotalPlays(
                    std::num::NonZeroU16::new(3).unwrap(),
                ),
            }),
        )
        .unwrap();
    let manifest = decoded_gif_manifest(&gif.bytes, gif.frame_timestamps_ms.clone());
    assert_eq!(
        manifest,
        AnimationGoldenManifest {
            timestamps: vec![0, 100, 200, 300, 400, 500],
            frame_hashes: vec![
                4_894_345_659_775_260_357,
                14_975_209_435_305_666_729,
                11_017_240_483_202_348_157,
                4_894_345_659_775_260_357,
                12_281_991_332_647_577_373,
                439_196_692_808_906_197,
            ],
            loop_repetitions: gif::Repeat::Finite(2),
            width: 96,
            height: 54,
            container_hash: 1_682_901_777_930_996_407,
        }
    );

    let avi = presentation
        .export_animation_deterministic(&animation_test_segments(), animation_test_avi_options(80))
        .unwrap();
    let avi_golden = decoded_avi_golden_manifest(&avi);
    assert_eq!(
        avi_golden,
        AviGoldenManifest {
            timestamps: vec![0, 100, 200, 300, 400, 500],
            avi: AviManifest {
                frame_rate: 10,
                frame_count: 6,
                width: 96,
                height: 54,
                duration_ms: 600,
                payload_sizes: vec![968, 974, 1_015, 968, 907, 908],
                payload_hashes: vec![
                    9_345_256_692_286_977_543,
                    2_927_541_614_763_063_172,
                    6_127_686_133_210_419_864,
                    9_345_256_692_286_977_543,
                    10_353_497_591_034_183_350,
                    7_116_127_623_826_450_847,
                ],
                decoded_frame_hashes: vec![
                    15_016_817_725_945_965_948,
                    7_383_799_236_496_809_804,
                    9_379_829_375_826_415_478,
                    15_016_817_725_945_965_948,
                    12_438_950_088_017_235_599,
                    3_796_532_621_390_706_249,
                ],
            },
            container_hash: 6_525_351_511_319_371_367,
            diagnostics: vec![
                "media shape 6 rendered as deterministic Video placeholder".to_owned(),
                "media shape 6 has unsupported content type `video/x-f227-opaque`".to_owned(),
                "media shape 6 rendered as deterministic Video placeholder".to_owned(),
                "media shape 6 has unsupported content type `video/x-f227-opaque`".to_owned(),
                "media shape 6 rendered as deterministic Video placeholder".to_owned(),
                "media shape 6 has unsupported content type `video/x-f227-opaque`".to_owned(),
                "media shape 6 rendered as deterministic Video placeholder".to_owned(),
                "media shape 6 has unsupported content type `video/x-f227-opaque`".to_owned(),
                "media shape 6 rendered as deterministic Video placeholder".to_owned(),
                "media shape 6 has unsupported content type `video/x-f227-opaque`".to_owned(),
                "media shape 6 rendered as deterministic Video placeholder".to_owned(),
                "media shape 6 has unsupported content type `video/x-f227-opaque`".to_owned(),
            ],
        }
    );

    let lower_quality = presentation
        .export_animation_deterministic(&animation_test_segments(), animation_test_avi_options(40))
        .unwrap();
    let lower_quality_manifest = parse_avi_manifest(&lower_quality.bytes);
    assert_eq!(
        (
            lower_quality_manifest.frame_rate,
            lower_quality_manifest.frame_count,
            lower_quality_manifest.width,
            lower_quality_manifest.height,
            lower_quality_manifest.duration_ms,
        ),
        (10, 6, 96, 54, 600)
    );
    assert_ne!(
        lower_quality_manifest.payload_hashes,
        avi_golden.avi.payload_hashes
    );
    assert_ne!(
        lower_quality_manifest.decoded_frame_hashes,
        avi_golden.avi.decoded_frame_hashes
    );
}

#[test]
fn the_animated_export_manifest_rejects_one_frame_timestamp_loop_and_dimension_mutations() {
    let (presentation, _, _) = animation_test_presentation();
    let exported = presentation
        .export_animation_deterministic(
            &animation_test_segments(),
            animation_test_options(rpptx::AnimationFormat::Gif {
                loop_behavior: rpptx::GifLoopBehavior::Infinite,
            }),
        )
        .unwrap();
    let manifest = decoded_gif_manifest(&exported.bytes, exported.frame_timestamps_ms);
    let mut mutated = manifest.clone();
    mutated.frame_hashes[0] ^= 1;
    assert_ne!(mutated, manifest);
    let mut mutated = manifest.clone();
    mutated.timestamps[1] += 1;
    assert_ne!(mutated, manifest);
    let mut mutated = manifest.clone();
    mutated.loop_repetitions = gif::Repeat::Finite(1);
    assert_ne!(mutated, manifest);
    let mut mutated = manifest.clone();
    mutated.width += 1;
    assert_ne!(mutated, manifest);
    let mut mutated = manifest.clone();
    mutated.height += 1;
    assert_ne!(mutated, manifest);

    let avi = presentation
        .export_animation_deterministic(&animation_test_segments(), animation_test_avi_options(80))
        .unwrap();
    let avi_manifest = decoded_avi_golden_manifest(&avi);
    let mut mutated = avi_manifest.clone();
    mutated.container_hash ^= 1;
    assert_ne!(mutated, avi_manifest);
    for index in 0..avi_manifest.avi.payload_hashes.len() {
        let mut mutated = avi_manifest.clone();
        mutated.avi.payload_hashes[index] ^= 1;
        assert_ne!(mutated, avi_manifest);
    }
    for index in 0..avi_manifest.avi.decoded_frame_hashes.len() {
        let mut mutated = avi_manifest.clone();
        mutated.avi.decoded_frame_hashes[index] ^= 1;
        assert_ne!(mutated, avi_manifest);
    }
    for index in 0..avi_manifest.diagnostics.len() {
        let mut mutated = avi_manifest.clone();
        mutated.diagnostics[index].push_str(" mutation");
        assert_ne!(mutated, avi_manifest);
    }
    let mut reordered = avi_manifest.clone();
    reordered.diagnostics.swap(0, 1);
    assert_ne!(reordered, avi_manifest);
}

#[test]
fn add_replace_extract_and_remove_embedded_media_are_atomic() {
    use rpptx::{
        EmbeddedMediaInput, MediaDiagnostic, MediaKind, MediaLocation, MediaPlaybackSettings,
        MediaPlaybackTrigger, MediaPoster, MediaSourceInput,
    };

    let mut presentation = Presentation::from_bytes(&fixture_bytes()).unwrap();
    let poster = valid_one_pixel_png();
    let initial = presentation.to_bytes().unwrap();
    assert!(
        presentation
            .add_media(
                0,
                MediaKind::Video,
                MediaSourceInput::Embedded(EmbeddedMediaInput {
                    bytes: b"not an mp4",
                    filename: "broken.mp4",
                    content_type: "video/mp4",
                }),
                MediaPoster {
                    bytes: &poster,
                    filename: "poster.png",
                },
                Emu(1),
                Emu(2),
                Emu(300),
                Emu(200),
                MediaPlaybackSettings::default(),
            )
            .is_err()
    );
    assert_eq!(presentation.to_bytes().unwrap(), initial);

    let audio = b"ID3\x04\0\0opaque-mp3-payload";
    let settings = MediaPlaybackSettings {
        trim_start_ms: Some(875),
        trim_end_ms: Some(125),
        volume: 80_000,
        looped: true,
        show_when_stopped: true,
        trigger: MediaPlaybackTrigger::Automatic,
    };
    presentation
        .add_media(
            0,
            MediaKind::Audio,
            MediaSourceInput::Embedded(EmbeddedMediaInput {
                bytes: audio,
                filename: "sample.mp3",
                content_type: "audio/mpeg",
            }),
            MediaPoster {
                bytes: &poster,
                filename: "poster.png",
            },
            Emu(10),
            Emu(20),
            Emu(300),
            Emu(200),
            settings,
        )
        .unwrap();
    let first = presentation.media(0).unwrap().remove(0);
    assert_eq!(first.kind, MediaKind::Audio);
    assert_eq!(first.settings, settings);
    assert!(first.diagnostics.is_empty());
    assert_eq!(
        presentation.extract_media(0, first.shape_id).unwrap(),
        Some(audio.to_vec())
    );
    let authored_package =
        OpcPackage::from_reader(Cursor::new(presentation.to_bytes().unwrap())).unwrap();
    let authored_slide =
        String::from_utf8(authored_package.get_part(SLIDE_TWO_PART).unwrap().to_vec()).unwrap();
    assert!(authored_slide.contains(r#"<p14:trim st="875" end="125"/>"#));
    assert!(!authored_slide.contains("trimStart="));
    assert!(!authored_slide.contains("trimEnd="));
    let first_part = match &first.source {
        MediaLocation::Embedded { part_name, .. } => part_name.clone(),
        MediaLocation::Linked { .. } => panic!("expected embedded audio"),
    };

    presentation
        .add_media(
            0,
            MediaKind::Audio,
            MediaSourceInput::Embedded(EmbeddedMediaInput {
                bytes: audio,
                filename: "same.mp3",
                content_type: "audio/mpeg",
            }),
            MediaPoster {
                bytes: &poster,
                filename: "same-poster.png",
            },
            Emu(40),
            Emu(50),
            Emu(300),
            Emu(200),
            MediaPlaybackSettings::default(),
        )
        .unwrap();
    let media = presentation.media(0).unwrap();
    let second = media
        .iter()
        .find(|item| item.shape_id != first.shape_id)
        .unwrap();
    assert_eq!(
        second.source,
        MediaLocation::Embedded {
            part_name: first_part.clone(),
            content_type: "audio/mpeg".to_owned(),
        }
    );

    let before_failed_replace = presentation.to_bytes().unwrap();
    assert!(
        presentation
            .replace_media(
                0,
                first.shape_id,
                MediaSourceInput::Embedded(EmbeddedMediaInput {
                    bytes: b"wrong signature",
                    filename: "wrong.wav",
                    content_type: "audio/wav",
                }),
            )
            .is_err()
    );
    assert_eq!(presentation.to_bytes().unwrap(), before_failed_replace);

    let opaque = b"unsupported codec bytes stay exact";
    presentation
        .replace_media(
            0,
            first.shape_id,
            MediaSourceInput::Embedded(EmbeddedMediaInput {
                bytes: opaque,
                filename: "opaque.cstm",
                content_type: "application/x-example-media",
            }),
        )
        .unwrap();
    let replaced = presentation
        .media(0)
        .unwrap()
        .into_iter()
        .find(|item| item.shape_id == first.shape_id)
        .unwrap();
    assert_eq!(replaced.settings, settings);
    assert_eq!(
        replaced.poster_relationship_id,
        first.poster_relationship_id
    );
    assert_eq!(
        replaced.diagnostics,
        vec![MediaDiagnostic::UnsupportedContentType {
            content_type: "application/x-example-media".to_owned(),
        }]
    );
    assert_eq!(
        presentation.extract_media(0, first.shape_id).unwrap(),
        Some(opaque.to_vec())
    );

    presentation
        .add_media(
            0,
            MediaKind::Video,
            MediaSourceInput::Linked {
                target: "https://example.invalid/movie.webm?sig=A%2FB",
                content_type: "video/webm",
            },
            MediaPoster {
                bytes: &poster,
                filename: "linked-poster.png",
            },
            Emu(70),
            Emu(80),
            Emu(300),
            Emu(200),
            MediaPlaybackSettings {
                trigger: MediaPlaybackTrigger::InClickSequence,
                ..MediaPlaybackSettings::default()
            },
        )
        .unwrap();
    let linked = presentation
        .media(0)
        .unwrap()
        .into_iter()
        .find(|item| matches!(item.source, MediaLocation::Linked { .. }))
        .unwrap();
    assert_eq!(
        linked.source,
        MediaLocation::Linked {
            target: "https://example.invalid/movie.webm?sig=A%2FB".to_owned(),
        }
    );
    assert_eq!(
        linked.settings.trigger,
        MediaPlaybackTrigger::InClickSequence
    );
    assert_eq!(
        presentation.extract_media(0, linked.shape_id).unwrap(),
        None
    );

    let reopened = Presentation::from_bytes(&presentation.to_bytes().unwrap()).unwrap();
    assert_eq!(reopened.media(0).unwrap(), presentation.media(0).unwrap());
    assert_eq!(
        reopened.extract_media(0, first.shape_id).unwrap(),
        Some(opaque.to_vec())
    );

    let second_shape_id = second.shape_id;
    presentation.remove_media(0, first.shape_id).unwrap();
    assert_eq!(
        presentation.extract_media(0, second_shape_id).unwrap(),
        Some(audio.to_vec())
    );
    let package = OpcPackage::from_reader(Cursor::new(presentation.to_bytes().unwrap())).unwrap();
    assert!(package.get_part(&first_part).is_some());
    presentation.remove_media(0, second_shape_id).unwrap();
    let package = OpcPackage::from_reader(Cursor::new(presentation.to_bytes().unwrap())).unwrap();
    assert!(package.get_part(&first_part).is_none());
    presentation.remove_media(0, linked.shape_id).unwrap();
    assert!(presentation.media(0).unwrap().is_empty());
}

#[test]
fn media_replacement_switches_office_source_kind_in_both_directions() {
    use rpptx::{EmbeddedMediaInput, MediaKind, MediaLocation, MediaPoster, MediaSourceInput};
    use rpptx_oxml::picture::MediaSource as PictureMediaSource;

    let poster = valid_one_pixel_png();
    let original = b"ID3original-replacement-audio";
    let replacement = b"ID3embedded-after-link";
    let linked_target = "https://example.invalid/replacement.mp3?token=A%2FB";
    let mut presentation = Presentation::from_bytes(&fixture_bytes()).unwrap();
    presentation
        .add_media(
            0,
            MediaKind::Audio,
            MediaSourceInput::Embedded(EmbeddedMediaInput {
                bytes: original,
                filename: "original.mp3",
                content_type: "audio/mpeg",
            }),
            MediaPoster {
                bytes: &poster,
                filename: "poster.png",
            },
            Emu(1),
            Emu(2),
            Emu(30),
            Emu(20),
            rpptx::MediaPlaybackSettings::default(),
        )
        .unwrap();
    let shape_id = presentation.media(0).unwrap()[0].shape_id;

    presentation
        .replace_media(
            0,
            shape_id,
            MediaSourceInput::Linked {
                target: linked_target,
                content_type: "audio/mpeg",
            },
        )
        .unwrap();
    assert_eq!(
        presentation.media(0).unwrap()[0].source,
        MediaLocation::Linked {
            target: linked_target.to_owned()
        }
    );
    assert_eq!(presentation.extract_media(0, shape_id).unwrap(), None);
    let linked_package =
        OpcPackage::from_reader(Cursor::new(presentation.to_bytes().unwrap())).unwrap();
    let linked_slide =
        CT_Slide::from_xml(linked_package.get_part(SLIDE_TWO_PART).unwrap()).unwrap();
    let linked_picture = linked_slide
        .common_slide_data
        .shape_tree
        .children
        .iter()
        .find(|child| child.non_visual_id() == Some(shape_id))
        .and_then(|child| match child {
            ShapeTreeChild::Picture(picture) => Some(picture),
            _ => None,
        })
        .unwrap();
    let PictureMediaSource::Linked {
        relationship_id: linked_office_id,
    } = &linked_picture.media.as_ref().unwrap().source
    else {
        panic!("expected linked Office media source");
    };
    assert_eq!(linked_picture.office_media_relationship_ids().len(), 1);
    let linked_relationships = linked_package.get_part_rels(SLIDE_TWO_PART).unwrap();
    for relationship_id in [
        linked_picture.standard_media_relationship_id().unwrap(),
        linked_office_id,
    ] {
        assert_eq!(
            linked_relationships
                .get_by_id(relationship_id)
                .unwrap()
                .target_mode
                .as_deref(),
            Some("External")
        );
    }

    presentation
        .replace_media(
            0,
            shape_id,
            MediaSourceInput::Embedded(EmbeddedMediaInput {
                bytes: replacement,
                filename: "replacement.mp3",
                content_type: "audio/mpeg",
            }),
        )
        .unwrap();
    assert_eq!(
        presentation.extract_media(0, shape_id).unwrap().as_deref(),
        Some(replacement.as_slice())
    );
    let embedded_package =
        OpcPackage::from_reader(Cursor::new(presentation.to_bytes().unwrap())).unwrap();
    let embedded_slide =
        CT_Slide::from_xml(embedded_package.get_part(SLIDE_TWO_PART).unwrap()).unwrap();
    let embedded_picture = embedded_slide
        .common_slide_data
        .shape_tree
        .children
        .iter()
        .find(|child| child.non_visual_id() == Some(shape_id))
        .and_then(|child| match child {
            ShapeTreeChild::Picture(picture) => Some(picture),
            _ => None,
        })
        .unwrap();
    let PictureMediaSource::Embedded {
        relationship_id: embedded_office_id,
    } = &embedded_picture.media.as_ref().unwrap().source
    else {
        panic!("expected embedded Office media source");
    };
    assert_eq!(embedded_picture.office_media_relationship_ids().len(), 1);
    let embedded_relationships = embedded_package.get_part_rels(SLIDE_TWO_PART).unwrap();
    assert!(
        embedded_relationships
            .items
            .iter()
            .all(|relationship| relationship.target != linked_target)
    );
    for relationship_id in [
        embedded_picture.standard_media_relationship_id().unwrap(),
        embedded_office_id,
    ] {
        assert_eq!(
            embedded_relationships
                .get_by_id(relationship_id)
                .unwrap()
                .target_mode,
            None
        );
    }
}

#[test]
fn media_replacement_preserves_unrelated_raw_standard_relationship_references() {
    use rpptx::{EmbeddedMediaInput, MediaKind, MediaLocation, MediaPoster, MediaSourceInput};

    let poster = valid_one_pixel_png();
    let original = b"ID3raw-reference-original";
    let linked_target = "https://example.invalid/raw-reference-replacement.mp3";
    let mut authored = Presentation::from_bytes(&fixture_bytes()).unwrap();
    authored
        .add_media(
            0,
            MediaKind::Audio,
            MediaSourceInput::Embedded(EmbeddedMediaInput {
                bytes: original,
                filename: "original.mp3",
                content_type: "audio/mpeg",
            }),
            MediaPoster {
                bytes: &poster,
                filename: "poster.png",
            },
            Emu(1),
            Emu(2),
            Emu(30),
            Emu(20),
            rpptx::MediaPlaybackSettings::default(),
        )
        .unwrap();
    let media = authored.media(0).unwrap().remove(0);
    let shape_id = media.shape_id;
    let old_part = match media.source {
        MediaLocation::Embedded { part_name, .. } => part_name,
        MediaLocation::Linked { .. } => panic!("expected embedded media"),
    };

    let mut package = OpcPackage::from_reader(Cursor::new(authored.to_bytes().unwrap())).unwrap();
    let slide = CT_Slide::from_xml(package.get_part(SLIDE_TWO_PART).unwrap()).unwrap();
    let picture = slide
        .common_slide_data
        .shape_tree
        .children
        .iter()
        .find(|child| child.non_visual_id() == Some(shape_id))
        .and_then(|child| match child {
            ShapeTreeChild::Picture(picture) => Some(picture),
            _ => None,
        })
        .unwrap();
    let old_standard_id = picture.standard_media_relationship_id().unwrap().to_owned();
    let marker = format!(r#"<a:audioFile r:link="{old_standard_id}"/>"#);
    let raw_reference = format!(
        r#"<a:extLst data-f215='keep'><a:ext uri="{{F215-RAW-RELATIONSHIP}}"><a:hlinkClick r:id="{old_standard_id}"/></a:ext></a:extLst>"#
    );
    let slide_xml = String::from_utf8(package.get_part(SLIDE_TWO_PART).unwrap().to_vec()).unwrap();
    assert_eq!(slide_xml.matches(&marker).count(), 1);
    package.set_part(
        SLIDE_TWO_PART,
        slide_xml
            .replacen(&marker, &format!("{marker}{raw_reference}"), 1)
            .into_bytes(),
    );

    let mut presentation = Presentation::from_bytes(&package_bytes(package)).unwrap();
    presentation
        .replace_media(
            0,
            shape_id,
            MediaSourceInput::Linked {
                target: linked_target,
                content_type: "audio/mpeg",
            },
        )
        .unwrap();
    assert_eq!(
        presentation.media(0).unwrap()[0].source,
        MediaLocation::Linked {
            target: linked_target.to_owned(),
        }
    );
    let replaced = OpcPackage::from_reader(Cursor::new(presentation.to_bytes().unwrap())).unwrap();
    let replaced_slide =
        String::from_utf8(replaced.get_part(SLIDE_TWO_PART).unwrap().to_vec()).unwrap();
    assert!(replaced_slide.contains(&raw_reference));
    assert!(
        replaced
            .get_part_rels(SLIDE_TWO_PART)
            .unwrap()
            .get_by_id(&old_standard_id)
            .is_some()
    );
    assert_eq!(replaced.get_part(&old_part), Some(original.as_slice()));
}

#[test]
fn standard_only_media_replacement_rewrites_only_the_standard_relationship() {
    use rpptx::{EmbeddedMediaInput, MediaKind, MediaLocation, MediaPoster, MediaSourceInput};

    let poster = valid_one_pixel_png();
    let original = b"ID3standard-only-original";
    let replacement = b"ID3standard-only-replacement";
    let linked_target = "https://example.invalid/standard-only.mp3?token=A%2FB";
    let mut authored = Presentation::from_bytes(&fixture_bytes()).unwrap();
    authored
        .add_media(
            0,
            MediaKind::Audio,
            MediaSourceInput::Embedded(EmbeddedMediaInput {
                bytes: original,
                filename: "original.mp3",
                content_type: "audio/mpeg",
            }),
            MediaPoster {
                bytes: &poster,
                filename: "poster.png",
            },
            Emu(1),
            Emu(2),
            Emu(30),
            Emu(20),
            rpptx::MediaPlaybackSettings::default(),
        )
        .unwrap();
    let shape_id = authored.media(0).unwrap()[0].shape_id;
    let mut package = OpcPackage::from_reader(Cursor::new(authored.to_bytes().unwrap())).unwrap();
    let slide = CT_Slide::from_xml(package.get_part(SLIDE_TWO_PART).unwrap()).unwrap();
    let picture = slide
        .common_slide_data
        .shape_tree
        .children
        .iter()
        .find(|child| child.non_visual_id() == Some(shape_id))
        .and_then(|child| match child {
            ShapeTreeChild::Picture(picture) => Some(picture),
            _ => None,
        })
        .unwrap();
    let office_id = picture.office_media_relationship_ids()[0].clone();
    let extension = format!(
        r#"<p:extLst><p:ext uri="{{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}}"><p14:media r:embed="{office_id}"/></p:ext></p:extLst>"#
    );
    let slide_xml = String::from_utf8(package.get_part(SLIDE_TWO_PART).unwrap().to_vec()).unwrap();
    assert_eq!(slide_xml.matches(&extension).count(), 1);
    package.set_part(
        SLIDE_TWO_PART,
        slide_xml.replacen(&extension, "", 1).into_bytes(),
    );
    package
        .part_rels
        .get_mut(SLIDE_TWO_PART)
        .unwrap()
        .items
        .retain(|relationship| relationship.id != office_id);

    let mut presentation = Presentation::from_bytes(&package_bytes(package)).unwrap();
    assert_eq!(
        presentation.extract_media(0, shape_id).unwrap().as_deref(),
        Some(original.as_slice())
    );
    presentation
        .replace_media(
            0,
            shape_id,
            MediaSourceInput::Linked {
                target: linked_target,
                content_type: "audio/mpeg",
            },
        )
        .unwrap();
    assert_eq!(
        presentation.media(0).unwrap()[0].source,
        MediaLocation::Linked {
            target: linked_target.to_owned()
        }
    );
    assert_eq!(presentation.extract_media(0, shape_id).unwrap(), None);
    let linked_package =
        OpcPackage::from_reader(Cursor::new(presentation.to_bytes().unwrap())).unwrap();
    let linked_relationships = linked_package.get_part_rels(SLIDE_TWO_PART).unwrap();
    assert!(
        linked_relationships
            .get_all_by_type(rel_types::POWERPOINT_MEDIA)
            .is_empty()
    );
    let linked_standard = linked_relationships.get_all_by_type(rel_types::AUDIO);
    assert_eq!(linked_standard.len(), 1);
    assert_eq!(linked_standard[0].target, linked_target);
    assert_eq!(linked_standard[0].target_mode.as_deref(), Some("External"));
    assert!(
        !String::from_utf8(linked_package.get_part(SLIDE_TWO_PART).unwrap().to_vec())
            .unwrap()
            .contains("p14:media")
    );

    let before_failed_replace = presentation.to_bytes().unwrap();
    assert!(
        presentation
            .replace_media(
                0,
                shape_id,
                MediaSourceInput::Embedded(EmbeddedMediaInput {
                    bytes: b"not a wave",
                    filename: "invalid.wav",
                    content_type: "audio/wav",
                }),
            )
            .is_err()
    );
    assert_eq!(presentation.to_bytes().unwrap(), before_failed_replace);

    presentation
        .replace_media(
            0,
            shape_id,
            MediaSourceInput::Embedded(EmbeddedMediaInput {
                bytes: replacement,
                filename: "replacement.mp3",
                content_type: "audio/mpeg",
            }),
        )
        .unwrap();
    assert_eq!(
        presentation.extract_media(0, shape_id).unwrap().as_deref(),
        Some(replacement.as_slice())
    );
    let embedded_package =
        OpcPackage::from_reader(Cursor::new(presentation.to_bytes().unwrap())).unwrap();
    let embedded_relationships = embedded_package.get_part_rels(SLIDE_TWO_PART).unwrap();
    assert!(
        embedded_relationships
            .get_all_by_type(rel_types::POWERPOINT_MEDIA)
            .is_empty()
    );
    let embedded_standard = embedded_relationships.get_all_by_type(rel_types::AUDIO);
    assert_eq!(embedded_standard.len(), 1);
    assert_eq!(embedded_standard[0].target_mode, None);
    assert!(
        !String::from_utf8(embedded_package.get_part(SLIDE_TWO_PART).unwrap().to_vec())
            .unwrap()
            .contains("p14:media")
    );
}

#[test]
fn dual_source_removal_drops_both_office_relationships_and_keeps_shared_payloads() {
    use rpptx::{EmbeddedMediaInput, MediaKind, MediaLocation, MediaPoster, MediaSourceInput};

    let payload = b"ID3dual-source-shared-audio";
    let poster = valid_one_pixel_png();
    let mut presentation = Presentation::from_bytes(&fixture_bytes()).unwrap();
    for left in [Emu(1), Emu(40)] {
        presentation
            .add_media(
                0,
                MediaKind::Audio,
                MediaSourceInput::Embedded(EmbeddedMediaInput {
                    bytes: payload,
                    filename: "dual.mp3",
                    content_type: "audio/mpeg",
                }),
                MediaPoster {
                    bytes: &poster,
                    filename: "poster.png",
                },
                left,
                Emu(2),
                Emu(30),
                Emu(20),
                rpptx::MediaPlaybackSettings::default(),
            )
            .unwrap();
    }
    let media = presentation.media(0).unwrap();
    let first_shape_id = media[0].shape_id;
    let second_shape_id = media[1].shape_id;
    let shared_part = match &media[0].source {
        MediaLocation::Embedded { part_name, .. } => part_name.clone(),
        MediaLocation::Linked { .. } => panic!("expected embedded media"),
    };
    let mut package =
        OpcPackage::from_reader(Cursor::new(presentation.to_bytes().unwrap())).unwrap();
    let slide = CT_Slide::from_xml(package.get_part(SLIDE_TWO_PART).unwrap()).unwrap();
    let first_picture = slide
        .common_slide_data
        .shape_tree
        .children
        .iter()
        .find(|child| child.non_visual_id() == Some(first_shape_id))
        .and_then(|child| match child {
            ShapeTreeChild::Picture(picture) => Some(picture),
            _ => None,
        })
        .unwrap();
    let embedded_office_id = first_picture.office_media_relationship_ids()[0].clone();
    let first_standard_id = first_picture
        .standard_media_relationship_id()
        .unwrap()
        .to_owned();
    let first_poster_id = first_picture
        .media
        .as_ref()
        .unwrap()
        .poster_relationship_id
        .as_ref()
        .unwrap()
        .clone();
    let linked_office_id = package
        .part_rels
        .get_mut(SLIDE_TWO_PART)
        .unwrap()
        .add_external(
            rel_types::POWERPOINT_MEDIA,
            "https://example.invalid/producer-link.mp3",
        );
    let slide_xml = String::from_utf8(package.get_part(SLIDE_TWO_PART).unwrap().to_vec()).unwrap();
    let source_attribute = format!(r#"r:embed="{embedded_office_id}""#);
    assert_eq!(slide_xml.matches(&source_attribute).count(), 1);
    let dual_attribute = format!(r#"r:embed="{embedded_office_id}" r:link="{linked_office_id}""#);
    package.set_part(
        SLIDE_TWO_PART,
        slide_xml
            .replacen(&source_attribute, &dual_attribute, 1)
            .into_bytes(),
    );

    let dual_bytes = package_bytes(package);
    let replacement = b"ID3dual-source-replacement-audio";
    let mut replacement_case = Presentation::from_bytes(&dual_bytes).unwrap();
    replacement_case
        .replace_media(
            0,
            first_shape_id,
            MediaSourceInput::Embedded(EmbeddedMediaInput {
                bytes: replacement,
                filename: "replacement.mp3",
                content_type: "audio/mpeg",
            }),
        )
        .unwrap();
    assert_eq!(
        replacement_case
            .extract_media(0, first_shape_id)
            .unwrap()
            .as_deref(),
        Some(replacement.as_slice())
    );
    let replaced_package =
        OpcPackage::from_reader(Cursor::new(replacement_case.to_bytes().unwrap())).unwrap();
    let replaced_relationships = replaced_package.get_part_rels(SLIDE_TWO_PART).unwrap();
    for relationship_id in [&embedded_office_id, &linked_office_id, &first_standard_id] {
        assert!(replaced_relationships.get_by_id(relationship_id).is_none());
    }
    assert!(replaced_relationships.get_by_id(&first_poster_id).is_some());
    assert!(replaced_package.get_part(&shared_part).is_some());

    let mut dual = Presentation::from_bytes(&dual_bytes).unwrap();
    assert_eq!(
        dual.media(0).unwrap()[0].source,
        MediaLocation::Linked {
            target: "https://example.invalid/producer-link.mp3".to_owned()
        }
    );
    dual.remove_media(0, first_shape_id).unwrap();
    let after_first = OpcPackage::from_reader(Cursor::new(dual.to_bytes().unwrap())).unwrap();
    let relationships = after_first.get_part_rels(SLIDE_TWO_PART).unwrap();
    for relationship_id in [
        embedded_office_id,
        linked_office_id,
        first_standard_id,
        first_poster_id,
    ] {
        assert!(relationships.get_by_id(&relationship_id).is_none());
    }
    assert!(after_first.get_part(&shared_part).is_some());
    assert_eq!(
        dual.extract_media(0, second_shape_id).unwrap().as_deref(),
        Some(payload.as_slice())
    );

    dual.remove_media(0, second_shape_id).unwrap();
    let after_second = OpcPackage::from_reader(Cursor::new(dual.to_bytes().unwrap())).unwrap();
    assert!(after_second.get_part(&shared_part).is_none());
}

#[test]
fn embedded_audio_and_video_corpus_media_round_trip_without_duplication() {
    use rpptx::{MediaKind, MediaLocation, MediaPlaybackSettings, MediaPlaybackTrigger};

    let corpus = corpus_dir();
    let cases = [
        (
            corpus.join("EmbeddedAudio.pptx"),
            MediaKind::Audio,
            4,
            "/ppt/media/media1.mp3",
            "audio/mpeg",
            rel_types::AUDIO,
            br#"<a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{9A81383C-4007-750D-FE55-453D231AD88A}"/>"#.as_slice(),
        ),
        (
            corpus.join("EmbeddedVideo.pptx"),
            MediaKind::Video,
            2,
            "/ppt/media/media1.mp4",
            "video/mp4",
            rel_types::VIDEO,
            br#"<a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{A48A45DC-D4F4-4A17-9406-AF61E7DDE9CC}"/>"#.as_slice(),
        ),
    ];
    if cases.iter().any(|case| !case.0.is_file()) {
        assert!(
            std::env::var_os("RDOCX_PPTX_CORPUS_REQUIRED").is_none(),
            "required embedded-media corpus decks are missing under {}",
            corpus.display()
        );
        return;
    }
    for (path, kind, shape_id, media_part, content_type, standard_type, creation_id) in cases {
        let source = fs::read(&path).unwrap();
        let original_package = OpcPackage::from_reader(Cursor::new(&source)).unwrap();
        let presentation = Presentation::from_bytes(&source).unwrap();
        let media = presentation.media(0).unwrap();
        assert_eq!(media.len(), 1, "{}", path.display());
        assert_eq!(media[0].shape_id, shape_id, "{}", path.display());
        assert_eq!(media[0].kind, kind, "{}", path.display());
        assert_eq!(
            media[0].source,
            MediaLocation::Embedded {
                part_name: media_part.to_owned(),
                content_type: content_type.to_owned(),
            },
            "{}",
            path.display()
        );
        assert_eq!(
            media[0].poster_relationship_id.as_deref(),
            Some("rId4"),
            "{}",
            path.display()
        );
        assert_eq!(
            media[0].settings,
            MediaPlaybackSettings {
                trim_start_ms: None,
                trim_end_ms: None,
                volume: 80_000,
                looped: false,
                show_when_stopped: false,
                trigger: MediaPlaybackTrigger::InClickSequence,
            },
            "{}",
            path.display()
        );
        assert!(media[0].diagnostics.is_empty(), "{}", path.display());
        let expected_media_bytes = original_package.get_part(media_part).unwrap();
        assert_eq!(
            presentation.extract_media(0, shape_id).unwrap().as_deref(),
            Some(expected_media_bytes),
            "{}",
            path.display()
        );
        let original_relationships = original_package
            .get_part_rels("/ppt/slides/slide1.xml")
            .unwrap();
        for (id, expected_type, expected_target) in [
            ("rId1", rel_types::POWERPOINT_MEDIA, &media_part[5..]),
            ("rId2", standard_type, &media_part[5..]),
            ("rId4", rel_types::IMAGE, "../media/image1.png"),
        ] {
            let relationship = original_relationships.get_by_id(id).unwrap();
            assert_eq!(relationship.rel_type, expected_type, "{}", path.display());
            let expected_target = if id == "rId4" {
                expected_target.to_owned()
            } else {
                format!("../{}", expected_target)
            };
            assert_eq!(relationship.target, expected_target, "{}", path.display());
        }
        let saved = presentation.to_bytes().unwrap();
        let saved_package = OpcPackage::from_reader(Cursor::new(&saved)).unwrap();
        let reopened = Presentation::from_bytes(&saved).unwrap();
        assert_eq!(
            reopened.extract_media(0, shape_id).unwrap().as_deref(),
            Some(expected_media_bytes),
            "{}",
            path.display()
        );
        let saved_media_parts = saved_package
            .parts
            .keys()
            .filter(|part| part.starts_with("/ppt/media/"))
            .count();
        assert_eq!(saved_media_parts, 2, "{}", path.display());
        let slide = saved_package.get_part("/ppt/slides/slide1.xml").unwrap();
        assert!(
            slide
                .windows(creation_id.len())
                .any(|value| value == creation_id)
        );
        let saved_relationships = saved_package
            .get_part_rels("/ppt/slides/slide1.xml")
            .unwrap();
        for (id, expected_type, expected_target) in [
            (
                "rId1",
                rel_types::POWERPOINT_MEDIA,
                format!("../{}", &media_part[5..]),
            ),
            ("rId2", standard_type, format!("../{}", &media_part[5..])),
            ("rId4", rel_types::IMAGE, "../media/image1.png".to_owned()),
        ] {
            let relationship = saved_relationships.get_by_id(id).unwrap();
            assert_eq!(relationship.rel_type, expected_type, "{}", path.display());
            assert_eq!(relationship.target, expected_target, "{}", path.display());
        }
        assert_eq!(
            saved_package.content_types.content_type_for(media_part),
            Some(content_type)
        );
    }
}

#[test]
fn linked_media_keeps_external_targets_and_never_fetches_them() {
    use rpptx::{
        MediaKind, MediaLocation, MediaPlaybackSettings, MediaPlaybackTrigger, MediaPoster,
        MediaSourceInput,
    };

    let mut presentation = Presentation::from_bytes(&fixture_bytes()).unwrap();
    let poster = valid_one_pixel_png();
    let target = "https://127.0.0.1:9/media/movie.webm?token=A%2FB";
    let settings = MediaPlaybackSettings {
        trim_start_ms: Some(250),
        trim_end_ms: Some(1_750),
        volume: 65_000,
        looped: true,
        show_when_stopped: true,
        trigger: MediaPlaybackTrigger::InClickSequence,
    };
    presentation
        .add_media(
            0,
            MediaKind::Video,
            MediaSourceInput::Linked {
                target,
                content_type: "video/webm",
            },
            MediaPoster {
                bytes: &poster,
                filename: "poster.png",
            },
            Emu(1),
            Emu(2),
            Emu(30),
            Emu(20),
            settings,
        )
        .unwrap();
    let before_save = presentation.media(0).unwrap();
    let bytes = presentation.to_bytes().unwrap();
    let reopened = Presentation::from_bytes(&bytes).unwrap();
    let media = reopened.media(0).unwrap().remove(0);
    assert_eq!(reopened.media(0).unwrap(), before_save);
    assert_eq!(
        media.source,
        MediaLocation::Linked {
            target: target.to_owned()
        }
    );
    assert!(media.poster_relationship_id.is_some());
    assert_eq!(media.settings, settings);
    assert!(media.diagnostics.is_empty());
    assert_eq!(reopened.extract_media(0, media.shape_id).unwrap(), None);
    let package = OpcPackage::from_reader(Cursor::new(bytes)).unwrap();
    let relationships = package.get_part_rels(SLIDE_TWO_PART).unwrap();
    let external = relationships
        .items
        .iter()
        .filter(|relationship| relationship.target == target)
        .collect::<Vec<_>>();
    assert_eq!(external.len(), 2);
    assert!(
        external
            .iter()
            .all(|relationship| { relationship.target_mode.as_deref() == Some("External") })
    );
}

#[test]
fn unsupported_codec_bytes_remain_packaged_extractable_and_diagnosable() {
    use rpptx::{EmbeddedMediaInput, MediaDiagnostic, MediaKind, MediaPoster, MediaSourceInput};

    let payload = b"opaque unsupported codec payload";
    let poster = valid_one_pixel_png();
    let mut presentation = Presentation::from_bytes(&fixture_bytes()).unwrap();
    presentation
        .add_media(
            0,
            MediaKind::Audio,
            MediaSourceInput::Embedded(EmbeddedMediaInput {
                bytes: payload,
                filename: "payload.xyzcodec",
                content_type: "audio/x-example-codec",
            }),
            MediaPoster {
                bytes: &poster,
                filename: "poster.png",
            },
            Emu(1),
            Emu(2),
            Emu(30),
            Emu(20),
            rpptx::MediaPlaybackSettings::default(),
        )
        .unwrap();
    let reopened = Presentation::from_bytes(&presentation.to_bytes().unwrap()).unwrap();
    let media = reopened.media(0).unwrap().remove(0);
    assert_eq!(
        media.diagnostics,
        vec![MediaDiagnostic::UnsupportedContentType {
            content_type: "audio/x-example-codec".to_owned()
        }]
    );
    assert_eq!(
        reopened.extract_media(0, media.shape_id).unwrap(),
        Some(payload.to_vec())
    );
}

#[test]
fn media_dedup_requires_compatible_extension_and_content_type() {
    use rpptx::{EmbeddedMediaInput, MediaKind, MediaLocation, MediaPoster, MediaSourceInput};

    let poster = valid_one_pixel_png();
    let mut presentation = Presentation::from_bytes(&fixture_bytes()).unwrap();
    presentation
        .add_media(
            0,
            MediaKind::Audio,
            MediaSourceInput::Embedded(EmbeddedMediaInput {
                bytes: &poster,
                filename: "opaque.binmedia",
                content_type: "application/x-opaque-media",
            }),
            MediaPoster {
                bytes: &poster,
                filename: "poster.png",
            },
            Emu(1),
            Emu(2),
            Emu(30),
            Emu(20),
            rpptx::MediaPlaybackSettings::default(),
        )
        .unwrap();
    presentation
        .add_media(
            0,
            MediaKind::Audio,
            MediaSourceInput::Embedded(EmbeddedMediaInput {
                bytes: &poster,
                filename: "opaque.binmedia",
                content_type: "audio/x-different-opaque",
            }),
            MediaPoster {
                bytes: &poster,
                filename: "poster.png",
            },
            Emu(40),
            Emu(2),
            Emu(30),
            Emu(20),
            rpptx::MediaPlaybackSettings::default(),
        )
        .unwrap();

    let media = presentation.media(0).unwrap();
    let (first_part, second_part) = match (&media[0].source, &media[1].source) {
        (
            MediaLocation::Embedded {
                part_name: first,
                content_type: first_type,
            },
            MediaLocation::Embedded {
                part_name: second,
                content_type: second_type,
            },
        ) => {
            assert_eq!(first_type, "application/x-opaque-media");
            assert_eq!(second_type, "audio/x-different-opaque");
            (first.clone(), second.clone())
        }
        _ => panic!("expected two embedded media sources"),
    };
    assert_ne!(first_part, second_part);
    assert!(first_part.ends_with(".binmedia"));
    assert!(second_part.ends_with(".binmedia"));

    let package = OpcPackage::from_reader(Cursor::new(presentation.to_bytes().unwrap())).unwrap();
    let poster_part = package
        .get_part_rels(SLIDE_TWO_PART)
        .unwrap()
        .get_by_id(media[0].poster_relationship_id.as_deref().unwrap())
        .map(|relationship| OpcPackage::resolve_rel_target(SLIDE_TWO_PART, &relationship.target))
        .unwrap();
    assert_ne!(first_part, poster_part);
    assert_eq!(
        package.content_types.content_type_for(&poster_part),
        Some("image/png")
    );
}

#[test]
fn invalid_media_content_types_fail_atomically() {
    use rpptx::{EmbeddedMediaInput, MediaKind, MediaPoster, MediaSourceInput};

    let poster = valid_one_pixel_png();
    let mut presentation = Presentation::from_bytes(&fixture_bytes()).unwrap();
    for content_type in ["/", "audio /mpeg", "not a/type with spaces"] {
        let before = presentation.to_bytes().unwrap();
        assert!(
            presentation
                .add_media(
                    0,
                    MediaKind::Audio,
                    MediaSourceInput::Embedded(EmbeddedMediaInput {
                        bytes: b"opaque",
                        filename: "opaque.bin",
                        content_type,
                    }),
                    MediaPoster {
                        bytes: &poster,
                        filename: "poster.png",
                    },
                    Emu(1),
                    Emu(2),
                    Emu(30),
                    Emu(20),
                    rpptx::MediaPlaybackSettings::default(),
                )
                .is_err(),
            "{content_type}"
        );
        assert_eq!(presentation.to_bytes().unwrap(), before, "{content_type}");
    }
    let before = presentation.to_bytes().unwrap();
    assert!(
        presentation
            .add_media(
                0,
                MediaKind::Audio,
                MediaSourceInput::Linked {
                    target: "https://example.invalid/audio",
                    content_type: "audio /mpeg",
                },
                MediaPoster {
                    bytes: &poster,
                    filename: "poster.png",
                },
                Emu(1),
                Emu(2),
                Emu(30),
                Emu(20),
                rpptx::MediaPlaybackSettings::default(),
            )
            .is_err()
    );
    assert_eq!(presentation.to_bytes().unwrap(), before);
}

#[test]
fn known_media_content_types_are_case_insensitive_for_signature_validation() {
    use rpptx::{EmbeddedMediaInput, MediaKind, MediaLocation, MediaPoster, MediaSourceInput};

    let poster = valid_one_pixel_png();
    let mut presentation = Presentation::from_bytes(&fixture_bytes()).unwrap();
    let before = presentation.to_bytes().unwrap();
    assert!(
        presentation
            .add_media(
                0,
                MediaKind::Audio,
                MediaSourceInput::Embedded(EmbeddedMediaInput {
                    bytes: b"not an mp3",
                    filename: "invalid.mp3",
                    content_type: "Audio/MPEG",
                }),
                MediaPoster {
                    bytes: &poster,
                    filename: "poster.png",
                },
                Emu(1),
                Emu(2),
                Emu(30),
                Emu(20),
                rpptx::MediaPlaybackSettings::default(),
            )
            .is_err()
    );
    assert_eq!(presentation.to_bytes().unwrap(), before);

    let valid = b"ID3\x04\0\0uppercase-mime";
    presentation
        .add_media(
            0,
            MediaKind::Audio,
            MediaSourceInput::Embedded(EmbeddedMediaInput {
                bytes: valid,
                filename: "valid.mp3",
                content_type: "Audio/MPEG",
            }),
            MediaPoster {
                bytes: &poster,
                filename: "poster.png",
            },
            Emu(1),
            Emu(2),
            Emu(30),
            Emu(20),
            rpptx::MediaPlaybackSettings::default(),
        )
        .unwrap();
    let media = presentation.media(0).unwrap().remove(0);
    let MediaLocation::Embedded { content_type, .. } = &media.source else {
        panic!("expected embedded media");
    };
    assert_eq!(content_type, "Audio/MPEG");
    assert_eq!(
        presentation.extract_media(0, media.shape_id).unwrap(),
        Some(valid.to_vec())
    );
}

#[test]
fn media_removal_deletes_only_parts_owned_by_the_removed_relationships() {
    use rpptx::{EmbeddedMediaInput, MediaKind, MediaLocation, MediaPoster, MediaSourceInput};

    let payload = b"ID3shared-audio";
    let poster = valid_one_pixel_png();
    let mut presentation = Presentation::from_bytes(&fixture_bytes()).unwrap();
    for left in [Emu(1), Emu(40)] {
        presentation
            .add_media(
                0,
                MediaKind::Audio,
                MediaSourceInput::Embedded(EmbeddedMediaInput {
                    bytes: payload,
                    filename: "shared.mp3",
                    content_type: "audio/mpeg",
                }),
                MediaPoster {
                    bytes: &poster,
                    filename: "poster.png",
                },
                left,
                Emu(2),
                Emu(30),
                Emu(20),
                rpptx::MediaPlaybackSettings::default(),
            )
            .unwrap();
    }
    let media = presentation.media(0).unwrap();
    let part = match &media[0].source {
        MediaLocation::Embedded { part_name, .. } => part_name.clone(),
        MediaLocation::Linked { .. } => panic!("expected embedded media"),
    };
    assert_eq!(media[1].source, media[0].source);
    presentation.remove_media(0, media[0].shape_id).unwrap();
    let package = OpcPackage::from_reader(Cursor::new(presentation.to_bytes().unwrap())).unwrap();
    assert!(package.get_part(&part).is_some());
    presentation.remove_media(0, media[1].shape_id).unwrap();
    let package = OpcPackage::from_reader(Cursor::new(presentation.to_bytes().unwrap())).unwrap();
    assert!(package.get_part(&part).is_none());
}

#[test]
fn replace_and_remove_preserve_shared_relationship_references() {
    use std::collections::HashMap;

    use rpptx::{EmbeddedMediaInput, MediaKind, MediaPoster, MediaSourceInput};
    use rpptx_oxml::relmap::rewrite_exact_rel_ids;

    let poster = valid_one_pixel_png();
    let payload = b"ID3shared-relationship-audio";
    let mut presentation = Presentation::from_bytes(&fixture_bytes()).unwrap();
    for left in [Emu(1), Emu(40)] {
        presentation
            .add_media(
                0,
                MediaKind::Audio,
                MediaSourceInput::Embedded(EmbeddedMediaInput {
                    bytes: payload,
                    filename: "shared.mp3",
                    content_type: "audio/mpeg",
                }),
                MediaPoster {
                    bytes: &poster,
                    filename: "poster.png",
                },
                left,
                Emu(2),
                Emu(30),
                Emu(20),
                rpptx::MediaPlaybackSettings::default(),
            )
            .unwrap();
    }
    let media = presentation.media(0).unwrap();
    let first_shape_id = media[0].shape_id;
    let second_shape_id = media[1].shape_id;
    let mut package =
        OpcPackage::from_reader(Cursor::new(presentation.to_bytes().unwrap())).unwrap();
    let slide = CT_Slide::from_xml(package.get_part(SLIDE_TWO_PART).unwrap()).unwrap();
    let pictures = slide
        .common_slide_data
        .shape_tree
        .children
        .iter()
        .filter_map(|child| match child {
            ShapeTreeChild::Picture(picture) if picture.media.is_some() => Some(picture),
            _ => None,
        })
        .collect::<Vec<_>>();
    let first_standard = pictures[0]
        .standard_media_relationship_id()
        .unwrap()
        .to_owned();
    let first_microsoft = pictures[0]
        .media
        .as_ref()
        .unwrap()
        .source
        .relationship_id()
        .to_owned();
    let first_poster = pictures[0]
        .media
        .as_ref()
        .unwrap()
        .poster_relationship_id
        .as_ref()
        .unwrap()
        .to_owned();
    let second_standard = pictures[1]
        .standard_media_relationship_id()
        .unwrap()
        .to_owned();
    let second_microsoft = pictures[1]
        .media
        .as_ref()
        .unwrap()
        .source
        .relationship_id()
        .to_owned();
    let second_poster = pictures[1]
        .media
        .as_ref()
        .unwrap()
        .poster_relationship_id
        .as_ref()
        .unwrap()
        .to_owned();
    let replacements = HashMap::from([
        (second_standard.clone(), first_standard.clone()),
        (second_microsoft.clone(), first_microsoft.clone()),
        (second_poster.clone(), first_poster.clone()),
    ]);
    let shared_slide =
        rewrite_exact_rel_ids(package.get_part(SLIDE_TWO_PART).unwrap(), &replacements).unwrap();
    package.set_part(SLIDE_TWO_PART, shared_slide);
    package
        .part_rels
        .get_mut(SLIDE_TWO_PART)
        .unwrap()
        .items
        .retain(|relationship| {
            !matches!(
                relationship.id.as_str(),
                id if id == second_standard || id == second_microsoft || id == second_poster
            )
        });
    let shared_bytes = package_bytes(package);

    let mut replacement_case = Presentation::from_bytes(&shared_bytes).unwrap();
    let replacement = b"ID3replacement-audio";
    replacement_case
        .replace_media(
            0,
            first_shape_id,
            MediaSourceInput::Embedded(EmbeddedMediaInput {
                bytes: replacement,
                filename: "replacement.mp3",
                content_type: "audio/mpeg",
            }),
        )
        .unwrap();
    assert_eq!(
        replacement_case
            .extract_media(0, first_shape_id)
            .unwrap()
            .as_deref(),
        Some(replacement.as_slice())
    );
    assert_eq!(
        replacement_case
            .extract_media(0, second_shape_id)
            .unwrap()
            .as_deref(),
        Some(payload.as_slice())
    );

    let mut removal_case = Presentation::from_bytes(&shared_bytes).unwrap();
    removal_case.remove_media(0, first_shape_id).unwrap();
    assert_eq!(
        removal_case
            .extract_media(0, second_shape_id)
            .unwrap()
            .as_deref(),
        Some(payload.as_slice())
    );
    let saved = OpcPackage::from_reader(Cursor::new(removal_case.to_bytes().unwrap())).unwrap();
    let relationships = saved.get_part_rels(SLIDE_TWO_PART).unwrap();
    for id in [first_standard, first_microsoft, first_poster] {
        assert!(relationships.get_by_id(&id).is_some(), "{id}");
    }
}

#[test]
fn duplicated_media_slides_rewrite_every_retained_relationship_id() {
    use rpptx::{EmbeddedMediaInput, MediaKind, MediaLocation, MediaPoster, MediaSourceInput};

    let payload = b"ID3duplicate-audio";
    let poster = valid_one_pixel_png();
    let mut presentation = Presentation::from_bytes(&fixture_bytes()).unwrap();
    presentation
        .add_media(
            0,
            MediaKind::Audio,
            MediaSourceInput::Embedded(EmbeddedMediaInput {
                bytes: payload,
                filename: "duplicate.mp3",
                content_type: "audio/mpeg",
            }),
            MediaPoster {
                bytes: &poster,
                filename: "poster.png",
            },
            Emu(1),
            Emu(2),
            Emu(30),
            Emu(20),
            rpptx::MediaPlaybackSettings::default(),
        )
        .unwrap();
    presentation.duplicate_slide(0).unwrap();
    let first = presentation.media(0).unwrap().remove(0);
    let duplicate = presentation.media(1).unwrap().remove(0);
    assert_ne!(first.shape_id, duplicate.shape_id);
    assert_eq!(
        presentation.extract_media(0, first.shape_id).unwrap(),
        Some(payload.to_vec())
    );
    assert_eq!(
        presentation.extract_media(1, duplicate.shape_id).unwrap(),
        Some(payload.to_vec())
    );
    let (
        MediaLocation::Embedded {
            part_name: first_part,
            ..
        },
        MediaLocation::Embedded {
            part_name: duplicate_part,
            ..
        },
    ) = (&first.source, &duplicate.source)
    else {
        panic!("expected embedded media on both slides");
    };
    assert_eq!(first_part, duplicate_part);
    let reopened = Presentation::from_bytes(&presentation.to_bytes().unwrap()).unwrap();
    assert_eq!(reopened.media(0).unwrap(), vec![first]);
    assert_eq!(reopened.media(1).unwrap(), vec![duplicate]);
}

use oxml_chart::{AxisData, CT_ChartSpace};
use oxml_drawing::color::ColorMap;
use oxml_drawing::theme::CT_OfficeStyleSheet;
use oxml_layout::{MediaId, PageFrame, Paint, PositionedElement, Rect, walk};
use oxml_opc::relationship::rel_types;
use oxml_opc::{OpcPackage, content_types};
use rpptx::{
    Angle, CT_LineProperties, CT_TextCharacterProperties, CT_TextParagraphProperties, ChartData,
    ChartKind, ConnectorType, EmbeddedContentKind, EmbeddedMediaInput, EmbeddedMutationPolicy,
    EmbeddedSignatureState, Emu, Error, Fill, HandoutLayout, MediaDiagnostic, MediaFallbackPolicy,
    MediaKind, MediaPlaybackPhase, MediaPlaybackSettings, MediaPoster, MediaSourceInput,
    Presentation, PresentationPackageClass, ShapeKind, ShapeRef, TextBullet, TextBulletCharacter,
    TextBulletChoice, TextFont, TimelinePosition,
};
use rpptx_layout::{
    FlattenedItem, ResolveCtx, ResolvedContent, ResolvedSlide, ResolvedTextBody, ResolvedTextRun,
};
use rpptx_oxml::graphic_frame::GraphicDataPayload;
use rpptx_oxml::notes_parts::{CT_NotesMaster, CT_NotesSlide};
use rpptx_oxml::placeholder::PhType;
use rpptx_oxml::presentation::CT_Presentation;
use rpptx_oxml::shape_tree::ShapeTreeChild;
use rpptx_oxml::slide_parts::{
    BackgroundRendering, CT_Slide, CT_SlideLayout, CT_SlideMaster, ColorMapOverrideKind,
};

#[path = "../examples/dump_deck.rs"]
mod dump_deck;
#[allow(dead_code)]
#[path = "../examples/render_deck.rs"]
mod render_deck_example;

const P_NS: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const A_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MC_NS: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const PRESENTATION_PART: &str = "/custom/presentation-main.xml";
const SLIDE_ONE_PART: &str = "/custom/slides/first.xml";
const SLIDE_TWO_PART: &str = "/custom/slides/second.xml";
const NOTES_PART: &str = "/custom/notes/speaker.xml";
const LAYOUT_PART: &str = "/custom/layouts/validation.xml";
const POWERPOINT_VERSION: &str = "16.104";
const POWERPOINT_BUILD: &str = "16.104.25121423";
const POWERPOINT_APP_BUILD: &str = "1214";
const F214_POWERPOINT_ORACLE_WIDTH: u32 = 1920;
const F214_POWERPOINT_ORACLE_HEIGHT: u32 = 1080;
const F214_POWERPOINT_COMPARISON_WIDTH: u32 = 2001;
const F214_POWERPOINT_COMPARISON_HEIGHT: u32 = 1125;
const F214_POWERPOINT_COMPARISON_DPI: f64 = 150.0;
const F214_POWERPOINT_FOREGROUND_CHANNEL_CUTOFF: u8 = 240;
const F214_POWERPOINT_ORACLE_PROVENANCE: [&str; 8] = [
    "capture_method\tPowerPoint save as movie",
    "frame_extractor\tAVAssetImageGenerator",
    "requested_time_tolerance_before_ms\t0",
    "requested_time_tolerance_after_ms\t0",
    "normalization_operation\toxml-pdf image render with tiny-skia bilinear filtering",
    "normalization_input_px\t1920x1080",
    "normalization_output_px\t2001x1125",
    "normalization_dpi\t150",
];
const F214_POWERPOINT_ORACLE_CLASSIFICATION: &str =
    "PowerPoint reference frame, no divergence allowed";
const F214_POWERPOINT_ORACLE_MOVIE_SHA256_ENV: &str = "RDOCX_PPTX_TIMELINE_ORACLE_MOVIE_SHA256";
const F214_POWERPOINT_FRAME_EXTRACTOR_SWIFT: &str = r#"import AVFoundation
import AppKit

let movie = URL(fileURLWithPath: CommandLine.arguments[1])
let output = URL(fileURLWithPath: CommandLine.arguments[2])
let timeValue = CMTimeValue(CommandLine.arguments[3])!
let timeScale = CMTimeScale(CommandLine.arguments[4])!
guard timeScale > 0 else {
    throw NSError(domain: "rpptx.timeline.oracle", code: 2, userInfo: [NSLocalizedDescriptionKey: "movie sample timescale must be positive"])
}
let asset = AVURLAsset(url: movie)
let generator = AVAssetImageGenerator(asset: asset)
generator.appliesPreferredTrackTransform = true
generator.requestedTimeToleranceBefore = .zero
generator.requestedTimeToleranceAfter = .zero
let requestedTime = CMTime(value: timeValue, timescale: timeScale)
var actualTime = CMTime.invalid
let image = try generator.copyCGImage(at: requestedTime, actualTime: &actualTime)
guard CMTimeCompare(actualTime, requestedTime) == 0 else {
    throw NSError(domain: "rpptx.timeline.oracle", code: 1, userInfo: [NSLocalizedDescriptionKey: "requested \(requestedTime.value)/\(requestedTime.timescale), received \(actualTime.value)/\(actualTime.timescale)"])
}
let bitmap = NSBitmapImageRep(cgImage: image)
let png = bitmap.representation(using: .png, properties: [:])!
try png.write(to: output)
print(image.width, image.height, requestedTime.value, requestedTime.timescale, actualTime.value, actualTime.timescale)
"#;
const F214_POWERPOINT_ORACLE_CASES: [(&str, usize, u64, u32, i64, i32); 9] = [
    ("timeline-appear-active", 0, 330, 0, 198, 600),
    ("timeline-wipe-mid", 0, 825, 0, 495, 600),
    ("timeline-opacity-mid", 0, 1_320, 0, 792, 600),
    ("timeline-parallel-mid", 0, 1_980, 0, 1_188, 600),
    ("timeline-exit-mid", 0, 2_805, 0, 1_683, 600),
    (
        "ordinary-transition-outgoing-terminal",
        1,
        u64::MAX,
        u32::MAX,
        4_198,
        600,
    ),
    ("ordinary-transition", 2, 425, 0, 4_455, 600),
    ("morph", 3, 700, 0, 6_336, 600),
    ("push-transition", 4, 264, 0, 8_118, 600),
];
const KEYNOTE_VERSION: &str = "14.4";
const KEYNOTE_BUILD: &str = "7043.0.93";
const LIBREOFFICE_VERSION: &str = "LibreOffice 26.2.5.2 cd7284b4cbbfeb507e630c1aac019f4157393acb";
const F116_CANDIDATE_PATH: &str = "/private/tmp/rdocx-f116-m11-write-api.pptx";
const F124_CANDIDATE_PATH: &str = "/private/tmp/rdocx-f124-add-chart.pptx";
const F221_CANDIDATE_PATH: &str = "/private/tmp/rdocx-f221-agile-powerpoint.pptx";
const F221_POWERPOINT_PASSWORD: &str = "F-221 PowerPoint oracle";
const F221_ARTIFACT_SHA256: &str =
    "a0d33171c63ec084231daeef3b35718f5a2d709a5c92c9c0e2017ccaf9fa52d6";
const F124_ARTIFACT_SHA256: &str =
    "e6e9f7eef1c774d0414c5d0c3f1202da1a28635b5d089e15455b7adc3f66cb00";
const F116_ARTIFACT_SHA256: &str =
    "d36da6e8849eabd4487d2572baea19c3716ee7d0fe03aaa4714a28ce3c41de4f";
const F116_FINAL_TITLES: [&str; 10] = [
    "F-116 slide 10",
    "F-116 slide 02",
    "F-116 slide 03",
    "F-116 slide 04",
    "F-116 slide 05",
    "F-116 slide 06",
    "F-116 slide 07",
    "F-116 slide 08",
    "F-116 slide 08",
    "F-116 slide 09",
];
static F116_TEMP_FILE_COUNTER: AtomicUsize = AtomicUsize::new(0);

#[cfg(feature = "digital-signatures")]
const F221_PRIVATE_KEY_PKCS8_BASE64: &str = "MIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQCXVziWlLINGRFmqL/FTgWy1/o2zEkmQOHt23YQZDm7PiiBjq+ib85OF6UH9nH9VTr3rE4njpTj0TgzbTaY0ftxQqQg+sdSLZu6+FGAo+erEQAEoa9HZ6ys/9cuFbeYLUVVSQ33wtdq+VMDWMZqXxXd53QX9aqxBg1cBY4XAHgqWQ50X3AQ/IBhMTpZ0Wte3wGZeC+i1NQ45Ws9HwpwTgjg1mX3BTUaPtZW801SBOC2b3ug1jGJXO/fGigdzDwGcJKAXt/NZwa+PIhKv8GyebstGwCEu/detSmaODbMs7cYb4liJK2NLCY99Y7BLO4xMNg2hwploUwAdu+RgEf5U8LvAgMBAAECggEAEg5S7wxAjfV+sPvTHWwom+TOsnj/BTRagDFdzajXhnJtDMAETmH+gCysAN4zTWE8zs3c6TVGqEOO6/vMtsDeue2UfWbOHwzX9p+nwaxMeIlnsiXELsW8wUso1hO7Osmz6u/zXar+XoHumIif65L6neX+YNlriwFI2MDE6hOhQpP8eqA+u8DX1g2Oq0AhpPX6g5ABC422GjB6Z/NsXnwrFb4SpUy7aQNCHjDBQsetc2uE9hStvGvfwg4qvQl4NtL2k2udU/wOO66jvLBMbsF6Blz5vpljGpDuYJk/9sTTmzH+cJNob8nlOWwXVwNrAPMHe0l4RK8+DUqWvdEqkxZBoQKBgQDLOmt9DPX80aY4QucFNOGjjqH+Gk+6jEI6LrKfXiCswCiziWhoVCSKkQSi4k2Ol0NEKfcKy5jUYxglwjLDF/MBqz0rCKX/XjZ/hO8f8DlQFeZdLxvsFVH/lUl98JIifrrpMnO9za/0SfcMsy3CCglz8ho9+LUch8fQx4mTX3MCfwKBgQC+o5VqOZIJJXuYq6nWdjj5BPWKC8kdNq+t2Z5FnlsDuwNmSzCJI8juUPyId3pI0fiSbmdTfp8z4OrV46FrkFJ/oY8JJ5X3s2sfvviSVgvgZ/G+wJ4fQpN9yRg4Dp/RCgT7xrqxXtLTpjl7X0RrKri2uxyhh7AU/9mEuviwmTInkQKBgBed8l/V4cA/nNFs9Ovl+VLIgIrHA/zpz8hzJM7gYWux6Qj0Lu3w2U5BDAjhw6GOcoK5XbwjbN9BpMy+hKenYNYQ0Erv9lp22F55VFCh2gc0hFDP6K7Gy4CoGKJKErFviMkQ0+J6xLfe4JbZO7gQ8ohG2kXZYTKvlMjuZ055CSSBAoGAdTgOgl9dzSPwCGLdLlJJG80R0U0H31+lzAb4S6RgID4YjAiFkn2fafIAJUUZurbo2djqzasY5wRQQS4TLhlysKm9Uoq1qrX2k3GQVCJ2cQhY28qCL4R3PiutKaLMX/OCNvHuD2vXxG38AEEGx8JgC3On2iadfXwH2pZAng3EihECgYEAn9DRQhgazqs35kZgEwqkpBKovhd32pDRkA6XEoNHB1uTbHTAMuo53SOZbS1bjGB2B2fKwvxMkAFgKAkEhVqOIXXzgxmn8cv4C40QZbrC8Mjrfax6wbNmAGfmKawxxJtRrx3EoAB2ti9D7JaQxgZiIxdvrv4bd/8Ejd7d8P1as7c=";

#[cfg(feature = "digital-signatures")]
const F221_CERTIFICATE_DER_BASE64: &str = "MIIDdDCCAlygAwIBAgIUFuFlT5/whakB7xH7oRLpOMWRK6swDQYJKoZIhvcNAQELBQAwQzEaMBgGA1UEAwwRRi0xNzIgVGVzdCBTaWduZXIxGDAWBgNVBAoMD1RlbnNvcmJlZSBUZXN0czELMAkGA1UEBhMCR0IwHhcNMjYwODIyMjIyOTQ4WhcNMzYwODE5MjIyOTQ4WjBDMRowGAYDVQQDDBFGLTE3MiBUZXN0IFNpZ25lcjEYMBYGA1UECgwPVGVuc29yYmVlIFRlc3RzMQswCQYDVQQGEwJHQjCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAJdXOJaUsg0ZEWaov8VOBbLX+jbMSSZA4e3bdhBkObs+KIGOr6Jvzk4XpQf2cf1VOvesTieOlOPRODNtNpjR+3FCpCD6x1Itm7r4UYCj56sRAAShr0dnrKz/1y4Vt5gtRVVJDffC12r5UwNYxmpfFd3ndBf1qrEGDVwFjhcAeCpZDnRfcBD8gGExOlnRa17fAZl4L6LU1Djlaz0fCnBOCODWZfcFNRo+1lbzTVIE4LZve6DWMYlc798aKB3MPAZwkoBe381nBr48iEq/wbJ5uy0bAIS79161KZo4NsyztxhviWIkrY0sJj31jsEs7jEw2DaHCmWhTAB275GAR/lTwu8CAwEAAaNgMF4wHQYDVR0OBBYEFN+iqzK8SA1T4nFG9QkuPzX2343DMB8GA1UdIwQYMBaAFN+iqzK8SA1T4nFG9QkuPzX2343DMAwGA1UdEwEB/wQCMAAwDgYDVR0PAQH/BAQDAgeAMA0GCSqGSIb3DQEBCwUAA4IBAQBqjXqb8Wm/QzMHl7jLk42TJCwpTBW86nK/KuQb6XLfjEQ3hy4nSchLGZHlb5yladTu7KFp2DIDHRRVMDaxSUWd0bVcwus3D8JCjuydeLsOiGlQTtrDMp0eEBR3NU//047vDLRodwv7q0cE+r9d6z8bXprXBKqKrQrVweaZP6eNxb7Cz6aALENNDrSd5sGy1CfOC5hZzBzFcyQUp+bBmc5XL5MzO3sm5yQebCfqiwBiBwkNEGkjfOIxU0O+LxHjOzXrU2zLbOcn5WvPl+UkD/OzD5NjobsZk4wUtGqSe4+KGgb1VxK8JSnJOBqnUyPc/4KNl8JQROopBKTUxENfWyKM";

#[cfg(feature = "digital-signatures")]
fn f221_signing_material() -> (Vec<u8>, Vec<u8>) {
    (
        decode_base64_fixture(F221_PRIVATE_KEY_PKCS8_BASE64),
        decode_base64_fixture(F221_CERTIFICATE_DER_BASE64),
    )
}

fn f226_shape(
    id: u32,
    name: &str,
    kind: &str,
    idx: u32,
    bounds: (i64, i64, i64, i64),
    text: &str,
) -> String {
    let (x, y, cx, cy) = bounds;
    format!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="{id}" name="{name}"/><p:cNvSpPr/><p:nvPr><p:ph type="{kind}" idx="{idx}"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:defRPr sz="1200"/></a:lvl1pPr></a:lstStyle><a:p><a:r><a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp>"#
    )
}

fn f226_background_shape(id: u32) -> String {
    format!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="{id}" name="Master background"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="6858000" cy="9144000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr></p:sp>"#
    )
}

fn f226_picture(id: u32, relationship_id: &str, bounds: (i64, i64, i64, i64)) -> String {
    let (x, y, cx, cy) = bounds;
    format!(
        r#"<p:pic><p:nvPicPr><p:cNvPr id="{id}" name="Master image"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="{relationship_id}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>"#
    )
}

fn f226_group_scaffold() -> &'static str {
    r#"<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>"#
}

fn f226_notes_master_xml() -> Vec<u8> {
    let shapes = [
        f226_background_shape(9),
        f226_picture(8, "image", (3_200_400, 152_400, 304_800, 304_800)),
        f226_shape(
            2,
            "Header",
            "hdr",
            0,
            (0, 0, 2_971_800, 457_200),
            "F-226 header",
        ),
        f226_shape(
            3,
            "Date",
            "dt",
            1,
            (3_884_613, 0, 2_971_800, 457_200),
            "2030-01-02",
        ),
        f226_shape(
            4,
            "Slide",
            "sldImg",
            2,
            (1_143_000, 685_800, 4_572_000, 3_429_000),
            "",
        ),
        f226_shape(
            5,
            "Notes",
            "body",
            3,
            (685_800, 4_343_400, 5_486_400, 3_657_600),
            "Master prompt",
        ),
        f226_shape(
            6,
            "Footer",
            "ftr",
            4,
            (0, 8_685_213, 2_971_800, 457_200),
            "F-226 footer",
        ),
        f226_shape(
            7,
            "Number",
            "sldNum",
            5,
            (3_884_613, 8_685_213, 2_971_800, 457_200),
            "stored",
        ),
    ]
    .join("");
    let group = f226_group_scaffold();
    format!(
        r#"<p:notesMaster xmlns:p="{P_NS}" xmlns:a="{A_NS}" xmlns:r="{R_NS}"><p:cSld><p:spTree>{group}{shapes}</p:spTree></p:cSld><p:clrMap accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" bg1="lt1" bg2="lt2" folHlink="folHlink" hlink="hlink" tx1="dk1" tx2="dk2"/><p:hf dt="1" hdr="1" ftr="1" sldNum="1"/><p:notesStyle><a:lvl1pPr><a:defRPr sz="1400"/></a:lvl1pPr></p:notesStyle></p:notesMaster>"#
    )
    .into_bytes()
}

fn f226_handout_master_xml() -> Vec<u8> {
    let shapes = [
        f226_background_shape(7),
        f226_picture(6, "image", (3_200_400, 152_400, 304_800, 304_800)),
        f226_shape(
            2,
            "Header",
            "hdr",
            0,
            (0, 0, 2_971_800, 457_200),
            "F-226 handout header",
        ),
        f226_shape(
            3,
            "Date",
            "dt",
            1,
            (3_884_613, 0, 2_971_800, 457_200),
            "2030-01-02",
        ),
        f226_shape(
            4,
            "Footer",
            "ftr",
            2,
            (0, 8_685_213, 2_971_800, 457_200),
            "F-226 handout footer",
        ),
        f226_shape(
            5,
            "Number",
            "sldNum",
            3,
            (3_884_613, 8_685_213, 2_971_800, 457_200),
            "stored",
        ),
    ]
    .join("");
    let group = f226_group_scaffold();
    format!(
        r#"<p:handoutMaster xmlns:p="{P_NS}" xmlns:a="{A_NS}" xmlns:r="{R_NS}"><p:cSld><p:spTree>{group}{shapes}</p:spTree></p:cSld><p:clrMap accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" bg1="lt1" bg2="lt2" folHlink="folHlink" hlink="hlink" tx1="dk1" tx2="dk2"/><p:hf dt="1" hdr="1" ftr="1" sldNum="1"/></p:handoutMaster>"#
    )
    .into_bytes()
}

fn f226_notes_slide_xml(text: &str) -> Vec<u8> {
    let body = format!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="3"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>{text}</a:t></a:r><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>"#
    );
    let picture = f226_picture(3, "image", (3_657_600, 152_400, 304_800, 304_800));
    let group = f226_group_scaffold();
    format!(
        r#"<p:notes xmlns:p="{P_NS}" xmlns:a="{A_NS}" xmlns:r="{R_NS}"><p:cSld><p:spTree>{group}{body}{picture}</p:spTree></p:cSld></p:notes>"#
    )
    .into_bytes()
}

fn f226_blue_pixel_png() -> Vec<u8> {
    let mut pixmap = tiny_skia::Pixmap::new(1, 1).unwrap();
    pixmap.fill(tiny_skia::Color::from_rgba8(0, 0, 255, 255));
    pixmap.encode_png().unwrap()
}

fn f226_fixture_bytes() -> Vec<u8> {
    let mut source = Presentation::new().unwrap();
    for (index, text) in ["F-226 slide one", "F-226 slide two"]
        .into_iter()
        .enumerate()
    {
        source.add_slide(0).unwrap();
        source
            .slide_mut(index)
            .unwrap()
            .add_textbox(Emu(914_400), Emu(914_400), Emu(4_572_000), Emu(914_400))
            .unwrap()
            .set_text(text)
            .unwrap();
    }
    let mut package = open_opc(&source.to_bytes().unwrap(), "F-226 source fixture");
    let presentation_part = package.main_document_part().unwrap();
    let model = CT_Presentation::from_xml(package.get_part(&presentation_part).unwrap()).unwrap();
    let presentation_relationships = package.get_part_rels(&presentation_part).unwrap().clone();
    let slide_parts = model
        .slide_ids
        .iter()
        .map(|slide| {
            let relationship = presentation_relationships
                .get_by_id(&slide.relationship_id)
                .unwrap();
            OpcPackage::resolve_rel_target(&presentation_part, &relationship.target)
        })
        .collect::<Vec<_>>();
    let notes_master_relationship = package
        .get_or_create_part_rels(&presentation_part)
        .items
        .iter_mut()
        .find(|relationship| relationship.rel_type == rel_types::NOTES_MASTER)
        .unwrap();
    notes_master_relationship.target = "../custom/masters/notes.xml".to_owned();
    let notes_master_part = "/custom/masters/notes.xml";
    package.set_part(notes_master_part, f226_notes_master_xml());
    package
        .content_types
        .add_override(notes_master_part, content_types::NOTES_MASTER);
    package
        .get_or_create_part_rels(notes_master_part)
        .add_with_id("theme", rel_types::THEME, "../../ppt/theme/theme2.xml");
    package
        .get_or_create_part_rels(notes_master_part)
        .add_with_id("image", rel_types::IMAGE, "../media/master.png");
    let handout_master_part = "/custom/handouts/master.xml";
    package.set_part(handout_master_part, f226_handout_master_xml());
    package
        .content_types
        .add_override(handout_master_part, content_types::HANDOUT_MASTER);
    package
        .get_or_create_part_rels(handout_master_part)
        .add_with_id("theme", rel_types::THEME, "../../ppt/theme/theme2.xml");
    package
        .get_or_create_part_rels(handout_master_part)
        .add_with_id("image", rel_types::IMAGE, "../media/master.png");
    package.set_part("/custom/media/master.png", valid_one_pixel_png());
    package.content_types.add_default("png", "image/png");
    package
        .get_or_create_part_rels(&presentation_part)
        .add_with_id(
            "handout",
            rel_types::HANDOUT_MASTER,
            "../custom/handouts/master.xml",
        );
    for (index, slide_part) in slide_parts.iter().enumerate() {
        let notes_part = format!("/custom/notes/notes{}.xml", index + 1);
        package.set_part(
            &notes_part,
            f226_notes_slide_xml(&format!("F-226 speaker note {}", index + 1)),
        );
        package
            .content_types
            .add_override(&notes_part, content_types::NOTES_SLIDE);
        package.get_or_create_part_rels(slide_part).add_with_id(
            &format!("notes-{}", index + 1),
            rel_types::NOTES_SLIDE,
            &format!("../../custom/notes/notes{}.xml", index + 1),
        );
        package.get_or_create_part_rels(&notes_part).add_with_id(
            "master",
            rel_types::NOTES_MASTER,
            "../masters/notes.xml",
        );
        package.get_or_create_part_rels(&notes_part).add_with_id(
            "slide",
            rel_types::SLIDE,
            &format!("../../{}", slide_part.trim_start_matches('/')),
        );
        package.get_or_create_part_rels(&notes_part).add_with_id(
            "image",
            rel_types::IMAGE,
            "../media/notes.png",
        );
    }
    package.set_part("/custom/media/notes.png", f226_blue_pixel_png());
    package_bytes(package)
}

fn m21_notes_handout_fixture_bytes() -> Vec<u8> {
    let mut source = Presentation::new().unwrap();
    for (index, text) in ["M21 slide one", "M21 slide two"].into_iter().enumerate() {
        source.add_slide(0).unwrap();
        source
            .slide_mut(index)
            .unwrap()
            .add_textbox(Emu(914_400), Emu(914_400), Emu(4_572_000), Emu(914_400))
            .unwrap()
            .set_text(text)
            .unwrap();
    }
    let mut package = open_opc(&source.to_bytes().unwrap(), "M21 output fixture");
    let presentation_part = package.main_document_part().unwrap();
    let model = CT_Presentation::from_xml(package.get_part(&presentation_part).unwrap()).unwrap();
    let presentation_relationships = package.get_part_rels(&presentation_part).unwrap().clone();
    let slide_parts = model
        .slide_ids
        .iter()
        .map(|slide| {
            let relationship = presentation_relationships
                .get_by_id(&slide.relationship_id)
                .unwrap();
            OpcPackage::resolve_rel_target(&presentation_part, &relationship.target)
        })
        .collect::<Vec<_>>();
    let notes_master_relationship = package
        .get_or_create_part_rels(&presentation_part)
        .items
        .iter()
        .find(|relationship| relationship.rel_type == rel_types::NOTES_MASTER)
        .unwrap()
        .clone();
    let notes_master_part =
        OpcPackage::resolve_rel_target(&presentation_part, &notes_master_relationship.target);
    package.set_part(&notes_master_part, f226_notes_master_xml());
    package
        .content_types
        .add_override(&notes_master_part, content_types::NOTES_MASTER);
    let notes_master_rels = package.get_or_create_part_rels(&notes_master_part);
    notes_master_rels.items.clear();
    notes_master_rels.add_with_id("theme", rel_types::THEME, "../theme/theme2.xml");
    notes_master_rels.add_with_id("image", rel_types::IMAGE, "../media/m21-master.png");
    let handout_master_part = "/ppt/handoutMasters/handoutMaster1.xml";
    package.set_part(handout_master_part, f226_handout_master_xml());
    package
        .content_types
        .add_override(handout_master_part, content_types::HANDOUT_MASTER);
    package
        .get_or_create_part_rels(handout_master_part)
        .add_with_id("theme", rel_types::THEME, "../theme/theme2.xml");
    package
        .get_or_create_part_rels(handout_master_part)
        .add_with_id("image", rel_types::IMAGE, "../media/m21-master.png");
    package
        .get_or_create_part_rels(&presentation_part)
        .add_with_id(
            "handout",
            rel_types::HANDOUT_MASTER,
            "handoutMasters/handoutMaster1.xml",
        );
    package.set_part("/ppt/media/m21-master.png", valid_one_pixel_png());
    package.set_part("/ppt/media/m21-notes.png", f226_blue_pixel_png());
    package.content_types.add_default("png", "image/png");
    for (index, slide_part) in slide_parts.iter().enumerate() {
        let notes_part = format!("/ppt/notesSlides/notesSlide{}.xml", index + 1);
        package.set_part(
            &notes_part,
            f226_notes_slide_xml(&format!("M21 speaker note {}", index + 1)),
        );
        package
            .content_types
            .add_override(&notes_part, content_types::NOTES_SLIDE);
        package.get_or_create_part_rels(slide_part).add_with_id(
            &format!("notes-{}", index + 1),
            rel_types::NOTES_SLIDE,
            &format!("../notesSlides/notesSlide{}.xml", index + 1),
        );
        package.get_or_create_part_rels(&notes_part).add_with_id(
            "master",
            rel_types::NOTES_MASTER,
            "../notesMasters/notesMaster1.xml",
        );
        package.get_or_create_part_rels(&notes_part).add_with_id(
            "slide",
            rel_types::SLIDE,
            &format!("../slides/slide{}.xml", index + 1),
        );
        package.get_or_create_part_rels(&notes_part).add_with_id(
            "image",
            rel_types::IMAGE,
            "../media/m21-notes.png",
        );
    }
    package_bytes(package)
}

fn f226_pdf_text(pdf: &[u8], label: &str) -> String {
    let path = std::env::temp_dir().join(format!("rpptx-f226-{label}-{}.pdf", std::process::id()));
    fs::write(&path, pdf).unwrap();
    let output = Command::new("pdftotext")
        .arg(&path)
        .arg("-")
        .output()
        .unwrap();
    fs::remove_file(path).unwrap();
    assert!(output.status.success());
    String::from_utf8(output.stdout).unwrap()
}

fn f226_pdf_page_count(pdf: &[u8], label: &str) -> usize {
    let path = std::env::temp_dir().join(format!(
        "rpptx-f226-{label}-pages-{}.pdf",
        std::process::id()
    ));
    fs::write(&path, pdf).unwrap();
    let output = Command::new("pdfinfo").arg(&path).output().unwrap();
    fs::remove_file(path).unwrap();
    assert!(output.status.success());
    String::from_utf8(output.stdout)
        .unwrap()
        .lines()
        .find_map(|line| line.strip_prefix("Pages:").map(str::trim))
        .unwrap()
        .parse()
        .unwrap()
}

fn m21_pdf_page_size_path(pdf: &Path) -> (f64, f64) {
    let output = Command::new("pdfinfo").arg(pdf).output().unwrap();
    assert!(
        output.status.success(),
        "pdfinfo {}: {}",
        pdf.display(),
        String::from_utf8_lossy(&output.stderr)
    );
    let text = String::from_utf8(output.stdout).unwrap();
    let fields = text
        .lines()
        .find(|line| line.starts_with("Page size:"))
        .unwrap()
        .split_whitespace()
        .collect::<Vec<_>>();
    (fields[2].parse().unwrap(), fields[4].parse().unwrap())
}

fn m21_pdf_page_size(pdf: &[u8], label: &str) -> (f64, f64) {
    let path =
        std::env::temp_dir().join(format!("rpptx-m21-{label}-size-{}.pdf", std::process::id()));
    fs::write(&path, pdf).unwrap();
    let size = m21_pdf_page_size_path(&path);
    fs::remove_file(path).unwrap();
    size
}

fn f226_png_dimensions(png: &[u8]) -> (u32, u32) {
    assert_eq!(&png[..8], b"\x89PNG\r\n\x1a\n");
    (
        u32::from_be_bytes(png[16..20].try_into().unwrap()),
        u32::from_be_bytes(png[20..24].try_into().unwrap()),
    )
}

fn m21_pdf_page_pngs(pdf: &Path, output_prefix: &Path, dpi: u32) -> Vec<Vec<u8>> {
    let output = Command::new("pdftoppm")
        .args(["-r", &dpi.to_string(), "-png"])
        .arg(pdf)
        .arg(output_prefix)
        .output()
        .unwrap();
    assert!(
        output.status.success(),
        "render PowerPoint PDF: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let mut pages = Vec::new();
    for page_number in 1.. {
        let path = output_prefix.with_file_name(format!(
            "{}-{page_number}",
            output_prefix.file_name().unwrap().to_string_lossy()
        ));
        let path = path.with_extension("png");
        if !path.is_file() {
            break;
        }
        pages.push(fs::read(path).unwrap());
    }
    assert!(!pages.is_empty(), "PowerPoint PDF produced no raster pages");
    pages
}

fn m21_normalized_tokens(text: &str) -> Vec<String> {
    text.split_whitespace().map(str::to_owned).collect()
}

fn m21_pdf_text_pages(pdf: &[u8], label: &str) -> Vec<String> {
    let text = f226_pdf_text(pdf, label);
    let mut pages = text.split('\u{c}').map(str::to_owned).collect::<Vec<_>>();
    if pages.last().is_some_and(|page| page.trim().is_empty()) {
        pages.pop();
    }
    pages
}

fn m21_pdf_token_pages(pdf: &[u8], label: &str) -> Vec<Vec<String>> {
    m21_pdf_text_pages(pdf, label)
        .iter()
        .map(|page| m21_normalized_tokens(page))
        .collect()
}

fn m21_assert_text_mutations_fail(actual: &str, expected: &[&str]) {
    assert!(m21_tokens_match(actual, expected));
    assert!(!m21_tokens_match(&format!("{actual} unexpected"), expected));
    assert!(!m21_tokens_match(&format!("{actual} {actual}"), expected));
    if let Some(first) = expected.first() {
        assert!(!m21_tokens_match(
            &actual.replacen(first, &format!("{first}extra"), 1),
            expected
        ));
    }
    if expected.len() >= 2 {
        let mut reordered = m21_normalized_tokens(actual);
        reordered.swap(0, 1);
        assert!(!m21_tokens_match(&reordered.join(" "), expected));
    }
}

fn m21_normalize_movie_frame(png: &[u8]) -> Vec<u8> {
    let info = oxml_media::probe(png).unwrap();
    assert_eq!(u64::from(info.width_px) * 9, u64::from(info.height_px) * 16);
    let image = PositionedElement::Image {
        rect: Rect {
            x: 0.0,
            y: 0.0,
            width: 960.0,
            height: 540.0,
        },
        data: png.to_vec(),
        content_type: "image/png".to_owned(),
        media_id: MediaId(1),
    };
    let page = PageFrame::new(1, 960.0, 540.0, vec![image]);
    let layout = oxml_layout::LayoutResult::new(
        vec![std::sync::Arc::new(page)],
        Vec::new(),
        None,
        Vec::new(),
    );
    oxml_pdf::render_page_to_png(&layout, 0, F214_POWERPOINT_COMPARISON_DPI).unwrap()
}

fn m21_mask_audio_poster(png: &[u8]) -> Vec<u8> {
    let mut pixmap = tiny_skia::Pixmap::decode_png(png).unwrap();
    for y in 598..772 {
        for x in 823..1_128 {
            let offset = (y * pixmap.width() as usize + x) * 4;
            pixmap.data_mut()[offset..offset + 4].copy_from_slice(&[255, 255, 255, 255]);
        }
    }
    pixmap.encode_png().unwrap()
}

fn m21_ink_mass(png: &[u8], bounds: (u32, u32, u32, u32)) -> u64 {
    let pixmap = tiny_skia::Pixmap::decode_png(png).unwrap();
    let (left, top, width, height) = bounds;
    (top..top + height)
        .flat_map(|y| (left..left + width).map(move |x| (x, y)))
        .map(|(x, y)| {
            let pixel = pixmap.pixel(x, y).unwrap();
            let alpha_gap = 255 - pixel.alpha();
            let lightest_dark_channel = pixel
                .red()
                .min(pixel.green())
                .min(pixel.blue())
                .saturating_add(alpha_gap);
            u64::from(255 - lightest_dark_channel)
        })
        .sum()
}

fn m21_timeline_layout(
    frame: &rpptx::DeterministicTimelineFrame,
    deterministic_fonts: &[oxml_layout::FontData],
) -> oxml_layout::LayoutResult {
    let font_ids = deterministic_fonts
        .iter()
        .map(|font| font.id)
        .collect::<HashSet<_>>();
    let mut glyph_run_count = 0;
    oxml_layout::walk(&frame.page.elements, &mut |element, _| {
        let font_id = match element {
            PositionedElement::Text(run) => Some(run.font_id),
            PositionedElement::MultilingualText(run) => Some(run.font_id),
            _ => None,
        };
        if let Some(font_id) = font_id {
            glyph_run_count += 1;
            assert!(
                font_ids.contains(&font_id),
                "timeline FontId {font_id:?} must resolve through static deterministic fonts"
            );
        }
    });
    assert!(
        glyph_run_count > 0,
        "timeline frame must retain visible text"
    );
    oxml_layout::LayoutResult::new(
        vec![std::sync::Arc::new(frame.page.clone())],
        deterministic_fonts.to_vec(),
        None,
        Vec::new(),
    )
}

fn m21_visible_page_text(page: &PageFrame) -> String {
    fn visit(elements: &[PositionedElement], opacity: f64, output: &mut Vec<String>) {
        for element in elements {
            match element {
                PositionedElement::Text(run) if opacity > 0.0 => output.push(run.text.clone()),
                PositionedElement::MultilingualText(run) if opacity > 0.0 => {
                    output.push(run.logical_text.clone());
                }
                PositionedElement::Group(group) => {
                    visit(&group.children, opacity * group.opacity, output);
                }
                PositionedElement::MarkedContent { children, .. } => {
                    visit(children, opacity, output);
                }
                _ => {}
            }
        }
    }

    let mut output = Vec::new();
    visit(&page.elements, 1.0, &mut output);
    output.join("\n")
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
struct M21InkBounds {
    left: u32,
    top: u32,
    right: u32,
    bottom: u32,
}

fn m21_page_ink_bounds(png: &[u8], monochrome_only: bool) -> Vec<M21InkBounds> {
    const INK_CUTOFF: u8 = 245;
    const MAX_ROW_GAP: u32 = 8;
    let pixmap = tiny_skia::Pixmap::decode_png(png).unwrap();
    let row_has_ink = (0..pixmap.height())
        .map(|y| {
            (0..pixmap.width()).any(|x| {
                let pixel = pixmap.pixel(x, y).unwrap();
                pixel.red().min(pixel.green()).min(pixel.blue()) < INK_CUTOFF
                    && (!monochrome_only
                        || pixel.red().abs_diff(pixel.green()) < 16
                            && pixel.green().abs_diff(pixel.blue()) < 16)
            })
        })
        .collect::<Vec<_>>();
    let mut row_bands = Vec::new();
    let mut active = None;
    let mut last_ink = 0;
    for (y, has_ink) in row_has_ink.into_iter().enumerate() {
        let y = y as u32;
        if !has_ink {
            continue;
        }
        match active {
            Some(top) if y - last_ink > MAX_ROW_GAP => {
                row_bands.push((top, last_ink + 1));
                active = Some(y);
            }
            None => active = Some(y),
            Some(_) => {}
        }
        last_ink = y;
    }
    if let Some(top) = active {
        row_bands.push((top, last_ink + 1));
    }
    row_bands
        .into_iter()
        .map(|(top, bottom)| {
            let mut left = pixmap.width();
            let mut right = 0;
            for y in top..bottom {
                for x in 0..pixmap.width() {
                    let pixel = pixmap.pixel(x, y).unwrap();
                    if pixel.red().min(pixel.green()).min(pixel.blue()) < INK_CUTOFF
                        && (!monochrome_only
                            || pixel.red().abs_diff(pixel.green()) < 16
                                && pixel.green().abs_diff(pixel.blue()) < 16)
                    {
                        left = left.min(x);
                        right = right.max(x + 1);
                    }
                }
            }
            M21InkBounds {
                left,
                top,
                right,
                bottom,
            }
        })
        .collect()
}

fn m21_full_page_ink_bounds(png: &[u8]) -> Vec<M21InkBounds> {
    m21_page_ink_bounds(png, false)
}

fn m21_crop_png(png: &[u8], bounds: M21InkBounds) -> Vec<u8> {
    let source = tiny_skia::Pixmap::decode_png(png).unwrap();
    let width = bounds.right - bounds.left;
    let height = bounds.bottom - bounds.top;
    let mut cropped = tiny_skia::Pixmap::new(width, height).unwrap();
    for y in 0..height {
        let source_start = (((bounds.top + y) * source.width() + bounds.left) * 4) as usize;
        let source_end = source_start + (width * 4) as usize;
        let target_start = (y * width * 4) as usize;
        let target_end = target_start + (width * 4) as usize;
        cropped.data_mut()[target_start..target_end]
            .copy_from_slice(&source.data()[source_start..source_end]);
    }
    cropped.encode_png().unwrap()
}

#[derive(Clone, Debug)]
struct M21InkComparison {
    actual: Vec<M21InkBounds>,
    expected: Vec<M21InkBounds>,
    geometry_error_px: u32,
    region_ssim: Vec<f64>,
}

const M21_MAX_FULL_PAGE_INK_ERROR_PX: u32 = 6;
const M21_MIN_INK_REGION_SSIM: f64 = 0.45;

fn m21_visual_gate(text_order_matches: bool, comparison: &M21InkComparison) -> bool {
    text_order_matches
        && !comparison.actual.is_empty()
        && comparison.actual.len() == comparison.expected.len()
        && comparison.geometry_error_px <= M21_MAX_FULL_PAGE_INK_ERROR_PX
        && comparison
            .region_ssim
            .iter()
            .all(|score| *score >= M21_MIN_INK_REGION_SSIM)
}

fn m21_static_page_gate(
    text_matches: bool,
    actual_png: &[u8],
    expected_png: &[u8],
    comparison: &M21InkComparison,
) -> bool {
    if comparison.expected.is_empty() {
        text_matches
            && comparison.actual.is_empty()
            && f226_png_dimensions(actual_png) == f226_png_dimensions(expected_png)
            && smartart_png_ssim(actual_png, expected_png) >= 0.999_999
    } else {
        m21_visual_gate(text_matches, comparison)
    }
}

fn m21_solidify_first_ink_region(png: &[u8]) -> Vec<u8> {
    let bounds = m21_full_page_ink_bounds(png)[0];
    m21_solidify_ink_region(png, bounds)
}

fn m21_solidify_ink_region(png: &[u8], bounds: M21InkBounds) -> Vec<u8> {
    let mut pixmap = tiny_skia::Pixmap::decode_png(png).unwrap();
    for y in bounds.top..bounds.bottom {
        for x in bounds.left..bounds.right {
            let offset = ((y * pixmap.width() + x) * 4) as usize;
            pixmap.data_mut()[offset..offset + 4].copy_from_slice(&[0, 0, 0, 255]);
        }
    }
    pixmap.encode_png().unwrap()
}

fn m21_extend_ink_region(png: &[u8], bounds: M21InkBounds, pixels: u32) -> Vec<u8> {
    let mut pixmap = tiny_skia::Pixmap::decode_png(png).unwrap();
    let right = bounds.right.saturating_add(pixels).min(pixmap.width());
    assert!(
        right > bounds.right,
        "ink extension must remain inside the page"
    );
    for y in bounds.top..bounds.bottom {
        for x in bounds.right..right {
            let offset = ((y * pixmap.width() + x) * 4) as usize;
            pixmap.data_mut()[offset..offset + 4].copy_from_slice(&[0, 0, 0, 255]);
        }
    }
    pixmap.encode_png().unwrap()
}

fn m21_shift_ink_up(png: &[u8], pixels: u32) -> Vec<u8> {
    let source = tiny_skia::Pixmap::decode_png(png).unwrap();
    let mut shifted = source.clone();
    for bounds in m21_full_page_ink_bounds(png) {
        let padded = M21InkBounds {
            left: bounds.left.saturating_sub(2),
            top: bounds.top.saturating_sub(2),
            right: (bounds.right + 2).min(source.width()),
            bottom: (bounds.bottom + 2).min(source.height()),
        };
        for y in padded.top..padded.bottom {
            for x in padded.left..padded.right {
                let target_offset = ((y * shifted.width() + x) * 4) as usize;
                shifted.data_mut()[target_offset..target_offset + 4]
                    .copy_from_slice(&[255, 255, 255, 255]);
            }
        }
        for y in padded.top..padded.bottom {
            let Some(target_y) = y.checked_sub(pixels) else {
                continue;
            };
            for x in padded.left..padded.right {
                let source_offset = ((y * source.width() + x) * 4) as usize;
                let target_offset = ((target_y * shifted.width() + x) * 4) as usize;
                shifted.data_mut()[target_offset..target_offset + 4]
                    .copy_from_slice(&source.data()[source_offset..source_offset + 4]);
            }
        }
    }
    shifted.encode_png().unwrap()
}

fn m21_tokens_match(text: &str, expected: &[&str]) -> bool {
    m21_normalized_tokens(text) == expected
}

#[derive(Clone, Copy, Debug)]
struct M21NormalizedBounds {
    left: f64,
    top: f64,
    right: f64,
    bottom: f64,
}

fn m21_normalized_geometry_error(
    actual: &[M21NormalizedBounds],
    expected: &[M21NormalizedBounds],
) -> f64 {
    assert_eq!(actual.len(), expected.len());
    actual
        .iter()
        .zip(expected)
        .flat_map(|(actual, expected)| {
            [
                (actual.left - expected.left).abs(),
                (actual.top - expected.top).abs(),
                (actual.right - expected.right).abs(),
                (actual.bottom - expected.bottom).abs(),
            ]
        })
        .fold(0.0_f64, f64::max)
}

fn m21_handout_thumbnail_bounds(png: &[u8]) -> Vec<M21NormalizedBounds> {
    let pixmap = tiny_skia::Pixmap::decode_png(png).unwrap();
    let minimum_run = pixmap.width() * 3 / 10;
    let maximum_left = pixmap.width() * 2 / 5;
    let maximum_right = pixmap.width() * 3 / 5;
    let mut horizontal_edges = Vec::new();
    for y in 0..pixmap.height() {
        let mut start = None;
        let mut best = None;
        for x in 0..=pixmap.width() {
            let black = x < pixmap.width() && {
                let pixel = pixmap.pixel(x, y).unwrap();
                pixel.red() < 192 && pixel.green() < 192 && pixel.blue() < 192
            };
            match (start, black) {
                (None, true) => start = Some(x),
                (Some(run_start), false) => {
                    if x - run_start >= minimum_run && run_start < maximum_left && x < maximum_right
                    {
                        best = Some((run_start, x));
                    }
                    start = None;
                }
                _ => {}
            }
        }
        if let Some((left, right)) = best {
            if horizontal_edges
                .last()
                .is_some_and(|(previous_y, _, _)| y - previous_y <= 2)
            {
                continue;
            }
            horizontal_edges.push((y, left, right));
        }
    }
    assert_eq!(
        horizontal_edges.len(),
        6,
        "expected three thumbnail border pairs: {horizontal_edges:?}"
    );
    horizontal_edges
        .chunks_exact(2)
        .map(|pair| M21NormalizedBounds {
            left: f64::from(pair[0].1) / f64::from(pixmap.width()),
            top: f64::from(pair[0].0) / f64::from(pixmap.height()),
            right: f64::from(pair[0].2) / f64::from(pixmap.width()),
            bottom: f64::from(pair[1].0 + 1) / f64::from(pixmap.height()),
        })
        .collect()
}

#[derive(Clone, Debug)]
struct M21NotesInkComparison {
    rust_band_count: usize,
    powerpoint_band_count: usize,
    component_size_error: f64,
    rust_ink_occupancy: f64,
    powerpoint_ink_occupancy: f64,
    component_raster_error: f64,
}

const M21_MAX_NOTES_COMPONENT_SIZE_ERROR: f64 = 0.06;
const M21_MAX_NOTES_COMPONENT_RASTER_ERROR: f64 = 0.35;

fn m21_component_ink_occupancy(png: &[u8], bounds: M21InkBounds) -> f64 {
    const INK_CUTOFF: u8 = 245;
    let width = bounds.right - bounds.left;
    let height = bounds.bottom - bounds.top;
    let pixmap = tiny_skia::Pixmap::decode_png(png).unwrap();
    let ink_pixels = (bounds.top..bounds.bottom)
        .flat_map(|y| (bounds.left..bounds.right).map(move |x| (x, y)))
        .filter(|(x, y)| {
            let pixel = pixmap.pixel(*x, *y).unwrap();
            pixel.red().min(pixel.green()).min(pixel.blue()) < INK_CUTOFF
                && pixel.red().abs_diff(pixel.green()) < 16
                && pixel.green().abs_diff(pixel.blue()) < 16
        })
        .count();
    ink_pixels as f64 / (f64::from(width) * f64::from(height))
}

fn m21_compare_notes_ink(
    rust_png: &[u8],
    powerpoint_png: &[u8],
    rust_component_index: usize,
) -> M21NotesInkComparison {
    let rust_bands = m21_page_ink_bounds(rust_png, true);
    let powerpoint_bands = m21_page_ink_bounds(powerpoint_png, true);
    let rust_component = rust_bands[rust_component_index];
    let powerpoint_component = powerpoint_bands[0];
    let (rust_width, rust_height) = f226_png_dimensions(rust_png);
    let (powerpoint_width, powerpoint_height) = f226_png_dimensions(powerpoint_png);
    let rust_component_size = (
        f64::from(rust_component.right - rust_component.left) / f64::from(rust_width),
        f64::from(rust_component.bottom - rust_component.top) / f64::from(rust_height),
    );
    let powerpoint_component_size = (
        f64::from(powerpoint_component.right - powerpoint_component.left)
            / f64::from(powerpoint_width),
        f64::from(powerpoint_component.bottom - powerpoint_component.top)
            / f64::from(powerpoint_height),
    );
    let rust_ink_occupancy = m21_component_ink_occupancy(rust_png, rust_component);
    let powerpoint_ink_occupancy =
        m21_component_ink_occupancy(powerpoint_png, powerpoint_component);
    M21NotesInkComparison {
        rust_band_count: rust_bands.len(),
        powerpoint_band_count: powerpoint_bands.len(),
        component_size_error: (rust_component_size.0 - powerpoint_component_size.0)
            .abs()
            .max((rust_component_size.1 - powerpoint_component_size.1).abs()),
        rust_ink_occupancy,
        powerpoint_ink_occupancy,
        component_raster_error: (rust_ink_occupancy - powerpoint_ink_occupancy).abs(),
    }
}

fn m21_notes_ink_gate(
    text_matches: bool,
    comparison: &M21NotesInkComparison,
    expected_rust_bands: usize,
) -> bool {
    text_matches
        && comparison.rust_band_count == expected_rust_bands
        && comparison.powerpoint_band_count == 1
        && comparison.component_size_error <= M21_MAX_NOTES_COMPONENT_SIZE_ERROR
        && comparison.rust_ink_occupancy > 0.0
        && comparison.rust_ink_occupancy <= 1.0
        && comparison.powerpoint_ink_occupancy > 0.0
        && comparison.powerpoint_ink_occupancy <= 1.0
        && comparison.component_raster_error <= M21_MAX_NOTES_COMPONENT_RASTER_ERROR
}

fn m21_assert_page_size(actual: (f64, f64), expected: (f64, f64)) {
    assert!(
        (actual.0 - expected.0).abs() <= 0.01 && (actual.1 - expected.1).abs() <= 0.01,
        "page size {actual:?} differs from {expected:?}"
    );
}

fn m21_compare_full_page_ink(actual_png: &[u8], expected_png: &[u8]) -> M21InkComparison {
    let actual = m21_full_page_ink_bounds(actual_png);
    let expected = m21_full_page_ink_bounds(expected_png);
    let mut geometry_error_px = 0;
    let region_ssim = actual
        .iter()
        .zip(&expected)
        .map(|(actual, expected)| {
            for delta in [
                actual.left.abs_diff(expected.left),
                actual.top.abs_diff(expected.top),
                actual.right.abs_diff(expected.right),
                actual.bottom.abs_diff(expected.bottom),
            ] {
                geometry_error_px = geometry_error_px.max(delta);
            }
            let union = M21InkBounds {
                left: actual.left.min(expected.left).saturating_sub(8),
                top: actual.top.min(expected.top).saturating_sub(8),
                right: (actual.right.max(expected.right) + 8)
                    .min(f226_png_dimensions(actual_png).0),
                bottom: (actual.bottom.max(expected.bottom) + 8)
                    .min(f226_png_dimensions(actual_png).1),
            };
            smartart_png_ssim(
                &m21_crop_png(actual_png, union),
                &m21_crop_png(expected_png, union),
            )
        })
        .collect();
    M21InkComparison {
        actual,
        expected,
        geometry_error_px,
        region_ssim,
    }
}

#[test]
fn notes_pages_follow_master_geometry_text_metadata_and_slide_order() {
    let presentation = Presentation::from_bytes(&f226_fixture_bytes()).unwrap();
    let pdf = presentation.to_notes_pdf_deterministic().unwrap();
    assert_eq!(f226_pdf_page_count(&pdf, "notes-golden"), 2);
    let text = f226_pdf_text(&pdf, "notes-golden");
    let pages = text
        .split('\u{c}')
        .filter(|page| !page.trim().is_empty())
        .collect::<Vec<_>>();
    assert_eq!(pages.len(), 2, "unexpected extracted pages: {text:?}");
    for (page, slide_word, page_number) in [(pages[0], "one", "1"), (pages[1], "two", "2")] {
        let words = page.split_whitespace().collect::<Vec<_>>();
        for expected in [
            slide_word,
            "speaker",
            "note",
            page_number,
            "header",
            "footer",
            "2030-01-02",
        ] {
            assert!(
                words.contains(&expected),
                "missing {expected:?} in {page:?}"
            );
        }
    }
    let pngs = presentation.notes_page_pngs_deterministic(72.0).unwrap();
    let page = tiny_skia::Pixmap::decode_png(&pngs[0]).unwrap();
    let outside = page.pixel(88, 100).unwrap();
    assert_eq!(
        (outside.red(), outside.green(), outside.blue()),
        (255, 0, 0)
    );
    let inside = page.pixel(92, 100).unwrap();
    assert_eq!(
        (inside.red(), inside.green(), inside.blue()),
        (255, 255, 255),
        "the source thumbnail must start at the notes-master slide-image edge"
    );
}

#[test]
fn handouts_follow_master_metadata_and_all_six_audience_layouts() {
    let presentation = Presentation::from_bytes(&f226_fixture_bytes()).unwrap();
    for (layout, pages) in [
        (HandoutLayout::One, 2),
        (HandoutLayout::Two, 1),
        (HandoutLayout::Three, 1),
        (HandoutLayout::Four, 1),
        (HandoutLayout::Six, 1),
        (HandoutLayout::Nine, 1),
    ] {
        let pdf = presentation.to_handout_pdf_deterministic(layout).unwrap();
        assert_eq!(f226_pdf_page_count(&pdf, "handout-golden"), pages);
        let text = f226_pdf_text(&pdf, "handout-golden");
        let words = text.split_whitespace().collect::<Vec<_>>();
        for expected in [
            "one", "two", "handout", "header", "footer", "2030", "01", "02",
        ] {
            assert!(
                words.contains(&expected),
                "missing {expected:?} in {text:?}"
            );
        }
    }
    let pngs = presentation
        .handout_page_pngs_deterministic(HandoutLayout::One, 72.0)
        .unwrap();
    let page = tiny_skia::Pixmap::decode_png(&pngs[0]).unwrap();
    let background = page.pixel(12, 100).unwrap();
    assert_eq!(
        (background.red(), background.green(), background.blue()),
        (255, 0, 0)
    );
    let thumbnail = page.pixel(270, 353).unwrap();
    assert_eq!(
        (thumbnail.red(), thumbnail.green(), thumbnail.blue()),
        (255, 255, 255),
        "the master background must remain below the source thumbnail"
    );
}

#[test]
fn notes_and_handouts_export_pdf_and_png_with_deterministic_dimensions() {
    let presentation = Presentation::from_bytes(&f226_fixture_bytes()).unwrap();
    assert_eq!(
        presentation.to_notes_pdf_deterministic().unwrap(),
        presentation.to_notes_pdf_deterministic().unwrap()
    );
    let notes = presentation.notes_page_pngs_deterministic(144.0).unwrap();
    let notes_again = presentation.notes_page_pngs_deterministic(144.0).unwrap();
    assert_eq!(notes, notes_again);
    assert_eq!(notes.len(), 2);
    assert!(
        notes
            .iter()
            .all(|png| f226_png_dimensions(png) == (1080, 1440))
    );
    let handouts = presentation
        .handout_page_pngs_deterministic(HandoutLayout::Six, 144.0)
        .unwrap();
    assert_eq!(handouts.len(), 1);
    assert_eq!(f226_png_dimensions(&handouts[0]), (1080, 1440));
}

#[test]
fn notes_and_handout_export_resolve_noncanonical_master_theme_and_media_targets() {
    let presentation = Presentation::from_bytes(&f226_fixture_bytes()).unwrap();
    assert!(
        presentation
            .to_notes_pdf_deterministic()
            .unwrap()
            .starts_with(b"%PDF")
    );
    assert!(
        presentation
            .to_handout_pdf_deterministic(HandoutLayout::Three)
            .unwrap()
            .starts_with(b"%PDF")
    );
    let notes = presentation.notes_page_pngs_deterministic(72.0).unwrap();
    let page = tiny_skia::Pixmap::decode_png(&notes[0]).unwrap();
    let image = page.pixel(264, 24).unwrap();
    assert_eq!((image.red(), image.green(), image.blue()), (255, 0, 0));
    let notes_image = page.pixel(300, 24).unwrap();
    assert_eq!(
        (notes_image.red(), notes_image.green(), notes_image.blue()),
        (0, 0, 255),
        "the notes-slide image must not alias the master's same-id relationship"
    );

    let mut package = open_opc(&f226_fixture_bytes(), "F-226 fallback placeholder");
    let notes_part = "/custom/notes/notes1.xml";
    let notes = String::from_utf8(package.get_part(notes_part).unwrap().to_vec()).unwrap();
    assert!(notes.contains("type=\"body\" idx=\"3\""));
    package.set_part(
        notes_part,
        notes
            .replacen("type=\"body\" idx=\"3\"", "type=\"body\"", 1)
            .into_bytes(),
    );
    let fallback = Presentation::from_bytes(&package_bytes(package)).unwrap();
    assert!(
        f226_pdf_text(&fallback.to_notes_pdf_deterministic().unwrap(), "fallback")
            .contains("speaker")
    );
}

#[test]
fn notes_and_handout_export_fail_closed_for_broken_hierarchy_and_invalid_dpi() {
    let bytes = f226_fixture_bytes();
    let presentation = Presentation::from_bytes(&bytes).unwrap();
    for dpi in [-1.0, 0.0, f64::NAN, f64::INFINITY, 601.0] {
        assert!(presentation.notes_page_pngs_deterministic(dpi).is_err());
    }
    let rejects_notes =
        |package: OpcPackage| match Presentation::from_bytes(&package_bytes(package)) {
            Ok(presentation) => presentation.to_notes_pdf_deterministic().is_err(),
            Err(_) => true,
        };
    let rejects_handout =
        |package: OpcPackage| match Presentation::from_bytes(&package_bytes(package)) {
            Ok(presentation) => presentation
                .to_handout_pdf_deterministic(HandoutLayout::Three)
                .is_err(),
            Err(_) => true,
        };

    let mut package = open_opc(&bytes, "external F-226 theme");
    package
        .get_or_create_part_rels("/custom/masters/notes.xml")
        .items
        .iter_mut()
        .find(|relationship| relationship.rel_type == rel_types::THEME)
        .unwrap()
        .target_mode = Some("External".to_owned());
    assert!(rejects_notes(package));

    let mut package = open_opc(&bytes, "duplicate F-226 theme");
    package
        .get_or_create_part_rels("/custom/masters/notes.xml")
        .add_with_id("theme-two", rel_types::THEME, "../../ppt/theme/theme2.xml");
    assert!(rejects_notes(package));

    let mut package = open_opc(&bytes, "wrong-type F-226 theme");
    package
        .get_or_create_part_rels("/custom/masters/notes.xml")
        .items
        .iter_mut()
        .find(|relationship| relationship.rel_type == rel_types::THEME)
        .unwrap()
        .rel_type = rel_types::IMAGE.to_owned();
    assert!(rejects_notes(package));

    let mut package = open_opc(&bytes, "missing F-226 theme");
    package
        .get_or_create_part_rels("/custom/masters/notes.xml")
        .items
        .iter_mut()
        .find(|relationship| relationship.rel_type == rel_types::THEME)
        .unwrap()
        .target = "../missing/theme.xml".to_owned();
    assert!(rejects_notes(package));

    let mut package = open_opc(&bytes, "malformed F-226 master");
    package.set_part("/custom/masters/notes.xml", b"<p:notesMaster".to_vec());
    assert!(rejects_notes(package));

    let mut package = open_opc(&bytes, "unmatched F-226 placeholder");
    let notes_part = "/custom/notes/notes1.xml";
    let notes = String::from_utf8(package.get_part(notes_part).unwrap().to_vec()).unwrap();
    assert!(notes.contains("type=\"body\" idx=\"3\""));
    package.set_part(
        notes_part,
        notes
            .replacen("type=\"body\" idx=\"3\"", "type=\"body\" idx=\"99\"", 1)
            .into_bytes(),
    );
    assert!(rejects_notes(package));

    let mut package = open_opc(&bytes, "ambiguous F-226 placeholder");
    let notes_part = "/custom/notes/notes1.xml";
    let notes = String::from_utf8(package.get_part(notes_part).unwrap().to_vec()).unwrap();
    let extra = r#"<p:sp><p:nvSpPr><p:cNvPr id="9" name="Ambiguous notes"/><p:cNvSpPr/><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>ambiguous</a:t></a:r></a:p></p:txBody></p:sp>"#;
    assert!(notes.contains("</p:spTree>"));
    package.set_part(
        notes_part,
        notes
            .replacen("</p:spTree>", &format!("{extra}</p:spTree>"), 1)
            .into_bytes(),
    );
    assert!(rejects_notes(package));

    let mut package = open_opc(&bytes, "duplicate F-226 handout master");
    let presentation_part = package.main_document_part().unwrap();
    package
        .get_or_create_part_rels(&presentation_part)
        .add_with_id(
            "handout-two",
            rel_types::HANDOUT_MASTER,
            "../custom/handouts/master.xml",
        );
    assert!(rejects_handout(package));

    let mut package = open_opc(&bytes, "external F-226 handout master");
    let presentation_part = package.main_document_part().unwrap();
    package
        .get_or_create_part_rels(&presentation_part)
        .items
        .iter_mut()
        .find(|relationship| relationship.rel_type == rel_types::HANDOUT_MASTER)
        .unwrap()
        .target_mode = Some("External".to_owned());
    assert!(rejects_handout(package));
}

#[test]
fn notes_and_handout_export_leave_the_opened_package_byte_identical() {
    let presentation = Presentation::from_bytes(&f226_fixture_bytes()).unwrap();
    let before = presentation.to_bytes().unwrap();
    presentation.to_notes_pdf_deterministic().unwrap();
    presentation
        .to_handout_pdf_deterministic(HandoutLayout::Nine)
        .unwrap();
    presentation.notes_page_pngs_deterministic(72.0).unwrap();
    presentation
        .handout_page_pngs_deterministic(HandoutLayout::Two, 72.0)
        .unwrap();
    assert_eq!(presentation.to_bytes().unwrap(), before);
}

#[cfg(all(feature = "digital-signatures", feature = "render"))]
const M21_AUDIO_WAV: &[u8] = b"RIFF\x26\0\0\0WAVEfmt \x10\0\0\0\x01\0\x01\0\x40\x1f\0\0\x40\x1f\0\0\x01\0\x08\0data\x02\0\0\0\x80\x80";

#[cfg(all(feature = "digital-signatures", feature = "render"))]
fn m21_representative_unsigned_deck_bytes_from_smartart(smartart_bytes: Option<&[u8]>) -> Vec<u8> {
    use rpptx::{Comment, CommentAuthor, Section};

    let mut presentation = Presentation::from_bytes(&m21_notes_handout_fixture_bytes())
        .expect("open source-built notes and handout deck");
    if let Some(smartart_bytes) = smartart_bytes {
        let smartart = Presentation::from_bytes(smartart_bytes).expect("open SmartArt source");
        presentation
            .transfer_smartart_slide_from(&smartart, 0, 0)
            .expect("transfer representative SmartArt slide");
    }
    let mut representative_slide = presentation.slide_mut(0).unwrap();
    let mut animated_shape = representative_slide
        .add_textbox(Emu(914_400), Emu(4_114_800), Emu(3_657_600), Emu(685_800))
        .unwrap();
    animated_shape
        .set_name("M21 representative animated text")
        .unwrap();
    animated_shape.set_text("M21 representative").unwrap();
    presentation
        .add_media(
            0,
            MediaKind::Audio,
            MediaSourceInput::Embedded(EmbeddedMediaInput {
                bytes: M21_AUDIO_WAV,
                filename: "m21-audio.wav",
                content_type: "audio/wav",
            }),
            MediaPoster {
                bytes: &valid_one_pixel_png(),
                filename: "m21-poster.png",
            },
            Emu(5_029_200),
            Emu(3_657_600),
            Emu(1_828_800),
            Emu(1_028_700),
            MediaPlaybackSettings {
                trigger: rpptx::MediaPlaybackTrigger::OnClick,
                ..MediaPlaybackSettings::default()
            },
        )
        .unwrap();
    presentation
        .add_comment_author(
            CommentAuthor::new(
                "{11111111-1111-1111-1111-111111111111}",
                "M21 reviewer",
                Some("M21"),
                "m21@example.test",
                "representative gate",
            )
            .unwrap(),
        )
        .unwrap();
    presentation
        .add_comment(
            0,
            Comment::new(
                "{22222222-2222-2222-2222-222222222222}",
                "{11111111-1111-1111-1111-111111111111}",
                "2026-09-02T11:00:00Z",
                "M21 representative comment",
            )
            .unwrap(),
        )
        .unwrap();
    let slide_ids = presentation
        .slides()
        .map(|slide| slide.id())
        .collect::<Vec<_>>();
    presentation
        .set_sections(vec![
            Section::new(
                "{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}",
                "Modern content",
                slide_ids,
            )
            .unwrap(),
        ])
        .unwrap();

    let mut package = open_opc(
        &presentation.to_bytes().unwrap(),
        "M21 representative timing source",
    );
    let presentation_part = package.main_document_part().unwrap();
    let model = CT_Presentation::from_xml(package.get_part(&presentation_part).unwrap()).unwrap();
    let slide_relationship = package
        .get_part_rels(&presentation_part)
        .unwrap()
        .get_by_id(&model.slide_ids[0].relationship_id)
        .unwrap();
    let slide_part = OpcPackage::resolve_rel_target(&presentation_part, &slide_relationship.target);
    let mut slide = CT_Slide::from_xml(package.get_part(&slide_part).unwrap()).unwrap();
    let animated_shape_id = slide
        .common_slide_data
        .shape_tree
        .children
        .iter()
        .find(|child| {
            child.non_visual_name().as_deref() == Some("M21 representative animated text")
        })
        .and_then(ShapeTreeChild::non_visual_id)
        .unwrap();
    slide.timing = Some(
        rpptx_oxml::timing::CT_Timing::from_xml(
            format!(
                r#"<p:timing xmlns:p="{P_NS}"><p:tnLst><p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="100" dur="1000" fill="hold"/><p:tgtEl><p:spTgt spid="{animated_shape_id}"/></p:tgtEl></p:cBhvr></p:animEffect></p:tnLst></p:timing>"#
            )
            .as_bytes(),
        )
        .unwrap(),
    );
    package.set_part(&slide_part, slide.to_xml().unwrap());

    let timed = Presentation::from_bytes(&package_bytes(package)).unwrap();
    timed
        .to_bytes_as(PresentationPackageClass::MacroEnabledPresentation)
        .unwrap()
}

#[cfg(all(feature = "digital-signatures", feature = "render"))]
fn m21_representative_unsigned_deck_bytes(include_smartart: bool) -> Vec<u8> {
    let smartart = include_smartart.then(|| authentic_smartart_oracle_source_bytes("list"));
    m21_representative_unsigned_deck_bytes_from_smartart(smartart.as_deref())
}

#[cfg(all(feature = "digital-signatures", feature = "render"))]
fn m21_minimal_smartart_fallback_deck_bytes() -> Vec<u8> {
    let minimal_smartart = smartart_rust_source_bytes("list");
    m21_representative_unsigned_deck_bytes_from_smartart(Some(&minimal_smartart))
}

#[cfg(all(feature = "digital-signatures", feature = "render"))]
fn m21_sign_deck(unsigned: &[u8]) -> Vec<u8> {
    let mut signed = Presentation::from_bytes(unsigned).unwrap();
    let (private_key, certificate) = f221_signing_material();
    let report = signed.sign(&private_key, &certificate).unwrap();
    assert!(report.cryptographically_valid);
    assert!(report.coverage_complete);
    signed.to_bytes().unwrap()
}

#[cfg(all(feature = "digital-signatures", feature = "render"))]
#[derive(Clone, Copy)]
enum M21SmartArtExpectation {
    MinimalFallback,
    Authentic,
}

#[cfg(all(feature = "digital-signatures", feature = "render"))]
fn m21_assert_representative_semantics(
    bytes: &[u8],
    smartart_expectation: M21SmartArtExpectation,
) -> Presentation {
    let opened = Presentation::from_bytes(bytes).expect("open representative modern deck bytes");
    assert!(opened.validate().is_empty());
    let saved = opened.to_bytes().expect("save representative modern deck");
    let reopened =
        Presentation::from_bytes(&saved).expect("reopen saved representative modern deck");
    assert!(reopened.validate().is_empty());

    let assert_loaded = |presentation: &Presentation, package_bytes: &[u8]| {
        assert_eq!(
            presentation.package_class().unwrap(),
            PresentationPackageClass::MacroEnabledPresentation
        );
        assert_eq!(
            presentation
                .slides()
                .map(|slide| slide.id())
                .collect::<Vec<_>>(),
            [256, 257, 258]
        );
        assert_eq!(presentation.len(), 3);
        assert_eq!(
            presentation
                .slide(0)
                .unwrap()
                .shapes()
                .filter_map(|shape| shape.text())
                .filter(|text| !text.is_empty())
                .collect::<Vec<_>>(),
            ["M21 slide one", "M21 representative"]
        );
        assert_eq!(
            presentation
                .slide(1)
                .unwrap()
                .shapes()
                .filter_map(|shape| shape.text())
                .filter(|text| !text.is_empty())
                .collect::<Vec<_>>(),
            ["M21 slide two"]
        );
        assert_eq!(
            presentation.slide(0).unwrap().notes_text().as_deref(),
            Some("M21 speaker note 1")
        );
        assert_eq!(
            presentation.slide(1).unwrap().notes_text().as_deref(),
            Some("M21 speaker note 2")
        );
        assert_eq!(presentation.slide(2).unwrap().notes_text(), None);

        let authors = presentation.comment_authors();
        assert_eq!(authors.len(), 1);
        assert_eq!(authors[0].id, "{11111111-1111-1111-1111-111111111111}");
        assert_eq!(authors[0].name, "M21 reviewer");
        assert_eq!(authors[0].initials.as_deref(), Some("M21"));
        assert_eq!(authors[0].user_id, "m21@example.test");
        assert_eq!(authors[0].provider_id, "representative gate");
        let comments = presentation.comments(0).unwrap();
        assert_eq!(comments.len(), 1);
        assert_eq!(comments[0].id, "{22222222-2222-2222-2222-222222222222}");
        assert_eq!(comments[0].author_id, authors[0].id);
        assert_eq!(comments[0].created, "2026-09-02T11:00:00Z");
        assert_eq!(comments[0].text(), "M21 representative comment");
        assert!(comments[0].replies().is_empty());
        assert_eq!(presentation.comments(1), None);
        assert_eq!(presentation.comments(2), None);

        let sections = presentation.sections();
        assert_eq!(sections.len(), 1);
        assert_eq!(
            sections[0].id.as_deref(),
            Some("{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}")
        );
        assert_eq!(sections[0].name.as_deref(), Some("Modern content"));
        assert_eq!(sections[0].slide_ids, [256, 257, 258]);

        let media = presentation.media(0).unwrap();
        assert_eq!(media.len(), 1);
        assert_eq!(media[0].slide_index, 0);
        assert_eq!(media[0].kind, MediaKind::Audio);
        let rpptx::MediaLocation::Embedded {
            part_name,
            content_type,
        } = &media[0].source
        else {
            panic!("representative audio must remain embedded")
        };
        assert_eq!(part_name, "/ppt/media/media1.wav");
        assert_eq!(content_type, "audio/wav");
        assert_eq!(media[0].settings, MediaPlaybackSettings::default());
        assert_eq!(
            media[0].diagnostics,
            [rpptx::MediaDiagnostic::MissingPlaybackTiming]
        );
        assert_eq!(
            presentation
                .extract_media(0, media[0].shape_id)
                .unwrap()
                .as_deref(),
            Some(M21_AUDIO_WAV)
        );

        let signatures = presentation.verify_signatures().unwrap();
        assert_eq!(signatures.len(), 1);
        assert!(signatures[0].cryptographically_valid);
        assert!(signatures[0].coverage_complete);

        assert!(presentation.smart_art(0).unwrap().is_empty());
        assert!(presentation.smart_art(1).unwrap().is_empty());
        let smartart = presentation.smart_art(2).unwrap();
        assert_eq!(smartart.len(), 1);
        assert_eq!(smartart[0].relationships.data, "smart-data");
        assert_eq!(smartart[0].relationships.layout, "smart-layout");
        assert_eq!(smartart[0].relationships.style, "smart-style");
        assert_eq!(smartart[0].relationships.colors, "smart-colors");
        assert!(smartart[0].relationships.drawing.is_none());
        let rpptx::DiagramPart::Parsed(data) = &smartart[0].data else {
            panic!("representative SmartArt data must parse")
        };
        assert_eq!(
            data.points()
                .iter()
                .filter_map(|point| point.text.as_ref().map(|_| point.model_id.as_str()))
                .collect::<Vec<_>>(),
            ["1", "2", "3"]
        );
        assert_eq!(
            data.points()
                .iter()
                .filter_map(|point| point.text.as_ref().map(|text| text.plain_text()))
                .collect::<Vec<_>>(),
            ["F220 list 1", "F220 list 2", "F220 list 3"]
        );
        let connections = data
            .connections()
            .iter()
            .map(|connection| {
                (
                    connection.model_id.as_str(),
                    connection.source_id.as_str(),
                    connection.destination_id.as_str(),
                )
            })
            .collect::<Vec<_>>();
        let rpptx::DiagramPart::Parsed(layout) = &smartart[0].layout else {
            panic!("representative SmartArt layout must parse")
        };
        let rpptx::DiagramPart::Parsed(style) = &smartart[0].style else {
            panic!("representative SmartArt style must parse")
        };
        let rpptx::DiagramPart::Parsed(colors) = &smartart[0].colors else {
            panic!("representative SmartArt colours must parse")
        };
        assert!(smartart[0].drawing.is_none());
        let (expected_connections, expected_layout_id, expected_style_id, expected_colors_id) =
            match smartart_expectation {
                M21SmartArtExpectation::MinimalFallback => (
                    vec![
                        ("201", "0", "1"),
                        ("202", "1", "101"),
                        ("203", "2", "102"),
                        ("204", "3", "103"),
                        ("205", "1", "2"),
                        ("206", "1", "3"),
                    ],
                    "urn:f220:list",
                    "urn:f220:style",
                    "urn:f220:colors",
                ),
                M21SmartArtExpectation::Authentic => (
                    vec![("4", "0", "1"), ("5", "0", "2"), ("6", "0", "3")],
                    "urn:microsoft.com/office/officeart/2005/8/layout/list1",
                    "urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1",
                    "urn:microsoft.com/office/officeart/2005/8/colors/accent1_1",
                ),
            };
        assert_eq!(connections, expected_connections);
        assert_eq!(layout.unique_id.as_deref(), Some(expected_layout_id));
        match smartart_expectation {
            M21SmartArtExpectation::MinimalFallback => assert_eq!(
                layout.family,
                rpptx_oxml::diagram::DiagramLayoutFamily::Unsupported(
                    expected_layout_id.to_owned()
                )
            ),
            M21SmartArtExpectation::Authentic => assert_eq!(
                layout.family,
                rpptx_oxml::diagram::DiagramLayoutFamily::List
            ),
        }
        assert_eq!(style.unique_id.as_deref(), Some(expected_style_id));
        assert_eq!(colors.unique_id.as_deref(), Some(expected_colors_id));

        let package = open_opc(package_bytes, "M21 representative semantic contract");
        assert_eq!(package.get_part(part_name).unwrap(), M21_AUDIO_WAV);
        assert_eq!(
            package.content_types.content_type_for(part_name),
            Some("audio/wav")
        );
        let presentation_part = package.main_document_part().unwrap();
        let presentation_model =
            CT_Presentation::from_xml(package.get_part(&presentation_part).unwrap()).unwrap();
        let slide_relationship = package
            .get_part_rels(&presentation_part)
            .unwrap()
            .get_by_id(&presentation_model.slide_ids[0].relationship_id)
            .unwrap();
        let slide_part =
            OpcPackage::resolve_rel_target(&presentation_part, &slide_relationship.target);
        let slide_relationships = package.get_part_rels(&slide_part).unwrap();
        let poster_relationship = slide_relationships
            .get_by_id(media[0].poster_relationship_id.as_deref().unwrap())
            .unwrap();
        assert_eq!(poster_relationship.rel_type, rel_types::IMAGE);
        assert_eq!(poster_relationship.target_mode, None);
        let poster_part = OpcPackage::resolve_rel_target(&slide_part, &poster_relationship.target);
        assert_eq!(
            package.get_part(&poster_part).unwrap(),
            valid_one_pixel_png()
        );
        assert_eq!(
            package.content_types.content_type_for(&poster_part),
            Some("image/png")
        );
        let slide = CT_Slide::from_xml(package.get_part(&slide_part).unwrap()).unwrap();
        let animated_shape_id = slide
            .common_slide_data
            .shape_tree
            .children
            .iter()
            .find(|child| {
                child.non_visual_name().as_deref() == Some("M21 representative animated text")
            })
            .and_then(ShapeTreeChild::non_visual_id)
            .unwrap();
        let timing = slide.timing.as_ref().expect("representative typed timing");
        assert_eq!(timing.nodes().len(), 1);
        let rpptx_oxml::timing::TimingNode::Effect(effect) = &timing.nodes()[0] else {
            panic!("representative timing must be one fade effect")
        };
        assert_eq!(effect.common.id, 100);
        assert_eq!(
            effect.common.duration,
            rpptx_oxml::timing::TimingDuration::Finite(1_000)
        );
        assert_eq!(
            effect.target,
            rpptx_oxml::timing::TimingTarget::Shape(animated_shape_id)
        );
        assert_eq!(effect.transition.as_deref(), Some("in"));
        assert_eq!(effect.filter.as_deref(), Some("fade"));

        let (input, layout_result) = presentation.render_deterministic().unwrap();
        let visible_tokens = m21_normalized_tokens(&m21_visible_page_text(&layout_result.pages[2]));
        match smartart_expectation {
            M21SmartArtExpectation::MinimalFallback => {
                assert_eq!(visible_tokens, ["Unsupported", "SmartArt"]);
                assert!(
                    input.slides[2]
                        .diagnostics
                        .iter()
                        .any(|diagnostic| diagnostic.message.contains("unsupported SmartArt"))
                );
            }
            M21SmartArtExpectation::Authentic => {
                assert_eq!(visible_tokens, M21_RUST_STATIC_TOKENS[2]);
                assert!(
                    input.slides[2]
                        .diagnostics
                        .iter()
                        .all(|diagnostic| !diagnostic.message.contains("unsupported SmartArt"))
                );
            }
        }
    };

    assert_loaded(&opened, bytes);
    assert_loaded(&reopened, &saved);
    reopened
}

#[cfg(all(feature = "digital-signatures", feature = "render"))]
#[test]
fn m21_minimal_smartart_source_classifies_the_unsupported_fallback() {
    let bytes = m21_minimal_smartart_fallback_deck_bytes();
    assert_eq!(sha256_bytes(&bytes), M21_UNSIGNED_SOURCE_SHA256);
    let presentation = Presentation::from_bytes(&bytes).unwrap();
    let (_, layout) = presentation.render_deterministic().unwrap();
    assert_eq!(
        m21_normalized_tokens(&m21_visible_page_text(&layout.pages[2])),
        ["Unsupported", "SmartArt"]
    );
    assert_eq!(
        m21_full_page_ink_bounds(&oxml_pdf::render_page_to_png(&layout, 2, 150.0).unwrap()).len(),
        1
    );
}

#[cfg(all(feature = "digital-signatures", feature = "render"))]
#[test]
#[ignore = "macOS reference-only generator requiring pinned external SmartArt resources"]
fn m21_write_corrected_representative_source_candidates() {
    let output_dir = PathBuf::from(
        std::env::var_os("RPPTX_M21_CORRECTED_SOURCE_OUTPUT_DIR")
            .expect("set RPPTX_M21_CORRECTED_SOURCE_OUTPUT_DIR to an output-only directory"),
    );
    fs::create_dir_all(&output_dir).unwrap();
    let unsigned = m21_representative_unsigned_deck_bytes(true);
    let signed = m21_sign_deck(&unsigned);
    let unsigned_path = output_dir.join("m21-representative-corrected-unsigned.pptm");
    let signed_path = output_dir.join("m21-representative-corrected-signed.pptm");
    fs::write(&unsigned_path, &unsigned).unwrap();
    fs::write(&signed_path, &signed).unwrap();

    let reopened = Presentation::from_bytes(&signed).unwrap();
    assert!(reopened.validate().is_empty());
    let signatures = reopened.verify_signatures().unwrap();
    assert_eq!(signatures.len(), 1);
    assert!(signatures[0].cryptographically_valid);
    assert!(signatures[0].coverage_complete);
    let smartart = reopened.smart_art(2).unwrap();
    assert_eq!(smartart.len(), 1);
    assert_eq!(smartart[0].relationships.data, "smart-data");
    assert_eq!(smartart[0].relationships.layout, "smart-layout");
    assert_eq!(smartart[0].relationships.style, "smart-style");
    assert_eq!(smartart[0].relationships.colors, "smart-colors");
    assert!(smartart[0].relationships.drawing.is_none());
    assert!(matches!(smartart[0].data, rpptx::DiagramPart::Parsed(_)));
    assert!(matches!(smartart[0].layout, rpptx::DiagramPart::Parsed(_)));
    assert!(matches!(smartart[0].style, rpptx::DiagramPart::Parsed(_)));
    assert!(matches!(smartart[0].colors, rpptx::DiagramPart::Parsed(_)));
    let (_, layout) = reopened.render_deterministic().unwrap();
    assert_eq!(
        m21_normalized_tokens(&m21_visible_page_text(&layout.pages[2])),
        M21_RUST_STATIC_TOKENS[2]
    );
    assert!(
        !m21_full_page_ink_bounds(&oxml_pdf::render_page_to_png(&layout, 2, 150.0).unwrap())
            .is_empty()
    );
    eprintln!(
        "M21 corrected source candidates: unsigned={} sha256={} signed={} sha256={}",
        unsigned_path.display(),
        sha256(&unsigned_path),
        signed_path.display(),
        sha256(&signed_path),
    );
}

#[cfg(all(feature = "digital-signatures", feature = "render"))]
#[test]
fn m21_portable_representative_package_round_trips_and_outputs_are_sensitive() {
    let unsigned = m21_minimal_smartart_fallback_deck_bytes();
    let bytes = m21_sign_deck(&unsigned);
    m21_assert_signature_surrogate(&unsigned, &bytes);
    let reopened =
        m21_assert_representative_semantics(&bytes, M21SmartArtExpectation::MinimalFallback);

    let (_, static_layout) = reopened.render_deterministic().unwrap();
    assert_eq!(static_layout.pages.len(), 3);
    let static_tokens = static_layout
        .pages
        .iter()
        .map(|page| m21_normalized_tokens(&m21_visible_page_text(page)))
        .collect::<Vec<_>>();
    let portable_static_tokens: [&[&str]; 3] = [
        M21_RUST_STATIC_TOKENS[0],
        M21_RUST_STATIC_TOKENS[1],
        &["Unsupported", "SmartArt"],
    ];
    assert_eq!(static_tokens, portable_static_tokens);
    for (page, expected) in static_layout.pages.iter().zip(portable_static_tokens) {
        m21_assert_text_mutations_fail(&m21_visible_page_text(page), expected);
    }
    assert_eq!(static_tokens[2], ["Unsupported", "SmartArt"]);
    let static_png = oxml_pdf::render_page_to_png(&static_layout, 0, 150.0).unwrap();
    let smartart_png = oxml_pdf::render_page_to_png(&static_layout, 2, 150.0).unwrap();
    assert!(!m21_full_page_ink_bounds(&smartart_png).is_empty());
    for (page_index, source_text, replacement) in [
        (0, "M21 slide one", "M21 slide one extra"),
        (1, "M21 slide two", "M21 slide two two"),
    ] {
        let mut page_mutated = Presentation::from_bytes(&bytes).unwrap();
        let shape_index = page_mutated
            .slide(page_index)
            .unwrap()
            .shapes()
            .position(|shape| shape.text().as_deref() == Some(source_text))
            .unwrap();
        page_mutated
            .slide_mut(page_index)
            .unwrap()
            .shape_mut(shape_index)
            .unwrap()
            .set_text(replacement)
            .unwrap();
        let (_, page_mutated_layout) = page_mutated.render_deterministic().unwrap();
        assert_ne!(
            m21_normalized_tokens(&m21_visible_page_text(
                &page_mutated_layout.pages[page_index]
            )),
            portable_static_tokens[page_index]
        );
        assert_ne!(
            oxml_pdf::render_page_to_png(&page_mutated_layout, page_index, 150.0).unwrap(),
            oxml_pdf::render_page_to_png(&static_layout, page_index, 150.0).unwrap()
        );
    }
    let mut smartart_mutated = Presentation::from_bytes(&bytes).unwrap();
    smartart_mutated
        .slide_mut(2)
        .unwrap()
        .add_textbox(Emu(914_400), Emu(4_114_800), Emu(3_657_600), Emu(685_800))
        .unwrap()
        .set_text("F220 list extra")
        .unwrap();
    let (_, smartart_mutated_layout) = smartart_mutated.render_deterministic().unwrap();
    assert_ne!(
        m21_normalized_tokens(&m21_visible_page_text(&smartart_mutated_layout.pages[2])),
        portable_static_tokens[2]
    );
    assert_ne!(
        oxml_pdf::render_page_to_png(&smartart_mutated_layout, 2, 150.0).unwrap(),
        smartart_png
    );
    let animation = reopened
        .export_animation_deterministic(
            &[rpptx::AnimationSegment {
                slide_index: 0,
                duration_ms: 1_000,
                click_count: 0,
                transition: rpptx::AnimationTransition::None,
            }],
            rpptx::AnimationExportOptions {
                frame_rate: 2,
                width_px: 320,
                height_px: 180,
                format: rpptx::AnimationFormat::Gif {
                    loop_behavior: rpptx::GifLoopBehavior::Once,
                },
                media_fallback: MediaFallbackPolicy::PosterFrame,
            },
        )
        .unwrap();
    assert!(animation.bytes.starts_with(b"GIF89a"));
    let mut timeline_ink = Vec::new();
    for elapsed_ms in [0, 495, 1_000] {
        let frame = reopened
            .render_timeline_deterministic(
                0,
                TimelinePosition {
                    elapsed_ms,
                    click_count: 0,
                },
                None,
            )
            .unwrap();
        let expected_text = if elapsed_ms == 0 {
            &["M21", "slide", "one"][..]
        } else {
            &["M21", "slide", "one", "M21", "representative"][..]
        };
        m21_assert_text_mutations_fail(&m21_visible_page_text(&frame.page), expected_text);
        let layout = m21_timeline_layout(&frame, &static_layout.fonts);
        let png = oxml_pdf::render_page_to_png(&layout, 0, 150.0).unwrap();
        timeline_ink.push((
            m21_ink_mass(&png, (150, 140, 350, 100)),
            m21_ink_mass(&png, (150, 650, 600, 100)),
        ));
    }
    assert!(timeline_ink.iter().all(|(stable, _)| *stable > 100_000));
    assert_eq!(timeline_ink[0].0, timeline_ink[1].0);
    assert_eq!(timeline_ink[1].0, timeline_ink[2].0);
    assert_eq!(timeline_ink[0].1, 0);
    assert!(timeline_ink[1].1 > 0);
    assert!(timeline_ink[2].1 > timeline_ink[1].1);
    let notes_pdf = reopened.to_notes_pdf_deterministic().unwrap();
    let notes_pngs = reopened.notes_page_pngs_deterministic(72.0).unwrap();
    assert_eq!(f226_pdf_page_count(&notes_pdf, "m21-notes"), 3);
    assert_eq!(notes_pngs.len(), 3);
    let handout_pdf = reopened
        .to_handout_pdf_deterministic(HandoutLayout::Three)
        .unwrap();
    let handout_pngs = reopened
        .handout_page_pngs_deterministic(HandoutLayout::Three, 72.0)
        .unwrap();
    assert_eq!(f226_pdf_page_count(&handout_pdf, "m21-handout"), 1);
    assert_eq!(handout_pngs.len(), 1);
    assert_eq!(m21_handout_thumbnail_bounds(&handout_pngs[0]).len(), 3);
    let mut mutated = Presentation::from_bytes(&bytes).unwrap();
    let animated_shape_index = mutated
        .slide(0)
        .unwrap()
        .shapes()
        .position(|shape| shape.text().as_deref() == Some("M21 representative"))
        .unwrap();
    mutated
        .slide_mut(0)
        .unwrap()
        .shape_mut(animated_shape_index)
        .unwrap()
        .set_text("M21 sensitivity mutation")
        .unwrap();
    let (_, mutated_layout) = mutated.render_deterministic().unwrap();
    assert_ne!(
        oxml_pdf::render_page_to_png(&mutated_layout, 0, 150.0).unwrap(),
        static_png
    );
    assert_ne!(
        mutated
            .export_animation_deterministic(
                &[rpptx::AnimationSegment {
                    slide_index: 0,
                    duration_ms: 1_000,
                    click_count: 0,
                    transition: rpptx::AnimationTransition::None,
                }],
                rpptx::AnimationExportOptions {
                    frame_rate: 2,
                    width_px: 320,
                    height_px: 180,
                    format: rpptx::AnimationFormat::Gif {
                        loop_behavior: rpptx::GifLoopBehavior::Once,
                    },
                    media_fallback: MediaFallbackPolicy::PosterFrame,
                },
            )
            .unwrap()
            .bytes,
        animation.bytes
    );
    assert_ne!(mutated.to_notes_pdf_deterministic().unwrap(), notes_pdf);
    assert_ne!(
        mutated
            .to_handout_pdf_deterministic(HandoutLayout::Three)
            .unwrap(),
        handout_pdf
    );
}

#[cfg(all(feature = "digital-signatures", feature = "render"))]
#[derive(Clone, Copy)]
struct M21RecordedArtifact {
    file_name: &'static str,
    sha256: &'static str,
    source_sha256: &'static str,
}

#[cfg(all(feature = "digital-signatures", feature = "render"))]
struct M21RecordedMovieSample {
    movie_time: i64,
    rust_time_ms: u64,
    observed_visible_text: &'static [&'static str],
    observed_ink_band_count: usize,
}

#[cfg(all(feature = "digital-signatures", feature = "render"))]
const M21_UNSIGNED_SOURCE_SHA256: &str =
    "00c98ad616bd4cc851065dca65c7252d964c1b66eedbf620027ace6341fbabde";

#[cfg(all(feature = "digital-signatures", feature = "render"))]
const M21_LEGACY_STATIC_PDF_SHA256: &str =
    "42800576d3e9fc3c265ad57fb863d0aef7a5a670c81ee342b96b43f3b2a179de";

#[cfg(all(feature = "digital-signatures", feature = "render"))]
struct M21CorrectedPowerPointManifest {
    version: &'static str,
    bundle_build: &'static str,
    app_build: &'static str,
    signed: bool,
    signed_open_no_repair: bool,
    observed_active_name: &'static str,
    observed_source_sha256: &'static str,
    signed_source: M21RecordedArtifact,
    static_pdf: M21RecordedArtifact,
    movie: M21RecordedArtifact,
    notes_pdf: M21RecordedArtifact,
    handout_pdf: M21RecordedArtifact,
}

#[cfg(all(feature = "digital-signatures", feature = "render"))]
const M21_CORRECTED_POWERPOINT_MANIFEST: M21CorrectedPowerPointManifest =
    M21CorrectedPowerPointManifest {
        version: "16.104",
        bundle_build: "16.104.25121423",
        app_build: "1214",
        signed: true,
        signed_open_no_repair: true,
        observed_active_name: "m21-corrected-signed.pptm",
        observed_source_sha256: "74fe838af835fbf9852d232d1eb39683bfbb1381b86095073e9e96974b50aac9",
        signed_source: M21RecordedArtifact {
            file_name: "m21-corrected-signed.pptm",
            sha256: "74fe838af835fbf9852d232d1eb39683bfbb1381b86095073e9e96974b50aac9",
            source_sha256: "74fe838af835fbf9852d232d1eb39683bfbb1381b86095073e9e96974b50aac9",
        },
        static_pdf: M21RecordedArtifact {
            file_name: "m21-corrected-static.pdf",
            sha256: "aebe97df20d029a611afa935fad0e96653e0b515396ce7ec1f5e2c665d92f8de",
            source_sha256: "74fe838af835fbf9852d232d1eb39683bfbb1381b86095073e9e96974b50aac9",
        },
        movie: M21RecordedArtifact {
            file_name: "m21-corrected-signed.mp4",
            sha256: "4643c6cb25222b343067364a8983673c79962e32378809206b7f9e6f5306e5e9",
            source_sha256: "74fe838af835fbf9852d232d1eb39683bfbb1381b86095073e9e96974b50aac9",
        },
        notes_pdf: M21RecordedArtifact {
            file_name: "m21-corrected-notes.pdf",
            sha256: "d940316865a28e626c2cc7756d9bef4f132c516d03cba63387e1f6f0ca0dba2a",
            source_sha256: "74fe838af835fbf9852d232d1eb39683bfbb1381b86095073e9e96974b50aac9",
        },
        handout_pdf: M21RecordedArtifact {
            file_name: "m21-corrected-handout.pdf",
            sha256: "77345fd00914bb2b233bf548530bd2f6de05c25b53a08cd7392bf38be696d05f",
            source_sha256: "74fe838af835fbf9852d232d1eb39683bfbb1381b86095073e9e96974b50aac9",
        },
    };

#[cfg(all(feature = "digital-signatures", feature = "render"))]
const M21_POWERPOINT_MOVIE_SAMPLES: [M21RecordedMovieSample; 3] = [
    M21RecordedMovieSample {
        movie_time: 0,
        rust_time_ms: 0,
        observed_visible_text: &["M21", "slide", "one"],
        observed_ink_band_count: 1,
    },
    M21RecordedMovieSample {
        movie_time: 297,
        rust_time_ms: 495,
        observed_visible_text: &["M21", "slide", "one", "M21", "representative"],
        observed_ink_band_count: 2,
    },
    M21RecordedMovieSample {
        movie_time: 594,
        rust_time_ms: 990,
        observed_visible_text: &["M21", "slide", "one", "M21", "representative"],
        observed_ink_band_count: 2,
    },
];

#[cfg(all(feature = "digital-signatures", feature = "render"))]
const M21_POWERPOINT_STATIC_TOKENS: [&[&str]; 3] = [
    &["M21", "slide", "one", "M21", "representative"],
    &["M21", "slide", "two"],
    &[
        "F220", "list", "1", "F220", "list", "2", "F220", "list", "3",
    ],
];

#[cfg(all(feature = "digital-signatures", feature = "render"))]
const M21_RUST_STATIC_TOKENS: [&[&str]; 3] = [
    &["M21", "slide", "one", "M21", "representative"],
    &["M21", "slide", "two"],
    &[
        "F220", "list", "1", "F220", "list", "2", "F220", "list", "3",
    ],
];

#[cfg(all(feature = "digital-signatures", feature = "render"))]
const M21_POWERPOINT_NOTES_TOKENS: [&[&str]; 3] = [
    &["M21", "speaker", "note", "1"],
    &["M21", "speaker", "note", "2"],
    &[
        "F220", "list", "1", "F220", "list", "2", "F220", "list", "3",
    ],
];

#[cfg(all(feature = "digital-signatures", feature = "render"))]
const M21_POWERPOINT_HANDOUT_TOKENS: &[&str] = &[
    "9/2/26",
    "M21",
    "slide",
    "one",
    "M21",
    "representative",
    "1",
    "M21",
    "slide",
    "two",
    "2",
    "F220",
    "list",
    "1",
    "F220",
    "list",
    "2",
    "F220",
    "list",
    "3",
    "3",
    "1",
];

#[cfg(all(feature = "digital-signatures", feature = "render"))]
const M21_RUST_NOTES_TOKENS: [&[&str]; 3] = [
    &[
        "representative",
        "one",
        "slide",
        "M21",
        "2030-01-02",
        "M21",
        "F-226",
        "header",
        "M21",
        "speaker",
        "note",
        "1",
        "F-226",
        "footer",
        "1",
    ],
    &[
        "two",
        "slide",
        "2030-01-02",
        "M21",
        "F-226",
        "header",
        "M21",
        "speaker",
        "note",
        "2",
        "F-226",
        "footer",
        "2",
    ],
    &[
        "F-226",
        "footer",
        "3",
        "list",
        "1",
        "2",
        "list",
        "F220",
        "F220",
        "list",
        "2030-01-02",
        "F220",
        "F-226",
        "header",
        "3",
    ],
];

#[cfg(all(feature = "digital-signatures", feature = "render"))]
const M21_RUST_HANDOUT_TOKENS: &[&str] = &[
    "1",
    "footer",
    "handout",
    "226",
    "-",
    "F",
    "2",
    "list",
    "3",
    "list",
    "F220",
    "F220",
    "1",
    "list",
    "F220",
    "two",
    "representative",
    "M21",
    "slide",
    "M21",
    "1",
    "2",
    "3",
    "one",
    "slide",
    "M21",
    "02",
    "-",
    "01",
    "-",
    "2030",
    "header",
    "handout",
    "226",
    "-",
    "F",
];

#[cfg(all(feature = "digital-signatures", feature = "render"))]
fn m21_assert_signature_surrogate(unsigned_bytes: &[u8], signed_bytes: &[u8]) {
    let unsigned = open_opc(unsigned_bytes, "M21 unsigned signature surrogate");
    let signed = open_opc(signed_bytes, "M21 signed signature surrogate");
    let signature_parts = ["/_xmlsignatures/origin.sigs", "/_xmlsignatures/sig1.xml"];
    let mut signed_only_parts = signed
        .parts
        .keys()
        .filter(|name| !unsigned.parts.contains_key(*name))
        .map(String::as_str)
        .collect::<Vec<_>>();
    signed_only_parts.sort_unstable();
    assert_eq!(signed_only_parts, signature_parts);
    assert_eq!(signed.parts[signature_parts[0]], []);
    assert!(!signed.parts[signature_parts[1]].is_empty());
    assert_eq!(
        unsigned.parts.len() + signature_parts.len(),
        signed.parts.len()
    );
    for (name, data) in &unsigned.parts {
        assert_eq!(
            signed.parts.get(name),
            Some(data),
            "non-signature part changed: {name}"
        );
    }

    assert_eq!(
        unsigned.content_types.defaults,
        signed.content_types.defaults
    );
    let mut signed_overrides = signed.content_types.overrides.clone();
    assert_eq!(
        signed_overrides.remove(signature_parts[0]),
        Some("application/vnd.openxmlformats-package.digital-signature-origin".to_owned())
    );
    assert_eq!(
        signed_overrides.remove(signature_parts[1]),
        Some(
            "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml".to_owned()
        )
    );
    assert_eq!(signed_overrides, unsigned.content_types.overrides);

    let non_signature_package_rels = signed
        .package_rels
        .items
        .iter()
        .filter(|relationship| relationship.rel_type != rel_types::DIGITAL_SIGNATURE_ORIGIN)
        .cloned()
        .collect::<Vec<_>>();
    assert_eq!(non_signature_package_rels, unsigned.package_rels.items);
    let signature_origin = signed
        .package_rels
        .items
        .iter()
        .filter(|relationship| relationship.rel_type == rel_types::DIGITAL_SIGNATURE_ORIGIN)
        .collect::<Vec<_>>();
    assert_eq!(signature_origin.len(), 1);
    assert_eq!(signature_origin[0].id, "rId5");
    assert_eq!(signature_origin[0].target, "_xmlsignatures/origin.sigs");
    assert_eq!(signature_origin[0].target_mode, None);

    let mut signed_only_rel_owners = signed
        .part_rels
        .keys()
        .filter(|owner| !unsigned.part_rels.contains_key(*owner))
        .map(String::as_str)
        .collect::<Vec<_>>();
    signed_only_rel_owners.sort_unstable();
    assert_eq!(signed_only_rel_owners, ["/_xmlsignatures/origin.sigs"]);
    for (owner, relationships) in &unsigned.part_rels {
        assert_eq!(
            signed.part_rels.get(owner).map(|signed| &signed.items),
            Some(&relationships.items),
            "non-signature relationships changed: {owner}"
        );
    }
    let signature_rels = &signed.part_rels["/_xmlsignatures/origin.sigs"].items;
    assert_eq!(signature_rels.len(), 1);
    assert_eq!(signature_rels[0].id, "rId0");
    assert_eq!(signature_rels[0].rel_type, rel_types::DIGITAL_SIGNATURE);
    assert_eq!(signature_rels[0].target, "sig1.xml");
    assert_eq!(signature_rels[0].target_mode, None);

    for required_part in [
        "/ppt/presentation.xml",
        "/ppt/media/media1.wav",
        "/ppt/theme/theme1.xml",
        "/ppt/theme/theme2.xml",
        "/ppt/diagrams/data1.xml",
        "/ppt/diagrams/layout1.xml",
        "/ppt/diagrams/quickStyle1.xml",
        "/ppt/diagrams/colors1.xml",
        "/ppt/slides/slide1.xml",
        "/ppt/notesMasters/notesMaster1.xml",
        "/ppt/notesSlides/notesSlide1.xml",
        "/ppt/handoutMasters/handoutMaster1.xml",
    ] {
        assert_eq!(signed.parts[required_part], unsigned.parts[required_part]);
    }
    assert!(String::from_utf8_lossy(&signed.parts["/ppt/slides/slide1.xml"]).contains("<p:timing"));
    for relationship_owner in [
        "/ppt/presentation.xml",
        "/ppt/slides/slide1.xml",
        "/ppt/notesMasters/notesMaster1.xml",
        "/ppt/notesSlides/notesSlide1.xml",
        "/ppt/handoutMasters/handoutMaster1.xml",
    ] {
        assert_eq!(
            signed.part_rels[relationship_owner].items,
            unsigned.part_rels[relationship_owner].items
        );
    }
}

#[cfg(all(feature = "digital-signatures", feature = "render"))]
#[test]
#[ignore = "requires RPPTX_M21_CORRECTED_POWERPOINT_ORACLE_DIR"]
fn m21_corrected_signed_outputs_match_powerpoint() {
    let manifest = &M21_CORRECTED_POWERPOINT_MANIFEST;
    assert_eq!(manifest.version, POWERPOINT_VERSION);
    assert_eq!(manifest.bundle_build, POWERPOINT_BUILD);
    assert_eq!(manifest.app_build, POWERPOINT_APP_BUILD);
    assert!(manifest.signed);
    assert!(manifest.signed_open_no_repair);
    assert_eq!(
        manifest.observed_active_name,
        manifest.signed_source.file_name
    );
    assert_eq!(
        manifest.observed_source_sha256,
        manifest.signed_source.sha256
    );
    assert_eq!(
        manifest.static_pdf.source_sha256,
        manifest.signed_source.sha256
    );
    assert_eq!(manifest.movie.source_sha256, manifest.signed_source.sha256);
    assert_eq!(
        manifest.notes_pdf.source_sha256,
        manifest.signed_source.sha256
    );
    assert_eq!(
        manifest.handout_pdf.source_sha256,
        manifest.signed_source.sha256
    );
    let oracle_dir = PathBuf::from(
        std::env::var_os("RPPTX_M21_CORRECTED_POWERPOINT_ORACLE_DIR")
            .expect("set RPPTX_M21_CORRECTED_POWERPOINT_ORACLE_DIR"),
    );
    let source = oracle_dir.join(manifest.signed_source.file_name);
    let static_pdf = oracle_dir.join(manifest.static_pdf.file_name);
    let movie = oracle_dir.join(manifest.movie.file_name);
    let notes_pdf = oracle_dir.join(manifest.notes_pdf.file_name);
    let handout_pdf = oracle_dir.join(manifest.handout_pdf.file_name);
    for path in [&source, &static_pdf, &movie, &notes_pdf, &handout_pdf] {
        assert!(
            path.is_file(),
            "missing corrected oracle {}",
            path.display()
        );
    }
    assert_eq!(sha256(&source), manifest.signed_source.sha256);
    assert_eq!(sha256(&static_pdf), manifest.static_pdf.sha256);
    assert_eq!(sha256(&movie), manifest.movie.sha256);
    assert_eq!(sha256(&notes_pdf), manifest.notes_pdf.sha256);
    assert_eq!(sha256(&handout_pdf), manifest.handout_pdf.sha256);

    let source_bytes = fs::read(&source).unwrap();
    let presentation =
        m21_assert_representative_semantics(&source_bytes, M21SmartArtExpectation::Authentic);
    let (_, static_layout) = presentation.render_deterministic().unwrap();
    let powerpoint_static_tokens =
        m21_pdf_token_pages(&fs::read(&static_pdf).unwrap(), "m21-corrected-static-text");
    assert_eq!(powerpoint_static_tokens, M21_POWERPOINT_STATIC_TOKENS);

    let root =
        std::env::temp_dir().join(format!("rpptx-m21-corrected-oracle-{}", std::process::id()));
    fs::create_dir_all(&root).unwrap();
    let powerpoint_pages = m21_pdf_page_pngs(&static_pdf, &root.join("m21-corrected-static"), 150)
        .iter()
        .map(|page| m21_normalize_movie_frame(page))
        .collect::<Vec<_>>();
    assert_eq!(powerpoint_pages.len(), 3);
    let mut static_scores = Vec::new();
    for page_index in 0..3 {
        let rust_page = oxml_pdf::render_page_to_png(&static_layout, page_index, 150.0).unwrap();
        let rust_page = if page_index == 0 {
            m21_mask_audio_poster(&rust_page)
        } else {
            rust_page
        };
        let powerpoint_page = if page_index == 0 {
            m21_mask_audio_poster(&powerpoint_pages[page_index])
        } else {
            powerpoint_pages[page_index].clone()
        };
        let rust_text = m21_visible_page_text(&static_layout.pages[page_index]);
        m21_assert_text_mutations_fail(&rust_text, M21_RUST_STATIC_TOKENS[page_index]);
        let comparison = m21_compare_full_page_ink(&rust_page, &powerpoint_page);
        assert!(
            m21_static_page_gate(true, &rust_page, &powerpoint_page, &comparison),
            "corrected static page {}: {comparison:?}",
            page_index + 1
        );
        let shifted = m21_shift_ink_up(&rust_page, 7);
        let shifted_comparison = m21_compare_full_page_ink(&shifted, &powerpoint_page);
        assert!(!m21_static_page_gate(
            true,
            &shifted,
            &powerpoint_page,
            &shifted_comparison,
        ));
        let solidified = m21_solidify_first_ink_region(&powerpoint_page);
        let solid_comparison = m21_compare_full_page_ink(&rust_page, &solidified);
        assert!(!m21_static_page_gate(
            true,
            &rust_page,
            &solidified,
            &solid_comparison,
        ));
        static_scores.push((
            page_index + 1,
            smartart_png_ssim(&rust_page, &powerpoint_page),
            comparison,
        ));
    }

    let mut movie_text_mutated = Presentation::from_bytes(&source_bytes).unwrap();
    let movie_title_index = movie_text_mutated
        .slide(0)
        .unwrap()
        .shapes()
        .position(|shape| shape.text().as_deref() == Some("M21 slide one"))
        .unwrap();
    movie_text_mutated
        .slide_mut(0)
        .unwrap()
        .shape_mut(movie_title_index)
        .unwrap()
        .set_text("M21 slide one extra")
        .unwrap();
    let (_, movie_text_mutated_static) = movie_text_mutated.render_deterministic().unwrap();
    let mut movie_scores = Vec::new();
    for sample in &M21_POWERPOINT_MOVIE_SAMPLES {
        let movie_frame = root.join(format!("m21-corrected-movie-{}.png", sample.movie_time));
        let extraction =
            extract_powerpoint_movie_frame(&movie, &movie_frame, sample.movie_time, 600);
        assert!(
            extraction.status.success(),
            "extract corrected movie frame: {}",
            String::from_utf8_lossy(&extraction.stderr)
        );
        let powerpoint_frame =
            m21_mask_audio_poster(&m21_normalize_movie_frame(&fs::read(&movie_frame).unwrap()));
        let rust_frame = presentation
            .render_timeline_deterministic(
                0,
                TimelinePosition {
                    elapsed_ms: sample.rust_time_ms,
                    click_count: 0,
                },
                None,
            )
            .unwrap();
        let rust_movie_text = m21_visible_page_text(&rust_frame.page);
        m21_assert_text_mutations_fail(&rust_movie_text, sample.observed_visible_text);
        let rust_layout = m21_timeline_layout(&rust_frame, &static_layout.fonts);
        let rust_frame_png = m21_mask_audio_poster(
            &oxml_pdf::render_page_to_png(&rust_layout, 0, F214_POWERPOINT_COMPARISON_DPI).unwrap(),
        );
        let comparison = m21_compare_full_page_ink(&rust_frame_png, &powerpoint_frame);
        let movie_text_matches = m21_tokens_match(&rust_movie_text, sample.observed_visible_text)
            && comparison.expected.len() == sample.observed_ink_band_count;
        assert!(m21_visual_gate(movie_text_matches, &comparison));

        let shifted_powerpoint = m21_shift_ink_up(&powerpoint_frame, 7);
        let shifted_comparison = m21_compare_full_page_ink(&rust_frame_png, &shifted_powerpoint);
        assert!(!m21_visual_gate(movie_text_matches, &shifted_comparison));
        let solid_powerpoint = m21_solidify_first_ink_region(&powerpoint_frame);
        let solid_comparison = m21_compare_full_page_ink(&rust_frame_png, &solid_powerpoint);
        assert!(!m21_visual_gate(movie_text_matches, &solid_comparison));

        let mutated_frame = movie_text_mutated
            .render_timeline_deterministic(
                0,
                TimelinePosition {
                    elapsed_ms: sample.rust_time_ms,
                    click_count: 0,
                },
                None,
            )
            .unwrap();
        let mutated_text = m21_visible_page_text(&mutated_frame.page);
        let mutated_layout = m21_timeline_layout(&mutated_frame, &movie_text_mutated_static.fonts);
        let mutated_png = m21_mask_audio_poster(
            &oxml_pdf::render_page_to_png(&mutated_layout, 0, F214_POWERPOINT_COMPARISON_DPI)
                .unwrap(),
        );
        let mutated_comparison = m21_compare_full_page_ink(&mutated_png, &powerpoint_frame);
        let mutated_text_matches = m21_tokens_match(&mutated_text, sample.observed_visible_text)
            && mutated_comparison.expected.len() == sample.observed_ink_band_count;
        assert!(!mutated_text_matches);
        assert!(!m21_visual_gate(mutated_text_matches, &mutated_comparison));
        movie_scores.push((sample.movie_time, comparison));
    }

    let notes_bytes = fs::read(&notes_pdf).unwrap();
    let handout_bytes = fs::read(&handout_pdf).unwrap();
    assert_eq!(f226_pdf_page_count(&notes_bytes, "m21-corrected-notes"), 3);
    assert_eq!(
        f226_pdf_page_count(&handout_bytes, "m21-corrected-handout"),
        1
    );
    m21_assert_page_size(m21_pdf_page_size_path(&notes_pdf), (595.276, 841.89));
    m21_assert_page_size(m21_pdf_page_size_path(&handout_pdf), (595.276, 841.89));
    let powerpoint_note_text_pages = m21_pdf_text_pages(&notes_bytes, "m21-corrected-notes-text");
    assert_eq!(powerpoint_note_text_pages.len(), 3);
    for (page, expected) in powerpoint_note_text_pages
        .iter()
        .zip(M21_POWERPOINT_NOTES_TOKENS)
    {
        m21_assert_text_mutations_fail(page, expected);
    }
    let powerpoint_handout_text_pages =
        m21_pdf_text_pages(&handout_bytes, "m21-corrected-handout-text");
    assert_eq!(powerpoint_handout_text_pages.len(), 1);
    m21_assert_text_mutations_fail(
        &powerpoint_handout_text_pages[0],
        M21_POWERPOINT_HANDOUT_TOKENS,
    );

    let rust_notes = presentation.to_notes_pdf_deterministic().unwrap();
    let rust_handout = presentation
        .to_handout_pdf_deterministic(HandoutLayout::Three)
        .unwrap();
    let rust_note_tokens = m21_pdf_token_pages(&rust_notes, "m21-corrected-rust-notes");
    assert_eq!(rust_note_tokens, M21_RUST_NOTES_TOKENS);
    assert_eq!(
        m21_pdf_token_pages(&rust_handout, "m21-corrected-rust-handout")[0],
        M21_RUST_HANDOUT_TOKENS
    );
    m21_assert_page_size(
        m21_pdf_page_size(&rust_notes, "m21-rust-notes-size"),
        (540.0, 720.0),
    );
    m21_assert_page_size(
        m21_pdf_page_size(&rust_handout, "m21-rust-handout-size"),
        (540.0, 720.0),
    );
    let rust_note_pngs = presentation.notes_page_pngs_deterministic(72.0).unwrap();
    let rust_handout_pngs = presentation
        .handout_page_pngs_deterministic(HandoutLayout::Three, 72.0)
        .unwrap();
    let powerpoint_note_pngs = m21_pdf_page_pngs(&notes_pdf, &root.join("m21-corrected-notes"), 72);
    let powerpoint_handout_pngs =
        m21_pdf_page_pngs(&handout_pdf, &root.join("m21-corrected-handout"), 72);
    assert_eq!(rust_note_pngs.len(), 3);
    assert_eq!(powerpoint_note_pngs.len(), 3);
    assert_eq!(rust_handout_pngs.len(), 1);
    assert_eq!(powerpoint_handout_pngs.len(), 1);
    let mut notes_scores = Vec::new();
    for (page_index, ((rust_component_index, expected_rust_bands), powerpoint_page)) in
        [(3, 5), (2, 4), (1, 3)]
            .into_iter()
            .zip(&powerpoint_note_pngs)
            .enumerate()
    {
        let rust_page = &rust_note_pngs[page_index];
        let text_matches = rust_note_tokens[page_index] == M21_RUST_NOTES_TOKENS[page_index]
            && m21_normalized_tokens(&powerpoint_note_text_pages[page_index])
                == M21_POWERPOINT_NOTES_TOKENS[page_index];
        let comparison = m21_compare_notes_ink(rust_page, powerpoint_page, rust_component_index);
        assert!(
            m21_notes_ink_gate(text_matches, &comparison, expected_rust_bands),
            "corrected notes page {} semantic-component gate: {comparison:?}",
            page_index + 1
        );

        let rust_component = m21_page_ink_bounds(rust_page, true)[rust_component_index];
        let rust_geometry_mutation = m21_extend_ink_region(rust_page, rust_component, 80);
        let rust_geometry_comparison = m21_compare_notes_ink(
            &rust_geometry_mutation,
            powerpoint_page,
            rust_component_index,
        );
        assert!(!m21_notes_ink_gate(
            text_matches,
            &rust_geometry_comparison,
            expected_rust_bands,
        ));
        let powerpoint_component = m21_page_ink_bounds(powerpoint_page, true)[0];
        let powerpoint_geometry_mutation =
            m21_extend_ink_region(powerpoint_page, powerpoint_component, 80);
        let powerpoint_geometry_comparison = m21_compare_notes_ink(
            rust_page,
            &powerpoint_geometry_mutation,
            rust_component_index,
        );
        assert!(!m21_notes_ink_gate(
            text_matches,
            &powerpoint_geometry_comparison,
            expected_rust_bands,
        ));

        let rust_paint_mutation = m21_solidify_ink_region(rust_page, rust_component);
        let rust_paint_comparison =
            m21_compare_notes_ink(&rust_paint_mutation, powerpoint_page, rust_component_index);
        assert!(!m21_notes_ink_gate(
            text_matches,
            &rust_paint_comparison,
            expected_rust_bands,
        ));
        let powerpoint_paint_mutation =
            m21_solidify_ink_region(powerpoint_page, powerpoint_component);
        let powerpoint_paint_comparison =
            m21_compare_notes_ink(rust_page, &powerpoint_paint_mutation, rust_component_index);
        assert!(!m21_notes_ink_gate(
            text_matches,
            &powerpoint_paint_comparison,
            expected_rust_bands,
        ));
        notes_scores.push((page_index + 1, comparison));
    }
    let rust_thumbnails = m21_handout_thumbnail_bounds(&rust_handout_pngs[0]);
    let powerpoint_thumbnails = m21_handout_thumbnail_bounds(&powerpoint_handout_pngs[0]);
    let handout_geometry_error =
        m21_normalized_geometry_error(&rust_thumbnails, &powerpoint_thumbnails);
    assert!(
        handout_geometry_error <= 0.05,
        "corrected normalized handout geometry differs by {handout_geometry_error}: rust={rust_thumbnails:?} PowerPoint={powerpoint_thumbnails:?}"
    );
    let mut handout_geometry_mutation = powerpoint_thumbnails.clone();
    handout_geometry_mutation[0].top += 0.051;
    assert!(m21_normalized_geometry_error(&rust_thumbnails, &handout_geometry_mutation) > 0.05);
    eprintln!(
        "M21 corrected signed oracle: source_sha256={} static_sha256={} movie_sha256={} notes_sha256={} handout_sha256={} static_scores={static_scores:?} movie_scores={movie_scores:?} notes_scores={notes_scores:?} handout_geometry_error={handout_geometry_error}",
        sha256(&source),
        sha256(&static_pdf),
        sha256(&movie),
        sha256(&notes_pdf),
        sha256(&handout_pdf),
    );
    fs::remove_dir_all(root).unwrap();
}

#[cfg(all(feature = "digital-signatures", feature = "render"))]
#[test]
#[ignore = "requires RPPTX_M21_LEGACY_STATIC_PDF"]
fn m21_recorded_minimal_smartart_oracle_classifies_blank_powerpoint_page() {
    let unsigned_source_bytes = m21_minimal_smartart_fallback_deck_bytes();
    assert_eq!(
        sha256_bytes(&unsigned_source_bytes),
        M21_UNSIGNED_SOURCE_SHA256
    );
    let static_pdf = PathBuf::from(
        std::env::var_os("RPPTX_M21_LEGACY_STATIC_PDF")
            .expect("set RPPTX_M21_LEGACY_STATIC_PDF to the recorded minimal-source static PDF"),
    );
    assert_eq!(sha256(&static_pdf), M21_LEGACY_STATIC_PDF_SHA256);
    let presentation = Presentation::from_bytes(&unsigned_source_bytes).unwrap();
    let (_, legacy_layout) = presentation.render_deterministic().unwrap();
    assert_eq!(
        m21_normalized_tokens(&m21_visible_page_text(&legacy_layout.pages[2])),
        ["Unsupported", "SmartArt"]
    );
    let rust_page = oxml_pdf::render_page_to_png(&legacy_layout, 2, 150.0).unwrap();
    let root = std::env::temp_dir().join(format!(
        "rpptx-m21-minimal-smartart-classification-{}",
        std::process::id()
    ));
    fs::create_dir_all(&root).unwrap();
    let powerpoint_pages = m21_pdf_page_pngs(&static_pdf, &root.join("m21-legacy-static"), 150);
    assert_eq!(powerpoint_pages.len(), 3);
    assert_eq!(m21_full_page_ink_bounds(&powerpoint_pages[2]), []);
    assert_eq!(m21_full_page_ink_bounds(&rust_page).len(), 1);
    assert_ne!(rust_page, powerpoint_pages[2]);
    eprintln!(
        "M21 legacy minimal SmartArt classification: Rust renders the unsupported fallback and PowerPoint renders a blank third page"
    );
    fs::remove_dir_all(root).unwrap();
}
#[cfg(feature = "agile-encryption")]
#[test]
fn encrypted_presentation_round_trips_with_password_and_preserves_package() {
    let source = fixture_bytes();
    let presentation = Presentation::from_bytes(&source).expect("open source presentation");
    let encrypted = presentation
        .to_encrypted_bytes("correct horse battery staple")
        .expect("encrypt presentation");
    assert!(Presentation::from_encrypted_bytes(&encrypted, "wrong password").is_err());
    assert!(
        Presentation::from_encrypted_bytes_with_limits(
            &encrypted,
            "correct horse battery staple",
            rpptx::PackageReadLimits {
                max_entries: 1,
                max_part_uncompressed_bytes: u64::MAX,
                max_total_uncompressed_bytes: u64::MAX,
            },
        )
        .is_err()
    );
    let reopened = Presentation::from_encrypted_bytes(&encrypted, "correct horse battery staple")
        .expect("decrypt presentation");
    assert_eq!(
        reopened.to_bytes().unwrap(),
        presentation.to_bytes().unwrap()
    );

    let path = f116_temp_path("f221-encrypted-round-trip", "pptx");
    presentation
        .save_encrypted(&path, "correct horse battery staple")
        .expect("save encrypted presentation");
    let reopened_from_path =
        Presentation::open_encrypted(&path, "correct horse battery staple").unwrap();
    assert_eq!(
        reopened_from_path.to_bytes().unwrap(),
        presentation.to_bytes().unwrap()
    );
    fs::remove_file(path).unwrap();
}

#[cfg(feature = "agile-encryption")]
#[test]
#[ignore = "requires pinned Microsoft PowerPoint and recorded password observations"]
fn powerpoint_opens_the_written_agile_presentation() {
    assert_powerpoint_build();
    let path = Path::new(F221_CANDIDATE_PATH);
    if !path.exists() {
        Presentation::from_bytes(&fixture_bytes())
            .unwrap()
            .save_encrypted(path, F221_POWERPOINT_PASSWORD)
            .unwrap();
    }
    let artifact_sha256 = sha256(path);
    eprintln!("F-221 PowerPoint candidate: {F221_CANDIDATE_PATH}, SHA-256 {artifact_sha256}");
    assert_eq!(artifact_sha256, F221_ARTIFACT_SHA256);
    assert_eq!(
        std::env::var("RPPTX_F221_POWERPOINT_OBSERVATIONS").as_deref(),
        Ok("correct-password-opened;wrong-password-rejected"),
        "record both pinned PowerPoint password observations for {artifact_sha256}"
    );
}

#[cfg(feature = "agile-encryption")]
#[test]
fn failed_encrypted_save_leaves_destination_and_presentation_unchanged() {
    let presentation = Presentation::from_bytes(&fixture_bytes()).unwrap();
    let before = presentation.to_bytes().unwrap();

    let password_failure = f116_temp_path("f221-empty-password", "pptx");
    fs::write(&password_failure, b"sentinel destination").unwrap();
    assert!(presentation.save_encrypted(&password_failure, "").is_err());
    assert_eq!(
        fs::read(&password_failure).unwrap(),
        b"sentinel destination"
    );
    assert_eq!(presentation.to_bytes().unwrap(), before);
    fs::remove_file(password_failure).unwrap();

    let destination = f116_temp_path("f221-encrypted-directory", "pptx");
    fs::create_dir(&destination).unwrap();
    assert!(
        presentation
            .save_encrypted(&destination, "password")
            .is_err()
    );
    assert!(destination.is_dir());
    assert_eq!(presentation.to_bytes().unwrap(), before);
    fs::remove_dir(destination).unwrap();
}

#[cfg(feature = "digital-signatures")]
#[test]
fn signed_presentation_reopens_with_complete_fixture_coverage() {
    let (private_key, certificate) = f221_signing_material();
    let mut presentation = Presentation::from_bytes(&fixture_bytes()).unwrap();
    let before_failed_sign = presentation.to_bytes().unwrap();
    assert!(presentation.sign(b"not-pkcs8", &certificate).is_err());
    assert_eq!(presentation.to_bytes().unwrap(), before_failed_sign);
    let report = presentation.sign(&private_key, &certificate).unwrap();
    assert!(report.cryptographically_valid);
    assert!(report.coverage_complete);
    let bytes = presentation.to_bytes().unwrap();
    let reopened = Presentation::from_bytes(&bytes).unwrap();
    let reports = reopened.verify_signatures().unwrap();
    assert_eq!(reports.len(), 1);
    assert!(reports[0].cryptographically_valid);
    assert!(reports[0].coverage_complete);
}

#[cfg(feature = "digital-signatures")]
#[test]
fn untouched_producer_shaped_presentation_keeps_its_signature_valid() {
    let (private_key, certificate) = f221_signing_material();
    let mut package = open_opc(&fixture_bytes(), "producer-shaped signed presentation");
    let presentation_xml = String::from_utf8(package.get_part(PRESENTATION_PART).unwrap().to_vec())
        .unwrap()
        .replace("xmlns:p=", "xmlns:producer=")
        .replace("<p:", "<producer:")
        .replace("</p:", "</producer:");
    package.set_part(PRESENTATION_PART, presentation_xml.into_bytes());
    let report = package.sign(&private_key, &certificate).unwrap();
    assert!(report.cryptographically_valid);
    assert!(report.coverage_complete);
    let mut signed_bytes = Cursor::new(Vec::new());
    package.write_to(&mut signed_bytes).unwrap();

    let presentation = Presentation::from_bytes(&signed_bytes.into_inner()).unwrap();
    let reports = presentation.verify_signatures().unwrap();
    assert_eq!(reports.len(), 1);
    assert!(reports[0].cryptographically_valid);
    assert!(reports[0].coverage_complete);
}

#[cfg(feature = "digital-signatures")]
#[test]
fn mutating_a_signed_presentation_never_reports_the_stale_signature_as_valid() {
    let (private_key, certificate) = f221_signing_material();

    let signed = || {
        let mut presentation = Presentation::from_bytes(&fixture_bytes()).unwrap();
        presentation.sign(&private_key, &certificate).unwrap();
        presentation
    };
    let assert_retained_but_invalid = |presentation: &Presentation| {
        let reports = presentation.verify_signatures().unwrap();
        assert_eq!(reports.len(), 1, "retained signature remains inspectable");
        assert!(!reports[0].cryptographically_valid);
    };

    let mut slide_changed = signed();
    slide_changed.slide_mut(0).unwrap().set_hidden(true);
    assert_retained_but_invalid(&slide_changed);

    let mut shape_changed = signed();
    shape_changed
        .slide_mut(0)
        .unwrap()
        .shape_mut(0)
        .unwrap()
        .set_name("changed after signing")
        .unwrap();
    assert_retained_but_invalid(&shape_changed);

    let mut text_changed = signed();
    text_changed
        .slide_mut(0)
        .unwrap()
        .shape_mut(0)
        .unwrap()
        .text_frame()
        .unwrap()
        .paragraph_mut(0)
        .unwrap()
        .run_mut(0)
        .unwrap()
        .set_text("changed after signing");
    assert_retained_but_invalid(&text_changed);

    let mut property_changed = signed();
    property_changed.core_properties_mut().title = Some("changed after signing".to_owned());
    assert_retained_but_invalid(&property_changed);

    let mut graph_changed = signed();
    graph_changed.duplicate_slide(0).unwrap();
    assert_retained_but_invalid(&graph_changed);
}

#[cfg(feature = "digital-signatures")]
#[test]
fn collaboration_mutations_never_report_the_retained_signature_as_valid() {
    use rpptx::{Comment, CommentAuthor, CommentReply, Section};

    let (private_key, certificate) = f221_signing_material();
    let fresh = || open_package(collaboration_fixture_package()).unwrap();
    let signed = |mut presentation: Presentation| {
        let report = presentation.sign(&private_key, &certificate).unwrap();
        assert!(report.cryptographically_valid);
        assert!(report.coverage_complete);
        presentation
    };
    let assert_retained_but_invalid = |label: &str, presentation: &Presentation| {
        let reports = presentation.verify_signatures().unwrap();
        assert_eq!(reports.len(), 1, "{label}: retained signature report");
        assert!(
            reports.iter().all(|report| !report.cryptographically_valid),
            "{label}: no retained signature remains valid"
        );
    };
    let author = || {
        CommentAuthor::new(
            "{11111111-1111-1111-1111-111111111111}",
            "Ada",
            Some("A"),
            "ada@example.test",
            "test",
        )
        .unwrap()
    };
    let comment = || {
        Comment::new(
            "{22222222-2222-2222-2222-222222222222}",
            "{11111111-1111-1111-1111-111111111111}",
            "2026-08-29T10:11:12Z",
            "first",
        )
        .unwrap()
    };
    let reply = || {
        CommentReply::new(
            "{33333333-3333-3333-3333-333333333333}",
            "{11111111-1111-1111-1111-111111111111}",
            "2026-08-29T10:12:13Z",
            "reply",
        )
        .unwrap()
    };

    let mut author_changed = signed(fresh());
    author_changed.add_comment_author(author()).unwrap();
    assert_retained_but_invalid("comment author", &author_changed);

    let mut comment_changed = fresh();
    comment_changed.add_comment_author(author()).unwrap();
    let mut comment_changed = signed(comment_changed);
    comment_changed.add_comment(0, comment()).unwrap();
    assert_retained_but_invalid("comment", &comment_changed);

    let mut reply_changed = fresh();
    reply_changed.add_comment_author(author()).unwrap();
    reply_changed.add_comment(0, comment()).unwrap();
    let mut reply_changed = signed(reply_changed);
    reply_changed
        .reply_to_comment(0, "{22222222-2222-2222-2222-222222222222}", reply())
        .unwrap();
    assert_retained_but_invalid("comment reply", &reply_changed);

    let mut sections_changed = signed(fresh());
    sections_changed
        .set_sections(vec![
            Section::new(
                "{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}",
                "Opening",
                vec![900, 256],
            )
            .unwrap(),
        ])
        .unwrap();
    assert_retained_but_invalid("sections", &sections_changed);

    let mut notes_master_changed = signed(fresh());
    notes_master_changed
        .notes_header_footer_mut()
        .unwrap()
        .footer = Some(false);
    assert_retained_but_invalid("notes master header and footer", &notes_master_changed);

    let mut handout_master_changed = signed(fresh());
    handout_master_changed
        .handout_header_footer_mut()
        .unwrap()
        .header = Some(false);
    assert_retained_but_invalid("handout master header and footer", &handout_master_changed);
}

#[test]
fn presentation_security_features_are_default_off_and_binding_manifests_do_not_enable_them() {
    let manifest = include_str!("../Cargo.toml");
    assert!(manifest.contains("agile-encryption = [\"oxml-opc/agile-encryption\"]"));
    assert!(manifest.contains("digital-signatures = [\"oxml-opc/digital-signatures\"]"));
    let default = manifest
        .lines()
        .find(|line| line.starts_with("default ="))
        .expect("rpptx default feature declaration");
    assert!(!default.contains("agile-encryption"));
    assert!(!default.contains("digital-signatures"));
    for binding in [
        include_str!("../../rpptx-py/Cargo.toml"),
        include_str!("../../rpptx-wasm/Cargo.toml"),
        include_str!("../../rpptx-cli/Cargo.toml"),
    ] {
        assert!(!binding.contains("agile-encryption"));
        assert!(!binding.contains("digital-signatures"));
    }
}

#[test]
fn ordinary_save_still_canonicalizes_untouched_modelled_notes_parts() {
    let mut package = fixture_package();
    let source_notes = String::from_utf8(notes_xml()).unwrap();
    let producer_notes = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n{}",
        source_notes.replace(
            &format!(r#"<p:notes xmlns:p="{P_NS}" xmlns:a="{A_NS}">"#),
            &format!(r#"<p:notes xmlns:a="{A_NS}" xmlns:p="{P_NS}">"#),
        )
    )
    .into_bytes();
    let canonical_notes = CT_NotesSlide::from_xml(&producer_notes)
        .unwrap()
        .to_xml()
        .unwrap();
    assert_ne!(producer_notes, canonical_notes);
    package.set_part(NOTES_PART, producer_notes);

    let saved = Presentation::from_bytes(&package_bytes(package))
        .unwrap()
        .to_bytes()
        .unwrap();
    let saved_package = open_opc(&saved, "ordinary canonical notes save");
    assert_eq!(
        saved_package.get_part(NOTES_PART),
        Some(canonical_notes.as_slice())
    );
}

fn f124_chart_data() -> ChartData {
    ChartData {
        categories: vec!["North".to_owned(), "South".to_owned(), "West".to_owned()],
        series: vec![
            ("Revenue".to_owned(), vec![12.5, 19.0, 14.25]),
            ("Cost".to_owned(), vec![8.0, 11.5, 9.75]),
        ],
        number_format: Some("0.00".to_owned()),
    }
}

#[test]
fn python_binding_facade_accessors_are_total_and_immutable() {
    let mut presentation = Presentation::new().expect("create presentation");
    presentation.add_slide(6).expect("add blank slide");
    presentation
        .slide_mut(0)
        .expect("new slide")
        .add_textbox(Emu(914_400), Emu(914_400), Emu(2 * 914_400), Emu(914_400))
        .expect("add textbox")
        .set_text("hello")
        .expect("set text");

    let slide = presentation.slide(0).expect("slide by index");
    assert!(slide.title().is_none());
    assert!(slide.placeholder(0).is_none());
    let shape = slide.shape(0).expect("shape by index");
    let frame = shape.text_frame().expect("immutable text frame");
    assert_eq!(frame.paragraph_count(), 1);
    let paragraph = frame.paragraph(0).expect("paragraph by index");
    assert_eq!(paragraph.text(), "hello");
    assert_eq!(paragraph.run_count(), 1);
    assert_eq!(paragraph.run(0).expect("run by index").text(), "hello");
    assert!(frame.paragraph(1).is_none());
}

#[test]
fn facade_replacement_counts_same_run_cross_run_group_and_table_matches() {
    let mut presentation = Presentation::from_bytes(&fixture_bytes()).unwrap();
    {
        let mut slide = presentation.slide_mut(0).unwrap();
        let mut shape = slide.shape_mut(0).unwrap();
        shape.set_text("same target").unwrap();
        let mut frame = shape.text_frame().unwrap();
        let mut paragraph = frame.paragraph_mut(0).unwrap();
        paragraph.add_run(" and cross ");
        paragraph.add_run("target");

        let mut table = slide.shape_mut(2).unwrap().into_table_mut().unwrap();
        table.cell_mut(0, 0).unwrap().set_text("table target");

        let mut group = slide.shape_mut(3).unwrap();
        group
            .child_mut(0)
            .unwrap()
            .set_text("nested target")
            .unwrap();
    }

    assert_eq!(presentation.replace_text("target", "done"), 4);
    let slide = presentation.slide(0).unwrap();
    assert_eq!(
        slide.shape(0).unwrap().text().unwrap(),
        "same done and cross done"
    );
    assert_eq!(
        slide
            .shape(2)
            .unwrap()
            .table()
            .unwrap()
            .cell(0, 0)
            .unwrap()
            .text(),
        "table done"
    );
    assert_eq!(
        slide.shape(3).unwrap().child(0).unwrap().text().unwrap(),
        "nested done"
    );
}

fn presentation_with_authored_chart() -> Presentation {
    let mut presentation = Presentation::new().expect("open bundled template");
    presentation.add_slide(0).expect("add chart slide");
    presentation
        .add_chart(
            0,
            ChartKind::Bar,
            Emu(914_400),
            Emu(914_400),
            Emu(7_315_200),
            Emu(4_572_000),
            &f124_chart_data(),
        )
        .expect("add authored chart");
    presentation
}

fn override_color_map(accent1: &str) -> String {
    format!(
        r#"<a:overrideClrMapping bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="{accent1}" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>"#
    )
}

fn authored_chart_with_color_maps(slide_mapping: Option<&str>, layout_mapping: &str) -> Vec<u8> {
    let bytes = presentation_with_authored_chart().to_bytes().unwrap();
    let mut package = open_opc(&bytes, "authored chart color maps");
    let slide_part = package
        .content_types
        .overrides
        .iter()
        .find_map(|(part, content_type)| {
            (content_type == content_types::SLIDE).then_some(part.clone())
        })
        .expect("slide content type override");
    let layout_part = {
        let relationship = package
            .get_part_rels(&slide_part)
            .and_then(|relationships| relationships.get_by_type(rel_types::SLIDE_LAYOUT))
            .expect("slide layout relationship");
        OpcPackage::resolve_rel_target(&slide_part, &relationship.target)
    };
    let master_mapping = "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>";

    let slide_xml = String::from_utf8(package.get_part(&slide_part).unwrap().to_vec()).unwrap();
    assert!(slide_xml.contains(master_mapping));
    let slide_xml = match slide_mapping {
        Some(mapping) => slide_xml.replace(
            master_mapping,
            &format!("<p:clrMapOvr>{mapping}</p:clrMapOvr>"),
        ),
        None => slide_xml.replace(master_mapping, ""),
    };
    package.set_part(&slide_part, slide_xml.into_bytes());

    let layout_xml = String::from_utf8(package.get_part(&layout_part).unwrap().to_vec()).unwrap();
    assert!(layout_xml.contains(master_mapping));
    let layout_xml = layout_xml.replace(
        master_mapping,
        &format!("<p:clrMapOvr>{layout_mapping}</p:clrMapOvr>"),
    );
    package.set_part(&layout_part, layout_xml.into_bytes());

    write_package(&package, "authored chart color maps", "write color maps")
}

#[test]
fn add_chart_writes_complete_relationship_graph() {
    let bytes = presentation_with_authored_chart()
        .to_bytes()
        .expect("serialize authored chart");
    let package = open_opc(&bytes, "authored chart graph");
    let chart_part = package
        .content_types
        .overrides
        .iter()
        .find_map(|(part, content_type)| {
            (content_type == content_types::CHART).then_some(part.as_str())
        })
        .expect("chart content type override");
    let workbook_part = package
        .content_types
        .overrides
        .iter()
        .find_map(|(part, content_type)| {
            (content_type == content_types::EMBEDDED_WORKBOOK).then_some(part.as_str())
        })
        .expect("workbook content type override");
    let slide_part = package
        .content_types
        .overrides
        .iter()
        .find_map(|(part, content_type)| {
            (content_type == content_types::SLIDE).then_some(part.as_str())
        })
        .expect("slide content type override");
    let chart_relationship = package
        .get_part_rels(slide_part)
        .expect("slide relationships")
        .get_by_type(rel_types::CHART)
        .expect("slide to chart relationship");
    assert_eq!(
        OpcPackage::resolve_rel_target(slide_part, &chart_relationship.target),
        chart_part
    );
    let workbook_relationship = package
        .get_part_rels(chart_part)
        .expect("chart relationships")
        .get_by_type(rel_types::PACKAGE)
        .expect("chart to workbook relationship");
    assert_eq!(
        OpcPackage::resolve_rel_target(chart_part, &workbook_relationship.target),
        workbook_part
    );
    let slide_xml = String::from_utf8(package.get_part(slide_part).unwrap().to_vec()).unwrap();
    assert!(slide_xml.contains(&format!(r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="{R_NS}" r:id="{}"/>"#, chart_relationship.id)));
    let chart_xml = String::from_utf8(package.get_part(chart_part).unwrap().to_vec()).unwrap();
    assert!(chart_xml.contains(&format!(
        r#"<c:externalData r:id="{}">"#,
        workbook_relationship.id
    )));
    assert!(package.get_part(workbook_part).is_some());
    assert!(
        Presentation::from_bytes(&bytes)
            .unwrap()
            .validate()
            .is_empty()
    );
}

#[test]
fn authored_chart_honors_slide_layout_and_master_color_maps() {
    let baseline = deterministic_render(&presentation_with_authored_chart().to_bytes().unwrap());
    let layout_mapping = override_color_map("accent6");
    let slide_mapping = override_color_map("accent5");

    let master = deterministic_render(&authored_chart_with_color_maps(
        Some("<a:masterClrMapping/>"),
        &layout_mapping,
    ));
    let layout = deterministic_render(&authored_chart_with_color_maps(None, &layout_mapping));
    let slide = deterministic_render(&authored_chart_with_color_maps(
        Some(&slide_mapping),
        &layout_mapping,
    ));

    assert_eq!(master, baseline);
    assert_ne!(layout, baseline);
    assert_ne!(slide, baseline);
    assert_ne!(slide, layout);
}

#[test]
fn add_chart_uses_collision_free_part_numbers() {
    let mut source = Presentation::new().unwrap();
    source.add_slide(0).unwrap();
    let mut package = open_opc(&source.to_bytes().unwrap(), "sparse chart fixture");
    package.set_part("/ppt/charts/chart2.xml", b"occupied chart".to_vec());
    package.set_part(
        "/ppt/embeddings/Workbook7.xlsx",
        b"occupied workbook".to_vec(),
    );
    package
        .content_types
        .add_override("/ppt/charts/chart2.xml", content_types::CHART);
    package.content_types.add_override(
        "/ppt/embeddings/Workbook7.xlsx",
        content_types::EMBEDDED_WORKBOOK,
    );
    let mut bytes = Cursor::new(Vec::new());
    package.write_to(&mut bytes).unwrap();
    let mut presentation = Presentation::from_bytes(&bytes.into_inner()).unwrap();
    presentation
        .add_chart(
            0,
            ChartKind::Line,
            Emu(1),
            Emu(2),
            Emu(3),
            Emu(4),
            &f124_chart_data(),
        )
        .unwrap();
    let package = open_opc(&presentation.to_bytes().unwrap(), "numbered chart");
    assert!(package.parts.contains_key("/ppt/charts/chart3.xml"));
    assert!(package.parts.contains_key("/ppt/embeddings/Workbook8.xlsx"));
    assert_eq!(
        package.get_part("/ppt/charts/chart2.xml"),
        Some(b"occupied chart".as_slice())
    );
}

#[test]
fn add_chart_caches_and_workbook_share_one_source() {
    let bytes = presentation_with_authored_chart().to_bytes().unwrap();
    let package = open_opc(&bytes, "authored chart cache");
    let chart_part = package
        .content_types
        .overrides
        .iter()
        .find_map(|(part, content_type)| {
            (content_type == content_types::CHART).then_some(part.as_str())
        })
        .unwrap();
    let chart = CT_ChartSpace::from_xml(package.get_part(chart_part).unwrap()).unwrap();
    let series = chart.chart.plot_area.series().unwrap();
    assert_eq!(series.len(), 2);
    let expected_series = [
        (
            "Sheet1!$B$1",
            "Revenue",
            "Sheet1!$B$2:$B$4",
            [12.5, 19.0, 14.25],
        ),
        ("Sheet1!$C$1", "Cost", "Sheet1!$C$2:$C$4", [8.0, 11.5, 9.75]),
    ];
    for (actual, (name_formula, name, value_formula, values)) in series.iter().zip(expected_series)
    {
        let actual_name = actual.name.as_ref().unwrap();
        assert_eq!(actual_name.formula, name_formula);
        assert_eq!(actual_name.values, [name]);
        let AxisData::String(categories) = actual.categories.as_ref().unwrap() else {
            panic!("category series must use a string cache");
        };
        assert_eq!(categories.formula, "Sheet1!$A$2:$A$4");
        assert_eq!(categories.values, ["North", "South", "West"]);
        assert_eq!(actual.values.formula, value_formula);
        assert_eq!(actual.values.format_code, "0.00");
        assert_eq!(actual.values.values, values);
    }
    let workbook_relationship = package
        .get_part_rels(chart_part)
        .unwrap()
        .get_by_type(rel_types::PACKAGE)
        .unwrap();
    let workbook_part = OpcPackage::resolve_rel_target(chart_part, &workbook_relationship.target);
    let workbook =
        OpcPackage::from_reader(Cursor::new(package.get_part(&workbook_part).unwrap())).unwrap();
    let worksheet = String::from_utf8(
        workbook
            .get_part("/xl/worksheets/sheet1.xml")
            .unwrap()
            .to_vec(),
    )
    .unwrap();
    let shared_strings =
        String::from_utf8(workbook.get_part("/xl/sharedStrings.xml").unwrap().to_vec()).unwrap();
    let shared_strings = shared_string_values(&shared_strings);
    let expected_cells = BTreeMap::from([
        ("A1", "Category"),
        ("A2", "North"),
        ("A3", "South"),
        ("A4", "West"),
        ("B1", "Revenue"),
        ("B2", "12.5"),
        ("B3", "19"),
        ("B4", "14.25"),
        ("C1", "Cost"),
        ("C2", "8"),
        ("C3", "11.5"),
        ("C4", "9.75"),
    ]);
    let actual_cells = expected_cells
        .keys()
        .map(|address| {
            let (shared, value) = worksheet_cell(&worksheet, address);
            let value = if shared {
                shared_strings[value.parse::<usize>().unwrap()].as_str()
            } else {
                value
            };
            (*address, value)
        })
        .collect::<BTreeMap<_, _>>();
    assert_eq!(actual_cells, expected_cells);
}

#[test]
fn add_chart_rejects_invalid_data_without_mutation() {
    let invalid = [
        ChartData {
            categories: Vec::new(),
            series: vec![("Revenue".to_owned(), Vec::new())],
            number_format: None,
        },
        ChartData {
            categories: vec!["North".to_owned()],
            series: Vec::new(),
            number_format: None,
        },
        ChartData {
            categories: vec!["North".to_owned(), "South".to_owned()],
            series: vec![("Revenue".to_owned(), vec![12.5])],
            number_format: None,
        },
        ChartData {
            categories: vec!["North".to_owned()],
            series: vec![("Revenue".to_owned(), vec![f64::NAN])],
            number_format: None,
        },
        ChartData {
            categories: vec!["North".to_owned()],
            series: vec![("Revenue".to_owned(), vec![12.5])],
            number_format: Some(String::new()),
        },
    ];
    for data in invalid {
        let mut presentation = Presentation::new().unwrap();
        presentation.add_slide(0).unwrap();
        let before = presentation.to_bytes().unwrap();
        assert!(
            presentation
                .add_chart(0, ChartKind::Bar, Emu(1), Emu(2), Emu(3), Emu(4), &data,)
                .is_err()
        );
        assert_eq!(presentation.to_bytes().unwrap(), before);
    }

    let mut presentation = Presentation::new().unwrap();
    presentation.add_slide(0).unwrap();
    let before = presentation.to_bytes().unwrap();
    assert!(
        presentation
            .add_chart(
                0,
                ChartKind::Pie,
                Emu(1),
                Emu(2),
                Emu(3),
                Emu(4),
                &f124_chart_data(),
            )
            .is_err()
    );
    assert_eq!(presentation.to_bytes().unwrap(), before);

    for (width, height) in [
        (Emu(0), Emu(1)),
        (Emu(-1), Emu(1)),
        (Emu(1), Emu(0)),
        (Emu(1), Emu(-1)),
    ] {
        assert!(
            presentation
                .add_chart(
                    0,
                    ChartKind::Bar,
                    Emu(1),
                    Emu(2),
                    width,
                    height,
                    &f124_chart_data(),
                )
                .is_err()
        );
        assert_eq!(presentation.to_bytes().unwrap(), before);
    }

    let scatter = ChartData {
        categories: vec!["not numeric".to_owned()],
        series: vec![("Revenue".to_owned(), vec![12.5])],
        number_format: None,
    };
    assert!(
        presentation
            .add_chart(
                0,
                ChartKind::Scatter,
                Emu(1),
                Emu(2),
                Emu(3),
                Emu(4),
                &scatter,
            )
            .is_err()
    );
    assert_eq!(presentation.to_bytes().unwrap(), before);
}

#[test]
fn add_chart_authors_each_supported_family() {
    let one_series = ChartData {
        categories: vec!["1".to_owned(), "2".to_owned(), "3".to_owned()],
        series: vec![("Revenue".to_owned(), vec![12.5, 19.0, 14.25])],
        number_format: None,
    };
    for (kind, plot_element) in [
        (ChartKind::Bar, "c:barChart"),
        (ChartKind::Line, "c:lineChart"),
        (ChartKind::Pie, "c:pieChart"),
        (ChartKind::Doughnut, "c:doughnutChart"),
        (ChartKind::Area, "c:areaChart"),
        (ChartKind::Scatter, "c:scatterChart"),
        (ChartKind::Radar, "c:radarChart"),
    ] {
        let mut presentation = Presentation::new().unwrap();
        presentation.add_slide(0).unwrap();
        presentation
            .add_chart(0, kind, Emu(1), Emu(2), Emu(3), Emu(4), &one_series)
            .unwrap();
        let bytes = presentation.to_bytes().unwrap();
        let package = open_opc(&bytes, "supported authored chart");
        let chart = package
            .content_types
            .overrides
            .iter()
            .find_map(|(part, content_type)| {
                (content_type == content_types::CHART).then(|| package.get_part(part).unwrap())
            })
            .unwrap();
        let chart = String::from_utf8(chart.to_vec()).unwrap();
        assert!(
            chart.contains(&format!("<{plot_element}>")),
            "{kind:?}: {chart}"
        );
        assert!(
            Presentation::from_bytes(&bytes)
                .unwrap()
                .validate()
                .is_empty()
        );
    }
}

#[test]
fn authored_chart_relationship_enters_presentation_renderer() {
    let bytes = presentation_with_authored_chart()
        .to_bytes()
        .expect("serialize authored chart for rendering");
    let rendered = render_presentation_package(&bytes);

    let mut paths = 0;
    let mut text_runs = 0;
    walk(
        &rendered.pages[0].elements,
        &mut |element, _| match element {
            PositionedElement::Path(path) => {
                let path_bounds = path.path.bounds().expect("nonempty authored chart path");
                assert!(path_bounds.x.is_finite());
                assert!(path_bounds.y.is_finite());
                assert!(path_bounds.width.is_finite());
                assert!(path_bounds.height.is_finite());
                paths += 1;
            }
            PositionedElement::Text(run) => {
                assert!(run.origin.x.is_finite());
                assert!(run.origin.y.is_finite());
                assert!(run.font_size.is_finite());
                assert!(run.advances.iter().all(|advance| advance.is_finite()));
                assert!(!run.text.is_empty());
                text_runs += 1;
            }
            PositionedElement::Line {
                start, end, width, ..
            } => {
                assert!(start.x.is_finite());
                assert!(start.y.is_finite());
                assert!(end.x.is_finite());
                assert!(end.y.is_finite());
                assert!(width.is_finite());
            }
            _ => {}
        },
    );
    assert!(paths >= 6, "expected authored bars and annotation paths");
    assert!(text_runs >= 6, "expected authored axis and legend labels");
}

#[test]
fn supported_and_fallback_charts_render_deterministically() {
    let supported = presentation_with_authored_chart().to_bytes().unwrap();
    let preview = valid_one_pixel_png();
    let fallback =
        presentation_with_three_dimensional_chart_fallback(Some((&preview, "image/png")));
    let missing = presentation_with_three_dimensional_chart_fallback(None);
    let malformed =
        presentation_with_three_dimensional_chart_fallback(Some((b"not a PNG", "image/png")));
    let mut corrupt_png = valid_one_pixel_png();
    corrupt_png[41] = 0;
    let corrupt =
        presentation_with_three_dimensional_chart_fallback(Some((&corrupt_png, "image/png")));
    let mismatched =
        presentation_with_three_dimensional_chart_fallback(Some((&preview, "image/jpeg")));
    let jpeg = valid_template_jpeg();
    let jpeg_fallback =
        presentation_with_three_dimensional_chart_fallback(Some((&jpeg, "image/jpeg")));
    let sof = jpeg
        .windows(2)
        .position(|window| matches!(window, [0xff, 0xc0] | [0xff, 0xc2]))
        .unwrap();
    let sof_length = usize::from(u16::from_be_bytes([jpeg[sof + 2], jpeg[sof + 3]]));
    let corrupt_jpeg = jpeg[..sof + 2 + sof_length].to_vec();
    assert!(oxml_media::probe(&corrupt_jpeg).is_some());
    let corrupt_jpeg_fallback =
        presentation_with_three_dimensional_chart_fallback(Some((&corrupt_jpeg, "image/jpeg")));
    let mismatched_jpeg =
        presentation_with_three_dimensional_chart_fallback(Some((&jpeg, "image/png")));
    let inflation_bomb = png_inflation_bomb();
    assert!(oxml_media::probe(&inflation_bomb).is_some());
    let inflation_bomb_fallback =
        presentation_with_three_dimensional_chart_fallback(Some((&inflation_bomb, "image/png")));
    let oversized = png_header(u32::MAX, u32::MAX);
    let oversized_fallback =
        presentation_with_three_dimensional_chart_fallback(Some((&oversized, "image/png")));
    let sparse = sparse_preview_png();
    let sparse_fallback =
        presentation_with_three_dimensional_chart_fallback(Some((&sparse, "image/png")));
    let grayscale_jpeg = valid_grayscale_jpeg();
    assert_eq!(oxml_media::probe(&grayscale_jpeg).unwrap().channels, 1);
    let grayscale_jpeg_fallback =
        presentation_with_three_dimensional_chart_fallback(Some((&grayscale_jpeg, "image/jpeg")));
    let cmyk_jpeg = valid_cmyk_jpeg();
    assert_eq!(oxml_media::probe(&cmyk_jpeg).unwrap().channels, 4);
    let cmyk_jpeg_fallback =
        presentation_with_three_dimensional_chart_fallback(Some((&cmyk_jpeg, "image/jpeg")));

    for bytes in [
        &supported,
        &fallback,
        &missing,
        &malformed,
        &corrupt,
        &mismatched,
        &jpeg_fallback,
        &corrupt_jpeg_fallback,
        &mismatched_jpeg,
        &inflation_bomb_fallback,
        &oversized_fallback,
        &sparse_fallback,
        &grayscale_jpeg_fallback,
        &cmyk_jpeg_fallback,
    ] {
        let first = render_presentation_package(bytes);
        let second = render_presentation_package(bytes);
        assert_eq!(format!("{first:?}"), format!("{second:?}"));
        assert_eq!(
            oxml_pdf::render_page_to_png(&first, 0, 96.0).unwrap(),
            oxml_pdf::render_page_to_png(&second, 0, 96.0).unwrap()
        );
    }

    let rendered_fallback = render_presentation_package(&fallback);
    let mut cached_images = Vec::new();
    walk(&rendered_fallback.pages[0].elements, &mut |element, _| {
        if let PositionedElement::Image { data, .. } = element {
            cached_images.push(data.clone());
        }
    });
    assert_eq!(cached_images, vec![preview]);

    let rendered_jpeg = render_presentation_package(&jpeg_fallback);
    let mut cached_jpegs = Vec::new();
    walk(&rendered_jpeg.pages[0].elements, &mut |element, _| {
        if let PositionedElement::Image { data, .. } = element {
            cached_jpegs.push(data.clone());
        }
    });
    assert_eq!(cached_jpegs, vec![jpeg]);

    let rendered_sparse = render_presentation_package(&sparse_fallback);
    let mut cached_sparse_images = Vec::new();
    walk(&rendered_sparse.pages[0].elements, &mut |element, _| {
        if let PositionedElement::Image { data, .. } = element {
            cached_sparse_images.push(data.clone());
        }
    });
    assert_eq!(cached_sparse_images, vec![sparse]);

    for bytes in [
        &missing,
        &malformed,
        &corrupt,
        &mismatched,
        &corrupt_jpeg_fallback,
        &mismatched_jpeg,
        &inflation_bomb_fallback,
        &oversized_fallback,
        &grayscale_jpeg_fallback,
        &cmyk_jpeg_fallback,
    ] {
        let rendered = render_presentation_package(bytes);
        let mut labels = Vec::new();
        walk(&rendered.pages[0].elements, &mut |element, _| {
            if let PositionedElement::Text(run) = element {
                labels.push(run.text.clone());
            }
        });
        assert!(labels.iter().any(|label| label == "Unsupported chart"));
    }
}

#[test]
#[ignore = "manual PowerPoint open and Edit Data acceptance"]
fn created_chart_opens_and_edit_data_matches_source() {
    let bytes = presentation_with_authored_chart().to_bytes().unwrap();
    fs::write(F124_CANDIDATE_PATH, &bytes).unwrap();
    let sha = sha256(Path::new(F124_CANDIDATE_PATH));
    eprintln!("F-124 PowerPoint candidate: {F124_CANDIDATE_PATH}, SHA-256 {sha}");
    assert_eq!(sha, F124_ARTIFACT_SHA256);
    assert!(open_opc(&bytes, "F-124 candidate").parts.len() > 1);
}

fn shared_string_values(xml: &str) -> Vec<String> {
    xml.split("<t>")
        .skip(1)
        .map(|tail| tail.split_once("</t>").unwrap().0.to_owned())
        .collect()
}

fn worksheet_cell<'a>(xml: &'a str, address: &str) -> (bool, &'a str) {
    let marker = format!(r#"<c r="{address}""#);
    let cell = xml
        .split_once(&marker)
        .unwrap_or_else(|| panic!("missing worksheet cell {address}"))
        .1;
    let (attributes, cell) = cell.split_once('>').unwrap();
    let cell = cell.split_once("</c>").unwrap().0;
    let value = cell
        .split_once("<v>")
        .and_then(|(_, tail)| tail.split_once("</v>"))
        .map(|(value, _)| value)
        .unwrap_or_else(|| panic!("missing worksheet value {address}"));
    (attributes.contains(r#"t="s""#), value)
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum CrossViewerObservation {
    Clean {
        opened_or_imported: bool,
        observed_slide_count: usize,
        no_repair_or_conversion_error: bool,
        export_or_close_result: &'static str,
    },
    Pending {
        reason: &'static str,
    },
}

#[derive(Clone, Copy, Debug)]
struct CrossViewerEvidence {
    application: &'static str,
    version_or_date: Option<&'static str>,
    build: Option<&'static str>,
    input_sha256: &'static str,
    observation: CrossViewerObservation,
}

const CROSS_VIEWER_ACCEPTANCE: [CrossViewerEvidence; 4] = [
    CrossViewerEvidence {
        application: "Microsoft PowerPoint",
        version_or_date: Some(POWERPOINT_VERSION),
        build: Some("Info.plist 16.104.25121423, AppleScript 1214"),
        input_sha256: F116_ARTIFACT_SHA256,
        observation: CrossViewerObservation::Clean {
            opened_or_imported: true,
            observed_slide_count: 10,
            no_repair_or_conversion_error: true,
            export_or_close_result: "closed without saving",
        },
    },
    CrossViewerEvidence {
        application: "Apple Keynote",
        version_or_date: Some(KEYNOTE_VERSION),
        build: Some(KEYNOTE_BUILD),
        input_sha256: F116_ARTIFACT_SHA256,
        observation: CrossViewerObservation::Clean {
            opened_or_imported: true,
            observed_slide_count: 10,
            no_repair_or_conversion_error: true,
            export_or_close_result: "user confirmed clean open and closed the presentation",
        },
    },
    CrossViewerEvidence {
        application: "Google Slides",
        version_or_date: Some("accepted 2026-08-09"),
        build: Some("Google Chrome 151.0.7922.76, build 7922.76"),
        input_sha256: F116_ARTIFACT_SHA256,
        observation: CrossViewerObservation::Clean {
            opened_or_imported: true,
            observed_slide_count: 10,
            no_repair_or_conversion_error: true,
            export_or_close_result: "Microsoft PowerPoint download started once",
        },
    },
    CrossViewerEvidence {
        application: "LibreOffice Impress",
        version_or_date: Some("26.2.5.2"),
        build: Some("cd7284b4cbbfeb507e630c1aac019f4157393acb"),
        input_sha256: F116_ARTIFACT_SHA256,
        observation: CrossViewerObservation::Clean {
            opened_or_imported: true,
            observed_slide_count: 10,
            no_repair_or_conversion_error: true,
            export_or_close_result: "hidden-slide PDF export succeeded with ten pages",
        },
    },
];

#[test]
fn ten_slide_write_api_deck_validates_and_reopens() {
    let presentation = build_f116_ten_slide_deck();
    assert_f116_deck_structure(&presentation);
    let bytes = presentation.to_bytes().expect("serialize F-116 deck");
    let candidate_path = f116_temp_path("validation", "pptx");
    fs::write(&candidate_path, &bytes).expect("write temporary F-116 candidate");
    let candidate_sha = sha256(&candidate_path);
    fs::remove_file(&candidate_path).expect("remove temporary F-116 candidate");
    eprintln!("F-116 candidate SHA-256: {candidate_sha}");
    assert_eq!(candidate_sha, F116_ARTIFACT_SHA256);

    let reopened = Presentation::from_bytes(&bytes).expect("reopen F-116 deck");
    assert_eq!(reopened.len(), 10);
    assert_f116_deck_structure(&reopened);

    let package = open_opc(&bytes, "F-116 ten-slide candidate");
    let presentation_part = package.main_document_part().unwrap();
    let slide_relationships = package
        .get_part_rels(&presentation_part)
        .unwrap()
        .get_all_by_type(rel_types::SLIDE);
    assert_eq!(slide_relationships.len(), 10);
    let image_scopes = package
        .part_rels
        .values()
        .filter(|relationships| {
            relationships
                .items
                .iter()
                .any(|relationship| relationship.rel_type == rel_types::IMAGE)
        })
        .count();
    assert_eq!(image_scopes, 3);
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
fn ten_slide_write_api_deck_saves_as_presentation_and_show() {
    let presentation = build_f116_ten_slide_deck();
    assert!(presentation.validate().is_empty());
    let ordinary_path = f116_temp_path("ordinary-save", "pptx");
    presentation
        .save(&ordinary_path)
        .expect("save F-116 presentation");
    let ordinary_bytes = fs::read(&ordinary_path).expect("read F-116 presentation");
    fs::remove_file(&ordinary_path).expect("remove temporary F-116 presentation");
    let ordinary = open_opc(&ordinary_bytes, "F-116 presentation package");
    let ordinary_main = ordinary.main_document_part().unwrap();
    assert_eq!(
        ordinary.content_types.content_type_for(&ordinary_main),
        Some(content_types::PRESENTATION)
    );

    let show_path = f116_temp_path("slideshow-save", "ppsx");
    presentation
        .save_as_show(&show_path)
        .expect("save F-116 slideshow");
    let show_bytes = fs::read(&show_path).expect("read F-116 slideshow");
    fs::remove_file(&show_path).expect("remove temporary F-116 slideshow");
    let show = open_opc(&show_bytes, "F-116 slideshow package");
    let show_main = show.main_document_part().unwrap();
    assert_eq!(
        show.content_types.content_type_for(&show_main),
        Some(content_types::SLIDESHOW)
    );
    let reopened = Presentation::from_bytes(&show_bytes).expect("reopen F-116 slideshow");
    assert_f116_deck_structure(&reopened);
}

#[test]
#[ignore = "requires pinned PowerPoint, Keynote, Google Slides, and LibreOffice"]
fn generated_ten_slide_write_api_deck_opens_clean_in_all_four_viewers() {
    let path = write_f116_candidate();
    assert_eq!(sha256(&path), F116_ARTIFACT_SHA256);
    assert_f116_powerpoint_acceptance(&path);
    assert_f116_libreoffice_acceptance(&path);
    for row in CROSS_VIEWER_ACCEPTANCE {
        assert_clean_cross_viewer_evidence(row).unwrap_or_else(|error| panic!("{error}"));
    }
}

#[test]
fn cross_viewer_acceptance_evidence_is_complete_and_bound_to_one_artifact() {
    assert_eq!(CROSS_VIEWER_ACCEPTANCE.len(), 4);
    assert_eq!(
        CROSS_VIEWER_ACCEPTANCE
            .iter()
            .map(|row| row.application)
            .collect::<HashSet<_>>(),
        HashSet::from([
            "Microsoft PowerPoint",
            "Apple Keynote",
            "Google Slides",
            "LibreOffice Impress",
        ])
    );
    let candidate = write_f116_temporary_candidate("evidence");
    let actual_sha = sha256(&candidate);
    fs::remove_file(&candidate).expect("remove temporary F-116 evidence candidate");
    assert_eq!(actual_sha, F116_ARTIFACT_SHA256);
    for row in CROSS_VIEWER_ACCEPTANCE {
        assert_eq!(row.input_sha256, F116_ARTIFACT_SHA256);
        if matches!(row.observation, CrossViewerObservation::Clean { .. }) {
            assert_clean_cross_viewer_evidence(row).unwrap();
        }
    }
    let pending = CROSS_VIEWER_ACCEPTANCE
        .iter()
        .filter_map(|row| match row.observation {
            CrossViewerObservation::Pending { reason } => {
                assert!(!reason.is_empty());
                Some(row.application)
            }
            CrossViewerObservation::Clean { .. } => None,
        })
        .collect::<Vec<_>>();
    assert!(pending.is_empty(), "pending viewer evidence: {pending:?}");
}

#[test]
fn pending_cross_viewer_evidence_cannot_pass_as_clean() {
    let pending = CrossViewerEvidence {
        application: "pending regression fixture",
        version_or_date: Some("observed"),
        build: Some("observed"),
        input_sha256: F116_ARTIFACT_SHA256,
        observation: CrossViewerObservation::Pending {
            reason: "acceptance operation remains incomplete",
        },
    };
    assert!(
        assert_clean_cross_viewer_evidence(pending).is_err(),
        "pending evidence passed as clean"
    );
}

#[test]
fn clean_cross_viewer_evidence_requires_each_positive_operation() {
    let clean = CrossViewerEvidence {
        application: "regression fixture",
        version_or_date: Some("observed"),
        build: Some("observed"),
        input_sha256: F116_ARTIFACT_SHA256,
        observation: CrossViewerObservation::Clean {
            opened_or_imported: true,
            observed_slide_count: 10,
            no_repair_or_conversion_error: true,
            export_or_close_result: "closed or exported successfully",
        },
    };
    assert!(assert_clean_cross_viewer_evidence(clean).is_ok());

    for incomplete in [
        CrossViewerObservation::Clean {
            opened_or_imported: false,
            observed_slide_count: 10,
            no_repair_or_conversion_error: true,
            export_or_close_result: "closed or exported successfully",
        },
        CrossViewerObservation::Clean {
            opened_or_imported: true,
            observed_slide_count: 0,
            no_repair_or_conversion_error: true,
            export_or_close_result: "closed or exported successfully",
        },
        CrossViewerObservation::Clean {
            opened_or_imported: true,
            observed_slide_count: 10,
            no_repair_or_conversion_error: false,
            export_or_close_result: "closed or exported successfully",
        },
        CrossViewerObservation::Clean {
            opened_or_imported: true,
            observed_slide_count: 10,
            no_repair_or_conversion_error: true,
            export_or_close_result: "",
        },
        CrossViewerObservation::Clean {
            opened_or_imported: true,
            observed_slide_count: 10,
            no_repair_or_conversion_error: true,
            export_or_close_result: " \t\n",
        },
    ] {
        assert!(
            assert_clean_cross_viewer_evidence(CrossViewerEvidence {
                observation: incomplete,
                ..clean
            })
            .is_err()
        );
    }
}

#[test]
fn clean_cross_viewer_evidence_rejects_missing_or_blank_metadata() {
    let clean = CrossViewerEvidence {
        application: "regression fixture",
        version_or_date: Some("observed"),
        build: Some("observed"),
        input_sha256: F116_ARTIFACT_SHA256,
        observation: CrossViewerObservation::Clean {
            opened_or_imported: true,
            observed_slide_count: 10,
            no_repair_or_conversion_error: true,
            export_or_close_result: "closed or exported successfully",
        },
    };

    for version_or_date in [None, Some(""), Some(" \t\n")] {
        assert!(
            assert_clean_cross_viewer_evidence(CrossViewerEvidence {
                version_or_date,
                ..clean
            })
            .is_err()
        );
    }
    for build in [None, Some(""), Some(" \t\n")] {
        assert!(
            assert_clean_cross_viewer_evidence(CrossViewerEvidence { build, ..clean }).is_err()
        );
    }
}

#[test]
fn slide_and_presentation_properties_round_trip() {
    let mut presentation = Presentation::new().expect("open bundled template");
    presentation.add_slide(0).expect("add slide");
    presentation
        .set_slide_size(Emu(10_000_001), Emu(5_000_003))
        .expect("set slide size");
    let fill = Fill::from_xml(br#"<a:solidFill><a:srgbClr val="345678"/></a:solidFill>"#).unwrap();
    let mut slide = presentation.slide_mut(0).unwrap();
    slide.set_hidden(true);
    slide.set_background(fill).expect("set direct background");
    presentation.core_properties_mut().title = Some("F-115 properties".to_owned());

    let bytes = presentation.to_bytes().expect("serialize properties");
    let package = open_opc(&bytes, "property round trip");
    let slide_part = package
        .content_types
        .overrides
        .iter()
        .find_map(|(part_name, content_type)| {
            (content_type == content_types::SLIDE).then_some(part_name)
        })
        .unwrap();
    let slide = CT_Slide::from_xml(package.get_part(slide_part).unwrap()).unwrap();
    let BackgroundRendering::Properties(Some(background_fill)) = slide
        .common_slide_data
        .background
        .as_ref()
        .unwrap()
        .rendering()
    else {
        panic!("expected a direct slide background fill");
    };
    assert_eq!(
        background_fill.to_xml().unwrap(),
        br#"<a:solidFill><a:srgbClr val="345678"/></a:solidFill>"#
    );
    let reopened = Presentation::from_bytes(&bytes).expect("reopen properties");

    assert_eq!(
        reopened.slide_size(),
        Some((Emu(10_000_001), Emu(5_000_003)))
    );
    assert!(reopened.slide(0).unwrap().hidden());
    assert!(reopened.slide(0).unwrap().has_explicit_background());
    assert_eq!(
        reopened
            .core_properties()
            .and_then(|props| props.title.as_deref()),
        Some("F-115 properties")
    );
}

#[test]
fn core_properties_are_loaded_lazily_and_written_with_valid_graph() {
    const CORE_PART: &str = "/metadata/nonstandard-core.xml";
    let core_xml = br#"<?xml version="1.0"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title> Original &amp; exact </dc:title><!--producer--></cp:coreProperties>"#;
    let mut package = fixture_package();
    package.set_part(CORE_PART, core_xml.to_vec());
    package
        .content_types
        .add_override(CORE_PART, content_types::CORE_PROPERTIES);
    package.package_rels.add_with_id(
        "core-link",
        rel_types::CORE_PROPERTIES,
        "metadata/nonstandard-core.xml",
    );
    let mut presentation = Presentation::from_bytes(&package_bytes(package)).unwrap();

    assert_eq!(
        presentation
            .core_properties()
            .and_then(|props| props.title.as_deref()),
        Some(" Original & exact ")
    );
    let immutable = open_opc(
        &presentation.to_bytes().unwrap(),
        "immutable core properties",
    );
    assert_eq!(immutable.get_part(CORE_PART).unwrap(), core_xml);

    presentation.core_properties_mut().creator = Some("F-115".to_owned());
    let mutated = open_opc(&presentation.to_bytes().unwrap(), "mutated core properties");
    assert_eq!(
        mutated.content_types.content_type_for(CORE_PART),
        Some(content_types::CORE_PROPERTIES)
    );
    let relationship = mutated
        .package_rels
        .get_by_type(rel_types::CORE_PROPERTIES)
        .unwrap();
    assert_eq!(
        OpcPackage::resolve_rel_target("/", &relationship.target),
        CORE_PART
    );
    assert_eq!(
        oxml_core::core_properties::CoreProperties::from_xml(mutated.get_part(CORE_PART).unwrap())
            .unwrap()
            .creator
            .as_deref(),
        Some("F-115")
    );

    let mut absent = Presentation::from_bytes(&package_bytes(fixture_package())).unwrap();
    assert!(absent.core_properties().is_none());
    absent.core_properties_mut().subject = Some("created".to_owned());
    let created = open_opc(&absent.to_bytes().unwrap(), "created core properties");
    let relationship = created
        .package_rels
        .get_by_type(rel_types::CORE_PROPERTIES)
        .unwrap();
    let part = OpcPackage::resolve_rel_target("/", &relationship.target);
    assert_eq!(part, "/docProps/core.xml");
    assert_eq!(
        created.content_types.content_type_for(&part),
        Some(content_types::CORE_PROPERTIES)
    );
    let created_xml = created.get_part(&part).expect("created core part");
    assert_eq!(
        oxml_core::core_properties::CoreProperties::from_xml(created_xml)
            .unwrap()
            .subject
            .as_deref(),
        Some("created")
    );
}

#[test]
fn occupied_conventional_core_part_is_not_overwritten_without_relationship() {
    const OCCUPIED_CORE_PART: &str = "/docProps/core.xml";
    let occupied = b"producer-owned bytes outside the core relationship graph";
    let mut package = fixture_package();
    package.set_part(OCCUPIED_CORE_PART, occupied.to_vec());
    let source_bytes = package_bytes(package);
    let mut presentation = Presentation::from_bytes(&source_bytes).unwrap();
    presentation.core_properties_mut().subject = Some("must not overwrite".to_owned());

    let error = presentation.to_bytes().unwrap_err();
    assert!(matches!(
        error,
        Error::CorePropertiesPartCollision { ref part_name }
            if part_name == OCCUPIED_CORE_PART
    ));

    let output = std::env::temp_dir().join(format!(
        "rpptx-f115-core-collision-{}.pptx",
        std::process::id()
    ));
    fs::write(&output, b"existing output").unwrap();
    assert!(presentation.save(&output).is_err());
    assert_eq!(fs::read(&output).unwrap(), b"existing output");
    fs::remove_file(&output).unwrap();

    let source = open_opc(&source_bytes, "occupied core source");
    assert_eq!(source.get_part(OCCUPIED_CORE_PART).unwrap(), occupied);
    assert!(
        source
            .package_rels
            .get_by_type(rel_types::CORE_PROPERTIES)
            .is_none()
    );
}

#[test]
fn modern_presentation_package_variants_reopen_with_original_class_and_payloads() {
    for class in f223_variant_classes() {
        let source = f223_fixture_package(class);
        let source_parts = source.parts.clone();
        let source_relationships = source.part_rels.clone();
        let opened = Presentation::from_bytes(&package_bytes(source)).unwrap();
        assert_eq!(opened.package_class().unwrap(), class);

        let saved = opened.to_bytes().unwrap();
        let package = open_opc(&saved, "F-223 saved variant");
        let reopened = Presentation::from_bytes(&saved).unwrap();
        assert_eq!(reopened.package_class().unwrap(), class);
        for (part, bytes) in &source_parts {
            if part.ends_with(".bin") {
                assert_eq!(
                    package.get_part(part),
                    Some(bytes.as_slice()),
                    "{class:?}: {part}"
                );
            }
        }
        assert_f223_relationships_equal(&package.part_rels, &source_relationships, class);
        if matches!(
            class,
            PresentationPackageClass::MacroEnabledPresentation
                | PresentationPackageClass::MacroEnabledTemplate
                | PresentationPackageClass::MacroEnabledSlideshow
        ) {
            assert_eq!(
                reopened
                    .extract_embedded_content(PRESENTATION_PART, "vba-rel")
                    .unwrap(),
                b"vba-executable"
            );
        }
    }
}

#[test]
fn package_class_conversion_changes_only_the_main_content_type() {
    let presentation =
        Presentation::from_bytes(&package_bytes(embedded_fixture_package(false))).unwrap();
    let original = presentation.to_bytes().unwrap();
    let original_package = open_opc(&original, "F-223 original class");

    for class in [
        PresentationPackageClass::Presentation,
        PresentationPackageClass::MacroEnabledPresentation,
        PresentationPackageClass::Template,
        PresentationPackageClass::MacroEnabledTemplate,
        PresentationPackageClass::Slideshow,
        PresentationPackageClass::MacroEnabledSlideshow,
    ] {
        let converted = presentation.to_bytes_as(class).unwrap();
        let package = open_opc(&converted, "F-223 converted class");
        assert_eq!(
            Presentation::from_bytes(&converted)
                .unwrap()
                .package_class()
                .unwrap(),
            class
        );
        assert_eq!(package.parts, original_package.parts, "{class:?}");
        assert_eq!(
            package.package_rels.items, original_package.package_rels.items,
            "{class:?}"
        );
        assert_f223_relationships_equal(&package.part_rels, &original_package.part_rels, class);
        let mut expected = original_package.content_types.clone();
        expected.overrides.insert(
            PRESENTATION_PART.to_owned(),
            f223_content_type(class).to_owned(),
        );
        assert_eq!(package.content_types, expected, "{class:?}");
    }
    assert_eq!(
        presentation.package_class().unwrap(),
        PresentationPackageClass::Presentation
    );
    assert_eq!(presentation.to_bytes().unwrap(), original);
}

#[test]
fn ordinary_save_preserves_opened_template_and_macro_classes() {
    for class in f223_variant_classes() {
        let presentation =
            Presentation::from_bytes(&package_bytes(f223_fixture_package(class))).unwrap();
        let path =
            std::env::temp_dir().join(format!("rpptx-f223-{}-{:?}.bin", std::process::id(), class));
        presentation.save(&path).unwrap();
        let reopened = Presentation::open(&path).unwrap();
        fs::remove_file(path).unwrap();
        assert_eq!(reopened.package_class().unwrap(), class);
    }
}

#[test]
fn package_class_conversion_preserves_and_invalidates_signature_evidence() {
    let presentation =
        Presentation::from_bytes(&package_bytes(embedded_fixture_package(true))).unwrap();
    let converted = presentation
        .to_bytes_as(PresentationPackageClass::MacroEnabledTemplate)
        .unwrap();
    let reopened = Presentation::from_bytes(&converted).unwrap();
    assert_eq!(
        reopened.package_class().unwrap(),
        PresentationPackageClass::MacroEnabledTemplate
    );
    let inventory = reopened.embedded_content().unwrap();
    assert_eq!(inventory.len(), 3);
    assert!(inventory.iter().all(|item| {
        item.signature_state == EmbeddedSignatureState::Invalidated
            && reopened
                .extract_embedded_content(&item.source_part, &item.relationship_id)
                .is_ok()
    }));
    let package = open_opc(&converted, "F-223 signed class conversion");
    assert_eq!(
        package.get_part("/_xmlsignatures/sig1.xml"),
        Some(b"package-signature-evidence".as_slice())
    );
    assert_eq!(
        package.get_part("/custom/vbaSignature.bin"),
        Some(b"vba-signature-evidence".as_slice())
    );
}

#[test]
fn unknown_presentation_main_content_type_fails_closed() {
    let mut package = fixture_package();
    package.content_types.add_override(
        PRESENTATION_PART,
        "application/vnd.example.presentation.unknown.main+xml",
    );
    assert!(matches!(
        Presentation::from_bytes(&package_bytes(package)),
        Err(Error::UnsupportedPackageClass {
            content_type: Some(content_type)
        }) if content_type == "application/vnd.example.presentation.unknown.main+xml"
    ));
}

fn f223_variant_classes() -> [PresentationPackageClass; 5] {
    [
        PresentationPackageClass::MacroEnabledPresentation,
        PresentationPackageClass::Template,
        PresentationPackageClass::MacroEnabledTemplate,
        PresentationPackageClass::Slideshow,
        PresentationPackageClass::MacroEnabledSlideshow,
    ]
}

fn f223_fixture_package(class: PresentationPackageClass) -> OpcPackage {
    let mut package = if matches!(
        class,
        PresentationPackageClass::MacroEnabledPresentation
            | PresentationPackageClass::MacroEnabledTemplate
            | PresentationPackageClass::MacroEnabledSlideshow
    ) {
        embedded_fixture_package(false)
    } else {
        let mut package = fixture_package();
        package.set_part("/custom/variant-opaque.bin", b"variant-opaque".to_vec());
        package.content_types.add_override(
            "/custom/variant-opaque.bin",
            "application/vnd.example.variant-opaque",
        );
        package
    };
    package
        .content_types
        .add_override(PRESENTATION_PART, f223_content_type(class));
    package
}

fn f223_content_type(class: PresentationPackageClass) -> &'static str {
    match class {
        PresentationPackageClass::Presentation => content_types::PRESENTATION,
        PresentationPackageClass::MacroEnabledPresentation => {
            content_types::PRESENTATION_MACRO_ENABLED
        }
        PresentationPackageClass::Template => content_types::PRESENTATION_TEMPLATE,
        PresentationPackageClass::MacroEnabledTemplate => {
            content_types::PRESENTATION_TEMPLATE_MACRO_ENABLED
        }
        PresentationPackageClass::Slideshow => content_types::SLIDESHOW,
        PresentationPackageClass::MacroEnabledSlideshow => content_types::SLIDESHOW_MACRO_ENABLED,
    }
}

fn assert_f223_relationships_equal(
    actual: &std::collections::HashMap<String, oxml_opc::Relationships>,
    expected: &std::collections::HashMap<String, oxml_opc::Relationships>,
    class: PresentationPackageClass,
) {
    assert_eq!(actual.len(), expected.len(), "{class:?}");
    for (part, relationships) in expected {
        assert_eq!(
            actual.get(part).map(|actual| &actual.items),
            Some(&relationships.items),
            "{class:?}: {part}"
        );
    }
}

#[test]
fn save_as_show_changes_only_the_main_content_type() {
    let presentation = Presentation::from_bytes(&package_bytes(fixture_package())).unwrap();
    let ordinary_bytes = presentation.to_bytes().unwrap();
    let output = std::env::temp_dir().join(format!("rpptx-f115-show-{}.ppsx", std::process::id()));
    presentation.save_as_show(&output).expect("save slideshow");
    let show_bytes = fs::read(&output).expect("read slideshow");
    fs::remove_file(&output).expect("remove slideshow");

    let ordinary = open_opc(&ordinary_bytes, "ordinary presentation");
    let show = open_opc(&show_bytes, "slideshow presentation");
    assert_eq!(
        show.content_types.content_type_for(PRESENTATION_PART),
        Some(content_types::SLIDESHOW)
    );
    assert_eq!(ordinary.parts, show.parts);
    assert_eq!(ordinary.package_rels.items, show.package_rels.items);
    assert_eq!(ordinary.part_rels.len(), show.part_rels.len());
    for (part_name, relationships) in &ordinary.part_rels {
        assert_eq!(
            relationships.items,
            show.part_rels.get(part_name).unwrap().items,
            "{part_name}: relationship scope"
        );
    }
    let mut ordinary_overrides = ordinary.content_types.overrides;
    ordinary_overrides.insert(
        PRESENTATION_PART.to_owned(),
        content_types::SLIDESHOW.to_owned(),
    );
    assert_eq!(ordinary_overrides, show.content_types.overrides);
    assert_eq!(ordinary.content_types.defaults, show.content_types.defaults);
    assert_eq!(presentation.to_bytes().unwrap(), ordinary_bytes);
    assert_eq!(
        Presentation::from_bytes(&show_bytes).unwrap().len(),
        presentation.len()
    );
}

#[test]
fn invalid_slide_size_does_not_mutate_the_presentation() {
    let mut presentation = Presentation::new().expect("open bundled template");
    let before = presentation.to_bytes().unwrap();

    assert!(presentation.set_slide_size(Emu(0), Emu(5_000_000)).is_err());
    assert!(
        presentation
            .set_slide_size(Emu(10_000_000), Emu(-1))
            .is_err()
    );

    assert_eq!(presentation.to_bytes().unwrap(), before);
}

#[test]
fn invalid_slide_collection_indices_do_not_mutate_the_presentation() {
    let mut presentation = Presentation::new().expect("open bundled template");
    presentation.add_slide(0).expect("add slide");
    let before = presentation.to_bytes().expect("serialize baseline");

    assert!(presentation.remove_slide(1).is_err());
    assert!(presentation.move_slide(1, 0).is_err());
    assert!(presentation.move_slide(0, 1).is_err());
    assert!(presentation.duplicate_slide(1).is_err());
    assert_eq!(presentation.to_bytes().unwrap(), before);
}

#[test]
fn move_slide_reorders_the_slide_id_list_without_rewriting_relationships() {
    let mut package = fixture_package();
    let marked = String::from_utf8(package.get_part(PRESENTATION_PART).unwrap().to_vec())
        .unwrap()
        .replacen(
            r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#,
            r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:f114""#,
            1,
        )
        .replacen(
            r#"r:id="ordered-first"/>"#,
            r#"r:id="ordered-first"/><x:between producer="kept"/>"#,
            1,
        );
    package.set_part(PRESENTATION_PART, marked.into_bytes());
    let before_relationships = package
        .get_part_rels(PRESENTATION_PART)
        .unwrap()
        .items
        .clone();
    let mut presentation = Presentation::from_bytes(&package_bytes(package)).unwrap();
    let before = presentation
        .slides()
        .map(|slide| slide.id())
        .collect::<Vec<_>>();

    presentation.move_slide(0, 1).expect("move first slide");

    let after = presentation
        .slides()
        .map(|slide| slide.id())
        .collect::<Vec<_>>();
    assert_eq!(after, vec![before[1], before[0]]);
    let bytes = presentation.to_bytes().unwrap();
    let package = open_opc(&bytes, "moved slide package");
    assert_eq!(
        package.get_part_rels(PRESENTATION_PART).unwrap().items,
        before_relationships
    );
    let presentation_xml =
        String::from_utf8(package.get_part(PRESENTATION_PART).unwrap().to_vec()).unwrap();
    assert!(
        presentation_xml
            .find("<x:between producer=\"kept\"/>")
            .unwrap()
            < presentation_xml.find(r#"r:id="ordered-second""#).unwrap()
    );
    let reopened = Presentation::from_bytes(&bytes).unwrap();
    assert_eq!(
        reopened
            .slides()
            .map(|slide| slide.id())
            .collect::<Vec<_>>(),
        after
    );
}

#[test]
fn remove_slide_removes_its_part_relationship_notes_and_custom_show_entries() {
    let mut package = fixture_package();
    let with_custom_show = String::from_utf8(package.get_part(PRESENTATION_PART).unwrap().to_vec())
        .unwrap()
        .replacen(
            "<p:notesSz",
            r#"<p:custShowLst xmlns:x="urn:f114" x:producer="kept"><p:custShow name="review" id="1"><p:sldLst><p:sld r:id="ordered-first"/><x:marker/><p:sld r:id="ordered-second"/></p:sldLst></p:custShow></p:custShowLst><p:notesSz"#,
            1,
        );
    package.set_part(PRESENTATION_PART, with_custom_show.into_bytes());
    let mut presentation = Presentation::from_bytes(&package_bytes(package)).unwrap();
    assert!(presentation.slide(0).unwrap().notes_text().is_some());

    presentation
        .remove_slide(0)
        .expect("remove slide with notes");

    let bytes = presentation.to_bytes().unwrap();
    let package = open_opc(&bytes, "removed slide package");
    assert!(!package.parts.contains_key(SLIDE_TWO_PART));
    assert!(!package.parts.contains_key(NOTES_PART));
    assert!(!package.part_rels.contains_key(SLIDE_TWO_PART));
    assert!(!package.part_rels.contains_key(NOTES_PART));
    assert!(!package.content_types.overrides.contains_key(SLIDE_TWO_PART));
    assert!(!package.content_types.overrides.contains_key(NOTES_PART));
    assert!(
        package
            .get_part_rels(PRESENTATION_PART)
            .unwrap()
            .get_by_id("ordered-first")
            .is_none()
    );
    let presentation_xml =
        String::from_utf8(package.get_part(PRESENTATION_PART).unwrap().to_vec()).unwrap();
    assert_eq!(
        presentation_xml.matches(r#"r:id="ordered-first""#).count(),
        0
    );
    assert!(presentation_xml.contains("<x:marker/>"));
    assert!(presentation_xml.contains(r#"r:id="ordered-second""#));
    let reopened = Presentation::from_bytes(&bytes).unwrap();
    assert_eq!(reopened.len(), 1);
    assert!(reopened.validate().is_empty());
}

#[test]
fn removing_shared_and_last_image_users_prunes_only_new_orphans() {
    let png = png_header(4, 3);
    let mut presentation = Presentation::new().expect("open bundled template");
    for _ in 0..2 {
        presentation.add_slide(0).expect("add slide");
    }
    for index in 0..2 {
        presentation
            .add_picture(index, &png, "shared.png", Emu(0), Emu(0), None, None)
            .unwrap();
    }

    presentation.remove_slide(0).expect("remove one image user");
    let shared = open_opc(&presentation.to_bytes().unwrap(), "shared image remains");
    assert_eq!(
        shared
            .parts
            .iter()
            .filter(|(part, bytes)| part.starts_with("/ppt/media/") && *bytes == &png)
            .count(),
        1
    );

    presentation
        .remove_slide(0)
        .expect("remove last image user");
    let pruned = open_opc(&presentation.to_bytes().unwrap(), "last image pruned");
    assert_eq!(
        pruned
            .parts
            .iter()
            .filter(|(part, bytes)| part.starts_with("/ppt/media/") && *bytes == &png)
            .count(),
        0
    );
}

#[test]
fn removing_last_slide_image_preserves_a_package_root_relationship_target() {
    let png = png_header(4, 3);
    let mut presentation = Presentation::new().expect("open bundled template");
    presentation.add_slide(0).expect("add slide");
    presentation
        .add_picture(0, &png, "root-shared.png", Emu(0), Emu(0), None, None)
        .unwrap();
    let mut package = open_opc(&presentation.to_bytes().unwrap(), "root media fixture");
    let media_part = package
        .parts
        .keys()
        .find(|part| part.starts_with("/ppt/media/"))
        .unwrap()
        .clone();
    package
        .package_rels
        .add(rel_types::THUMBNAIL, media_part.trim_start_matches('/'));
    let mut presentation = Presentation::from_bytes(&package_bytes(package)).unwrap();

    presentation.remove_slide(0).expect("remove final slide");

    let package = open_opc(&presentation.to_bytes().unwrap(), "root media result");
    let relationship = package
        .package_rels
        .get_all_by_type(rel_types::THUMBNAIL)
        .into_iter()
        .find(|relationship| {
            OpcPackage::resolve_rel_target("/", &relationship.target) == media_part
        })
        .unwrap();
    assert_eq!(
        OpcPackage::resolve_rel_target("/", &relationship.target),
        media_part
    );
    assert_eq!(package.get_part(&media_part).unwrap(), png);
    assert!(presentation.validate().is_empty());
}

#[test]
fn duplicated_slides_images_resolve_to_the_new_slides_own_relationships() {
    let png = png_header(4, 3);
    let mut presentation = Presentation::new().expect("open bundled template");
    presentation.add_slide(0).expect("add source slide");
    presentation
        .add_picture(0, &png, "source.png", Emu(0), Emu(0), None, None)
        .unwrap();

    presentation.duplicate_slide(0).expect("duplicate slide");

    let bytes = presentation.to_bytes().unwrap();
    let package = open_opc(&bytes, "duplicated image scopes");
    let presentation_part = package.main_document_part().unwrap();
    let model = CT_Presentation::from_xml(package.get_part(&presentation_part).unwrap()).unwrap();
    let mut image_relationship_ids = Vec::new();
    let mut image_targets = Vec::new();
    for slide_id in &model.slide_ids {
        let slide_relationship = package
            .get_part_rels(&presentation_part)
            .unwrap()
            .get_by_id(&slide_id.relationship_id)
            .unwrap();
        let slide_part =
            OpcPackage::resolve_rel_target(&presentation_part, &slide_relationship.target);
        let slide = CT_Slide::from_xml(package.get_part(&slide_part).unwrap()).unwrap();
        let embed = slide
            .common_slide_data
            .shape_tree
            .children
            .iter()
            .find_map(|child| match child {
                ShapeTreeChild::Picture(picture) => picture
                    .blip_fill
                    .as_ref()
                    .and_then(|fill| fill.blip.as_ref())
                    .and_then(|blip| blip.embed.clone()),
                _ => None,
            })
            .unwrap();
        let image_relationship = package
            .get_part_rels(&slide_part)
            .unwrap()
            .get_by_id(&embed)
            .unwrap();
        assert_eq!(image_relationship.rel_type, rel_types::IMAGE);
        image_relationship_ids.push(embed);
        image_targets.push(OpcPackage::resolve_rel_target(
            &slide_part,
            &image_relationship.target,
        ));
    }
    assert_eq!(image_relationship_ids.len(), 2);
    assert_eq!(image_targets[0], image_targets[1]);
    assert_eq!(package.get_part(&image_targets[0]).unwrap(), png);
    assert_eq!(
        package
            .parts
            .keys()
            .filter(|part| part.starts_with("/ppt/media/"))
            .count(),
        1
    );
    let reopened = Presentation::from_bytes(&bytes).unwrap();
    assert_eq!(reopened.len(), 2);
    assert!(reopened.validate().is_empty());
}

#[test]
fn duplicate_slide_rewrites_typed_and_preserved_relationship_ids_without_other_byte_changes() {
    let mut package = open_opc(&mutation_fixture_bytes(), "relationship rewrite fixture");
    let image_part = "/custom/media/preserved.png";
    let image = png_header(3, 2);
    package.set_part(image_part, image.clone());
    package.content_types.add_default("png", "image/png");
    package.get_or_create_part_rels(SLIDE_TWO_PART).add_with_id(
        "rId7",
        rel_types::IMAGE,
        "../media/preserved.png",
    );
    let marked = String::from_utf8(package.get_part(SLIDE_TWO_PART).unwrap().to_vec())
        .unwrap()
        .replacen(
            "<p:blipFill/>",
            &format!(
                r#"<p:blipFill><a:blip xmlns:r="{R_NS}" r:embed="rId7"><x:effect xmlns:x="urn:f114"> A &amp; B </x:effect></a:blip></p:blipFill>"#
            ),
            1,
        )
        .replacen(
            "</p:cSld>",
            &format!(
                r#"<x:preserved xmlns:x="urn:f114" xmlns:r="{R_NS}" r:id="rId7" syntax=" A &amp; B "/></p:cSld>"#
            ),
            1,
        );
    package.set_part(SLIDE_TWO_PART, marked.into_bytes());
    let mut presentation = Presentation::from_bytes(&package_bytes(package)).unwrap();
    let source_text = presentation.slide(0).unwrap().text();

    presentation
        .duplicate_slide(0)
        .expect("duplicate marked slide");

    assert_eq!(presentation.slide(1).unwrap().text(), source_text);
    let bytes = presentation.to_bytes().unwrap();
    let package = open_opc(&bytes, "relationship rewrite result");
    let model = CT_Presentation::from_xml(package.get_part(PRESENTATION_PART).unwrap()).unwrap();
    let duplicate_relationship = package
        .get_part_rels(PRESENTATION_PART)
        .unwrap()
        .get_by_id(&model.slide_ids[1].relationship_id)
        .unwrap();
    let duplicate_part =
        OpcPackage::resolve_rel_target(PRESENTATION_PART, &duplicate_relationship.target);
    let duplicate_xml =
        String::from_utf8(package.get_part(&duplicate_part).unwrap().to_vec()).unwrap();
    assert!(duplicate_xml.contains(r#"r:embed="rId2""#));
    assert!(duplicate_xml.contains(r#"r:id="rId2" syntax=" A &amp; B "/>"#));
    assert!(duplicate_xml.contains(r#"<x:effect xmlns:x="urn:f114"> A &amp; B </x:effect>"#));
    assert!(!duplicate_xml.contains("rId7"));
    let image_relationship = package
        .get_part_rels(&duplicate_part)
        .unwrap()
        .get_by_id("rId2")
        .unwrap();
    let duplicate_image =
        OpcPackage::resolve_rel_target(&duplicate_part, &image_relationship.target);
    assert_eq!(package.get_part(&duplicate_image).unwrap(), image);
    assert!(Presentation::from_bytes(&bytes).is_ok());
}

#[test]
fn duplicate_slide_assigns_fresh_shape_ids_and_rewrites_connector_endpoints() {
    let mut presentation = Presentation::from_bytes(&mutation_fixture_bytes()).unwrap();

    presentation
        .duplicate_slide(0)
        .expect("duplicate connected shapes");

    let package = open_opc(
        &presentation.to_bytes().unwrap(),
        "duplicated connector ids",
    );
    let model = CT_Presentation::from_xml(package.get_part(PRESENTATION_PART).unwrap()).unwrap();
    let slide_parts = model
        .slide_ids
        .iter()
        .map(|slide_id| {
            let relationship = package
                .get_part_rels(PRESENTATION_PART)
                .unwrap()
                .get_by_id(&slide_id.relationship_id)
                .unwrap();
            OpcPackage::resolve_rel_target(PRESENTATION_PART, &relationship.target)
        })
        .collect::<Vec<_>>();
    let source = String::from_utf8(package.get_part(&slide_parts[0]).unwrap().to_vec()).unwrap();
    let duplicate = String::from_utf8(package.get_part(&slide_parts[1]).unwrap().to_vec()).unwrap();
    assert!(source.contains(r#"<a:stCxn id="2" idx="0"/><a:endCxn id="3" idx="1"/>"#));
    assert!(duplicate.contains(r#"<a:stCxn id="9" idx="0"/><a:endCxn id="10" idx="1"/>"#));
    assert!(source.contains(r#"<p:cNvPr id="20" name="choice shape"/>"#));
    assert!(source.contains(r#"<a:stCxn id="20" idx="0"/><a:endCxn id="20" idx="1"/>"#));
    assert!(source.contains(r#"<p:cNvPr id="22" name="fallback shape"/>"#));
    assert!(source.contains(r#"<a:stCxn id="22" idx="0"/><a:endCxn id="22" idx="1"/>"#));
    assert!(duplicate.contains(r#"<p:cNvPr id="16" name="choice shape"/>"#));
    assert!(duplicate.contains(r#"<p:cNvPr id="17" name="choice connector"/>"#));
    assert!(duplicate.contains(r#"<a:stCxn id="16" idx="0"/><a:endCxn id="16" idx="1"/>"#));
    assert!(duplicate.contains(r#"<p:cNvPr id="18" name="fallback shape"/>"#));
    assert!(duplicate.contains(r#"<p:cNvPr id="19" name="fallback connector"/>"#));
    assert!(duplicate.contains(r#"<a:stCxn id="18" idx="0"/><a:endCxn id="18" idx="1"/>"#));
    for stale_id in [20, 21, 22, 23] {
        assert!(!duplicate.contains(&format!(r#"id="{stale_id}""#)));
    }
    let duplicate_slide = CT_Slide::from_xml(duplicate.as_bytes()).unwrap();
    let connector = duplicate_slide
        .common_slide_data
        .shape_tree
        .children
        .iter()
        .find_map(|child| match child {
            ShapeTreeChild::Connector(connector) => Some(connector),
            _ => None,
        })
        .unwrap();
    assert_eq!(connector.start_connection.as_ref().unwrap().id, 9);
    assert_eq!(connector.end_connection.as_ref().unwrap().id, 10);
    assert_eq!(
        duplicate_slide
            .common_slide_data
            .shape_tree
            .children
            .iter()
            .find_map(|child| matches!(child, ShapeTreeChild::Connector(_))
                .then(|| child.non_visual_id()))
            .flatten(),
        Some(15)
    );
    assert!(!duplicate.contains(r#"<p:cNvPr id="2""#));
    assert!(presentation.validate().is_empty());
}

#[test]
fn duplicate_slide_rewrites_schema_shape_targets_without_touching_raw_lookalikes() {
    let mut package = open_opc(&mutation_fixture_bytes(), "schema shape target fixture");
    let source_xml = String::from_utf8(package.get_part(SLIDE_TWO_PART).unwrap().to_vec()).unwrap();
    let raw_build_lookalikes = r#"<x:bldOleChart xmlns:x="urn:f215" spid="2" keep='A&#x20;B'/><p:bldOleChart xmlns:x="urn:f215" x:spid="2"/>"#;
    let raw_ink_lookalikes = r#"<x:inkTgt xmlns:x="urn:f215" spid="2" keep='C&#x20;D'/><x:wrapper xmlns:x="urn:f215"><p:inkTgt x:spid="2"/></x:wrapper>"#;
    let timing = format!(
        r#"<p:timing><p:tnLst><p:set><p:cBhvr><p:cTn id="1"/><p:tgtEl><p:inkTgt spid="2"/>{raw_ink_lookalikes}</p:tgtEl></p:cBhvr></p:set></p:tnLst><p:bldLst><p:bldOleChart spid="2" grpId="7"/>{raw_build_lookalikes}</p:bldLst></p:timing>"#
    );
    assert!(!source_xml.contains("<p:timing"));
    package.set_part(
        SLIDE_TWO_PART,
        source_xml
            .replacen("</p:sld>", &format!("{timing}</p:sld>"), 1)
            .into_bytes(),
    );
    let mut presentation = Presentation::from_bytes(&package_bytes(package)).unwrap();

    presentation
        .duplicate_slide(0)
        .expect("duplicate schema shape targets");

    let package = open_opc(
        &presentation.to_bytes().unwrap(),
        "duplicated schema shape targets",
    );
    let model = CT_Presentation::from_xml(package.get_part(PRESENTATION_PART).unwrap()).unwrap();
    let duplicate_relationship = package
        .get_part_rels(PRESENTATION_PART)
        .unwrap()
        .get_by_id(&model.slide_ids[1].relationship_id)
        .unwrap();
    let duplicate_part =
        OpcPackage::resolve_rel_target(PRESENTATION_PART, &duplicate_relationship.target);
    let duplicate = String::from_utf8(package.get_part(&duplicate_part).unwrap().to_vec()).unwrap();
    assert!(duplicate.contains(r#"<p:cNvPr id="9" name="old""#));
    assert!(duplicate.contains(r#"<p:inkTgt spid="9"/>"#));
    assert!(duplicate.contains(r#"<p:bldOleChart spid="9" grpId="7"/>"#));
    assert!(duplicate.contains(raw_ink_lookalikes));
    assert!(duplicate.contains(raw_build_lookalikes));
    assert!(presentation.validate().is_empty());
}

#[test]
fn duplicate_slide_copies_notes_with_a_fresh_back_relationship() {
    let mut presentation = Presentation::from_bytes(&fixture_bytes()).unwrap();
    let notes = presentation.slide(0).unwrap().notes_text();

    presentation
        .duplicate_slide(0)
        .expect("duplicate slide with notes");

    assert_eq!(presentation.slide(1).unwrap().notes_text(), notes);
    let bytes = presentation.to_bytes().unwrap();
    let package = open_opc(&bytes, "duplicated notes scope");
    let model = CT_Presentation::from_xml(package.get_part(PRESENTATION_PART).unwrap()).unwrap();
    let mut notes_parts = Vec::new();
    for slide_id in model.slide_ids.iter().take(2) {
        let slide_relationship = package
            .get_part_rels(PRESENTATION_PART)
            .unwrap()
            .get_by_id(&slide_id.relationship_id)
            .unwrap();
        let slide_part =
            OpcPackage::resolve_rel_target(PRESENTATION_PART, &slide_relationship.target);
        let notes_relationship = package
            .get_part_rels(&slide_part)
            .unwrap()
            .items
            .iter()
            .find(|relationship| relationship.rel_type == rel_types::NOTES_SLIDE)
            .unwrap();
        let notes_part = OpcPackage::resolve_rel_target(&slide_part, &notes_relationship.target);
        let back_relationship = package
            .get_part_rels(&notes_part)
            .unwrap()
            .items
            .iter()
            .find(|relationship| relationship.rel_type == rel_types::SLIDE)
            .unwrap();
        assert_eq!(
            OpcPackage::resolve_rel_target(&notes_part, &back_relationship.target),
            slide_part
        );
        notes_parts.push(notes_part);
    }
    assert_ne!(notes_parts[0], notes_parts[1]);
    assert!(Presentation::from_bytes(&bytes).is_ok());
}

#[test]
fn duplicate_slide_adds_a_missing_notes_back_relationship() {
    let mut package = fixture_package();
    package
        .get_or_create_part_rels(NOTES_PART)
        .items
        .retain(|relationship| relationship.rel_type != rel_types::SLIDE);

    assert_notes_back_relationship_is_normalized(package, &HashSet::new());
}

#[test]
fn duplicate_slide_collapses_multiple_notes_back_relationships() {
    let mut package = fixture_package();
    package.get_or_create_part_rels(NOTES_PART).add_with_id(
        "second-slide-link",
        rel_types::SLIDE,
        "../slides/first.xml",
    );
    let source_ids = package
        .get_part_rels(NOTES_PART)
        .unwrap()
        .items
        .iter()
        .filter(|relationship| relationship.rel_type == rel_types::SLIDE)
        .map(|relationship| relationship.id.clone())
        .collect();

    assert_notes_back_relationship_is_normalized(package, &source_ids);
}

#[test]
fn duplicate_slide_replaces_an_external_notes_back_relationship() {
    let mut package = fixture_package();
    let relationship = package
        .get_or_create_part_rels(NOTES_PART)
        .items
        .iter_mut()
        .find(|relationship| relationship.rel_type == rel_types::SLIDE)
        .unwrap();
    relationship.target = "https://example.com/source-slide".to_owned();
    relationship.target_mode = Some("External".to_owned());
    let source_ids = HashSet::from([relationship.id.clone()]);

    assert_notes_back_relationship_is_normalized(package, &source_ids);
}

#[test]
fn duplicate_slide_rewrites_preserved_nonnumeric_notes_back_references() {
    let mut package = fixture_package();
    let marked_notes = String::from_utf8(package.get_part(NOTES_PART).unwrap().to_vec())
        .unwrap()
        .replacen(
            "</p:cSld>",
            &format!(
                r#"<x:back xmlns:x="urn:f114" xmlns:r="{R_NS}" r:id="slide-link" syntax=" A &amp; B "/><x:other xmlns:x="urn:f114" xmlns:r="{R_NS}" r:id="other-link" syntax=" untouched "/></p:cSld>"#
            ),
            1,
        );
    package.set_part(NOTES_PART, marked_notes.into_bytes());
    package.get_or_create_part_rels(NOTES_PART).add_with_id(
        "other-link",
        "urn:f114:unrelated",
        "../opaque/data.bin",
    );
    let mut presentation = Presentation::from_bytes(&package_bytes(package)).unwrap();

    presentation.duplicate_slide(0).expect("duplicate slide");
    let saved = presentation.to_bytes().unwrap();
    let reopened = Presentation::from_bytes(&saved).unwrap();
    let resaved = reopened.to_bytes().unwrap();
    let package = open_opc(&resaved, "nonnumeric notes back reference result");
    let model = CT_Presentation::from_xml(package.get_part(PRESENTATION_PART).unwrap()).unwrap();
    let duplicate_relationship = package
        .get_part_rels(PRESENTATION_PART)
        .unwrap()
        .get_by_id(&model.slide_ids[1].relationship_id)
        .unwrap();
    let duplicate_slide =
        OpcPackage::resolve_rel_target(PRESENTATION_PART, &duplicate_relationship.target);
    let notes_relationship = package
        .get_part_rels(&duplicate_slide)
        .unwrap()
        .items
        .iter()
        .find(|relationship| relationship.rel_type == rel_types::NOTES_SLIDE)
        .unwrap();
    let duplicate_notes =
        OpcPackage::resolve_rel_target(&duplicate_slide, &notes_relationship.target);
    let notes_relationships = package.get_part_rels(&duplicate_notes).unwrap();
    let back_relationships = notes_relationships
        .items
        .iter()
        .filter(|relationship| relationship.rel_type == rel_types::SLIDE)
        .collect::<Vec<_>>();
    assert_eq!(back_relationships.len(), 1);
    let back_relationship = back_relationships[0];
    assert_ne!(back_relationship.id, "slide-link");
    assert_eq!(back_relationship.target_mode, None);
    assert_eq!(
        OpcPackage::resolve_rel_target(&duplicate_notes, &back_relationship.target),
        duplicate_slide
    );
    let duplicate_xml =
        String::from_utf8(package.get_part(&duplicate_notes).unwrap().to_vec()).unwrap();
    assert!(duplicate_xml.contains(&format!(
        r#"r:id="{}" syntax=" A &amp; B "/>"#,
        back_relationship.id
    )));
    assert!(duplicate_xml.contains(
        r#"<x:other xmlns:x="urn:f114" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="other-link" syntax=" untouched "/>"#
    ));
    assert!(!duplicate_xml.contains(r#"r:id="slide-link""#));
    let unrelated = notes_relationships.get_by_id("other-link").unwrap();
    assert_eq!(unrelated.rel_type, "urn:f114:unrelated");
    assert_eq!(unrelated.target_mode, None);
    assert_eq!(
        OpcPackage::resolve_rel_target(&duplicate_notes, &unrelated.target),
        "/custom/opaque/data.bin"
    );
    assert!(reopened.validate().is_empty());
}

#[test]
fn remove_move_and_duplicate_round_trip_with_clean_validation() {
    let mut presentation = Presentation::new().expect("open bundled template");
    for _ in 0..3 {
        presentation.add_slide(0).expect("add slide");
    }

    presentation
        .duplicate_slide(1)
        .expect("duplicate middle slide");
    presentation.move_slide(3, 0).expect("move duplicate first");
    presentation.remove_slide(2).expect("remove one slide");

    let reopened = Presentation::from_bytes(&presentation.to_bytes().unwrap()).unwrap();
    assert_eq!(reopened.len(), 3);
    assert!(reopened.validate().is_empty());
}

#[test]
fn merge_then_split_restores_the_original_grid() {
    let mut presentation = Presentation::new().expect("open bundled template");
    presentation.add_slide(0).expect("add slide");
    {
        let mut slide = presentation.slide_mut(0).expect("borrow slide");
        let mut shape = slide
            .add_table(2, 3, Emu(10), Emu(20), Emu(302), Emu(201))
            .expect("add table");
        let mut table = shape.table_mut().expect("table handle");
        let widths = (0..3)
            .map(|column| table.column_width(column).unwrap())
            .collect::<Vec<_>>();
        table.cell_mut(0, 0).unwrap().merge_to(1, 2).unwrap();
        table.cell_mut(0, 0).unwrap().split().unwrap();
        assert_eq!(
            (0..3)
                .map(|column| table.column_width(column).unwrap())
                .collect::<Vec<_>>(),
            widths
        );
    }

    let saved = presentation.to_bytes().expect("save presentation");
    let reopened = Presentation::from_bytes(&saved).expect("reopen presentation");
    let slide = reopened.slide(0).unwrap();
    let table = slide.shapes().last().unwrap().table().unwrap();
    assert_eq!(table.row_count(), 2);
    assert_eq!(table.column_count(), 3);
    assert!((0..2).all(|row| (0..3).all(|column| {
        let cell = table.cell(row, column).unwrap();
        !cell.is_merge_origin()
            && !cell.is_spanned()
            && cell.span_height() == 1
            && cell.span_width() == 1
    })));
    assert_eq!(
        (0..3)
            .map(|column| table.column_width(column).unwrap())
            .collect::<Vec<_>>(),
        vec![Emu(100), Emu(100), Emu(102)]
    );

    let package = open_opc(&saved, "F-113 merge then split");
    let slide = CT_Slide::from_xml(package.get_part("/ppt/slides/slide1.xml").unwrap()).unwrap();
    let ShapeTreeChild::GraphicFrame(frame) =
        slide.common_slide_data.shape_tree.children.last().unwrap()
    else {
        panic!("expected table graphic frame");
    };
    let GraphicDataPayload::Table(table) = frame.graphic_data.payload() else {
        panic!("expected table payload");
    };
    assert_eq!(
        table.rows.iter().map(|row| row.height).collect::<Vec<_>>(),
        vec![Emu(100), Emu(101)]
    );
    assert!(table.rows.iter().all(|row| row.cells.len() == 3));
}

#[test]
fn add_table_round_trips_cells_formatting_banding_and_widths() {
    let mut presentation = Presentation::new().expect("open bundled template");
    presentation.add_slide(0).expect("add slide");
    let fill = Fill::from_xml(br#"<a:solidFill><a:srgbClr val="112233"/></a:solidFill>"#).unwrap();
    {
        let mut slide = presentation.slide_mut(0).unwrap();
        let mut shape = slide
            .add_table(2, 2, Emu(10), Emu(20), Emu(201), Emu(101))
            .expect("add table");
        let mut table = shape.table_mut().unwrap();
        table.set_first_row(true);
        table.set_last_row(true);
        table.set_first_column(true);
        table.set_last_column(true);
        table.set_horizontal_banding(true);
        table.set_vertical_banding(true);
        table.set_column_width(1, Emu(150)).unwrap();
        let mut cell = table.cell_mut(0, 0).unwrap();
        cell.set_text("formatted");
        cell.text_frame()
            .add_paragraph()
            .set_text("second paragraph");
        cell.set_fill(Some(fill.clone()));
        cell.set_margins(Some(Emu(10)), Some(Emu(20)), Some(Emu(30)), Some(Emu(40)));
    }

    let saved = presentation.to_bytes().unwrap();
    let reopened = Presentation::from_bytes(&saved).unwrap();
    let slide = reopened.slide(0).unwrap();
    let table = slide.shapes().last().unwrap().table().unwrap();
    assert_eq!(table.row_count(), 2);
    assert_eq!(table.column_count(), 2);
    assert_eq!(table.column_width(0), Some(Emu(100)));
    assert_eq!(table.column_width(1), Some(Emu(150)));
    assert!(table.first_row());
    assert!(table.last_row());
    assert!(table.first_column());
    assert!(table.last_column());
    assert!(table.horizontal_banding());
    assert!(table.vertical_banding());
    let cell = table.cell(0, 0).unwrap();
    assert_eq!(cell.text(), "formatted\nsecond paragraph");
    assert_eq!(cell.fill(), Some(&fill));
    assert_eq!(
        cell.margins(),
        (Some(Emu(10)), Some(Emu(20)), Some(Emu(30)), Some(Emu(40)),)
    );

    let package = open_opc(&saved, "F-113 table formatting");
    let slide = CT_Slide::from_xml(package.get_part("/ppt/slides/slide1.xml").unwrap()).unwrap();
    let ShapeTreeChild::GraphicFrame(frame) =
        slide.common_slide_data.shape_tree.children.last().unwrap()
    else {
        panic!("expected table graphic frame");
    };
    assert_eq!(frame.transform.extent.unwrap().cx, Emu(250));
    assert_eq!(frame.transform.extent.unwrap().cy, Emu(101));
}

#[test]
fn table_mutation_rejects_invalid_ranges_without_partial_changes() {
    let mut presentation = Presentation::new().expect("open bundled template");
    presentation.add_slide(0).expect("add slide");
    let before = presentation.to_bytes().unwrap();
    let rejected = {
        let mut slide = presentation.slide_mut(0).unwrap();
        matches!(
            slide.add_table(0, 2, Emu(0), Emu(0), Emu(200), Emu(200)),
            Err(Error::InvalidTableMutation { .. })
        )
    };
    assert!(rejected);
    assert_eq!(presentation.to_bytes().unwrap(), before);

    let rejected = {
        let mut slide = presentation.slide_mut(0).unwrap();
        matches!(
            slide.add_table(2, 2, Emu(0), Emu(0), Emu(-1), Emu(200)),
            Err(Error::InvalidTableMutation { .. })
        )
    };
    assert!(rejected);
    assert_eq!(presentation.to_bytes().unwrap(), before);

    presentation
        .slide_mut(0)
        .unwrap()
        .add_table(2, 3, Emu(0), Emu(0), Emu(300), Emu(200))
        .unwrap();
    let table_index = presentation.slide(0).unwrap().shapes().len() - 1;
    let before_invalid_cell = presentation.to_bytes().unwrap();
    {
        let mut slide = presentation.slide_mut(0).unwrap();
        let mut shape = slide.shape_mut(table_index).unwrap();
        let mut table = shape.table_mut().unwrap();
        assert!(table.cell_mut(2, 0).is_none());
        assert!(table.cell_mut(0, 3).is_none());
        assert!(table.cell_mut(0, 0).unwrap().merge_to(2, 0).is_err());
        assert!(table.cell_mut(0, 0).unwrap().merge_to(0, 0).is_err());
    }
    assert_eq!(presentation.to_bytes().unwrap(), before_invalid_cell);

    {
        let mut slide = presentation.slide_mut(0).unwrap();
        let mut shape = slide.shape_mut(table_index).unwrap();
        let mut table = shape.table_mut().unwrap();
        assert!(table.cell_mut(0, 0).unwrap().split().is_err());
        assert!(table.set_column_width(0, Emu(0)).is_err());
        assert!(table.set_column_width(0, Emu(i64::MAX)).is_err());
    }
    assert_eq!(presentation.to_bytes().unwrap(), before_invalid_cell);

    {
        let mut slide = presentation.slide_mut(0).unwrap();
        let mut shape = slide.shape_mut(table_index).unwrap();
        shape
            .table_mut()
            .unwrap()
            .cell_mut(0, 0)
            .unwrap()
            .merge_to(1, 1)
            .unwrap();
    }
    let before_overlap = presentation.to_bytes().unwrap();
    {
        let mut slide = presentation.slide_mut(0).unwrap();
        let mut shape = slide.shape_mut(table_index).unwrap();
        let result = shape
            .table_mut()
            .unwrap()
            .cell_mut(0, 1)
            .unwrap()
            .merge_to(1, 2);
        assert!(matches!(result, Err(Error::InvalidTableMutation { .. })));
    }
    assert_eq!(presentation.to_bytes().unwrap(), before_overlap);
}

#[test]
fn table_mutation_preserves_unmodelled_xml_and_schema_order() {
    let fixture = table_mutation_fixture_bytes();
    let original_package = open_opc(&fixture, "F-113 original table preservation");
    let original_xml = original_package.get_part(SLIDE_TWO_PART).unwrap();
    let original_raw = capture_exact_subtree(original_xml, b"<x:raw-payload", b"</x:raw-payload>");
    let mut presentation = Presentation::from_bytes(&fixture).unwrap();
    {
        let mut slide = presentation.slide_mut(0).unwrap();
        let mut table_shape = slide.shape_mut(2).unwrap();
        let mut table = table_shape.table_mut().unwrap();
        table.set_column_width(0, Emu(125)).unwrap();
    }
    let saved = presentation.to_bytes().unwrap();
    let package = open_opc(&saved, "F-113 table preservation");
    let saved_xml = package.get_part(SLIDE_TWO_PART).unwrap();
    let saved_raw = capture_exact_subtree(saved_xml, b"<x:raw-payload", b"</x:raw-payload>");
    assert_eq!(saved_raw, original_raw, "unsupported table XML bytes");
    let xml = String::from_utf8(saved_xml.to_vec()).unwrap();
    for marker in [
        r#"producer="kept""#,
        r#"<x:before-table/>"#,
        r#"x:grid="kept""#,
        r#"<x:before-column/>"#,
        r#"x:column="kept""#,
        r#"<x:column-child/>"#,
        r#"<x:after-grid/>"#,
        r#"x:row="kept""#,
        r#"<x:before-cell/>"#,
        r#"x:cell="kept""#,
        r#"<x:before-text/>"#,
        r#"x:cell-properties="kept""#,
        r#"<x:cell-properties-child/>"#,
    ] {
        assert!(xml.contains(marker), "missing preserved marker {marker}");
    }
    assert!(xml.contains(r#"<a:gridCol w="125" x:column="kept">"#));
    let table_start = xml.find("<a:tbl ").unwrap();
    let grid_start = xml[table_start..].find("<a:tblGrid").unwrap() + table_start;
    let row_start = xml[grid_start..].find("<a:tr ").unwrap() + grid_start;
    assert!(table_start < grid_start && grid_start < row_start);
    let cell_start = xml[row_start..].find("<a:tc ").unwrap() + row_start;
    let text_start = xml[cell_start..].find("<a:txBody>").unwrap() + cell_start;
    let properties_start = xml[cell_start..].find("<a:tcPr ").unwrap() + cell_start;
    assert!(cell_start < text_start && text_start < properties_start);
}

#[test]
#[ignore = "requires uv and pinned python-pptx 1.0.2"]
fn add_table_matches_pinned_python_pptx_table_semantics() {
    let mut presentation = Presentation::new().expect("open bundled template");
    presentation.add_slide(0).expect("add slide");
    for (index, split) in [false, true].into_iter().enumerate() {
        let mut slide = presentation.slide_mut(0).unwrap();
        let mut shape = slide
            .add_table(
                2,
                3,
                Emu(10),
                Emu(20 + i64::try_from(index).unwrap() * 300),
                Emu(302),
                Emu(201),
            )
            .expect("add table");
        let mut table = shape.table_mut().unwrap();
        for (cell_index, text) in ["A", "B", "C", "D", "E", "F"].into_iter().enumerate() {
            table
                .cell_mut(cell_index / 3, cell_index % 3)
                .unwrap()
                .set_text(text);
        }
        table.cell_mut(1, 1).unwrap().merge_to(0, 0).unwrap();
        if split {
            table.cell_mut(0, 0).unwrap().split().unwrap();
        }
    }

    let output_path = std::env::temp_dir().join(format!(
        "rpptx-f113-python-oracle-{}.pptx",
        std::process::id()
    ));
    fs::write(&output_path, presentation.to_bytes().unwrap()).unwrap();
    let script = r#"
import json
import sys
import pptx
from pptx import Presentation

assert pptx.__version__ == "1.0.2", pptx.__version__

def normalize(shape):
    table = shape.table
    return {
        "frame": [shape.left, shape.top, shape.width, shape.height],
        "columns": [column.width for column in table.columns],
        "rows": [row.height for row in table.rows],
        "flags": [
            table.first_row,
            table.last_row,
            table.first_col,
            table.last_col,
            table.horz_banding,
            table.vert_banding,
        ],
        "cells": [[
            cell.text,
            cell.is_merge_origin,
            cell.is_spanned,
            cell.span_height,
            cell.span_width,
        ] for row in table.rows for cell in row.cells],
    }

ours = Presentation(sys.argv[1])
ours_records = [normalize(shape) for shape in list(ours.slides[0].shapes)[-2:]]

oracle = Presentation()
slide = oracle.slides.add_slide(oracle.slide_layouts[6])
for index, split in enumerate((False, True)):
    shape = slide.shapes.add_table(2, 3, 10, 20 + index * 300, 302, 201)
    table = shape.table
    for cell, text in zip((cell for row in table.rows for cell in row.cells), "ABCDEF"):
        cell.text = text
    table.cell(0, 0).merge(table.cell(1, 1))
    if split:
        table.cell(0, 0).split()
oracle_records = [normalize(shape) for shape in list(slide.shapes)[-2:]]

print(json.dumps(ours_records, sort_keys=True))
print(json.dumps(oracle_records, sort_keys=True))
"#;
    let output = Command::new("uv")
        .args([
            "run",
            "--with",
            "python-pptx==1.0.2",
            "python",
            "-c",
            script,
        ])
        .arg(&output_path)
        .output()
        .expect("run pinned python-pptx table oracle");
    fs::remove_file(&output_path).unwrap();
    assert!(
        output.status.success(),
        "python-pptx oracle failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let records = String::from_utf8(output.stdout).unwrap();
    let mut lines = records.lines();
    let ours = lines.next().unwrap();
    let oracle = lines.next().unwrap();
    assert_eq!(ours, oracle, "normalized table semantics");
    assert!(lines.next().is_none());
}

fn png_header(width: u32, height: u32) -> Vec<u8> {
    let mut bytes = b"\x89PNG\r\n\x1a\n\0\0\0\rIHDR".to_vec();
    bytes.extend_from_slice(&width.to_be_bytes());
    bytes.extend_from_slice(&height.to_be_bytes());
    bytes.extend_from_slice(&[8, 6, 0, 0, 0]);
    bytes.extend_from_slice(&[0; 4]);
    bytes
}

fn valid_one_pixel_png() -> Vec<u8> {
    vec![
        0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44,
        0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00, 0x90,
        0x77, 0x53, 0xde, 0x00, 0x00, 0x00, 0x0c, 0x49, 0x44, 0x41, 0x54, 0x78, 0xda, 0x63, 0xf8,
        0xcf, 0xc0, 0x00, 0x00, 0x03, 0x01, 0x01, 0x00, 0xf7, 0x03, 0x41, 0x43, 0x00, 0x00, 0x00,
        0x00, 0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
    ]
}

fn png_inflation_bomb() -> Vec<u8> {
    let inflated = vec![0; 1024 * 1024];
    let compressed = miniz_oxide::deflate::compress_to_vec_zlib(&inflated, 6);
    let mut png = png_header(1, 1);
    png.extend_from_slice(&(compressed.len() as u32).to_be_bytes());
    png.extend_from_slice(b"IDAT");
    png.extend_from_slice(&compressed);
    png.extend_from_slice(&[0; 4]);
    png.extend_from_slice(&0_u32.to_be_bytes());
    png.extend_from_slice(b"IEND");
    png.extend_from_slice(&[0; 4]);
    png
}

fn sparse_preview_png() -> Vec<u8> {
    let width = 9_u32;
    let height = 9_u32;
    let mut pixels = Vec::with_capacity((width * height * 4 + height) as usize);
    for row in 0..height {
        pixels.push(0);
        for column in 0..width {
            if row == 0 && column == 0 {
                pixels.extend_from_slice(&[255, 0, 0, 255]);
            } else {
                pixels.extend_from_slice(&[0, 0, 0, 0]);
            }
        }
    }
    let mut ihdr = Vec::with_capacity(13);
    ihdr.extend_from_slice(&width.to_be_bytes());
    ihdr.extend_from_slice(&height.to_be_bytes());
    ihdr.extend_from_slice(&[8, 6, 0, 0, 0]);
    let compressed = miniz_oxide::deflate::compress_to_vec_zlib(&pixels, 6);
    let mut png = b"\x89PNG\r\n\x1a\n".to_vec();
    append_png_chunk(&mut png, b"IHDR", &ihdr);
    append_png_chunk(&mut png, b"IDAT", &compressed);
    append_png_chunk(&mut png, b"IEND", &[]);
    png
}

fn append_png_chunk(png: &mut Vec<u8>, kind: &[u8; 4], data: &[u8]) {
    png.extend_from_slice(&(data.len() as u32).to_be_bytes());
    png.extend_from_slice(kind);
    png.extend_from_slice(data);
    let mut crc = 0xffff_ffff_u32;
    for byte in kind.iter().chain(data) {
        crc ^= u32::from(*byte);
        for _ in 0..8 {
            crc = (crc >> 1) ^ (0xedb8_8320 & 0_u32.wrapping_sub(crc & 1));
        }
    }
    png.extend_from_slice(&(!crc).to_be_bytes());
}

fn valid_template_jpeg() -> Vec<u8> {
    let bytes = Presentation::new().unwrap().to_bytes().unwrap();
    open_opc(&bytes, "bundled template JPEG")
        .get_part("/docProps/thumbnail.jpeg")
        .unwrap()
        .to_vec()
}

fn valid_grayscale_jpeg() -> Vec<u8> {
    decode_base64_fixture(
        "/9j/4AAQSkZJRgABAQAASABIAAD/4QNERXhpZgAATU0AKgAAAAgACQEPAAIAAAAIAAAAegEQAAIAAAAOAAAAggEaAAUAAAABAAAAkAEbAAUAAAABAAAAmAEoAAMAAAABAAIAAAExAAIAAAAMAAAAoAEyAAIAAAAUAAAArIdpAAQAAAABAAAAwIglAAQAAAABAAACegAAAABPbmVQbHVzAE9ORVBMVVMgQTUwMTAAAAAASAAAAAEAAABIAAAAAUdJTVAgMi44LjIyADIwMTk6MTE6MjYgMjI6Mzc6NDQAABuCmgAFAAAAAQAAAgqCnQAFAAAAAQAAAhKIIgADAAAAAQAAAACIJwADAAAAAQBkAACQAAAHAAAABDAyMjCQAwACAAAAFAAAAhqQBAACAAAAFAAAAi6RAQAHAAAABAECAwCSAQAKAAAAAQAAAkKSAgAFAAAAAQAAAkqSAwAKAAAAAQAAAlKSBwADAAAAAQAAAACSCQADAAAAAQAQAACSCgAFAAAAAQAAAlqSkAACAAAABwAAAmKSkQACAAAABwAAAmqSkgACAAAABwAAAnKgAAAHAAAABDAxMDCgAQADAAAAAf//AACgAgAEAAAAAQAAAAGgAwAEAAAAAQAAAAGiFwADAAAAAQACAACjAQAHAAAAAQEAAACkAgADAAAAAQAAAACkAwADAAAAAQAAAACkBQADAAAAAQAYAACkBgADAAAAAQAAAAAAAAAAAAAAAQAABgsAAAARAAAACjIwMTk6MTE6MDkgMTA6NTk6MzYAMjAxOToxMTowOSAxMDo1OTozNgAAAAhHAAAAyAAAAJkAAABkAAAAJQAAAAUAABAHAAAD6DYwMjY5NQAANjAyNjk1AAA2MDI2OTUAAAAIAAEAAgAAAAJOAAAAAAIABQAAAAMAAALgAAMAAgAAAAJFAAAAAAQABQAAAAMAAAL4AAUAAQAAAAEAAAAAAAYABQAAAAEAAAMQAAcABQAAAAMAAAMYAB0AAgAAAAsAAAMwAAAAAAAAADAAAAABAAAAMwAAAAEAAAZnAAAAZAAAAAIAAAABAAAAEgAAAAEAAABoAAAAZAABK6cAAAPoAAAACQAAAAEAAAA7AAAAAQAAACQAAAABMjAxOToxMTowOQAA/+0AeFBob3Rvc2hvcCAzLjAAOEJJTQQEAAAAAAA/HAFaAAMbJUccAgAAAgACHAI/AAYxMDU5MzYcAj4ACDIwMTkxMTA5HAI3AAgyMDE5MTEwORwCPAAGMTA1OTM2ADhCSU0EJQAAAAAAEG0diLVAz066cPvLhCa6YLX/wAALCAABAAEBAREA/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/9sAQwACAgICAgIDAgIDBQMDAwUGBQUFBQYIBgYGBgYICggICAgICAoKCgoKCgoKDAwMDAwMDg4ODg4PDw8PDw8PDw8P/90ABAAB/9oACAEBAAA/APTK/9k=",
    )
}

fn valid_cmyk_jpeg() -> Vec<u8> {
    decode_base64_fixture(
        "/9j/4AAQSkZJRgABAQAASABIAAD/4QEGRXhpZgAATU0AKgAAAAgACAEGAAMAAAABAAIAAAESAAMAAAABAAEAAAEaAAUAAAABAAAAbgEbAAUAAAABAAAAdgEoAAMAAAABAAIAAAExAAIAAAAhAAAAfgEyAAIAAAAUAAAAoIdpAAQAAAABAAAAtAAAAAAAAABIAAAAAQAAAEgAAAABQWRvYmUgUGhvdG9zaG9wIDI0LjYgKE1hY2ludG9zaCkAADIwMjQ6MDM6MjMgMDk6NTY6MDMAAASQAAAHAAAABDAyMzGQBAACAAAAFAAAAOqgAgAEAAAAAQAAAAGgAwAEAAAAAQAAAAEAAAAAMjAyMjowNDoxMiAxMjowNTo1MQD/7QBkUGhvdG9zaG9wIDMuMAA4QklNBAQAAAAAACwcAVoAAxslRxwCAAACAAIcAj4ACDIwMjIwNDEyHAI/AAsxMjA1NTErMDIwMDhCSU0EJQAAAAAAEOtjiDMDJLv+fLJTuujhSZT/7gAOQWRvYmUAZAAAAAAA/8AAFAgAAQABBAERAAIRAQMRAQQRAf/EAB8AAAEFAQEBAQEBAAAAAAAAAAABAgMEBQYHCAkKC//EALUQAAIBAwMCBAMFBQQEAAABfQECAwAEEQUSITFBBhNRYQcicRQygZGhCCNCscEVUtHwJDNicoIJChYXGBkaJSYnKCkqNDU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6g4SFhoeIiYqSk5SVlpeYmZqio6Slpqeoqaqys7S1tre4ubrCw8TFxsfIycrS09TV1tfY2drh4uPk5ebn6Onq8fLz9PX29/j5+v/EAB8BAAMBAQEBAQEBAQEAAAAAAAABAgMEBQYHCAkKC//EALURAAIBAgQEAwQHBQQEAAECdwABAgMRBAUhMQYSQVEHYXETIjKBCBRCkaGxwQkjM1LwFWJy0QoWJDThJfEXGBkaJicoKSo1Njc4OTpDREVGR0hJSlNUVVZXWFlaY2RlZmdoaWpzdHV2d3h5eoKDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uLj5OXm5+jp6vLz9PX29/j5+v/bAEMAAgICAgICAwICAwUDAwMFBgUFBQUGCAYGBgYGCAoICAgICAgKCgoKCgoKCgwMDAwMDA4ODg4ODw8PDw8PDw8PD//bAEMBAgMDBAQEBwQEBxALCQsQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEP/dAAQAAf/aAA4EAQACEQMRBBEAPwDj6/qA/wA7z+/D/9k=",
    )
}

fn decode_base64_fixture(encoded: &str) -> Vec<u8> {
    let mut output = Vec::with_capacity(encoded.len() * 3 / 4);
    let mut accumulator = 0_u32;
    let mut bits = 0_u32;
    for byte in encoded.bytes().take_while(|byte| *byte != b'=') {
        let value = match byte {
            b'A'..=b'Z' => u32::from(byte - b'A'),
            b'a'..=b'z' => u32::from(byte - b'a') + 26,
            b'0'..=b'9' => u32::from(byte - b'0') + 52,
            b'+' => 62,
            b'/' => 63,
            _ => panic!("invalid fixture base64"),
        };
        accumulator = (accumulator << 6) | value;
        bits += 6;
        if bits >= 8 {
            bits -= 8;
            output.push((accumulator >> bits) as u8);
            accumulator &= (1_u32 << bits).wrapping_sub(1);
        }
    }
    output
}

#[test]
fn setting_text_on_placeholder_round_trips_and_renders() {
    let mut presentation = Presentation::new().expect("open bundled template");
    presentation.add_slide(0).expect("add title slide");
    let placeholder_index = presentation
        .slide(0)
        .unwrap()
        .shapes()
        .position(|shape| shape.text().is_some())
        .expect("layout clone supplies a text placeholder");
    presentation
        .slide_mut(0)
        .unwrap()
        .shape_mut(placeholder_index)
        .unwrap()
        .set_text("")
        .unwrap();
    let blank_render = deterministic_render(&presentation.to_bytes().unwrap());
    presentation
        .slide_mut(0)
        .unwrap()
        .shape_mut(placeholder_index)
        .unwrap()
        .set_text("Visible changed placeholder")
        .unwrap();

    let bytes = presentation
        .to_bytes()
        .expect("serialize changed placeholder");
    let reopened = Presentation::from_bytes(&bytes).expect("reopen changed placeholder");
    assert_eq!(
        reopened
            .slide(0)
            .unwrap()
            .shape(placeholder_index)
            .unwrap()
            .text(),
        Some("Visible changed placeholder".to_owned())
    );
    let changed_render = deterministic_render(&bytes);
    assert_ne!(
        changed_render, blank_render,
        "changed text must alter pixels"
    );
}

#[test]
fn clearing_text_preserves_required_paragraph() {
    let mut presentation = Presentation::from_bytes(&mutation_fixture_bytes()).unwrap();
    presentation
        .slide_mut(0)
        .unwrap()
        .shape_mut(0)
        .unwrap()
        .set_text("")
        .unwrap();
    let bytes = presentation.to_bytes().unwrap();
    let reopened = Presentation::from_bytes(&bytes).unwrap();
    assert!(reopened.validate().is_empty());
    let package = open_opc(&bytes, "F-112 empty text");
    let slide = CT_Slide::from_xml(package.get_part(SLIDE_TWO_PART).unwrap()).unwrap();
    let ShapeTreeChild::Shape(shape) = &slide.common_slide_data.shape_tree.children[0] else {
        panic!("expected ordinary shape");
    };
    assert_eq!(shape.text_body.as_ref().unwrap().paragraph_count(), 1);
}

#[test]
fn paragraph_run_font_and_bullet_properties_round_trip() {
    let mut presentation = Presentation::from_bytes(&mutation_fixture_bytes()).unwrap();
    let mut slide = presentation.slide_mut(0).unwrap();
    let mut shape = slide.shape_mut(0).unwrap();
    shape.set_text("first").unwrap();
    let mut frame = shape.text_frame().unwrap();
    let mut paragraph = frame.paragraph_mut(0).unwrap();
    let mut paragraph_properties = CT_TextParagraphProperties::default();
    paragraph_properties.level = Some(2);
    paragraph.set_properties(paragraph_properties);
    paragraph.set_bullet(Some(TextBullet {
        choice: Some(TextBulletChoice::Character(
            TextBulletCharacter::new("•").unwrap(),
        )),
        ..TextBullet::default()
    }));
    let mut run = paragraph.add_run(" second");
    let mut run_properties = CT_TextCharacterProperties::default();
    run_properties.font_size = Some(1_800);
    run_properties.bold = Some(true);
    run.set_properties(run_properties);
    run.set_font(Some(TextFont::new("Carlito").unwrap()));

    let bytes = presentation.to_bytes().unwrap();
    let package = open_opc(&bytes, "F-112 formatting round trip");
    let slide = CT_Slide::from_xml(package.get_part(SLIDE_TWO_PART).unwrap()).unwrap();
    let ShapeTreeChild::Shape(shape) = &slide.common_slide_data.shape_tree.children[0] else {
        panic!("expected ordinary shape");
    };
    let paragraph = &shape.text_body.as_ref().unwrap().paragraphs()[0];
    let properties = paragraph.properties.as_ref().unwrap();
    assert_eq!(properties.level, Some(2));
    assert!(matches!(
        properties.bullet.as_ref().and_then(|bullet| bullet.choice.as_ref()),
        Some(TextBulletChoice::Character(character)) if character.character == "•"
    ));
    let oxml_drawing::text::TextRun::Run(run) = &paragraph.runs[1] else {
        panic!("expected appended regular run");
    };
    let properties = run.properties.as_ref().unwrap();
    assert_eq!(properties.font_size, Some(1_800));
    assert_eq!(properties.bold, Some(true));
    assert_eq!(properties.latin.as_ref().unwrap().typeface, "Carlito");
}

#[test]
fn text_mutation_preserves_placeholder_identity() {
    let mut package = open_opc(&mutation_fixture_bytes(), "F-112 placeholder fixture");
    let original = String::from_utf8(package.get_part(SLIDE_TWO_PART).unwrap().to_vec()).unwrap();
    let with_placeholder = original.replacen(
        "<p:nvPr/>",
        r#"<p:nvPr><p:ph type="body" idx="7"/></p:nvPr>"#,
        1,
    );
    package.set_part(SLIDE_TWO_PART, with_placeholder.into_bytes());
    let mut presentation = Presentation::from_bytes(&package_bytes(package)).unwrap();
    presentation
        .slide_mut(0)
        .unwrap()
        .shape_mut(0)
        .unwrap()
        .set_text("replacement")
        .unwrap();

    let output = open_opc(
        &presentation.to_bytes().unwrap(),
        "F-112 placeholder output",
    );
    let slide = CT_Slide::from_xml(output.get_part(SLIDE_TWO_PART).unwrap()).unwrap();
    let ShapeTreeChild::Shape(shape) = &slide.common_slide_data.shape_tree.children[0] else {
        panic!("expected ordinary shape");
    };
    let placeholder = shape.placeholder.as_ref().unwrap();
    assert_eq!(placeholder.ph_type, Some(PhType::Body));
    assert_eq!(placeholder.idx, Some(7));
}

#[test]
fn text_mutation_preserves_unmodelled_xml_and_schema_order() {
    let mut package = open_opc(&mutation_fixture_bytes(), "F-112 raw fixture");
    let original = String::from_utf8(package.get_part(SLIDE_TWO_PART).unwrap().to_vec()).unwrap();
    let enriched = original.replacen(
        "<a:bodyPr/>",
        r#"<x:before xmlns:x="urn:f112"/><a:bodyPr/><x:after-body xmlns:x="urn:f112"/>"#,
        1,
    ).replacen(
        "<a:r><a:t>plain</a:t></a:r>",
        r#"<mc:AlternateContent xmlns:x="urn:f112"><mc:Choice Requires="x"><x:producer-marker value="kept"/></mc:Choice><mc:Fallback><a:r><a:t>fallback</a:t></a:r></mc:Fallback></mc:AlternateContent><a:fld id="{00000000-0000-0000-0000-000000000001}"><a:t>field</a:t></a:fld><a:br/><x:after-runs xmlns:x="urn:f112"/>"#,
        1,
    );
    package.set_part(SLIDE_TWO_PART, enriched.clone().into_bytes());
    let mut formatted = Presentation::from_bytes(&package_bytes(package)).unwrap();
    let mut slide = formatted.slide_mut(0).unwrap();
    let mut shape = slide.shape_mut(0).unwrap();
    let mut frame = shape.text_frame().unwrap();
    let mut paragraph = frame.paragraph_mut(0).unwrap();
    let mut properties = CT_TextParagraphProperties::default();
    properties.level = Some(1);
    paragraph.set_properties(properties);
    let formatted_output = open_opc(&formatted.to_bytes().unwrap(), "F-112 formatted raw output");
    let formatted_xml =
        String::from_utf8(formatted_output.get_part(SLIDE_TWO_PART).unwrap().to_vec()).unwrap();
    let marker = r#"<mc:AlternateContent xmlns:x="urn:f112"><mc:Choice Requires="x"><x:producer-marker value="kept"/></mc:Choice><mc:Fallback><a:r><a:t>fallback</a:t></a:r></mc:Fallback></mc:AlternateContent>"#;
    let field =
        r#"<a:fld id="{00000000-0000-0000-0000-000000000001}"><a:t>field</a:t></a:fld><a:br/>"#;
    let properties_index = formatted_xml.find("<a:pPr lvl=\"1\"").unwrap();
    let marker_index = formatted_xml.find(marker).unwrap();
    let field_index = formatted_xml.find(field).unwrap();
    assert!(properties_index < marker_index);
    assert!(marker_index < field_index);
    assert!(formatted_xml.contains(field));
    let reparsed_xml = String::from_utf8(
        CT_Slide::from_xml(formatted_output.get_part(SLIDE_TWO_PART).unwrap())
            .unwrap()
            .to_xml()
            .unwrap(),
    )
    .unwrap();
    let properties_index = reparsed_xml.find("<a:pPr lvl=\"1\"").unwrap();
    let marker_index = reparsed_xml.find(marker).unwrap();
    let field_index = reparsed_xml.find(field).unwrap();
    assert!(properties_index < marker_index);
    assert!(marker_index < field_index);

    let mut package = open_opc(&mutation_fixture_bytes(), "F-112 raw replacement fixture");
    package.set_part(SLIDE_TWO_PART, enriched.into_bytes());
    let mut presentation = Presentation::from_bytes(&package_bytes(package)).unwrap();
    presentation
        .slide_mut(0)
        .unwrap()
        .shape_mut(0)
        .unwrap()
        .set_text("ordered")
        .unwrap();

    let output = open_opc(&presentation.to_bytes().unwrap(), "F-112 raw output");
    let xml = String::from_utf8(output.get_part(SLIDE_TWO_PART).unwrap().to_vec()).unwrap();
    for marker in [
        "<x:before ",
        "<x:after-body ",
        "<x:producer-marker ",
        "<x:after-runs ",
    ] {
        assert!(xml.contains(marker), "missing preserved marker {marker}");
    }
    assert!(xml.find("<a:bodyPr").unwrap() < xml.find("<a:p>").unwrap());
    assert!(xml.find("<a:r>").unwrap() < xml.find("<a:t>ordered</a:t>").unwrap());
}

#[test]
fn text_mutation_indices_and_shape_kinds_are_total() {
    let mut presentation = Presentation::from_bytes(&mutation_fixture_bytes()).unwrap();
    let mut slide = presentation.slide_mut(0).unwrap();
    assert!(slide.shape_mut(999).is_none());
    let mut shape = slide.shape_mut(0).unwrap();
    {
        let mut frame = shape.text_frame().unwrap();
        assert!(frame.paragraph_mut(999).is_none());
    }
    let mut picture = slide.shape_mut(1).unwrap();
    assert!(picture.text_frame().is_none());
    assert!(matches!(
        picture.set_text("unsupported"),
        Err(Error::UnsupportedShapeMutation { .. })
    ));
}

#[test]
fn text_frame_handles_append_paragraphs_and_runs_in_order() {
    let mut presentation = Presentation::from_bytes(&mutation_fixture_bytes()).unwrap();
    let mut slide = presentation.slide_mut(0).unwrap();
    let mut shape = slide.shape_mut(0).unwrap();
    shape.set_text("one").unwrap();
    let mut frame = shape.text_frame().unwrap();
    frame.paragraph_mut(0).unwrap().add_run(" two");
    frame.add_paragraph().set_text("three");
    assert_eq!(frame.text(), "one two\nthree");

    let reopened = Presentation::from_bytes(&presentation.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened.slide(0).unwrap().shape(0).unwrap().text(),
        Some("one two\nthree".to_owned())
    );
}

#[test]
fn picture_without_explicit_size_uses_native_dimensions() {
    let png = png_header(32, 16);
    let mut presentation = Presentation::new().expect("open bundled template");
    presentation.add_slide(0).expect("add slide");
    let original_count = presentation.slide(0).unwrap().shapes().len();

    let picture = presentation
        .add_picture(0, &png, "image.png", Emu(10), Emu(20), None, None)
        .expect("add native-size picture");
    assert_eq!(picture.kind(), ShapeKind::Picture);

    let bytes = presentation.to_bytes().expect("serialize picture");
    let reopened = Presentation::from_bytes(&bytes).expect("reopen picture");
    assert!(reopened.validate().is_empty());
    let package = open_opc(&bytes, "F-111 native picture");
    let slide = CT_Slide::from_xml(package.get_part("/ppt/slides/slide1.xml").unwrap()).unwrap();
    let ShapeTreeChild::Picture(picture) =
        &slide.common_slide_data.shape_tree.children[original_count]
    else {
        panic!("expected appended picture");
    };
    let transform = picture.shape_properties.transform.as_ref().unwrap();
    assert_eq!(transform.offset.unwrap().x, Emu(10));
    assert_eq!(transform.offset.unwrap().y, Emu(20));
    assert_eq!(transform.extent.unwrap().cx, Emu(406_400));
    assert_eq!(transform.extent.unwrap().cy, Emu(203_200));
}

#[test]
fn picture_one_dimension_preserves_aspect_ratio_with_truncation() {
    let png = png_header(3, 2);
    let mut presentation = Presentation::new().expect("open bundled template");
    presentation.add_slide(0).expect("add slide");
    let original_count = presentation.slide(0).unwrap().shapes().len();
    presentation
        .add_picture(0, &png, "ratio.png", Emu(0), Emu(0), Some(Emu(10)), None)
        .unwrap();
    presentation
        .add_picture(0, &png, "ratio.png", Emu(0), Emu(0), None, Some(Emu(10)))
        .unwrap();

    let package = open_opc(&presentation.to_bytes().unwrap(), "F-111 aspect ratio");
    assert_eq!(
        package
            .get_part_rels("/ppt/slides/slide1.xml")
            .unwrap()
            .get_all_by_type(rel_types::IMAGE)
            .len(),
        1,
        "equal bytes on one slide reuse the slide-scoped relationship"
    );
    let slide = CT_Slide::from_xml(package.get_part("/ppt/slides/slide1.xml").unwrap()).unwrap();
    let pictures = &slide.common_slide_data.shape_tree.children[original_count..];
    let ShapeTreeChild::Picture(width_only) = &pictures[0] else {
        panic!("expected width-only picture");
    };
    let ShapeTreeChild::Picture(height_only) = &pictures[1] else {
        panic!("expected height-only picture");
    };
    let width_only = width_only.shape_properties.transform.as_ref().unwrap();
    let height_only = height_only.shape_properties.transform.as_ref().unwrap();
    assert_eq!(width_only.extent.unwrap().cx, Emu(10));
    assert_eq!(width_only.extent.unwrap().cy, Emu(6));
    assert_eq!(height_only.extent.unwrap().cx, Emu(15));
    assert_eq!(height_only.extent.unwrap().cy, Emu(10));
}

#[test]
fn duplicate_picture_bytes_share_one_media_part_across_slides() {
    let png = png_header(4, 3);
    let mut presentation = Presentation::new().expect("open bundled template");
    presentation.add_slide(0).expect("add first slide");
    presentation.add_slide(0).expect("add second slide");
    for slide_index in 0..2 {
        presentation
            .add_picture(slide_index, &png, "shared.png", Emu(0), Emu(0), None, None)
            .unwrap();
    }

    let bytes = presentation.to_bytes().unwrap();
    let reopened = Presentation::from_bytes(&bytes).unwrap();
    assert!(reopened.validate().is_empty());
    let package = open_opc(&bytes, "F-111 deduplicated media");
    let media = package
        .parts
        .iter()
        .filter(|(part_name, bytes)| part_name.starts_with("/ppt/media/") && *bytes == &png)
        .map(|(part_name, _)| part_name.clone())
        .collect::<Vec<_>>();
    assert_eq!(media.len(), 1);
    for slide_part in ["/ppt/slides/slide1.xml", "/ppt/slides/slide2.xml"] {
        let relationships = package.get_part_rels(slide_part).unwrap();
        let images = relationships.get_all_by_type(rel_types::IMAGE);
        assert_eq!(images.len(), 1);
        assert_eq!(
            OpcPackage::resolve_rel_target(slide_part, &images[0].target),
            media[0]
        );
    }
}

#[test]
fn picture_sniffs_bytes_when_extension_is_misleading() {
    let png = png_header(5, 4);
    let mut presentation = Presentation::new().expect("open bundled template");
    presentation.add_slide(0).expect("add slide");
    presentation
        .add_picture(0, &png, "misleading.jpeg", Emu(0), Emu(0), None, None)
        .unwrap();

    let package = open_opc(&presentation.to_bytes().unwrap(), "F-111 sniffed media");
    let part = package
        .parts
        .iter()
        .find(|(part_name, bytes)| part_name.starts_with("/ppt/media/") && *bytes == &png)
        .map(|(part_name, _)| part_name)
        .unwrap();
    assert!(part.ends_with(".png"));
    assert_eq!(
        package.content_types.content_type_for(part),
        Some("image/png")
    );
}

#[test]
fn picture_with_both_dimensions_does_not_require_native_metadata() {
    let svg = br#"<svg xmlns="http://www.w3.org/2000/svg" width="2" height="1"/>"#;
    let mut presentation = Presentation::new().expect("open bundled template");
    presentation.add_slide(0).expect("add slide");
    let original_count = presentation.slide(0).unwrap().shapes().len();
    presentation
        .add_picture(
            0,
            svg,
            "vector.svg",
            Emu(1),
            Emu(2),
            Some(Emu(300)),
            Some(Emu(400)),
        )
        .unwrap();

    let package = open_opc(&presentation.to_bytes().unwrap(), "F-111 explicit size");
    let slide = CT_Slide::from_xml(package.get_part("/ppt/slides/slide1.xml").unwrap()).unwrap();
    let ShapeTreeChild::Picture(picture) =
        &slide.common_slide_data.shape_tree.children[original_count]
    else {
        panic!("expected explicit-size picture");
    };
    let extent = picture
        .shape_properties
        .transform
        .as_ref()
        .unwrap()
        .extent
        .unwrap();
    assert_eq!(extent.cx, Emu(300));
    assert_eq!(extent.cy, Emu(400));
    assert!(package.parts.keys().any(|part| part.ends_with(".svg")));
}

#[test]
fn picture_append_preserves_schema_final_content_and_raw_reserved_ids() {
    const EXTENSION: &str = r#"<p:extLst><p:ext uri="{F110-APPEND-ORDER}"><x:marker xmlns:x="urn:f110" value="kept"/></p:ext></p:extLst>"#;
    let png = png_header(2, 1);
    let mut presentation = Presentation::from_bytes(&append_boundary_fixture_bytes()).unwrap();
    presentation
        .add_picture(0, &png, "ordered.png", Emu(1), Emu(2), None, None)
        .unwrap();

    let package = open_opc(&presentation.to_bytes().unwrap(), "F-111 append boundary");
    let xml = String::from_utf8(package.get_part(SLIDE_TWO_PART).unwrap().to_vec()).unwrap();
    assert!(xml.contains(EXTENSION));
    assert!(xml.rfind("<p:pic>").unwrap() < xml.find(EXTENSION).unwrap());
    let slide = CT_Slide::from_xml(xml.as_bytes()).unwrap();
    let appended = slide.common_slide_data.shape_tree.children.last().unwrap();
    assert_eq!(appended.non_visual_id(), Some(2));
    let ShapeTreeChild::Picture(_) = appended else {
        panic!("expected appended picture");
    };
}

#[test]
fn invalid_picture_input_does_not_mutate_the_presentation() {
    let png = png_header(5, 4);
    let mut presentation = Presentation::new().expect("open bundled template");
    presentation.add_slide(0).expect("add slide");
    let before = presentation.to_bytes().unwrap();

    assert!(matches!(
        presentation.add_picture(1, &png, "image.png", Emu(0), Emu(0), None, None),
        Err(Error::UnknownSlideIndex { index: 1, .. })
    ));
    assert_eq!(presentation.to_bytes().unwrap(), before);

    assert!(matches!(
        presentation.add_picture(0, b"not an image", "image.bin", Emu(0), Emu(0), None, None,),
        Err(Error::UnsupportedPicture { .. })
    ));
    assert_eq!(presentation.to_bytes().unwrap(), before);

    assert!(matches!(
        presentation.add_picture(
            0,
            br#"<svg xmlns="http://www.w3.org/2000/svg"/>"#,
            "image.svg",
            Emu(0),
            Emu(0),
            None,
            None,
        ),
        Err(Error::UnavailablePictureDimensions { .. })
    ));
    assert_eq!(presentation.to_bytes().unwrap(), before);

    assert!(matches!(
        presentation.add_picture(
            0,
            &png_header(1, u32::MAX),
            "too-tall.png",
            Emu(0),
            Emu(0),
            Some(Emu(i64::MAX)),
            None,
        ),
        Err(Error::PictureDimensionOutOfRange { axis: "height", .. })
    ));
    assert_eq!(presentation.to_bytes().unwrap(), before);
}

#[test]
#[ignore = "requires uv and pinned python-pptx 1.0.2"]
fn picture_native_size_matches_python_pptx_1_0_2() {
    let png = valid_one_pixel_png();
    let mut presentation = Presentation::new().expect("open bundled template");
    presentation.add_slide(0).expect("add slide");
    presentation
        .add_picture(0, &png, "pixel.png", Emu(10), Emu(20), None, None)
        .unwrap();
    let output_path = std::env::temp_dir().join(format!(
        "rpptx-f111-python-oracle-{}.pptx",
        std::process::id()
    ));
    fs::write(&output_path, presentation.to_bytes().unwrap()).unwrap();
    let image_hex = png
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect::<String>();
    let script = r#"
import io
import sys
from pptx import Presentation

ours = Presentation(sys.argv[1])
ours_picture = ours.slides[0].shapes[-1]
print(f"ours\t{ours_picture.shape_type.name}\t{ours_picture.width}\t{ours_picture.height}")

oracle = Presentation()
slide = oracle.slides.add_slide(oracle.slide_layouts[6])
oracle_picture = slide.shapes.add_picture(io.BytesIO(bytes.fromhex(sys.argv[2])), 10, 20)
print(f"oracle\t{oracle_picture.shape_type.name}\t{oracle_picture.width}\t{oracle_picture.height}")
"#;
    let output = Command::new("uv")
        .args([
            "run",
            "--with",
            "python-pptx==1.0.2",
            "python",
            "-c",
            script,
        ])
        .arg(&output_path)
        .arg(image_hex)
        .output()
        .expect("run pinned python-pptx picture oracle");
    fs::remove_file(&output_path).unwrap();
    assert!(
        output.status.success(),
        "python-pptx oracle failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    assert_eq!(
        String::from_utf8(output.stdout).unwrap(),
        "ours\tPICTURE\t12700\t12700\noracle\tPICTURE\t12700\t12700\n"
    );
}

#[test]
#[ignore = "requires pinned Microsoft PowerPoint"]
fn added_picture_validates_and_opens_without_repair() {
    assert_powerpoint_build();
    let png = valid_one_pixel_png();
    let mut presentation = Presentation::new().expect("open bundled template");
    presentation.add_slide(0).expect("add slide");
    presentation
        .add_picture(
            0,
            &png,
            "pixel.png",
            Emu(914_400),
            Emu(914_400),
            Some(Emu(914_400)),
            Some(Emu(914_400)),
        )
        .unwrap();
    assert!(presentation.validate().is_empty());
    let output =
        std::env::temp_dir().join(format!("rpptx-f111-picture-{}.pptx", std::process::id()));
    fs::write(&output, presentation.to_bytes().expect("serialize deck"))
        .expect("write native acceptance deck");
    let name = output.file_name().unwrap().to_string_lossy();
    let path = output.to_string_lossy();
    let script = format!(
        "with timeout of 120 seconds\ntell application \"Microsoft PowerPoint\"\nset previousStartUpDialog to start up dialog\ntry\nif (Version as text) is not \"{POWERPOINT_VERSION}\" then error \"PowerPoint version mismatch\"\nif (build as text) is not \"{POWERPOINT_APP_BUILD}\" then error \"PowerPoint build mismatch\"\nset start up dialog to false\nset deckPath to \"{path}\"\nopen my POSIX file deckPath\nset checkedDeck to presentation \"{name}\"\nif (count of slides of checkedDeck) is not 1 then error \"slide count mismatch\"\nclose checkedDeck saving no\nset start up dialog to previousStartUpDialog\non error errorMessage number errorNumber\ntry\nclose checkedDeck saving no\nend try\nset start up dialog to previousStartUpDialog\nerror errorMessage number errorNumber\nend try\nend tell\nend timeout\n"
    );
    let result = Command::new("osascript")
        .args(["-e", &script])
        .output()
        .expect("launch PowerPoint acceptance script");
    fs::remove_file(&output).expect("remove native acceptance deck");
    assert!(
        result.status.success(),
        "PowerPoint no-repair open failed: {}",
        String::from_utf8_lossy(&result.stderr)
    );
}
const MODELLED_CONTENT_TYPES: [&str; 7] = [
    content_types::PRESENTATION,
    content_types::SLIDE,
    content_types::SLIDE_LAYOUT,
    content_types::SLIDE_MASTER,
    content_types::NOTES_SLIDE,
    content_types::NOTES_MASTER,
    content_types::THEME,
];
const VISUAL_DECKS: [&str; 5] = [
    "WithMaster.pptx",
    "backgrounds.pptx",
    "placeholder-layout-color.pptx",
    "bug58144-headers-footers-2007.pptx",
    "60810.pptx",
];
const POWERPOINT_VISUAL_ACCEPTANCE: &str = r#"application	Microsoft PowerPoint 16.104	bundle 16.104.25121423
deck	/Users/atulsharma/Documents/projects/tensorbee/rdocx/corpus/pptx/WithMaster.pptx
deck	/Users/atulsharma/Documents/projects/tensorbee/rdocx/corpus/pptx/backgrounds.pptx
deck	/Users/atulsharma/Documents/projects/tensorbee/rdocx/corpus/pptx/placeholder-layout-color.pptx
deck	/Users/atulsharma/Documents/projects/tensorbee/rdocx/corpus/pptx/bug58144-headers-footers-2007.pptx
native_pdf	/private/tmp/F088-WithMaster.pdf
native_pdf	/private/tmp/F088-backgrounds.pdf
native_pdf	/private/tmp/F088-placeholder-layout-color.pdf
native_pdf	/private/tmp/F088-headers-footers.pdf
rendered_pages	/private/tmp/F088-pdf-render
normalized_evidence	/private/tmp/F-088-normalized-visual.txt
result	clean	opened and exported without repair, clipping, or prompt text
WithMaster	master artwork once, slide footer once on slide 2, no latent date or number
backgrounds	distinct solid, gradient, picture, and texture visuals
placeholder-layout-color	exact cyan on black, RGBA #00FFFF
bug58144-headers-footers-2007	Slide footer once
"#;

#[test]
fn four_appended_shapes_have_unique_ids_and_reopen() {
    let mut presentation = Presentation::new().expect("open bundled template");
    presentation.add_slide(0).expect("add slide");
    let original_count = presentation.slide(0).unwrap().shapes().len();
    let mut slide = presentation.slide_mut(0).unwrap();
    assert_eq!(
        slide
            .add_textbox(Emu(10), Emu(20), Emu(300), Emu(400))
            .unwrap()
            .kind(),
        ShapeKind::Shape
    );
    assert_eq!(
        slide
            .add_shape("triangle", Emu(30), Emu(40), Emu(500), Emu(600))
            .unwrap()
            .kind(),
        ShapeKind::Shape
    );
    assert_eq!(
        slide
            .add_connector(ConnectorType::Elbow, Emu(700), Emu(800), Emu(100), Emu(200),)
            .unwrap()
            .kind(),
        ShapeKind::Connector
    );
    assert_eq!(slide.add_group_shape().unwrap().kind(), ShapeKind::Group);

    let bytes = presentation
        .to_bytes()
        .expect("serialize constructed shapes");
    let reopened = Presentation::from_bytes(&bytes).expect("reopen constructed shapes");
    assert!(reopened.validate().is_empty());
    let reopened_slide = reopened.slide(0).unwrap();
    let shapes = reopened_slide.shapes().collect::<Vec<_>>();
    assert_eq!(shapes.len(), original_count + 4);
    assert_eq!(shapes[original_count].kind(), ShapeKind::Shape);
    assert_eq!(shapes[original_count + 1].kind(), ShapeKind::Shape);
    assert_eq!(shapes[original_count + 2].kind(), ShapeKind::Connector);
    assert_eq!(shapes[original_count + 3].kind(), ShapeKind::Group);

    let package = open_opc(&bytes, "F-110 constructor reopen");
    let slide = CT_Slide::from_xml(package.get_part("/ppt/slides/slide1.xml").unwrap()).unwrap();
    let appended = &slide.common_slide_data.shape_tree.children[original_count..];
    let ids = appended
        .iter()
        .map(|child| child.non_visual_id().unwrap())
        .collect::<HashSet<_>>();
    assert_eq!(ids.len(), 4);

    let ShapeTreeChild::Shape(textbox) = &appended[0] else {
        panic!("expected textbox shape");
    };
    let transform = textbox.shape_properties.transform.as_ref().unwrap();
    assert_eq!(transform.offset.unwrap().x, Emu(10));
    assert_eq!(transform.offset.unwrap().y, Emu(20));
    assert_eq!(transform.extent.unwrap().cx, Emu(300));
    assert_eq!(transform.extent.unwrap().cy, Emu(400));
    assert_eq!(
        textbox
            .shape_properties
            .preset_geometry
            .as_ref()
            .unwrap()
            .preset,
        "rect"
    );
    assert!(textbox.shape_properties.fill.is_some());
    let text_body = textbox.text_body.as_ref().unwrap();
    assert!(text_body.has_list_style());
    assert_eq!(text_body.paragraph_count(), 1);
    assert!(text_body.paragraphs()[0].runs.is_empty());
    let textbox_xml = String::from_utf8(textbox.to_xml().unwrap()).unwrap();
    assert!(textbox_xml.contains("<p:cNvSpPr txBox=\"1\"/>"));
    assert!(textbox_xml.contains("<a:noFill/>"));

    let ShapeTreeChild::Shape(shape) = &appended[1] else {
        panic!("expected ordinary shape");
    };
    let transform = shape.shape_properties.transform.as_ref().unwrap();
    assert_eq!(transform.offset.unwrap().x, Emu(30));
    assert_eq!(transform.offset.unwrap().y, Emu(40));
    assert_eq!(transform.extent.unwrap().cx, Emu(500));
    assert_eq!(transform.extent.unwrap().cy, Emu(600));
    assert_eq!(
        shape
            .shape_properties
            .preset_geometry
            .as_ref()
            .unwrap()
            .preset,
        "triangle"
    );
    assert!(shape.shape_properties.fill.is_none());
    let text_body = shape.text_body.as_ref().unwrap();
    assert!(text_body.has_list_style());
    assert_eq!(text_body.paragraph_count(), 1);
    assert!(text_body.paragraphs()[0].runs.is_empty());

    let ShapeTreeChild::Connector(connector) = &appended[2] else {
        panic!("expected connector");
    };
    let transform = connector.shape_properties.transform.as_ref().unwrap();
    assert_eq!(transform.offset.unwrap().x, Emu(100));
    assert_eq!(transform.offset.unwrap().y, Emu(200));
    assert_eq!(transform.extent.unwrap().cx, Emu(600));
    assert_eq!(transform.extent.unwrap().cy, Emu(600));
    assert!(transform.flip_horizontal);
    assert!(transform.flip_vertical);
    assert_eq!(
        connector
            .shape_properties
            .preset_geometry
            .as_ref()
            .unwrap()
            .preset,
        "bentConnector3"
    );

    let ShapeTreeChild::GroupShape(group) = &appended[3] else {
        panic!("expected empty group");
    };
    assert!(group.children.is_empty());
    assert!(group.group_transform().is_none());
    let group_xml = String::from_utf8(group.to_xml().unwrap()).unwrap();
    assert!(group_xml.contains("<p:cNvGrpSpPr/>"));
    assert!(group_xml.contains("<p:nvPr/>"));
    assert!(group_xml.contains("<p:grpSpPr/>"));
}

#[test]
fn unknown_preset_does_not_mutate_the_slide() {
    let mut presentation = Presentation::new().expect("open bundled template");
    presentation.add_slide(0).expect("add slide");
    let before = presentation.to_bytes().unwrap();

    let error = match presentation.slide_mut(0).unwrap().add_shape(
        "not-a-preset",
        Emu(1),
        Emu(2),
        Emu(3),
        Emu(4),
    ) {
        Ok(_) => panic!("unknown preset was accepted"),
        Err(error) => error,
    };

    assert!(matches!(
        error,
        Error::InvalidShapeMutation {
            operation: "add shape",
            ref message,
        } if message.contains("unknown DrawingML preset geometry: not-a-preset")
    ));
    assert_eq!(presentation.to_bytes().unwrap(), before);
}

#[test]
fn every_constructor_appends_before_preserved_shape_tree_extensions() {
    const EXTENSION: &str = r#"<p:extLst><p:ext uri="{F110-APPEND-ORDER}"><x:marker xmlns:x="urn:f110" value="kept"/></p:ext></p:extLst>"#;

    for constructor in 0..4 {
        let mut presentation = Presentation::from_bytes(&append_boundary_fixture_bytes()).unwrap();
        let original_count = presentation.slide(0).unwrap().shapes().len();
        let mut slide = presentation.slide_mut(0).unwrap();
        let expected_tag = match constructor {
            0 => {
                slide.add_textbox(Emu(1), Emu(2), Emu(3), Emu(4)).unwrap();
                "<p:sp>"
            }
            1 => {
                slide
                    .add_shape("triangle", Emu(1), Emu(2), Emu(3), Emu(4))
                    .unwrap();
                "<p:sp>"
            }
            2 => {
                slide
                    .add_connector(ConnectorType::Straight, Emu(1), Emu(2), Emu(3), Emu(4))
                    .unwrap();
                "<p:cxnSp>"
            }
            3 => {
                slide.add_group_shape().unwrap();
                "<p:grpSp>"
            }
            _ => unreachable!(),
        };

        let bytes = presentation.to_bytes().unwrap();
        let package = open_opc(&bytes, "F-110 append boundary");
        let xml = String::from_utf8(package.get_part(SLIDE_TWO_PART).unwrap().to_vec()).unwrap();
        assert!(xml.contains(EXTENSION));
        assert!(xml.rfind(expected_tag).unwrap() < xml.find(EXTENSION).unwrap());
        let reparsed = CT_Slide::from_xml(xml.as_bytes()).unwrap();
        assert_eq!(
            reparsed.common_slide_data.shape_tree.children.len(),
            original_count + 1
        );
    }
}

#[test]
fn constructor_names_are_deterministic_from_allocated_ids() {
    let mut presentation = Presentation::new().expect("open bundled template");
    presentation.add_slide(0).expect("add slide");
    let original_count = presentation.slide(0).unwrap().shapes().len();
    let mut slide = presentation.slide_mut(0).unwrap();
    slide.add_textbox(Emu(1), Emu(2), Emu(3), Emu(4)).unwrap();
    slide
        .add_shape("rect", Emu(5), Emu(6), Emu(7), Emu(8))
        .unwrap();
    slide
        .add_connector(ConnectorType::Straight, Emu(9), Emu(10), Emu(11), Emu(12))
        .unwrap();
    slide.add_group_shape().unwrap();

    let bytes = presentation.to_bytes().unwrap();
    let package = open_opc(&bytes, "F-110 deterministic names");
    let xml =
        String::from_utf8(package.get_part("/ppt/slides/slide1.xml").unwrap().to_vec()).unwrap();
    let slide = CT_Slide::from_xml(xml.as_bytes()).unwrap();
    let appended = &slide.common_slide_data.shape_tree.children[original_count..];
    for (child, prefix) in appended
        .iter()
        .zip(["TextBox", "Shape", "Connector", "Group"])
    {
        let id = child.non_visual_id().unwrap();
        assert!(xml.contains(&format!("name=\"{prefix} {id}\"")));
    }
}

#[test]
fn shape_mutation_setters_survive_save_and_reload() {
    let mut presentation = Presentation::from_bytes(&mutation_fixture_bytes()).unwrap();
    let mut slide = presentation.slide_mut(0).unwrap();
    let mut shape = slide.shape_mut(0).unwrap();
    shape.set_position(Emu(101), Emu(202)).unwrap();
    shape.set_size(Emu(303), Emu(404)).unwrap();
    shape.set_rotation(Angle(5_400_000)).unwrap();
    shape.set_name("A & B \"quoted\"").unwrap();
    shape
        .set_fill(
            Fill::from_xml(
                br#"<a:noFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>"#,
            )
            .unwrap(),
        )
        .unwrap();
    shape
        .set_line(
            CT_LineProperties::from_xml(br#"<a:ln xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" w="12700"/>"#).unwrap(),
        )
        .unwrap();
    shape.set_adjust_value("adj", 25_000.0).unwrap();

    let bytes = presentation.to_bytes().unwrap();
    let reopened = Presentation::from_bytes(&bytes).unwrap();
    assert_eq!(
        reopened.slide(0).unwrap().shape(0).unwrap().kind(),
        ShapeKind::Shape
    );
    let package = open_opc(&bytes, "F-109 setter gate");
    let slide = CT_Slide::from_xml(package.get_part(SLIDE_TWO_PART).unwrap()).unwrap();
    let ShapeTreeChild::Shape(shape) = &slide.common_slide_data.shape_tree.children[0] else {
        panic!("expected ordinary shape");
    };
    let transform = shape.shape_properties.transform.as_ref().unwrap();
    assert_eq!(transform.offset.unwrap().x, Emu(101));
    assert_eq!(transform.offset.unwrap().y, Emu(202));
    assert_eq!(transform.extent.unwrap().cx, Emu(303));
    assert_eq!(transform.extent.unwrap().cy, Emu(404));
    assert_eq!(transform.rotation, Angle(5_400_000));
    assert!(shape.shape_properties.fill.is_some());
    assert_eq!(
        shape.shape_properties.line.as_ref().unwrap().width,
        Some(12_700)
    );
    let adjustments = shape
        .shape_properties
        .preset_geometry
        .as_ref()
        .unwrap()
        .adjust_values();
    assert_eq!(adjustments.len(), 1);
    assert_eq!(adjustments[0].name, "adj");
    assert_eq!(
        adjustments[0].args[0],
        oxml_drawing::geometry::GuideOperand::Literal(25_000.0)
    );
    let xml = String::from_utf8(package.get_part(SLIDE_TWO_PART).unwrap().to_vec()).unwrap();
    assert!(xml.contains("name=\"A &amp; B &quot;quoted&quot;\""));
}

#[test]
fn shape_mutation_preserves_unmodelled_xml_and_schema_order() {
    let mut presentation = Presentation::from_bytes(&mutation_fixture_bytes()).unwrap();
    let mut slide = presentation.slide_mut(0).unwrap();
    let mut shape = slide.shape_mut(0).unwrap();
    shape.set_position(Emu(1), Emu(2)).unwrap();
    shape
        .set_fill(
            Fill::from_xml(
                br#"<a:noFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>"#,
            )
            .unwrap(),
        )
        .unwrap();
    shape
        .set_line(
            CT_LineProperties::from_xml(
                br#"<a:ln xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>"#,
            )
            .unwrap(),
        )
        .unwrap();

    let bytes = presentation.to_bytes().unwrap();
    let package = open_opc(&bytes, "F-109 preservation");
    let xml = String::from_utf8(package.get_part(SLIDE_TWO_PART).unwrap().to_vec()).unwrap();
    assert!(xml.contains("producer=\"kept\""));
    assert!(xml.contains("<ext:nv xmlns:ext=\"urn:ext\">one &amp; two</ext:nv>"));
    assert!(xml.contains("<ext:before xmlns:ext=\"urn:ext\"/>"));
    assert!(xml.contains("<ext:after xmlns:ext=\"urn:ext\"/>"));
    let transform = xml.find("<a:xfrm").unwrap();
    let geometry = xml.find("<a:prstGeom").unwrap();
    let fill = xml.find("<a:noFill").unwrap();
    let line = xml.find("<a:ln").unwrap();
    assert!(transform < geometry && geometry < fill && fill < line);
    Presentation::from_bytes(&bytes).expect("schema-ordered output reparses");
}

#[test]
fn shape_mutation_handles_nested_group_children() {
    let mut presentation = Presentation::from_bytes(&fixture_bytes()).unwrap();
    let mut slide = presentation.slide_mut(0).unwrap();
    let mut group = slide.shape_mut(3).unwrap();
    assert_eq!(group.kind(), ShapeKind::Group);
    assert!(group.child_mut(2).is_none());
    let mut nested = group.child_mut(0).unwrap();
    nested.set_position(Emu(41), Emu(42)).unwrap();
    nested.set_name("nested changed").unwrap();

    let bytes = presentation.to_bytes().unwrap();
    let package = open_opc(&bytes, "F-109 nested group");
    let slide = CT_Slide::from_xml(package.get_part(SLIDE_TWO_PART).unwrap()).unwrap();
    let ShapeTreeChild::GroupShape(group) = &slide.common_slide_data.shape_tree.children[3] else {
        panic!("expected group");
    };
    assert_eq!(group.children.len(), 2);
    assert_eq!(
        group
            .children
            .iter()
            .map(ShapeTreeChild::non_visual_id)
            .collect::<Vec<_>>(),
        vec![Some(7), Some(8)]
    );
    let ShapeTreeChild::Shape(nested) = &group.children[0] else {
        panic!("expected nested shape");
    };
    assert_eq!(
        nested
            .shape_properties
            .transform
            .as_ref()
            .unwrap()
            .offset
            .unwrap()
            .x,
        Emu(41)
    );
    let ShapeTreeChild::Shape(sibling) = &group.children[1] else {
        panic!("expected sibling shape");
    };
    assert_eq!(nested.text_body.as_ref().unwrap().plain_text(), "nested");
    assert_eq!(sibling.text_body.as_ref().unwrap().plain_text(), "sibling");
    assert!(sibling.shape_properties.transform.is_none());
    let xml = String::from_utf8(package.get_part(SLIDE_TWO_PART).unwrap().to_vec()).unwrap();
    let changed = xml.find("id=\"7\" name=\"nested changed\"").unwrap();
    let unchanged = xml.find("id=\"8\" name=\"nested sibling\"").unwrap();
    assert!(changed < unchanged);
    assert_eq!(slide.common_slide_data.shape_tree.children.len(), 6);
}

#[test]
fn inserted_group_transform_precedes_preserved_group_properties() {
    let mut presentation = Presentation::from_bytes(&mutation_fixture_bytes()).unwrap();
    let mut slide = presentation.slide_mut(0).unwrap();
    let mut group = slide.shape_mut(3).unwrap();
    group.set_position(Emu(111), Emu(222)).unwrap();
    group.set_size(Emu(333), Emu(444)).unwrap();
    group.set_rotation(Angle(2_700_000)).unwrap();

    let bytes = presentation.to_bytes().unwrap();
    let package = open_opc(&bytes, "F-109 group transform insertion");
    let xml = String::from_utf8(package.get_part(SLIDE_TWO_PART).unwrap().to_vec()).unwrap();
    let group_start = xml.find("<p:grpSp>").unwrap();
    let transform = xml[group_start..].find("<a:xfrm").unwrap() + group_start;
    let preserved = xml[group_start..]
        .find("<a:solidFill><a:srgbClr val=\"112233\"/></a:solidFill>")
        .unwrap()
        + group_start;
    assert!(transform < preserved);

    let slide = CT_Slide::from_xml(package.get_part(SLIDE_TWO_PART).unwrap()).unwrap();
    let ShapeTreeChild::GroupShape(group) = &slide.common_slide_data.shape_tree.children[3] else {
        panic!("expected group");
    };
    let transform = group.group_transform().unwrap();
    assert_eq!(transform.offset.unwrap().x, Emu(111));
    assert_eq!(transform.offset.unwrap().y, Emu(222));
    assert_eq!(transform.extent.unwrap().cx, Emu(333));
    assert_eq!(transform.extent.unwrap().cy, Emu(444));
    assert_eq!(transform.rotation, Angle(2_700_000));
    let group_xml = String::from_utf8(group.to_xml().unwrap()).unwrap();
    assert!(group_xml.contains("<a:xfrm rot=\"2700000\">"));
    assert!(group_xml.contains("<a:solidFill><a:srgbClr val=\"112233\"/></a:solidFill>"));
}

#[test]
fn alternate_content_fallback_is_not_mutable() {
    let original = fixture_bytes();
    let original_package = open_opc(&original, "F-109 original AlternateContent");
    let original_slide =
        CT_Slide::from_xml(original_package.get_part(SLIDE_TWO_PART).unwrap()).unwrap();
    let ShapeTreeChild::AlternateContent(original_alternate) =
        &original_slide.common_slide_data.shape_tree.children[5]
    else {
        panic!("expected AlternateContent");
    };

    let mut presentation = Presentation::from_bytes(&original).unwrap();
    let mut slide = presentation.slide_mut(0).unwrap();
    let mut alternate = slide.shape_mut(5).unwrap();
    assert!(alternate.child_mut(0).is_none());
    assert!(matches!(
        alternate.set_position(Emu(1), Emu(2)),
        Err(Error::UnsupportedShapeMutation {
            shape_kind: ShapeKind::AlternateContent,
            ..
        })
    ));

    let bytes = presentation.to_bytes().unwrap();
    let saved_package = open_opc(&bytes, "F-109 saved AlternateContent");
    let saved_slide = CT_Slide::from_xml(saved_package.get_part(SLIDE_TWO_PART).unwrap()).unwrap();
    let ShapeTreeChild::AlternateContent(saved_alternate) =
        &saved_slide.common_slide_data.shape_tree.children[5]
    else {
        panic!("expected AlternateContent");
    };
    assert_eq!(saved_alternate.raw_xml(), original_alternate.raw_xml());
}

#[test]
fn shape_mutation_indices_and_kinds_are_total() {
    let mut presentation = Presentation::from_bytes(&fixture_bytes()).unwrap();
    assert!(presentation.slide_mut(2).is_none());
    let mut slide = presentation.slide_mut(0).unwrap();
    assert!(slide.shape_mut(6).is_none());
    for index in 0..5 {
        slide
            .shape_mut(index)
            .unwrap()
            .set_position(Emu(index as i64), Emu(index as i64 + 10))
            .unwrap();
    }
    let mut picture = slide.shape_mut(1).unwrap();
    assert!(picture.child_mut(0).is_none());
    picture
        .set_fill(
            Fill::from_xml(
                br#"<a:noFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>"#,
            )
            .unwrap(),
        )
        .unwrap();
    picture
        .set_line(
            CT_LineProperties::from_xml(
                br#"<a:ln xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>"#,
            )
            .unwrap(),
        )
        .unwrap();
    let mut frame = slide.shape_mut(2).unwrap();
    assert!(matches!(
        frame.set_fill(
            Fill::from_xml(
                br#"<a:noFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>"#
            )
            .unwrap()
        ),
        Err(Error::UnsupportedShapeMutation {
            shape_kind: ShapeKind::GraphicFrame,
            ..
        })
    ));
    let mut shape = slide.shape_mut(0).unwrap();
    assert!(matches!(
        shape.set_adjust_value("adj", f64::NAN),
        Err(Error::NonFiniteAdjustmentValue { .. })
    ));
    assert!(matches!(
        shape.set_adjust_value("adj", 1.0),
        Err(Error::UnsupportedShapeMutation {
            operation: "set adjustment",
            ..
        })
    ));
    Presentation::from_bytes(&presentation.to_bytes().unwrap())
        .expect("all supported transform variants reparse");
}

#[test]
fn shape_name_mutation_escapes_xml_and_preserves_children() {
    let mut presentation = Presentation::from_bytes(&mutation_fixture_bytes()).unwrap();
    let mut slide = presentation.slide_mut(0).unwrap();
    for index in 0..5 {
        slide
            .shape_mut(index)
            .unwrap()
            .set_name(&format!("A & B \"quoted\" {index}"))
            .unwrap();
    }

    let package = open_opc(&presentation.to_bytes().unwrap(), "F-109 name mutation");
    let xml = String::from_utf8(package.get_part(SLIDE_TWO_PART).unwrap().to_vec()).unwrap();
    for index in 0..5 {
        assert!(xml.contains(&format!("name=\"A &amp; B &quot;quoted&quot; {index}\"")));
    }
    assert!(xml.contains("<ext:nv xmlns:ext=\"urn:ext\">one &amp; two</ext:nv>"));
}

#[test]
fn all_corpus_modelled_parts_reparse_structurally() {
    let Some(paths) = corpus_paths() else {
        return;
    };
    let mut coverage = BTreeMap::new();
    for path in paths {
        let deck = deck_name(&path);
        let bytes = fs::read(&path).unwrap_or_else(|error| panic!("{deck}: {error}"));
        let package = open_opc(&bytes, deck);
        for (content_type, count) in modelled_part_bytes(&package, deck).1 {
            *coverage.entry(content_type).or_insert(0) += count;
        }
    }
    for content_type in MODELLED_CONTENT_TYPES {
        assert!(
            coverage.get(content_type).copied().unwrap_or(0) > 0,
            "corpus has no modelled part with content type {content_type}"
        );
    }
}

#[test]
fn three_added_slides_have_unique_ids_and_reopen() {
    let mut presentation = Presentation::new().expect("open bundled template");
    assert!(presentation.layout_count() > 0);
    assert!(presentation.layout_name(0).is_some());
    for _ in 0..3 {
        presentation.add_slide(0).expect("add slide");
    }

    let ids = presentation
        .slides()
        .map(|slide| slide.id())
        .collect::<HashSet<_>>();
    assert_eq!(ids.len(), 3);
    assert!(ids.iter().all(|id| *id >= 256));
    let bytes = presentation.to_bytes().expect("serialize three-slide deck");
    assert_eq!(
        Presentation::from_bytes(&bytes)
            .expect("reopen three-slide deck")
            .len(),
        3
    );

    if let Some(path) = std::env::var_os("RPPTX_ACCEPTANCE_OUTPUT") {
        fs::write(path, bytes).expect("write native acceptance deck");
    }
}

#[test]
fn all_pinned_corpus_decks_validate_cleanly() {
    let Some(paths) = corpus_paths() else {
        return;
    };
    for path in paths {
        let deck = deck_name(&path);
        let bytes = fs::read(&path).unwrap_or_else(|error| panic!("{deck}: {error}"));
        let presentation = Presentation::from_bytes(&bytes)
            .unwrap_or_else(|error| panic!("{deck}: open presentation: {error}"));
        let issues = presentation.validate();
        assert!(issues.is_empty(), "{deck}: {issues:?}");
    }
}

#[test]
fn save_writes_the_same_bytes_as_to_bytes() {
    let presentation = Presentation::new().expect("open bundled template");
    let output = std::env::temp_dir().join(format!("rpptx-f108-save-{}.pptx", std::process::id()));
    presentation.save(&output).expect("save presentation");
    assert_eq!(
        fs::read(&output).expect("read saved deck"),
        presentation.to_bytes().unwrap()
    );
    fs::remove_file(output).expect("remove saved deck");
}

#[test]
#[ignore = "requires pinned Microsoft PowerPoint"]
fn three_added_slides_open_in_powerpoint_without_repair() {
    assert_powerpoint_build();
    let mut presentation = Presentation::new().expect("open bundled template");
    for _ in 0..3 {
        presentation.add_slide(0).expect("add slide");
    }
    let output = std::env::temp_dir().join(format!(
        "rpptx-f107-three-slides-{}.pptx",
        std::process::id()
    ));
    fs::write(&output, presentation.to_bytes().expect("serialize deck"))
        .expect("write native acceptance deck");
    let name = output.file_name().unwrap().to_string_lossy();
    let path = output.to_string_lossy();
    let script = format!(
        "with timeout of 120 seconds\ntell application \"Microsoft PowerPoint\"\nset previousStartUpDialog to start up dialog\ntry\nif (Version as text) is not \"{POWERPOINT_VERSION}\" then error \"PowerPoint version mismatch\"\nif (build as text) is not \"{POWERPOINT_APP_BUILD}\" then error \"PowerPoint build mismatch\"\nset start up dialog to false\nset deckPath to \"{path}\"\nopen my POSIX file deckPath\nset checkedDeck to presentation \"{name}\"\nif (count of slides of checkedDeck) is not 3 then error \"slide count mismatch\"\nclose checkedDeck saving no\nset start up dialog to previousStartUpDialog\non error errorMessage number errorNumber\ntry\nclose checkedDeck saving no\nend try\nset start up dialog to previousStartUpDialog\nerror errorMessage number errorNumber\nend try\nend tell\nend timeout\n"
    );
    let result = Command::new("osascript")
        .args(["-e", &script])
        .output()
        .expect("launch PowerPoint acceptance script");
    fs::remove_file(&output).expect("remove native acceptance deck");
    assert!(
        result.status.success(),
        "PowerPoint no-repair open failed: {}",
        String::from_utf8_lossy(&result.stderr)
    );
}

#[test]
#[ignore = "requires pinned Microsoft PowerPoint"]
fn all_shape_constructors_open_in_powerpoint_without_repair() {
    assert_powerpoint_build();
    let mut presentation = Presentation::new().expect("open bundled template");
    presentation.add_slide(0).expect("add slide");
    let mut package = open_opc(
        &presentation.to_bytes().unwrap(),
        "F-110 native append boundary",
    );
    let original =
        String::from_utf8(package.get_part("/ppt/slides/slide1.xml").unwrap().to_vec()).unwrap();
    let with_extension = original.replacen(
        "</p:spTree>",
        r#"<p:extLst><p:ext uri="{F110-NATIVE-APPEND-ORDER}"><x:marker xmlns:x="urn:f110" value="kept"/></p:ext></p:extLst></p:spTree>"#,
        1,
    );
    assert_ne!(with_extension, original);
    package.set_part("/ppt/slides/slide1.xml", with_extension.into_bytes());
    let mut presentation = Presentation::from_bytes(&package_bytes(package)).unwrap();
    let mut slide = presentation.slide_mut(0).unwrap();
    slide
        .add_textbox(Emu(914_400), Emu(914_400), Emu(2_000_000), Emu(900_000))
        .unwrap();
    slide
        .add_shape(
            "triangle",
            Emu(3_000_000),
            Emu(914_400),
            Emu(2_000_000),
            Emu(2_000_000),
        )
        .unwrap();
    slide
        .add_connector(
            ConnectorType::Curve,
            Emu(1_000_000),
            Emu(4_000_000),
            Emu(5_000_000),
            Emu(3_000_000),
        )
        .unwrap();
    slide.add_group_shape().unwrap();
    let output = std::env::temp_dir().join(format!(
        "rpptx-f110-shape-constructors-{}.pptx",
        std::process::id()
    ));
    fs::write(&output, presentation.to_bytes().expect("serialize deck"))
        .expect("write native acceptance deck");
    let name = output.file_name().unwrap().to_string_lossy();
    let path = output.to_string_lossy();
    let script = format!(
        "with timeout of 120 seconds\ntell application \"Microsoft PowerPoint\"\nset previousStartUpDialog to start up dialog\ntry\nif (Version as text) is not \"{POWERPOINT_VERSION}\" then error \"PowerPoint version mismatch\"\nif (build as text) is not \"{POWERPOINT_APP_BUILD}\" then error \"PowerPoint build mismatch\"\nset start up dialog to false\nset deckPath to \"{path}\"\nopen my POSIX file deckPath\nset checkedDeck to presentation \"{name}\"\nif (count of slides of checkedDeck) is not 1 then error \"slide count mismatch\"\nclose checkedDeck saving no\nset start up dialog to previousStartUpDialog\non error errorMessage number errorNumber\ntry\nclose checkedDeck saving no\nend try\nset start up dialog to previousStartUpDialog\nerror errorMessage number errorNumber\nend try\nend tell\nend timeout\n"
    );
    let result = Command::new("osascript")
        .args(["-e", &script])
        .output()
        .expect("launch PowerPoint acceptance script");
    fs::remove_file(&output).expect("remove native acceptance deck");
    assert!(
        result.status.success(),
        "PowerPoint no-repair open failed: {}",
        String::from_utf8_lossy(&result.stderr)
    );
}

#[test]
fn all_modelled_corpus_packages_match_expected_parts() {
    let Some(paths) = corpus_paths() else {
        return;
    };
    let save_dir = std::env::var_os("RDOCX_PPTX_SAVE_DIR").map(PathBuf::from);
    if let Some(directory) = &save_dir {
        fs::create_dir_all(directory)
            .unwrap_or_else(|error| panic!("{}: {error}", directory.display()));
    }

    let mut saved_paths = Vec::new();
    for path in paths {
        let deck = deck_name(&path);
        let original_bytes = fs::read(&path).unwrap_or_else(|error| panic!("{deck}: {error}"));
        let original = open_opc(&original_bytes, deck);
        let (rewritten, _) = modelled_part_bytes(&original, deck);

        let mut expected = original.clone();
        for (part_name, bytes) in &rewritten {
            expected.set_part(part_name, bytes.clone());
        }
        let expected = reopen_written_package(&expected, deck, "expected package");

        let facade_bytes = Presentation::from_bytes(&original_bytes)
            .unwrap_or_else(|error| panic!("{deck}: open facade: {error}"))
            .to_bytes()
            .unwrap_or_else(|error| panic!("{deck}: save facade: {error}"));
        let mut actual = open_opc(&facade_bytes, deck);
        for (part_name, bytes) in &rewritten {
            let content_type = original
                .content_types
                .overrides
                .get(part_name)
                .map(String::as_str)
                .unwrap_or_else(|| panic!("{deck} {part_name}: missing content-type override"));
            if matches!(
                content_type,
                content_types::SLIDE_LAYOUT
                    | content_types::SLIDE_MASTER
                    | content_types::NOTES_MASTER
                    | content_types::THEME
            ) {
                actual.set_part(part_name, bytes.clone());
            }
        }
        let actual_bytes = write_package(&actual, deck, "modelled package");
        let actual = open_opc(&actual_bytes, deck);
        assert_packages_equal(&expected, &actual, deck);

        if let Some(directory) = &save_dir {
            let saved_path = directory.join(deck);
            fs::write(&saved_path, &actual_bytes)
                .unwrap_or_else(|error| panic!("{}: {error}", saved_path.display()));
            saved_paths.push(saved_path);
        }
    }

    saved_paths.sort();
    for path in saved_paths {
        eprintln!(
            "sha256\t{}\t{}",
            sha256(&path),
            path.file_name().unwrap().to_string_lossy()
        );
    }
}

#[test]
fn facade_saved_corpus_reopens_with_the_same_read_surface() {
    let Some(paths) = corpus_paths() else {
        return;
    };
    for path in paths {
        let deck = deck_name(&path);
        let bytes = fs::read(&path).unwrap_or_else(|error| panic!("{deck}: {error}"));
        let before = Presentation::from_bytes(&bytes)
            .unwrap_or_else(|error| panic!("{deck}: open facade: {error}"));
        let snapshot = read_surface(&before);
        let saved = before
            .to_bytes()
            .unwrap_or_else(|error| panic!("{deck}: save facade: {error}"));
        let after = Presentation::from_bytes(&saved)
            .unwrap_or_else(|error| panic!("{deck}: reopen facade: {error}"));
        assert_eq!(snapshot, read_surface(&after), "{deck}: read surface");
    }
}

#[test]
fn rpptx_is_an_explicit_publication_candidate() {
    let workspace = include_str!("../../../Cargo.toml");
    let manifest = include_str!("../Cargo.toml");
    assert!(workspace.contains("\"crates/rpptx\""));
    assert!(workspace.contains(
        "rpptx = { path = \"crates/rpptx\", version = \"0.11.0\", default-features = false }"
    ));
    assert!(manifest.contains("name = \"rpptx\""));
    assert!(manifest.contains("version = \"0.11.0\""));
    assert!(manifest.contains("publish = true"));
    assert!(manifest.contains("default = [\"default-template\", \"render\", \"system-fonts\"]"));
    assert!(manifest.contains("default-template = [\"dep:scraper\"]"));
    assert!(manifest.contains("scraper = { workspace = true, optional = true }"));
    for dependency in [
        "dep:gif",
        "dep:jpeg-encoder",
        "dep:miniz_oxide",
        "dep:oxml-pdf",
        "dep:rpptx-layout",
        "dep:rpptx-render",
        "dep:tiny-skia",
    ] {
        assert!(manifest.contains(dependency));
    }
    assert!(
        manifest.contains(
            "system-fonts = [\"oxml-layout/system-fonts\", \"rpptx-render?/system-fonts\"]"
        )
    );
}

#[test]
#[cfg(feature = "default-template")]
fn new_presentation_uses_the_bundled_zero_slide_template() {
    let presentation = Presentation::new().expect("open bundled presentation template");
    assert!(presentation.is_empty());

    let bytes = presentation.to_bytes().expect("serialize new presentation");
    if let Some(path) = std::env::var_os("RDOCX_F105_SAVE_PATH") {
        fs::write(&path, &bytes)
            .unwrap_or_else(|error| panic!("write {}: {error}", Path::new(&path).display()));
    }
    let reopened = Presentation::from_bytes(&bytes).expect("reopen new presentation");
    assert!(reopened.is_empty());
}

#[test]
#[cfg(feature = "default-template")]
fn bundled_template_has_the_documented_part_graph() {
    let asset = workspace_root().join("crates/rpptx/assets/default.pptx");
    let bytes = fs::read(&asset)
        .unwrap_or_else(|error| panic!("read bundled template {}: {error}", asset.display()));
    let package = open_opc(&bytes, "default.pptx");
    let presentation_part = package
        .main_document_part()
        .expect("bundled template main presentation part");
    let presentation = CT_Presentation::from_xml(
        package
            .get_part(&presentation_part)
            .expect("bundled template presentation XML"),
    )
    .expect("parse bundled template presentation XML");

    assert!(presentation.slide_ids.is_empty());
    assert_eq!(presentation.slide_master_ids.len(), 1);
    let slide_size = presentation
        .slide_size
        .expect("bundled template slide size");
    assert_eq!((slide_size.cx.0, slide_size.cy.0), (12_192_000, 6_858_000));

    let count = |content_type: &str| {
        package
            .content_types
            .overrides
            .values()
            .filter(|actual| actual.as_str() == content_type)
            .count()
    };
    assert_eq!(count(content_types::SLIDE_MASTER), 1);
    assert_eq!(count(content_types::SLIDE_LAYOUT), 11);
    assert!(count(content_types::THEME) >= 1);
    assert_eq!(count(content_types::PRES_PROPS), 1);
    assert_eq!(count(content_types::VIEW_PROPS), 1);
    assert_eq!(count(content_types::TABLE_STYLES), 1);
    assert_eq!(count(content_types::NOTES_MASTER), 1);
    assert_eq!(count(content_types::SLIDE), 0);

    for (part_name, content_type) in &package.content_types.overrides {
        if content_type == content_types::THEME {
            CT_OfficeStyleSheet::from_xml(
                package
                    .get_part(part_name)
                    .unwrap_or_else(|| panic!("missing bundled theme {part_name}")),
            )
            .unwrap_or_else(|error| panic!("parse bundled theme {part_name}: {error}"));
        }
    }

    let presentation_relationships = package
        .get_part_rels(&presentation_part)
        .expect("bundled template presentation relationships");
    for relationship_type in [
        rel_types::SLIDE_MASTER,
        rel_types::PRES_PROPS,
        rel_types::VIEW_PROPS,
        rel_types::TABLE_STYLES,
        rel_types::NOTES_MASTER,
    ] {
        assert!(
            presentation_relationships
                .get_by_type(relationship_type)
                .is_some(),
            "bundled template lacks relationship type {relationship_type}"
        );
    }

    let table_styles = package
        .parts
        .iter()
        .find(|(part_name, _)| {
            package.content_types.content_type_for(part_name) == Some(content_types::TABLE_STYLES)
        })
        .map(|(_, bytes)| bytes)
        .expect("bundled template table styles");
    assert!(
        table_styles
            .windows(b"{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}".len())
            .any(|window| window == b"{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}")
    );
}

#[test]
fn presentation_resolves_ordered_slides_and_notes() {
    let presentation = Presentation::from_bytes(&fixture_bytes()).unwrap();
    assert_eq!(presentation.len(), 2);
    assert!(!presentation.is_empty());
    let slides = presentation.slides().collect::<Vec<_>>();
    assert_eq!(slides[0].id(), 900);
    assert_eq!(slides[0].name(), Some("Second in storage, first in order"));
    assert_eq!(
        slides[0].text(),
        "plain\nA\tB\nC\tD\nnested\nsibling\nfallback"
    );
    assert_eq!(slides[0].notes_text().as_deref(), Some("speaker\nnote"));
    assert_eq!(slides[1].id(), 256);
    assert_eq!(slides[1].name(), Some("First in storage, second in order"));
    assert_eq!(slides[1].notes_text(), None);
}

#[test]
fn package_root_targets_with_dot_segments_resolve_to_their_parts() {
    for target in [
        "./custom/presentation-main.xml",
        "../../custom/presentation-main.xml",
    ] {
        let mut package = fixture_package();
        package
            .package_rels
            .items
            .iter_mut()
            .find(|relationship| relationship.rel_type == rel_types::DOCUMENT)
            .unwrap()
            .target = target.to_owned();

        let presentation = open_package(package).unwrap();
        assert_eq!(presentation.len(), 2, "target {target}");
        assert_eq!(presentation.slide(0).unwrap().id(), 900, "target {target}");
    }
}

#[test]
fn indexed_read_access_is_total() {
    let presentation = Presentation::from_bytes(&fixture_bytes()).unwrap();
    assert!(presentation.slide(2).is_none());
    let slide = presentation.slide(0).unwrap();
    assert!(slide.shape(6).is_none());
    let picture = slide.shape(1).unwrap();
    assert_eq!(picture.child_count(), 0);
    assert!(picture.child(0).is_none());
    assert_eq!(picture.children().len(), 0);
    let group = slide.shape(3).unwrap();
    assert_eq!(group.child_count(), 2);
    assert!(group.child(2).is_none());
    let alternate = slide.shape(5).unwrap();
    assert_eq!(alternate.child_count(), 1);
    assert!(alternate.child(1).is_none());

    let empty = Presentation::from_bytes(&package_bytes(empty_package())).unwrap();
    assert!(empty.is_empty());
    assert!(empty.slide(0).is_none());
    assert_eq!(empty.slides().len(), 0);
}

#[test]
fn shape_refs_cover_the_typed_shape_tree() {
    let presentation = Presentation::from_bytes(&fixture_bytes()).unwrap();
    let slide = presentation.slide(0).unwrap();
    let shapes = slide.shapes().collect::<Vec<_>>();
    assert_eq!(
        shapes.iter().map(|shape| shape.kind()).collect::<Vec<_>>(),
        vec![
            ShapeKind::Shape,
            ShapeKind::Picture,
            ShapeKind::GraphicFrame,
            ShapeKind::Group,
            ShapeKind::Connector,
            ShapeKind::AlternateContent,
        ]
    );
    assert_eq!(shapes[0].text().as_deref(), Some("plain"));
    assert_eq!(shapes[1].text(), None);
    assert_eq!(shapes[2].text().as_deref(), Some("A\tB\nC\tD"));
    assert_eq!(
        shapes[3].children().next().unwrap().text().as_deref(),
        Some("nested")
    );
    assert_eq!(
        shapes[3].children().nth(1).unwrap().text().as_deref(),
        Some("sibling")
    );
    assert_eq!(
        shapes[5].children().next().unwrap().text().as_deref(),
        Some("fallback")
    );
}

#[test]
fn broken_presentation_graph_returns_contextual_errors() {
    let mut missing = fixture_package();
    missing.parts.remove(SLIDE_TWO_PART);
    assert!(matches!(
        open_package(missing),
        Err(Error::MissingPart { part_name }) if part_name == SLIDE_TWO_PART
    ));

    let mut wrong_type = fixture_package();
    wrong_type
        .get_or_create_part_rels(PRESENTATION_PART)
        .add_with_id(
            "ordered-first",
            rel_types::SLIDE_LAYOUT,
            "slides/second.xml",
        );
    assert!(matches!(
        open_package(wrong_type),
        Err(Error::WrongRelationshipType { relationship_id, .. })
            if relationship_id == "ordered-first"
    ));

    let mut external = fixture_package();
    let relationships = external.get_or_create_part_rels(PRESENTATION_PART);
    let relationship = relationships
        .items
        .iter_mut()
        .find(|relationship| relationship.id == "ordered-first")
        .unwrap();
    relationship.target_mode = Some("External".to_owned());
    assert!(matches!(
        open_package(external),
        Err(Error::ExternalRelationship { relationship_id, .. })
            if relationship_id == "ordered-first"
    ));

    let mut duplicate_notes = fixture_package();
    duplicate_notes
        .get_or_create_part_rels(SLIDE_TWO_PART)
        .add(rel_types::NOTES_SLIDE, "../notes/other.xml");
    assert!(matches!(
        open_package(duplicate_notes),
        Err(Error::DuplicateNotesSlides { count: 2, .. })
    ));

    let mut malformed = fixture_package();
    malformed.set_part(SLIDE_TWO_PART, b"<p:sld".to_vec());
    assert!(matches!(
        open_package(malformed),
        Err(Error::MalformedPart { part_name, .. }) if part_name == SLIDE_TWO_PART
    ));
}

#[test]
fn facade_bytes_reopen_with_the_same_read_model() {
    let original = fixture_bytes();
    let presentation = Presentation::from_bytes(&original).unwrap();
    let saved = presentation.to_bytes().unwrap();
    let reopened = Presentation::from_bytes(&saved).unwrap();
    assert_eq!(presentation.len(), reopened.len());
    for (before, after) in presentation.slides().zip(reopened.slides()) {
        assert_eq!(before.id(), after.id());
        assert_eq!(before.name(), after.name());
        assert_eq!(before.text(), after.text());
        assert_eq!(before.notes_text(), after.notes_text());
    }

    let original_package = OpcPackage::from_reader(Cursor::new(original)).unwrap();
    let saved_package = OpcPackage::from_reader(Cursor::new(saved)).unwrap();
    assert_eq!(
        original_package.get_part("/custom/opaque/data.bin"),
        saved_package.get_part("/custom/opaque/data.bin")
    );
    assert_eq!(
        original_package.get_part("/custom/opaque/raw.xml"),
        saved_package.get_part("/custom/opaque/raw.xml")
    );
}

#[test]
fn dump_deck_matches_python_pptx_1_0_2_for_the_corpus() {
    let corpus = corpus_dir();
    if !corpus.is_dir() {
        assert!(
            std::env::var_os("RDOCX_PPTX_CORPUS_REQUIRED").is_none(),
            "required corpus is missing at {}",
            corpus.display()
        );
        eprintln!(
            "corpus differential skipped because {} is absent",
            corpus.display()
        );
        return;
    }
    let entries = manifest_entries();
    assert_eq!(entries.len(), 50);
    let paths = entries
        .iter()
        .map(|entry| corpus.join(entry))
        .collect::<Vec<_>>();
    assert!(paths.iter().all(|path| path.is_file()));

    let mut command = Command::new("uv");
    command.args([
        "run",
        "--with",
        "python-pptx==1.0.2",
        "python",
        "-c",
        PYTHON_ORACLE,
    ]);
    command.args(&paths);
    let output = command.output().expect("run pinned python-pptx oracle");
    assert!(
        output.status.success(),
        "python-pptx oracle failed:\n{}",
        String::from_utf8_lossy(&output.stderr)
    );
    let oracle = String::from_utf8(output.stdout).unwrap();

    let mut rust = String::new();
    for path in &paths {
        rust.push_str("deck\t");
        rust.push_str(path.file_name().unwrap().to_str().unwrap());
        rust.push('\n');
        rust.push_str(&dump_deck::records_for_path(path).unwrap());
        rust.push('\n');
    }
    if rust != oracle {
        let rust_lines = rust.lines().collect::<Vec<_>>();
        let oracle_lines = oracle.lines().collect::<Vec<_>>();
        let mismatch = rust_lines
            .iter()
            .zip(&oracle_lines)
            .position(|(rust, oracle)| rust != oracle)
            .unwrap_or_else(|| rust_lines.len().min(oracle_lines.len()));
        panic!(
            "normalized output differs at line {}:\nrpptx: {}\npython-pptx: {}\nline counts: {} and {}",
            mismatch + 1,
            rust_lines.get(mismatch).copied().unwrap_or("<missing>"),
            oracle_lines.get(mismatch).copied().unwrap_or("<missing>"),
            rust_lines.len(),
            oracle_lines.len(),
        );
    }
}

#[test]
fn resolved_visual_dump_matches_python_pptx_1_0_2() {
    let Some(decks) = resolve_visual_decks() else {
        return;
    };
    let rust = normalized_visual_oracle_records(&decks);
    let paths = VISUAL_DECKS
        .iter()
        .map(|name| corpus_dir().join(name))
        .collect::<Vec<_>>();
    let mut command = Command::new("uv");
    command.args([
        "run",
        "--with",
        "python-pptx==1.0.2",
        "python",
        "-c",
        PYTHON_VISUAL_ORACLE,
    ]);
    command.args(&paths);
    let output = command
        .output()
        .expect("run pinned visual python-pptx oracle");
    assert!(
        output.status.success(),
        "visual python-pptx oracle failed:\n{}",
        String::from_utf8_lossy(&output.stderr)
    );
    let python = String::from_utf8(output.stdout).unwrap();
    assert_eq!(rust, python, "normalized resolved visual records");

    let cyan_deck = decks
        .iter()
        .find(|deck| deck.name == "placeholder-layout-color.pptx")
        .expect("cyan placeholder deck");
    let cyan_fills = cyan_deck.slides[0]
        .resolved
        .shapes
        .iter()
        .filter(|shape| resolved_content_text(&shape.content) == "CYAN TEXT")
        .flat_map(|shape| resolved_run_fills(&shape.content))
        .collect::<Vec<_>>();
    assert_eq!(
        cyan_fills,
        ["Solid(Color { r: 0.0, g: 1.0, b: 1.0, a: 1.0 })"],
        "placeholder run must resolve to exact cyan RGBA"
    );

    let with_master = decks
        .iter()
        .find(|deck| deck.name == "WithMaster.pptx")
        .expect("latent-placeholder deck");
    let first_text = with_master.slides[0]
        .resolved
        .shapes
        .iter()
        .map(|shape| resolved_content_text(&shape.content))
        .collect::<Vec<_>>();
    let second_text = with_master.slides[1]
        .resolved
        .shapes
        .iter()
        .map(|shape| resolved_content_text(&shape.content))
        .collect::<Vec<_>>();
    for text in [&first_text, &second_text] {
        assert!(!text.iter().any(|value| value == "9/26/2011"));
        assert!(!text.iter().any(|value| value == "‹#›"));
    }
    assert_eq!(
        second_text
            .iter()
            .filter(|value| value.as_str() == "Footer from the master slide")
            .count(),
        1,
        "occupied slide footer must resolve exactly once"
    );

    let header_footer = decks
        .iter()
        .find(|deck| deck.name == "bug58144-headers-footers-2007.pptx")
        .expect("header-footer policy deck");
    let header_footer_text = header_footer.slides[0]
        .resolved
        .shapes
        .iter()
        .map(|shape| resolved_content_text(&shape.content))
        .collect::<Vec<_>>();
    assert_eq!(
        header_footer_text
            .iter()
            .filter(|value| value.as_str() == "Slide footer")
            .count(),
        1,
        "enabled slide footer must remain visible exactly once"
    );

    if let Some(path) = std::env::var_os("RDOCX_PPTX_VISUAL_EVIDENCE") {
        let path = PathBuf::from(path);
        fs::write(&path, normalized_resolved_evidence(&decks))
            .unwrap_or_else(|error| panic!("{}: {error}", path.display()));
        eprintln!("normalized visual evidence: {}", path.display());
    }
}

#[test]
fn visual_decks_resolve_in_expected_draw_order() {
    let Some(decks) = resolve_visual_decks() else {
        return;
    };
    let records = normalized_visual_oracle_records(&decks);
    assert!(records.contains("This text comes from the Master Slide"));
    assert!(records.contains("First page title"));
    let master_text = records
        .find("This text comes from the Master Slide")
        .unwrap();
    let slide_text = records.find("First page title").unwrap();
    assert!(
        master_text < slide_text,
        "layout artwork must precede slide content"
    );
}

#[test]
fn visual_decks_never_emit_template_prompt_text() {
    let Some(decks) = resolve_visual_decks() else {
        return;
    };
    let evidence = normalized_resolved_evidence(&decks);
    for prompt in ["Click to edit", "Click to add"] {
        assert!(
            !evidence.contains(prompt),
            "template prompt leaked: {prompt}"
        );
    }
}

#[test]
fn visual_decks_emit_each_master_logo_once() {
    let Some(decks) = resolve_visual_decks() else {
        return;
    };
    let deck = decks
        .iter()
        .find(|deck| deck.name == "60810.pptx")
        .expect("master-logo deck");
    for (slide_index, slide) in deck.slides.iter().enumerate() {
        let logos = slide
            .resolved
            .shapes
            .iter()
            .filter(|shape| {
                shape.unsupported == Some("image media pending relationship resolution")
                    && (shape.bounds.x - 611.751).abs() < 0.001
                    && (shape.bounds.y - 27.000).abs() < 0.001
                    && (shape.bounds.width - 101.732).abs() < 0.001
                    && (shape.bounds.height - 20.312).abs() < 0.001
            })
            .count();
        let expected = usize::from(!matches!(slide_index, 2 | 3));
        assert_eq!(
            logos, expected,
            "60810.pptx slide {slide_index} master logo"
        );
    }
}

#[test]
fn selected_visual_decks_reviewed_once_in_powerpoint() {
    assert!(POWERPOINT_VISUAL_ACCEPTANCE.contains("bundle 16.104.25121423"));
    assert_eq!(
        POWERPOINT_VISUAL_ACCEPTANCE
            .lines()
            .filter(|line| line.starts_with("deck\t"))
            .count(),
        4
    );
    assert!(POWERPOINT_VISUAL_ACCEPTANCE.contains("result\tclean"));
    assert!(POWERPOINT_VISUAL_ACCEPTANCE.contains("RGBA #00FFFF"));
    assert!(POWERPOINT_VISUAL_ACCEPTANCE.contains("slide footer once on slide 2"));
}

struct ResolvedVisualDeck {
    name: &'static str,
    slides: Vec<ResolvedVisualSlide>,
}

struct ResolvedVisualSlide {
    resolved: ResolvedSlide,
    shape_metadata: Vec<ResolvedVisualShapeMetadata>,
}

struct ResolvedVisualShapeMetadata {
    kind: &'static str,
    is_latent: bool,
}

fn resolve_visual_decks() -> Option<Vec<ResolvedVisualDeck>> {
    let corpus = corpus_dir();
    if !corpus.is_dir() {
        assert!(
            std::env::var_os("RDOCX_PPTX_CORPUS_REQUIRED").is_none(),
            "required visual corpus is missing at {}",
            corpus.display()
        );
        eprintln!(
            "visual differential skipped because {} is absent",
            corpus.display()
        );
        return None;
    }
    Some(resolve_visual_decks_from_corpus(&corpus))
}

fn resolve_visual_decks_from_corpus(corpus: &Path) -> Vec<ResolvedVisualDeck> {
    VISUAL_DECKS
        .iter()
        .map(|&name| {
            let path = corpus.join(name);
            assert!(
                path.is_file(),
                "missing selected visual deck {}",
                path.display()
            );
            let package = OpcPackage::open(&path)
                .unwrap_or_else(|error| panic!("{}: {error}", path.display()));
            let presentation_part = package
                .main_document_part()
                .unwrap_or_else(|| panic!("{}: no presentation part", path.display()));
            let presentation = CT_Presentation::from_xml(visual_part(&package, &presentation_part))
                .unwrap_or_else(|error| panic!("{}: {error}", path.display()));
            let default_text_style = presentation.default_text_style.unwrap_or_default();
            let size = presentation.slide_size.map_or((720.0, 540.0), |size| {
                (size.cx.0 as f64 / 12_700.0, size.cy.0 as f64 / 12_700.0)
            });
            let presentation_rels = package
                .get_part_rels(&presentation_part)
                .unwrap_or_else(|| panic!("{}: no presentation relationships", path.display()));
            let slides = presentation
                .slide_ids
                .iter()
                .map(|slide_id| {
                    let slide_rel = presentation_rels
                        .get_by_id(&slide_id.relationship_id)
                        .unwrap_or_else(|| {
                            panic!(
                                "{}: missing slide relationship {}",
                                path.display(),
                                slide_id.relationship_id
                            )
                        });
                    assert_eq!(slide_rel.rel_type, rel_types::SLIDE, "{}", path.display());
                    let slide_part =
                        OpcPackage::resolve_rel_target(&presentation_part, &slide_rel.target);
                    let layout_part =
                        related_visual_part(&package, &slide_part, rel_types::SLIDE_LAYOUT, &path);
                    let master_part =
                        related_visual_part(&package, &layout_part, rel_types::SLIDE_MASTER, &path);
                    let theme_part =
                        related_visual_part(&package, &master_part, rel_types::THEME, &path);
                    let slide = CT_Slide::from_xml(visual_part(&package, &slide_part))
                        .unwrap_or_else(|error| panic!("{} {slide_part}: {error}", path.display()));
                    let layout = CT_SlideLayout::from_xml(visual_part(&package, &layout_part))
                        .unwrap_or_else(|error| {
                            panic!("{} {layout_part}: {error}", path.display())
                        });
                    let master = CT_SlideMaster::from_xml(visual_part(&package, &master_part))
                        .unwrap_or_else(|error| {
                            panic!("{} {master_part}: {error}", path.display())
                        });
                    let theme = CT_OfficeStyleSheet::from_xml(visual_part(&package, &theme_part))
                        .unwrap_or_else(|error| panic!("{} {theme_part}: {error}", path.display()));
                    let color_map = effective_visual_color_map(&master, &layout, &slide);
                    let context = ResolveCtx::new(
                        &theme,
                        color_map,
                        &master,
                        &layout,
                        &slide,
                        &default_text_style,
                    );
                    let shape_metadata = context
                        .flatten()
                        .into_iter()
                        .filter_map(|item| match item {
                            FlattenedItem::Background(_) => None,
                            FlattenedItem::Shape { child, .. } => {
                                flattened_child_has_bounds(&context, child).then(|| {
                                    ResolvedVisualShapeMetadata {
                                        kind: flattened_child_kind(child),
                                        is_latent: flattened_child_is_latent(child),
                                    }
                                })
                            }
                        })
                        .collect::<Vec<_>>();
                    let resolved = context
                        .resolve_slide(size)
                        .unwrap_or_else(|error| panic!("{} {slide_part}: {error}", path.display()));
                    assert_eq!(
                        shape_metadata.len(),
                        resolved.shapes.len(),
                        "{} {slide_part}: flattened and resolved shape counts",
                        path.display()
                    );
                    ResolvedVisualSlide {
                        resolved,
                        shape_metadata,
                    }
                })
                .collect();
            ResolvedVisualDeck { name, slides }
        })
        .collect()
}

fn flattened_child_kind(child: &ShapeTreeChild) -> &'static str {
    match child {
        ShapeTreeChild::Shape(_) => "Shape",
        ShapeTreeChild::Picture(_) => "Picture",
        ShapeTreeChild::GraphicFrame(_) => "GraphicFrame",
        ShapeTreeChild::Connector(_) => "Connector",
        ShapeTreeChild::GroupShape(_) => "Group",
        ShapeTreeChild::AlternateContent(_) => "AlternateContent",
    }
}

fn flattened_child_has_bounds(context: &ResolveCtx<'_>, child: &ShapeTreeChild) -> bool {
    match child {
        ShapeTreeChild::Shape(shape) => context
            .effective_xfrm(shape)
            .as_ref()
            .is_some_and(transform_has_bounds),
        ShapeTreeChild::Picture(picture) => picture
            .shape_properties
            .transform
            .as_ref()
            .is_some_and(transform_has_bounds),
        ShapeTreeChild::GraphicFrame(frame) => transform_has_bounds(&frame.transform),
        ShapeTreeChild::Connector(connector) => connector
            .shape_properties
            .transform
            .as_ref()
            .is_some_and(transform_has_bounds),
        ShapeTreeChild::GroupShape(_) | ShapeTreeChild::AlternateContent(_) => false,
    }
}

fn transform_has_bounds(transform: &oxml_drawing::xfrm::CT_Transform2D) -> bool {
    transform.offset.is_some() && transform.extent.is_some()
}

fn flattened_child_is_latent(child: &ShapeTreeChild) -> bool {
    let placeholder = match child {
        ShapeTreeChild::Shape(shape) => shape.placeholder.as_ref(),
        ShapeTreeChild::Picture(picture) => picture.placeholder.as_ref(),
        _ => None,
    };
    placeholder.is_some_and(|placeholder| {
        matches!(
            placeholder.effective_type(),
            PhType::DateTime | PhType::Footer | PhType::SlideNumber
        )
    })
}

fn visual_part<'a>(package: &'a OpcPackage, part_name: &str) -> &'a [u8] {
    package
        .get_part(part_name)
        .unwrap_or_else(|| panic!("missing package part {part_name}"))
}

fn related_visual_part(
    package: &OpcPackage,
    source_part: &str,
    relationship_type: &str,
    deck: &Path,
) -> String {
    let relationship = package
        .get_part_rels(source_part)
        .and_then(|relationships| relationships.get_by_type(relationship_type))
        .unwrap_or_else(|| {
            panic!(
                "{} {source_part}: missing relationship {relationship_type}",
                deck.display()
            )
        });
    OpcPackage::resolve_rel_target(source_part, &relationship.target)
}

fn effective_visual_color_map(
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

fn normalized_visual_oracle_records(decks: &[ResolvedVisualDeck]) -> String {
    let mut output = String::new();
    for deck in decks {
        writeln!(output, "deck\t{}", deck.name).unwrap();
        for (slide_index, visual_slide) in deck.slides.iter().enumerate() {
            let slide = &visual_slide.resolved;
            writeln!(
                output,
                "slide\t{slide_index}\t{:.3}\t{:.3}",
                slide.size.0, slide.size.1
            )
            .unwrap();
            for (shape, metadata) in slide.shapes.iter().zip(&visual_slide.shape_metadata) {
                // python-pptx exposes header and footer placeholders as raw
                // collection members rather than applying the effective p:hf
                // policy. Their policy is asserted by the Rust regressions.
                if metadata.is_latent {
                    continue;
                }
                let bounds = shape.group_transform.transform_rect_bbox(shape.bounds);
                writeln!(
                    output,
                    "shape\t{slide_index}\t{}\t{:.3}\t{:.3}\t{:.3}\t{:.3}\t{}",
                    metadata.kind,
                    bounds.x,
                    bounds.y,
                    bounds.width,
                    bounds.height,
                    escape_record(&resolved_content_text(&shape.content))
                )
                .unwrap();
            }
        }
    }
    output
}

fn normalized_resolved_evidence(decks: &[ResolvedVisualDeck]) -> String {
    let mut output = String::new();
    for deck in decks {
        writeln!(output, "deck\t{}", deck.name).unwrap();
        for (slide_index, visual_slide) in deck.slides.iter().enumerate() {
            let slide = &visual_slide.resolved;
            writeln!(
                output,
                "slide\t{slide_index}\tsize={:.3}x{:.3}\tbackground={:?}\tdiagnostics={:?}",
                slide.size.0, slide.size.1, slide.background, slide.diagnostics
            )
            .unwrap();
            for (shape_index, (shape, metadata)) in slide
                .shapes
                .iter()
                .zip(&visual_slide.shape_metadata)
                .enumerate()
            {
                writeln!(
                    output,
                    "shape\t{slide_index}.{shape_index}\tkind={}\tbounds={:.3},{:.3},{:.3},{:.3}\ttransform={:?}\tfill={:?}\tline={:?}\trun_fills={:?}\ttext={}\tunsupported={}",
                    metadata.kind,
                    shape.bounds.x,
                    shape.bounds.y,
                    shape.bounds.width,
                    shape.bounds.height,
                    shape.group_transform,
                    shape.fill,
                    shape.line,
                    resolved_run_fills(&shape.content),
                    escape_record(&resolved_content_text(&shape.content)),
                    shape.unsupported.unwrap_or("-")
                )
                .unwrap();
            }
        }
    }
    output
}

fn resolved_run_fills(content: &ResolvedContent) -> Vec<String> {
    let mut fills = Vec::new();
    let mut collect_body = |body: &ResolvedTextBody| {
        for paragraph in &body.paragraphs {
            for run in &paragraph.runs {
                let style = match run {
                    ResolvedTextRun::Text { style, .. } | ResolvedTextRun::Field { style, .. } => {
                        style
                    }
                    ResolvedTextRun::Break => continue,
                };
                if let Some(fill) = &style.fill {
                    fills.push(format!("{fill:?}"));
                }
            }
        }
    };
    match content {
        ResolvedContent::Text(body) => collect_body(body),
        ResolvedContent::Table(table) => {
            for body in table
                .rows
                .iter()
                .flat_map(|row| row.cells.iter())
                .filter_map(|cell| cell.text.as_ref())
            {
                collect_body(body);
            }
        }
        ResolvedContent::None | ResolvedContent::Image(_) | ResolvedContent::Group(_) => {}
    }
    fills
}

fn resolved_content_text(content: &ResolvedContent) -> String {
    match content {
        ResolvedContent::Text(body) => resolved_text(body),
        ResolvedContent::Table(table) => table
            .rows
            .iter()
            .map(|row| {
                row.cells
                    .iter()
                    .map(|cell| cell.text.as_ref().map(resolved_text).unwrap_or_default())
                    .collect::<Vec<_>>()
                    .join("\t")
            })
            .collect::<Vec<_>>()
            .join("\n"),
        ResolvedContent::Group(group) => {
            let mut text = Vec::new();
            walk(&group.children, &mut |element, _| {
                if let PositionedElement::Text(run) = element {
                    text.push(run.text.clone());
                }
            });
            text.join("\n")
        }
        ResolvedContent::None | ResolvedContent::Image(_) => String::new(),
    }
}

fn resolved_text(body: &ResolvedTextBody) -> String {
    body.paragraphs
        .iter()
        .map(|paragraph| {
            paragraph
                .runs
                .iter()
                .map(|run| match run {
                    ResolvedTextRun::Text { text, .. } | ResolvedTextRun::Field { text, .. } => {
                        text.as_str()
                    }
                    ResolvedTextRun::Break => "\n",
                })
                .collect::<String>()
        })
        .collect::<Vec<_>>()
        .join("\n")
}

fn escape_record(value: &str) -> String {
    value
        .replace('\\', "\\\\")
        .replace('\t', "\\t")
        .replace('\r', "\\r")
        .replace('\n', "\\n")
}

fn corpus_paths() -> Option<Vec<PathBuf>> {
    let corpus = corpus_dir();
    if !corpus.is_dir() {
        assert!(
            std::env::var_os("RDOCX_PPTX_CORPUS_REQUIRED").is_none(),
            "required corpus is missing at {}",
            corpus.display()
        );
        eprintln!(
            "corpus round-trip skipped because {} is absent",
            corpus.display()
        );
        return None;
    }
    let entries = manifest_entries();
    assert_eq!(
        entries.len(),
        50,
        "the corpus manifest must contain 50 decks"
    );
    let mut paths = entries
        .into_iter()
        .map(|entry| corpus.join(entry))
        .collect::<Vec<_>>();
    paths.sort();
    assert!(
        paths.iter().all(|path| path.is_file()),
        "every corpus manifest entry must exist under {}",
        corpus.display()
    );
    Some(paths)
}

fn deck_name(path: &Path) -> &str {
    path.file_name()
        .and_then(|name| name.to_str())
        .unwrap_or_else(|| panic!("non-UTF-8 corpus filename: {}", path.display()))
}

fn open_opc(bytes: &[u8], deck: &str) -> OpcPackage {
    OpcPackage::from_reader(Cursor::new(bytes))
        .unwrap_or_else(|error| panic!("{deck}: open OPC package: {error}"))
}

fn deterministic_render(bytes: &[u8]) -> Vec<u8> {
    let layout = render_presentation_package(bytes);
    oxml_pdf::render_page_to_png(&layout, 0, 72.0).unwrap()
}

fn render_presentation_package(bytes: &[u8]) -> oxml_layout::LayoutResult {
    Presentation::from_bytes(bytes)
        .expect("F-112 deterministic render package")
        .render_deterministic()
        .unwrap()
        .1
}

#[test]
fn corpus_example_and_facade_rendering_are_identical() {
    let mut presentation = Presentation::new().expect("create presentation");
    presentation.add_slide(0).expect("add slide");
    let bytes = presentation.to_bytes().expect("serialize presentation");
    let package = open_opc(&bytes, "F-142 example parity");
    let example = render_deck_example::render_package(&package, Path::new("F-142 parity"))
        .expect("example render");
    let (_, facade) = Presentation::from_bytes(&bytes)
        .expect("open facade")
        .render_deterministic()
        .expect("facade render");

    assert_eq!(
        oxml_pdf::render_to_pdf(&example.layout),
        oxml_pdf::render_to_pdf(&facade)
    );
    assert_eq!(example.input.slides.len(), facade.pages.len());
}

fn timeline_fixture_bytes() -> Vec<u8> {
    let mut presentation = Presentation::new().expect("create timeline presentation");
    presentation.add_slide(0).expect("add timeline slide");
    presentation
        .slide_mut(0)
        .unwrap()
        .add_shape(
            "rect",
            Emu(914_400),
            Emu(914_400),
            Emu(1_828_800),
            Emu(914_400),
        )
        .unwrap();
    let bytes = presentation.to_bytes().expect("serialize timeline source");
    let mut package = open_opc(&bytes, "timeline source");
    let presentation_part = package.main_document_part().unwrap();
    let model = CT_Presentation::from_xml(package.get_part(&presentation_part).unwrap()).unwrap();
    let relationship = package
        .get_part_rels(&presentation_part)
        .unwrap()
        .get_by_id(&model.slide_ids[0].relationship_id)
        .unwrap();
    let slide_part = OpcPackage::resolve_rel_target(&presentation_part, &relationship.target);
    let slide = String::from_utf8(package.get_part(&slide_part).unwrap().to_vec()).unwrap();
    let slide_model = CT_Slide::from_xml(slide.as_bytes()).unwrap();
    let shape_id = slide_model.common_slide_data.shape_tree.children[0]
        .non_visual_id()
        .unwrap();
    let timing = format!(
        r#"<p:timing><p:tnLst><p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="1" dur="1000" fill="hold"/><p:tgtEl><p:spTgt spid="{shape_id}"/></p:tgtEl></p:cBhvr></p:animEffect></p:tnLst></p:timing>"#
    );
    package.set_part(
        &slide_part,
        slide
            .replacen("</p:sld>", &format!("{timing}</p:sld>"), 1)
            .into_bytes(),
    );
    package_bytes(package)
}

#[test]
fn smartart_oracle_sources_preserve_the_valid_default_presentation_graph() {
    if !pinned_smartart_resources_available() {
        eprintln!(
            "authentic SmartArt source structure skipped because pinned PowerPoint resources are absent"
        );
        return;
    }
    for family in [
        "list",
        "hierarchy",
        "cycle",
        "relationship",
        "matrix",
        "pyramid",
    ] {
        let bytes = authentic_smartart_oracle_source_bytes(family);
        let package = OpcPackage::from_reader(Cursor::new(&bytes)).unwrap();
        assert!(!package.parts.keys().any(|part| part.ends_with(".png")));
        assert!(!package.content_types.defaults.contains_key("png"));
        assert_eq!(
            package
                .parts
                .keys()
                .filter(|part| part.starts_with("/ppt/diagrams/"))
                .map(String::as_str)
                .collect::<HashSet<_>>(),
            HashSet::from([
                "/ppt/diagrams/data1.xml",
                "/ppt/diagrams/layout1.xml",
                "/ppt/diagrams/quickStyle1.xml",
                "/ppt/diagrams/colors1.xml",
            ])
        );
        let presentation_part = package.main_document_part().unwrap();
        let presentation =
            CT_Presentation::from_xml(package.get_part(&presentation_part).unwrap()).unwrap();
        assert_eq!(presentation.slide_ids.len(), 1, "{family} slide count");
        assert_eq!(
            presentation.slide_master_ids.len(),
            1,
            "{family} master count"
        );
        assert!(presentation.slide_size.is_some(), "{family} slide size");
        let presentation_relationships = package.get_part_rels(&presentation_part).unwrap();
        let master_relationship = presentation_relationships
            .items
            .iter()
            .find(|relationship| relationship.rel_type == rel_types::SLIDE_MASTER)
            .expect("presentation master relationship");
        let master_part =
            OpcPackage::resolve_rel_target(&presentation_part, &master_relationship.target);
        let slide_relationship = presentation_relationships
            .get_by_id(&presentation.slide_ids[0].relationship_id)
            .unwrap();
        let slide_part =
            OpcPackage::resolve_rel_target(&presentation_part, &slide_relationship.target);
        let slide = CT_Slide::from_xml(package.get_part(&slide_part).unwrap()).unwrap();
        assert_eq!(slide.common_slide_data.shape_tree.children.len(), 1);
        let slide_relationships = package.get_part_rels(&slide_part).unwrap();
        assert_eq!(
            slide_relationships
                .items
                .iter()
                .map(|relationship| relationship.rel_type.as_str())
                .collect::<HashSet<_>>(),
            HashSet::from([
                rel_types::SLIDE_LAYOUT,
                rel_types::DIAGRAM_DATA,
                rel_types::DIAGRAM_LAYOUT,
                rel_types::DIAGRAM_QUICK_STYLE,
                rel_types::DIAGRAM_COLORS,
            ])
        );
        let data_relationship = slide_relationships
            .items
            .iter()
            .find(|relationship| relationship.rel_type == rel_types::DIAGRAM_DATA)
            .unwrap();
        let data_part = OpcPackage::resolve_rel_target(&slide_part, &data_relationship.target);
        let data_xml = String::from_utf8(package.get_part(&data_part).unwrap().to_vec()).unwrap();
        let data = rpptx::CT_DiagramData::from_xml(data_xml.as_bytes()).unwrap();
        assert_eq!(data.points().len(), 4, "{family} point count");
        assert_eq!(data.connections().len(), 3, "{family} connection count");
        assert!(data.points().iter().all(|point| {
            !point.model_id.is_empty() && point.model_id.chars().all(|value| value.is_ascii_digit())
        }));
        assert!(data.connections().iter().all(|connection| {
            connection
                .model_id
                .chars()
                .all(|value| value.is_ascii_digit())
                && connection
                    .source_id
                    .chars()
                    .all(|value| value.is_ascii_digit())
                && connection
                    .destination_id
                    .chars()
                    .all(|value| value.is_ascii_digit())
        }));
        assert_eq!(data_xml.matches(" srcOrd=").count(), 3);
        assert_eq!(data_xml.matches(" destOrd=").count(), 3);
        assert!(data_xml.contains(r#"modelId="0" type="doc""#));
        assert!(data_xml.contains(r#"srcId="0" destId="1""#));
        let layout_relationship = slide_relationships
            .items
            .iter()
            .find(|relationship| relationship.rel_type == rel_types::DIAGRAM_LAYOUT)
            .unwrap();
        let diagram_layout_part =
            OpcPackage::resolve_rel_target(&slide_part, &layout_relationship.target);
        let diagram_layout_xml =
            String::from_utf8(package.get_part(&diagram_layout_part).unwrap().to_vec()).unwrap();
        for invalid_token in [r#"type="titleBand""#, r#"type="cols""#, r#"type="weights""#] {
            assert!(
                !diagram_layout_xml.contains(invalid_token),
                "{family} contained non-schema DiagramML token {invalid_token}"
            );
        }
        let layout_relationship = slide_relationships
            .items
            .iter()
            .find(|relationship| relationship.rel_type == rel_types::SLIDE_LAYOUT)
            .unwrap();
        let layout_part = OpcPackage::resolve_rel_target(&slide_part, &layout_relationship.target);
        CT_SlideLayout::from_xml(package.get_part(&layout_part).unwrap()).unwrap();
        assert!(
            package
                .get_part_rels(&layout_part)
                .unwrap()
                .items
                .iter()
                .any(
                    |relationship| relationship.rel_type == rel_types::SLIDE_MASTER
                        && OpcPackage::resolve_rel_target(&layout_part, &relationship.target)
                            == master_part
                )
        );
        CT_SlideMaster::from_xml(package.get_part(&master_part).unwrap()).unwrap();
        let master_relationships = package.get_part_rels(&master_part).unwrap();
        assert!(
            master_relationships
                .items
                .iter()
                .any(|relationship| relationship.rel_type == rel_types::THEME)
        );
        assert!(
            master_relationships
                .items
                .iter()
                .any(|relationship| relationship.rel_type == rel_types::SLIDE_LAYOUT)
        );
        assert!(
            Presentation::from_bytes(&bytes)
                .unwrap()
                .validate()
                .is_empty()
        );
    }
}

fn powerpoint_timeline_oracle_source_bytes() -> Vec<u8> {
    let mut presentation = Presentation::new().expect("create PowerPoint timeline oracle source");
    for slide_index in 0..6 {
        presentation.add_slide(0).expect("add oracle slide");
        let mut slide = presentation.slide_mut(slide_index).unwrap();
        let mut shape = slide
            .add_shape(
                "rect",
                Emu(if slide_index == 3 { 6_858_000 } else { 914_400 }),
                Emu(1_371_600),
                Emu(3_657_600),
                Emu(2_743_200),
            )
            .unwrap();
        shape.set_name("!!Hero").unwrap();
        shape
            .set_fill(
                Fill::from_xml(
                    if slide_index == 0 {
                        br#"<a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:srgbClr val="D63C32"/></a:solidFill>"#
                    } else {
                        br#"<a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:srgbClr val="3278D6"/></a:solidFill>"#
                    },
                )
                .unwrap(),
            )
            .unwrap();
        if slide_index == 1 {
            let mut click_shape = slide
                .add_shape(
                    "rect",
                    Emu(6_858_000),
                    Emu(4_572_000),
                    Emu(914_400),
                    Emu(914_400),
                )
                .unwrap();
            click_shape.set_name("Click fill marker").unwrap();
            click_shape
                .set_fill(
                    Fill::from_xml(
                        br#"<a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:srgbClr val="F0C541"/></a:solidFill>"#,
                    )
                    .unwrap(),
                )
                .unwrap();
        }
    }
    let bytes = presentation.to_bytes().expect("serialize oracle source");
    let mut package = open_opc(&bytes, "PowerPoint timeline oracle source");
    let presentation_part = package.main_document_part().unwrap();
    let model = CT_Presentation::from_xml(package.get_part(&presentation_part).unwrap()).unwrap();
    let slide_parts = model
        .slide_ids
        .iter()
        .map(|slide_id| {
            let relationship = package
                .get_part_rels(&presentation_part)
                .unwrap()
                .get_by_id(&slide_id.relationship_id)
                .unwrap();
            OpcPackage::resolve_rel_target(&presentation_part, &relationship.target)
        })
        .collect::<Vec<_>>();
    let evaluator_part = &slide_parts[0];
    let mut evaluator_slide =
        CT_Slide::from_xml(package.get_part(evaluator_part).unwrap()).unwrap();
    let evaluator_target = evaluator_slide.common_slide_data.shape_tree.children[0]
        .non_visual_id()
        .unwrap();
    let timing = format!(
        r#"<p:timing xmlns:p="{P_NS}"><p:tnLst><p:seq><p:cTn id="100" dur="4000"><p:childTnLst>
        <p:animEffect transition="in" filter="appear"><p:cBhvr><p:cTn id="101" dur="500" fill="hold"/><p:tgtEl><p:spTgt spid="{evaluator_target}"/></p:tgtEl></p:cBhvr></p:animEffect>
        <p:animEffect transition="in" filter="wipe(left)"><p:cBhvr><p:cTn id="102" dur="500" fill="hold"/><p:tgtEl><p:spTgt spid="{evaluator_target}"/></p:tgtEl></p:cBhvr></p:animEffect>
        <p:anim from="100000" to="50000"><p:cBhvr><p:cTn id="103" dur="500" fill="hold"/><p:tgtEl><p:spTgt spid="{evaluator_target}"/></p:tgtEl><p:attrNameLst><p:attrName>style.opacity</p:attrName></p:attrNameLst></p:cBhvr></p:anim>
        <p:par><p:cTn id="104" dur="1000"><p:childTnLst>
        <p:anim from="100000" to="150000"><p:cBhvr><p:cTn id="105" dur="1000" fill="hold"/><p:tgtEl><p:spTgt spid="{evaluator_target}"/></p:tgtEl><p:attrNameLst><p:attrName>scale</p:attrName></p:attrNameLst></p:cBhvr></p:anim>
        <p:anim from="0" to="5400000"><p:cBhvr><p:cTn id="106" dur="1000" fill="hold"/><p:tgtEl><p:spTgt spid="{evaluator_target}"/></p:tgtEl><p:attrNameLst><p:attrName>spin</p:attrName></p:attrNameLst></p:cBhvr></p:anim>
        <p:animMotion origin="layout" path="M 0 0 L .15 .1 E"><p:cBhvr><p:cTn id="107" dur="1000" fill="hold"/><p:tgtEl><p:spTgt spid="{evaluator_target}"/></p:tgtEl></p:cBhvr></p:animMotion>
        </p:childTnLst></p:cTn></p:par>
        <p:animEffect transition="out" filter="fade"><p:cBhvr><p:cTn id="108" dur="500" fill="hold"/><p:tgtEl><p:spTgt spid="{evaluator_target}"/></p:tgtEl></p:cBhvr></p:animEffect>
        </p:childTnLst></p:cTn></p:seq>
        </p:tnLst></p:timing>"#
    );
    let parsed_timing = rpptx_oxml::timing::CT_Timing::from_xml(timing.as_bytes()).unwrap();
    assert_eq!(parsed_timing.nodes().len(), 1);
    evaluator_slide.timing = Some(parsed_timing);
    evaluator_slide.transition = Some(
        rpptx_oxml::timing::CT_SlideTransition::from_xml(
            format!(r#"<p:transition xmlns:p="{P_NS}" advClick="0" advTm="5000"/>"#).as_bytes(),
        )
        .unwrap(),
    );
    package.set_part(evaluator_part, evaluator_slide.to_xml().unwrap());

    let click_part = &slide_parts[1];
    let mut click_slide = CT_Slide::from_xml(package.get_part(click_part).unwrap()).unwrap();
    let click_target = click_slide.common_slide_data.shape_tree.children[1]
        .non_visual_id()
        .unwrap();
    let click_timing = format!(
        r#"<p:timing xmlns:p="{P_NS}"><p:tnLst><p:set><p:cBhvr><p:cTn id="109" dur="500" fill="remove" nodeType="clickEffect"><p:stCondLst><p:cond evt="onClick" delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="{click_target}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="hidden"/></p:to></p:set></p:tnLst></p:timing>"#
    );
    let parsed_click_timing =
        rpptx_oxml::timing::CT_Timing::from_xml(click_timing.as_bytes()).unwrap();
    assert_eq!(parsed_click_timing.nodes().len(), 1);
    let rpptx_oxml::timing::TimingNode::Set(click_node) = &parsed_click_timing.nodes()[0] else {
        panic!("oracle click/fill node must remain typed")
    };
    assert_eq!(
        click_node.target,
        rpptx_oxml::timing::TimingTarget::Shape(click_target)
    );
    click_slide.timing = Some(parsed_click_timing);
    click_slide.transition = Some(
        rpptx_oxml::timing::CT_SlideTransition::from_xml(
            format!(r#"<p:transition xmlns:p="{P_NS}" advClick="0" advTm="2000"/>"#).as_bytes(),
        )
        .unwrap(),
    );
    package.set_part(click_part, click_slide.to_xml().unwrap());

    for (index, effect) in [
        (2usize, "<p:fade/>"),
        (4usize, "<p:push dir=\"l\"/>"),
        (5usize, "<p:zoom dir=\"in\"/>"),
    ] {
        let slide_part = &slide_parts[index];
        let mut slide = CT_Slide::from_xml(package.get_part(slide_part).unwrap()).unwrap();
        let transition = format!(
            r#"<p:transition xmlns:p="{P_NS}" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" advClick="0" advTm="2000" p14:dur="1000">{effect}</p:transition>"#
        );
        slide.transition =
            Some(rpptx_oxml::timing::CT_SlideTransition::from_xml(transition.as_bytes()).unwrap());
        package.set_part(slide_part, slide.to_xml().unwrap());
    }
    let morph_part = &slide_parts[3];
    let morph_slide = String::from_utf8(package.get_part(morph_part).unwrap().to_vec()).unwrap();
    let morph_slide = morph_slide
        .replacen(
            "<p:sld ",
            &format!(
                r#"<p:sld xmlns:mc="{MC_NS}" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" mc:Ignorable="p14 p159" "#
            ),
            1,
        )
        .replacen(
            "</p:sld>",
            r#"<mc:AlternateContent><mc:Choice Requires="p159"><p:transition advClick="0" advTm="2000" p14:dur="1000"><p159:morph option="byObject"/></p:transition></mc:Choice><mc:Fallback><p:transition advClick="0" advTm="2000" p14:dur="1000"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent></p:sld>"#,
            1,
        );
    let morph_slide = CT_Slide::from_xml(morph_slide.as_bytes()).unwrap();
    assert_eq!(
        morph_slide.transition.as_ref().unwrap().effect,
        Some(rpptx_oxml::timing::TransitionEffect::Morph)
    );
    package.set_part(morph_part, morph_slide.to_xml().unwrap());
    let bytes = package_bytes(package);
    let source = Presentation::from_bytes(&bytes).expect("reopen timeline oracle source");
    assert!(source.validate().is_empty());
    bytes
}

fn is_approved_powerpoint_timeline_oracle_source(source: &[u8]) -> bool {
    source == powerpoint_timeline_oracle_source_bytes()
}

fn approved_powerpoint_timeline_oracle_case(case: &str) -> Option<(usize, u64, u32, i64, i32)> {
    F214_POWERPOINT_ORACLE_CASES.iter().find_map(
        |&(
            approved_case,
            slide,
            local_timestamp_ms,
            click_count,
            movie_time_value,
            movie_time_timescale,
        )| {
            (approved_case == case).then_some((
                slide,
                local_timestamp_ms,
                click_count,
                movie_time_value,
                movie_time_timescale,
            ))
        },
    )
}

fn is_approved_powerpoint_timeline_oracle_case(
    case: &str,
    slide: usize,
    local_timestamp_ms: u64,
    click_count: u32,
    movie_time_value: i64,
    movie_time_timescale: i32,
) -> bool {
    approved_powerpoint_timeline_oracle_case(case)
        == Some((
            slide,
            local_timestamp_ms,
            click_count,
            movie_time_value,
            movie_time_timescale,
        ))
}

fn has_approved_powerpoint_fade_outgoing_binding(
    cases: &[(&str, usize, u64, u32, i64, i32)],
) -> bool {
    let outgoing = cases.iter().enumerate().find(|(_, case)| {
        **case
            == (
                "ordinary-transition-outgoing-terminal",
                1,
                u64::MAX,
                u32::MAX,
                4_198,
                600,
            )
    });
    let fade = cases
        .iter()
        .enumerate()
        .find(|(_, case)| **case == ("ordinary-transition", 2, 425, 0, 4_455, 600));
    matches!((outgoing, fade), (Some((outgoing_index, _)), Some((fade_index, _))) if outgoing_index < fade_index)
}

fn has_approved_powerpoint_timeline_oracle_provenance(lines: &[&str]) -> bool {
    F214_POWERPOINT_ORACLE_PROVENANCE.iter().all(|expected| {
        let key = expected.split_once('\t').unwrap().0;
        let matching = lines
            .iter()
            .filter(|line| line.split_once('\t').map(|(field, _)| field) == Some(key))
            .collect::<Vec<_>>();
        matching.len() == 1 && *matching[0] == *expected
    })
}

fn is_approved_powerpoint_timeline_oracle_classification(classification: &str) -> bool {
    classification == F214_POWERPOINT_ORACLE_CLASSIFICATION
}

fn is_approved_powerpoint_timeline_oracle_movie_pin(
    actual_movie_sha256: &str,
    pinned_movie_sha256: &str,
) -> bool {
    pinned_movie_sha256.len() == 64
        && pinned_movie_sha256
            .bytes()
            .all(|byte| byte.is_ascii_digit() || (b'a'..=b'f').contains(&byte))
        && actual_movie_sha256 == pinned_movie_sha256
}

fn extract_powerpoint_movie_frame(
    movie: &Path,
    output: &Path,
    movie_time_value: i64,
    movie_time_timescale: i32,
) -> std::process::Output {
    Command::new("swift")
        .args(["-e", F214_POWERPOINT_FRAME_EXTRACTOR_SWIFT])
        .arg(movie)
        .arg(output)
        .arg(movie_time_value.to_string())
        .arg(movie_time_timescale.to_string())
        .output()
        .expect("extract exact PowerPoint movie frame")
}

fn powerpoint_movie_extractor_returned_exact_time(
    stdout: &[u8],
    expected_width: u32,
    expected_height: u32,
    movie_time_value: i64,
    movie_time_timescale: i32,
) -> bool {
    let Ok(output) = std::str::from_utf8(stdout) else {
        return false;
    };
    let values = output
        .split_whitespace()
        .map(str::parse::<i128>)
        .collect::<Result<Vec<_>, _>>();
    let Ok(values) = values else {
        return false;
    };
    values.len() == 6
        && values[0] == i128::from(expected_width)
        && values[1] == i128::from(expected_height)
        && values[2] == i128::from(movie_time_value)
        && values[3] == i128::from(movie_time_timescale)
        && values[4] == i128::from(movie_time_value)
        && values[5] == i128::from(movie_time_timescale)
        && movie_time_timescale > 0
}

fn is_approved_powerpoint_timeline_oracle_movie_binding(
    fields: &[&str],
    actual_movie_sha256: &str,
    pinned_movie_sha256: &str,
    source_sha256: &str,
) -> bool {
    fields.len() == 7
        && fields[0] == "movie"
        && fields[1] == "timeline-oracle.mp4"
        && fields[2] == actual_movie_sha256
        && fields[2] == pinned_movie_sha256
        && fields[3] == source_sha256
        && fields[4] == POWERPOINT_VERSION
        && fields[5] == POWERPOINT_BUILD
        && fields[6] == POWERPOINT_APP_BUILD
}

fn powerpoint_timeline_oracle_frame_matches_reextraction(
    stored_png: &[u8],
    reextracted_png: &[u8],
) -> bool {
    stored_png == reextracted_png
}

fn normalize_powerpoint_timeline_oracle_frame_for_comparison(
    raw_oracle_png: &[u8],
) -> Option<Vec<u8>> {
    let raw = oxml_media::probe(raw_oracle_png)?;
    if (raw.width_px, raw.height_px)
        != (F214_POWERPOINT_ORACLE_WIDTH, F214_POWERPOINT_ORACLE_HEIGHT)
    {
        return None;
    }
    let page_width = 960.0;
    let page_height = 540.0;
    let image = PositionedElement::Image {
        rect: Rect {
            x: 0.0,
            y: 0.0,
            width: page_width,
            height: page_height,
        },
        data: raw_oracle_png.to_vec(),
        content_type: "image/png".to_owned(),
        media_id: MediaId(1),
    };
    let page = PageFrame::new(1, page_width, page_height, vec![image]);
    let layout = oxml_layout::LayoutResult::new(
        vec![std::sync::Arc::new(page)],
        Vec::new(),
        None,
        Vec::new(),
    );
    let normalized = oxml_pdf::render_page_to_png(&layout, 0, F214_POWERPOINT_COMPARISON_DPI)?;
    let normalized_info = oxml_media::probe(&normalized)?;
    ((normalized_info.width_px, normalized_info.height_px)
        == (
            F214_POWERPOINT_COMPARISON_WIDTH,
            F214_POWERPOINT_COMPARISON_HEIGHT,
        ))
        .then_some(normalized)
}

#[test]
fn powerpoint_timeline_oracle_source_round_trips_through_f213_model() {
    let bytes = powerpoint_timeline_oracle_source_bytes();
    let presentation = Presentation::from_bytes(&bytes).unwrap();
    assert!(presentation.validate().is_empty());

    let at = |elapsed_ms, click_count| {
        presentation
            .render_timeline_deterministic(
                0,
                TimelinePosition {
                    elapsed_ms,
                    click_count,
                },
                None,
            )
            .unwrap()
    };
    let wipe = at(750, 0);
    assert!(wipe.state.shapes.values().next().unwrap().clip.is_some());
    let opacity = at(1_250, 0);
    assert_eq!(opacity.state.shapes.values().next().unwrap().opacity, 0.75);
    let parallel = at(2_000, 0);
    assert!(
        !parallel
            .state
            .shapes
            .values()
            .next()
            .unwrap()
            .transform
            .is_identity()
    );
    let click_at = |elapsed_ms| {
        presentation
            .render_timeline_deterministic(
                1,
                TimelinePosition {
                    elapsed_ms,
                    click_count: 1,
                },
                None,
            )
            .unwrap()
    };
    let click_start = click_at(0);
    assert_eq!(click_start.state.shapes.len(), 1, "{:?}", click_start.state);
    assert!(!click_start.state.shapes.values().next().unwrap().visible);
    assert_eq!(click_at(499).state.shapes.len(), 1);
    assert_eq!(click_at(500).state.shapes.len(), 1);
    assert!(click_at(501).state.shapes.is_empty());

    for (slide, expected) in [
        (2, rpptx_oxml::timing::TransitionEffect::Fade),
        (3, rpptx_oxml::timing::TransitionEffect::Morph),
        (4, rpptx_oxml::timing::TransitionEffect::Push),
        (5, rpptx_oxml::timing::TransitionEffect::Zoom),
    ] {
        let frame = presentation
            .render_timeline_deterministic(
                slide,
                TimelinePosition {
                    elapsed_ms: 500,
                    click_count: 0,
                },
                slide.checked_sub(1),
            )
            .unwrap();
        assert_eq!(frame.state.transition.unwrap().effect, expected);
    }
}

#[test]
fn powerpoint_timeline_oracle_fade_starts_from_the_bound_terminal_outgoing_state() {
    let presentation =
        Presentation::from_bytes(&powerpoint_timeline_oracle_source_bytes()).unwrap();
    let terminal_outgoing = presentation
        .render_timeline_deterministic(
            1,
            TimelinePosition {
                elapsed_ms: u64::MAX,
                click_count: u32::MAX,
            },
            None,
        )
        .unwrap();
    let fade_start = presentation
        .render_timeline_deterministic(
            2,
            TimelinePosition {
                elapsed_ms: 0,
                click_count: 0,
            },
            Some(1),
        )
        .unwrap();
    let render = |page| {
        let layout = oxml_layout::LayoutResult::new(
            vec![std::sync::Arc::new(page)],
            Vec::new(),
            None,
            Vec::new(),
        );
        oxml_pdf::render_page_to_png(&layout, 0, F214_POWERPOINT_COMPARISON_DPI).unwrap()
    };
    assert_eq!(
        render(fade_start.page),
        render(terminal_outgoing.page),
        "fade progress zero must use the externally bound terminal outgoing state"
    );
}

#[test]
fn powerpoint_timeline_oracle_source_binding_rejects_substitution() {
    let source = powerpoint_timeline_oracle_source_bytes();
    assert!(is_approved_powerpoint_timeline_oracle_source(&source));
    let mut substituted = source;
    substituted.push(0);
    assert!(!is_approved_powerpoint_timeline_oracle_source(&substituted));
}

#[test]
fn powerpoint_timeline_oracle_case_binding_rejects_relabelled_or_substituted_coordinates() {
    for &(case, slide, local_timestamp_ms, click_count, movie_time_value, movie_time_timescale) in
        &F214_POWERPOINT_ORACLE_CASES
    {
        assert!(is_approved_powerpoint_timeline_oracle_case(
            case,
            slide,
            local_timestamp_ms,
            click_count,
            movie_time_value,
            movie_time_timescale
        ));
    }

    assert!(
        !is_approved_powerpoint_timeline_oracle_case("morph", 2, 700, 0, 6_336, 600),
        "relabelled transition coordinates must not satisfy another case"
    );
    assert!(
        !is_approved_powerpoint_timeline_oracle_case(
            "timeline-parallel-mid",
            0,
            1_980,
            1,
            1_188,
            600
        ),
        "substituting a click count must be rejected"
    );
    assert!(
        !is_approved_powerpoint_timeline_oracle_case("timeline-wipe-mid", 0, 826, 0, 495, 600),
        "substituting a local timestamp must be rejected"
    );
    assert!(
        !is_approved_powerpoint_timeline_oracle_case("push-transition", 5, 264, 0, 8_118, 600),
        "substituting a generator slide must be rejected"
    );
    assert!(
        !is_approved_powerpoint_timeline_oracle_case("timeline-opacity-mid", 0, 1_320, 0, 791, 600),
        "substituting a movie time value must be rejected"
    );
    assert!(
        !is_approved_powerpoint_timeline_oracle_case("morph", 3, 700, 0, 6_335, 600),
        "a neighboring morph sample must be rejected"
    );
    assert!(
        !is_approved_powerpoint_timeline_oracle_case("push-transition", 4, 264, 0, 8_117, 600),
        "a neighboring push sample must be rejected"
    );
    assert!(
        !is_approved_powerpoint_timeline_oracle_case("morph", 3, 570, 0, 5_742, 600),
        "the pre-morph sample from the uncorrected schedule must be rejected"
    );
    assert!(
        !is_approved_powerpoint_timeline_oracle_case("push-transition", 4, 550, 0, 6_930, 600),
        "the pre-push sample from the uncorrected schedule must be rejected"
    );
    assert!(
        !is_approved_powerpoint_timeline_oracle_case(
            "timeline-opacity-mid",
            0,
            1_320,
            0,
            792,
            1_000
        ),
        "substituting a movie timescale must be rejected"
    );
    assert!(
        !is_approved_powerpoint_timeline_oracle_case(
            "ordinary-transition-outgoing-terminal",
            1,
            u64::MAX,
            0,
            4_198,
            600
        ),
        "the outgoing fade page must use terminal click semantics"
    );
    assert!(
        !is_approved_powerpoint_timeline_oracle_case(
            "ordinary-transition-outgoing-terminal",
            1,
            0,
            u32::MAX,
            4_198,
            600
        ),
        "the outgoing fade page must use terminal elapsed-time semantics"
    );
    assert!(!is_approved_powerpoint_timeline_oracle_case(
        "unapproved-case",
        0,
        0,
        0,
        0,
        0
    ));
}

#[test]
fn powerpoint_timeline_oracle_uses_unique_samples_and_excludes_collapsed_click_boundaries() {
    assert_eq!(
        approved_powerpoint_timeline_oracle_case("timeline-appear-active"),
        Some((0, 330, 0, 198, 600))
    );

    let mut samples = HashSet::new();
    for &(case, slide, local_timestamp_ms, click_count, movie_time_value, movie_time_timescale) in
        &F214_POWERPOINT_ORACLE_CASES
    {
        assert!(
            samples.insert((movie_time_value, movie_time_timescale)),
            "{case}: external cases must use unique encoded movie samples"
        );
        if case == "ordinary-transition-outgoing-terminal" {
            assert_eq!(slide, 1);
            assert_eq!(local_timestamp_ms, u64::MAX);
            assert_eq!(click_count, u32::MAX);
            continue;
        }
        assert_eq!(click_count, 0, "{case}: incoming cases must be automatic");
        if case.starts_with("timeline-") {
            assert_eq!(slide, 0, "{case}: automatic state must use its own slide");
        }
        if case == "morph" {
            assert_eq!((slide, local_timestamp_ms), (3, 700));
            assert_eq!((movie_time_value, movie_time_timescale), (6_336, 600));
            assert!(
                (6_000..6_600).contains(&movie_time_value),
                "the calibrated morph observation must be inside the corrected 10 to 11 second transition window"
            );
            continue;
        }
        if case == "push-transition" {
            assert_eq!((slide, local_timestamp_ms), (4, 264));
            assert_eq!((movie_time_value, movie_time_timescale), (8_118, 600));
            assert!(
                (7_800..8_400).contains(&movie_time_value),
                "the calibrated push observation must be inside the corrected 13 to 14 second transition window"
            );
            continue;
        }
        let slide_movie_start_ms = match slide {
            0 => 0i128,
            2 => 7_000,
            _ => panic!("{case}: unsupported external-oracle slide {slide}"),
        };
        let movie_time_numerator = i128::from(movie_time_value) * 1_000;
        assert_eq!(
            movie_time_numerator % i128::from(movie_time_timescale),
            0,
            "{case}: the selected movie sample must map to an exact Rust millisecond"
        );
        assert_eq!(
            movie_time_numerator / i128::from(movie_time_timescale) - slide_movie_start_ms,
            i128::from(local_timestamp_ms),
            "{case}: Rust local time must correspond to the encoded movie sample"
        );
    }

    for automatic_case in [
        "timeline-appear-active",
        "timeline-wipe-mid",
        "timeline-opacity-mid",
        "timeline-parallel-mid",
        "timeline-exit-mid",
    ] {
        assert!(approved_powerpoint_timeline_oracle_case(automatic_case).is_some());
    }

    for collapsed_case in [
        "timeline-click-start",
        "timeline-fill-remove-before-end",
        "timeline-fill-remove-end",
        "timeline-fill-remove-after-end",
    ] {
        assert_eq!(
            approved_powerpoint_timeline_oracle_case(collapsed_case),
            None,
            "{collapsed_case}: 1 ms evaluator boundaries must not be claimed by a movie frame"
        );
    }
}

#[test]
fn powerpoint_timeline_oracle_excludes_unobservable_zoom_from_identical_source_frames() {
    assert_eq!(
        approved_powerpoint_timeline_oracle_case("zoom-transition"),
        None,
        "the fixed movie must not claim a zoom observation that is byte-identical to static frames"
    );

    let (_, layout) = Presentation::from_bytes(&powerpoint_timeline_oracle_source_bytes())
        .unwrap()
        .render_deterministic()
        .unwrap();
    let outgoing =
        oxml_pdf::render_page_to_png(&layout, 4, F214_POWERPOINT_COMPARISON_DPI).unwrap();
    let incoming =
        oxml_pdf::render_page_to_png(&layout, 5, F214_POWERPOINT_COMPARISON_DPI).unwrap();
    assert_eq!(
        outgoing, incoming,
        "the unchanged source offers no visual discriminator for its zoom transition"
    );
}

#[test]
fn powerpoint_timeline_oracle_fade_requires_the_terminal_outgoing_binding() {
    assert!(has_approved_powerpoint_fade_outgoing_binding(
        &F214_POWERPOINT_ORACLE_CASES
    ));

    let missing = F214_POWERPOINT_ORACLE_CASES
        .iter()
        .copied()
        .filter(|case| case.0 != "ordinary-transition-outgoing-terminal")
        .collect::<Vec<_>>();
    assert!(
        !has_approved_powerpoint_fade_outgoing_binding(&missing),
        "the fade case must not remain approved without its outgoing terminal observation"
    );

    let mut substituted = F214_POWERPOINT_ORACLE_CASES.to_vec();
    let outgoing = substituted
        .iter_mut()
        .find(|case| case.0 == "ordinary-transition-outgoing-terminal")
        .unwrap();
    outgoing.3 = 0;
    assert!(
        !has_approved_powerpoint_fade_outgoing_binding(&substituted),
        "substituted outgoing click semantics must break the fade binding"
    );
}

#[test]
fn powerpoint_timeline_oracle_rejects_substituted_provenance_and_classification() {
    let approved = F214_POWERPOINT_ORACLE_PROVENANCE.to_vec();
    assert!(has_approved_powerpoint_timeline_oracle_provenance(
        &approved
    ));

    let mut substituted = approved.clone();
    substituted[0] = "capture_method\tmanual screenshot";
    assert!(!has_approved_powerpoint_timeline_oracle_provenance(
        &substituted
    ));

    let mut substituted_normalization = approved.clone();
    substituted_normalization[4] = "normalization_operation\tnearest-neighbor resize";
    assert!(!has_approved_powerpoint_timeline_oracle_provenance(
        &substituted_normalization
    ));

    let mut substituted_dimensions = approved.clone();
    substituted_dimensions[6] = "normalization_output_px\t2000x1125";
    assert!(!has_approved_powerpoint_timeline_oracle_provenance(
        &substituted_dimensions
    ));

    let mut duplicated = approved;
    duplicated.push("frame_extractor\tffmpeg");
    assert!(!has_approved_powerpoint_timeline_oracle_provenance(
        &duplicated
    ));
    assert!(is_approved_powerpoint_timeline_oracle_classification(
        F214_POWERPOINT_ORACLE_CLASSIFICATION
    ));
    assert!(!is_approved_powerpoint_timeline_oracle_classification(
        "PowerPoint movie frame at 3000 ms"
    ));
    assert!(!is_approved_powerpoint_timeline_oracle_classification(""));
}

#[test]
fn powerpoint_timeline_oracle_rejects_self_authenticated_movie_and_png_substitutions() {
    const PINNED_MOVIE: &str = "0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdef";
    const SUBSTITUTED_MOVIE: &str =
        "abcdef0123456789abcdef0123456789abcdef0123456789abcdef0123456789";
    const SOURCE: &str = "source-sha256";
    let approved = [
        "movie",
        "timeline-oracle.mp4",
        PINNED_MOVIE,
        SOURCE,
        POWERPOINT_VERSION,
        POWERPOINT_BUILD,
        POWERPOINT_APP_BUILD,
    ];
    assert!(is_approved_powerpoint_timeline_oracle_movie_binding(
        &approved,
        PINNED_MOVIE,
        PINNED_MOVIE,
        SOURCE
    ));

    let mut self_authenticated = approved;
    self_authenticated[2] = SUBSTITUTED_MOVIE;
    assert!(
        !is_approved_powerpoint_timeline_oracle_movie_binding(
            &self_authenticated,
            SUBSTITUTED_MOVIE,
            PINNED_MOVIE,
            SOURCE
        ),
        "a substituted movie and its adjacent manifest hash must not replace the independent pin"
    );
    let mut wrong_source = approved;
    wrong_source[3] = "other-source";
    assert!(!is_approved_powerpoint_timeline_oracle_movie_binding(
        &wrong_source,
        PINNED_MOVIE,
        PINNED_MOVIE,
        SOURCE
    ));
    let mut wrong_version = approved;
    wrong_version[4] = "16.105";
    assert!(!is_approved_powerpoint_timeline_oracle_movie_binding(
        &wrong_version,
        PINNED_MOVIE,
        PINNED_MOVIE,
        SOURCE
    ));

    assert!(powerpoint_timeline_oracle_frame_matches_reextraction(
        b"captured png",
        b"captured png"
    ));
    assert!(
        !powerpoint_timeline_oracle_frame_matches_reextraction(
            b"stale png from another movie time",
            b"re-extracted png"
        ),
        "a stored PNG and its self-supplied manifest hash cannot replace gate-side re-extraction"
    );
}

#[test]
fn powerpoint_timeline_oracle_existing_movie_requires_the_exact_lowercase_pin() {
    const EXPORTED_MOVIE_SHA256: &str =
        "28514432f4aafae9d6c5ddd522d23e87458e08ec59ebf5555398a74e712fa83e";
    assert!(is_approved_powerpoint_timeline_oracle_movie_pin(
        EXPORTED_MOVIE_SHA256,
        EXPORTED_MOVIE_SHA256
    ));
    assert!(!is_approved_powerpoint_timeline_oracle_movie_pin(
        EXPORTED_MOVIE_SHA256,
        "18514432f4aafae9d6c5ddd522d23e87458e08ec59ebf5555398a74e712fa83e"
    ));
    assert!(!is_approved_powerpoint_timeline_oracle_movie_pin(
        EXPORTED_MOVIE_SHA256,
        &EXPORTED_MOVIE_SHA256.to_ascii_uppercase()
    ));
    assert!(!is_approved_powerpoint_timeline_oracle_movie_pin(
        EXPORTED_MOVIE_SHA256,
        "not-a-sha256"
    ));
}

#[test]
fn powerpoint_timeline_oracle_extractor_requires_the_exact_encoded_sample_time() {
    for &(_, _, _, _, movie_time_value, movie_time_timescale) in &F214_POWERPOINT_ORACLE_CASES {
        let exact = format!(
            "{F214_POWERPOINT_ORACLE_WIDTH} {F214_POWERPOINT_ORACLE_HEIGHT} {movie_time_value} {movie_time_timescale} {movie_time_value} {movie_time_timescale}"
        );
        assert!(powerpoint_movie_extractor_returned_exact_time(
            exact.as_bytes(),
            F214_POWERPOINT_ORACLE_WIDTH,
            F214_POWERPOINT_ORACLE_HEIGHT,
            movie_time_value,
            movie_time_timescale
        ));
    }

    let substituted_actual =
        format!("{F214_POWERPOINT_ORACLE_WIDTH} {F214_POWERPOINT_ORACLE_HEIGHT} 495 600 494 600");
    assert!(
        !powerpoint_movie_extractor_returned_exact_time(
            substituted_actual.as_bytes(),
            F214_POWERPOINT_ORACLE_WIDTH,
            F214_POWERPOINT_ORACLE_HEIGHT,
            495,
            600
        ),
        "a neighboring encoded sample must not satisfy the requested rational time"
    );
    let substituted_request =
        format!("{F214_POWERPOINT_ORACLE_WIDTH} {F214_POWERPOINT_ORACLE_HEIGHT} 825 1000 495 600");
    assert!(
        !powerpoint_movie_extractor_returned_exact_time(
            substituted_request.as_bytes(),
            F214_POWERPOINT_ORACLE_WIDTH,
            F214_POWERPOINT_ORACLE_HEIGHT,
            495,
            600
        ),
        "an equivalent millisecond request must not replace the encoded sample representation"
    );
    assert!(
        !powerpoint_movie_extractor_returned_exact_time(
            b"2000 1125 198 600 198 600",
            F214_POWERPOINT_ORACLE_WIDTH,
            F214_POWERPOINT_ORACLE_HEIGHT,
            198,
            600
        ),
        "the unsupported 2000 by 1125 export assumption must be rejected"
    );
    assert!(!powerpoint_movie_extractor_returned_exact_time(
        b"1920 1080 198 0 198 0",
        F214_POWERPOINT_ORACLE_WIDTH,
        F214_POWERPOINT_ORACLE_HEIGHT,
        198,
        0
    ));
}

#[test]
fn powerpoint_timeline_oracle_normalization_is_bounded_deterministic_and_non_mutating() {
    let source_layout = Presentation::from_bytes(&powerpoint_timeline_oracle_source_bytes())
        .unwrap()
        .render_deterministic()
        .unwrap()
        .1;
    let static_before =
        oxml_pdf::render_page_to_png(&source_layout, 0, F214_POWERPOINT_COMPARISON_DPI).unwrap();
    let static_info = oxml_media::probe(&static_before).unwrap();
    assert_eq!(
        (static_info.width_px, static_info.height_px),
        (
            F214_POWERPOINT_COMPARISON_WIDTH,
            F214_POWERPOINT_COMPARISON_HEIGHT
        ),
        "the approved source's unchanged static path must render at literal 150 dpi"
    );

    let raw = oxml_pdf::render_page_to_png(&source_layout, 0, 144.0).unwrap();
    let raw_info = oxml_media::probe(&raw).unwrap();
    assert_eq!(
        (raw_info.width_px, raw_info.height_px),
        (F214_POWERPOINT_ORACLE_WIDTH, F214_POWERPOINT_ORACLE_HEIGHT)
    );
    let first = normalize_powerpoint_timeline_oracle_frame_for_comparison(&raw).unwrap();
    let second = normalize_powerpoint_timeline_oracle_frame_for_comparison(&raw).unwrap();
    assert_eq!(first, second, "normalization output must be deterministic");
    let normalized_info = oxml_media::probe(&first).unwrap();
    assert_eq!(
        (normalized_info.width_px, normalized_info.height_px),
        (
            F214_POWERPOINT_COMPARISON_WIDTH,
            F214_POWERPOINT_COMPARISON_HEIGHT
        )
    );
    assert!(
        normalize_powerpoint_timeline_oracle_frame_for_comparison(&static_before).is_none(),
        "normalization must reject input that is not the exact raw movie geometry"
    );

    let static_after =
        oxml_pdf::render_page_to_png(&source_layout, 0, F214_POWERPOINT_COMPARISON_DPI).unwrap();
    assert_eq!(
        static_before, static_after,
        "oracle normalization must not alter static Rust rendering at 150 dpi"
    );
}

#[test]
fn powerpoint_timeline_oracle_foreground_cutoff_ignores_fixed_opacity_halo() {
    let oracle_dir = std::env::var_os("RDOCX_PPTX_TIMELINE_ORACLE_DIR")
        .map(PathBuf::from)
        .unwrap_or_else(|| workspace_root().join("corpus/pptx-timeline-oracle"));
    let appear = oracle_dir.join("timeline-appear-active-330.png");
    let opacity = oracle_dir.join("timeline-opacity-mid-1320.png");
    if !appear.is_file() || !opacity.is_file() {
        assert!(
            std::env::var_os("RDOCX_PPTX_CORPUS_REQUIRED").is_none(),
            "required fixed opacity-halo oracle frames are missing"
        );
        eprintln!("PowerPoint opacity-halo regression skipped because fixed frames are absent");
        return;
    }
    assert_eq!(
        sha256(&appear),
        "3970c922c4105875bb57cfd61896cafc345c0aaa520434651415a5e9b2f57f16"
    );
    assert_eq!(
        sha256(&opacity),
        "14c836fc873c5080e3834eab04b72e2c93b30e814440cd0f845942faa5d71c69"
    );

    let script = r#"from pathlib import Path
import sys
sys.path.insert(0, sys.argv[1])
from pptx_ssim_harness import decode_png

def foreground_bounds(image, cutoff):
    width, height, rgba = image
    points = []
    for index in range(width * height):
        red, green, blue, alpha = rgba[index * 4:index * 4 + 4]
        red = (red * alpha + 255 * (255 - alpha) + 127) // 255
        green = (green * alpha + 255 * (255 - alpha) + 127) // 255
        blue = (blue * alpha + 255 * (255 - alpha) + 127) // 255
        if min(red, green, blue) < cutoff:
            points.append((index % width, index // width))
    xs = [point[0] for point in points]
    ys = [point[1] for point in points]
    return min(xs), min(ys), max(xs) + 1, max(ys) + 1

appear = decode_png(Path(sys.argv[2]))
opacity = decode_png(Path(sys.argv[3]))
cutoff = int(sys.argv[4])
print(*foreground_bounds(appear, cutoff), *foreground_bounds(opacity, cutoff), *foreground_bounds(appear, 250), *foreground_bounds(opacity, 250))
"#;
    let output = Command::new("python3")
        .args(["-c", script])
        .arg(workspace_root().join("scripts"))
        .arg(&appear)
        .arg(&opacity)
        .arg(F214_POWERPOINT_FOREGROUND_CHANNEL_CUTOFF.to_string())
        .output()
        .unwrap();
    assert!(
        output.status.success(),
        "fixed opacity-halo helper failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let bounds = String::from_utf8(output.stdout)
        .unwrap()
        .split_whitespace()
        .map(str::parse::<u32>)
        .collect::<Result<Vec<_>, _>>()
        .unwrap();
    assert_eq!(bounds.len(), 16);
    assert_eq!(
        &bounds[0..4],
        &bounds[4..8],
        "approved foreground cutoff must keep geometry invariant under opacity"
    );
    assert_ne!(
        &bounds[8..12],
        &bounds[12..16],
        "legacy near-white cutoff must expose the fixed movie's opacity halo"
    );
}

#[test]
fn fade_composition_matches_the_fixed_powerpoint_observation() {
    let oracle_dir = std::env::var_os("RDOCX_PPTX_TIMELINE_ORACLE_DIR")
        .map(PathBuf::from)
        .unwrap_or_else(|| workspace_root().join("corpus/pptx-timeline-oracle"));
    let oracle = oracle_dir.join("ordinary-transition-425.png");
    if !oracle.is_file() {
        assert!(
            std::env::var_os("RDOCX_PPTX_CORPUS_REQUIRED").is_none(),
            "required fixed fade oracle frame is missing"
        );
        eprintln!("PowerPoint fade regression skipped because the fixed frame is absent");
        return;
    }
    assert_eq!(
        sha256(&oracle),
        "61d45d85bd3d6a36361cdad6d438fc60e0a3295e932b582f6e046af16879818d"
    );

    let presentation =
        Presentation::from_bytes(&powerpoint_timeline_oracle_source_bytes()).unwrap();
    let frame = presentation
        .render_timeline_deterministic(
            2,
            TimelinePosition {
                elapsed_ms: 425,
                click_count: 0,
            },
            Some(1),
        )
        .unwrap();
    let layout = oxml_layout::LayoutResult::new(
        vec![std::sync::Arc::new(frame.page)],
        Vec::new(),
        None,
        Vec::new(),
    );
    let rust_png =
        oxml_pdf::render_page_to_png(&layout, 0, F214_POWERPOINT_COMPARISON_DPI).unwrap();
    let normalized_oracle =
        normalize_powerpoint_timeline_oracle_frame_for_comparison(&fs::read(&oracle).unwrap())
            .unwrap();
    let rust_path = std::env::temp_dir().join(format!(
        "rpptx-f214-fixed-fade-rust-{}.png",
        std::process::id()
    ));
    let oracle_path = std::env::temp_dir().join(format!(
        "rpptx-f214-fixed-fade-oracle-{}.png",
        std::process::id()
    ));
    fs::write(&rust_path, rust_png).unwrap();
    fs::write(&oracle_path, normalized_oracle).unwrap();
    let output = Command::new("python3")
        .args([
            "-c",
            "from pathlib import Path; import sys; sys.path.insert(0, sys.argv[1]); from pptx_ssim_harness import decode_png, structural_similarity; print(structural_similarity(decode_png(Path(sys.argv[2])), decode_png(Path(sys.argv[3]))))",
        ])
        .arg(workspace_root().join("scripts"))
        .arg(&rust_path)
        .arg(&oracle_path)
        .output()
        .unwrap();
    fs::remove_file(&rust_path).unwrap();
    fs::remove_file(&oracle_path).unwrap();
    assert!(
        output.status.success(),
        "fixed fade helper failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let score = String::from_utf8(output.stdout)
        .unwrap()
        .trim()
        .parse::<f64>()
        .unwrap();
    assert!(score >= 0.99, "fixed PowerPoint fade SSIM {score}");
}

#[test]
#[ignore = "launches pinned PowerPoint and writes ignored oracle artifacts"]
fn generate_powerpoint_timeline_frame_oracle() {
    let oracle_dir = std::env::var_os("RDOCX_PPTX_TIMELINE_ORACLE_DIR")
        .map(PathBuf::from)
        .unwrap_or_else(|| workspace_root().join("corpus/pptx-timeline-oracle"));
    fs::create_dir_all(&oracle_dir).unwrap();
    let source = oracle_dir.join("timeline-oracle-source.pptx");
    let movie = oracle_dir.join("timeline-oracle.mp4");
    fs::write(&source, powerpoint_timeline_oracle_source_bytes()).unwrap();
    let pinned_movie_sha256 = std::env::var(F214_POWERPOINT_ORACLE_MOVIE_SHA256_ENV).ok();
    if let Some(pinned_movie_sha256) = pinned_movie_sha256.as_deref() {
        assert!(
            movie.is_file(),
            "{F214_POWERPOINT_ORACLE_MOVIE_SHA256_ENV} was supplied but {} is missing",
            movie.display()
        );
        let actual_movie_sha256 = sha256(&movie);
        assert!(
            is_approved_powerpoint_timeline_oracle_movie_pin(
                &actual_movie_sha256,
                pinned_movie_sha256
            ),
            "existing PowerPoint movie SHA-256 does not match {F214_POWERPOINT_ORACLE_MOVIE_SHA256_ENV}"
        );
    } else {
        assert!(
            !movie.exists(),
            "refusing to overwrite {}; supply its exact SHA-256 in {F214_POWERPOINT_ORACLE_MOVIE_SHA256_ENV} to ingest it",
            movie.display()
        );
        assert_powerpoint_build();
        let app = Path::new("/Applications/Microsoft PowerPoint.app");
        assert!(app.is_dir(), "pinned Microsoft PowerPoint is not installed");
        let source_argument = format!(
            "set sourceFile to POSIX file {:?}",
            source.to_string_lossy()
        );
        let movie_argument = format!("set movieFile to POSIX file {:?}", movie.to_string_lossy());
        let width_argument = format!(
            "set «class mFrW» of «class pSMs» of item 1 of presentations to {F214_POWERPOINT_ORACLE_WIDTH}"
        );
        let height_argument = format!(
            "set «class mFrH» of «class pSMs» of item 1 of presentations to {F214_POWERPOINT_ORACLE_HEIGHT}"
        );
        let output = Command::new("osascript")
            .args(["-e", "tell application \"Microsoft PowerPoint\""])
            .args(["-e", &source_argument])
            .args(["-e", &movie_argument])
            .args(["-e", "open sourceFile"])
            .args(["-e", &width_argument])
            .args(["-e", &height_argument])
            .args([
                "-e",
                "save item 1 of presentations in movieFile as save as movie",
            ])
            .args(["-e", "close item 1 of presentations saving no"])
            .args(["-e", "end tell"])
            .output()
            .expect("launch pinned PowerPoint movie export");
        assert!(
            output.status.success(),
            "PowerPoint movie export failed: {}",
            String::from_utf8_lossy(&output.stderr)
        );
    }
    assert!(
        movie.is_file(),
        "PowerPoint did not create {}",
        movie.display()
    );
    assert!(has_approved_powerpoint_fade_outgoing_binding(
        &F214_POWERPOINT_ORACLE_CASES
    ));

    let mut frame_lines = Vec::new();
    for &(case, slide, local_timestamp_ms, click_count, movie_time_value, movie_time_timescale) in
        &F214_POWERPOINT_ORACLE_CASES
    {
        let png = oracle_dir.join(format!("{case}-{local_timestamp_ms}.png"));
        let output =
            extract_powerpoint_movie_frame(&movie, &png, movie_time_value, movie_time_timescale);
        assert!(
            output.status.success(),
            "PowerPoint frame extraction failed: {}",
            String::from_utf8_lossy(&output.stderr)
        );
        assert!(
            powerpoint_movie_extractor_returned_exact_time(
                &output.stdout,
                F214_POWERPOINT_ORACLE_WIDTH,
                F214_POWERPOINT_ORACLE_HEIGHT,
                movie_time_value,
                movie_time_timescale
            ),
            "{case}: extractor did not return the exact requested movie time: {}",
            String::from_utf8_lossy(&output.stdout)
        );
        frame_lines.push(format!(
            "frame\t{case}\t{slide}\t{local_timestamp_ms}\t{click_count}\t{movie_time_value}\t{movie_time_timescale}\t{}\t{}\t{F214_POWERPOINT_ORACLE_WIDTH}\t{F214_POWERPOINT_ORACLE_HEIGHT}\t{F214_POWERPOINT_ORACLE_CLASSIFICATION}",
            png.file_name().unwrap().to_string_lossy(),
            sha256(&png)
        ));
    }
    let source_sha256 = sha256(&source);
    let movie_sha256 = sha256(&movie);
    if let Some(pinned_movie_sha256) = pinned_movie_sha256.as_deref() {
        assert!(is_approved_powerpoint_timeline_oracle_movie_pin(
            &movie_sha256,
            pinned_movie_sha256
        ));
    }
    let manifest = format!(
        "application\tMicrosoft PowerPoint\nversion\t{POWERPOINT_VERSION}\nbundle_build\t{POWERPOINT_BUILD}\napp_build\t{POWERPOINT_APP_BUILD}\n{}\ndpi\t150\ngeometry_tolerance_pt\t1\nssim_threshold\t0.99\nsource\t{}\t{}\nmovie\t{}\t{}\t{}\t{POWERPOINT_VERSION}\t{POWERPOINT_BUILD}\t{POWERPOINT_APP_BUILD}\nframe\tcase\tslide\ttimestamp_ms\tclick_count\tmovie_time_value\tmovie_time_timescale\tpng\tsha256\twidth\theight\tclassification\n{}\n",
        F214_POWERPOINT_ORACLE_PROVENANCE.join("\n"),
        source.file_name().unwrap().to_string_lossy(),
        source_sha256,
        movie.file_name().unwrap().to_string_lossy(),
        movie_sha256,
        source_sha256,
        frame_lines.join("\n")
    );
    fs::write(oracle_dir.join("manifest.tsv"), manifest).unwrap();
}

#[test]
fn timeline_frames_render_through_the_existing_resolved_slide_boundary() {
    let presentation = Presentation::from_bytes(&timeline_fixture_bytes()).unwrap();
    let frame = presentation
        .render_timeline_deterministic(
            0,
            TimelinePosition {
                elapsed_ms: 500,
                click_count: 0,
            },
            None,
        )
        .unwrap();

    assert_eq!(frame.state.shapes.len(), 1);
    assert_eq!(frame.state.shapes.values().next().unwrap().opacity, 0.5);
    let PositionedElement::Group(shape) = &frame.page.elements[0] else {
        panic!("timeline shape should reuse the normal renderer group")
    };
    assert_eq!(shape.opacity, 0.5);
}

#[test]
fn timeline_render_rejects_a_supplied_outgoing_index_out_of_bounds() {
    let presentation = Presentation::from_bytes(&timeline_fixture_bytes()).unwrap();
    let error = presentation
        .render_timeline_deterministic(
            0,
            TimelinePosition {
                elapsed_ms: 500,
                click_count: 0,
            },
            Some(1),
        )
        .unwrap_err();

    assert!(matches!(
        error,
        Error::UnknownSlideIndex {
            index: 1,
            slide_count: 1
        }
    ));
}

#[test]
fn morph_uses_the_timestamped_outgoing_page_that_matches_its_evaluated_state() {
    let bytes = powerpoint_timeline_oracle_source_bytes();
    let mut package = open_opc(&bytes, "timestamped outgoing morph source");
    let presentation_part = package.main_document_part().unwrap();
    let model = CT_Presentation::from_xml(package.get_part(&presentation_part).unwrap()).unwrap();
    let outgoing_relationship = package
        .get_part_rels(&presentation_part)
        .unwrap()
        .get_by_id(&model.slide_ids[2].relationship_id)
        .unwrap();
    let outgoing_part =
        OpcPackage::resolve_rel_target(&presentation_part, &outgoing_relationship.target);
    let outgoing_xml =
        String::from_utf8(package.get_part(&outgoing_part).unwrap().to_vec()).unwrap();
    let outgoing_model = CT_Slide::from_xml(outgoing_xml.as_bytes()).unwrap();
    let shape_id = outgoing_model.common_slide_data.shape_tree.children[0]
        .non_visual_id()
        .unwrap();
    let timing = format!(
        r#"<p:timing><p:tnLst><p:anim from="0" to="100000"><p:cBhvr><p:cTn id="1" dur="1000" fill="hold"/><p:tgtEl><p:spTgt spid="{shape_id}"/></p:tgtEl><p:attrNameLst><p:attrName>style.opacity</p:attrName></p:attrNameLst></p:cBhvr></p:anim><p:cmd type="call" cmd="unsupported"/></p:tnLst></p:timing>"#
    );
    package.set_part(
        &outgoing_part,
        outgoing_xml
            .replacen("</p:sld>", &format!("{timing}</p:sld>"), 1)
            .into_bytes(),
    );

    let presentation = Presentation::from_bytes(&package_bytes(package)).unwrap();
    let outgoing_start = presentation
        .render_timeline_deterministic(
            2,
            TimelinePosition {
                elapsed_ms: 0,
                click_count: 0,
            },
            None,
        )
        .unwrap();
    assert_eq!(outgoing_start.state.shapes[&shape_id].opacity, 0.0);
    let outgoing = presentation
        .render_timeline_deterministic(
            2,
            TimelinePosition {
                elapsed_ms: u64::MAX,
                click_count: u32::MAX,
            },
            None,
        )
        .unwrap();
    assert_eq!(outgoing.state.shapes[&shape_id].opacity, 1.0);
    let frame = presentation
        .render_timeline_deterministic(
            3,
            TimelinePosition {
                elapsed_ms: 0,
                click_count: 0,
            },
            Some(2),
        )
        .unwrap();
    assert!(
        frame
            .diagnostics
            .iter()
            .any(|diagnostic| diagnostic.message == "unsupported timing node cmd"),
        "{:?}",
        frame.diagnostics
    );
    assert!(matches!(
        frame
            .state
            .transition
            .as_ref()
            .map(|transition| &transition.effect),
        Some(rpptx_oxml::timing::TransitionEffect::Morph)
    ));
    let render = |page| {
        let layout = oxml_layout::LayoutResult::new(
            vec![std::sync::Arc::new(page)],
            Vec::new(),
            None,
            Vec::new(),
        );
        oxml_pdf::render_page_to_png(&layout, 0, 150.0).unwrap()
    };
    assert_eq!(
        render(frame.page),
        render(outgoing.page),
        "morph progress zero must equal the terminal evaluated outgoing page"
    );
}

#[test]
fn timeline_facade_retains_resolver_diagnostics() {
    let bytes = timeline_fixture_bytes();
    let mut package = open_opc(&bytes, "timeline resolver diagnostic");
    let presentation_part = package.main_document_part().unwrap();
    let model = CT_Presentation::from_xml(package.get_part(&presentation_part).unwrap()).unwrap();
    let relationship = package
        .get_part_rels(&presentation_part)
        .unwrap()
        .get_by_id(&model.slide_ids[0].relationship_id)
        .unwrap();
    let slide_part = OpcPackage::resolve_rel_target(&presentation_part, &relationship.target);
    let slide = String::from_utf8(package.get_part(&slide_part).unwrap().to_vec()).unwrap();
    let unsupported = slide.replacen("prst=\"rect\"", "prst=\"not-a-preset\"", 1);
    assert_ne!(unsupported, slide);
    package.set_part(&slide_part, unsupported.into_bytes());

    let frame = Presentation::from_bytes(&package_bytes(package))
        .unwrap()
        .render_timeline_deterministic(
            0,
            TimelinePosition {
                elapsed_ms: 500,
                click_count: 0,
            },
            None,
        )
        .unwrap();

    assert!(frame.diagnostics.iter().any(|diagnostic| {
        diagnostic
            .message
            .contains("unknown preset geometry `not-a-preset`")
    }));
}

#[test]
fn ordinary_static_rendering_does_not_execute_the_timeline() {
    let timeline = Presentation::from_bytes(&timeline_fixture_bytes()).unwrap();
    let timeline_layout = timeline.render_deterministic().unwrap().1;

    let mut static_presentation = Presentation::new().unwrap();
    static_presentation.add_slide(0).unwrap();
    static_presentation
        .slide_mut(0)
        .unwrap()
        .add_shape(
            "rect",
            Emu(914_400),
            Emu(914_400),
            Emu(1_828_800),
            Emu(914_400),
        )
        .unwrap();
    let static_layout = static_presentation.render_deterministic().unwrap().1;

    assert_eq!(
        oxml_pdf::render_to_pdf(&timeline_layout),
        oxml_pdf::render_to_pdf(&static_layout)
    );
    assert_eq!(
        oxml_pdf::render_page_to_png(&timeline_layout, 0, 150.0).unwrap(),
        oxml_pdf::render_page_to_png(&static_layout, 0, 150.0).unwrap()
    );
}

fn media_render_fixture(kind: MediaKind, content_type: &str) -> Presentation {
    let mut presentation = Presentation::new().unwrap();
    presentation.add_slide(0).unwrap();
    let poster = valid_one_pixel_png();
    let (bytes, filename) = match kind {
        MediaKind::Audio => (b"ID3f216-audio".as_slice(), "media.mp3"),
        MediaKind::Video => (b"\0\0\0\x18ftypisom-f216-video".as_slice(), "media.mp4"),
    };
    presentation
        .add_media(
            0,
            kind,
            MediaSourceInput::Embedded(EmbeddedMediaInput {
                bytes,
                filename,
                content_type,
            }),
            MediaPoster {
                bytes: &poster,
                filename: "poster.png",
            },
            Emu(914_400),
            Emu(914_400),
            Emu(2_743_200),
            Emu(1_828_800),
            MediaPlaybackSettings {
                trigger: rpptx::MediaPlaybackTrigger::Automatic,
                ..MediaPlaybackSettings::default()
            },
        )
        .unwrap();
    presentation
}

fn page_text(page: &PageFrame) -> Vec<String> {
    let mut labels = Vec::new();
    walk(&page.elements, &mut |element, _| match element {
        PositionedElement::Text(run) => labels.push(run.text.clone()),
        PositionedElement::MultilingualText(run) => labels.push(run.logical_text.clone()),
        _ => {}
    });
    labels
}

fn page_image_count(page: &PageFrame) -> usize {
    let mut count = 0;
    walk(&page.elements, &mut |element, _| {
        if matches!(element, PositionedElement::Image { .. }) {
            count += 1;
        }
    });
    count
}

fn decoded_png_rgba_sha256(png: &[u8]) -> String {
    let path = std::env::temp_dir().join(format!(
        "rpptx-f216-rgba-{}-{}.png",
        std::process::id(),
        MediaId::from_bytes(png).0,
    ));
    fs::write(&path, png).unwrap();
    let script = r#"from pathlib import Path
import hashlib
import sys
sys.path.insert(0, sys.argv[1])
from pptx_ssim_harness import decode_png
_, _, rgba = decode_png(Path(sys.argv[2]))
print(hashlib.sha256(bytes(rgba)).hexdigest())
"#;
    let output = Command::new("python3")
        .args(["-c", script])
        .arg(workspace_root().join("scripts"))
        .arg(&path)
        .output()
        .unwrap();
    fs::remove_file(&path).unwrap();
    assert!(
        output.status.success(),
        "F-216 RGBA decoder failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    String::from_utf8(output.stdout).unwrap().trim().to_owned()
}

#[test]
fn valid_media_posters_resolve_through_the_existing_image_content_path() {
    let presentation = media_render_fixture(MediaKind::Video, "video/mp4");
    let rendered = presentation
        .render_media_timeline_deterministic(
            0,
            TimelinePosition {
                elapsed_ms: 250,
                click_count: 0,
            },
            None,
            MediaFallbackPolicy::PosterFrame,
        )
        .unwrap();

    assert_eq!(page_image_count(&rendered.frame.page), 1);
    assert!(!page_text(&rendered.frame.page).contains(&"Video".to_owned()));
    let (input, _) = presentation.render_deterministic().unwrap();
    assert_eq!(
        input.media.len(),
        1,
        "only the poster enters renderer media"
    );
    assert!(
        !input
            .media
            .contains_key(&MediaId::from_bytes(b"\0\0\0\x18ftypisom-f216-video"))
    );

    let placeholder = presentation
        .render_media_timeline_deterministic(
            0,
            TimelinePosition::default(),
            None,
            MediaFallbackPolicy::DeterministicPlaceholder,
        )
        .unwrap();
    assert!(page_text(&placeholder.frame.page).contains(&"Video".to_owned()));
}

#[test]
fn unsupported_codecs_keep_the_poster_without_attempting_to_decode_the_payload() {
    let presentation = media_render_fixture(MediaKind::Video, "video/x-f216-opaque");
    let media = presentation.media(0).unwrap().remove(0);
    assert!(
        media
            .diagnostics
            .contains(&MediaDiagnostic::UnsupportedContentType {
                content_type: "video/x-f216-opaque".to_owned(),
            })
    );
    let payload = presentation
        .extract_media(0, media.shape_id)
        .unwrap()
        .unwrap();
    let rendered = presentation
        .render_media_timeline_deterministic(
            0,
            TimelinePosition::default(),
            None,
            MediaFallbackPolicy::PosterFrame,
        )
        .unwrap();
    assert_eq!(payload, b"\0\0\0\x18ftypisom-f216-video");
    assert_eq!(page_image_count(&rendered.frame.page), 1);
    assert_eq!(
        rendered
            .frame
            .diagnostics
            .iter()
            .map(|diagnostic| diagnostic.message.clone())
            .collect::<Vec<_>>(),
        vec![format!(
            "media shape {} has unsupported content type `video/x-f216-opaque`",
            media.shape_id
        )]
    );
}

#[test]
fn orphan_incoming_media_commands_keep_their_retained_timing_diagnostic() {
    let presentation = media_render_fixture(MediaKind::Audio, "audio/mpeg");
    let mut package = open_opc(
        &presentation.to_bytes().unwrap(),
        "orphan incoming media command",
    );
    let presentation_part = package.main_document_part().unwrap();
    let model = CT_Presentation::from_xml(package.get_part(&presentation_part).unwrap()).unwrap();
    let slide_relationship = package
        .get_part_rels(&presentation_part)
        .unwrap()
        .get_by_id(&model.slide_ids[0].relationship_id)
        .unwrap();
    let slide_part = OpcPackage::resolve_rel_target(&presentation_part, &slide_relationship.target);
    let slide = String::from_utf8(package.get_part(&slide_part).unwrap().to_vec()).unwrap();
    let orphan_command = r#"<p:cmd type="call" cmd="play"><p:cBhvr><p:cTn id="999" dur="1"/><p:tgtEl><p:spTgt spid="999999"/></p:tgtEl></p:cBhvr></p:cmd>"#;
    let slide = slide.replacen("</p:tnLst>", &format!("{orphan_command}</p:tnLst>"), 1);
    package.set_part(&slide_part, slide.into_bytes());

    let rendered = Presentation::from_bytes(&package_bytes(package))
        .unwrap()
        .render_media_timeline_deterministic(
            0,
            TimelinePosition::default(),
            None,
            MediaFallbackPolicy::PosterFrame,
        )
        .unwrap();

    assert_eq!(rendered.media.len(), 1);
    assert_eq!(
        rendered
            .frame
            .diagnostics
            .iter()
            .filter(|diagnostic| {
                diagnostic.message == "media timing is retained for synchronized playback"
            })
            .count(),
        1
    );
}

#[test]
fn outgoing_media_timing_diagnostics_remain_when_only_incoming_state_is_returned() {
    let mut presentation = media_render_fixture(MediaKind::Video, "video/mp4");
    presentation.add_slide(0).unwrap();

    let rendered = presentation
        .render_media_timeline_deterministic(
            1,
            TimelinePosition::default(),
            Some(0),
            MediaFallbackPolicy::PosterFrame,
        )
        .unwrap();

    assert!(rendered.media.is_empty());
    assert_eq!(
        rendered
            .frame
            .diagnostics
            .iter()
            .filter(|diagnostic| {
                diagnostic.message == "media timing is retained for synchronized playback"
            })
            .count(),
        2
    );
}

#[test]
fn missing_media_posters_remain_visible_and_do_not_change_timeline_siblings() {
    let presentation = media_render_fixture(MediaKind::Audio, "audio/mpeg");
    let poster_relationship_id = presentation.media(0).unwrap()[0]
        .poster_relationship_id
        .clone()
        .unwrap();
    let mut package = open_opc(&presentation.to_bytes().unwrap(), "missing media poster");
    let presentation_part = package.main_document_part().unwrap();
    let model = CT_Presentation::from_xml(package.get_part(&presentation_part).unwrap()).unwrap();
    let slide_relationship = package
        .get_part_rels(&presentation_part)
        .unwrap()
        .get_by_id(&model.slide_ids[0].relationship_id)
        .unwrap();
    let slide_part = OpcPackage::resolve_rel_target(&presentation_part, &slide_relationship.target);
    package
        .part_rels
        .get_mut(&slide_part)
        .unwrap()
        .items
        .retain(|relationship| relationship.id != poster_relationship_id);
    let presentation = Presentation::from_bytes(&package_bytes(package)).unwrap();
    let position = TimelinePosition {
        elapsed_ms: 250,
        click_count: 0,
    };
    let ordinary = presentation
        .render_timeline_deterministic(0, position, None)
        .unwrap();
    let rendered = presentation
        .render_media_timeline_deterministic(0, position, None, MediaFallbackPolicy::PosterFrame)
        .unwrap();

    assert!(page_text(&rendered.frame.page).contains(&"Audio".to_owned()));
    assert_eq!(rendered.frame.state.shapes, ordinary.state.shapes);
    assert_eq!(rendered.frame.state.transition, ordinary.state.transition);
    assert_eq!(rendered.media.len(), 1);
    assert!(rendered.frame.diagnostics.iter().any(|diagnostic| {
        diagnostic
            .message
            .contains("rendered as deterministic Audio placeholder")
    }));
    assert!(
        presentation
            .render_media_timeline_deterministic(0, position, None, MediaFallbackPolicy::Fail,)
            .is_err()
    );
}

#[test]
fn media_fallback_preserves_a_preceding_ordinary_missing_picture_diagnostic() {
    let mut presentation = Presentation::new().unwrap();
    presentation.add_slide(0).unwrap();
    let ordinary_picture = sparse_preview_png();
    presentation
        .add_picture(
            0,
            &ordinary_picture,
            "ordinary.png",
            Emu(0),
            Emu(0),
            None,
            None,
        )
        .unwrap();
    let poster = valid_one_pixel_png();
    presentation
        .add_media(
            0,
            MediaKind::Audio,
            MediaSourceInput::Embedded(EmbeddedMediaInput {
                bytes: b"ID3f216-diagnostic-identity",
                filename: "identity.mp3",
                content_type: "audio/mpeg",
            }),
            MediaPoster {
                bytes: &poster,
                filename: "poster.png",
            },
            Emu(914_400),
            Emu(914_400),
            Emu(2_743_200),
            Emu(1_828_800),
            MediaPlaybackSettings::default(),
        )
        .unwrap();
    let media = presentation.media(0).unwrap().remove(0);
    let poster_relationship_id = media.poster_relationship_id.unwrap();
    let mut package = open_opc(
        &presentation.to_bytes().unwrap(),
        "ordinary picture before media",
    );
    let presentation_part = package.main_document_part().unwrap();
    let model = CT_Presentation::from_xml(package.get_part(&presentation_part).unwrap()).unwrap();
    let slide_relationship = package
        .get_part_rels(&presentation_part)
        .unwrap()
        .get_by_id(&model.slide_ids[0].relationship_id)
        .unwrap();
    let slide_part = OpcPackage::resolve_rel_target(&presentation_part, &slide_relationship.target);
    let slide = CT_Slide::from_xml(package.get_part(&slide_part).unwrap()).unwrap();
    let ordinary_relationship_id = slide
        .common_slide_data
        .shape_tree
        .children
        .iter()
        .find_map(|child| match child {
            ShapeTreeChild::Picture(picture) if picture.media.is_none() => picture
                .blip_fill
                .as_ref()
                .and_then(|fill| fill.blip.as_ref())
                .and_then(|blip| blip.embed.clone()),
            _ => None,
        })
        .unwrap();
    assert_ne!(ordinary_relationship_id, poster_relationship_id);
    package
        .part_rels
        .get_mut(&slide_part)
        .unwrap()
        .items
        .retain(|relationship| {
            relationship.id != ordinary_relationship_id && relationship.id != poster_relationship_id
        });

    let rendered = Presentation::from_bytes(&package_bytes(package))
        .unwrap()
        .render_media_timeline_deterministic(
            0,
            TimelinePosition::default(),
            None,
            MediaFallbackPolicy::PosterFrame,
        )
        .unwrap();
    let ordinary_message =
        format!("missing slide picture image relationship `{ordinary_relationship_id}`");
    let fallback_message = format!(
        "media shape {} rendered as deterministic Audio placeholder",
        media.shape_id
    );

    assert_eq!(
        rendered
            .frame
            .diagnostics
            .iter()
            .filter(|diagnostic| diagnostic.message == ordinary_message)
            .count(),
        1
    );
    assert_eq!(
        rendered
            .frame
            .diagnostics
            .iter()
            .filter(|diagnostic| diagnostic.message == fallback_message)
            .count(),
        1
    );
    assert!(
        !rendered
            .frame
            .diagnostics
            .iter()
            .any(|diagnostic| { diagnostic.message.contains("media poster for shape") })
    );
}

#[test]
fn legacy_render_entry_points_keep_exact_unresolved_poster_diagnostics() {
    let presentation = media_render_fixture(MediaKind::Audio, "audio/mpeg");
    let poster_relationship_id = presentation.media(0).unwrap()[0]
        .poster_relationship_id
        .clone()
        .unwrap();
    let mut package = open_opc(
        &presentation.to_bytes().unwrap(),
        "legacy unresolved media poster diagnostic",
    );
    let presentation_part = package.main_document_part().unwrap();
    let model = CT_Presentation::from_xml(package.get_part(&presentation_part).unwrap()).unwrap();
    let slide_relationship = package
        .get_part_rels(&presentation_part)
        .unwrap()
        .get_by_id(&model.slide_ids[0].relationship_id)
        .unwrap();
    let slide_part = OpcPackage::resolve_rel_target(&presentation_part, &slide_relationship.target);
    package
        .part_rels
        .get_mut(&slide_part)
        .unwrap()
        .items
        .retain(|relationship| relationship.id != poster_relationship_id);
    let presentation = Presentation::from_bytes(&package_bytes(package)).unwrap();
    let expected = format!("missing slide picture image relationship `{poster_relationship_id}`");

    let (input, layout) = presentation.render_deterministic().unwrap();
    assert_eq!(
        input.slides[0]
            .diagnostics
            .iter()
            .map(|diagnostic| diagnostic.message.as_str())
            .collect::<Vec<_>>(),
        vec![expected.as_str()]
    );
    assert_eq!(
        input.slides[0].diagnostics[0].message.as_bytes(),
        expected.as_bytes()
    );
    assert_eq!(
        layout
            .diagnostics
            .iter()
            .map(|diagnostic| diagnostic.message.as_str())
            .collect::<Vec<_>>(),
        vec![expected.as_str()]
    );

    let timeline = presentation
        .render_timeline_deterministic(0, TimelinePosition::default(), None)
        .unwrap();
    assert_eq!(
        timeline
            .diagnostics
            .iter()
            .map(|diagnostic| diagnostic.message.as_str())
            .collect::<Vec<_>>(),
        vec![
            expected.as_str(),
            "media timing is retained for synchronized playback",
            "media timing is retained for synchronized playback",
        ]
    );
    assert_eq!(
        timeline.diagnostics[0].message.as_bytes(),
        expected.as_bytes()
    );
}

#[test]
fn media_fallback_identity_distinguishes_inherited_and_slide_shape_ids() {
    let presentation = media_render_fixture(MediaKind::Audio, "audio/mpeg");
    let media = presentation.media(0).unwrap().remove(0);
    let original_shape_id = media.shape_id;
    let shared_shape_id = 424_242;
    let poster_relationship_id = media.poster_relationship_id.unwrap();
    let mut package = open_opc(
        &presentation.to_bytes().unwrap(),
        "source scoped media poster diagnostic",
    );
    let presentation_part = package.main_document_part().unwrap();
    let model = CT_Presentation::from_xml(package.get_part(&presentation_part).unwrap()).unwrap();
    let slide_relationship = package
        .get_part_rels(&presentation_part)
        .unwrap()
        .get_by_id(&model.slide_ids[0].relationship_id)
        .unwrap();
    let slide_part = OpcPackage::resolve_rel_target(&presentation_part, &slide_relationship.target);
    let layout_relationship = package
        .get_part_rels(&slide_part)
        .unwrap()
        .get_by_type(rel_types::SLIDE_LAYOUT)
        .unwrap();
    let layout_part = OpcPackage::resolve_rel_target(&slide_part, &layout_relationship.target);
    let master_relationship = package
        .get_part_rels(&layout_part)
        .unwrap()
        .get_by_type(rel_types::SLIDE_MASTER)
        .unwrap();
    let master_part = OpcPackage::resolve_rel_target(&layout_part, &master_relationship.target);

    let slide_xml = String::from_utf8(package.get_part(&slide_part).unwrap().to_vec()).unwrap();
    let slide_xml = slide_xml.replacen(
        &format!(r#"<p:cNvPr id="{original_shape_id}""#),
        &format!(r#"<p:cNvPr id="{shared_shape_id}""#),
        1,
    );
    let slide_xml = slide_xml.replace(
        &format!(r#"spid="{original_shape_id}""#),
        &format!(r#"spid="{shared_shape_id}""#),
    );
    let picture_start = slide_xml.find("<p:pic").unwrap();
    let picture_end =
        slide_xml[picture_start..].find("</p:pic>").unwrap() + picture_start + "</p:pic>".len();
    let inherited_poster_relationship_id = "rIdF216MissingMasterPoster";
    let inherited_picture = slide_xml[picture_start..picture_end].replace(
        &format!(r#"r:embed="{poster_relationship_id}""#),
        &format!(r#"r:embed="{inherited_poster_relationship_id}""#),
    );
    package.set_part(&slide_part, slide_xml.into_bytes());
    package
        .part_rels
        .get_mut(&slide_part)
        .unwrap()
        .items
        .retain(|relationship| relationship.id != poster_relationship_id);

    let mut master_xml =
        String::from_utf8(package.get_part(&master_part).unwrap().to_vec()).unwrap();
    if !master_xml.contains("xmlns:p14=") {
        master_xml = master_xml.replacen(
            "<p:sldMaster",
            r#"<p:sldMaster xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main""#,
            1,
        );
    }
    master_xml = master_xml.replacen("</p:spTree>", &format!("{inherited_picture}</p:spTree>"), 1);
    package.set_part(&master_part, master_xml.into_bytes());

    let rendered = Presentation::from_bytes(&package_bytes(package))
        .unwrap()
        .render_media_timeline_deterministic(
            0,
            TimelinePosition::default(),
            None,
            MediaFallbackPolicy::PosterFrame,
        )
        .unwrap();
    let inherited_message =
        format!("missing master picture image relationship `{inherited_poster_relationship_id}`");
    let slide_message =
        format!("missing slide picture image relationship `{poster_relationship_id}`");
    let fallback_message =
        format!("media shape {shared_shape_id} rendered as deterministic Audio placeholder");

    assert_eq!(
        rendered
            .frame
            .diagnostics
            .iter()
            .filter(|diagnostic| diagnostic.message == inherited_message)
            .count(),
        1
    );
    assert_eq!(
        rendered
            .frame
            .diagnostics
            .iter()
            .filter(|diagnostic| diagnostic.message == fallback_message)
            .count(),
        1
    );
    assert!(
        !rendered
            .frame
            .diagnostics
            .iter()
            .any(|diagnostic| diagnostic.message == slide_message)
    );
    assert!(
        !rendered
            .frame
            .diagnostics
            .iter()
            .any(|diagnostic| diagnostic.message.contains("media poster for shape"))
    );
}

#[test]
fn media_timeline_frames_return_the_page_and_synchronized_playback_state_once() {
    let presentation = media_render_fixture(MediaKind::Video, "video/mp4");
    let position = TimelinePosition {
        elapsed_ms: 375,
        click_count: 0,
    };
    let ordinary = presentation
        .render_timeline_deterministic(0, position, None)
        .unwrap();
    let rendered = presentation
        .render_media_timeline_deterministic(0, position, None, MediaFallbackPolicy::PosterFrame)
        .unwrap();

    assert_eq!(rendered.frame.state.shapes, ordinary.state.shapes);
    assert_eq!(rendered.frame.state.transition, ordinary.state.transition);
    assert_eq!(rendered.frame.page.elements, ordinary.page.elements);
    assert_eq!(rendered.media.len(), 1);
    assert_eq!(rendered.media[0].phase, MediaPlaybackPhase::Playing);
    assert_eq!(rendered.media[0].source_position_ms, 375);
    assert_eq!(rendered.media[0].volume, 1.0);
}

fn media_playback_golden_fixture() -> Presentation {
    let mut presentation = Presentation::new().unwrap();
    presentation.add_slide(0).unwrap();
    let poster = valid_one_pixel_png();
    presentation
        .add_media(
            0,
            MediaKind::Audio,
            MediaSourceInput::Embedded(EmbeddedMediaInput {
                bytes: b"ID3f216-playback-golden",
                filename: "playback.mp3",
                content_type: "audio/mpeg",
            }),
            MediaPoster {
                bytes: &poster,
                filename: "poster.png",
            },
            Emu(914_400),
            Emu(914_400),
            Emu(2_743_200),
            Emu(1_828_800),
            MediaPlaybackSettings {
                trim_start_ms: Some(50),
                trim_end_ms: Some(250),
                volume: 25_000,
                looped: true,
                show_when_stopped: false,
                trigger: rpptx::MediaPlaybackTrigger::Automatic,
            },
        )
        .unwrap();
    let shape_id = presentation.media(0).unwrap()[0].shape_id;
    let mut package = open_opc(&presentation.to_bytes().unwrap(), "F-216 playback golden");
    let presentation_part = package.main_document_part().unwrap();
    let model = CT_Presentation::from_xml(package.get_part(&presentation_part).unwrap()).unwrap();
    let slide_relationship = package
        .get_part_rels(&presentation_part)
        .unwrap()
        .get_by_id(&model.slide_ids[0].relationship_id)
        .unwrap();
    let slide_part = OpcPackage::resolve_rel_target(&presentation_part, &slide_relationship.target);
    let slide = String::from_utf8(package.get_part(&slide_part).unwrap().to_vec()).unwrap();
    let commands = format!(
        r#"<p:cmd type="call" cmd="pause"><p:cBhvr><p:cTn id="100" dur="1"><p:stCondLst><p:cond delay="300"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="{shape_id}"/></p:tgtEl></p:cBhvr></p:cmd><p:cmd type="call" cmd="playFrom(0.2)"><p:cBhvr><p:cTn id="101" dur="1"><p:stCondLst><p:cond delay="400"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="{shape_id}"/></p:tgtEl></p:cBhvr></p:cmd><p:cmd type="call" cmd="stop"><p:cBhvr><p:cTn id="102" dur="1"><p:stCondLst><p:cond delay="700"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="{shape_id}"/></p:tgtEl></p:cBhvr></p:cmd>"#
    );
    let slide = slide.replacen("</p:tnLst>", &format!("{commands}</p:tnLst>"), 1);
    package.set_part(&slide_part, slide.into_bytes());
    Presentation::from_bytes(&package_bytes(package)).unwrap()
}

#[test]
fn static_poster_output_and_timestamped_playback_state_match_source_built_oracle_fixtures() {
    let presentation = media_render_fixture(MediaKind::Video, "video/mp4");
    let rendered = presentation
        .render_media_timeline_deterministic(
            0,
            TimelinePosition {
                elapsed_ms: 375,
                click_count: 0,
            },
            None,
            MediaFallbackPolicy::PosterFrame,
        )
        .unwrap();
    let layout = oxml_layout::LayoutResult::new(
        vec![std::sync::Arc::new(rendered.frame.page)],
        Vec::new(),
        None,
        Vec::new(),
    );
    let png = oxml_pdf::render_page_to_png(&layout, 0, 150.0).unwrap();
    let row = format!(
        "shape={} phase={:?} position={} volume={:.2} loop={}",
        rendered.media[0].shape_id,
        rendered.media[0].phase,
        rendered.media[0].source_position_ms,
        rendered.media[0].volume,
        rendered.media[0].looping,
    );

    assert_eq!(
        decoded_png_rgba_sha256(&png),
        "d635fc93667da98961df92e0158c9baf3fb1e3bd0fa05006913bf2e2c55b1cac"
    );
    assert_eq!(
        row,
        "shape=4 phase=Playing position=375 volume=1.00 loop=false"
    );

    let fallback_hashes = [
        (MediaKind::Audio, "audio/mpeg"),
        (MediaKind::Video, "video/mp4"),
    ]
    .into_iter()
    .map(|(kind, content_type)| {
        let presentation = media_render_fixture(kind, content_type);
        let deterministic_fonts = presentation.render_deterministic().unwrap().1.fonts;
        let fallback = presentation
            .render_media_timeline_deterministic(
                0,
                TimelinePosition::default(),
                None,
                MediaFallbackPolicy::DeterministicPlaceholder,
            )
            .unwrap();
        let layout = oxml_layout::LayoutResult::new(
            vec![std::sync::Arc::new(fallback.frame.page)],
            deterministic_fonts,
            None,
            Vec::new(),
        );
        decoded_png_rgba_sha256(&oxml_pdf::render_page_to_png(&layout, 0, 150.0).unwrap())
    })
    .collect::<Vec<_>>();
    assert_eq!(
        fallback_hashes,
        vec![
            "79e8da66b2cedf6a8b37b1eb723d4e278e22076aff5cdcff0616e0e7f2decec5",
            "bca557835b53eb1ac671297c1d641c4f2e335fa51aa5515651ca95cc104e6917",
        ]
    );

    let opaque = media_render_fixture(MediaKind::Video, "video/x-f216-opaque");
    let opaque_frame = opaque
        .render_media_timeline_deterministic(
            0,
            TimelinePosition::default(),
            None,
            MediaFallbackPolicy::PosterFrame,
        )
        .unwrap();
    assert_eq!(
        opaque_frame
            .frame
            .diagnostics
            .iter()
            .map(|diagnostic| diagnostic.message.as_str())
            .collect::<Vec<_>>(),
        vec!["media shape 4 has unsupported content type `video/x-f216-opaque`"]
    );

    let mut unknown_duration = Presentation::new().unwrap();
    unknown_duration.add_slide(0).unwrap();
    unknown_duration
        .add_media(
            0,
            MediaKind::Audio,
            MediaSourceInput::Embedded(EmbeddedMediaInput {
                bytes: b"ID3f216-unknown-duration-loop",
                filename: "unknown.mp3",
                content_type: "audio/mpeg",
            }),
            MediaPoster {
                bytes: &valid_one_pixel_png(),
                filename: "poster.png",
            },
            Emu(914_400),
            Emu(914_400),
            Emu(2_743_200),
            Emu(1_828_800),
            MediaPlaybackSettings {
                volume: 50_000,
                looped: true,
                trigger: rpptx::MediaPlaybackTrigger::Automatic,
                ..MediaPlaybackSettings::default()
            },
        )
        .unwrap();
    let unknown_frame = unknown_duration
        .render_media_timeline_deterministic(
            0,
            TimelinePosition {
                elapsed_ms: 500,
                click_count: 0,
            },
            None,
            MediaFallbackPolicy::PosterFrame,
        )
        .unwrap();
    let unknown_state = &unknown_frame.media[0];
    let unknown_row = format!(
        "shape={} phase={:?} position={} volume={:.2} loop={} diagnostic={}",
        unknown_state.shape_id,
        unknown_state.phase,
        unknown_state.source_position_ms,
        unknown_state.volume,
        unknown_state.looping,
        unknown_frame.frame.diagnostics[0].message,
    );
    assert_eq!(
        unknown_row,
        "shape=4 phase=Playing position=500 volume=0.50 loop=true diagnostic=looping media shape 4 has no finite positive playback interval"
    );

    let playback = media_playback_golden_fixture();
    let rows = [0, 200, 300, 400, 450, 700]
        .into_iter()
        .map(|elapsed_ms| {
            let frame = playback
                .render_media_timeline_deterministic(
                    0,
                    TimelinePosition {
                        elapsed_ms,
                        click_count: 0,
                    },
                    None,
                    MediaFallbackPolicy::PosterFrame,
                )
                .unwrap();
            let state = &frame.media[0];
            format!(
                "{elapsed_ms}\t{:?}\t{}\t{:.2}\t{}",
                state.phase, state.source_position_ms, state.volume, state.looping,
            )
        })
        .collect::<Vec<_>>();
    assert_eq!(
        rows,
        vec![
            "0\tPlaying\t50\t0.25\ttrue",
            "200\tPlaying\t50\t0.25\ttrue",
            "300\tPaused\t150\t0.25\ttrue",
            "400\tPlaying\t200\t0.25\ttrue",
            "450\tPlaying\t50\t0.25\ttrue",
            "700\tStopped\t50\t0.25\ttrue",
        ]
    );

    let mut click = Presentation::new().unwrap();
    click.add_slide(0).unwrap();
    let poster = valid_one_pixel_png();
    click
        .add_media(
            0,
            MediaKind::Audio,
            MediaSourceInput::Embedded(EmbeddedMediaInput {
                bytes: b"ID3f216-click-golden",
                filename: "click.mp3",
                content_type: "audio/mpeg",
            }),
            MediaPoster {
                bytes: &poster,
                filename: "poster.png",
            },
            Emu(914_400),
            Emu(914_400),
            Emu(2_743_200),
            Emu(1_828_800),
            MediaPlaybackSettings {
                trigger: rpptx::MediaPlaybackTrigger::OnClick,
                ..MediaPlaybackSettings::default()
            },
        )
        .unwrap();
    let click_rows = [0, 1]
        .into_iter()
        .map(|click_count| {
            click
                .render_media_timeline_deterministic(
                    0,
                    TimelinePosition {
                        elapsed_ms: 100,
                        click_count,
                    },
                    None,
                    MediaFallbackPolicy::PosterFrame,
                )
                .unwrap()
                .media[0]
                .phase
        })
        .collect::<Vec<_>>();
    assert_eq!(
        click_rows,
        vec![MediaPlaybackPhase::Stopped, MediaPlaybackPhase::Playing]
    );

    let mut linked = Presentation::new().unwrap();
    linked.add_slide(0).unwrap();
    linked
        .add_media(
            0,
            MediaKind::Video,
            MediaSourceInput::Linked {
                target: "https://example.invalid/f216.mp4",
                content_type: "video/mp4",
            },
            MediaPoster {
                bytes: &poster,
                filename: "poster.png",
            },
            Emu(914_400),
            Emu(914_400),
            Emu(2_743_200),
            Emu(1_828_800),
            MediaPlaybackSettings {
                trigger: rpptx::MediaPlaybackTrigger::Automatic,
                ..MediaPlaybackSettings::default()
            },
        )
        .unwrap();
    let linked_info = linked.media(0).unwrap().remove(0);
    assert_eq!(
        linked_info.source,
        rpptx::MediaLocation::Linked {
            target: "https://example.invalid/f216.mp4".to_owned(),
        }
    );
    assert_eq!(linked.extract_media(0, linked_info.shape_id).unwrap(), None);
    let linked_frame = linked
        .render_media_timeline_deterministic(
            0,
            TimelinePosition::default(),
            None,
            MediaFallbackPolicy::PosterFrame,
        )
        .unwrap();
    assert_eq!(page_image_count(&linked_frame.frame.page), 1);
    assert_eq!(linked_frame.media[0].phase, MediaPlaybackPhase::Playing);
}

#[test]
fn pinned_timestamps_match_powerpoint_frame_oracle_within_declared_tolerances() {
    const HEADER: &str = "frame\tcase\tslide\ttimestamp_ms\tclick_count\tmovie_time_value\tmovie_time_timescale\tpng\tsha256\twidth\theight\tclassification";
    let oracle_dir = std::env::var_os("RDOCX_PPTX_TIMELINE_ORACLE_DIR")
        .map(PathBuf::from)
        .unwrap_or_else(|| workspace_root().join("corpus/pptx-timeline-oracle"));
    if !oracle_dir.is_dir() {
        assert!(
            std::env::var_os("RDOCX_PPTX_CORPUS_REQUIRED").is_none(),
            "required PowerPoint timeline oracle is missing at {}",
            oracle_dir.display()
        );
        eprintln!(
            "PowerPoint timeline oracle skipped because {} is absent",
            oracle_dir.display()
        );
        return;
    }
    let manifest_path = oracle_dir.join("manifest.tsv");
    if !manifest_path.is_file() {
        assert!(
            std::env::var_os("RDOCX_PPTX_CORPUS_REQUIRED").is_none(),
            "required PowerPoint timeline oracle manifest is missing at {}",
            manifest_path.display()
        );
        eprintln!(
            "PowerPoint timeline oracle skipped because {} is absent",
            manifest_path.display()
        );
        return;
    }
    let manifest = fs::read_to_string(&manifest_path)
        .unwrap_or_else(|error| panic!("{}: {error}", manifest_path.display()));
    let lines = manifest.lines().collect::<Vec<_>>();
    assert!(lines.contains(&"application\tMicrosoft PowerPoint"));
    assert!(lines.contains(&format!("version\t{POWERPOINT_VERSION}").as_str()));
    assert!(lines.contains(&format!("bundle_build\t{POWERPOINT_BUILD}").as_str()));
    assert!(lines.contains(&format!("app_build\t{POWERPOINT_APP_BUILD}").as_str()));
    assert!(lines.contains(&"dpi\t150"));
    assert!(lines.contains(&"geometry_tolerance_pt\t1"));
    assert!(lines.contains(&"ssim_threshold\t0.99"));
    assert!(
        has_approved_powerpoint_timeline_oracle_provenance(&lines),
        "PowerPoint timeline oracle capture provenance is missing, duplicated, or substituted"
    );
    let source_fields = lines
        .iter()
        .find_map(|line| line.strip_prefix("source\t"))
        .expect("timeline oracle manifest source")
        .split('\t')
        .collect::<Vec<_>>();
    assert_eq!(source_fields.len(), 2);
    let source = oracle_dir.join(source_fields[0]);
    assert_eq!(sha256(&source), source_fields[1]);
    let source_bytes = fs::read(&source).unwrap();
    assert!(
        is_approved_powerpoint_timeline_oracle_source(&source_bytes),
        "PowerPoint timeline oracle source does not match the approved deterministic generator"
    );
    let presentation = Presentation::from_bytes(&source_bytes).unwrap();
    let pinned_movie_sha256 = match std::env::var(F214_POWERPOINT_ORACLE_MOVIE_SHA256_ENV) {
        Ok(value) => value,
        Err(_) => {
            assert!(
                std::env::var_os("RDOCX_PPTX_CORPUS_REQUIRED").is_none(),
                "required PowerPoint timeline oracle movie SHA-256 is not pinned in {F214_POWERPOINT_ORACLE_MOVIE_SHA256_ENV}"
            );
            eprintln!(
                "PowerPoint timeline oracle skipped because {F214_POWERPOINT_ORACLE_MOVIE_SHA256_ENV} is unset"
            );
            return;
        }
    };
    let movie_fields = lines
        .iter()
        .find(|line| line.starts_with("movie\t"))
        .expect("timeline oracle manifest movie")
        .split('\t')
        .collect::<Vec<_>>();
    let movie = oracle_dir.join(movie_fields.get(1).copied().unwrap_or_default());
    let actual_movie_sha256 = sha256(&movie);
    assert!(
        is_approved_powerpoint_timeline_oracle_movie_pin(
            &actual_movie_sha256,
            &pinned_movie_sha256
        ),
        "{F214_POWERPOINT_ORACLE_MOVIE_SHA256_ENV} must be the actual lowercase movie SHA-256"
    );
    assert!(
        is_approved_powerpoint_timeline_oracle_movie_binding(
            &movie_fields,
            &actual_movie_sha256,
            &pinned_movie_sha256,
            source_fields[1]
        ),
        "PowerPoint timeline oracle movie is not independently pinned to the exact source and PowerPoint builds"
    );
    let header_index = lines.iter().position(|line| *line == HEADER).unwrap();
    let frames = &lines[header_index + 1..];
    assert!(!frames.is_empty());
    let mut covered_cases = HashSet::new();
    let mut fade_outgoing_bound = false;
    for (frame_index, line) in frames.iter().enumerate() {
        let fields = line.split('\t').collect::<Vec<_>>();
        assert_eq!(fields.len(), 12, "{}", manifest_path.display());
        assert_eq!(fields[0], "frame");
        let case = fields[1];
        assert!(covered_cases.insert(case), "duplicate oracle case {case}");
        let slide = fields[2].parse::<usize>().unwrap();
        let timestamp_ms = fields[3].parse::<u64>().unwrap();
        let click_count = fields[4].parse::<u32>().unwrap();
        let movie_time_value = fields[5].parse::<i64>().unwrap();
        let movie_time_timescale = fields[6].parse::<i32>().unwrap();
        assert!(
            is_approved_powerpoint_timeline_oracle_case(
                case,
                slide,
                timestamp_ms,
                click_count,
                movie_time_value,
                movie_time_timescale
            ),
            "{case}: manifest slide, local timestamp, click count, and encoded movie sample time must match the approved generator capture"
        );
        if case == "ordinary-transition" {
            assert!(
                fade_outgoing_bound,
                "ordinary-transition requires the preceding verified terminal outgoing observation"
            );
        }
        let oracle_png = oracle_dir.join(fields[7]);
        assert_eq!(sha256(&oracle_png), fields[8]);
        let width = fields[9].parse::<u32>().unwrap();
        let height = fields[10].parse::<u32>().unwrap();
        assert_eq!(
            (width, height),
            (F214_POWERPOINT_ORACLE_WIDTH, F214_POWERPOINT_ORACLE_HEIGHT),
            "{case}: PowerPoint oracle frame dimensions must match the standard movie export"
        );
        assert!(
            is_approved_powerpoint_timeline_oracle_classification(fields[11]),
            "{case}: substituted or unsupported divergence classification"
        );
        let reextracted_path = std::env::temp_dir().join(format!(
            "rpptx-f214-powerpoint-{}-{frame_index}.png",
            std::process::id()
        ));
        let extraction = extract_powerpoint_movie_frame(
            &movie,
            &reextracted_path,
            movie_time_value,
            movie_time_timescale,
        );
        assert!(
            extraction.status.success(),
            "{case}: gate-side PowerPoint frame re-extraction failed: {}",
            String::from_utf8_lossy(&extraction.stderr)
        );
        assert!(
            powerpoint_movie_extractor_returned_exact_time(
                &extraction.stdout,
                width,
                height,
                movie_time_value,
                movie_time_timescale
            ),
            "{case}: gate-side extractor did not return the exact requested movie time: {}",
            String::from_utf8_lossy(&extraction.stdout)
        );
        let oracle_png_bytes = fs::read(&oracle_png).unwrap();
        let reextracted_png_bytes = fs::read(&reextracted_path).unwrap();
        fs::remove_file(&reextracted_path).unwrap();
        assert!(
            powerpoint_timeline_oracle_frame_matches_reextraction(
                &oracle_png_bytes,
                &reextracted_png_bytes
            ),
            "{case}: stored PNG bytes do not match gate-side extraction from the pinned movie sample time"
        );
        let normalized_oracle_png =
            normalize_powerpoint_timeline_oracle_frame_for_comparison(&oracle_png_bytes)
                .unwrap_or_else(|| {
                    panic!(
                        "{case}: verified raw PowerPoint frame could not be normalized to the comparison geometry"
                    )
                });
        let normalized_oracle_path = std::env::temp_dir().join(format!(
            "rpptx-f214-powerpoint-normalized-{}-{frame_index}.png",
            std::process::id()
        ));
        fs::write(&normalized_oracle_path, normalized_oracle_png).unwrap();
        let outgoing_slide_index = if case == "ordinary-transition-outgoing-terminal" {
            None
        } else {
            slide.checked_sub(1)
        };
        let frame = presentation
            .render_timeline_deterministic(
                slide,
                TimelinePosition {
                    elapsed_ms: timestamp_ms,
                    click_count,
                },
                outgoing_slide_index,
            )
            .unwrap();
        let layout = oxml_layout::LayoutResult::new(
            vec![std::sync::Arc::new(frame.page)],
            Vec::new(),
            None,
            Vec::new(),
        );
        let rust_png =
            oxml_pdf::render_page_to_png(&layout, 0, F214_POWERPOINT_COMPARISON_DPI).unwrap();
        let rust_path = std::env::temp_dir().join(format!(
            "rpptx-f214-{}-{frame_index}.png",
            std::process::id()
        ));
        fs::write(&rust_path, rust_png).unwrap();
        let script = r#"from pathlib import Path
import sys
sys.path.insert(0, sys.argv[1])
from pptx_ssim_harness import decode_png, structural_similarity

def foreground_bounds(image, cutoff):
    width, height, rgba = image
    points = []
    for index in range(width * height):
        red, green, blue, alpha = rgba[index * 4:index * 4 + 4]
        red = (red * alpha + 255 * (255 - alpha) + 127) // 255
        green = (green * alpha + 255 * (255 - alpha) + 127) // 255
        blue = (blue * alpha + 255 * (255 - alpha) + 127) // 255
        if min(red, green, blue) < cutoff:
            points.append((index % width, index // width))
    if not points:
        raise ValueError("image has no measurable foreground geometry")
    xs = [point[0] for point in points]
    ys = [point[1] for point in points]
    return min(xs), min(ys), max(xs) + 1, max(ys) + 1

rust = decode_png(Path(sys.argv[2]))
oracle = decode_png(Path(sys.argv[3]))
cutoff = int(sys.argv[4])
rust_bounds = foreground_bounds(rust, cutoff)
oracle_bounds = foreground_bounds(oracle, cutoff)
geometry_error_pt = max(abs(left - right) for left, right in zip(rust_bounds, oracle_bounds)) * 72 / 150
print(rust[0], rust[1], oracle[0], oracle[1], structural_similarity(rust, oracle), geometry_error_pt, *rust_bounds, *oracle_bounds)
"#;
        let output = Command::new("python3")
            .args(["-c", script])
            .arg(workspace_root().join("scripts"))
            .arg(&rust_path)
            .arg(&normalized_oracle_path)
            .arg(F214_POWERPOINT_FOREGROUND_CHANNEL_CUTOFF.to_string())
            .output()
            .unwrap();
        fs::remove_file(&rust_path).unwrap();
        fs::remove_file(&normalized_oracle_path).unwrap();
        assert!(
            output.status.success(),
            "{case}: SSIM helper failed: {}",
            String::from_utf8_lossy(&output.stderr)
        );
        let values = String::from_utf8(output.stdout)
            .unwrap()
            .split_whitespace()
            .map(str::to_owned)
            .collect::<Vec<_>>();
        assert_eq!(values.len(), 14);
        assert_eq!(
            values[0].parse::<u32>().unwrap(),
            F214_POWERPOINT_COMPARISON_WIDTH
        );
        assert_eq!(
            values[1].parse::<u32>().unwrap(),
            F214_POWERPOINT_COMPARISON_HEIGHT
        );
        assert_eq!(
            values[2].parse::<u32>().unwrap(),
            F214_POWERPOINT_COMPARISON_WIDTH
        );
        assert_eq!(
            values[3].parse::<u32>().unwrap(),
            F214_POWERPOINT_COMPARISON_HEIGHT
        );
        let score = values[4].parse::<f64>().unwrap();
        let geometry_error = values[5].parse::<f64>().unwrap();
        let rust_bounds = &values[6..10];
        let oracle_bounds = &values[10..14];
        eprintln!(
            "{case}: geometry_error_pt={geometry_error:.6}, ssim={score:.6}, rust_bounds={}, oracle_bounds={}",
            rust_bounds.join(","),
            oracle_bounds.join(",")
        );
        assert!(
            geometry_error <= 1.0,
            "{case}: measured geometry error {geometry_error} pt, {}",
            fields[11]
        );
        assert!(score >= 0.99, "{case}: SSIM {score}, {}", fields[11]);
        if case == "ordinary-transition-outgoing-terminal" {
            fade_outgoing_bound = true;
        }
    }
    assert!(
        fade_outgoing_bound,
        "PowerPoint timeline oracle must verify the fade's terminal outgoing observation"
    );
    assert_eq!(
        covered_cases,
        F214_POWERPOINT_ORACLE_CASES
            .iter()
            .map(|(case, ..)| *case)
            .collect(),
        "PowerPoint timeline oracle manifest must cover every evaluator and compositor case"
    );
}

#[test]
fn native_presentation_pdfa_method_selects_the_requested_profile() {
    let mut presentation = Presentation::new().expect("create presentation");
    presentation.add_slide(0).expect("add slide");

    for (profile, part) in [
        (oxml_pdf::PdfConformance::PdfA2b, "2"),
        (oxml_pdf::PdfConformance::PdfA3b, "3"),
    ] {
        let pdf = presentation.to_pdfa_deterministic(profile).unwrap();
        assert!(
            String::from_utf8_lossy(&pdf).contains(&format!("<pdfaid:part>{part}</pdfaid:part>"))
        );
    }
}

fn presentation_with_three_dimensional_chart_fallback(preview: Option<(&[u8], &str)>) -> Vec<u8> {
    let bytes = presentation_with_authored_chart().to_bytes().unwrap();
    let mut package = open_opc(&bytes, "three-dimensional chart fallback");
    let chart_part = package
        .content_types
        .overrides
        .iter()
        .find_map(|(part, content_type)| {
            (content_type == content_types::CHART).then_some(part.clone())
        })
        .unwrap();
    package.set_part(
        &chart_part,
        br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:bar3DChart><c:barDir val="col"/></c:bar3DChart></c:plotArea></c:chart></c:chartSpace>"#.to_vec(),
    );
    let slide_part = package
        .content_types
        .overrides
        .iter()
        .find_map(|(part, content_type)| {
            (content_type == content_types::SLIDE).then_some(part.clone())
        })
        .unwrap();
    let preview_target = "../media/f128-chart-preview.png";
    let preview_id = package
        .get_or_create_part_rels(&slide_part)
        .add(rel_types::IMAGE, preview_target);
    let slide_xml = String::from_utf8(package.get_part(&slide_part).unwrap().to_vec()).unwrap();
    let frame_start = slide_xml.find("<p:graphicFrame").unwrap();
    let frame_close = "</p:graphicFrame>";
    let frame_end =
        frame_start + slide_xml[frame_start..].find(frame_close).unwrap() + frame_close.len();
    let frame = &slide_xml[frame_start..frame_end];
    let fallback_picture = format!(
        r#"<p:pic><p:nvPicPr><p:cNvPr/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip xmlns:r="{R_NS}" r:embed="{preview_id}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>"#
    );
    let alternate = format!(
        r#"<mc:AlternateContent xmlns:mc="{MC_NS}"><mc:Choice xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" Requires="c">{frame}</mc:Choice><mc:Fallback>{fallback_picture}</mc:Fallback></mc:AlternateContent>"#
    );
    package.set_part(
        &slide_part,
        format!(
            "{}{}{}",
            &slide_xml[..frame_start],
            alternate,
            &slide_xml[frame_end..]
        )
        .into_bytes(),
    );
    if let Some((preview, content_type)) = preview {
        let preview_part = OpcPackage::resolve_rel_target(&slide_part, preview_target);
        package.set_part(&preview_part, preview.to_vec());
        package
            .content_types
            .add_override(&preview_part, content_type);
    }
    let mut output = Cursor::new(Vec::new());
    package.write_to(&mut output).unwrap();
    output.into_inner()
}

fn modelled_part_bytes(
    package: &OpcPackage,
    deck: &str,
) -> (BTreeMap<String, Vec<u8>>, BTreeMap<&'static str, usize>) {
    let mut overrides = package.content_types.overrides.iter().collect::<Vec<_>>();
    overrides.sort_by_key(|(part_name, _)| *part_name);
    let mut rewritten = BTreeMap::new();
    let mut coverage = BTreeMap::new();
    for (part_name, content_type) in overrides {
        let Some(canonical_type) = MODELLED_CONTENT_TYPES
            .iter()
            .copied()
            .find(|candidate| *candidate == content_type)
        else {
            continue;
        };
        let Some(bytes) = package.get_part(part_name) else {
            panic!("{deck} {part_name}: content-type override has no package part");
        };
        let serialised = serialise_modelled_part(content_type, bytes, deck, part_name);
        if let Some(serialised) = serialised {
            *coverage.entry(canonical_type).or_insert(0) += 1;
            rewritten.insert(part_name.clone(), serialised);
        }
    }
    (rewritten, coverage)
}

fn serialise_modelled_part(
    content_type: &str,
    xml: &[u8],
    deck: &str,
    part_name: &str,
) -> Option<Vec<u8>> {
    let context = || format!("{deck} {part_name}");
    match content_type {
        content_types::PRESENTATION => {
            let parsed = CT_Presentation::from_xml(xml)
                .unwrap_or_else(|error| panic!("{}: parse presentation: {error}", context()));
            let serialised = parsed
                .to_xml()
                .unwrap_or_else(|error| panic!("{}: serialise presentation: {error}", context()));
            let reparsed = CT_Presentation::from_xml(&serialised)
                .unwrap_or_else(|error| panic!("{}: reparse presentation: {error}", context()));
            assert_eq!(parsed, reparsed, "{}: presentation structure", context());
            Some(serialised)
        }
        content_types::SLIDE => {
            let parsed = CT_Slide::from_xml(xml)
                .unwrap_or_else(|error| panic!("{}: parse slide: {error}", context()));
            let serialised = parsed
                .to_xml()
                .unwrap_or_else(|error| panic!("{}: serialise slide: {error}", context()));
            let reparsed = CT_Slide::from_xml(&serialised)
                .unwrap_or_else(|error| panic!("{}: reparse slide: {error}", context()));
            assert_eq!(parsed, reparsed, "{}: slide structure", context());
            Some(serialised)
        }
        content_types::SLIDE_LAYOUT => {
            let parsed = CT_SlideLayout::from_xml(xml)
                .unwrap_or_else(|error| panic!("{}: parse slide layout: {error}", context()));
            let serialised = parsed
                .to_xml()
                .unwrap_or_else(|error| panic!("{}: serialise slide layout: {error}", context()));
            let reparsed = CT_SlideLayout::from_xml(&serialised)
                .unwrap_or_else(|error| panic!("{}: reparse slide layout: {error}", context()));
            assert_eq!(parsed, reparsed, "{}: slide layout structure", context());
            Some(serialised)
        }
        content_types::SLIDE_MASTER => {
            let parsed = CT_SlideMaster::from_xml(xml)
                .unwrap_or_else(|error| panic!("{}: parse slide master: {error}", context()));
            let serialised = parsed
                .to_xml()
                .unwrap_or_else(|error| panic!("{}: serialise slide master: {error}", context()));
            let reparsed = CT_SlideMaster::from_xml(&serialised)
                .unwrap_or_else(|error| panic!("{}: reparse slide master: {error}", context()));
            assert_eq!(parsed, reparsed, "{}: slide master structure", context());
            Some(serialised)
        }
        content_types::NOTES_SLIDE => {
            let parsed = CT_NotesSlide::from_xml(xml)
                .unwrap_or_else(|error| panic!("{}: parse notes slide: {error}", context()));
            let serialised = parsed
                .to_xml()
                .unwrap_or_else(|error| panic!("{}: serialise notes slide: {error}", context()));
            let reparsed = CT_NotesSlide::from_xml(&serialised)
                .unwrap_or_else(|error| panic!("{}: reparse notes slide: {error}", context()));
            assert_eq!(parsed, reparsed, "{}: notes slide structure", context());
            Some(serialised)
        }
        content_types::NOTES_MASTER => {
            let parsed = CT_NotesMaster::from_xml(xml)
                .unwrap_or_else(|error| panic!("{}: parse notes master: {error}", context()));
            let serialised = parsed
                .to_xml()
                .unwrap_or_else(|error| panic!("{}: serialise notes master: {error}", context()));
            let reparsed = CT_NotesMaster::from_xml(&serialised)
                .unwrap_or_else(|error| panic!("{}: reparse notes master: {error}", context()));
            assert_eq!(parsed, reparsed, "{}: notes master structure", context());
            Some(serialised)
        }
        content_types::THEME => {
            let parsed = CT_OfficeStyleSheet::from_xml(xml)
                .unwrap_or_else(|error| panic!("{}: parse theme: {error}", context()));
            let serialised = parsed
                .to_xml()
                .unwrap_or_else(|error| panic!("{}: serialise theme: {error}", context()));
            let reparsed = CT_OfficeStyleSheet::from_xml(&serialised)
                .unwrap_or_else(|error| panic!("{}: reparse theme: {error}", context()));
            assert_eq!(parsed, reparsed, "{}: theme structure", context());
            Some(serialised)
        }
        _ => None,
    }
}

fn reopen_written_package(package: &OpcPackage, deck: &str, action: &str) -> OpcPackage {
    open_opc(&write_package(package, deck, action), deck)
}

fn write_package(package: &OpcPackage, deck: &str, action: &str) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    package
        .write_to(&mut output)
        .unwrap_or_else(|error| panic!("{deck}: {action}: {error}"));
    output.into_inner()
}

fn assert_packages_equal(expected: &OpcPackage, actual: &OpcPackage, deck: &str) {
    assert_eq!(
        expected.content_types.defaults, actual.content_types.defaults,
        "{deck}: content-type defaults"
    );
    assert_eq!(
        expected.content_types.overrides, actual.content_types.overrides,
        "{deck}: content-type overrides"
    );
    assert_eq!(
        expected.package_rels.items, actual.package_rels.items,
        "{deck}: package relationships"
    );
    let mut expected_relationship_parts = expected.part_rels.keys().collect::<Vec<_>>();
    let mut actual_relationship_parts = actual.part_rels.keys().collect::<Vec<_>>();
    expected_relationship_parts.sort();
    actual_relationship_parts.sort();
    assert_eq!(
        expected_relationship_parts, actual_relationship_parts,
        "{deck}: relationship part names"
    );
    for part_name in expected_relationship_parts {
        assert_eq!(
            expected.part_rels[part_name].items, actual.part_rels[part_name].items,
            "{deck} {part_name}: relationships"
        );
    }
    assert_eq!(
        expected.parts.len(),
        actual.parts.len(),
        "{deck}: part count"
    );
    let mut expected_part_names = expected.parts.keys().collect::<Vec<_>>();
    let mut actual_part_names = actual.parts.keys().collect::<Vec<_>>();
    expected_part_names.sort();
    actual_part_names.sort();
    assert_eq!(expected_part_names, actual_part_names, "{deck}: part names");
    for part_name in expected_part_names {
        assert_eq!(
            expected.parts[part_name], actual.parts[part_name],
            "{deck} {part_name}: part bytes"
        );
    }
}

fn read_surface(presentation: &Presentation) -> String {
    let mut output = String::new();
    for (slide_index, slide) in presentation.slides().enumerate() {
        writeln!(
            output,
            "slide\t{slide_index}\t{}\t{:?}\t{:?}\t{:?}",
            slide.id(),
            slide.name(),
            slide.text(),
            slide.notes_text()
        )
        .unwrap();
        for (shape_index, shape) in slide.shapes().enumerate() {
            append_shape_surface(&mut output, shape, &shape_index.to_string());
        }
    }
    output
}

fn append_shape_surface(output: &mut String, shape: ShapeRef<'_>, path: &str) {
    writeln!(
        output,
        "shape\t{path}\t{:?}\t{:?}",
        shape.kind(),
        shape.text()
    )
    .unwrap();
    for (child_index, child) in shape.children().enumerate() {
        append_shape_surface(output, child, &format!("{path}.{child_index}"));
    }
}

fn build_f116_ten_slide_deck() -> Presentation {
    let mut presentation = Presentation::new().expect("open bundled template");
    presentation
        .set_slide_size(Emu(12_800_000), Emu(7_200_000))
        .expect("set F-116 slide size");
    let core = presentation.core_properties_mut();
    core.title = Some("rdocx M11 cross-viewer acceptance".to_owned());
    core.creator = Some("rdocx deterministic acceptance generator".to_owned());
    core.subject = Some("F-107 through F-115".to_owned());

    for number in 1..=10 {
        presentation.add_slide(0).expect("add F-116 slide");
        let mut slide = presentation.slide_mut(number - 1).unwrap();
        slide
            .add_textbox(Emu(450_000), Emu(220_000), Emu(5_600_000), Emu(650_000))
            .expect("add slide label")
            .set_text(&format!("F-116 slide {number:02}"))
            .expect("set slide label");
    }

    let blue_fill = Fill::from_xml(
        br#"<a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:srgbClr val="2F5597"/></a:solidFill>"#,
    )
    .unwrap();
    let green_fill = Fill::from_xml(
        br#"<a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:srgbClr val="70AD47"/></a:solidFill>"#,
    )
    .unwrap();
    let line = CT_LineProperties::from_xml(
        br#"<a:ln xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" w="25400"><a:solidFill><a:srgbClr val="203864"/></a:solidFill></a:ln>"#,
    )
    .unwrap();

    {
        let mut slide = presentation.slide_mut(1).unwrap();
        let mut shape = slide
            .add_shape(
                "roundRect",
                Emu(900_000),
                Emu(1_250_000),
                Emu(3_200_000),
                Emu(1_650_000),
            )
            .expect("add mutable shape");
        shape.set_position(Emu(1_000_000), Emu(1_350_000)).unwrap();
        shape.set_size(Emu(3_400_000), Emu(1_800_000)).unwrap();
        shape.set_rotation(Angle(600_000)).unwrap();
        shape.set_name("M11 mutable shape").unwrap();
        shape.set_fill(blue_fill.clone()).unwrap();
        shape.set_line(line).unwrap();
        shape.set_adjust_value("adj", 30_000.0).unwrap();
        shape
            .set_text("Position, size, rotation, fill, line, and adjustment")
            .unwrap();
    }

    {
        let mut slide = presentation.slide_mut(2).unwrap();
        slide
            .add_textbox(Emu(700_000), Emu(1_200_000), Emu(2_900_000), Emu(800_000))
            .unwrap()
            .set_text("Textbox constructor")
            .unwrap();
        slide
            .add_shape(
                "triangle",
                Emu(4_000_000),
                Emu(1_200_000),
                Emu(1_800_000),
                Emu(1_800_000),
            )
            .unwrap();
        slide
            .add_connector(
                ConnectorType::Straight,
                Emu(900_000),
                Emu(3_600_000),
                Emu(3_000_000),
                Emu(4_300_000),
            )
            .unwrap();
        slide
            .add_connector(
                ConnectorType::Elbow,
                Emu(3_300_000),
                Emu(3_600_000),
                Emu(5_400_000),
                Emu(4_300_000),
            )
            .unwrap();
        slide
            .add_connector(
                ConnectorType::Curve,
                Emu(5_700_000),
                Emu(3_600_000),
                Emu(7_800_000),
                Emu(4_300_000),
            )
            .unwrap();
        slide.add_group_shape().unwrap();
    }

    let png = valid_one_pixel_png();
    presentation
        .add_picture(
            3,
            &png,
            "f116-pixel.png",
            Emu(1_000_000),
            Emu(1_300_000),
            Some(Emu(2_000_000)),
            Some(Emu(2_000_000)),
        )
        .expect("add F-116 picture");

    {
        let mut slide = presentation.slide_mut(4).unwrap();
        let mut shape = slide
            .add_textbox(Emu(800_000), Emu(1_250_000), Emu(7_000_000), Emu(2_600_000))
            .unwrap();
        shape.set_text("Formatted paragraph").unwrap();
        let mut frame = shape.text_frame().unwrap();
        let mut paragraph = frame.paragraph_mut(0).unwrap();
        let mut paragraph_properties = CT_TextParagraphProperties::default();
        paragraph_properties.level = Some(1);
        paragraph.set_properties(paragraph_properties);
        paragraph.set_bullet(Some(TextBullet {
            choice: Some(TextBulletChoice::Character(
                TextBulletCharacter::new("•").unwrap(),
            )),
            ..TextBullet::default()
        }));
        let mut run = paragraph.add_run(" with a bold Carlito run");
        let mut run_properties = CT_TextCharacterProperties::default();
        run_properties.font_size = Some(2_000);
        run_properties.bold = Some(true);
        run.set_properties(run_properties);
        run.set_font(Some(TextFont::new("Carlito").unwrap()));
        frame
            .add_paragraph()
            .set_text("Second paragraph exercises structural append");
    }

    {
        let mut slide = presentation.slide_mut(5).unwrap();
        let mut shape = slide
            .add_table(
                3,
                3,
                Emu(750_000),
                Emu(1_250_000),
                Emu(7_500_000),
                Emu(3_600_000),
            )
            .expect("add F-116 table");
        let mut table = shape.table_mut().unwrap();
        table.set_first_row(true);
        table.set_last_column(true);
        table.set_horizontal_banding(true);
        table.set_vertical_banding(true);
        table.set_column_width(2, Emu(2_700_000)).unwrap();
        for row in 0..3 {
            for column in 0..3 {
                table.cell_mut(row, column).unwrap().set_text(&format!(
                    "R{}C{}",
                    row + 1,
                    column + 1
                ));
            }
        }
        let mut first = table.cell_mut(0, 0).unwrap();
        first.set_fill(Some(green_fill.clone()));
        first.set_margins(
            Some(Emu(72_000)),
            Some(Emu(72_000)),
            Some(Emu(36_000)),
            Some(Emu(36_000)),
        );
        first.merge_to(1, 1).unwrap();
        table.cell_mut(2, 1).unwrap().merge_to(2, 2).unwrap();
        table.cell_mut(2, 1).unwrap().split().unwrap();
    }

    {
        let mut slide = presentation.slide_mut(6).unwrap();
        slide.set_hidden(true);
        slide
            .set_background(green_fill.clone())
            .expect("set F-116 background");
    }

    presentation
        .add_picture(
            7,
            &png,
            "f116-reused-pixel.png",
            Emu(1_000_000),
            Emu(1_300_000),
            Some(Emu(1_500_000)),
            Some(Emu(1_500_000)),
        )
        .expect("add duplicated-slide picture");

    {
        let mut slide = presentation.slide_mut(8).unwrap();
        slide.set_background(blue_fill).unwrap();
        slide.clear_background();
    }

    presentation
        .duplicate_slide(7)
        .expect("duplicate relationship-bearing slide");
    presentation
        .remove_slide(0)
        .expect("remove collection slide");
    presentation
        .move_slide(9, 0)
        .expect("move final slide first");
    presentation
}

fn assert_f116_deck_structure(presentation: &Presentation) {
    assert_eq!(presentation.len(), 10);
    assert!(
        presentation.validate().is_empty(),
        "F-116 validation issues: {:?}",
        presentation.validate()
    );
    assert_eq!(
        presentation.slide_size(),
        Some((Emu(12_800_000), Emu(7_200_000)))
    );
    assert_eq!(
        presentation
            .core_properties()
            .and_then(|properties| properties.title.as_deref()),
        Some("rdocx M11 cross-viewer acceptance")
    );
    for (slide, expected_title) in presentation.slides().zip(F116_FINAL_TITLES) {
        assert!(
            slide.text().contains(expected_title),
            "slide {} lacks {expected_title:?}",
            slide.id()
        );
    }
    assert_eq!(
        presentation
            .slides()
            .map(|slide| slide.id())
            .collect::<HashSet<_>>()
            .len(),
        10
    );
    assert!(presentation.slide(6).unwrap().hidden());
    assert!(presentation.slide(6).unwrap().has_explicit_background());
    let constructor_kinds = presentation
        .slide(2)
        .unwrap()
        .shapes()
        .map(|shape| shape.kind())
        .collect::<Vec<_>>();
    for kind in [ShapeKind::Shape, ShapeKind::Connector, ShapeKind::Group] {
        assert!(constructor_kinds.contains(&kind));
    }
    assert!(
        presentation
            .slide(5)
            .unwrap()
            .shapes()
            .any(|shape| shape.table().is_some())
    );
    assert_eq!(
        presentation
            .slides()
            .map(|slide| {
                slide
                    .shapes()
                    .filter(|shape| shape.kind() == ShapeKind::Picture)
                    .count()
            })
            .sum::<usize>(),
        3
    );
}

fn write_f116_candidate() -> PathBuf {
    let presentation = build_f116_ten_slide_deck();
    assert_f116_deck_structure(&presentation);
    let path = PathBuf::from(F116_CANDIDATE_PATH);
    presentation
        .save(&path)
        .unwrap_or_else(|error| panic!("{}: {error}", path.display()));
    path
}

fn write_f116_temporary_candidate(label: &str) -> PathBuf {
    let presentation = build_f116_ten_slide_deck();
    assert_f116_deck_structure(&presentation);
    let path = f116_temp_path(label, "pptx");
    presentation
        .save(&path)
        .unwrap_or_else(|error| panic!("{}: {error}", path.display()));
    path
}

fn f116_temp_path(label: &str, extension: &str) -> PathBuf {
    let ordinal = F116_TEMP_FILE_COUNTER.fetch_add(1, Ordering::Relaxed);
    std::env::temp_dir().join(format!(
        "rdocx-f116-{label}-{}-{ordinal}.{extension}",
        std::process::id()
    ))
}

fn assert_clean_cross_viewer_evidence(row: CrossViewerEvidence) -> Result<(), String> {
    let CrossViewerObservation::Clean {
        opened_or_imported,
        observed_slide_count,
        no_repair_or_conversion_error,
        export_or_close_result,
    } = row.observation
    else {
        return Err(format!("{} evidence is pending", row.application));
    };
    if row
        .version_or_date
        .is_none_or(|value| value.trim().is_empty())
    {
        return Err(format!(
            "{} lacks a nonblank version or acceptance date",
            row.application
        ));
    }
    if row.build.is_none_or(|value| value.trim().is_empty()) {
        return Err(format!("{} lacks a nonblank build", row.application));
    }
    if !opened_or_imported {
        return Err(format!(
            "{} lacks a positive open or import",
            row.application
        ));
    }
    if observed_slide_count != 10 {
        return Err(format!(
            "{} observed {observed_slide_count} slides instead of 10",
            row.application
        ));
    }
    if !no_repair_or_conversion_error {
        return Err(format!(
            "{} observed a repair or conversion error",
            row.application
        ));
    }
    if export_or_close_result.trim().is_empty() {
        return Err(format!(
            "{} lacks an export or close result",
            row.application
        ));
    }
    Ok(())
}

fn assert_f116_powerpoint_acceptance(path: &Path) {
    assert_powerpoint_build();
    let name = path.file_name().unwrap().to_string_lossy();
    let path = path.to_string_lossy();
    let script = format!(
        "with timeout of 120 seconds\ntell application \"Microsoft PowerPoint\"\nset previousStartUpDialog to start up dialog\ntry\nif (Version as text) is not \"{POWERPOINT_VERSION}\" then error \"PowerPoint version mismatch\"\nif (build as text) is not \"{POWERPOINT_APP_BUILD}\" then error \"PowerPoint build mismatch\"\nset start up dialog to false\nset deckPath to \"{path}\"\nopen my POSIX file deckPath\nset checkedDeck to presentation \"{name}\"\nif (count of slides of checkedDeck) is not 10 then error \"slide count mismatch\"\nclose checkedDeck saving no\nset start up dialog to previousStartUpDialog\non error errorMessage number errorNumber\ntry\nclose checkedDeck saving no\nend try\nset start up dialog to previousStartUpDialog\nerror errorMessage number errorNumber\nend try\nend tell\nend timeout\n"
    );
    let result = Command::new("osascript")
        .args(["-e", &script])
        .output()
        .expect("launch PowerPoint F-116 acceptance script");
    assert!(
        result.status.success(),
        "PowerPoint F-116 acceptance failed: {}",
        String::from_utf8_lossy(&result.stderr)
    );
}

fn assert_f116_libreoffice_acceptance(path: &Path) {
    let version = Command::new("soffice")
        .arg("--version")
        .output()
        .expect("read LibreOffice version");
    assert!(version.status.success());
    assert_eq!(
        String::from_utf8(version.stdout).unwrap().trim(),
        LIBREOFFICE_VERSION
    );

    let output_dir = PathBuf::from(format!(
        "/private/tmp/rdocx-f116-libreoffice-{}",
        std::process::id()
    ));
    let profile_dir = PathBuf::from(format!(
        "/private/tmp/rdocx-f116-libreoffice-profile-{}",
        std::process::id()
    ));
    if output_dir.exists() {
        fs::remove_dir_all(&output_dir).expect("remove stale F-116 LibreOffice output");
    }
    if profile_dir.exists() {
        fs::remove_dir_all(&profile_dir).expect("remove stale F-116 LibreOffice profile");
    }
    fs::create_dir_all(&output_dir).expect("create F-116 LibreOffice output");
    let profile = format!("-env:UserInstallation=file://{}", profile_dir.display());
    let filter =
        r#"pdf:impress_pdf_Export:{"ExportHiddenSlides":{"type":"boolean","value":"true"}}"#;
    let result = Command::new("soffice")
        .args(["--headless", &profile, "--convert-to", filter, "--outdir"])
        .arg(&output_dir)
        .arg(path)
        .output()
        .expect("run LibreOffice F-116 acceptance");
    assert!(
        result.status.success(),
        "LibreOffice F-116 import or export failed: {}",
        String::from_utf8_lossy(&result.stderr)
    );
    let pdf = output_dir.join("rdocx-f116-m11-write-api.pdf");
    assert!(
        pdf.is_file(),
        "LibreOffice did not create {}",
        pdf.display()
    );
    let info = Command::new("pdfinfo")
        .arg(&pdf)
        .output()
        .expect("inspect LibreOffice F-116 PDF");
    assert!(
        info.status.success(),
        "pdfinfo failed: {}",
        String::from_utf8_lossy(&info.stderr)
    );
    let info = String::from_utf8(info.stdout).unwrap();
    let pages = info
        .lines()
        .find_map(|line| line.strip_prefix("Pages:").map(str::trim))
        .and_then(|value| value.parse::<usize>().ok());
    assert_eq!(pages, Some(10));
    fs::remove_dir_all(&output_dir).expect("remove F-116 LibreOffice output");
    if profile_dir.exists() {
        fs::remove_dir_all(&profile_dir).expect("remove F-116 LibreOffice profile");
    }
}

fn sha256(path: &Path) -> String {
    let output = Command::new("shasum")
        .args(["-a", "256"])
        .arg(path)
        .output()
        .unwrap_or_else(|error| panic!("{}: run shasum: {error}", path.display()));
    assert!(
        output.status.success(),
        "{}: shasum failed: {}",
        path.display(),
        String::from_utf8_lossy(&output.stderr)
    );
    String::from_utf8(output.stdout)
        .unwrap()
        .split_whitespace()
        .next()
        .unwrap()
        .to_owned()
}

fn open_package(package: OpcPackage) -> Result<Presentation, Error> {
    Presentation::from_bytes(&package_bytes(package))
}

fn assert_powerpoint_build() {
    let app = "/Applications/Microsoft PowerPoint.app/Contents/Info.plist";
    let version = Command::new("/usr/libexec/PlistBuddy")
        .args(["-c", "Print :CFBundleShortVersionString", app])
        .output()
        .expect("read PowerPoint version");
    let build = Command::new("/usr/libexec/PlistBuddy")
        .args(["-c", "Print :CFBundleVersion", app])
        .output()
        .expect("read PowerPoint build");
    let app_build = Command::new("osascript")
        .args([
            "-e",
            "tell application \"Microsoft PowerPoint\" to return build as text",
        ])
        .output()
        .expect("read PowerPoint AppleScript build");
    assert!(version.status.success());
    assert!(build.status.success());
    assert!(
        app_build.status.success(),
        "PowerPoint AppleScript build query failed: {}",
        String::from_utf8_lossy(&app_build.stderr)
    );
    assert_eq!(
        String::from_utf8(version.stdout).unwrap().trim(),
        POWERPOINT_VERSION
    );
    assert_eq!(
        String::from_utf8(build.stdout).unwrap().trim(),
        POWERPOINT_BUILD
    );
    assert_eq!(
        String::from_utf8(app_build.stdout).unwrap().trim(),
        POWERPOINT_APP_BUILD
    );
}

fn fixture_bytes() -> Vec<u8> {
    package_bytes(fixture_package())
}

fn collaboration_fixture_package() -> OpcPackage {
    let mut package = fixture_package();
    package.set_part(
        "/custom/notes/master.xml",
        format!(r#"<p:notesMaster xmlns:p="{P_NS}" xmlns:a="{A_NS}"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld><p:clrMap accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" bg1="lt1" bg2="lt2" folHlink="folHlink" hlink="hlink" tx1="dk1" tx2="dk2"/><p:hf sldNum="1"/></p:notesMaster>"#).into_bytes(),
    );
    package.set_part(
        "/custom/handouts/master.xml",
        format!(r#"<p:handoutMaster xmlns:p="{P_NS}" xmlns:a="{A_NS}"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld><p:clrMap accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" bg1="lt1" bg2="lt2" folHlink="folHlink" hlink="hlink" tx1="dk1" tx2="dk2"/><p:hf hdr="1"/></p:handoutMaster>"#).into_bytes(),
    );
    package
        .content_types
        .add_override("/custom/notes/master.xml", content_types::NOTES_MASTER);
    package
        .content_types
        .add_override("/custom/handouts/master.xml", content_types::HANDOUT_MASTER);
    package
        .get_or_create_part_rels(PRESENTATION_PART)
        .add(rel_types::NOTES_MASTER, "notes/master.xml");
    package
        .get_or_create_part_rels(PRESENTATION_PART)
        .add(rel_types::HANDOUT_MASTER, "handouts/master.xml");
    package
}

#[test]
fn modern_comments_replies_sections_and_handout_settings_survive_ordered_mutation_save_and_reopen()
{
    use rpptx::{Comment, CommentAuthor, CommentReply, Section};

    let mut presentation =
        open_package(collaboration_fixture_package()).expect("open collaboration fixture");
    presentation
        .add_comment_author(
            CommentAuthor::new(
                "{11111111-1111-1111-1111-111111111111}",
                "Ada",
                Some("A"),
                "ada@example.test",
                "test",
            )
            .unwrap(),
        )
        .unwrap();
    presentation
        .add_comment(
            0,
            Comment::new(
                "{22222222-2222-2222-2222-222222222222}",
                "{11111111-1111-1111-1111-111111111111}",
                "2026-08-29T10:11:12Z",
                "first",
            )
            .unwrap(),
        )
        .unwrap();
    presentation
        .reply_to_comment(
            0,
            "{22222222-2222-2222-2222-222222222222}",
            CommentReply::new(
                "{33333333-3333-3333-3333-333333333333}",
                "{11111111-1111-1111-1111-111111111111}",
                "2026-08-29T10:12:13Z",
                "reply",
            )
            .unwrap(),
        )
        .unwrap();
    presentation
        .reply_to_comment(
            0,
            "{22222222-2222-2222-2222-222222222222}",
            CommentReply::new(
                "{55555555-5555-5555-5555-555555555555}",
                "{11111111-1111-1111-1111-111111111111}",
                "2026-08-29T10:12:14Z",
                "reply two",
            )
            .unwrap(),
        )
        .unwrap();
    presentation
        .move_reply(0, "{22222222-2222-2222-2222-222222222222}", 1, 0)
        .unwrap();
    presentation
        .add_comment(
            0,
            Comment::new(
                "{44444444-4444-4444-4444-444444444444}",
                "{11111111-1111-1111-1111-111111111111}",
                "2026-08-29T10:13:14Z",
                "second",
            )
            .unwrap(),
        )
        .unwrap();
    presentation.move_comment(0, 1, 0).unwrap();
    presentation
        .set_sections(vec![
            Section::new(
                "{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}",
                "Opening",
                vec![900, 256],
            )
            .unwrap(),
        ])
        .unwrap();
    presentation.notes_header_footer_mut().unwrap().footer = Some(false);
    presentation.handout_header_footer_mut().unwrap().header = Some(false);
    presentation.move_slide(0, 1).unwrap();

    let saved = presentation.to_bytes().unwrap();
    let saved_package = open_opc(&saved, "saved collaboration package");
    assert_eq!(
        saved_package
            .content_types
            .content_type_for("/ppt/authors.xml"),
        Some(content_types::POWERPOINT_AUTHORS)
    );
    assert_eq!(
        saved_package
            .content_types
            .content_type_for("/ppt/comments/comment1.xml"),
        Some(content_types::POWERPOINT_COMMENTS)
    );
    let mut reopened = Presentation::from_bytes(&saved).unwrap();
    assert_eq!(reopened.comment_authors()[0].name, "Ada");
    let comments = reopened.comments(1).unwrap();
    assert_eq!(
        comments
            .iter()
            .map(|comment| comment.text())
            .collect::<Vec<_>>(),
        ["second", "first"]
    );
    assert_eq!(
        comments[1]
            .replies()
            .iter()
            .map(|reply| reply.text())
            .collect::<Vec<_>>(),
        ["reply two", "reply"]
    );
    assert_eq!(reopened.sections()[0].slide_ids, vec![900, 256]);
    assert_eq!(
        reopened.notes_header_footer_mut().unwrap().footer,
        Some(false)
    );
    assert_eq!(
        reopened.handout_header_footer_mut().unwrap().header,
        Some(false)
    );
}

#[test]
fn invalid_collaboration_graph_does_not_mutate_the_presentation() {
    use rpptx::{Comment, CommentAuthor, Section};

    let mut presentation = Presentation::from_bytes(&fixture_bytes()).unwrap();
    presentation
        .add_comment_author(
            CommentAuthor::new(
                "{11111111-1111-1111-1111-111111111111}",
                "Ada",
                None,
                "ada@example.test",
                "test",
            )
            .unwrap(),
        )
        .unwrap();
    let before = presentation.to_bytes().unwrap();
    let duplicate = presentation.add_comment_author(
        CommentAuthor::new(
            "{11111111-1111-1111-1111-111111111111}",
            "Other",
            None,
            "other@example.test",
            "test",
        )
        .unwrap(),
    );
    assert!(duplicate.is_err());
    let unknown_author = presentation.add_comment(
        0,
        Comment::new(
            "{22222222-2222-2222-2222-222222222222}",
            "{99999999-9999-9999-9999-999999999999}",
            "2026-08-29T10:11:12Z",
            "bad",
        )
        .unwrap(),
    );
    assert!(unknown_author.is_err());
    assert_eq!(presentation.to_bytes().unwrap(), before);
    presentation
        .add_comment(
            0,
            Comment::new(
                "{22222222-2222-2222-2222-222222222222}",
                "{11111111-1111-1111-1111-111111111111}",
                "2026-08-29T10:11:12Z",
                "valid",
            )
            .unwrap(),
        )
        .unwrap();
    let with_comment = presentation.to_bytes().unwrap();
    assert!(
        presentation
            .add_comment(
                1,
                Comment::new(
                    "{22222222-2222-2222-2222-222222222222}",
                    "{11111111-1111-1111-1111-111111111111}",
                    "2026-08-29T10:12:13Z",
                    "duplicate",
                )
                .unwrap()
            )
            .is_err()
    );
    assert_eq!(presentation.to_bytes().unwrap(), with_comment);
    assert!(
        presentation
            .set_sections(vec![
                Section::new(
                    "{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}",
                    "Bad",
                    vec![123456],
                )
                .unwrap()
            ])
            .is_err()
    );
    assert_eq!(presentation.to_bytes().unwrap(), with_comment);

    let mut occupied = fixture_package();
    occupied.set_part("/ppt/authors.xml", b"unrelated".to_vec());
    let mut occupied = open_package(occupied).unwrap();
    assert!(
        occupied
            .add_comment_author(
                CommentAuthor::new(
                    "{11111111-1111-1111-1111-111111111111}",
                    "Ada",
                    None,
                    "ada@example.test",
                    "test",
                )
                .unwrap()
            )
            .is_err()
    );

    let mut occupied_comments = fixture_package();
    occupied_comments.set_part("/ppt/comments/comment1.xml", b"unrelated".to_vec());
    let mut occupied_comments = open_package(occupied_comments).unwrap();
    occupied_comments
        .add_comment_author(
            CommentAuthor::new(
                "{11111111-1111-1111-1111-111111111111}",
                "Ada",
                None,
                "ada@example.test",
                "test",
            )
            .unwrap(),
        )
        .unwrap();
    assert!(
        occupied_comments
            .add_comment(
                0,
                Comment::new(
                    "{22222222-2222-2222-2222-222222222222}",
                    "{11111111-1111-1111-1111-111111111111}",
                    "2026-08-29T10:11:12Z",
                    "blocked",
                )
                .unwrap()
            )
            .is_err()
    );

    let comment_extension = |relationship_id: &str| {
        format!(
            r#"<p:extLst><p:ext uri="{{6950BFC3-D8DA-4A85-94F7-54DA5524770B}}"><p188:commentRel xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main" xmlns:r="{R_NS}" r:id="{relationship_id}"/></p:ext></p:extLst></p:sld>"#
        )
    };
    let mut wrong_type = fixture_package();
    wrong_type
        .get_or_create_part_rels(SLIDE_ONE_PART)
        .add_with_id("wrong-comment", rel_types::IMAGE, "../opaque/raw.xml");
    let slide = String::from_utf8(wrong_type.get_part(SLIDE_ONE_PART).unwrap().to_vec()).unwrap();
    wrong_type.set_part(
        SLIDE_ONE_PART,
        slide
            .replacen("</p:sld>", &comment_extension("wrong-comment"), 1)
            .into_bytes(),
    );
    assert!(matches!(
        open_package(wrong_type),
        Err(Error::WrongRelationshipType { .. })
    ));

    let mut external = fixture_package();
    let relationship_id = external
        .get_or_create_part_rels(SLIDE_ONE_PART)
        .add_external(
            rel_types::POWERPOINT_COMMENTS,
            "https://example.test/comments",
        );
    let slide = String::from_utf8(external.get_part(SLIDE_ONE_PART).unwrap().to_vec()).unwrap();
    external.set_part(
        SLIDE_ONE_PART,
        slide
            .replacen("</p:sld>", &comment_extension(&relationship_id), 1)
            .into_bytes(),
    );
    assert!(matches!(
        open_package(external),
        Err(Error::ExternalRelationship { .. })
    ));
}

#[test]
fn matching_mime_on_occupied_conventional_comment_part_fails_atomically() {
    use rpptx::{Comment, CommentAuthor};

    let mut package = fixture_package();
    package.set_part(
        "/ppt/comments/comment1.xml",
        br#"<unrelated xmlns="urn:f217"/>"#.to_vec(),
    );
    package.content_types.overrides.insert(
        "/ppt/comments/comment1.xml".to_owned(),
        content_types::POWERPOINT_COMMENTS.to_owned(),
    );
    let mut presentation = open_package(package).unwrap();
    presentation
        .add_comment_author(
            CommentAuthor::new(
                "{11111111-1111-1111-1111-111111111111}",
                "Ada",
                None,
                "ada@example.test",
                "test",
            )
            .unwrap(),
        )
        .unwrap();
    let before = presentation.to_bytes().unwrap();
    assert!(matches!(
        presentation.add_comment(
            0,
            Comment::new(
                "{22222222-2222-2222-2222-222222222222}",
                "{11111111-1111-1111-1111-111111111111}",
                "2026-08-29T10:11:12Z",
                "blocked",
            )
            .unwrap(),
        ),
        Err(Error::CollaborationPartCollision { part_name })
            if part_name == "/ppt/comments/comment1.xml"
    ));
    assert_eq!(presentation.to_bytes().unwrap(), before);
}

#[test]
fn second_commented_slide_allocates_comment2_after_owned_comment1() {
    use rpptx::{Comment, CommentAuthor};

    let mut presentation = Presentation::from_bytes(&fixture_bytes()).unwrap();
    let author_id = "{11111111-1111-1111-1111-111111111111}";
    presentation
        .add_comment_author(
            CommentAuthor::new(author_id, "Ada", None, "ada@example.test", "test").unwrap(),
        )
        .unwrap();
    presentation
        .add_comment(
            0,
            Comment::new(
                "{22222222-2222-2222-2222-222222222222}",
                author_id,
                "2026-08-29T10:11:12Z",
                "first",
            )
            .unwrap(),
        )
        .unwrap();
    presentation
        .add_comment(
            1,
            Comment::new(
                "{33333333-3333-3333-3333-333333333333}",
                author_id,
                "2026-08-29T10:12:13Z",
                "second",
            )
            .unwrap(),
        )
        .unwrap();
    let bytes = presentation.to_bytes().unwrap();
    let package = open_opc(&bytes, "two commented slides");
    assert!(package.get_part("/ppt/comments/comment1.xml").is_some());
    assert!(package.get_part("/ppt/comments/comment2.xml").is_some());
    let reopened = Presentation::from_bytes(&bytes).unwrap();
    assert_eq!(reopened.comments(0).unwrap()[0].text(), "first");
    assert_eq!(reopened.comments(1).unwrap()[0].text(), "second");
}

#[test]
fn remove_slide_rejects_unserializable_section_mutation_atomically() {
    let mut package = fixture_package();
    let xml = String::from_utf8(package.get_part(PRESENTATION_PART).unwrap().to_vec()).unwrap();
    let sections = r#"<p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"><q:sectionLst xmlns:q="http://schemas.microsoft.com/office/powerpoint/2010/main"><q:section name="Unsafe" id="{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}" xmlns:p14="urn:f217:p14" xmlns:p14m="urn:f217:p14m" xmlns:p14model="urn:f217:p14model"><p14:one/><p14m:two/><p14model:three/><q:sldIdLst><q:sldId id="900"/><q:sldId id="256"/></q:sldIdLst></q:section></q:sectionLst></p:ext></p:extLst>"#;
    package.set_part(
        PRESENTATION_PART,
        xml.replace("</p:presentation>", &format!("{sections}</p:presentation>"))
            .into_bytes(),
    );
    let mut presentation = open_package(package).unwrap();
    let before = presentation.to_bytes().unwrap();
    assert!(presentation.remove_slide(0).is_err());
    assert_eq!(presentation.to_bytes().unwrap(), before);
    assert_eq!(presentation.len(), 2);
    assert_eq!(Presentation::from_bytes(&before).unwrap().len(), 2);
}

#[test]
fn commented_slide_duplication_is_atomic_and_removal_keeps_comment_ownership_isolated() {
    use rpptx::{Comment, CommentAuthor};

    let mut presentation = Presentation::from_bytes(&fixture_bytes()).unwrap();
    presentation
        .add_comment_author(
            CommentAuthor::new(
                "{11111111-1111-1111-1111-111111111111}",
                "Ada",
                None,
                "ada@example.test",
                "test",
            )
            .unwrap(),
        )
        .unwrap();
    presentation
        .add_comment(
            0,
            Comment::new(
                "{22222222-2222-2222-2222-222222222222}",
                "{11111111-1111-1111-1111-111111111111}",
                "2026-08-29T10:11:12Z",
                "owned",
            )
            .unwrap(),
        )
        .unwrap();
    let before = presentation.to_bytes().unwrap();
    assert!(presentation.duplicate_slide(0).is_err());
    assert_eq!(presentation.to_bytes().unwrap(), before);
    presentation.remove_slide(1).unwrap();
    let reopened = Presentation::from_bytes(&presentation.to_bytes().unwrap()).unwrap();
    assert_eq!(reopened.comments(0).unwrap()[0].text(), "owned");
}

#[test]
fn opening_rejects_two_slides_that_share_one_empty_comment_part() {
    let mut package = fixture_package();
    let comment_part = "/custom/comments/shared.xml";
    package.set_part(
        comment_part,
        br#"<p188:cmLst xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main"/>"#
            .to_vec(),
    );
    package
        .content_types
        .add_override(comment_part, content_types::POWERPOINT_COMMENTS);
    for slide_part in [SLIDE_ONE_PART, SLIDE_TWO_PART] {
        let relationship_id = package
            .get_or_create_part_rels(slide_part)
            .add(rel_types::POWERPOINT_COMMENTS, "../comments/shared.xml");
        let original = String::from_utf8(package.get_part(slide_part).unwrap().to_vec()).unwrap();
        let extension = format!(
            r#"<p:extLst><p:ext uri="{{6950BFC3-D8DA-4A85-94F7-54DA5524770B}}"><p188:commentRel xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main" r:id="{relationship_id}"/></p:ext></p:extLst></p:sld>"#
        );
        package.set_part(
            slide_part,
            original.replacen("</p:sld>", &extension, 1).into_bytes(),
        );
    }
    assert!(Presentation::from_bytes(&package_bytes(package)).is_err());
}

fn mutation_fixture_bytes() -> Vec<u8> {
    let mut package = fixture_package();
    let original = String::from_utf8(package.get_part(SLIDE_TWO_PART).unwrap().to_vec()).unwrap();
    let with_non_visual = original.replacen(
        "<p:cNvPr/>",
        r#"<p:cNvPr id="2" name="old"><ext:nv xmlns:ext="urn:ext">one &amp; two</ext:nv></p:cNvPr>"#,
        1,
    );
    let with_shape_properties = with_non_visual.replacen(
        "<p:spPr/>",
        r#"<p:spPr producer="kept"><ext:before xmlns:ext="urn:ext"/><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom><ext:after xmlns:ext="urn:ext"/></p:spPr>"#,
        1,
    );
    let with_picture_name = with_shape_properties.replacen(
        "<p:pic><p:nvPicPr><p:cNvPr/>",
        "<p:pic><p:nvPicPr><p:cNvPr id=\"3\" name=\"picture\"/>",
        1,
    );
    let with_frame_name = with_picture_name.replacen(
        "<p:graphicFrame><p:nvGraphicFramePr/>",
        "<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id=\"4\" name=\"frame\"/></p:nvGraphicFramePr>",
        1,
    );
    let with_group_name = with_frame_name.replacen(
        "<p:grpSp><p:nvGrpSpPr/>",
        "<p:grpSp><p:nvGrpSpPr><p:cNvPr id=\"5\" name=\"group\"/></p:nvGrpSpPr>",
        1,
    );
    let with_group_property = with_group_name.replacen(
        "</p:nvGrpSpPr><p:grpSpPr/>",
        "</p:nvGrpSpPr><p:grpSpPr><a:solidFill><a:srgbClr val=\"112233\"/></a:solidFill></p:grpSpPr>",
        1,
    );
    let complete = with_group_property.replacen(
        "<p:cxnSp><p:nvCxnSpPr><p:cNvPr/>",
        "<p:cxnSp><p:nvCxnSpPr><p:cNvPr id=\"6\" name=\"connector\"/>",
        1,
    );
    let complete = complete.replacen(
        "<p:cNvCxnSpPr/>",
        r#"<p:cNvCxnSpPr><a:stCxn id="2" idx="0"/><a:endCxn id="3" idx="1"/></p:cNvCxnSpPr>"#,
        1,
    );
    let complete = complete
        .replacen(
            r#"<mc:Choice Requires="p14"><p:sp/></mc:Choice>"#,
            r#"<mc:Choice Requires="p14"><p:sp><p:nvSpPr><p:cNvPr id="20" name="choice shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp><p:cxnSp><p:nvCxnSpPr><p:cNvPr id="21" name="choice connector"/><p:cNvCxnSpPr><a:stCxn id="20" idx="0"/><a:endCxn id="20" idx="1"/></p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr><p:spPr/></p:cxnSp></mc:Choice>"#,
            1,
        )
        .replacen(
            "<mc:Fallback>",
            r#"<mc:Fallback><p:sp><p:nvSpPr><p:cNvPr id="22" name="fallback shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp><p:cxnSp><p:nvCxnSpPr><p:cNvPr id="23" name="fallback connector"/><p:cNvCxnSpPr><a:stCxn id="22" idx="0"/><a:endCxn id="22" idx="1"/></p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr><p:spPr/></p:cxnSp>"#,
            1,
        );
    assert_ne!(complete, original);
    package.set_part(SLIDE_TWO_PART, complete.into_bytes());
    package_bytes(package)
}

fn table_mutation_fixture_bytes() -> Vec<u8> {
    let mut package = fixture_package();
    let original = String::from_utf8(package.get_part(SLIDE_TWO_PART).unwrap().to_vec()).unwrap();
    let marked = original
        .replacen(
            "<a:tbl>",
            r#"<a:tbl xmlns:x="urn:f113" producer="kept"><x:before-table/><x:raw-payload xmlns:z="urn:f113:attributes" z:second="2" first="1">
  one &amp; two<!-- preserve spacing --><?f113 raw?><x:nested z:value=" A &amp; B "/>
</x:raw-payload>"#,
            1,
        )
        .replacen(
            r#"<a:tblGrid><a:gridCol w="100"/><a:gridCol w="100"/></a:tblGrid>"#,
            r#"<a:tblGrid x:grid="kept"><x:before-column/><a:gridCol w="100" x:column="kept"><x:column-child/></a:gridCol><a:gridCol w="101"/></a:tblGrid><x:after-grid/>"#,
            1,
        )
        .replacen(
            r#"<a:tr h="100"><a:tc><a:txBody>"#,
            r#"<a:tr h="100" x:row="kept"><x:before-cell/><a:tc x:cell="kept"><x:before-text/><a:txBody>"#,
            1,
        )
        .replacen(
            "<a:tcPr/>",
            r#"<a:tcPr x:cell-properties="kept"><x:cell-properties-child/></a:tcPr>"#,
            1,
        );
    assert_ne!(marked, original);
    package.set_part(SLIDE_TWO_PART, marked.into_bytes());
    package_bytes(package)
}

fn append_boundary_fixture_bytes() -> Vec<u8> {
    let mut package = fixture_package();
    let original = String::from_utf8(package.get_part(SLIDE_TWO_PART).unwrap().to_vec()).unwrap();
    let with_extension = original.replacen(
        "</p:spTree>",
        r#"<p:extLst><p:ext uri="{F110-APPEND-ORDER}"><x:marker xmlns:x="urn:f110" value="kept"/></p:ext></p:extLst></p:spTree>"#,
        1,
    );
    assert_ne!(with_extension, original);
    package.set_part(SLIDE_TWO_PART, with_extension.into_bytes());
    package_bytes(package)
}

fn package_bytes(package: OpcPackage) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    package.write_to(&mut output).unwrap();
    output.into_inner()
}

fn assert_notes_back_relationship_is_normalized(
    mut package: OpcPackage,
    source_back_relationship_ids: &HashSet<String>,
) {
    let unrelated_type = "urn:f114:unrelated";
    package
        .get_or_create_part_rels(NOTES_PART)
        .add_external(unrelated_type, "https://example.com/preserved");
    let mut presentation = Presentation::from_bytes(&package_bytes(package)).unwrap();
    presentation.duplicate_slide(0).expect("duplicate slide");

    let bytes = presentation.to_bytes().unwrap();
    let package = open_opc(&bytes, "normalized notes relationships");
    let model = CT_Presentation::from_xml(package.get_part(PRESENTATION_PART).unwrap()).unwrap();
    let duplicate_relationship = package
        .get_part_rels(PRESENTATION_PART)
        .unwrap()
        .get_by_id(&model.slide_ids[1].relationship_id)
        .unwrap();
    let duplicate_slide =
        OpcPackage::resolve_rel_target(PRESENTATION_PART, &duplicate_relationship.target);
    let notes_relationship = package
        .get_part_rels(&duplicate_slide)
        .unwrap()
        .items
        .iter()
        .find(|relationship| relationship.rel_type == rel_types::NOTES_SLIDE)
        .unwrap();
    let duplicate_notes =
        OpcPackage::resolve_rel_target(&duplicate_slide, &notes_relationship.target);
    let notes_relationships = package.get_part_rels(&duplicate_notes).unwrap();
    let back_relationships = notes_relationships
        .items
        .iter()
        .filter(|relationship| relationship.rel_type == rel_types::SLIDE)
        .collect::<Vec<_>>();
    assert_eq!(back_relationships.len(), 1);
    assert_eq!(back_relationships[0].target_mode, None);
    assert_eq!(
        OpcPackage::resolve_rel_target(&duplicate_notes, &back_relationships[0].target),
        duplicate_slide
    );
    assert!(!source_back_relationship_ids.contains(&back_relationships[0].id));
    let preserved = notes_relationships
        .items
        .iter()
        .find(|relationship| relationship.rel_type == unrelated_type)
        .unwrap();
    assert_eq!(preserved.target, "https://example.com/preserved");
    assert_eq!(preserved.target_mode.as_deref(), Some("External"));
    assert!(
        Presentation::from_bytes(&bytes)
            .unwrap()
            .validate()
            .is_empty()
    );
}

fn capture_exact_subtree(xml: &[u8], opening: &[u8], closing: &[u8]) -> Vec<u8> {
    let start = xml
        .windows(opening.len())
        .position(|window| window == opening)
        .expect("opening marker exists");
    let relative_end = xml[start..]
        .windows(closing.len())
        .position(|window| window == closing)
        .expect("closing marker exists");
    let end = start + relative_end + closing.len();
    xml[start..end].to_vec()
}

fn empty_package() -> OpcPackage {
    let mut package = OpcPackage::with_main_part(
        PRESENTATION_PART.trim_start_matches('/'),
        content_types::PRESENTATION,
    );
    package.set_part(PRESENTATION_PART, presentation_xml(&[]));
    package
}

fn fixture_package() -> OpcPackage {
    let mut package = OpcPackage::with_main_part(
        PRESENTATION_PART.trim_start_matches('/'),
        content_types::PRESENTATION,
    );
    package.set_part(
        PRESENTATION_PART,
        presentation_xml(&[(900, "ordered-first"), (256, "ordered-second")]),
    );
    package.set_part(SLIDE_ONE_PART, simple_slide_xml());
    package.set_part(SLIDE_TWO_PART, rich_slide_xml());
    package.set_part(NOTES_PART, notes_xml());
    package.set_part(LAYOUT_PART, simple_layout_xml());
    package.set_part("/custom/opaque/data.bin", vec![0, 1, 2, 255]);
    package.set_part(
        "/custom/opaque/raw.xml",
        b"<producer:raw xmlns:producer=\"urn:producer\">exact</producer:raw>".to_vec(),
    );
    package
        .content_types
        .add_override(SLIDE_ONE_PART, content_types::SLIDE);
    package
        .content_types
        .add_override(SLIDE_TWO_PART, content_types::SLIDE);
    package
        .content_types
        .add_override(NOTES_PART, content_types::NOTES_SLIDE);
    package
        .content_types
        .add_override(LAYOUT_PART, content_types::SLIDE_LAYOUT);
    package
        .content_types
        .add_default("bin", "application/octet-stream");
    let relationships = package.get_or_create_part_rels(PRESENTATION_PART);
    relationships.add_with_id("ordered-first", rel_types::SLIDE, "slides/second.xml");
    relationships.add_with_id("ordered-second", rel_types::SLIDE, "slides/first.xml");
    package.get_or_create_part_rels(SLIDE_TWO_PART).add_with_id(
        "notes-link",
        rel_types::NOTES_SLIDE,
        "../notes/speaker.xml",
    );
    package.get_or_create_part_rels(NOTES_PART).add_with_id(
        "slide-link",
        rel_types::SLIDE,
        "../slides/second.xml",
    );
    package
        .get_or_create_part_rels(SLIDE_ONE_PART)
        .add(rel_types::SLIDE_LAYOUT, "../layouts/validation.xml");
    package
        .get_or_create_part_rels(SLIDE_TWO_PART)
        .add(rel_types::SLIDE_LAYOUT, "../layouts/validation.xml");
    package
}

fn presentation_xml(slides: &[(u32, &str)]) -> Vec<u8> {
    let slide_ids = slides
        .iter()
        .map(|(id, relationship)| format!(r#"<p:sldId id="{id}" r:id="{relationship}"/>"#))
        .collect::<String>();
    format!(
        r#"<p:presentation xmlns:p="{P_NS}" xmlns:r="{R_NS}"><p:sldIdLst>{slide_ids}</p:sldIdLst><p:notesSz cx="6858000" cy="9144000"/></p:presentation>"#
    )
    .into_bytes()
}

fn simple_slide_xml() -> Vec<u8> {
    format!(
        r#"<p:sld xmlns:p="{P_NS}" xmlns:a="{A_NS}"><p:cSld name="First in storage, second in order"><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld></p:sld>"#
    )
    .into_bytes()
}

fn simple_layout_xml() -> Vec<u8> {
    format!(
        r#"<p:sldLayout xmlns:p="{P_NS}"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld></p:sldLayout>"#
    )
    .into_bytes()
}

fn rich_slide_xml() -> Vec<u8> {
    let text_shape = |text: &str| {
        format!(
            r#"<p:sp><p:nvSpPr><p:cNvPr/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp>"#
        )
    };
    let identified_text_shape = |id: u32, name: &str, text: &str| {
        format!(
            r#"<p:sp><p:nvSpPr><p:cNvPr id="{id}" name="{name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp>"#
        )
    };
    let picture = r#"<p:pic><p:nvPicPr><p:cNvPr/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill/><p:spPr/></p:pic>"#;
    let table = r#"<p:graphicFrame><p:nvGraphicFramePr/><p:xfrm/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblGrid><a:gridCol w="100"/><a:gridCol w="100"/></a:tblGrid><a:tr h="100"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>A</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>B</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr><a:tr h="100"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>C</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>D</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>"#;
    let group = format!(
        "<p:grpSp><p:nvGrpSpPr/><p:grpSpPr/>{}{}</p:grpSp>",
        identified_text_shape(7, "nested original", "nested"),
        identified_text_shape(8, "nested sibling", "sibling")
    );
    let connector = r#"<p:cxnSp><p:nvCxnSpPr><p:cNvPr/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr/></p:cxnSp>"#;
    let alternate = format!(
        r#"<mc:AlternateContent><mc:Choice Requires="p14"><p:sp/></mc:Choice><mc:Fallback>{}</mc:Fallback></mc:AlternateContent>"#,
        text_shape("fallback")
    );
    format!(
        r#"<p:sld xmlns:p="{P_NS}" xmlns:a="{A_NS}" xmlns:mc="{MC_NS}"><p:cSld name="Second in storage, first in order"><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>{}{picture}{table}{group}{connector}{alternate}</p:spTree></p:cSld></p:sld>"#,
        text_shape("plain")
    )
    .into_bytes()
}

fn notes_xml() -> Vec<u8> {
    format!(
        r#"<p:notes xmlns:p="{P_NS}" xmlns:a="{A_NS}"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr/><p:cNvSpPr/><p:nvPr><p:ph/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>speaker</a:t></a:r><a:br/><a:r><a:t>note</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notes>"#
    )
    .into_bytes()
}

fn workspace_root() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .and_then(Path::parent)
        .unwrap()
        .to_path_buf()
}

fn corpus_dir() -> PathBuf {
    std::env::var_os("RDOCX_PPTX_CORPUS_DIR")
        .map(PathBuf::from)
        .unwrap_or_else(|| workspace_root().join("corpus/pptx"))
}

fn manifest_entries() -> Vec<&'static str> {
    include_str!("../../../scripts/pptx-corpus-manifest.tsv")
        .lines()
        .skip(1)
        .map(|line| line.split('\t').next().unwrap())
        .collect()
}

const PYTHON_ORACLE: &str = r#"
from pathlib import Path
import sys
import pptx
from pptx import Presentation
from pptx.shapes.shapetree import SlideShapeFactory

assert pptx.__version__ == "1.0.2", pptx.__version__
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
KINDS = {
    (P, "sp"): "Shape",
    (P, "pic"): "Picture",
    (P, "graphicFrame"): "GraphicFrame",
    (P, "grpSp"): "Group",
    (P, "cxnSp"): "Connector",
    (MC, "AlternateContent"): "AlternateContent",
}

def qname(element):
    tag = element.tag
    if not isinstance(tag, str) or not tag.startswith("{"):
        return (None, tag)
    namespace, local = tag[1:].split("}", 1)
    return (namespace, local)

def escape(value):
    return value.replace("\\", "\\\\").replace("\t", "\\t").replace("\r", "\\r").replace("\n", "\\n")

def option(value):
    return "-" if value is None else "+" + escape(value)

def normalized_text(value):
    return value.replace("\v", "\n")

def shape_text(element, shape_collection):
    kind = KINDS.get(qname(element))
    if kind == "Shape":
        if not any(qname(child) == (P, "txBody") for child in element):
            return None
        shape = SlideShapeFactory(element, shape_collection)
        return normalized_text(shape.text)
    if kind == "GraphicFrame":
        shape = SlideShapeFactory(element, shape_collection)
        if not shape.has_table:
            return None
        return "\n".join("\t".join(normalized_text(cell.text) for cell in row.cells) for row in shape.table.rows)
    return None

def shape_children(element):
    kind = KINDS.get(qname(element))
    if kind == "Group":
        return [child for child in element if qname(child) in KINDS]
    if kind == "AlternateContent":
        fallback = next((child for child in element if qname(child) == (MC, "Fallback")), None)
        return [] if fallback is None else [child for child in fallback if qname(child) in KINDS]
    return []

def append_shape(lines, slide_index, path, element, shape_collection):
    kind = KINDS[qname(element)]
    lines.append(f"shape\t{slide_index}\t{path}\t{kind}\t{option(shape_text(element, shape_collection))}")
    for index, child in enumerate(shape_children(element)):
        append_shape(lines, slide_index, f"{path}.{index}", child, shape_collection)

for filename in sys.argv[1:]:
    path = Path(filename)
    print("deck\t" + path.name)
    presentation = Presentation(path)
    for slide_index, slide in enumerate(presentation.slides):
        top_level = [child for child in slide.shapes._spTree if qname(child) in KINDS]
        shape_lines = []
        shape_texts = []
        def collect_text(element):
            value = shape_text(element, slide.shapes)
            if value:
                shape_texts.append(value)
            for child in shape_children(element):
                collect_text(child)
        for element in top_level:
            collect_text(element)
        notes = None
        if slide.has_notes_slide:
            frame = slide.notes_slide.notes_text_frame
            notes = "" if frame is None else normalized_text(frame.text)
        print(f"slide\t{slide_index}\t{slide.slide_id}\t{option(slide.name or None)}\t{escape(chr(10).join(shape_texts))}\t{option(notes)}")
        for shape_index, element in enumerate(top_level):
            append_shape(shape_lines, slide_index, str(shape_index), element, slide.shapes)
        print("\n".join(shape_lines))
"#;

const PYTHON_VISUAL_ORACLE: &str = r#"
from pathlib import Path
import sys
import pptx
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER

assert pptx.__version__ == "1.0.2", pptx.__version__

def escape(value):
    return value.replace("\\", "\\\\").replace("\t", "\\t").replace("\r", "\\r").replace("\n", "\\n")

def shape_text(shape):
    if getattr(shape, "has_table", False):
        return "\n".join("\t".join(cell.text.replace("\v", "\n") for cell in row.cells) for row in shape.table.rows)
    if getattr(shape, "has_text_frame", False):
        return shape.text.replace("\v", "\n")
    return ""

def shape_kind(shape):
    local = shape._element.tag.rsplit("}", 1)[-1]
    return {
        "sp": "Shape",
        "pic": "Picture",
        "graphicFrame": "GraphicFrame",
        "cxnSp": "Connector",
    }[local]

IDENTITY = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)

def then(first, second):
    a, b, c, d, e, f = first
    g, h, i, j, k, l = second
    return (
        g * a + i * b,
        h * a + j * b,
        g * c + i * d,
        h * c + j * d,
        g * e + i * f + k,
        h * e + j * f + l,
    )

def rotate_about(degrees, cx, cy):
    import math
    radians = math.radians(degrees)
    sine, cosine = math.sin(radians), math.cos(radians)
    return (
        cosine,
        sine,
        -sine,
        cosine,
        cx - cosine * cx + sine * cy,
        cy - sine * cx - cosine * cy,
    )

def group_mapping(group):
    xfrm = group._element.grpSpPr.xfrm
    width, height = group.width, group.height
    child_width, child_height = xfrm.chExt.cx, xfrm.chExt.cy
    scale_x = 1.0 if child_width == 0 else width / child_width
    scale_y = 1.0 if child_height == 0 else height / child_height
    mapping = (
        scale_x,
        0.0,
        0.0,
        scale_y,
        group.left - xfrm.chOff.x * scale_x,
        group.top - xfrm.chOff.y * scale_y,
    )
    center_x = group.left + width / 2.0
    center_y = group.top + height / 2.0
    mapping = then(mapping, rotate_about(group.rotation, center_x, center_y))
    if xfrm.flipH or xfrm.flipV:
        mapping = then(
            mapping,
            (
                -1.0 if xfrm.flipH else 1.0,
                0.0,
                0.0,
                -1.0 if xfrm.flipV else 1.0,
                2.0 * center_x if xfrm.flipH else 0.0,
                2.0 * center_y if xfrm.flipV else 0.0,
            ),
        )
    return mapping

def mapped_bounds(shape, mapping):
    a, b, c, d, e, f = mapping
    corners = (
        (shape.left, shape.top),
        (shape.left + shape.width, shape.top),
        (shape.left, shape.top + shape.height),
        (shape.left + shape.width, shape.top + shape.height),
    )
    points = tuple((a * x + c * y + e, b * x + d * y + f) for x, y in corners)
    xs, ys = tuple(point[0] for point in points), tuple(point[1] for point in points)
    return min(xs), min(ys), max(xs) - min(xs), max(ys) - min(ys)

def leaves(shapes, parent_mapping=IDENTITY):
    for shape in shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            mapping = then(group_mapping(shape), parent_mapping)
            yield from leaves(shape.shapes, mapping)
        elif not (
            shape.is_placeholder
            and shape.placeholder_format.type
            in (PP_PLACEHOLDER.DATE, PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.SLIDE_NUMBER)
        ):
            yield shape, parent_mapping

def visible_shapes(slide):
    layout = slide.slide_layout
    master = layout.slide_master
    if layout._element.get("showMasterSp") != "0":
        yield from leaves(shape for shape in master.shapes if not shape.is_placeholder)
    if slide._element.get("showMasterSp") != "0":
        yield from leaves(shape for shape in layout.shapes if not shape.is_placeholder)
    yield from leaves(slide.shapes)

for filename in sys.argv[1:]:
    path = Path(filename)
    presentation = Presentation(path)
    print("deck\t" + path.name)
    for slide_index, slide in enumerate(presentation.slides):
        print(f"slide\t{slide_index}\t{presentation.slide_width / 12700:.3f}\t{presentation.slide_height / 12700:.3f}")
        for shape, mapping in visible_shapes(slide):
            bounds = mapped_bounds(shape, mapping)
            if any(value is None for value in bounds):
                continue
            print(
                f"shape\t{slide_index}\t{shape_kind(shape)}\t{bounds[0] / 12700:.3f}\t{bounds[1] / 12700:.3f}\t"
                f"{bounds[2] / 12700:.3f}\t{bounds[3] / 12700:.3f}\t{escape(shape_text(shape))}"
            )
"#;
