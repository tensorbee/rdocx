//! Custom header banner with DrawingML group shapes.
//!
//! This example demonstrates how to build a professional header banner
//! using raw DrawingML XML with set_raw_header_with_images(). The banner
//! consists of a colored rectangle background with a logo image overlaid.
//!
//! Run with: cargo run --example header_banner

use std::fmt::Write;
use std::path::Path;

use rdocx::{Document, Length};

fn main() {
    let samples_dir = Path::new(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .unwrap()
        .parent()
        .unwrap()
        .join("samples");
    std::fs::create_dir_all(&samples_dir).unwrap();

    let out = samples_dir.join("header_banner.docx");
    generate_header_banner_doc(&out);
    println!("  Created: header_banner.docx");
    println!("\nDone!");
}

fn generate_header_banner_doc(path: &Path) {
    let mut doc = Document::new();

    // Page setup with extra top margin for the banner
    doc.set_page_size(Length::inches(8.5), Length::inches(11.0));
    doc.set_margins(
        Length::twips(2292), // top — extra tall for header banner
        Length::twips(1440), // right
        Length::twips(1440), // bottom
        Length::twips(1440), // left
    );
    doc.set_header_footer_distance(Length::twips(720), Length::twips(432));

    // Generate a simple logo image (white text on transparent background)
    let logo_img = create_logo_png(220, 48);

    // ── Dark blue banner ──
    let banner = build_header_banner_xml(
        "rId1",
        &BannerOpts {
            bg_color: "1A3C6E",
            banner_width: 7772400, // full page width in EMU (~8.5")
            banner_height: 969026, // banner height in EMU (~1.06")
            logo_width: 2011680,   // logo display width (~2.2")
            logo_height: 438912,   // logo display height (~0.48")
            logo_x_offset: 295125, // left padding
            logo_y_offset: 265057, // vertical centering
        },
    );

    doc.set_raw_header_with_images(
        banner.clone(),
        &[("rId1", &logo_img, "logo.png")],
        rdocx_oxml::header_footer::HdrFtrType::Default,
    );

    // Use a different first page header (same banner, different color)
    doc.set_different_first_page(true);
    let first_page_banner = build_header_banner_xml(
        "rId1",
        &BannerOpts {
            bg_color: "2E75B6", // lighter blue for cover
            banner_width: 7772400,
            banner_height: 969026,
            logo_width: 2011680,
            logo_height: 438912,
            logo_x_offset: 295125,
            logo_y_offset: 265057,
        },
    );
    doc.set_raw_header_with_images(
        first_page_banner,
        &[("rId1", &logo_img, "logo.png")],
        rdocx_oxml::header_footer::HdrFtrType::First,
    );

    // Footer
    doc.set_footer("Confidential — Internal Use Only");

    // ── Page 1: Cover ──
    doc.add_paragraph("Company Report").style("Heading1");

    doc.add_paragraph(
        "This document demonstrates a custom header banner built with DrawingML \
         group shapes. The banner uses a colored rectangle with a logo image overlaid, \
         positioned at the top of each page.",
    );

    doc.add_paragraph("");

    doc.add_paragraph("How the Header Banner Works")
        .style("Heading2");

    doc.add_paragraph(
        "The header banner is built using set_raw_header_with_images(), which \
         accepts raw XML and a list of (rel_id, image_data, filename) tuples. \
         The XML uses a DrawingML group shape (wpg:wgp) containing:",
    );

    doc.add_bullet_list_item(
        "A wps:wsp rectangle shape with a solid color fill (the background bar)",
        0,
    );
    doc.add_bullet_list_item(
        "A pic:pic image element positioned within the group (the logo)",
        0,
    );
    doc.add_bullet_list_item(
        "The group is wrapped in a wp:anchor element for absolute page positioning",
        0,
    );

    doc.add_paragraph("");

    doc.add_paragraph("Customization").style("Heading2");

    doc.add_paragraph(
        "All dimensions are in EMU (English Metric Units) where 914400 EMU = 1 inch. \
         You can customize:",
    );

    doc.add_bullet_list_item("bg_color — any hex color for the rectangle background", 0);
    doc.add_bullet_list_item("banner_width / banner_height — size of the full banner", 0);
    doc.add_bullet_list_item("logo_width / logo_height — display size of the logo", 0);
    doc.add_bullet_list_item(
        "logo_x_offset / logo_y_offset — logo position within the banner",
        0,
    );

    doc.add_paragraph("");

    doc.add_paragraph("Different First Page").style("Heading2");

    doc.add_paragraph(
        "This page uses a lighter blue banner (first page header). \
         Subsequent pages use a darker blue banner (default header). \
         Use set_different_first_page(true) to enable this.",
    );

    // ── Page 2 ──
    doc.add_paragraph("").page_break_before(true);

    doc.add_paragraph("Second Page").style("Heading1");

    doc.add_paragraph(
        "This page shows the default header banner (dark blue). The first page \
         had a lighter blue banner because we set a different first-page header.",
    );

    doc.add_paragraph("");

    doc.add_paragraph(
        "The banner repeats on every page because it is placed in the header part. \
         You can have different banners for default, first-page, and even-page headers.",
    );

    doc.set_title("Header Banner Example");
    doc.set_author("rdocx");

    doc.save(path).unwrap();
}

// ─────────────────────────────────────────────────────────────────────────────
// Header banner builder
// ─────────────────────────────────────────────────────────────────────────────

struct BannerOpts<'a> {
    bg_color: &'a str,
    banner_width: i64,
    banner_height: i64,
    logo_width: i64,
    logo_height: i64,
    logo_x_offset: i64,
    logo_y_offset: i64,
}

/// Build complete `<w:hdr>` XML for a banner header (colored rect + logo image).
///
/// `image_rel_id` is the rId that will reference the logo image (e.g. "rId1").
fn build_header_banner_xml(image_rel_id: &str, opts: &BannerOpts) -> Vec<u8> {
    let mut xml = String::with_capacity(2048);

    write!(
        xml,
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#
    )
    .unwrap();
    write!(xml, r#"<w:hdr "#).unwrap();
    write!(
        xml,
        r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" "#
    )
    .unwrap();
    write!(
        xml,
        r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" "#
    )
    .unwrap();
    write!(
        xml,
        r#"xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" "#
    )
    .unwrap();
    write!(
        xml,
        r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" "#
    )
    .unwrap();
    write!(
        xml,
        r#"xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" "#
    )
    .unwrap();
    write!(
        xml,
        r#"xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" "#
    )
    .unwrap();
    write!(
        xml,
        r#"xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" "#
    )
    .unwrap();
    write!(
        xml,
        r#"xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">"#
    )
    .unwrap();

    write!(xml, r#"<w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr>"#).unwrap();
    write!(xml, r#"<w:r><w:rPr><w:noProof/></w:rPr>"#).unwrap();
    write!(xml, r#"<mc:AlternateContent><mc:Choice Requires="wpg">"#).unwrap();
    write!(xml, r#"<w:drawing>"#).unwrap();

    // Anchor at top-left of page
    write!(
        xml,
        r#"<wp:anchor distT="0" distB="0" distL="0" distR="0" "#
    )
    .unwrap();
    write!(xml, r#"simplePos="0" relativeHeight="251658240" "#).unwrap();
    write!(xml, r#"behindDoc="0" locked="0" layoutInCell="1" "#).unwrap();
    write!(xml, r#"hidden="0" allowOverlap="1">"#).unwrap();
    write!(xml, r#"<wp:simplePos x="0" y="0"/>"#).unwrap();
    write!(
        xml,
        r#"<wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH>"#
    )
    .unwrap();
    write!(
        xml,
        r#"<wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV>"#
    )
    .unwrap();
    write!(
        xml,
        r#"<wp:extent cx="{}" cy="{}"/>"#,
        opts.banner_width, opts.banner_height
    )
    .unwrap();
    write!(xml, r#"<wp:effectExtent l="0" t="0" r="0" b="0"/>"#).unwrap();
    write!(xml, r#"<wp:wrapNone/>"#).unwrap();
    write!(xml, r#"<wp:docPr id="1" name="Header Banner"/>"#).unwrap();
    write!(xml, r#"<wp:cNvGraphicFramePr/>"#).unwrap();

    // Group shape containing rect + image
    write!(xml, r#"<a:graphic>"#).unwrap();
    write!(
        xml,
        r#"<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">"#
    )
    .unwrap();
    write!(xml, r#"<wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm>"#).unwrap();
    write!(
        xml,
        r#"<a:off x="0" y="0"/><a:ext cx="{w}" cy="{h}"/>"#,
        w = opts.banner_width,
        h = opts.banner_height
    )
    .unwrap();
    write!(
        xml,
        r#"<a:chOff x="0" y="0"/><a:chExt cx="{w}" cy="{h}"/>"#,
        w = opts.banner_width,
        h = opts.banner_height
    )
    .unwrap();
    write!(xml, r#"</a:xfrm></wpg:grpSpPr>"#).unwrap();

    // Background rectangle
    write!(
        xml,
        r#"<wps:wsp><wps:cNvPr id="2" name="Background"/><wps:cNvSpPr/><wps:spPr>"#
    )
    .unwrap();
    write!(
        xml,
        r#"<a:xfrm><a:off x="0" y="0"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>"#,
        w = opts.banner_width,
        h = opts.banner_height
    )
    .unwrap();
    write!(xml, r#"<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>"#).unwrap();
    write!(
        xml,
        r#"<a:solidFill><a:srgbClr val="{}"/></a:solidFill>"#,
        opts.bg_color
    )
    .unwrap();
    write!(xml, r#"<a:ln><a:noFill/></a:ln>"#).unwrap();
    write!(xml, r#"</wps:spPr><wps:bodyPr/></wps:wsp>"#).unwrap();

    // Logo image
    write!(
        xml,
        r#"<pic:pic><pic:nvPicPr><pic:cNvPr id="3" name="Logo"/><pic:cNvPicPr/></pic:nvPicPr>"#
    )
    .unwrap();
    write!(xml, r#"<pic:blipFill><a:blip r:embed="{}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>"#, image_rel_id).unwrap();
    write!(
        xml,
        r#"<pic:spPr><a:xfrm><a:off x="{}" y="{}"/><a:ext cx="{}" cy="{}"/></a:xfrm>"#,
        opts.logo_x_offset, opts.logo_y_offset, opts.logo_width, opts.logo_height
    )
    .unwrap();
    write!(xml, r#"<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>"#).unwrap();
    write!(
        xml,
        r#"<a:noFill/><a:ln><a:noFill/></a:ln></pic:spPr></pic:pic>"#
    )
    .unwrap();

    // Close all
    write!(xml, r#"</wpg:wgp></a:graphicData></a:graphic>"#).unwrap();
    write!(xml, r#"</wp:anchor></w:drawing>"#).unwrap();
    write!(xml, r#"</mc:Choice></mc:AlternateContent>"#).unwrap();
    write!(xml, r#"</w:r></w:p></w:hdr>"#).unwrap();

    xml.into_bytes()
}

// ─────────────────────────────────────────────────────────────────────────────
// PNG generation helpers
// ─────────────────────────────────────────────────────────────────────────────

/// Create a simple logo PNG (white text area on dark background).
fn create_logo_png(width: u32, height: u32) -> Vec<u8> {
    let mut pixels = Vec::with_capacity((width * height * 4) as usize);
    for y in 0..height {
        for x in 0..width {
            // White rectangle in the center area (simulates logo text)
            let in_text_area =
                x > width / 8 && x < width * 7 / 8 && y > height / 4 && y < height * 3 / 4;
            if in_text_area {
                pixels.extend_from_slice(&[255, 255, 255, 255]);
            } else {
                pixels.extend_from_slice(&[255, 255, 255, 40]); // mostly transparent
            }
        }
    }
    encode_png(width, height, &pixels)
}

fn encode_png(width: u32, height: u32, pixels: &[u8]) -> Vec<u8> {
    let mut png = Vec::new();
    {
        use std::io::Write as _;
        png.write_all(&[137, 80, 78, 71, 13, 10, 26, 10]).unwrap();

        let mut ihdr = Vec::new();
        ihdr.extend_from_slice(&width.to_be_bytes());
        ihdr.extend_from_slice(&height.to_be_bytes());
        ihdr.extend_from_slice(&[8, 6, 0, 0, 0]); // 8-bit RGBA
        write_chunk(&mut png, b"IHDR", &ihdr);

        let mut raw = Vec::new();
        for y in 0..height {
            raw.push(0);
            let s = (y * width * 4) as usize;
            raw.extend_from_slice(&pixels[s..s + (width * 4) as usize]);
        }
        write_chunk(&mut png, b"IDAT", &zlib_store(&raw));
        write_chunk(&mut png, b"IEND", &[]);
    }
    png
}

fn write_chunk(out: &mut Vec<u8>, ct: &[u8; 4], data: &[u8]) {
    use std::io::Write as _;
    out.write_all(&(data.len() as u32).to_be_bytes()).unwrap();
    out.write_all(ct).unwrap();
    out.write_all(data).unwrap();
    out.write_all(&crc32(ct, data).to_be_bytes()).unwrap();
}

fn crc32(ct: &[u8], data: &[u8]) -> u32 {
    static T: std::sync::LazyLock<[u32; 256]> = std::sync::LazyLock::new(|| {
        let mut t = [0u32; 256];
        for n in 0..256u32 {
            let mut c = n;
            for _ in 0..8 {
                c = if c & 1 != 0 {
                    0xEDB88320 ^ (c >> 1)
                } else {
                    c >> 1
                };
            }
            t[n as usize] = c;
        }
        t
    });
    let mut c = 0xFFFFFFFF_u32;
    for &b in ct.iter().chain(data) {
        c = T[((c ^ b as u32) & 0xFF) as usize] ^ (c >> 8);
    }
    c ^ 0xFFFFFFFF
}

fn zlib_store(data: &[u8]) -> Vec<u8> {
    let mut out = vec![0x78, 0x01];
    for (i, chunk) in data.chunks(65535).enumerate() {
        let last = i == data.chunks(65535).count() - 1;
        out.push(if last { 0x01 } else { 0x00 });
        let len = chunk.len() as u16;
        out.extend_from_slice(&len.to_le_bytes());
        out.extend_from_slice(&(!len).to_le_bytes());
        out.extend_from_slice(chunk);
    }
    let (mut a, mut b) = (1u32, 0u32);
    for &byte in data {
        a = (a + byte as u32) % 65521;
        b = (b + a) % 65521;
    }
    out.extend_from_slice(&((b << 16) | a).to_be_bytes());
    out
}
