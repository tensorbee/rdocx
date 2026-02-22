//! Page-to-image rendering using tiny-skia software rasterizer.

use rdocx_layout::{LayoutResult, PageFrame, PositionedElement};
use tiny_skia::{FillRule, Paint, PathBuilder, Pixmap, Stroke, Transform};

/// Render a single page to PNG bytes.
///
/// # Arguments
/// * `layout` - The layout result containing all pages and font data
/// * `page_index` - 0-based page index to render
/// * `dpi` - Dots per inch (72 = 1:1 with points, 150 = 2x, 300 = 4.17x)
pub fn render_page_to_png(layout: &LayoutResult, page_index: usize, dpi: f64) -> Option<Vec<u8>> {
    let page = layout.pages.get(page_index)?;
    let pixmap = render_page_to_pixmap(page, &layout.fonts, dpi)?;
    pixmap.encode_png().ok()
}

/// Render all pages to PNG bytes.
pub fn render_all_pages(layout: &LayoutResult, dpi: f64) -> Vec<Vec<u8>> {
    layout
        .pages
        .iter()
        .filter_map(|page| {
            let pixmap = render_page_to_pixmap(page, &layout.fonts, dpi)?;
            pixmap.encode_png().ok()
        })
        .collect()
}

/// Render a page to a Pixmap.
fn render_page_to_pixmap(
    page: &PageFrame,
    fonts: &[rdocx_layout::FontData],
    dpi: f64,
) -> Option<Pixmap> {
    let scale = dpi / 72.0; // points to pixels
    let width = (page.width * scale).ceil() as u32;
    let height = (page.height * scale).ceil() as u32;

    let mut pixmap = Pixmap::new(width, height)?;

    // Fill with white background
    pixmap.fill(tiny_skia::Color::WHITE);

    let transform = Transform::from_scale(scale as f32, scale as f32);

    for element in &page.elements {
        match element {
            PositionedElement::FilledRect { rect, color } => {
                let sk_rect = tiny_skia::Rect::from_xywh(
                    rect.x as f32,
                    rect.y as f32,
                    rect.width as f32,
                    rect.height as f32,
                );
                if let Some(sk_rect) = sk_rect {
                    let mut paint = Paint::default();
                    paint.set_color_rgba8(
                        (color.r * 255.0) as u8,
                        (color.g * 255.0) as u8,
                        (color.b * 255.0) as u8,
                        (color.a * 255.0) as u8,
                    );
                    paint.anti_alias = false;
                    pixmap.fill_rect(sk_rect, &paint, transform, None);
                }
            }
            PositionedElement::Line {
                start,
                end,
                width: line_width,
                color,
                dash_pattern: _,
            } => {
                let mut pb = PathBuilder::new();
                pb.move_to(start.x as f32, start.y as f32);
                pb.line_to(end.x as f32, end.y as f32);
                if let Some(path) = pb.finish() {
                    let mut paint = Paint::default();
                    paint.set_color_rgba8(
                        (color.r * 255.0) as u8,
                        (color.g * 255.0) as u8,
                        (color.b * 255.0) as u8,
                        (color.a * 255.0) as u8,
                    );
                    paint.anti_alias = true;
                    let stroke = Stroke { width: *line_width as f32, ..Stroke::default() };
                    pixmap.stroke_path(&path, &paint, &stroke, transform, None);
                }
            }
            PositionedElement::Text(glyph_run) => {
                // Render text by extracting glyph outlines from the font
                let font_data = fonts.iter().find(|f| f.id == glyph_run.font_id);
                if let Some(font_data) = font_data {
                    render_glyph_run(&mut pixmap, glyph_run, font_data, transform);
                }
            }
            PositionedElement::Image {
                rect, data, content_type, ..
            } => {
                if !data.is_empty() {
                    render_image(&mut pixmap, rect, data, content_type, transform);
                }
            }
            PositionedElement::LinkAnnotation { .. } => {
                // Link annotations are not visual elements in raster output
            }
        }
    }

    Some(pixmap)
}

/// Render a glyph run by extracting glyph outlines from the font.
fn render_glyph_run(
    pixmap: &mut Pixmap,
    glyph_run: &rdocx_layout::GlyphRun,
    font_data: &rdocx_layout::FontData,
    transform: Transform,
) {
    let Ok(face) = ttf_parser::Face::parse(&font_data.data, font_data.face_index) else {
        return;
    };

    let upem = face.units_per_em() as f64;
    let scale = glyph_run.font_size / upem;

    let mut paint = Paint::default();
    paint.set_color_rgba8(
        (glyph_run.color.r * 255.0) as u8,
        (glyph_run.color.g * 255.0) as u8,
        (glyph_run.color.b * 255.0) as u8,
        (glyph_run.color.a * 255.0) as u8,
    );
    paint.anti_alias = true;

    let mut x = glyph_run.origin.x;
    let y = glyph_run.origin.y;

    for (i, &glyph_id) in glyph_run.glyph_ids.iter().enumerate() {
        let gid = ttf_parser::GlyphId(glyph_id);

        // Build path from glyph outline
        let mut builder = GlyphPathBuilder::new();
        if face.outline_glyph(gid, &mut builder).is_some() {
            if let Some(path) = builder.finish() {
                // Transform: translate to glyph position, scale from font units to points,
                // and flip Y (font coordinates are Y-up, pixmap is Y-down)
                let glyph_transform = transform
                    .pre_translate(x as f32, y as f32)
                    .pre_scale(scale as f32, -(scale as f32));

                pixmap.fill_path(&path, &paint, FillRule::Winding, glyph_transform, None);
            }
        }

        if i < glyph_run.advances.len() {
            x += glyph_run.advances[i];
        }
    }
}

/// Convert a glyph outline to a tiny-skia Path.
struct GlyphPathBuilder {
    pb: PathBuilder,
}

impl GlyphPathBuilder {
    fn new() -> Self {
        GlyphPathBuilder {
            pb: PathBuilder::new(),
        }
    }

    fn finish(self) -> Option<tiny_skia::Path> {
        self.pb.finish()
    }
}

impl ttf_parser::OutlineBuilder for GlyphPathBuilder {
    fn move_to(&mut self, x: f32, y: f32) {
        self.pb.move_to(x, y);
    }

    fn line_to(&mut self, x: f32, y: f32) {
        self.pb.line_to(x, y);
    }

    fn quad_to(&mut self, x1: f32, y1: f32, x: f32, y: f32) {
        self.pb.quad_to(x1, y1, x, y);
    }

    fn curve_to(&mut self, x1: f32, y1: f32, x2: f32, y2: f32, x: f32, y: f32) {
        self.pb.cubic_to(x1, y1, x2, y2, x, y);
    }

    fn close(&mut self) {
        self.pb.close();
    }
}

/// Render an image onto the pixmap.
fn render_image(
    pixmap: &mut Pixmap,
    rect: &rdocx_layout::Rect,
    data: &[u8],
    _content_type: &str,
    transform: Transform,
) {
    // Decode the image
    let decoded = crate::image::decode_image(data, _content_type);
    let Some(decoded) = decoded else {
        return;
    };

    // Create a pixmap from the decoded image data
    let img_pixmap = if decoded.is_jpeg || decoded.color_space == "DeviceRGB" {
        // Convert RGB to RGBA
        let mut rgba = Vec::with_capacity(decoded.width as usize * decoded.height as usize * 4);
        if let Some(alpha) = &decoded.alpha {
            for (rgb, &a) in decoded.data.chunks_exact(3).zip(alpha.iter()) {
                rgba.extend_from_slice(&[rgb[0], rgb[1], rgb[2], a]);
            }
        } else {
            for rgb in decoded.data.chunks_exact(3) {
                rgba.extend_from_slice(&[rgb[0], rgb[1], rgb[2], 255]);
            }
        }
        let size = tiny_skia::IntSize::from_wh(decoded.width, decoded.height);
        size.and_then(|s| Pixmap::from_vec(rgba, s))
    } else if decoded.color_space == "DeviceGray" {
        // Convert grayscale to RGBA
        let mut rgba = Vec::with_capacity(decoded.width as usize * decoded.height as usize * 4);
        for &gray in &decoded.data {
            rgba.push(gray);
            rgba.push(gray);
            rgba.push(gray);
            rgba.push(255);
        }
        let size = tiny_skia::IntSize::from_wh(decoded.width, decoded.height);
        size.and_then(|s| Pixmap::from_vec(rgba, s))
    } else {
        return;
    };

    let Some(img_pixmap) = img_pixmap else {
        return;
    };

    // Calculate the transform to position and scale the image
    let sx = rect.width as f32 / decoded.width as f32;
    let sy = rect.height as f32 / decoded.height as f32;

    let img_transform = transform
        .pre_translate(rect.x as f32, rect.y as f32)
        .pre_scale(sx, sy);

    let pattern = tiny_skia::Pattern::new(
        img_pixmap.as_ref(),
        tiny_skia::SpreadMode::Pad,
        tiny_skia::FilterQuality::Bilinear,
        1.0,
        Transform::identity(),
    );

    let paint = Paint { shader: pattern, ..Paint::default() };

    let fill_rect = tiny_skia::Rect::from_xywh(
        0.0,
        0.0,
        decoded.width as f32,
        decoded.height as f32,
    );
    if let Some(fill_rect) = fill_rect {
        pixmap.fill_rect(fill_rect, &paint, img_transform, None);
    }
}
