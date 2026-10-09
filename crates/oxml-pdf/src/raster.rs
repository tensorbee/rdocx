//! Page-to-image rendering using tiny-skia software rasterizer.

use std::collections::HashSet;
use std::error::Error;
use std::fmt;
use std::io::Cursor;

use jpeg_encoder::{ColorType as JpegColorType, Encoder as JpegEncoder};
use oxml_layout::{
    Effect, LayoutResult, LineCap, LineJoin, PageFrame, Paint as LayoutPaint, Path as LayoutPath,
    PathCommand, PositionedElement, Stroke as LayoutStroke,
};
use tiny_skia::{
    FillRule, Mask, Paint, PathBuilder, Pixmap, PixmapPaint, Stroke, StrokeDash, Transform,
};
use zune_jpeg::JpegDecoder;
use zune_jpeg::zune_core::bytestream::ZCursor;
use zune_jpeg::zune_core::colorspace::ColorSpace;
use zune_jpeg::zune_core::options::DecoderOptions;

/// Image container and encoding options for raster output.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RasterFormat {
    /// Encode each selected page as a PNG.
    Png {
        /// Leave unpainted pixels transparent instead of compositing on white.
        transparent_background: bool,
    },
    /// Encode each selected page as an opaque JPEG.
    Jpeg {
        /// JPEG quality from 1 through 100.
        quality: u8,
    },
    /// Encode selected pages into one opaque, uncompressed, multi-page TIFF.
    Tiff,
}

/// Options for raster image export.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct RasterOptions {
    /// Dots per inch. Must be finite and positive.
    pub dpi: f64,
    /// Requested image container and format-specific options.
    pub format: RasterFormat,
}

/// Encoded raster output.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum RasterOutput {
    /// One image per selected page, used for PNG and JPEG.
    SeparatePages(Vec<Vec<u8>>),
    /// One multi-directory TIFF stream for every selected page.
    MultiPageTiff(Vec<u8>),
}

/// A reason a raster export request could not be completed.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum RasterError {
    /// DPI was NaN, infinite, zero, or negative.
    InvalidDpi,
    /// No page was selected.
    EmptyPageSelection,
    /// A selected zero-based page index does not exist.
    PageOutOfRange {
        page_index: usize,
        page_count: usize,
    },
    /// A selected zero-based page index appears more than once.
    DuplicatePage { page_index: usize },
    /// JPEG quality must be in the inclusive 1 through 100 range.
    InvalidJpegQuality { quality: u8 },
    /// The requested raster dimensions cannot be represented.
    InvalidDimensions,
    /// The rendered pixmap could not be encoded into the requested format.
    EncodeFailed { format: &'static str },
}

impl fmt::Display for RasterError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::InvalidDpi => formatter.write_str("raster DPI must be finite and positive"),
            Self::EmptyPageSelection => formatter.write_str("at least one raster page is required"),
            Self::PageOutOfRange {
                page_index,
                page_count,
            } => write!(
                formatter,
                "raster page index {page_index} is out of range for {page_count} pages"
            ),
            Self::DuplicatePage { page_index } => {
                write!(
                    formatter,
                    "raster page index {page_index} was selected more than once"
                )
            }
            Self::InvalidJpegQuality { quality } => {
                write!(formatter, "JPEG quality {quality} is outside 1 through 100")
            }
            Self::InvalidDimensions => {
                formatter.write_str("raster dimensions are invalid for the requested format")
            }
            Self::EncodeFailed { format } => write!(formatter, "{format} raster encoding failed"),
        }
    }
}

impl Error for RasterError {}

/// Render a single page to PNG bytes.
///
/// # Arguments
/// * `layout` - The layout result containing all pages and font data
/// * `page_index` - 0-based page index to render
/// * `dpi` - Dots per inch (72 = 1:1 with points, 150 = 2x, 300 = 4.17x)
pub fn render_page_to_png(layout: &LayoutResult, page_index: usize, dpi: f64) -> Option<Vec<u8>> {
    let page = layout.pages.get(page_index)?;
    let pixmap = render_page_to_pixmap_with_background(page, &layout.fonts, dpi, false)?;
    pixmap.encode_png().ok()
}

/// Render all pages to PNG bytes.
pub fn render_all_pages(layout: &LayoutResult, dpi: f64) -> Vec<Vec<u8>> {
    layout
        .pages
        .iter()
        .filter_map(|page| {
            let pixmap = render_page_to_pixmap_with_background(page, &layout.fonts, dpi, false)?;
            pixmap.encode_png().ok()
        })
        .collect()
}

/// Render selected zero-based pages to the requested raster image format.
pub fn render_pages(
    layout: &LayoutResult,
    page_indices: &[usize],
    options: RasterOptions,
) -> Result<RasterOutput, RasterError> {
    validate_request(layout, page_indices, options)?;

    match options.format {
        RasterFormat::Png {
            transparent_background,
        } => page_indices
            .iter()
            .map(|&page_index| {
                let pixmap = render_selected_pixmap(
                    layout,
                    page_index,
                    options.dpi,
                    transparent_background,
                )?;
                pixmap
                    .encode_png()
                    .map_err(|_| RasterError::EncodeFailed { format: "PNG" })
            })
            .collect::<Result<Vec<_>, _>>()
            .map(RasterOutput::SeparatePages),
        RasterFormat::Jpeg { quality } => page_indices
            .iter()
            .map(|&page_index| {
                let pixmap = render_selected_pixmap(layout, page_index, options.dpi, false)?;
                encode_jpeg(&pixmap, quality)
            })
            .collect::<Result<Vec<_>, _>>()
            .map(RasterOutput::SeparatePages),
        RasterFormat::Tiff => {
            encode_tiff(layout, page_indices, options.dpi).map(RasterOutput::MultiPageTiff)
        }
    }
}

fn validate_request(
    layout: &LayoutResult,
    page_indices: &[usize],
    options: RasterOptions,
) -> Result<(), RasterError> {
    if !options.dpi.is_finite() || options.dpi <= 0.0 {
        return Err(RasterError::InvalidDpi);
    }
    if page_indices.is_empty() {
        return Err(RasterError::EmptyPageSelection);
    }
    if let RasterFormat::Jpeg { quality } = options.format
        && !(1..=100).contains(&quality)
    {
        return Err(RasterError::InvalidJpegQuality { quality });
    }
    let page_count = layout.pages.len();
    let mut seen = HashSet::with_capacity(page_indices.len());
    for &page_index in page_indices {
        if page_index >= page_count {
            return Err(RasterError::PageOutOfRange {
                page_index,
                page_count,
            });
        }
        if !seen.insert(page_index) {
            return Err(RasterError::DuplicatePage { page_index });
        }
    }
    Ok(())
}

fn render_selected_pixmap(
    layout: &LayoutResult,
    page_index: usize,
    dpi: f64,
    transparent_background: bool,
) -> Result<Pixmap, RasterError> {
    let page = layout
        .pages
        .get(page_index)
        .ok_or(RasterError::InvalidDimensions)?;
    render_page_to_pixmap_with_background(page, &layout.fonts, dpi, transparent_background)
        .ok_or(RasterError::InvalidDimensions)
}

fn encode_jpeg(pixmap: &Pixmap, quality: u8) -> Result<Vec<u8>, RasterError> {
    let width = u16::try_from(pixmap.width()).map_err(|_| RasterError::InvalidDimensions)?;
    let height = u16::try_from(pixmap.height()).map_err(|_| RasterError::InvalidDimensions)?;
    let rgb = opaque_rgb(pixmap);
    let mut output = Vec::new();
    JpegEncoder::new(&mut output, quality)
        .encode(&rgb, width, height, JpegColorType::Rgb)
        .map_err(|_| RasterError::EncodeFailed { format: "JPEG" })?;
    Ok(output)
}

fn encode_tiff(
    layout: &LayoutResult,
    page_indices: &[usize],
    dpi: f64,
) -> Result<Vec<u8>, RasterError> {
    let mut cursor = Cursor::new(Vec::new());
    {
        let mut encoder = tiff::encoder::TiffEncoder::new(&mut cursor)
            .map_err(|_| RasterError::EncodeFailed { format: "TIFF" })?;
        for &page_index in page_indices {
            let pixmap = render_selected_pixmap(layout, page_index, dpi, false)?;
            let rgb = opaque_rgb(&pixmap);
            encoder
                .write_image::<tiff::encoder::colortype::RGB8>(
                    pixmap.width(),
                    pixmap.height(),
                    &rgb,
                )
                .map_err(|_| RasterError::EncodeFailed { format: "TIFF" })?;
        }
    }
    Ok(cursor.into_inner())
}

fn opaque_rgb(pixmap: &Pixmap) -> Vec<u8> {
    let mut rgb = Vec::with_capacity(pixmap.width() as usize * pixmap.height() as usize * 3);
    for pixel in pixmap.pixels() {
        let alpha = pixel.alpha();
        rgb.push(pixel.red().saturating_add(255 - alpha));
        rgb.push(pixel.green().saturating_add(255 - alpha));
        rgb.push(pixel.blue().saturating_add(255 - alpha));
    }
    rgb
}

/// Render a page to a Pixmap.
#[cfg(test)]
fn render_page_to_pixmap(
    page: &PageFrame,
    fonts: &[oxml_layout::FontData],
    dpi: f64,
) -> Option<Pixmap> {
    render_page_to_pixmap_with_background(page, fonts, dpi, false)
}

fn render_page_to_pixmap_with_background(
    page: &PageFrame,
    fonts: &[oxml_layout::FontData],
    dpi: f64,
    transparent_background: bool,
) -> Option<Pixmap> {
    let scale = dpi / 72.0; // points to pixels
    let width = (page.width * scale).ceil() as u32;
    let height = (page.height * scale).ceil() as u32;

    let mut pixmap = Pixmap::new(width, height)?;

    if !transparent_background {
        pixmap.fill(tiny_skia::Color::WHITE);
    }

    let transform = Transform::from_scale(scale as f32, scale as f32);

    if let Some(background) = &page.background {
        let path = LayoutPath::rect(oxml_layout::Rect {
            x: 0.0,
            y: 0.0,
            width: page.width,
            height: page.height,
        });
        render_path(&mut pixmap, &path, Some(background), None, transform, None);
    }
    render_elements(&mut pixmap, fonts, &page.elements, transform, None);

    Some(pixmap)
}

fn render_elements(
    pixmap: &mut Pixmap,
    fonts: &[oxml_layout::FontData],
    elements: &[PositionedElement],
    transform: Transform,
    mask: Option<&Mask>,
) {
    for element in elements {
        match element {
            PositionedElement::FilledRect { rect, color } => {
                let Some(rect) = tiny_skia::Rect::from_xywh(
                    rect.x as f32,
                    rect.y as f32,
                    rect.width as f32,
                    rect.height as f32,
                ) else {
                    continue;
                };
                let mut paint = solid_paint(*color);
                paint.anti_alias = false;
                pixmap.fill_rect(rect, &paint, transform, mask);
            }
            PositionedElement::Line {
                start,
                end,
                width,
                color,
                dash_pattern,
            } => {
                let mut builder = PathBuilder::new();
                builder.move_to(start.x as f32, start.y as f32);
                builder.line_to(end.x as f32, end.y as f32);
                let Some(path) = builder.finish() else {
                    continue;
                };
                let dash = dash_pattern
                    .and_then(|(on, off)| StrokeDash::new(vec![on as f32, off as f32], 0.0));
                let stroke = Stroke {
                    width: *width as f32,
                    dash,
                    ..Stroke::default()
                };
                pixmap.stroke_path(&path, &solid_paint(*color), &stroke, transform, mask);
            }
            PositionedElement::Text(glyph_run) => {
                if let Some(font_data) = fonts.iter().find(|font| font.id == glyph_run.font_id) {
                    render_glyph_run(pixmap, glyph_run, font_data, transform, mask);
                }
            }
            PositionedElement::MultilingualText(glyph_run) => {
                if glyph_run.is_valid()
                    && let Some(font_data) = fonts.iter().find(|font| font.id == glyph_run.font_id)
                {
                    render_positioned_glyph_run(
                        pixmap,
                        &glyph_run.legacy_projection(),
                        font_data,
                        transform,
                        mask,
                        &glyph_run.x_offsets,
                        &glyph_run.y_offsets,
                        &glyph_run.y_advances,
                    );
                }
            }
            PositionedElement::Image {
                rect,
                data,
                content_type,
                ..
            } => {
                if !data.is_empty() {
                    render_image(pixmap, rect, data, content_type, transform, mask);
                }
            }
            PositionedElement::LinkAnnotation { .. } => {}
            PositionedElement::Path(element) => render_path(
                pixmap,
                &element.path,
                element.fill.as_ref(),
                element.stroke.as_ref(),
                transform,
                mask,
            ),
            PositionedElement::Group(group) => {
                let child_transform = transform.pre_concat(layout_transform(group.transform));
                let child_mask = match group.clip.as_ref() {
                    Some(clip) => {
                        let Some(path) = build_path(clip) else {
                            continue;
                        };
                        let mut child_mask = match mask {
                            Some(parent) => parent.clone(),
                            None => {
                                let Some(mask) = Mask::new(pixmap.width(), pixmap.height()) else {
                                    continue;
                                };
                                mask
                            }
                        };
                        if mask.is_some() {
                            child_mask.intersect_path(
                                &path,
                                fill_rule(clip.fill_rule),
                                true,
                                child_transform,
                            );
                        } else {
                            child_mask.fill_path(
                                &path,
                                fill_rule(clip.fill_rule),
                                true,
                                child_transform,
                            );
                        }
                        Some(child_mask)
                    }
                    None => None,
                };
                let child_mask_ref = child_mask.as_ref().or(mask);
                let opacity = if group.opacity.is_finite() {
                    group.opacity.clamp(0.0, 1.0) as f32
                } else {
                    1.0
                };
                if opacity == 0.0 {
                    continue;
                }
                if opacity == 1.0 && group.effects.is_empty() {
                    render_elements(
                        pixmap,
                        fonts,
                        &group.children,
                        child_transform,
                        child_mask_ref,
                    );
                } else if let Some(mut layer) = Pixmap::new(pixmap.width(), pixmap.height()) {
                    render_elements(
                        &mut layer,
                        fonts,
                        &group.children,
                        child_transform,
                        child_mask_ref,
                    );
                    if group.effects.is_empty() {
                        pixmap.draw_pixmap(
                            0,
                            0,
                            layer.as_ref(),
                            &PixmapPaint {
                                opacity,
                                ..PixmapPaint::default()
                            },
                            Transform::identity(),
                            None,
                        );
                    } else if let Some(mut composited) =
                        Pixmap::new(pixmap.width(), pixmap.height())
                    {
                        render_group_effects(
                            &mut composited,
                            &layer,
                            &group.effects,
                            child_transform,
                            mask,
                        );
                        composited.draw_pixmap(
                            0,
                            0,
                            layer.as_ref(),
                            &PixmapPaint::default(),
                            Transform::identity(),
                            None,
                        );
                        pixmap.draw_pixmap(
                            0,
                            0,
                            composited.as_ref(),
                            &PixmapPaint {
                                opacity,
                                ..PixmapPaint::default()
                            },
                            Transform::identity(),
                            None,
                        );
                    }
                }
            }
            PositionedElement::MarkedContent { children, .. } => {
                render_elements(pixmap, fonts, children, transform, mask);
            }
            _ => {}
        }
    }
}

fn render_group_effects(
    pixmap: &mut Pixmap,
    layer: &Pixmap,
    effects: &[Effect],
    transform: Transform,
    mask: Option<&Mask>,
) {
    for effect in effects {
        if let Effect::OuterShadow {
            dx,
            dy,
            blur,
            color,
        } = effect
        {
            let Some((left, top, right, bottom)) = alpha_bounds(layer) else {
                continue;
            };
            let (scale_x, scale_y) = transform.get_scale();
            let scale = ((scale_x + scale_y) * 0.5).max(0.0);
            let blur_radius = ((*blur as f32 * scale).max(0.0).ceil() as usize)
                .min(layer.width().max(layer.height()) as usize);
            let source_width = right - left + 1;
            let source_height = bottom - top + 1;
            let width = source_width.saturating_add(blur_radius.saturating_mul(2));
            let height = source_height.saturating_add(blur_radius.saturating_mul(2));
            let Some(pixel_count) = width.checked_mul(height) else {
                continue;
            };
            let mut alpha = vec![0_u8; pixel_count];
            for y in 0..source_height {
                let source_start = ((top + y) * layer.width() as usize + left) * 4;
                let source = &layer.data()[source_start..source_start + source_width * 4];
                let target_start = (y + blur_radius) * width + blur_radius;
                for (target, pixel) in alpha[target_start..target_start + source_width]
                    .iter_mut()
                    .zip(source.chunks_exact(4))
                {
                    *target = pixel[3];
                }
            }
            box_blur_alpha(&mut alpha, width, height, blur_radius);

            let (Ok(shadow_width), Ok(shadow_height)) =
                (u32::try_from(width), u32::try_from(height))
            else {
                continue;
            };
            let Some(mut shadow) = Pixmap::new(shadow_width, shadow_height) else {
                continue;
            };
            let color_alpha = color.a.clamp(0.0, 1.0);
            for (pixel, source_alpha) in shadow.data_mut().chunks_exact_mut(4).zip(alpha) {
                let alpha = (f64::from(source_alpha) * color_alpha).round() as u8;
                pixel[0] = (color.r.clamp(0.0, 1.0) * f64::from(alpha)).round() as u8;
                pixel[1] = (color.g.clamp(0.0, 1.0) * f64::from(alpha)).round() as u8;
                pixel[2] = (color.b.clamp(0.0, 1.0) * f64::from(alpha)).round() as u8;
                pixel[3] = alpha;
            }

            let mut origin = tiny_skia::Point::from_xy(0.0, 0.0);
            let mut offset = tiny_skia::Point::from_xy(*dx as f32, *dy as f32);
            transform.map_point(&mut origin);
            transform.map_point(&mut offset);
            pixmap.draw_pixmap(
                left as i32 - blur_radius as i32 + (offset.x - origin.x).round() as i32,
                top as i32 - blur_radius as i32 + (offset.y - origin.y).round() as i32,
                shadow.as_ref(),
                &PixmapPaint::default(),
                Transform::identity(),
                mask,
            );
        }
    }
}

fn alpha_bounds(layer: &Pixmap) -> Option<(usize, usize, usize, usize)> {
    let width = layer.width() as usize;
    let mut left = width;
    let mut top = layer.height() as usize;
    let mut right = 0;
    let mut bottom = 0;
    let mut found = false;
    for (index, pixel) in layer.data().chunks_exact(4).enumerate() {
        if pixel[3] == 0 {
            continue;
        }
        let x = index % width;
        let y = index / width;
        left = left.min(x);
        top = top.min(y);
        right = right.max(x);
        bottom = bottom.max(y);
        found = true;
    }
    found.then_some((left, top, right, bottom))
}

fn box_blur_alpha(alpha: &mut [u8], width: usize, height: usize, radius: usize) {
    if radius == 0 || width == 0 || height == 0 {
        return;
    }
    let mut horizontal = vec![0_u8; alpha.len()];
    let mut prefix = vec![0_u64; width.max(height) + 1];
    let diameter = radius.saturating_mul(2).saturating_add(1) as u64;

    for y in 0..height {
        let row = &alpha[y * width..(y + 1) * width];
        let target = &mut horizontal[y * width..(y + 1) * width];
        prefix[0] = 0;
        for (index, value) in row.iter().enumerate() {
            prefix[index + 1] = prefix[index] + u64::from(*value);
        }
        for (x, target) in target.iter_mut().enumerate() {
            let left = x.saturating_sub(radius);
            let right = (x + radius + 1).min(width);
            *target = ((prefix[right] - prefix[left]) / diameter) as u8;
        }
    }

    for x in 0..width {
        prefix[0] = 0;
        for y in 0..height {
            prefix[y + 1] = prefix[y] + u64::from(horizontal[y * width + x]);
        }
        for y in 0..height {
            let top = y.saturating_sub(radius);
            let bottom = (y + radius + 1).min(height);
            alpha[y * width + x] = ((prefix[bottom] - prefix[top]) / diameter) as u8;
        }
    }
}

fn render_path(
    pixmap: &mut Pixmap,
    path: &LayoutPath,
    fill: Option<&LayoutPaint>,
    stroke: Option<&LayoutStroke>,
    transform: Transform,
    mask: Option<&Mask>,
) {
    let Some(path_geometry) = build_path(path) else {
        return;
    };
    if let Some(paint) = fill.and_then(raster_paint) {
        let paint_mask =
            fill.and_then(|paint| gradient_domain_mask(pixmap, paint, transform, mask));
        pixmap.fill_path(
            &path_geometry,
            &paint,
            fill_rule(path.fill_rule),
            transform,
            paint_mask.as_ref().or(mask),
        );
    }
    if let Some(stroke) = stroke
        && let Some(paint) = raster_paint(&stroke.paint)
    {
        let paint_mask = gradient_domain_mask(pixmap, &stroke.paint, transform, mask);
        let stroke = raster_stroke(stroke);
        pixmap.stroke_path(
            &path_geometry,
            &paint,
            &stroke,
            transform,
            paint_mask.as_ref().or(mask),
        );
    }
}

fn build_path(path: &LayoutPath) -> Option<tiny_skia::Path> {
    let mut builder = PathBuilder::new();
    for command in &path.commands {
        match command {
            PathCommand::MoveTo(point) => builder.move_to(point.x as f32, point.y as f32),
            PathCommand::LineTo(point) => builder.line_to(point.x as f32, point.y as f32),
            PathCommand::CurveTo { c1, c2, to } => builder.cubic_to(
                c1.x as f32,
                c1.y as f32,
                c2.x as f32,
                c2.y as f32,
                to.x as f32,
                to.y as f32,
            ),
            PathCommand::Close => builder.close(),
        }
    }
    builder.finish()
}

fn fill_rule(rule: oxml_layout::FillRule) -> FillRule {
    match rule {
        oxml_layout::FillRule::NonZero => FillRule::Winding,
        oxml_layout::FillRule::EvenOdd => FillRule::EvenOdd,
    }
}

fn raster_paint(paint: &LayoutPaint) -> Option<Paint<'static>> {
    match paint {
        LayoutPaint::Solid(color) => Some(solid_paint(*color)),
        LayoutPaint::Linear {
            start, end, stops, ..
        } => {
            let shader = tiny_skia::LinearGradient::new(
                tiny_skia::Point::from_xy(start.x as f32, start.y as f32),
                tiny_skia::Point::from_xy(end.x as f32, end.y as f32),
                gradient_stops(stops),
                tiny_skia::SpreadMode::Pad,
                Transform::identity(),
            )?;
            Some(Paint {
                shader,
                ..Paint::default()
            })
        }
        LayoutPaint::Radial {
            center,
            radius,
            focal,
            stops,
            ..
        } => {
            let shader = tiny_skia::RadialGradient::new(
                tiny_skia::Point::from_xy(focal.x as f32, focal.y as f32),
                0.0,
                tiny_skia::Point::from_xy(center.x as f32, center.y as f32),
                *radius as f32,
                gradient_stops(stops),
                tiny_skia::SpreadMode::Pad,
                Transform::identity(),
            )?;
            Some(Paint {
                shader,
                ..Paint::default()
            })
        }
        LayoutPaint::Tile { .. } => None,
    }
}

fn gradient_domain_mask(
    pixmap: &Pixmap,
    paint: &LayoutPaint,
    transform: Transform,
    parent: Option<&Mask>,
) -> Option<Mask> {
    let path = match paint {
        LayoutPaint::Linear {
            start, end, extend, ..
        } if *extend != (true, true) => linear_domain_path(
            pixmap.width(),
            pixmap.height(),
            *start,
            *end,
            *extend,
            transform,
        )?,
        LayoutPaint::Radial {
            center,
            radius,
            extend,
            ..
        } if !extend.1 => build_path(&LayoutPath::ellipse(oxml_layout::Rect {
            x: center.x - radius,
            y: center.y - radius,
            width: radius * 2.0,
            height: radius * 2.0,
        }))?,
        _ => return None,
    };
    let mut domain = match parent {
        Some(parent) => parent.clone(),
        None => Mask::new(pixmap.width(), pixmap.height())?,
    };
    if parent.is_some() {
        domain.intersect_path(&path, FillRule::Winding, true, transform);
    } else {
        domain.fill_path(&path, FillRule::Winding, true, transform);
    }
    Some(domain)
}

fn linear_domain_path(
    width: u32,
    height: u32,
    start: oxml_layout::Point,
    end: oxml_layout::Point,
    extend: (bool, bool),
    transform: Transform,
) -> Option<tiny_skia::Path> {
    let inverse = transform.invert()?;
    let mut corners = [
        tiny_skia::Point::from_xy(0.0, 0.0),
        tiny_skia::Point::from_xy(width as f32, 0.0),
        tiny_skia::Point::from_xy(width as f32, height as f32),
        tiny_skia::Point::from_xy(0.0, height as f32),
    ];
    inverse.map_points(&mut corners);

    let dx = (end.x - start.x) as f32;
    let dy = (end.y - start.y) as f32;
    let length_squared = dx * dx + dy * dy;
    if !length_squared.is_finite() || length_squared <= f32::EPSILON {
        return None;
    }
    let length = length_squared.sqrt();
    let normal_x = -dy / length;
    let normal_y = dx / length;
    let mut min_t = f32::INFINITY;
    let mut max_t = f32::NEG_INFINITY;
    let mut min_normal = f32::INFINITY;
    let mut max_normal = f32::NEG_INFINITY;
    for corner in corners {
        let x = corner.x - start.x as f32;
        let y = corner.y - start.y as f32;
        let t = (x * dx + y * dy) / length_squared;
        let normal = x * normal_x + y * normal_y;
        min_t = min_t.min(t);
        max_t = max_t.max(t);
        min_normal = min_normal.min(normal);
        max_normal = max_normal.max(normal);
    }
    let min_t = if extend.0 { min_t - 1.0 } else { 0.0 };
    let max_t = if extend.1 { max_t + 1.0 } else { 1.0 };
    let min_normal = min_normal - 1.0;
    let max_normal = max_normal + 1.0;
    let point = |t: f32, normal: f32| {
        tiny_skia::Point::from_xy(
            start.x as f32 + dx * t + normal_x * normal,
            start.y as f32 + dy * t + normal_y * normal,
        )
    };
    let mut builder = PathBuilder::new();
    let first = point(min_t, min_normal);
    builder.move_to(first.x, first.y);
    for corner in [
        point(max_t, min_normal),
        point(max_t, max_normal),
        point(min_t, max_normal),
    ] {
        builder.line_to(corner.x, corner.y);
    }
    builder.close();
    builder.finish()
}

fn gradient_stops(stops: &[oxml_layout::GradientStop]) -> Vec<tiny_skia::GradientStop> {
    let mut normalized: Vec<(f64, oxml_layout::Color)> = stops
        .iter()
        .map(|stop| {
            let offset = if stop.offset.is_nan() {
                0.0
            } else {
                stop.offset.clamp(0.0, 1.0)
            };
            (offset, stop.color)
        })
        .collect();
    normalized.sort_by(|left, right| left.0.total_cmp(&right.0));

    let mut deduplicated: Vec<(f64, oxml_layout::Color)> = Vec::with_capacity(normalized.len());
    for stop in normalized {
        if let Some(previous) = deduplicated.last_mut()
            && previous.0 == stop.0
        {
            *previous = stop;
        } else {
            deduplicated.push(stop);
        }
    }

    deduplicated
        .into_iter()
        .map(|(offset, color)| tiny_skia::GradientStop::new(offset as f32, skia_color(color)))
        .collect()
}

fn solid_paint(color: oxml_layout::Color) -> Paint<'static> {
    let mut paint = Paint::default();
    paint.set_color_rgba8(
        (color.r.clamp(0.0, 1.0) * 255.0) as u8,
        (color.g.clamp(0.0, 1.0) * 255.0) as u8,
        (color.b.clamp(0.0, 1.0) * 255.0) as u8,
        (color.a.clamp(0.0, 1.0) * 255.0) as u8,
    );
    paint
}

fn skia_color(color: oxml_layout::Color) -> tiny_skia::Color {
    tiny_skia::Color::from_rgba(
        color.r.clamp(0.0, 1.0) as f32,
        color.g.clamp(0.0, 1.0) as f32,
        color.b.clamp(0.0, 1.0) as f32,
        color.a.clamp(0.0, 1.0) as f32,
    )
    .unwrap_or(tiny_skia::Color::TRANSPARENT)
}

fn raster_stroke(stroke: &LayoutStroke) -> Stroke {
    Stroke {
        width: stroke.width as f32,
        line_cap: match stroke.cap {
            LineCap::Butt => tiny_skia::LineCap::Butt,
            LineCap::Round => tiny_skia::LineCap::Round,
            LineCap::Square => tiny_skia::LineCap::Square,
        },
        line_join: match stroke.join {
            LineJoin::Miter => tiny_skia::LineJoin::Miter,
            LineJoin::Round => tiny_skia::LineJoin::Round,
            LineJoin::Bevel => tiny_skia::LineJoin::Bevel,
        },
        dash: stroke.dash.as_ref().and_then(|dash| {
            StrokeDash::new(dash.iter().map(|value| *value as f32).collect(), 0.0)
        }),
        ..Stroke::default()
    }
}

fn layout_transform(transform: oxml_layout::Transform) -> Transform {
    Transform::from_row(
        transform.a as f32,
        transform.b as f32,
        transform.c as f32,
        transform.d as f32,
        transform.e as f32,
        transform.f as f32,
    )
}

/// Render a glyph run by extracting glyph outlines from the font.
fn render_glyph_run(
    pixmap: &mut Pixmap,
    glyph_run: &oxml_layout::GlyphRun,
    font_data: &oxml_layout::FontData,
    transform: Transform,
    mask: Option<&Mask>,
) {
    render_positioned_glyph_run(pixmap, glyph_run, font_data, transform, mask, &[], &[], &[]);
}

#[allow(clippy::too_many_arguments)]
fn render_positioned_glyph_run(
    pixmap: &mut Pixmap,
    glyph_run: &oxml_layout::GlyphRun,
    font_data: &oxml_layout::FontData,
    transform: Transform,
    mask: Option<&Mask>,
    x_offsets: &[f64],
    y_offsets: &[f64],
    y_advances: &[f64],
) {
    let Some(face) = crate::font::instanced_face(font_data) else {
        return;
    };

    let upem = face.units_per_em() as f64;
    let scale = glyph_run.font_size / upem;
    // A synthetic bold outline gains the PDF writer's stroke, in font units.
    let embolden = font_data.synthetic_bold.then(|| Stroke {
        width: (upem * oxml_layout::SYNTHETIC_BOLD_STROKE_EM) as f32,
        line_join: tiny_skia::LineJoin::Round,
        ..Stroke::default()
    });

    let mut paint = Paint::default();
    paint.set_color_rgba8(
        (glyph_run.color.r * 255.0) as u8,
        (glyph_run.color.g * 255.0) as u8,
        (glyph_run.color.b * 255.0) as u8,
        (glyph_run.color.a * 255.0) as u8,
    );
    paint.anti_alias = true;

    let mut x = glyph_run.origin.x;
    let mut y = glyph_run.origin.y;

    for (i, &glyph_id) in glyph_run.glyph_ids.iter().enumerate() {
        let gid = ttf_parser::GlyphId(glyph_id);

        // Build path from glyph outline
        let mut builder = GlyphPathBuilder::new();
        if face.outline_glyph(gid, &mut builder).is_some()
            && let Some(path) = builder.finish()
        {
            // Transform: translate to glyph position, scale from font units to points,
            // and flip Y (font coordinates are Y-up, pixmap is Y-down)
            let glyph_transform = transform
                .pre_translate(
                    (x + x_offsets.get(i).copied().unwrap_or(0.0)) as f32,
                    (y - y_offsets.get(i).copied().unwrap_or(0.0)) as f32,
                )
                .pre_scale(scale as f32, -(scale as f32));

            pixmap.fill_path(&path, &paint, FillRule::Winding, glyph_transform, mask);
            if let Some(stroke) = &embolden {
                pixmap.stroke_path(&path, &paint, stroke, glyph_transform, mask);
            }
        }

        if i < glyph_run.advances.len() {
            x += glyph_run.advances[i];
        }
        y -= y_advances.get(i).copied().unwrap_or(0.0);
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
    rect: &oxml_layout::Rect,
    data: &[u8],
    content_type: &str,
    transform: Transform,
    mask: Option<&Mask>,
) {
    // Decode the image
    let decoded = crate::image::decode_image(data, content_type);
    let Some(decoded) = decoded else {
        return;
    };

    // Create a pixmap from the decoded image data
    let img_pixmap = if decoded.is_jpeg {
        decode_jpeg_pixmap(data, decoded.width, decoded.height)
    } else if decoded.color_space == "DeviceRGB" {
        // Convert RGB to RGBA
        let mut rgba = Vec::with_capacity(decoded.width as usize * decoded.height as usize * 4);
        if let Some(alpha) = &decoded.alpha {
            for (rgb, &a) in decoded.data.chunks_exact(3).zip(alpha.iter()) {
                let color = tiny_skia::ColorU8::from_rgba(rgb[0], rgb[1], rgb[2], a).premultiply();
                rgba.extend_from_slice(&[color.red(), color.green(), color.blue(), color.alpha()]);
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

    // Calculate the transform to position and scale the image. A zero
    // dimension from a malformed image would make this an infinite scale.
    if decoded.width == 0 || decoded.height == 0 {
        return;
    }
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

    let paint = Paint {
        shader: pattern,
        ..Paint::default()
    };

    let fill_rect =
        tiny_skia::Rect::from_xywh(0.0, 0.0, decoded.width as f32, decoded.height as f32);
    if let Some(fill_rect) = fill_rect {
        pixmap.fill_rect(fill_rect, &paint, img_transform, mask);
    }
}

fn decode_jpeg_pixmap(data: &[u8], width: u32, height: u32) -> Option<Pixmap> {
    let options = DecoderOptions::default().jpeg_set_out_colorspace(ColorSpace::RGBA);
    let mut decoder = JpegDecoder::new_with_options(ZCursor::new(data), options);
    let rgba = decoder.decode().ok()?;
    let (decoded_width, decoded_height) = decoder.dimensions()?;
    if decoded_width != width as usize || decoded_height != height as usize {
        return None;
    }
    let size = tiny_skia::IntSize::from_wh(width, height)?;
    Pixmap::from_vec(rgba, size)
}

#[cfg(test)]
mod tests {
    use std::io::Cursor;

    use super::{
        RasterFormat, RasterOptions, RasterOutput, decode_jpeg_pixmap, render_page_to_pixmap,
        render_pages,
    };
    use oxml_layout::{
        Color, Effect, FillRule, GradientStop, GroupElement, LayoutResult, MediaId, PageFrame,
        Paint, Path, PathCommand, PathElement, Point, PositionedElement, Rect, Stroke, Transform,
    };

    fn rgb(pixel: tiny_skia::PremultipliedColorU8) -> (u8, u8, u8) {
        (pixel.red(), pixel.green(), pixel.blue())
    }

    fn color(r: f32, g: f32, b: f32) -> Color {
        Color {
            r: r.into(),
            g: g.into(),
            b: b.into(),
            a: 1.0,
        }
    }

    fn solid_path(path: Path, color: Color) -> PositionedElement {
        PositionedElement::Path(PathElement {
            path,
            fill: Some(Paint::Solid(color)),
            stroke: None,
        })
    }

    fn page(elements: Vec<PositionedElement>) -> PageFrame {
        PageFrame::new(1, 32.0, 32.0, elements)
    }

    fn layout(pages: Vec<PageFrame>) -> LayoutResult {
        LayoutResult::new(
            pages.into_iter().map(Into::into).collect(),
            Vec::new(),
            None,
            Vec::new(),
        )
    }

    fn assert_near_rgb(actual: (u8, u8, u8), expected: (u8, u8, u8), tolerance: u8) {
        for (actual, expected) in [actual.0, actual.1, actual.2]
            .into_iter()
            .zip([expected.0, expected.1, expected.2])
        {
            let delta = actual.abs_diff(expected);
            assert!(
                delta <= tolerance,
                "channel {actual} differed from {expected} by {delta}, tolerance {tolerance}"
            );
        }
    }

    fn decode_tiff_pages(bytes: &[u8]) -> Vec<(u32, u32, Vec<u8>)> {
        let mut decoder =
            tiff::decoder::Decoder::new(Cursor::new(bytes)).expect("decode TIFF header");
        let mut pages = Vec::new();
        loop {
            let (width, height) = decoder.dimensions().expect("TIFF dimensions");
            let tiff::decoder::DecodingResult::U8(data) =
                decoder.read_image().expect("read TIFF image")
            else {
                panic!("TIFF image should decode as RGB8");
            };
            pages.push((width, height, data));
            if !decoder.more_images() {
                break;
            }
            decoder.next_image().expect("advance TIFF image");
        }
        pages
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

    fn one_pixel_rgba_png(pixel: [u8; 4]) -> Vec<u8> {
        let mut ihdr = Vec::with_capacity(13);
        ihdr.extend_from_slice(&1_u32.to_be_bytes());
        ihdr.extend_from_slice(&1_u32.to_be_bytes());
        ihdr.extend_from_slice(&[8, 6, 0, 0, 0]);
        let mut png = b"\x89PNG\r\n\x1a\n".to_vec();
        append_png_chunk(&mut png, b"IHDR", &ihdr);
        append_png_chunk(
            &mut png,
            b"IDAT",
            &miniz_oxide::deflate::compress_to_vec_zlib(
                &[0, pixel[0], pixel[1], pixel[2], pixel[3]],
                6,
            ),
        );
        append_png_chunk(&mut png, b"IEND", &[]);
        png
    }

    fn transparent_pixel_layout(stored_rgb: [u8; 3]) -> LayoutResult {
        let mut frame = page(vec![
            PositionedElement::FilledRect {
                rect: Rect {
                    x: 0.0,
                    y: 0.0,
                    width: 32.0,
                    height: 32.0,
                },
                color: Color {
                    r: 0.0,
                    g: 32.0 / 255.0,
                    b: 96.0 / 255.0,
                    a: 1.0,
                },
            },
            PositionedElement::Image {
                rect: Rect {
                    x: 0.0,
                    y: 0.0,
                    width: 32.0,
                    height: 32.0,
                },
                data: one_pixel_rgba_png([stored_rgb[0], stored_rgb[1], stored_rgb[2], 0]),
                content_type: "image/png".to_owned(),
                media_id: MediaId(1),
            },
        ]);
        frame.background = None;
        layout(vec![frame])
    }

    #[test]
    fn image_export_options_produce_the_declared_formats_and_exact_pages() {
        let layout = layout(vec![
            page(vec![PositionedElement::FilledRect {
                rect: Rect {
                    x: 0.0,
                    y: 0.0,
                    width: 32.0,
                    height: 32.0,
                },
                color: color(1.0, 0.0, 0.0),
            }]),
            page(vec![PositionedElement::FilledRect {
                rect: Rect {
                    x: 0.0,
                    y: 0.0,
                    width: 32.0,
                    height: 32.0,
                },
                color: color(0.0, 1.0, 0.0),
            }]),
            page(vec![PositionedElement::FilledRect {
                rect: Rect {
                    x: 0.0,
                    y: 0.0,
                    width: 32.0,
                    height: 32.0,
                },
                color: color(0.0, 0.0, 1.0),
            }]),
        ]);

        let png = render_pages(
            &layout,
            &[2, 0],
            RasterOptions {
                dpi: 72.0,
                format: RasterFormat::Png {
                    transparent_background: false,
                },
            },
        )
        .expect("selected PNG pages render");
        let RasterOutput::SeparatePages(png_pages) = png else {
            panic!("PNG renders as separate pages");
        };
        assert_eq!(png_pages.len(), 2);
        let png_page_2 = tiny_skia::Pixmap::decode_png(&png_pages[0]).expect("decode page 2 PNG");
        let png_page_0 = tiny_skia::Pixmap::decode_png(&png_pages[1]).expect("decode page 0 PNG");
        assert_eq!((png_page_2.width(), png_page_2.height()), (32, 32));
        assert_eq!((png_page_0.width(), png_page_0.height()), (32, 32));
        assert_eq!(rgb(png_page_2.pixel(16, 16).unwrap()), (0, 0, 255));
        assert_eq!(rgb(png_page_0.pixel(16, 16).unwrap()), (255, 0, 0));

        let jpeg = render_pages(
            &layout,
            &[2, 0],
            RasterOptions {
                dpi: 72.0,
                format: RasterFormat::Jpeg { quality: 80 },
            },
        )
        .expect("selected JPEG pages render");
        let RasterOutput::SeparatePages(jpeg_pages) = jpeg else {
            panic!("JPEG renders as separate pages");
        };
        assert_eq!(jpeg_pages.len(), 2);
        let jpeg_page_2 = decode_jpeg_pixmap(&jpeg_pages[0], 32, 32).expect("decode page 2 JPEG");
        let jpeg_page_0 = decode_jpeg_pixmap(&jpeg_pages[1], 32, 32).expect("decode page 0 JPEG");
        assert_ne!(jpeg_pages[0], jpeg_pages[1]);
        assert_near_rgb(rgb(jpeg_page_2.pixel(16, 16).unwrap()), (0, 0, 255), 8);
        assert_near_rgb(rgb(jpeg_page_0.pixel(16, 16).unwrap()), (255, 0, 0), 8);

        let tiff = render_pages(
            &layout,
            &[2, 0],
            RasterOptions {
                dpi: 72.0,
                format: RasterFormat::Tiff,
            },
        )
        .expect("selected TIFF pages render");
        let RasterOutput::MultiPageTiff(tiff) = tiff else {
            panic!("TIFF renders as one multi-page stream");
        };
        assert!(tiff.starts_with(b"II*\0") || tiff.starts_with(b"MM\0*"));
        let tiff_pages = decode_tiff_pages(&tiff);
        assert_eq!(tiff_pages.len(), 2);
        assert_eq!((tiff_pages[0].0, tiff_pages[0].1), (32, 32));
        assert_eq!((tiff_pages[1].0, tiff_pages[1].1), (32, 32));
        let first_center = 3 * (16 * 32 + 16);
        let second_center = 3 * (16 * 32 + 16);
        assert_eq!(
            &tiff_pages[0].2[first_center..first_center + 3],
            &[0, 0, 255]
        );
        assert_eq!(
            &tiff_pages[1].2[second_center..second_center + 3],
            &[255, 0, 0]
        );
    }

    #[test]
    fn transparent_png_keeps_clear_pixels_but_authored_background_paints() {
        let blank = layout(vec![PageFrame::new(1, 1.0, 1.0, Vec::new())]);
        let transparent = render_pages(
            &blank,
            &[0],
            RasterOptions {
                dpi: 72.0,
                format: RasterFormat::Png {
                    transparent_background: true,
                },
            },
        )
        .expect("transparent PNG renders");
        let RasterOutput::SeparatePages(pages) = transparent else {
            panic!("PNG renders as separate pages");
        };
        let clear = tiny_skia::Pixmap::decode_png(&pages[0]).expect("decode transparent PNG");
        assert_eq!(clear.pixel(0, 0).unwrap().alpha(), 0);

        let opaque = render_pages(
            &blank,
            &[0],
            RasterOptions {
                dpi: 72.0,
                format: RasterFormat::Png {
                    transparent_background: false,
                },
            },
        )
        .expect("opaque PNG renders");
        let RasterOutput::SeparatePages(pages) = opaque else {
            panic!("PNG renders as separate pages");
        };
        let white = tiny_skia::Pixmap::decode_png(&pages[0]).expect("decode opaque PNG");
        assert_eq!(white.pixel(0, 0).unwrap().alpha(), 255);

        let mut painted = PageFrame::new(1, 1.0, 1.0, Vec::new());
        painted.background = Some(Paint::Solid(color(0.0, 0.0, 1.0)));
        let painted = render_pages(
            &layout(vec![painted]),
            &[0],
            RasterOptions {
                dpi: 72.0,
                format: RasterFormat::Png {
                    transparent_background: true,
                },
            },
        )
        .expect("painted transparent PNG renders");
        let RasterOutput::SeparatePages(pages) = painted else {
            panic!("PNG renders as separate pages");
        };
        let blue = tiny_skia::Pixmap::decode_png(&pages[0]).expect("decode painted PNG");
        assert_eq!(blue.pixel(0, 0).unwrap().alpha(), 255);
    }

    #[test]
    fn straight_alpha_images_composite_with_premultiplied_pixels() {
        let white_storage = transparent_pixel_layout([255, 255, 255]);
        let black_storage = transparent_pixel_layout([0, 0, 0]);

        for format in [
            RasterFormat::Png {
                transparent_background: false,
            },
            RasterFormat::Jpeg { quality: 90 },
            RasterFormat::Tiff,
        ] {
            let white = render_pages(&white_storage, &[0], RasterOptions { dpi: 72.0, format })
                .expect("transparent white storage renders");
            let black = render_pages(&black_storage, &[0], RasterOptions { dpi: 72.0, format })
                .expect("transparent black storage renders");
            assert_eq!(
                white, black,
                "stored RGB must not affect transparent pixels"
            );
        }
    }

    #[test]
    fn jpeg_quality_is_validated_and_changes_the_encoded_result() {
        let layout = layout(vec![page(vec![PositionedElement::FilledRect {
            rect: Rect {
                x: 0.0,
                y: 0.0,
                width: 32.0,
                height: 32.0,
            },
            color: color(1.0, 0.0, 0.0),
        }])]);

        for quality in [0, 101] {
            assert!(
                render_pages(
                    &layout,
                    &[0],
                    RasterOptions {
                        dpi: 72.0,
                        format: RasterFormat::Jpeg { quality },
                    },
                )
                .is_err()
            );
        }

        let low = render_pages(
            &layout,
            &[0],
            RasterOptions {
                dpi: 72.0,
                format: RasterFormat::Jpeg { quality: 30 },
            },
        )
        .expect("low-quality JPEG renders");
        let high = render_pages(
            &layout,
            &[0],
            RasterOptions {
                dpi: 72.0,
                format: RasterFormat::Jpeg { quality: 90 },
            },
        )
        .expect("high-quality JPEG renders");
        assert_ne!(low, high);
    }

    #[test]
    fn existing_png_entry_points_equal_opaque_option_defaults() {
        let layout = layout(vec![page(vec![PositionedElement::FilledRect {
            rect: Rect {
                x: 0.0,
                y: 0.0,
                width: 32.0,
                height: 32.0,
            },
            color: Color::BLACK,
        }])]);
        let old_page = super::render_page_to_png(&layout, 0, 72.0).expect("old page API renders");
        let new_page = render_pages(
            &layout,
            &[0],
            RasterOptions {
                dpi: 72.0,
                format: RasterFormat::Png {
                    transparent_background: false,
                },
            },
        )
        .expect("new page API renders");
        let RasterOutput::SeparatePages(pages) = new_page else {
            panic!("PNG renders as separate pages");
        };
        assert_eq!(old_page, pages[0]);
        assert_eq!(super::render_all_pages(&layout, 72.0), pages);
    }

    #[test]
    fn half_alpha_over_white_rasterises_to_midpoint() {
        let page = PageFrame::new(
            1,
            1.0,
            1.0,
            vec![PositionedElement::FilledRect {
                rect: Rect {
                    x: 0.0,
                    y: 0.0,
                    width: 1.0,
                    height: 1.0,
                },
                color: Color {
                    r: 0.0,
                    g: 0.0,
                    b: 0.0,
                    a: 0.5,
                },
            }],
        );

        let pixmap = render_page_to_pixmap(&page, &[], 72.0).expect("one-pixel page");
        let pixel = pixmap.pixel(0, 0).expect("one-pixel page has one pixel");
        assert_eq!((pixel.red(), pixel.green(), pixel.blue()), (128, 128, 128));
    }

    #[test]
    fn jpeg_is_decoded_before_raster_composition() {
        let jpeg = vec![
            0xff, 0xd8, 0xff, 0xe0, 0x00, 0x10, 0x4a, 0x46, 0x49, 0x46, 0x00, 0x01, 0x01, 0x01,
            0x00, 0x60, 0x00, 0x60, 0x00, 0x00, 0xff, 0xdb, 0x00, 0x43, 0x00, 0x08, 0x06, 0x06,
            0x07, 0x06, 0x05, 0x08, 0x07, 0x07, 0x07, 0x09, 0x09, 0x08, 0x0a, 0x0c, 0x14, 0x0d,
            0x0c, 0x0b, 0x0b, 0x0c, 0x19, 0x12, 0x13, 0x0f, 0x14, 0x1d, 0x1a, 0x1f, 0x1e, 0x1d,
            0x1a, 0x1c, 0x1c, 0x20, 0x24, 0x2e, 0x27, 0x20, 0x22, 0x2c, 0x23, 0x1c, 0x1c, 0x28,
            0x37, 0x29, 0x2c, 0x30, 0x31, 0x34, 0x34, 0x34, 0x1f, 0x27, 0x39, 0x3d, 0x38, 0x32,
            0x3c, 0x2e, 0x33, 0x34, 0x32, 0xff, 0xdb, 0x00, 0x43, 0x01, 0x09, 0x09, 0x09, 0x0c,
            0x0b, 0x0c, 0x18, 0x0d, 0x0d, 0x18, 0x32, 0x21, 0x1c, 0x21, 0x32, 0x32, 0x32, 0x32,
            0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32,
            0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32,
            0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32, 0x32,
            0x32, 0x32, 0x32, 0x32, 0xff, 0xc0, 0x00, 0x11, 0x08, 0x00, 0x19, 0x00, 0x06, 0x03,
            0x01, 0x22, 0x00, 0x02, 0x11, 0x01, 0x03, 0x11, 0x01, 0xff, 0xc4, 0x00, 0x1f, 0x00,
            0x00, 0x01, 0x05, 0x01, 0x01, 0x01, 0x01, 0x01, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0a, 0x0b,
            0xff, 0xc4, 0x00, 0xb5, 0x10, 0x00, 0x02, 0x01, 0x03, 0x03, 0x02, 0x04, 0x03, 0x05,
            0x05, 0x04, 0x04, 0x00, 0x00, 0x01, 0x7d, 0x01, 0x02, 0x03, 0x00, 0x04, 0x11, 0x05,
            0x12, 0x21, 0x31, 0x41, 0x06, 0x13, 0x51, 0x61, 0x07, 0x22, 0x71, 0x14, 0x32, 0x81,
            0x91, 0xa1, 0x08, 0x23, 0x42, 0xb1, 0xc1, 0x15, 0x52, 0xd1, 0xf0, 0x24, 0x33, 0x62,
            0x72, 0x82, 0x09, 0x0a, 0x16, 0x17, 0x18, 0x19, 0x1a, 0x25, 0x26, 0x27, 0x28, 0x29,
            0x2a, 0x34, 0x35, 0x36, 0x37, 0x38, 0x39, 0x3a, 0x43, 0x44, 0x45, 0x46, 0x47, 0x48,
            0x49, 0x4a, 0x53, 0x54, 0x55, 0x56, 0x57, 0x58, 0x59, 0x5a, 0x63, 0x64, 0x65, 0x66,
            0x67, 0x68, 0x69, 0x6a, 0x73, 0x74, 0x75, 0x76, 0x77, 0x78, 0x79, 0x7a, 0x83, 0x84,
            0x85, 0x86, 0x87, 0x88, 0x89, 0x8a, 0x92, 0x93, 0x94, 0x95, 0x96, 0x97, 0x98, 0x99,
            0x9a, 0xa2, 0xa3, 0xa4, 0xa5, 0xa6, 0xa7, 0xa8, 0xa9, 0xaa, 0xb2, 0xb3, 0xb4, 0xb5,
            0xb6, 0xb7, 0xb8, 0xb9, 0xba, 0xc2, 0xc3, 0xc4, 0xc5, 0xc6, 0xc7, 0xc8, 0xc9, 0xca,
            0xd2, 0xd3, 0xd4, 0xd5, 0xd6, 0xd7, 0xd8, 0xd9, 0xda, 0xe1, 0xe2, 0xe3, 0xe4, 0xe5,
            0xe6, 0xe7, 0xe8, 0xe9, 0xea, 0xf1, 0xf2, 0xf3, 0xf4, 0xf5, 0xf6, 0xf7, 0xf8, 0xf9,
            0xfa, 0xff, 0xc4, 0x00, 0x1f, 0x01, 0x00, 0x03, 0x01, 0x01, 0x01, 0x01, 0x01, 0x01,
            0x01, 0x01, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x02, 0x03, 0x04, 0x05,
            0x06, 0x07, 0x08, 0x09, 0x0a, 0x0b, 0xff, 0xc4, 0x00, 0xb5, 0x11, 0x00, 0x02, 0x01,
            0x02, 0x04, 0x04, 0x03, 0x04, 0x07, 0x05, 0x04, 0x04, 0x00, 0x01, 0x02, 0x77, 0x00,
            0x01, 0x02, 0x03, 0x11, 0x04, 0x05, 0x21, 0x31, 0x06, 0x12, 0x41, 0x51, 0x07, 0x61,
            0x71, 0x13, 0x22, 0x32, 0x81, 0x08, 0x14, 0x42, 0x91, 0xa1, 0xb1, 0xc1, 0x09, 0x23,
            0x33, 0x52, 0xf0, 0x15, 0x62, 0x72, 0xd1, 0x0a, 0x16, 0x24, 0x34, 0xe1, 0x25, 0xf1,
            0x17, 0x18, 0x19, 0x1a, 0x26, 0x27, 0x28, 0x29, 0x2a, 0x35, 0x36, 0x37, 0x38, 0x39,
            0x3a, 0x43, 0x44, 0x45, 0x46, 0x47, 0x48, 0x49, 0x4a, 0x53, 0x54, 0x55, 0x56, 0x57,
            0x58, 0x59, 0x5a, 0x63, 0x64, 0x65, 0x66, 0x67, 0x68, 0x69, 0x6a, 0x73, 0x74, 0x75,
            0x76, 0x77, 0x78, 0x79, 0x7a, 0x82, 0x83, 0x84, 0x85, 0x86, 0x87, 0x88, 0x89, 0x8a,
            0x92, 0x93, 0x94, 0x95, 0x96, 0x97, 0x98, 0x99, 0x9a, 0xa2, 0xa3, 0xa4, 0xa5, 0xa6,
            0xa7, 0xa8, 0xa9, 0xaa, 0xb2, 0xb3, 0xb4, 0xb5, 0xb6, 0xb7, 0xb8, 0xb9, 0xba, 0xc2,
            0xc3, 0xc4, 0xc5, 0xc6, 0xc7, 0xc8, 0xc9, 0xca, 0xd2, 0xd3, 0xd4, 0xd5, 0xd6, 0xd7,
            0xd8, 0xd9, 0xda, 0xe2, 0xe3, 0xe4, 0xe5, 0xe6, 0xe7, 0xe8, 0xe9, 0xea, 0xf2, 0xf3,
            0xf4, 0xf5, 0xf6, 0xf7, 0xf8, 0xf9, 0xfa, 0xff, 0xda, 0x00, 0x0c, 0x03, 0x01, 0x00,
            0x02, 0x11, 0x03, 0x11, 0x00, 0x3f, 0x00, 0xaf, 0xf0, 0x53, 0xfe, 0x44, 0xdb, 0xcf,
            0xfb, 0x08, 0x3f, 0xfe, 0x8b, 0x8e, 0x8a, 0x3e, 0x0a, 0x7f, 0xc8, 0x9b, 0x79, 0xff,
            0x00, 0x61, 0x07, 0xff, 0x00, 0xd1, 0x71, 0xd1, 0x40, 0x07, 0xc1, 0x4f, 0xf9, 0x13,
            0x6f, 0x3f, 0xec, 0x20, 0xff, 0x00, 0xfa, 0x2e, 0x3a, 0x28, 0xf8, 0x29, 0xff, 0x00,
            0x22, 0x6d, 0xe7, 0xfd, 0x84, 0x1f, 0xff, 0x00, 0x45, 0xc7, 0x45, 0x00, 0x7f, 0xff,
            0xd9,
        ];
        let image = PositionedElement::Image {
            rect: Rect {
                x: 0.0,
                y: 0.0,
                width: 6.0,
                height: 25.0,
            },
            data: jpeg.clone(),
            content_type: "image/jpeg".to_owned(),
            media_id: MediaId(1),
        };
        let decoded =
            decode_jpeg_pixmap(&jpeg, 6, 25).expect("valid JPEG fixture must decode to a pixmap");
        assert_eq!(rgb(decoded.pixel(0, 0).unwrap()), (1, 1, 1));
        assert_eq!(rgb(decoded.pixel(5, 0).unwrap()), (192, 192, 192));
        let pixmap = render_page_to_pixmap(&PageFrame::new(1, 6.0, 25.0, vec![image]), &[], 72.0)
            .expect("JPEG page must rasterise");

        assert_eq!(rgb(pixmap.pixel(0, 0).unwrap()), (1, 1, 1));
        assert_eq!(rgb(pixmap.pixel(5, 0).unwrap()), (192, 192, 192));
    }

    #[test]
    fn rotated_rectangle_has_a_filled_interior_and_empty_corner() {
        let element = PositionedElement::Group(GroupElement {
            transform: Transform::rotate_about(45.0, 16.0, 16.0),
            clip: None,
            opacity: 1.0,
            effects: Vec::new(),
            children: vec![solid_path(
                Path::rect(Rect {
                    x: 10.0,
                    y: 12.0,
                    width: 12.0,
                    height: 8.0,
                }),
                Color::BLACK,
            )],
        });
        let pixmap = render_page_to_pixmap(&page(vec![element]), &[], 72.0).unwrap();

        assert_eq!(rgb(pixmap.pixel(16, 16).unwrap()), (0, 0, 0));
        assert_eq!(rgb(pixmap.pixel(9, 9).unwrap()), (255, 255, 255));
    }

    #[test]
    fn dashed_line_contains_deterministic_gaps() {
        let element = PositionedElement::Line {
            start: Point { x: 2.0, y: 4.5 },
            end: Point { x: 30.0, y: 4.5 },
            width: 1.0,
            color: Color::BLACK,
            dash_pattern: Some((4.0, 4.0)),
        };
        let pixmap = render_page_to_pixmap(&page(vec![element]), &[], 72.0).unwrap();

        assert_eq!(rgb(pixmap.pixel(3, 4).unwrap()), (0, 0, 0));
        assert_eq!(rgb(pixmap.pixel(7, 4).unwrap()), (255, 255, 255));
        assert_eq!(rgb(pixmap.pixel(11, 4).unwrap()), (0, 0, 0));
    }

    #[test]
    fn nested_group_transforms_apply_child_before_parent() {
        let translated = |transform, children| {
            PositionedElement::Group(GroupElement {
                transform,
                clip: None,
                opacity: 1.0,
                effects: Vec::new(),
                children,
            })
        };
        let leaf = solid_path(
            Path::rect(Rect {
                x: 0.0,
                y: 0.0,
                width: 2.0,
                height: 2.0,
            }),
            Color::BLACK,
        );
        let element = translated(
            Transform {
                e: 4.0,
                ..Transform::IDENTITY
            },
            vec![translated(
                Transform {
                    e: 3.0,
                    ..Transform::IDENTITY
                },
                vec![translated(
                    Transform {
                        f: 5.0,
                        ..Transform::IDENTITY
                    },
                    vec![leaf],
                )],
            )],
        );
        let pixmap = render_page_to_pixmap(&page(vec![element]), &[], 72.0).unwrap();

        assert_eq!(rgb(pixmap.pixel(7, 5).unwrap()), (0, 0, 0));
        assert_eq!(rgb(pixmap.pixel(1, 1).unwrap()), (255, 255, 255));
    }

    #[test]
    fn group_clip_masks_children_and_preserves_outside_pixels() {
        let element = PositionedElement::Group(GroupElement {
            transform: Transform::IDENTITY,
            clip: Some(Path::rect(Rect {
                x: 4.0,
                y: 4.0,
                width: 8.0,
                height: 12.0,
            })),
            opacity: 1.0,
            effects: Vec::new(),
            children: vec![PositionedElement::Group(GroupElement {
                transform: Transform::IDENTITY,
                clip: Some(Path::rect(Rect {
                    x: 8.0,
                    y: 4.0,
                    width: 8.0,
                    height: 12.0,
                })),
                opacity: 1.0,
                effects: Vec::new(),
                children: vec![solid_path(
                    Path::rect(Rect {
                        x: 0.0,
                        y: 0.0,
                        width: 24.0,
                        height: 24.0,
                    }),
                    Color::BLACK,
                )],
            })],
        });
        let pixmap = render_page_to_pixmap(&page(vec![element]), &[], 72.0).unwrap();

        assert_eq!(rgb(pixmap.pixel(10, 10).unwrap()), (0, 0, 0));
        assert_eq!(rgb(pixmap.pixel(6, 10).unwrap()), (255, 255, 255));
        assert_eq!(rgb(pixmap.pixel(14, 10).unwrap()), (255, 255, 255));
    }

    #[test]
    fn group_opacity_is_applied_once_to_the_composited_subtree() {
        let element = PositionedElement::Group(GroupElement {
            transform: Transform::IDENTITY,
            clip: None,
            opacity: 0.5,
            effects: Vec::new(),
            children: vec![
                solid_path(
                    Path::rect(Rect {
                        x: 2.0,
                        y: 2.0,
                        width: 12.0,
                        height: 12.0,
                    }),
                    Color::BLACK,
                ),
                solid_path(
                    Path::rect(Rect {
                        x: 8.0,
                        y: 2.0,
                        width: 12.0,
                        height: 12.0,
                    }),
                    Color::BLACK,
                ),
            ],
        });
        let pixmap = render_page_to_pixmap(&page(vec![element]), &[], 72.0).unwrap();

        assert_eq!(rgb(pixmap.pixel(4, 4).unwrap()), (128, 128, 128));
        assert_eq!(rgb(pixmap.pixel(10, 4).unwrap()), (128, 128, 128));
    }

    #[test]
    fn outer_shadow_is_coloured_offset_and_behind_group_content() {
        let element = PositionedElement::Group(GroupElement {
            transform: Transform::IDENTITY,
            clip: None,
            opacity: 1.0,
            effects: vec![Effect::OuterShadow {
                dx: 4.0,
                dy: 0.0,
                blur: 0.0,
                color: Color {
                    r: 1.0,
                    g: 0.0,
                    b: 0.0,
                    a: 1.0,
                },
            }],
            children: vec![solid_path(
                Path::rect(Rect {
                    x: 4.0,
                    y: 4.0,
                    width: 8.0,
                    height: 8.0,
                }),
                Color::BLACK,
            )],
        });
        let pixmap = render_page_to_pixmap(&page(vec![element]), &[], 72.0).unwrap();

        assert_eq!(rgb(pixmap.pixel(6, 6).unwrap()), (0, 0, 0));
        assert_eq!(rgb(pixmap.pixel(14, 6).unwrap()), (255, 0, 0));
        assert_eq!(rgb(pixmap.pixel(2, 6).unwrap()), (255, 255, 255));
    }

    #[test]
    fn outer_shadow_blur_extends_alpha_and_obeys_group_opacity() {
        let element = PositionedElement::Group(GroupElement {
            transform: Transform::IDENTITY,
            clip: None,
            opacity: 0.5,
            effects: vec![Effect::OuterShadow {
                dx: 0.0,
                dy: 0.0,
                blur: 2.0,
                color: Color::BLACK,
            }],
            children: vec![solid_path(
                Path::rect(Rect {
                    x: 12.0,
                    y: 12.0,
                    width: 4.0,
                    height: 4.0,
                }),
                Color::BLACK,
            )],
        });
        let pixmap = render_page_to_pixmap(&page(vec![element]), &[], 72.0).unwrap();

        assert_eq!(rgb(pixmap.pixel(13, 13).unwrap()), (128, 128, 128));
        let blurred = rgb(pixmap.pixel(10, 13).unwrap());
        assert!(blurred.0 < 255 && blurred.0 > 128, "{blurred:?}");
        assert_eq!(blurred.0, blurred.1);
        assert_eq!(blurred.1, blurred.2);
        assert_eq!(rgb(pixmap.pixel(8, 13).unwrap()), (255, 255, 255));
    }

    #[test]
    fn path_fill_rule_selects_the_tiny_skia_rule() {
        let nested = |fill_rule| Path {
            commands: vec![
                PathCommand::MoveTo(Point { x: 2.0, y: 2.0 }),
                PathCommand::LineTo(Point { x: 14.0, y: 2.0 }),
                PathCommand::LineTo(Point { x: 14.0, y: 14.0 }),
                PathCommand::LineTo(Point { x: 2.0, y: 14.0 }),
                PathCommand::Close,
                PathCommand::MoveTo(Point { x: 5.0, y: 5.0 }),
                PathCommand::LineTo(Point { x: 11.0, y: 5.0 }),
                PathCommand::LineTo(Point { x: 11.0, y: 11.0 }),
                PathCommand::LineTo(Point { x: 5.0, y: 11.0 }),
                PathCommand::Close,
            ],
            fill_rule,
        };
        let nonzero = render_page_to_pixmap(
            &page(vec![solid_path(nested(FillRule::NonZero), Color::BLACK)]),
            &[],
            72.0,
        )
        .unwrap();
        let evenodd = render_page_to_pixmap(
            &page(vec![solid_path(nested(FillRule::EvenOdd), Color::BLACK)]),
            &[],
            72.0,
        )
        .unwrap();

        assert_eq!(rgb(nonzero.pixel(8, 8).unwrap()), (0, 0, 0));
        assert_eq!(rgb(evenodd.pixel(8, 8).unwrap()), (255, 255, 255));
    }

    #[test]
    fn linear_and_radial_gradients_sample_expected_colours() {
        let stops = vec![
            GradientStop {
                offset: 0.0,
                color: Color {
                    r: 1.0,
                    g: 0.0,
                    b: 0.0,
                    a: 1.0,
                },
            },
            GradientStop {
                offset: 1.0,
                color: Color {
                    r: 0.0,
                    g: 0.0,
                    b: 1.0,
                    a: 1.0,
                },
            },
        ];
        let linear = PositionedElement::Path(PathElement {
            path: Path::rect(Rect {
                x: 0.0,
                y: 0.0,
                width: 16.0,
                height: 8.0,
            }),
            fill: Some(Paint::linear(
                Point { x: 0.0, y: 0.0 },
                Point { x: 16.0, y: 0.0 },
                stops.clone(),
                (true, true),
            )),
            stroke: None,
        });
        let radial = PositionedElement::Path(PathElement {
            path: Path::rect(Rect {
                x: 0.0,
                y: 16.0,
                width: 16.0,
                height: 16.0,
            }),
            fill: Some(Paint::radial(
                Point { x: 8.0, y: 24.0 },
                8.0,
                Point { x: 8.0, y: 24.0 },
                stops,
                (true, true),
            )),
            stroke: None,
        });
        let pixmap = render_page_to_pixmap(&page(vec![linear, radial]), &[], 72.0).unwrap();

        let linear_start = rgb(pixmap.pixel(1, 4).unwrap());
        let linear_end = rgb(pixmap.pixel(14, 4).unwrap());
        assert!(linear_start.0 > linear_start.2);
        assert!(linear_end.2 > linear_end.0);
        let radial_center = rgb(pixmap.pixel(8, 24).unwrap());
        let radial_edge = rgb(pixmap.pixel(1, 24).unwrap());
        assert!(radial_center.0 > radial_center.2);
        assert!(radial_edge.2 > radial_edge.0);
    }

    #[test]
    fn translated_group_gradient_uses_local_coordinates_exactly_once() {
        let red = Color {
            r: 1.0,
            g: 0.0,
            b: 0.0,
            a: 1.0,
        };
        let blue = Color {
            r: 0.0,
            g: 0.0,
            b: 1.0,
            a: 1.0,
        };
        let element = PositionedElement::Group(GroupElement {
            transform: Transform {
                e: 12.0,
                ..Transform::IDENTITY
            },
            clip: None,
            opacity: 1.0,
            effects: Vec::new(),
            children: vec![PositionedElement::Path(PathElement {
                path: Path::rect(Rect {
                    x: 0.0,
                    y: 2.0,
                    width: 8.0,
                    height: 8.0,
                }),
                fill: Some(Paint::linear(
                    Point { x: 0.0, y: 0.0 },
                    Point { x: 8.0, y: 0.0 },
                    vec![
                        GradientStop {
                            offset: 0.0,
                            color: red,
                        },
                        GradientStop {
                            offset: 1.0,
                            color: blue,
                        },
                    ],
                    (true, true),
                )),
                stroke: None,
            })],
        });
        let pixmap = render_page_to_pixmap(&page(vec![element]), &[], 72.0).unwrap();

        let start = rgb(pixmap.pixel(13, 6).unwrap());
        let end = rgb(pixmap.pixel(18, 6).unwrap());
        assert!(
            start.0 > start.2,
            "expected red-dominant start, got {start:?}"
        );
        assert!(end.2 > end.0, "expected blue-dominant end, got {end:?}");
        assert_eq!(rgb(pixmap.pixel(9, 6).unwrap()), (255, 255, 255));
    }

    #[test]
    fn gradient_extend_flags_leave_outside_regions_unpainted() {
        let stops = vec![
            GradientStop {
                offset: 0.0,
                color: Color::BLACK,
            },
            GradientStop {
                offset: 1.0,
                color: Color::WHITE,
            },
        ];
        let linear = PositionedElement::Path(PathElement {
            path: Path::rect(Rect {
                x: 0.0,
                y: 0.0,
                width: 32.0,
                height: 12.0,
            }),
            fill: Some(Paint::linear(
                Point { x: 8.0, y: 0.0 },
                Point { x: 24.0, y: 0.0 },
                stops.clone(),
                (false, false),
            )),
            stroke: None,
        });
        let radial = PositionedElement::Path(PathElement {
            path: Path::rect(Rect {
                x: 0.0,
                y: 16.0,
                width: 32.0,
                height: 16.0,
            }),
            fill: Some(Paint::radial(
                Point { x: 16.0, y: 24.0 },
                6.0,
                Point { x: 16.0, y: 24.0 },
                stops,
                (false, false),
            )),
            stroke: None,
        });
        let pixmap = render_page_to_pixmap(&page(vec![linear, radial]), &[], 72.0).unwrap();

        assert_eq!(rgb(pixmap.pixel(4, 6).unwrap()), (255, 255, 255));
        assert_ne!(rgb(pixmap.pixel(12, 6).unwrap()), (255, 255, 255));
        assert_ne!(rgb(pixmap.pixel(16, 24).unwrap()), (255, 255, 255));
        assert_eq!(rgb(pixmap.pixel(24, 24).unwrap()), (255, 255, 255));
    }

    #[test]
    fn gradient_stops_are_sorted_clamped_and_last_repeated_offset_wins() {
        let element = PositionedElement::Path(PathElement {
            path: Path::rect(Rect {
                x: 0.0,
                y: 0.0,
                width: 16.0,
                height: 8.0,
            }),
            fill: Some(Paint::linear(
                Point { x: 0.0, y: 0.0 },
                Point { x: 16.0, y: 0.0 },
                vec![
                    GradientStop {
                        offset: 2.0,
                        color: Color {
                            r: 0.0,
                            g: 0.0,
                            b: 1.0,
                            a: 1.0,
                        },
                    },
                    GradientStop {
                        offset: -1.0,
                        color: Color {
                            r: 1.0,
                            g: 0.0,
                            b: 0.0,
                            a: 1.0,
                        },
                    },
                    GradientStop {
                        offset: f64::NAN,
                        color: Color {
                            r: 0.0,
                            g: 1.0,
                            b: 0.0,
                            a: 1.0,
                        },
                    },
                ],
                (true, true),
            )),
            stroke: None,
        });
        let pixmap = render_page_to_pixmap(&page(vec![element]), &[], 72.0).unwrap();

        let start = rgb(pixmap.pixel(1, 4).unwrap());
        let end = rgb(pixmap.pixel(14, 4).unwrap());
        assert!(start.1 > start.0 && start.1 > start.2);
        assert!(end.2 > end.0 && end.2 > end.1);
    }

    #[test]
    fn page_background_replaces_the_white_default() {
        let mut painted = page(Vec::new());
        painted.background = Some(Paint::Solid(Color {
            r: 0.0,
            g: 1.0,
            b: 0.0,
            a: 1.0,
        }));
        let painted = render_page_to_pixmap(&painted, &[], 72.0).unwrap();
        let defaulted = render_page_to_pixmap(&page(Vec::new()), &[], 72.0).unwrap();
        let mut gradient = page(Vec::new());
        gradient.background = Some(Paint::linear(
            Point { x: 0.0, y: 0.0 },
            Point { x: 32.0, y: 0.0 },
            vec![
                GradientStop {
                    offset: 0.0,
                    color: Color::BLACK,
                },
                GradientStop {
                    offset: 1.0,
                    color: Color::WHITE,
                },
            ],
            (true, true),
        ));
        let gradient = render_page_to_pixmap(&gradient, &[], 72.0).unwrap();

        assert_eq!(rgb(painted.pixel(4, 4).unwrap()), (0, 255, 0));
        assert_eq!(rgb(defaulted.pixel(4, 4).unwrap()), (255, 255, 255));
        assert!(gradient.pixel(4, 4).unwrap().red() < gradient.pixel(28, 4).unwrap().red());
    }

    #[test]
    fn path_dash_array_contains_deterministic_gaps() {
        let mut stroke = Stroke::new(Paint::Solid(Color::BLACK), 1.0);
        stroke.dash = Some(vec![4.0, 4.0]);
        let element = PositionedElement::Path(PathElement {
            path: Path {
                commands: vec![
                    PathCommand::MoveTo(Point { x: 2.0, y: 8.5 }),
                    PathCommand::LineTo(Point { x: 30.0, y: 8.5 }),
                ],
                fill_rule: FillRule::NonZero,
            },
            fill: None,
            stroke: Some(stroke),
        });
        let pixmap = render_page_to_pixmap(&page(vec![element]), &[], 72.0).unwrap();

        assert_eq!(rgb(pixmap.pixel(3, 8).unwrap()), (0, 0, 0));
        assert_eq!(rgb(pixmap.pixel(7, 8).unwrap()), (255, 255, 255));
    }

    /// Issue 297: a bold run draws more ink than a regular one, through a
    /// variable face's bold instance or a regular-only face's synthetic bold.
    #[test]
    fn caller_bold_runs_draw_bolder_than_regular_runs() {
        let mut fonts = oxml_layout::FontManager::new_deterministic().unwrap();
        fonts.load_additional_fonts(&[
            oxml_layout::FontFile {
                family: "Caller Variable".to_owned(),
                data: include_bytes!("../../oxml-layout/fonts/NotoSansSC-FX058-subset.ttf")
                    .to_vec(),
            },
            oxml_layout::FontFile {
                family: "Caller Mono".to_owned(),
                data: include_bytes!("../../oxml-layout/fonts/LiberationMono-Regular.ttf").to_vec(),
            },
        ]);
        for family in ["Caller Variable", "Caller Mono"] {
            let mut ink = Vec::new();
            for bold in [false, true] {
                let font_id = fonts.resolve_font(Some(family), bold, false).unwrap();
                let shaped = fonts.shape_text(font_id, "Hill bold", 24.0).unwrap();
                let run = PositionedElement::Text(oxml_layout::GlyphRun {
                    origin: Point { x: 10.0, y: 40.0 },
                    font_id,
                    font_size: 24.0,
                    glyph_ids: shaped.glyph_ids,
                    advances: shaped.advances,
                    text: "Hill bold".to_owned(),
                    source: None,
                    color: Color::BLACK,
                    bold,
                    italic: false,
                    field_kind: None,
                    field_source: None,
                    tab_aligned: None,
                    note: None,
                    note_reference_source: None,
                });
                let page = PageFrame::new(1, 200.0, 60.0, vec![run]);
                let pixmap = render_page_to_pixmap(&page, &fonts.all_font_data(), 144.0).unwrap();
                ink.push(
                    pixmap
                        .pixels()
                        .iter()
                        .map(|pixel| 255 - u32::from(pixel.red()))
                        .sum::<u32>(),
                );
            }
            assert!(
                f64::from(ink[1]) > f64::from(ink[0]) * 1.15,
                "{family}: {ink:?}"
            );
        }
    }
}
