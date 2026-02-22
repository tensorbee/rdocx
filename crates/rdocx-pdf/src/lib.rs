//! PDF renderer for rdocx layout output.
//!
//! Converts `LayoutResult` (positioned page frames with glyph runs, lines,
//! rectangles, and images) into a PDF document.

mod font;
mod image;
pub mod raster;
mod writer;

use rdocx_layout::LayoutResult;

/// Render a laid-out document to PDF bytes.
///
/// The `LayoutResult` must contain all pages, fonts, metadata, and outlines
/// produced by `rdocx_layout::layout_document()`.
pub fn render_to_pdf(layout: &LayoutResult) -> Vec<u8> {
    writer::write_pdf(layout)
}

/// Render a single page to PNG bytes.
///
/// # Arguments
/// * `layout` - The layout result from `rdocx_layout::layout_document()`
/// * `page_index` - 0-based page index
/// * `dpi` - Resolution (72 = 1:1, 150 = standard, 300 = high quality)
pub fn render_page_to_png(layout: &LayoutResult, page_index: usize, dpi: f64) -> Option<Vec<u8>> {
    raster::render_page_to_png(layout, page_index, dpi)
}

/// Render all pages to PNG bytes.
pub fn render_all_pages(layout: &LayoutResult, dpi: f64) -> Vec<Vec<u8>> {
    raster::render_all_pages(layout, dpi)
}

#[cfg(test)]
mod tests {
    use super::*;
    use rdocx_layout::output::*;

    #[test]
    fn render_empty_layout() {
        let layout = LayoutResult {
            pages: vec![PageFrame {
                page_number: 1,
                width: 612.0,
                height: 792.0,
                elements: vec![],
            }],
            fonts: vec![],
            metadata: None,
            outlines: vec![],
        };

        let pdf = render_to_pdf(&layout);
        assert!(pdf.starts_with(b"%PDF"));
        let pdf_str = String::from_utf8_lossy(&pdf);
        assert!(pdf_str.contains("%%EOF"));
    }

    #[test]
    fn render_with_lines_and_rects() {
        let layout = LayoutResult {
            pages: vec![PageFrame {
                page_number: 1,
                width: 612.0,
                height: 792.0,
                elements: vec![
                    PositionedElement::Line {
                        start: Point { x: 72.0, y: 72.0 },
                        end: Point { x: 540.0, y: 72.0 },
                        width: 1.0,
                        color: Color::BLACK,
                        dash_pattern: None,
                    },
                    PositionedElement::FilledRect {
                        rect: Rect {
                            x: 72.0,
                            y: 100.0,
                            width: 468.0,
                            height: 20.0,
                        },
                        color: Color {
                            r: 0.9,
                            g: 0.9,
                            b: 0.9,
                            a: 1.0,
                        },
                    },
                ],
            }],
            fonts: vec![],
            metadata: None,
            outlines: vec![],
        };

        let pdf = render_to_pdf(&layout);
        assert!(pdf.starts_with(b"%PDF"));
        assert!(pdf.len() > 100);
    }

    #[test]
    fn render_with_metadata() {
        let layout = LayoutResult {
            pages: vec![PageFrame {
                page_number: 1,
                width: 612.0,
                height: 792.0,
                elements: vec![],
            }],
            fonts: vec![],
            metadata: Some(DocumentMetadata {
                title: Some("Test Title".to_string()),
                author: Some("Test Author".to_string()),
                subject: None,
                keywords: None,
                creator: Some("rdocx".to_string()),
            }),
            outlines: vec![],
        };

        let pdf = render_to_pdf(&layout);
        let pdf_str = String::from_utf8_lossy(&pdf);
        assert!(pdf_str.contains("Test Title"));
        assert!(pdf_str.contains("Test Author"));
    }

    #[test]
    fn render_with_link_annotation() {
        let layout = LayoutResult {
            pages: vec![PageFrame {
                page_number: 1,
                width: 612.0,
                height: 792.0,
                elements: vec![PositionedElement::LinkAnnotation {
                    rect: Rect {
                        x: 72.0,
                        y: 100.0,
                        width: 100.0,
                        height: 15.0,
                    },
                    url: "https://example.com".to_string(),
                }],
            }],
            fonts: vec![],
            metadata: None,
            outlines: vec![],
        };

        let pdf = render_to_pdf(&layout);
        let pdf_str = String::from_utf8_lossy(&pdf);
        assert!(pdf_str.contains("example.com"));
    }

    #[test]
    fn render_with_outlines() {
        let layout = LayoutResult {
            pages: vec![PageFrame {
                page_number: 1,
                width: 612.0,
                height: 792.0,
                elements: vec![],
            }],
            fonts: vec![],
            metadata: None,
            outlines: vec![
                OutlineEntry {
                    title: "Chapter 1".to_string(),
                    level: 1,
                    page_index: 0,
                    y_position: 72.0,
                },
                OutlineEntry {
                    title: "Section 1.1".to_string(),
                    level: 2,
                    page_index: 0,
                    y_position: 200.0,
                },
            ],
        };

        let pdf = render_to_pdf(&layout);
        let pdf_str = String::from_utf8_lossy(&pdf);
        assert!(pdf_str.contains("Chapter 1"));
        assert!(pdf_str.contains("Section 1.1"));
    }
}
