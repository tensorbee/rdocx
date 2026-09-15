//! PDF renderer for format-neutral OOXML layout output.
//!
//! Converts `LayoutResult` (positioned page frames with glyph runs, lines,
//! rectangles, and images) into a PDF document.

mod conformance;
mod font;
mod image;
pub mod raster;
mod structure;
mod writer;

use oxml_layout::LayoutResult;

pub use conformance::{PdfConformance, PdfError, PdfOptions};
pub use raster::{RasterError, RasterFormat, RasterOptions, RasterOutput};

/// Render a laid-out document to PDF bytes.
///
/// The `LayoutResult` must contain all pages, fonts, metadata, and outlines
/// produced by a format-specific layout engine.
pub fn render_to_pdf(layout: &LayoutResult) -> Vec<u8> {
    writer::write_pdf(layout)
}

/// Render a laid-out document to one explicitly selected archival PDF profile.
///
/// The complete layout is validated before the writer allocates output bytes.
pub fn render_to_pdf_with_options(
    layout: &LayoutResult,
    options: PdfOptions,
) -> Result<Vec<u8>, PdfError> {
    conformance::preflight(layout)?;
    Ok(writer::write_archival_pdf(layout, options.profile))
}

/// Render a single page to PNG bytes.
///
/// # Arguments
/// * `layout` - The format-neutral layout result to rasterise
/// * `page_index` - 0-based page index
/// * `dpi` - Resolution (72 = 1:1, 150 = standard, 300 = high quality)
pub fn render_page_to_png(layout: &LayoutResult, page_index: usize, dpi: f64) -> Option<Vec<u8>> {
    raster::render_page_to_png(layout, page_index, dpi)
}

/// Render all pages to PNG bytes.
pub fn render_all_pages(layout: &LayoutResult, dpi: f64) -> Vec<Vec<u8>> {
    raster::render_all_pages(layout, dpi)
}

/// Render selected zero-based pages to the requested raster image format.
///
/// Page selections preserve caller order. Empty selections, duplicate pages,
/// out-of-range pages, invalid DPI, and invalid format options are rejected
/// before any output is encoded.
pub fn render_pages(
    layout: &LayoutResult,
    page_indices: &[usize],
    options: RasterOptions,
) -> Result<RasterOutput, RasterError> {
    raster::render_pages(layout, page_indices, options)
}

#[cfg(test)]
mod tests {
    use super::*;
    use oxml_layout::{
        Color, DocumentMetadata, DocumentStructure, FontId, GlyphRun, LayoutResult, MediaId,
        OutlineEntry, PageFrame, Paint, Point, PositionedElement, Rect, StructureId, StructureNode,
        StructureRole, Transform,
    };

    fn layout_with(elements: Vec<PositionedElement>) -> LayoutResult {
        LayoutResult::new(
            vec![PageFrame::new(1, 612.0, 792.0, elements).into()],
            vec![],
            None,
            vec![],
        )
    }

    #[test]
    fn render_empty_layout() {
        let pdf = render_to_pdf(&layout_with(vec![]));
        assert!(pdf.starts_with(b"%PDF"));
        assert!(String::from_utf8_lossy(&pdf).contains("%%EOF"));
    }

    #[test]
    fn render_with_lines_and_rects() {
        let elements = vec![
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
        ];
        let pdf = render_to_pdf(&layout_with(elements));
        assert!(pdf.starts_with(b"%PDF"));
        assert!(pdf.len() > 100);
    }

    #[test]
    fn tagging_preserves_visible_pdf_and_raster_output() {
        let leaf = PositionedElement::FilledRect {
            rect: Rect {
                x: 72.0,
                y: 100.0,
                width: 120.0,
                height: 30.0,
            },
            color: Color {
                r: 0.25,
                g: 0.5,
                b: 0.75,
                a: 1.0,
            },
        };
        let plain = layout_with(vec![leaf.clone()]);
        let document = StructureId::new(1).expect("non-zero document id");
        let paragraph = StructureId::new(2).expect("non-zero paragraph id");
        let mut tagged = layout_with(vec![PositionedElement::MarkedContent {
            structure: Some(paragraph),
            children: vec![leaf],
        }]);
        tagged.structure = Some(DocumentStructure {
            root: document,
            nodes: vec![
                StructureNode {
                    id: document,
                    role: StructureRole::Document,
                    children: vec![paragraph],
                    alternate_text: None,
                },
                StructureNode {
                    id: paragraph,
                    role: StructureRole::Paragraph,
                    children: Vec::new(),
                    alternate_text: None,
                },
            ],
        });

        assert_eq!(
            render_page_to_png(&plain, 0, 144.0),
            render_page_to_png(&tagged, 0, 144.0)
        );
        assert_ne!(render_to_pdf(&plain), render_to_pdf(&tagged));
    }

    #[test]
    fn render_with_metadata() {
        let layout = LayoutResult::new(
            vec![PageFrame::new(1, 612.0, 792.0, vec![]).into()],
            vec![],
            Some(DocumentMetadata {
                title: Some("Test Title".to_owned()),
                author: Some("Test Author".to_owned()),
                subject: None,
                keywords: None,
                creator: Some("oxml-pdf".to_owned()),
            }),
            vec![],
        );
        let pdf = render_to_pdf(&layout);
        let pdf = String::from_utf8_lossy(&pdf);
        assert!(pdf.contains("Test Title"));
        assert!(pdf.contains("Test Author"));
    }

    #[test]
    fn render_with_link_annotation() {
        let layout = layout_with(vec![PositionedElement::LinkAnnotation {
            rect: Rect {
                x: 72.0,
                y: 100.0,
                width: 100.0,
                height: 15.0,
            },
            url: "https://example.com".to_owned(),
        }]);
        let pdf = render_to_pdf(&layout);
        let pdf = String::from_utf8_lossy(&pdf);
        assert!(pdf.contains("example.com"));
    }

    #[test]
    fn render_with_outlines() {
        let layout = LayoutResult::new(
            vec![PageFrame::new(1, 612.0, 792.0, vec![]).into()],
            vec![],
            None,
            vec![
                OutlineEntry {
                    title: "Chapter 1".to_owned(),
                    level: 1,
                    page_index: 0,
                    y_position: 72.0,
                },
                OutlineEntry {
                    title: "Section 1.1".to_owned(),
                    level: 2,
                    page_index: 0,
                    y_position: 200.0,
                },
            ],
        );
        let pdf = render_to_pdf(&layout);
        let pdf = String::from_utf8_lossy(&pdf);
        assert!(pdf.contains("Chapter 1"));
        assert!(pdf.contains("Section 1.1"));
    }

    #[test]
    fn pdfa_profiles_emit_matching_xmp_and_output_intent() {
        for (profile, part) in [(PdfConformance::PdfA2b, "2"), (PdfConformance::PdfA3b, "3")] {
            let pdf = render_to_pdf_with_options(&layout_with(vec![]), PdfOptions { profile })
                .expect("render conforming archival PDF");
            let text = String::from_utf8_lossy(&pdf);
            assert!(text.contains(&format!("<pdfaid:part>{part}</pdfaid:part>")));
            assert!(text.contains("<pdfaid:conformance>B</pdfaid:conformance>"));
            assert!(text.contains("/Type /Metadata"));
            assert!(text.contains("/Type /OutputIntent"));
            assert!(text.contains("/S /GTS_PDFA1"));
            assert!(text.contains("/OutputConditionIdentifier (sRGB2014)"));
        }
    }

    #[test]
    fn pdfa_rejects_prohibited_or_incomplete_features_before_output() {
        let missing_font = layout_with(vec![PositionedElement::Text(GlyphRun {
            origin: Point { x: 10.0, y: 10.0 },
            font_id: FontId(91),
            font_size: 12.0,
            glyph_ids: vec![1],
            advances: vec![6.0],
            text: "A".to_owned(),
            source: None,
            color: Color::BLACK,
            bold: false,
            italic: false,
            field_kind: None,
            field_source: None,
            note: None,
        })]);
        assert_eq!(
            render_to_pdf_with_options(&missing_font, PdfOptions::new(PdfConformance::PdfA2b)),
            Err(PdfError::MissingEmbeddedFont { font_id: 91 })
        );

        let forbidden_action = layout_with(vec![PositionedElement::LinkAnnotation {
            rect: Rect {
                x: 0.0,
                y: 0.0,
                width: 10.0,
                height: 10.0,
            },
            url: "javascript:alert(1)".to_owned(),
        }]);
        assert!(matches!(
            render_to_pdf_with_options(&forbidden_action, PdfOptions::new(PdfConformance::PdfA2b)),
            Err(PdfError::ForbiddenAction { .. })
        ));

        let unsupported_color = layout_with(vec![PositionedElement::FilledRect {
            rect: Rect {
                x: 0.0,
                y: 0.0,
                width: 10.0,
                height: 10.0,
            },
            color: Color {
                r: f64::NAN,
                ..Color::BLACK
            },
        }]);
        assert_eq!(
            render_to_pdf_with_options(&unsupported_color, PdfOptions::new(PdfConformance::PdfA2b)),
            Err(PdfError::UnsupportedColor)
        );

        let root = StructureId::new(1).unwrap();
        let paragraph = StructureId::new(2).unwrap();
        let mut invalid_structure = layout_with(vec![PositionedElement::MarkedContent {
            structure: Some(paragraph),
            children: vec![PositionedElement::FilledRect {
                rect: Rect {
                    x: 0.0,
                    y: 0.0,
                    width: 10.0,
                    height: 10.0,
                },
                color: Color::BLACK,
            }],
        }]);
        invalid_structure.structure = Some(DocumentStructure {
            root,
            nodes: vec![StructureNode {
                id: root,
                role: StructureRole::Document,
                children: vec![paragraph],
                alternate_text: None,
            }],
        });
        assert_eq!(
            render_to_pdf_with_options(&invalid_structure, PdfOptions::new(PdfConformance::PdfA3b)),
            Err(PdfError::InvalidStructure)
        );

        let mut tile_layout = layout_with(Vec::new());
        std::sync::Arc::make_mut(&mut tile_layout.pages[0]).background = Some(Paint::Tile {
            image: MediaId(1),
            tile: Rect {
                x: 0.0,
                y: 0.0,
                width: 10.0,
                height: 10.0,
            },
            transform: Transform::IDENTITY,
        });
        assert_eq!(
            render_to_pdf_with_options(&tile_layout, PdfOptions::new(PdfConformance::PdfA2b)),
            Err(PdfError::UnsupportedPaint { paint: "tile" })
        );
    }

    #[test]
    fn pdfa_link_annotations_set_the_print_flag() {
        let layout = layout_with(vec![PositionedElement::LinkAnnotation {
            rect: Rect {
                x: 0.0,
                y: 0.0,
                width: 10.0,
                height: 10.0,
            },
            url: "https://example.com".to_owned(),
        }]);
        let ordinary = String::from_utf8_lossy(&render_to_pdf(&layout)).into_owned();
        assert!(!ordinary.contains("/F 4"), "{ordinary}");
        for profile in [PdfConformance::PdfA2b, PdfConformance::PdfA3b] {
            let archival = render_to_pdf_with_options(&layout, PdfOptions::new(profile)).unwrap();
            let archival = String::from_utf8_lossy(&archival);
            assert!(archival.contains("/F 4"), "{archival}");
        }
    }

    #[test]
    fn ordinary_pdf_api_remains_byte_identical() {
        let pdf = render_to_pdf(&layout_with(vec![PositionedElement::FilledRect {
            rect: Rect {
                x: 12.0,
                y: 24.0,
                width: 36.0,
                height: 48.0,
            },
            color: Color {
                r: 0.25,
                g: 0.5,
                b: 0.75,
                a: 1.0,
            },
        }]));
        let digest = pdf.iter().fold(0xcbf29ce484222325_u64, |hash, byte| {
            hash.wrapping_mul(0x100000001b3) ^ u64::from(*byte)
        });
        assert_eq!(digest, 13_222_022_444_707_008_020);
    }

    #[test]
    fn pdfa_retains_tagged_structure_tree() {
        let root = StructureId::new(1).unwrap();
        let heading = StructureId::new(2).unwrap();
        let list = StructureId::new(3).unwrap();
        let item = StructureId::new(4).unwrap();
        let paragraph = StructureId::new(5).unwrap();
        let table = StructureId::new(6).unwrap();
        let row = StructureId::new(7).unwrap();
        let header = StructureId::new(8).unwrap();
        let cell = StructureId::new(9).unwrap();
        let figure = StructureId::new(10).unwrap();
        let mark = |structure| PositionedElement::MarkedContent {
            structure: Some(structure),
            children: vec![PositionedElement::FilledRect {
                rect: Rect {
                    x: 1.0,
                    y: 1.0,
                    width: 1.0,
                    height: 1.0,
                },
                color: Color::BLACK,
            }],
        };
        let mut layout = layout_with(vec![
            mark(heading),
            mark(paragraph),
            mark(header),
            mark(cell),
            mark(figure),
        ]);
        layout.structure = Some(DocumentStructure {
            root,
            nodes: vec![
                StructureNode {
                    id: root,
                    role: StructureRole::Document,
                    children: vec![heading, list, table, figure],
                    alternate_text: None,
                },
                StructureNode {
                    id: heading,
                    role: StructureRole::Heading(1),
                    children: vec![],
                    alternate_text: None,
                },
                StructureNode {
                    id: list,
                    role: StructureRole::List,
                    children: vec![item],
                    alternate_text: None,
                },
                StructureNode {
                    id: item,
                    role: StructureRole::ListItem,
                    children: vec![paragraph],
                    alternate_text: None,
                },
                StructureNode {
                    id: paragraph,
                    role: StructureRole::Paragraph,
                    children: vec![],
                    alternate_text: None,
                },
                StructureNode {
                    id: table,
                    role: StructureRole::Table,
                    children: vec![row],
                    alternate_text: None,
                },
                StructureNode {
                    id: row,
                    role: StructureRole::TableRow,
                    children: vec![header, cell],
                    alternate_text: None,
                },
                StructureNode {
                    id: header,
                    role: StructureRole::TableHeaderCell,
                    children: vec![],
                    alternate_text: None,
                },
                StructureNode {
                    id: cell,
                    role: StructureRole::TableCell,
                    children: vec![],
                    alternate_text: None,
                },
                StructureNode {
                    id: figure,
                    role: StructureRole::Figure,
                    children: vec![],
                    alternate_text: Some("Quarterly revenue".to_owned()),
                },
            ],
        });
        for profile in [PdfConformance::PdfA2b, PdfConformance::PdfA3b] {
            let pdf = render_to_pdf_with_options(&layout, PdfOptions::new(profile)).unwrap();
            let text = String::from_utf8_lossy(&pdf);
            for marker in [
                "/StructTreeRoot",
                "/S /H1",
                "/S /L",
                "/S /LI",
                "/S /Table",
                "/S /TH",
                "/S /TD",
                "Quarterly revenue",
            ] {
                assert!(text.contains(marker), "missing {marker}: {text}");
            }
            assert!(text.contains("pdfuaid:part"), "{text}");
        }
    }
}
