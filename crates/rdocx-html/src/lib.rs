//! DOCX-to-HTML and DOCX-to-Markdown conversion.
//!
//! Works directly from semantic OXML types — no layout engine needed.

mod emitter;
mod css;
mod markdown;

use std::collections::HashMap;

use rdocx_oxml::document::CT_Document;
use rdocx_oxml::numbering::CT_Numbering;
use rdocx_oxml::styles::CT_Styles;

/// Options for HTML conversion.
#[derive(Debug, Clone)]
pub struct HtmlOptions {
    /// Whether to inline images as base64 data URIs (default: true).
    pub inline_images: bool,
}

impl Default for HtmlOptions {
    fn default() -> Self {
        Self {
            inline_images: true,
        }
    }
}

/// Input for HTML conversion.
pub struct HtmlInput {
    pub document: CT_Document,
    pub styles: CT_Styles,
    pub numbering: Option<CT_Numbering>,
    /// Images keyed by embed/relationship ID.
    pub images: HashMap<String, ImageData>,
    /// Hyperlink URLs keyed by relationship ID.
    pub hyperlink_urls: HashMap<String, String>,
}

/// Image data for HTML embedding.
pub struct ImageData {
    pub data: Vec<u8>,
    pub content_type: String,
}

/// Convert a DOCX document to a complete HTML document string.
pub fn to_html_document(input: &HtmlInput, options: &HtmlOptions) -> String {
    let body = to_html_fragment(input, options);
    let css = css::generate_base_css();
    format!(
        "<!DOCTYPE html>\n<html>\n<head>\n<meta charset=\"UTF-8\">\n<style>\n{css}\n</style>\n</head>\n<body>\n{body}\n</body>\n</html>"
    )
}

/// Convert a DOCX document to an HTML fragment (body content only).
pub fn to_html_fragment(input: &HtmlInput, options: &HtmlOptions) -> String {
    emitter::emit_body(&input.document.body, &input.styles, input.numbering.as_ref(), &input.images, &input.hyperlink_urls, options)
}

/// Convert a DOCX document to Markdown.
pub fn to_markdown(input: &HtmlInput) -> String {
    markdown::emit_markdown(&input.document.body, &input.styles, input.numbering.as_ref(), &input.hyperlink_urls)
}

#[cfg(test)]
mod tests {
    use super::*;
    use rdocx_oxml::document::{BodyContent, CT_Document};
    use rdocx_oxml::styles::CT_Styles;
    use rdocx_oxml::text::CT_P;

    fn simple_input(text: &str) -> HtmlInput {
        let mut doc = CT_Document::new();
        let mut p = CT_P::new();
        p.add_run(text);
        doc.body.add_paragraph(p);

        HtmlInput {
            document: doc,
            styles: CT_Styles::new_default(),
            numbering: None,
            images: HashMap::new(),
            hyperlink_urls: HashMap::new(),
        }
    }

    #[test]
    fn html_document_basic() {
        let input = simple_input("Hello, World!");
        let html = to_html_document(&input, &HtmlOptions::default());
        assert!(html.contains("<!DOCTYPE html>"));
        assert!(html.contains("Hello, World!"));
        assert!(html.contains("<p"));
    }

    #[test]
    fn html_fragment_basic() {
        let input = simple_input("Test paragraph");
        let html = to_html_fragment(&input, &HtmlOptions::default());
        assert!(html.contains("Test paragraph"));
        assert!(html.contains("<p"));
        assert!(!html.contains("<!DOCTYPE"));
    }

    #[test]
    fn markdown_basic() {
        let input = simple_input("Test paragraph");
        let md = to_markdown(&input);
        assert!(md.contains("Test paragraph"));
    }

    #[test]
    fn html_heading() {
        let mut doc = CT_Document::new();
        let mut p = CT_P::new();
        p.add_run("Chapter 1");
        p.properties = Some(rdocx_oxml::properties::CT_PPr {
            style_id: Some("Heading1".to_string()),
            ..Default::default()
        });
        doc.body.add_paragraph(p);

        let input = HtmlInput {
            document: doc,
            styles: CT_Styles::new_default(),
            numbering: None,
            images: HashMap::new(),
            hyperlink_urls: HashMap::new(),
        };

        let html = to_html_fragment(&input, &HtmlOptions::default());
        assert!(html.contains("<h1"));
        assert!(html.contains("Chapter 1"));
    }

    #[test]
    fn html_table() {
        let mut doc = CT_Document::new();
        let mut tbl = rdocx_oxml::table::CT_Tbl::new();
        let mut row = rdocx_oxml::table::CT_Row::new();
        let mut cell = rdocx_oxml::table::CT_Tc::new();
        let mut p = CT_P::new();
        p.add_run("Cell text");
        cell.content = vec![rdocx_oxml::table::CellContent::Paragraph(p)];
        row.cells.push(cell);
        tbl.rows.push(row);
        doc.body.content.push(BodyContent::Table(tbl));

        let input = HtmlInput {
            document: doc,
            styles: CT_Styles::new_default(),
            numbering: None,
            images: HashMap::new(),
            hyperlink_urls: HashMap::new(),
        };

        let html = to_html_fragment(&input, &HtmlOptions::default());
        assert!(html.contains("<table"));
        assert!(html.contains("<td"));
        assert!(html.contains("Cell text"));
    }

    #[test]
    fn markdown_heading() {
        let mut doc = CT_Document::new();
        let mut p = CT_P::new();
        p.add_run("Title");
        p.properties = Some(rdocx_oxml::properties::CT_PPr {
            style_id: Some("Heading1".to_string()),
            ..Default::default()
        });
        doc.body.add_paragraph(p);

        let input = HtmlInput {
            document: doc,
            styles: CT_Styles::new_default(),
            numbering: None,
            images: HashMap::new(),
            hyperlink_urls: HashMap::new(),
        };

        let md = to_markdown(&input);
        assert!(md.contains("# Title"));
    }

    #[test]
    fn html_bold_italic() {
        let mut doc = CT_Document::new();
        let mut p = CT_P::new();
        let mut r = rdocx_oxml::text::CT_R::new("bold text");
        r.properties = Some(rdocx_oxml::properties::CT_RPr {
            bold: Some(true),
            italic: Some(true),
            ..Default::default()
        });
        p.runs.push(r);
        doc.body.add_paragraph(p);

        let input = HtmlInput {
            document: doc,
            styles: CT_Styles::new_default(),
            numbering: None,
            images: HashMap::new(),
            hyperlink_urls: HashMap::new(),
        };

        let html = to_html_fragment(&input, &HtmlOptions::default());
        assert!(html.contains("<strong"));
        assert!(html.contains("<em"));
        assert!(html.contains("bold text"));
    }
}
