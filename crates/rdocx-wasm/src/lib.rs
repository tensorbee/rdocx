//! WebAssembly bindings for rdocx.
//!
//! Provides JavaScript-friendly API for creating, opening, and converting
//! DOCX documents entirely in the browser or Node.js — no server needed.

use wasm_bindgen::prelude::*;

use rdocx_opc::OpcPackage;
use rdocx_oxml::document::{BodyContent, CT_Document};
use rdocx_oxml::properties::{CT_PPr, CT_RPr};
use rdocx_oxml::styles::CT_Styles;
use rdocx_oxml::table::{CT_Row, CT_Tbl, CT_Tc, CellContent};
use rdocx_oxml::text::{CT_P, CT_R};

use std::collections::HashMap;

/// A Word document (.docx) that can be created, modified, and exported.
#[wasm_bindgen]
pub struct WasmDocument {
    document: CT_Document,
    styles: CT_Styles,
    /// Raw OPC package bytes (set when opened from existing file).
    #[allow(dead_code)]
    package_bytes: Option<Vec<u8>>,
}

#[wasm_bindgen]
impl WasmDocument {
    /// Create a new, empty document.
    #[wasm_bindgen(constructor)]
    pub fn new() -> Self {
        WasmDocument {
            document: CT_Document::new(),
            styles: CT_Styles::new_default(),
            package_bytes: None,
        }
    }

    /// Open a document from DOCX bytes.
    #[wasm_bindgen(js_name = "fromBytes")]
    pub fn from_bytes(data: &[u8]) -> Result<WasmDocument, JsValue> {
        let cursor = std::io::Cursor::new(data);
        let package =
            OpcPackage::from_reader(cursor).map_err(|e| JsValue::from_str(&e.to_string()))?;

        let doc_part = package
            .main_document_part()
            .ok_or_else(|| JsValue::from_str("No main document part found"))?;

        let xml = package
            .get_part(&doc_part)
            .ok_or_else(|| JsValue::from_str("Document part not found in package"))?;

        let document =
            CT_Document::from_xml(xml).map_err(|e| JsValue::from_str(&e.to_string()))?;

        // Try to load styles
        let styles = package
            .get_part_rels(&doc_part)
            .and_then(|rels| {
                rels.get_by_type(rdocx_opc::relationship::rel_types::STYLES)
                    .map(|r| OpcPackage::resolve_rel_target(&doc_part, &r.target))
            })
            .and_then(|part_name| package.get_part(&part_name))
            .and_then(|xml| CT_Styles::from_xml(xml).ok())
            .unwrap_or_else(CT_Styles::new_default);

        Ok(WasmDocument {
            document,
            styles,
            package_bytes: Some(data.to_vec()),
        })
    }

    /// Add a paragraph with the given text.
    #[wasm_bindgen(js_name = "addParagraph")]
    pub fn add_paragraph(&mut self, text: &str) {
        let mut p = CT_P::new();
        p.add_run(text);
        self.document.body.add_paragraph(p);
    }

    /// Add a heading paragraph (level 1-6).
    #[wasm_bindgen(js_name = "addHeading")]
    pub fn add_heading(&mut self, text: &str, level: u32) {
        let mut p = CT_P::new();
        p.add_run(text);
        let lvl = level.clamp(1, 6);
        p.properties = Some(CT_PPr {
            style_id: Some(format!("Heading{lvl}")),
            ..Default::default()
        });
        self.document.body.add_paragraph(p);
    }

    /// Add a paragraph with bold text.
    #[wasm_bindgen(js_name = "addBoldParagraph")]
    pub fn add_bold_paragraph(&mut self, text: &str) {
        let mut p = CT_P::new();
        let mut r = CT_R::new(text);
        r.properties = Some(CT_RPr {
            bold: Some(true),
            ..Default::default()
        });
        p.runs.push(r);
        self.document.body.add_paragraph(p);
    }

    /// Add a simple table with the given number of rows and columns.
    #[wasm_bindgen(js_name = "addTable")]
    pub fn add_table(&mut self, rows: u32, cols: u32) {
        let mut tbl = CT_Tbl::new();
        for _ in 0..rows {
            let mut row = CT_Row::new();
            for _ in 0..cols {
                let mut cell = CT_Tc::new();
                cell.content = vec![CellContent::Paragraph(CT_P::new())];
                row.cells.push(cell);
            }
            tbl.rows.push(row);
        }
        self.document
            .body
            .content
            .push(BodyContent::Table(tbl));
    }

    /// Get the text content of the entire document.
    #[wasm_bindgen(js_name = "getText")]
    pub fn get_text(&self) -> String {
        let mut result = String::new();
        for content in &self.document.body.content {
            match content {
                BodyContent::Paragraph(p) => {
                    result.push_str(&p.text());
                    result.push('\n');
                }
                BodyContent::Table(tbl) => {
                    for row in &tbl.rows {
                        for cell in &row.cells {
                            for cc in &cell.content {
                                if let CellContent::Paragraph(p) = cc {
                                    result.push_str(&p.text());
                                    result.push('\t');
                                }
                            }
                        }
                        result.push('\n');
                    }
                }
                BodyContent::RawXml(_) => {}
            }
        }
        result
    }

    /// Get the number of paragraphs in the document.
    #[wasm_bindgen(js_name = "paragraphCount")]
    pub fn paragraph_count(&self) -> u32 {
        self.document
            .body
            .content
            .iter()
            .filter(|c| matches!(c, BodyContent::Paragraph(_)))
            .count() as u32
    }

    /// Export as DOCX bytes.
    #[wasm_bindgen(js_name = "toDocxBytes")]
    pub fn to_docx_bytes(&self) -> Result<Vec<u8>, JsValue> {
        // Build a minimal OPC package
        let mut package = OpcPackage::new_docx();

        // Add styles relationship
        package
            .get_or_create_part_rels("/word/document.xml")
            .add(rdocx_opc::relationship::rel_types::STYLES, "styles.xml");

        // Serialize document
        let doc_xml = self
            .document
            .to_xml()
            .map_err(|e| JsValue::from_str(&e.to_string()))?;
        package.set_part("/word/document.xml", doc_xml);

        // Serialize styles
        let styles_xml = self
            .styles
            .to_xml()
            .map_err(|e| JsValue::from_str(&e.to_string()))?;
        package.set_part("/word/styles.xml", styles_xml);

        let mut output = Vec::new();
        package
            .write_to(&mut std::io::Cursor::new(&mut output))
            .map_err(|e| JsValue::from_str(&e.to_string()))?;

        Ok(output)
    }

    /// Convert to a complete HTML document string.
    #[wasm_bindgen(js_name = "toHtml")]
    pub fn to_html(&self) -> String {
        let input = self.build_html_input();
        rdocx_html::to_html_document(&input, &rdocx_html::HtmlOptions::default())
    }

    /// Convert to an HTML fragment (body content only).
    #[wasm_bindgen(js_name = "toHtmlFragment")]
    pub fn to_html_fragment(&self) -> String {
        let input = self.build_html_input();
        rdocx_html::to_html_fragment(&input, &rdocx_html::HtmlOptions::default())
    }

    /// Convert to Markdown.
    #[wasm_bindgen(js_name = "toMarkdown")]
    pub fn to_markdown(&self) -> String {
        let input = self.build_html_input();
        rdocx_html::to_markdown(&input)
    }

    /// Replace all occurrences of a placeholder with a value.
    #[wasm_bindgen(js_name = "replacePlaceholder")]
    pub fn replace_placeholder(&mut self, placeholder: &str, value: &str) -> u32 {
        use rdocx_oxml::placeholder;
        let mut count = 0u32;
        for content in &mut self.document.body.content {
            match content {
                BodyContent::Paragraph(p) => {
                    count += placeholder::replace_in_paragraph(p, placeholder, value) as u32;
                }
                BodyContent::Table(tbl) => {
                    count += placeholder::replace_in_table(tbl, placeholder, value) as u32;
                }
                BodyContent::RawXml(_) => {}
            }
        }
        count
    }
}

impl WasmDocument {
    fn build_html_input(&self) -> rdocx_html::HtmlInput {
        rdocx_html::HtmlInput {
            document: self.document.clone(),
            styles: self.styles.clone(),
            numbering: None,
            images: HashMap::new(),
            hyperlink_urls: HashMap::new(),
        }
    }
}
