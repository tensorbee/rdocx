//! WebAssembly bindings for rdocx.
//!
//! Provides JavaScript-friendly API for creating, opening, and converting
//! DOCX documents entirely in the browser or Node.js — no server needed.

use wasm_bindgen::prelude::*;

/// A Word document (.docx) that can be created, modified, and exported.
#[wasm_bindgen]
pub struct WasmDocument {
    inner: rdocx::Document,
}

impl Default for WasmDocument {
    fn default() -> Self {
        Self::new()
    }
}

#[wasm_bindgen]
impl WasmDocument {
    /// Create a new, empty document.
    #[wasm_bindgen(constructor)]
    pub fn new() -> Self {
        Self {
            inner: rdocx::Document::new(),
        }
    }

    /// Open a document from DOCX bytes.
    #[wasm_bindgen(js_name = "fromBytes")]
    pub fn from_bytes(data: &[u8]) -> Result<WasmDocument, JsValue> {
        let inner = rdocx::Document::from_bytes(data).map_err(to_js_error)?;
        Ok(Self { inner })
    }

    /// Add a paragraph with the given text.
    #[wasm_bindgen(js_name = "addParagraph")]
    pub fn add_paragraph(&mut self, text: &str) {
        self.inner.add_paragraph(text);
    }

    /// Add a heading paragraph (level 1-6).
    #[wasm_bindgen(js_name = "addHeading")]
    pub fn add_heading(&mut self, text: &str, level: u32) {
        let level = level.clamp(1, 6);
        self.inner
            .add_paragraph(text)
            .set_style(&format!("Heading{level}"));
    }

    /// Add a paragraph with bold text.
    #[wasm_bindgen(js_name = "addBoldParagraph")]
    pub fn add_bold_paragraph(&mut self, text: &str) {
        self.inner.add_paragraph("").add_run(text).set_bold(true);
    }

    /// Add a simple table with the given number of rows and columns.
    #[wasm_bindgen(js_name = "addTable")]
    pub fn add_table(&mut self, rows: u32, cols: u32) {
        self.inner.add_table(rows as usize, cols as usize);
    }

    /// Get the text content of the entire document.
    #[wasm_bindgen(js_name = "getText")]
    pub fn get_text(&self) -> String {
        self.inner.text()
    }

    /// Get the number of paragraphs in the document.
    #[wasm_bindgen(js_name = "paragraphCount")]
    pub fn paragraph_count(&self) -> u32 {
        self.inner.paragraph_count() as u32
    }

    /// Export as DOCX bytes.
    #[wasm_bindgen(js_name = "toDocxBytes")]
    pub fn to_docx_bytes(&mut self) -> Result<Vec<u8>, JsValue> {
        self.inner.to_bytes().map_err(to_js_error)
    }

    /// Render the document as PDF bytes.
    #[wasm_bindgen(js_name = "toPdf")]
    pub fn to_pdf(&self) -> Result<Vec<u8>, JsValue> {
        self.inner.to_pdf().map_err(to_js_error)
    }

    /// Convert to a complete HTML document string.
    #[wasm_bindgen(js_name = "toHtml")]
    pub fn to_html(&self) -> String {
        self.inner.to_html()
    }

    /// Convert to an HTML fragment (body content only).
    #[wasm_bindgen(js_name = "toHtmlFragment")]
    pub fn to_html_fragment(&self) -> String {
        self.inner.to_html_fragment()
    }

    /// Convert to Markdown.
    #[wasm_bindgen(js_name = "toMarkdown")]
    pub fn to_markdown(&self) -> String {
        self.inner.to_markdown()
    }

    /// Replace all occurrences of a placeholder with a value.
    #[wasm_bindgen(js_name = "replacePlaceholder")]
    pub fn replace_placeholder(&mut self, placeholder: &str, value: &str) -> u32 {
        self.inner.replace_text(placeholder, value) as u32
    }
}

fn to_js_error(error: rdocx::Error) -> JsValue {
    JsValue::from_str(&error.to_string())
}

#[cfg(test)]
mod tests {
    use super::*;
    use oxml_opc::OpcPackage;
    use std::collections::BTreeSet;

    const DOCUMENT_CONTENT_TYPE: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
    const STYLES_CONTENT_TYPE: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
    const PNG_1_BY_1: &[u8] = &[
        0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44,
        0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00, 0x90,
        0x77, 0x53, 0xde, 0x00, 0x00, 0x00, 0x0c, 0x49, 0x44, 0x41, 0x54, 0x08, 0xd7, 0x63, 0xf8,
        0xcf, 0xc0, 0x00, 0x00, 0x00, 0x02, 0x00, 0x01, 0xe2, 0x21, 0xbc, 0x33, 0x00, 0x00, 0x00,
        0x00, 0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
    ];

    fn package_fixture() -> Vec<u8> {
        let mut document = rdocx::Document::new();
        document.add_paragraph("Before image");
        document.add_picture(
            PNG_1_BY_1,
            "pixel.png",
            rdocx::Length::inches(1.0),
            rdocx::Length::inches(1.0),
        );
        document.set_header("Preserved header");
        document.add_numbered_list_item("Numbered item", 0);
        document.to_bytes().expect("fixture should serialize")
    }

    fn assert_complete_package_round_trip(source: &[u8], output: &[u8]) {
        let source_package = OpcPackage::from_reader(std::io::Cursor::new(source)).unwrap();
        let output_package = OpcPackage::from_reader(std::io::Cursor::new(output)).unwrap();

        assert_eq!(
            source_package.parts.keys().collect::<BTreeSet<_>>(),
            output_package.parts.keys().collect::<BTreeSet<_>>(),
            "the complete part inventory must survive"
        );
        assert_eq!(
            source_package.content_types.defaults,
            output_package.content_types.defaults
        );
        assert_eq!(
            source_package.content_types.overrides,
            output_package.content_types.overrides
        );
        assert_eq!(
            source_package.package_rels.items,
            output_package.package_rels.items
        );
        assert_eq!(
            source_package.part_rels.keys().collect::<BTreeSet<_>>(),
            output_package.part_rels.keys().collect::<BTreeSet<_>>()
        );
        for (part_name, relationships) in &source_package.part_rels {
            assert_eq!(
                relationships.items, output_package.part_rels[part_name].items,
                "relationships for {part_name} must survive"
            );
        }

        for part_name in source_package.parts.keys().filter(|name| {
            !matches!(
                name.as_str(),
                "/word/document.xml"
                    | "/word/styles.xml"
                    | "/word/numbering.xml"
                    | "/docProps/core.xml"
                    | "/word/footnotes.xml"
            )
        }) {
            assert_eq!(
                source_package.parts[part_name], output_package.parts[part_name],
                "opaque part {part_name} must survive byte-identically"
            );
        }

        let reopened = rdocx::Document::from_bytes(output).expect("output should reopen");
        assert_eq!(reopened.header_text().as_deref(), Some("Preserved header"));
        assert_eq!(reopened.images().len(), 1);
        let numbered = reopened
            .paragraphs()
            .into_iter()
            .find(|paragraph| paragraph.text() == "Numbered item")
            .expect("numbered paragraph should survive");
        assert!(numbered.numbering().is_some());
    }

    #[test]
    fn document_with_images_headers_and_numbering_round_trips_every_part_intact() {
        let source = package_fixture();
        let mut document = WasmDocument::from_bytes(&source).expect("fixture should open");
        let output = document
            .to_docx_bytes()
            .expect("WASM document should serialize");

        assert_complete_package_round_trip(&source, &output);
    }

    #[cfg(target_arch = "wasm32")]
    #[wasm_bindgen_test::wasm_bindgen_test]
    fn wasm_round_trip_preserves_the_complete_package_in_node() {
        use js_sys::{Function, Reflect, Uint8Array};
        use wasm_bindgen::JsCast;

        let source = package_fixture();
        let probe = JsValue::from(WasmDocument::new());
        let constructor = Reflect::get(&probe, &JsValue::from_str("constructor"))
            .expect("generated class constructor should be visible");
        let from_bytes: Function = Reflect::get(&constructor, &JsValue::from_str("fromBytes"))
            .expect("generated fromBytes should be visible")
            .dyn_into()
            .expect("fromBytes should be callable");
        let input = Uint8Array::from(source.as_slice());
        let document = from_bytes
            .call1(&constructor, &input)
            .expect("fromBytes should accept Uint8Array");
        let to_docx_bytes: Function = Reflect::get(&document, &JsValue::from_str("toDocxBytes"))
            .expect("generated toDocxBytes should be visible")
            .dyn_into()
            .expect("toDocxBytes should be callable");
        let output = to_docx_bytes
            .call0(&document)
            .expect("toDocxBytes should return bytes");
        assert!(output.is_instance_of::<Uint8Array>());
        let output = Uint8Array::new(&output).to_vec();

        assert_complete_package_round_trip(&source, &output);
    }

    #[cfg(target_arch = "wasm32")]
    #[wasm_bindgen_test::wasm_bindgen_test]
    fn to_pdf_in_node_returns_a_complete_pdf_with_an_embedded_bundled_font() {
        use js_sys::{Function, Reflect, Uint8Array};
        use wasm_bindgen::JsCast;

        let document = JsValue::from(WasmDocument::new());
        let add_paragraph: Function = Reflect::get(&document, &JsValue::from_str("addParagraph"))
            .expect("generated addParagraph should be visible")
            .dyn_into()
            .expect("addParagraph should be callable");
        add_paragraph
            .call1(
                &document,
                &JsValue::from_str("Bundled font PDF from WebAssembly"),
            )
            .expect("addParagraph should accept text");
        let to_pdf: Function = Reflect::get(&document, &JsValue::from_str("toPdf"))
            .expect("generated toPdf should be visible")
            .dyn_into()
            .expect("toPdf should be callable");
        let pdf = to_pdf.call0(&document).expect("toPdf should return bytes");
        assert!(pdf.is_instance_of::<Uint8Array>());
        let pdf = Uint8Array::new(&pdf).to_vec();
        let pdf_text = String::from_utf8_lossy(&pdf);

        assert!(pdf.starts_with(b"%PDF-"));
        assert!(pdf.ends_with(b"%%EOF"));
        assert!(pdf_text.contains("/Subtype /Type0"));
        assert!(pdf_text.contains("/FontFile2"));
        assert!(pdf_text.contains("/BaseFont /Carlito"));
    }

    #[test]
    fn wasm_new_document_uses_the_shared_word_package_setup() {
        let mut document = WasmDocument::new();
        let bytes = document.to_docx_bytes().unwrap();
        let package = OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();

        assert_eq!(
            package.main_document_part().as_deref(),
            Some("/word/document.xml")
        );
        assert_eq!(
            package.content_types.content_type_for("/word/document.xml"),
            Some(DOCUMENT_CONTENT_TYPE)
        );
        assert_eq!(
            package.content_types.content_type_for("/word/styles.xml"),
            Some(STYLES_CONTENT_TYPE)
        );
        assert!(package.get_part("/word/styles.xml").is_some());
        let styles = package
            .get_part_rels("/word/document.xml")
            .and_then(|rels| rels.get_by_type(oxml_opc::relationship::rel_types::STYLES))
            .expect("WASM document must relate its styles part");
        assert_eq!(styles.target, "styles.xml");
    }

    #[test]
    fn word_native_defaults_and_wasm_isolation_are_manifest_contracts() {
        let workspace_manifest = include_str!("../../../Cargo.toml");
        let layout_manifest = include_str!("../../rdocx-layout/Cargo.toml");
        let facade_manifest = include_str!("../../rdocx/Cargo.toml");
        let wasm_manifest = include_str!("../Cargo.toml");

        for dependency in [
            "oxml-layout = { path = \"crates/oxml-layout\", version = \"0.10.0\", default-features = false }",
            "rdocx = { path = \"crates/rdocx\", version = \"0.12.0\", default-features = false }",
            "rdocx-layout = { path = \"crates/rdocx-layout\", version = \"0.12.0\", default-features = false }",
        ] {
            assert!(
                workspace_manifest.contains(dependency),
                "workspace dependency must stay defaults-off: {dependency}"
            );
        }
        assert!(layout_manifest.contains(
            "[features]\ndefault = [\"system-fonts\"]\nsystem-fonts = [\"oxml-layout/system-fonts\"]"
        ));
        assert!(
            layout_manifest
                .contains("oxml-layout = { workspace = true, default-features = false }")
        );
        assert!(facade_manifest.contains(
            "[features]\ndefault = [\"system-fonts\"]\nsystem-fonts = [\"oxml-layout/system-fonts\", \"rdocx-layout/system-fonts\"]"
        ));
        assert!(
            facade_manifest
                .contains("rdocx-layout = { workspace = true, default-features = false }")
        );
        assert!(wasm_manifest.contains("rdocx = { workspace = true, default-features = false }"));
        assert!(!wasm_manifest.contains("rdocx/system-fonts"));
    }
}
