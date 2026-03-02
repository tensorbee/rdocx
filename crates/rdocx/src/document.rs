//! The main Document type — entry point for the rdocx API.

use std::path::Path;

use rdocx_opc::OpcPackage;
use rdocx_opc::relationship::rel_types;
use rdocx_oxml::document::{BodyContent, CT_Columns, CT_Document, CT_SectPr};
use rdocx_oxml::drawing::{CT_Anchor, CT_Drawing, CT_Inline};
use rdocx_oxml::header_footer::{CT_HdrFtr, HdrFtrRef, HdrFtrType};
use rdocx_oxml::numbering::CT_Numbering;
use rdocx_oxml::properties::{CT_PPr, CT_RPr};
use rdocx_oxml::shared::{ST_PageOrientation, ST_SectionType};
use rdocx_oxml::styles::CT_Styles;
use rdocx_oxml::table::CT_Tbl;
use rdocx_oxml::text::{CT_P, CT_R, RunContent};

use rdocx_oxml::core_properties::CoreProperties;

use crate::Length;
use crate::error::{Error, Result};
use crate::paragraph::{Paragraph, ParagraphRef};
use crate::style::{self, Style, StyleBuilder};
use crate::table::{Table, TableRef};

/// A Word document (.docx file).
///
/// This is the main entry point for reading, creating, and modifying
/// DOCX documents.
pub struct Document {
    package: OpcPackage,
    document: CT_Document,
    styles: CT_Styles,
    numbering: Option<CT_Numbering>,
    core_properties: Option<CoreProperties>,
    /// Part name for the main document
    doc_part_name: String,
    /// Cached count of image media parts (avoids rescanning parts on each embed).
    image_counter: usize,
}

impl Document {
    /// Create a new, empty document with default page setup and styles.
    pub fn new() -> Self {
        let mut package = OpcPackage::new_docx();
        let document = CT_Document::new();
        let styles = CT_Styles::new_default();

        // Set up styles relationship
        package
            .get_or_create_part_rels("/word/document.xml")
            .add(rel_types::STYLES, "styles.xml");

        Document {
            package,
            document,
            styles,
            numbering: None,
            core_properties: None,
            doc_part_name: "/word/document.xml".to_string(),
            image_counter: 0,
        }
    }

    /// Open a document from a file path.
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        let package = OpcPackage::open(path)?;
        Self::from_package(package)
    }

    /// Open a document from bytes.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        let cursor = std::io::Cursor::new(bytes);
        let package = OpcPackage::from_reader(cursor)?;
        Self::from_package(package)
    }

    fn from_package(package: OpcPackage) -> Result<Self> {
        let doc_part_name = package.main_document_part().ok_or(Error::NoDocumentPart)?;

        let doc_xml = package
            .get_part(&doc_part_name)
            .ok_or(Error::NoDocumentPart)?;
        let document = CT_Document::from_xml(doc_xml)?;

        // Try to load styles
        let styles = if let Some(rels) = package.get_part_rels(&doc_part_name) {
            if let Some(styles_rel) = rels.get_by_type(rel_types::STYLES) {
                let styles_part =
                    OpcPackage::resolve_rel_target(&doc_part_name, &styles_rel.target);
                if let Some(styles_xml) = package.get_part(&styles_part) {
                    CT_Styles::from_xml(styles_xml)?
                } else {
                    CT_Styles::new_default()
                }
            } else {
                CT_Styles::new_default()
            }
        } else {
            CT_Styles::new_default()
        };

        // Try to load numbering definitions
        let numbering = if let Some(rels) = package.get_part_rels(&doc_part_name) {
            if let Some(num_rel) = rels.get_by_type(rel_types::NUMBERING) {
                let num_part = OpcPackage::resolve_rel_target(&doc_part_name, &num_rel.target);
                if let Some(num_xml) = package.get_part(&num_part) {
                    Some(CT_Numbering::from_xml(num_xml)?)
                } else {
                    None
                }
            } else {
                None
            }
        } else {
            None
        };

        // Try to load core properties from docProps/core.xml
        let core_properties = package
            .get_part("/docProps/core.xml")
            .and_then(|xml| CoreProperties::from_xml(xml).ok());

        let image_counter = package
            .parts
            .keys()
            .filter(|k| k.starts_with("/word/media/image"))
            .count();

        Ok(Document {
            package,
            document,
            styles,
            numbering,
            core_properties,
            doc_part_name,
            image_counter,
        })
    }

    /// Save the document to a file path.
    pub fn save<P: AsRef<Path>>(&mut self, path: P) -> Result<()> {
        self.flush_to_package()?;
        self.package.save(path)?;
        Ok(())
    }

    /// Save the document to a byte vector.
    pub fn to_bytes(&mut self) -> Result<Vec<u8>> {
        self.flush_to_package()?;
        let mut buf = std::io::Cursor::new(Vec::new());
        self.package.write_to(&mut buf)?;
        Ok(buf.into_inner())
    }

    /// Write the in-memory document/styles back into the OPC package parts.
    fn flush_to_package(&mut self) -> Result<()> {
        // Serialize document.xml
        let doc_xml = self.document.to_xml()?;
        self.package.set_part(&self.doc_part_name, doc_xml);

        // Serialize styles.xml
        let styles_xml = self.styles.to_xml()?;
        self.package.set_part("/word/styles.xml", styles_xml);

        // Serialize numbering.xml if we have numbering definitions
        if let Some(ref numbering) = self.numbering {
            let numbering_xml = numbering.to_xml()?;
            self.package.set_part("/word/numbering.xml", numbering_xml);
        }

        // Serialize docProps/core.xml if we have metadata
        if let Some(ref props) = self.core_properties {
            let core_xml = props.to_xml()?;
            self.package.set_part("/docProps/core.xml", core_xml);
            self.package.content_types.add_override(
                "/docProps/core.xml",
                "application/vnd.openxmlformats-package.core-properties+xml",
            );
        }

        Ok(())
    }

    // ---- Paragraph access ----

    /// Get immutable references to all paragraphs.
    pub fn paragraphs(&self) -> Vec<ParagraphRef<'_>> {
        self.document
            .body
            .paragraphs()
            .map(|p| ParagraphRef { inner: p })
            .collect()
    }

    /// Add a paragraph with the given text and return a mutable reference.
    pub fn add_paragraph(&mut self, text: &str) -> Paragraph<'_> {
        let mut p = CT_P::new();
        if !text.is_empty() {
            p.add_run(text);
        }
        self.document.body.content.push(BodyContent::Paragraph(p));
        match self.document.body.content.last_mut().unwrap() {
            BodyContent::Paragraph(p) => Paragraph { inner: p },
            _ => unreachable!(),
        }
    }

    /// Get the number of paragraphs.
    pub fn paragraph_count(&self) -> usize {
        self.document.body.paragraphs().count()
    }

    /// Get a mutable reference to a paragraph by index (among paragraphs only).
    pub fn paragraph_mut(&mut self, index: usize) -> Option<Paragraph<'_>> {
        self.document
            .body
            .paragraphs_mut()
            .nth(index)
            .map(|p| Paragraph { inner: p })
    }

    // ---- Table access ----

    /// Get immutable references to all tables.
    pub fn tables(&self) -> Vec<TableRef<'_>> {
        self.document
            .body
            .tables()
            .map(|t| TableRef { inner: t })
            .collect()
    }

    /// Add a table with the specified number of rows and columns.
    /// Returns a mutable reference for further configuration.
    pub fn add_table(&mut self, rows: usize, cols: usize) -> Table<'_> {
        use rdocx_oxml::table::{CT_Row, CT_TblGrid, CT_TblGridCol, CT_TblPr, CT_TblWidth, CT_Tc};
        use rdocx_oxml::units::Twips;

        // Default column width: divide 9360tw (6.5" printable at 1" margins) evenly
        let col_width = Twips(9360 / cols as i32);

        let grid = CT_TblGrid {
            columns: (0..cols)
                .map(|_| CT_TblGridCol { width: col_width })
                .collect(),
        };

        let mut tbl = CT_Tbl::new();
        tbl.properties = Some(CT_TblPr {
            width: Some(CT_TblWidth::dxa(col_width.0 * cols as i32)),
            ..Default::default()
        });
        tbl.grid = Some(grid);

        for _ in 0..rows {
            let mut row = CT_Row::new();
            for _ in 0..cols {
                row.cells.push(CT_Tc::new());
            }
            tbl.rows.push(row);
        }

        self.document.body.content.push(BodyContent::Table(tbl));
        match self.document.body.content.last_mut().unwrap() {
            BodyContent::Table(t) => Table { inner: t },
            _ => unreachable!(),
        }
    }

    /// Get the number of tables.
    pub fn table_count(&self) -> usize {
        self.document.body.tables().count()
    }

    // ---- Content insertion ----

    /// Get the number of body content elements (paragraphs + tables).
    pub fn content_count(&self) -> usize {
        self.document.body.content_count()
    }

    /// Insert a paragraph at the given body index.
    ///
    /// Returns a mutable `Paragraph` for further configuration.
    /// Panics if `index > content_count()`.
    pub fn insert_paragraph(&mut self, index: usize, text: &str) -> Paragraph<'_> {
        let mut p = CT_P::new();
        if !text.is_empty() {
            p.add_run(text);
        }
        self.document.body.insert_paragraph(index, p);
        match &mut self.document.body.content[index] {
            BodyContent::Paragraph(p) => Paragraph { inner: p },
            _ => unreachable!(),
        }
    }

    /// Insert a table at the given body index.
    ///
    /// Returns a mutable `Table` for further configuration.
    /// Panics if `index > content_count()`.
    pub fn insert_table(&mut self, index: usize, rows: usize, cols: usize) -> Table<'_> {
        use rdocx_oxml::table::{CT_Row, CT_TblGrid, CT_TblGridCol, CT_TblPr, CT_TblWidth, CT_Tc};
        use rdocx_oxml::units::Twips;

        let col_width = Twips(9360 / cols as i32);
        let grid = CT_TblGrid {
            columns: (0..cols)
                .map(|_| CT_TblGridCol { width: col_width })
                .collect(),
        };

        let mut tbl = CT_Tbl::new();
        tbl.properties = Some(CT_TblPr {
            width: Some(CT_TblWidth::dxa(col_width.0 * cols as i32)),
            ..Default::default()
        });
        tbl.grid = Some(grid);

        for _ in 0..rows {
            let mut row = CT_Row::new();
            for _ in 0..cols {
                row.cells.push(CT_Tc::new());
            }
            tbl.rows.push(row);
        }

        self.document.body.insert_table(index, tbl);
        match &mut self.document.body.content[index] {
            BodyContent::Table(t) => Table { inner: t },
            _ => unreachable!(),
        }
    }

    /// Find the body content index of the first paragraph containing the given text.
    pub fn find_content_index(&self, text: &str) -> Option<usize> {
        self.document.body.find_paragraph_index(text)
    }

    /// Remove the content at the given body index.
    ///
    /// Returns `true` if an element was removed, `false` if the index was out of bounds.
    pub fn remove_content(&mut self, index: usize) -> bool {
        self.document.body.remove(index).is_some()
    }

    // ---- Image support ----

    /// Add an inline image to the document.
    ///
    /// Embeds the image data (PNG, JPEG, etc.) into the package and adds a
    /// paragraph containing the image. Returns a mutable reference to the
    /// paragraph for further configuration.
    ///
    /// `width` and `height` specify the display size.
    pub fn add_picture(
        &mut self,
        image_data: &[u8],
        image_filename: &str,
        width: Length,
        height: Length,
    ) -> Paragraph<'_> {
        let rel_id = self.embed_image(image_data, image_filename);

        let inline = CT_Inline::new(&rel_id, width.to_emu(), height.to_emu());

        let drawing = CT_Drawing::inline(inline);
        let run = CT_R {
            properties: None,
            content: vec![RunContent::Drawing(drawing)],
            extra_xml: Vec::new(),
        };

        let mut p = CT_P::new();
        p.runs.push(run);
        self.document.body.content.push(BodyContent::Paragraph(p));
        match self.document.body.content.last_mut().unwrap() {
            BodyContent::Paragraph(p) => Paragraph { inner: p },
            _ => unreachable!(),
        }
    }

    /// Add a full-page background image behind text.
    ///
    /// The image is placed at position (0,0) relative to the page with
    /// dimensions matching the page size from section properties.
    /// It is inserted at the beginning of the document body so it renders
    /// behind all other content.
    pub fn add_background_image(
        &mut self,
        image_data: &[u8],
        image_filename: &str,
    ) -> Paragraph<'_> {
        let rel_id = self.embed_image(image_data, image_filename);

        // Get page dimensions from section properties (default US Letter)
        let sect = self
            .document
            .body
            .sect_pr
            .as_ref()
            .cloned()
            .unwrap_or_else(CT_SectPr::default_letter);
        let page_width_emu = sect
            .page_width
            .unwrap_or(rdocx_oxml::units::Twips(12240))
            .to_emu()
            .0;
        let page_height_emu = sect
            .page_height
            .unwrap_or(rdocx_oxml::units::Twips(15840))
            .to_emu()
            .0;

        let anchor = CT_Anchor::background(&rel_id, page_width_emu, page_height_emu);
        let drawing = CT_Drawing::anchor(anchor);
        let run = CT_R {
            properties: None,
            content: vec![RunContent::Drawing(drawing)],
            extra_xml: Vec::new(),
        };

        let mut p = CT_P::new();
        p.runs.push(run);
        self.document.body.insert_paragraph(0, p);
        match &mut self.document.body.content[0] {
            BodyContent::Paragraph(p) => Paragraph { inner: p },
            _ => unreachable!(),
        }
    }

    /// Add an anchored (floating) image to the document.
    ///
    /// If `behind_text` is true, the image renders behind text content.
    /// The image is inserted at the beginning of the document body.
    pub fn add_anchored_image(
        &mut self,
        image_data: &[u8],
        image_filename: &str,
        width: Length,
        height: Length,
        behind_text: bool,
    ) -> Paragraph<'_> {
        let rel_id = self.embed_image(image_data, image_filename);

        let mut anchor = CT_Anchor::background(&rel_id, width.to_emu(), height.to_emu());
        anchor.behind_doc = behind_text;

        let drawing = CT_Drawing::anchor(anchor);
        let run = CT_R {
            properties: None,
            content: vec![RunContent::Drawing(drawing)],
            extra_xml: Vec::new(),
        };

        let mut p = CT_P::new();
        p.runs.push(run);
        self.document.body.insert_paragraph(0, p);
        match &mut self.document.body.content[0] {
            BodyContent::Paragraph(p) => Paragraph { inner: p },
            _ => unreachable!(),
        }
    }

    /// Return the next unique image number and bump the counter.
    fn next_image_number(&mut self) -> usize {
        self.image_counter += 1;
        self.image_counter
    }

    /// Embed an image into the OPC package and return the relationship ID.
    ///
    /// Public so callers can pre-embed an image and then pass the returned
    /// `rel_id` to [`Cell::add_picture`] for inline cell images.
    pub fn embed_image(&mut self, image_data: &[u8], filename: &str) -> String {
        use rdocx_opc::relationship::rel_types;

        // Determine content type from extension
        let ext = filename.rsplit('.').next().unwrap_or("png").to_lowercase();
        let content_type = match ext.as_str() {
            "png" => "image/png",
            "jpg" | "jpeg" => "image/jpeg",
            "gif" => "image/gif",
            "bmp" => "image/bmp",
            "tiff" | "tif" => "image/tiff",
            "svg" => "image/svg+xml",
            _ => "image/png",
        };

        // Generate a unique part name using cached counter
        let image_num = self.next_image_number();
        let part_name = format!("/word/media/image{image_num}.{ext}");

        // Store the image data
        self.package.set_part(&part_name, image_data.to_vec());

        // Add content type override
        self.package.content_types.add_default(&ext, content_type);

        // Add relationship
        let rel_target = format!("media/image{image_num}.{ext}");
        let rels = self.package.get_or_create_part_rels(&self.doc_part_name);
        rels.add(rel_types::IMAGE, &rel_target)
    }

    // ---- Header/Footer ----

    /// Set the default header text.
    ///
    /// Creates a header part with the given text and references it from
    /// the section properties.
    pub fn set_header(&mut self, text: &str) {
        self.set_header_footer_part(text, true, HdrFtrType::Default);
    }

    /// Set the default footer text.
    pub fn set_footer(&mut self, text: &str) {
        self.set_header_footer_part(text, false, HdrFtrType::Default);
    }

    /// Set the first-page header text.
    pub fn set_first_page_header(&mut self, text: &str) {
        self.set_different_first_page(true);
        self.set_header_footer_part(text, true, HdrFtrType::First);
    }

    /// Set the first-page footer text.
    pub fn set_first_page_footer(&mut self, text: &str) {
        self.set_different_first_page(true);
        self.set_header_footer_part(text, false, HdrFtrType::First);
    }

    fn set_header_footer_part(&mut self, text: &str, is_header: bool, hdr_type: HdrFtrType) {
        use rdocx_opc::relationship::rel_types;

        let mut hdr_ftr = CT_HdrFtr::new();
        let mut p = CT_P::new();
        if !text.is_empty() {
            p.add_run(text);
        }
        hdr_ftr.paragraphs.push(p);

        // Determine part name based on type
        let type_suffix = match hdr_type {
            HdrFtrType::Default => "",
            HdrFtrType::First => "First",
            HdrFtrType::Even => "Even",
        };
        let (part_name, rel_type, content_type) = if is_header {
            (
                format!("/word/header{type_suffix}1.xml"),
                rel_types::HEADER,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
            )
        } else {
            (
                format!("/word/footer{type_suffix}1.xml"),
                rel_types::FOOTER,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
            )
        };

        // Serialize the header/footer
        let xml = if is_header {
            hdr_ftr
                .to_xml_header()
                .expect("header serialization failed")
        } else {
            hdr_ftr
                .to_xml_footer()
                .expect("footer serialization failed")
        };

        self.package.set_part(&part_name, xml);
        self.package
            .content_types
            .add_override(&part_name, content_type);

        // Add relationship
        let rel_target = part_name.trim_start_matches("/word/");
        let rels = self.package.get_or_create_part_rels(&self.doc_part_name);
        let rel_id = rels.add(rel_type, rel_target);

        // Add reference in section properties
        let sect = self.section_properties_mut();
        let refs = if is_header {
            &mut sect.header_refs
        } else {
            &mut sect.footer_refs
        };

        // Remove existing ref of same type
        refs.retain(|r| r.hdr_ftr_type != hdr_type);
        refs.push(HdrFtrRef {
            hdr_ftr_type: hdr_type,
            rel_id,
        });
    }

    /// Get the default header text, if set.
    pub fn header_text(&self) -> Option<String> {
        self.get_header_footer_text(true, HdrFtrType::Default)
    }

    /// Get the default footer text, if set.
    pub fn footer_text(&self) -> Option<String> {
        self.get_header_footer_text(false, HdrFtrType::Default)
    }

    /// Set the default header to an inline image.
    ///
    /// Creates a header part with an image paragraph. The image is embedded
    /// in the header part's relationships.
    pub fn set_header_image(
        &mut self,
        image_data: &[u8],
        image_filename: &str,
        width: Length,
        height: Length,
    ) {
        self.set_header_footer_image_part(
            image_data,
            image_filename,
            width,
            height,
            true,
            HdrFtrType::Default,
        );
    }

    /// Set the default footer to an inline image.
    pub fn set_footer_image(
        &mut self,
        image_data: &[u8],
        image_filename: &str,
        width: Length,
        height: Length,
    ) {
        self.set_header_footer_image_part(
            image_data,
            image_filename,
            width,
            height,
            false,
            HdrFtrType::Default,
        );
    }

    /// Set a header from raw XML bytes with associated images.
    ///
    /// This is useful for copying complex headers from template documents
    /// that contain grouped shapes, VML, or other elements not easily
    /// recreated through the high-level API.
    ///
    /// Each entry in `images` is `(rel_id, image_data, image_filename)`:
    /// - `rel_id`: the relationship ID referenced in the header XML (e.g. "rId1")
    /// - `image_data`: the raw image bytes
    /// - `image_filename`: used to derive the part name and content type (e.g. "image5.png")
    pub fn set_raw_header_with_images(
        &mut self,
        header_xml: Vec<u8>,
        images: &[(&str, &[u8], &str)],
        hdr_type: HdrFtrType,
    ) {
        self.set_raw_hdr_ftr_with_images(header_xml, images, true, hdr_type);
    }

    /// Set a footer from raw XML bytes with associated images.
    pub fn set_raw_footer_with_images(
        &mut self,
        footer_xml: Vec<u8>,
        images: &[(&str, &[u8], &str)],
        hdr_type: HdrFtrType,
    ) {
        self.set_raw_hdr_ftr_with_images(footer_xml, images, false, hdr_type);
    }

    fn set_raw_hdr_ftr_with_images(
        &mut self,
        xml: Vec<u8>,
        images: &[(&str, &[u8], &str)],
        is_header: bool,
        hdr_type: HdrFtrType,
    ) {
        use rdocx_opc::relationship::rel_types;

        let type_suffix = match hdr_type {
            HdrFtrType::Default => "",
            HdrFtrType::First => "First",
            HdrFtrType::Even => "Even",
        };
        let (part_name, rel_type, content_type) = if is_header {
            (
                format!("/word/header{type_suffix}1.xml"),
                rel_types::HEADER,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
            )
        } else {
            (
                format!("/word/footer{type_suffix}1.xml"),
                rel_types::FOOTER,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
            )
        };

        // Store the raw header/footer XML
        self.package.set_part(&part_name, xml);
        self.package
            .content_types
            .add_override(&part_name, content_type);

        // Store each image and create relationships with the specified rel_ids
        for &(rel_id, image_data, image_filename) in images {
            let ext = image_filename
                .rsplit('.')
                .next()
                .unwrap_or("png")
                .to_lowercase();
            let img_content_type = match ext.as_str() {
                "png" => "image/png",
                "jpg" | "jpeg" => "image/jpeg",
                _ => "image/png",
            };

            let image_num = self.next_image_number();
            let img_part_name = format!("/word/media/image{image_num}.{ext}");
            self.package.set_part(&img_part_name, image_data.to_vec());
            self.package
                .content_types
                .add_default(&ext, img_content_type);

            // Create relationship in the header/footer part's rels with the EXACT rel_id
            let img_rel_target = format!("media/image{image_num}.{ext}");
            let hdr_rels = self.package.get_or_create_part_rels(&part_name);
            hdr_rels.add_with_id(rel_id, rel_types::IMAGE, &img_rel_target);
        }

        // Add relationship from document to header/footer
        let rel_target = part_name.trim_start_matches("/word/");
        let rels = self.package.get_or_create_part_rels(&self.doc_part_name);
        let rel_id = rels.add(rel_type, rel_target);

        // Add reference in section properties
        let sect = self.section_properties_mut();
        let refs = if is_header {
            &mut sect.header_refs
        } else {
            &mut sect.footer_refs
        };

        refs.retain(|r| r.hdr_ftr_type != hdr_type);
        refs.push(HdrFtrRef {
            hdr_ftr_type: hdr_type,
            rel_id,
        });
    }

    /// Set the default header to an inline image with a colored background.
    ///
    /// Creates a header part where the paragraph has shading fill set to
    /// `bg_color` (hex string, e.g. "000000" for black) and contains the
    /// inline image.
    pub fn set_header_image_with_background(
        &mut self,
        image_data: &[u8],
        image_filename: &str,
        width: Length,
        height: Length,
        bg_color: &str,
    ) {
        self.set_header_footer_image_bg_part(
            image_data,
            image_filename,
            width,
            height,
            Some(bg_color),
            true,
            HdrFtrType::Default,
        );
    }

    /// Set the first-page header to an inline image.
    pub fn set_first_page_header_image(
        &mut self,
        image_data: &[u8],
        image_filename: &str,
        width: Length,
        height: Length,
    ) {
        self.set_different_first_page(true);
        self.set_header_footer_image_part(
            image_data,
            image_filename,
            width,
            height,
            true,
            HdrFtrType::First,
        );
    }

    fn set_header_footer_image_part(
        &mut self,
        image_data: &[u8],
        image_filename: &str,
        width: Length,
        height: Length,
        is_header: bool,
        hdr_type: HdrFtrType,
    ) {
        use rdocx_opc::relationship::rel_types;

        // Determine part name based on type
        let type_suffix = match hdr_type {
            HdrFtrType::Default => "",
            HdrFtrType::First => "First",
            HdrFtrType::Even => "Even",
        };
        let (part_name, rel_type, content_type) = if is_header {
            (
                format!("/word/header{type_suffix}1.xml"),
                rel_types::HEADER,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
            )
        } else {
            (
                format!("/word/footer{type_suffix}1.xml"),
                rel_types::FOOTER,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
            )
        };

        // Embed the image in the package
        let ext = image_filename
            .rsplit('.')
            .next()
            .unwrap_or("png")
            .to_lowercase();
        let img_content_type = match ext.as_str() {
            "png" => "image/png",
            "jpg" | "jpeg" => "image/jpeg",
            _ => "image/png",
        };

        // Generate unique image name using cached counter
        let image_num = self.next_image_number();
        let img_part_name = format!("/word/media/image{image_num}.{ext}");
        self.package.set_part(&img_part_name, image_data.to_vec());
        self.package
            .content_types
            .add_default(&ext, img_content_type);

        // Create image relationship in the HEADER/FOOTER part's rels
        let img_rel_target = format!("media/image{image_num}.{ext}");
        let hdr_rels = self.package.get_or_create_part_rels(&part_name);
        let img_rel_id = hdr_rels.add(rel_types::IMAGE, &img_rel_target);

        // Build header/footer with image paragraph
        let inline = CT_Inline::new(&img_rel_id, width.to_emu(), height.to_emu());
        let drawing = CT_Drawing::inline(inline);
        let run = CT_R {
            properties: None,
            content: vec![RunContent::Drawing(drawing)],
            extra_xml: Vec::new(),
        };

        let mut hdr_ftr = CT_HdrFtr::new();
        let mut p = CT_P::new();
        p.runs.push(run);
        hdr_ftr.paragraphs.push(p);

        // Serialize
        let xml = if is_header {
            hdr_ftr
                .to_xml_header()
                .expect("header serialization failed")
        } else {
            hdr_ftr
                .to_xml_footer()
                .expect("footer serialization failed")
        };

        self.package.set_part(&part_name, xml);
        self.package
            .content_types
            .add_override(&part_name, content_type);

        // Add relationship from document to header/footer
        let rel_target = part_name.trim_start_matches("/word/");
        let rels = self.package.get_or_create_part_rels(&self.doc_part_name);
        let rel_id = rels.add(rel_type, rel_target);

        // Add reference in section properties
        let sect = self.section_properties_mut();
        let refs = if is_header {
            &mut sect.header_refs
        } else {
            &mut sect.footer_refs
        };

        refs.retain(|r| r.hdr_ftr_type != hdr_type);
        refs.push(HdrFtrRef {
            hdr_ftr_type: hdr_type,
            rel_id,
        });
    }

    fn set_header_footer_image_bg_part(
        &mut self,
        image_data: &[u8],
        image_filename: &str,
        width: Length,
        height: Length,
        bg_color: Option<&str>,
        is_header: bool,
        hdr_type: HdrFtrType,
    ) {
        use rdocx_opc::relationship::rel_types;
        use rdocx_oxml::properties::CT_Shd;

        // Determine part name based on type
        let type_suffix = match hdr_type {
            HdrFtrType::Default => "",
            HdrFtrType::First => "First",
            HdrFtrType::Even => "Even",
        };
        let (part_name, rel_type, content_type) = if is_header {
            (
                format!("/word/header{type_suffix}1.xml"),
                rel_types::HEADER,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
            )
        } else {
            (
                format!("/word/footer{type_suffix}1.xml"),
                rel_types::FOOTER,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
            )
        };

        // Embed the image in the package
        let ext = image_filename
            .rsplit('.')
            .next()
            .unwrap_or("png")
            .to_lowercase();
        let img_content_type = match ext.as_str() {
            "png" => "image/png",
            "jpg" | "jpeg" => "image/jpeg",
            _ => "image/png",
        };

        let image_num = self.next_image_number();
        let img_part_name = format!("/word/media/image{image_num}.{ext}");
        self.package.set_part(&img_part_name, image_data.to_vec());
        self.package
            .content_types
            .add_default(&ext, img_content_type);

        // Create image relationship in the HEADER/FOOTER part's rels
        let img_rel_target = format!("media/image{image_num}.{ext}");
        let hdr_rels = self.package.get_or_create_part_rels(&part_name);
        let img_rel_id = hdr_rels.add(rel_types::IMAGE, &img_rel_target);

        // Build header/footer with image paragraph
        let inline = CT_Inline::new(&img_rel_id, width.to_emu(), height.to_emu());
        let drawing = CT_Drawing::inline(inline);
        let run = CT_R {
            properties: None,
            content: vec![RunContent::Drawing(drawing)],
            extra_xml: Vec::new(),
        };

        let mut hdr_ftr = CT_HdrFtr::new();
        let mut p = CT_P::new();
        p.runs.push(run);

        // Add background shading if requested
        if let Some(color) = bg_color {
            let ppr = CT_PPr {
                shading: Some(CT_Shd {
                    val: "clear".to_string(),
                    color: Some("auto".to_string()),
                    fill: Some(color.to_string()),
                }),
                ..Default::default()
            };
            p.properties = Some(ppr);
        }

        hdr_ftr.paragraphs.push(p);

        // Serialize
        let xml = if is_header {
            hdr_ftr
                .to_xml_header()
                .expect("header serialization failed")
        } else {
            hdr_ftr
                .to_xml_footer()
                .expect("footer serialization failed")
        };

        self.package.set_part(&part_name, xml);
        self.package
            .content_types
            .add_override(&part_name, content_type);

        // Add relationship from document to header/footer
        let rel_target = part_name.trim_start_matches("/word/");
        let rels = self.package.get_or_create_part_rels(&self.doc_part_name);
        let rel_id = rels.add(rel_type, rel_target);

        // Add reference in section properties
        let sect = self.section_properties_mut();
        let refs = if is_header {
            &mut sect.header_refs
        } else {
            &mut sect.footer_refs
        };

        refs.retain(|r| r.hdr_ftr_type != hdr_type);
        refs.push(HdrFtrRef {
            hdr_ftr_type: hdr_type,
            rel_id,
        });
    }

    fn get_header_footer_text(&self, is_header: bool, hdr_type: HdrFtrType) -> Option<String> {
        let sect = self.document.body.sect_pr.as_ref()?;
        let refs = if is_header {
            &sect.header_refs
        } else {
            &sect.footer_refs
        };
        let hdr_ref = refs.iter().find(|r| r.hdr_ftr_type == hdr_type)?;

        // Resolve the part
        let rels = self.package.get_part_rels(&self.doc_part_name)?;
        let rel = rels.get_by_id(&hdr_ref.rel_id)?;
        let part_name = OpcPackage::resolve_rel_target(&self.doc_part_name, &rel.target);
        let xml = self.package.get_part(&part_name)?;
        let hdr_ftr = CT_HdrFtr::from_xml(xml).ok()?;
        Some(hdr_ftr.text())
    }

    // ---- Numbering/Lists ----

    /// Ensure a numbering part exists, creating it and its relationship if needed.
    fn ensure_numbering(&mut self) -> &mut CT_Numbering {
        if self.numbering.is_none() {
            self.numbering = Some(CT_Numbering::new());

            // Set up numbering relationship and content type
            self.package
                .get_or_create_part_rels(&self.doc_part_name)
                .add(rel_types::NUMBERING, "numbering.xml");
            self.package.content_types.add_override(
                "/word/numbering.xml",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
            );
        }
        self.numbering.as_mut().unwrap()
    }

    /// Add a bullet list item at the given indentation level (0-based).
    ///
    /// If no bullet list definition exists yet, one is created automatically.
    /// Returns a mutable `Paragraph` for further configuration.
    pub fn add_bullet_list_item(&mut self, text: &str, level: u32) -> Paragraph<'_> {
        // Find or create a bullet list numId
        let num_id = {
            let numbering = self.ensure_numbering();
            // Look for an existing bullet list
            let existing = numbering.nums.iter().find(|n| {
                numbering
                    .get_abstract_num_for(n.num_id)
                    .map(|a| {
                        a.levels.first().and_then(|l| l.num_fmt)
                            == Some(rdocx_oxml::numbering::ST_NumberFormat::Bullet)
                    })
                    .unwrap_or(false)
            });
            if let Some(existing) = existing {
                existing.num_id
            } else {
                numbering.add_bullet_list()
            }
        };

        let mut p = CT_P::new();
        if !text.is_empty() {
            p.add_run(text);
        }
        let ppr = CT_PPr {
            num_id: Some(num_id),
            num_ilvl: Some(level),
            ..Default::default()
        };
        p.properties = Some(ppr);

        self.document.body.content.push(BodyContent::Paragraph(p));
        match self.document.body.content.last_mut().unwrap() {
            BodyContent::Paragraph(p) => Paragraph { inner: p },
            _ => unreachable!(),
        }
    }

    /// Add a numbered list item at the given indentation level (0-based).
    ///
    /// If no numbered list definition exists yet, one is created automatically.
    /// Returns a mutable `Paragraph` for further configuration.
    pub fn add_numbered_list_item(&mut self, text: &str, level: u32) -> Paragraph<'_> {
        // Find or create a numbered list numId
        let num_id = {
            let numbering = self.ensure_numbering();
            // Look for an existing numbered list
            let existing = numbering.nums.iter().find(|n| {
                numbering
                    .get_abstract_num_for(n.num_id)
                    .map(|a| {
                        a.levels.first().and_then(|l| l.num_fmt)
                            == Some(rdocx_oxml::numbering::ST_NumberFormat::Decimal)
                    })
                    .unwrap_or(false)
            });
            if let Some(existing) = existing {
                existing.num_id
            } else {
                numbering.add_numbered_list()
            }
        };

        let mut p = CT_P::new();
        if !text.is_empty() {
            p.add_run(text);
        }
        let ppr = CT_PPr {
            num_id: Some(num_id),
            num_ilvl: Some(level),
            ..Default::default()
        };
        p.properties = Some(ppr);

        self.document.body.content.push(BodyContent::Paragraph(p));
        match self.document.body.content.last_mut().unwrap() {
            BodyContent::Paragraph(p) => Paragraph { inner: p },
            _ => unreachable!(),
        }
    }

    // ---- Style access ----

    /// Get all styles.
    pub fn styles(&self) -> Vec<Style<'_>> {
        self.styles
            .styles
            .iter()
            .map(|s| Style { inner: s })
            .collect()
    }

    /// Find a style by its ID.
    pub fn style(&self, style_id: &str) -> Option<Style<'_>> {
        self.styles.get_by_id(style_id).map(|s| Style { inner: s })
    }

    // ---- Style manipulation ----

    /// Add a custom style to the document.
    pub fn add_style(&mut self, builder: StyleBuilder) {
        self.styles.styles.push(builder.build());
    }

    /// Resolve the effective paragraph properties for a given style ID,
    /// walking the full inheritance chain (docDefaults → basedOn → ...).
    pub fn resolve_paragraph_properties(&self, style_id: Option<&str>) -> CT_PPr {
        style::resolve_paragraph_properties(style_id, &self.styles)
    }

    /// Resolve the effective run properties for the given paragraph and character styles,
    /// walking the full inheritance chain.
    pub fn resolve_run_properties(
        &self,
        para_style_id: Option<&str>,
        run_style_id: Option<&str>,
    ) -> CT_RPr {
        style::resolve_run_properties(para_style_id, run_style_id, &self.styles)
    }

    // ---- Section/Page setup ----

    /// Get the section properties (page size, margins).
    pub fn section_properties(&self) -> Option<&CT_SectPr> {
        self.document.body.sect_pr.as_ref()
    }

    /// Get a mutable reference to section properties, creating defaults if needed.
    pub fn section_properties_mut(&mut self) -> &mut CT_SectPr {
        self.document
            .body
            .sect_pr
            .get_or_insert_with(CT_SectPr::default_letter)
    }

    /// Set page size.
    pub fn set_page_size(&mut self, width: Length, height: Length) {
        let sect = self.section_properties_mut();
        sect.page_width = Some(width.as_twips());
        sect.page_height = Some(height.as_twips());
    }

    /// Set page orientation to landscape (swaps width and height if needed).
    pub fn set_landscape(&mut self) {
        let sect = self.section_properties_mut();
        sect.orientation = Some(ST_PageOrientation::Landscape);
        // Swap width/height if portrait dimensions
        if let (Some(w), Some(h)) = (sect.page_width, sect.page_height)
            && w.0 < h.0
        {
            sect.page_width = Some(h);
            sect.page_height = Some(w);
        }
    }

    /// Set page orientation to portrait (swaps width and height if needed).
    pub fn set_portrait(&mut self) {
        let sect = self.section_properties_mut();
        sect.orientation = Some(ST_PageOrientation::Portrait);
        // Swap width/height if landscape dimensions
        if let (Some(w), Some(h)) = (sect.page_width, sect.page_height)
            && w.0 > h.0
        {
            sect.page_width = Some(h);
            sect.page_height = Some(w);
        }
    }

    /// Set all page margins.
    pub fn set_margins(&mut self, top: Length, right: Length, bottom: Length, left: Length) {
        let sect = self.section_properties_mut();
        sect.margin_top = Some(top.as_twips());
        sect.margin_right = Some(right.as_twips());
        sect.margin_bottom = Some(bottom.as_twips());
        sect.margin_left = Some(left.as_twips());
    }

    /// Set equal-width column layout.
    pub fn set_columns(&mut self, num: u32, spacing: Length) {
        let sect = self.section_properties_mut();
        sect.columns = Some(CT_Columns {
            num: Some(num),
            space: Some(spacing.as_twips()),
            equal_width: Some(true),
            sep: None,
            columns: Vec::new(),
        });
    }

    /// Set header and footer distances from page edges.
    pub fn set_header_footer_distance(&mut self, header: Length, footer: Length) {
        let sect = self.section_properties_mut();
        sect.header_distance = Some(header.as_twips());
        sect.footer_distance = Some(footer.as_twips());
    }

    /// Set the gutter margin.
    pub fn set_gutter(&mut self, gutter: Length) {
        self.section_properties_mut().gutter = Some(gutter.as_twips());
    }

    /// Enable or disable different first page header/footer.
    pub fn set_different_first_page(&mut self, val: bool) {
        self.section_properties_mut().title_pg = Some(val);
    }

    // ---- Metadata access ----

    /// Get the document title.
    pub fn title(&self) -> Option<&str> {
        self.core_properties.as_ref()?.title.as_deref()
    }

    /// Set the document title.
    pub fn set_title(&mut self, title: &str) {
        self.ensure_core_properties().title = Some(title.to_string());
    }

    /// Get the document author/creator.
    pub fn author(&self) -> Option<&str> {
        self.core_properties.as_ref()?.creator.as_deref()
    }

    /// Set the document author/creator.
    pub fn set_author(&mut self, author: &str) {
        self.ensure_core_properties().creator = Some(author.to_string());
    }

    /// Get the document subject.
    pub fn subject(&self) -> Option<&str> {
        self.core_properties.as_ref()?.subject.as_deref()
    }

    /// Set the document subject.
    pub fn set_subject(&mut self, subject: &str) {
        self.ensure_core_properties().subject = Some(subject.to_string());
    }

    /// Get the document keywords.
    pub fn keywords(&self) -> Option<&str> {
        self.core_properties.as_ref()?.keywords.as_deref()
    }

    /// Set the document keywords.
    pub fn set_keywords(&mut self, keywords: &str) {
        self.ensure_core_properties().keywords = Some(keywords.to_string());
    }

    fn ensure_core_properties(&mut self) -> &mut CoreProperties {
        self.core_properties
            .get_or_insert_with(CoreProperties::default)
    }

    // ---- Document Merging ----

    /// Append the content of another document to this document.
    ///
    /// Copies all body content (paragraphs and tables) from the other document.
    /// Handles style deduplication and numbering remapping.
    pub fn append(&mut self, other: &Document) {
        self.merge_styles(other);

        let start_idx = self.document.body.content.len();
        for content in &other.document.body.content {
            self.document.body.content.push(content.clone());
        }

        self.remap_merged_numbering(other, start_idx);
    }

    /// Append the content of another document with a section break.
    pub fn append_with_break(&mut self, other: &Document, break_type: crate::SectionBreak) {
        // Insert a section break paragraph before the merged content
        let mut p = CT_P::new();
        let sect_pr = match break_type {
            crate::SectionBreak::NextPage => CT_SectPr::default_letter(),
            crate::SectionBreak::Continuous => {
                let mut sp = CT_SectPr::default_letter();
                sp.section_type = Some(ST_SectionType::Continuous);
                sp
            }
            crate::SectionBreak::EvenPage => {
                let mut sp = CT_SectPr::default_letter();
                sp.section_type = Some(ST_SectionType::EvenPage);
                sp
            }
            crate::SectionBreak::OddPage => {
                let mut sp = CT_SectPr::default_letter();
                sp.section_type = Some(ST_SectionType::OddPage);
                sp
            }
        };
        p.properties = Some(CT_PPr {
            sect_pr: Some(sect_pr),
            ..Default::default()
        });
        self.document.body.content.push(BodyContent::Paragraph(p));

        self.append(other);
    }

    /// Insert the content of another document at a specified body index.
    pub fn insert_document(&mut self, index: usize, other: &Document) {
        self.merge_styles(other);

        let insert_at = index.min(self.document.body.content.len());
        for (i, content) in other.document.body.content.iter().enumerate() {
            self.document
                .body
                .content
                .insert(insert_at + i, content.clone());
        }

        self.remap_merged_numbering(other, insert_at);
    }

    /// Merge styles from another document, avoiding duplicates.
    fn merge_styles(&mut self, other: &Document) {
        for style in &other.styles.styles {
            if self.styles.get_by_id(&style.style_id).is_none() {
                self.styles.styles.push(style.clone());
            }
        }
    }

    /// Merge numbering from another document and remap IDs in the merged content.
    /// `start_idx` is the index where the other document's content starts in self.
    fn remap_merged_numbering(&mut self, other: &Document, start_idx: usize) {
        let Some(other_numbering) = &other.numbering else {
            return;
        };

        let numbering = self
            .numbering
            .get_or_insert_with(|| rdocx_oxml::numbering::CT_Numbering {
                abstract_nums: Vec::new(),
                nums: Vec::new(),
            });

        // Find max existing IDs to avoid collision
        let max_abstract_id = numbering
            .abstract_nums
            .iter()
            .map(|a| a.abstract_num_id)
            .max()
            .unwrap_or(0);
        let max_num_id = numbering.nums.iter().map(|n| n.num_id).max().unwrap_or(0);

        let abstract_offset = max_abstract_id + 1;
        let num_offset = max_num_id + 1;

        // Copy abstract nums with remapped IDs
        for abs_num in &other_numbering.abstract_nums {
            let mut new_abs = abs_num.clone();
            new_abs.abstract_num_id += abstract_offset;
            numbering.abstract_nums.push(new_abs);
        }

        // Copy num instances with remapped IDs
        for num in &other_numbering.nums {
            let mut new_num = num.clone();
            new_num.num_id += num_offset;
            new_num.abstract_num_id += abstract_offset;
            numbering.nums.push(new_num);
        }

        // Remap numId references in the merged content
        let incoming_count = other.document.body.content.len();
        for content in self.document.body.content[start_idx..start_idx + incoming_count].iter_mut()
        {
            Self::remap_num_ids(content, num_offset);
        }
    }

    /// Remap numId references in body content by adding an offset.
    fn remap_num_ids(content: &mut BodyContent, offset: u32) {
        match content {
            BodyContent::Paragraph(p) => {
                Self::remap_paragraph_num_id(p, offset);
            }
            BodyContent::Table(tbl) => {
                Self::remap_table_num_ids(tbl, offset);
            }
            BodyContent::RawXml(_) => {}
        }
    }

    fn remap_paragraph_num_id(p: &mut CT_P, offset: u32) {
        if let Some(ppr) = &mut p.properties
            && let Some(num_id) = &mut ppr.num_id
            && *num_id > 0
        {
            *num_id += offset;
        }
    }

    fn remap_table_num_ids(tbl: &mut CT_Tbl, offset: u32) {
        for row in &mut tbl.rows {
            for cell in &mut row.cells {
                for cc in &mut cell.content {
                    match cc {
                        rdocx_oxml::table::CellContent::Paragraph(p) => {
                            Self::remap_paragraph_num_id(p, offset);
                        }
                        rdocx_oxml::table::CellContent::Table(nested) => {
                            Self::remap_table_num_ids(nested, offset);
                        }
                    }
                }
            }
        }
    }

    // ---- Table of Contents ----

    /// Insert a Table of Contents at the given body content index.
    ///
    /// Scans the document for heading paragraphs (Heading1..HeadingN where N <= max_level),
    /// inserts bookmark markers at each heading, and generates TOC entry paragraphs
    /// with internal hyperlinks and dot-leader tab stops.
    ///
    /// # Arguments
    /// * `index` - Body content index at which to insert the TOC
    /// * `max_level` - Maximum heading level to include (1-9, typically 3)
    pub fn insert_toc(&mut self, index: usize, max_level: u32) {
        use rdocx_oxml::borders::{CT_TabStop, CT_Tabs};
        use rdocx_oxml::shared::{ST_TabJc, ST_TabLeader};
        use rdocx_oxml::text::HyperlinkSpan;
        use rdocx_oxml::units::Twips;

        let max_level = max_level.clamp(1, 9);

        // Step 1: Collect heading info from the document body
        struct HeadingInfo {
            content_index: usize,
            level: u32,
            text: String,
            bookmark_name: String,
        }

        let mut headings = Vec::new();
        let mut toc_counter = 0u32;

        for (idx, content) in self.document.body.content.iter().enumerate() {
            if let BodyContent::Paragraph(p) = content
                && let Some(level) = Self::detect_heading_level_for_toc(p)
                && level <= max_level
            {
                let text = p.text();
                if !text.trim().is_empty() {
                    toc_counter += 1;
                    headings.push(HeadingInfo {
                        content_index: idx,
                        level,
                        text,
                        bookmark_name: format!("_Toc{toc_counter}"),
                    });
                }
            }
        }

        // Step 2: Insert bookmark markers at each heading paragraph (as raw XML in extra_xml)
        // We insert bookmarkStart/bookmarkEnd as extra_xml at position 0 in the paragraph.
        // Adjust for insertions that shift indices.
        let mut bookmark_id = 100; // Start at a high ID to avoid collision
        for heading in &headings {
            if let Some(BodyContent::Paragraph(p)) =
                self.document.body.content.get_mut(heading.content_index)
            {
                let bm_start = format!(
                    "<w:bookmarkStart w:id=\"{bookmark_id}\" w:name=\"{}\"/>",
                    heading.bookmark_name
                );
                let bm_end = format!("<w:bookmarkEnd w:id=\"{bookmark_id}\"/>");
                // Insert at position 0 (before runs)
                p.extra_xml.push((0, bm_start.into_bytes()));
                // Insert at end (after runs)
                p.extra_xml.push((p.runs.len(), bm_end.into_bytes()));
                bookmark_id += 1;
            }
        }

        // Step 3: Build TOC entry paragraphs
        // Right margin tab stop at 9360 twips (6.5") with dot leader
        let right_tab = CT_Tabs {
            tabs: vec![CT_TabStop {
                val: ST_TabJc::Right,
                pos: Twips(9360),
                leader: Some(ST_TabLeader::Dot),
            }],
        };

        let mut toc_paragraphs: Vec<CT_P> = Vec::new();

        // TOC title
        let mut title_p = CT_P::new();
        let mut title_r = CT_R::new("Table of Contents");
        title_r.properties = Some(CT_RPr {
            bold: Some(true),
            ..Default::default()
        });
        title_p.runs.push(title_r);
        title_p.properties = Some(CT_PPr {
            space_after: Some(Twips(120)),
            ..Default::default()
        });
        toc_paragraphs.push(title_p);

        for heading in &headings {
            let mut p = CT_P::new();

            // Indentation based on heading level (each level indented 360 twips = 0.25")
            let indent = Twips(360 * (heading.level as i32 - 1));

            p.properties = Some(CT_PPr {
                tabs: Some(right_tab.clone()),
                ind_left: if indent.0 > 0 { Some(indent) } else { None },
                ..Default::default()
            });

            // Run with heading text
            let text_run = CT_R::new(&heading.text);
            p.runs.push(text_run);

            // Tab run (separates text from page number)
            p.runs.push(CT_R {
                properties: None,
                content: vec![rdocx_oxml::text::RunContent::Tab],
                extra_xml: Vec::new(),
            });

            // Wrap the text run in a hyperlink to the bookmark
            p.hyperlinks.push(HyperlinkSpan {
                rel_id: None,
                anchor: Some(heading.bookmark_name.clone()),
                run_start: 0,
                run_end: 1, // Just the text run, not the tab
            });

            toc_paragraphs.push(p);
        }

        // Step 4: Insert TOC paragraphs at the specified index
        let insert_at = index.min(self.document.body.content.len());
        for (i, p) in toc_paragraphs.into_iter().enumerate() {
            self.document
                .body
                .content
                .insert(insert_at + i, BodyContent::Paragraph(p));
        }
    }

    /// Detect heading level from a paragraph's style ID.
    fn detect_heading_level_for_toc(para: &CT_P) -> Option<u32> {
        let ppr = para.properties.as_ref()?;
        let style_id = ppr.style_id.as_deref()?;
        let rest = style_id.strip_prefix("Heading")?;
        rest.parse::<u32>().ok().filter(|n| (1..=9).contains(n))
    }

    // ---- Placeholder replacement ----

    /// Replace all occurrences of `placeholder` with `replacement` throughout the document.
    ///
    /// Searches body paragraphs, tables (including nested), headers, and footers.
    /// Handles placeholders split across multiple runs. Returns the total number
    /// of replacements made.
    pub fn replace_text(&mut self, placeholder: &str, replacement: &str) -> usize {
        use rdocx_oxml::placeholder;

        let mut count = 0;

        // Replace in body paragraphs
        for content in &mut self.document.body.content {
            match content {
                BodyContent::Paragraph(p) => {
                    count += placeholder::replace_in_paragraph(p, placeholder, replacement);
                }
                BodyContent::Table(t) => {
                    count += placeholder::replace_in_table(t, placeholder, replacement);
                }
                _ => {} // Skip RawXml elements
            }
        }

        // Replace in headers and footers
        if let Some(sect_pr) = self.document.body.sect_pr.as_ref() {
            let hdr_rel_ids: Vec<String> = sect_pr
                .header_refs
                .iter()
                .map(|r| r.rel_id.clone())
                .collect();
            let ftr_rel_ids: Vec<String> = sect_pr
                .footer_refs
                .iter()
                .map(|r| r.rel_id.clone())
                .collect();

            for rel_id in hdr_rel_ids {
                if let Some(mut hf) = self.load_header_footer(&rel_id) {
                    let n =
                        placeholder::replace_in_header_footer(&mut hf, placeholder, replacement);
                    if n > 0 {
                        self.save_header_footer(&rel_id, &hf, true);
                        count += n;
                    }
                }
            }
            for rel_id in ftr_rel_ids {
                if let Some(mut hf) = self.load_header_footer(&rel_id) {
                    let n =
                        placeholder::replace_in_header_footer(&mut hf, placeholder, replacement);
                    if n > 0 {
                        self.save_header_footer(&rel_id, &hf, false);
                        count += n;
                    }
                }
            }
        }

        // Flush document to package, then do raw XML pass for text boxes/shapes
        if let Ok(()) = self.flush_to_package() {
            count += self.replace_in_xml_parts(placeholder, replacement);
        }

        count
    }

    /// Replace multiple placeholders at once. Returns total replacements.
    pub fn replace_all(&mut self, replacements: &std::collections::HashMap<&str, &str>) -> usize {
        let mut count = 0;
        for (placeholder, replacement) in replacements {
            count += self.replace_text(placeholder, replacement);
        }
        count
    }

    // ---- Regex replacement ----

    /// Replace all regex matches with `replacement` throughout the document.
    ///
    /// The `replacement` string supports capture groups: `$1`, `$2`, etc.
    /// Searches body paragraphs, tables (including nested), headers, and footers.
    /// Returns the total number of replacements made, or an error if the regex is invalid.
    pub fn replace_regex(&mut self, pattern: &str, replacement: &str) -> Result<usize> {
        let re =
            regex::Regex::new(pattern).map_err(|e| Error::Other(format!("invalid regex: {e}")))?;
        Ok(self.replace_regex_compiled(&re, replacement))
    }

    /// Replace multiple regex patterns at once. Returns total replacements.
    pub fn replace_all_regex(&mut self, patterns: &[(String, String)]) -> Result<usize> {
        let mut count = 0;
        for (pattern, replacement) in patterns {
            count += self.replace_regex(pattern, replacement)?;
        }
        Ok(count)
    }

    /// Internal: replace using a pre-compiled regex.
    fn replace_regex_compiled(&mut self, re: &regex::Regex, replacement: &str) -> usize {
        use rdocx_oxml::placeholder;

        let mut count = 0;

        // Replace in body paragraphs and tables
        for content in &mut self.document.body.content {
            match content {
                BodyContent::Paragraph(p) => {
                    count += placeholder::replace_regex_in_paragraph(p, re, replacement);
                }
                BodyContent::Table(t) => {
                    count += placeholder::replace_regex_in_table(t, re, replacement);
                }
                _ => {}
            }
        }

        // Replace in headers and footers
        if let Some(sect_pr) = self.document.body.sect_pr.as_ref() {
            let hdr_rel_ids: Vec<String> = sect_pr
                .header_refs
                .iter()
                .map(|r| r.rel_id.clone())
                .collect();
            let ftr_rel_ids: Vec<String> = sect_pr
                .footer_refs
                .iter()
                .map(|r| r.rel_id.clone())
                .collect();

            for rel_id in hdr_rel_ids {
                if let Some(mut hf) = self.load_header_footer(&rel_id) {
                    let n = placeholder::replace_regex_in_header_footer(&mut hf, re, replacement);
                    if n > 0 {
                        self.save_header_footer(&rel_id, &hf, true);
                        count += n;
                    }
                }
            }
            for rel_id in ftr_rel_ids {
                if let Some(mut hf) = self.load_header_footer(&rel_id) {
                    let n = placeholder::replace_regex_in_header_footer(&mut hf, re, replacement);
                    if n > 0 {
                        self.save_header_footer(&rel_id, &hf, false);
                        count += n;
                    }
                }
            }
        }

        count
    }

    /// Load a header/footer part by its relationship ID.
    fn load_header_footer(&self, rel_id: &str) -> Option<CT_HdrFtr> {
        let rels = self.package.get_part_rels(&self.doc_part_name)?;
        let rel = rels.get_by_id(rel_id)?;
        let part_name = OpcPackage::resolve_rel_target(&self.doc_part_name, &rel.target);
        let xml = self.package.get_part(&part_name)?;
        CT_HdrFtr::from_xml(xml).ok()
    }

    /// Run raw XML replacement on all XML parts (for text boxes, shapes, charts, etc.).
    ///
    /// This is called after the typed-model replacement and flush_to_package.
    fn replace_in_xml_parts(&mut self, placeholder: &str, replacement: &str) -> usize {
        use rdocx_oxml::placeholder::{replace_in_chart_xml, replace_in_xml_part};

        let mut count = 0;

        // Collect part names for XML parts to process (text boxes/shapes)
        let mut xml_parts: Vec<String> = vec![self.doc_part_name.clone()];
        if let Some(sect_pr) = self.document.body.sect_pr.as_ref()
            && let Some(rels) = self.package.get_part_rels(&self.doc_part_name)
        {
            for href in &sect_pr.header_refs {
                if let Some(rel) = rels.get_by_id(&href.rel_id) {
                    xml_parts.push(OpcPackage::resolve_rel_target(
                        &self.doc_part_name,
                        &rel.target,
                    ));
                }
            }
            for fref in &sect_pr.footer_refs {
                if let Some(rel) = rels.get_by_id(&fref.rel_id) {
                    xml_parts.push(OpcPackage::resolve_rel_target(
                        &self.doc_part_name,
                        &rel.target,
                    ));
                }
            }
        }

        for part_name in xml_parts {
            if let Some(xml) = self.package.get_part(&part_name) {
                let xml = xml.to_vec();
                if let Ok((new_xml, n)) = replace_in_xml_part(&xml, placeholder, replacement)
                    && n > 0
                {
                    self.package.set_part(&part_name, new_xml);
                    count += n;
                }
            }
        }

        // Collect chart part names
        let chart_parts: Vec<String> = self
            .package
            .get_part_rels(&self.doc_part_name)
            .map(|rels| {
                rels.get_all_by_type(rel_types::CHART)
                    .iter()
                    .map(|rel| OpcPackage::resolve_rel_target(&self.doc_part_name, &rel.target))
                    .collect()
            })
            .unwrap_or_default();

        for part_name in chart_parts {
            if let Some(xml) = self.package.get_part(&part_name) {
                let xml = xml.to_vec();
                if let Ok((new_xml, n)) = replace_in_chart_xml(&xml, placeholder, replacement)
                    && n > 0
                {
                    self.package.set_part(&part_name, new_xml);
                    count += n;
                }
            }
        }

        // Re-parse document from the (possibly modified) package XML
        if count > 0
            && let Some(doc_xml) = self.package.get_part(&self.doc_part_name)
            && let Ok(doc) = CT_Document::from_xml(doc_xml)
        {
            self.document = doc;
        }

        count
    }

    // ---- PDF conversion ----

    /// Render the document to PDF bytes.
    ///
    /// This performs a full layout pass (font shaping, line breaking, pagination)
    /// and then renders the result to a PDF document.
    ///
    /// Font resolution order:
    /// 1. Fonts embedded in the DOCX file (word/fonts/)
    /// 2. System fonts
    /// 3. Bundled fonts (if `bundled-fonts` feature is enabled)
    pub fn to_pdf(&self) -> Result<Vec<u8>> {
        self.to_pdf_with_fonts(&[])
    }

    /// Render the document to PDF bytes with user-provided font files.
    ///
    /// User-provided fonts take highest priority in font resolution.
    ///
    /// # Arguments
    /// * `font_files` - Additional font files to use. Each entry is `(family_name, font_bytes)`.
    ///
    /// Font resolution order:
    /// 1. User-provided fonts (this parameter)
    /// 2. Fonts embedded in the DOCX file (word/fonts/)
    /// 3. System fonts
    /// 4. Bundled fonts (if `bundled-fonts` feature is enabled)
    pub fn to_pdf_with_fonts(&self, font_files: &[(&str, &[u8])]) -> Result<Vec<u8>> {
        let mut input = self.build_layout_input();
        for (family, data) in font_files {
            input.fonts.push(rdocx_layout::FontFile {
                family: family.to_string(),
                data: data.to_vec(),
            });
        }
        let layout = rdocx_layout::layout_document(&input)?;
        Ok(rdocx_pdf::render_to_pdf(&layout))
    }

    /// Save the document as a PDF file.
    pub fn save_pdf<P: AsRef<Path>>(&self, path: P) -> Result<()> {
        let pdf_bytes = self.to_pdf()?;
        std::fs::write(path, pdf_bytes)?;
        Ok(())
    }

    /// Convert the document to a complete HTML document string.
    pub fn to_html(&self) -> String {
        let input = self.build_html_input();
        rdocx_html::to_html_document(&input, &rdocx_html::HtmlOptions::default())
    }

    /// Convert the document to an HTML fragment (body content only, no `<html>` wrapper).
    pub fn to_html_fragment(&self) -> String {
        let input = self.build_html_input();
        rdocx_html::to_html_fragment(&input, &rdocx_html::HtmlOptions::default())
    }

    /// Convert the document to Markdown.
    pub fn to_markdown(&self) -> String {
        let input = self.build_html_input();
        rdocx_html::to_markdown(&input)
    }

    /// Build an HtmlInput from the document's current state.
    fn build_html_input(&self) -> rdocx_html::HtmlInput {
        use rdocx_opc::relationship::rel_types;
        use std::collections::HashMap;

        let mut images: HashMap<String, rdocx_html::ImageData> = HashMap::new();
        let mut hyperlink_urls: HashMap<String, String> = HashMap::new();

        if let Some(rels) = self.package.get_part_rels(&self.doc_part_name) {
            for rel in &rels.items {
                match rel.rel_type.as_str() {
                    t if t == rel_types::IMAGE => {
                        let part_name =
                            OpcPackage::resolve_rel_target(&self.doc_part_name, &rel.target);
                        if let Some(data) = self.package.get_part(&part_name) {
                            let content_type = guess_image_content_type(&part_name);
                            images.insert(
                                rel.id.clone(),
                                rdocx_html::ImageData {
                                    data: data.to_vec(),
                                    content_type,
                                },
                            );
                        }
                    }
                    t if t == rel_types::HYPERLINK => {
                        if rel.target_mode.as_ref().is_some_and(|m| m == "External") {
                            hyperlink_urls.insert(rel.id.clone(), rel.target.clone());
                        }
                    }
                    _ => {}
                }
            }
        }

        rdocx_html::HtmlInput {
            document: self.document.clone(),
            styles: self.styles.clone(),
            numbering: self.numbering.clone(),
            images,
            hyperlink_urls,
        }
    }

    /// Render a single page of the document to PNG bytes.
    ///
    /// # Arguments
    /// * `page_index` - 0-based page index
    /// * `dpi` - Resolution (72 = 1:1, 150 = standard, 300 = high quality)
    pub fn render_page_to_png(&self, page_index: usize, dpi: f64) -> Result<Option<Vec<u8>>> {
        let input = self.build_layout_input();
        let layout = rdocx_layout::layout_document(&input)?;
        Ok(rdocx_pdf::render_page_to_png(&layout, page_index, dpi))
    }

    /// Render all pages of the document to PNG bytes.
    pub fn render_all_pages(&self, dpi: f64) -> Result<Vec<Vec<u8>>> {
        let input = self.build_layout_input();
        let layout = rdocx_layout::layout_document(&input)?;
        Ok(rdocx_pdf::render_all_pages(&layout, dpi))
    }

    /// Build a LayoutInput from the document's current state.
    fn build_layout_input(&self) -> rdocx_layout::LayoutInput {
        use rdocx_layout::{ImageData, LayoutInput};
        use rdocx_opc::relationship::rel_types;
        use std::collections::HashMap;

        let mut headers: HashMap<String, CT_HdrFtr> = HashMap::new();
        let mut footers: HashMap<String, CT_HdrFtr> = HashMap::new();
        let mut images: HashMap<String, ImageData> = HashMap::new();
        let mut hyperlink_urls: HashMap<String, String> = HashMap::new();
        let mut footnotes = None;
        let mut endnotes = None;

        // Extract embedded fonts from the DOCX package
        let fonts = self.extract_embedded_fonts();

        if let Some(rels) = self.package.get_part_rels(&self.doc_part_name) {
            for rel in &rels.items {
                match rel.rel_type.as_str() {
                    t if t == rel_types::HEADER => {
                        let part_name =
                            OpcPackage::resolve_rel_target(&self.doc_part_name, &rel.target);
                        if let Some(xml) = self.package.get_part(&part_name)
                            && let Ok(hf) = CT_HdrFtr::from_xml(xml)
                        {
                            headers.insert(rel.id.clone(), hf);
                        }
                    }
                    t if t == rel_types::FOOTER => {
                        let part_name =
                            OpcPackage::resolve_rel_target(&self.doc_part_name, &rel.target);
                        if let Some(xml) = self.package.get_part(&part_name)
                            && let Ok(hf) = CT_HdrFtr::from_xml(xml)
                        {
                            footers.insert(rel.id.clone(), hf);
                        }
                    }
                    t if t == rel_types::IMAGE => {
                        let part_name =
                            OpcPackage::resolve_rel_target(&self.doc_part_name, &rel.target);
                        if let Some(data) = self.package.get_part(&part_name) {
                            let content_type = guess_image_content_type(&part_name);
                            images.insert(
                                rel.id.clone(),
                                ImageData {
                                    data: data.to_vec(),
                                    content_type,
                                },
                            );
                        }
                    }
                    t if t == rel_types::HYPERLINK => {
                        if rel.target_mode.as_ref().is_some_and(|m| m == "External") {
                            hyperlink_urls.insert(rel.id.clone(), rel.target.clone());
                        }
                    }
                    t if t == rel_types::FOOTNOTES => {
                        let part_name =
                            OpcPackage::resolve_rel_target(&self.doc_part_name, &rel.target);
                        if let Some(xml) = self.package.get_part(&part_name) {
                            footnotes = rdocx_oxml::footnotes::CT_Footnotes::from_xml(xml).ok();
                        }
                    }
                    t if t == rel_types::ENDNOTES => {
                        let part_name =
                            OpcPackage::resolve_rel_target(&self.doc_part_name, &rel.target);
                        if let Some(xml) = self.package.get_part(&part_name) {
                            endnotes = rdocx_oxml::footnotes::CT_Footnotes::from_xml(xml).ok();
                        }
                    }
                    _ => {}
                }
            }
        }

        // Parse theme if available
        let theme = self
            .package
            .get_part("/word/theme/theme1.xml")
            .and_then(|data| rdocx_oxml::theme::Theme::from_xml(data).ok());

        LayoutInput {
            document: self.document.clone(),
            styles: self.styles.clone(),
            numbering: self.numbering.clone(),
            headers,
            footers,
            images,
            core_properties: self.core_properties.clone(),
            hyperlink_urls,
            footnotes,
            endnotes,
            theme,
            fonts,
        }
    }

    /// Extract embedded fonts from the DOCX package.
    ///
    /// Word can embed fonts as `.odttf` (obfuscated TrueType) or regular `.ttf`/`.otf`
    /// files in the `word/fonts/` directory. ODTTF files have the first 32 bytes
    /// XOR'd with a 16-byte GUID derived from the font's relationship ID.
    fn extract_embedded_fonts(&self) -> Vec<rdocx_layout::FontFile> {
        let mut fonts = Vec::new();

        // Look for font parts in word/fonts/ directory
        for (part_name, data) in &self.package.parts {
            let lower = part_name.to_lowercase();
            if !lower.contains("/word/fonts/") && !lower.contains("/word/font") {
                continue;
            }

            // Determine font family name from the file name
            let file_name = part_name.rsplit('/').next().unwrap_or(part_name);
            let family = file_name.split('.').next().unwrap_or(file_name).to_string();

            if lower.ends_with(".odttf") {
                // Deobfuscate ODTTF: XOR first 32 bytes with GUID from the file name
                if let Some(deobfuscated) = deobfuscate_odttf(data, file_name) {
                    fonts.push(rdocx_layout::FontFile {
                        family,
                        data: deobfuscated,
                    });
                }
            } else if lower.ends_with(".ttf") || lower.ends_with(".otf") || lower.ends_with(".ttc")
            {
                fonts.push(rdocx_layout::FontFile {
                    family,
                    data: data.clone(),
                });
            }
        }

        fonts
    }

    /// Load font files from a directory and return them as FontFile entries.
    ///
    /// This is useful for CLI tools that accept a `--font-dir` argument.
    /// Supports `.ttf`, `.otf`, and `.ttc` files.
    pub fn load_fonts_from_dir<P: AsRef<Path>>(dir: P) -> Vec<rdocx_layout::FontFile> {
        let mut fonts = Vec::new();
        let dir = dir.as_ref();
        if let Ok(entries) = std::fs::read_dir(dir) {
            for entry in entries.flatten() {
                let path = entry.path();
                let ext = path
                    .extension()
                    .and_then(|e| e.to_str())
                    .unwrap_or("")
                    .to_lowercase();
                if (ext == "ttf" || ext == "otf" || ext == "ttc")
                    && let Ok(data) = std::fs::read(&path)
                {
                    let family = path
                        .file_stem()
                        .and_then(|s| s.to_str())
                        .unwrap_or("Unknown")
                        .to_string();
                    fonts.push(rdocx_layout::FontFile { family, data });
                }
            }
        }
        fonts
    }

    /// Save a header/footer part back to the OPC package.
    fn save_header_footer(&mut self, rel_id: &str, hf: &CT_HdrFtr, is_header: bool) {
        let part_name = {
            let rels = self.package.get_part_rels(&self.doc_part_name);
            rels.and_then(|r| r.get_by_id(rel_id))
                .map(|rel| OpcPackage::resolve_rel_target(&self.doc_part_name, &rel.target))
        };
        if let Some(part_name) = part_name {
            let xml = if is_header {
                hf.to_xml_header()
            } else {
                hf.to_xml_footer()
            };
            if let Ok(xml) = xml {
                self.package.set_part(&part_name, xml);
            }
        }
    }

    // ---- Document Intelligence API ----

    /// Get all headings in the document as (level, text) pairs.
    ///
    /// Detects heading paragraphs by their style ID (e.g. "Heading1", "Heading2").
    pub fn headings(&self) -> Vec<(u32, String)> {
        let mut result = Vec::new();
        for content in &self.document.body.content {
            if let BodyContent::Paragraph(p) = content
                && let Some(level) = Self::detect_heading_level_for_toc(p)
            {
                result.push((level, p.text()));
            }
        }
        result
    }

    /// Get a hierarchical outline of the document headings.
    ///
    /// Returns a tree structure where each node contains the heading level,
    /// text, and children (sub-headings).
    pub fn document_outline(&self) -> Vec<OutlineNode> {
        let headings = self.headings();
        build_outline_tree(&headings)
    }

    /// Get information about all images in the document.
    ///
    /// Returns metadata for each inline and anchored image found in body paragraphs.
    pub fn images(&self) -> Vec<ImageInfo> {
        let mut result = Vec::new();

        for content in &self.document.body.content {
            Self::collect_images_from_content(content, &mut result);
        }
        result
    }

    fn collect_images_from_content(content: &BodyContent, result: &mut Vec<ImageInfo>) {
        match content {
            BodyContent::Paragraph(p) => {
                for run in &p.runs {
                    for rc in &run.content {
                        if let RunContent::Drawing(drawing) = rc {
                            if let Some(inline) = &drawing.inline {
                                result.push(ImageInfo {
                                    embed_id: inline.embed_id.clone(),
                                    name: inline.name.clone(),
                                    description: inline.description.clone(),
                                    width_emu: inline.extent_cx.0,
                                    height_emu: inline.extent_cy.0,
                                    is_anchor: false,
                                });
                            }
                            if let Some(anchor) = &drawing.anchor {
                                result.push(ImageInfo {
                                    embed_id: anchor.embed_id.clone(),
                                    name: anchor.name.clone(),
                                    description: anchor.description.clone(),
                                    width_emu: anchor.extent_cx.0,
                                    height_emu: anchor.extent_cy.0,
                                    is_anchor: true,
                                });
                            }
                        }
                    }
                }
            }
            BodyContent::Table(tbl) => {
                for row in &tbl.rows {
                    for cell in &row.cells {
                        for cc in &cell.content {
                            match cc {
                                rdocx_oxml::table::CellContent::Paragraph(p) => {
                                    Self::collect_images_from_content(
                                        &BodyContent::Paragraph(p.clone()),
                                        result,
                                    );
                                }
                                rdocx_oxml::table::CellContent::Table(nested) => {
                                    Self::collect_images_from_content(
                                        &BodyContent::Table(nested.clone()),
                                        result,
                                    );
                                }
                            }
                        }
                    }
                }
            }
            BodyContent::RawXml(_) => {}
        }
    }

    /// Get information about all hyperlinks in the document.
    ///
    /// Resolves hyperlink relationship IDs to their target URLs where possible.
    pub fn links(&self) -> Vec<LinkInfo> {
        use rdocx_opc::relationship::rel_types;

        // Build a map of hyperlink rel_id -> target URL
        let mut url_map = std::collections::HashMap::new();
        if let Some(rels) = self.package.get_part_rels(&self.doc_part_name) {
            for rel in &rels.items {
                if rel.rel_type == rel_types::HYPERLINK
                    && rel.target_mode.as_ref().is_some_and(|m| m == "External")
                {
                    url_map.insert(rel.id.clone(), rel.target.clone());
                }
            }
        }

        let mut result = Vec::new();
        for content in &self.document.body.content {
            if let BodyContent::Paragraph(p) = content {
                for hl in &p.hyperlinks {
                    let text: String = p.runs[hl.run_start..hl.run_end]
                        .iter()
                        .map(|r| r.text())
                        .collect::<Vec<_>>()
                        .join("");

                    let url = hl.rel_id.as_ref().and_then(|id| url_map.get(id)).cloned();

                    result.push(LinkInfo {
                        text,
                        url,
                        anchor: hl.anchor.clone(),
                        rel_id: hl.rel_id.clone(),
                    });
                }
            }
        }
        result
    }

    /// Count the number of words in the document.
    ///
    /// Counts whitespace-separated tokens across all paragraphs (including
    /// paragraphs inside table cells).
    pub fn word_count(&self) -> usize {
        let mut count = 0;
        for content in &self.document.body.content {
            count += Self::word_count_in_content(content);
        }
        count
    }

    fn word_count_in_content(content: &BodyContent) -> usize {
        match content {
            BodyContent::Paragraph(p) => p.text().split_whitespace().count(),
            BodyContent::Table(tbl) => {
                let mut count = 0;
                for row in &tbl.rows {
                    for cell in &row.cells {
                        for cc in &cell.content {
                            match cc {
                                rdocx_oxml::table::CellContent::Paragraph(p) => {
                                    count += p.text().split_whitespace().count();
                                }
                                rdocx_oxml::table::CellContent::Table(nested) => {
                                    count += Self::word_count_in_content(&BodyContent::Table(
                                        nested.clone(),
                                    ));
                                }
                            }
                        }
                    }
                }
                count
            }
            BodyContent::RawXml(_) => 0,
        }
    }

    /// Audit the document for accessibility issues.
    ///
    /// Checks for common problems: missing image alt text, heading level gaps,
    /// empty paragraphs, missing document metadata.
    pub fn audit_accessibility(&self) -> Vec<AccessibilityIssue> {
        let mut issues = Vec::new();

        // Check: missing document title
        if self.title().is_none() {
            issues.push(AccessibilityIssue {
                severity: IssueSeverity::Warning,
                message: "Document has no title".to_string(),
            });
        }

        // Check: missing document language (author as a proxy for basic metadata)
        if self.author().is_none() {
            issues.push(AccessibilityIssue {
                severity: IssueSeverity::Info,
                message: "Document has no author".to_string(),
            });
        }

        // Check: images without alt text
        let images = self.images();
        for img in &images {
            let has_alt = img
                .description
                .as_ref()
                .is_some_and(|d| !d.is_empty() && d != "Background");
            if !has_alt {
                let name = img
                    .name
                    .as_deref()
                    .or(Some(&img.embed_id))
                    .unwrap_or("unknown");
                issues.push(AccessibilityIssue {
                    severity: IssueSeverity::Error,
                    message: format!("Image \"{name}\" has no alt text"),
                });
            }
        }

        // Check: heading level gaps
        let headings = self.headings();
        let mut prev_level: Option<u32> = None;
        for (level, text) in &headings {
            if let Some(prev) = prev_level
                && *level > prev + 1
            {
                issues.push(AccessibilityIssue {
                    severity: IssueSeverity::Warning,
                    message: format!(
                        "Heading level gap: h{prev} -> h{level} (\"{}\")",
                        truncate_str(text, 40)
                    ),
                });
            }
            prev_level = Some(*level);
        }

        // Check: excessive empty paragraphs
        let mut consecutive_empty = 0u32;
        for content in &self.document.body.content {
            if let BodyContent::Paragraph(p) = content {
                if p.text().trim().is_empty() {
                    consecutive_empty += 1;
                    if consecutive_empty >= 3 {
                        issues.push(AccessibilityIssue {
                            severity: IssueSeverity::Info,
                            message: format!(
                                "{consecutive_empty} consecutive empty paragraphs (consider using spacing instead)"
                            ),
                        });
                    }
                } else {
                    consecutive_empty = 0;
                }
            } else {
                consecutive_empty = 0;
            }
        }

        issues
    }
}

impl Default for Document {
    fn default() -> Self {
        Self::new()
    }
}

/// Guess image content type from the part name extension.
fn guess_image_content_type(part_name: &str) -> String {
    let ext = part_name.rsplit('.').next().unwrap_or("").to_lowercase();
    match ext.as_str() {
        "png" => "image/png",
        "jpg" | "jpeg" => "image/jpeg",
        "gif" => "image/gif",
        "bmp" => "image/bmp",
        "tiff" | "tif" => "image/tiff",
        _ => "image/png",
    }
    .to_string()
}

/// A node in the document outline tree.
#[derive(Debug, Clone, PartialEq)]
pub struct OutlineNode {
    /// The heading level (1-9).
    pub level: u32,
    /// The heading text.
    pub text: String,
    /// Child headings (sub-headings).
    pub children: Vec<OutlineNode>,
}

/// Information about an image in the document.
#[derive(Debug, Clone, PartialEq)]
pub struct ImageInfo {
    /// The relationship ID for the embedded image.
    pub embed_id: String,
    /// Optional name attribute.
    pub name: Option<String>,
    /// Optional description (alt text).
    pub description: Option<String>,
    /// Width in EMUs (English Metric Units, 914400 EMU = 1 inch).
    pub width_emu: i64,
    /// Height in EMUs.
    pub height_emu: i64,
    /// Whether this is an anchored (floating) image vs inline.
    pub is_anchor: bool,
}

/// Information about a hyperlink in the document.
#[derive(Debug, Clone, PartialEq)]
pub struct LinkInfo {
    /// The display text of the hyperlink.
    pub text: String,
    /// The resolved target URL (if external).
    pub url: Option<String>,
    /// Internal document anchor (if any).
    pub anchor: Option<String>,
    /// The relationship ID.
    pub rel_id: Option<String>,
}

/// Severity level for accessibility issues.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum IssueSeverity {
    /// Informational suggestion.
    Info,
    /// Potential problem.
    Warning,
    /// Definite accessibility barrier.
    Error,
}

/// An accessibility issue found during audit.
#[derive(Debug, Clone, PartialEq)]
pub struct AccessibilityIssue {
    /// How severe the issue is.
    pub severity: IssueSeverity,
    /// Human-readable description of the issue.
    pub message: String,
}

/// Build a hierarchical outline tree from a flat list of (level, text) headings.
fn build_outline_tree(headings: &[(u32, String)]) -> Vec<OutlineNode> {
    let mut root: Vec<OutlineNode> = Vec::new();
    let mut stack: Vec<(u32, usize)> = Vec::new(); // (level, index in parent's children)

    for (level, text) in headings {
        let node = OutlineNode {
            level: *level,
            text: text.clone(),
            children: Vec::new(),
        };

        // Pop stack until we find a parent with a lower level
        while let Some(&(stack_level, _)) = stack.last() {
            if stack_level >= *level {
                stack.pop();
            } else {
                break;
            }
        }

        if stack.is_empty() {
            root.push(node);
            let idx = root.len() - 1;
            stack.push((*level, idx));
        } else {
            // Navigate to the correct parent in the tree
            let target = get_outline_parent_mut(&mut root, &stack);
            target.children.push(node);
            let idx = target.children.len() - 1;
            stack.push((*level, idx));
        }
    }

    root
}

/// Navigate to the parent node indicated by the stack.
fn get_outline_parent_mut<'a>(
    root: &'a mut [OutlineNode],
    stack: &[(u32, usize)],
) -> &'a mut OutlineNode {
    let mut current = &mut root[stack[0].1];
    for &(_, idx) in &stack[1..] {
        current = &mut current.children[idx];
    }
    current
}

/// Truncate a string to a maximum length, appending "..." if truncated.
fn truncate_str(s: &str, max_len: usize) -> String {
    if s.len() <= max_len {
        s.to_string()
    } else {
        let truncated: String = s.chars().take(max_len.saturating_sub(3)).collect();
        format!("{truncated}...")
    }
}

/// Deobfuscate an ODTTF (obfuscated TrueType) font file.
///
/// Word embeds fonts as `.odttf` files where the first 32 bytes are XOR'd
/// with a 16-byte GUID derived from the file name. The file name follows
/// the pattern `{GUID}.odttf` where GUID is a hex string without hyphens.
fn deobfuscate_odttf(data: &[u8], file_name: &str) -> Option<Vec<u8>> {
    if data.len() < 32 {
        return None;
    }

    // Extract GUID from file name: "00112233-4455-6677-8899-AABBCCDDEEFF.odttf"
    // or "{00112233-4455-6677-8899-AABBCCDDEEFF}.odttf"
    let name = file_name
        .split('.')
        .next()
        .unwrap_or("")
        .trim_start_matches('{')
        .trim_end_matches('}');

    // Remove hyphens and parse as hex bytes
    let hex: String = name.chars().filter(|c| c.is_ascii_hexdigit()).collect();
    if hex.len() != 32 {
        return None;
    }

    let mut guid = [0u8; 16];
    for (i, byte) in guid.iter_mut().enumerate() {
        *byte = u8::from_str_radix(&hex[i * 2..i * 2 + 2], 16).ok()?;
    }

    // Per OOXML spec, the GUID bytes are reordered for XOR key:
    // bytes 0-3 reversed, 4-5 reversed, 6-7 reversed, 8-15 as-is
    let key: [u8; 16] = [
        guid[3], guid[2], guid[1], guid[0], guid[5], guid[4], guid[7], guid[6], guid[8], guid[9],
        guid[10], guid[11], guid[12], guid[13], guid[14], guid[15],
    ];

    let mut result = data.to_vec();
    // XOR first 32 bytes with the 16-byte key (repeated twice)
    for i in 0..32 {
        result[i] ^= key[i % 16];
    }

    Some(result)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::paragraph::Alignment;
    use rdocx_oxml::units::{HalfPoint, Twips};

    #[test]
    fn create_new_document() {
        let doc = Document::new();
        assert_eq!(doc.paragraph_count(), 0);
        assert!(doc.section_properties().is_some());
    }

    #[test]
    fn add_paragraphs() {
        let mut doc = Document::new();
        doc.add_paragraph("First paragraph");
        doc.add_paragraph("Second paragraph");
        assert_eq!(doc.paragraph_count(), 2);

        let paras = doc.paragraphs();
        assert_eq!(paras[0].text(), "First paragraph");
        assert_eq!(paras[1].text(), "Second paragraph");
    }

    #[test]
    fn paragraph_formatting() {
        let mut doc = Document::new();
        doc.add_paragraph("Centered").alignment(Alignment::Center);

        let paras = doc.paragraphs();
        assert_eq!(paras[0].alignment(), Some(Alignment::Center));
    }

    #[test]
    fn run_formatting() {
        let mut doc = Document::new();
        let mut para = doc.add_paragraph("");
        para.add_run("Bold text").bold(true).size(14.0);

        let paras = doc.paragraphs();
        let runs: Vec<_> = paras[0].runs().collect();
        assert!(runs[0].is_bold());
        assert_eq!(runs[0].size(), Some(14.0));
    }

    #[test]
    fn round_trip_in_memory() {
        let mut doc = Document::new();
        doc.add_paragraph("Hello, World!");
        doc.add_paragraph("Second paragraph")
            .alignment(Alignment::Center);

        let bytes = doc.to_bytes().unwrap();
        let doc2 = Document::from_bytes(&bytes).unwrap();

        assert_eq!(doc2.paragraph_count(), 2);
        let paras = doc2.paragraphs();
        assert_eq!(paras[0].text(), "Hello, World!");
        assert_eq!(paras[1].text(), "Second paragraph");
        assert_eq!(paras[1].alignment(), Some(Alignment::Center));
    }

    #[test]
    fn styles_present() {
        let doc = Document::new();
        assert!(doc.style("Normal").is_some());
        assert!(doc.style("Heading1").is_some());
    }

    #[test]
    fn paragraph_with_style() {
        let mut doc = Document::new();
        doc.add_paragraph("Title").style("Heading1");

        let paras = doc.paragraphs();
        assert_eq!(paras[0].style_id(), Some("Heading1"));
    }

    #[test]
    fn multiple_runs_in_paragraph() {
        let mut doc = Document::new();
        let mut para = doc.add_paragraph("");
        para.add_run("Normal ");
        para.add_run("bold ").bold(true);
        para.add_run("italic").italic(true);

        let paras = doc.paragraphs();
        assert_eq!(paras[0].text(), "Normal bold italic");
        let runs: Vec<_> = paras[0].runs().collect();
        assert_eq!(runs.len(), 3);
        assert!(!runs[0].is_bold());
        assert!(runs[1].is_bold());
        assert!(runs[2].is_italic());
    }

    #[test]
    fn add_custom_style() {
        let mut doc = Document::new();
        doc.add_style(StyleBuilder::paragraph("MyCustom", "My Custom Style").based_on("Normal"));
        assert!(doc.style("MyCustom").is_some());
        let s = doc.style("MyCustom").unwrap();
        assert_eq!(s.name(), Some("My Custom Style"));
        assert_eq!(s.based_on(), Some("Normal"));
    }

    #[test]
    fn resolve_style_properties() {
        let doc = Document::new();
        // Heading1 should inherit from docDefaults and have its own overrides
        let ppr = doc.resolve_paragraph_properties(Some("Heading1"));
        assert_eq!(ppr.keep_next, Some(true));
        assert_eq!(ppr.space_before, Some(Twips(240)));

        // Default (None) should apply Normal style
        let ppr = doc.resolve_paragraph_properties(None);
        assert_eq!(ppr.space_after, Some(Twips(160)));
    }

    #[test]
    fn resolve_run_style_properties() {
        let doc = Document::new();
        let rpr = doc.resolve_run_properties(Some("Heading1"), None);
        assert_eq!(rpr.bold, Some(true));
        assert_eq!(rpr.sz, Some(HalfPoint(32)));
        assert_eq!(rpr.font_ascii, Some("Calibri".to_string()));
    }

    #[test]
    fn set_landscape() {
        let mut doc = Document::new();
        doc.set_landscape();
        let sect = doc.section_properties().unwrap();
        assert_eq!(sect.orientation, Some(ST_PageOrientation::Landscape));
        // Width should be > height in landscape
        assert!(sect.page_width.unwrap().0 > sect.page_height.unwrap().0);
    }

    #[test]
    fn set_margins() {
        let mut doc = Document::new();
        doc.set_margins(
            Length::inches(0.5),
            Length::inches(0.75),
            Length::inches(0.5),
            Length::inches(0.75),
        );
        let sect = doc.section_properties().unwrap();
        assert_eq!(sect.margin_top, Some(Twips(720)));
        assert_eq!(sect.margin_right, Some(Twips(1080)));
    }

    #[test]
    fn set_columns() {
        let mut doc = Document::new();
        doc.set_columns(2, Length::inches(0.5));
        let sect = doc.section_properties().unwrap();
        let cols = sect.columns.as_ref().unwrap();
        assert_eq!(cols.num, Some(2));
        assert_eq!(cols.space, Some(Twips(720)));
        assert_eq!(cols.equal_width, Some(true));
    }

    #[test]
    fn set_page_size() {
        let mut doc = Document::new();
        doc.set_page_size(Length::cm(21.0), Length::cm(29.7));
        let sect = doc.section_properties().unwrap();
        // A4: ~11906tw x ~16838tw
        let w = sect.page_width.unwrap().0;
        let h = sect.page_height.unwrap().0;
        assert!((w - 11906).abs() < 5);
        assert!((h - 16838).abs() < 5);
    }

    #[test]
    fn set_different_first_page() {
        let mut doc = Document::new();
        doc.set_different_first_page(true);
        assert_eq!(doc.section_properties().unwrap().title_pg, Some(true));
    }

    #[test]
    fn content_insertion_api() {
        let mut doc = Document::new();
        doc.add_paragraph("First");
        doc.add_paragraph("Third");

        // Insert in middle
        doc.insert_paragraph(1, "Second");
        assert_eq!(doc.content_count(), 3);
        let paras = doc.paragraphs();
        assert_eq!(paras[0].text(), "First");
        assert_eq!(paras[1].text(), "Second");
        assert_eq!(paras[2].text(), "Third");

        // Insert at beginning
        doc.insert_paragraph(0, "Zeroth");
        assert_eq!(doc.content_count(), 4);
        assert_eq!(doc.paragraphs()[0].text(), "Zeroth");
    }

    #[test]
    fn find_content_index_and_remove() {
        let mut doc = Document::new();
        doc.add_paragraph("Hello");
        doc.add_paragraph("{{PLACEHOLDER}}");
        doc.add_paragraph("World");

        assert_eq!(doc.find_content_index("{{PLACEHOLDER}}"), Some(1));
        assert_eq!(doc.find_content_index("NONEXISTENT"), None);

        assert!(doc.remove_content(1));
        assert_eq!(doc.content_count(), 2);
        assert_eq!(doc.paragraphs()[1].text(), "World");

        // Out of bounds
        assert!(!doc.remove_content(10));
    }

    #[test]
    fn insert_table_at_index() {
        let mut doc = Document::new();
        doc.add_paragraph("Before");
        doc.add_paragraph("After");

        doc.insert_table(1, 2, 3);
        assert_eq!(doc.content_count(), 3);
        assert_eq!(doc.table_count(), 1);
        // Paragraphs are still in correct order
        let paras = doc.paragraphs();
        assert_eq!(paras[0].text(), "Before");
        assert_eq!(paras[1].text(), "After");
    }

    #[test]
    fn replace_text_in_body() {
        let mut doc = Document::new();
        doc.add_paragraph("Hello {{name}}!");
        doc.add_paragraph("Welcome to {{company}}.");

        let count = doc.replace_text("{{name}}", "Alice");
        assert_eq!(count, 1);
        assert_eq!(doc.paragraphs()[0].text(), "Hello Alice!");

        let count = doc.replace_text("{{company}}", "Acme");
        assert_eq!(count, 1);
        assert_eq!(doc.paragraphs()[1].text(), "Welcome to Acme.");
    }

    #[test]
    fn replace_text_in_header_and_footer() {
        let mut doc = Document::new();
        doc.set_header("Header: {{title}}");
        doc.set_footer("Footer: {{title}}");
        doc.add_paragraph("Body: {{title}}");

        let count = doc.replace_text("{{title}}", "My Doc");
        assert_eq!(count, 3);

        assert_eq!(doc.paragraphs()[0].text(), "Body: My Doc");
        assert_eq!(doc.header_text().unwrap(), "Header: My Doc");
        assert_eq!(doc.footer_text().unwrap(), "Footer: My Doc");
    }

    #[test]
    fn replace_all_batch() {
        let mut doc = Document::new();
        doc.add_paragraph("{{a}} and {{b}}");

        let mut map = std::collections::HashMap::new();
        map.insert("{{a}}", "X");
        map.insert("{{b}}", "Y");
        let count = doc.replace_all(&map);
        assert_eq!(count, 2);
        assert_eq!(doc.paragraphs()[0].text(), "X and Y");
    }

    #[test]
    fn template_workflow_round_trip() {
        let mut doc = Document::new();
        doc.add_paragraph("Company: {{company}}");
        doc.add_paragraph("Date: {{date}}");

        doc.replace_text("{{company}}", "Acme Corp");
        doc.replace_text("{{date}}", "2026-02-22");

        // Round-trip
        let bytes = doc.to_bytes().unwrap();
        let doc2 = Document::from_bytes(&bytes).unwrap();
        assert_eq!(doc2.paragraphs()[0].text(), "Company: Acme Corp");
        assert_eq!(doc2.paragraphs()[1].text(), "Date: 2026-02-22");
    }

    #[test]
    fn add_background_image_round_trip() {
        // Create a minimal 1x1 PNG
        let png_data: Vec<u8> = vec![
            0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, // PNG signature
            0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52, // IHDR chunk
            0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, // 1x1
            0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53, 0xde, 0x00, 0x00, 0x00, 0x0c, 0x49,
            0x44, 0x41, 0x54, // IDAT chunk
            0x08, 0xd7, 0x63, 0xf8, 0xcf, 0xc0, 0x00, 0x00, 0x00, 0x02, 0x00, 0x01, 0xe2, 0x21,
            0xbc, 0x33, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, // IEND chunk
            0xae, 0x42, 0x60, 0x82,
        ];

        let mut doc = Document::new();
        doc.add_paragraph("Hello World");
        doc.add_background_image(&png_data, "bg.png");

        // Background image paragraph should be at index 0
        assert_eq!(doc.content_count(), 2);

        // Round-trip
        let bytes = doc.to_bytes().unwrap();
        let doc2 = Document::from_bytes(&bytes).unwrap();

        // Should still have 2 content items
        assert_eq!(doc2.content_count(), 2);
        // The second paragraph should have our text
        assert_eq!(doc2.paragraphs().last().unwrap().text(), "Hello World");
    }

    #[test]
    fn add_anchored_image() {
        let png_data: Vec<u8> = vec![
            0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48,
            0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00,
            0x00, 0x90, 0x77, 0x53, 0xde, 0x00, 0x00, 0x00, 0x0c, 0x49, 0x44, 0x41, 0x54, 0x08,
            0xd7, 0x63, 0xf8, 0xcf, 0xc0, 0x00, 0x00, 0x00, 0x02, 0x00, 0x01, 0xe2, 0x21, 0xbc,
            0x33, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
        ];

        let mut doc = Document::new();
        doc.add_paragraph("Content");
        doc.add_anchored_image(
            &png_data,
            "overlay.png",
            Length::inches(4.0),
            Length::inches(3.0),
            false,
        );
        assert_eq!(doc.content_count(), 2);
    }

    #[test]
    fn insert_toc_basic() {
        let mut doc = Document::new();
        doc.add_paragraph("Introduction");
        doc.add_paragraph("Chapter 1").style("Heading1");
        doc.add_paragraph("Some text in chapter 1.");
        doc.add_paragraph("Section 1.1").style("Heading2");
        doc.add_paragraph("Text in section 1.1.");
        doc.add_paragraph("Chapter 2").style("Heading1");
        doc.add_paragraph("Text in chapter 2.");

        // Before TOC: 7 content elements
        assert_eq!(doc.content_count(), 7);

        // Insert TOC at index 0 with max_level 2
        doc.insert_toc(0, 2);

        // TOC adds: 1 title + 3 heading entries (Ch1, Sec1.1, Ch2) = 4 paragraphs
        assert_eq!(doc.content_count(), 11);

        // Verify TOC title
        let paras = doc.paragraphs();
        assert_eq!(paras[0].text(), "Table of Contents");

        // Verify TOC entries contain heading text
        assert_eq!(paras[1].text(), "Chapter 1\t");
        assert_eq!(paras[2].text(), "Section 1.1\t");
        assert_eq!(paras[3].text(), "Chapter 2\t");

        // Verify round-trip: save and re-open
        let bytes = doc.to_bytes().expect("should serialize");
        let doc2 = Document::from_bytes(&bytes).expect("should open");
        assert_eq!(doc2.content_count(), 11);
        let paras2 = doc2.paragraphs();
        assert_eq!(paras2[0].text(), "Table of Contents");
    }

    #[test]
    fn append_documents() {
        let mut doc_a = Document::new();
        doc_a.add_paragraph("Paragraph A1");
        doc_a.add_paragraph("Paragraph A2");

        let mut doc_b = Document::new();
        doc_b.add_paragraph("Paragraph B1");
        doc_b.add_paragraph("Paragraph B2");
        doc_b.add_paragraph("Paragraph B3");

        assert_eq!(doc_a.content_count(), 2);
        doc_a.append(&doc_b);
        assert_eq!(doc_a.content_count(), 5);

        let paras = doc_a.paragraphs();
        assert_eq!(paras[0].text(), "Paragraph A1");
        assert_eq!(paras[1].text(), "Paragraph A2");
        assert_eq!(paras[2].text(), "Paragraph B1");
        assert_eq!(paras[3].text(), "Paragraph B2");
        assert_eq!(paras[4].text(), "Paragraph B3");

        // Verify round-trip
        let bytes = doc_a.to_bytes().expect("serialize");
        let reopened = Document::from_bytes(&bytes).expect("open");
        assert_eq!(reopened.content_count(), 5);
    }

    #[test]
    fn append_with_section_break() {
        let mut doc_a = Document::new();
        doc_a.add_paragraph("A1");

        let mut doc_b = Document::new();
        doc_b.add_paragraph("B1");

        doc_a.append_with_break(&doc_b, crate::SectionBreak::Continuous);
        // 1 original + 1 section break paragraph + 1 merged = 3
        assert_eq!(doc_a.content_count(), 3);
    }

    #[test]
    fn insert_document_at_index() {
        let mut doc_a = Document::new();
        doc_a.add_paragraph("First");
        doc_a.add_paragraph("Last");

        let mut doc_b = Document::new();
        doc_b.add_paragraph("Middle 1");
        doc_b.add_paragraph("Middle 2");

        doc_a.insert_document(1, &doc_b);
        assert_eq!(doc_a.content_count(), 4);

        let paras = doc_a.paragraphs();
        assert_eq!(paras[0].text(), "First");
        assert_eq!(paras[1].text(), "Middle 1");
        assert_eq!(paras[2].text(), "Middle 2");
        assert_eq!(paras[3].text(), "Last");
    }

    #[test]
    fn merge_deduplicates_styles() {
        let mut doc_a = Document::new();
        doc_a.add_paragraph("A").style("Heading1");

        let mut doc_b = Document::new();
        doc_b.add_paragraph("B").style("Heading1");
        doc_b.add_style(
            crate::style::StyleBuilder::paragraph("CustomB", "Custom B").based_on("Normal"),
        );
        doc_b.add_paragraph("C").style("CustomB");

        let styles_before = doc_a.styles.styles.len();
        doc_a.append(&doc_b);
        let styles_after = doc_a.styles.styles.len();

        // Heading1 already existed, so only CustomB should be added
        assert_eq!(styles_after, styles_before + 1);
    }

    #[test]
    fn headings_and_outline() {
        let mut doc = Document::new();
        doc.add_paragraph("Intro");
        doc.add_paragraph("Chapter 1").style("Heading1");
        doc.add_paragraph("Section 1.1").style("Heading2");
        doc.add_paragraph("Section 1.2").style("Heading2");
        doc.add_paragraph("Chapter 2").style("Heading1");
        doc.add_paragraph("Section 2.1").style("Heading2");
        doc.add_paragraph("Sub 2.1.1").style("Heading3");

        let headings = doc.headings();
        assert_eq!(headings.len(), 6);
        assert_eq!(headings[0], (1, "Chapter 1".to_string()));
        assert_eq!(headings[1], (2, "Section 1.1".to_string()));
        assert_eq!(headings[5], (3, "Sub 2.1.1".to_string()));

        let outline = doc.document_outline();
        assert_eq!(outline.len(), 2); // Two h1 nodes
        assert_eq!(outline[0].text, "Chapter 1");
        assert_eq!(outline[0].children.len(), 2); // 1.1 and 1.2
        assert_eq!(outline[1].text, "Chapter 2");
        assert_eq!(outline[1].children.len(), 1); // 2.1
        assert_eq!(outline[1].children[0].children.len(), 1); // 2.1.1
    }

    #[test]
    fn word_count_basic() {
        let mut doc = Document::new();
        doc.add_paragraph("Hello world");
        doc.add_paragraph("Three more words");
        assert_eq!(doc.word_count(), 5);
    }

    #[test]
    fn audit_accessibility_missing_metadata() {
        let doc = Document::new();
        let issues = doc.audit_accessibility();
        // New document has no title or author
        assert!(issues.iter().any(|i| i.message.contains("no title")));
        assert!(issues.iter().any(|i| i.message.contains("no author")));
    }

    #[test]
    fn audit_heading_level_gap() {
        let mut doc = Document::new();
        doc.set_title("Test");
        doc.set_author("Test");
        doc.add_paragraph("Ch 1").style("Heading1");
        doc.add_paragraph("Skip to 3").style("Heading3");

        let issues = doc.audit_accessibility();
        assert!(
            issues
                .iter()
                .any(|i| i.message.contains("Heading level gap"))
        );
    }

    #[test]
    fn links_returns_empty_for_no_hyperlinks() {
        let mut doc = Document::new();
        doc.add_paragraph("No links here.");
        assert!(doc.links().is_empty());
    }

    #[test]
    fn images_returns_empty_for_text_only() {
        let mut doc = Document::new();
        doc.add_paragraph("Just text.");
        assert!(doc.images().is_empty());
    }
}
