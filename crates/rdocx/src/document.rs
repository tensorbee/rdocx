//! The main Document type — entry point for the rdocx API.

use std::path::Path;
use std::sync::{Arc, Mutex};

#[cfg(test)]
use std::cell::Cell;

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
    /// Package part containing the core properties, resolved from `_rels/.rels`.
    core_properties_part_name: String,
    /// Part name for the main document
    doc_part_name: String,
    /// Part name the styles were loaded from, and where they are written back.
    /// Resolved through the relationship rather than assumed, so a document
    /// that keeps its styles somewhere other than `/word/styles.xml` is
    /// updated in place instead of gaining an orphaned second part.
    styles_part_name: String,
    /// Part name for numbering definitions, resolved the same way.
    numbering_part_name: String,
    /// Greatest numeric suffix among existing image media parts.
    image_counter: usize,
    /// Footnotes: loaded from word/footnotes.xml on open, written back on save.
    footnotes: rdocx_oxml::footnotes::CT_Footnotes,
    /// Normal layout, including system font discovery, computed on first use.
    layout_cache: Mutex<Option<Arc<rdocx_layout::LayoutResult>>>,
    /// Bundled-font-only layout used by deterministic rendering.
    deterministic_layout_cache: Mutex<Option<Arc<rdocx_layout::LayoutResult>>>,
}

/// Fallback part names used when a document does not already declare one.
const DEFAULT_STYLES_PART: &str = "/word/styles.xml";
const DEFAULT_NUMBERING_PART: &str = "/word/numbering.xml";
const DEFAULT_CORE_PROPERTIES_PART: &str = "/docProps/core.xml";
const STYLES_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
const NUMBERING_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";
const CORE_PROPERTIES_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
const CORE_PROPERTIES_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-package.core-properties+xml";

#[cfg(test)]
thread_local! {
    static LAYOUT_INVOCATIONS: Cell<usize> = const { Cell::new(0) };
}

#[cfg(test)]
fn record_layout_invocation() {
    LAYOUT_INVOCATIONS.set(LAYOUT_INVOCATIONS.get() + 1);
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
            core_properties_part_name: DEFAULT_CORE_PROPERTIES_PART.to_string(),
            doc_part_name: "/word/document.xml".to_string(),
            styles_part_name: DEFAULT_STYLES_PART.to_string(),
            numbering_part_name: DEFAULT_NUMBERING_PART.to_string(),
            image_counter: 0,
            footnotes: rdocx_oxml::footnotes::CT_Footnotes::new(),
            layout_cache: Mutex::new(None),
            deterministic_layout_cache: Mutex::new(None),
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

        // Resolve the part a relationship of the given type points at.
        let resolve_part = |rel_type: &str| -> Option<String> {
            let rels = package.get_part_rels(&doc_part_name)?;
            let rel = rels.get_by_type(rel_type)?;
            Some(OpcPackage::resolve_rel_target(&doc_part_name, &rel.target))
        };

        // Try to load styles, remembering where they came from.
        let styles_part_name = resolve_part(rel_types::STYLES);
        let styles = match styles_part_name
            .as_deref()
            .and_then(|p| package.get_part(p))
        {
            Some(styles_xml) => CT_Styles::from_xml(styles_xml)?,
            None => CT_Styles::new_default(),
        };

        // Try to load numbering definitions
        let numbering_part_name = resolve_part(rel_types::NUMBERING);
        let numbering = match numbering_part_name
            .as_deref()
            .and_then(|p| package.get_part(p))
        {
            Some(num_xml) => Some(CT_Numbering::from_xml(num_xml)?),
            None => None,
        };

        // Core properties are a package-level relationship, not a document part.
        let core_properties_part_name = package
            .package_rels
            .get_by_type(CORE_PROPERTIES_REL_TYPE)
            .map(|rel| OpcPackage::resolve_rel_target("/", &rel.target));
        let core_properties = core_properties_part_name
            .as_deref()
            .and_then(|part| package.get_part(part))
            .and_then(|xml| CoreProperties::from_xml(xml).ok());

        let image_counter = package
            .parts
            .keys()
            .filter_map(|name| image_number_from_part_name(name))
            .max()
            .unwrap_or(0);

        let footnotes = package
            .get_part_rels(&doc_part_name)
            .and_then(|rels| rels.get_by_type(rel_types::FOOTNOTES))
            .map(|rel| OpcPackage::resolve_rel_target(&doc_part_name, &rel.target))
            .and_then(|part| package.get_part(&part))
            .and_then(|xml| rdocx_oxml::footnotes::CT_Footnotes::from_xml(xml).ok())
            .unwrap_or_default();

        Ok(Document {
            package,
            document,
            styles,
            numbering,
            core_properties,
            core_properties_part_name: core_properties_part_name
                .unwrap_or_else(|| DEFAULT_CORE_PROPERTIES_PART.to_string()),
            doc_part_name,
            styles_part_name: styles_part_name.unwrap_or_else(|| DEFAULT_STYLES_PART.to_string()),
            numbering_part_name: numbering_part_name
                .unwrap_or_else(|| DEFAULT_NUMBERING_PART.to_string()),
            image_counter,
            footnotes,
            layout_cache: Mutex::new(None),
            deterministic_layout_cache: Mutex::new(None),
        })
    }

    /// Clear layouts derived from the current document state.
    fn invalidate_layout(&mut self) {
        self.layout_cache
            .get_mut()
            .unwrap_or_else(std::sync::PoisonError::into_inner)
            .take();
        self.deterministic_layout_cache
            .get_mut()
            .unwrap_or_else(std::sync::PoisonError::into_inner)
            .take();
    }

    /// Return the normal-font layout, computing it once after each mutation.
    fn cached_layout(&self) -> Result<Arc<rdocx_layout::LayoutResult>> {
        let mut cache = self
            .layout_cache
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner);
        if let Some(layout) = cache.as_ref() {
            return Ok(Arc::clone(layout));
        }

        let input = self.build_layout_input();
        #[cfg(test)]
        record_layout_invocation();
        let layout = Arc::new(rdocx_layout::layout_document(&input)?);
        *cache = Some(Arc::clone(&layout));
        Ok(layout)
    }

    /// Return the bundled-font-only layout, computing it once after mutation.
    fn cached_deterministic_layout(&self) -> Result<Arc<rdocx_layout::LayoutResult>> {
        let mut cache = self
            .deterministic_layout_cache
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner);
        if let Some(layout) = cache.as_ref() {
            return Ok(Arc::clone(layout));
        }

        let input = self.build_layout_input();
        #[cfg(test)]
        record_layout_invocation();
        let layout = Arc::new(rdocx_layout::layout_document_deterministic(&input)?);
        *cache = Some(Arc::clone(&layout));
        Ok(layout)
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

        // Serialize the styles part. A document opened without one still gets
        // rdocx's defaults written out, so make sure it is reachable: an
        // unreferenced, untyped part would simply be ignored by Word.
        let styles_xml = self.styles.to_xml()?;
        let styles_part = self.styles_part_name.clone();
        self.package.set_part(&styles_part, styles_xml);
        self.ensure_part_relationship(&styles_part, rel_types::STYLES, STYLES_CONTENT_TYPE);

        // Serialize numbering definitions if we have any
        if let Some(ref numbering) = self.numbering {
            let numbering_xml = numbering.to_xml()?;
            let numbering_part = self.numbering_part_name.clone();
            self.package.set_part(&numbering_part, numbering_xml);
            self.ensure_part_relationship(
                &numbering_part,
                rel_types::NUMBERING,
                NUMBERING_CONTENT_TYPE,
            );
        }

        // Serialize footnotes.xml when any footnotes exist
        if !self.footnotes.footnotes.is_empty() {
            let fx = self.footnotes.to_xml_footnotes()?;
            self.package.set_part("/word/footnotes.xml", fx);
            self.package.content_types.add_override(
                "/word/footnotes.xml",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
            );
            let rels = self
                .package
                .get_or_create_part_rels(&self.doc_part_name.clone());
            if rels.get_by_type(rel_types::FOOTNOTES).is_none() {
                rels.add(rel_types::FOOTNOTES, "footnotes.xml");
            }
        }

        // Serialize core properties to the package relationship's target.
        if let Some(ref props) = self.core_properties {
            let core_xml = props.to_xml()?;
            self.package
                .set_part(&self.core_properties_part_name, core_xml);
            self.package.content_types.add_override(
                &self.core_properties_part_name,
                CORE_PROPERTIES_CONTENT_TYPE,
            );
            if self
                .package
                .package_rels
                .get_by_type(CORE_PROPERTIES_REL_TYPE)
                .is_none()
            {
                let target = self
                    .core_properties_part_name
                    .strip_prefix('/')
                    .unwrap_or(&self.core_properties_part_name);
                self.package
                    .package_rels
                    .add(CORE_PROPERTIES_REL_TYPE, target);
            }
        }

        Ok(())
    }

    /// Make sure `part_name` is reachable from the main document: it needs a
    /// relationship of `rel_type` and a content-type override.
    fn ensure_part_relationship(&mut self, part_name: &str, rel_type: &str, content_type: &str) {
        self.package
            .content_types
            .add_override(part_name, content_type);

        let doc_part_name = self.doc_part_name.clone();
        let already_linked = self
            .package
            .get_part_rels(&doc_part_name)
            .and_then(|rels| rels.get_by_type(rel_type))
            .map(|rel| OpcPackage::resolve_rel_target(&doc_part_name, &rel.target))
            .is_some_and(|target| target == part_name);
        if already_linked {
            return;
        }

        // Relationship targets are relative to the source part's directory.
        let target = relative_target(&doc_part_name, part_name);
        self.package
            .get_or_create_part_rels(&doc_part_name)
            .add(rel_type, &target);
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

    /// All footnotes as (id, plain text), in file order.
    pub fn footnotes(&self) -> Vec<(i32, String)> {
        self.footnotes
            .footnotes
            .iter()
            .map(|f| {
                let text = f
                    .paragraphs
                    .iter()
                    .map(|p| p.text())
                    .collect::<Vec<_>>()
                    .join("\n");
                (f.id, text)
            })
            .collect()
    }

    /// Add a footnote with the given text; returns its id. Pair with
    /// `Paragraph::add_footnote_ref` to reference it from the body.
    pub fn add_footnote(&mut self, text: &str) -> i32 {
        self.invalidate_layout();
        use rdocx_oxml::footnotes::CT_Footnote;
        use rdocx_oxml::text::CT_P;
        let id = self
            .footnotes
            .footnotes
            .iter()
            .map(|f| f.id)
            .max()
            .unwrap_or(1)
            + 1;
        let mut p = CT_P::new();
        p.add_run(text);
        self.footnotes.footnotes.push(CT_Footnote {
            id,
            paragraphs: vec![p],
        });
        id
    }

    /// Add a paragraph with the given text and return a mutable reference.
    pub fn add_paragraph(&mut self, text: &str) -> Paragraph<'_> {
        self.invalidate_layout();
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
        self.invalidate_layout();
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
        self.invalidate_layout();
        use rdocx_oxml::table::{CT_Row, CT_TblGrid, CT_TblGridCol, CT_TblPr, CT_TblWidth, CT_Tc};
        use rdocx_oxml::units::Twips;

        // Default column width: divide 9360tw (6.5" printable at 1" margins) evenly.
        // A zero-column table has no grid to divide; clamp so this cannot divide by zero.
        let col_width = Twips(9360 / cols.max(1) as i32);

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
    /// # Panics
    ///
    /// Panics if `index > content_count()`. (Unlike [`Self::insert_document`]
    /// and [`Self::insert_toc`], which clamp an out-of-range index to the end.)
    pub fn insert_paragraph(&mut self, index: usize, text: &str) -> Paragraph<'_> {
        self.invalidate_layout();
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
    /// A `cols` of 0 produces a table with no columns rather than panicking.
    ///
    /// # Panics
    ///
    /// Panics if `index > content_count()`. (Unlike [`Self::insert_document`]
    /// and [`Self::insert_toc`], which clamp an out-of-range index to the end.)
    pub fn insert_table(&mut self, index: usize, rows: usize, cols: usize) -> Table<'_> {
        self.invalidate_layout();
        use rdocx_oxml::table::{CT_Row, CT_TblGrid, CT_TblGridCol, CT_TblPr, CT_TblWidth, CT_Tc};
        use rdocx_oxml::units::Twips;

        let col_width = Twips(9360 / cols.max(1) as i32);
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
        self.invalidate_layout();
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
        self.invalidate_layout();
        let rel_id = self.embed_image(image_data, image_filename);

        let inline = CT_Inline::new(&rel_id, width.to_emu(), height.to_emu());

        let drawing = CT_Drawing::inline(inline);
        let run = CT_R {
            alt_drawings: Vec::new(),
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
        self.invalidate_layout();
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
            alt_drawings: Vec::new(),
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
        self.invalidate_layout();
        let rel_id = self.embed_image(image_data, image_filename);

        let mut anchor = CT_Anchor::background(&rel_id, width.to_emu(), height.to_emu());
        anchor.behind_doc = behind_text;

        let drawing = CT_Drawing::anchor(anchor);
        let run = CT_R {
            alt_drawings: Vec::new(),
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
        let mut candidate = self.image_counter.checked_add(1).unwrap_or(1);
        while self
            .package
            .parts
            .keys()
            .filter_map(|name| image_number_from_part_name(name))
            .any(|number| number == candidate)
        {
            candidate = candidate.checked_add(1).unwrap_or(1);
        }
        self.image_counter = candidate;
        candidate
    }

    /// Store image bytes as a new media part and declare its content type.
    ///
    /// Returns the relationship target to use when referencing it, e.g.
    /// `media/image3.png`. No relationship is created here: an image referenced
    /// from a header or footer must be related to *that* part, not the
    /// document, so the caller decides where it is attached.
    fn store_image_part(&mut self, image_data: &[u8], filename: &str) -> String {
        let ext = image_extension(filename);
        let image_num = self.next_image_number();

        self.package.set_part(
            &format!("/word/media/image{image_num}.{ext}"),
            image_data.to_vec(),
        );
        self.package
            .content_types
            .add_default(&ext, image_content_type(&ext));

        format!("media/image{image_num}.{ext}")
    }

    /// Embed an image into the OPC package and return the relationship ID.
    ///
    /// Public so callers can pre-embed an image and then pass the returned
    /// `rel_id` to [`crate::Cell::add_picture`] for inline cell images.
    pub fn embed_image(&mut self, image_data: &[u8], filename: &str) -> String {
        self.invalidate_layout();
        let rel_target = self.store_image_part(image_data, filename);
        self.package
            .get_or_create_part_rels(&self.doc_part_name)
            .add(rel_types::IMAGE, &rel_target)
    }

    /// Whether the given numbering definition renders as bullets (true)
    /// or numbers (false). None if the id is unknown.
    pub fn numbering_is_bullet(&self, num_id: u32) -> Option<bool> {
        let numbering = self.numbering.as_ref()?;
        let abstract_num = numbering.get_abstract_num_for(num_id)?;
        let fmt = abstract_num.levels.first()?.num_fmt?;
        Some(fmt == rdocx_oxml::numbering::ST_NumberFormat::Bullet)
    }

    /// Append an external hyperlink to the last paragraph (creating one if
    /// the document is empty): adds the External relationship and wraps the
    /// new run in a hyperlink span.
    pub fn append_hyperlink(&mut self, text: &str, url: &str) {
        self.invalidate_layout();
        use rdocx_opc::relationship::rel_types;

        let rel_id = {
            let rels = self.package.get_or_create_part_rels(&self.doc_part_name);
            rels.add_external(rel_types::HYPERLINK, url)
        };

        if !matches!(
            self.document.body.content.last(),
            Some(BodyContent::Paragraph(_))
        ) {
            self.document
                .body
                .content
                .push(BodyContent::Paragraph(CT_P::new()));
        }
        let Some(BodyContent::Paragraph(p)) = self.document.body.content.last_mut() else {
            unreachable!();
        };
        let run_start = p.runs.len();
        p.add_run(text);
        p.hyperlinks.push(rdocx_oxml::text::HyperlinkSpan {
            rel_id: Some(rel_id),
            anchor: None,
            run_start,
            run_end: run_start + 1,
        });
    }

    /// Get a builder for the last paragraph in the body, if any. Lets
    /// callers interleave plain runs with `append_hyperlink` calls.
    pub fn last_paragraph_mut(&mut self) -> Option<Paragraph<'_>> {
        self.invalidate_layout();
        match self.document.body.content.last_mut() {
            Some(BodyContent::Paragraph(p)) => Some(Paragraph { inner: p }),
            _ => None,
        }
    }

    /// Fetch the raw bytes of an embedded image by its relationship ID.
    pub fn image_data(&self, rel_id: &str) -> Option<Vec<u8>> {
        let rels = self.package.get_part_rels(&self.doc_part_name)?;
        let rel = rels.items.iter().find(|r| r.id == rel_id)?;
        let target = OpcPackage::resolve_rel_target(&self.doc_part_name, &rel.target);
        self.package.get_part(&target).map(|b| b.to_vec())
    }

    /// Resolve a hyperlink relationship ID to its external URL.
    pub fn hyperlink_url(&self, rel_id: &str) -> Option<String> {
        use rdocx_opc::relationship::rel_types;
        let rels = self.package.get_part_rels(&self.doc_part_name)?;
        rels.items
            .iter()
            .find(|r| r.id == rel_id && r.rel_type == rel_types::HYPERLINK)
            .map(|r| r.target.clone())
    }

    // ---- Header/Footer ----

    /// Set the default header text.
    ///
    /// Creates a header part with the given text and references it from
    /// the section properties.
    pub fn set_header(&mut self, text: &str) {
        self.invalidate_layout();
        self.set_header_footer_part(text, true, HdrFtrType::Default);
    }

    /// Set the default footer text.
    pub fn set_footer(&mut self, text: &str) {
        self.invalidate_layout();
        self.set_header_footer_part(text, false, HdrFtrType::Default);
    }

    /// Set the first-page header text.
    pub fn set_first_page_header(&mut self, text: &str) {
        self.invalidate_layout();
        self.set_different_first_page(true);
        self.set_header_footer_part(text, true, HdrFtrType::First);
    }

    /// Set the first-page footer text.
    pub fn set_first_page_footer(&mut self, text: &str) {
        self.invalidate_layout();
        self.set_different_first_page(true);
        self.set_header_footer_part(text, false, HdrFtrType::First);
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
        self.invalidate_layout();
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
        self.invalidate_layout();
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
        self.invalidate_layout();
        self.set_raw_hdr_ftr_with_images(header_xml, images, true, hdr_type);
    }

    /// Set a footer from raw XML bytes with associated images.
    pub fn set_raw_footer_with_images(
        &mut self,
        footer_xml: Vec<u8>,
        images: &[(&str, &[u8], &str)],
        hdr_type: HdrFtrType,
    ) {
        self.invalidate_layout();
        self.set_raw_hdr_ftr_with_images(footer_xml, images, false, hdr_type);
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
        self.invalidate_layout();
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
        self.invalidate_layout();
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

    /// Where a header/footer of this kind lives, and how to declare it.
    ///
    /// All four public entry points differ only in what goes *inside* the part;
    /// the surrounding bookkeeping — part name, content type, relationship,
    /// section reference — is identical, and lives here.
    ///
    /// Note the fixed `1` in the part name: rdocx manages one header and one
    /// footer per [`HdrFtrType`] for the document's single section. Setting a
    /// header of the same type again replaces the existing part.
    fn hdr_ftr_slots(
        is_header: bool,
        hdr_type: HdrFtrType,
    ) -> (String, &'static str, &'static str) {
        let type_suffix = match hdr_type {
            HdrFtrType::Default => "",
            HdrFtrType::First => "First",
            HdrFtrType::Even => "Even",
        };
        if is_header {
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
        }
    }

    /// Install a header/footer part: store its bytes, declare the content type,
    /// relate it to the document, and point the section properties at it.
    ///
    /// Any previous reference of the same [`HdrFtrType`] is replaced.
    fn install_hdr_ftr_part(
        &mut self,
        xml: Vec<u8>,
        is_header: bool,
        hdr_type: HdrFtrType,
    ) -> String {
        let (part_name, rel_type, content_type) = Self::hdr_ftr_slots(is_header, hdr_type);

        self.package.set_part(&part_name, xml);
        self.package
            .content_types
            .add_override(&part_name, content_type);

        // Setting the same header twice must not leave the first relationship
        // behind pointing at the same part.
        let rel_target = relative_target(&self.doc_part_name, &part_name);
        let rels = self.package.get_or_create_part_rels(&self.doc_part_name);
        let rel_id = match rels
            .items
            .iter()
            .find(|r| r.rel_type == rel_type && r.target == rel_target)
        {
            Some(existing) => existing.id.clone(),
            None => rels.add(rel_type, &rel_target),
        };

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

        part_name
    }

    /// Serialize a header/footer body, choosing the right root element.
    fn serialize_hdr_ftr(hdr_ftr: &CT_HdrFtr, is_header: bool) -> Result<Vec<u8>> {
        let xml = if is_header {
            hdr_ftr.to_xml_header()
        } else {
            hdr_ftr.to_xml_footer()
        };
        Ok(xml?)
    }

    fn set_header_footer_part(&mut self, text: &str, is_header: bool, hdr_type: HdrFtrType) {
        let mut hdr_ftr = CT_HdrFtr::new();
        let mut p = CT_P::new();
        if !text.is_empty() {
            p.add_run(text);
        }
        hdr_ftr.paragraphs.push(p);

        let Ok(xml) = Self::serialize_hdr_ftr(&hdr_ftr, is_header) else {
            return;
        };
        self.install_hdr_ftr_part(xml, is_header, hdr_type);
    }

    fn set_raw_hdr_ftr_with_images(
        &mut self,
        xml: Vec<u8>,
        images: &[(&str, &[u8], &str)],
        is_header: bool,
        hdr_type: HdrFtrType,
    ) {
        let part_name = self.install_hdr_ftr_part(xml, is_header, hdr_type);

        // The supplied markup already references these images by ID, so each
        // relationship has to be created with that exact ID.
        for &(rel_id, image_data, image_filename) in images {
            let img_rel_target = self.store_image_part(image_data, image_filename);
            self.package
                .get_or_create_part_rels(&part_name)
                .add_with_id(rel_id, rel_types::IMAGE, &img_rel_target);
        }
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
        self.set_header_footer_image_bg_part(
            image_data,
            image_filename,
            width,
            height,
            None,
            is_header,
            hdr_type,
        );
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
        use rdocx_oxml::properties::CT_Shd;

        let (part_name, _, _) = Self::hdr_ftr_slots(is_header, hdr_type);

        // The image relationship belongs to the header/footer part, not the
        // document, because that is where the drawing referencing it lives.
        let img_rel_target = self.store_image_part(image_data, image_filename);
        let img_rel_id = self
            .package
            .get_or_create_part_rels(&part_name)
            .add(rel_types::IMAGE, &img_rel_target);

        let inline = CT_Inline::new(&img_rel_id, width.to_emu(), height.to_emu());
        let run = CT_R {
            alt_drawings: Vec::new(),
            properties: None,
            content: vec![RunContent::Drawing(CT_Drawing::inline(inline))],
            extra_xml: Vec::new(),
        };

        let mut p = CT_P::new();
        p.runs.push(run);
        if let Some(color) = bg_color {
            p.properties = Some(CT_PPr {
                shading: Some(CT_Shd {
                    val: "clear".to_string(),
                    color: Some("auto".to_string()),
                    fill: Some(color.to_string()),
                }),
                ..Default::default()
            });
        }

        let mut hdr_ftr = CT_HdrFtr::new();
        hdr_ftr.paragraphs.push(p);

        let Ok(xml) = Self::serialize_hdr_ftr(&hdr_ftr, is_header) else {
            return;
        };
        self.install_hdr_ftr_part(xml, is_header, hdr_type);
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

    /// Ensure a numbering part exists.
    ///
    /// The relationship and content-type override are added by
    /// [`Self::flush_to_package`], which knows the resolved part name and will
    /// not create a second numbering relationship if one already exists.
    fn ensure_numbering(&mut self) -> &mut CT_Numbering {
        self.numbering.get_or_insert_with(CT_Numbering::new)
    }

    /// Add a bullet list item at the given indentation level (0-based).
    ///
    /// If no bullet list definition exists yet, one is created automatically.
    /// Returns a mutable `Paragraph` for further configuration.
    pub fn add_bullet_list_item(&mut self, text: &str, level: u32) -> Paragraph<'_> {
        self.invalidate_layout();
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
        self.invalidate_layout();
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
        self.invalidate_layout();
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
        self.invalidate_layout();
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
        self.invalidate_layout();
        self.ensure_core_properties().title = Some(title.to_string());
    }

    /// Get the document author/creator.
    pub fn author(&self) -> Option<&str> {
        self.core_properties.as_ref()?.creator.as_deref()
    }

    /// Set the document author/creator.
    pub fn set_author(&mut self, author: &str) {
        self.invalidate_layout();
        self.ensure_core_properties().creator = Some(author.to_string());
    }

    /// Get the document subject.
    pub fn subject(&self) -> Option<&str> {
        self.core_properties.as_ref()?.subject.as_deref()
    }

    /// Set the document subject.
    pub fn set_subject(&mut self, subject: &str) {
        self.invalidate_layout();
        self.ensure_core_properties().subject = Some(subject.to_string());
    }

    /// Get the document keywords.
    pub fn keywords(&self) -> Option<&str> {
        self.core_properties.as_ref()?.keywords.as_deref()
    }

    /// Set the document keywords.
    pub fn set_keywords(&mut self, keywords: &str) {
        self.invalidate_layout();
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
        self.invalidate_layout();
        self.merge_styles(other);

        let start_idx = self.document.body.content.len();
        for content in &other.document.body.content {
            self.document.body.content.push(content.clone());
        }

        self.remap_merged_numbering(other, start_idx);
    }

    /// Append the content of another document with a section break.
    pub fn append_with_break(&mut self, other: &Document, break_type: crate::SectionBreak) {
        self.invalidate_layout();
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
    ///
    /// An `index` past the end is clamped to the end rather than panicking.
    pub fn insert_document(&mut self, index: usize, other: &Document) {
        self.invalidate_layout();
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
        self.invalidate_layout();
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

        // Calling insert_toc twice must not mint bookmarks that collide with
        // the ones the first call left behind — duplicate `w:name` values make
        // the internal links ambiguous. Continue numbering past whatever is
        // already there.
        let mut toc_counter = self.highest_toc_bookmark();
        let mut bookmark_id = 100 + toc_counter;

        let mut headings = Vec::new();

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

        // Step 3: Build TOC entry paragraphs.
        // The dot leader runs to the right text margin, which depends on the
        // section's page size and margins rather than being a fixed 6.5".
        let right_tab = CT_Tabs {
            tabs: vec![CT_TabStop {
                val: ST_TabJc::Right,
                pos: Twips(self.text_width_twips()),
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
                alt_drawings: Vec::new(),
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

    /// The highest `_TocN` bookmark number already present in the body.
    ///
    /// Returns 0 when there are none, so the next bookmark is `_Toc1`.
    fn highest_toc_bookmark(&self) -> u32 {
        let mut highest = 0;
        for content in &self.document.body.content {
            let BodyContent::Paragraph(p) = content else {
                continue;
            };
            for (_, raw) in &p.extra_xml {
                let Ok(text) = std::str::from_utf8(raw) else {
                    continue;
                };
                for (_, after) in text.match_indices("_Toc") {
                    let digits: String = after
                        .trim_start_matches("_Toc")
                        .chars()
                        .take_while(char::is_ascii_digit)
                        .collect();
                    if let Ok(n) = digits.parse::<u32>() {
                        highest = highest.max(n);
                    }
                }
            }
        }
        highest
    }

    /// Width of the text column in twips: page width less both side margins.
    ///
    /// Falls back to the US Letter default (6.5") when the section does not
    /// specify a size, and never returns a non-positive width.
    fn text_width_twips(&self) -> i32 {
        const DEFAULT_TEXT_WIDTH: i32 = 9360;

        let Some(sect) = self.document.body.sect_pr.as_ref() else {
            return DEFAULT_TEXT_WIDTH;
        };
        let page_width = sect.page_width.map(|w| w.0).unwrap_or(12240);
        let left = sect.margin_left.map(|m| m.0).unwrap_or(1440);
        let right = sect.margin_right.map(|m| m.0).unwrap_or(1440);

        let width = page_width - left - right;
        if width > 0 { width } else { DEFAULT_TEXT_WIDTH }
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
    /// Searches body paragraphs, tables (including nested), headers, footers,
    /// text boxes and chart labels. Handles placeholders split across multiple
    /// runs. Returns the total number of replacements made.
    ///
    /// A `replacement` that contains `placeholder` is substituted once, not
    /// repeatedly.
    pub fn replace_text(&mut self, placeholder: &str, replacement: &str) -> usize {
        self.invalidate_layout();
        self.replace_batch(&[(placeholder, replacement)])
    }

    /// Replace multiple placeholders at once. Returns total replacements.
    ///
    /// Cheaper than calling [`Self::replace_text`] per entry: the document is
    /// serialised and re-parsed once for the whole batch rather than once per
    /// placeholder.
    pub fn replace_all(&mut self, replacements: &std::collections::HashMap<&str, &str>) -> usize {
        self.invalidate_layout();
        let pairs: Vec<(&str, &str)> = replacements.iter().map(|(k, v)| (*k, *v)).collect();
        self.replace_batch(&pairs)
    }

    /// Apply a batch of literal replacements across the whole document.
    fn replace_batch(&mut self, pairs: &[(&str, &str)]) -> usize {
        if pairs.is_empty() {
            return 0;
        }

        let mut count = 0;

        // Typed model: body content, then headers and footers.
        for (placeholder, replacement) in pairs {
            count += self.replace_in_body(placeholder, replacement);
        }
        count += self.replace_in_headers_footers(pairs);

        // Raw XML: text boxes, shapes and charts live in markup the typed model
        // does not cover, so flush first and work on the serialised parts.
        if self.flush_to_package().is_ok() {
            count += self.replace_in_xml_parts(pairs);
        }

        count
    }

    /// Run the typed replacement over body paragraphs and tables.
    fn replace_in_body(&mut self, placeholder: &str, replacement: &str) -> usize {
        use rdocx_oxml::placeholder;

        let mut count = 0;
        for content in &mut self.document.body.content {
            match content {
                BodyContent::Paragraph(p) => {
                    count += placeholder::replace_in_paragraph(p, placeholder, replacement);
                }
                BodyContent::Table(t) => {
                    count += placeholder::replace_in_table(t, placeholder, replacement);
                }
                BodyContent::RawXml(_) => {}
            }
        }
        count
    }

    /// Run the typed replacement over every referenced header and footer part.
    fn replace_in_headers_footers(&mut self, pairs: &[(&str, &str)]) -> usize {
        use rdocx_oxml::placeholder;

        let mut count = 0;
        for (rel_id, is_header) in self.header_footer_rel_ids() {
            let Some(mut hf) = self.load_header_footer(&rel_id) else {
                continue;
            };
            let mut part_count = 0;
            for (placeholder, replacement) in pairs {
                part_count +=
                    placeholder::replace_in_header_footer(&mut hf, placeholder, replacement);
            }
            if part_count > 0 {
                self.save_header_footer(&rel_id, &hf, is_header);
                count += part_count;
            }
        }
        count
    }

    /// Relationship IDs of the section's headers and footers, with a flag
    /// saying which kind each one is.
    fn header_footer_rel_ids(&self) -> Vec<(String, bool)> {
        let Some(sect_pr) = self.document.body.sect_pr.as_ref() else {
            return Vec::new();
        };
        sect_pr
            .header_refs
            .iter()
            .map(|r| (r.rel_id.clone(), true))
            .chain(
                sect_pr
                    .footer_refs
                    .iter()
                    .map(|r| (r.rel_id.clone(), false)),
            )
            .collect()
    }

    // ---- Regex replacement ----

    /// Replace all regex matches with `replacement` throughout the document.
    ///
    /// The `replacement` string supports capture groups: `$1`, `$2`, etc.
    /// Searches body paragraphs, tables (including nested), headers, and footers.
    /// Returns the total number of replacements made, or an error if the regex is invalid.
    pub fn replace_regex(&mut self, pattern: &str, replacement: &str) -> Result<usize> {
        self.invalidate_layout();
        let re =
            regex::Regex::new(pattern).map_err(|e| Error::Other(format!("invalid regex: {e}")))?;
        Ok(self.replace_regex_compiled(&re, replacement))
    }

    /// Replace multiple regex patterns at once. Returns total replacements.
    pub fn replace_all_regex(&mut self, patterns: &[(String, String)]) -> Result<usize> {
        self.invalidate_layout();
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
                BodyContent::RawXml(_) => {}
            }
        }

        // Replace in headers and footers
        for (rel_id, is_header) in self.header_footer_rel_ids() {
            let Some(mut hf) = self.load_header_footer(&rel_id) else {
                continue;
            };
            let n = placeholder::replace_regex_in_header_footer(&mut hf, re, replacement);
            if n > 0 {
                self.save_header_footer(&rel_id, &hf, is_header);
                count += n;
            }
        }

        // Text boxes and shapes live in raw markup the typed model does not
        // reach. `replace_text` has always covered them; do the same here so
        // the two entry points search the same places.
        if self.flush_to_package().is_ok() {
            count += self.replace_regex_in_xml_parts(re, replacement);
        }

        count
    }

    /// Apply a regex replacement to the text-box content of the raw XML parts.
    fn replace_regex_in_xml_parts(&mut self, re: &regex::Regex, replacement: &str) -> usize {
        let mut count = 0;

        for part_name in self.text_bearing_part_names() {
            let Some(xml) = self.package.get_part(&part_name).map(<[u8]>::to_vec) else {
                continue;
            };
            if let Ok((new_xml, n)) =
                rdocx_oxml::placeholder::replace_regex_in_xml_part(&xml, re, replacement)
                && n > 0
            {
                self.package.set_part(&part_name, new_xml);
                count += n;
            }
        }

        // Re-parse so the in-memory model reflects the edited markup; otherwise
        // the next flush would write the pre-replacement document back out.
        if count > 0
            && let Some(doc_xml) = self.package.get_part(&self.doc_part_name)
            && let Ok(doc) = CT_Document::from_xml(doc_xml)
        {
            self.document = doc;
        }

        count
    }

    /// The main document part plus every header and footer part: everywhere
    /// text boxes and shapes with replaceable text can appear.
    fn text_bearing_part_names(&self) -> Vec<String> {
        let mut names = vec![self.doc_part_name.clone()];
        if let Some(rels) = self.package.get_part_rels(&self.doc_part_name) {
            for (rel_id, _) in self.header_footer_rel_ids() {
                if let Some(rel) = rels.get_by_id(&rel_id) {
                    names.push(OpcPackage::resolve_rel_target(
                        &self.doc_part_name,
                        &rel.target,
                    ));
                }
            }
        }
        names
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
    fn replace_in_xml_parts(&mut self, pairs: &[(&str, &str)]) -> usize {
        use rdocx_oxml::placeholder::{replace_many_in_chart_xml, replace_many_in_xml_part};

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
                if let Ok((new_xml, n)) = replace_many_in_xml_part(&xml, pairs)
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
                if let Ok((new_xml, n)) = replace_many_in_chart_xml(&xml, pairs)
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
        let layout = self.cached_layout()?;
        Ok(rdocx_pdf::render_to_pdf(&layout))
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
        #[cfg(test)]
        record_layout_invocation();
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
                    t if t == rel_types::HYPERLINK
                        && rel.target_mode.as_ref().is_some_and(|m| m == "External") =>
                    {
                        hyperlink_urls.insert(rel.id.clone(), rel.target.clone());
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
        let layout = self.cached_layout()?;
        Ok(rdocx_pdf::render_page_to_png(&layout, page_index, dpi))
    }

    /// Render a single page to PNG using bundled fonts without system font
    /// discovery.
    ///
    /// # Arguments
    /// * `page_index` - 0-based page index
    /// * `dpi` - Resolution (72 = 1:1, 150 = standard, 300 = high quality)
    pub fn render_page_to_png_deterministic(
        &self,
        page_index: usize,
        dpi: f64,
    ) -> Result<Option<Vec<u8>>> {
        let layout = self.cached_deterministic_layout()?;
        Ok(rdocx_pdf::render_page_to_png(&layout, page_index, dpi))
    }

    /// Render all pages of the document to PNG bytes.
    pub fn render_all_pages(&self, dpi: f64) -> Result<Vec<Vec<u8>>> {
        let layout = self.cached_layout()?;
        Ok(rdocx_pdf::render_all_pages(&layout, dpi))
    }

    /// Return a cloned positioned page from the cached normal-font layout.
    ///
    /// `page_index` is zero-based. An index beyond the document returns `None`.
    pub fn layout_page(&self, page_index: usize) -> Result<Option<rdocx_layout::PageFrame>> {
        let layout = self.cached_layout()?;
        Ok(layout.pages.get(page_index).cloned())
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
            BodyContent::Paragraph(p) => Self::collect_images_from_paragraph(p, result),
            BodyContent::Table(tbl) => Self::collect_images_from_table(tbl, result),
            BodyContent::RawXml(_) => {}
        }
    }

    fn collect_images_from_paragraph(p: &CT_P, result: &mut Vec<ImageInfo>) {
        for run in &p.runs {
            for rc in &run.content {
                let RunContent::Drawing(drawing) = rc else {
                    continue;
                };
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

    fn collect_images_from_table(tbl: &CT_Tbl, result: &mut Vec<ImageInfo>) {
        use rdocx_oxml::table::CellContent;

        for row in &tbl.rows {
            for cell in &row.cells {
                for cc in &cell.content {
                    match cc {
                        CellContent::Paragraph(p) => Self::collect_images_from_paragraph(p, result),
                        CellContent::Table(nested) => {
                            Self::collect_images_from_table(nested, result)
                        }
                    }
                }
            }
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
                    // `HyperlinkSpan`'s bounds are public and can be set by
                    // hand, so clamp rather than slice-panic on a bad range.
                    let start = hl.run_start.min(p.runs.len());
                    let end = hl.run_end.clamp(start, p.runs.len());
                    let text: String = p.runs[start..end].iter().map(|r| r.text()).collect();

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
            BodyContent::Table(tbl) => Self::word_count_in_table(tbl),
            BodyContent::RawXml(_) => 0,
        }
    }

    fn word_count_in_table(tbl: &CT_Tbl) -> usize {
        use rdocx_oxml::table::CellContent;

        let mut count = 0;
        for row in &tbl.rows {
            for cell in &row.cells {
                for cc in &cell.content {
                    match cc {
                        CellContent::Paragraph(p) => {
                            count += p.text().split_whitespace().count();
                        }
                        CellContent::Table(nested) => {
                            count += Self::word_count_in_table(nested);
                        }
                    }
                }
            }
        }
        count
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

/// Express `target_part` relative to the directory holding `source_part`.
///
/// Falls back to the absolute part name when the two live in different
/// directories, which OPC also permits.
fn relative_target(source_part: &str, target_part: &str) -> String {
    let dir = match source_part.rfind('/') {
        Some(pos) => &source_part[..=pos],
        None => "/",
    };
    match target_part.strip_prefix(dir) {
        Some(rest) if !rest.contains('/') => rest.to_string(),
        _ => target_part.to_string(),
    }
}

/// The lower-cased file extension of `filename`, defaulting to `png`.
fn image_number_from_part_name(name: &str) -> Option<usize> {
    let suffix = name.strip_prefix("/word/media/image")?;
    let digit_count = suffix
        .bytes()
        .take_while(|byte| byte.is_ascii_digit())
        .count();
    if digit_count == 0 {
        return None;
    }
    suffix[..digit_count]
        .parse::<usize>()
        .ok()
        .filter(|index| *index > 0)
}

fn image_extension(filename: &str) -> String {
    match filename.rsplit_once('.') {
        Some((_, ext)) if !ext.is_empty() => ext.to_lowercase(),
        _ => "png".to_string(),
    }
}

/// Map an image file extension to its MIME type.
///
/// This is the single place the mapping lives; header, footer, body and
/// raw-XML image paths all go through it, so they cannot drift apart and
/// start disagreeing about, say, whether GIF is supported.
fn image_content_type(ext: &str) -> &'static str {
    match ext {
        "jpg" | "jpeg" => "image/jpeg",
        "gif" => "image/gif",
        "bmp" => "image/bmp",
        "tiff" | "tif" => "image/tiff",
        "svg" => "image/svg+xml",
        "webp" => "image/webp",
        // PNG is both the common case and a safe default for unknown types.
        _ => "image/png",
    }
}

/// Guess image content type from the part name extension.
fn guess_image_content_type(part_name: &str) -> String {
    image_content_type(&image_extension(part_name)).to_string()
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

/// Truncate a string to at most `max_len` characters, appending "..." if it
/// was cut short.
///
/// Both the comparison and the cut are in characters; mixing byte length with
/// character counts would truncate non-ASCII text earlier than asked.
fn truncate_str(s: &str, max_len: usize) -> String {
    if s.chars().count() <= max_len {
        return s.to_string();
    }
    let truncated: String = s.chars().take(max_len.saturating_sub(3)).collect();
    format!("{truncated}...")
}

/// Deobfuscate an ODTTF (obfuscated TrueType) font file.
///
/// Word embeds fonts as `.odttf` files whose first 32 bytes are XOR'd with a
/// 16-byte key derived from the GUID in the part name (ECMA-376 Part 1,
/// "Embedded Font Obfuscation"). The GUID hex is read into the key *backwards*,
/// but implementations differ in whether they reverse the raw hex string or the
/// mixed-endian layout .NET's `Guid.ToByteArray` produces — the two agree on
/// the first eight key bytes and disagree on the rest.
///
/// Rather than pick one and hope, both orders are tried and the result is only
/// accepted if it starts with a recognised sfnt version. A wrong key yields
/// bytes that no font parser can use, so validating here means a bad guess
/// degrades to "font not embedded" instead of feeding garbage downstream.
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

    let candidates = odttf_key_candidates(&guid);
    let decoded: Vec<Vec<u8>> = candidates
        .iter()
        .map(|key| {
            let mut result = data.to_vec();
            // XOR the first 32 bytes with the 16-byte key, applied twice.
            for (i, byte) in result.iter_mut().take(32).enumerate() {
                *byte ^= key[i % 16];
            }
            result
        })
        .collect();

    // A well-formed table directory pins down which key was used. Fall back to
    // the weaker signature check for fonts whose header arithmetic is wrong —
    // subsetting tools do emit those — so they still load rather than being
    // dropped entirely.
    decoded
        .iter()
        .find(|d| has_consistent_sfnt_header(d))
        .or_else(|| decoded.iter().find(|d| looks_like_sfnt(d)))
        .cloned()
}

/// The two candidate XOR keys for ODTTF deobfuscation, most likely first.
fn odttf_key_candidates(guid: &[u8; 16]) -> [[u8; 16]; 2] {
    // Read the hex string end-first, as the spec prose describes.
    let mut plain_reversed = *guid;
    plain_reversed.reverse();

    // The .NET route: `Guid.ToByteArray` byte-swaps the first three groups,
    // and the whole array is then reversed.
    let dotnet = [
        guid[3], guid[2], guid[1], guid[0], guid[5], guid[4], guid[7], guid[6], guid[8], guid[9],
        guid[10], guid[11], guid[12], guid[13], guid[14], guid[15],
    ];
    let mut dotnet_reversed = dotnet;
    dotnet_reversed.reverse();

    [plain_reversed, dotnet_reversed]
}

/// Check that `data` opens with a plausible sfnt (TrueType/OpenType) header.
///
/// This is the weak test: signature plus printable table tags. It cannot always
/// tell the two ODTTF key conventions apart, since they produce identical
/// output for the first eight bytes.
fn looks_like_sfnt(data: &[u8]) -> bool {
    let Some(signature) = data.first_chunk::<4>() else {
        return false;
    };
    match signature {
        b"\x00\x01\x00\x00" | b"OTTO" | b"true" => {}
        // A collection header has a different layout; take it on signature.
        b"ttcf" => return true,
        _ => return false,
    }

    if data.len() < 32 {
        return false;
    }

    let num_tables = u16::from_be_bytes([data[4], data[5]]);
    if num_tables == 0 || num_tables > 512 {
        return false;
    }

    // Table records begin at offset 12 and are 16 bytes each, so the first
    // record's tag is at 12..16 and the second record's tag at 28..32 — both
    // inside the 32 bytes the obfuscation touches.
    let is_tag = |tag: &[u8]| tag.iter().all(|b| (0x20..=0x7E).contains(b));
    is_tag(&data[12..16]) && (num_tables < 2 || is_tag(&data[28..32]))
}

/// The strong test: the sfnt header's binary-search hints must agree with the
/// table count.
///
/// `searchRange`, `entrySelector` and `rangeShift` are all derived from
/// `numTables`, and `entrySelector`/`rangeShift` sit in the byte range where
/// the two ODTTF key conventions differ — so this identifies the right key
/// outright whenever the font's header is spec-conformant.
fn has_consistent_sfnt_header(data: &[u8]) -> bool {
    if !looks_like_sfnt(data) || data.len() < 12 {
        return false;
    }
    if data.first_chunk::<4>() == Some(b"ttcf") {
        return false; // no table directory at this offset
    }

    let num_tables = u16::from_be_bytes([data[4], data[5]]);
    let search_range = u16::from_be_bytes([data[6], data[7]]);
    let entry_selector = u16::from_be_bytes([data[8], data[9]]);
    let range_shift = u16::from_be_bytes([data[10], data[11]]);

    let expected_selector = num_tables.ilog2() as u16;
    let expected_search_range = (1u16 << expected_selector) * 16;
    let expected_range_shift = num_tables
        .wrapping_mul(16)
        .wrapping_sub(expected_search_range);

    search_range == expected_search_range
        && entry_selector == expected_selector
        && range_shift == expected_range_shift
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::paragraph::Alignment;
    use rdocx_oxml::units::{HalfPoint, Twips};

    fn reset_layout_invocations() {
        LAYOUT_INVOCATIONS.set(0);
    }

    fn layout_invocations() -> usize {
        LAYOUT_INVOCATIONS.get()
    }

    #[test]
    fn rendering_all_pages_performs_one_layout() {
        let mut doc = Document::new();
        doc.add_paragraph("Page 1");
        for page in 2..=20 {
            doc.add_paragraph(&format!("Page {page}"))
                .page_break_before(true);
        }

        reset_layout_invocations();
        for page_index in 0..20 {
            assert!(
                doc.render_page_to_png_deterministic(page_index, 1.0)
                    .expect("deterministic layout should succeed")
                    .is_some(),
                "page {page_index} should exist"
            );
        }

        assert_eq!(layout_invocations(), 1);
    }

    #[test]
    fn document_mutation_invalidates_cached_layout() {
        let mut doc = Document::new();
        doc.add_paragraph("Before mutation");

        reset_layout_invocations();
        doc.render_page_to_png_deterministic(0, 1.0).unwrap();
        doc.render_page_to_png_deterministic(0, 1.0).unwrap();
        assert_eq!(layout_invocations(), 1);

        doc.add_paragraph("After mutation");
        doc.render_page_to_png_deterministic(0, 1.0).unwrap();
        doc.render_page_to_png_deterministic(0, 1.0).unwrap();
        assert_eq!(layout_invocations(), 2);
    }

    #[test]
    fn mutable_accessor_invalidates_cached_layout() {
        let mut doc = Document::new();
        doc.add_paragraph("Before wrapper mutation");

        reset_layout_invocations();
        doc.render_page_to_png_deterministic(0, 1.0).unwrap();
        doc.render_page_to_png_deterministic(0, 1.0).unwrap();
        assert_eq!(layout_invocations(), 1);

        doc.paragraph_mut(0)
            .expect("paragraph should exist")
            .add_run(" changed");
        doc.render_page_to_png_deterministic(0, 1.0).unwrap();
        doc.render_page_to_png_deterministic(0, 1.0).unwrap();
        assert_eq!(layout_invocations(), 2);

        let mut table = doc.add_table(1, 1);
        table
            .cell(0, 0)
            .expect("cell should exist")
            .set_text("table mutation");
        doc.render_page_to_png_deterministic(0, 1.0).unwrap();
        doc.render_page_to_png_deterministic(0, 1.0).unwrap();
        assert_eq!(layout_invocations(), 3);
    }

    #[test]
    fn font_modes_use_isolated_layout_caches() {
        let mut doc = Document::new();
        doc.add_paragraph("Font mode isolation");

        reset_layout_invocations();
        doc.render_page_to_png(0, 1.0).unwrap();
        doc.render_page_to_png(0, 1.0).unwrap();
        assert!(doc.layout_page(0).unwrap().is_some());
        assert!(doc.layout_page(usize::MAX).unwrap().is_none());
        assert_eq!(doc.render_all_pages(1.0).unwrap().len(), 1);
        assert!(!doc.to_pdf().unwrap().is_empty());
        assert_eq!(layout_invocations(), 1);

        doc.render_page_to_png_deterministic(0, 1.0).unwrap();
        doc.render_page_to_png_deterministic(0, 1.0).unwrap();
        assert_eq!(layout_invocations(), 2);

        doc.render_page_to_png(0, 1.0).unwrap();
        doc.render_page_to_png_deterministic(0, 1.0).unwrap();
        assert_eq!(layout_invocations(), 2);

        let (family, font_data) = rdocx_layout::bundled_fonts::bundled_font_data()[0];
        doc.to_pdf_with_fonts(&[(family, font_data)]).unwrap();
        doc.to_pdf_with_fonts(&[(family, font_data)]).unwrap();
        assert_eq!(layout_invocations(), 4);
    }

    #[test]
    fn document_remains_send_and_sync() {
        fn assert_send_and_sync<T: Send + Sync>() {}
        assert_send_and_sync::<Document>();
    }

    #[test]
    fn deterministic_render_is_independent_of_system_fonts() {
        let mut doc = Document::new();
        doc.add_paragraph("Deterministic rendering");

        let input = doc.build_layout_input();
        let layout = rdocx_layout::layout_document_deterministic(&input)
            .expect("deterministic layout should succeed");
        let bundled_fonts = rdocx_layout::bundled_fonts::bundled_font_data();

        assert!(!layout.fonts.is_empty());
        for font in &layout.fonts {
            assert!(!font.data.is_empty());
            assert!(
                bundled_fonts
                    .iter()
                    .any(|(_family, data)| *data == font.data.as_slice()),
                "resolved font '{}' did not come from the bundled font set",
                font.family
            );
        }

        let inspected = rdocx_pdf::render_page_to_png(&layout, 0, 150.0)
            .expect("document should have a first page");
        let facade = doc
            .render_page_to_png_deterministic(0, 150.0)
            .expect("deterministic layout should succeed")
            .expect("document should have a first page");

        assert!(!inspected.is_empty());
        assert_eq!(facade, inspected);
    }

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

    #[test]
    fn numbering_getter_round_trips() {
        let mut doc = Document::new();
        doc.add_bullet_list_item("bullet item", 0);
        doc.add_numbered_list_item("numbered item", 0);
        doc.add_paragraph("plain");

        let bytes = doc.to_bytes().unwrap();
        let doc2 = Document::from_bytes(&bytes).unwrap();
        let paras = doc2.paragraphs();

        let (bullet_id, bullet_lvl) = paras[0].numbering().expect("bullet numbering");
        assert_eq!(bullet_lvl, 0);
        assert_eq!(doc2.numbering_is_bullet(bullet_id), Some(true));

        let (num_id, _) = paras[1].numbering().expect("numbered numbering");
        assert_eq!(doc2.numbering_is_bullet(num_id), Some(false));

        assert!(paras[2].numbering().is_none());
    }

    #[test]
    fn highlight_getter_round_trips() {
        let mut doc = Document::new();
        {
            let mut p = doc.add_paragraph("");
            let mut r = p.add_run("glowing");
            r = r.highlight("yellow");
            let _ = r;
        }

        let bytes = doc.to_bytes().unwrap();
        let doc2 = Document::from_bytes(&bytes).unwrap();
        let paras = doc2.paragraphs();
        let run = paras[0].runs().next().expect("run");
        assert_eq!(run.highlight().as_deref(), Some("yellow"));
    }

    #[test]
    fn run_style_id_getter_round_trips() {
        let mut doc = Document::new();
        {
            let mut p = doc.add_paragraph("");
            let mut r = p.add_run("code text");
            r = r.style("SourceText");
            let _ = r;
        }

        let bytes = doc.to_bytes().unwrap();
        let doc2 = Document::from_bytes(&bytes).unwrap();
        let paras = doc2.paragraphs();
        let run = paras[0].runs().next().expect("run");
        assert_eq!(run.style_id(), Some("SourceText"));
    }

    #[test]
    fn append_hyperlink_round_trips() {
        let mut doc = Document::new();
        doc.add_paragraph("visit ");
        doc.append_hyperlink("GNOME", "https://gnome.org");

        let bytes = doc.to_bytes().unwrap();
        let doc2 = Document::from_bytes(&bytes).unwrap();

        let links = doc2.links();
        assert_eq!(links.len(), 1);
        assert_eq!(links[0].text, "GNOME");
        assert_eq!(links[0].url.as_deref(), Some("https://gnome.org"));
        assert_eq!(doc2.paragraphs()[0].text(), "visit GNOME");

        let paras = doc2.paragraphs();
        let spans = paras[0].hyperlink_spans();
        assert_eq!(spans.len(), 1);
        let (start, end, rel_id) = (spans[0].0, spans[0].1, spans[0].2);
        assert_eq!(end - start, 1);
        let url = doc2.hyperlink_url(rel_id.expect("rel id"));
        assert_eq!(url.as_deref(), Some("https://gnome.org"));
    }

    #[test]
    fn picture_round_trips() {
        // 1x1 red PNG
        let png: &[u8] = &[
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48,
            0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00,
            0x00, 0x90, 0x77, 0x53, 0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41, 0x54, 0x08,
            0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00, 0x00, 0x00, 0x03, 0x00, 0x01, 0x9E, 0xDD, 0x22,
            0x71, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
        ];
        let mut doc = Document::new();
        doc.add_paragraph("before");
        doc.add_picture(png, "dot.png", Length::inches(1.0), Length::inches(1.0));

        let bytes = doc.to_bytes().unwrap();
        let doc2 = Document::from_bytes(&bytes).unwrap();
        let paras = doc2.paragraphs();
        let mut found = None;
        for p in &paras {
            for r in p.runs() {
                if let Some((rel, _alt)) = r.inline_image() {
                    found = Some(rel.to_string());
                }
            }
        }
        let rel = found.expect("no inline image found on read");
        let data = doc2.image_data(&rel).expect("image bytes missing");
        assert_eq!(data, png);
    }
}

#[cfg(test)]
mod hyperlink_span_tests {
    use super::*;
    use rdocx_oxml::text::HyperlinkSpan;

    /// `HyperlinkSpan`'s bounds are public, so a caller building the OXML model
    /// by hand can hand us a range past the end of `runs`. `links()` used to
    /// slice with it and panic.
    #[test]
    fn links_clamps_out_of_range_spans() {
        let mut doc = Document::new();
        {
            let mut para = doc.add_paragraph("");
            para.add_run("one");
            para.add_run("two");
        }

        let BodyContent::Paragraph(p) = &mut doc.document.body.content[0] else {
            unreachable!("just added a paragraph")
        };
        p.hyperlinks.push(HyperlinkSpan {
            rel_id: None,
            anchor: Some("bookmark".to_string()),
            run_start: 1,
            run_end: 99,
        });
        p.hyperlinks.push(HyperlinkSpan {
            rel_id: None,
            anchor: Some("inverted".to_string()),
            run_start: 5,
            run_end: 1,
        });

        let links = doc.links();

        assert_eq!(links.len(), 2);
        assert_eq!(links[0].text, "two");
        assert_eq!(links[1].text, "");
    }
}

#[cfg(test)]
mod odttf_tests {
    use super::*;

    /// Build a TrueType header with a two-entry table directory.
    fn fake_font() -> Vec<u8> {
        let mut data = Vec::new();
        data.extend(b"\x00\x01\x00\x00"); // sfntVersion
        data.extend(2u16.to_be_bytes()); // numTables
        data.extend(32u16.to_be_bytes()); // searchRange
        data.extend(1u16.to_be_bytes()); // entrySelector
        data.extend(0u16.to_be_bytes()); // rangeShift
        for (tag, offset, length) in [(b"cmap", 96u32, 40u32), (b"head", 136, 54)] {
            data.extend(tag); // tag
            data.extend(0u32.to_be_bytes()); // checksum
            data.extend(offset.to_be_bytes());
            data.extend(length.to_be_bytes());
        }
        data.extend((0u8..64).map(|i| i.wrapping_mul(7)));
        data
    }

    fn obfuscate(font: &[u8], key: &[u8; 16]) -> Vec<u8> {
        let mut out = font.to_vec();
        for (i, byte) in out.iter_mut().take(32).enumerate() {
            *byte ^= key[i % 16];
        }
        out
    }

    const GUID_HEX: &str = "00112233445566778899AABBCCDDEEFF";

    fn guid_bytes() -> [u8; 16] {
        let mut g = [0u8; 16];
        for (i, b) in g.iter_mut().enumerate() {
            *b = u8::from_str_radix(&GUID_HEX[i * 2..i * 2 + 2], 16).unwrap();
        }
        g
    }

    #[test]
    fn recovers_font_under_either_key_convention() {
        let font = fake_font();
        let name = format!("{GUID_HEX}.odttf");
        for key in odttf_key_candidates(&guid_bytes()) {
            let obfuscated = obfuscate(&font, &key);
            assert_eq!(
                deobfuscate_odttf(&obfuscated, &name).as_deref(),
                Some(font.as_slice()),
                "failed to recover font for key {key:02x?}",
            );
        }
    }

    #[test]
    fn rejects_data_that_does_not_decode_to_a_font() {
        // A GUID that matches nothing in the payload must not yield garbage.
        let junk = vec![0xAB; 64];
        let name = format!("{GUID_HEX}.odttf");
        assert_eq!(deobfuscate_odttf(&junk, &name), None);
    }

    #[test]
    fn rejects_short_or_malformed_input() {
        assert_eq!(deobfuscate_odttf(&[0u8; 8], "abc.odttf"), None);
        assert_eq!(deobfuscate_odttf(&[0u8; 64], "not-a-guid.odttf"), None);
    }

    #[test]
    fn accepts_braced_and_hyphenated_names() {
        let font = fake_font();
        let key = odttf_key_candidates(&guid_bytes())[0];
        let obfuscated = obfuscate(&font, &key);
        for name in [
            "{00112233-4455-6677-8899-AABBCCDDEEFF}.odttf",
            "00112233-4455-6677-8899-AABBCCDDEEFF.odttf",
        ] {
            assert_eq!(
                deobfuscate_odttf(&obfuscated, name).as_deref(),
                Some(font.as_slice()),
                "failed for {name}"
            );
        }
    }
}
