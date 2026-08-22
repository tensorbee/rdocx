//! The main Document type — entry point for the rdocx API.

use std::collections::HashSet;
use std::path::Path;
use std::sync::{Arc, Mutex};

#[cfg(test)]
use std::cell::Cell;

use oxml_chart::{CT_ChartSpace, ChartData, ChartKind};
use oxml_media::MediaNamer;
use oxml_opc::content_types;
use oxml_opc::relationship::rel_types;
use oxml_opc::{OpcPackage, PackageReadLimits};
use oxml_sml::Workbook;
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use rdocx_oxml::content_control::{CT_Sdt, SdtContent};
use rdocx_oxml::document::{BodyContent, CT_Columns, CT_Document, CT_SectPr};
use rdocx_oxml::drawing::{CT_Anchor, CT_Drawing, CT_Inline};
use rdocx_oxml::header_footer::{
    CT_HdrFtr, HdrFtrRef, HdrFtrType, VmlWatermark, replace_authored_watermark,
};
use rdocx_oxml::namespace::matches_local_name;
use rdocx_oxml::numbering::{CT_Numbering, ST_NumberFormat};
use rdocx_oxml::properties::{CT_PPr, CT_RPr};
use rdocx_oxml::settings::{CT_Settings, DocumentProtection};
use rdocx_oxml::shared::{ST_PageOrientation, ST_SectionType};
use rdocx_oxml::styles::CT_Styles;
use rdocx_oxml::table::{CT_Row, CT_Tbl, CT_Tc, CellContent};
use rdocx_oxml::text::{CT_P, CT_R, RunContent};

use oxml_core::custom_properties::CustomProperties;
use rdocx_oxml::core_properties::CoreProperties;

use crate::Length;
use crate::content_control::ContentControlRef;
use crate::error::{Error, Result};
use crate::paragraph::{Paragraph, ParagraphRef};
use crate::revision::RevisionRef;
use crate::style::{self, Style, StyleBuilder};
use crate::table::{Table, TableRef};

/// Options that select a native document render projection.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct RenderOptions {
    /// The tracked-revision view to render.
    pub revision_view: rdocx_layout::RevisionView,
}

/// One direct child of a document body, in source order.
pub enum BodyItemRef<'a> {
    /// A body paragraph.
    Paragraph(ParagraphRef<'a>),
    /// A body table.
    Table(TableRef<'a>),
    /// A body-level content control.
    ContentControl(ContentControlRef<'a>),
    /// A preserved body child that rdocx does not model.
    UnsupportedXml(&'a [u8]),
}

/// A Word document (.docx file).
///
/// This is the main entry point for reading, creating, and modifying
/// DOCX documents.
pub struct Document {
    pub(crate) package: OpcPackage,
    pub(crate) document: CT_Document,
    pub(crate) styles: CT_Styles,
    numbering: Option<CT_Numbering>,
    pub(crate) core_properties: Option<CoreProperties>,
    /// Read-only custom document properties resolved from package relationships.
    pub(crate) custom_properties: Option<CustomProperties>,
    /// Package part containing the core properties, resolved from `_rels/.rels`.
    core_properties_part_name: String,
    /// Part name for the main document
    pub(crate) doc_part_name: String,
    /// Part name the styles were loaded from, and where they are written back.
    /// Resolved through the relationship rather than assumed, so a document
    /// that keeps its styles somewhere other than `/word/styles.xml` is
    /// updated in place instead of gaining an orphaned second part.
    styles_part_name: String,
    /// Part name for numbering definitions, resolved the same way.
    numbering_part_name: String,
    /// Typed document settings loaded through the main document relationship.
    pub(crate) settings: Option<CT_Settings>,
    /// Existing settings relationship target. No conventional target is assumed.
    settings_part_name: Option<String>,
    /// Collision-free allocator for image media parts.
    image_namer: MediaNamer,
    /// Typed footnotes loaded through the main document relationship.
    pub(crate) footnotes: rdocx_oxml::footnotes::CT_Footnotes,
    /// Existing footnotes relationship target. No conventional target is assumed on read.
    pub(crate) footnotes_part_name: Option<String>,
    /// Whether a facade mutation requires complete typed footnote serialization.
    pub(crate) footnotes_dirty: bool,
    /// Typed comments loaded through the main document relationship.
    pub(crate) comments: Option<rdocx_oxml::comments::CT_Comments>,
    /// Existing comments relationship target. No target is invented on read.
    pub(crate) comments_part_name: Option<String>,
    /// Typed reply linkage and resolved state for comments.
    pub(crate) comments_extended: Option<rdocx_oxml::comments_extended::CT_CommentsEx>,
    /// Existing comments-extended relationship target.
    pub(crate) comments_extended_part_name: Option<String>,
    /// Whether this facade created the comments part and may remove it when empty.
    pub(crate) comments_owned: bool,
    /// Whether this facade created the comments-extended part and may remove it when empty.
    pub(crate) comments_extended_owned: bool,
    /// Normal layout, including system font discovery, computed on first use.
    layout_cache: Mutex<Option<Arc<rdocx_layout::WordLayoutResult>>>,
    /// Reusable normal-font engine retained across document edits.
    normal_layout_engine: Mutex<Option<rdocx_layout::engine::Engine>>,
    /// Bundled-font-only layout used by deterministic rendering.
    deterministic_layout_cache: Mutex<Option<Arc<rdocx_layout::WordLayoutResult>>>,
    /// SVG PoC patch: reusable engine for the bundled-fallback caller-fonts
    /// path (wasm editors), kept across edits like `normal_layout_engine`.
    fallback_layout_engine: Mutex<Option<rdocx_layout::engine::Engine>>,
}

enum ChartPackageSource<'a> {
    #[allow(dead_code)]
    Typed {
        chart: &'a CT_ChartSpace,
        workbook: &'a Workbook,
    },
    Authored {
        kind: ChartKind,
        data: &'a ChartData,
    },
}

/// Fallback part names used when a document does not already declare one.
const DEFAULT_STYLES_PART: &str = "/word/styles.xml";
const DEFAULT_NUMBERING_PART: &str = "/word/numbering.xml";
const DEFAULT_CORE_PROPERTIES_PART: &str = "/docProps/core.xml";
const DOCUMENT_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
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

fn new_word_package() -> OpcPackage {
    let mut package = OpcPackage::with_main_part("word/document.xml", DOCUMENT_CONTENT_TYPE);
    package
        .content_types
        .add_override(DEFAULT_STYLES_PART, STYLES_CONTENT_TYPE);
    package
}

fn take_paragraph<'a>(paragraph: &'a mut CT_P, index: &mut usize) -> Option<&'a mut CT_P> {
    if *index == 0 {
        Some(paragraph)
    } else {
        *index -= 1;
        None
    }
}

fn nth_paragraph_in_body<'a>(
    content: &'a mut [BodyContent],
    index: &mut usize,
) -> Option<&'a mut CT_P> {
    for child in content {
        let paragraph = match child {
            BodyContent::Paragraph(paragraph) => take_paragraph(paragraph, index),
            BodyContent::ContentControl(control) => nth_paragraph_in_control(control, index),
            BodyContent::Table(_) | BodyContent::RawXml(_) => None,
        };
        if paragraph.is_some() {
            return paragraph;
        }
    }
    None
}

fn nth_paragraph_in_control<'a>(
    control: &'a mut CT_Sdt,
    index: &mut usize,
) -> Option<&'a mut CT_P> {
    for child in &mut control.content {
        let paragraph = match child {
            SdtContent::Paragraph(paragraph) => take_paragraph(paragraph, index),
            SdtContent::Table(table) => nth_paragraph_in_table(table, index),
            SdtContent::Row(row) => nth_paragraph_in_row(row, index),
            SdtContent::Cell(cell) => nth_paragraph_in_cell(cell, index),
            SdtContent::ContentControl(control) => nth_paragraph_in_control(control, index),
            SdtContent::Run(_) | SdtContent::RawXml(_) => None,
        };
        if paragraph.is_some() {
            return paragraph;
        }
    }
    None
}

fn paragraph_count_in_control(control: &CT_Sdt) -> usize {
    control
        .content
        .iter()
        .map(|child| match child {
            SdtContent::Paragraph(_) => 1,
            SdtContent::Table(table) => paragraph_count_in_table(table),
            SdtContent::Row(row) => paragraph_count_in_row(row),
            SdtContent::Cell(cell) => paragraph_count_in_cell(cell),
            SdtContent::ContentControl(control) => paragraph_count_in_control(control),
            SdtContent::Run(_) | SdtContent::RawXml(_) => 0,
        })
        .sum()
}

fn paragraph_count_in_table(table: &CT_Tbl) -> usize {
    table
        .content_controls
        .iter()
        .map(|(_, _, control)| paragraph_count_in_control(control))
        .sum::<usize>()
        + table.rows.iter().map(paragraph_count_in_row).sum::<usize>()
}

fn paragraph_count_in_row(row: &CT_Row) -> usize {
    row.content_controls
        .iter()
        .map(|(_, _, control)| paragraph_count_in_control(control))
        .sum::<usize>()
        + row.cells.iter().map(paragraph_count_in_cell).sum::<usize>()
}

fn paragraph_count_in_cell(cell: &CT_Tc) -> usize {
    cell.content
        .iter()
        .map(|child| match child {
            CellContent::Paragraph(_) => 1,
            CellContent::ContentControl(control) => paragraph_count_in_control(control),
            CellContent::Table(_) => 0,
        })
        .sum()
}

fn nth_paragraph_in_table<'a>(table: &'a mut CT_Tbl, index: &mut usize) -> Option<&'a mut CT_P> {
    let CT_Tbl {
        rows,
        content_controls,
        ..
    } = table;
    let mut selected_control = None;
    let mut selected_row = None;
    for boundary in 0..=rows.len() {
        for (control_index, (_, _, control)) in content_controls
            .iter()
            .enumerate()
            .filter(|(_, (at, _, _))| *at == boundary)
        {
            let count = paragraph_count_in_control(control);
            if *index < count {
                selected_control = Some(control_index);
                break;
            }
            *index -= count;
        }
        if selected_control.is_some() {
            break;
        }
        if let Some(row) = rows.get(boundary) {
            let count = paragraph_count_in_row(row);
            if *index < count {
                selected_row = Some(boundary);
                break;
            }
            *index -= count;
        }
    }
    if let Some(control_index) = selected_control {
        nth_paragraph_in_control(&mut content_controls[control_index].2, index)
    } else if let Some(row_index) = selected_row {
        nth_paragraph_in_row(&mut rows[row_index], index)
    } else {
        None
    }
}

fn nth_paragraph_in_row<'a>(row: &'a mut CT_Row, index: &mut usize) -> Option<&'a mut CT_P> {
    let CT_Row {
        cells,
        content_controls,
        ..
    } = row;
    let mut selected_control = None;
    let mut selected_cell = None;
    for boundary in 0..=cells.len() {
        for (control_index, (_, _, control)) in content_controls
            .iter()
            .enumerate()
            .filter(|(_, (at, _, _))| *at == boundary)
        {
            let count = paragraph_count_in_control(control);
            if *index < count {
                selected_control = Some(control_index);
                break;
            }
            *index -= count;
        }
        if selected_control.is_some() {
            break;
        }
        if let Some(cell) = cells.get(boundary) {
            let count = paragraph_count_in_cell(cell);
            if *index < count {
                selected_cell = Some(boundary);
                break;
            }
            *index -= count;
        }
    }
    if let Some(control_index) = selected_control {
        nth_paragraph_in_control(&mut content_controls[control_index].2, index)
    } else if let Some(cell_index) = selected_cell {
        nth_paragraph_in_cell(&mut cells[cell_index], index)
    } else {
        None
    }
}

fn nth_paragraph_in_cell<'a>(cell: &'a mut CT_Tc, index: &mut usize) -> Option<&'a mut CT_P> {
    for child in &mut cell.content {
        let paragraph = match child {
            CellContent::Paragraph(paragraph) => take_paragraph(paragraph, index),
            CellContent::ContentControl(control) => nth_paragraph_in_control(control, index),
            CellContent::Table(_) => None,
        };
        if paragraph.is_some() {
            return paragraph;
        }
    }
    None
}

fn take_table<'a>(table: &'a mut CT_Tbl, index: &mut usize) -> Option<&'a mut CT_Tbl> {
    if *index == 0 {
        Some(table)
    } else {
        *index -= 1;
        None
    }
}

fn nth_table_in_body<'a>(
    content: &'a mut [BodyContent],
    index: &mut usize,
) -> Option<&'a mut CT_Tbl> {
    for child in content {
        let table = match child {
            BodyContent::Table(table) => take_table(table, index),
            BodyContent::ContentControl(control) => nth_table_in_control(control, index),
            BodyContent::Paragraph(_) | BodyContent::RawXml(_) => None,
        };
        if table.is_some() {
            return table;
        }
    }
    None
}

fn nth_table_in_control<'a>(control: &'a mut CT_Sdt, index: &mut usize) -> Option<&'a mut CT_Tbl> {
    for child in &mut control.content {
        let table = match child {
            SdtContent::Table(table) => take_table(table, index),
            SdtContent::Cell(cell) => nth_table_in_cell(cell, index),
            SdtContent::ContentControl(control) => nth_table_in_control(control, index),
            SdtContent::Paragraph(_)
            | SdtContent::Row(_)
            | SdtContent::Run(_)
            | SdtContent::RawXml(_) => None,
        };
        if table.is_some() {
            return table;
        }
    }
    None
}

fn nth_table_in_cell<'a>(cell: &'a mut CT_Tc, index: &mut usize) -> Option<&'a mut CT_Tbl> {
    for child in &mut cell.content {
        let table = match child {
            CellContent::Table(table) => take_table(table, index),
            CellContent::ContentControl(control) => nth_table_in_control(control, index),
            CellContent::Paragraph(_) => None,
        };
        if table.is_some() {
            return table;
        }
    }
    None
}

impl Document {
    /// Create a new, empty document with default page setup and styles.
    pub fn new() -> Self {
        let mut package = new_word_package();
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
            custom_properties: None,
            core_properties_part_name: DEFAULT_CORE_PROPERTIES_PART.to_string(),
            doc_part_name: "/word/document.xml".to_string(),
            styles_part_name: DEFAULT_STYLES_PART.to_string(),
            numbering_part_name: DEFAULT_NUMBERING_PART.to_string(),
            settings: None,
            settings_part_name: None,
            image_namer: MediaNamer::scan("/word/media", "image", std::iter::empty()),
            footnotes: rdocx_oxml::footnotes::CT_Footnotes::new(),
            footnotes_part_name: None,
            footnotes_dirty: false,
            comments: None,
            comments_part_name: None,
            comments_extended: None,
            comments_extended_part_name: None,
            comments_owned: false,
            comments_extended_owned: false,
            layout_cache: Mutex::new(None),
            normal_layout_engine: Mutex::new(None),
            deterministic_layout_cache: Mutex::new(None),
            fallback_layout_engine: Mutex::new(None),
        }
    }

    /// Clone all package and typed state while discarding derived layout caches.
    pub(crate) fn clone_for_staging(&self) -> Self {
        Self {
            package: self.package.clone(),
            document: self.document.clone(),
            styles: self.styles.clone(),
            numbering: self.numbering.clone(),
            core_properties: self.core_properties.clone(),
            custom_properties: self.custom_properties.clone(),
            core_properties_part_name: self.core_properties_part_name.clone(),
            doc_part_name: self.doc_part_name.clone(),
            styles_part_name: self.styles_part_name.clone(),
            numbering_part_name: self.numbering_part_name.clone(),
            settings: self.settings.clone(),
            settings_part_name: self.settings_part_name.clone(),
            image_namer: self.image_namer.clone(),
            footnotes: self.footnotes.clone(),
            footnotes_part_name: self.footnotes_part_name.clone(),
            footnotes_dirty: self.footnotes_dirty,
            comments: self.comments.clone(),
            comments_part_name: self.comments_part_name.clone(),
            comments_extended: self.comments_extended.clone(),
            comments_extended_part_name: self.comments_extended_part_name.clone(),
            comments_owned: self.comments_owned,
            comments_extended_owned: self.comments_extended_owned,
            layout_cache: Mutex::new(None),
            normal_layout_engine: Mutex::new(None),
            deterministic_layout_cache: Mutex::new(None),
            fallback_layout_engine: Mutex::new(None),
        }
    }

    /// Commit staged package state without discarding reusable layout work.
    fn commit_staged_mutation(&mut self, mut candidate: Self) {
        std::mem::swap(
            &mut self.normal_layout_engine,
            &mut candidate.normal_layout_engine,
        );
        *self = candidate;
    }

    /// Open a document from a file path.
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        let package = OpcPackage::open(path)?;
        Self::from_package(package)
    }

    /// Open a document from bytes.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_bytes_with_limits(bytes, PackageReadLimits::UNBOUNDED)
    }

    /// Open a document from bytes while bounding OPC archive expansion.
    pub fn from_bytes_with_limits(bytes: &[u8], limits: PackageReadLimits) -> Result<Self> {
        let cursor = std::io::Cursor::new(bytes);
        let package = OpcPackage::from_reader_with_limits(cursor, limits)?;
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

        let settings_part_name = resolve_part(rel_types::SETTINGS);
        let settings = match settings_part_name
            .as_deref()
            .and_then(|part| package.get_part(part))
        {
            Some(xml) => Some(CT_Settings::from_xml(xml)?),
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

        let custom_properties = package
            .package_rels
            .get_by_type(rel_types::CUSTOM_PROPERTIES)
            .map(|rel| OpcPackage::resolve_rel_target("/", &rel.target))
            .and_then(|part| package.get_part(&part))
            .and_then(|xml| CustomProperties::from_xml(xml).ok());

        let image_namer = MediaNamer::scan(
            "/word/media",
            "image",
            package.parts.keys().map(String::as_str),
        );

        let footnotes_part_name = resolve_part(rel_types::FOOTNOTES);
        let footnotes = footnotes_part_name
            .as_deref()
            .and_then(|part| package.get_part(part))
            .and_then(|xml| rdocx_oxml::footnotes::CT_Footnotes::from_xml(xml).ok())
            .unwrap_or_default();

        let comments_part_name = resolve_part(rel_types::COMMENTS);
        let comments = match comments_part_name
            .as_deref()
            .and_then(|part| package.get_part(part))
        {
            Some(xml) => Some(rdocx_oxml::comments::CT_Comments::from_xml(xml)?),
            None => None,
        };
        let comments_extended_part_name = resolve_part(crate::comments::COMMENTS_EXTENDED_REL_TYPE);
        let comments_extended = match comments_extended_part_name
            .as_deref()
            .and_then(|part| package.get_part(part))
        {
            Some(xml) => Some(rdocx_oxml::comments_extended::CT_CommentsEx::from_xml(xml)?),
            None => None,
        };

        Ok(Document {
            package,
            document,
            styles,
            numbering,
            core_properties,
            custom_properties,
            core_properties_part_name: core_properties_part_name
                .unwrap_or_else(|| DEFAULT_CORE_PROPERTIES_PART.to_string()),
            doc_part_name,
            styles_part_name: styles_part_name.unwrap_or_else(|| DEFAULT_STYLES_PART.to_string()),
            numbering_part_name: numbering_part_name
                .unwrap_or_else(|| DEFAULT_NUMBERING_PART.to_string()),
            settings,
            settings_part_name,
            image_namer,
            footnotes,
            footnotes_part_name,
            footnotes_dirty: false,
            comments,
            comments_part_name,
            comments_extended,
            comments_extended_part_name,
            comments_owned: false,
            comments_extended_owned: false,
            layout_cache: Mutex::new(None),
            normal_layout_engine: Mutex::new(None),
            deterministic_layout_cache: Mutex::new(None),
            fallback_layout_engine: Mutex::new(None),
        })
    }

    /// Clear layouts derived from the current document state.
    pub(crate) fn invalidate_layout(&mut self) {
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
    fn cached_layout(&self) -> Result<Arc<rdocx_layout::WordLayoutResult>> {
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
        let mut engine = self
            .normal_layout_engine
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner);
        let engine = engine.get_or_insert_with(rdocx_layout::engine::Engine::new);
        let layout = Arc::new(rdocx_layout::layout_document_with_reusable_engine(
            engine, &input,
        )?);
        *cache = Some(Arc::clone(&layout));
        Ok(layout)
    }

    /// Return the bundled-font-only layout, computing it once after mutation.
    fn cached_deterministic_layout(&self) -> Result<Arc<rdocx_layout::WordLayoutResult>> {
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
        let layout = Arc::new(rdocx_layout::layout_document_deterministic_with_provenance(
            &input,
        )?);
        *cache = Some(Arc::clone(&layout));
        Ok(layout)
    }

    fn layout_for_options(
        &self,
        options: RenderOptions,
        deterministic: bool,
    ) -> Result<Arc<rdocx_layout::WordLayoutResult>> {
        if options.revision_view == rdocx_layout::RevisionView::Accepted {
            return if deterministic {
                self.cached_deterministic_layout()
            } else {
                self.cached_layout()
            };
        }

        let mut input = self.build_layout_input();
        input.revision_view = options.revision_view;
        #[cfg(test)]
        record_layout_invocation();
        let layout = if deterministic {
            rdocx_layout::layout_document_deterministic_with_provenance(&input)?
        } else {
            let mut engine = self
                .normal_layout_engine
                .lock()
                .unwrap_or_else(std::sync::PoisonError::into_inner);
            let engine = engine.get_or_insert_with(rdocx_layout::engine::Engine::new);
            rdocx_layout::layout_document_with_reusable_engine(engine, &input)?
        };
        Ok(Arc::new(layout))
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

        // F-155 exposes settings as a read-only projection. Parsed settings
        // retain their complete producer bytes and are written back only to
        // the relationship-resolved part they came from.
        if let (Some(settings), Some(part_name)) = (&self.settings, &self.settings_part_name) {
            self.package.set_part(part_name, settings.to_xml()?);
        }

        // Preserve parsed footnote bytes until a facade mutation makes the typed view dirty.
        if self.footnotes_dirty && !self.footnotes.footnotes.is_empty() {
            let fx = self.footnotes.to_xml_footnotes()?;
            let footnotes_part = self
                .footnotes_part_name
                .clone()
                .unwrap_or_else(|| "/word/footnotes.xml".to_owned());
            self.package.set_part(&footnotes_part, fx);
            self.package.content_types.add_override(
                &footnotes_part,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
            );
            let rels = self
                .package
                .get_or_create_part_rels(&self.doc_part_name.clone());
            if rels.get_by_type(rel_types::FOOTNOTES).is_none() {
                rels.add(rel_types::FOOTNOTES, "footnotes.xml");
            }
        }

        // An existing comments part is modelled and flushed to its resolved
        // relationship target. An absent part remains absent until the later
        // comment API creates one deliberately.
        if let (Some(comments), Some(part_name)) = (&self.comments, self.comments_part_name.clone())
        {
            let xml = comments.to_xml()?;
            self.package.set_part(&part_name, xml);
            self.ensure_part_relationship(
                &part_name,
                rel_types::COMMENTS,
                crate::comments::COMMENTS_CONTENT_TYPE,
            );
        }

        if let (Some(comments), Some(part_name)) = (
            &self.comments_extended,
            self.comments_extended_part_name.clone(),
        ) {
            let xml = comments.to_xml()?;
            self.package.set_part(&part_name, xml);
            self.ensure_part_relationship(
                &part_name,
                crate::comments::COMMENTS_EXTENDED_REL_TYPE,
                crate::comments::COMMENTS_EXTENDED_CONTENT_TYPE,
            );
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

    /// Stage a chart, its editable workbook, and the Word drawing that reaches them.
    fn add_chart_package(
        &mut self,
        source: ChartPackageSource<'_>,
        width: Length,
        height: Length,
    ) -> Result<()> {
        if width.to_emu() <= 0 || height.to_emu() <= 0 {
            return Err(Error::Other(
                "chart width and height must be positive".to_owned(),
            ));
        }

        let mut package = self.package.clone();
        let mut document = self.document.clone();
        let mut chart_namer = MediaNamer::scan(
            "/word/charts",
            "chart",
            package.parts.keys().map(String::as_str),
        );
        let mut workbook_namer = MediaNamer::scan(
            "/word/embeddings",
            "Workbook",
            package.parts.keys().map(String::as_str),
        );
        let chart_part = chart_namer.next_part_name("xml");
        let workbook_part = workbook_namer.next_part_name("xlsx");

        let document_relationship_id = package.get_or_create_part_rels(&self.doc_part_name).add(
            rel_types::CHART,
            &relative_target(&self.doc_part_name, &chart_part),
        );
        let workbook_relationship_id = package.get_or_create_part_rels(&chart_part).add(
            rel_types::PACKAGE,
            &relative_target(&chart_part, &workbook_part),
        );
        let (chart_xml, workbook_bytes) = match source {
            ChartPackageSource::Typed { chart, workbook } => (
                chart_with_workbook_relationship(chart, &workbook_relationship_id)?,
                workbook
                    .to_xlsx_bytes()
                    .map_err(|error| Error::Other(format!("invalid chart workbook: {error}")))?,
            ),
            ChartPackageSource::Authored { kind, data } => {
                oxml_chart::authored_chart_parts(kind, data, &workbook_relationship_id)
                    .map_err(|error| Error::Other(format!("invalid chart data: {error}")))?
            }
        };

        let inline =
            CT_Inline::new_chart(&document_relationship_id, width.to_emu(), height.to_emu());
        let drawing = CT_Drawing::inline(inline);
        let run = CT_R {
            alt_drawings: Vec::new(),
            properties: None,
            content: vec![RunContent::Drawing(drawing)],
            extra_xml: Vec::new(),
            extra_xml_positions: Vec::new(),
        };
        let mut paragraph = CT_P::new();
        paragraph.runs.push(run);
        document
            .body
            .content
            .push(BodyContent::Paragraph(paragraph));
        document.to_xml()?;

        package.set_part(&chart_part, chart_xml);
        package.set_part(&workbook_part, workbook_bytes);
        package
            .content_types
            .add_override(&chart_part, content_types::CHART);
        package
            .content_types
            .add_override(&workbook_part, content_types::EMBEDDED_WORKBOOK);

        self.package = package;
        self.document = document;
        self.invalidate_layout();
        Ok(())
    }

    // ---- Paragraph access ----

    /// Iterate over direct body items in source order.
    ///
    /// Unlike [`Self::paragraphs`] and [`Self::tables`], this retains the
    /// interleaving of paragraphs, tables, content controls, and preserved
    /// unmodelled XML.
    pub fn body_items(&self) -> impl Iterator<Item = BodyItemRef<'_>> {
        self.document.body.content.iter().map(|item| match item {
            BodyContent::Paragraph(paragraph) => {
                BodyItemRef::Paragraph(ParagraphRef { inner: paragraph })
            }
            BodyContent::Table(table) => BodyItemRef::Table(TableRef { inner: table }),
            BodyContent::ContentControl(control) => {
                BodyItemRef::ContentControl(ContentControlRef { inner: control })
            }
            BodyContent::RawXml(raw) => BodyItemRef::UnsupportedXml(raw),
        })
    }

    /// Get immutable references to all paragraphs.
    pub fn paragraphs(&self) -> Vec<ParagraphRef<'_>> {
        self.document
            .body
            .paragraphs()
            .map(|p| ParagraphRef { inner: p })
            .collect()
    }

    /// Get an immutable reference to a paragraph by index (among paragraphs only).
    pub fn paragraph(&self, index: usize) -> Option<ParagraphRef<'_>> {
        self.document
            .body
            .paragraphs()
            .nth(index)
            .map(|p| ParagraphRef { inner: p })
    }

    /// All footnotes as (id, plain text), in file order.
    ///
    /// Separator entries are excluded. They live in the same stream and are
    /// retained by the model so a round trip preserves them, but they are not
    /// notes and never were part of this listing.
    pub fn footnotes(&self) -> Vec<(i32, String)> {
        self.footnotes
            .footnotes
            .iter()
            .filter(|f| f.note_type == rdocx_oxml::footnotes::NoteType::Normal)
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
        self.footnotes_dirty = true;
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
            note_type: rdocx_oxml::footnotes::NoteType::Normal,
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

    /// Return every valid modeled main-document revision in document order.
    pub fn revisions(&self) -> Vec<RevisionRef<'_>> {
        self.document
            .revisions()
            .into_iter()
            .map(|inner| RevisionRef { inner })
            .collect()
    }

    /// Return valid document-protection metadata recorded in the settings part.
    ///
    /// This reports author intent and password-verification metadata. It does
    /// not enforce an access-control boundary or verify a password.
    pub fn document_protection(&self) -> Option<&DocumentProtection> {
        self.settings.as_ref()?.document_protection()
    }

    /// Get the plain text of body paragraphs and table cells in document order.
    pub fn text(&self) -> String {
        let mut result = String::new();
        for content in &self.document.body.content {
            match content {
                BodyContent::Paragraph(paragraph) => {
                    result.push_str(&paragraph.text());
                    result.push('\n');
                }
                BodyContent::Table(table) => {
                    for row in &table.rows {
                        for cell in &row.cells {
                            for content in &cell.content {
                                if let CellContent::Paragraph(paragraph) = content {
                                    result.push_str(&paragraph.text());
                                    result.push('\t');
                                }
                            }
                        }
                        result.push('\n');
                    }
                }
                BodyContent::ContentControl(_) => {}
                BodyContent::RawXml(_) => {}
            }
        }
        result
    }

    /// Get a mutable reference to a paragraph by index (among paragraphs only).
    pub fn paragraph_mut(&mut self, index: usize) -> Option<Paragraph<'_>> {
        self.invalidate_layout();
        let mut remaining = index;
        nth_paragraph_in_body(&mut self.document.body.content, &mut remaining)
            .map(|inner| Paragraph { inner })
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

    /// Get an immutable table by index among tables only.
    pub fn table(&self, index: usize) -> Option<TableRef<'_>> {
        self.document
            .body
            .tables()
            .nth(index)
            .map(|inner| TableRef { inner })
    }

    /// Get a mutable table by index among tables only.
    pub fn table_mut(&mut self, index: usize) -> Option<Table<'_>> {
        self.invalidate_layout();
        let mut remaining = index;
        nth_table_in_body(&mut self.document.body.content, &mut remaining)
            .map(|inner| Table { inner })
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
            extra_xml_positions: Vec::new(),
        };

        let mut p = CT_P::new();
        p.runs.push(run);
        self.document.body.content.push(BodyContent::Paragraph(p));
        match self.document.body.content.last_mut().unwrap() {
            BodyContent::Paragraph(p) => Paragraph { inner: p },
            _ => unreachable!(),
        }
    }

    /// Add an editable inline chart to the document.
    ///
    /// The ChartML cache and embedded workbook are authored from the same
    /// validated data. Invalid dimensions or data leave the document unchanged.
    pub fn add_chart(
        &mut self,
        kind: ChartKind,
        width: Length,
        height: Length,
        data: &ChartData,
    ) -> Result<Paragraph<'_>> {
        self.add_chart_package(ChartPackageSource::Authored { kind, data }, width, height)?;
        let Some(BodyContent::Paragraph(paragraph)) = self.document.body.content.last_mut() else {
            unreachable!("chart package assembly appends one paragraph");
        };
        Ok(Paragraph { inner: paragraph })
    }

    /// Add an inline image at its native size using 72 DPI when none is declared.
    ///
    /// Returns an error without changing the document when the image dimensions
    /// cannot be determined.
    pub fn add_picture_auto(
        &mut self,
        image_data: &[u8],
        image_filename: &str,
    ) -> Result<Paragraph<'_>> {
        let native_size = oxml_media::probe(image_data)
            .and_then(|info| info.native_size(72.0))
            .ok_or_else(|| Error::UnavailableImageDimensions {
                filename: image_filename.to_owned(),
            })?;

        Ok(self.add_picture(
            image_data,
            image_filename,
            Length::emu(native_size.width_emu),
            Length::emu(native_size.height_emu),
        ))
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
            extra_xml_positions: Vec::new(),
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
            extra_xml_positions: Vec::new(),
        };

        let mut p = CT_P::new();
        p.runs.push(run);
        self.document.body.insert_paragraph(0, p);
        match &mut self.document.body.content[0] {
            BodyContent::Paragraph(p) => Paragraph { inner: p },
            _ => unreachable!(),
        }
    }

    /// Store image bytes as a new media part and declare its content type.
    ///
    /// Returns the relationship target to use when referencing it, e.g.
    /// `media/image3.png`. No relationship is created here: an image referenced
    /// from a header or footer must be related to *that* part, not the
    /// document, so the caller decides where it is attached.
    fn store_image_part(&mut self, image_data: &[u8], filename: &str) -> String {
        let format = oxml_media::resolve(image_data, filename);
        let extension = format.extension();
        let part_name = self.image_namer.next_part_name(extension);

        self.package.set_part(&part_name, image_data.to_vec());
        let content_type = format.content_type();
        match self.package.content_types.content_type_for(&part_name) {
            Some(existing) if existing == content_type => {}
            Some(_) => self
                .package
                .content_types
                .add_override(&part_name, content_type),
            None => self
                .package
                .content_types
                .add_default(extension, content_type),
        }

        part_name
            .strip_prefix("/word/")
            .unwrap_or(&part_name)
            .to_owned()
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
        let rel_id = self.add_hyperlink_relationship(url);

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
        crate::Paragraph { inner: p }.add_hyperlink(text, &rel_id);
    }

    /// Add an external hyperlink relationship and return its relationship ID.
    ///
    /// Use this with [`crate::Paragraph::add_hyperlink`] when the target
    /// paragraph is not the last body paragraph, such as a paragraph inside a
    /// table cell.
    pub fn add_hyperlink_relationship(&mut self, url: &str) -> String {
        self.invalidate_layout();
        self.package
            .get_or_create_part_rels(&self.doc_part_name)
            .add_external(rel_types::HYPERLINK, url)
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
        use oxml_opc::relationship::rel_types;
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

    /// Whether the document references any header or footer part.
    ///
    /// This is broader than [`Self::header_text`] and [`Self::footer_text`]:
    /// a referenced part can contain drawings, fields, tables, or other visible
    /// content without contributing literal text.
    pub fn has_header_footer_content(&self) -> bool {
        !self.header_footer_rel_ids().is_empty()
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

    /// Set a Word-compatible text watermark in every active header variant.
    pub fn set_text_watermark(&mut self, text: &str) -> Result<()> {
        let mut candidate = self.clone_for_staging();
        candidate.apply_watermark(|_, _| VmlWatermark::Text {
            text: text.to_owned(),
            width_pt: 468.0,
            height_pt: 117.0,
            rotation_degrees: 315.0,
            color: "D9D9D9".to_owned(),
            font_family: Some("Calibri".to_owned()),
            opacity: 0.5,
        })?;
        self.commit_staged_mutation(candidate);
        Ok(())
    }

    /// Set an image watermark in every active header variant.
    pub fn set_image_watermark(
        &mut self,
        image_data: &[u8],
        image_filename: &str,
        width: Length,
        height: Length,
    ) -> Result<()> {
        if width.to_emu() <= 0 || height.to_emu() <= 0 {
            return Err(Error::Other(
                "watermark image width and height must be positive".to_owned(),
            ));
        }

        let mut candidate = self.clone_for_staging();
        let image_target = candidate.store_image_part(image_data, image_filename);
        let image_part_name =
            OpcPackage::resolve_rel_target(&candidate.doc_part_name, &image_target);
        candidate.apply_watermark(|package, part_name| {
            let target = relative_target(part_name, &image_part_name);
            let relationship_id = package
                .get_or_create_part_rels(part_name)
                .add(rel_types::IMAGE, &target);
            VmlWatermark::Image {
                relationship_id,
                width_pt: width.to_pt(),
                height_pt: height.to_pt(),
                rotation_degrees: 0.0,
                opacity: 0.5,
            }
        })?;
        self.commit_staged_mutation(candidate);
        Ok(())
    }

    fn apply_watermark(
        &mut self,
        mut watermark_for_part: impl FnMut(&mut OpcPackage, &str) -> VmlWatermark,
    ) -> Result<()> {
        self.ensure_watermark_header_inheritance()?;
        let header_ids = self
            .header_footer_rel_ids()
            .into_iter()
            .filter_map(|(relationship_id, is_header)| is_header.then_some(relationship_id))
            .collect::<Vec<_>>();

        for relationship_id in header_ids {
            let target = self
                .package
                .get_part_rels(&self.doc_part_name)
                .and_then(|relationships| relationships.get_by_id(&relationship_id))
                .map(|relationship| relationship.target.clone())
                .ok_or_else(|| {
                    Error::Other(format!(
                        "active header relationship {relationship_id} is missing"
                    ))
                })?;
            let part_name = OpcPackage::resolve_rel_target(&self.doc_part_name, &target);
            let xml = self
                .package
                .get_part(&part_name)
                .map(<[u8]>::to_vec)
                .ok_or_else(|| {
                    Error::Other(format!("active header part {part_name} is missing"))
                })?;
            let watermark = watermark_for_part(&mut self.package, &part_name);
            let updated = replace_authored_watermark(&xml, &watermark)?;
            self.package.set_part(&part_name, updated);
        }
        self.invalidate_layout();
        Ok(())
    }

    fn ensure_watermark_header_inheritance(&mut self) -> Result<()> {
        self.section_properties_mut();
        let even_enabled = self.even_headers_enabled();
        let insertions = {
            let mut effective = [false; 3];
            let mut insertions = Vec::new();
            let mut inspect = |location: Option<usize>, section: &CT_SectPr| {
                for reference in &section.header_refs {
                    effective[header_type_index(reference.hdr_ftr_type)] = true;
                }
                for hdr_type in [HdrFtrType::Default, HdrFtrType::First, HdrFtrType::Even] {
                    let active = match hdr_type {
                        HdrFtrType::Default => true,
                        HdrFtrType::First => section.title_pg.unwrap_or(false),
                        HdrFtrType::Even => even_enabled,
                    };
                    let index = header_type_index(hdr_type);
                    if active && !effective[index] {
                        insertions.push((location, hdr_type));
                        effective[index] = true;
                    }
                }
            };
            for (index, content) in self.document.body.content.iter().enumerate() {
                if let BodyContent::Paragraph(paragraph) = content
                    && let Some(section) = paragraph
                        .properties
                        .as_ref()
                        .and_then(|properties| properties.sect_pr.as_ref())
                {
                    inspect(Some(index), section);
                }
            }
            if let Some(section) = self.document.body.sect_pr.as_ref() {
                inspect(None, section);
            }
            insertions
        };

        for (location, hdr_type) in insertions {
            let relationship_id = self.create_watermark_header_relationship(hdr_type)?;
            let reference = HdrFtrRef {
                hdr_ftr_type: hdr_type,
                rel_id: relationship_id,
            };
            match location {
                Some(index) => {
                    let BodyContent::Paragraph(paragraph) = &mut self.document.body.content[index]
                    else {
                        unreachable!("recorded section owner changed")
                    };
                    paragraph
                        .properties
                        .as_mut()
                        .and_then(|properties| properties.sect_pr.as_mut())
                        .expect("recorded section disappeared")
                        .header_refs
                        .push(reference);
                }
                None => self
                    .document
                    .body
                    .sect_pr
                    .as_mut()
                    .expect("final section disappeared")
                    .header_refs
                    .push(reference),
            }
        }
        Ok(())
    }

    fn create_watermark_header_relationship(&mut self, hdr_type: HdrFtrType) -> Result<String> {
        let label = match hdr_type {
            HdrFtrType::Default => "Default",
            HdrFtrType::First => "First",
            HdrFtrType::Even => "Even",
        };
        let mut index = 1usize;
        let part_name = loop {
            let candidate = format!("/word/headerWatermark{label}{index}.xml");
            if self.package.get_part(&candidate).is_none() {
                break candidate;
            }
            index += 1;
        };
        let empty_header = CT_HdrFtr::new().to_xml_header()?;
        self.package.set_part(&part_name, empty_header);
        self.package.content_types.add_override(
            &part_name,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
        );
        let target = relative_target(&self.doc_part_name, &part_name);
        Ok(self
            .package
            .get_or_create_part_rels(&self.doc_part_name)
            .add(rel_types::HEADER, &target))
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
            extra_xml_positions: Vec::new(),
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

    /// Create a list definition with explicit per-level formats and return
    /// its numId.
    ///
    /// Unlike [`Self::add_bullet_list_item`] / [`Self::add_numbered_list_item`],
    /// which share one bullet and one numbered definition per document, every
    /// call creates a fresh definition — so separate lists restart their
    /// numbering, and one definition can mix formats across levels (e.g. a
    /// bullet list whose nested level is decimal). Attach paragraphs with
    /// [`crate::Paragraph::set_numbering`].
    ///
    /// `levels[i]` configures level `i`; deeper unspecified levels fall back
    /// to the standard template rotation for the last specified format's
    /// family. An empty slice produces the standard numbered template. Word
    /// supports nine levels, so entries after index eight are ignored.
    ///
    /// ```no_run
    /// use rdocx::{Document, ListLevel};
    ///
    /// let mut doc = Document::new();
    /// let num_id = doc.add_list_definition(&[
    ///     ListLevel::bullet(),
    ///     ListLevel::decimal().start(3),
    /// ]);
    /// doc.add_paragraph("first bullet").set_numbering(num_id, 0);
    /// doc.add_paragraph("third decimal").set_numbering(num_id, 1);
    /// ```
    pub fn add_list_definition(&mut self, levels: &[ListLevel]) -> u32 {
        self.invalidate_layout();
        let levels: Vec<(ST_NumberFormat, Option<u32>)> = levels
            .iter()
            .take(9)
            .map(|level| (level.format.to_st(), level.start))
            .collect();
        self.ensure_numbering().add_list(&levels)
    }

    /// Redefine one level (0–8) of an existing list definition, for callers
    /// that only learn a deeper level's format when content first reaches it.
    ///
    /// Returns `false` when `num_id` is unknown or `level` is out of range.
    pub fn set_list_level(&mut self, num_id: u32, level: u32, spec: ListLevel) -> bool {
        let updated = self.numbering.as_mut().is_some_and(|numbering| {
            numbering.set_list_level(num_id, level, spec.format.to_st(), spec.start)
        });
        if updated {
            self.invalidate_layout();
        }
        updated
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
                root_attributes: Vec::new(),
                extra_xml: Vec::new(),
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
            BodyContent::ContentControl(_) => {}
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
                        rdocx_oxml::table::CellContent::ContentControl(_) => {}
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
            bookmark_id: i32,
        }

        // Calling insert_toc twice must not mint bookmarks that collide with
        // the ones the first call left behind — duplicate `w:name` values make
        // the internal links ambiguous. Continue numbering past whatever is
        // already there.
        let mut occupied_suffixes = self.toc_bookmark_suffixes();
        let mut toc_counter = occupied_suffixes.iter().copied().max().unwrap_or(0);
        let mut occupied_ids = self
            .document
            .body
            .content
            .iter()
            .filter_map(|content| match content {
                BodyContent::Paragraph(paragraph) => Some(&paragraph.bookmark_markers),
                _ => None,
            })
            .flatten()
            .filter_map(|marker| marker.id())
            .filter(|id| *id >= 0)
            .collect::<HashSet<_>>();

        let mut headings = Vec::new();

        for (idx, content) in self.document.body.content.iter().enumerate() {
            if let BodyContent::Paragraph(p) = content
                && let Some(level) = Self::detect_heading_level_for_toc(p)
                && level <= max_level
            {
                let text = p.text();
                if !text.trim().is_empty() {
                    let Some(suffix) = next_toc_bookmark_suffix(&occupied_suffixes, toc_counter)
                    else {
                        return;
                    };
                    toc_counter = suffix;
                    occupied_suffixes.insert(suffix);
                    let preferred_id = suffix
                        .checked_add(99)
                        .and_then(|candidate| i32::try_from(candidate).ok())
                        .filter(|candidate| !occupied_ids.contains(candidate));
                    let Some(bookmark_id) = preferred_id.or_else(|| {
                        (0..=i32::MAX).find(|candidate| !occupied_ids.contains(candidate))
                    }) else {
                        return;
                    };
                    occupied_ids.insert(bookmark_id);
                    headings.push(HeadingInfo {
                        content_index: idx,
                        level,
                        text,
                        bookmark_name: format!("_Toc{suffix}"),
                        bookmark_id,
                    });
                }
            }
        }

        // Step 2: Insert typed bookmark markers at each heading paragraph.
        for heading in &headings {
            if let Some(BodyContent::Paragraph(p)) =
                self.document.body.content.get_mut(heading.content_index)
            {
                let run_count = p.runs.len();
                let inserted_start =
                    p.insert_bookmark_start(0, heading.bookmark_id, &heading.bookmark_name);
                let inserted_end = p.insert_bookmark_end(run_count, heading.bookmark_id);
                debug_assert!(inserted_start && inserted_end);
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
                source_occurrence: None,
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
                extra_xml_positions: Vec::new(),
            });

            // Wrap the text run in a hyperlink to the bookmark
            p.hyperlinks.push(HyperlinkSpan {
                rel_id: None,
                anchor: Some(heading.bookmark_name.clone()),
                run_start: 0,
                run_end: 1, // Just the text run, not the tab
                extra_attributes: Vec::new(),
                extra_xml: Vec::new(),
                preserved_raw_before: None,
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

    /// Numeric `_TocN` bookmark suffixes already present in the body.
    fn toc_bookmark_suffixes(&self) -> HashSet<u64> {
        let mut suffixes = HashSet::new();
        for content in &self.document.body.content {
            let BodyContent::Paragraph(p) = content else {
                continue;
            };
            for marker in &p.bookmark_markers {
                let Some(name) = marker.name() else {
                    continue;
                };
                let Some(after) = name.strip_prefix("_Toc") else {
                    continue;
                };
                if let Ok(suffix) = after.parse::<u64>() {
                    suffixes.insert(suffix);
                }
            }
        }
        suffixes
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

    /// Render scalar and structural template tags from structured JSON data.
    ///
    /// Tags use `{{ path.to.value }}` syntax and may cross ordinary Word run
    /// boundaries. String, number and boolean leaves render as text, while
    /// `null` renders as an empty string. Missing paths, malformed tags, and
    /// object or array leaves return an error without changing the document.
    /// Dedicated main-body paragraphs and table rows may contain nested
    /// `{% for item in path %}` or `{% if path %}` blocks with their matching
    /// end markers. Loop paths require arrays and introduce lexical scopes.
    /// Conditions use JSON truthiness, where false, null, zero, and empty
    /// strings or collections are false.
    ///
    /// This additive API is native-only. Python, WASM, and CLI binding surfaces
    /// remain unchanged and continue to preserve documents rendered here.
    pub fn render_template(&mut self, data: &serde_json::Value) -> Result<usize> {
        crate::template::render(self, data)
    }

    pub(crate) fn template_numbering_reference_exists(&self, num_id: u32, level: u32) -> bool {
        self.numbering
            .as_ref()
            .and_then(|numbering| numbering.get_abstract_num_for(num_id))
            .is_some_and(|abstract_numbering| {
                abstract_numbering
                    .levels
                    .iter()
                    .any(|candidate| candidate.ilvl == level)
            })
    }

    pub(crate) fn template_sources(&mut self) -> Result<Vec<String>> {
        self.flush_to_package()?;
        let mut sources = crate::template::body_sources(&self.document);

        for (rel_id, _) in self.header_footer_rel_ids() {
            if let Some(header_footer) = self.load_header_footer(&rel_id) {
                sources.extend(crate::template::header_footer_sources(&header_footer));
            }
        }

        for part_name in self.raw_text_bearing_part_names() {
            if let Some(xml) = self.package.get_part(&part_name) {
                sources.extend(crate::template::text_box_sources(xml)?);
            }
        }

        for part_name in self.chart_part_names() {
            if let Some(xml) = self.package.get_part(&part_name) {
                sources.extend(crate::template::chart_sources(xml)?);
            }
        }

        Ok(sources)
    }

    pub(crate) fn apply_template_pairs(&mut self, pairs: &[(&str, &str)]) -> usize {
        self.replace_batch(pairs)
    }

    pub(crate) fn commit_template(&mut self, candidate: Self) {
        self.package = candidate.package;
        self.document = candidate.document;
        self.invalidate_layout();
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
                BodyContent::ContentControl(_) => {}
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

    /// Relationship IDs of every section's headers and footers, with a flag
    /// saying which kind each one is.
    fn header_footer_rel_ids(&self) -> Vec<(String, bool)> {
        let mut rel_ids = Vec::new();
        let mut seen = HashSet::new();
        let sections = self
            .document
            .body
            .content
            .iter()
            .filter_map(|content| match content {
                BodyContent::Paragraph(paragraph) => paragraph
                    .properties
                    .as_ref()
                    .and_then(|properties| properties.sect_pr.as_ref()),
                BodyContent::Table(_) | BodyContent::ContentControl(_) | BodyContent::RawXml(_) => {
                    None
                }
            })
            .chain(self.document.body.sect_pr.iter());

        for section in sections {
            for reference in &section.header_refs {
                let rel_id = (reference.rel_id.clone(), true);
                if seen.insert(rel_id.clone()) {
                    rel_ids.push(rel_id);
                }
            }
            for reference in &section.footer_refs {
                let rel_id = (reference.rel_id.clone(), false);
                if seen.insert(rel_id.clone()) {
                    rel_ids.push(rel_id);
                }
            }
        }

        rel_ids
    }

    fn header_footer_rel_ids_for_layout(&self) -> HashSet<String> {
        let even_headers_enabled = self.even_headers_enabled();
        let mut rel_ids = HashSet::new();
        let sections = self
            .document
            .body
            .content
            .iter()
            .filter_map(|content| match content {
                BodyContent::Paragraph(paragraph) => paragraph
                    .properties
                    .as_ref()
                    .and_then(|properties| properties.sect_pr.as_ref()),
                BodyContent::Table(_) | BodyContent::ContentControl(_) | BodyContent::RawXml(_) => {
                    None
                }
            })
            .chain(self.document.body.sect_pr.iter());
        for section in sections {
            for reference in section.header_refs.iter().chain(&section.footer_refs) {
                if reference.hdr_ftr_type != HdrFtrType::Even || even_headers_enabled {
                    rel_ids.insert(reference.rel_id.clone());
                }
            }
        }
        rel_ids
    }

    fn even_headers_enabled(&self) -> bool {
        let Some(settings) = self.settings.as_ref() else {
            return false;
        };
        settings
            .to_xml()
            .ok()
            .is_some_and(|xml| settings_enable_even_headers(&xml))
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
                BodyContent::ContentControl(_) => {}
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

    fn raw_text_bearing_part_names(&self) -> Vec<String> {
        let mut names = vec![self.doc_part_name.clone()];
        if let Some(section) = self.document.body.sect_pr.as_ref()
            && let Some(rels) = self.package.get_part_rels(&self.doc_part_name)
        {
            for reference in section.header_refs.iter().chain(&section.footer_refs) {
                if let Some(relationship) = rels.get_by_id(&reference.rel_id) {
                    names.push(OpcPackage::resolve_rel_target(
                        &self.doc_part_name,
                        &relationship.target,
                    ));
                }
            }
        }
        names
    }

    fn chart_part_names(&self) -> Vec<String> {
        self.package
            .get_part_rels(&self.doc_part_name)
            .map(|relationships| {
                relationships
                    .get_all_by_type(rel_types::CHART)
                    .iter()
                    .map(|relationship| {
                        OpcPackage::resolve_rel_target(&self.doc_part_name, &relationship.target)
                    })
                    .collect()
            })
            .unwrap_or_default()
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
        for part_name in self.raw_text_bearing_part_names() {
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
        for part_name in self.chart_part_names() {
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

    // ---- Layout and PDF conversion ----

    /// Return the cached normal-font layout with its Word source map.
    ///
    /// Repeated calls share the same accepted-view result until the document
    /// is mutated.
    pub fn layout(&self) -> Result<Arc<rdocx_layout::WordLayoutResult>> {
        self.layout_with_options(RenderOptions::default())
    }

    /// Return a normal-font layout with the selected revision view.
    ///
    /// Accepted-view calls share the normal layout cache. Tracked-view calls
    /// remain uncached because they do not replace the accepted-view cache.
    pub fn layout_with_options(
        &self,
        options: RenderOptions,
    ) -> Result<Arc<rdocx_layout::WordLayoutResult>> {
        self.layout_for_options(options, false)
    }

    /// Return an uncached layout using user-provided font files.
    ///
    /// User-provided fonts take highest priority in font resolution. The
    /// returned owned result retains the exact font bytes and Word source map
    /// used by layout.
    pub fn layout_with_fonts(
        &self,
        font_files: &[(&str, &[u8])],
    ) -> Result<rdocx_layout::WordLayoutResult> {
        self.layout_with_fonts_and_options(font_files, RenderOptions::default())
    }

    /// Return an uncached caller-font layout with the selected revision view.
    pub fn layout_with_fonts_and_options(
        &self,
        font_files: &[(&str, &[u8])],
        options: RenderOptions,
    ) -> Result<rdocx_layout::WordLayoutResult> {
        let mut input = self.build_layout_input();
        input.revision_view = options.revision_view;
        for (family, data) in font_files {
            input.fonts.push(rdocx_layout::FontFile {
                family: family.to_string(),
                data: data.to_vec(),
            });
        }
        #[cfg(test)]
        record_layout_invocation();
        Ok(rdocx_layout::layout_document_with_caller_fonts_and_provenance(&input)?)
    }

    /// SVG PoC patch: like [`Self::layout_with_fonts`], but the injected
    /// fonts sit on top of the bundled metric-compatible set, so a wasm
    /// editor that only injects a Korean font still resolves Calibri/Times.
    /// Separate from the strictly isolated caller-fonts path on purpose.
    pub fn layout_with_fonts_and_bundled_fallback(
        &self,
        font_files: &[(&str, &[u8])],
    ) -> Result<rdocx_layout::WordLayoutResult> {
        let mut input = self.build_layout_input();
        for (family, data) in font_files {
            input.fonts.push(rdocx_layout::FontFile {
                family: family.to_string(),
                data: data.to_vec(),
            });
        }
        let mut engine = self
            .fallback_layout_engine
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner);
        let engine = match engine.as_mut() {
            Some(engine) => engine,
            None => {
                *engine = Some(rdocx_layout::engine::Engine::new_deterministic()?);
                engine.as_mut().expect("just inserted")
            }
        };
        Ok(rdocx_layout::layout_document_with_fallback_fonts_engine(
            engine, &input,
        )?)
    }

    /// SVG PoC patch: hand the bundled-fallback engine (and its content-keyed
    /// relayout caches) to a caller restoring an undo snapshot, so the
    /// rebuilt Document does not go cache-cold. Interim shape until the
    /// upstream F-X039 session/handle design lands.
    pub fn take_layout_engine(&self) -> Option<rdocx_layout::engine::Engine> {
        self.fallback_layout_engine
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner)
            .take()
    }

    /// SVG PoC patch: counterpart of [`Self::take_layout_engine`].
    pub fn set_layout_engine(&self, engine: rdocx_layout::engine::Engine) {
        *self
            .fallback_layout_engine
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner) = Some(engine);
    }

    /// Render the document to PDF bytes.
    ///
    /// This performs a full layout pass (font shaping, line breaking, pagination)
    /// and then renders the result to a PDF document.
    ///
    /// Font resolution order:
    /// 1. Fonts embedded in the DOCX file (word/fonts/)
    /// 2. System fonts when the default `system-fonts` feature is enabled
    /// 3. Always-available bundled metric-compatible fonts
    pub fn to_pdf(&self) -> Result<Vec<u8>> {
        self.to_pdf_with_options(RenderOptions::default())
    }

    /// Render the document to PDF bytes with the selected revision view.
    pub fn to_pdf_with_options(&self, options: RenderOptions) -> Result<Vec<u8>> {
        let layout = self.layout_with_options(options)?;
        Ok(oxml_pdf::render_to_pdf(&layout.layout))
    }

    /// Render the document to PDF bytes using bundled fonts without system
    /// font discovery.
    ///
    /// The deterministic layout is cached independently from the normal-font
    /// layout and is suitable for reproducible render baselines.
    pub fn to_pdf_deterministic(&self) -> Result<Vec<u8>> {
        self.to_pdf_deterministic_with_options(RenderOptions::default())
    }

    /// Render the selected revision view to deterministic PDF bytes.
    pub fn to_pdf_deterministic_with_options(&self, options: RenderOptions) -> Result<Vec<u8>> {
        let layout = self.layout_for_options(options, true)?;
        Ok(oxml_pdf::render_to_pdf(&layout.layout))
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
    /// 3. System fonts when the default `system-fonts` feature is enabled
    /// 4. Always-available bundled metric-compatible fonts
    pub fn to_pdf_with_fonts(&self, font_files: &[(&str, &[u8])]) -> Result<Vec<u8>> {
        self.to_pdf_with_fonts_and_options(font_files, RenderOptions::default())
    }

    /// Render the selected revision view to PDF with user-provided fonts.
    pub fn to_pdf_with_fonts_and_options(
        &self,
        font_files: &[(&str, &[u8])],
        options: RenderOptions,
    ) -> Result<Vec<u8>> {
        let layout = self.layout_with_fonts_and_options(font_files, options)?;
        Ok(oxml_pdf::render_to_pdf(&layout.layout))
    }

    /// Save the document as a PDF file.
    pub fn save_pdf<P: AsRef<Path>>(&self, path: P) -> Result<()> {
        self.save_pdf_with_options(path, RenderOptions::default())
    }

    /// Save the selected revision view as a PDF file.
    pub fn save_pdf_with_options<P: AsRef<Path>>(
        &self,
        path: P,
        options: RenderOptions,
    ) -> Result<()> {
        let pdf_bytes = self.to_pdf_with_options(options)?;
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
        use oxml_opc::relationship::rel_types;
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
                            let content_type = oxml_media::resolve(data, &part_name)
                                .content_type()
                                .to_owned();
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
        self.render_page_to_png_with_options(page_index, dpi, RenderOptions::default())
    }

    /// Render one page to PNG with the selected revision view.
    pub fn render_page_to_png_with_options(
        &self,
        page_index: usize,
        dpi: f64,
        options: RenderOptions,
    ) -> Result<Option<Vec<u8>>> {
        let layout = self.layout_with_options(options)?;
        Ok(oxml_pdf::render_page_to_png(
            &layout.layout,
            page_index,
            dpi,
        ))
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
        self.render_page_to_png_deterministic_with_options(
            page_index,
            dpi,
            RenderOptions::default(),
        )
    }

    /// Render one page to deterministic PNG with the selected revision view.
    pub fn render_page_to_png_deterministic_with_options(
        &self,
        page_index: usize,
        dpi: f64,
        options: RenderOptions,
    ) -> Result<Option<Vec<u8>>> {
        let layout = self.layout_for_options(options, true)?;
        Ok(oxml_pdf::render_page_to_png(
            &layout.layout,
            page_index,
            dpi,
        ))
    }

    /// Render all pages of the document to PNG bytes.
    pub fn render_all_pages(&self, dpi: f64) -> Result<Vec<Vec<u8>>> {
        self.render_all_pages_with_options(dpi, RenderOptions::default())
    }

    /// Render every page to PNG with the selected revision view.
    pub fn render_all_pages_with_options(
        &self,
        dpi: f64,
        options: RenderOptions,
    ) -> Result<Vec<Vec<u8>>> {
        let layout = self.layout_with_options(options)?;
        Ok(oxml_pdf::render_all_pages(&layout.layout, dpi))
    }

    /// Return a cloned positioned page from the cached normal-font layout.
    ///
    /// `page_index` is zero-based. An index beyond the document returns `None`.
    pub fn layout_page(
        &self,
        page_index: usize,
    ) -> Result<Option<std::sync::Arc<oxml_layout::PageFrame>>> {
        self.layout_page_with_options(page_index, RenderOptions::default())
    }

    /// Return one positioned page from the selected revision view.
    pub fn layout_page_with_options(
        &self,
        page_index: usize,
        options: RenderOptions,
    ) -> Result<Option<std::sync::Arc<oxml_layout::PageFrame>>> {
        let layout = self.layout_with_options(options)?;
        Ok(layout.layout.pages.get(page_index).cloned())
    }

    /// Build a LayoutInput from the document's current state.
    fn build_layout_input(&self) -> rdocx_layout::LayoutInput {
        use oxml_opc::relationship::rel_types;
        use rdocx_layout::{ImageData, LayoutInput};
        use std::collections::HashMap;

        let mut headers: HashMap<String, CT_HdrFtr> = HashMap::new();
        let mut footers: HashMap<String, CT_HdrFtr> = HashMap::new();
        let mut images: HashMap<String, ImageData> = HashMap::new();
        let mut charts = HashMap::new();
        let mut hyperlink_urls: HashMap<String, String> = HashMap::new();
        let mut footnotes = None;
        let mut endnotes = None;
        let mut theme_part_name = None;
        let active_header_footer_ids = self.header_footer_rel_ids_for_layout();
        let even_headers_enabled = self.even_headers_enabled();

        // Extract embedded fonts from the DOCX package
        let fonts = self.extract_embedded_fonts();

        if let Some(rels) = self.package.get_part_rels(&self.doc_part_name) {
            for rel in &rels.items {
                match rel.rel_type.as_str() {
                    t if t == rel_types::HEADER => {
                        if !active_header_footer_ids.contains(&rel.id) {
                            continue;
                        }
                        let part_name =
                            OpcPackage::resolve_rel_target(&self.doc_part_name, &rel.target);
                        if let Some(xml) = self.package.get_part(&part_name)
                            && let Ok(hf) = CT_HdrFtr::from_xml(xml)
                        {
                            headers.insert(rel.id.clone(), hf);
                        }
                        if let Some(header_relationships) = self.package.get_part_rels(&part_name) {
                            for image_relationship in
                                header_relationships.items.iter().filter(|item| {
                                    item.rel_type == rel_types::IMAGE
                                        && item.target_mode.as_deref() != Some("External")
                                })
                            {
                                let image_part = OpcPackage::resolve_rel_target(
                                    &part_name,
                                    &image_relationship.target,
                                );
                                if let Some(data) = self.package.get_part(&image_part) {
                                    images.insert(
                                        format!("{}\0{}", rel.id, image_relationship.id),
                                        ImageData {
                                            data: data.to_vec(),
                                            content_type: oxml_media::resolve(data, &image_part)
                                                .content_type()
                                                .to_owned(),
                                        },
                                    );
                                }
                            }
                        }
                    }
                    t if t == rel_types::FOOTER => {
                        if !active_header_footer_ids.contains(&rel.id) {
                            continue;
                        }
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
                            let content_type = oxml_media::resolve(data, &part_name)
                                .content_type()
                                .to_owned();
                            images.insert(
                                rel.id.clone(),
                                ImageData {
                                    data: data.to_vec(),
                                    content_type,
                                },
                            );
                        }
                    }
                    t if t == rel_types::CHART => {
                        let chart = if rel.target_mode.as_deref() == Some("External") {
                            Err(format!("external target {}", rel.target))
                        } else {
                            let part_name =
                                OpcPackage::resolve_rel_target(&self.doc_part_name, &rel.target);
                            match self.package.get_part(&part_name) {
                                Some(xml) => {
                                    CT_ChartSpace::from_xml(xml).map(Box::new).map_err(|error| {
                                        format!("malformed target {part_name}: {error}")
                                    })
                                }
                                None => Err(format!("missing target {part_name}")),
                            }
                        };
                        charts.insert(rel.id.clone(), chart);
                    }
                    t if t == rel_types::THEME => {
                        if rel.target_mode.as_deref() != Some("External") {
                            theme_part_name = Some(OpcPackage::resolve_rel_target(
                                &self.doc_part_name,
                                &rel.target,
                            ));
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
        let theme_xml = theme_part_name
            .as_deref()
            .and_then(|part_name| self.package.get_part(part_name));
        let chart_theme = theme_xml
            .and_then(|data| oxml_drawing::theme::CT_OfficeStyleSheet::from_xml(data).ok())
            .unwrap_or_else(oxml_drawing::theme::CT_OfficeStyleSheet::office_default);
        let theme = theme_xml.and_then(|data| rdocx_oxml::theme::Theme::from_xml(data).ok());

        let mut document = self.document.clone();
        if !even_headers_enabled {
            for content in &mut document.body.content {
                if let BodyContent::Paragraph(paragraph) = content
                    && let Some(section) = paragraph
                        .properties
                        .as_mut()
                        .and_then(|properties| properties.sect_pr.as_mut())
                {
                    section
                        .header_refs
                        .retain(|reference| reference.hdr_ftr_type != HdrFtrType::Even);
                    section
                        .footer_refs
                        .retain(|reference| reference.hdr_ftr_type != HdrFtrType::Even);
                }
            }
            if let Some(section) = document.body.sect_pr.as_mut() {
                section
                    .header_refs
                    .retain(|reference| reference.hdr_ftr_type != HdrFtrType::Even);
                section
                    .footer_refs
                    .retain(|reference| reference.hdr_ftr_type != HdrFtrType::Even);
            }
        }
        materialize_header_inheritance(&mut document, even_headers_enabled);

        LayoutInput {
            revision_view: rdocx_layout::RevisionView::Accepted,
            document,
            styles: self.styles.clone(),
            numbering: self.numbering.clone(),
            headers,
            footers,
            images,
            charts,
            chart_theme,
            chart_color_map: oxml_drawing::color::ColorMap::default(),
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
            BodyContent::ContentControl(_) => {}
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
                        CellContent::ContentControl(_) => {}
                    }
                }
            }
        }
    }

    /// Get information about all hyperlinks in the document.
    ///
    /// Resolves hyperlink relationship IDs to their target URLs where possible.
    pub fn links(&self) -> Vec<LinkInfo> {
        use oxml_opc::relationship::rel_types;

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
                for field in p.complex_field_hyperlinks() {
                    let start = field.run_start.min(p.runs.len());
                    let end = field.run_end.clamp(start, p.runs.len());
                    let text: String = p.runs[start..end].iter().map(|run| run.text()).collect();
                    result.push(LinkInfo {
                        text,
                        url: Some(field.target),
                        anchor: None,
                        rel_id: None,
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
            BodyContent::ContentControl(_) => 0,
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
                        CellContent::ContentControl(_) => {}
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

fn next_toc_bookmark_suffix(occupied: &HashSet<u64>, after: u64) -> Option<u64> {
    after
        .checked_add(1)
        .filter(|candidate| !occupied.contains(candidate))
        .or_else(|| (1..=u64::MAX).find(|candidate| !occupied.contains(candidate)))
}

fn chart_with_workbook_relationship(
    chart: &CT_ChartSpace,
    relationship_id: &str,
) -> Result<Vec<u8>> {
    const END: &[u8] = b"</c:chartSpace>";
    let mut xml = chart
        .to_xml()
        .map_err(|error| Error::Other(format!("invalid chart part: {error}")))?;
    if chart_has_external_data(&xml)? {
        return Err(Error::Other(
            "chart already contains an external workbook relationship".to_owned(),
        ));
    }
    if !xml.ends_with(END) {
        return Err(Error::Other(
            "serialized chart lacks the chartSpace closing element".to_owned(),
        ));
    }

    xml.truncate(xml.len() - END.len());
    xml.extend_from_slice(
        format!(
            r#"<c:externalData r:id="{relationship_id}"><c:autoUpdate val="0"/></c:externalData>"#
        )
        .as_bytes(),
    );
    xml.extend_from_slice(END);
    CT_ChartSpace::from_xml(&xml)
        .and_then(|validated| validated.to_xml())
        .map_err(|error| Error::Other(format!("invalid chart part: {error}")))
}

fn chart_has_external_data(xml: &[u8]) -> Result<bool> {
    let mut reader = NsReader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut chart_space_depth = None;
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("invalid chart part XML: {error}")))?;
        match event {
            Event::Start(ref element) => {
                if chart_space_depth.is_none()
                    && chart_namespace_matches(&namespace)
                    && matches_local_name(element.name().as_ref(), b"chartSpace")
                {
                    chart_space_depth = Some(depth);
                } else if chart_space_depth.is_some_and(|root| depth == root + 1)
                    && chart_namespace_matches(&namespace)
                    && matches_local_name(element.name().as_ref(), b"externalData")
                {
                    return Ok(true);
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| Error::Other("chart part XML depth overflow".to_owned()))?;
            }
            Event::Empty(ref element)
                if chart_space_depth.is_some_and(|root| depth == root + 1)
                    && chart_namespace_matches(&namespace)
                    && matches_local_name(element.name().as_ref(), b"externalData") =>
            {
                return Ok(true);
            }
            Event::End(_) => depth = depth.saturating_sub(1),
            Event::Eof => return Ok(false),
            _ => {}
        }
        buffer.clear();
    }
}

fn chart_namespace_matches(namespace: &ResolveResult<'_>) -> bool {
    match namespace {
        ResolveResult::Bound(Namespace(uri)) => *uri == oxml_chart::C_NS.as_bytes(),
        ResolveResult::Unknown(prefix) => *prefix == b"c",
        ResolveResult::Unbound => false,
    }
}

fn settings_enable_even_headers(xml: &[u8]) -> bool {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    loop {
        let Ok((namespace, event)) = reader.read_resolved_event_into(&mut buffer) else {
            return false;
        };
        match event {
            Event::Start(ref element) => {
                if depth == 1
                    && matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == rdocx_oxml::namespace::W_NS.as_bytes())
                    && matches_local_name(element.name().as_ref(), b"evenAndOddHeaders")
                {
                    return word_on_off_value(&reader, element);
                }
                depth += 1;
            }
            Event::Empty(ref element)
                if depth == 1
                    && matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == rdocx_oxml::namespace::W_NS.as_bytes())
                    && matches_local_name(element.name().as_ref(), b"evenAndOddHeaders") =>
            {
                return word_on_off_value(&reader, element);
            }
            Event::End(_) => depth = depth.saturating_sub(1),
            Event::Eof => return false,
            _ => {}
        }
        buffer.clear();
    }
}

fn header_type_index(hdr_type: HdrFtrType) -> usize {
    match hdr_type {
        HdrFtrType::Default => 0,
        HdrFtrType::First => 1,
        HdrFtrType::Even => 2,
    }
}

fn materialize_header_inheritance(document: &mut CT_Document, even_headers_enabled: bool) {
    let mut effective: [Option<HdrFtrRef>; 3] = [None, None, None];
    let inherit = |section: &mut CT_SectPr, effective: &mut [Option<HdrFtrRef>; 3]| {
        for hdr_type in [HdrFtrType::Default, HdrFtrType::First, HdrFtrType::Even] {
            let index = header_type_index(hdr_type);
            if let Some(reference) = section
                .header_refs
                .iter()
                .find(|reference| reference.hdr_ftr_type == hdr_type)
                .cloned()
            {
                effective[index] = Some(reference);
            } else if let Some(reference) = effective[index].clone() {
                section.header_refs.push(reference);
            } else if hdr_type == HdrFtrType::Even && even_headers_enabled {
                section.header_refs.push(HdrFtrRef {
                    hdr_ftr_type: HdrFtrType::Even,
                    rel_id: String::new(),
                });
            }
        }
    };
    for content in &mut document.body.content {
        if let BodyContent::Paragraph(paragraph) = content
            && let Some(section) = paragraph
                .properties
                .as_mut()
                .and_then(|properties| properties.sect_pr.as_mut())
        {
            inherit(section, &mut effective);
        }
    }
    if let Some(section) = document.body.sect_pr.as_mut() {
        inherit(section, &mut effective);
    }
}

fn word_on_off_value(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
) -> bool {
    element
        .attributes()
        .flatten()
        .find_map(|attribute| {
            let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
            if matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == rdocx_oxml::namespace::W_NS.as_bytes())
                && local.as_ref() == b"val"
            {
                attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, element.decoder())
                    .ok()
                    .map(|value| !matches!(value.as_ref(), "0" | "false" | "off"))
            } else {
                None
            }
        })
        .unwrap_or(true)
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

/// Numbering format for one level of a custom list definition.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ListNumberFormat {
    Bullet,
    Decimal,
    LowerLetter,
    UpperLetter,
    LowerRoman,
    UpperRoman,
    Ordinal,
}

impl ListNumberFormat {
    fn to_st(self) -> ST_NumberFormat {
        match self {
            Self::Bullet => ST_NumberFormat::Bullet,
            Self::Decimal => ST_NumberFormat::Decimal,
            Self::LowerLetter => ST_NumberFormat::LowerLetter,
            Self::UpperLetter => ST_NumberFormat::UpperLetter,
            Self::LowerRoman => ST_NumberFormat::LowerRoman,
            Self::UpperRoman => ST_NumberFormat::UpperRoman,
            Self::Ordinal => ST_NumberFormat::Ordinal,
        }
    }
}

/// One level of a custom list definition for [`Document::add_list_definition`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ListLevel {
    /// Numbering format for this level.
    pub format: ListNumberFormat,
    /// Starting number (defaults to 1; ignored for bullet levels).
    pub start: Option<u32>,
}

impl ListLevel {
    /// A level with the given format, starting at 1.
    pub fn new(format: ListNumberFormat) -> Self {
        ListLevel {
            format,
            start: None,
        }
    }

    /// A bullet level.
    pub fn bullet() -> Self {
        Self::new(ListNumberFormat::Bullet)
    }

    /// A decimal-numbered level.
    pub fn decimal() -> Self {
        Self::new(ListNumberFormat::Decimal)
    }

    /// Override the starting number for this level.
    pub fn start(mut self, start: u32) -> Self {
        self.start = Some(start);
        self
    }
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
    use crate::FieldEvaluationContext;
    use crate::paragraph::Alignment;
    use oxml_chart::{
        Axis, AxisData, AxisId, AxisKind, AxisPosition, BarDirection, BarGrouping, CT_PlotArea,
        ChartData, ChartKind, NumericData, Plot, Series, StringRef,
    };
    use oxml_sml::Column;
    use rdocx_oxml::text::Field;
    use rdocx_oxml::units::{HalfPoint, Twips};
    use std::fs;
    use std::io::Cursor;
    use std::process::Command;

    const WORD_VERSION: &str = "16.104";
    const WORD_BUILD: &str = "16.104.25121423";
    const WORD_CHART_CANDIDATE_SHA256: &str =
        "79e9b9ff9e7557dbd09a365bb8c189806e700ed48ca768b27d7158cf2b41370b";

    fn minimal_chart_workbook() -> Workbook {
        Workbook::new(
            "Sheet1",
            vec![
                Column::Text {
                    header: "Category".to_owned(),
                    values: vec!["North".to_owned(), "South".to_owned(), "West".to_owned()],
                },
                Column::Number {
                    header: "Revenue".to_owned(),
                    values: vec![12.5, 19.0, 14.25],
                    number_format: None,
                },
            ],
        )
        .expect("valid chart workbook")
    }

    fn minimal_typed_chart() -> CT_ChartSpace {
        let values = NumericData::new(
            "Sheet1!$B$2:$B$4".to_owned(),
            "General".to_owned(),
            vec![12.5, 19.0, 14.25],
        )
        .expect("valid values");
        let mut series = Series::new(0, 0, values);
        series.name = Some(
            StringRef::new("Sheet1!$B$1".to_owned(), vec!["Revenue".to_owned()])
                .expect("valid series name"),
        );
        series.categories = Some(AxisData::String(
            StringRef::new(
                "Sheet1!$A$2:$A$4".to_owned(),
                vec!["North".to_owned(), "South".to_owned(), "West".to_owned()],
            )
            .expect("valid categories"),
        ));
        let category_axis = AxisId::new(48_650_112).expect("valid category axis ID");
        let value_axis = AxisId::new(48_672_768).expect("valid value axis ID");
        let plot = Plot::bar(
            BarDirection::Column,
            BarGrouping::Clustered,
            vec![series],
            [category_axis, value_axis],
        )
        .expect("valid bar plot");
        let axes = vec![
            Axis::new(
                AxisKind::Category,
                category_axis,
                AxisPosition::Bottom,
                value_axis,
            ),
            Axis::new(
                AxisKind::Value,
                value_axis,
                AxisPosition::Left,
                category_axis,
            ),
        ];
        let mut chart = CT_ChartSpace::from_xml(
            format!(
                r#"<c:chartSpace xmlns:c="{}" xmlns:a="{}" xmlns:r="{}"><c:chart><c:plotArea/></c:chart></c:chartSpace>"#,
                oxml_chart::C_NS,
                oxml_chart::A_NS,
                oxml_chart::R_NS,
            )
            .as_bytes(),
        )
        .expect("valid chart shell");
        chart.chart.auto_title_deleted = true;
        chart.chart.plot_area = CT_PlotArea::new(vec![plot], axes).expect("valid plot area");
        chart
    }

    fn document_with_minimal_chart() -> Document {
        let mut document = Document::new();
        document
            .add_chart_package(
                ChartPackageSource::Typed {
                    chart: &minimal_typed_chart(),
                    workbook: &minimal_chart_workbook(),
                },
                Length::inches(5.0),
                Length::inches(3.0),
            )
            .expect("assemble chart package");
        document
    }

    fn reset_layout_invocations() {
        LAYOUT_INVOCATIONS.set(0);
    }

    fn layout_invocations() -> usize {
        LAYOUT_INVOCATIONS.get()
    }

    fn caller_only_font() -> (&'static str, Vec<u8>) {
        const FAMILY: &str = "Callira";
        const SOURCE: &[u8] =
            include_bytes!("../../oxml-layout/fonts/Carlito-Regular.ttf").as_slice();

        fn replace_all_same_length(data: &mut [u8], from: &[u8], to: &[u8]) -> usize {
            assert_eq!(from.len(), to.len());
            let mut replaced = 0;
            let mut offset = 0;
            while let Some(index) = data[offset..]
                .windows(from.len())
                .position(|window| window == from)
            {
                let start = offset + index;
                data[start..start + from.len()].copy_from_slice(to);
                offset = start + from.len();
                replaced += 1;
            }
            replaced
        }

        let mut bytes = SOURCE.to_vec();
        let ascii = replace_all_same_length(&mut bytes, b"Carlito", FAMILY.as_bytes());
        let source_utf16 = "Carlito"
            .encode_utf16()
            .flat_map(u16::to_be_bytes)
            .collect::<Vec<_>>();
        let family_utf16 = FAMILY
            .encode_utf16()
            .flat_map(u16::to_be_bytes)
            .collect::<Vec<_>>();
        let utf16 = replace_all_same_length(&mut bytes, &source_utf16, &family_utf16);
        assert!(ascii > 0 && utf16 > 0, "font family records were renamed");
        assert!(
            oxml_layout::bundled_fonts::bundled_font_data()
                .iter()
                .all(|(family, bundled)| *family != FAMILY && *bundled != bytes),
            "the caller-only font must not match bundled family names or bytes"
        );
        (FAMILY, bytes)
    }

    #[test]
    fn body_items_preserve_paragraph_table_control_and_raw_order() {
        let mut doc = Document::new();
        doc.add_paragraph("first");
        doc.add_table(1, 1);
        doc.document
            .body
            .content
            .push(BodyContent::RawXml(b"<w:custom/>".to_vec()));

        let mut reader = quick_xml::Reader::from_reader(
            br#"<w:sdt xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:sdtContent><w:p><w:r><w:t>inside</w:t></w:r></w:p></w:sdtContent></w:sdt>"#
                .as_slice(),
        );
        let mut buffer = Vec::new();
        let control = match reader.read_event_into(&mut buffer).unwrap() {
            Event::Start(start) => CT_Sdt::from_xml(&mut reader, &start).unwrap(),
            event => panic!("expected content control start, got {event:?}"),
        };
        doc.document
            .body
            .content
            .push(BodyContent::ContentControl(control));
        doc.add_paragraph("last");

        let items = doc
            .body_items()
            .map(|item| match item {
                BodyItemRef::Paragraph(paragraph) => format!("paragraph:{}", paragraph.text()),
                BodyItemRef::Table(_) => "table".to_owned(),
                BodyItemRef::ContentControl(control) => format!("control:{}", control.text()),
                BodyItemRef::UnsupportedXml(raw) => {
                    format!("raw:{}", std::str::from_utf8(raw).unwrap())
                }
            })
            .collect::<Vec<_>>();

        assert_eq!(
            items,
            [
                "paragraph:first",
                "table",
                "raw:<w:custom/>",
                "control:inside",
                "paragraph:last",
            ]
        );
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
    fn tracked_option_layouts_are_not_cached() {
        let mut doc = Document::new();
        doc.add_paragraph("tracked view");
        let options = RenderOptions {
            revision_view: rdocx_layout::RevisionView::Tracked,
        };

        reset_layout_invocations();
        doc.render_page_to_png_deterministic_with_options(0, 1.0, options)
            .unwrap();
        doc.render_page_to_png_deterministic_with_options(0, 1.0, options)
            .unwrap();
        assert_eq!(layout_invocations(), 2);
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
    fn field_update_batch_invalidates_cached_layout_once() {
        let mut doc = Document::new();
        let mut paragraph = CT_P::new();
        paragraph.runs.push(CT_R {
            properties: None,
            content: vec![RunContent::Field(Field::new("PAGE", "4"))],
            extra_xml: Vec::new(),
            extra_xml_positions: Vec::new(),
            alt_drawings: Vec::new(),
        });
        doc.document
            .body
            .content
            .push(BodyContent::Paragraph(paragraph));

        reset_layout_invocations();
        doc.render_page_to_png_deterministic(0, 1.0).unwrap();
        doc.render_page_to_png_deterministic(0, 1.0).unwrap();
        assert_eq!(layout_invocations(), 1);

        assert_eq!(
            doc.update_fields(&FieldEvaluationContext::default())
                .unwrap(),
            1
        );
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
    fn immutable_run_accessors_preserve_cached_layout() {
        let mut doc = Document::new();
        doc.add_paragraph("Before immutable access")
            .add_run(" remains cached");

        reset_layout_invocations();
        doc.render_page_to_png_deterministic(0, 1.0).unwrap();
        assert_eq!(layout_invocations(), 1);

        let paragraph = doc.paragraph(0).expect("paragraph should exist");
        assert_eq!(paragraph.run_count(), 2);
        assert_eq!(
            paragraph.run(1).expect("run should exist").text(),
            " remains cached"
        );
        assert!(paragraph.run(2).is_none());

        doc.render_page_to_png_deterministic(0, 1.0).unwrap();
        assert_eq!(layout_invocations(), 1);
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

        let (family, font_data) = oxml_layout::bundled_fonts::bundled_font_data()[0];
        doc.to_pdf_with_fonts(&[(family, font_data)]).unwrap();
        doc.to_pdf_with_fonts(&[(family, font_data)]).unwrap();
        assert_eq!(layout_invocations(), 4);
    }

    #[test]
    fn full_layout_exposes_resolvable_font_data_and_reuses_the_cache() {
        let mut doc = Document::new();
        doc.add_paragraph("complete layout result");

        reset_layout_invocations();
        let first = doc.layout().expect("normal layout should succeed");
        let second = doc
            .layout_with_options(RenderOptions::default())
            .expect("accepted layout should succeed");
        assert!(Arc::ptr_eq(&first, &second));
        assert_eq!(layout_invocations(), 1);

        let mut sourced_runs = 0;
        for page in &first.layout.pages {
            oxml_layout::walk(&page.elements, &mut |element, _| {
                if let oxml_layout::PositionedElement::Text(run) = element {
                    assert!(first.layout.fonts.iter().any(|font| font.id == run.font_id));
                    if let Some(source) = run.source {
                        assert!(first.source_node(source.node).is_some());
                        sourced_runs += 1;
                    }
                }
            });
        }
        assert!(sourced_runs > 0);

        assert!(
            !doc.to_pdf()
                .expect("PDF should use cached layout")
                .is_empty()
        );
        assert_eq!(layout_invocations(), 1);
    }

    #[test]
    fn layout_with_fonts_returns_the_caller_font_mapping_without_caching() {
        let mut doc = Document::new();
        let (family, bytes) = caller_only_font();
        doc.add_paragraph("")
            .add_run("caller font result")
            .font(family);

        reset_layout_invocations();
        let first = doc
            .layout_with_fonts(&[(family, &bytes)])
            .expect("caller-font layout should succeed");
        let second = doc
            .layout_with_fonts_and_options(&[(family, &bytes)], RenderOptions::default())
            .expect("caller-font options layout should succeed");
        assert_eq!(layout_invocations(), 2);

        for result in [&first, &second] {
            let mut positioned_runs = Vec::new();
            for page in &result.layout.pages {
                oxml_layout::walk(&page.elements, &mut |element, _| {
                    if let oxml_layout::PositionedElement::Text(run) = element
                        && run.source.is_some()
                    {
                        positioned_runs.push(run.clone());
                    }
                });
            }
            assert!(
                !positioned_runs.is_empty(),
                "caller-font text was positioned"
            );
            for run in positioned_runs {
                let font = result
                    .layout
                    .fonts
                    .iter()
                    .find(|font| font.id == run.font_id)
                    .expect("caller-font glyph run should resolve its font id");
                assert_eq!(font.family, family);
                assert_eq!(font.data.as_ref(), bytes.as_slice());
                let source = run
                    .source
                    .expect("caller-font text should retain source provenance");
                assert!(result.source_node(source.node).is_some());
            }
        }
    }

    #[test]
    fn caller_font_layout_does_not_fall_through_to_system_or_bundled_fonts() {
        let mut doc = Document::new();
        doc.add_paragraph("caller isolation requires an explicit font universe");
        assert!(doc.layout_with_fonts(&[]).is_err());
        assert!(
            doc.normal_layout_engine
                .lock()
                .unwrap_or_else(std::sync::PoisonError::into_inner)
                .is_none()
        );
    }

    #[test]
    fn layout_options_keep_tracked_and_accepted_cache_ownership_separate() {
        let mut doc = Document::new();
        doc.add_paragraph("revision cache separation");
        let tracked = RenderOptions {
            revision_view: rdocx_layout::RevisionView::Tracked,
        };

        reset_layout_invocations();
        let accepted_first = doc.layout().expect("accepted layout should succeed");
        assert_eq!(layout_invocations(), 1);

        let tracked_first = doc
            .layout_with_options(tracked)
            .expect("tracked layout should succeed");
        let tracked_second = doc
            .layout_with_options(tracked)
            .expect("second tracked layout should succeed");
        assert!(!Arc::ptr_eq(&tracked_first, &tracked_second));
        assert_eq!(layout_invocations(), 3);
        assert!(
            doc.normal_layout_engine
                .lock()
                .unwrap_or_else(std::sync::PoisonError::into_inner)
                .is_some()
        );

        let accepted_second = doc.layout().expect("accepted cache should succeed");
        assert!(Arc::ptr_eq(&accepted_first, &accepted_second));
        assert_eq!(layout_invocations(), 3);
        assert_eq!(
            tracked_first.revision_view,
            rdocx_layout::RevisionView::Tracked
        );
        assert_eq!(
            accepted_first.revision_view,
            rdocx_layout::RevisionView::Accepted
        );
    }

    #[test]
    fn tracked_normal_layouts_retain_the_document_engine_without_caching_results() {
        let mut doc = Document::new();
        doc.add_paragraph("tracked engine reuse");
        let tracked = RenderOptions {
            revision_view: rdocx_layout::RevisionView::Tracked,
        };

        let first = doc
            .layout_with_options(tracked)
            .expect("first tracked layout");
        let second = doc
            .layout_with_options(tracked)
            .expect("second tracked layout");
        assert!(!Arc::ptr_eq(&first, &second));
        assert!(
            doc.normal_layout_engine
                .lock()
                .unwrap_or_else(std::sync::PoisonError::into_inner)
                .is_some()
        );
        assert!(
            doc.layout_cache
                .lock()
                .unwrap_or_else(std::sync::PoisonError::into_inner)
                .is_none()
        );
    }

    #[test]
    fn document_remains_send_and_sync() {
        fn assert_send_and_sync<T: Send + Sync>() {}
        assert_send_and_sync::<Document>();
    }

    #[test]
    fn relayout_caches_are_bounded_and_recover_from_poison() {
        let mut document = Document::new();
        document.add_paragraph("layout after engine lock poison");
        let document = Arc::new(document);
        let poison = Arc::clone(&document);
        assert!(
            std::thread::spawn(move || {
                let _engine = poison.normal_layout_engine.lock().unwrap();
                panic!("poison normal layout engine lock for recovery coverage");
            })
            .join()
            .is_err()
        );

        let first = document
            .layout()
            .expect("layout recovers from poisoned engine lock");
        let second = document.layout().expect("recovered layout remains cached");
        assert!(Arc::ptr_eq(&first, &second));
    }

    #[test]
    fn word_chart_part_and_workbook_round_trip() {
        let workbook_bytes = minimal_chart_workbook()
            .to_xlsx_bytes()
            .expect("serialize expected workbook");
        let mut document = document_with_minimal_chart();
        let bytes = document.to_bytes().expect("save document");
        let reopened = Document::from_bytes(&bytes).expect("reopen document");
        let chart_part = "/word/charts/chart1.xml";
        let workbook_part = "/word/embeddings/Workbook1.xlsx";
        let chart_xml = reopened.package.get_part(chart_part).expect("chart part");
        assert!(CT_ChartSpace::from_xml(chart_xml).is_ok());
        assert_eq!(
            reopened.package.get_part(workbook_part),
            Some(workbook_bytes.as_slice())
        );
        assert_eq!(
            reopened.package.content_types.content_type_for(chart_part),
            Some(content_types::CHART)
        );
        assert_eq!(
            reopened
                .package
                .content_types
                .content_type_for(workbook_part),
            Some(content_types::EMBEDDED_WORKBOOK)
        );
        let document_relationship = reopened
            .package
            .get_part_rels(&reopened.doc_part_name)
            .and_then(|relationships| relationships.get_by_type(rel_types::CHART))
            .expect("document to chart relationship");
        assert_eq!(
            OpcPackage::resolve_rel_target(&reopened.doc_part_name, &document_relationship.target,),
            chart_part
        );
        let workbook_relationship = reopened
            .package
            .get_part_rels(chart_part)
            .and_then(|relationships| relationships.get_by_type(rel_types::PACKAGE))
            .expect("chart to workbook relationship");
        assert_eq!(
            OpcPackage::resolve_rel_target(chart_part, &workbook_relationship.target),
            workbook_part
        );
        assert!(
            std::str::from_utf8(chart_xml)
                .expect("chart XML is utf8")
                .contains(&format!(
                    r#"<c:externalData r:id="{}">"#,
                    workbook_relationship.id
                ))
        );
        let document_xml = std::str::from_utf8(
            reopened
                .package
                .get_part(&reopened.doc_part_name)
                .expect("document part"),
        )
        .expect("document XML is utf8");
        assert!(document_xml.contains(&format!(r#"r:id="{}""#, document_relationship.id)));
    }

    #[test]
    fn word_chart_parts_allocate_after_sparse_suffixes() {
        let mut document = Document::new();
        document
            .package
            .set_part("/word/charts/chart3.xml", b"occupied".to_vec());
        document
            .package
            .set_part("/word/embeddings/Workbook7.xlsx", b"occupied".to_vec());
        document
            .add_chart_package(
                ChartPackageSource::Typed {
                    chart: &minimal_typed_chart(),
                    workbook: &minimal_chart_workbook(),
                },
                Length::inches(5.0),
                Length::inches(3.0),
            )
            .expect("assemble after sparse suffixes");

        assert!(
            document
                .package
                .get_part("/word/charts/chart4.xml")
                .is_some()
        );
        assert!(
            document
                .package
                .get_part("/word/embeddings/Workbook8.xlsx")
                .is_some()
        );
        assert_eq!(
            document.package.get_part("/word/charts/chart3.xml"),
            Some(b"occupied".as_slice())
        );
        assert_eq!(
            document.package.get_part("/word/embeddings/Workbook7.xlsx"),
            Some(b"occupied".as_slice())
        );
    }

    #[test]
    fn invalid_chart_package_assembly_is_atomic() {
        let mut document = Document::new();
        let before = document
            .to_bytes()
            .expect("serialize before failed mutation");
        let chart = CT_ChartSpace::from_xml(
            format!(
                r#"<c:chartSpace xmlns:c="{}" xmlns:a="{}" xmlns:r="{}" xmlns:q="{}"><c:chart><c:plotArea/></c:chart><q:externalData r:id="rId99"/></c:chartSpace>"#,
                oxml_chart::C_NS,
                oxml_chart::A_NS,
                oxml_chart::R_NS,
                oxml_chart::C_NS,
            )
            .as_bytes(),
        )
        .expect("valid chart with an occupied workbook relationship");
        let error = document
            .add_chart_package(
                ChartPackageSource::Typed {
                    chart: &chart,
                    workbook: &minimal_chart_workbook(),
                },
                Length::inches(5.0),
                Length::inches(3.0),
            )
            .expect_err("a second workbook relationship is invalid");
        assert!(error.to_string().contains("already contains"));
        assert_eq!(
            document
                .to_bytes()
                .expect("serialize after failed mutation"),
            before
        );
    }

    #[test]
    fn nested_external_data_lookalike_does_not_occupy_workbook_slot() {
        let chart = CT_ChartSpace::from_xml(
            format!(
                r#"<c:chartSpace xmlns:c="{}" xmlns:a="{}" xmlns:r="{}" xmlns:q="{}"><c:chart><c:plotArea/></c:chart><c:extLst><c:ext uri="urn:producer"><q:externalData r:id="rId99"/></c:ext></c:extLst></c:chartSpace>"#,
                oxml_chart::C_NS,
                oxml_chart::A_NS,
                oxml_chart::R_NS,
                oxml_chart::C_NS,
            )
            .as_bytes(),
        )
        .expect("valid chart with a nested producer lookalike");
        let xml = chart_with_workbook_relationship(&chart, "rId1")
            .expect("nested lookalike leaves the workbook slot available");
        let xml = std::str::from_utf8(&xml).expect("chart XML is utf8");
        assert!(xml.contains(r#"<q:externalData r:id="rId99"/>"#));
        assert!(xml.contains(r#"<c:externalData r:id="rId1">"#));
    }

    fn f158_chart_data(series_count: usize) -> ChartData {
        ChartData {
            categories: vec!["North".to_owned(), "South".to_owned(), "West".to_owned()],
            series: vec![
                ("Revenue".to_owned(), vec![12.5, 19.0, 14.25]),
                ("Cost".to_owned(), vec![8.0, 11.5, 9.75]),
            ]
            .into_iter()
            .take(series_count)
            .collect(),
            number_format: Some("0.00".to_owned()),
        }
    }

    fn assert_authored_chart_matches(chart: &CT_ChartSpace, data: &ChartData) {
        let series = chart.chart.plot_area.series().expect("typed chart series");
        assert_eq!(series.len(), data.series.len());
        for (actual, (name, values)) in series.iter().zip(&data.series) {
            assert_eq!(
                actual.name.as_ref().expect("series name").values.as_slice(),
                std::slice::from_ref(name)
            );
            let AxisData::String(categories) = actual.categories.as_ref().expect("categories")
            else {
                panic!("Word bar, line, and pie charts use string categories");
            };
            assert_eq!(categories.values, data.categories);
            assert_eq!(actual.values.values, *values);
            assert_eq!(
                actual.values.format_code,
                data.number_format.as_deref().unwrap_or("General")
            );
        }
    }

    fn deterministic_chart_layout(document: &Document) -> oxml_layout::LayoutResult {
        rdocx_layout::layout_document_deterministic(&document.build_layout_input())
            .expect("deterministic Word chart layout")
    }

    fn chart_leaf_counts(layout: &oxml_layout::LayoutResult) -> (usize, usize, usize) {
        let mut paths = 0;
        let mut text = 0;
        let mut images = 0;
        oxml_layout::walk(&layout.pages[0].elements, &mut |element, _| match element {
            oxml_layout::PositionedElement::Path(_) => paths += 1,
            oxml_layout::PositionedElement::Text(_) => text += 1,
            oxml_layout::PositionedElement::Image { .. } => images += 1,
            _ => {}
        });
        (paths, text, images)
    }

    #[test]
    fn inline_word_chart_renders_backend_neutral_group() {
        let mut document = Document::new();
        document
            .add_chart(
                ChartKind::Bar,
                Length::inches(5.0),
                Length::inches(3.0),
                &f158_chart_data(2),
            )
            .expect("author inline chart");

        let layout = deterministic_chart_layout(&document);
        assert!(layout.diagnostics.is_empty());
        assert!(
            layout.pages[0]
                .elements
                .iter()
                .any(|element| matches!(element, oxml_layout::PositionedElement::Group(_)))
        );
        let (paths, text, images) = chart_leaf_counts(&layout);
        assert!(paths > 0, "chart should lower to backend-neutral paths");
        assert!(text > 0, "chart should lower labels to shaped text");
        assert_eq!(images, 0, "chart must not be rasterized before pagination");
    }

    #[test]
    fn anchored_word_chart_uses_existing_wrap_and_z_order() {
        let mut document = Document::new();
        document
            .add_chart(
                ChartKind::Line,
                Length::inches(3.0),
                Length::inches(2.0),
                &f158_chart_data(2),
            )
            .expect("author chart");
        let BodyContent::Paragraph(paragraph) = &mut document.document.body.content[0] else {
            panic!("chart paragraph");
        };
        let RunContent::Drawing(drawing) = &mut paragraph.runs[0].content[0] else {
            panic!("chart drawing");
        };
        let inline = drawing.inline.take().expect("authored inline chart");
        let mut anchor = rdocx_oxml::drawing::CT_Anchor::new_chart(
            inline.chart_rel_id.as_deref().expect("chart relationship"),
            inline.extent_cx.0,
            inline.extent_cy.0,
        );
        anchor.behind_doc = true;
        anchor.wrap = rdocx_oxml::drawing::WrapType::Square;
        anchor.pos_h_relative_from = rdocx_oxml::drawing::ST_RelativeFromH::Page;
        anchor.pos_v_relative_from = rdocx_oxml::drawing::ST_RelativeFromV::Page;
        anchor.pos_h_offset = Length::inches(1.0).as_emu();
        anchor.pos_v_offset = Length::inches(0.5).as_emu();
        anchor.dist_l = Length::pt(12.0).as_emu();
        anchor.dist_r = Length::pt(12.0).as_emu();
        drawing.anchor = Some(anchor);
        paragraph.add_run("Foreground text");

        let layout = deterministic_chart_layout(&document);
        assert!(layout.diagnostics.is_empty());
        let group_index = layout.pages[0]
            .elements
            .iter()
            .position(|element| matches!(element, oxml_layout::PositionedElement::Group(_)))
            .expect("anchored chart group");
        let text_index = layout.pages[0]
            .elements
            .iter()
            .position(|element| matches!(element, oxml_layout::PositionedElement::Text(_)))
            .expect("chart label text");
        assert!(
            group_index <= text_index,
            "behind-text group must be emitted first"
        );
        let oxml_layout::PositionedElement::Group(group) = &layout.pages[0].elements[group_index]
        else {
            unreachable!()
        };
        assert_eq!((group.transform.e, group.transform.f), (72.0, 36.0));
        let oxml_layout::PositionedElement::Text(foreground) =
            &layout.pages[0].elements[text_index]
        else {
            unreachable!()
        };
        assert!(
            foreground.origin.x >= 300.0,
            "12 point wrap distance should clear the chart's 288 point right edge"
        );
    }

    #[test]
    fn word_chart_uses_document_theme_and_default_color_map() {
        let mut document = Document::new();
        document
            .add_chart(
                ChartKind::Bar,
                Length::inches(5.0),
                Length::inches(3.0),
                &f158_chart_data(1),
            )
            .expect("author themed chart");
        let themed = oxml_drawing::theme::OFFICE_DEFAULT_XML.replace("156082", "12AB34");
        document
            .package
            .set_part("/word/theme/custom-chart-theme.xml", themed.into_bytes());
        document.package.parts.remove("/word/theme/theme1.xml");
        document
            .package
            .get_or_create_part_rels("/word/document.xml")
            .add(rel_types::THEME, "theme/custom-chart-theme.xml");

        let input = document.build_layout_input();
        assert_eq!(
            input.chart_color_map,
            oxml_drawing::color::ColorMap::default()
        );
        let layout =
            rdocx_layout::layout_document_deterministic(&input).expect("layout themed Word chart");
        let expected = oxml_layout::Color::from_hex("12AB34");
        let mut found = false;
        oxml_layout::walk(&layout.pages[0].elements, &mut |element, _| {
            if let oxml_layout::PositionedElement::Path(path) = element
                && path.fill == Some(oxml_layout::Paint::Solid(expected))
            {
                found = true;
            }
        });
        assert!(found, "chart series should use the document theme accent");
    }

    #[test]
    fn missing_or_malformed_word_chart_is_visible() {
        for (label, mutate) in [
            ("missing", 0_u8),
            ("malformed", 1_u8),
            ("external", 2_u8),
            ("unsupported", 3_u8),
        ] {
            let mut document = Document::new();
            document
                .add_chart(
                    ChartKind::Bar,
                    Length::inches(5.0),
                    Length::inches(3.0),
                    &f158_chart_data(1),
                )
                .expect("author chart before corrupting target");
            match mutate {
                0 => {
                    document.package.parts.remove("/word/charts/chart1.xml");
                }
                1 => document
                    .package
                    .set_part("/word/charts/chart1.xml", b"<c:chartSpace".to_vec()),
                2 => {
                    let relationship = document
                        .package
                        .get_or_create_part_rels("/word/document.xml")
                        .items
                        .iter_mut()
                        .find(|relationship| relationship.rel_type == rel_types::CHART)
                        .expect("chart relationship");
                    relationship.target = "https://example.invalid/chart.xml".to_owned();
                    relationship.target_mode = Some("External".to_owned());
                }
                3 => document.package.set_part(
                    "/word/charts/chart1.xml",
                    format!(
                        r#"<c:chartSpace xmlns:c="{}"><c:chart><c:plotArea/></c:chart></c:chartSpace>"#,
                        oxml_chart::C_NS,
                    )
                    .into_bytes(),
                ),
                _ => unreachable!(),
            }

            let layout = deterministic_chart_layout(&document);
            assert_eq!(layout.diagnostics.len(), 1, "{label}");
            assert!(
                layout.diagnostics[0]
                    .message
                    .contains("Word chart relationship")
            );
            let (_, text, images) = chart_leaf_counts(&layout);
            assert!(text > 0, "{label} chart fallback should be visible");
            assert_eq!(images, 0, "{label} chart must not become an empty image");
        }
    }

    #[test]
    fn word_and_powerpoint_chart_pixels_are_identical() {
        const DPI: &str = "150";
        const RASTERIZER: &str = "pdftoppm version 26.01.0";
        const CROP_X: &str = "150";
        const CROP_Y: &str = "150";
        const CROP_WIDTH: &str = "750";
        const CROP_HEIGHT: &str = "450";
        const WORD_SHA256: &str =
            "e50845637449e2af4b8e2dbf16f5f6f53e5f598a00401fcc34c13f5d5716a1c4";
        const POWERPOINT_SHA256: &str =
            "7525e9a088c5fbf58fa1ed98cdfa0ec2fabf998662112ced7a6b6521f2c4edfc";

        let data = f158_chart_data(2);
        let evidence_dir = std::env::temp_dir();
        let word_path =
            evidence_dir.join(format!("rdocx-f159-word-chart-{}.docx", std::process::id()));
        let powerpoint_path = evidence_dir.join(format!(
            "rdocx-f159-powerpoint-chart-{}.pptx",
            std::process::id()
        ));

        let mut word = Document::new();
        word.add_chart(
            ChartKind::Bar,
            Length::inches(5.0),
            Length::inches(3.0),
            &data,
        )
        .expect("author Word chart from shared data");
        let mut powerpoint = rpptx::Presentation::new().expect("open bundled PowerPoint template");
        powerpoint
            .set_slide_size(Length::inches(8.5).as_emu(), Length::inches(11.0).as_emu())
            .expect("match the Word page size");
        powerpoint.add_slide(0).expect("add chart slide");
        powerpoint
            .add_chart(
                0,
                ChartKind::Bar,
                Length::inches(1.0).as_emu(),
                Length::inches(1.0).as_emu(),
                Length::inches(5.0).as_emu(),
                Length::inches(3.0).as_emu(),
                &data,
            )
            .expect("author PowerPoint chart from shared data");
        powerpoint
            .save(&powerpoint_path)
            .expect("save PowerPoint chart artifact");
        let powerpoint_package = OpcPackage::from_reader(Cursor::new(
            powerpoint
                .to_bytes()
                .expect("serialize PowerPoint artifact"),
        ))
        .expect("open PowerPoint artifact package");
        let effective_theme = powerpoint_package
            .get_part("/ppt/theme/theme1.xml")
            .expect("PowerPoint effective theme")
            .to_vec();
        word.package
            .set_part("/word/theme/theme1.xml", effective_theme);
        word.package
            .get_or_create_part_rels("/word/document.xml")
            .add(rel_types::THEME, "theme/theme1.xml");
        word.save(&word_path).expect("save Word chart artifact");

        let word_sha = sha256(&word_path);
        let powerpoint_sha = sha256(&powerpoint_path);
        assert_eq!(word_sha, WORD_SHA256);
        assert_eq!(powerpoint_sha, POWERPOINT_SHA256);

        let version = Command::new("pdftoppm")
            .arg("-v")
            .output()
            .expect("run pinned rasterizer");
        let reported = String::from_utf8_lossy(&version.stderr);
        assert!(
            reported
                .lines()
                .next()
                .is_some_and(|line| line == RASTERIZER)
        );

        let word_pdf =
            std::env::temp_dir().join(format!("rdocx-f159-word-chart-{}.pdf", std::process::id()));
        let powerpoint_pdf = std::env::temp_dir().join(format!(
            "rdocx-f159-powerpoint-chart-{}.pdf",
            std::process::id()
        ));
        fs::write(
            &word_pdf,
            word.to_pdf_deterministic().expect("render Word chart PDF"),
        )
        .expect("write Word chart PDF");
        fs::write(
            &powerpoint_pdf,
            powerpoint
                .to_pdf_deterministic()
                .expect("render PowerPoint chart PDF"),
        )
        .expect("write PowerPoint chart PDF");

        let word_crop =
            std::env::temp_dir().join(format!("rdocx-f159-word-chart-crop-{}", std::process::id()));
        let powerpoint_crop = std::env::temp_dir().join(format!(
            "rdocx-f159-powerpoint-chart-crop-{}",
            std::process::id()
        ));
        for (pdf, crop) in [(&word_pdf, &word_crop), (&powerpoint_pdf, &powerpoint_crop)] {
            let output = Command::new("pdftoppm")
                .args([
                    "-f",
                    "1",
                    "-l",
                    "1",
                    "-singlefile",
                    "-r",
                    DPI,
                    "-x",
                    CROP_X,
                    "-y",
                    CROP_Y,
                    "-W",
                    CROP_WIDTH,
                    "-H",
                    CROP_HEIGHT,
                    "-png",
                ])
                .arg(pdf)
                .arg(crop)
                .output()
                .expect("rasterize chart crop");
            assert!(
                output.status.success(),
                "pdftoppm chart crop failed: {}",
                String::from_utf8_lossy(&output.stderr)
            );
        }

        let word_png = word_crop.with_extension("png");
        let powerpoint_png = powerpoint_crop.with_extension("png");
        let comparison = Command::new("python3")
            .args([
                "-c",
                "import sys\nsys.path.insert(0, sys.argv[3])\nfrom scripts.golden_png_harness import decode_png\na=decode_png(__import__('pathlib').Path(sys.argv[1]))\nb=decode_png(__import__('pathlib').Path(sys.argv[2]))\nassert a[:2] == b[:2] == (750, 450), (a[:2], b[:2])\ndiff=sum(a[2][i:i+4] != b[2][i:i+4] for i in range(0, len(a[2]), 4))\nprint(f'{a[0]}x{a[1]} differing={diff}')\nassert diff == 0, diff",
            ])
            .arg(&word_png)
            .arg(&powerpoint_png)
            .arg(Path::new(env!("CARGO_MANIFEST_DIR")).join("../.."))
            .output()
            .expect("decode and compare chart RGBA pixels");
        assert!(
            comparison.status.success(),
            "chart pixel comparison failed: {}{}",
            String::from_utf8_lossy(&comparison.stdout),
            String::from_utf8_lossy(&comparison.stderr)
        );
        assert_eq!(
            String::from_utf8(comparison.stdout)
                .expect("pixel comparison output is utf8")
                .trim(),
            "750x450 differing=0"
        );

        for path in [
            word_path,
            powerpoint_path,
            word_pdf,
            powerpoint_pdf,
            word_png,
            powerpoint_png,
        ] {
            fs::remove_file(path).expect("remove temporary golden evidence");
        }
    }

    #[test]
    fn added_bar_line_and_pie_charts_keep_source_data() {
        let mut document = Document::new();
        let expected = [
            (ChartKind::Bar, f158_chart_data(2), "<c:barChart>"),
            (ChartKind::Line, f158_chart_data(2), "<c:lineChart>"),
            (ChartKind::Pie, f158_chart_data(1), "<c:pieChart>"),
        ];
        for (kind, data, _) in &expected {
            document
                .add_chart(*kind, Length::inches(5.0), Length::inches(3.0), data)
                .expect("public Word chart authoring");
        }

        let bytes = document.to_bytes().expect("save authored Word charts");
        let reopened = Document::from_bytes(&bytes).expect("reopen authored Word charts");
        let mut chart_parts = reopened
            .package
            .content_types
            .overrides
            .iter()
            .filter_map(|(part, content_type)| {
                (content_type == content_types::CHART).then_some(part.as_str())
            })
            .collect::<Vec<_>>();
        chart_parts.sort_unstable();
        assert_eq!(chart_parts.len(), expected.len());
        for (part, (_, data, plot_tag)) in chart_parts.into_iter().zip(expected) {
            let xml = reopened
                .package
                .get_part(part)
                .expect("authored chart part");
            assert!(
                std::str::from_utf8(xml)
                    .expect("chart XML is utf8")
                    .contains(plot_tag)
            );
            let chart = CT_ChartSpace::from_xml(xml).expect("typed authored chart");
            assert_authored_chart_matches(&chart, &data);
        }
    }

    #[test]
    fn word_add_chart_writes_cache_and_workbook_from_one_source() {
        let data = f158_chart_data(2);
        let mut document = Document::new();
        document
            .add_chart(
                ChartKind::Bar,
                Length::inches(5.0),
                Length::inches(3.0),
                &data,
            )
            .expect("author Word chart");
        let bytes = document.to_bytes().expect("save Word chart");
        let package = Document::from_bytes(&bytes)
            .expect("reopen Word chart")
            .package;
        let chart_part = "/word/charts/chart1.xml";
        let chart = CT_ChartSpace::from_xml(package.get_part(chart_part).expect("chart part"))
            .expect("parse chart part");
        assert_authored_chart_matches(&chart, &data);
        let series = chart.chart.plot_area.series().expect("typed chart series");
        for (actual, column) in series.iter().zip(['B', 'C']) {
            assert_eq!(
                actual.name.as_ref().expect("series name").formula,
                format!("Sheet1!${column}$1")
            );
            let AxisData::String(categories) = actual.categories.as_ref().expect("categories")
            else {
                panic!("authored bar chart uses string categories");
            };
            assert_eq!(categories.formula, "Sheet1!$A$2:$A$4");
            assert_eq!(
                actual.values.formula,
                format!("Sheet1!${column}$2:${column}$4")
            );
        }
        let workbook_relationship = package
            .get_part_rels(chart_part)
            .and_then(|relationships| relationships.get_by_type(rel_types::PACKAGE))
            .expect("chart workbook relationship");
        let workbook_part =
            OpcPackage::resolve_rel_target(chart_part, &workbook_relationship.target);
        let workbook = OpcPackage::from_reader(Cursor::new(
            package.get_part(&workbook_part).expect("workbook part"),
        ))
        .expect("open embedded workbook");
        let worksheet = std::str::from_utf8(
            workbook
                .get_part("/xl/worksheets/sheet1.xml")
                .expect("worksheet"),
        )
        .expect("worksheet XML is utf8");
        let shared_strings = std::str::from_utf8(
            workbook
                .get_part("/xl/sharedStrings.xml")
                .expect("shared strings"),
        )
        .expect("shared strings XML is utf8");
        let shared_values = shared_strings
            .split("<si>")
            .skip(1)
            .map(|item| {
                item.split_once("<t>")
                    .expect("shared string text start")
                    .1
                    .split_once("</t>")
                    .expect("shared string text end")
                    .0
            })
            .collect::<Vec<_>>();
        for (address, value) in [
            ("A1", "Category"),
            ("A2", "North"),
            ("A3", "South"),
            ("A4", "West"),
            ("B1", "Revenue"),
            ("C1", "Cost"),
        ] {
            let index = worksheet_cell_value(worksheet, address)
                .parse::<usize>()
                .expect("shared string index");
            assert_eq!(shared_values[index], value);
        }
        for (address, value) in [
            ("B2", "12.5"),
            ("B3", "19"),
            ("B4", "14.25"),
            ("C2", "8"),
            ("C3", "11.5"),
            ("C4", "9.75"),
        ] {
            assert_eq!(worksheet_cell_value(worksheet, address), value);
        }
    }

    fn worksheet_cell_value<'a>(worksheet: &'a str, address: &str) -> &'a str {
        let marker = format!(r#"<c r="{address}""#);
        worksheet
            .split_once(&marker)
            .unwrap_or_else(|| panic!("missing worksheet cell {address}"))
            .1
            .split_once("</c>")
            .expect("worksheet cell end")
            .0
            .split_once("<v>")
            .expect("worksheet value start")
            .1
            .split_once("</v>")
            .expect("worksheet value end")
            .0
    }

    #[test]
    fn word_add_chart_rejects_invalid_data_without_mutation() {
        let invalid = [
            ChartData {
                categories: Vec::new(),
                series: vec![("Revenue".to_owned(), Vec::new())],
                number_format: None,
            },
            ChartData {
                categories: vec!["North".to_owned()],
                series: Vec::new(),
                number_format: None,
            },
            ChartData {
                categories: vec!["North".to_owned(), "South".to_owned()],
                series: vec![("Revenue".to_owned(), vec![12.5])],
                number_format: None,
            },
            ChartData {
                categories: vec!["North".to_owned()],
                series: vec![("Revenue".to_owned(), vec![f64::NAN])],
                number_format: None,
            },
            ChartData {
                categories: vec!["North".to_owned()],
                series: vec![("Revenue".to_owned(), vec![12.5])],
                number_format: Some(String::new()),
            },
        ];
        for data in invalid {
            let mut document = Document::new();
            let before = document.to_bytes().expect("state before invalid chart");
            assert!(
                document
                    .add_chart(
                        ChartKind::Bar,
                        Length::inches(5.0),
                        Length::inches(3.0),
                        &data,
                    )
                    .is_err()
            );
            assert_eq!(
                document.to_bytes().expect("state after invalid chart"),
                before
            );
        }

        let invalid_cases = [
            (ChartKind::Pie, Length::inches(5.0), Length::inches(3.0)),
            (ChartKind::Bar, Length::emu(0), Length::inches(3.0)),
            (ChartKind::Bar, Length::emu(-1), Length::inches(3.0)),
            (ChartKind::Bar, Length::inches(5.0), Length::emu(0)),
            (ChartKind::Bar, Length::inches(5.0), Length::emu(-1)),
        ];
        for (kind, width, height) in invalid_cases {
            let mut document = Document::new();
            let before = document.to_bytes().expect("state before invalid chart");
            assert!(
                document
                    .add_chart(kind, width, height, &f158_chart_data(2))
                    .is_err()
            );
            assert_eq!(
                document.to_bytes().expect("state after invalid chart"),
                before
            );
        }
    }

    #[test]
    fn word_add_chart_uses_inline_flow_placement() {
        let mut document = Document::new();
        let width = Length::inches(5.0);
        let height = Length::inches(3.0);
        let paragraph = document
            .add_chart(ChartKind::Line, width, height, &f158_chart_data(2))
            .expect("author inline Word chart");
        assert_eq!(paragraph.run_count(), 1);
        assert_eq!(document.content_count(), 1);
        let BodyContent::Paragraph(paragraph) = &document.document.body.content[0] else {
            panic!("Word chart must be inline flow content");
        };
        let RunContent::Drawing(drawing) = &paragraph.runs[0].content[0] else {
            panic!("Word chart paragraph must carry a drawing run");
        };
        let inline = drawing.inline.as_ref().expect("inline chart drawing");
        assert!(drawing.anchor.is_none());
        assert_eq!(inline.extent_cx.0, width.to_emu());
        assert_eq!(inline.extent_cy.0, height.to_emu());
        assert!(inline.chart_rel_id.is_some());
    }

    #[test]
    #[ignore = "requires pinned Microsoft Word and human Edit Data evidence"]
    fn word_opens_native_chart_without_repair() {
        let output = std::env::var_os("RDOCX_WORD_CHART_GATE_OUTPUT")
            .map(std::path::PathBuf::from)
            .expect("set RDOCX_WORD_CHART_GATE_OUTPUT to the SHA-bound .docx path");
        let mut document = document_with_minimal_chart();
        document.save(&output).expect("write Word chart candidate");
        assert_eq!(sha256(&output), WORD_CHART_CANDIDATE_SHA256);
        let plist = "/Applications/Microsoft Word.app/Contents/Info.plist";
        assert_eq!(
            plist_value(plist, "CFBundleShortVersionString"),
            WORD_VERSION
        );
        assert_eq!(plist_value(plist, "CFBundleVersion"), WORD_BUILD);
        let script = format!(
            "with timeout of 120 seconds\ntell application \"Microsoft Word\"\nactivate\nopen POSIX file \"{}\"\ndelay 3\nset gateDocument to active document\ntry\nset openedPath to POSIX path of (full name of gateDocument as alias)\nif openedPath is not \"{}\" then error \"Word chart candidate path mismatch\"\nclose gateDocument saving no\non error errorMessage number errorNumber\ntry\nclose gateDocument saving no\nend try\nerror errorMessage number errorNumber\nend try\nend tell\nend timeout\n",
            output.display(),
            output.display(),
        );
        let result = Command::new("osascript")
            .args(["-e", &script])
            .output()
            .expect("launch Word acceptance script");
        assert!(
            result.status.success(),
            "Microsoft Word F-157 acceptance failed: {}",
            String::from_utf8_lossy(&result.stderr)
        );
    }

    #[test]
    fn word_chart_candidate_is_bound_to_recorded_sha() {
        let output =
            std::env::temp_dir().join(format!("rdocx-f157-word-chart-{}.docx", std::process::id()));
        document_with_minimal_chart()
            .save(&output)
            .expect("write SHA-bound candidate");
        assert_eq!(sha256(&output), WORD_CHART_CANDIDATE_SHA256);
        fs::remove_file(output).expect("remove temporary candidate");
    }

    fn sha256(path: &Path) -> String {
        let output = Command::new("shasum")
            .args(["-a", "256"])
            .arg(path)
            .output()
            .unwrap_or_else(|error| panic!("{}: run shasum: {error}", path.display()));
        assert!(
            output.status.success(),
            "shasum failed: {}",
            String::from_utf8_lossy(&output.stderr)
        );
        String::from_utf8(output.stdout)
            .expect("shasum output is utf8")
            .split_whitespace()
            .next()
            .expect("shasum digest")
            .to_owned()
    }

    fn plist_value(path: &str, key: &str) -> String {
        let output = Command::new("defaults")
            .args(["read", path.trim_end_matches(".plist"), key])
            .output()
            .expect("read application plist");
        assert!(output.status.success());
        String::from_utf8(output.stdout)
            .expect("plist value is utf8")
            .trim()
            .to_owned()
    }

    #[test]
    fn html_and_layout_media_use_sniffed_content_type() {
        let jpeg = [0xff, 0xd8, 0xff, 0xd9];
        let mut document = Document::new();
        document
            .package
            .set_part("/word/media/misleading.png", jpeg.to_vec());
        let relationship_id = document
            .package
            .get_or_create_part_rels("/word/document.xml")
            .add(rel_types::IMAGE, "media/misleading.png");

        let html_input = document.build_html_input();
        let layout_input = document.build_layout_input();

        assert_eq!(
            html_input.images[&relationship_id].content_type,
            "image/jpeg"
        );
        assert_eq!(
            layout_input.images[&relationship_id].content_type,
            "image/jpeg"
        );
    }

    #[test]
    fn deterministic_render_is_independent_of_system_fonts() {
        let mut doc = Document::new();
        doc.add_paragraph("Deterministic rendering");

        let input = doc.build_layout_input();
        let layout = rdocx_layout::layout_document_deterministic(&input)
            .expect("deterministic layout should succeed");
        let bundled_fonts = oxml_layout::bundled_fonts::bundled_font_data();

        assert!(!layout.fonts.is_empty());
        for font in &layout.fonts {
            assert!(!font.data.is_empty());
            assert!(
                bundled_fonts
                    .iter()
                    .any(|(_family, data)| *data == font.data.as_ref()),
                "resolved font '{}' did not come from the bundled font set",
                font.family
            );
        }

        let inspected = oxml_pdf::render_page_to_png(&layout, 0, 150.0)
            .expect("document should have a first page");
        let facade = doc
            .render_page_to_png_deterministic(0, 150.0)
            .expect("deterministic layout should succeed")
            .expect("document should have a first page");

        assert!(!inspected.is_empty());
        assert_eq!(facade, inspected);
    }

    #[test]
    fn deterministic_pdf_facade_reuses_bundled_font_layout() {
        let mut doc = Document::new();
        doc.add_paragraph("Deterministic PDF rendering");

        reset_layout_invocations();
        let first = doc
            .to_pdf_deterministic()
            .expect("deterministic PDF rendering should succeed");
        let second = doc
            .to_pdf_deterministic()
            .expect("cached deterministic PDF rendering should succeed");

        assert!(first.starts_with(b"%PDF-"));
        assert!(second.starts_with(b"%PDF-"));
        assert_eq!(layout_invocations(), 1);
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
    fn document_text_preserves_body_and_table_order() {
        let mut doc = Document::new();
        doc.add_paragraph("Before");
        let mut table = doc.add_table(1, 2);
        table.cell(0, 0).unwrap().set_text("Left");
        table.cell(0, 1).unwrap().set_text("Right");
        doc.add_paragraph("After");

        assert_eq!(doc.text(), "Before\nLeft\tRight\t\nAfter\n");
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
    fn a_tag_split_across_five_formatted_runs_preserves_surrounding_formatting() {
        let mut doc = Document::new();
        let mut paragraph = CT_P::new();
        for (text, properties) in [
            (
                "Before {",
                CT_RPr {
                    bold: Some(true),
                    ..Default::default()
                },
            ),
            (
                "{ pro",
                CT_RPr {
                    italic: Some(true),
                    ..Default::default()
                },
            ),
            (
                "file.",
                CT_RPr {
                    color: Some("112233".to_owned()),
                    ..Default::default()
                },
            ),
            (
                "name ",
                CT_RPr {
                    strike: Some(true),
                    ..Default::default()
                },
            ),
            (
                "}} after",
                CT_RPr {
                    italic: Some(false),
                    ..Default::default()
                },
            ),
        ] {
            let mut run = CT_R::new(text);
            run.properties = Some(properties);
            paragraph.runs.push(run);
        }
        doc.document
            .body
            .content
            .push(BodyContent::Paragraph(paragraph));

        let count = doc
            .render_template(&serde_json::json!({"profile": {"name": "Ada"}}))
            .expect("valid template should render");

        assert_eq!(count, 1);
        let paragraph = doc.paragraph(0).expect("rendered paragraph");
        assert_eq!(paragraph.text(), "Before Ada after");
        assert_eq!(paragraph.run_count(), 2);
        assert_eq!(paragraph.run(0).unwrap().text(), "Before Ada");
        assert!(paragraph.run(0).unwrap().is_bold());
        assert_eq!(paragraph.run(1).unwrap().text(), " after");
        assert_eq!(paragraph.run(1).unwrap().italic_value(), Some(false));
    }

    #[test]
    fn dotted_scalar_paths_render_supported_json_leaves() {
        let mut doc = Document::new();
        doc.add_paragraph("{{ person.name }}");
        doc.add_paragraph("{{person.age}}");
        doc.add_paragraph("{{ person.active }}");
        doc.add_paragraph("x{{person.middle}}y");

        let count = doc
            .render_template(&serde_json::json!({
                "person": {
                    "name": "Ada",
                    "age": 37,
                    "active": true,
                    "middle": null
                }
            }))
            .expect("scalar leaves should render");

        assert_eq!(count, 4);
        assert_eq!(
            doc.paragraphs()
                .into_iter()
                .map(|paragraph| paragraph.text())
                .collect::<Vec<_>>(),
            ["Ada", "37", "true", "xy"]
        );
    }

    #[test]
    fn invalid_template_input_leaves_the_document_unchanged() {
        for (template, data) in [
            (
                "{{present}} {{missing}}",
                serde_json::json!({"present": "value"}),
            ),
            ("{{value}}", serde_json::json!({"value": [1, 2]})),
            ("{{value}}", serde_json::json!({"value": {"nested": 1}})),
            ("{{ malformed", serde_json::json!({"malformed": "value"})),
        ] {
            let mut doc = Document::new();
            doc.add_paragraph(template);
            let before_xml = doc.document.to_xml().expect("serialize before render");
            let before_parts = doc.package.parts.clone();

            reset_layout_invocations();
            doc.render_page_to_png_deterministic(0, 1.0)
                .expect("warm deterministic layout cache");
            assert_eq!(layout_invocations(), 1);

            assert!(doc.render_template(&data).is_err());
            assert_eq!(
                doc.document.to_xml().expect("serialize after rejection"),
                before_xml
            );
            assert_eq!(doc.package.parts, before_parts);
            doc.render_page_to_png_deterministic(0, 1.0)
                .expect("reuse deterministic layout after rejection");
            assert_eq!(layout_invocations(), 1);
        }
    }

    #[test]
    fn template_scalar_coverage_matches_literal_replacement() {
        let mut doc = Document::new();
        doc.add_paragraph("Body {{body}}");
        doc.set_header("Header {{header}}");
        doc.set_footer("Footer {{footer}}");
        doc.add_table(1, 1)
            .cell(0, 0)
            .expect("template table cell")
            .set_text("Table {{table}}");
        doc.document.body.content.push(BodyContent::RawXml(
            br#"<w:custom><w:txbxContent><w:p><w:r><w:t>Box {{box}}</w:t></w:r></w:p></w:txbxContent></w:custom>"#.to_vec(),
        ));

        let chart_part = "/word/charts/chart99.xml";
        doc.package.set_part(
            chart_part,
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:title><a:p><a:r><a:t>{{chart}}</a:t></a:r></a:p></c:title><c:strCache><c:pt><c:v>{{cache}}</c:v></c:pt></c:strCache></c:chartSpace>"#.to_vec(),
        );
        doc.package
            .get_or_create_part_rels(&doc.doc_part_name.clone())
            .add(rel_types::CHART, "charts/chart99.xml");

        let count = doc
            .render_template(&serde_json::json!({
                "body": "B",
                "header": "H",
                "footer": "F",
                "table": "T",
                "box": "X",
                "chart": "C",
                "cache": "K"
            }))
            .expect("all literal-replacement locations should render");

        assert_eq!(count, 7);
        assert_eq!(doc.paragraphs()[0].text(), "Body B");
        assert_eq!(
            doc.table(0)
                .expect("template table")
                .cell(0, 0)
                .expect("rendered table cell")
                .text(),
            "Table T"
        );
        assert_eq!(doc.header_text().as_deref(), Some("Header H"));
        assert_eq!(doc.footer_text().as_deref(), Some("Footer F"));
        let document_xml = std::str::from_utf8(
            doc.package
                .get_part(&doc.doc_part_name)
                .expect("main document package part"),
        )
        .expect("main document XML");
        assert!(document_xml.contains("Box X"));
        let chart_xml = std::str::from_utf8(
            doc.package
                .get_part(chart_part)
                .expect("chart package part"),
        )
        .expect("chart XML");
        assert!(chart_xml.contains("<a:t>C</a:t>"));
        assert!(chart_xml.contains("<c:v>K</c:v>"));
    }

    #[test]
    fn template_values_are_not_recursively_interpreted() {
        let mut doc = Document::new();
        doc.add_paragraph("{{first}} {{second}}");
        doc.add_paragraph("{% if later %}");
        doc.add_paragraph("hidden");
        doc.add_paragraph("{% endif %}");

        assert_eq!(
            doc.render_template(&serde_json::json!({
                "first": "{{second}}",
                "second": "done",
                "later": false
            }))
            .expect("scalar values and controls should render"),
            2
        );
        assert_eq!(doc.paragraphs().len(), 1);
        assert_eq!(doc.paragraphs()[0].text(), "{{second}} done");
    }

    #[test]
    fn template_render_preserves_unmodelled_paragraph_xml() {
        let raw = br#"<w:proofErr w:type="spellStart"/>"#.to_vec();
        let mut doc = Document::new();
        let mut paragraph = CT_P::new();
        paragraph.runs.push(CT_R::new("Value: {{ value }}"));
        paragraph.extra_xml.push((1, raw.clone()));
        doc.document
            .body
            .content
            .push(BodyContent::Paragraph(paragraph));

        assert_eq!(
            doc.render_template(&serde_json::json!({"value": "kept"}))
                .expect("valid template should render"),
            1
        );
        let reopened = Document::from_bytes(&doc.to_bytes().expect("save rendered document"))
            .expect("reopen rendered document");
        let BodyContent::Paragraph(paragraph) = &reopened.document.body.content[0] else {
            panic!("expected paragraph");
        };
        assert_eq!(paragraph.text(), "Value: kept");
        assert_eq!(paragraph.extra_xml, vec![(1, raw)]);
    }

    #[test]
    fn mismatched_or_cross_container_blocks_fail_without_mutation() {
        let cases = [
            vec!["{% if show %}", "content"],
            vec!["{% endif %}"],
            vec![
                "{% for item in items %}",
                "{% if item.show %}",
                "{% endfor %}",
                "{% endif %}",
            ],
        ];
        for paragraphs in cases {
            let mut document = Document::new();
            for paragraph in paragraphs {
                document.add_paragraph(paragraph);
            }
            let before = document.document.to_xml().unwrap();
            assert!(
                document
                    .render_template(&serde_json::json!({
                        "show": true,
                        "items": [{"show": true}]
                    }))
                    .is_err()
            );
            assert_eq!(document.document.to_xml().unwrap(), before);
        }

        let mut document = Document::new();
        document.add_paragraph("{% for item in items %}");
        document
            .add_table(1, 1)
            .cell(0, 0)
            .unwrap()
            .set_text("{% endfor %}");
        let before = document.document.to_xml().unwrap();
        assert!(
            document
                .render_template(&serde_json::json!({"items": [1]}))
                .is_err()
        );
        assert_eq!(document.document.to_xml().unwrap(), before);
    }

    #[test]
    fn loop_scopes_shadow_root_values_and_restore_after_exit() {
        let mut document = Document::new();
        for text in [
            "{% for item in items %}",
            "outer {{ item.name }}",
            "{% for item in item.children %}",
            "inner {{ item.name }}",
            "{% endfor %}",
            "restored {{ item.name }}",
            "{% endfor %}",
            "root {{ item.name }}",
        ] {
            document.add_paragraph(text);
        }

        document
            .render_template(&serde_json::json!({
                "item": {"name": "root"},
                "items": [{
                    "name": "outer",
                    "children": [{"name": "inner-a"}, {"name": "inner-b"}]
                }]
            }))
            .unwrap();
        assert_eq!(
            document
                .paragraphs()
                .into_iter()
                .map(|paragraph| paragraph.text())
                .collect::<Vec<_>>(),
            [
                "outer outer",
                "inner inner-a",
                "inner inner-b",
                "restored outer",
                "root root"
            ]
        );
    }

    #[test]
    fn structural_generation_preserves_schema_order_and_raw_xml() {
        let mut document = Document::new();
        document.add_paragraph("{% for item in items %}");
        let mut generated = CT_P::new();
        generated.runs.push(CT_R::new("section {{ item }}"));
        generated
            .extra_xml
            .push((1, br#"<w:proofErr w:type="spellStart"/>"#.to_vec()));
        generated.properties = Some(CT_PPr {
            sect_pr: Some(CT_SectPr::default_letter()),
            ..Default::default()
        });
        document
            .document
            .body
            .content
            .push(BodyContent::Paragraph(generated));
        document.add_paragraph("{% endfor %}");

        let row = |text: &str, raw: Option<Vec<u8>>| {
            let mut paragraph = CT_P::new();
            paragraph.runs.push(CT_R::new(text));
            let mut cell = CT_Tc::new();
            cell.content = vec![CellContent::Paragraph(paragraph)];
            let mut row = CT_Row::new();
            row.cells.push(cell);
            if let Some(raw) = raw {
                row.extra_xml.push((0, raw));
            }
            row
        };
        let mut table = CT_Tbl::new();
        table.rows.push(row("{% for item in items %}", None));
        table.rows.push(row(
            "row {{ item }}",
            Some(br#"<w:customRow w:val="kept"/>"#.to_vec()),
        ));
        table.rows.push(row("{% endfor %}", None));
        document
            .document
            .body
            .content
            .push(BodyContent::Table(table));

        document
            .render_template(&serde_json::json!({"items": ["one", "two"]}))
            .unwrap();
        let bytes = document.to_bytes().unwrap();
        let reopened = Document::from_bytes(&bytes).unwrap();
        assert_eq!(reopened.document.body.content.len(), 3);
        for content in &reopened.document.body.content[..2] {
            let BodyContent::Paragraph(paragraph) = content else {
                panic!("expected generated section-ending paragraph");
            };
            assert_eq!(
                paragraph.extra_xml,
                vec![(1, br#"<w:proofErr w:type="spellStart"/>"#.to_vec())]
            );
            assert!(
                paragraph
                    .properties
                    .as_ref()
                    .and_then(|properties| properties.sect_pr.as_ref())
                    .is_some()
            );
        }
        let BodyContent::Table(table) = &reopened.document.body.content[2] else {
            panic!("expected generated table");
        };
        assert_eq!(table.rows.len(), 2);
        assert_eq!(
            table
                .rows
                .iter()
                .map(|row| row.cells[0].text())
                .collect::<Vec<_>>(),
            ["row one", "row two"]
        );
        assert!(
            table.rows.iter().all(|row| {
                row.extra_xml == vec![(0, br#"<w:customRow w:val="kept"/>"#.to_vec())]
            })
        );
        let package = OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
        let xml = std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
        assert!(xml.rfind("<w:p>").unwrap() < xml.rfind("<w:sectPr>").unwrap());
    }

    #[test]
    fn repeated_rows_and_lists_preserve_properties_and_raw_xml() {
        use rdocx_oxml::table::{
            CT_TblGrid, CT_TblGridCol, CT_TblLook, CT_TblPr, CT_TcPr, CT_TrPr, VMerge,
        };

        let mut document = Document::new();
        document.add_paragraph("{% for item in items %}");
        document.add_numbered_list_item("List {{ item }}", 1);
        document.add_paragraph("Detail {{ item }}");
        document.add_paragraph("{% endfor %}");

        let row_raw = br#"<q:rowData xmlns:q="urn:rdocx:f165" q:value="kept"/>"#.to_vec();
        let cell_property_raw =
            br#"<q:cellProperty xmlns:q="urn:rdocx:f165" q:value="kept"/>"#.to_vec();
        let cell_raw = br#"<q:cellData xmlns:q="urn:rdocx:f165" q:value="kept"/>"#.to_vec();
        let table_raw = br#"<q:tableData xmlns:q="urn:rdocx:f165" q:value="kept"/>"#.to_vec();
        let control_xml = br#"<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:sdt><w:sdtPr><w:tag w:val="row-control"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>controlled</w:t></w:r></w:p></w:sdtContent></w:sdt><w:sectPr/></w:body></w:document>"#;
        let mut control_document = CT_Document::from_xml(control_xml).unwrap();
        let BodyContent::ContentControl(control) = control_document.body.content.remove(0) else {
            panic!("expected parsed content control");
        };

        let marker_row = |text: &str| {
            let mut paragraph = CT_P::new();
            paragraph.runs.push(CT_R::new(text));
            let mut cell = CT_Tc::new();
            cell.content = vec![CellContent::Paragraph(paragraph)];
            let mut row = CT_Row::new();
            row.cells.push(cell);
            row
        };
        let template_row = |text: &str, merge: VMerge, header: bool| {
            let mut paragraph = CT_P::new();
            paragraph.runs.push(CT_R::new(text));
            let mut properties = CT_TcPr {
                grid_span: Some(2),
                v_merge: Some(merge),
                ..Default::default()
            };
            properties.extra_xml.push((3, cell_property_raw.clone()));
            let mut cell = CT_Tc::new();
            cell.properties = Some(properties);
            cell.content = vec![
                CellContent::Paragraph(paragraph),
                CellContent::ContentControl(control.clone()),
            ];
            cell.extra_xml.push((0, cell_raw.clone()));
            let mut row = CT_Row::new();
            row.properties = Some(CT_TrPr {
                header: Some(header),
                cnf_style: Some("100000000000".to_owned()),
                ..Default::default()
            });
            row.cells.push(cell);
            row.extra_xml.push((0, row_raw.clone()));
            row
        };
        let mut table = CT_Tbl::new();
        table.properties = Some(CT_TblPr {
            style_id: Some("BandedRows".to_owned()),
            look: Some(CT_TblLook {
                first_row: Some(true),
                no_h_band: Some(false),
                ..Default::default()
            }),
            ..Default::default()
        });
        table.grid = Some(CT_TblGrid {
            columns: vec![
                CT_TblGridCol { width: Twips(1800) },
                CT_TblGridCol { width: Twips(1800) },
                CT_TblGridCol { width: Twips(1800) },
            ],
        });
        table.rows.push(marker_row("{% for item in items %}"));
        table
            .rows
            .push(template_row("A {{ item }}", VMerge::Restart, true));
        table
            .rows
            .push(template_row("B {{ item }}", VMerge::Continue, false));
        table
            .rows
            .push(template_row("C {{ item }}", VMerge::Continue, false));
        table.rows.push(marker_row("{% endfor %}"));
        table.extra_xml.push((table.rows.len(), table_raw.clone()));
        document
            .document
            .body
            .content
            .push(BodyContent::Table(table));

        assert_eq!(
            document
                .render_template(&serde_json::json!({"items": ["one", "two"]}))
                .unwrap(),
            10
        );
        let bytes = document.to_bytes().unwrap();
        let reopened = Document::from_bytes(&bytes).unwrap();
        let paragraphs = reopened.paragraphs();
        let numbered = paragraphs
            .iter()
            .filter_map(|paragraph| paragraph.numbering())
            .collect::<Vec<_>>();
        assert_eq!(numbered.len(), 2);
        assert_eq!(numbered[0], numbered[1]);
        assert_eq!(numbered[0].1, 1);

        let table = reopened.document.body.tables().next().unwrap();
        assert_eq!(table.rows.len(), 6);
        assert_eq!(table.grid.as_ref().unwrap().columns.len(), 3);
        assert_eq!(
            table
                .properties
                .as_ref()
                .unwrap()
                .look
                .as_ref()
                .unwrap()
                .no_h_band,
            Some(false)
        );
        assert_eq!(table.extra_xml, vec![(6, table_raw)]);
        for (index, row) in table.rows.iter().enumerate() {
            assert_eq!(row.extra_xml, vec![(0, row_raw.clone())]);
            assert_eq!(
                row.properties.as_ref().unwrap().header,
                (index % 3 == 0).then_some(true)
            );
            let cell = &row.cells[0];
            assert_eq!(cell.extra_xml, vec![(0, cell_raw.clone())]);
            let CellContent::ContentControl(control) = &cell.content[1] else {
                panic!("expected repeated cell content control");
            };
            assert_eq!(
                control.properties.as_ref().unwrap().tag.as_deref(),
                Some("row-control")
            );
            let paragraph = control
                .content
                .iter()
                .find_map(|content| match content {
                    SdtContent::Paragraph(paragraph) => Some(paragraph),
                    _ => None,
                })
                .expect("expected controlled paragraph");
            assert_eq!(paragraph.text(), "controlled");
            let properties = cell.properties.as_ref().unwrap();
            assert_eq!(properties.grid_span, Some(2));
            assert_eq!(
                properties.v_merge,
                Some(if index % 3 == 0 {
                    VMerge::Restart
                } else {
                    VMerge::Continue
                })
            );
            assert_eq!(properties.extra_xml, vec![(3, cell_property_raw.clone())]);
        }

        let package = OpcPackage::from_reader(std::io::Cursor::new(bytes)).unwrap();
        let xml = std::str::from_utf8(package.get_part("/word/document.xml").unwrap()).unwrap();
        assert!(xml.find("<w:trPr>").unwrap() < xml.find("<q:rowData").unwrap());
        assert!(xml.find("<w:gridSpan").unwrap() < xml.find("<q:cellProperty").unwrap());
        assert!(xml.find("<q:cellProperty").unwrap() < xml.find("<w:vMerge").unwrap());
        assert!(xml.rfind("<w:tr>").unwrap() < xml.rfind("<q:tableData").unwrap());
        assert!(xml.rfind("<q:tableData").unwrap() < xml.rfind("</w:tbl>").unwrap());
    }

    #[test]
    fn conditions_use_explicit_json_truthiness() {
        for (value, included) in [
            (serde_json::json!(false), false),
            (serde_json::Value::Null, false),
            (serde_json::json!(0), false),
            (serde_json::json!(""), false),
            (serde_json::json!([]), false),
            (serde_json::json!({}), false),
            (serde_json::json!(true), true),
            (serde_json::json!(-1), true),
            (serde_json::json!("x"), true),
            (serde_json::json!([1]), true),
            (serde_json::json!({"x": 1}), true),
        ] {
            let mut document = Document::new();
            document.add_paragraph("{% if value %}");
            document.add_paragraph("included");
            document.add_paragraph("{% endif %}");
            assert_eq!(
                document
                    .render_template(&serde_json::json!({"value": value}))
                    .unwrap(),
                0
            );
            assert_eq!(document.paragraphs().len(), usize::from(included));
        }
    }

    #[test]
    fn structural_only_table_blocks_commit_the_candidate() {
        let row = |text: &str| {
            let mut cell = CT_Tc::new();
            let mut paragraph = CT_P::new();
            paragraph.runs.push(CT_R::new(text));
            cell.content = vec![CellContent::Paragraph(paragraph)];
            let mut row = CT_Row::new();
            row.cells.push(cell);
            row
        };
        let mut table = CT_Tbl::new();
        table.rows.push(row("{% if show %}"));
        table.rows.push(row("included"));
        table.rows.push(row("{% endif %}"));
        let mut document = Document::new();
        document
            .document
            .body
            .content
            .push(BodyContent::Table(table));

        assert_eq!(
            document
                .render_template(&serde_json::json!({"show": false}))
                .unwrap(),
            0
        );
        let BodyContent::Table(table) = &document.document.body.content[0] else {
            panic!("expected table");
        };
        assert!(table.rows.is_empty());
    }

    #[test]
    fn nested_table_controls_stay_with_their_direct_table() {
        let row = |text: &str| {
            let mut cell = CT_Tc::new();
            let mut paragraph = CT_P::new();
            paragraph.runs.push(CT_R::new(text));
            cell.content = vec![CellContent::Paragraph(paragraph)];
            let mut row = CT_Row::new();
            row.cells.push(cell);
            row
        };
        let mut nested = CT_Tbl::new();
        nested.rows.push(row("{% if show %}"));
        nested.rows.push(row("nested"));
        nested.rows.push(row("{% endif %}"));
        let mut outer_cell = CT_Tc::new();
        let mut outer_paragraph = CT_P::new();
        outer_paragraph.runs.push(CT_R::new("outer"));
        outer_cell
            .content
            .push(CellContent::Paragraph(outer_paragraph));
        outer_cell.content.push(CellContent::Table(nested));
        let mut outer_row = CT_Row::new();
        outer_row.cells.push(outer_cell);
        let mut outer = CT_Tbl::new();
        outer.rows.push(outer_row);
        let mut document = Document::new();
        document
            .document
            .body
            .content
            .push(BodyContent::Table(outer));

        document
            .render_template(&serde_json::json!({"show": true}))
            .unwrap();
        let BodyContent::Table(outer) = &document.document.body.content[0] else {
            panic!("expected outer table");
        };
        let nested = outer.rows[0].cells[0]
            .content
            .iter()
            .find_map(|content| match content {
                CellContent::Table(table) => Some(table),
                CellContent::Paragraph(_) | CellContent::ContentControl(_) => None,
            })
            .expect("expected nested table");
        assert_eq!(nested.rows.len(), 1);
        assert_eq!(nested.rows[0].cells[0].text(), "nested");
    }

    #[test]
    fn direct_row_markers_reject_nested_table_content() {
        let mut nested_cell = CT_Tc::new();
        let mut nested_paragraph = CT_P::new();
        nested_paragraph.runs.push(CT_R::new("must remain"));
        nested_cell.content = vec![CellContent::Paragraph(nested_paragraph)];
        let mut nested_row = CT_Row::new();
        nested_row.cells.push(nested_cell);
        let mut nested = CT_Tbl::new();
        nested.rows.push(nested_row);

        let mut marker_cell = CT_Tc::new();
        let mut marker = CT_P::new();
        marker.runs.push(CT_R::new("{% if show %}"));
        marker_cell.content = vec![CellContent::Paragraph(marker)];
        let mut nested_cell = CT_Tc::new();
        nested_cell.content = vec![CellContent::Table(nested)];
        let mut marker_row = CT_Row::new();
        marker_row.cells.push(marker_cell);
        marker_row.cells.push(nested_cell);
        let mut table = CT_Tbl::new();
        table.rows.push(marker_row);
        for text in ["included", "{% endif %}"] {
            let mut paragraph = CT_P::new();
            paragraph.runs.push(CT_R::new(text));
            let mut cell = CT_Tc::new();
            cell.content = vec![CellContent::Paragraph(paragraph)];
            let mut row = CT_Row::new();
            row.cells.push(cell);
            table.rows.push(row);
        }
        let mut document = Document::new();
        document
            .document
            .body
            .content
            .push(BodyContent::Table(table));
        let before = document.document.to_xml().unwrap();

        assert!(
            document
                .render_template(&serde_json::json!({"show": true}))
                .is_err()
        );
        assert_eq!(document.document.to_xml().unwrap(), before);
    }

    #[test]
    fn false_condition_scalar_leaves_are_preflighted() {
        let mut document = Document::new();
        document.add_paragraph("{% if show %}");
        document.add_paragraph("{{ missing }}");
        document.add_paragraph("{% endif %}");
        let before = document.document.to_xml().unwrap();

        assert!(
            document
                .render_template(&serde_json::json!({"show": false}))
                .is_err()
        );
        assert_eq!(document.document.to_xml().unwrap(), before);
    }

    #[test]
    fn empty_loop_bodies_are_preflighted_without_requiring_an_item() {
        let mut valid = Document::new();
        valid.add_paragraph("{% for item in items %}");
        valid.add_paragraph("{{ item.name }}");
        valid.add_paragraph("{% endfor %}");
        assert_eq!(
            valid
                .render_template(&serde_json::json!({"items": []}))
                .unwrap(),
            0
        );
        assert!(valid.paragraphs().is_empty());

        for invalid_body in ["{{ item.name", "{{ config.missing }}"] {
            let mut document = Document::new();
            document.add_paragraph("{% for item in items %}");
            document.add_paragraph(invalid_body);
            document.add_paragraph("{% endfor %}");
            let before = document.document.to_xml().unwrap();

            assert!(
                document
                    .render_template(&serde_json::json!({"items": [], "config": {}}))
                    .is_err()
            );
            assert_eq!(document.document.to_xml().unwrap(), before);
        }

        let mut nested = Document::new();
        for text in [
            "{% for item in items %}",
            "{% for child in config.missing %}",
            "{{ child.name }}",
            "{% endfor %}",
            "{% endfor %}",
        ] {
            nested.add_paragraph(text);
        }
        let before = nested.document.to_xml().unwrap();
        assert!(
            nested
                .render_template(&serde_json::json!({"items": [], "config": {}}))
                .is_err()
        );
        assert_eq!(nested.document.to_xml().unwrap(), before);

        let mut nested_scope = Document::new();
        for text in [
            "{% for item in items %}",
            "{% for child in item.children %}",
            "{{ settings.missing }}",
            "{% endfor %}",
            "{% endfor %}",
        ] {
            nested_scope.add_paragraph(text);
        }
        let before = nested_scope.document.to_xml().unwrap();
        assert!(
            nested_scope
                .render_template(&serde_json::json!({
                    "items": [{"children": [], "settings": {}}]
                }))
                .is_err()
        );
        assert_eq!(nested_scope.document.to_xml().unwrap(), before);
    }

    #[test]
    fn repeated_table_level_controls_use_the_row_loop_scope() {
        let xml = br#"<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl><w:tr><w:tc><w:p><w:r><w:t>{% for item in items %}</w:t></w:r></w:p></w:tc></w:tr><w:sdt><w:sdtPr><w:tag w:val="row-control"/></w:sdtPr><w:sdtContent><w:tr><w:tc><w:p><w:r><w:t>control {{ item.name }}</w:t></w:r></w:p></w:tc></w:tr></w:sdtContent></w:sdt><w:tr><w:tc><w:p><w:r><w:t>row {{ item.name }}</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>{% endfor %}</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>"#;
        let mut document = Document::new();
        document.document = CT_Document::from_xml(xml).unwrap();

        assert_eq!(
            document
                .render_template(&serde_json::json!({
                    "items": [{"name": "one"}, {"name": "two"}]
                }))
                .unwrap(),
            4
        );
        let reopened = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
        let BodyContent::Table(table) = &reopened.document.body.content[0] else {
            panic!("expected table");
        };
        assert_eq!(
            table
                .rows
                .iter()
                .map(|row| row.cells[0].text())
                .collect::<Vec<_>>(),
            ["row one", "row two"]
        );
        assert_eq!(table.content_controls.len(), 2);
        assert_eq!(
            table
                .content_controls
                .iter()
                .map(|(_, _, control)| {
                    control
                        .content
                        .iter()
                        .find_map(|content| match content {
                            SdtContent::Row(row) => Some(row.cells[0].text()),
                            _ => None,
                        })
                        .expect("expected controlled row")
                })
                .collect::<Vec<_>>(),
            ["control one", "control two"]
        );
    }

    fn template_test_row(text: &str) -> CT_Row {
        let mut paragraph = CT_P::new();
        paragraph.runs.push(CT_R::new(text));
        let mut cell = CT_Tc::new();
        cell.content.push(CellContent::Paragraph(paragraph));
        let mut row = CT_Row::new();
        row.cells.push(cell);
        row
    }

    fn template_test_row_control(row: CT_Row) -> CT_Sdt {
        let xml = br#"<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:sdt><w:sdtContent><w:p><w:r><w:t>placeholder</w:t></w:r></w:p></w:sdtContent></w:sdt><w:sectPr/></w:body></w:document>"#;
        let mut parsed = CT_Document::from_xml(xml).unwrap();
        let BodyContent::ContentControl(mut control) = parsed.body.content.remove(0) else {
            panic!("expected parsed content control");
        };
        control.content = vec![SdtContent::Row(row)];
        control
    }

    fn template_test_table_with_control(open: &str, control: CT_Sdt, close: &str) -> Document {
        let mut table = CT_Tbl::new();
        table.rows.push(template_test_row(open));
        table.rows.push(template_test_row("source"));
        table.rows.push(template_test_row(close));
        table.content_controls.push((1, 0, control));
        let mut document = Document::new();
        document
            .document
            .body
            .content
            .push(BodyContent::Table(table));
        document
    }

    #[test]
    fn excluded_table_level_controls_are_preflighted() {
        let cases = [
            (
                "{% for item in items %}",
                "{% endfor %}",
                "{{ config.missing }}",
                serde_json::json!({"items": [], "config": {}}),
            ),
            (
                "{% if show %}",
                "{% endif %}",
                "{{ unclosed",
                serde_json::json!({"show": false}),
            ),
        ];
        for (open, close, control_text, data) in cases {
            let control = template_test_row_control(template_test_row(control_text));
            let mut document = template_test_table_with_control(open, control, close);
            let before = document.document.to_xml().unwrap();

            assert!(document.render_template(&data).is_err());
            assert_eq!(document.document.to_xml().unwrap(), before);
        }
    }

    #[test]
    fn repeated_table_level_controls_validate_numbering() {
        let mut numbered_row = template_test_row("{{ item.name }}");
        let CellContent::Paragraph(paragraph) = &mut numbered_row.cells[0].content[0] else {
            panic!("expected paragraph");
        };
        paragraph.properties = Some(CT_PPr {
            num_id: Some(99),
            num_ilvl: Some(0),
            ..Default::default()
        });
        let control = template_test_row_control(numbered_row);
        let mut document =
            template_test_table_with_control("{% for item in items %}", control, "{% endfor %}");
        let before = document.document.to_xml().unwrap();

        assert!(
            document
                .render_template(&serde_json::json!({"items": [{"name": "one"}]}))
                .is_err()
        );
        assert_eq!(document.document.to_xml().unwrap(), before);
    }

    #[test]
    fn replace_text_in_header_and_footer() {
        let mut doc = Document::new();
        assert!(!doc.has_header_footer_content());
        doc.set_header("Header: {{title}}");
        doc.set_footer("Footer: {{title}}");
        doc.add_paragraph("Body: {{title}}");

        assert!(doc.has_header_footer_content());

        let count = doc.replace_text("{{title}}", "My Doc");
        assert_eq!(count, 3);

        assert_eq!(doc.paragraphs()[0].text(), "Body: My Doc");
        assert_eq!(doc.header_text().unwrap(), "Header: My Doc");
        assert_eq!(doc.footer_text().unwrap(), "Footer: My Doc");

        let reopened = Document::from_bytes(&doc.to_bytes().unwrap()).unwrap();
        assert!(reopened.has_header_footer_content());
    }

    #[test]
    fn header_footer_content_includes_earlier_sections_after_round_trip() {
        let mut doc = Document::new();
        doc.set_header("Earlier section header");
        doc.add_paragraph("First section");

        let header_reference = doc
            .document
            .body
            .sect_pr
            .as_mut()
            .expect("final section")
            .header_refs
            .pop()
            .expect("header reference");
        let BodyContent::Paragraph(paragraph) = doc
            .document
            .body
            .content
            .last_mut()
            .expect("first section paragraph")
        else {
            panic!("expected paragraph");
        };
        paragraph.properties = Some(CT_PPr {
            sect_pr: Some(CT_SectPr {
                header_refs: vec![header_reference],
                ..CT_SectPr::default_letter()
            }),
            ..Default::default()
        });

        assert!(doc.has_header_footer_content());

        let reopened = Document::from_bytes(&doc.to_bytes().unwrap()).unwrap();
        assert!(reopened.has_header_footer_content());
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
        assert_eq!(
            doc.bookmarks()
                .iter()
                .filter_map(|bookmark| bookmark.name())
                .collect::<Vec<_>>(),
            vec!["_Toc1", "_Toc2", "_Toc3"]
        );

        // Verify round-trip: save and re-open
        let bytes = doc.to_bytes().expect("should serialize");
        let doc2 = Document::from_bytes(&bytes).expect("should open");
        assert_eq!(doc2.content_count(), 11);
        let paras2 = doc2.paragraphs();
        assert_eq!(paras2[0].text(), "Table of Contents");
    }

    #[test]
    fn insert_toc_avoids_an_existing_bookmark_id() {
        let mut doc = Document::new();
        doc.add_paragraph("Chapter").style("Heading1");
        let BodyContent::Paragraph(paragraph) = &mut doc.document.body.content[0] else {
            panic!("heading paragraph");
        };
        assert!(paragraph.insert_bookmark_start(0, 100, "existing"));
        assert!(paragraph.insert_bookmark_end(1, 100));

        doc.insert_toc(0, 1);

        let bookmarks = doc.bookmarks();
        let existing = bookmarks
            .iter()
            .find(|bookmark| bookmark.name() == Some("existing"))
            .expect("existing bookmark");
        let toc = bookmarks
            .iter()
            .find(|bookmark| bookmark.name() == Some("_Toc1"))
            .expect("TOC bookmark");
        assert_eq!(existing.id(), Some(100));
        assert_ne!(toc.id(), existing.id());
        assert_eq!(toc.id(), Some(0));
    }

    #[test]
    fn insert_toc_handles_a_maximum_numeric_suffix_without_panicking() {
        let mut doc = Document::new();
        doc.add_paragraph("Chapter").style("Heading1");
        let BodyContent::Paragraph(paragraph) = &mut doc.document.body.content[0] else {
            panic!("heading paragraph");
        };
        assert!(paragraph.insert_bookmark_start(0, 9, "_Toc18446744073709551615"));
        assert!(paragraph.insert_bookmark_end(1, 9));

        doc.insert_toc(0, 1);

        assert!(
            doc.bookmarks()
                .iter()
                .any(|bookmark| bookmark.name() == Some("_Toc1"))
        );
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
    fn links_exposes_a_complex_field_hyperlink_target() {
        let mut doc = Document::new();
        let xml = concat!(
            r#"<w:p>"#,
            r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#,
            r#"<w:r><w:instrText> HYPERLINK &quot;https://example.test&quot; </w:instrText></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
            r#"<w:r><w:t>Cached link</w:t></w:r>"#,
            r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
            r#"</w:p>"#,
        );
        let mut reader = quick_xml::Reader::from_str(xml);
        let mut buffer = Vec::new();
        let paragraph = loop {
            match reader.read_event_into(&mut buffer).unwrap() {
                Event::Start(element) if matches_local_name(element.name().as_ref(), b"p") => {
                    break CT_P::from_xml(&mut reader).unwrap();
                }
                Event::Eof => panic!("missing paragraph"),
                _ => {}
            }
            buffer.clear();
        };
        doc.document
            .body
            .content
            .push(BodyContent::Paragraph(paragraph));

        assert_eq!(
            doc.links(),
            vec![LinkInfo {
                text: "Cached link".to_owned(),
                url: Some("https://example.test".to_owned()),
                anchor: None,
                rel_id: None,
            }]
        );
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
    fn paragraph_hard_break_and_table_cell_hyperlink_round_trip() {
        let mut doc = Document::new();
        let relationship_id = doc.add_hyperlink_relationship("https://example.com/table");

        let mut paragraph = doc.add_paragraph("");
        paragraph.add_run("before");
        paragraph.add_line_break();
        paragraph.add_run("after");

        let mut table = doc.add_table(1, 1);
        let mut cell = table.cell(0, 0).expect("cell");
        cell.remove_first_empty_paragraph();
        cell.add_paragraph("")
            .add_hyperlink("table link", &relationship_id)
            .bold(true);

        let bytes = doc.to_bytes().unwrap();
        let reopened = Document::from_bytes(&bytes).unwrap();

        assert_eq!(reopened.paragraphs()[0].text(), "before\nafter");
        let tables = reopened.tables();
        let cell = tables[0].cell(0, 0).expect("cell");
        let paragraph = cell.paragraphs().next().expect("paragraph");
        assert_eq!(paragraph.text(), "table link");
        assert!(paragraph.runs().next().expect("run").is_bold());
        let spans = paragraph.hyperlink_spans();
        assert_eq!(spans.len(), 1);
        assert_eq!(
            reopened.hyperlink_url(spans[0].2.expect("relationship id")),
            Some("https://example.com/table".to_string())
        );
    }

    #[test]
    fn writer_hyperlink_tooltip_and_table_indent_round_trip() {
        let mut doc = Document::new();
        let relationship_id = doc.add_hyperlink_relationship("https://example.com");
        doc.add_paragraph("").add_hyperlink_with_tooltip(
            "linked",
            &relationship_id,
            Some("Example site"),
        );
        doc.add_table(1, 1).set_indent(Length::twips(720));

        let bytes = doc.to_bytes().expect("document writes");
        let reopened = Document::from_bytes(&bytes).expect("document reopens");

        let BodyContent::Paragraph(paragraph) = &reopened.document.body.content[0] else {
            panic!("expected paragraph");
        };
        assert_eq!(
            paragraph.hyperlinks[0].extra_attributes,
            vec![("w:tooltip".to_string(), "Example site".to_string())]
        );
        let BodyContent::Table(table) = &reopened.document.body.content[1] else {
            panic!("expected table");
        };
        assert_eq!(
            table
                .properties
                .as_ref()
                .and_then(|properties| properties.indent.as_ref()),
            Some(&rdocx_oxml::table::CT_TblWidth::dxa(720))
        );
    }

    #[test]
    fn rejected_list_level_update_does_not_materialize_numbering() {
        let mut doc = Document::new();
        assert!(doc.numbering.is_none());

        assert!(!doc.set_list_level(999, 1, ListLevel::decimal()));

        assert!(
            doc.numbering.is_none(),
            "a rejected setter must not add an empty numbering part"
        );
    }

    #[test]
    fn custom_list_and_paragraph_numbering_enforce_the_nine_level_contract() {
        let mut doc = Document::new();
        let levels = vec![ListLevel::decimal(); 10];
        let num_id = doc.add_list_definition(&levels);
        assert_eq!(
            doc.numbering.as_ref().unwrap().abstract_nums[0]
                .levels
                .len(),
            9
        );

        let mut paragraph = doc.add_paragraph("item");
        assert!(!paragraph.set_numbering(num_id, 9));
        assert_eq!(
            paragraph.inner.properties.as_ref().and_then(|p| p.num_id),
            None
        );
        assert!(paragraph.set_numbering(num_id, 8));
        assert_eq!(
            paragraph.inner.properties.as_ref().unwrap().num_ilvl,
            Some(8)
        );
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

    #[test]
    fn layout_resolves_relationship_images_to_shared_media() {
        let png: &[u8] = &[
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48,
            0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00,
            0x00, 0x90, 0x77, 0x53, 0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41, 0x54, 0x08,
            0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00, 0x00, 0x00, 0x03, 0x00, 0x01, 0x9E, 0xDD, 0x22,
            0x71, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
        ];
        let mut document = Document::new();
        document.add_picture(png, "first.png", Length::inches(1.0), Length::inches(1.0));
        document.add_picture(png, "second.png", Length::inches(1.0), Length::inches(1.0));

        let page = document
            .layout_page(0)
            .expect("layout should succeed")
            .expect("document should have a first page");
        let images = page
            .elements
            .iter()
            .filter_map(|element| match element {
                oxml_layout::PositionedElement::Image {
                    data,
                    content_type,
                    media_id,
                    ..
                } => Some((data, content_type, media_id)),
                _ => None,
            })
            .collect::<Vec<_>>();

        assert_eq!(images.len(), 2);
        assert!(images.iter().all(|(data, _, _)| data.as_slice() == png));
        assert!(
            images
                .iter()
                .all(|(_, content_type, _)| *content_type == "image/png")
        );
        assert_eq!(*images[0].2, oxml_layout::MediaId::from_bytes(png));
        assert_eq!(images[0].2, images[1].2);
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
            extra_attributes: Vec::new(),
            extra_xml: Vec::new(),
            preserved_raw_before: None,
        });
        p.hyperlinks.push(HyperlinkSpan {
            rel_id: None,
            anchor: Some("inverted".to_string()),
            run_start: 5,
            run_end: 1,
            extra_attributes: Vec::new(),
            extra_xml: Vec::new(),
            preserved_raw_before: None,
        });

        let links = doc.links();

        assert_eq!(links.len(), 2);
        assert_eq!(links[0].text, "two");
        assert_eq!(links[1].text, "");
    }
}

#[cfg(test)]
mod watermark_tests {
    use super::*;

    const PNG: &[u8] = &[
        0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44,
        0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00, 0x90,
        0x77, 0x53, 0xde, 0x00, 0x00, 0x00, 0x0c, 0x49, 0x44, 0x41, 0x54, 0x08, 0xd7, 0x63, 0xf8,
        0xcf, 0xc0, 0x00, 0x00, 0x00, 0x02, 0x00, 0x01, 0xe2, 0x21, 0xbc, 0x33, 0x00, 0x00, 0x00,
        0x00, 0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
    ];

    fn text_vml_header(text: &str, color: &str) -> Vec<u8> {
        format!(
            r#"<w:hdr xmlns:w="{}" xmlns:v="urn:schemas-microsoft-com:vml"><w:p><w:r><w:pict><v:shape style="width:468pt;height:117pt;rotation:315" fillcolor="{color}"><v:fill opacity=".5"/><v:textpath string="{text}" style="font-family:&quot;Calibri&quot;"/></v:shape></w:pict></w:r></w:p></w:hdr>"#,
            rdocx_oxml::namespace::W_NS
        )
        .into_bytes()
    }

    fn page_text(layout: &oxml_layout::LayoutResult, index: usize) -> String {
        layout.pages[index]
            .elements
            .iter()
            .filter_map(|element| match element {
                oxml_layout::PositionedElement::Text(run) => Some(run.text.as_str()),
                _ => None,
            })
            .collect::<Vec<_>>()
            .concat()
    }

    fn enable_even_headers(document: &mut Document) {
        set_even_headers_value(document, None);
    }

    fn set_even_headers_value(document: &mut Document, value: Option<&str>) {
        let part_name = "/word/settings.xml";
        let setting = value.map_or_else(
            || "<w:evenAndOddHeaders/>".to_owned(),
            |value| format!(r#"<w:evenAndOddHeaders w:val="{value}"/>"#),
        );
        let xml = format!(
            r#"<w:settings xmlns:w="{}">{setting}</w:settings>"#,
            rdocx_oxml::namespace::W_NS
        )
        .into_bytes();
        document.settings = Some(CT_Settings::from_xml(&xml).unwrap());
        document.settings_part_name = Some(part_name.to_owned());
        document.package.set_part(part_name, xml);
        document.ensure_part_relationship(
            part_name,
            rel_types::SETTINGS,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
        );
    }

    fn assert_watermark_mutation_preserves_reusable_engine(
        mut document: Document,
        mutate: impl FnOnce(&mut Document),
    ) {
        let layout_state = |document: &Document| {
            (
                document
                    .layout_cache
                    .lock()
                    .unwrap_or_else(std::sync::PoisonError::into_inner)
                    .is_some(),
                document
                    .normal_layout_engine
                    .lock()
                    .unwrap_or_else(std::sync::PoisonError::into_inner)
                    .is_some(),
                document
                    .deterministic_layout_cache
                    .lock()
                    .unwrap_or_else(std::sync::PoisonError::into_inner)
                    .is_some(),
            )
        };
        document.add_paragraph("unchanged cacheable body paragraph");

        LAYOUT_INVOCATIONS.set(0);
        let accepted_before = document.layout().expect("populate normal layout cache");
        document
            .to_pdf_deterministic()
            .expect("populate deterministic layout cache");
        assert_eq!(LAYOUT_INVOCATIONS.get(), 2);
        assert_eq!(layout_state(&document), (true, true, true));

        mutate(&mut document);

        assert_eq!(
            layout_state(&document),
            (false, true, false),
            "completed results are invalidated while reusable work survives"
        );

        let accepted_after = document
            .layout()
            .expect("relayout after watermark mutation");
        document
            .to_pdf_deterministic()
            .expect("deterministic relayout after watermark mutation");
        assert!(!Arc::ptr_eq(&accepted_before, &accepted_after));
        assert_eq!(LAYOUT_INVOCATIONS.get(), 4);
    }

    #[test]
    fn watermark_mutations_preserve_reusable_engine_and_invalidate_completed_layouts() {
        assert_watermark_mutation_preserves_reusable_engine(Document::new(), |document| {
            document.set_text_watermark("DRAFT").unwrap();
        });
        assert_watermark_mutation_preserves_reusable_engine(Document::new(), |document| {
            document
                .set_image_watermark(PNG, "watermark.png", Length::pt(72.0), Length::pt(36.0))
                .unwrap();
        });
    }

    #[test]
    fn text_and_image_watermarks_round_trip_through_header_relationships() {
        let mut text = Document::new();
        text.set_header("default header");
        text.set_first_page_header("first header");
        text.set_raw_header_with_images(
            format!(
                r#"<w:hdr xmlns:w="{}"><w:p><w:r><w:t>even header</w:t></w:r></w:p></w:hdr>"#,
                rdocx_oxml::namespace::W_NS
            )
            .into_bytes(),
            &[],
            HdrFtrType::Even,
        );
        enable_even_headers(&mut text);
        text.set_text_watermark("DRAFT").unwrap();
        text.set_text_watermark("FINAL").unwrap();
        let reopened = Document::from_bytes(&text.to_bytes().unwrap()).unwrap();
        let input = reopened.build_layout_input();
        assert_eq!(input.headers.len(), 3);
        assert!(input.headers.values().all(|header| {
            matches!(
                header.watermarks(),
                [rdocx_oxml::header_footer::VmlWatermark::Text { text, .. }]
                    if text == "FINAL"
            )
        }));

        let mut image = Document::new();
        image
            .set_image_watermark(PNG, "watermark.png", Length::pt(72.0), Length::pt(36.0))
            .unwrap();
        let reopened = Document::from_bytes(&image.to_bytes().unwrap()).unwrap();
        assert!(!reopened.build_layout_input().headers.is_empty());

        let mut invalid = Document::new();
        invalid.set_header("unchanged");
        let before = invalid.to_bytes().unwrap();
        assert!(
            invalid
                .set_image_watermark(PNG, "bad.png", Length::emu(0), Length::pt(36.0))
                .is_err()
        );
        assert_eq!(invalid.to_bytes().unwrap(), before);
    }

    #[test]
    fn header_image_relationship_ids_are_scoped_per_part() {
        let mut document = Document::new();
        let image_watermark = |text: &str| {
            format!(
                r#"<w:hdr xmlns:w="{}" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:p><w:r><w:pict><v:shape style="width:72pt;height:36pt"><v:fill opacity=".5"/><v:imagedata r:id="rId1"/></v:shape></w:pict><w:t>{text}</w:t></w:r></w:p></w:hdr>"#,
                rdocx_oxml::namespace::W_NS
            )
            .into_bytes()
        };
        let mut alternate = PNG.to_vec();
        let last = alternate.len() - 1;
        alternate[last] ^= 1;
        document.set_raw_header_with_images(
            image_watermark("default"),
            &[("rId1", PNG, "default.png")],
            HdrFtrType::Default,
        );
        document.set_raw_header_with_images(
            image_watermark("first"),
            &[("rId1", &alternate, "first.png")],
            HdrFtrType::First,
        );
        let input = document.build_layout_input();
        let scoped = input
            .images
            .iter()
            .filter(|(id, _)| id.ends_with("\0rId1"))
            .map(|(_, image)| image.data.as_slice())
            .collect::<Vec<_>>();
        assert_eq!(scoped.len(), 2);
        assert!(scoped.contains(&PNG));
        assert!(scoped.contains(&alternate.as_slice()));
    }

    #[test]
    fn watermark_covers_title_page_fallback_and_every_section() {
        let mut document = Document::new();
        document.add_paragraph("first section");
        let BodyContent::Paragraph(first) = document.document.body.content.last_mut().unwrap()
        else {
            panic!("expected first section paragraph");
        };
        first.properties = Some(CT_PPr {
            sect_pr: Some(CT_SectPr::default_letter()),
            ..Default::default()
        });
        document.add_paragraph("final section");
        document.set_different_first_page(true);

        document.set_text_watermark("DRAFT").unwrap();
        let sections = document
            .document
            .body
            .content
            .iter()
            .filter_map(|content| match content {
                BodyContent::Paragraph(paragraph) => paragraph
                    .properties
                    .as_ref()
                    .and_then(|properties| properties.sect_pr.as_ref()),
                _ => None,
            })
            .chain(document.document.body.sect_pr.iter())
            .collect::<Vec<_>>();
        assert_eq!(sections.len(), 2);
        assert!(
            sections[0]
                .header_refs
                .iter()
                .any(|reference| reference.hdr_ftr_type == HdrFtrType::Default)
        );
        assert!(
            !sections[1]
                .header_refs
                .iter()
                .any(|reference| reference.hdr_ftr_type == HdrFtrType::Default)
        );
        assert!(
            sections[1]
                .header_refs
                .iter()
                .any(|reference| reference.hdr_ftr_type == HdrFtrType::First)
        );

        let layout = document
            .layout_for_options(RenderOptions::default(), true)
            .unwrap()
            .layout
            .clone();
        assert_eq!(layout.pages.len(), 2);
        assert!(layout.pages.iter().all(|page| {
            matches!(
                page.elements.first(),
                Some(oxml_layout::PositionedElement::Group(_))
            )
        }));
    }

    #[test]
    fn saved_watermark_materializes_first_and_even_fallback_headers() {
        let mut document = Document::new();
        document.set_different_first_page(true);
        enable_even_headers(&mut document);
        document.set_text_watermark("DRAFT").unwrap();

        let reopened = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
        let section = reopened.section_properties().unwrap();
        for hdr_type in [HdrFtrType::Default, HdrFtrType::First, HdrFtrType::Even] {
            let reference = section
                .header_refs
                .iter()
                .find(|reference| reference.hdr_ftr_type == hdr_type)
                .unwrap_or_else(|| panic!("missing saved {hdr_type:?} header"));
            let header = reopened.load_header_footer(&reference.rel_id).unwrap();
            assert!(matches!(
                header.watermarks(),
                [rdocx_oxml::header_footer::VmlWatermark::Text { text, .. }]
                    if text == "DRAFT"
            ));
        }
    }

    #[test]
    fn inherited_ordinary_header_remains_inherited_after_watermarking() {
        let mut document = Document::new();
        document.set_header("Company header");
        document.add_paragraph("first section");
        let company = document
            .section_properties_mut()
            .header_refs
            .pop()
            .expect("company header reference");
        let BodyContent::Paragraph(first) = document.document.body.content.last_mut().unwrap()
        else {
            panic!("expected first section paragraph");
        };
        first.properties = Some(CT_PPr {
            sect_pr: Some(CT_SectPr {
                header_refs: vec![company],
                ..CT_SectPr::default_letter()
            }),
            ..Default::default()
        });
        document.add_paragraph("second section");

        document.set_text_watermark("DRAFT").unwrap();
        assert!(
            document
                .section_properties()
                .unwrap()
                .header_refs
                .is_empty()
        );
        let reopened = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
        assert!(
            reopened
                .section_properties()
                .unwrap()
                .header_refs
                .is_empty()
        );
        let layout = reopened
            .layout_for_options(RenderOptions::default(), true)
            .unwrap()
            .layout
            .clone();
        assert_eq!(layout.pages.len(), 2);
        assert!(page_text(&layout, 1).contains("Company header"));
        assert!(matches!(
            layout.pages[1].elements.first(),
            Some(oxml_layout::PositionedElement::Group(_))
        ));
    }

    #[test]
    fn section_page_number_restart_controls_even_header_parity() {
        let mut document = Document::new();
        document.add_paragraph("first section");
        let BodyContent::Paragraph(first) = document.document.body.content.last_mut().unwrap()
        else {
            panic!("expected first section paragraph");
        };
        first.properties = Some(CT_PPr {
            sect_pr: Some(CT_SectPr::default_letter()),
            ..Default::default()
        });
        document.add_paragraph("second section");
        document.set_header("default header");
        document.set_raw_header_with_images(
            format!(
                r#"<w:hdr xmlns:w="{}"><w:p><w:r><w:t>even header</w:t></w:r></w:p></w:hdr>"#,
                rdocx_oxml::namespace::W_NS
            )
            .into_bytes(),
            &[],
            HdrFtrType::Even,
        );
        enable_even_headers(&mut document);
        document
            .section_properties_mut()
            .extra_xml
            .push(br#"<w:pgNumType w:start="1"/>"#.to_vec());
        document.set_text_watermark("DRAFT").unwrap();

        let reopened = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
        let layout = reopened
            .layout_for_options(RenderOptions::default(), true)
            .unwrap()
            .layout
            .clone();
        assert_eq!(layout.pages.len(), 2);
        assert!(page_text(&layout, 1).contains("default header"));
        assert!(!page_text(&layout, 1).contains("even header"));
    }

    #[test]
    fn blank_first_and_even_variants_do_not_borrow_default_content() {
        let mut first = Document::new();
        first.set_raw_header_with_images(
            format!(
                r#"<w:hdr xmlns:w="{}" xmlns:v="urn:schemas-microsoft-com:vml"><w:p><w:r><w:pict><v:shape style="width:468pt;height:117pt;rotation:315" fillcolor="D9D9D9"><v:fill opacity=".5"/><v:textpath string="DRAFT"/></v:shape></w:pict><w:t>company header</w:t></w:r></w:p></w:hdr>"#,
                rdocx_oxml::namespace::W_NS
            )
            .into_bytes(),
            &[],
            HdrFtrType::Default,
        );
        first.set_footer("company footer");
        first.set_first_page_header("");
        first.set_first_page_footer("");
        first.add_paragraph("body");
        let first_layout = first
            .layout_for_options(RenderOptions::default(), true)
            .unwrap()
            .layout
            .clone();
        assert!(
            !first_layout.pages[0]
                .elements
                .iter()
                .any(|element| matches!(element, oxml_layout::PositionedElement::Group(_)))
        );
        assert!(!page_text(&first_layout, 0).contains("company header"));
        assert!(!page_text(&first_layout, 0).contains("company footer"));

        let mut even = Document::new();
        even.set_header("company header");
        even.set_footer("company footer");
        even.set_raw_footer_with_images(
            format!(r#"<w:ftr xmlns:w="{}"/>"#, rdocx_oxml::namespace::W_NS).into_bytes(),
            &[],
            HdrFtrType::Even,
        );
        enable_even_headers(&mut even);
        even.set_text_watermark("DRAFT").unwrap();
        even.add_paragraph(&"body ".repeat(4_000));
        let even_layout = even
            .layout_for_options(RenderOptions::default(), true)
            .unwrap()
            .layout
            .clone();
        assert!(matches!(
            even_layout.pages[1].elements.first(),
            Some(oxml_layout::PositionedElement::Group(_))
        ));
        assert!(!page_text(&even_layout, 1).contains("company header"));
        assert!(!page_text(&even_layout, 1).contains("company footer"));
    }

    #[test]
    fn image_watermark_target_is_relative_to_a_custom_header_part() {
        let mut document = Document::new();
        let header_part = "/custom/headers/header.xml";
        document.package.set_part(
            header_part,
            format!(
                r#"<w:hdr xmlns:w="{}"><w:p/></w:hdr>"#,
                rdocx_oxml::namespace::W_NS
            )
            .into_bytes(),
        );
        document.package.content_types.add_override(
            header_part,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
        );
        let target = relative_target(&document.doc_part_name, header_part);
        let rel_id = document
            .package
            .get_or_create_part_rels(&document.doc_part_name)
            .add(rel_types::HEADER, &target);
        document.section_properties_mut().header_refs = vec![HdrFtrRef {
            hdr_ftr_type: HdrFtrType::Default,
            rel_id,
        }];

        document
            .set_image_watermark(PNG, "custom.png", Length::pt(72.0), Length::pt(36.0))
            .unwrap();
        let image_relationship = document
            .package
            .get_part_rels(header_part)
            .unwrap()
            .items
            .iter()
            .find(|relationship| relationship.rel_type == rel_types::IMAGE)
            .unwrap();
        let image_part = OpcPackage::resolve_rel_target(header_part, &image_relationship.target);
        assert_eq!(document.package.get_part(&image_part), Some(PNG));
    }

    #[test]
    fn named_vml_colour_renders_its_defined_rgb_value() {
        let mut document = Document::new();
        document.set_raw_header_with_images(
            text_vml_header("DRAFT", "silver"),
            &[],
            HdrFtrType::Default,
        );
        document.add_paragraph("body");
        let layout = document
            .layout_for_options(RenderOptions::default(), true)
            .unwrap()
            .layout
            .clone();
        let oxml_layout::PositionedElement::Group(group) = &layout.pages[0].elements[0] else {
            panic!("expected watermark group");
        };
        let oxml_layout::PositionedElement::Text(run) = &group.children[0] else {
            panic!("expected watermark text");
        };
        assert_eq!(
            run.color,
            oxml_layout::Color {
                r: 192.0 / 255.0,
                g: 192.0 / 255.0,
                b: 192.0 / 255.0,
                a: 1.0,
            }
        );
    }

    #[test]
    fn invalid_multibyte_vml_colour_is_suppressed_without_panicking() {
        let mut document = Document::new();
        document.set_raw_header_with_images(
            text_vml_header("DRAFT", "€€"),
            &[],
            HdrFtrType::Default,
        );
        document.add_paragraph("body");
        let layout = document
            .layout_for_options(RenderOptions::default(), true)
            .unwrap()
            .layout
            .clone();
        assert!(
            !layout.pages[0]
                .elements
                .iter()
                .any(|element| matches!(element, oxml_layout::PositionedElement::Group(_)))
        );
        assert!(
            layout
                .diagnostics
                .iter()
                .any(|diagnostic| diagnostic.message.contains("colour")
                    && diagnostic.message.contains("unsupported"))
        );
    }

    #[test]
    fn even_header_selection_follows_the_document_setting() {
        let render = |setting: Option<Option<&str>>| {
            let mut document = Document::new();
            document.set_header("default header");
            document.set_raw_header_with_images(
                format!(
                    r#"<w:hdr xmlns:w="{}"><w:p><w:r><w:t>even header</w:t></w:r></w:p></w:hdr>"#,
                    rdocx_oxml::namespace::W_NS
                )
                .into_bytes(),
                &[],
                HdrFtrType::Even,
            );
            if let Some(value) = setting {
                set_even_headers_value(&mut document, value);
            }
            document.add_paragraph(&"body ".repeat(4_000));
            document
                .layout_for_options(RenderOptions::default(), true)
                .unwrap()
                .layout
                .clone()
        };
        let disabled = render(None);
        let entity_false = render(Some(Some("&#48;")));
        let enabled = render(Some(None));
        assert!(page_text(&disabled, 1).contains("default header"));
        assert!(!page_text(&disabled, 1).contains("even header"));
        assert!(page_text(&entity_false, 1).contains("default header"));
        assert!(!page_text(&entity_false, 1).contains("even header"));
        assert!(page_text(&enabled, 1).contains("even header"));
    }

    #[test]
    fn unresolved_image_watermark_is_suppressed_with_a_diagnostic() {
        let mut document = Document::new();
        let xml = format!(
            r#"<w:hdr xmlns:w="{}" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:p><w:r><w:pict><v:shape style="width:72pt;height:36pt"><v:imagedata r:id="rIdMissing"/></v:shape></w:pict></w:r></w:p></w:hdr>"#,
            rdocx_oxml::namespace::W_NS
        );
        document.set_raw_header_with_images(xml.into_bytes(), &[], HdrFtrType::Default);
        document.add_paragraph("body");
        let layout = document
            .layout_for_options(RenderOptions::default(), true)
            .unwrap()
            .layout
            .clone();
        assert!(
            !layout.pages[0]
                .elements
                .iter()
                .any(|element| matches!(element, oxml_layout::PositionedElement::Group(_)))
        );
        assert!(layout.diagnostics.iter().any(|diagnostic| {
            diagnostic.message.contains("rIdMissing") && diagnostic.message.contains("not resolved")
        }));
    }

    #[test]
    fn watermark_is_centered_in_the_margin_rectangle() {
        let mut document = Document::new();
        document.set_margins(
            Length::pt(36.0),
            Length::pt(144.0),
            Length::pt(108.0),
            Length::pt(72.0),
        );
        document.set_text_watermark("DRAFT").unwrap();
        document.add_paragraph("body");
        let layout = document
            .layout_for_options(RenderOptions::default(), true)
            .unwrap()
            .layout
            .clone();
        let oxml_layout::PositionedElement::Group(group) = &layout.pages[0].elements[0] else {
            panic!("expected watermark group");
        };
        let center = group.transform.apply(oxml_layout::Point {
            x: 468.0 / 2.0,
            y: 117.0 / 2.0,
        });
        assert!((center.x - (72.0 + (612.0 - 72.0 - 144.0) / 2.0)).abs() < 1.0e-9);
        assert!((center.y - (36.0 + (792.0 - 36.0 - 108.0) / 2.0)).abs() < 1.0e-9);
    }

    #[test]
    fn watermark_group_precedes_body_elements_on_every_page() {
        let mut document = Document::new();
        document.set_header("default header");
        document.set_first_page_header("first header");
        document.set_raw_header_with_images(
            format!(
                r#"<w:hdr xmlns:w="{}"><w:p><w:r><w:t>even header</w:t></w:r></w:p></w:hdr>"#,
                rdocx_oxml::namespace::W_NS
            )
            .into_bytes(),
            &[],
            HdrFtrType::Even,
        );
        enable_even_headers(&mut document);
        document.set_text_watermark("DRAFT").unwrap();
        document.add_paragraph(&"body ".repeat(4_000));
        let layout = document
            .layout_for_options(RenderOptions::default(), true)
            .unwrap()
            .layout
            .clone();
        assert!(layout.pages.len() > 2);
        assert!(layout.pages.iter().all(|page| {
            matches!(
                page.elements.first(),
                Some(oxml_layout::PositionedElement::Group(_))
            )
        }));
        assert!(
            page_text(&layout, 0).contains("first header"),
            "{:?}",
            page_text(&layout, 0)
        );
        assert!(
            page_text(&layout, 1).contains("even header"),
            "{:?}",
            page_text(&layout, 1)
        );
        assert!(
            page_text(&layout, 2).contains("default header"),
            "{:?}",
            page_text(&layout, 2)
        );
    }

    #[test]
    fn watermark_renders_behind_body_text_on_every_page() {
        let mut document = Document::new();
        document.set_text_watermark("DRAFT").unwrap();
        document.add_paragraph(&"body ".repeat(4_000));
        let layout = document
            .layout_for_options(RenderOptions::default(), true)
            .unwrap()
            .layout
            .clone();
        let pngs = oxml_pdf::render_all_pages(&layout, 72.0);
        assert!(pngs.len() > 1);
        assert!(pngs.iter().all(|png| png.starts_with(b"\x89PNG\r\n\x1a\n")));
        let digests = pngs
            .iter()
            .map(|png| {
                png.iter().fold(0xcbf2_9ce4_8422_2325_u64, |hash, byte| {
                    (hash ^ u64::from(*byte)).wrapping_mul(0x0000_0100_0000_01b3)
                })
            })
            .collect::<Vec<_>>();
        assert_eq!(
            digests,
            [
                740_018_920_125_384_146,
                740_018_920_125_384_146,
                740_018_920_125_384_146,
                740_018_920_125_384_146,
                1_020_215_290_976_271_429,
            ]
        );
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
