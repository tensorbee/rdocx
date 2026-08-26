//! Bounded OpenDocument Text import and export for the native Word document model.

use std::collections::{BTreeMap, BTreeSet, HashMap, HashSet};
use std::fmt::Write as _;
use std::io::{Cursor, Read, Write};
use std::path::Path;

use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;
use rdocx_oxml::document::BodyContent;
use rdocx_oxml::numbering::ST_NumberFormat;
use rdocx_oxml::properties::{CT_PPr, CT_RPr};
use rdocx_oxml::shared::{ST_HighlightColor, ST_Jc, ST_Underline};
use rdocx_oxml::table::{
    CT_Row, CT_Tbl, CT_TblGrid, CT_TblGridCol, CT_TblPr, CT_TblWidth, CT_Tc, CT_TcPr, CellContent,
    VMerge,
};
use rdocx_oxml::text::{BreakType, CT_P, CT_R, RunContent};
use rdocx_oxml::units::Twips;
use zip::CompressionMethod;
use zip::ZipWriter;
use zip::read::ZipArchive;
use zip::write::SimpleFileOptions;

use crate::paragraph::{Alignment, Paragraph};
use crate::run::Run;
use crate::{Document, Error, Length, ListLevel, PackageReadLimits, Result};

const OFFICE_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TEXT_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const STYLE_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const FO_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
const TABLE_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const DRAW_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const XLINK_NS: &str = "http://www.w3.org/1999/xlink";
const SVG_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
const MANIFEST_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";
const ODT_MIMETYPE: &[u8] = b"application/vnd.oasis.opendocument.text";

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OdtDiagnostic {
    pub path: String,
    pub message: String,
}

pub struct OdtReadResult {
    pub document: Document,
    pub diagnostics: Vec<OdtDiagnostic>,
}

/// Serialized ODT bytes together with every lossy-conversion diagnostic.
pub struct OdtWriteResult {
    pub bytes: Vec<u8>,
    pub diagnostics: Vec<OdtDiagnostic>,
}

#[derive(Clone, Copy)]
struct OdtLimits {
    archive: PackageReadLimits,
    xml_depth: usize,
    xml_nodes: usize,
    retained_text: usize,
    blocks: usize,
    runs: usize,
    rows: usize,
    columns: usize,
    cells: usize,
    diagnostics: usize,
}

impl OdtLimits {
    const DEFAULT: Self = Self {
        archive: PackageReadLimits {
            max_entries: 4_096,
            max_part_uncompressed_bytes: 64 * 1024 * 1024,
            max_total_uncompressed_bytes: 128 * 1024 * 1024,
        },
        xml_depth: 256,
        xml_nodes: 300_000,
        retained_text: 64 * 1024 * 1024,
        blocks: 100_000,
        runs: 100_000,
        rows: 10_000,
        columns: 256,
        cells: 50_000,
        diagnostics: 10_000,
    };
}

impl Document {
    /// Convert an ODT package into a fresh editable Word document.
    pub fn from_odt_bytes(bytes: &[u8]) -> Result<OdtReadResult> {
        from_odt_with_limits(bytes, OdtLimits::DEFAULT)
    }

    /// Convert an ODT package while applying caller-supplied archive bounds.
    pub fn from_odt_bytes_with_limits(
        bytes: &[u8],
        limits: PackageReadLimits,
    ) -> Result<OdtReadResult> {
        from_odt_with_limits(
            bytes,
            OdtLimits {
                archive: limits,
                ..OdtLimits::DEFAULT
            },
        )
    }

    /// Open and convert an ODT package from a path.
    pub fn open_odt<P: AsRef<Path>>(path: P) -> Result<OdtReadResult> {
        let file = std::fs::File::open(path)?;
        let declared = file.metadata()?.len();
        if declared > OdtLimits::DEFAULT.archive.max_total_uncompressed_bytes {
            return Err(odt_error(None, 0, "ODT input exceeds the size limit"));
        }
        let mut bytes = Vec::with_capacity(usize::try_from(declared).unwrap_or(0));
        file.take(
            OdtLimits::DEFAULT
                .archive
                .max_total_uncompressed_bytes
                .saturating_add(1),
        )
        .read_to_end(&mut bytes)?;
        if bytes.len() as u64 > OdtLimits::DEFAULT.archive.max_total_uncompressed_bytes {
            return Err(odt_error(None, 0, "ODT input exceeds the size limit"));
        }
        Self::from_odt_bytes(&bytes)
    }

    /// Serialize the editable document to the supported ODT subset.
    pub fn to_odt_bytes(&self) -> Result<OdtWriteResult> {
        OdtWriter::new(self).write()
    }

    /// Serialize and save ODT to a path, returning lossy-conversion diagnostics.
    pub fn save_odt<P: AsRef<Path>>(&self, path: P) -> Result<Vec<OdtDiagnostic>> {
        let result = self.to_odt_bytes()?;
        crate::document::write_atomic_file(
            path.as_ref(),
            &result.bytes,
            "invalid file name",
            "could not allocate ODT-save staging file",
        )?;
        Ok(result.diagnostics)
    }
}

#[derive(Clone)]
struct OdtMedia {
    path: String,
    media_type: &'static str,
    bytes: Vec<u8>,
}

struct OdtWriter<'a> {
    document: &'a Document,
    limits: OdtLimits,
    paragraph_styles: BTreeMap<String, String>,
    text_styles: BTreeMap<String, String>,
    list_styles: BTreeMap<u32, String>,
    used_list_levels: BTreeSet<(u32, u32)>,
    media: Vec<OdtMedia>,
    media_at: BTreeMap<String, usize>,
    diagnostics: Vec<OdtDiagnostic>,
    diagnostic_keys: BTreeSet<(String, String)>,
    retained_output_estimate: usize,
    retained_media_bytes: usize,
    output_blocks: usize,
    output_rows: usize,
    output_cells: usize,
    output_runs: usize,
}

impl<'a> OdtWriter<'a> {
    fn new(document: &'a Document) -> Self {
        Self::new_with_limits(document, OdtLimits::DEFAULT)
    }

    fn new_with_limits(document: &'a Document, limits: OdtLimits) -> Self {
        Self {
            document,
            limits,
            paragraph_styles: BTreeMap::new(),
            text_styles: BTreeMap::new(),
            list_styles: BTreeMap::new(),
            used_list_levels: BTreeSet::new(),
            media: Vec::new(),
            media_at: BTreeMap::new(),
            diagnostics: Vec::new(),
            diagnostic_keys: BTreeSet::new(),
            retained_output_estimate: 2_048,
            retained_media_bytes: 0,
            output_blocks: 0,
            output_rows: 0,
            output_cells: 0,
            output_runs: 0,
        }
    }

    fn write(mut self) -> Result<OdtWriteResult> {
        self.scan_document_losses()?;
        for (index, content) in self.document.document.body.content.iter().enumerate() {
            self.scan_body(content, &format!("body[{index}]"))?;
        }

        let content = self.content_xml()?;
        let manifest = self.manifest_xml();
        parse_xml("content.xml", content.as_bytes(), self.limits)?;
        parse_xml("META-INF/manifest.xml", manifest.as_bytes(), self.limits)?;
        let part_limit = self.limits.archive.max_part_uncompressed_bytes as usize;
        if content.len() > part_limit || manifest.len() > part_limit {
            return Err(odt_error(None, 0, "ODT XML output exceeds the size limit"));
        }
        let media_total = self.media.iter().try_fold(0_usize, |total, media| {
            if media.bytes.len() > part_limit {
                return None;
            }
            total.checked_add(media.bytes.len())
        });
        let total = media_total
            .and_then(|total| total.checked_add(content.len()))
            .and_then(|total| total.checked_add(manifest.len()))
            .and_then(|total| total.checked_add(ODT_MIMETYPE.len()))
            .ok_or_else(|| odt_error(None, 0, "ODT output exceeds the size limit"))?;
        if total as u64 > self.limits.archive.max_total_uncompressed_bytes {
            return Err(odt_error(None, 0, "ODT output exceeds the size limit"));
        }
        if self.media.len() + 3 > self.limits.archive.max_entries {
            return Err(odt_error(None, 0, "ODT output exceeds the entry limit"));
        }

        let mut output = Cursor::new(Vec::new());
        {
            let mut archive = ZipWriter::new(&mut output);
            write_odt_entry(
                &mut archive,
                "mimetype",
                ODT_MIMETYPE,
                CompressionMethod::Stored,
            )?;
            write_odt_entry(
                &mut archive,
                "content.xml",
                content.as_bytes(),
                CompressionMethod::Deflated,
            )?;
            for media in &self.media {
                write_odt_entry(
                    &mut archive,
                    &media.path,
                    &media.bytes,
                    CompressionMethod::Deflated,
                )?;
            }
            write_odt_entry(
                &mut archive,
                "META-INF/manifest.xml",
                manifest.as_bytes(),
                CompressionMethod::Deflated,
            )?;
            archive
                .finish()
                .map_err(|error| odt_error(None, 0, format!("cannot finish ODT ZIP: {error}")))?;
        }
        let bytes = output.into_inner();
        if bytes.len() as u64 > self.limits.archive.max_total_uncompressed_bytes {
            return Err(odt_error(None, 0, "ODT output exceeds the size limit"));
        }
        Ok(OdtWriteResult {
            bytes,
            diagnostics: self.diagnostics,
        })
    }

    fn scan_body(&mut self, content: &BodyContent, path: &str) -> Result<()> {
        self.reserve_output(64)?;
        match content {
            BodyContent::Paragraph(paragraph) => self.scan_paragraph(paragraph, path),
            BodyContent::Table(table) => self.scan_table(table, path),
            BodyContent::ContentControl(_) => {
                self.diagnose(path, "body content control was dropped during ODT export")
            }
            BodyContent::RawXml(_) => {
                self.diagnose(path, "unmodelled body XML was dropped during ODT export")
            }
        }
    }

    fn scan_table(&mut self, table: &CT_Tbl, path: &str) -> Result<()> {
        self.charge_output_block()?;
        self.reserve_output(256)?;
        validate_word_table(table, path)?;
        if table
            .properties
            .as_ref()
            .is_some_and(table_has_lossy_properties)
        {
            self.diagnose(
                &format!("{path}/tblPr"),
                "table properties were dropped during ODT export",
            )?;
        }
        if table
            .grid
            .as_ref()
            .is_some_and(|grid| !grid.columns.is_empty())
        {
            self.diagnose(
                &format!("{path}/tblGrid"),
                "table grid column widths were dropped during ODT export",
            )?;
        }
        for (index, _) in &table.extra_xml {
            self.diagnose(
                &format!("{path}/raw[{index}]"),
                "unmodelled table XML was dropped during ODT export",
            )?;
        }
        if !table.content_controls.is_empty() {
            self.diagnose(
                &format!("{path}/content-controls"),
                "table row content controls were dropped during ODT export",
            )?;
        }
        for (row_index, row) in table.rows.iter().enumerate() {
            self.output_rows = self
                .output_rows
                .checked_add(1)
                .filter(|rows| *rows <= self.limits.rows)
                .ok_or_else(|| odt_error(None, 0, "ODT output exceeds the row limit"))?;
            self.reserve_output(128)?;
            let row_path = format!("{path}/row[{row_index}]");
            if row.properties.is_some() {
                self.diagnose(
                    &format!("{row_path}/trPr"),
                    "table-row properties were dropped during ODT export",
                )?;
            }
            for (index, _) in &row.extra_xml {
                self.diagnose(
                    &format!("{row_path}/raw[{index}]"),
                    "unmodelled table-row XML was dropped during ODT export",
                )?;
            }
            if !row.content_controls.is_empty() {
                self.diagnose(
                    &format!("{row_path}/content-controls"),
                    "table cell content controls were dropped during ODT export",
                )?;
            }
            for (cell_index, cell) in row.cells.iter().enumerate() {
                let span = cell
                    .properties
                    .as_ref()
                    .and_then(|properties| properties.grid_span)
                    .unwrap_or(1);
                let span = usize::try_from(span)
                    .map_err(|_| odt_error(None, 0, "ODT output cell count overflowed"))?;
                self.output_cells = self
                    .output_cells
                    .checked_add(span)
                    .filter(|cells| *cells <= self.limits.cells)
                    .ok_or_else(|| odt_error(None, 0, "ODT output exceeds the cell limit"))?;
                self.reserve_output(256)?;
                let cell_path = format!("{row_path}/cell[{cell_index}]");
                if cell
                    .properties
                    .as_ref()
                    .is_some_and(cell_has_lossy_properties)
                {
                    self.diagnose(
                        &format!("{cell_path}/tcPr"),
                        "unsupported table-cell properties were dropped during ODT export",
                    )?;
                }
                for (index, _) in &cell.extra_xml {
                    self.diagnose(
                        &format!("{cell_path}/raw[{index}]"),
                        "unmodelled table-cell XML was dropped during ODT export",
                    )?;
                }
                if matches!(
                    cell.properties
                        .as_ref()
                        .and_then(|properties| properties.v_merge.as_ref()),
                    Some(VMerge::Continue)
                ) {
                    if cell_has_substantive_content(cell) {
                        self.diagnose(
                            &format!("{cell_path}/content"),
                            "vertical-merge continuation content was dropped during ODT export",
                        )?;
                    }
                    continue;
                }
                if !cell
                    .content
                    .iter()
                    .any(|content| matches!(content, CellContent::Paragraph(_)))
                {
                    self.charge_output_block()?;
                }
                for (content_index, content) in cell.content.iter().enumerate() {
                    let content_path = format!("{cell_path}/content[{content_index}]");
                    match content {
                        CellContent::Paragraph(paragraph) => {
                            self.scan_paragraph(paragraph, &content_path)?
                        }
                        CellContent::Table(_) => self.diagnose(
                            &content_path,
                            "nested table was dropped during ODT export",
                        )?,
                        CellContent::ContentControl(_) => self.diagnose(
                            &content_path,
                            "table-cell content control was dropped during ODT export",
                        )?,
                    }
                }
            }
        }
        Ok(())
    }

    fn scan_paragraph(&mut self, paragraph: &CT_P, path: &str) -> Result<()> {
        self.charge_output_block()?;
        self.reserve_output(256)?;
        let paragraph_properties = self.effective_paragraph_properties(paragraph);
        validate_paragraph_projection(&paragraph_properties, path)?;
        let paragraph_xml = paragraph_style_xml(&paragraph_properties);
        self.ensure_paragraph_style(paragraph_xml)?;
        self.scan_paragraph_losses(paragraph, &paragraph_properties, path)?;
        if let Some((num_id, level)) = paragraph_numbering_properties(&paragraph_properties) {
            if level > 8 {
                return Err(odt_error(
                    Some("content.xml"),
                    0,
                    format!("ODT list level exceeds 8 at {path}"),
                ));
            }
            if self.list_level_has_producer_format(num_id, level) {
                self.diagnose(
                    &format!("numbering[{num_id}]/level[{level}]"),
                    "producer-defined numbering format was flattened without a marker during ODT export",
                )?;
                self.scan_flattened_producer_list_losses(num_id, level)?;
            } else {
                self.used_list_levels.insert((num_id, level));
                self.ensure_list_style(num_id, path)?;
            }
        }
        let mut trailing_text_style = None;
        let mut projected_runs = 0_usize;
        for (run_index, run) in paragraph.runs.iter().enumerate() {
            self.reserve_output(192)?;
            let run_path = format!("{path}/run[{run_index}]");
            let run_properties = self.effective_run_properties(paragraph, run);
            validate_run_projection(&run_properties, &run_path)?;
            if let Some(font) = selected_run_font(&run_properties) {
                validate_xml_value(font, &format!("{run_path}/rPr/font"))?;
                validate_font_family_projection(font, &format!("{run_path}/rPr/font"))?;
            }
            let run_style = text_style_xml(&run_properties);
            self.ensure_text_style(run_style.clone())?;
            self.scan_run_losses(run, &run_properties, &run_path)?;
            for (content_index, content) in run.content.iter().enumerate() {
                if let RunContent::Drawing(drawing) = content {
                    self.scan_drawing(drawing, &format!("{run_path}/content[{content_index}]"))?;
                }
            }
            for (content_index, content) in run.content.iter().enumerate() {
                let pieces = match content {
                    RunContent::Text(text) | RunContent::DeletedText(text) => {
                        projected_text_run_pieces(&text.text, &run_style, &mut trailing_text_style)
                    }
                    RunContent::Field(field) => projected_text_run_pieces(
                        field
                            .projected_text()
                            .unwrap_or(field.cached_result.as_str()),
                        &run_style,
                        &mut trailing_text_style,
                    ),
                    RunContent::Tab | RunContent::Break(BreakType::Line) => {
                        trailing_text_style = None;
                        1
                    }
                    RunContent::Drawing(_) => {
                        if self
                            .media_at
                            .contains_key(&format!("{run_path}/content[{content_index}]"))
                        {
                            trailing_text_style = None;
                            1
                        } else {
                            0
                        }
                    }
                    RunContent::Break(_)
                    | RunContent::FootnoteRef { .. }
                    | RunContent::EndnoteRef { .. }
                    | RunContent::CommentReference { .. } => 0,
                };
                projected_runs = projected_runs
                    .checked_add(pieces)
                    .ok_or_else(|| odt_error(None, 0, "ODT output exceeds the run limit"))?;
            }
        }
        self.output_runs = self
            .output_runs
            .checked_add(projected_runs)
            .filter(|runs| *runs <= self.limits.runs)
            .ok_or_else(|| odt_error(None, 0, "ODT output exceeds the run limit"))?;
        Ok(())
    }

    fn scan_document_losses(&mut self) -> Result<()> {
        if self.document.document.background_xml.is_some() {
            self.diagnose(
                "document/background",
                "document background was dropped during ODT export",
            )?;
        }
        if self
            .document
            .document
            .body
            .sect_pr
            .as_ref()
            .is_some_and(final_section_has_unsupported_properties)
        {
            self.diagnose(
                "document/sectPr",
                "final section properties were dropped during ODT export",
            )?;
        }
        let relationships = self
            .document
            .package
            .get_part_rels(&self.document.doc_part_name);
        if let Some(section) = self.document.document.body.sect_pr.as_ref() {
            for (index, reference) in section.header_refs.iter().enumerate() {
                let relationship = relationships.and_then(|relationships| {
                    relationships
                        .items
                        .iter()
                        .find(|relationship| relationship.id == reference.rel_id)
                });
                if relationship.is_none_or(|relationship| {
                    relationship.rel_type != oxml_opc::relationship::rel_types::HEADER
                }) {
                    self.diagnose(
                        &format!("document/sectPr/headerReference[{index}]"),
                        "unresolved or wrong-type header reference was dropped during ODT export",
                    )?;
                }
            }
            for (index, reference) in section.footer_refs.iter().enumerate() {
                let relationship = relationships.and_then(|relationships| {
                    relationships
                        .items
                        .iter()
                        .find(|relationship| relationship.id == reference.rel_id)
                });
                if relationship.is_none_or(|relationship| {
                    relationship.rel_type != oxml_opc::relationship::rel_types::FOOTER
                }) {
                    self.diagnose(
                        &format!("document/sectPr/footerReference[{index}]"),
                        "unresolved or wrong-type footer reference was dropped during ODT export",
                    )?;
                }
            }
        }
        if relationships.is_some_and(|relationships| {
            relationships.items.iter().any(|relationship| {
                relationship.rel_type == oxml_opc::relationship::rel_types::HEADER
            })
        }) {
            self.diagnose(
                "document/headers",
                "header stories were dropped during ODT export",
            )?;
        }
        if relationships.is_some_and(|relationships| {
            relationships.items.iter().any(|relationship| {
                relationship.rel_type == oxml_opc::relationship::rel_types::FOOTER
            })
        }) {
            self.diagnose(
                "document/footers",
                "footer stories were dropped during ODT export",
            )?;
        }
        Ok(())
    }

    fn charge_output_block(&mut self) -> Result<()> {
        self.output_blocks = self
            .output_blocks
            .checked_add(1)
            .filter(|blocks| *blocks <= self.limits.blocks)
            .ok_or_else(|| odt_error(None, 0, "ODT output exceeds the block limit"))?;
        Ok(())
    }

    fn scan_paragraph_losses(
        &mut self,
        paragraph: &CT_P,
        effective: &CT_PPr,
        path: &str,
    ) -> Result<()> {
        if paragraph_properties_have_unsupported(effective)
            || paragraph
                .properties
                .as_ref()
                .is_some_and(paragraph_properties_have_unsupported)
        {
            self.diagnose(
                &format!("{path}/pPr"),
                "unsupported paragraph properties were dropped during ODT export",
            )?;
        }
        if [Some(effective), paragraph.properties.as_ref()]
            .into_iter()
            .flatten()
            .any(|properties| properties.style_id.is_some())
        {
            self.diagnose(
                &format!("{path}/pPr/pStyle"),
                "paragraph style identity was materialized and dropped during ODT export",
            )?;
        }
        if [Some(effective), paragraph.properties.as_ref()]
            .into_iter()
            .flatten()
            .any(|properties| {
                properties
                    .line_rule
                    .as_deref()
                    .is_some_and(|rule| rule != "auto" && rule != "exact")
            })
        {
            self.diagnose(
                &format!("{path}/pPr/spacing"),
                "unsupported line-spacing rule was simplified during ODT export",
            )?;
        }
        if [Some(effective), paragraph.properties.as_ref()]
            .into_iter()
            .flatten()
            .any(|properties| properties.line_rule.is_some() && properties.line_spacing.is_none())
        {
            self.diagnose(
                &format!("{path}/pPr/spacing"),
                "line-spacing rule without line spacing was dropped during ODT export",
            )?;
        }
        if [Some(effective), paragraph.properties.as_ref()]
            .into_iter()
            .flatten()
            .any(|properties| properties.num_ilvl.is_some() && properties.num_id.is_none())
        {
            self.diagnose(
                &format!("{path}/pPr/numPr/ilvl"),
                "numbering level without a numbering id was dropped during ODT export",
            )?;
        }
        if [Some(effective), paragraph.properties.as_ref()]
            .into_iter()
            .flatten()
            .any(|properties| {
                properties.ind_first_line.is_some() && properties.ind_hanging.is_some()
            })
        {
            self.diagnose(
                &format!("{path}/pPr/ind/hanging"),
                "hanging indent was dropped because first-line indent takes precedence during ODT export",
            )?;
        }
        if [Some(effective), paragraph.properties.as_ref()]
            .into_iter()
            .flatten()
            .any(|properties| properties.jc == Some(ST_Jc::Distribute))
        {
            self.diagnose(
                &format!("{path}/pPr/jc"),
                "distributed paragraph alignment was simplified to justify during ODT export",
            )?;
        }
        for (index, _) in &paragraph.extra_xml {
            self.diagnose(
                &format!("{path}/raw[{index}]"),
                "unmodelled paragraph XML was dropped during ODT export",
            )?;
        }
        if !paragraph.content_controls.is_empty() {
            self.diagnose(
                &format!("{path}/content-controls"),
                "run content controls were dropped during ODT export",
            )?;
        }
        if !paragraph.revisions.is_empty() {
            self.diagnose(
                &format!("{path}/revisions"),
                "paragraph revisions were flattened during ODT export",
            )?;
        }
        if !paragraph.comment_ranges.is_empty() {
            self.diagnose(
                &format!("{path}/comments"),
                "comment markers were dropped during ODT export",
            )?;
        }
        if !paragraph.bookmark_markers.is_empty() {
            self.diagnose(
                &format!("{path}/bookmarks"),
                "bookmark markers were dropped during ODT export",
            )?;
        }
        if !paragraph.hyperlinks.is_empty() {
            self.diagnose(
                &format!("{path}/hyperlinks"),
                "hyperlink wrappers were flattened during ODT export",
            )?;
        }
        Ok(())
    }

    fn scan_run_losses(&mut self, run: &CT_R, effective: &CT_RPr, path: &str) -> Result<()> {
        if run_properties_have_unsupported(effective)
            || run
                .properties
                .as_ref()
                .is_some_and(run_properties_have_unsupported)
        {
            self.diagnose(
                &format!("{path}/rPr"),
                "unsupported run properties were dropped during ODT export",
            )?;
        }
        if [Some(effective), run.properties.as_ref()]
            .into_iter()
            .flatten()
            .any(|properties| properties.style_id.is_some())
        {
            self.diagnose(
                &format!("{path}/rPr/rStyle"),
                "run style identity was materialized and dropped during ODT export",
            )?;
        }
        if !run.extra_xml.is_empty() || !run.alt_drawings.is_empty() {
            self.diagnose(
                &format!("{path}/raw"),
                "unmodelled run XML was dropped during ODT export",
            )?;
        }
        for properties in [Some(effective), run.properties.as_ref()]
            .into_iter()
            .flatten()
        {
            if properties
                .vert_align
                .as_deref()
                .is_some_and(|value| !matches!(value, "superscript" | "subscript"))
            {
                self.diagnose(
                    &format!("{path}/rPr/vertAlign"),
                    "unsupported run vertical alignment was dropped during ODT export",
                )?;
            }
            if properties
                .color
                .as_deref()
                .is_some_and(|color| normalized_color(color).is_none())
            {
                self.diagnose(
                    &format!("{path}/rPr/color"),
                    "unsupported run color was dropped during ODT export",
                )?;
            }
            if let Some(shading) = properties.shading.as_ref() {
                if !matches!(shading.val.as_str(), "clear" | "nil") {
                    self.diagnose(
                        &format!("{path}/rPr/shading-pattern"),
                        "run shading pattern was simplified during ODT export",
                    )?;
                }
                if shading
                    .color
                    .as_deref()
                    .is_some_and(|color| color != "auto")
                {
                    self.diagnose(
                        &format!("{path}/rPr/shading-color"),
                        "run shading foreground color was dropped during ODT export",
                    )?;
                }
                let valid_fill = shading
                    .fill
                    .as_deref()
                    .is_some_and(|fill| normalized_color(fill).is_some());
                if shading
                    .fill
                    .as_deref()
                    .is_some_and(|fill| normalized_color(fill).is_none())
                {
                    self.diagnose(
                        &format!("{path}/rPr/shading-fill"),
                        "unsupported run shading fill was dropped during ODT export",
                    )?;
                }
                if valid_fill
                    && properties
                        .highlight
                        .as_ref()
                        .is_some_and(|highlight| highlight != &ST_HighlightColor::None)
                {
                    self.diagnose(
                        &format!("{path}/rPr/highlight"),
                        "run highlight was replaced by shading fill during ODT export",
                    )?;
                }
            }
        }
        for (index, content) in run.content.iter().enumerate() {
            let content_path = format!("{path}/content[{index}]");
            match content {
                RunContent::Text(text) | RunContent::DeletedText(text) => {
                    if !text.text.chars().all(valid_xml_character) {
                        return Err(odt_error(
                            Some("content.xml"),
                            0,
                            format!("invalid XML character at {content_path}"),
                        ));
                    }
                    if text.text.chars().any(|character| {
                        character.is_whitespace() && !matches!(character, ' ' | '\t' | '\r' | '\n')
                    }) {
                        return Err(odt_error(
                            Some("content.xml"),
                            0,
                            format!("ODT cannot preserve Unicode whitespace at {content_path}"),
                        ));
                    }
                    self.reserve_output(text.text.len().saturating_mul(18))?;
                    if matches!(content, RunContent::DeletedText(_)) {
                        self.diagnose(
                            &content_path,
                            "deleted text was flattened during ODT export",
                        )?;
                    }
                }
                RunContent::Break(BreakType::Page | BreakType::Column) => self.diagnose(
                    &content_path,
                    "unsupported break type was dropped during ODT export",
                )?,
                RunContent::Field(field) => {
                    let display = field
                        .projected_text()
                        .unwrap_or(field.cached_result.as_str());
                    validate_xml_value(display, &content_path)?;
                    validate_unicode_whitespace(display, &content_path)?;
                    self.reserve_output(display.len().saturating_mul(18))?;
                    self.diagnose(&content_path, "field was flattened during ODT export")?
                }
                RunContent::FootnoteRef { .. } => self.diagnose(
                    &content_path,
                    "footnote reference was dropped during ODT export",
                )?,
                RunContent::EndnoteRef { .. } => self.diagnose(
                    &content_path,
                    "endnote reference was dropped during ODT export",
                )?,
                RunContent::CommentReference { .. } => self.diagnose(
                    &content_path,
                    "comment reference was dropped during ODT export",
                )?,
                RunContent::Tab | RunContent::Break(BreakType::Line) | RunContent::Drawing(_) => {}
            }
        }
        Ok(())
    }

    fn scan_drawing(
        &mut self,
        drawing: &rdocx_oxml::drawing::CT_Drawing,
        path: &str,
    ) -> Result<()> {
        if drawing.anchor.is_some() {
            self.diagnose(path, "anchored drawing was dropped during ODT export")?;
        }
        let Some(inline) = &drawing.inline else {
            if drawing.anchor.is_none() {
                self.diagnose(path, "empty drawing was dropped during ODT export")?;
            }
            return Ok(());
        };
        if inline.description.is_some() {
            self.diagnose(
                path,
                "inline image description was dropped during ODT export",
            )?;
        }
        if inline.name.is_some() {
            self.diagnose(path, "inline image name was replaced during ODT export")?;
        }
        if inline.raw_xml.is_some() {
            self.diagnose(
                path,
                "retained inline drawing XML was dropped during ODT export",
            )?;
        }
        if inline.chart_rel_id.is_some() || inline.embed_id.is_empty() {
            self.diagnose(
                path,
                "non-image inline drawing was dropped during ODT export",
            )?;
            return Ok(());
        }
        if inline.extent_cx.0 <= 0 || inline.extent_cy.0 <= 0 {
            self.diagnose(
                path,
                "non-positive inline image was dropped during ODT export",
            )?;
            return Ok(());
        }
        const MAX_IMAGE_EMU: i64 = 12_700_000_000;
        if inline.extent_cx.0 > MAX_IMAGE_EMU || inline.extent_cy.0 > MAX_IMAGE_EMU {
            return Err(odt_error(
                Some("content.xml"),
                0,
                format!("ODT cannot preserve inline image dimensions at {path}"),
            ));
        }
        let Some(relationship) = self.image_relationship(&inline.embed_id) else {
            self.diagnose(
                path,
                "unresolved inline image was dropped during ODT export",
            )?;
            return Ok(());
        };
        if relationship.rel_type != oxml_opc::relationship::rel_types::IMAGE {
            self.diagnose(
                path,
                "inline drawing with a non-image relationship was dropped during ODT export",
            )?;
            return Ok(());
        }
        match relationship.target_mode.as_deref() {
            None | Some("Internal") => {}
            Some("External") => {
                self.diagnose(path, "external inline image was dropped during ODT export")?;
                return Ok(());
            }
            Some(_) => {
                self.diagnose(
                    path,
                    "invalid inline image target mode was dropped during ODT export",
                )?;
                return Ok(());
            }
        }
        let Some(bytes) = self.relationship_bytes(relationship) else {
            self.diagnose(
                path,
                "unresolved inline image was dropped during ODT export",
            )?;
            return Ok(());
        };
        let Some(info) = oxml_media::probe(bytes) else {
            self.diagnose(
                path,
                "unsupported inline image was dropped during ODT export",
            )?;
            return Ok(());
        };
        let format = info.format;
        if !matches!(
            format,
            oxml_media::ImageFormat::Png
                | oxml_media::ImageFormat::Jpeg
                | oxml_media::ImageFormat::Gif
                | oxml_media::ImageFormat::Bmp
                | oxml_media::ImageFormat::Webp
        ) {
            self.diagnose(
                path,
                "unsupported inline image was dropped during ODT export",
            )?;
            return Ok(());
        }
        let media_bytes = bytes.len();
        let part_limit = self.limits.archive.max_part_uncompressed_bytes as usize;
        if media_bytes > part_limit {
            return Err(odt_error(
                None,
                0,
                "ODT media output exceeds the part size limit",
            ));
        }
        if self.media.len().saturating_add(4) > self.limits.archive.max_entries {
            return Err(odt_error(None, 0, "ODT output exceeds the entry limit"));
        }
        self.reserve_output(512)?;
        let retained_media_bytes = self
            .retained_media_bytes
            .checked_add(media_bytes)
            .filter(|media| {
                media
                    .checked_add(self.retained_output_estimate)
                    .is_some_and(|total| {
                        total as u64 <= self.limits.archive.max_total_uncompressed_bytes
                    })
            })
            .ok_or_else(|| odt_error(None, 0, "ODT output exceeds the size limit"))?;
        let index = self.media.len();
        let bytes = self
            .image_relationship(&inline.embed_id)
            .and_then(|relationship| self.relationship_bytes(relationship))
            .expect("image relationship remained immutable and internal")
            .to_vec();
        self.media.push(OdtMedia {
            path: format!("Pictures/image{}.{}", index + 1, format.extension()),
            media_type: format.content_type(),
            bytes,
        });
        self.retained_media_bytes = retained_media_bytes;
        self.media_at.insert(path.to_owned(), index);
        Ok(())
    }

    fn image_relationship(
        &self,
        relationship_id: &str,
    ) -> Option<&oxml_opc::relationship::Relationship> {
        let relationships = self
            .document
            .package
            .get_part_rels(&self.document.doc_part_name)?;
        relationships
            .items
            .iter()
            .find(|relationship| relationship.id == relationship_id)
    }

    fn relationship_bytes(
        &self,
        relationship: &oxml_opc::relationship::Relationship,
    ) -> Option<&[u8]> {
        let target = oxml_opc::OpcPackage::resolve_rel_target(
            &self.document.doc_part_name,
            &relationship.target,
        );
        self.document.package.get_part(&target)
    }

    fn effective_paragraph_properties(&self, paragraph: &CT_P) -> CT_PPr {
        let direct = paragraph.properties.as_ref();
        let mut effective = self.document.resolve_paragraph_properties(
            direct.and_then(|properties| properties.style_id.as_deref()),
        );
        if let Some(direct) = direct {
            effective.merge_from(direct);
        }
        effective
    }

    fn paragraph_numbering(&self, paragraph: &CT_P) -> Option<(u32, u32)> {
        let numbering =
            paragraph_numbering_properties(&self.effective_paragraph_properties(paragraph))?;
        (!self.list_level_has_producer_format(numbering.0, numbering.1)).then_some(numbering)
    }

    fn list_level_has_producer_format(&self, num_id: u32, level: u32) -> bool {
        self.document
            .numbering
            .as_ref()
            .and_then(|numbering| numbering.get_abstract_num_for(num_id))
            .and_then(|abstract_num| {
                abstract_num
                    .levels
                    .iter()
                    .find(|definition| definition.ilvl == level)
            })
            .and_then(|definition| definition.num_fmt.as_ref())
            .is_some_and(|format| matches!(format, ST_NumberFormat::Other(_)))
    }

    fn effective_run_properties(&self, paragraph: &CT_P, run: &CT_R) -> CT_RPr {
        let paragraph_style = paragraph
            .properties
            .as_ref()
            .and_then(|properties| properties.style_id.as_deref());
        let direct = run.properties.as_ref();
        let mut effective = self.document.resolve_run_properties(
            paragraph_style,
            direct.and_then(|properties| properties.style_id.as_deref()),
        );
        if let Some(direct) = direct {
            effective.merge_from(direct);
        }
        effective
    }

    fn ensure_paragraph_style(&mut self, xml: String) -> Result<Option<String>> {
        if xml.is_empty() {
            return Ok(None);
        }
        if let Some(name) = self.paragraph_styles.get(&xml) {
            return Ok(Some(name.clone()));
        }
        self.reserve_output(xml.len().saturating_add(160))?;
        let name = format!("P{}", self.paragraph_styles.len() + 1);
        self.paragraph_styles.insert(xml, name.clone());
        Ok(Some(name))
    }

    fn ensure_text_style(&mut self, xml: String) -> Result<Option<String>> {
        if xml.is_empty() {
            return Ok(None);
        }
        if let Some(name) = self.text_styles.get(&xml) {
            return Ok(Some(name.clone()));
        }
        self.reserve_output(xml.len().saturating_add(160))?;
        let name = format!("T{}", self.text_styles.len() + 1);
        self.text_styles.insert(xml, name.clone());
        Ok(Some(name))
    }

    fn ensure_list_style(&mut self, num_id: u32, path: &str) -> Result<String> {
        if let Some(name) = self.list_styles.get(&num_id) {
            return Ok(name.clone());
        }
        if self
            .document
            .numbering
            .as_ref()
            .and_then(|numbering| numbering.get_abstract_num_for(num_id))
            .is_none()
        {
            self.diagnose(
                path,
                "unknown list definition was exported as decimal ODT list",
            )?;
        } else {
            self.scan_numbering_container_losses(num_id)?;
        }
        self.reserve_output(1_024)?;
        let name = format!("L{}", self.list_styles.len() + 1);
        self.list_styles.insert(num_id, name.clone());
        Ok(name)
    }

    fn scan_numbering_container_losses(&mut self, num_id: u32) -> Result<()> {
        let Some(numbering) = self.document.numbering.as_ref() else {
            return Ok(());
        };
        let instance = numbering
            .nums
            .iter()
            .find(|instance| instance.num_id == num_id);
        let abstract_num = numbering.get_abstract_num_for(num_id);
        let losses = (
            !numbering.root_attributes.is_empty(),
            !numbering.extra_xml.is_empty(),
            instance.is_some_and(|value| !value.extra_attributes.is_empty()),
            instance.is_some_and(|value| !value.extra_xml.is_empty()),
            abstract_num.is_some_and(|value| value.multi_level_type.is_some()),
            abstract_num.is_some_and(|value| !value.extra_attributes.is_empty()),
            abstract_num.is_some_and(|value| !value.extra_xml.is_empty()),
        );
        let path = format!("numbering[{num_id}]");
        for (present, suffix, message) in [
            (
                losses.0,
                "root-attributes",
                "retained numbering root attributes were dropped during ODT export",
            ),
            (
                losses.1,
                "root-xml",
                "retained numbering root XML was dropped during ODT export",
            ),
            (
                losses.2,
                "instance-attributes",
                "retained numbering instance attributes were dropped during ODT export",
            ),
            (
                losses.3,
                "instance-overrides",
                "numbering instance overrides or retained XML were dropped during ODT export",
            ),
            (
                losses.4,
                "abstract-type",
                "abstract numbering type metadata was dropped during ODT export",
            ),
            (
                losses.5,
                "abstract-attributes",
                "retained abstract numbering attributes were dropped during ODT export",
            ),
            (
                losses.6,
                "abstract-xml",
                "retained abstract numbering XML was dropped during ODT export",
            ),
        ] {
            if present {
                self.diagnose(&format!("{path}/{suffix}"), message)?;
            }
        }
        Ok(())
    }

    fn scan_flattened_producer_list_losses(&mut self, num_id: u32, level: u32) -> Result<()> {
        self.scan_numbering_container_losses(num_id)?;
        let definition = self
            .document
            .numbering
            .as_ref()
            .and_then(|numbering| numbering.get_abstract_num_for(num_id))
            .and_then(|abstract_num| {
                abstract_num
                    .levels
                    .iter()
                    .find(|definition| definition.ilvl == level)
            });
        let losses = definition.map_or(
            (false, false, false, false, false, false, false),
            |definition| {
                (
                    definition.start.is_some_and(|start| start != 1),
                    definition.suffix.is_some(),
                    definition.lvl_text.is_some(),
                    definition.lvl_jc.is_some(),
                    definition.ppr.is_some() || definition.ppr_raw.is_some(),
                    definition.rpr.is_some() || definition.rpr_raw.is_some(),
                    !definition.extra_xml.is_empty() || !definition.extra_attributes.is_empty(),
                )
            },
        );
        let path = format!("numbering[{num_id}]/level[{level}]");
        for (present, message) in [
            (
                losses.0,
                "custom list start value was dropped during ODT export",
            ),
            (losses.1, "list marker suffix was dropped during ODT export"),
            (
                losses.2,
                "list marker text or bullet glyph was dropped during ODT export",
            ),
            (
                losses.3,
                "list marker justification was dropped during ODT export",
            ),
            (
                losses.4,
                "list level paragraph formatting was dropped during ODT export",
            ),
            (
                losses.5,
                "list marker run formatting was dropped during ODT export",
            ),
            (
                losses.6,
                "retained list level XML or attributes were dropped during ODT export",
            ),
        ] {
            if present {
                self.diagnose(&path, message)?;
            }
        }
        Ok(())
    }

    fn content_xml(&mut self) -> Result<String> {
        let mut output = String::new();
        output.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
        write!(
            output,
            "<office:document-content xmlns:office=\"{OFFICE_NS}\" xmlns:text=\"{TEXT_NS}\" xmlns:style=\"{STYLE_NS}\" xmlns:fo=\"{FO_NS}\" xmlns:table=\"{TABLE_NS}\" xmlns:draw=\"{DRAW_NS}\" xmlns:xlink=\"{XLINK_NS}\" xmlns:svg=\"{SVG_NS}\" office:version=\"1.3\"><office:automatic-styles>"
        )
        .unwrap();
        for (xml, name) in &self.paragraph_styles {
            write!(
                output,
                "<style:style style:name=\"{name}\" style:family=\"paragraph\"><style:paragraph-properties{xml}/></style:style>"
            )
            .unwrap();
        }
        for (xml, name) in &self.text_styles {
            write!(
                output,
                "<style:style style:name=\"{name}\" style:family=\"text\"><style:text-properties{xml}/></style:style>"
            )
            .unwrap();
        }
        let list_styles = self.list_styles.clone();
        for (num_id, name) in list_styles {
            write!(output, "<text:list-style style:name=\"{name}\">").unwrap();
            for level in 0..9_u32 {
                if self.list_level_is_bullet(
                    num_id,
                    level,
                    &format!("numbering[{num_id}]/level[{level}]"),
                )? {
                    write!(
                        output,
                        "<text:list-level-style-bullet text:level=\"{}\" text:bullet-char=\"•\"/>",
                        level + 1
                    )
                    .unwrap();
                } else {
                    write!(
                        output,
                        "<text:list-level-style-number text:level=\"{}\" style:num-format=\"1\"/>",
                        level + 1
                    )
                    .unwrap();
                }
            }
            output.push_str("</text:list-style>");
        }
        output.push_str("</office:automatic-styles><office:body><office:text>");
        let mut index = 0_usize;
        let mut written_lists = BTreeSet::new();
        while index < self.document.document.body.content.len() {
            let content = &self.document.document.body.content[index];
            if let BodyContent::Paragraph(paragraph) = content
                && let Some((num_id, _)) = self.paragraph_numbering(paragraph)
            {
                let start = index;
                while index < self.document.document.body.content.len()
                    && matches!(
                        &self.document.document.body.content[index],
                        BodyContent::Paragraph(candidate)
                            if self.paragraph_numbering(candidate).is_some_and(|(candidate_id, _)| candidate_id == num_id)
                    )
                {
                    index += 1;
                }
                let continue_numbering = !written_lists.insert(num_id);
                self.write_list(&mut output, start, index, num_id, continue_numbering)?;
                continue;
            }
            self.write_body(&mut output, content, &format!("body[{index}]"))?;
            index += 1;
        }
        output.push_str("</office:text></office:body></office:document-content>");
        Ok(output)
    }

    fn write_body(&self, output: &mut String, content: &BodyContent, path: &str) -> Result<()> {
        match content {
            BodyContent::Paragraph(paragraph) => self.write_paragraph(output, paragraph, path),
            BodyContent::Table(table) => self.write_table(output, table, path),
            BodyContent::ContentControl(_) | BodyContent::RawXml(_) => Ok(()),
        }
    }

    fn write_list(
        &self,
        output: &mut String,
        start: usize,
        end: usize,
        num_id: u32,
        continue_numbering: bool,
    ) -> Result<()> {
        let style = self.list_styles.get(&num_id).expect("scanned list style");
        let mut items = Vec::with_capacity(end - start);
        for index in start..end {
            let BodyContent::Paragraph(paragraph) = &self.document.document.body.content[index]
            else {
                unreachable!();
            };
            let (_, level) = self
                .paragraph_numbering(paragraph)
                .expect("scanned list paragraph");
            items.push((paragraph, format!("body[{index}]"), level as usize));
        }
        write_list_level(self, output, &items, 0, style, continue_numbering)?;
        Ok(())
    }

    fn write_table(&self, output: &mut String, table: &CT_Tbl, path: &str) -> Result<()> {
        let columns = table_column_count(table)?;
        write!(output, "<table:table table:name=\"Table\">").unwrap();
        write!(
            output,
            "<table:table-column table:number-columns-repeated=\"{columns}\"/>"
        )
        .unwrap();
        for (row_index, row) in table.rows.iter().enumerate() {
            output.push_str("<table:table-row>");
            for (cell_index, cell) in row.cells.iter().enumerate() {
                let cell_path = format!("{path}/row[{row_index}]/cell[{cell_index}]");
                let properties = cell.properties.as_ref();
                let span = properties.and_then(|value| value.grid_span).unwrap_or(1);
                if matches!(
                    properties.and_then(|value| value.v_merge.as_ref()),
                    Some(VMerge::Continue)
                ) {
                    output.push_str("<table:covered-table-cell/>");
                    for _ in 1..span {
                        output.push_str("<table:covered-table-cell/>");
                    }
                    continue;
                }
                output.push_str("<table:table-cell");
                if span > 1 {
                    write!(output, " table:number-columns-spanned=\"{span}\"").unwrap();
                }
                if matches!(
                    properties.and_then(|value| value.v_merge.as_ref()),
                    Some(VMerge::Restart)
                ) {
                    let rowspan = table_rowspan(table, row_index, cell_index)?;
                    write!(output, " table:number-rows-spanned=\"{rowspan}\"").unwrap();
                }
                output.push('>');
                let mut wrote_paragraph = false;
                let mut content_index = 0_usize;
                let mut written_lists = BTreeSet::new();
                while content_index < cell.content.len() {
                    match &cell.content[content_index] {
                        CellContent::Paragraph(paragraph) => {
                            if let Some((num_id, _)) = self.paragraph_numbering(paragraph) {
                                let start = content_index;
                                while content_index < cell.content.len()
                                    && matches!(
                                        &cell.content[content_index],
                                        CellContent::Paragraph(candidate)
                                            if self.paragraph_numbering(candidate).is_some_and(|(candidate_id, _)| candidate_id == num_id)
                                    )
                                {
                                    content_index += 1;
                                }
                                let style =
                                    self.list_styles.get(&num_id).expect("scanned list style");
                                let mut items = Vec::with_capacity(content_index - start);
                                for item_index in start..content_index {
                                    let CellContent::Paragraph(item) = &cell.content[item_index]
                                    else {
                                        unreachable!();
                                    };
                                    let (_, level) = self
                                        .paragraph_numbering(item)
                                        .expect("scanned list paragraph");
                                    items.push((
                                        item,
                                        format!("{cell_path}/content[{item_index}]"),
                                        level as usize,
                                    ));
                                }
                                let continue_numbering = !written_lists.insert(num_id);
                                write_list_level(
                                    self,
                                    output,
                                    &items,
                                    0,
                                    style,
                                    continue_numbering,
                                )?;
                            } else {
                                self.write_paragraph(
                                    output,
                                    paragraph,
                                    &format!("{cell_path}/content[{content_index}]"),
                                )?;
                                content_index += 1;
                            }
                            wrote_paragraph = true;
                        }
                        CellContent::Table(_) | CellContent::ContentControl(_) => {
                            content_index += 1;
                        }
                    }
                }
                if !wrote_paragraph {
                    output.push_str("<text:p/>");
                }
                output.push_str("</table:table-cell>");
                for _ in 1..span {
                    output.push_str("<table:covered-table-cell/>");
                }
            }
            output.push_str("</table:table-row>");
        }
        output.push_str("</table:table>");
        Ok(())
    }

    fn write_paragraph(&self, output: &mut String, paragraph: &CT_P, path: &str) -> Result<()> {
        let properties = self.effective_paragraph_properties(paragraph);
        let style_xml = paragraph_style_xml(&properties);
        let style = self.paragraph_styles.get(&style_xml);
        let heading = properties
            .outline_lvl
            .or_else(|| heading_level_from_style(properties.style_id.as_deref()));
        let tag = if heading.is_some() {
            "text:h"
        } else {
            "text:p"
        };
        write!(output, "<{tag}").unwrap();
        if let Some(style) = style {
            write!(output, " text:style-name=\"{style}\"").unwrap();
        }
        if let Some(level) = heading {
            write!(output, " text:outline-level=\"{}\"", level.clamp(0, 8) + 1).unwrap();
        }
        output.push('>');
        self.write_runs(output, paragraph, path)?;
        write!(output, "</{tag}>").unwrap();
        Ok(())
    }

    fn write_runs(&self, output: &mut String, paragraph: &CT_P, path: &str) -> Result<()> {
        for (run_index, run) in paragraph.runs.iter().enumerate() {
            let run_path = format!("{path}/run[{run_index}]");
            let properties = self.effective_run_properties(paragraph, run);
            let style_xml = text_style_xml(&properties);
            let style = self.text_styles.get(&style_xml);
            if let Some(style) = style {
                write!(output, "<text:span text:style-name=\"{style}\">").unwrap();
            }
            for (content_index, content) in run.content.iter().enumerate() {
                let content_path = format!("{run_path}/content[{content_index}]");
                match content {
                    RunContent::Text(text) | RunContent::DeletedText(text) => {
                        write_odf_text(output, &text.text)
                    }
                    RunContent::Tab => output.push_str("<text:tab/>"),
                    RunContent::Break(BreakType::Line) => output.push_str("<text:line-break/>"),
                    RunContent::Break(_) => {}
                    RunContent::Drawing(drawing) => {
                        if let Some(inline) = &drawing.inline
                            && let Some(media_index) = self.media_at.get(&content_path)
                        {
                            let media = &self.media[*media_index];
                            write!(
                                output,
                                "<draw:frame draw:name=\"Image{}\" text:anchor-type=\"as-char\" svg:width=\"{}pt\" svg:height=\"{}pt\"><draw:image xlink:href=\"{}\" xlink:type=\"simple\" xlink:show=\"embed\" xlink:actuate=\"onLoad\"/></draw:frame>",
                                media_index + 1,
                                emu_points(inline.extent_cx.0),
                                emu_points(inline.extent_cy.0),
                                media.path
                            )
                            .unwrap();
                        }
                    }
                    RunContent::Field(field) => {
                        write_odf_text(
                            output,
                            field
                                .projected_text()
                                .unwrap_or(field.cached_result.as_str()),
                        );
                    }
                    RunContent::FootnoteRef { .. }
                    | RunContent::EndnoteRef { .. }
                    | RunContent::CommentReference { .. } => {}
                }
            }
            if style.is_some() {
                output.push_str("</text:span>");
            }
        }
        Ok(())
    }

    fn list_level_is_bullet(&mut self, num_id: u32, level: u32, path: &str) -> Result<bool> {
        let abstract_num = self
            .document
            .numbering
            .as_ref()
            .and_then(|numbering| numbering.get_abstract_num_for(num_id));
        let level_definition = abstract_num
            .and_then(|abstract_num| abstract_num.levels.iter().find(|item| item.ilvl == level));
        let (
            is_bullet,
            is_decimal,
            has_other_format,
            start,
            suffix,
            level_text,
            justification,
            paragraph_formatting,
            marker_formatting,
            retained_xml,
        ) = level_definition.map_or(
            (
                false, false, false, 1, false, false, false, false, false, false,
            ),
            |definition| {
                let format = definition.num_fmt.as_ref();
                (
                    matches!(format, Some(ST_NumberFormat::Bullet)),
                    matches!(format, Some(ST_NumberFormat::Decimal)),
                    format.is_some()
                        && !matches!(
                            format,
                            Some(ST_NumberFormat::Bullet | ST_NumberFormat::Decimal)
                        ),
                    definition.start.unwrap_or(1),
                    definition.suffix.is_some(),
                    definition.lvl_text.is_some(),
                    definition.lvl_jc.is_some(),
                    definition.ppr.is_some() || definition.ppr_raw.is_some(),
                    definition.rpr.is_some() || definition.rpr_raw.is_some(),
                    !definition.extra_xml.is_empty() || !definition.extra_attributes.is_empty(),
                )
            },
        );
        let level_is_used = self.used_list_levels.contains(&(num_id, level));
        if level_is_used && abstract_num.is_some() && level_definition.is_none() {
            self.diagnose(
                path,
                "undefined numbering level was exported as decimal ODT list",
            )?;
        }
        if level_is_used && start != 1 {
            self.diagnose(path, "custom list start value was reset during ODT export")?;
        }
        for (present, message) in [
            (suffix, "list marker suffix was dropped during ODT export"),
            (
                level_text,
                "list marker text or bullet glyph was simplified during ODT export",
            ),
            (
                justification,
                "list marker justification was dropped during ODT export",
            ),
            (
                paragraph_formatting,
                "list level paragraph formatting was dropped during ODT export",
            ),
            (
                marker_formatting,
                "list marker run formatting was dropped during ODT export",
            ),
            (
                retained_xml,
                "retained list level XML or attributes were dropped during ODT export",
            ),
        ] {
            if level_is_used && present {
                self.diagnose(path, message)?;
            }
        }
        if is_bullet {
            return Ok(true);
        }
        if is_decimal {
            return Ok(false);
        }
        if level_is_used && has_other_format {
            self.diagnose(
                path,
                "non-decimal numbering format was exported as decimal ODT list",
            )?;
        }
        Ok(false)
    }

    fn manifest_xml(&self) -> String {
        let mut output = format!(
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?><manifest:manifest xmlns:manifest=\"{MANIFEST_NS}\" manifest:version=\"1.3\"><manifest:file-entry manifest:full-path=\"/\" manifest:media-type=\"{}\"/><manifest:file-entry manifest:full-path=\"content.xml\" manifest:media-type=\"text/xml\"/>",
            std::str::from_utf8(ODT_MIMETYPE).expect("ASCII MIME type")
        );
        for media in &self.media {
            write!(
                output,
                "<manifest:file-entry manifest:full-path=\"{}\" manifest:media-type=\"{}\"/>",
                media.path, media.media_type
            )
            .unwrap();
        }
        output.push_str("</manifest:manifest>");
        output
    }

    fn diagnose(&mut self, path: &str, message: &str) -> Result<()> {
        let key = (path.to_owned(), message.to_owned());
        if !self.diagnostic_keys.insert(key) {
            return Ok(());
        }
        if self.diagnostics.len() >= self.limits.diagnostics {
            return Err(odt_error(
                Some("content.xml"),
                0,
                "ODT exceeds the diagnostic limit",
            ));
        }
        self.diagnostics.push(OdtDiagnostic {
            path: path.to_owned(),
            message: message.to_owned(),
        });
        Ok(())
    }

    fn reserve_output(&mut self, amount: usize) -> Result<()> {
        self.retained_output_estimate = self
            .retained_output_estimate
            .checked_add(amount)
            .filter(|value| *value as u64 <= self.limits.archive.max_part_uncompressed_bytes)
            .ok_or_else(|| odt_error(None, 0, "ODT XML output exceeds the size limit"))?;
        Ok(())
    }
}

fn valid_xml_character(value: char) -> bool {
    matches!(value, '\u{9}' | '\u{A}' | '\u{D}')
        || ('\u{20}'..='\u{D7FF}').contains(&value)
        || ('\u{E000}'..='\u{FFFD}').contains(&value)
        || ('\u{10000}'..='\u{10FFFF}').contains(&value)
}

fn validate_xml_value(value: &str, path: &str) -> Result<()> {
    if value.chars().all(valid_xml_character) {
        Ok(())
    } else {
        Err(odt_error(
            Some("content.xml"),
            0,
            format!("invalid XML character at {path}"),
        ))
    }
}

fn validate_unicode_whitespace(value: &str, path: &str) -> Result<()> {
    if value.chars().any(|character| {
        character.is_whitespace() && !matches!(character, ' ' | '\t' | '\r' | '\n')
    }) {
        return Err(odt_error(
            Some("content.xml"),
            0,
            format!("ODT cannot preserve Unicode whitespace at {path}"),
        ));
    }
    Ok(())
}

fn write_odt_entry(
    archive: &mut ZipWriter<&mut Cursor<Vec<u8>>>,
    name: &str,
    bytes: &[u8],
    method: CompressionMethod,
) -> Result<()> {
    archive
        .start_file(
            name,
            SimpleFileOptions::DEFAULT
                .compression_method(method)
                .unix_permissions(0o644),
        )
        .map_err(|error| odt_error(Some(name), 0, format!("cannot create ODT entry: {error}")))?;
    archive
        .write_all(bytes)
        .map_err(|error| odt_error(Some(name), 0, format!("cannot write ODT entry: {error}")))
}

fn write_list_level(
    writer: &OdtWriter<'_>,
    output: &mut String,
    items: &[(&CT_P, String, usize)],
    level: usize,
    style: &str,
    continue_numbering: bool,
) -> Result<()> {
    write!(output, "<text:list text:style-name=\"{style}\"").unwrap();
    if continue_numbering {
        output.push_str(" text:continue-numbering=\"true\"");
    }
    output.push('>');
    let mut index = 0_usize;
    while index < items.len() {
        let (paragraph, path, item_level) = &items[index];
        if *item_level < level {
            break;
        }
        if *item_level > level {
            let nested_start = index;
            while index < items.len() && items[index].2 > level {
                index += 1;
            }
            output.push_str("<text:list-item>");
            write_list_level(
                writer,
                output,
                &items[nested_start..index],
                level + 1,
                style,
                false,
            )?;
            output.push_str("</text:list-item>");
            continue;
        }
        output.push_str("<text:list-item>");
        writer.write_paragraph(output, paragraph, path)?;
        index += 1;
        if index < items.len() && items[index].2 > level {
            let nested_start = index;
            while index < items.len() && items[index].2 > level {
                index += 1;
            }
            write_list_level(
                writer,
                output,
                &items[nested_start..index],
                level + 1,
                style,
                false,
            )?;
        }
        output.push_str("</text:list-item>");
    }
    output.push_str("</text:list>");
    Ok(())
}

fn paragraph_numbering_properties(properties: &CT_PPr) -> Option<(u32, u32)> {
    let num_id = properties.num_id?;
    (num_id != 0).then_some((num_id, properties.num_ilvl.unwrap_or(0)))
}

fn validate_paragraph_projection(properties: &CT_PPr, path: &str) -> Result<()> {
    const MAX_LENGTH_TWIPS: i32 = 20_000_000;
    for (name, value) in [
        ("space-before", properties.space_before),
        ("space-after", properties.space_after),
        ("left-indent", properties.ind_left),
        ("right-indent", properties.ind_right),
    ] {
        if value.is_some_and(|value| !(0..=MAX_LENGTH_TWIPS).contains(&value.0)) {
            return Err(odt_error(
                Some("content.xml"),
                0,
                format!("ODT cannot preserve {name} at {path}"),
            ));
        }
    }
    if properties
        .ind_first_line
        .is_some_and(|value| value.0.unsigned_abs() > MAX_LENGTH_TWIPS as u32)
    {
        return Err(odt_error(
            Some("content.xml"),
            0,
            format!("ODT cannot preserve first-line indent at {path}"),
        ));
    }
    if properties
        .ind_hanging
        .is_some_and(|value| !(0..=MAX_LENGTH_TWIPS).contains(&value.0))
    {
        return Err(odt_error(
            Some("content.xml"),
            0,
            format!("ODT cannot preserve hanging indent at {path}"),
        ));
    }
    if let Some(line) = properties.line_spacing {
        let valid = match properties.line_rule.as_deref() {
            Some("exact" | "atLeast") => (1..=MAX_LENGTH_TWIPS).contains(&line.0),
            _ => (1..=24_000).contains(&line.0),
        };
        if !valid {
            return Err(odt_error(
                Some("content.xml"),
                0,
                format!("ODT cannot preserve line spacing at {path}"),
            ));
        }
    }
    if properties.outline_lvl.is_some_and(|level| level > 8) {
        return Err(odt_error(
            Some("content.xml"),
            0,
            format!("ODT cannot preserve outline level at {path}"),
        ));
    }
    if properties.outline_lvl.is_none()
        && let Some(suffix) = properties
            .style_id
            .as_deref()
            .and_then(|style| style.strip_prefix("Heading"))
        && !suffix.is_empty()
        && suffix.bytes().all(|byte| byte.is_ascii_digit())
        && !suffix
            .parse::<u32>()
            .is_ok_and(|level| (1..=9).contains(&level))
    {
        return Err(odt_error(
            Some("content.xml"),
            0,
            format!("ODT cannot preserve outline level at {path}"),
        ));
    }
    Ok(())
}

fn validate_run_projection(properties: &CT_RPr, path: &str) -> Result<()> {
    const MAX_FONT_HALF_POINTS: u32 = 2_000_000;
    if properties
        .sz
        .is_some_and(|size| !(1..=MAX_FONT_HALF_POINTS).contains(&size.0))
    {
        return Err(odt_error(
            Some("content.xml"),
            0,
            format!("ODT cannot preserve font size at {path}"),
        ));
    }
    Ok(())
}

fn validate_font_family_projection(font: &str, path: &str) -> Result<()> {
    let normalized = font
        .chars()
        .map(|character| match character {
            '\t' | '\r' | '\n' => ' ',
            other => other,
        })
        .collect::<String>();
    if normalized.trim().trim_matches(['\'', '"']) != font {
        return Err(odt_error(
            Some("content.xml"),
            0,
            format!("ODT cannot preserve font family at {path}"),
        ));
    }
    Ok(())
}

fn final_section_has_unsupported_properties(section: &rdocx_oxml::document::CT_SectPr) -> bool {
    let mut projected = section.clone();
    projected.header_refs.clear();
    projected.footer_refs.clear();
    projected != rdocx_oxml::document::CT_SectPr::default_letter()
}

fn paragraph_style_xml(properties: &CT_PPr) -> String {
    let mut output = String::new();
    if let Some(alignment) = properties.jc {
        let value = match alignment {
            ST_Jc::Start | ST_Jc::Left => "left",
            ST_Jc::End | ST_Jc::Right => "right",
            ST_Jc::Center => "center",
            ST_Jc::Both | ST_Jc::Distribute => "justify",
        };
        write!(output, " fo:text-align=\"{value}\"").unwrap();
    }
    for (name, value) in [
        ("margin-top", properties.space_before),
        ("margin-bottom", properties.space_after),
        ("margin-left", properties.ind_left),
        ("margin-right", properties.ind_right),
    ] {
        if let Some(value) = value {
            write!(output, " fo:{name}=\"{}pt\"", twips_points(value.0)).unwrap();
        }
    }
    if let Some(value) = properties.ind_first_line {
        write!(output, " fo:text-indent=\"{}pt\"", twips_points(value.0)).unwrap();
    } else if let Some(value) = properties.ind_hanging {
        write!(
            output,
            " fo:text-indent=\"{}pt\"",
            twips_points(value.0.saturating_neg())
        )
        .unwrap();
    }
    if let Some(value) = properties.line_spacing {
        match properties.line_rule.as_deref() {
            Some("exact" | "atLeast") => {
                write!(output, " fo:line-height=\"{}pt\"", twips_points(value.0)).unwrap();
            }
            _ => {
                write!(
                    output,
                    " fo:line-height=\"{}%\"",
                    automatic_line_height_percent(value.0)
                )
                .unwrap();
            }
        }
    }
    output
}

fn text_style_xml(properties: &CT_RPr) -> String {
    let mut output = String::new();
    if let Some(font) = selected_run_font(properties) {
        write!(output, " fo:font-family=\"{}\"", escape_xml(font)).unwrap();
    }
    if let Some(size) = properties.sz {
        write!(output, " fo:font-size=\"{}pt\"", decimal(size.to_pt())).unwrap();
    }
    if let Some(value) = properties.bold {
        write!(
            output,
            " fo:font-weight=\"{}\"",
            if value { "bold" } else { "normal" }
        )
        .unwrap();
    }
    if let Some(value) = properties.italic {
        write!(
            output,
            " fo:font-style=\"{}\"",
            if value { "italic" } else { "normal" }
        )
        .unwrap();
    }
    if let Some(value) = properties.underline {
        write!(
            output,
            " style:text-underline-style=\"{}\"",
            if value == ST_Underline::None {
                "none"
            } else {
                "solid"
            }
        )
        .unwrap();
    }
    if let Some(value) = properties.strike {
        write!(
            output,
            " style:text-line-through-style=\"{}\"",
            if value { "solid" } else { "none" }
        )
        .unwrap();
    }
    if let Some(color) = properties.color.as_deref().and_then(normalized_color) {
        write!(output, " fo:color=\"#{color}\"").unwrap();
    }
    let background = properties
        .shading
        .as_ref()
        .and_then(|shading| shading.fill.as_deref())
        .and_then(normalized_color)
        .or_else(|| properties.highlight.and_then(highlight_color));
    if let Some(background) = background {
        write!(output, " fo:background-color=\"#{background}\"").unwrap();
    }
    match properties.vert_align.as_deref() {
        Some("superscript") => output.push_str(" style:text-position=\"super\""),
        Some("subscript") => output.push_str(" style:text-position=\"sub\""),
        _ => {}
    }
    output
}

fn selected_run_font(properties: &CT_RPr) -> Option<&str> {
    properties
        .font_ascii
        .as_deref()
        .or(properties.font_hansi.as_deref())
        .or(properties.font_east_asia.as_deref())
        .or(properties.font_cs.as_deref())
}

fn normalized_color(value: &str) -> Option<String> {
    let value = value.trim_start_matches('#');
    (value.len() == 6 && value.bytes().all(|byte| byte.is_ascii_hexdigit()))
        .then(|| value.to_ascii_uppercase())
}

fn highlight_color(value: ST_HighlightColor) -> Option<String> {
    let value = match value {
        ST_HighlightColor::Black => "000000",
        ST_HighlightColor::Blue => "0000FF",
        ST_HighlightColor::Cyan => "00FFFF",
        ST_HighlightColor::DarkBlue => "000080",
        ST_HighlightColor::DarkCyan => "008080",
        ST_HighlightColor::DarkGray => "808080",
        ST_HighlightColor::DarkGreen => "008000",
        ST_HighlightColor::DarkMagenta => "800080",
        ST_HighlightColor::DarkRed => "800000",
        ST_HighlightColor::DarkYellow => "808000",
        ST_HighlightColor::Green => "00FF00",
        ST_HighlightColor::LightGray => "C0C0C0",
        ST_HighlightColor::Magenta => "FF00FF",
        ST_HighlightColor::None => return None,
        ST_HighlightColor::Red => "FF0000",
        ST_HighlightColor::White => "FFFFFF",
        ST_HighlightColor::Yellow => "FFFF00",
    };
    Some(value.to_owned())
}

fn heading_level_from_style(style: Option<&str>) -> Option<u32> {
    let level = style?.strip_prefix("Heading")?.parse::<u32>().ok()?;
    (1..=9).contains(&level).then_some(level - 1)
}

fn projected_text_run_pieces(
    text: &str,
    style: &str,
    trailing_text_style: &mut Option<String>,
) -> usize {
    let mut pieces = 0;
    let mut characters = text.chars().peekable();
    while let Some(character) = characters.next() {
        match character {
            '\t' | '\n' => {
                pieces += 1;
                *trailing_text_style = None;
            }
            '\r' => {
                if characters.peek() == Some(&'\n') {
                    characters.next();
                }
                pieces += 1;
                *trailing_text_style = None;
            }
            _ if trailing_text_style.as_deref() != Some(style) => {
                pieces += 1;
                *trailing_text_style = Some(style.to_owned());
            }
            _ => {}
        }
    }
    pieces
}

fn write_odf_text(output: &mut String, text: &str) {
    let mut ordinary = String::new();
    let mut characters = text.chars().peekable();
    while let Some(character) = characters.next() {
        match character {
            ' ' => {
                if !ordinary.is_empty() {
                    output.push_str(&escape_xml(&ordinary));
                    ordinary.clear();
                }
                output.push_str("<text:s/>");
            }
            '\t' => {
                if !ordinary.is_empty() {
                    output.push_str(&escape_xml(&ordinary));
                    ordinary.clear();
                }
                output.push_str("<text:tab/>");
            }
            '\r' => {
                if characters.peek() == Some(&'\n') {
                    characters.next();
                }
                if !ordinary.is_empty() {
                    output.push_str(&escape_xml(&ordinary));
                    ordinary.clear();
                }
                output.push_str("<text:line-break/>");
            }
            '\n' => {
                if !ordinary.is_empty() {
                    output.push_str(&escape_xml(&ordinary));
                    ordinary.clear();
                }
                output.push_str("<text:line-break/>");
            }
            value => ordinary.push(value),
        }
    }
    if !ordinary.is_empty() {
        output.push_str(&escape_xml(&ordinary));
    }
}

fn escape_xml(value: &str) -> String {
    quick_xml::escape::escape(value).into_owned()
}

fn decimal(value: f64) -> String {
    let mut output = format!("{value:.10}");
    while output.contains('.') && output.ends_with('0') {
        output.pop();
    }
    if output.ends_with('.') {
        output.pop();
    }
    output
}

fn automatic_line_height_percent(value: i32) -> String {
    const SCALE: i64 = 10_000_000_000;
    let numerator = i64::from(value) * 5 * SCALE;
    let scaled = if value == 24_000 {
        numerator / 12
    } else {
        numerator / 12 + 1
    };
    let whole = scaled / SCALE;
    let fraction = scaled % SCALE;
    if fraction == 0 {
        return whole.to_string();
    }
    let mut fraction = format!("{fraction:010}");
    while fraction.ends_with('0') {
        fraction.pop();
    }
    format!("{whole}.{fraction}")
}

fn twips_points(value: i32) -> String {
    if value.unsigned_abs() == 20_000_000 {
        return (value / 20).to_string();
    }
    let points = f64::from(value) / 20.0;
    decimal(points + points.signum() * 0.000_000_001)
}

fn emu_points(value: i64) -> String {
    if value == 12_700_000_000 {
        return "1000000".to_string();
    }
    decimal((value as f64 + 0.25) / 12_700.0)
}

fn paragraph_properties_have_unsupported(properties: &CT_PPr) -> bool {
    properties.style_id.is_some()
        || properties.jc == Some(ST_Jc::Distribute)
        || (properties.line_rule.is_some() && properties.line_spacing.is_none())
        || (properties.num_ilvl.is_some() && properties.num_id.is_none())
        || (properties.ind_first_line.is_some() && properties.ind_hanging.is_some())
        || properties.before_autospacing.is_some()
        || properties.after_autospacing.is_some()
        || properties.keep_next.is_some()
        || properties.keep_lines.is_some()
        || properties.page_break_before.is_some()
        || properties.widow_control.is_some()
        || properties.suppress_auto_hyphens.is_some()
        || properties.borders.is_some()
        || properties.tabs.is_some()
        || properties.shading.is_some()
        || properties.rpr.is_some()
        || properties.sect_pr.is_some()
        || properties.numbering_revision.is_some()
        || !properties.numbering_revision_xml.is_empty()
        || properties.change.is_some()
        || !properties.revision_xml.is_empty()
}

fn run_properties_have_unsupported(properties: &CT_RPr) -> bool {
    let selected = selected_run_font(properties);
    let alternate_font = [
        properties.font_hansi.as_deref(),
        properties.font_east_asia.as_deref(),
        properties.font_cs.as_deref(),
    ]
    .into_iter()
    .flatten()
    .any(|font| Some(font) != selected);
    properties.style_id.is_some()
        || alternate_font
        || properties.font_ascii_theme.is_some()
        || properties.font_hansi_theme.is_some()
        || properties
            .bold_cs
            .is_some_and(|value| Some(value) != properties.bold)
        || properties
            .italic_cs
            .is_some_and(|value| Some(value) != properties.italic)
        || properties
            .underline
            .is_some_and(|value| !matches!(value, ST_Underline::None | ST_Underline::Single))
        || properties.dstrike.is_some()
        || properties
            .sz_cs
            .is_some_and(|value| Some(value) != properties.sz)
        || properties.color_theme.is_some()
        || properties.spacing.is_some()
        || properties.width_scale.is_some()
        || properties.position.is_some()
        || properties
            .vert_align
            .as_deref()
            .is_some_and(|value| !matches!(value, "superscript" | "subscript"))
        || properties.caps.is_some()
        || properties.small_caps.is_some()
        || properties.vanish.is_some()
        || !properties.revision_markers.is_empty()
        || properties.change.is_some()
        || !properties.revision_xml.is_empty()
}

fn cell_has_lossy_properties(properties: &CT_TcPr) -> bool {
    properties.width.is_some()
        || properties.borders.is_some()
        || properties.shading.is_some()
        || properties.v_align.is_some()
        || properties.no_wrap.is_some()
        || properties.text_direction.is_some()
        || properties.cnf_style.is_some()
        || !properties.extra_xml.is_empty()
}

fn cell_has_substantive_content(cell: &CT_Tc) -> bool {
    cell.content.iter().any(|content| match content {
        CellContent::Paragraph(paragraph) => paragraph != &CT_P::new(),
        CellContent::Table(_) | CellContent::ContentControl(_) => true,
    })
}

fn table_has_lossy_properties(properties: &CT_TblPr) -> bool {
    properties.width.is_some()
        || properties.style_id.is_some()
        || properties.jc.is_some()
        || properties.borders.is_some()
        || properties.cell_margin.is_some()
        || properties.layout.is_some()
        || properties.indent.is_some()
        || properties.shading.is_some()
        || properties.look.is_some()
        || properties.change.is_some()
        || !properties.revision_xml.is_empty()
}

fn table_column_count(table: &CT_Tbl) -> Result<usize> {
    let grid = table.grid.as_ref().map_or(0, |grid| grid.columns.len());
    let rows = table.rows.iter().try_fold(0_usize, |maximum, row| {
        row_width(row).map(|width| maximum.max(width))
    })?;
    let columns = grid.max(rows);
    if columns == 0 || columns > OdtLimits::DEFAULT.columns {
        return Err(odt_error(
            Some("content.xml"),
            0,
            "ODT table column count is invalid",
        ));
    }
    Ok(columns)
}

fn validate_word_table(table: &CT_Tbl, path: &str) -> Result<()> {
    let columns = table_column_count(table)?;
    for (row_index, row) in table.rows.iter().enumerate() {
        if row_width(row)? != columns {
            return Err(odt_error(
                Some("content.xml"),
                0,
                format!("malformed Word table grid at {path}/row[{row_index}]"),
            ));
        }
    }
    for (row_index, row) in table.rows.iter().enumerate() {
        for (cell_index, cell) in row.cells.iter().enumerate() {
            if matches!(
                cell.properties
                    .as_ref()
                    .and_then(|properties| properties.v_merge.as_ref()),
                Some(VMerge::Continue)
            ) {
                let column = row_cell_column(row, cell_index)?;
                let span = cell
                    .properties
                    .as_ref()
                    .and_then(|properties| properties.grid_span)
                    .unwrap_or(1);
                let valid = if row_index == 0 {
                    false
                } else {
                    match row_cell_covering_column(&table.rows[row_index - 1], column)? {
                        Some((candidate_column, candidate)) => {
                            let properties = candidate.properties.as_ref();
                            let candidate_span =
                                properties.and_then(|value| value.grid_span).unwrap_or(1);
                            let merge = properties.and_then(|value| value.v_merge.as_ref());
                            candidate_column == column
                                && candidate_span == span
                                && matches!(merge, Some(VMerge::Restart | VMerge::Continue))
                        }
                        None => false,
                    }
                };
                if !valid {
                    return Err(odt_error(
                        Some("content.xml"),
                        0,
                        format!(
                            "malformed vertical table span at {path}/row[{row_index}]/cell[{cell_index}]"
                        ),
                    ));
                }
            }
        }
    }
    Ok(())
}

fn row_cell_covering_column(row: &CT_Row, column: usize) -> Result<Option<(usize, &CT_Tc)>> {
    let mut start = 0_usize;
    for cell in &row.cells {
        let end = start
            .checked_add(cell_grid_span(cell)?)
            .ok_or_else(|| odt_error(Some("content.xml"), 0, "ODT table grid span overflows"))?;
        if (start..end).contains(&column) {
            return Ok(Some((start, cell)));
        }
        start = end;
    }
    Ok(None)
}

fn cell_grid_span(cell: &CT_Tc) -> Result<usize> {
    let span = cell
        .properties
        .as_ref()
        .and_then(|properties| properties.grid_span)
        .unwrap_or(1);
    usize::try_from(span)
        .ok()
        .filter(|span| *span > 0)
        .ok_or_else(|| odt_error(Some("content.xml"), 0, "ODT table grid span is invalid"))
}

fn row_width(row: &CT_Row) -> Result<usize> {
    row.cells.iter().try_fold(0_usize, |column, cell| {
        column
            .checked_add(cell_grid_span(cell)?)
            .ok_or_else(|| odt_error(Some("content.xml"), 0, "ODT table grid span overflows"))
    })
}

fn row_cell_column(row: &CT_Row, cell_index: usize) -> Result<usize> {
    row.cells[..cell_index]
        .iter()
        .try_fold(0_usize, |column, cell| {
            column
                .checked_add(cell_grid_span(cell)?)
                .ok_or_else(|| odt_error(Some("content.xml"), 0, "ODT table grid span overflows"))
        })
}

fn table_rowspan(table: &CT_Tbl, row_index: usize, cell_index: usize) -> Result<usize> {
    let row = &table.rows[row_index];
    let column = row_cell_column(row, cell_index)?;
    let span = row.cells[cell_index]
        .properties
        .as_ref()
        .and_then(|properties| properties.grid_span)
        .unwrap_or(1);
    let mut rowspan = 1_usize;
    for candidate_row in table.rows.iter().skip(row_index + 1) {
        let Some((_, candidate)) = candidate_row.cells.iter().enumerate().find(|(index, _)| {
            row_cell_column(candidate_row, *index)
                .is_ok_and(|candidate_column| candidate_column == column)
        }) else {
            break;
        };
        let properties = candidate.properties.as_ref();
        if !matches!(
            properties.and_then(|value| value.v_merge.as_ref()),
            Some(VMerge::Continue)
        ) || properties.and_then(|value| value.grid_span).unwrap_or(1) != span
        {
            break;
        }
        rowspan += 1;
    }
    if rowspan == 1 {
        return Err(odt_error(
            Some("content.xml"),
            0,
            "vertical table merge restart has no continuation",
        ));
    }
    Ok(rowspan)
}

fn from_odt_with_limits(bytes: &[u8], limits: OdtLimits) -> Result<OdtReadResult> {
    let mut archive = OdtArchive::open(bytes, limits.archive)?;
    let mimetype = archive.read_required("mimetype")?;
    if mimetype.as_slice() != ODT_MIMETYPE {
        return Err(odt_error(
            Some("mimetype"),
            0,
            "ODT mimetype is missing or invalid",
        ));
    }
    let content_bytes = archive.read_required("content.xml")?;
    let styles_bytes = archive.read_optional("styles.xml")?;
    let manifest_bytes = archive.read_optional("META-INF/manifest.xml")?;

    let manifest = manifest_bytes
        .as_deref()
        .map(|xml| parse_xml("META-INF/manifest.xml", xml, limits))
        .transpose()?;
    let encrypted = encrypted_manifest_paths(manifest.as_ref())?;
    for required_part in ["/", "content.xml"] {
        if encrypted.contains(required_part) {
            return Err(odt_error(
                Some(required_part),
                0,
                "encrypted required ODT content is unsupported",
            ));
        }
    }
    if styles_bytes.is_some() && encrypted.contains("styles.xml") {
        return Err(odt_error(
            Some("styles.xml"),
            0,
            "encrypted ODT styles are unsupported",
        ));
    }

    let content = parse_xml("content.xml", &content_bytes, limits)?;
    let styles = styles_bytes
        .as_deref()
        .map(|xml| parse_xml("styles.xml", xml, limits))
        .transpose()?;

    Importer::new(archive, content, styles, encrypted, limits).project()
}

struct OdtArchive<'a> {
    archive: ZipArchive<Cursor<&'a [u8]>>,
    names: HashSet<String>,
    limit: u64,
}

impl<'a> OdtArchive<'a> {
    fn open(bytes: &'a [u8], limits: PackageReadLimits) -> Result<Self> {
        let raw_names = central_directory_names(bytes)?;
        if raw_names.len() > limits.max_entries {
            return Err(odt_error(None, 0, "ODT archive exceeds the entry limit"));
        }
        let mut names = HashSet::with_capacity(raw_names.len());
        for name in raw_names {
            validate_entry_name(&name)?;
            if !names.insert(name.clone()) {
                return Err(odt_error(Some(&name), 0, "duplicate ODT archive entry"));
            }
        }
        let mut archive = ZipArchive::new(Cursor::new(bytes))
            .map_err(|error| odt_error(None, 0, format!("invalid ODT ZIP: {error}")))?;
        if archive.len() > limits.max_entries {
            return Err(odt_error(None, 0, "ODT archive exceeds the entry limit"));
        }
        let mut total = 0_u64;
        for index in 0..archive.len() {
            let entry = archive.by_index(index).map_err(|error| {
                odt_error(None, 0, format!("cannot index ODT archive: {error}"))
            })?;
            let name = entry.name();
            validate_entry_name(name)?;
            if !entry.is_file() {
                return Err(odt_error(Some(name), 0, "ODT entry is not a regular file"));
            }
            if entry.encrypted() {
                return Err(odt_error(
                    Some(name),
                    0,
                    "encrypted ZIP entries are unsupported",
                ));
            }
            if !matches!(
                entry.compression(),
                CompressionMethod::Stored | CompressionMethod::Deflated
            ) {
                return Err(odt_error(
                    Some(name),
                    0,
                    "ODT entry uses unsupported compression",
                ));
            }
            if entry.size() > limits.max_part_uncompressed_bytes {
                return Err(odt_error(Some(name), 0, "ODT part exceeds the size limit"));
            }
            total = total
                .checked_add(entry.size())
                .filter(|total| *total <= limits.max_total_uncompressed_bytes)
                .ok_or_else(|| odt_error(None, 0, "ODT archive exceeds the expansion limit"))?;
        }
        if !names.contains("mimetype") || !names.contains("content.xml") {
            return Err(odt_error(None, 0, "ODT requires mimetype and content.xml"));
        }
        Ok(Self {
            archive,
            names,
            limit: limits.max_part_uncompressed_bytes,
        })
    }

    fn read_required(&mut self, name: &str) -> Result<Vec<u8>> {
        self.read_optional(name)?.ok_or_else(|| {
            odt_error(
                Some(name),
                0,
                format!("required ODT part {name} is missing"),
            )
        })
    }

    fn read_optional(&mut self, name: &str) -> Result<Option<Vec<u8>>> {
        if !self.names.contains(name) {
            return Ok(None);
        }
        let mut entry = self
            .archive
            .by_name(name)
            .map_err(|error| odt_error(Some(name), 0, format!("cannot read ODT part: {error}")))?;
        let mut bytes = Vec::with_capacity(usize::try_from(entry.size()).unwrap_or(0));
        entry
            .by_ref()
            .take(self.limit.saturating_add(1))
            .read_to_end(&mut bytes)
            .map_err(|error| {
                odt_error(Some(name), 0, format!("cannot expand ODT part: {error}"))
            })?;
        if bytes.len() as u64 > self.limit {
            return Err(odt_error(Some(name), 0, "ODT part exceeds the size limit"));
        }
        Ok(Some(bytes))
    }
}

fn central_directory_names(bytes: &[u8]) -> Result<Vec<String>> {
    const EOCD: &[u8; 4] = b"PK\x05\x06";
    const CENTRAL: &[u8; 4] = b"PK\x01\x02";
    let search_start = bytes.len().saturating_sub(65_557);
    let eocd = (search_start..bytes.len().saturating_sub(3))
        .rev()
        .find(|offset| &bytes[*offset..*offset + 4] == EOCD)
        .ok_or_else(|| odt_error(None, 0, "ODT ZIP end record is missing"))?;
    if eocd + 22 > bytes.len() {
        return Err(odt_error(None, eocd as u64, "truncated ODT ZIP end record"));
    }
    let disk = read_u16(bytes, eocd + 4)?;
    let central_disk = read_u16(bytes, eocd + 6)?;
    let disk_entries = read_u16(bytes, eocd + 8)?;
    let entries = read_u16(bytes, eocd + 10)?;
    let central_size = read_u32(bytes, eocd + 12)?;
    let central_offset = read_u32(bytes, eocd + 16)?;
    let comment = read_u16(bytes, eocd + 20)? as usize;
    if disk != 0 || central_disk != 0 || disk_entries != entries {
        return Err(odt_error(
            None,
            eocd as u64,
            "multi-disk ODT ZIP is unsupported",
        ));
    }
    if entries == u16::MAX || central_size == u32::MAX || central_offset == u32::MAX {
        return Err(odt_error(
            None,
            eocd as u64,
            "ZIP64 ODT archives are unsupported",
        ));
    }
    if eocd + 22 + comment != bytes.len() {
        return Err(odt_error(
            None,
            eocd as u64,
            "invalid ODT ZIP comment length",
        ));
    }
    let mut position = central_offset as usize;
    let central_end = position
        .checked_add(central_size as usize)
        .filter(|end| *end <= eocd)
        .ok_or_else(|| {
            odt_error(
                None,
                position as u64,
                "invalid ODT central directory bounds",
            )
        })?;
    let mut names = Vec::with_capacity(entries as usize);
    for _ in 0..entries {
        if position + 46 > central_end || &bytes[position..position + 4] != CENTRAL {
            return Err(odt_error(
                None,
                position as u64,
                "invalid ODT central directory entry",
            ));
        }
        let name_len = read_u16(bytes, position + 28)? as usize;
        let extra_len = read_u16(bytes, position + 30)? as usize;
        let comment_len = read_u16(bytes, position + 32)? as usize;
        let name_start = position + 46;
        let next = name_start
            .checked_add(name_len)
            .and_then(|value| value.checked_add(extra_len))
            .and_then(|value| value.checked_add(comment_len))
            .filter(|next| *next <= central_end)
            .ok_or_else(|| odt_error(None, position as u64, "truncated ODT central entry"))?;
        let name = std::str::from_utf8(&bytes[name_start..name_start + name_len])
            .map_err(|_| odt_error(None, position as u64, "ODT entry name is not UTF-8"))?;
        names.push(name.to_string());
        position = next;
    }
    if position != central_end {
        return Err(odt_error(
            None,
            position as u64,
            "unexpected ODT central directory bytes",
        ));
    }
    Ok(names)
}

fn read_u16(bytes: &[u8], offset: usize) -> Result<u16> {
    let value = bytes
        .get(offset..offset + 2)
        .ok_or_else(|| odt_error(None, offset as u64, "truncated ODT ZIP field"))?;
    Ok(u16::from_le_bytes([value[0], value[1]]))
}

fn read_u32(bytes: &[u8], offset: usize) -> Result<u32> {
    let value = bytes
        .get(offset..offset + 4)
        .ok_or_else(|| odt_error(None, offset as u64, "truncated ODT ZIP field"))?;
    Ok(u32::from_le_bytes([value[0], value[1], value[2], value[3]]))
}

fn validate_entry_name(name: &str) -> Result<()> {
    if name.is_empty()
        || name.starts_with('/')
        || name.starts_with('\\')
        || name.contains('\\')
        || name.contains('\0')
        || name
            .split('/')
            .any(|component| component.is_empty() || component == "." || component == "..")
    {
        return Err(odt_error(Some(name), 0, "unsafe ODT archive entry name"));
    }
    Ok(())
}

#[derive(Clone, Debug)]
struct XmlName {
    namespace: Option<String>,
    local: String,
}

#[derive(Clone, Debug)]
struct XmlAttribute {
    name: XmlName,
    value: String,
}

#[derive(Clone, Debug)]
enum XmlChild {
    Element(XmlNode),
    Text(String),
}

#[derive(Clone, Debug)]
struct XmlNode {
    name: XmlName,
    attributes: Vec<XmlAttribute>,
    children: Vec<XmlChild>,
}

impl XmlNode {
    fn is(&self, namespace: &str, local: &str) -> bool {
        self.name.namespace.as_deref() == Some(namespace) && self.name.local == local
    }

    fn attr(&self, namespace: Option<&str>, local: &str) -> Option<&str> {
        self.attributes
            .iter()
            .find(|attribute| {
                attribute.name.namespace.as_deref() == namespace && attribute.name.local == local
            })
            .map(|attribute| attribute.value.as_str())
    }

    fn elements(&self) -> impl Iterator<Item = &XmlNode> {
        self.children.iter().filter_map(|child| match child {
            XmlChild::Element(element) => Some(element),
            XmlChild::Text(_) => None,
        })
    }
}

fn parse_xml(part: &str, xml: &[u8], limits: OdtLimits) -> Result<XmlNode> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut stack: Vec<XmlNode> = Vec::new();
    let mut root = None;
    let mut nodes = 0_usize;
    let mut text_bytes = 0_usize;
    loop {
        let offset = reader.buffer_position();
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| odt_error(Some(part), offset, format!("malformed XML: {error}")))?;
        let namespace = namespace_value(namespace, part, offset)?;
        match event {
            Event::Start(element) => {
                nodes = nodes.checked_add(1).ok_or_else(|| {
                    odt_error(Some(part), offset, "ODT XML node count overflowed")
                })?;
                if nodes > limits.xml_nodes {
                    return Err(odt_error(
                        Some(part),
                        offset,
                        "ODT XML exceeds the node limit",
                    ));
                }
                if stack.len() >= limits.xml_depth {
                    return Err(odt_error(
                        Some(part),
                        offset,
                        "ODT XML exceeds the depth limit",
                    ));
                }
                stack.push(xml_node(&reader, namespace, &element, part, offset)?);
            }
            Event::Empty(element) => {
                nodes = nodes.checked_add(1).ok_or_else(|| {
                    odt_error(Some(part), offset, "ODT XML node count overflowed")
                })?;
                if nodes > limits.xml_nodes {
                    return Err(odt_error(
                        Some(part),
                        offset,
                        "ODT XML exceeds the node limit",
                    ));
                }
                if stack.len() >= limits.xml_depth {
                    return Err(odt_error(
                        Some(part),
                        offset,
                        "ODT XML exceeds the depth limit",
                    ));
                }
                attach_node(
                    &mut stack,
                    &mut root,
                    xml_node(&reader, namespace, &element, part, offset)?,
                    part,
                    offset,
                )?;
            }
            Event::End(_) => {
                let node = stack.pop().ok_or_else(|| {
                    odt_error(Some(part), offset, "ODT XML end tag has no open element")
                })?;
                attach_node(&mut stack, &mut root, node, part, offset)?;
            }
            Event::Text(text) => {
                let decoded = text.xml10_content().map_err(|error| {
                    odt_error(Some(part), offset, format!("invalid XML text: {error}"))
                })?;
                let value = quick_xml::escape::unescape(&decoded)
                    .map_err(|error| {
                        odt_error(
                            Some(part),
                            offset,
                            format!("unresolved XML entity: {error}"),
                        )
                    })?
                    .into_owned();
                append_text(&mut stack, value, &mut text_bytes, limits, part, offset)?;
            }
            Event::CData(text) => {
                let value = text
                    .xml10_content()
                    .map_err(|error| {
                        odt_error(Some(part), offset, format!("invalid XML CDATA: {error}"))
                    })?
                    .into_owned();
                append_text(&mut stack, value, &mut text_bytes, limits, part, offset)?;
            }
            Event::GeneralRef(reference) => {
                let value = if let Some(character) =
                    reference.resolve_char_ref().map_err(|error| {
                        odt_error(
                            Some(part),
                            offset,
                            format!("invalid XML reference: {error}"),
                        )
                    })? {
                    character.to_string()
                } else {
                    let reference_bytes: &[u8] = &reference;
                    match reference_bytes {
                        b"amp" => "&".to_string(),
                        b"lt" => "<".to_string(),
                        b"gt" => ">".to_string(),
                        b"apos" => "'".to_string(),
                        b"quot" => "\"".to_string(),
                        _ => {
                            return Err(odt_error(
                                Some(part),
                                offset,
                                "unresolved XML entity is unsupported",
                            ));
                        }
                    }
                };
                append_text(&mut stack, value, &mut text_bytes, limits, part, offset)?;
            }
            Event::DocType(_) => {
                return Err(odt_error(
                    Some(part),
                    offset,
                    "DTD declarations are forbidden",
                ));
            }
            Event::Eof => break,
            Event::Decl(_) | Event::Comment(_) | Event::PI(_) => {}
        }
        buffer.clear();
    }
    if !stack.is_empty() {
        return Err(odt_error(
            Some(part),
            xml.len() as u64,
            "unclosed XML element",
        ));
    }
    root.ok_or_else(|| odt_error(Some(part), 0, "ODT XML has no root element"))
}

fn namespace_value(
    namespace: ResolveResult<'_>,
    part: &str,
    offset: u64,
) -> Result<Option<String>> {
    match namespace {
        ResolveResult::Bound(namespace) => Ok(Some(
            std::str::from_utf8(namespace.as_ref())
                .map_err(|_| odt_error(Some(part), offset, "namespace URI is not UTF-8"))?
                .to_string(),
        )),
        ResolveResult::Unbound => Ok(None),
        ResolveResult::Unknown(_) => Err(odt_error(
            Some(part),
            offset,
            "ODT XML uses an unresolved namespace prefix",
        )),
    }
}

fn xml_node(
    reader: &NsReader<&[u8]>,
    namespace: Option<String>,
    element: &BytesStart<'_>,
    part: &str,
    offset: u64,
) -> Result<XmlNode> {
    let local = std::str::from_utf8(element.local_name().as_ref())
        .map_err(|_| odt_error(Some(part), offset, "element local name is not UTF-8"))?
        .to_string();
    let mut attributes = Vec::new();
    let mut seen = HashSet::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| {
            odt_error(
                Some(part),
                offset,
                format!("malformed XML attribute: {error}"),
            )
        })?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = namespace_value(namespace, part, offset)?;
        let local = std::str::from_utf8(local.as_ref())
            .map_err(|_| odt_error(Some(part), offset, "attribute local name is not UTF-8"))?
            .to_string();
        if !seen.insert((namespace.clone(), local.clone())) {
            return Err(odt_error(
                Some(part),
                offset,
                "duplicate expanded XML attribute",
            ));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| {
                odt_error(
                    Some(part),
                    offset,
                    format!("invalid XML attribute value: {error}"),
                )
            })?
            .into_owned();
        attributes.push(XmlAttribute {
            name: XmlName { namespace, local },
            value,
        });
    }
    Ok(XmlNode {
        name: XmlName { namespace, local },
        attributes,
        children: Vec::new(),
    })
}

fn attach_node(
    stack: &mut [XmlNode],
    root: &mut Option<XmlNode>,
    node: XmlNode,
    part: &str,
    offset: u64,
) -> Result<()> {
    if let Some(parent) = stack.last_mut() {
        parent.children.push(XmlChild::Element(node));
    } else if root.replace(node).is_some() {
        return Err(odt_error(
            Some(part),
            offset,
            "ODT XML has multiple root elements",
        ));
    }
    Ok(())
}

fn append_text(
    stack: &mut [XmlNode],
    value: String,
    retained: &mut usize,
    limits: OdtLimits,
    part: &str,
    offset: u64,
) -> Result<()> {
    if value.is_empty() {
        return Ok(());
    }
    if stack.is_empty() && value.chars().all(char::is_whitespace) {
        return Ok(());
    }
    *retained = retained
        .checked_add(value.len())
        .filter(|retained| *retained <= limits.retained_text)
        .ok_or_else(|| {
            odt_error(
                Some(part),
                offset,
                "ODT XML exceeds the retained text limit",
            )
        })?;
    let parent = stack
        .last_mut()
        .ok_or_else(|| odt_error(Some(part), offset, "text appears outside the XML root"))?;
    parent.children.push(XmlChild::Text(value));
    Ok(())
}

fn encrypted_manifest_paths(manifest: Option<&XmlNode>) -> Result<HashSet<String>> {
    let mut encrypted = HashSet::new();
    let Some(manifest) = manifest else {
        return Ok(encrypted);
    };
    if !manifest.is(MANIFEST_NS, "manifest") {
        return Err(odt_error(
            Some("META-INF/manifest.xml"),
            0,
            "invalid ODT manifest root",
        ));
    }
    for entry in manifest
        .elements()
        .filter(|entry| entry.is(MANIFEST_NS, "file-entry"))
    {
        let path = entry
            .attr(Some(MANIFEST_NS), "full-path")
            .or_else(|| entry.attr(None, "full-path"));
        if entry
            .elements()
            .any(|child| child.is(MANIFEST_NS, "encryption-data"))
        {
            let path = path.ok_or_else(|| {
                odt_error(
                    Some("META-INF/manifest.xml"),
                    0,
                    "encrypted manifest entry has no full path",
                )
            })?;
            encrypted.insert(path.to_string());
        }
    }
    Ok(encrypted)
}

#[derive(Clone, Debug, Default, PartialEq)]
struct EffectiveStyle {
    font: Option<String>,
    size: Option<f64>,
    bold: Option<bool>,
    italic: Option<bool>,
    underline: Option<bool>,
    strike: Option<bool>,
    color: Option<String>,
    background: Option<String>,
    vertical: Option<VerticalText>,
    alignment: Option<Alignment>,
    before: Option<Length>,
    after: Option<Length>,
    left: Option<Length>,
    right: Option<Length>,
    first: Option<Length>,
    line: Option<LineHeight>,
}

impl EffectiveStyle {
    fn overlay(&mut self, other: &Self) {
        macro_rules! replace {
            ($field:ident) => {
                if other.$field.is_some() {
                    self.$field = other.$field.clone();
                }
            };
        }
        replace!(font);
        replace!(size);
        replace!(bold);
        replace!(italic);
        replace!(underline);
        replace!(strike);
        replace!(color);
        replace!(background);
        replace!(vertical);
        replace!(alignment);
        replace!(before);
        replace!(after);
        replace!(left);
        replace!(right);
        replace!(first);
        replace!(line);
    }
}

#[derive(Clone, Copy, Debug, PartialEq)]
enum VerticalText {
    Superscript,
    Subscript,
}

#[derive(Clone, Copy, Debug, PartialEq)]
enum LineHeight {
    Exact(f64),
    Multiple(f64),
}

#[derive(Clone, Debug)]
struct StyleDef {
    family: String,
    name: String,
    parent: Option<String>,
    values: EffectiveStyle,
}

struct Importer<'a> {
    archive: OdtArchive<'a>,
    content: XmlNode,
    styles_root: Option<XmlNode>,
    encrypted: HashSet<String>,
    styles: HashMap<(String, String), StyleDef>,
    defaults: HashMap<String, EffectiveStyle>,
    resolved: HashMap<(String, String), EffectiveStyle>,
    list_kinds: HashMap<String, Vec<bool>>,
    fonts: HashMap<String, String>,
    document: Document,
    diagnostics: Vec<OdtDiagnostic>,
    diagnostic_keys: HashSet<(String, String)>,
    limits: OdtLimits,
    blocks: usize,
    runs: usize,
    rows: usize,
    cells: usize,
    projected_text: usize,
}

impl<'a> Importer<'a> {
    fn new(
        archive: OdtArchive<'a>,
        content: XmlNode,
        styles_root: Option<XmlNode>,
        encrypted: HashSet<String>,
        limits: OdtLimits,
    ) -> Self {
        Self {
            archive,
            content,
            styles_root,
            encrypted,
            styles: HashMap::new(),
            defaults: HashMap::new(),
            resolved: HashMap::new(),
            list_kinds: HashMap::new(),
            fonts: HashMap::new(),
            document: Document::new(),
            diagnostics: Vec::new(),
            diagnostic_keys: HashSet::new(),
            limits,
            blocks: 0,
            runs: 0,
            rows: 0,
            cells: 0,
            projected_text: 0,
        }
    }

    fn project(mut self) -> Result<OdtReadResult> {
        if !self.content.is(OFFICE_NS, "document-content") {
            return Err(odt_error(
                Some("content.xml"),
                0,
                "invalid ODT content root",
            ));
        }
        if let Some(styles) = self.styles_root.clone() {
            if !styles.is(OFFICE_NS, "document-styles") {
                return Err(odt_error(Some("styles.xml"), 0, "invalid ODT styles root"));
            }
            self.collect_style_document(&styles, "styles.xml")?;
        }
        let content = self.content.clone();
        self.collect_style_document(&content, "content.xml")?;
        self.resolve_all_styles()?;

        let body = content
            .elements()
            .find(|node| node.is(OFFICE_NS, "body"))
            .and_then(|node| node.elements().find(|node| node.is(OFFICE_NS, "text")))
            .ok_or_else(|| odt_error(Some("content.xml"), 0, "ODT text body is missing"))?;
        self.project_container(body, "office:body/office:text", None)?;

        let bytes = self.document.to_bytes()?;
        let document = Document::from_bytes(&bytes)?;
        Ok(OdtReadResult {
            document,
            diagnostics: self.diagnostics,
        })
    }

    fn collect_style_document(&mut self, root: &XmlNode, part: &str) -> Result<()> {
        self.collect_fonts(root);
        self.walk_style_containers(root, part)
    }

    fn collect_fonts(&mut self, node: &XmlNode) {
        if node.is(STYLE_NS, "font-face")
            && let Some(name) = node.attr(Some(STYLE_NS), "name")
            && let Some(family) = node.attr(Some(SVG_NS), "font-family")
        {
            self.fonts.insert(
                name.to_string(),
                family.trim_matches(['\'', '"']).to_string(),
            );
        }
        for child in node.elements() {
            self.collect_fonts(child);
        }
    }

    fn walk_style_containers(&mut self, node: &XmlNode, part: &str) -> Result<()> {
        if node.is(STYLE_NS, "default-style") {
            let family = required_attr(node, STYLE_NS, "family", part)?;
            let values =
                self.parse_style_values(node, part, &format!("style:default-style[{family}]"))?;
            self.defaults
                .entry(family.to_string())
                .or_default()
                .overlay(&values);
        } else if node.is(STYLE_NS, "style") {
            let family = required_attr(node, STYLE_NS, "family", part)?;
            let name = required_attr(node, STYLE_NS, "name", part)?;
            let values = self.parse_style_values(node, part, &format!("style:style[{name}]"))?;
            let def = StyleDef {
                family: family.to_string(),
                name: name.to_string(),
                parent: node
                    .attr(Some(STYLE_NS), "parent-style-name")
                    .map(str::to_string),
                values,
            };
            let key = (def.family.clone(), def.name.clone());
            if self.styles.insert(key, def).is_some() {
                return Err(odt_error(Some(part), 0, "duplicate ODT style definition"));
            }
        } else if node.is(TEXT_NS, "list-style") {
            let name = required_attr(node, STYLE_NS, "name", part)?.to_string();
            let mut levels = vec![false; 9];
            for level in node.elements() {
                let index = level
                    .attr(Some(TEXT_NS), "level")
                    .and_then(|value| value.parse::<usize>().ok())
                    .unwrap_or(1)
                    .saturating_sub(1)
                    .min(8);
                if level.is(TEXT_NS, "list-level-style-bullet") {
                    levels[index] = true;
                } else if level.is(TEXT_NS, "list-level-style-number") {
                    levels[index] = false;
                }
            }
            self.list_kinds.insert(name, levels);
        }
        for child in node.elements() {
            self.walk_style_containers(child, part)?;
        }
        Ok(())
    }

    fn parse_style_values(
        &mut self,
        style: &XmlNode,
        part: &str,
        path: &str,
    ) -> Result<EffectiveStyle> {
        let mut values = EffectiveStyle::default();
        for properties in style.elements() {
            if properties.is(STYLE_NS, "text-properties") {
                for attribute in &properties.attributes {
                    let result = parse_text_property(attribute, &self.fonts);
                    match result {
                        Some(Ok(change)) => change.apply(&mut values),
                        Some(Err(message)) => self.diagnostic(path, message)?,
                        None if !is_ignorable_style_attribute(attribute) => self.diagnostic(
                            path,
                            format!("unsupported text property {}", attribute.name.local),
                        )?,
                        None => {}
                    }
                }
            } else if properties.is(STYLE_NS, "paragraph-properties") {
                for attribute in &properties.attributes {
                    let result = parse_paragraph_property(attribute);
                    match result {
                        Some(Ok(change)) => change.apply(&mut values),
                        Some(Err(message)) => self.diagnostic(path, message)?,
                        None if !is_ignorable_style_attribute(attribute) => self.diagnostic(
                            path,
                            format!("unsupported paragraph property {}", attribute.name.local),
                        )?,
                        None => {}
                    }
                }
            }
        }
        let _ = part;
        Ok(values)
    }

    fn resolve_all_styles(&mut self) -> Result<()> {
        let keys: Vec<(String, String)> = self.styles.keys().cloned().collect();
        for key in keys {
            self.resolve_style(&key.0, &key.1, &mut Vec::new())?;
        }
        Ok(())
    }

    fn resolve_style(
        &mut self,
        family: &str,
        name: &str,
        visiting: &mut Vec<(String, String)>,
    ) -> Result<EffectiveStyle> {
        let key = (family.to_string(), name.to_string());
        if let Some(value) = self.resolved.get(&key) {
            return Ok(value.clone());
        }
        if visiting.contains(&key) {
            return Err(odt_error(
                Some("styles.xml"),
                0,
                format!("ODT style inheritance cycle at {family}/{name}"),
            ));
        }
        let Some(definition) = self.styles.get(&key).cloned() else {
            self.diagnostic(
                &format!("style:{family}/{name}"),
                format!("missing ODT style {name}"),
            )?;
            return Ok(self.defaults.get(family).cloned().unwrap_or_default());
        };
        visiting.push(key.clone());
        let mut value = if let Some(parent) = definition.parent.as_deref() {
            self.resolve_style(family, parent, visiting)?
        } else {
            self.defaults.get(family).cloned().unwrap_or_default()
        };
        visiting.pop();
        value.overlay(&definition.values);
        self.resolved.insert(key, value.clone());
        Ok(value)
    }

    fn effective_style(&mut self, family: &str, name: Option<&str>) -> Result<EffectiveStyle> {
        match name {
            Some(name) => self.resolve_style(family, name, &mut Vec::new()),
            None => Ok(self.defaults.get(family).cloned().unwrap_or_default()),
        }
    }

    fn project_container(
        &mut self,
        container: &XmlNode,
        path: &str,
        list: Option<(u32, usize)>,
    ) -> Result<()> {
        let mut positions: HashMap<(Option<String>, String), usize> = HashMap::new();
        for child in container.elements() {
            let key = (child.name.namespace.clone(), child.name.local.clone());
            let index = positions.entry(key).or_default();
            *index += 1;
            let child_path = format!("{path}/{}[{}]", display_name(child), *index);
            if child.is(TEXT_NS, "p") || child.is(TEXT_NS, "h") {
                self.project_paragraph(child, &child_path, list)?;
            } else if child.is(TEXT_NS, "list") {
                self.project_list(child, &child_path, list.map_or(0, |(_, level)| level + 1))?;
            } else if child.is(TABLE_NS, "table") {
                self.project_table(child, &child_path)?;
            } else if child.is(TEXT_NS, "section") || child.is(OFFICE_NS, "text") {
                self.project_container(child, &child_path, list)?;
            } else if child.is(DRAW_NS, "frame") {
                let paragraph = XmlNode {
                    name: XmlName {
                        namespace: Some(TEXT_NS.to_string()),
                        local: "p".to_string(),
                    },
                    attributes: Vec::new(),
                    children: vec![XmlChild::Element(child.clone())],
                };
                self.project_paragraph(&paragraph, &child_path, list)?;
            } else {
                self.diagnostic(
                    &child_path,
                    format!("unsupported ODT subtree {}", display_name(child)),
                )?;
            }
        }
        Ok(())
    }

    fn project_list(&mut self, list: &XmlNode, path: &str, level: usize) -> Result<()> {
        let mut paragraphs = Vec::new();
        self.collect_list_paragraphs(list, path, level, &mut paragraphs)?;
        self.document
            .document
            .body
            .content
            .extend(paragraphs.into_iter().map(BodyContent::Paragraph));
        Ok(())
    }

    fn collect_list_paragraphs(
        &mut self,
        list: &XmlNode,
        path: &str,
        level: usize,
        output: &mut Vec<CT_P>,
    ) -> Result<()> {
        if level >= 9 {
            return Err(odt_error(
                Some("content.xml"),
                0,
                "ODT list exceeds nine levels",
            ));
        }
        if list.attr(Some(TEXT_NS), "continue-numbering").is_some()
            || list.attr(Some(TEXT_NS), "continue-list").is_some()
        {
            self.diagnostic(path, "unsupported ODT list continuation semantics")?;
        }
        let style_name = list.attr(Some(TEXT_NS), "style-name");
        let levels = style_name
            .and_then(|name| self.list_kinds.get(name))
            .cloned()
            .unwrap_or_else(|| vec![true; 9]);
        if style_name.is_some_and(|name| !self.list_kinds.contains_key(name)) {
            self.diagnostic(path, "missing ODT list style, using bullets")?;
        }
        let definitions: Vec<ListLevel> = levels
            .iter()
            .map(|bullet| {
                if *bullet {
                    ListLevel::bullet()
                } else {
                    ListLevel::decimal()
                }
            })
            .collect();
        let num_id = self.document.add_list_definition(&definitions);
        for (item_index, item) in list
            .elements()
            .filter(|item| item.is(TEXT_NS, "list-item") || item.is(TEXT_NS, "list-header"))
            .enumerate()
        {
            let item_name = if item.is(TEXT_NS, "list-header") {
                "list-header"
            } else {
                "list-item"
            };
            let item_path = format!("{path}/text:{item_name}[{}]", item_index + 1);
            if item.attr(Some(TEXT_NS), "start-value").is_some() {
                self.diagnostic(&item_path, "unsupported ODT list start value")?;
            }
            let numbering = item.is(TEXT_NS, "list-item").then_some((num_id, level));
            self.collect_list_item_paragraphs(item, &item_path, numbering, level, output)?;
        }
        Ok(())
    }

    fn collect_list_item_paragraphs(
        &mut self,
        container: &XmlNode,
        path: &str,
        numbering: Option<(u32, usize)>,
        level: usize,
        output: &mut Vec<CT_P>,
    ) -> Result<()> {
        let mut positions: HashMap<(Option<String>, String), usize> = HashMap::new();
        for child in container.elements() {
            let key = (child.name.namespace.clone(), child.name.local.clone());
            let child_index = positions.entry(key).or_default();
            *child_index += 1;
            let child_path = format!("{path}/{}[{}]", display_name(child), *child_index);
            if child.is(TEXT_NS, "p") || child.is(TEXT_NS, "h") {
                output.push(self.build_paragraph(child, &child_path, numbering)?);
            } else if child.is(TEXT_NS, "list") {
                self.collect_list_paragraphs(child, &child_path, level + 1, output)?;
            } else if child.is(TEXT_NS, "section") || child.is(OFFICE_NS, "text") {
                self.collect_list_item_paragraphs(child, &child_path, numbering, level, output)?;
            } else if child.is(DRAW_NS, "frame") {
                let paragraph = XmlNode {
                    name: XmlName {
                        namespace: Some(TEXT_NS.to_string()),
                        local: "p".to_string(),
                    },
                    attributes: Vec::new(),
                    children: vec![XmlChild::Element(child.clone())],
                };
                output.push(self.build_paragraph(&paragraph, &child_path, numbering)?);
            } else {
                self.diagnostic(
                    &child_path,
                    format!("unsupported ODT list subtree {}", display_name(child)),
                )?;
            }
        }
        Ok(())
    }

    fn project_paragraph(
        &mut self,
        node: &XmlNode,
        path: &str,
        list: Option<(u32, usize)>,
    ) -> Result<()> {
        let paragraph = self.build_paragraph(node, path, list)?;
        self.document
            .document
            .body
            .content
            .push(BodyContent::Paragraph(paragraph));
        Ok(())
    }

    fn build_paragraph(
        &mut self,
        node: &XmlNode,
        path: &str,
        list: Option<(u32, usize)>,
    ) -> Result<CT_P> {
        self.bump_blocks()?;
        let style_name = node.attr(Some(TEXT_NS), "style-name");
        let paragraph_style = self.effective_style("paragraph", style_name)?;
        let mut pieces = Vec::new();
        self.collect_inline(node, path, &paragraph_style, &mut pieces)?;
        let mut paragraph = CT_P::new();
        {
            let mut target = Paragraph {
                inner: &mut paragraph,
            };
            apply_paragraph_style(&mut target, &paragraph_style);
            if node.is(TEXT_NS, "h") {
                let level = node
                    .attr(Some(TEXT_NS), "outline-level")
                    .and_then(|value| value.parse::<u32>().ok())
                    .unwrap_or(1)
                    .clamp(1, 9);
                target.set_style(&format!("Heading{level}"));
                target.set_outline_level(level - 1);
            }
            if let Some((num_id, level)) = list {
                target.set_numbering(num_id, level as u32);
            }
            for piece in pieces {
                match piece {
                    InlinePiece::Text(text, style) => {
                        self.bump_runs()?;
                        let mut run = target.add_run(&text);
                        apply_run_style(&mut run, &style);
                    }
                    InlinePiece::Break => {
                        self.bump_runs()?;
                        target.add_line_break();
                    }
                    InlinePiece::Tab => {
                        self.bump_runs()?;
                        target.add_tab();
                    }
                    InlinePiece::Image {
                        relationship,
                        width,
                        height,
                    } => {
                        self.bump_runs()?;
                        target.add_picture(&relationship, width, height);
                    }
                }
            }
        }
        Ok(paragraph)
    }

    fn collect_inline(
        &mut self,
        node: &XmlNode,
        path: &str,
        inherited: &EffectiveStyle,
        output: &mut Vec<InlinePiece>,
    ) -> Result<()> {
        let mut whitespace = InlineWhitespace::default();
        self.collect_inline_inner(node, path, inherited, output, &mut whitespace)
    }

    fn collect_inline_inner(
        &mut self,
        node: &XmlNode,
        path: &str,
        inherited: &EffectiveStyle,
        output: &mut Vec<InlinePiece>,
        whitespace: &mut InlineWhitespace,
    ) -> Result<()> {
        let mut element_positions: HashMap<(Option<String>, String), usize> = HashMap::new();
        for child in &node.children {
            match child {
                XmlChild::Text(text) => {
                    self.bump_projected_text(text.len())?;
                    collapse_odt_text(text, inherited, whitespace, output);
                }
                XmlChild::Element(element) if element.is(TEXT_NS, "span") => {
                    let key = (element.name.namespace.clone(), element.name.local.clone());
                    let index = element_positions.entry(key).or_default();
                    *index += 1;
                    let child_path = format!("{path}/text:span[{}]", *index);
                    let mut style = inherited.clone();
                    let own =
                        self.effective_style("text", element.attr(Some(TEXT_NS), "style-name"))?;
                    style.overlay(&own);
                    self.collect_inline_inner(element, &child_path, &style, output, whitespace)?;
                }
                XmlChild::Element(element) if element.is(TEXT_NS, "s") => {
                    let count = repeated_count(element, TEXT_NS, "c", self.limits.retained_text)?;
                    self.bump_projected_text(count)?;
                    push_inline_text(output, " ".repeat(count), inherited.clone());
                    whitespace.pending = false;
                    whitespace.style = None;
                    whitespace.emitted = true;
                }
                XmlChild::Element(element) if element.is(TEXT_NS, "tab") => {
                    output.push(InlinePiece::Tab);
                    whitespace.pending = false;
                    whitespace.style = None;
                    whitespace.emitted = true;
                }
                XmlChild::Element(element) if element.is(TEXT_NS, "line-break") => {
                    output.push(InlinePiece::Break);
                    whitespace.pending = false;
                    whitespace.style = None;
                    whitespace.emitted = true;
                }
                XmlChild::Element(element) if element.is(DRAW_NS, "frame") => {
                    let key = (element.name.namespace.clone(), element.name.local.clone());
                    let index = element_positions.entry(key).or_default();
                    *index += 1;
                    let child_path = format!("{path}/draw:frame[{}]", *index);
                    if let Some(image) = self.project_image(element, &child_path)? {
                        output.push(image);
                        whitespace.pending = false;
                        whitespace.style = None;
                        whitespace.emitted = true;
                    }
                }
                XmlChild::Element(element) => {
                    let key = (element.name.namespace.clone(), element.name.local.clone());
                    let index = element_positions.entry(key).or_default();
                    *index += 1;
                    let child_path = format!("{path}/{}[{}]", display_name(element), *index);
                    self.diagnostic(
                        &child_path,
                        format!("unsupported ODT inline subtree {}", display_name(element)),
                    )?;
                }
            }
        }
        Ok(())
    }

    fn project_image(&mut self, frame: &XmlNode, path: &str) -> Result<Option<InlinePiece>> {
        let Some(image) = frame.elements().find(|node| node.is(DRAW_NS, "image")) else {
            self.diagnostic(path, "drawing frame has no supported image")?;
            return Ok(None);
        };
        let Some(href) = image.attr(Some(XLINK_NS), "href") else {
            self.diagnostic(path, "ODT image has no package target")?;
            return Ok(None);
        };
        let target = safe_image_target(href).ok_or_else(|| {
            odt_error(
                Some("content.xml"),
                0,
                format!("unsafe ODT image target {href}"),
            )
        })?;
        if self.encrypted.contains(&target) {
            return Err(odt_error(
                Some(&target),
                0,
                "referenced ODT image is encrypted",
            ));
        }
        let Some(bytes) = self.archive.read_optional(&target)? else {
            self.diagnostic(path, format!("missing ODT image target {target}"))?;
            return Ok(None);
        };
        let Some(info) = oxml_media::probe(&bytes) else {
            self.diagnostic(path, format!("unsupported ODT image content {target}"))?;
            return Ok(None);
        };
        let explicit = match (
            frame.attr(Some(SVG_NS), "width"),
            frame.attr(Some(SVG_NS), "height"),
        ) {
            (Some(width), Some(height)) => Some((
                parse_positive_length(width)
                    .map_err(|message| odt_error(Some("content.xml"), 0, message))?,
                parse_positive_length(height)
                    .map_err(|message| odt_error(Some("content.xml"), 0, message))?,
            )),
            (None, None) => None,
            _ => {
                self.diagnostic(
                    path,
                    "ODT image frame has only one dimension, using intrinsic size",
                )?;
                None
            }
        };
        let (width, height) = explicit
            .or_else(|| {
                info.native_size(72.0)
                    .map(|size| (Length::emu(size.width_emu), Length::emu(size.height_emu)))
            })
            .ok_or_else(|| odt_error(Some(&target), 0, "ODT image dimensions are unavailable"))?;
        let relationship = self.document.embed_image(&bytes, &target);
        Ok(Some(InlinePiece::Image {
            relationship,
            width,
            height,
        }))
    }

    fn project_table(&mut self, table: &XmlNode, path: &str) -> Result<()> {
        self.bump_blocks()?;
        let mut column_hint = 0_usize;
        for column in table
            .elements()
            .filter(|node| node.is(TABLE_NS, "table-column"))
        {
            column_hint = column_hint
                .checked_add(repeated_count(
                    column,
                    TABLE_NS,
                    "number-columns-repeated",
                    self.limits.columns,
                )?)
                .ok_or_else(|| odt_error(Some("content.xml"), 0, "ODT column count overflowed"))?;
        }
        let mut source_rows = Vec::new();
        collect_table_rows(table, &mut source_rows);
        let mut rows = Vec::new();
        for (row_index, row) in source_rows.into_iter().enumerate() {
            let repeated = repeated_count(row, TABLE_NS, "number-rows-repeated", self.limits.rows)?;
            for repeat in 0..repeated {
                self.rows = self
                    .rows
                    .checked_add(1)
                    .filter(|value| *value <= self.limits.rows)
                    .ok_or_else(|| {
                        odt_error(Some("content.xml"), 0, "ODT table exceeds the row limit")
                    })?;
                rows.push(self.parse_table_row(
                    row,
                    &format!("{path}/table:table-row[{}]", row_index + repeat + 1),
                )?);
            }
        }
        let columns = rows
            .iter()
            .map(|row| row.iter().map(|cell| cell.colspan).sum::<usize>())
            .max()
            .unwrap_or(column_hint)
            .max(column_hint);
        if columns == 0 || columns > self.limits.columns {
            return Err(odt_error(
                Some("content.xml"),
                0,
                "ODT table column count is invalid",
            ));
        }
        let mut tbl = CT_Tbl::new();
        let width = Twips(9000 / columns as i32);
        tbl.properties = Some(CT_TblPr {
            width: Some(CT_TblWidth::dxa(width.0 * columns as i32)),
            ..Default::default()
        });
        tbl.grid = Some(CT_TblGrid {
            columns: (0..columns).map(|_| CT_TblGridCol { width }).collect(),
            grid_change_xml: None,
        });
        let mut active: Vec<Option<(usize, usize)>> = vec![None; columns];
        for row in rows {
            let mut target = CT_Row::new();
            let mut column = 0_usize;
            let mut source = row.into_iter();
            while column < columns {
                if let Some((remaining, span)) = active[column] {
                    let mut cell = CT_Tc::new();
                    cell.properties = Some(CT_TcPr {
                        grid_span: (span > 1).then_some(span as u32),
                        v_merge: Some(VMerge::Continue),
                        ..Default::default()
                    });
                    target.cells.push(cell);
                    for slot in active.iter_mut().skip(column).take(span) {
                        *slot = if remaining > 1 {
                            Some((remaining - 1, span))
                        } else {
                            None
                        };
                    }
                    column += span;
                    continue;
                }
                let Some(model) = source.next() else {
                    target.cells.push(CT_Tc::new());
                    column += 1;
                    continue;
                };
                if column + model.colspan > columns {
                    return Err(odt_error(
                        Some("content.xml"),
                        0,
                        "ODT table span exceeds the grid",
                    ));
                }
                let mut cell = CT_Tc::new();
                cell.properties = Some(CT_TcPr {
                    grid_span: (model.colspan > 1).then_some(model.colspan as u32),
                    v_merge: (model.rowspan > 1).then_some(VMerge::Restart),
                    ..Default::default()
                });
                cell.content.clear();
                for paragraph in model.paragraphs {
                    cell.content.push(CellContent::Paragraph(paragraph));
                }
                if cell.paragraphs().is_empty() {
                    cell.content.push(CellContent::Paragraph(CT_P::new()));
                }
                if model.rowspan > 1 {
                    for slot in active.iter_mut().skip(column).take(model.colspan) {
                        if slot.is_some() {
                            return Err(odt_error(
                                Some("content.xml"),
                                0,
                                "overlapping ODT table spans",
                            ));
                        }
                        *slot = Some((model.rowspan - 1, model.colspan));
                    }
                }
                target.cells.push(cell);
                column += model.colspan;
            }
            if source.next().is_some() {
                return Err(odt_error(
                    Some("content.xml"),
                    0,
                    "ODT table row exceeds the grid",
                ));
            }
            tbl.rows.push(target);
        }
        self.document
            .document
            .body
            .content
            .push(BodyContent::Table(tbl));
        Ok(())
    }

    fn parse_table_row(&mut self, row: &XmlNode, path: &str) -> Result<Vec<TableCellModel>> {
        let mut output = Vec::new();
        for (index, cell) in row
            .elements()
            .filter(|node| {
                node.is(TABLE_NS, "table-cell") || node.is(TABLE_NS, "covered-table-cell")
            })
            .enumerate()
        {
            let repeat = repeated_count(
                cell,
                TABLE_NS,
                "number-columns-repeated",
                self.limits.columns,
            )?;
            for repeat_index in 0..repeat {
                self.cells = self
                    .cells
                    .checked_add(1)
                    .filter(|value| *value <= self.limits.cells)
                    .ok_or_else(|| {
                        odt_error(Some("content.xml"), 0, "ODT table exceeds the cell limit")
                    })?;
                if cell.is(TABLE_NS, "covered-table-cell") {
                    continue;
                }
                let colspan = repeated_count(
                    cell,
                    TABLE_NS,
                    "number-columns-spanned",
                    self.limits.columns,
                )?;
                let rowspan =
                    repeated_count(cell, TABLE_NS, "number-rows-spanned", self.limits.rows)?;
                let mut paragraphs = Vec::new();
                let cell_path = format!("{path}/table:table-cell[{}]", index + repeat_index + 1);
                let mut positions: HashMap<(Option<String>, String), usize> = HashMap::new();
                for child in cell.elements() {
                    let key = (child.name.namespace.clone(), child.name.local.clone());
                    let child_index = positions.entry(key).or_default();
                    *child_index += 1;
                    let child_path =
                        format!("{cell_path}/{}[{}]", display_name(child), *child_index);
                    if child.is(TEXT_NS, "p") || child.is(TEXT_NS, "h") {
                        paragraphs.push(self.build_paragraph(child, &child_path, None)?);
                    } else if child.is(TEXT_NS, "list") {
                        self.collect_list_paragraphs(child, &child_path, 0, &mut paragraphs)?;
                    } else {
                        self.diagnostic(
                            &child_path,
                            format!("unsupported ODT table-cell subtree {}", display_name(child)),
                        )?;
                    }
                }
                output.push(TableCellModel {
                    colspan,
                    rowspan,
                    paragraphs,
                });
            }
        }
        Ok(output)
    }

    fn diagnostic(&mut self, path: &str, message: impl Into<String>) -> Result<()> {
        let message = message.into();
        if !self
            .diagnostic_keys
            .insert((path.to_string(), message.clone()))
        {
            return Ok(());
        }
        if self.diagnostics.len() >= self.limits.diagnostics {
            return Err(odt_error(
                Some("content.xml"),
                0,
                "ODT exceeds the diagnostic limit",
            ));
        }
        self.diagnostics.push(OdtDiagnostic {
            path: path.to_string(),
            message,
        });
        Ok(())
    }

    fn bump_blocks(&mut self) -> Result<()> {
        self.blocks = self
            .blocks
            .checked_add(1)
            .filter(|value| *value <= self.limits.blocks)
            .ok_or_else(|| {
                odt_error(
                    Some("content.xml"),
                    0,
                    "ODT exceeds the projected block limit",
                )
            })?;
        Ok(())
    }

    fn bump_runs(&mut self) -> Result<()> {
        self.runs = self
            .runs
            .checked_add(1)
            .filter(|value| *value <= self.limits.runs)
            .ok_or_else(|| {
                odt_error(
                    Some("content.xml"),
                    0,
                    "ODT exceeds the projected run limit",
                )
            })?;
        Ok(())
    }

    fn bump_projected_text(&mut self, amount: usize) -> Result<()> {
        self.projected_text = self
            .projected_text
            .checked_add(amount)
            .filter(|value| *value <= self.limits.retained_text)
            .ok_or_else(|| {
                odt_error(
                    Some("content.xml"),
                    0,
                    "ODT exceeds the projected text limit",
                )
            })?;
        Ok(())
    }
}

#[derive(Clone)]
enum InlinePiece {
    Text(String, EffectiveStyle),
    Break,
    Tab,
    Image {
        relationship: String,
        width: Length,
        height: Length,
    },
}

#[derive(Default)]
struct InlineWhitespace {
    pending: bool,
    emitted: bool,
    style: Option<EffectiveStyle>,
}

fn collapse_odt_text(
    text: &str,
    style: &EffectiveStyle,
    whitespace: &mut InlineWhitespace,
    output: &mut Vec<InlinePiece>,
) {
    let mut visible = String::new();
    for character in text.chars() {
        if character.is_whitespace() {
            if !visible.is_empty() {
                push_inline_text(output, std::mem::take(&mut visible), style.clone());
            }
            whitespace.pending = true;
            whitespace.style.get_or_insert_with(|| style.clone());
        } else {
            if whitespace.pending && whitespace.emitted {
                push_inline_text(
                    output,
                    " ".to_string(),
                    whitespace.style.take().unwrap_or_else(|| style.clone()),
                );
            }
            visible.push(character);
            whitespace.pending = false;
            whitespace.style = None;
            whitespace.emitted = true;
        }
    }
    if !visible.is_empty() {
        push_inline_text(output, visible, style.clone());
    }
}

fn push_inline_text(output: &mut Vec<InlinePiece>, text: String, style: EffectiveStyle) {
    if let Some(InlinePiece::Text(previous, previous_style)) = output.last_mut()
        && *previous_style == style
    {
        previous.push_str(&text);
    } else {
        output.push(InlinePiece::Text(text, style));
    }
}

struct TableCellModel {
    colspan: usize,
    rowspan: usize,
    paragraphs: Vec<CT_P>,
}

#[derive(Clone)]
enum StyleChange {
    Font(String),
    Size(f64),
    Bold(bool),
    Italic(bool),
    Underline(bool),
    Strike(bool),
    Color(String),
    Background(String),
    Vertical(VerticalText),
    Alignment(Alignment),
    Before(Length),
    After(Length),
    Left(Length),
    Right(Length),
    First(Length),
    Line(LineHeight),
}

impl StyleChange {
    fn apply(self, style: &mut EffectiveStyle) {
        match self {
            Self::Font(value) => style.font = Some(value),
            Self::Size(value) => style.size = Some(value),
            Self::Bold(value) => style.bold = Some(value),
            Self::Italic(value) => style.italic = Some(value),
            Self::Underline(value) => style.underline = Some(value),
            Self::Strike(value) => style.strike = Some(value),
            Self::Color(value) => style.color = Some(value),
            Self::Background(value) => style.background = Some(value),
            Self::Vertical(value) => style.vertical = Some(value),
            Self::Alignment(value) => style.alignment = Some(value),
            Self::Before(value) => style.before = Some(value),
            Self::After(value) => style.after = Some(value),
            Self::Left(value) => style.left = Some(value),
            Self::Right(value) => style.right = Some(value),
            Self::First(value) => style.first = Some(value),
            Self::Line(value) => style.line = Some(value),
        }
    }
}

fn parse_text_property(
    attribute: &XmlAttribute,
    fonts: &HashMap<String, String>,
) -> Option<std::result::Result<StyleChange, String>> {
    let namespace = attribute.name.namespace.as_deref();
    let name = attribute.name.local.as_str();
    let value = attribute.value.trim();
    let lower = value.to_ascii_lowercase();
    let result = match (namespace, name) {
        (Some(STYLE_NS), "font-name") => Ok(StyleChange::Font(
            fonts
                .get(value)
                .cloned()
                .unwrap_or_else(|| value.to_string()),
        )),
        (Some(FO_NS), "font-family") => Ok(StyleChange::Font(
            value.trim_matches(['\'', '"']).to_string(),
        )),
        (Some(FO_NS), "font-size") => {
            parse_positive_length(value).map(|length| StyleChange::Size(length.to_pt()))
        }
        (Some(FO_NS), "font-weight") => match lower.as_str() {
            "bold" | "600" | "700" | "800" | "900" => Ok(StyleChange::Bold(true)),
            "normal" | "400" | "500" => Ok(StyleChange::Bold(false)),
            _ => Err(format!("unsupported font weight {value}")),
        },
        (Some(FO_NS), "font-style") => match lower.as_str() {
            "italic" | "oblique" => Ok(StyleChange::Italic(true)),
            "normal" => Ok(StyleChange::Italic(false)),
            _ => Err(format!("unsupported font style {value}")),
        },
        (Some(STYLE_NS), "text-underline-style") => Ok(StyleChange::Underline(lower != "none")),
        (Some(STYLE_NS), "text-line-through-style") => Ok(StyleChange::Strike(lower != "none")),
        (Some(FO_NS), "color") => parse_color(value).map(StyleChange::Color),
        (Some(FO_NS), "background-color") => parse_color(value).map(StyleChange::Background),
        (Some(STYLE_NS), "text-position") => {
            if lower.starts_with("super") {
                Ok(StyleChange::Vertical(VerticalText::Superscript))
            } else if lower.starts_with("sub") {
                Ok(StyleChange::Vertical(VerticalText::Subscript))
            } else {
                Err(format!("unsupported text position {value}"))
            }
        }
        _ => return None,
    };
    Some(result)
}

fn parse_paragraph_property(
    attribute: &XmlAttribute,
) -> Option<std::result::Result<StyleChange, String>> {
    let namespace = attribute.name.namespace.as_deref();
    let name = attribute.name.local.as_str();
    let value = attribute.value.trim();
    let lower = value.to_ascii_lowercase();
    let result = match (namespace, name) {
        (Some(FO_NS), "text-align") => match lower.as_str() {
            "start" | "left" => Ok(StyleChange::Alignment(Alignment::Left)),
            "center" => Ok(StyleChange::Alignment(Alignment::Center)),
            "end" | "right" => Ok(StyleChange::Alignment(Alignment::Right)),
            "justify" => Ok(StyleChange::Alignment(Alignment::Justify)),
            _ => Err(format!("unsupported paragraph alignment {value}")),
        },
        (Some(FO_NS), "margin-top") => parse_length(value, false).map(StyleChange::Before),
        (Some(FO_NS), "margin-bottom") => parse_length(value, false).map(StyleChange::After),
        (Some(FO_NS), "margin-left") => parse_length(value, false).map(StyleChange::Left),
        (Some(FO_NS), "margin-right") => parse_length(value, false).map(StyleChange::Right),
        (Some(FO_NS), "text-indent") => parse_length(value, true).map(StyleChange::First),
        (Some(FO_NS), "line-height") if lower.ends_with('%') => {
            match lower.trim_end_matches('%').parse::<f64>() {
                Ok(percentage)
                    if percentage.is_finite() && percentage > 0.0 && percentage <= 10_000.0 =>
                {
                    Ok(StyleChange::Line(LineHeight::Multiple(percentage / 100.0)))
                }
                Ok(_) => Err(format!(
                    "percentage line height is outside the supported range {value}"
                )),
                Err(_) => Err(format!("invalid percentage line height {value}")),
            }
        }
        (Some(FO_NS), "line-height") => parse_positive_length(value)
            .map(|length| StyleChange::Line(LineHeight::Exact(length.to_pt()))),
        _ => return None,
    };
    Some(result)
}

fn is_ignorable_style_attribute(attribute: &XmlAttribute) -> bool {
    matches!(
        (
            attribute.name.namespace.as_deref(),
            attribute.name.local.as_str()
        ),
        (Some(STYLE_NS), "font-name-asian" | "font-name-complex")
            | (Some(STYLE_NS), "font-size-asian" | "font-size-complex")
            | (Some(STYLE_NS), "font-weight-asian" | "font-weight-complex")
            | (Some(STYLE_NS), "font-style-asian" | "font-style-complex")
            | (Some(STYLE_NS), "writing-mode")
    )
}

fn apply_paragraph_style(paragraph: &mut Paragraph<'_>, style: &EffectiveStyle) {
    if let Some(value) = style.alignment {
        paragraph.set_alignment(value);
    }
    if let Some(value) = style.before {
        paragraph.set_space_before(value);
    }
    if let Some(value) = style.after {
        paragraph.set_space_after(value);
    }
    if let Some(value) = style.left {
        paragraph.set_indent_left(value);
    }
    if let Some(value) = style.right {
        paragraph.set_indent_right(value);
    }
    if let Some(value) = style.first {
        paragraph.set_signed_first_line_indent_value(Some(value));
    }
    if let Some(value) = style.line {
        match value {
            LineHeight::Exact(points) => paragraph.set_line_spacing(points),
            LineHeight::Multiple(multiple) => paragraph.set_line_spacing_multiple(multiple),
        }
    }
}

fn apply_run_style(run: &mut Run<'_>, style: &EffectiveStyle) {
    if let Some(value) = &style.font {
        run.set_font(value);
    }
    if let Some(value) = style.size {
        run.set_size(value);
    }
    if let Some(value) = style.bold {
        run.set_bold(value);
    }
    if let Some(value) = style.italic {
        run.set_italic(value);
    }
    if let Some(value) = style.underline {
        run.set_underline(value);
    }
    if let Some(value) = style.strike {
        run.set_strike(value);
    }
    if let Some(value) = &style.color {
        run.set_color(value);
    }
    if let Some(value) = &style.background {
        run.set_highlight(value);
    }
    match style.vertical {
        Some(VerticalText::Superscript) => run.set_superscript(),
        Some(VerticalText::Subscript) => run.set_subscript(),
        None => {}
    }
}

fn parse_length(value: &str, signed: bool) -> std::result::Result<Length, String> {
    let value = value.trim().to_ascii_lowercase();
    let units = [
        ("cm", 72.0 / 2.54),
        ("mm", 72.0 / 25.4),
        ("in", 72.0),
        ("pt", 1.0),
        ("pc", 12.0),
        ("px", 0.75),
    ];
    let (number, factor) = units
        .iter()
        .find_map(|(unit, factor)| value.strip_suffix(unit).map(|number| (number, *factor)))
        .ok_or_else(|| format!("unsupported ODT length {value}"))?;
    let points = number
        .trim()
        .parse::<f64>()
        .map_err(|_| format!("invalid ODT length {value}"))?
        * factor;
    if !points.is_finite() || points.abs() > 1_000_000.0 || (!signed && points < 0.0) {
        return Err(format!("ODT length is outside the supported range {value}"));
    }
    Ok(Length::pt(points))
}

fn parse_positive_length(value: &str) -> std::result::Result<Length, String> {
    let length = parse_length(value, false)?;
    if length.to_emu() <= 0 {
        return Err(format!("ODT length must be positive {value}"));
    }
    Ok(length)
}

fn parse_color(value: &str) -> std::result::Result<String, String> {
    let value = value.trim();
    if value.len() == 7
        && value.starts_with('#')
        && value[1..].bytes().all(|byte| byte.is_ascii_hexdigit())
    {
        Ok(value[1..].to_ascii_uppercase())
    } else {
        Err(format!("unsupported ODT color {value}"))
    }
}

fn repeated_count(node: &XmlNode, namespace: &str, local: &str, maximum: usize) -> Result<usize> {
    let Some(value) = node.attr(Some(namespace), local) else {
        return Ok(1);
    };
    let value = value.parse::<usize>().map_err(|_| {
        odt_error(
            Some("content.xml"),
            0,
            format!("invalid ODT repeat or span value {value}"),
        )
    })?;
    if value == 0 || value > maximum {
        return Err(odt_error(
            Some("content.xml"),
            0,
            "ODT repeat or span exceeds its bound",
        ));
    }
    Ok(value)
}

fn collect_table_rows<'a>(node: &'a XmlNode, rows: &mut Vec<&'a XmlNode>) {
    for child in node.elements() {
        if child.is(TABLE_NS, "table-row") {
            rows.push(child);
        } else if child.is(TABLE_NS, "table-rows") || child.is(TABLE_NS, "table-header-rows") {
            collect_table_rows(child, rows);
        }
    }
}

fn safe_image_target(href: &str) -> Option<String> {
    if href.is_empty()
        || href.starts_with('/')
        || href.starts_with('#')
        || href.contains(':')
        || href.contains('\\')
        || href.contains('\0')
        || href
            .split('/')
            .any(|part| part.is_empty() || part == "." || part == "..")
    {
        None
    } else {
        Some(href.to_string())
    }
}

fn required_attr<'a>(
    node: &'a XmlNode,
    namespace: &str,
    local: &str,
    part: &str,
) -> Result<&'a str> {
    node.attr(Some(namespace), local).ok_or_else(|| {
        odt_error(
            Some(part),
            0,
            format!("{} requires {local}", display_name(node)),
        )
    })
}

fn display_name(node: &XmlNode) -> String {
    let prefix = match node.name.namespace.as_deref() {
        Some(OFFICE_NS) => "office",
        Some(TEXT_NS) => "text",
        Some(STYLE_NS) => "style",
        Some(TABLE_NS) => "table",
        Some(DRAW_NS) => "draw",
        Some(MANIFEST_NS) => "manifest",
        Some(_) => "foreign",
        None => "unbound",
    };
    format!("{prefix}:{}", node.name.local)
}

fn odt_error(part: Option<&str>, offset: u64, message: impl Into<String>) -> Error {
    Error::Odt {
        part: part.map(str::to_string),
        offset,
        message: message.into(),
    }
}

#[cfg(test)]
mod tests {
    use std::io::{Cursor, Write};

    use zip::write::SimpleFileOptions;
    use zip::{CompressionMethod, ZipWriter};

    use super::*;

    const PNG: &[u8] = &[
        0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44,
        0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00, 0x90,
        0x77, 0x53, 0xde, 0x00, 0x00, 0x00, 0x0c, 0x49, 0x44, 0x41, 0x54, 0x08, 0xd7, 0x63, 0xf8,
        0xcf, 0xc0, 0x00, 0x00, 0x00, 0x02, 0x00, 0x01, 0xe2, 0x21, 0xbc, 0x33, 0x00, 0x00, 0x00,
        0x00, 0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
    ];

    fn package_with(entries: &[(&str, &[u8])]) -> Vec<u8> {
        let mut output = Cursor::new(Vec::new());
        let mut archive = ZipWriter::new(&mut output);
        for (name, bytes) in entries {
            let method = if *name == "mimetype" {
                CompressionMethod::Stored
            } else {
                CompressionMethod::Deflated
            };
            archive
                .start_file(
                    *name,
                    SimpleFileOptions::default().compression_method(method),
                )
                .unwrap();
            archive.write_all(bytes).unwrap();
        }
        archive.finish().unwrap();
        output.into_inner()
    }

    fn package(content: &str, styles: Option<&str>, extra: &[(&str, &[u8])]) -> Vec<u8> {
        let mut entries = vec![
            ("mimetype", ODT_MIMETYPE),
            ("content.xml", content.as_bytes()),
        ];
        if let Some(styles) = styles {
            entries.push(("styles.xml", styles.as_bytes()));
        }
        entries.extend_from_slice(extra);
        package_with(&entries)
    }

    fn content(body: &str, automatic: &str) -> String {
        format!(
            r#"<?xml version="1.0"?><office:document-content xmlns:office="{OFFICE_NS}" xmlns:text="{TEXT_NS}" xmlns:style="{STYLE_NS}" xmlns:fo="{FO_NS}" xmlns:table="{TABLE_NS}" xmlns:draw="{DRAW_NS}" xmlns:xlink="{XLINK_NS}" xmlns:svg="{SVG_NS}"><office:automatic-styles>{automatic}</office:automatic-styles><office:body><office:text>{body}</office:text></office:body></office:document-content>"#
        )
    }

    fn assert_emitted_content_subset(node: &XmlNode) {
        let children = node.elements().collect::<Vec<_>>();
        let has_text = node
            .children
            .iter()
            .any(|child| matches!(child, XmlChild::Text(text) if !text.is_empty()));
        if has_text {
            assert!(node.is(TEXT_NS, "p") || node.is(TEXT_NS, "h") || node.is(TEXT_NS, "span"));
        }
        match (node.name.namespace.as_deref(), node.name.local.as_str()) {
            (Some(OFFICE_NS), "document-content") => {
                assert_eq!(node.attr(Some(OFFICE_NS), "version"), Some("1.3"));
                assert_eq!(children.len(), 2);
                assert!(children[0].is(OFFICE_NS, "automatic-styles"));
                assert!(children[1].is(OFFICE_NS, "body"));
            }
            (Some(OFFICE_NS), "automatic-styles") => {
                assert!(children.iter().all(|child| {
                    child.is(STYLE_NS, "style") || child.is(TEXT_NS, "list-style")
                }))
            }
            (Some(STYLE_NS), "style") => {
                assert!(node.attr(Some(STYLE_NS), "name").is_some());
                let family = node.attr(Some(STYLE_NS), "family").unwrap();
                assert!(matches!(family, "paragraph" | "text"));
                assert_eq!(children.len(), 1);
                assert!(
                    children[0].is(STYLE_NS, "paragraph-properties")
                        || children[0].is(STYLE_NS, "text-properties")
                );
            }
            (Some(STYLE_NS), "paragraph-properties" | "text-properties") => {
                assert!(children.is_empty());
            }
            (Some(TEXT_NS), "list-style") => {
                assert!(node.attr(Some(STYLE_NS), "name").is_some());
                assert_eq!(children.len(), 9);
                assert!(children.iter().all(|child| {
                    child.is(TEXT_NS, "list-level-style-bullet")
                        || child.is(TEXT_NS, "list-level-style-number")
                }));
            }
            (Some(TEXT_NS), "list-level-style-bullet") => {
                assert!(
                    (1..=9).contains(
                        &node
                            .attr(Some(TEXT_NS), "level")
                            .unwrap()
                            .parse::<u32>()
                            .unwrap()
                    )
                );
                assert_eq!(node.attr(Some(TEXT_NS), "bullet-char"), Some("•"));
                assert!(children.is_empty());
            }
            (Some(TEXT_NS), "list-level-style-number") => {
                assert!(
                    (1..=9).contains(
                        &node
                            .attr(Some(TEXT_NS), "level")
                            .unwrap()
                            .parse::<u32>()
                            .unwrap()
                    )
                );
                assert_eq!(node.attr(Some(STYLE_NS), "num-format"), Some("1"));
                assert!(children.is_empty());
            }
            (Some(OFFICE_NS), "body") => {
                assert_eq!(children.len(), 1);
                assert!(children[0].is(OFFICE_NS, "text"));
            }
            (Some(OFFICE_NS), "text") => assert!(children.iter().all(|child| {
                child.is(TEXT_NS, "p")
                    || child.is(TEXT_NS, "h")
                    || child.is(TEXT_NS, "list")
                    || child.is(TABLE_NS, "table")
            })),
            (Some(TEXT_NS), "p" | "h") => {
                if node.is(TEXT_NS, "h") {
                    assert!(
                        (1..=9).contains(
                            &node
                                .attr(Some(TEXT_NS), "outline-level")
                                .unwrap()
                                .parse::<u32>()
                                .unwrap()
                        )
                    );
                }
                assert!(children.iter().all(|child| {
                    child.is(TEXT_NS, "span")
                        || child.is(TEXT_NS, "s")
                        || child.is(TEXT_NS, "tab")
                        || child.is(TEXT_NS, "line-break")
                        || child.is(DRAW_NS, "frame")
                }));
            }
            (Some(TEXT_NS), "span") => {
                assert!(node.attr(Some(TEXT_NS), "style-name").is_some());
                assert!(children.iter().all(|child| {
                    child.is(TEXT_NS, "s")
                        || child.is(TEXT_NS, "tab")
                        || child.is(TEXT_NS, "line-break")
                        || child.is(DRAW_NS, "frame")
                }));
            }
            (Some(TEXT_NS), "s") => {
                if let Some(count) = node.attr(Some(TEXT_NS), "c") {
                    assert!(count.parse::<usize>().unwrap() > 0);
                }
                assert!(children.is_empty());
            }
            (Some(TEXT_NS), "tab" | "line-break") => assert!(children.is_empty()),
            (Some(TEXT_NS), "list") => {
                assert!(node.attr(Some(TEXT_NS), "style-name").is_some());
                assert!(!children.is_empty());
                assert!(children.iter().all(|child| child.is(TEXT_NS, "list-item")));
            }
            (Some(TEXT_NS), "list-item") => {
                assert!(!children.is_empty());
                assert!(children.iter().all(|child| {
                    child.is(TEXT_NS, "p") || child.is(TEXT_NS, "h") || child.is(TEXT_NS, "list")
                }));
            }
            (Some(TABLE_NS), "table") => {
                assert_eq!(node.attr(Some(TABLE_NS), "name"), Some("Table"));
                assert!(!children.is_empty());
                assert!(children[0].is(TABLE_NS, "table-column"));
                assert!(
                    children[1..]
                        .iter()
                        .all(|child| child.is(TABLE_NS, "table-row"))
                );
                assert_emitted_table_spans(node);
            }
            (Some(TABLE_NS), "table-column") => {
                assert!(
                    node.attr(Some(TABLE_NS), "number-columns-repeated")
                        .unwrap()
                        .parse::<usize>()
                        .unwrap()
                        > 0
                );
                assert!(children.is_empty());
            }
            (Some(TABLE_NS), "table-row") => assert!(children.iter().all(|child| {
                child.is(TABLE_NS, "table-cell") || child.is(TABLE_NS, "covered-table-cell")
            })),
            (Some(TABLE_NS), "table-cell") => {
                for attribute in ["number-columns-spanned", "number-rows-spanned"] {
                    if let Some(value) = node.attr(Some(TABLE_NS), attribute) {
                        assert!(value.parse::<usize>().unwrap() > 1);
                    }
                }
                assert!(children.iter().all(|child| {
                    child.is(TEXT_NS, "p") || child.is(TEXT_NS, "h") || child.is(TEXT_NS, "list")
                }));
            }
            (Some(TABLE_NS), "covered-table-cell") => assert!(children.is_empty()),
            (Some(DRAW_NS), "frame") => {
                assert!(node.attr(Some(DRAW_NS), "name").is_some());
                assert_eq!(node.attr(Some(TEXT_NS), "anchor-type"), Some("as-char"));
                assert!(parse_positive_length(node.attr(Some(SVG_NS), "width").unwrap()).is_ok());
                assert!(parse_positive_length(node.attr(Some(SVG_NS), "height").unwrap()).is_ok());
                assert_eq!(children.len(), 1);
                assert!(children[0].is(DRAW_NS, "image"));
            }
            (Some(DRAW_NS), "image") => {
                assert!(
                    node.attr(Some(XLINK_NS), "href")
                        .is_some_and(|href| href.starts_with("Pictures/"))
                );
                assert_eq!(node.attr(Some(XLINK_NS), "type"), Some("simple"));
                assert_eq!(node.attr(Some(XLINK_NS), "show"), Some("embed"));
                assert_eq!(node.attr(Some(XLINK_NS), "actuate"), Some("onLoad"));
                assert!(children.is_empty());
            }
            _ => panic!("writer emitted unsupported element {}", display_name(node)),
        }
        for child in children {
            assert_emitted_content_subset(child);
        }
    }

    fn assert_emitted_list_continuations(node: &XmlNode) {
        let mut seen = HashSet::new();
        for child in node.elements() {
            if child.is(TEXT_NS, "list") {
                let style = child.attr(Some(TEXT_NS), "style-name").unwrap();
                if seen.insert(style) {
                    assert_eq!(child.attr(Some(TEXT_NS), "continue-numbering"), None);
                } else {
                    assert_eq!(
                        child.attr(Some(TEXT_NS), "continue-numbering"),
                        Some("true")
                    );
                }
            }
        }
        for child in node.elements() {
            assert_emitted_list_continuations(child);
        }
    }

    fn assert_emitted_table_spans(table: &XmlNode) {
        let columns = table
            .elements()
            .next()
            .unwrap()
            .attr(Some(TABLE_NS), "number-columns-repeated")
            .unwrap()
            .parse::<usize>()
            .unwrap();
        let mut vertical = vec![0_usize; columns];
        for row in table
            .elements()
            .filter(|child| child.is(TABLE_NS, "table-row"))
        {
            let mut column = 0_usize;
            let mut horizontal = 0_usize;
            for cell in row.elements() {
                assert!(column < columns);
                if cell.is(TABLE_NS, "covered-table-cell") {
                    assert!(horizontal > 0 || vertical[column] > 0);
                    horizontal = horizontal.saturating_sub(1);
                    column += 1;
                    continue;
                }
                assert!(cell.is(TABLE_NS, "table-cell"));
                assert_eq!(horizontal, 0);
                assert_eq!(
                    vertical[column], 0,
                    "ordinary cell overlaps a vertical span"
                );
                let colspan = cell
                    .attr(Some(TABLE_NS), "number-columns-spanned")
                    .map(str::parse::<usize>)
                    .transpose()
                    .unwrap()
                    .unwrap_or(1);
                let rowspan = cell
                    .attr(Some(TABLE_NS), "number-rows-spanned")
                    .map(str::parse::<usize>)
                    .transpose()
                    .unwrap()
                    .unwrap_or(1);
                assert!(column + colspan <= columns);
                if rowspan > 1 {
                    for slot in &mut vertical[column..column + colspan] {
                        assert_eq!(*slot, 0);
                        *slot = rowspan;
                    }
                }
                horizontal = colspan - 1;
                column += 1;
            }
            assert_eq!(column, columns);
            assert_eq!(horizontal, 0);
            for slot in &mut vertical {
                *slot = slot.saturating_sub(1);
            }
        }
        assert!(vertical.into_iter().all(|remaining| remaining == 0));
    }

    #[test]
    fn odt_archive_rejects_unsafe_duplicate_encrypted_and_oversized_entries() {
        let safe = package(&content("<text:p>x</text:p>", ""), None, &[]);
        assert!(Document::from_odt_bytes(&safe).is_ok());

        for name in [
            "../content.xml",
            "/content.xml",
            "bad\\name",
            "./content.xml",
            "bad\0name",
        ] {
            let unsafe_package = package_with(&[("mimetype", ODT_MIMETYPE), (name, b"x")]);
            assert!(
                Document::from_odt_bytes(&unsafe_package).is_err(),
                "accepted {name}"
            );
        }
        let mut duplicate = package_with(&[
            ("mimetype", ODT_MIMETYPE),
            ("content.xml", content("<text:p>x</text:p>", "").as_bytes()),
            ("dontent.xml", content("<text:p>y</text:p>", "").as_bytes()),
        ]);
        for offset in 0..duplicate.len().saturating_sub(b"dontent.xml".len()) {
            if &duplicate[offset..offset + b"dontent.xml".len()] == b"dontent.xml" {
                duplicate[offset..offset + b"content.xml".len()].copy_from_slice(b"content.xml");
            }
        }
        assert!(!duplicate.windows(11).any(|window| window == b"dontent.xml"));
        assert!(Document::from_odt_bytes(&duplicate).is_err());

        let mut encrypted = safe.clone();
        for offset in 0..encrypted.len().saturating_sub(10) {
            if &encrypted[offset..offset + 4] == b"PK\x03\x04" {
                encrypted[offset + 6] |= 1;
            } else if &encrypted[offset..offset + 4] == b"PK\x01\x02" {
                encrypted[offset + 8] |= 1;
            }
        }
        assert!(Document::from_odt_bytes(&encrypted).is_err());

        let mut unsupported_compression = safe.clone();
        for offset in 0..unsupported_compression.len().saturating_sub(12) {
            if &unsupported_compression[offset..offset + 4] == b"PK\x03\x04" {
                unsupported_compression[offset + 8..offset + 10]
                    .copy_from_slice(&12_u16.to_le_bytes());
            } else if &unsupported_compression[offset..offset + 4] == b"PK\x01\x02" {
                unsupported_compression[offset + 10..offset + 12]
                    .copy_from_slice(&12_u16.to_le_bytes());
            }
        }
        assert!(Document::from_odt_bytes(&unsupported_compression).is_err());

        let mut directory = Cursor::new(Vec::new());
        let mut directory_archive = ZipWriter::new(&mut directory);
        directory_archive
            .add_directory("folder/", SimpleFileOptions::default())
            .unwrap();
        directory_archive.finish().unwrap();
        assert!(Document::from_odt_bytes(directory.get_ref()).is_err());

        let manifest = format!(
            r#"<manifest:manifest xmlns:manifest="{MANIFEST_NS}"><manifest:file-entry manifest:full-path="Pictures/pixel.png"><manifest:encryption-data/></manifest:file-entry></manifest:manifest>"#
        );
        let image_body = content(
            r#"<text:p><draw:frame svg:width="1in" svg:height="1in"><draw:image xlink:href="Pictures/pixel.png"/></draw:frame></text:p>"#,
            "",
        );
        let manifest_encrypted = package(
            &image_body,
            None,
            &[
                ("META-INF/manifest.xml", manifest.as_bytes()),
                ("Pictures/pixel.png", PNG),
            ],
        );
        assert!(Document::from_odt_bytes(&manifest_encrypted).is_err());

        let encrypted_content_manifest = format!(
            r#"<manifest:manifest xmlns:manifest="{MANIFEST_NS}"><manifest:file-entry manifest:full-path="content.xml"><manifest:encryption-data/></manifest:file-entry></manifest:manifest>"#
        );
        let encrypted_content = package(
            &content("<text:p>x</text:p>", ""),
            None,
            &[(
                "META-INF/manifest.xml",
                encrypted_content_manifest.as_bytes(),
            )],
        );
        assert!(Document::from_odt_bytes(&encrypted_content).is_err());

        let entry_limits = PackageReadLimits {
            max_entries: 1,
            max_part_uncompressed_bytes: u64::MAX,
            max_total_uncompressed_bytes: u64::MAX,
        };
        assert!(Document::from_odt_bytes_with_limits(&safe, entry_limits).is_err());

        let part_limits = PackageReadLimits {
            max_entries: usize::MAX,
            max_part_uncompressed_bytes: 16,
            max_total_uncompressed_bytes: u64::MAX,
        };
        assert!(Document::from_odt_bytes_with_limits(&safe, part_limits).is_err());

        let total_limits = PackageReadLimits {
            max_entries: usize::MAX,
            max_part_uncompressed_bytes: u64::MAX,
            max_total_uncompressed_bytes: 16,
        };
        assert!(Document::from_odt_bytes_with_limits(&safe, total_limits).is_err());
    }

    #[test]
    fn odt_styles_resolve_defaults_parents_and_automatic_overrides() {
        let styles = format!(
            r##"<o:document-styles xmlns:o="{OFFICE_NS}" xmlns:s="{STYLE_NS}" xmlns:f="{FO_NS}" xmlns:v="{SVG_NS}"><o:font-face-decls><s:font-face s:name="Alias" v:font-family="'Liberation Serif'"/></o:font-face-decls><o:styles><s:default-style s:family="paragraph"><s:text-properties f:font-size="10pt"/></s:default-style><s:style s:name="Parent" s:family="paragraph"><s:text-properties s:font-name="Alias" f:font-weight="bold" f:font-style="italic" s:text-underline-style="solid" s:text-line-through-style="solid" f:color="#123456" f:background-color="#FEDCBA" s:text-position="super"/></s:style><s:style s:name="Child" s:family="paragraph" s:parent-style-name="Parent"><s:paragraph-properties f:text-align="center" f:margin-top="3pt" f:margin-bottom="4pt" f:margin-left="5pt" f:margin-right="6pt" f:text-indent="-2pt" f:line-height="150%"/></s:style></o:styles></o:document-styles>"##
        );
        let input = content(r#"<text:p text:style-name="Child">styled</text:p>"#, "");
        let parsed = Document::from_odt_bytes(&package(&input, Some(&styles), &[])).unwrap();
        let paragraph = parsed.document.paragraph(0).unwrap();
        assert_eq!(paragraph.alignment(), Some(Alignment::Center));
        let run = paragraph.run(0).unwrap();
        assert_eq!(run.bold_value(), Some(true));
        assert_eq!(run.italic_value(), Some(true));
        assert_eq!(run.underline_code_value(), Some(1));
        assert_eq!(run.strike_value(), Some(true));
        assert_eq!(run.size(), Some(10.0));
        assert_eq!(run.font_name(), Some("Liberation Serif"));
        assert_eq!(run.color(), Some("123456"));
        assert_eq!(run.highlight().as_deref(), Some("FEDCBA"));
        assert_eq!(run.vert_align(), Some("superscript"));
        assert_eq!(paragraph.space_before(), Some(Length::pt(3.0)));
        assert_eq!(paragraph.space_after(), Some(Length::pt(4.0)));
        assert_eq!(paragraph.indent_left(), Some(Length::pt(5.0)));
        assert_eq!(paragraph.indent_right(), Some(Length::pt(6.0)));
        assert_eq!(paragraph.first_line_indent(), Some(Length::pt(-2.0)));
        assert_eq!(paragraph.line_spacing_multiple(), Some(1.5));

        let cycle = format!(
            r#"<office:document-styles xmlns:office="{OFFICE_NS}" xmlns:style="{STYLE_NS}"><office:styles><style:style style:name="A" style:family="paragraph" style:parent-style-name="B"/><style:style style:name="B" style:family="paragraph" style:parent-style-name="A"/></office:styles></office:document-styles>"#
        );
        assert!(Document::from_odt_bytes(&package(&input, Some(&cycle), &[])).is_err());

        let missing_parent = format!(
            r#"<office:document-styles xmlns:office="{OFFICE_NS}" xmlns:style="{STYLE_NS}"><office:styles><style:style style:name="Child" style:family="paragraph" style:parent-style-name="Missing"/></office:styles></office:document-styles>"#
        );
        let parsed =
            Document::from_odt_bytes(&package(&input, Some(&missing_parent), &[])).unwrap();
        assert!(
            parsed
                .diagnostics
                .iter()
                .any(|diagnostic| diagnostic.message.contains("missing ODT style Missing"))
        );
    }

    #[test]
    fn odt_reader_rejects_malformed_or_unbounded_xml() {
        let malformed = package("<!DOCTYPE x><x/>", None, &[]);
        assert!(Document::from_odt_bytes(&malformed).is_err());
        let deep = format!(
            "{}x{}",
            "<text:span>".repeat(257),
            "</text:span>".repeat(257)
        );
        let input = content(&format!("<text:p>{deep}</text:p>"), "");
        assert!(Document::from_odt_bytes(&package(&input, None, &[])).is_err());
        let repeated = content(
            r#"<table:table><table:table-row table:number-rows-repeated="10001"><table:table-cell/></table:table-row></table:table>"#,
            "",
        );
        assert!(Document::from_odt_bytes(&package(&repeated, None, &[])).is_err());

        let unresolved = content("<text:p>&missing;</text:p>", "");
        assert!(Document::from_odt_bytes(&package(&unresolved, None, &[])).is_err());
        let unknown_namespace = content("<text:p><bad:span>x</bad:span></text:p>", "");
        assert!(Document::from_odt_bytes(&package(&unknown_namespace, None, &[])).is_err());

        let shallow_limits = OdtLimits {
            xml_depth: 1,
            ..OdtLimits::DEFAULT
        };
        assert!(parse_xml("content.xml", b"<root><empty/></root>", shallow_limits).is_err());

        let bad_styles = "<root/>";
        let input = content("<text:p>x</text:p>", "");
        assert!(Document::from_odt_bytes(&package(&input, Some(bad_styles), &[])).is_err());

        for (body, limits) in [
            (
                "<text:p>x</text:p>",
                OdtLimits {
                    blocks: 0,
                    ..OdtLimits::DEFAULT
                },
            ),
            (
                "<text:p>x</text:p>",
                OdtLimits {
                    runs: 0,
                    ..OdtLimits::DEFAULT
                },
            ),
            (
                r#"<text:p><text:s text:c="3"/><text:s text:c="3"/></text:p>"#,
                OdtLimits {
                    retained_text: 5,
                    ..OdtLimits::DEFAULT
                },
            ),
            (
                r#"<table:table><table:table-column table:number-columns-repeated="2"/><table:table-row><table:table-cell/></table:table-row></table:table>"#,
                OdtLimits {
                    columns: 1,
                    ..OdtLimits::DEFAULT
                },
            ),
            (
                r#"<table:table><table:table-row><table:table-cell table:number-columns-repeated="2"/></table:table-row></table:table>"#,
                OdtLimits {
                    cells: 1,
                    ..OdtLimits::DEFAULT
                },
            ),
            (
                "<text:p>before<office:annotation/>after</text:p>",
                OdtLimits {
                    diagnostics: 0,
                    ..OdtLimits::DEFAULT
                },
            ),
        ] {
            let input = content(body, "");
            let package = package(&input, None, &[]);
            assert!(
                from_odt_with_limits(&package, limits).is_err(),
                "accepted {body}"
            );
        }
    }

    #[test]
    fn unsupported_odt_content_is_diagnosed_without_dropping_supported_siblings() {
        let input = content(
            r#"<text:p>before<office:annotation>drop</office:annotation>after</text:p><text:list><text:list-item><text:p>item</text:p></text:list-item></text:list><table:table><table:table-column/><table:table-row><table:table-cell><text:p>cell</text:p></table:table-cell></table:table-row></table:table><text:p><draw:frame svg:width="1in" svg:height="1in"><draw:image xlink:href="Pictures/pixel.png"/></draw:frame></text:p>"#,
            "",
        );
        let parsed =
            Document::from_odt_bytes(&package(&input, None, &[("Pictures/pixel.png", PNG)]))
                .unwrap();
        assert_eq!(parsed.document.paragraph(0).unwrap().text(), "beforeafter");
        assert_eq!(parsed.document.paragraph(1).unwrap().text(), "item");
        assert_eq!(parsed.document.table_count(), 1);
        assert_eq!(parsed.document.images().len(), 1);
        assert_eq!(parsed.diagnostics.len(), 1);
        assert_eq!(
            parsed.diagnostics[0].path,
            "office:body/office:text/text:p[1]/office:annotation[1]"
        );
    }

    #[test]
    fn odt_reader_projects_text_formatting_lists_tables_and_images() {
        let automatic = r##"<style:style style:name="P1" style:family="paragraph"><style:paragraph-properties fo:text-align="right"/></style:style><style:style style:name="T1" style:family="text"><style:text-properties fo:font-weight="bold" fo:font-style="italic" fo:color="#123456"/></style:style><text:list-style style:name="L1"><text:list-level-style-number text:level="1"/></text:list-style>"##;
        let body = r#"<text:p text:style-name="P1">A <text:span text:style-name="T1">formatted</text:span><text:tab/>line<text:line-break/>tail<draw:frame svg:width="1in" svg:height="0.5in"><draw:image xlink:href="Pictures/pixel.png"/></draw:frame></text:p><text:list text:style-name="L1"><text:list-item><text:p>one</text:p></text:list-item></text:list><table:table><table:table-column table:number-columns-repeated="2"/><table:table-row><table:table-cell table:number-columns-spanned="2"><text:p>wide</text:p><text:p>second</text:p></table:table-cell><table:covered-table-cell/></table:table-row><table:table-row><table:table-cell table:number-rows-spanned="2"><text:p>vertical</text:p></table:table-cell><table:table-cell><text:p>top</text:p></table:table-cell></table:table-row><table:table-row><table:covered-table-cell/><table:table-cell><text:p>bottom</text:p></table:table-cell></table:table-row></table:table>"#;
        let input = content(body, automatic);
        let parsed =
            Document::from_odt_bytes(&package(&input, None, &[("Pictures/pixel.png", PNG)]))
                .unwrap();
        let paragraph = parsed.document.paragraph(0).unwrap();
        assert_eq!(paragraph.alignment(), Some(Alignment::Right));
        assert_eq!(paragraph.run(1).unwrap().bold_value(), Some(true));
        assert_eq!(paragraph.run(1).unwrap().italic_value(), Some(true));
        assert_eq!(paragraph.run(1).unwrap().color(), Some("123456"));
        let numbering = parsed.document.paragraph(1).unwrap().numbering().unwrap();
        assert_eq!(
            parsed.document.numbering_is_bullet(numbering.0),
            Some(false)
        );
        let table = parsed.document.table(0).unwrap();
        let cell = table.cell(0, 0).unwrap();
        assert_eq!(cell.grid_span(), Some(2));
        assert_eq!(cell.paragraph_count(), 2);
        assert!(matches!(
            table.cell(1, 0).unwrap().v_merge(),
            Some(VMerge::Restart)
        ));
        assert!(matches!(
            table.cell(2, 0).unwrap().v_merge(),
            Some(VMerge::Continue)
        ));
        let images = parsed.document.images();
        assert_eq!(images.len(), 1);
        assert_eq!(
            (images[0].width_emu, images[0].height_emu),
            (914_400, 457_200)
        );

        let root = std::env::temp_dir().join(format!("rdocx-open-odt-{}", std::process::id()));
        let _ = std::fs::remove_dir_all(&root);
        std::fs::create_dir(&root).unwrap();
        let path = root.join("source.odt");
        std::fs::write(&path, package(&input, None, &[("Pictures/pixel.png", PNG)])).unwrap();
        let opened = Document::open_odt(&path).unwrap();
        assert_eq!(opened.document.text(), parsed.document.text());
        std::fs::remove_dir_all(root).unwrap();
    }

    #[test]
    fn odt_writer_emits_conforming_deterministic_package() {
        let mut document = Document::new();
        document
            .add_paragraph(" deterministic  text\tline\nbreak ")
            .add_run("styled")
            .bold(true);
        document.add_paragraph("heading").set_style("Heading1");
        let list = document.add_list_definition(&[ListLevel::bullet(), ListLevel::decimal()]);
        document.add_paragraph("top item").set_numbering(list, 0);
        document.add_paragraph("nested item").set_numbering(list, 1);
        document.add_paragraph("list interruption");
        document
            .add_paragraph("continued item")
            .set_numbering(list, 0);
        {
            let mut table = document.add_table(3, 2);
            table.cell(0, 0).unwrap().set_text("spanned");
            table.cell(0, 0).unwrap().set_grid_span(2);
            table.cell(1, 0).unwrap().set_text("vertical");
            table.cell(1, 0).unwrap().set_v_merge_restart();
            table.cell(2, 0).unwrap().set_v_merge_continue();
        }
        let BodyContent::Table(table) = document.document.body.content.last_mut().unwrap() else {
            unreachable!();
        };
        table.rows[0].cells.remove(1);
        document.add_picture(PNG, "pixel.png", Length::inches(1.0), Length::inches(1.0));

        let first = document.to_odt_bytes().unwrap();
        let second = document.to_odt_bytes().unwrap();
        assert_eq!(first.bytes, second.bytes);
        assert_eq!(first.diagnostics, second.diagnostics);

        let mut archive = ZipArchive::new(Cursor::new(first.bytes)).unwrap();
        assert_eq!(archive.len(), 4);
        assert!(archive.comment().is_empty());
        for (index, (name, compression)) in [
            ("mimetype", CompressionMethod::Stored),
            ("content.xml", CompressionMethod::Deflated),
            ("Pictures/image1.png", CompressionMethod::Deflated),
            ("META-INF/manifest.xml", CompressionMethod::Deflated),
        ]
        .into_iter()
        .enumerate()
        {
            let entry = archive.by_index(index).unwrap();
            assert_eq!(entry.name(), name);
            assert_eq!(entry.name_raw(), name.as_bytes());
            assert_eq!(entry.compression(), compression);
            assert_eq!(entry.last_modified(), Some(zip::DateTime::DEFAULT));
            assert_eq!(entry.unix_mode(), Some(0o100_644));
            assert!(entry.is_file());
            assert!(!entry.encrypted());
            assert!(entry.comment().is_empty());
            assert!(entry.extra_data().is_none_or(<[u8]>::is_empty));
            assert_eq!(
                entry.data_start().unwrap() - entry.header_start(),
                30 + entry.name_raw().len() as u64,
                "ODT local headers must not contain extra fields"
            );
        }

        let mut content_xml = String::new();
        archive
            .by_name("content.xml")
            .unwrap()
            .read_to_string(&mut content_xml)
            .unwrap();
        assert!(content_xml.starts_with(&format!(
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document-content xmlns:office=\"{OFFICE_NS}\" xmlns:text=\"{TEXT_NS}\" xmlns:style=\"{STYLE_NS}\" xmlns:fo=\"{FO_NS}\" xmlns:table=\"{TABLE_NS}\" xmlns:draw=\"{DRAW_NS}\" xmlns:xlink=\"{XLINK_NS}\" xmlns:svg=\"{SVG_NS}\" office:version=\"1.3\">"
        )));
        let automatic = content_xml.find("<office:automatic-styles>").unwrap();
        let body = content_xml.find("<office:body>").unwrap();
        assert!(automatic < body);
        let content_root =
            parse_xml("content.xml", content_xml.as_bytes(), OdtLimits::DEFAULT).unwrap();
        assert_emitted_content_subset(&content_root);
        assert_emitted_list_continuations(&content_root);

        let mut image = Vec::new();
        archive
            .by_name("Pictures/image1.png")
            .unwrap()
            .read_to_end(&mut image)
            .unwrap();
        assert_eq!(image, PNG);

        let mut manifest = String::new();
        archive
            .by_name("META-INF/manifest.xml")
            .unwrap()
            .read_to_string(&mut manifest)
            .unwrap();
        assert!(manifest.starts_with(&format!(
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?><manifest:manifest xmlns:manifest=\"{MANIFEST_NS}\" manifest:version=\"1.3\">"
        )));
        for entry in [
            format!(
                "manifest:full-path=\"/\" manifest:media-type=\"{}\"",
                std::str::from_utf8(ODT_MIMETYPE).unwrap()
            ),
            "manifest:full-path=\"content.xml\" manifest:media-type=\"text/xml\"".to_string(),
            "manifest:full-path=\"Pictures/image1.png\" manifest:media-type=\"image/png\""
                .to_string(),
        ] {
            assert!(manifest.contains(&entry), "missing manifest entry {entry}");
        }
        assert!(manifest.ends_with("</manifest:manifest>"));
        parse_xml(
            "META-INF/manifest.xml",
            manifest.as_bytes(),
            OdtLimits::DEFAULT,
        )
        .unwrap();
        assert!(
            content_xml.len() <= OdtLimits::DEFAULT.archive.max_part_uncompressed_bytes as usize
        );
        assert!(manifest.len() <= OdtLimits::DEFAULT.archive.max_part_uncompressed_bytes as usize);

        let tiny = OdtLimits {
            archive: PackageReadLimits {
                max_part_uncompressed_bytes: 32,
                ..OdtLimits::DEFAULT.archive
            },
            ..OdtLimits::DEFAULT
        };
        assert!(OdtWriter::new_with_limits(&document, tiny).write().is_err());

        for limits in [
            OdtLimits {
                archive: PackageReadLimits {
                    max_entries: 3,
                    ..OdtLimits::DEFAULT.archive
                },
                ..OdtLimits::DEFAULT
            },
            OdtLimits {
                archive: PackageReadLimits {
                    max_part_uncompressed_bytes: (PNG.len() - 1) as u64,
                    ..OdtLimits::DEFAULT.archive
                },
                ..OdtLimits::DEFAULT
            },
            OdtLimits {
                archive: PackageReadLimits {
                    max_total_uncompressed_bytes: 2_500,
                    ..OdtLimits::DEFAULT.archive
                },
                ..OdtLimits::DEFAULT
            },
        ] {
            assert!(
                OdtWriter::new_with_limits(&document, limits)
                    .write()
                    .is_err()
            );
        }

        let mut lossy = Document::new();
        lossy.add_paragraph("diagnostic").set_keep_with_next(true);
        let no_diagnostics = OdtLimits {
            diagnostics: 0,
            ..OdtLimits::DEFAULT
        };
        assert!(
            OdtWriter::new_with_limits(&lossy, no_diagnostics)
                .write()
                .is_err()
        );
    }

    #[test]
    fn odt_writer_flattens_simple_field_cached_display() {
        let mut document = Document::new();
        document.add_paragraph("");
        let BodyContent::Paragraph(paragraph) = &mut document.document.body.content[0] else {
            unreachable!();
        };
        paragraph.runs.push({
            let mut run = CT_R::new("");
            run.content = vec![RunContent::Field(rdocx_oxml::text::Field::new(
                "PAGE", "cached 7",
            ))];
            run
        });

        let written = document.to_odt_bytes().unwrap();
        let reopened = Document::from_odt_bytes(&written.bytes).unwrap().document;
        assert_eq!(reopened.paragraph(0).unwrap().text(), "cached 7");
        assert!(written.diagnostics.iter().any(|diagnostic| {
            diagnostic.path == "body[0]/run[0]/content[0]"
                && diagnostic.message == "field was flattened during ODT export"
        }));
    }

    #[test]
    fn odt_writer_rejects_paragraph_values_outside_reader_domain() {
        for properties in [
            CT_PPr {
                space_before: Some(Twips(-1)),
                ..Default::default()
            },
            CT_PPr {
                space_after: Some(Twips(-1)),
                ..Default::default()
            },
            CT_PPr {
                line_spacing: Some(Twips(0)),
                line_rule: Some("exact".to_string()),
                ..Default::default()
            },
            CT_PPr {
                line_spacing: Some(Twips(-1)),
                line_rule: Some("auto".to_string()),
                ..Default::default()
            },
            CT_PPr {
                ind_hanging: Some(Twips(-1)),
                ..Default::default()
            },
        ] {
            let mut document = Document::new();
            document.add_paragraph("unrepresentable");
            let BodyContent::Paragraph(paragraph) = &mut document.document.body.content[0] else {
                unreachable!();
            };
            paragraph.properties = Some(properties);
            let error = document.to_odt_bytes().err().unwrap();
            assert!(error.to_string().contains("ODT cannot preserve"));
        }
    }

    #[test]
    fn odt_writer_preserves_every_accepted_automatic_line_height_twip() {
        let mut document = Document::new();
        for value in 1..=24_000 {
            let mut paragraph = CT_P::new();
            paragraph.properties = Some(CT_PPr {
                line_spacing: Some(Twips(value)),
                line_rule: Some("auto".to_string()),
                ..Default::default()
            });
            document
                .document
                .body
                .content
                .push(BodyContent::Paragraph(paragraph));
        }

        let reopened = Document::from_odt_bytes(&document.to_odt_bytes().unwrap().bytes)
            .unwrap()
            .document;
        assert_eq!(reopened.document.body.content.len(), 24_000);
        for (index, expected) in (1..=24_000).enumerate() {
            let BodyContent::Paragraph(paragraph) = &reopened.document.body.content[index] else {
                unreachable!();
            };
            assert_eq!(
                paragraph
                    .properties
                    .as_ref()
                    .and_then(|properties| properties.line_spacing),
                Some(Twips(expected))
            );
        }
    }

    #[test]
    fn odt_writer_preserves_inclusive_paragraph_length_boundaries() {
        let mut document = Document::new();
        document.add_paragraph("maximum paragraph lengths");
        let BodyContent::Paragraph(paragraph) = &mut document.document.body.content[0] else {
            unreachable!();
        };
        paragraph.properties = Some(CT_PPr {
            space_before: Some(Twips(20_000_000)),
            space_after: Some(Twips(20_000_000)),
            ind_left: Some(Twips(20_000_000)),
            ind_right: Some(Twips(20_000_000)),
            ind_first_line: Some(Twips(-20_000_000)),
            line_spacing: Some(Twips(20_000_000)),
            line_rule: Some("exact".to_string()),
            ..Default::default()
        });

        let reopened = Document::from_odt_bytes(&document.to_odt_bytes().unwrap().bytes)
            .unwrap()
            .document;
        let paragraph = reopened.paragraph(0).unwrap();
        assert_eq!(paragraph.space_before(), Some(Length::twips(20_000_000)));
        assert_eq!(paragraph.space_after(), Some(Length::twips(20_000_000)));
        assert_eq!(paragraph.indent_left(), Some(Length::twips(20_000_000)));
        assert_eq!(paragraph.indent_right(), Some(Length::twips(20_000_000)));
        assert_eq!(
            paragraph.first_line_indent(),
            Some(Length::twips(-20_000_000))
        );
        assert_eq!(paragraph.line_spacing(), Some(Length::twips(20_000_000)));
    }

    #[test]
    fn odt_writer_rejects_derived_heading_levels_outside_odt_range() {
        let mut valid = Document::new();
        valid.add_paragraph("first heading").set_style("Heading1");
        valid.add_paragraph("last heading").set_style("Heading9");
        let reopened = Document::from_odt_bytes(&valid.to_odt_bytes().unwrap().bytes)
            .unwrap()
            .document;
        assert_eq!(reopened.paragraph(0).unwrap().style_id(), Some("Heading1"));
        assert_eq!(reopened.paragraph(1).unwrap().style_id(), Some("Heading9"));

        for style in ["Heading0", "Heading10", "Heading999999999999999999999999"] {
            let mut invalid = Document::new();
            invalid.add_paragraph("invalid heading").set_style(style);
            let error = invalid.to_odt_bytes().err().unwrap();
            assert!(error.to_string().contains("cannot preserve outline level"));
            assert!(error.to_string().contains("body[0]"));
        }
    }

    #[test]
    fn odt_writer_honors_direct_numbering_cancellation() {
        let mut document = Document::new();
        let list = document.add_list_definition(&[ListLevel::bullet()]);
        document.add_style(
            crate::StyleBuilder::paragraph("Numbered", "Numbered").paragraph_properties(CT_PPr {
                num_id: Some(list),
                num_ilvl: Some(0),
                ..Default::default()
            }),
        );
        let mut paragraph = document.add_paragraph("not a list");
        paragraph.set_style("Numbered");
        paragraph.set_numbering(0, 0);

        let written = document.to_odt_bytes().unwrap();
        let mut archive = ZipArchive::new(Cursor::new(&written.bytes)).unwrap();
        let mut content_xml = String::new();
        archive
            .by_name("content.xml")
            .unwrap()
            .read_to_string(&mut content_xml)
            .unwrap();
        assert!(!content_xml.contains("<text:list"));
        let reopened = Document::from_odt_bytes(&written.bytes).unwrap().document;
        assert_eq!(reopened.paragraph(0).unwrap().numbering(), None);
        assert!(
            written
                .diagnostics
                .iter()
                .all(|diagnostic| !diagnostic.path.starts_with("numbering["))
        );
    }

    #[test]
    fn odt_writer_does_not_invent_markers_for_producer_defined_numbering() {
        let mut document = Document::new();
        let number = document.add_list_definition(&[ListLevel::decimal()]);
        let abstract_id = document
            .numbering
            .as_ref()
            .unwrap()
            .nums
            .iter()
            .find(|item| item.num_id == number)
            .unwrap()
            .abstract_num_id;
        let numbering = document.numbering.as_mut().unwrap();
        numbering
            .root_attributes
            .push(("xmlns:x".to_owned(), "urn:producer".to_owned()));
        numbering
            .extra_xml
            .push((0, b"<x:root xmlns:x=\"urn:producer\"/>".to_vec()));
        let instance = numbering
            .nums
            .iter_mut()
            .find(|item| item.num_id == number)
            .unwrap();
        instance
            .extra_attributes
            .push(("x:instance".to_owned(), "retained".to_owned()));
        instance
            .extra_xml
            .push((0, b"<x:instance xmlns:x=\"urn:producer\"/>".to_vec()));
        let abstract_numbering = numbering
            .abstract_nums
            .iter_mut()
            .find(|item| item.abstract_num_id == abstract_id)
            .unwrap();
        abstract_numbering.multi_level_type = Some("hybridMultilevel".to_owned());
        abstract_numbering
            .extra_attributes
            .push(("x:abstract".to_owned(), "retained".to_owned()));
        abstract_numbering
            .extra_xml
            .push((0, b"<x:abstract xmlns:x=\"urn:producer\"/>".to_vec()));
        abstract_numbering.levels[0] = {
            let mut level = rdocx_oxml::numbering::CT_Lvl::new(0);
            level.num_fmt = Some(ST_NumberFormat::Other("chicago".to_owned()));
            level.start = Some(3);
            level.suffix = Some(rdocx_oxml::numbering::ST_LvlSuffix::Space);
            level.lvl_text = Some("custom".to_owned());
            level.lvl_jc = Some(ST_Jc::Right);
            level.ppr = Some(CT_PPr::default());
            level.rpr = Some(CT_RPr::default());
            level.extra_xml.push((0, b"<x:producer/>".to_vec()));
            level
        };
        document
            .add_paragraph("producer marker")
            .set_numbering(number, 0);
        document
            .add_paragraph("repeated producer marker")
            .set_numbering(number, 0);
        let mut table = CT_Tbl::new();
        let mut row = CT_Row::new();
        let mut cell = CT_Tc::new();
        let CellContent::Paragraph(cell_paragraph) = &mut cell.content[0] else {
            unreachable!();
        };
        cell_paragraph.add_run("cell producer marker");
        cell_paragraph.properties = Some(CT_PPr {
            num_id: Some(number),
            num_ilvl: Some(0),
            ..Default::default()
        });
        row.cells.push(cell);
        table.rows.push(row);
        document
            .document
            .body
            .content
            .push(BodyContent::Table(table));

        let written = document.to_odt_bytes().unwrap();
        let mut archive = ZipArchive::new(Cursor::new(&written.bytes)).unwrap();
        let mut content_xml = String::new();
        archive
            .by_name("content.xml")
            .unwrap()
            .read_to_string(&mut content_xml)
            .unwrap();
        assert!(!content_xml.contains("<text:list"), "{content_xml}");
        let numbering_diagnostics: Vec<_> = written
            .diagnostics
            .iter()
            .filter(|diagnostic| diagnostic.path.starts_with(&format!("numbering[{number}]")))
            .map(|diagnostic| (diagnostic.path.clone(), diagnostic.message.clone()))
            .collect();
        let level_path = format!("numbering[{number}]/level[0]");
        let root_path = format!("numbering[{number}]");
        assert_eq!(
            numbering_diagnostics,
            vec![
                (
                    level_path.clone(),
                    "producer-defined numbering format was flattened without a marker during ODT export"
                        .to_owned(),
                ),
                (
                    format!("{root_path}/root-attributes"),
                    "retained numbering root attributes were dropped during ODT export".to_owned(),
                ),
                (
                    format!("{root_path}/root-xml"),
                    "retained numbering root XML was dropped during ODT export".to_owned(),
                ),
                (
                    format!("{root_path}/instance-attributes"),
                    "retained numbering instance attributes were dropped during ODT export"
                        .to_owned(),
                ),
                (
                    format!("{root_path}/instance-overrides"),
                    "numbering instance overrides or retained XML were dropped during ODT export"
                        .to_owned(),
                ),
                (
                    format!("{root_path}/abstract-type"),
                    "abstract numbering type metadata was dropped during ODT export".to_owned(),
                ),
                (
                    format!("{root_path}/abstract-attributes"),
                    "retained abstract numbering attributes were dropped during ODT export"
                        .to_owned(),
                ),
                (
                    format!("{root_path}/abstract-xml"),
                    "retained abstract numbering XML was dropped during ODT export".to_owned(),
                ),
                (
                    level_path.clone(),
                    "custom list start value was dropped during ODT export".to_owned(),
                ),
                (
                    level_path.clone(),
                    "list marker suffix was dropped during ODT export".to_owned(),
                ),
                (
                    level_path.clone(),
                    "list marker text or bullet glyph was dropped during ODT export".to_owned(),
                ),
                (
                    level_path.clone(),
                    "list marker justification was dropped during ODT export".to_owned(),
                ),
                (
                    level_path.clone(),
                    "list level paragraph formatting was dropped during ODT export".to_owned(),
                ),
                (
                    level_path.clone(),
                    "list marker run formatting was dropped during ODT export".to_owned(),
                ),
                (
                    level_path,
                    "retained list level XML or attributes were dropped during ODT export"
                        .to_owned(),
                ),
            ]
        );
        let reopened = Document::from_odt_bytes(&written.bytes).unwrap().document;
        assert_eq!(
            reopened
                .paragraphs()
                .into_iter()
                .map(|paragraph| (paragraph.text(), paragraph.numbering()))
                .collect::<Vec<_>>(),
            vec![
                ("producer marker".to_owned(), None),
                ("repeated producer marker".to_owned(), None),
            ]
        );
        let BodyContent::Table(table) = &reopened.document.body.content[2] else {
            unreachable!();
        };
        let CellContent::Paragraph(cell_paragraph) = &table.rows[0].cells[0].content[0] else {
            unreachable!();
        };
        assert_eq!(cell_paragraph.text(), "cell producer marker");
        assert_eq!(
            cell_paragraph
                .properties
                .as_ref()
                .and_then(paragraph_numbering_properties),
            None
        );
    }

    #[test]
    fn odt_writer_rejects_non_image_and_external_image_relationships() {
        let mut document = Document::new();
        let wrong_type = document.embed_image(PNG, "wrong-type.png");
        let external = document.embed_image(PNG, "external.png");
        let invalid_mode = document.embed_image(PNG, "invalid-mode.png");
        document.add_paragraph("").add_picture(
            &wrong_type,
            Length::emu(12_700),
            Length::emu(12_700),
        );
        document
            .add_paragraph("")
            .add_picture(&external, Length::emu(12_700), Length::emu(12_700));
        document.add_paragraph("").add_picture(
            &invalid_mode,
            Length::emu(12_700),
            Length::emu(12_700),
        );
        let document_part = document.doc_part_name.clone();
        let relationships = document.package.get_or_create_part_rels(&document_part);
        relationships
            .items
            .iter_mut()
            .find(|relationship| relationship.id == wrong_type)
            .unwrap()
            .rel_type = oxml_opc::relationship::rel_types::HYPERLINK.to_string();
        relationships
            .items
            .iter_mut()
            .find(|relationship| relationship.id == external)
            .unwrap()
            .target_mode = Some("External".to_string());
        relationships
            .items
            .iter_mut()
            .find(|relationship| relationship.id == invalid_mode)
            .unwrap()
            .target_mode = Some("Bogus".to_string());

        let written = document.to_odt_bytes().unwrap();
        for expected in [
            (
                "body[0]/run[0]/content[0]",
                "inline drawing with a non-image relationship was dropped during ODT export",
            ),
            (
                "body[1]/run[0]/content[0]",
                "external inline image was dropped during ODT export",
            ),
            (
                "body[2]/run[0]/content[0]",
                "invalid inline image target mode was dropped during ODT export",
            ),
        ] {
            assert!(written.diagnostics.iter().any(|diagnostic| {
                (diagnostic.path.as_str(), diagnostic.message.as_str()) == expected
            }));
        }
        let mut archive = ZipArchive::new(Cursor::new(&written.bytes)).unwrap();
        assert_eq!(archive.len(), 3);
        assert!(archive.by_name("Pictures/image1.png").is_err());
    }

    #[test]
    fn odt_writer_rejects_values_above_image_and_font_reader_domains() {
        let mut maximum = Document::new();
        maximum.add_picture(
            PNG,
            "maximum.png",
            Length::emu(12_700_000_000),
            Length::emu(12_700_000_000),
        );
        maximum.add_paragraph("").add_run("maximum font");
        let BodyContent::Paragraph(paragraph) = &mut maximum.document.body.content[1] else {
            unreachable!();
        };
        paragraph.runs[0].properties = Some(CT_RPr {
            sz: Some(rdocx_oxml::units::HalfPoint(2_000_000)),
            ..Default::default()
        });
        let reopened = Document::from_odt_bytes(&maximum.to_odt_bytes().unwrap().bytes)
            .unwrap()
            .document;
        assert_eq!(reopened.images()[0].width_emu, 12_700_000_000);
        assert_eq!(reopened.images()[0].height_emu, 12_700_000_000);
        assert_eq!(
            reopened.paragraph(1).unwrap().run(0).unwrap().size(),
            Some(1_000_000.0)
        );

        let mut oversized_image = Document::new();
        oversized_image.add_picture(
            PNG,
            "oversized.png",
            Length::emu(12_700_000_001),
            Length::emu(1),
        );
        assert!(
            oversized_image
                .to_odt_bytes()
                .err()
                .unwrap()
                .to_string()
                .contains("cannot preserve inline image dimensions")
        );

        for size in [0, 2_000_001] {
            let mut invalid = Document::new();
            invalid.add_paragraph("invalid font");
            let BodyContent::Paragraph(paragraph) = &mut invalid.document.body.content[0] else {
                unreachable!();
            };
            paragraph.runs[0].properties = Some(CT_RPr {
                sz: Some(rdocx_oxml::units::HalfPoint(size)),
                ..Default::default()
            });
            assert!(
                invalid
                    .to_odt_bytes()
                    .err()
                    .unwrap()
                    .to_string()
                    .contains("cannot preserve font size")
            );
        }
    }

    #[test]
    fn odt_writer_enforces_reader_block_row_and_cell_ceilings() {
        let mut blocks = Document::new();
        blocks.add_paragraph("one");
        blocks.add_paragraph("two");
        let block_limits = OdtLimits {
            blocks: 1,
            ..OdtLimits::DEFAULT
        };
        assert!(
            OdtWriter::new_with_limits(&blocks, block_limits)
                .write()
                .err()
                .unwrap()
                .to_string()
                .contains("block limit")
        );

        let mut rows = Document::new();
        rows.add_table(2, 1);
        let row_limits = OdtLimits {
            rows: 1,
            ..OdtLimits::DEFAULT
        };
        assert!(
            OdtWriter::new_with_limits(&rows, row_limits)
                .write()
                .err()
                .unwrap()
                .to_string()
                .contains("row limit")
        );

        let mut cells = Document::new();
        cells.add_table(1, 2);
        let cell_limits = OdtLimits {
            cells: 1,
            ..OdtLimits::DEFAULT
        };
        assert!(
            OdtWriter::new_with_limits(&cells, cell_limits)
                .write()
                .err()
                .unwrap()
                .to_string()
                .contains("cell limit")
        );

        let mut synthetic = Document::new();
        synthetic.add_table(1, 1);
        let BodyContent::Table(table) = &mut synthetic.document.body.content[0] else {
            unreachable!();
        };
        table.rows[0].cells[0].content.clear();
        let synthetic_limits = OdtLimits {
            blocks: 1,
            ..OdtLimits::DEFAULT
        };
        assert!(
            OdtWriter::new_with_limits(&synthetic, synthetic_limits)
                .write()
                .err()
                .unwrap()
                .to_string()
                .contains("block limit")
        );
    }

    #[test]
    fn odt_writer_enforces_reader_run_and_xml_node_ceilings() {
        let mut runs = Document::new();
        runs.add_paragraph("one");
        runs.add_paragraph("two");
        let run_limits = OdtLimits {
            runs: 1,
            ..OdtLimits::DEFAULT
        };
        assert!(
            OdtWriter::new_with_limits(&runs, run_limits)
                .write()
                .err()
                .unwrap()
                .to_string()
                .contains("run limit")
        );

        let mut embedded = Document::new();
        embedded.add_paragraph("a\tb\r\n");
        let four_run_limits = OdtLimits {
            runs: 4,
            ..OdtLimits::DEFAULT
        };
        let written = OdtWriter::new_with_limits(&embedded, four_run_limits)
            .write()
            .unwrap();
        assert_eq!(
            Document::from_odt_bytes(&written.bytes)
                .unwrap()
                .document
                .paragraph(0)
                .unwrap()
                .run_count(),
            4
        );
        let three_run_limits = OdtLimits {
            runs: 3,
            ..OdtLimits::DEFAULT
        };
        assert!(
            OdtWriter::new_with_limits(&embedded, three_run_limits)
                .write()
                .err()
                .unwrap()
                .to_string()
                .contains("run limit")
        );

        let mut field = Document::new();
        field.add_paragraph("");
        let BodyContent::Paragraph(paragraph) = &mut field.document.body.content[0] else {
            unreachable!();
        };
        let mut field_run = CT_R::new("");
        field_run.content = vec![RunContent::Field(rdocx_oxml::text::Field::new(
            "PAGE", "a\tb\r\n",
        ))];
        paragraph.runs.push(field_run);
        let written = OdtWriter::new_with_limits(&field, four_run_limits)
            .write()
            .unwrap();
        assert_eq!(
            Document::from_odt_bytes(&written.bytes)
                .unwrap()
                .document
                .paragraph(0)
                .unwrap()
                .run_count(),
            4
        );
        assert!(
            OdtWriter::new_with_limits(&field, three_run_limits)
                .write()
                .err()
                .unwrap()
                .to_string()
                .contains("run limit")
        );

        let mut nodes = Document::new();
        nodes.add_paragraph("   ");
        let node_limits = OdtLimits {
            xml_nodes: 5,
            ..OdtLimits::DEFAULT
        };
        assert!(
            OdtWriter::new_with_limits(&nodes, node_limits)
                .write()
                .err()
                .unwrap()
                .to_string()
                .contains("node limit")
        );
    }

    #[test]
    fn odt_writer_preserves_list_continuation_across_body_and_cell_siblings() {
        let mut document = Document::new();
        let list = document.add_list_definition(&[ListLevel::decimal()]);
        document.add_paragraph("one").set_numbering(list, 0);
        document.add_paragraph("interruption");
        document.add_paragraph("two").set_numbering(list, 0);
        {
            let mut table = document.add_table(1, 1);
            let mut cell = table.cell(0, 0).unwrap();
            cell.set_text("cell one");
            cell.paragraph_mut(0).unwrap().set_numbering(list, 0);
            cell.add_paragraph("cell interruption");
            cell.add_paragraph("cell two").set_numbering(list, 0);
        }

        let written = document.to_odt_bytes().unwrap();
        let mut archive = ZipArchive::new(Cursor::new(&written.bytes)).unwrap();
        let mut content_xml = String::new();
        archive
            .by_name("content.xml")
            .unwrap()
            .read_to_string(&mut content_xml)
            .unwrap();
        assert_eq!(
            content_xml
                .matches("text:continue-numbering=\"true\"")
                .count(),
            2
        );
        let root = parse_xml("content.xml", content_xml.as_bytes(), OdtLimits::DEFAULT).unwrap();
        assert_emitted_list_continuations(&root);
    }

    #[test]
    fn odt_writer_rejects_unicode_whitespace_and_out_of_range_outline_levels() {
        for text in ["a\u{a0}b", "a\u{2003}b"] {
            let mut document = Document::new();
            document.add_paragraph(text);
            assert!(
                document
                    .to_odt_bytes()
                    .err()
                    .unwrap()
                    .to_string()
                    .contains("cannot preserve Unicode whitespace at body[0]/run[0]/content[0]")
            );
        }

        let mut field = Document::new();
        field.add_paragraph("");
        let BodyContent::Paragraph(paragraph) = &mut field.document.body.content[0] else {
            unreachable!();
        };
        paragraph.runs.push({
            let mut run = CT_R::new("");
            run.content = vec![RunContent::Field(rdocx_oxml::text::Field::new(
                "PAGE", "a\u{a0}b",
            ))];
            run
        });
        assert!(
            field
                .to_odt_bytes()
                .err()
                .unwrap()
                .to_string()
                .contains("cannot preserve Unicode whitespace at body[0]/run[0]/content[0]")
        );

        let mut outline = Document::new();
        outline.add_paragraph("outline");
        let BodyContent::Paragraph(paragraph) = &mut outline.document.body.content[0] else {
            unreachable!();
        };
        paragraph.properties = Some(CT_PPr {
            outline_lvl: Some(9),
            ..Default::default()
        });
        assert!(
            outline
                .to_odt_bytes()
                .err()
                .unwrap()
                .to_string()
                .contains("cannot preserve outline level at body[0]")
        );
    }

    #[test]
    fn odt_writer_diagnoses_unresolved_and_wrong_type_final_story_references() {
        let mut document = Document::new();
        document.set_header("header");
        document.set_footer("footer");
        let section = document.document.body.sect_pr.as_mut().unwrap();
        section.header_refs[0].rel_id = "missingHeader".to_string();
        let footer_id = section.footer_refs[0].rel_id.clone();
        let document_part = document.doc_part_name.clone();
        document
            .package
            .get_or_create_part_rels(&document_part)
            .items
            .iter_mut()
            .find(|relationship| relationship.id == footer_id)
            .unwrap()
            .rel_type = oxml_opc::relationship::rel_types::IMAGE.to_string();

        let written = document.to_odt_bytes().unwrap();
        for expected in [
            (
                "document/sectPr/headerReference[0]",
                "unresolved or wrong-type header reference was dropped during ODT export",
            ),
            (
                "document/sectPr/footerReference[0]",
                "unresolved or wrong-type footer reference was dropped during ODT export",
            ),
        ] {
            assert!(written.diagnostics.iter().any(|diagnostic| {
                (diagnostic.path.as_str(), diagnostic.message.as_str()) == expected
            }));
        }
    }

    #[test]
    fn odt_writer_rejects_vertical_continuation_after_an_overlapping_cell() {
        let mut document = Document::new();
        document.add_table(4, 2);
        let BodyContent::Table(table) = &mut document.document.body.content[0] else {
            unreachable!();
        };
        table.rows[0].cells[1].properties = Some(CT_TcPr {
            v_merge: Some(VMerge::Restart),
            ..Default::default()
        });
        table.rows[1].cells[1].properties = Some(CT_TcPr {
            v_merge: Some(VMerge::Continue),
            ..Default::default()
        });
        table.rows[2].cells[0].properties = Some(CT_TcPr {
            grid_span: Some(2),
            ..Default::default()
        });
        table.rows[2].cells.remove(1);
        table.rows[3].cells[1].properties = Some(CT_TcPr {
            v_merge: Some(VMerge::Continue),
            ..Default::default()
        });
        assert!(
            document
                .to_odt_bytes()
                .err()
                .unwrap()
                .to_string()
                .contains("malformed vertical table span at body[0]/row[3]/cell[1]")
        );
    }

    #[test]
    fn odt_writer_reports_every_color_shading_and_document_story_loss() {
        let mut document = Document::new();
        document.set_landscape();
        document.set_header("header");
        document.set_footer("footer");
        document.document.background_xml = Some(
            br#"<w:background xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:color="123456"/>"#
                .to_vec(),
        );
        document.add_paragraph("losses");
        let BodyContent::Paragraph(paragraph) = &mut document.document.body.content[0] else {
            unreachable!();
        };
        paragraph.runs[0].properties = Some(CT_RPr {
            color: Some("auto".to_string()),
            highlight: Some(ST_HighlightColor::Yellow),
            shading: Some(rdocx_oxml::properties::CT_Shd {
                val: "horzStripe".to_string(),
                color: Some("ABCDEF".to_string()),
                fill: Some("123456".to_string()),
            }),
            ..Default::default()
        });
        paragraph.runs.push({
            let mut run = CT_R::new("invalid fill");
            run.properties = Some(CT_RPr {
                shading: Some(rdocx_oxml::properties::CT_Shd {
                    val: "clear".to_string(),
                    color: None,
                    fill: Some("auto".to_string()),
                }),
                ..Default::default()
            });
            run
        });

        let written = document.to_odt_bytes().unwrap();
        assert_eq!(
            written
                .diagnostics
                .iter()
                .map(|diagnostic| (diagnostic.path.as_str(), diagnostic.message.as_str()))
                .collect::<Vec<_>>(),
            vec![
                (
                    "document/background",
                    "document background was dropped during ODT export",
                ),
                (
                    "document/sectPr",
                    "final section properties were dropped during ODT export",
                ),
                (
                    "document/headers",
                    "header stories were dropped during ODT export",
                ),
                (
                    "document/footers",
                    "footer stories were dropped during ODT export",
                ),
                (
                    "body[0]/run[0]/rPr",
                    "unsupported run properties were dropped during ODT export",
                ),
                (
                    "body[0]/run[0]/rPr/color",
                    "unsupported run color was dropped during ODT export",
                ),
                (
                    "body[0]/run[0]/rPr/shading-pattern",
                    "run shading pattern was simplified during ODT export",
                ),
                (
                    "body[0]/run[0]/rPr/shading-color",
                    "run shading foreground color was dropped during ODT export",
                ),
                (
                    "body[0]/run[0]/rPr/highlight",
                    "run highlight was replaced by shading fill during ODT export",
                ),
                (
                    "body[0]/run[1]/rPr",
                    "unsupported run properties were dropped during ODT export",
                ),
                (
                    "body[0]/run[1]/rPr/shading-fill",
                    "unsupported run shading fill was dropped during ODT export",
                ),
            ]
        );
        let reopened = Document::from_odt_bytes(&written.bytes).unwrap().document;
        let paragraph = reopened.paragraph(0).unwrap();
        let run = paragraph.run(0).unwrap();
        assert_eq!(run.color(), None);
        assert_eq!(run.highlight().as_deref(), Some("123456"));
    }

    #[test]
    fn odt_writer_validates_caller_strings_before_writing_xml() {
        let mut invalid_font = Document::new();
        invalid_font
            .add_paragraph("")
            .add_run("text")
            .font("Bad\u{1}Font");
        assert!(
            invalid_font
                .to_odt_bytes()
                .err()
                .unwrap()
                .to_string()
                .contains("invalid XML character at body[0]/run[0]/rPr/font")
        );

        let mut invalid_field = Document::new();
        invalid_field.add_paragraph("");
        let BodyContent::Paragraph(paragraph) = &mut invalid_field.document.body.content[0] else {
            unreachable!();
        };
        paragraph.runs.push({
            let mut run = CT_R::new("");
            run.content = vec![RunContent::Field(rdocx_oxml::text::Field::new(
                "PAGE",
                "Bad\u{1}Display",
            ))];
            run
        });
        assert!(
            invalid_field
                .to_odt_bytes()
                .err()
                .unwrap()
                .to_string()
                .contains("invalid XML character at body[0]/run[0]/content[0]")
        );

        for font in [
            " Leading",
            "Trailing ",
            "Tab\tFamily",
            "Line\nFamily",
            "Return\rFamily",
            "'Quoted'",
            "\"Quoted\"",
        ] {
            let mut normalized_font = Document::new();
            normalized_font.add_paragraph("").add_run("text").font(font);
            let error = normalized_font.to_odt_bytes().err().unwrap();
            assert!(
                error
                    .to_string()
                    .contains("ODT cannot preserve font family at body[0]/run[0]/rPr/font")
            );
        }
    }

    #[test]
    fn odt_writer_diagnoses_every_unsupported_vertical_alignment() {
        let mut document = Document::new();
        document.add_paragraph("");
        let BodyContent::Paragraph(paragraph) = &mut document.document.body.content[0] else {
            unreachable!();
        };
        for (text, alignment) in [("baseline", "baseline"), ("malformed", "sideways")] {
            let mut run = CT_R::new(text);
            run.properties = Some(CT_RPr {
                vert_align: Some(alignment.to_string()),
                ..Default::default()
            });
            paragraph.runs.push(run);
        }

        let written = document.to_odt_bytes().unwrap();
        for index in 0..2 {
            assert!(written.diagnostics.iter().any(|diagnostic| {
                diagnostic.path == format!("body[0]/run[{index}]/rPr/vertAlign")
                    && diagnostic.message
                        == "unsupported run vertical alignment was dropped during ODT export"
            }));
        }
        let reopened = Document::from_odt_bytes(&written.bytes).unwrap().document;
        let paragraph = reopened.paragraph(0).unwrap();
        assert_eq!(paragraph.text(), "baselinemalformed");
        assert!(paragraph.runs().all(|run| run.vert_align().is_none()));
    }

    #[test]
    fn odt_writer_diagnoses_empty_drawing_without_dropping_text_siblings() {
        let mut document = Document::new();
        document.add_paragraph("");
        let BodyContent::Paragraph(paragraph) = &mut document.document.body.content[0] else {
            unreachable!();
        };
        let mut run = CT_R::new("");
        run.content = vec![
            RunContent::Text(rdocx_oxml::text::CT_Text::new("before")),
            RunContent::Drawing(rdocx_oxml::drawing::CT_Drawing {
                inline: None,
                anchor: None,
            }),
            RunContent::Text(rdocx_oxml::text::CT_Text::new("after")),
        ];
        paragraph.runs.push(run);

        let written = document.to_odt_bytes().unwrap();
        assert!(written.diagnostics.iter().any(|diagnostic| {
            diagnostic.path == "body[0]/run[0]/content[1]"
                && diagnostic.message == "empty drawing was dropped during ODT export"
        }));
        let reopened = Document::from_odt_bytes(&written.bytes).unwrap().document;
        assert_eq!(reopened.paragraph(0).unwrap().text(), "beforeafter");
    }

    #[test]
    fn odt_writer_diagnoses_direct_and_inherited_distributed_alignment() {
        let mut document = Document::new();
        document.add_style(
            crate::StyleBuilder::paragraph("Distributed", "Distributed").paragraph_properties(
                CT_PPr {
                    jc: Some(ST_Jc::Distribute),
                    ..Default::default()
                },
            ),
        );
        document.add_paragraph("direct");
        let BodyContent::Paragraph(direct) = &mut document.document.body.content[0] else {
            unreachable!();
        };
        direct.properties = Some(CT_PPr {
            jc: Some(ST_Jc::Distribute),
            ..Default::default()
        });
        document.add_paragraph("inherited").set_style("Distributed");

        let written = document.to_odt_bytes().unwrap();
        for path in ["body[0]/pPr/jc", "body[1]/pPr/jc"] {
            assert!(written.diagnostics.iter().any(|diagnostic| {
                diagnostic.path == path
                    && diagnostic.message
                        == "distributed paragraph alignment was simplified to justify during ODT export"
            }));
        }
        let reopened = Document::from_odt_bytes(&written.bytes).unwrap().document;
        assert_eq!(
            reopened.paragraph(0).unwrap().alignment(),
            Some(Alignment::Justify)
        );
        assert_eq!(
            reopened.paragraph(1).unwrap().alignment(),
            Some(Alignment::Justify)
        );
    }

    #[test]
    fn odt_writer_bounds_repeated_media_occurrences_during_scan() {
        let mut document = Document::new();
        let relationship = document.embed_image(PNG, "shared.png");
        for _ in 0..3 {
            document.add_paragraph("").add_picture(
                &relationship,
                Length::inches(1.0),
                Length::inches(1.0),
            );
        }

        let entry_limits = OdtLimits {
            archive: PackageReadLimits {
                max_entries: 4,
                ..OdtLimits::DEFAULT.archive
            },
            ..OdtLimits::DEFAULT
        };
        let entry_error = OdtWriter::new_with_limits(&document, entry_limits)
            .write()
            .err()
            .unwrap();
        assert!(entry_error.to_string().contains("entry limit"));

        let total_limits = OdtLimits {
            archive: PackageReadLimits {
                max_total_uncompressed_bytes: 3_200,
                ..OdtLimits::DEFAULT.archive
            },
            ..OdtLimits::DEFAULT
        };
        let total_error = OdtWriter::new_with_limits(&document, total_limits)
            .write()
            .err()
            .unwrap();
        assert!(total_error.to_string().contains("size limit"));
    }

    #[test]
    fn odt_writer_materializes_effective_formatting_and_whitespace() {
        let mut document = Document::new();
        document.add_style(
            crate::StyleBuilder::paragraph("WriterStyle", "Writer Style").run_properties(CT_RPr {
                font_ascii: Some("Liberation Serif".to_owned()),
                font_hansi: Some("Liberation Serif".to_owned()),
                font_east_asia: Some("Liberation Serif".to_owned()),
                font_cs: Some("Liberation Serif".to_owned()),
                bold: Some(true),
                sz: Some(rdocx_oxml::units::HalfPoint::from_pt(13.0)),
                sz_cs: Some(rdocx_oxml::units::HalfPoint::from_pt(13.0)),
                ..Default::default()
            }),
        );
        document.add_style(
            crate::StyleBuilder::character("WriterCharacter", "Writer Character").run_properties(
                CT_RPr {
                    italic: Some(true),
                    ..Default::default()
                },
            ),
        );
        let mut paragraph = document.add_paragraph("");
        paragraph.set_style("WriterStyle");
        paragraph.set_alignment(Alignment::Center);
        paragraph.set_space_before(Length::twips(61));
        paragraph.set_space_after(Length::twips(119));
        paragraph.set_indent_left(Length::twips(721));
        paragraph.set_indent_right(Length::twips(359));
        paragraph.set_signed_first_line_indent_value(Some(Length::twips(-241)));
        paragraph.set_line_spacing_multiple(1.5);
        paragraph
            .add_run(" boundary  spaces\tline\nbreak ")
            .style("WriterCharacter")
            .underline(true)
            .strike(true)
            .color("123456")
            .highlight("ABCDEF")
            .superscript();

        let written = document.to_odt_bytes().unwrap();
        assert_eq!(
            written
                .diagnostics
                .iter()
                .map(|diagnostic| (diagnostic.path.as_str(), diagnostic.message.as_str()))
                .collect::<Vec<_>>(),
            vec![
                (
                    "body[0]/pPr",
                    "unsupported paragraph properties were dropped during ODT export",
                ),
                (
                    "body[0]/pPr/pStyle",
                    "paragraph style identity was materialized and dropped during ODT export",
                ),
                (
                    "body[0]/run[0]/rPr",
                    "unsupported run properties were dropped during ODT export",
                ),
                (
                    "body[0]/run[0]/rPr/rStyle",
                    "run style identity was materialized and dropped during ODT export",
                ),
            ]
        );
        let reopened = Document::from_odt_bytes(&written.bytes).unwrap().document;
        let paragraph = reopened.paragraph(0).unwrap();
        assert_eq!(paragraph.text(), " boundary  spaces\tline\nbreak ");
        assert_eq!(paragraph.alignment(), Some(Alignment::Center));
        assert_eq!(paragraph.space_before(), Some(Length::twips(61)));
        assert_eq!(paragraph.space_after(), Some(Length::twips(119)));
        assert_eq!(paragraph.indent_left(), Some(Length::twips(721)));
        assert_eq!(paragraph.indent_right(), Some(Length::twips(359)));
        assert_eq!(paragraph.first_line_indent(), Some(Length::twips(-241)));
        assert_eq!(paragraph.line_spacing_multiple(), Some(1.5));
        let run = paragraph.run(0).unwrap();
        assert_eq!(run.font_name(), Some("Liberation Serif"));
        assert_eq!(run.size(), Some(13.0));
        assert_eq!(run.bold_value(), Some(true));
        assert_eq!(run.italic_value(), Some(true));
        assert_eq!(run.underline_code_value(), Some(1));
        assert_eq!(run.strike_value(), Some(true));
        assert_eq!(run.color(), Some("123456"));
        assert_eq!(run.highlight().as_deref(), Some("ABCDEF"));
        assert_eq!(run.vert_align(), Some("superscript"));
    }

    #[test]
    fn odt_writer_emits_nested_lists_table_spans_and_images() {
        let mut document = Document::new();
        let list = document.add_list_definition(&[
            ListLevel::bullet(),
            ListLevel::decimal(),
            ListLevel::bullet(),
        ]);
        document.add_paragraph("top").set_numbering(list, 0);
        document.add_paragraph("nested").set_numbering(list, 1);
        document.add_paragraph("deep").set_numbering(list, 2);
        document.add_paragraph("top two").set_numbering(list, 0);
        {
            let mut table = document.add_table(3, 2);
            table.cell(0, 0).unwrap().set_text("wide");
            table.cell(0, 0).unwrap().set_grid_span(2);
            table.cell(1, 0).unwrap().set_text("vertical");
            table.cell(1, 0).unwrap().set_v_merge_restart();
            table.cell(1, 1).unwrap().set_text("right");
            table.cell(2, 0).unwrap().set_v_merge_continue();
            table.cell(2, 1).unwrap().set_text("bottom");
        }
        let BodyContent::Table(table) = document.document.body.content.last_mut().unwrap() else {
            unreachable!();
        };
        table.rows[0].cells.remove(1);
        document.add_picture(PNG, "pixel.png", Length::emu(914_400), Length::emu(457_200));

        let written = document.to_odt_bytes().unwrap();
        let reopened = Document::from_odt_bytes(&written.bytes).unwrap().document;
        for (index, expected_level) in [0, 1, 2, 0].into_iter().enumerate() {
            let (_, level) = reopened.paragraph(index).unwrap().numbering().unwrap();
            assert_eq!(level, expected_level);
        }
        assert_eq!(
            reopened.numbering_is_bullet(reopened.paragraph(0).unwrap().numbering().unwrap().0),
            Some(true)
        );
        let table = reopened.table(0).unwrap();
        assert_eq!(table.cell(0, 0).unwrap().grid_span(), Some(2));
        assert!(matches!(
            table.cell(1, 0).unwrap().v_merge(),
            Some(VMerge::Restart)
        ));
        assert!(matches!(
            table.cell(2, 0).unwrap().v_merge(),
            Some(VMerge::Continue)
        ));
        let images = reopened.images();
        assert_eq!(images.len(), 1);
        assert_eq!(
            (images[0].width_emu, images[0].height_emu),
            (914_400, 457_200)
        );
        assert_eq!(reopened.image_data(&images[0].embed_id).unwrap(), PNG);
    }

    #[test]
    fn odt_writer_round_trip_preserves_supported_document_content() {
        let mut document = Document::new();
        let mut paragraph = document.add_paragraph("");
        paragraph.set_alignment(Alignment::Right);
        paragraph.add_run("alpha").bold(true).color("123456");
        paragraph.add_run(" beta").italic(true);
        let list = document.add_list_definition(&[ListLevel::decimal()]);
        document.add_paragraph("one").set_numbering(list, 0);
        let mut table = document.add_table(1, 1);
        table.cell(0, 0).unwrap().set_text("cell");
        document.add_picture(PNG, "pixel.png", Length::inches(1.0), Length::inches(0.5));

        let written = document.to_odt_bytes().unwrap();
        assert!(
            written
                .diagnostics
                .iter()
                .any(|diagnostic| diagnostic.path == "body[2]/tblPr")
        );
        assert!(
            written
                .diagnostics
                .iter()
                .any(|diagnostic| diagnostic.path == "body[2]/tblGrid")
        );
        let reopened = Document::from_odt_bytes(&written.bytes).unwrap().document;
        assert_eq!(reopened.paragraph(0).unwrap().text(), "alpha beta");
        assert_eq!(
            reopened.paragraph(0).unwrap().alignment(),
            Some(Alignment::Right)
        );
        assert_eq!(
            reopened.paragraph(0).unwrap().run(0).unwrap().bold_value(),
            Some(true)
        );
        assert_eq!(
            reopened.table(0).unwrap().cell(0, 0).unwrap().text(),
            "cell"
        );
        assert_eq!(reopened.images().len(), 1);
    }

    #[test]
    fn odt_writer_preserves_small_positive_image_dimensions_exactly() {
        let mut document = Document::new();
        for dimension in [1_i64, 12_699, 12_700, 914_399] {
            document.add_picture(
                PNG,
                &format!("image-{dimension}.png"),
                Length::emu(dimension),
                Length::emu(dimension),
            );
        }

        let written = document.to_odt_bytes().unwrap();
        let reopened = Document::from_odt_bytes(&written.bytes).unwrap().document;
        assert_eq!(
            reopened
                .images()
                .iter()
                .map(|image| (image.width_emu, image.height_emu))
                .collect::<Vec<_>>(),
            [
                (1, 1),
                (12_699, 12_699),
                (12_700, 12_700),
                (914_399, 914_399)
            ]
        );
    }

    #[test]
    fn odt_writer_preserves_cell_deep_and_inherited_lists() {
        let mut deep = Document::new();
        let deep_list = deep.add_list_definition(&[ListLevel::bullet()]);
        let numbering = deep.numbering.as_mut().unwrap();
        let abstract_num_id = numbering
            .nums
            .iter()
            .find(|instance| instance.num_id == deep_list)
            .unwrap()
            .abstract_num_id;
        numbering
            .abstract_nums
            .iter_mut()
            .find(|definition| definition.abstract_num_id == abstract_num_id)
            .unwrap()
            .levels
            .retain(|level| level.ilvl != 8);
        deep.add_paragraph("deep first").set_numbering(deep_list, 8);
        let written = deep.to_odt_bytes().unwrap();
        assert!(written.diagnostics.iter().any(|diagnostic| {
            diagnostic.path == format!("numbering[{deep_list}]/level[8]")
                && diagnostic.message
                    == "undefined numbering level was exported as decimal ODT list"
        }));
        let reopened = Document::from_odt_bytes(&written.bytes).unwrap().document;
        assert_eq!(reopened.paragraph_count(), 1);
        assert_eq!(reopened.paragraph(0).unwrap().text(), "deep first");
        assert_eq!(reopened.paragraph(0).unwrap().numbering().unwrap().1, 8);

        let mut inherited = Document::new();
        let inherited_list = inherited.add_list_definition(&[ListLevel::decimal()]);
        inherited.add_style(
            crate::StyleBuilder::paragraph("InheritedList", "Inherited List").paragraph_properties(
                CT_PPr {
                    num_id: Some(inherited_list),
                    num_ilvl: Some(1),
                    ..Default::default()
                },
            ),
        );
        inherited
            .add_paragraph("styled list")
            .set_style("InheritedList");
        let reopened = Document::from_odt_bytes(&inherited.to_odt_bytes().unwrap().bytes)
            .unwrap()
            .document;
        assert_eq!(reopened.paragraph_count(), 1);
        assert_eq!(reopened.paragraph(0).unwrap().numbering().unwrap().1, 1);

        let mut cell_document = Document::new();
        let cell_list = cell_document.add_list_definition(&[ListLevel::bullet()]);
        cell_document
            .add_table(1, 1)
            .cell(0, 0)
            .unwrap()
            .set_text("cell item");
        let BodyContent::Table(table) = &mut cell_document.document.body.content[0] else {
            unreachable!();
        };
        let CellContent::Paragraph(paragraph) = &mut table.rows[0].cells[0].content[0] else {
            unreachable!();
        };
        Paragraph { inner: paragraph }.set_numbering(cell_list, 0);
        let reopened = Document::from_odt_bytes(&cell_document.to_odt_bytes().unwrap().bytes)
            .unwrap()
            .document;
        let table = reopened.table(0).unwrap();
        let cell = table.cell(0, 0).unwrap();
        assert_eq!(cell.paragraph_count(), 1);
        assert_eq!(cell.paragraph(0).unwrap().text(), "cell item");
        assert_eq!(cell.paragraph(0).unwrap().numbering().unwrap().1, 0);
    }

    #[test]
    fn odt_writer_diagnoses_continuation_content_and_all_widths() {
        let mut document = Document::new();
        let relationship = document.embed_image(PNG, "hidden.png");
        {
            let mut table = document.add_table(2, 1);
            table.cell(0, 0).unwrap().set_text("visible");
            table.cell(0, 0).unwrap().set_width(Length::twips(2_000));
            table.cell(0, 0).unwrap().set_v_merge_restart();
            table.cell(1, 0).unwrap().set_text("hidden");
            table.cell(1, 0).unwrap().set_width(Length::twips(2_000));
            table
                .cell(1, 0)
                .unwrap()
                .add_picture(&relationship, Length::emu(1), Length::emu(1));
            table.cell(1, 0).unwrap().set_v_merge_continue();
        }
        let BodyContent::Table(table) = &mut document.document.body.content[0] else {
            unreachable!();
        };
        table.rows[0].properties = Some(rdocx_oxml::table::CT_TrPr {
            cant_split: Some(true),
            ..Default::default()
        });

        let written = document.to_odt_bytes().unwrap();
        let paths = written
            .diagnostics
            .iter()
            .map(|diagnostic| diagnostic.path.as_str())
            .collect::<BTreeSet<_>>();
        for expected in [
            "body[0]/tblPr",
            "body[0]/tblGrid",
            "body[0]/row[0]/trPr",
            "body[0]/row[0]/cell[0]/tcPr",
            "body[0]/row[1]/cell[0]/tcPr",
            "body[0]/row[1]/cell[0]/content",
        ] {
            assert!(paths.contains(expected), "missing diagnostic at {expected}");
        }
        let mut archive = ZipArchive::new(Cursor::new(&written.bytes)).unwrap();
        assert_eq!(
            archive.len(),
            3,
            "unreferenced continuation media was packaged"
        );
        assert!(archive.by_name("Pictures/image1.png").is_err());
    }

    #[test]
    fn odt_writer_reports_inherited_losses_and_normalizes_crlf() {
        let mut document = Document::new();
        document.add_style(
            crate::StyleBuilder::paragraph("LossyStyle", "Lossy Style")
                .paragraph_properties(CT_PPr {
                    keep_next: Some(true),
                    ..Default::default()
                })
                .run_properties(CT_RPr {
                    font_ascii: Some("Liberation Serif".to_owned()),
                    font_hansi: Some("Liberation Serif".to_owned()),
                    font_east_asia: Some("Liberation Serif".to_owned()),
                    font_cs: Some("Liberation Serif".to_owned()),
                    caps: Some(true),
                    ..Default::default()
                }),
        );
        let mut paragraph = document.add_paragraph("a\r\nb\rc\n");
        paragraph.set_style("LossyStyle");

        let written = document.to_odt_bytes().unwrap();
        assert!(written.diagnostics.iter().any(|diagnostic| {
            diagnostic.path == "body[0]/pPr"
                && diagnostic.message
                    == "unsupported paragraph properties were dropped during ODT export"
        }));
        assert!(written.diagnostics.iter().any(|diagnostic| {
            diagnostic.path == "body[0]/run[0]/rPr"
                && diagnostic.message == "unsupported run properties were dropped during ODT export"
        }));
        let reopened = Document::from_odt_bytes(&written.bytes).unwrap().document;
        assert_eq!(reopened.paragraph(0).unwrap().text(), "a\nb\nc\n");
    }

    #[test]
    fn odt_writer_rejects_grid_span_overflow_without_panicking() {
        let mut document = Document::new();
        let mut table = CT_Tbl::new();
        let mut row = CT_Row::new();
        for _ in 0..2 {
            let mut cell = CT_Tc::new();
            cell.properties = Some(CT_TcPr {
                grid_span: Some(u32::MAX),
                ..Default::default()
            });
            row.cells.push(cell);
        }
        table.rows.push(row);
        document
            .document
            .body
            .content
            .push(BodyContent::Table(table));
        assert!(document.to_odt_bytes().is_err());
    }

    #[test]
    fn unsupported_document_content_is_diagnosed_without_dropping_supported_odt_siblings() {
        fn content_control(text: &str) -> rdocx_oxml::content_control::CT_Sdt {
            let xml = format!(
                r#"<w:sdt xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:sdtContent><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:sdtContent></w:sdt>"#
            );
            let mut reader = quick_xml::Reader::from_str(&xml);
            let mut buffer = Vec::new();
            match reader.read_event_into(&mut buffer).unwrap() {
                Event::Start(start) => {
                    rdocx_oxml::content_control::CT_Sdt::from_xml(&mut reader, &start).unwrap()
                }
                event => panic!("expected content control, got {event:?}"),
            }
        }

        fn revision_paragraph() -> CT_P {
            let xml = r#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:ins w:id="1" w:author="A"><w:r><w:t>revision</w:t></w:r></w:ins></w:p>"#;
            let mut reader = quick_xml::Reader::from_str(xml);
            let mut buffer = Vec::new();
            match reader.read_event_into(&mut buffer).unwrap() {
                Event::Start(_) => CT_P::from_xml(&mut reader).unwrap(),
                event => panic!("expected paragraph, got {event:?}"),
            }
        }

        let mut document = Document::new();
        document.set_landscape();
        document.set_header("header story");
        document.set_footer("footer story");
        document.document.background_xml = Some(
            br#"<w:background xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:color="123456"/>"#
                .to_vec(),
        );
        let defaults = document
            .styles
            .doc_defaults
            .as_mut()
            .and_then(|defaults| defaults.rpr.as_mut())
            .unwrap();
        defaults.font_cs = Some("Calibri".to_owned());

        document.add_paragraph("before");
        let BodyContent::Paragraph(paragraph) = &mut document.document.body.content[0] else {
            unreachable!();
        };
        paragraph.properties = Some(CT_PPr {
            style_id: Some("AppliedParagraphStyle".to_owned()),
            line_rule: Some("exact".to_owned()),
            num_ilvl: Some(2),
            ind_first_line: Some(Twips(120)),
            ind_hanging: Some(Twips(60)),
            keep_next: Some(true),
            ..Default::default()
        });
        paragraph.runs[0]
            .content
            .push(RunContent::DeletedText(rdocx_oxml::text::CT_Text::new(
                "deleted",
            )));
        paragraph.runs[0]
            .content
            .push(RunContent::Break(BreakType::Page));
        paragraph.runs[0]
            .content
            .push(RunContent::Break(BreakType::Column));
        paragraph.runs[0]
            .extra_xml
            .push(b"<w:retainedRun/>".to_vec());
        paragraph.runs[0].extra_xml_positions.push(1);
        paragraph.hyperlinks.push(rdocx_oxml::text::HyperlinkSpan {
            rel_id: None,
            anchor: Some("bookmark".to_owned()),
            run_start: 0,
            run_end: 1,
            extra_attributes: Vec::new(),
            extra_xml: Vec::new(),
            preserved_raw_before: None,
        });
        paragraph.runs[0].properties = Some(CT_RPr {
            style_id: Some("AppliedRunStyle".to_owned()),
            font_ascii: Some("Calibri".to_owned()),
            font_hansi: Some("Calibri".to_owned()),
            font_east_asia: Some("Calibri".to_owned()),
            font_cs: Some("Calibri".to_owned()),
            caps: Some(true),
            color: Some("auto".to_owned()),
            highlight: Some(ST_HighlightColor::Yellow),
            shading: Some(rdocx_oxml::properties::CT_Shd {
                val: "horzStripe".to_owned(),
                color: Some("ABCDEF".to_owned()),
                fill: Some("123456".to_owned()),
            }),
            ..Default::default()
        });
        paragraph.runs.push({
            let mut run = CT_R::new("");
            run.content = vec![RunContent::Field(rdocx_oxml::text::Field::new("PAGE", "7"))];
            run
        });
        paragraph.runs.push({
            let mut run = CT_R::new("");
            run.content = vec![RunContent::FootnoteRef { id: 2 }];
            run
        });
        paragraph.runs.push({
            let mut run = CT_R::new("");
            run.content = vec![RunContent::EndnoteRef { id: 3 }];
            run
        });
        paragraph.runs.push({
            let mut run = CT_R::new("");
            run.content = vec![RunContent::CommentReference {
                id: 4,
                raw_before: 0,
            }];
            run
        });
        paragraph
            .comment_ranges
            .push(rdocx_oxml::text::CommentRangeMarker::Start {
                id: 4,
                run_index: 0,
                raw_before: 0,
            });
        assert!(paragraph.insert_bookmark_start(0, 5, "bookmark"));
        assert!(paragraph.insert_bookmark_end(paragraph.runs.len(), 5));
        paragraph
            .content_controls
            .push((0, 0, 0, content_control("run control")));

        document
            .document
            .body
            .content
            .push(BodyContent::Paragraph(revision_paragraph()));
        document
            .document
            .body
            .content
            .push(BodyContent::RawXml(b"<w:custom/>".to_vec()));
        document
            .document
            .body
            .content
            .push(BodyContent::ContentControl(content_control("body control")));
        {
            let mut table = document.add_table(1, 1);
            table.cell(0, 0).unwrap().set_text("cell");
            table.cell(0, 0).unwrap().set_width(Length::twips(1_500));
            table.cell(0, 0).unwrap().set_no_wrap();
        }
        let BodyContent::Table(table) = &mut document.document.body.content[4] else {
            unreachable!();
        };
        table.extra_xml.push((0, b"<w:tableRaw/>".to_vec()));
        table
            .content_controls
            .push((0, 0, content_control("row control")));
        table.rows[0].properties = Some(rdocx_oxml::table::CT_TrPr {
            header: Some(true),
            ..Default::default()
        });
        table.rows[0].extra_xml.push((0, b"<w:rowRaw/>".to_vec()));
        table.rows[0]
            .content_controls
            .push((0, 0, content_control("cell control")));
        table.rows[0].cells[0]
            .extra_xml
            .push((0, b"<w:cellRaw/>".to_vec()));
        document.add_paragraph("after");
        document.add_picture(
            b"unsupported media",
            "unsupported.bin",
            Length::emu(12_700),
            Length::emu(12_700),
        );
        let BodyContent::Paragraph(unsupported_image) = &mut document.document.body.content[6]
        else {
            unreachable!();
        };
        let RunContent::Drawing(drawing) = &mut unsupported_image.runs[0].content[0] else {
            unreachable!();
        };
        let inline = drawing.inline.as_mut().unwrap();
        inline.description = Some("alternative text".to_string());
        inline.name = Some("source image name".to_string());
        inline.raw_xml = Some(
            br#"<wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"/>"#
                .to_vec(),
        );
        document.add_anchored_image(
            PNG,
            "anchored.png",
            Length::emu(12_700),
            Length::emu(12_700),
            false,
        );

        let before = document.to_bytes().unwrap();
        let written = document.to_odt_bytes().unwrap();
        let after = document.to_bytes().unwrap();
        assert_eq!(before, after);
        let actual = written
            .diagnostics
            .iter()
            .map(|diagnostic| (diagnostic.path.as_str(), diagnostic.message.as_str()))
            .collect::<Vec<_>>();
        let expected = vec![
            (
                "document/background",
                "document background was dropped during ODT export",
            ),
            (
                "document/sectPr",
                "final section properties were dropped during ODT export",
            ),
            (
                "document/headers",
                "header stories were dropped during ODT export",
            ),
            (
                "document/footers",
                "footer stories were dropped during ODT export",
            ),
            (
                "body[0]/run[0]/content[0]",
                "anchored drawing was dropped during ODT export",
            ),
            (
                "body[1]/pPr",
                "unsupported paragraph properties were dropped during ODT export",
            ),
            (
                "body[1]/pPr/pStyle",
                "paragraph style identity was materialized and dropped during ODT export",
            ),
            (
                "body[1]/pPr/spacing",
                "line-spacing rule without line spacing was dropped during ODT export",
            ),
            (
                "body[1]/pPr/numPr/ilvl",
                "numbering level without a numbering id was dropped during ODT export",
            ),
            (
                "body[1]/pPr/ind/hanging",
                "hanging indent was dropped because first-line indent takes precedence during ODT export",
            ),
            (
                "body[1]/raw[0]",
                "unmodelled paragraph XML was dropped during ODT export",
            ),
            (
                "body[1]/raw[5]",
                "unmodelled paragraph XML was dropped during ODT export",
            ),
            (
                "body[1]/content-controls",
                "run content controls were dropped during ODT export",
            ),
            (
                "body[1]/comments",
                "comment markers were dropped during ODT export",
            ),
            (
                "body[1]/bookmarks",
                "bookmark markers were dropped during ODT export",
            ),
            (
                "body[1]/hyperlinks",
                "hyperlink wrappers were flattened during ODT export",
            ),
            (
                "body[1]/run[0]/rPr",
                "unsupported run properties were dropped during ODT export",
            ),
            (
                "body[1]/run[0]/rPr/rStyle",
                "run style identity was materialized and dropped during ODT export",
            ),
            (
                "body[1]/run[0]/raw",
                "unmodelled run XML was dropped during ODT export",
            ),
            (
                "body[1]/run[0]/rPr/color",
                "unsupported run color was dropped during ODT export",
            ),
            (
                "body[1]/run[0]/rPr/shading-pattern",
                "run shading pattern was simplified during ODT export",
            ),
            (
                "body[1]/run[0]/rPr/shading-color",
                "run shading foreground color was dropped during ODT export",
            ),
            (
                "body[1]/run[0]/rPr/highlight",
                "run highlight was replaced by shading fill during ODT export",
            ),
            (
                "body[1]/run[0]/content[1]",
                "deleted text was flattened during ODT export",
            ),
            (
                "body[1]/run[0]/content[2]",
                "unsupported break type was dropped during ODT export",
            ),
            (
                "body[1]/run[0]/content[3]",
                "unsupported break type was dropped during ODT export",
            ),
            (
                "body[1]/run[1]/content[0]",
                "field was flattened during ODT export",
            ),
            (
                "body[1]/run[2]/content[0]",
                "footnote reference was dropped during ODT export",
            ),
            (
                "body[1]/run[3]/content[0]",
                "endnote reference was dropped during ODT export",
            ),
            (
                "body[1]/run[4]/content[0]",
                "comment reference was dropped during ODT export",
            ),
            (
                "body[2]/raw[0]",
                "unmodelled paragraph XML was dropped during ODT export",
            ),
            (
                "body[2]/revisions",
                "paragraph revisions were flattened during ODT export",
            ),
            (
                "body[3]",
                "unmodelled body XML was dropped during ODT export",
            ),
            (
                "body[4]",
                "body content control was dropped during ODT export",
            ),
            (
                "body[5]/tblPr",
                "table properties were dropped during ODT export",
            ),
            (
                "body[5]/tblGrid",
                "table grid column widths were dropped during ODT export",
            ),
            (
                "body[5]/raw[0]",
                "unmodelled table XML was dropped during ODT export",
            ),
            (
                "body[5]/content-controls",
                "table row content controls were dropped during ODT export",
            ),
            (
                "body[5]/row[0]/trPr",
                "table-row properties were dropped during ODT export",
            ),
            (
                "body[5]/row[0]/raw[0]",
                "unmodelled table-row XML was dropped during ODT export",
            ),
            (
                "body[5]/row[0]/content-controls",
                "table cell content controls were dropped during ODT export",
            ),
            (
                "body[5]/row[0]/cell[0]/tcPr",
                "unsupported table-cell properties were dropped during ODT export",
            ),
            (
                "body[5]/row[0]/cell[0]/raw[0]",
                "unmodelled table-cell XML was dropped during ODT export",
            ),
            (
                "body[7]/run[0]/content[0]",
                "inline image description was dropped during ODT export",
            ),
            (
                "body[7]/run[0]/content[0]",
                "inline image name was replaced during ODT export",
            ),
            (
                "body[7]/run[0]/content[0]",
                "retained inline drawing XML was dropped during ODT export",
            ),
            (
                "body[7]/run[0]/content[0]",
                "unsupported inline image was dropped during ODT export",
            ),
        ];
        assert_eq!(actual, expected);
        let reopened = Document::from_odt_bytes(&written.bytes).unwrap().document;
        assert!(reopened.text().contains("before"));
        assert!(reopened.text().contains("after"));
        assert_eq!(
            reopened.paragraph(1).unwrap().first_line_indent(),
            Some(Length::twips(120))
        );

        let mut numbered = Document::new();
        let list = numbered
            .add_list_definition(&[
                ListLevel::new(crate::document::ListNumberFormat::UpperRoman).start(3),
            ]);
        numbered.add_paragraph("roman").set_numbering(list, 0);
        let numbering = numbered.numbering.as_mut().unwrap();
        numbering
            .root_attributes
            .push(("xmlns:retained".to_string(), "urn:retained".to_string()));
        numbering.extra_xml.push((
            0,
            b"<retained:root xmlns:retained=\"urn:retained\"/>".to_vec(),
        ));
        let instance = numbering
            .nums
            .iter_mut()
            .find(|instance| instance.num_id == list)
            .unwrap();
        let abstract_num_id = instance.abstract_num_id;
        instance
            .extra_attributes
            .push(("retained:attribute".to_string(), "value".to_string()));
        instance.extra_xml.push((
            0,
            br#"<w:lvlOverride xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ilvl="0"/>"#
                .to_vec(),
        ));
        let abstract_num = numbering
            .abstract_nums
            .iter_mut()
            .find(|abstract_num| abstract_num.abstract_num_id == abstract_num_id)
            .unwrap();
        abstract_num
            .extra_attributes
            .push(("retained:abstract".to_string(), "value".to_string()));
        abstract_num.extra_xml.push((
            0,
            b"<retained:abstractChild xmlns:retained=\"urn:retained\"/>".to_vec(),
        ));
        let level = abstract_num
            .levels
            .iter_mut()
            .find(|level| level.ilvl == 0)
            .unwrap();
        level.suffix = Some(rdocx_oxml::numbering::ST_LvlSuffix::Space);
        level.rpr = Some(CT_RPr {
            bold: Some(true),
            ..Default::default()
        });
        level
            .extra_attributes
            .push(("retained:level".to_string(), "value".to_string()));
        level.extra_xml.push((
            0,
            b"<retained:levelChild xmlns:retained=\"urn:retained\"/>".to_vec(),
        ));
        let diagnostics = numbered.to_odt_bytes().unwrap().diagnostics;
        assert_eq!(
            diagnostics
                .iter()
                .map(|diagnostic| (diagnostic.path.as_str(), diagnostic.message.as_str()))
                .collect::<Vec<_>>(),
            vec![
                (
                    "numbering[1]/root-attributes",
                    "retained numbering root attributes were dropped during ODT export",
                ),
                (
                    "numbering[1]/root-xml",
                    "retained numbering root XML was dropped during ODT export",
                ),
                (
                    "numbering[1]/instance-attributes",
                    "retained numbering instance attributes were dropped during ODT export",
                ),
                (
                    "numbering[1]/instance-overrides",
                    "numbering instance overrides or retained XML were dropped during ODT export",
                ),
                (
                    "numbering[1]/abstract-type",
                    "abstract numbering type metadata was dropped during ODT export",
                ),
                (
                    "numbering[1]/abstract-attributes",
                    "retained abstract numbering attributes were dropped during ODT export",
                ),
                (
                    "numbering[1]/abstract-xml",
                    "retained abstract numbering XML was dropped during ODT export",
                ),
                (
                    "body[0]/run[0]/rPr",
                    "unsupported run properties were dropped during ODT export",
                ),
                (
                    "numbering[1]/level[0]",
                    "custom list start value was reset during ODT export",
                ),
                (
                    "numbering[1]/level[0]",
                    "list marker suffix was dropped during ODT export",
                ),
                (
                    "numbering[1]/level[0]",
                    "list marker text or bullet glyph was simplified during ODT export",
                ),
                (
                    "numbering[1]/level[0]",
                    "list marker justification was dropped during ODT export",
                ),
                (
                    "numbering[1]/level[0]",
                    "list level paragraph formatting was dropped during ODT export",
                ),
                (
                    "numbering[1]/level[0]",
                    "list marker run formatting was dropped during ODT export",
                ),
                (
                    "numbering[1]/level[0]",
                    "retained list level XML or attributes were dropped during ODT export",
                ),
                (
                    "numbering[1]/level[0]",
                    "non-decimal numbering format was exported as decimal ODT list",
                ),
            ]
        );
    }

    #[test]
    fn save_odt_preserves_existing_destination_when_staging_fails() {
        let root = std::env::temp_dir().join(format!(
            "rdocx-save-odt-failure-{}-{}",
            std::process::id(),
            std::thread::current().name().unwrap_or("test")
        ));
        let _ = std::fs::remove_dir_all(&root);
        std::fs::create_dir_all(&root).unwrap();
        let destination = root.join("existing.odt");
        let original = b"existing destination bytes";
        std::fs::write(&destination, original).unwrap();
        for attempt in 0..128_u8 {
            let staging = root.join(format!(
                ".existing.odt.rdocx-{}-{attempt}.tmp",
                std::process::id()
            ));
            std::fs::write(staging, b"occupied staging path").unwrap();
        }

        let mut document = Document::new();
        document.add_paragraph("preserve destination");
        assert!(document.save_odt(&destination).is_err());
        assert_eq!(std::fs::read(&destination).unwrap(), original);
        std::fs::remove_dir_all(root).unwrap();
    }
}
