//! Bounded RTF 1.9.1 reader for the subset emitted by Microsoft Word.

use std::borrow::Cow;
use std::collections::{BTreeMap, BTreeSet, HashMap};
use std::io::{Read, Write};
use std::path::Path;

use encoding_rs::Encoding;
use rdocx_oxml::document::BodyContent;
use rdocx_oxml::numbering::ST_NumberFormat;
use rdocx_oxml::properties::{CT_PPr, CT_RPr};
use rdocx_oxml::shared::ST_Jc;
use rdocx_oxml::table::{CT_Row, CT_Tbl, CT_TblPr, CT_TblWidth, CT_TcPr, CT_TrPr, CellContent};
use rdocx_oxml::text::{BreakType, CT_P, CT_R, RunContent};

use crate::{Alignment, Document, Error, Length, ListLevel, ListNumberFormat, Result};

const MAX_INPUT_BYTES: usize = 64 * 1024 * 1024;
const MAX_GROUP_DEPTH: usize = 256;
const MAX_LOOKUP_ENTRIES: usize = 4096;
const MAX_TABLE_ROWS: usize = 10_000;
const MAX_TABLE_COLUMNS: usize = 256;
const MAX_PICTURE_BYTES: usize = 64 * 1024 * 1024;
const MAX_DIAGNOSTICS: usize = 10_000;
const MAX_BLOCKS: usize = 100_000;
const MAX_RUNS: usize = 100_000;
const MAX_TABLE_CELLS: usize = 50_000;
const MAX_RETAINED_OUTPUT_BYTES: usize = 64 * 1024 * 1024;

/// One stable report of RTF content that could not be represented in DOCX.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RtfDiagnostic {
    pub offset: usize,
    pub destination: Option<String>,
    pub message: String,
}

/// A converted document together with every lossy-conversion diagnostic.
pub struct RtfReadResult {
    pub document: Document,
    pub diagnostics: Vec<RtfDiagnostic>,
}

/// Serialized RTF bytes together with every lossy-conversion diagnostic.
pub struct RtfWriteResult {
    pub bytes: Vec<u8>,
    pub diagnostics: Vec<RtfDiagnostic>,
}

impl Document {
    /// Convert an RTF byte stream into an editable document.
    pub fn from_rtf_bytes(bytes: &[u8]) -> Result<RtfReadResult> {
        Parser::parse(bytes)?.project()
    }

    /// Open and convert an RTF file into an editable document.
    pub fn open_rtf<P: AsRef<Path>>(path: P) -> Result<RtfReadResult> {
        let mut file = std::fs::File::open(path)?;
        let declared_len = file.metadata()?.len();
        if declared_len > MAX_INPUT_BYTES as u64 {
            return Err(rtf_error(0, "RTF input exceeds the size limit"));
        }
        let mut bytes = Vec::with_capacity(declared_len as usize);
        let mut chunk = [0_u8; 8192];
        while bytes.len() < MAX_INPUT_BYTES {
            let remaining = MAX_INPUT_BYTES - bytes.len();
            let read_len = remaining.min(chunk.len());
            let read = file.read(&mut chunk[..read_len])?;
            if read == 0 {
                return Self::from_rtf_bytes(&bytes);
            }
            if bytes.capacity() < bytes.len() + read {
                bytes.reserve_exact(bytes.len() + read - bytes.capacity());
            }
            bytes.extend_from_slice(&chunk[..read]);
        }
        if file.read(&mut chunk[..1])? != 0 {
            return Err(rtf_error(0, "RTF input exceeds the size limit"));
        }
        Self::from_rtf_bytes(&bytes)
    }

    /// Serialize the editable document to the supported RTF subset.
    pub fn to_rtf_bytes(&self) -> Result<RtfWriteResult> {
        RtfWriter::new(self).write()
    }

    /// Serialize and save RTF to a path, returning lossy-conversion diagnostics.
    pub fn save_rtf<P: AsRef<Path>>(&self, path: P) -> Result<Vec<RtfDiagnostic>> {
        let result = self.to_rtf_bytes()?;
        write_atomic_file(path.as_ref(), &result.bytes)?;
        Ok(result.diagnostics)
    }
}

fn write_atomic_file(path: &Path, bytes: &[u8]) -> std::io::Result<()> {
    let parent = path.parent().unwrap_or_else(|| Path::new("."));
    let file_name = path.file_name().ok_or_else(|| {
        std::io::Error::new(std::io::ErrorKind::InvalidInput, "invalid file name")
    })?;
    for attempt in 0..128_u8 {
        let mut temporary_name = std::ffi::OsString::from(".");
        temporary_name.push(file_name);
        temporary_name.push(format!(".rdocx-{}-{attempt}.tmp", std::process::id()));
        let temporary = parent.join(temporary_name);
        let mut file = match std::fs::OpenOptions::new()
            .write(true)
            .create_new(true)
            .open(&temporary)
        {
            Ok(file) => file,
            Err(error) if error.kind() == std::io::ErrorKind::AlreadyExists => continue,
            Err(error) => return Err(error),
        };
        let result = file.write_all(bytes).and_then(|()| file.sync_all());
        drop(file);
        let result = result.and_then(|()| crate::document::replace_file(&temporary, path));
        if result.is_err() {
            let _ = std::fs::remove_file(&temporary);
        }
        return result;
    }
    Err(std::io::Error::new(
        std::io::ErrorKind::AlreadyExists,
        "could not allocate RTF-save staging file",
    ))
}

#[derive(Clone, Debug)]
struct WriterList {
    list_id: i32,
    override_id: i32,
    levels: Vec<WriterListLevel>,
}

#[derive(Clone, Copy, Debug)]
struct WriterListLevel {
    format: Option<ListNumberFormat>,
    start: u32,
}

struct BoundedOutput {
    bytes: Vec<u8>,
    limit: usize,
}

impl BoundedOutput {
    fn new(limit: usize) -> Self {
        Self {
            bytes: Vec::new(),
            limit,
        }
    }

    fn into_vec(self) -> Vec<u8> {
        self.bytes
    }

    fn check_available(&self, additional: usize) -> std::io::Result<()> {
        let next = self
            .bytes
            .len()
            .checked_add(additional)
            .ok_or_else(output_limit_error)?;
        if next > self.limit {
            return Err(output_limit_error());
        }
        Ok(())
    }

    fn push(&mut self, byte: u8) -> Result<()> {
        self.write_all(&[byte])?;
        Ok(())
    }
}

impl Write for BoundedOutput {
    fn write(&mut self, buf: &[u8]) -> std::io::Result<usize> {
        self.write_all(buf)?;
        Ok(buf.len())
    }

    fn write_all(&mut self, buf: &[u8]) -> std::io::Result<()> {
        self.check_available(buf.len())?;
        self.bytes.extend_from_slice(buf);
        Ok(())
    }

    fn flush(&mut self) -> std::io::Result<()> {
        Ok(())
    }
}

fn output_limit_error() -> std::io::Error {
    std::io::Error::other("RTF output exceeds the size limit")
}

struct RtfWriter<'a> {
    document: &'a Document,
    fonts: BTreeMap<String, usize>,
    colors: BTreeMap<String, usize>,
    lists: Vec<WriterList>,
    list_indexes: BTreeMap<u32, usize>,
    diagnostics: Vec<RtfDiagnostic>,
    output_limit: usize,
}

impl<'a> RtfWriter<'a> {
    fn new(document: &'a Document) -> Self {
        Self::new_with_output_limit(document, MAX_RETAINED_OUTPUT_BYTES)
    }

    fn new_with_output_limit(document: &'a Document, output_limit: usize) -> Self {
        let mut fonts = BTreeMap::new();
        fonts.insert("Calibri".to_owned(), 0);
        Self {
            document,
            fonts,
            colors: BTreeMap::new(),
            lists: Vec::new(),
            list_indexes: BTreeMap::new(),
            diagnostics: Vec::new(),
            output_limit,
        }
    }

    fn write(mut self) -> Result<RtfWriteResult> {
        for (index, content) in self.document.document.body.content.iter().enumerate() {
            self.scan_body_content(content, format!("body[{index}]"));
        }

        let mut output = BoundedOutput::new(self.output_limit);
        write!(output, "{{\\rtf1\\ansi\\deff0")?;
        self.write_font_table(&mut output)?;
        self.write_color_table(&mut output)?;
        self.write_list_tables(&mut output)?;

        for (index, content) in self.document.document.body.content.iter().enumerate() {
            self.write_body_content(&mut output, content, &format!("body[{index}]"))?;
        }
        output.push(b'}')?;
        let output = output.into_vec();
        Ok(RtfWriteResult {
            bytes: output,
            diagnostics: self.diagnostics,
        })
    }

    fn scan_body_content(&mut self, content: &BodyContent, location: String) {
        match content {
            BodyContent::Paragraph(paragraph) => self.scan_paragraph(paragraph, location),
            BodyContent::Table(table) => self.scan_table(table, location),
            BodyContent::ContentControl(_) => self.diagnose(
                &location,
                "body content control was dropped during RTF export",
            ),
            BodyContent::RawXml(_) => self.diagnose(
                &location,
                "unmodelled body XML was dropped during RTF export",
            ),
        }
    }

    fn scan_table(&mut self, table: &CT_Tbl, location: String) {
        if let Some(properties) = &table.properties {
            self.scan_table_properties(table, properties, &location);
        }
        for (index, _) in &table.extra_xml {
            self.diagnose(
                &format!("{location}/raw[{index}]"),
                "unmodelled table XML was dropped during RTF export",
            );
        }
        for (row_index, row) in table.rows.iter().enumerate() {
            if let Some(properties) = &row.properties {
                self.scan_row_properties(properties, &format!("{location}/row[{row_index}]"));
            }
            for (index, _) in &row.extra_xml {
                self.diagnose(
                    &format!("{location}/row[{row_index}]/raw[{index}]"),
                    "unmodelled table-row XML was dropped during RTF export",
                );
            }
            for (cell_index, cell) in row.cells.iter().enumerate() {
                self.scan_cell_width(
                    table,
                    row,
                    cell_index,
                    &format!("{location}/row[{row_index}]/cell[{cell_index}]"),
                );
                if let Some(properties) = &cell.properties {
                    self.scan_cell_properties(
                        properties,
                        &format!("{location}/row[{row_index}]/cell[{cell_index}]"),
                    );
                }
                for (index, _) in &cell.extra_xml {
                    self.diagnose(
                        &format!("{location}/row[{row_index}]/cell[{cell_index}]/raw[{index}]"),
                        "unmodelled table-cell XML was dropped during RTF export",
                    );
                }
                for (content_index, content) in cell.content.iter().enumerate() {
                    let cell_location = format!(
                        "{location}/row[{row_index}]/cell[{cell_index}]/content[{content_index}]"
                    );
                    match content {
                        CellContent::Paragraph(paragraph) => {
                            self.scan_paragraph(paragraph, cell_location)
                        }
                        CellContent::Table(_) => self
                            .diagnose(&cell_location, "nested table was dropped during RTF export"),
                        CellContent::ContentControl(_) => self.diagnose(
                            &cell_location,
                            "table-cell content control was dropped during RTF export",
                        ),
                    }
                }
            }
        }
    }

    fn scan_table_properties(&mut self, table: &CT_Tbl, properties: &CT_TblPr, location: &str) {
        if properties.style_id.is_some() {
            self.diagnose(
                &format!("{location}/tblPr/tblStyle"),
                "table style was dropped during RTF export",
            );
        }
        if properties
            .width
            .as_ref()
            .is_some_and(|width| !table_width_is_preserved_by_grid(table, width))
        {
            self.diagnose(
                &format!("{location}/tblPr/tblW"),
                "table width was dropped during RTF export",
            );
        }
        if properties.jc.is_some() {
            self.diagnose(
                &format!("{location}/tblPr/jc"),
                "table alignment was dropped during RTF export",
            );
        }
        if properties.borders.is_some() {
            self.diagnose(
                &format!("{location}/tblPr/tblBorders"),
                "table borders were dropped during RTF export",
            );
        }
        if properties.cell_margin.is_some() {
            self.diagnose(
                &format!("{location}/tblPr/tblCellMar"),
                "table cell margins were dropped during RTF export",
            );
        }
        if properties.layout.is_some() {
            self.diagnose(
                &format!("{location}/tblPr/tblLayout"),
                "table layout was dropped during RTF export",
            );
        }
        if properties.indent.is_some() {
            self.diagnose(
                &format!("{location}/tblPr/tblInd"),
                "table indent was dropped during RTF export",
            );
        }
        if properties.shading.is_some() {
            self.diagnose(
                &format!("{location}/tblPr/shd"),
                "table shading was dropped during RTF export",
            );
        }
        if properties.look.is_some() {
            self.diagnose(
                &format!("{location}/tblPr/tblLook"),
                "table look was dropped during RTF export",
            );
        }
        if properties.change.is_some() {
            self.diagnose(
                &format!("{location}/tblPr/tblPrChange"),
                "table property revision was dropped during RTF export",
            );
        }
        if !properties.revision_xml.is_empty() {
            self.diagnose(
                &format!("{location}/tblPr/revisionXml"),
                "unmodelled table property revision XML was dropped during RTF export",
            );
        }
    }

    fn scan_row_properties(&mut self, properties: &CT_TrPr, location: &str) {
        if properties.height.is_some() {
            self.diagnose(
                &format!("{location}/trPr/trHeight"),
                "table-row height was dropped during RTF export",
            );
        }
        if properties.height_rule.is_some() {
            self.diagnose(
                &format!("{location}/trPr/hRule"),
                "table-row height rule was dropped during RTF export",
            );
        }
        if properties.header.is_some() {
            self.diagnose(
                &format!("{location}/trPr/tblHeader"),
                "table-row repeat header property was dropped during RTF export",
            );
        }
        if properties.jc.is_some() {
            self.diagnose(
                &format!("{location}/trPr/jc"),
                "table-row alignment was dropped during RTF export",
            );
        }
        if properties.cant_split.is_some() {
            self.diagnose(
                &format!("{location}/trPr/cantSplit"),
                "table-row cant-split property was dropped during RTF export",
            );
        }
        if properties.cnf_style.is_some() {
            self.diagnose(
                &format!("{location}/trPr/cnfStyle"),
                "table-row conditional style property was dropped during RTF export",
            );
        }
        if !properties.revision_markers.is_empty() {
            self.diagnose(
                &format!("{location}/trPr/revisions"),
                "table-row revision markers were dropped during RTF export",
            );
        }
        if !properties.revision_xml.is_empty() {
            self.diagnose(
                &format!("{location}/trPr/revisionXml"),
                "unmodelled table-row revision XML was dropped during RTF export",
            );
        }
    }

    fn scan_cell_properties(&mut self, properties: &CT_TcPr, location: &str) {
        if let Some(width) = &properties.width {
            if width.width_type != "dxa" {
                self.diagnose(
                    &format!("{location}/tcPr/tcW"),
                    "unsupported table-cell width type was dropped during RTF export",
                );
            } else if width.w <= 0 {
                self.diagnose(
                    &format!("{location}/tcPr/tcW"),
                    "invalid table-cell width was dropped during RTF export",
                );
            }
        }
        if properties.v_merge.is_some() {
            self.diagnose(
                &format!("{location}/tcPr/vMerge"),
                "vertical table-cell merge was dropped during RTF export",
            );
        }
        if properties.borders.is_some() {
            self.diagnose(
                &format!("{location}/tcPr/borders"),
                "table-cell borders were dropped during RTF export",
            );
        }
        if properties.shading.is_some() {
            self.diagnose(
                &format!("{location}/tcPr/shading"),
                "table-cell shading was dropped during RTF export",
            );
        }
        if properties.v_align.is_some() {
            self.diagnose(
                &format!("{location}/tcPr/vAlign"),
                "table-cell vertical alignment was dropped during RTF export",
            );
        }
        if properties.no_wrap.is_some() {
            self.diagnose(
                &format!("{location}/tcPr/noWrap"),
                "table-cell no-wrap property was dropped during RTF export",
            );
        }
        if properties.text_direction.is_some() {
            self.diagnose(
                &format!("{location}/tcPr/textDirection"),
                "table-cell text direction was dropped during RTF export",
            );
        }
        if properties.cnf_style.is_some() {
            self.diagnose(
                &format!("{location}/tcPr/cnfStyle"),
                "table-cell conditional style property was dropped during RTF export",
            );
        }
        let mut occurrences = BTreeMap::<usize, usize>::new();
        for (index, raw) in &properties.extra_xml {
            let occurrence = occurrences.entry(*index).or_insert(0);
            let suffix = if *occurrence == 0 {
                String::new()
            } else {
                format!("#{occurrence}")
            };
            let item_name = raw_xml_item_name(raw).unwrap_or("unknown");
            self.diagnose(
                &format!("{location}/tcPr/raw[{index}]{suffix}"),
                &format!(
                    "unmodelled table-cell property {item_name} was dropped during RTF export"
                ),
            );
            *occurrence += 1;
        }
    }

    fn scan_cell_width(&mut self, table: &CT_Tbl, row: &CT_Row, cell_index: usize, location: &str) {
        let cell = &row.cells[cell_index];
        let explicit_width = cell
            .properties
            .as_ref()
            .and_then(|properties| properties.width.as_ref());
        if explicit_width.is_some() {
            return;
        }
        let Some(grid) = &table.grid else {
            self.diagnose(
                &format!("{location}/tcPr/tcW"),
                "table-cell width could not be preserved because the table grid is missing",
            );
            return;
        };
        if row_grid_span_end(row, cell_index).is_none_or(|end| end > grid.columns.len()) {
            self.diagnose(
                &format!("{location}/tcPr/tcW"),
                "table-cell width could not be preserved because the table grid is too short",
            );
        }
    }

    fn scan_paragraph(&mut self, paragraph: &CT_P, location: String) {
        if let Some((num_id, level)) = paragraph_numbering(paragraph) {
            if level > 8 {
                self.diagnose(
                    &format!("{location}/ppr/numPr/ilvl"),
                    "numbering level above 8 was dropped during RTF export",
                );
            } else {
                self.ensure_writer_list(num_id, &location);
            }
        }
        if let Some(properties) = &paragraph.properties {
            self.scan_paragraph_properties(properties, &location);
        }
        let marker_raw_positions = paragraph_marker_raw_positions(paragraph);
        for (index, _) in &paragraph.extra_xml {
            if marker_raw_positions.contains(index) {
                continue;
            }
            self.diagnose(
                &format!("{location}/raw[{index}]"),
                "unmodelled paragraph XML was dropped during RTF export",
            );
        }
        for (index, _, _, _) in &paragraph.content_controls {
            self.diagnose(
                &format!("{location}/content-control[{index}]"),
                "run content control was dropped during RTF export",
            );
        }
        for (index, _, _) in &paragraph.revisions {
            self.diagnose(
                &format!("{location}/revision[{index}]"),
                "paragraph revision wrapper was flattened during RTF export",
            );
        }
        for marker in &paragraph.comment_ranges {
            self.diagnose(
                &format!("{location}/comment-range[{}]", marker_location(marker)),
                "comment range marker was dropped during RTF export",
            );
        }
        for marker in &paragraph.bookmark_markers {
            self.diagnose(
                &format!("{location}/bookmark[{}]", marker.run_index()),
                "bookmark marker was dropped during RTF export",
            );
        }
        for hyperlink in &paragraph.hyperlinks {
            self.diagnose(
                &format!("{location}/hyperlink[{}]", hyperlink.run_start),
                "hyperlink wrapper was flattened during RTF export",
            );
        }
        for (index, run) in paragraph.runs.iter().enumerate() {
            self.scan_run(run, format!("{location}/run[{index}]"));
        }
    }

    fn scan_paragraph_properties(&mut self, properties: &CT_PPr, location: &str) {
        if properties.style_id.is_some() {
            self.diagnose(
                &format!("{location}/ppr/pStyle"),
                "paragraph style was dropped during RTF export",
            );
        }
        if properties.line_rule.is_some() && properties.line_spacing.is_none() {
            self.diagnose(
                &format!("{location}/ppr/lineRule"),
                "paragraph line-spacing rule without line spacing was dropped during RTF export",
            );
        }
        if properties.before_autospacing.is_some() {
            self.diagnose(
                &format!("{location}/ppr/beforeAutospacing"),
                "paragraph before auto-spacing was dropped during RTF export",
            );
        }
        if properties.after_autospacing.is_some() {
            self.diagnose(
                &format!("{location}/ppr/afterAutospacing"),
                "paragraph after auto-spacing was dropped during RTF export",
            );
        }
        if properties.keep_next.is_some() {
            self.diagnose(
                &format!("{location}/ppr/keepNext"),
                "keep-with-next paragraph property was dropped during RTF export",
            );
        }
        if properties.keep_lines.is_some() {
            self.diagnose(
                &format!("{location}/ppr/keepLines"),
                "keep-lines paragraph property was dropped during RTF export",
            );
        }
        if properties.page_break_before.is_some() {
            self.diagnose(
                &format!("{location}/ppr/pageBreakBefore"),
                "page-break-before paragraph property was dropped during RTF export",
            );
        }
        if properties.widow_control.is_some() {
            self.diagnose(
                &format!("{location}/ppr/widowControl"),
                "widow-control paragraph property was dropped during RTF export",
            );
        }
        if properties.suppress_auto_hyphens.is_some() {
            self.diagnose(
                &format!("{location}/ppr/suppressAutoHyphens"),
                "paragraph suppress-auto-hyphens property was dropped during RTF export",
            );
        }
        if properties.outline_lvl.is_some() {
            self.diagnose(
                &format!("{location}/ppr/outlineLvl"),
                "paragraph outline level was dropped during RTF export",
            );
        }
        if properties.borders.is_some() {
            self.diagnose(
                &format!("{location}/ppr/borders"),
                "paragraph borders were dropped during RTF export",
            );
        }
        if properties.tabs.is_some() {
            self.diagnose(
                &format!("{location}/ppr/tabs"),
                "paragraph tab stops were dropped during RTF export",
            );
        }
        if properties.shading.is_some() {
            self.diagnose(
                &format!("{location}/ppr/shading"),
                "paragraph shading was dropped during RTF export",
            );
        }
        if properties.rpr.is_some() {
            self.diagnose(
                &format!("{location}/ppr/rPr"),
                "paragraph mark run properties were dropped during RTF export",
            );
        }
        if properties.num_ilvl.is_some() && properties.num_id.is_none() {
            self.diagnose(
                &format!("{location}/ppr/numPr"),
                "paragraph numbering level without numbering id was dropped during RTF export",
            );
        }
        if properties.sect_pr.is_some() {
            self.diagnose(
                &format!("{location}/ppr/sectPr"),
                "paragraph section properties were dropped during RTF export",
            );
        }
        if properties.numbering_revision.is_some() || !properties.numbering_revision_xml.is_empty()
        {
            self.diagnose(
                &format!("{location}/ppr/numPrChange"),
                "paragraph numbering revision was dropped during RTF export",
            );
        }
        if properties.change.is_some() || !properties.revision_xml.is_empty() {
            self.diagnose(
                &format!("{location}/ppr/pPrChange"),
                "paragraph property revision was dropped during RTF export",
            );
        }
    }

    fn scan_run(&mut self, run: &CT_R, location: String) {
        if let Some(properties) = &run.properties {
            if let Some(font) = run_font(properties) {
                self.ensure_font(font);
            }
            if let Some(color) = properties.color.as_deref() {
                self.ensure_hex_color(color, &format!("{location}/color"));
            }
            if let Some(fill) = properties
                .shading
                .as_ref()
                .and_then(|shading| shading.fill.as_deref())
            {
                self.ensure_hex_color(fill, &format!("{location}/highlight"));
            }
            self.scan_run_properties(properties, &location);
        }
        for index in &run.extra_xml_positions {
            self.diagnose(
                &format!("{location}/raw[{index}]"),
                "unmodelled run XML was dropped during RTF export",
            );
        }
        for (content_index, content) in run.content.iter().enumerate() {
            if let RunContent::Drawing(drawing) = content {
                let content_location = format!("{location}/content[{content_index}]");
                if let Some(inline) = &drawing.inline {
                    match self.document.image_data(&inline.embed_id) {
                        Some(bytes) => {
                            if picture_kind(&bytes).is_none() {
                                self.diagnose(
                                    &content_location,
                                    "unsupported inline image type was dropped during RTF export",
                                );
                            }
                        }
                        None => self.diagnose(
                            &content_location,
                            "unresolved inline image was dropped during RTF export",
                        ),
                    }
                }
                if drawing.anchor.is_some() {
                    self.diagnose(
                        &content_location,
                        "anchored drawing was dropped during RTF export",
                    );
                }
            }
        }
    }

    fn scan_run_properties(&mut self, properties: &CT_RPr, location: &str) {
        if properties.style_id.is_some() {
            self.diagnose(
                &format!("{location}/rPr/rStyle"),
                "run style was dropped during RTF export",
            );
        }
        let selected_font = run_font(properties);
        for (field, font) in [
            ("hAnsi", properties.font_hansi.as_deref()),
            ("eastAsia", properties.font_east_asia.as_deref()),
            ("cs", properties.font_cs.as_deref()),
        ] {
            if let Some(font) = font
                && Some(font) != selected_font
            {
                self.diagnose(
                    &format!("{location}/rPr/{field}"),
                    "alternate run font was dropped during RTF export",
                );
            }
        }
        for (field, font) in [
            ("asciiTheme", properties.font_ascii_theme.as_deref()),
            ("hAnsiTheme", properties.font_hansi_theme.as_deref()),
        ] {
            if font.is_some() {
                self.diagnose(
                    &format!("{location}/rPr/{field}"),
                    "theme run font was dropped during RTF export",
                );
            }
        }
        if properties.bold_cs.is_some() && properties.bold_cs != properties.bold {
            self.diagnose(
                &format!("{location}/rPr/bCs"),
                "complex-script bold was dropped during RTF export",
            );
        }
        if properties.italic_cs.is_some() && properties.italic_cs != properties.italic {
            self.diagnose(
                &format!("{location}/rPr/iCs"),
                "complex-script italic was dropped during RTF export",
            );
        }
        if properties.underline.is_some_and(|underline| {
            !matches!(
                underline,
                rdocx_oxml::shared::ST_Underline::None | rdocx_oxml::shared::ST_Underline::Single
            )
        }) {
            self.diagnose(
                &format!("{location}/rPr/u"),
                "non-basic underline style was simplified during RTF export",
            );
        }
        if properties.dstrike.is_some() {
            self.diagnose(
                &format!("{location}/rPr/dstrike"),
                "double strikethrough was dropped during RTF export",
            );
        }
        if properties.sz_cs.is_some() && properties.sz_cs != properties.sz {
            self.diagnose(
                &format!("{location}/rPr/szCs"),
                "complex-script font size was dropped during RTF export",
            );
        }
        if properties.color_theme.is_some() {
            self.diagnose(
                &format!("{location}/rPr/themeColor"),
                "theme run colour was dropped during RTF export",
            );
        }
        if properties.highlight.is_some() {
            self.diagnose(
                &format!("{location}/rPr/highlight"),
                "keyword highlight colour was dropped during RTF export",
            );
        }
        if properties.spacing.is_some() {
            self.diagnose(
                &format!("{location}/rPr/spacing"),
                "run character spacing was dropped during RTF export",
            );
        }
        if properties.width_scale.is_some() {
            self.diagnose(
                &format!("{location}/rPr/w"),
                "run width scale was dropped during RTF export",
            );
        }
        if properties.position.is_some() {
            self.diagnose(
                &format!("{location}/rPr/position"),
                "run text position was dropped during RTF export",
            );
        }
        if !properties.revision_markers.is_empty() {
            self.diagnose(
                &format!("{location}/rPr/revisions"),
                "run revision markers were dropped during RTF export",
            );
        }
        if properties.change.is_some() {
            self.diagnose(
                &format!("{location}/rPr/rPrChange"),
                "run property revision was dropped during RTF export",
            );
        }
        if !properties.revision_xml.is_empty() {
            self.diagnose(
                &format!("{location}/rPr/revisionXml"),
                "unmodelled run property revision XML was dropped during RTF export",
            );
        }
    }

    fn write_font_table(&self, output: &mut BoundedOutput) -> Result<()> {
        write!(output, "{{\\fonttbl")?;
        let mut fonts = self.fonts.iter().collect::<Vec<_>>();
        fonts.sort_by_key(|(_, id)| **id);
        for (font, id) in fonts {
            write!(output, "{{\\f{id}\\fcharset0 ")?;
            write_ascii_text(output, font)?;
            write!(output, ";}}")?;
        }
        write!(output, "}}")?;
        Ok(())
    }

    fn write_color_table(&self, output: &mut BoundedOutput) -> Result<()> {
        if self.colors.is_empty() {
            return Ok(());
        }
        write!(output, "{{\\colortbl;")?;
        let mut colors = self.colors.iter().collect::<Vec<_>>();
        colors.sort_by_key(|(_, id)| **id);
        for (color, _) in colors {
            let (red, green, blue) = parse_hex_color(color).unwrap_or((0, 0, 0));
            write!(output, "\\red{red}\\green{green}\\blue{blue};")?;
        }
        write!(output, "}}")?;
        Ok(())
    }

    fn write_list_tables(&self, output: &mut BoundedOutput) -> Result<()> {
        if self.lists.is_empty() {
            return Ok(());
        }
        write!(output, "{{\\*\\listtable")?;
        for list in &self.lists {
            write!(output, "{{\\list")?;
            for level in &list.levels {
                let format = level.format.unwrap_or(ListNumberFormat::Decimal);
                write!(
                    output,
                    "{{\\listlevel\\levelnfc{}\\levelnfcn{}\\levelstartat{}}}",
                    list_format_value(format),
                    list_format_value(format),
                    level.start
                )?;
            }
            write!(output, "\\listid{}}}", list.list_id)?;
        }
        write!(output, "}}{{\\*\\listoverridetable")?;
        for list in &self.lists {
            write!(
                output,
                "{{\\listoverride\\listid{}\\ls{}}}",
                list.list_id, list.override_id
            )?;
        }
        write!(output, "}}")?;
        Ok(())
    }

    fn write_body_content(
        &mut self,
        output: &mut BoundedOutput,
        content: &BodyContent,
        location: &str,
    ) -> Result<()> {
        match content {
            BodyContent::Paragraph(paragraph) => self.write_paragraph(output, paragraph, location),
            BodyContent::Table(table) => self.write_table(output, table, location),
            BodyContent::ContentControl(_) | BodyContent::RawXml(_) => Ok(()),
        }
    }

    fn write_table(
        &mut self,
        output: &mut BoundedOutput,
        table: &CT_Tbl,
        location: &str,
    ) -> Result<()> {
        for (row_index, row) in table.rows.iter().enumerate() {
            write!(output, "\\trowd")?;
            let widths = table_row_cell_widths(table, row);
            let mut boundary = 0_i32;
            for width in &widths {
                boundary = boundary.checked_add(*width).ok_or_else(|| {
                    rtf_error(0, "RTF table cell boundaries exceed the supported range")
                })?;
                write!(output, "\\cellx{boundary}")?;
            }
            output.push(b' ')?;
            for (cell_index, cell) in row.cells.iter().enumerate() {
                for (content_index, content) in cell.content.iter().enumerate() {
                    if let CellContent::Paragraph(paragraph) = content {
                        if content_index > 0 {
                            write!(output, "\\par ")?;
                        }
                        self.write_paragraph_contents(
                            output,
                            paragraph,
                            &format!("{location}/row[{row_index}]/cell[{cell_index}]/content[{content_index}]"),
                        )?;
                    }
                }
                write!(output, "\\cell ")?;
            }
            writeln!(output, "\\row")?;
        }
        Ok(())
    }

    fn write_paragraph(
        &mut self,
        output: &mut BoundedOutput,
        paragraph: &CT_P,
        location: &str,
    ) -> Result<()> {
        write!(output, "\\pard")?;
        self.write_paragraph_format(output, paragraph.properties.as_ref(), paragraph)?;
        self.write_runs(output, &paragraph.runs, location)?;
        writeln!(output, "\\par")?;
        Ok(())
    }

    fn write_paragraph_contents(
        &mut self,
        output: &mut BoundedOutput,
        paragraph: &CT_P,
        location: &str,
    ) -> Result<()> {
        write!(output, "\\pard\\intbl")?;
        self.write_paragraph_format(output, paragraph.properties.as_ref(), paragraph)?;
        self.write_runs(output, &paragraph.runs, location)
    }

    fn write_paragraph_format(
        &self,
        output: &mut BoundedOutput,
        properties: Option<&CT_PPr>,
        paragraph: &CT_P,
    ) -> Result<()> {
        let Some(properties) = properties else {
            return Ok(());
        };
        match properties.jc {
            Some(ST_Jc::Center) => write!(output, "\\qc")?,
            Some(ST_Jc::Right | ST_Jc::End) => write!(output, "\\qr")?,
            Some(ST_Jc::Both | ST_Jc::Distribute) => write!(output, "\\qj")?,
            Some(_) | None => {}
        }
        if let Some(value) = properties.ind_left {
            write!(output, "\\li{}", value.0)?;
        }
        if let Some(value) = properties.ind_right {
            write!(output, "\\ri{}", value.0)?;
        }
        if let Some(value) = properties.ind_first_line {
            write!(output, "\\fi{}", value.0)?;
        } else if let Some(value) = properties.ind_hanging {
            write!(output, "\\fi{}", value.0.saturating_neg())?;
        }
        if let Some(value) = properties.space_before {
            write!(output, "\\sb{}", value.0)?;
        }
        if let Some(value) = properties.space_after {
            write!(output, "\\sa{}", value.0)?;
        }
        if let Some(value) = properties.line_spacing {
            match properties.line_rule.as_deref() {
                Some("exact") => write!(output, "\\sl{}\\slmult0", value.0.saturating_neg())?,
                Some("atLeast") => write!(output, "\\sl{}\\slmult0", value.0)?,
                _ => write!(output, "\\sl{}\\slmult1", value.0)?,
            }
        }
        if let Some((num_id, level)) = paragraph_numbering(paragraph)
            && level <= 8
            && let Some(list_index) = self.list_indexes.get(&num_id)
        {
            let list = &self.lists[*list_index];
            let Some(writer_level) = list.levels.get(level as usize) else {
                return Ok(());
            };
            if writer_level.format.is_some() {
                write!(output, "\\ls{}\\ilvl{}", list.override_id, level)?;
            }
        }
        Ok(())
    }

    fn write_runs(
        &mut self,
        output: &mut BoundedOutput,
        runs: &[CT_R],
        location: &str,
    ) -> Result<()> {
        for (run_index, run) in runs.iter().enumerate() {
            write!(output, "{{\\plain")?;
            self.write_run_format(output, run.properties.as_ref())?;
            output.push(b' ')?;
            for (content_index, content) in run.content.iter().enumerate() {
                self.write_run_content(
                    output,
                    content,
                    &format!("{location}/run[{run_index}]/content[{content_index}]"),
                )?;
            }
            write!(output, "}}")?;
        }
        Ok(())
    }

    fn write_run_format(
        &self,
        output: &mut BoundedOutput,
        properties: Option<&CT_RPr>,
    ) -> Result<()> {
        let Some(properties) = properties else {
            write!(output, "\\f0")?;
            return Ok(());
        };
        let font_id = run_font(properties)
            .and_then(|font| self.fonts.get(font).copied())
            .unwrap_or(0);
        write!(output, "\\f{font_id}")?;
        if let Some(size) = properties.sz {
            write!(output, "\\fs{}", size.0)?;
        }
        write_toggle(output, "b", properties.bold)?;
        write_toggle(output, "i", properties.italic)?;
        match properties.underline {
            Some(rdocx_oxml::shared::ST_Underline::None) => write!(output, "\\ulnone")?,
            Some(_) => write!(output, "\\ul")?,
            None => {}
        }
        write_toggle(output, "strike", properties.strike)?;
        write_toggle(output, "caps", properties.caps)?;
        write_toggle(output, "scaps", properties.small_caps)?;
        write_toggle(output, "v", properties.vanish)?;
        match properties.vert_align.as_deref() {
            Some("superscript") => write!(output, "\\super")?,
            Some("subscript") => write!(output, "\\sub")?,
            _ => {}
        }
        if let Some(color) = properties.color.as_deref()
            && let Some(index) = self.colors.get(&normalize_hex_color(color))
        {
            write!(output, "\\cf{index}")?;
        }
        if let Some(fill) = properties
            .shading
            .as_ref()
            .and_then(|shading| shading.fill.as_deref())
            && let Some(index) = self.colors.get(&normalize_hex_color(fill))
        {
            write!(output, "\\highlight{index}")?;
        }
        Ok(())
    }

    fn write_run_content(
        &mut self,
        output: &mut BoundedOutput,
        content: &RunContent,
        location: &str,
    ) -> Result<()> {
        match content {
            RunContent::Text(text) => write_rtf_text(output, &text.text),
            RunContent::DeletedText(text) => {
                self.diagnose(location, "deleted text was flattened during RTF export");
                write_rtf_text(output, &text.text)
            }
            RunContent::Tab => {
                write!(output, "\\tab ")?;
                Ok(())
            }
            RunContent::Break(BreakType::Line) => {
                write!(output, "\\line ")?;
                Ok(())
            }
            RunContent::Break(_) => {
                self.diagnose(
                    location,
                    "unsupported break type was dropped during RTF export",
                );
                Ok(())
            }
            RunContent::Drawing(drawing) => {
                if let Some(inline) = &drawing.inline
                    && let Some(bytes) = self.document.image_data(&inline.embed_id)
                    && let Some(kind) = picture_kind(&bytes)
                {
                    write!(
                        output,
                        "{{\\pict\\{}\\picwgoal{}\\pichgoal{} ",
                        kind,
                        Length::emu(inline.extent_cx.0).to_twips(),
                        Length::emu(inline.extent_cy.0).to_twips()
                    )?;
                    write_hex_bytes(output, &bytes)?;
                    write!(output, "}}")?;
                }
                Ok(())
            }
            RunContent::Field(field) => {
                self.diagnose(location, "field was flattened during RTF export");
                if let Some(text) = field.projected_text() {
                    write_rtf_text(output, text)?;
                }
                Ok(())
            }
            RunContent::FootnoteRef { .. } => {
                self.diagnose(location, "footnote reference was dropped during RTF export");
                Ok(())
            }
            RunContent::EndnoteRef { .. } => {
                self.diagnose(location, "endnote reference was dropped during RTF export");
                Ok(())
            }
            RunContent::CommentReference { .. } => {
                self.diagnose(location, "comment reference was dropped during RTF export");
                Ok(())
            }
        }
    }

    fn ensure_font(&mut self, font: &str) {
        if !self.fonts.contains_key(font) {
            let next = self.fonts.len();
            self.fonts.insert(font.to_owned(), next);
        }
    }

    fn ensure_hex_color(&mut self, color: &str, location: &str) {
        let color = normalize_hex_color(color);
        if parse_hex_color(&color).is_none() {
            self.diagnose(
                location,
                "unsupported colour value was dropped during RTF export",
            );
            return;
        }
        if !self.colors.contains_key(&color) {
            let next = self.colors.len() + 1;
            self.colors.insert(color, next);
        }
    }

    fn ensure_writer_list(&mut self, num_id: u32, location: &str) {
        if self.list_indexes.contains_key(&num_id) {
            return;
        }
        let mut levels = Vec::new();
        if let Some(abstract_num) = self
            .document
            .numbering
            .as_ref()
            .and_then(|numbering| numbering.get_abstract_num_for(num_id))
        {
            for level_index in 0..9_u32 {
                let level = abstract_num
                    .levels
                    .iter()
                    .find(|level| level.ilvl == level_index);
                let format = level
                    .and_then(|level| level.num_fmt.clone())
                    .and_then(public_number_format);
                if level.and_then(|level| level.num_fmt.as_ref()).is_some() && format.is_none() {
                    self.diagnose(
                        &format!("numbering[numId={num_id}]/level[{level_index}]/numFmt"),
                        "unsupported numbering format was dropped during RTF export",
                    );
                }
                levels.push(WriterListLevel {
                    format,
                    start: level.and_then(|level| level.start).unwrap_or(1),
                });
            }
        } else {
            self.diagnose(
                location,
                "unknown list definition was exported as decimal RTF list",
            );
            levels.push(WriterListLevel {
                format: Some(ListNumberFormat::Decimal),
                start: 1,
            });
        }
        let index = self.lists.len();
        self.lists.push(WriterList {
            list_id: 10 + index as i32,
            override_id: 1 + index as i32,
            levels,
        });
        self.list_indexes.insert(num_id, index);
    }

    fn diagnose(&mut self, location: &str, message: &str) {
        if self.diagnostics.len() >= MAX_DIAGNOSTICS {
            return;
        }
        self.diagnostics.push(RtfDiagnostic {
            offset: stable_location_offset(location),
            destination: Some(location.to_owned()),
            message: message.to_owned(),
        });
    }
}

fn paragraph_numbering(paragraph: &CT_P) -> Option<(u32, u32)> {
    let properties = paragraph.properties.as_ref()?;
    Some((properties.num_id?, properties.num_ilvl.unwrap_or(0)))
}

fn paragraph_marker_raw_positions(paragraph: &CT_P) -> BTreeSet<usize> {
    let mut positions = BTreeSet::new();
    for marker in &paragraph.comment_ranges {
        match marker {
            rdocx_oxml::text::CommentRangeMarker::Start { run_index, .. }
            | rdocx_oxml::text::CommentRangeMarker::End { run_index, .. } => {
                positions.insert(*run_index);
            }
        }
    }
    for marker in &paragraph.bookmark_markers {
        positions.insert(marker.run_index());
    }
    positions
}

fn marker_location(marker: &rdocx_oxml::text::CommentRangeMarker) -> usize {
    match marker {
        rdocx_oxml::text::CommentRangeMarker::Start { run_index, .. }
        | rdocx_oxml::text::CommentRangeMarker::End { run_index, .. } => *run_index,
    }
}

fn run_font(properties: &CT_RPr) -> Option<&str> {
    properties
        .font_ascii
        .as_deref()
        .or(properties.font_hansi.as_deref())
        .or(properties.font_east_asia.as_deref())
        .or(properties.font_cs.as_deref())
}

fn write_toggle(output: &mut BoundedOutput, name: &str, value: Option<bool>) -> Result<()> {
    match value {
        Some(true) => write!(output, "\\{name}")?,
        Some(false) => write!(output, "\\{name}0")?,
        None => {}
    }
    Ok(())
}

fn write_ascii_text(output: &mut BoundedOutput, text: &str) -> Result<()> {
    for scalar in text.chars() {
        if scalar.is_ascii() && scalar != '\\' && scalar != '{' && scalar != '}' {
            write!(output, "{scalar}")?;
        } else {
            write_rtf_char(output, scalar)?;
        }
    }
    Ok(())
}

fn write_rtf_text(output: &mut BoundedOutput, text: &str) -> Result<()> {
    for scalar in text.chars() {
        write_rtf_char(output, scalar)?;
    }
    Ok(())
}

fn write_rtf_char(output: &mut BoundedOutput, scalar: char) -> Result<()> {
    match scalar {
        '\\' => write!(output, "\\\\")?,
        '{' => write!(output, "\\{{")?,
        '}' => write!(output, "\\}}")?,
        '\t' => write!(output, "\\tab ")?,
        '\n' => write!(output, "\\line ")?,
        '\r' => {}
        scalar if scalar.is_ascii() => write!(output, "{scalar}")?,
        scalar => {
            let mut units = [0_u16; 2];
            for unit in scalar.encode_utf16(&mut units) {
                let signed = i16::from_ne_bytes(unit.to_ne_bytes()) as i32;
                write!(output, "\\u{signed}?")?;
            }
        }
    }
    Ok(())
}

fn write_hex_bytes(output: &mut BoundedOutput, bytes: &[u8]) -> Result<()> {
    let expanded = bytes.len().checked_mul(2).ok_or_else(output_limit_error)?;
    output.check_available(expanded)?;
    for byte in bytes {
        write!(output, "{byte:02x}")?;
    }
    Ok(())
}

fn picture_kind(bytes: &[u8]) -> Option<&'static str> {
    match oxml_media::probe(bytes)?.format {
        oxml_media::ImageFormat::Png => Some("pngblip"),
        oxml_media::ImageFormat::Jpeg => Some("jpegblip"),
        _ => None,
    }
}

fn table_row_cell_widths(table: &CT_Tbl, row: &CT_Row) -> Vec<i32> {
    let mut grid_index = 0_usize;
    row.cells
        .iter()
        .map(|cell| {
            let span = cell
                .properties
                .as_ref()
                .and_then(|properties| properties.grid_span)
                .unwrap_or(1)
                .max(1) as usize;
            let width = cell
                .properties
                .as_ref()
                .and_then(|properties| properties.width.as_ref())
                .and_then(cell_width_twips)
                .or_else(|| {
                    let grid = table.grid.as_ref()?;
                    let end = grid_index.checked_add(span)?;
                    if end > grid.columns.len() {
                        return None;
                    }
                    let width = grid.columns[grid_index..end]
                        .iter()
                        .try_fold(0_i32, |total, column| total.checked_add(column.width.0))?;
                    (width > 0).then_some(width)
                })
                .unwrap_or(1440);
            grid_index = grid_index.saturating_add(span);
            if width == 0 { 1440 } else { width }
        })
        .collect()
}

fn table_width_is_preserved_by_grid(table: &CT_Tbl, width: &CT_TblWidth) -> bool {
    let Some(width) = cell_width_twips(width) else {
        return false;
    };
    let Some(grid) = &table.grid else {
        return false;
    };
    grid.columns
        .iter()
        .try_fold(0_i32, |total, column| total.checked_add(column.width.0))
        == Some(width)
}

fn cell_width_twips(width: &CT_TblWidth) -> Option<i32> {
    (width.width_type == "dxa" && width.w > 0).then_some(width.w)
}

fn row_grid_span_end(row: &CT_Row, cell_index: usize) -> Option<usize> {
    let mut grid_index = 0_usize;
    for cell in row.cells.iter().take(cell_index + 1) {
        let span = cell
            .properties
            .as_ref()
            .and_then(|properties| properties.grid_span)
            .unwrap_or(1)
            .max(1) as usize;
        grid_index = grid_index.checked_add(span)?;
    }
    Some(grid_index)
}

fn raw_xml_item_name(raw: &[u8]) -> Option<&str> {
    let start = raw.iter().position(|byte| *byte == b'<')? + 1;
    let name_start = start + usize::from(raw.get(start) == Some(&b'/'));
    let name_end = raw[name_start..]
        .iter()
        .position(|byte| matches!(*byte, b' ' | b'\t' | b'\r' | b'\n' | b'/' | b'>'))?
        + name_start;
    let qualified = std::str::from_utf8(&raw[name_start..name_end]).ok()?;
    Some(
        qualified
            .rsplit_once(':')
            .map_or(qualified, |(_, local)| local),
    )
}

fn normalize_hex_color(color: &str) -> String {
    color.trim_start_matches('#').to_ascii_uppercase()
}

fn parse_hex_color(color: &str) -> Option<(u8, u8, u8)> {
    let color = normalize_hex_color(color);
    if color.len() != 6 || !color.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return None;
    }
    let red = u8::from_str_radix(&color[0..2], 16).ok()?;
    let green = u8::from_str_radix(&color[2..4], 16).ok()?;
    let blue = u8::from_str_radix(&color[4..6], 16).ok()?;
    Some((red, green, blue))
}

fn public_number_format(format: ST_NumberFormat) -> Option<ListNumberFormat> {
    match format {
        ST_NumberFormat::Bullet => Some(ListNumberFormat::Bullet),
        ST_NumberFormat::Decimal => Some(ListNumberFormat::Decimal),
        ST_NumberFormat::LowerLetter => Some(ListNumberFormat::LowerLetter),
        ST_NumberFormat::UpperLetter => Some(ListNumberFormat::UpperLetter),
        ST_NumberFormat::LowerRoman => Some(ListNumberFormat::LowerRoman),
        ST_NumberFormat::UpperRoman => Some(ListNumberFormat::UpperRoman),
        ST_NumberFormat::Ordinal => Some(ListNumberFormat::Ordinal),
        _ => None,
    }
}

fn list_format_value(format: ListNumberFormat) -> i32 {
    match format {
        ListNumberFormat::Decimal => 0,
        ListNumberFormat::UpperRoman => 1,
        ListNumberFormat::LowerRoman => 2,
        ListNumberFormat::UpperLetter => 3,
        ListNumberFormat::LowerLetter => 4,
        ListNumberFormat::Ordinal => 5,
        ListNumberFormat::Bullet => 23,
    }
}

fn stable_location_offset(location: &str) -> usize {
    let mut hash = 0xcbf29ce484222325_u64;
    for byte in location.bytes() {
        hash ^= u64::from(byte);
        hash = hash.wrapping_mul(0x100000001b3);
    }
    hash as usize
}

#[cfg(test)]
mod writer_tests {
    use super::*;

    const TINY_PNG: &[u8] = &[
        0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44,
        0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f,
        0x15, 0xc4, 0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x60,
        0xf8, 0xcf, 0xf0, 0x00, 0x00, 0x04, 0x01, 0x01, 0x08, 0x9d, 0x1d, 0xe1, 0x00, 0x00, 0x00,
        0x00, 0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
    ];

    #[test]
    fn rtf_writer_checks_picture_hex_budget_before_expansion() {
        let mut document = Document::new();
        document.add_picture(
            TINY_PNG,
            "tiny.png",
            Length::emu(12_700),
            Length::emu(12_700),
        );

        let error = match RtfWriter::new_with_output_limit(&document, 180).write() {
            Ok(_) => panic!("picture hex expansion should be rejected before writing"),
            Err(error) => error,
        };

        assert!(
            error
                .to_string()
                .contains("RTF output exceeds the size limit")
        );
    }

    #[test]
    fn rtf_writer_caps_diagnostics_at_reader_limit() {
        let mut document = Document::new();
        document.document.body.content.clear();
        for index in 0..(MAX_DIAGNOSTICS + 1) {
            document.document.body.content.push(BodyContent::RawXml(
                format!("<p:item id=\"{index}\"/>").into_bytes(),
            ));
        }

        let written = document.to_rtf_bytes().unwrap();

        assert_eq!(written.diagnostics.len(), MAX_DIAGNOSTICS);
        assert_eq!(
            written.diagnostics.last().unwrap().destination.as_deref(),
            Some("body[9999]")
        );
    }
}

#[derive(Debug)]
enum Token<'a> {
    Open {
        offset: usize,
    },
    Close {
        offset: usize,
    },
    Word {
        offset: usize,
        name: String,
        parameter: Option<i32>,
    },
    Symbol {
        offset: usize,
        symbol: u8,
    },
    Hex {
        offset: usize,
        byte: u8,
    },
    Binary {
        offset: usize,
        bytes: &'a [u8],
    },
    Text {
        offset: usize,
        bytes: &'a [u8],
    },
}

struct Scanner<'a> {
    input: &'a [u8],
    position: usize,
}

impl<'a> Scanner<'a> {
    fn new(input: &'a [u8]) -> Self {
        Self { input, position: 0 }
    }

    fn next(&mut self) -> Result<Option<Token<'a>>> {
        while self
            .input
            .get(self.position)
            .is_some_and(|byte| matches!(byte, b'\r' | b'\n'))
        {
            self.position += 1;
        }
        let Some(&byte) = self.input.get(self.position) else {
            return Ok(None);
        };
        let offset = self.position;
        match byte {
            b'{' => {
                self.position += 1;
                Ok(Some(Token::Open { offset }))
            }
            b'}' => {
                self.position += 1;
                Ok(Some(Token::Close { offset }))
            }
            b'\\' => self.scan_control(offset).map(Some),
            _ => {
                let start = self.position;
                while self
                    .input
                    .get(self.position)
                    .is_some_and(|byte| !matches!(byte, b'{' | b'}' | b'\\' | b'\r' | b'\n'))
                {
                    self.position += 1;
                }
                Ok(Some(Token::Text {
                    offset: start,
                    bytes: &self.input[start..self.position],
                }))
            }
        }
    }

    fn scan_control(&mut self, offset: usize) -> Result<Token<'a>> {
        self.position += 1;
        let Some(&next) = self.input.get(self.position) else {
            return Err(rtf_error(offset, "trailing backslash"));
        };
        if next == b'\'' {
            if self.position + 2 >= self.input.len() {
                return Err(rtf_error(offset, "truncated hexadecimal escape"));
            }
            let high = hex_value(self.input[self.position + 1])
                .ok_or_else(|| rtf_error(offset, "invalid hexadecimal escape"))?;
            let low = hex_value(self.input[self.position + 2])
                .ok_or_else(|| rtf_error(offset, "invalid hexadecimal escape"))?;
            self.position += 3;
            return Ok(Token::Hex {
                offset,
                byte: high << 4 | low,
            });
        }
        if !next.is_ascii_alphabetic() {
            self.position += 1;
            return Ok(Token::Symbol {
                offset,
                symbol: next,
            });
        }

        let word_start = self.position;
        while self
            .input
            .get(self.position)
            .is_some_and(u8::is_ascii_alphabetic)
        {
            self.position += 1;
        }
        if self.position - word_start > 32 {
            return Err(rtf_error(offset, "control word exceeds 32 letters"));
        }
        let name = String::from_utf8(self.input[word_start..self.position].to_vec())
            .map_err(|_| rtf_error(offset, "control word is not ASCII"))?;
        let number_start = self.position;
        if self.input.get(self.position) == Some(&b'-') {
            if !self
                .input
                .get(self.position + 1)
                .is_some_and(u8::is_ascii_digit)
            {
                return Err(rtf_error(
                    offset,
                    "minus sign is not followed by a numeric parameter",
                ));
            }
            self.position += 1;
        }
        let digit_start = self.position;
        while self
            .input
            .get(self.position)
            .is_some_and(u8::is_ascii_digit)
        {
            self.position += 1;
        }
        let parameter = if self.position > digit_start {
            if self.position - digit_start > 10 {
                return Err(rtf_error(offset, "control parameter exceeds 10 digits"));
            }
            let value = std::str::from_utf8(&self.input[number_start..self.position])
                .ok()
                .and_then(|value| value.parse::<i64>().ok())
                .filter(|value| i32::try_from(*value).is_ok())
                .ok_or_else(|| rtf_error(offset, "control parameter is out of range"))?;
            Some(value as i32)
        } else {
            None
        };
        if self.input.get(self.position) == Some(&b' ') {
            self.position += 1;
        }
        if name == "bin" {
            let count = parameter.ok_or_else(|| rtf_error(offset, "bin requires a byte count"))?;
            let count = usize::try_from(count)
                .map_err(|_| rtf_error(offset, "bin byte count is negative"))?;
            if count > MAX_PICTURE_BYTES {
                return Err(rtf_error(offset, "binary payload exceeds the RTF limit"));
            }
            let end = self
                .position
                .checked_add(count)
                .filter(|end| *end <= self.input.len())
                .ok_or_else(|| rtf_error(offset, "truncated binary payload"))?;
            let bytes = &self.input[self.position..end];
            self.position = end;
            return Ok(Token::Binary { offset, bytes });
        }
        Ok(Token::Word {
            offset,
            name,
            parameter,
        })
    }
}

fn hex_value(byte: u8) -> Option<u8> {
    match byte {
        b'0'..=b'9' => Some(byte - b'0'),
        b'a'..=b'f' => Some(byte - b'a' + 10),
        b'A'..=b'F' => Some(byte - b'A' + 10),
        _ => None,
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
enum Destination {
    Body,
    FontTable,
    ColorTable,
    ListTable,
    ListOverrideTable,
    ListText,
    Picture,
    Container,
    UnicodeAlternatives,
    Skip(String),
}

#[derive(Clone, Debug, Default, PartialEq)]
struct RunFormat {
    bold: Option<bool>,
    italic: Option<bool>,
    underline: Option<bool>,
    strike: Option<bool>,
    font: Option<i32>,
    size_points: Option<f64>,
    color: Option<usize>,
    highlight: Option<usize>,
    caps: Option<bool>,
    small_caps: Option<bool>,
    hidden: Option<bool>,
    vertical: Option<VerticalPosition>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum VerticalPosition {
    Super,
    Sub,
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
struct ParagraphFormat {
    alignment: Option<Alignment>,
    list_override: Option<i32>,
    list_level: u32,
    indent_left: Option<i32>,
    indent_right: Option<i32>,
    first_line_indent: Option<i32>,
    space_before: Option<i32>,
    space_after: Option<i32>,
    line_spacing: Option<i32>,
    line_spacing_multiple: bool,
}

#[derive(Clone, Debug)]
struct FontEntry {
    name: String,
    code_page: Option<u16>,
}

#[derive(Debug)]
struct State {
    destination: Destination,
    starred: bool,
    at_group_start: bool,
    code_page: u16,
    uc_skip: usize,
    skip_remaining: usize,
    pending_high_surrogate: Option<u16>,
    format: RunFormat,
    paragraph: ParagraphFormat,
    text_bytes: Vec<u8>,
    text_offset: usize,
    font_id: Option<i32>,
    font_charset: Option<i32>,
    font_code_page: Option<u16>,
    font_name: Vec<u8>,
    color_red: Option<u8>,
    color_green: Option<u8>,
    color_blue: Option<u8>,
    unicode_alternative_children: usize,
    list_override_entry: bool,
}

impl Default for State {
    fn default() -> Self {
        Self {
            destination: Destination::Body,
            starred: false,
            at_group_start: true,
            code_page: 1252,
            uc_skip: 1,
            skip_remaining: 0,
            pending_high_surrogate: None,
            format: RunFormat::default(),
            paragraph: ParagraphFormat::default(),
            text_bytes: Vec::new(),
            text_offset: 0,
            font_id: None,
            font_charset: None,
            font_code_page: None,
            font_name: Vec::new(),
            color_red: None,
            color_green: None,
            color_blue: None,
            unicode_alternative_children: 0,
            list_override_entry: false,
        }
    }
}

impl State {
    fn child(&self) -> Self {
        let in_font_table = self.destination == Destination::FontTable;
        Self {
            destination: self.destination.clone(),
            starred: false,
            at_group_start: true,
            code_page: self.code_page,
            uc_skip: self.uc_skip,
            skip_remaining: self.skip_remaining,
            pending_high_surrogate: self.pending_high_surrogate,
            format: self.format.clone(),
            paragraph: self.paragraph.clone(),
            text_bytes: Vec::new(),
            text_offset: 0,
            font_id: (!in_font_table).then_some(self.font_id).flatten(),
            font_charset: (!in_font_table).then_some(self.font_charset).flatten(),
            font_code_page: (!in_font_table).then_some(self.font_code_page).flatten(),
            font_name: Vec::new(),
            color_red: self.color_red,
            color_green: self.color_green,
            color_blue: self.color_blue,
            unicode_alternative_children: self.unicode_alternative_children,
            list_override_entry: false,
        }
    }
}

#[derive(Clone, Debug, Default)]
struct ParagraphData {
    items: Vec<ParagraphItem>,
    format: ParagraphFormat,
    inferred_list: Option<ListNumberFormat>,
}

impl ParagraphData {
    fn has_content(&self) -> bool {
        self.items.iter().any(|item| match item {
            ParagraphItem::Text(run) => !run.text.is_empty(),
            ParagraphItem::Picture(_) | ParagraphItem::Break | ParagraphItem::Tab => true,
        })
    }
}

#[derive(Clone, Debug)]
enum ParagraphItem {
    Text(RunData),
    Picture(PictureData),
    Break,
    Tab,
}

#[derive(Clone, Debug)]
struct RunData {
    text: String,
    format: RunFormat,
}

#[derive(Debug)]
enum Block {
    Paragraph(ParagraphData),
    Table(Vec<TableRowData>),
}

#[derive(Debug)]
struct TableRowData {
    cells: Vec<Vec<ParagraphData>>,
    boundaries: Vec<i32>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum PictureKind {
    Png,
    Jpeg,
    Unsupported,
}

#[derive(Clone, Debug)]
struct PictureBuilder {
    offset: usize,
    kind: Option<PictureKind>,
    data: Vec<u8>,
    high_nibble: Option<u8>,
    width_goal: Option<i32>,
    height_goal: Option<i32>,
    scale_x: usize,
    scale_y: usize,
    crop_top: i32,
    crop_bottom: i32,
    crop_left: i32,
    crop_right: i32,
}

#[derive(Clone, Debug)]
struct PictureData {
    kind: PictureKind,
    data: Vec<u8>,
    width_goal: Option<i32>,
    height_goal: Option<i32>,
    scale_x: usize,
    scale_y: usize,
}

#[derive(Clone, Copy, Debug)]
struct ListLevelData {
    format: ListNumberFormat,
    has_new_format: bool,
    start: Option<u32>,
}

#[derive(Clone, Copy, Debug, Default)]
struct ListLevelOverrideData {
    format: Option<ListNumberFormat>,
    has_new_format: bool,
    start: Option<u32>,
}

#[derive(Clone, Debug)]
struct ListOverrideData {
    list_id: i32,
    levels: Vec<ListLevelOverrideData>,
}

impl Default for ListLevelData {
    fn default() -> Self {
        Self {
            format: ListNumberFormat::Decimal,
            has_new_format: false,
            start: None,
        }
    }
}

struct ParsedRtf {
    blocks: Vec<Block>,
    diagnostics: Vec<RtfDiagnostic>,
    fonts: HashMap<i32, FontEntry>,
    colors: Vec<Option<String>>,
    lists: HashMap<i32, Vec<ListLevelData>>,
    overrides: HashMap<i32, ListOverrideData>,
}

impl ParsedRtf {
    fn project(self) -> Result<RtfReadResult> {
        let ParsedRtf {
            blocks,
            diagnostics,
            fonts,
            colors,
            lists,
            overrides,
        } = self;
        let mut document = Document::new();
        let mut projected_lists = HashMap::new();
        for block in blocks {
            match block {
                Block::Paragraph(paragraph) => project_paragraph(
                    &mut document,
                    paragraph,
                    &fonts,
                    &colors,
                    &lists,
                    &overrides,
                    &mut projected_lists,
                )?,
                Block::Table(rows) => {
                    let columns = rows.iter().map(|row| row.cells.len()).max().unwrap_or(0);
                    let mut prepared = Vec::with_capacity(rows.len());
                    for row in &rows {
                        let mut prepared_row = Vec::with_capacity(row.cells.len());
                        for cell in &row.cells {
                            let mut prepared_cell = Vec::with_capacity(cell.len());
                            for paragraph in cell {
                                prepared_cell.push((
                                    ensure_list(
                                        &mut document,
                                        paragraph,
                                        &lists,
                                        &overrides,
                                        &mut projected_lists,
                                    )?,
                                    prepare_pictures(&mut document, paragraph)?,
                                ));
                            }
                            prepared_row.push(prepared_cell);
                        }
                        prepared.push(prepared_row);
                    }
                    let mut table = document.add_table(rows.len(), columns);
                    if let Some(boundaries) = rows.first().map(|row| &row.boundaries) {
                        let mut previous = 0;
                        for (column, boundary) in boundaries.iter().copied().enumerate() {
                            let width = boundary - previous;
                            if width <= 0 || !table.set_column_width(column, Length::twips(width)) {
                                return Err(rtf_error(0, "RTF table cell boundaries are invalid"));
                            }
                            previous = boundary;
                        }
                    }
                    for (row_index, (row, prepared_row)) in
                        rows.into_iter().zip(prepared).enumerate()
                    {
                        for (column_index, (paragraphs, prepared_cell)) in
                            row.cells.into_iter().zip(prepared_row).enumerate()
                        {
                            let mut cell = table.cell(row_index, column_index).unwrap();
                            cell.remove_first_empty_paragraph();
                            if paragraphs.is_empty() {
                                cell.add_paragraph("");
                            }
                            for (paragraph, (list_id, pictures)) in
                                paragraphs.into_iter().zip(prepared_cell)
                            {
                                let mut target = cell.add_paragraph("");
                                apply_paragraph_format(&mut target, &paragraph.format, list_id);
                                project_items(
                                    &mut target,
                                    paragraph.items,
                                    pictures,
                                    &fonts,
                                    &colors,
                                )?;
                            }
                        }
                    }
                }
            }
        }
        Ok(RtfReadResult {
            document,
            diagnostics,
        })
    }
}

fn project_paragraph(
    document: &mut Document,
    paragraph: ParagraphData,
    fonts: &HashMap<i32, FontEntry>,
    colors: &[Option<String>],
    lists: &HashMap<i32, Vec<ListLevelData>>,
    overrides: &HashMap<i32, ListOverrideData>,
    projected_lists: &mut HashMap<String, u32>,
) -> Result<()> {
    let list_id = ensure_list(document, &paragraph, lists, overrides, projected_lists)?;
    let pictures = prepare_pictures(document, &paragraph)?;
    let mut target = document.add_paragraph("");
    apply_paragraph_format(&mut target, &paragraph.format, list_id);
    project_items(&mut target, paragraph.items, pictures, fonts, colors)?;
    Ok(())
}

fn ensure_list(
    document: &mut Document,
    paragraph: &ParagraphData,
    lists: &HashMap<i32, Vec<ListLevelData>>,
    overrides: &HashMap<i32, ListOverrideData>,
    projected_lists: &mut HashMap<String, u32>,
) -> Result<Option<u32>> {
    if let Some(override_id) = paragraph.format.list_override {
        let key = format!("override:{override_id}");
        if let Some(num_id) = projected_lists.get(&key) {
            return Ok(Some(*num_id));
        }
        let override_data = overrides
            .get(&override_id)
            .ok_or_else(|| rtf_error(0, format!("RTF list override {override_id} is undefined")))?;
        let mut levels = lists
            .get(&override_data.list_id)
            .ok_or_else(|| {
                rtf_error(
                    0,
                    format!("RTF list {} is undefined", override_data.list_id),
                )
            })?
            .clone();
        for (level, replacement) in levels.iter_mut().zip(&override_data.levels) {
            if let Some(format) = replacement.format {
                level.format = format;
            }
            if let Some(start) = replacement.start {
                level.start = Some(start);
            }
        }
        let levels = levels
            .into_iter()
            .map(public_list_level)
            .collect::<Vec<_>>();
        let num_id = document.add_list_definition(&levels);
        projected_lists.insert(key, num_id);
        Ok(Some(num_id))
    } else if let Some(format) = paragraph.inferred_list {
        let key = format!("inferred:{format:?}");
        let num_id = *projected_lists
            .entry(key)
            .or_insert_with(|| document.add_list_definition(&[ListLevel::new(format)]));
        Ok(Some(num_id))
    } else {
        Ok(None)
    }
}

#[derive(Debug)]
struct PreparedPicture {
    relationship_id: String,
    width: Length,
    height: Length,
}

fn prepare_pictures(
    document: &mut Document,
    paragraph: &ParagraphData,
) -> Result<Vec<PreparedPicture>> {
    paragraph
        .items
        .iter()
        .filter_map(|item| match item {
            ParagraphItem::Picture(picture) => Some(picture),
            ParagraphItem::Text(_) | ParagraphItem::Break | ParagraphItem::Tab => None,
        })
        .map(|picture| {
            let filename = match picture.kind {
                PictureKind::Png => "image.png",
                PictureKind::Jpeg => "image.jpg",
                PictureKind::Unsupported => unreachable!(),
            };
            let (width, height) = picture_dimensions(picture)?;
            Ok(PreparedPicture {
                relationship_id: document.embed_image(&picture.data, filename),
                width,
                height,
            })
        })
        .collect()
}

fn picture_dimensions(picture: &PictureData) -> Result<(Length, Length)> {
    let native = oxml_media::probe(&picture.data)
        .and_then(|info| info.native_size(72.0))
        .ok_or_else(|| rtf_error(0, "RTF picture dimensions are unavailable"))?;
    let base_width = picture
        .width_goal
        .filter(|width| *width > 0)
        .map_or(native.width_emu, |width| Length::twips(width).to_emu());
    let base_height = picture
        .height_goal
        .filter(|height| *height > 0)
        .map_or(native.height_emu, |height| Length::twips(height).to_emu());
    let width = base_width
        .checked_mul(picture.scale_x as i64)
        .and_then(|value| value.checked_div(100))
        .ok_or_else(|| rtf_error(0, "RTF picture width scaling overflows"))?;
    let height = base_height
        .checked_mul(picture.scale_y as i64)
        .and_then(|value| value.checked_div(100))
        .ok_or_else(|| rtf_error(0, "RTF picture height scaling overflows"))?;
    Ok((Length::emu(width), Length::emu(height)))
}

fn project_items(
    target: &mut crate::Paragraph<'_>,
    items: Vec<ParagraphItem>,
    pictures: Vec<PreparedPicture>,
    fonts: &HashMap<i32, FontEntry>,
    colors: &[Option<String>],
) -> Result<()> {
    let mut pictures = pictures.into_iter();
    for item in items {
        match item {
            ParagraphItem::Text(run) => {
                let mut target_run = target.add_run(&run.text);
                apply_run_format(&mut target_run, &run.format, fonts, colors)?;
            }
            ParagraphItem::Picture(_) => {
                let picture = pictures
                    .next()
                    .ok_or_else(|| rtf_error(0, "RTF picture projection state is incomplete"))?;
                target.add_picture(&picture.relationship_id, picture.width, picture.height);
            }
            ParagraphItem::Break => target.add_line_break(),
            ParagraphItem::Tab => target.add_tab(),
        }
    }
    Ok(())
}

fn public_list_level(level: ListLevelData) -> ListLevel {
    let public = ListLevel::new(level.format);
    match level.start {
        Some(start) => public.start(start),
        None => public,
    }
}

fn apply_paragraph_format(
    paragraph: &mut crate::Paragraph<'_>,
    format: &ParagraphFormat,
    list_id: Option<u32>,
) {
    if let Some(alignment) = format.alignment {
        paragraph.set_alignment(alignment);
    }
    paragraph.set_indent_left_value(format.indent_left.map(Length::twips));
    paragraph.set_indent_right_value(format.indent_right.map(Length::twips));
    paragraph.set_signed_first_line_indent_value(format.first_line_indent.map(Length::twips));
    paragraph.set_space_before_value(format.space_before.map(Length::twips));
    paragraph.set_space_after_value(format.space_after.map(Length::twips));
    if let Some(line_spacing) = format.line_spacing {
        if line_spacing == 0 {
            paragraph.clear_line_spacing();
        } else if format.line_spacing_multiple {
            paragraph.set_line_spacing_multiple(line_spacing as f64 / 240.0);
        } else if line_spacing > 0 {
            paragraph.set_line_spacing_at_least(line_spacing as f64 / 20.0);
        } else {
            paragraph.set_line_spacing(line_spacing.saturating_abs() as f64 / 20.0);
        }
    }
    if let Some(list_id) = list_id {
        let _ = paragraph.set_numbering(list_id, format.list_level.min(8));
    }
}

fn apply_run_format(
    run: &mut crate::Run<'_>,
    format: &RunFormat,
    fonts: &HashMap<i32, FontEntry>,
    colors: &[Option<String>],
) -> Result<()> {
    run.set_bold_value(format.bold);
    run.set_italic_value(format.italic);
    if let Some(underline) = format.underline {
        run.set_underline(underline);
    }
    run.set_strike_value(format.strike);
    if let Some(id) = format.font {
        let font = fonts
            .get(&id)
            .ok_or_else(|| rtf_error(0, format!("RTF font {id} is undefined")))?;
        run.set_font(&font.name);
    }
    if let Some(size) = format.size_points {
        run.set_size(size);
    }
    if let Some(index) = format.color {
        let color = colors
            .get(index)
            .ok_or_else(|| rtf_error(0, format!("RTF colour {index} is undefined")))?;
        if let Some(color) = color {
            run.set_color(color);
        }
    }
    if let Some(index) = format.highlight {
        let color = colors
            .get(index)
            .ok_or_else(|| rtf_error(0, format!("RTF colour {index} is undefined")))?;
        if let Some(color) = color {
            run.set_highlight(color);
        }
    }
    if let Some(value) = format.caps {
        run.set_all_caps(value);
    }
    if let Some(value) = format.small_caps {
        run.set_small_caps(value);
    }
    if let Some(value) = format.hidden {
        run.set_hidden(value);
    }
    match format.vertical {
        Some(VerticalPosition::Super) => run.set_superscript(),
        Some(VerticalPosition::Sub) => run.set_subscript(),
        None => {}
    }
    Ok(())
}

struct Parser {
    states: Vec<State>,
    blocks: Vec<Block>,
    current_paragraph: ParagraphData,
    table_rows: Vec<TableRowData>,
    current_row: Vec<Vec<ParagraphData>>,
    current_cell: Vec<ParagraphData>,
    current_cell_boundaries: Vec<i32>,
    in_table: bool,
    expected_cells: usize,
    diagnostics: Vec<RtfDiagnostic>,
    fonts: HashMap<i32, FontEntry>,
    colors: Vec<Option<String>>,
    lists: HashMap<i32, Vec<ListLevelData>>,
    overrides: HashMap<i32, ListOverrideData>,
    active_list_levels: Vec<ListLevelData>,
    active_override_list: Option<i32>,
    active_override_id: Option<i32>,
    active_override_levels: Vec<ListLevelOverrideData>,
    active_picture: Option<PictureBuilder>,
    active_list_marker: Vec<u8>,
    active_list_marker_unicode: String,
    default_font: Option<i32>,
    seen_root: bool,
    seen_rtf: bool,
    expect_root_marker: bool,
    total_blocks: usize,
    total_runs: usize,
    total_cells: usize,
    retained_output_bytes: usize,
    body_started: bool,
    header_table_started: bool,
}

impl Parser {
    fn parse(input: &[u8]) -> Result<ParsedRtf> {
        if input.len() > MAX_INPUT_BYTES {
            return Err(rtf_error(0, "RTF input exceeds the size limit"));
        }
        let mut parser = Self {
            states: Vec::new(),
            blocks: Vec::new(),
            current_paragraph: ParagraphData::default(),
            table_rows: Vec::new(),
            current_row: Vec::new(),
            current_cell: Vec::new(),
            current_cell_boundaries: Vec::new(),
            in_table: false,
            expected_cells: 0,
            diagnostics: Vec::new(),
            fonts: HashMap::new(),
            colors: Vec::new(),
            lists: HashMap::new(),
            overrides: HashMap::new(),
            active_list_levels: Vec::new(),
            active_override_list: None,
            active_override_id: None,
            active_override_levels: Vec::new(),
            active_picture: None,
            active_list_marker: Vec::new(),
            active_list_marker_unicode: String::new(),
            default_font: None,
            seen_root: false,
            seen_rtf: false,
            expect_root_marker: false,
            total_blocks: 0,
            total_runs: 0,
            total_cells: 0,
            retained_output_bytes: 0,
            body_started: false,
            header_table_started: false,
        };
        let mut scanner = Scanner::new(input);
        while let Some(token) = scanner.next()? {
            let offset = token_offset(&token);
            parser.consume(token)?;
            if parser.diagnostics.len() > MAX_DIAGNOSTICS {
                return Err(rtf_error(offset, "RTF diagnostic limit exceeded"));
            }
        }
        if !parser.states.is_empty() {
            return Err(rtf_error(input.len(), "unbalanced RTF groups"));
        }
        if !parser.seen_root || !parser.seen_rtf {
            return Err(rtf_error(0, "input is not an RTF 1 document"));
        }
        parser.finish_document()?;
        Ok(ParsedRtf {
            blocks: parser.blocks,
            diagnostics: parser.diagnostics,
            fonts: parser.fonts,
            colors: parser.colors,
            lists: parser.lists,
            overrides: parser.overrides,
        })
    }

    fn consume(&mut self, token: Token<'_>) -> Result<()> {
        if self.expect_root_marker {
            match &token {
                Token::Word {
                    name, parameter, ..
                } if name == "rtf" && *parameter == Some(1) => {}
                _ => {
                    return Err(rtf_error(
                        token_offset(&token),
                        "\\rtf1 must immediately follow the root opening brace",
                    ));
                }
            }
            self.expect_root_marker = false;
        }
        match token {
            Token::Open { offset } => self.open_group(offset),
            Token::Close { offset } => self.close_group(offset),
            token if self.states.is_empty() => Err(rtf_error(
                token_offset(&token),
                "content appears outside the root RTF group",
            )),
            token => self.consume_inside(token),
        }
    }

    fn open_group(&mut self, offset: usize) -> Result<()> {
        if self.states.len() >= MAX_GROUP_DEPTH {
            return Err(rtf_error(offset, "RTF group nesting exceeds the limit"));
        }
        if let Some(mut parent) = self.states.pop() {
            if parent.starred {
                return Err(rtf_error(
                    offset,
                    "RTF star marker is not followed by a destination",
                ));
            }
            if self.awaiting_unicode_destination_with(&parent) {
                return Err(rtf_error(
                    offset,
                    "RTF Unicode alternate child must begin with \\*\\ud",
                ));
            }
            if parent.skip_remaining > 0 {
                return Err(rtf_error(
                    offset,
                    "Unicode fallback crosses a group boundary",
                ));
            }
            if parent.pending_high_surrogate.is_some() {
                return Err(rtf_error(
                    offset,
                    "RTF group boundary interrupts a Unicode surrogate pair",
                ));
            }
            self.flush_state(&mut parent)?;
            let mut child = parent.child();
            if parent.destination == Destination::UnicodeAlternatives {
                parent.unicode_alternative_children += 1;
                child.destination = match parent.unicode_alternative_children {
                    1 => Destination::Skip("upr-ansi".to_owned()),
                    2 => Destination::Container,
                    _ => {
                        return Err(rtf_error(
                            offset,
                            "Unicode alternate destination has more than two branches",
                        ));
                    }
                };
            }
            parent.at_group_start = false;
            self.states.push(parent);
            self.states.push(child);
        } else {
            if self.seen_root {
                return Err(rtf_error(offset, "more than one root RTF group"));
            }
            self.seen_root = true;
            self.states.push(State::default());
            self.expect_root_marker = true;
        }
        Ok(())
    }

    fn close_group(&mut self, offset: usize) -> Result<()> {
        let Some(mut state) = self.states.pop() else {
            return Err(rtf_error(offset, "unexpected closing brace"));
        };
        if state.skip_remaining > 0 {
            return Err(rtf_error(
                offset,
                "Unicode fallback crosses a group boundary",
            ));
        }
        self.flush_state(&mut state)?;
        if state.pending_high_surrogate.is_some() {
            return Err(rtf_error(offset, "unpaired Unicode high surrogate"));
        }
        if state.starred {
            return Err(rtf_error(
                offset,
                "RTF star marker is not followed by a destination",
            ));
        }
        if state.destination == Destination::UnicodeAlternatives
            && state.unicode_alternative_children != 2
        {
            return Err(rtf_error(
                offset,
                "RTF Unicode alternate destination requires exactly two children",
            ));
        }
        let parent_destination = self.states.last().map(|state| state.destination.clone());
        if state.destination == Destination::Picture
            && parent_destination != Some(Destination::Picture)
        {
            let picture = self.active_picture.take();
            self.finish_picture(picture, offset)?;
        }
        if state.destination == Destination::ListText
            && parent_destination != Some(Destination::ListText)
        {
            self.finish_list_marker(state.code_page)?;
        }
        if state.list_override_entry {
            self.finish_list_override(offset)?;
        }
        Ok(())
    }

    fn consume_inside(&mut self, token: Token<'_>) -> Result<()> {
        if matches!(
            self.states.last().unwrap().destination,
            Destination::Skip(_)
        ) {
            return Ok(());
        }
        if self.states.last().unwrap().skip_remaining > 0 {
            return self.skip_fallback(token);
        }
        let valid_unicode_destination_start = match &token {
            Token::Symbol { symbol: b'*', .. } => true,
            Token::Word { name, .. } if name == "ud" => true,
            _ => false,
        };
        if self.awaiting_unicode_destination() && !valid_unicode_destination_start {
            return Err(rtf_error(
                token_offset(&token),
                "RTF Unicode alternate child must begin with \\*\\ud",
            ));
        }
        match token {
            Token::Word {
                offset,
                name,
                parameter,
            } => self.control_word(offset, &name, parameter),
            Token::Symbol { offset, symbol } => self.control_symbol(offset, symbol),
            Token::Hex { offset, byte } => self.raw_text_byte(offset, byte),
            Token::Binary { offset, bytes } => self.binary(offset, bytes),
            Token::Text { offset, bytes } => self.text(offset, bytes),
            Token::Open { .. } | Token::Close { .. } => unreachable!(),
        }
    }

    fn skip_fallback(&mut self, token: Token<'_>) -> Result<()> {
        let state = self.states.last_mut().unwrap();
        match token {
            Token::Text { offset, bytes } => {
                let skipped = state.skip_remaining.min(bytes.len());
                state.skip_remaining -= skipped;
                if skipped < bytes.len() {
                    self.text(offset + skipped, &bytes[skipped..])?;
                }
            }
            Token::Open { offset } | Token::Close { offset } => {
                return Err(rtf_error(
                    offset,
                    "Unicode fallback crosses a group boundary",
                ));
            }
            _ => state.skip_remaining -= 1,
        }
        Ok(())
    }

    fn control_symbol(&mut self, offset: usize, symbol: u8) -> Result<()> {
        if self
            .states
            .last()
            .is_some_and(|state| state.pending_high_surrogate.is_some())
        {
            return Err(rtf_error(
                offset,
                "RTF control symbol interrupts a Unicode surrogate pair",
            ));
        }
        if symbol == b'*' {
            let state = self.states.last_mut().unwrap();
            if !state.at_group_start || state.starred {
                return Err(rtf_error(
                    offset,
                    "RTF star marker must begin a destination",
                ));
            }
            state.starred = true;
            return Ok(());
        }
        self.states.last_mut().unwrap().at_group_start = false;
        match symbol {
            b'\\' | b'{' | b'}' => self.raw_text_byte(offset, symbol)?,
            b'~' => self.append_unicode(offset, '\u{00a0}')?,
            b'_' => self.append_unicode(offset, '\u{2011}')?,
            b'-' => {}
            _ => self.diagnostics.push(RtfDiagnostic {
                offset,
                destination: None,
                message: format!(
                    "unsupported RTF control symbol \\{} ignored",
                    symbol as char
                ),
            }),
        }
        Ok(())
    }

    fn control_word(&mut self, offset: usize, name: &str, parameter: Option<i32>) -> Result<()> {
        if name != "u"
            && self
                .states
                .last()
                .is_some_and(|state| state.pending_high_surrogate.is_some())
        {
            return Err(rtf_error(
                offset,
                "RTF control word interrupts a Unicode surrogate pair",
            ));
        }
        let awaiting_unicode_destination = self.awaiting_unicode_destination();
        let mut state = self.states.pop().unwrap();
        self.flush_state(&mut state)?;
        let was_starred = std::mem::take(&mut state.starred);
        let at_group_start = state.at_group_start;
        state.at_group_start = false;
        self.states.push(state);
        let is_destination =
            is_supported_destination(name) || is_unsupported_destination(name) || was_starred;
        if is_destination && !at_group_start {
            return Err(rtf_error(
                offset,
                format!("RTF destination \\{name} does not begin its group"),
            ));
        }
        if name == "ud" && (!awaiting_unicode_destination || !was_starred) {
            return Err(rtf_error(
                offset,
                "RTF Unicode alternate child must begin with \\*\\ud",
            ));
        }
        if awaiting_unicode_destination && name != "ud" {
            return Err(rtf_error(
                offset,
                "RTF Unicode alternate child must begin with \\*\\ud",
            ));
        }
        if was_starred && !is_supported_destination(name) {
            self.states.last_mut().unwrap().destination = Destination::Skip(name.to_owned());
            self.diagnostics.push(RtfDiagnostic {
                offset,
                destination: Some(name.to_owned()),
                message: "unsupported RTF destination skipped".to_owned(),
            });
            return Ok(());
        }
        match name {
            "rtf" => {
                if parameter != Some(1) || self.seen_rtf || self.states.len() != 1 {
                    return Err(rtf_error(offset, "only RTF version 1 is supported"));
                }
                self.seen_rtf = true;
            }
            "ansi" => self.set_document_code_page(offset, 1252)?,
            "mac" => self.set_document_code_page(offset, 10000)?,
            "pc" => self.set_document_code_page(offset, 437)?,
            "pca" => self.set_document_code_page(offset, 850)?,
            "ansicpg" => {
                let value = required_u16(offset, name, parameter)?;
                if !supports_code_page(value) {
                    return Err(rtf_error(
                        offset,
                        format!("unsupported RTF code page {value}"),
                    ));
                }
                self.set_document_code_page(offset, value)?;
            }
            "uc" => {
                let value = required_nonnegative(offset, name, parameter)?;
                if value > 32 {
                    return Err(rtf_error(offset, "Unicode fallback width exceeds 32"));
                }
                self.states.last_mut().unwrap().uc_skip = value;
            }
            "u" => self.unicode_control(offset, parameter)?,
            "fonttbl" => {
                self.header_table_started = true;
                self.states.last_mut().unwrap().destination = Destination::FontTable;
            }
            "colortbl" => {
                self.header_table_started = true;
                self.states.last_mut().unwrap().destination = Destination::ColorTable;
            }
            "listtable" => {
                self.header_table_started = true;
                self.states.last_mut().unwrap().destination = Destination::ListTable;
            }
            "listoverridetable" => {
                self.header_table_started = true;
                self.states.last_mut().unwrap().destination = Destination::ListOverrideTable;
            }
            "listtext" | "pntext" => {
                let state = self.states.last_mut().unwrap();
                state.destination = Destination::ListText;
                self.active_list_marker.clear();
                self.active_list_marker_unicode.clear();
            }
            "pict" => {
                let state = self.states.last_mut().unwrap();
                state.destination = Destination::Picture;
                self.active_picture = Some(PictureBuilder {
                    offset,
                    kind: None,
                    data: Vec::new(),
                    high_nibble: None,
                    width_goal: None,
                    height_goal: None,
                    scale_x: 100,
                    scale_y: 100,
                    crop_top: 0,
                    crop_bottom: 0,
                    crop_left: 0,
                    crop_right: 0,
                });
            }
            "shppict" => self.states.last_mut().unwrap().destination = Destination::Container,
            "upr" => {
                self.states.last_mut().unwrap().destination = Destination::UnicodeAlternatives;
            }
            "ud" => self.states.last_mut().unwrap().destination = Destination::Body,
            "nonshppict" => {
                self.states.last_mut().unwrap().destination = Destination::Skip(name.to_owned());
                self.diagnostics.push(RtfDiagnostic {
                    offset,
                    destination: Some(name.to_owned()),
                    message: "unsupported RTF destination skipped".to_owned(),
                });
            }
            "f" => self.font_control(offset, parameter)?,
            "fcharset" => {
                let charset = required_i32(offset, name, parameter)?;
                if !is_known_font_charset(charset) {
                    return Err(rtf_error(
                        offset,
                        format!("unsupported RTF font charset {charset}"),
                    ));
                }
                self.states.last_mut().unwrap().font_charset = Some(charset);
            }
            "cpg" => {
                let code_page = required_u16(offset, name, parameter)?;
                if !supports_code_page(code_page) {
                    return Err(rtf_error(
                        offset,
                        format!("unsupported RTF font code page {code_page}"),
                    ));
                }
                self.states.last_mut().unwrap().font_code_page = Some(code_page);
            }
            "red" | "green" | "blue" => self.color_control(offset, name, parameter)?,
            "b" => self.states.last_mut().unwrap().format.bold = Some(toggle(parameter)),
            "i" => self.states.last_mut().unwrap().format.italic = Some(toggle(parameter)),
            "ul" => self.states.last_mut().unwrap().format.underline = Some(toggle(parameter)),
            "ulnone" => self.states.last_mut().unwrap().format.underline = Some(false),
            "strike" => self.states.last_mut().unwrap().format.strike = Some(toggle(parameter)),
            "caps" => self.states.last_mut().unwrap().format.caps = Some(toggle(parameter)),
            "scaps" => self.states.last_mut().unwrap().format.small_caps = Some(toggle(parameter)),
            "v" => self.states.last_mut().unwrap().format.hidden = Some(toggle(parameter)),
            "super" => {
                self.states.last_mut().unwrap().format.vertical = Some(VerticalPosition::Super)
            }
            "sub" => self.states.last_mut().unwrap().format.vertical = Some(VerticalPosition::Sub),
            "nosupersub" => self.states.last_mut().unwrap().format.vertical = None,
            "fs" => {
                let half_points = required_nonnegative(offset, name, parameter)?;
                self.states.last_mut().unwrap().format.size_points = Some(half_points as f64 / 2.0);
            }
            "cf" => {
                self.states.last_mut().unwrap().format.color =
                    Some(required_nonnegative(offset, name, parameter)?);
            }
            "highlight" | "cb" => {
                self.states.last_mut().unwrap().format.highlight =
                    Some(required_nonnegative(offset, name, parameter)?);
            }
            "plain" => {
                self.states.last_mut().unwrap().format = RunFormat {
                    font: self.default_font,
                    ..RunFormat::default()
                };
            }
            "pard" => self.states.last_mut().unwrap().paragraph = ParagraphFormat::default(),
            "ql" => self.states.last_mut().unwrap().paragraph.alignment = Some(Alignment::Left),
            "qc" => self.states.last_mut().unwrap().paragraph.alignment = Some(Alignment::Center),
            "qr" => self.states.last_mut().unwrap().paragraph.alignment = Some(Alignment::Right),
            "qj" => self.states.last_mut().unwrap().paragraph.alignment = Some(Alignment::Justify),
            "li" => {
                self.states.last_mut().unwrap().paragraph.indent_left =
                    Some(required_i32(offset, name, parameter)?);
            }
            "ri" => {
                self.states.last_mut().unwrap().paragraph.indent_right =
                    Some(required_i32(offset, name, parameter)?);
            }
            "fi" => {
                self.states.last_mut().unwrap().paragraph.first_line_indent =
                    Some(required_i32(offset, name, parameter)?);
            }
            "sb" => {
                self.states.last_mut().unwrap().paragraph.space_before =
                    Some(required_i32(offset, name, parameter)?);
            }
            "sa" => {
                self.states.last_mut().unwrap().paragraph.space_after =
                    Some(required_i32(offset, name, parameter)?);
            }
            "sl" => {
                self.states.last_mut().unwrap().paragraph.line_spacing =
                    Some(required_i32(offset, name, parameter)?);
            }
            "slmult" => {
                self.states
                    .last_mut()
                    .unwrap()
                    .paragraph
                    .line_spacing_multiple = required_i32(offset, name, parameter)? != 0;
            }
            "ls" if self.in_destination(Destination::ListOverrideTable) => {
                self.active_override_id = Some(required_i32(offset, name, parameter)?);
            }
            "ls" => {
                let override_id = required_i32(offset, name, parameter)?;
                self.states.last_mut().unwrap().paragraph.list_override =
                    (override_id != 0).then_some(override_id);
            }
            "ilvl" => {
                let level = required_nonnegative(offset, name, parameter)?;
                if level > 8 {
                    return Err(rtf_error(offset, "RTF list level exceeds 8"));
                }
                self.states.last_mut().unwrap().paragraph.list_level = level as u32;
            }
            "par" => self.finish_paragraph(true)?,
            "line" => self.append_structural_character(offset, ParagraphItem::Break, '\n')?,
            "tab" => self.append_structural_character(offset, ParagraphItem::Tab, '\t')?,
            "emdash" => self.append_unicode(offset, '\u{2014}')?,
            "endash" => self.append_unicode(offset, '\u{2013}')?,
            "emspace" => self.append_unicode(offset, '\u{2003}')?,
            "enspace" => self.append_unicode(offset, '\u{2002}')?,
            "qmspace" => self.append_unicode(offset, '\u{2005}')?,
            "bullet" => self.append_unicode(offset, '\u{2022}')?,
            "lquote" => self.append_unicode(offset, '\u{2018}')?,
            "rquote" => self.append_unicode(offset, '\u{2019}')?,
            "ldblquote" => self.append_unicode(offset, '\u{201c}')?,
            "rdblquote" => self.append_unicode(offset, '\u{201d}')?,
            "trowd" => self.start_row(offset)?,
            "cellx" => {
                if !self.in_table {
                    return Err(rtf_error(
                        offset,
                        "table cell boundary appears outside a row",
                    ));
                }
                let boundary = required_i32(offset, name, parameter)?;
                self.expected_cells = self.expected_cells.saturating_add(1);
                if self.expected_cells > MAX_TABLE_COLUMNS {
                    return Err(rtf_error(offset, "RTF table exceeds the column limit"));
                }
                if self
                    .current_cell_boundaries
                    .last()
                    .is_some_and(|previous| boundary <= *previous)
                {
                    return Err(rtf_error(
                        offset,
                        "RTF table cell boundaries are not increasing",
                    ));
                }
                self.current_cell_boundaries.push(boundary);
            }
            "cell" | "nestcell" => self.finish_cell()?,
            "row" | "nestrow" => self.finish_row(offset)?,
            "intbl" => self.in_table = true,
            "list" if self.in_destination(Destination::ListTable) => {
                self.active_list_levels.clear();
            }
            "listlevel" if self.in_destination(Destination::ListTable) => {
                if self.active_list_levels.len() < 9 {
                    self.active_list_levels.push(ListLevelData::default());
                }
            }
            "lfolevel" if self.in_destination(Destination::ListOverrideTable) => {
                if self.active_override_levels.len() < 9 {
                    self.active_override_levels
                        .push(ListLevelOverrideData::default());
                }
            }
            "levelnfcn" if self.in_destination(Destination::ListOverrideTable) => {
                let value = required_i32(offset, name, parameter)?;
                let format = self.list_format_or_diagnose(offset, value);
                if let Some(level) = self.active_override_levels.last_mut() {
                    level.format = Some(format);
                    level.has_new_format = true;
                }
            }
            "levelnfc" if self.in_destination(Destination::ListOverrideTable) => {
                let value = required_i32(offset, name, parameter)?;
                let format = self.list_format_or_diagnose(offset, value);
                if let Some(level) = self.active_override_levels.last_mut()
                    && !level.has_new_format
                {
                    level.format = Some(format);
                }
            }
            "levelnfcn" if self.in_destination(Destination::ListTable) => {
                let value = required_i32(offset, name, parameter)?;
                let format = self.list_format_or_diagnose(offset, value);
                if let Some(level) = self.active_list_levels.last_mut() {
                    level.format = format;
                    level.has_new_format = true;
                }
            }
            "levelnfc" if self.in_destination(Destination::ListTable) => {
                let value = required_i32(offset, name, parameter)?;
                let format = self.list_format_or_diagnose(offset, value);
                if let Some(level) = self.active_list_levels.last_mut()
                    && !level.has_new_format
                {
                    level.format = format;
                }
            }
            "levelstartat" if self.in_destination(Destination::ListOverrideTable) => {
                let value = required_nonnegative(offset, name, parameter)? as u32;
                if let Some(level) = self.active_override_levels.last_mut() {
                    level.start = Some(value);
                }
            }
            "levelstartat" if self.in_destination(Destination::ListTable) => {
                let value = required_nonnegative(offset, name, parameter)? as u32;
                if let Some(level) = self.active_list_levels.last_mut() {
                    level.start = Some(value);
                }
            }
            "listid" if self.in_destination(Destination::ListOverrideTable) => {
                self.active_override_list = Some(required_i32(offset, name, parameter)?);
            }
            "listid" if self.in_destination(Destination::ListTable) => {
                let id = required_i32(offset, name, parameter)?;
                let levels = if self.active_list_levels.is_empty() {
                    vec![ListLevelData::default()]
                } else {
                    self.active_list_levels.clone()
                };
                self.insert_list(offset, id, levels)?;
            }
            "listoverride" if self.in_destination(Destination::ListOverrideTable) => {
                if self.active_override_id.is_some() {
                    self.finish_list_override(offset)?;
                }
                self.active_override_list = None;
                self.active_override_id = None;
                self.active_override_levels.clear();
                self.states.last_mut().unwrap().list_override_entry = true;
            }
            "pngblip" => self.picture_kind(PictureKind::Png),
            "jpegblip" => self.picture_kind(PictureKind::Jpeg),
            "emfblip" | "macpict" => self.picture_kind(PictureKind::Unsupported),
            "pmmetafile" | "wmetafile" | "dibitmap" | "wbitmap" => {
                let _ = required_i32(offset, name, parameter)?;
                self.picture_kind(PictureKind::Unsupported);
            }
            "picwgoal" => {
                self.picture_mut(offset)?.width_goal = Some(required_i32(offset, name, parameter)?);
            }
            "pichgoal" => {
                self.picture_mut(offset)?.height_goal =
                    Some(required_i32(offset, name, parameter)?);
            }
            "picscalex" => {
                self.picture_mut(offset)?.scale_x = required_nonnegative(offset, name, parameter)?;
            }
            "picscaley" => {
                self.picture_mut(offset)?.scale_y = required_nonnegative(offset, name, parameter)?;
            }
            "piccropt" => {
                self.picture_mut(offset)?.crop_top = required_i32(offset, name, parameter)?;
            }
            "piccropb" => {
                self.picture_mut(offset)?.crop_bottom = required_i32(offset, name, parameter)?;
            }
            "piccropl" => {
                self.picture_mut(offset)?.crop_left = required_i32(offset, name, parameter)?;
            }
            "piccropr" => {
                self.picture_mut(offset)?.crop_right = required_i32(offset, name, parameter)?;
            }
            "picw" | "pich" | "bliptag" | "blipupi" | "wbmbitspixel" | "wbmplanes"
            | "wbmwidthbytes" => {
                let _ = required_i32(offset, name, parameter)?;
            }
            name @ ("leveljc" | "levelfollow" | "levelspace" | "levelindent") => {
                let _ = required_i32(offset, name, parameter)?;
                self.diagnostics.push(RtfDiagnostic {
                    offset,
                    destination: Some("listtable".to_owned()),
                    message: format!("RTF list property \\{name} was dropped"),
                });
            }
            "deff" => {
                let id = required_i32(offset, name, parameter)?;
                self.default_font = Some(id);
                self.states.last_mut().unwrap().format.font = Some(id);
            }
            name @ ("paperw" | "paperh" | "margl" | "margr" | "margt" | "margb" | "gutter") => {
                let _ = required_i32(offset, name, parameter)?;
                self.diagnostics.push(RtfDiagnostic {
                    offset,
                    destination: None,
                    message: format!("RTF page geometry control \\{name} was dropped"),
                });
            }
            name @ ("deflang" | "deflangfe" | "sectd" | "lang" | "langfe" | "loch" | "hich"
            | "dbch" | "af" | "rtlch" | "ltrch" | "widowctrl" | "nowidctlpar"
            | "hyphauto" | "formshade" | "headery" | "footery" | "endnhere") => {
                self.diagnostics.push(RtfDiagnostic {
                    offset,
                    destination: None,
                    message: format!("RTF document formatting control \\{name} was dropped"),
                });
            }
            name @ ("listsimple"
            | "listhybrid"
            | "listoverridecount"
            | "listoverrideformat"
            | "listoverridestartat") => {
                self.diagnostics.push(RtfDiagnostic {
                    offset,
                    destination: Some("listtable".to_owned()),
                    message: format!("RTF list property \\{name} was dropped"),
                });
            }
            "viewkind" | "viewscale" | "fet" | "nouicompat" | "listtemplateid" | "listname" => {}
            name if is_unsupported_destination(name) => {
                self.states.last_mut().unwrap().destination = Destination::Skip(name.to_owned());
                self.diagnostics.push(RtfDiagnostic {
                    offset,
                    destination: Some(name.to_owned()),
                    message: "unsupported RTF destination skipped".to_owned(),
                });
            }
            _ => self.diagnostics.push(RtfDiagnostic {
                offset,
                destination: None,
                message: format!("unsupported RTF control word \\{name} ignored"),
            }),
        }
        Ok(())
    }

    fn font_control(&mut self, offset: usize, parameter: Option<i32>) -> Result<()> {
        let id = required_i32(offset, "f", parameter)?;
        if self.in_destination(Destination::FontTable) {
            self.states.last_mut().unwrap().font_id = Some(id);
        } else {
            self.states.last_mut().unwrap().format.font = Some(id);
        }
        Ok(())
    }

    fn color_control(&mut self, offset: usize, name: &str, parameter: Option<i32>) -> Result<()> {
        let value = required_i32(offset, name, parameter)?;
        let value = u8::try_from(value)
            .map_err(|_| rtf_error(offset, format!("{name} value is outside 0 through 255")))?;
        let state = self.states.last_mut().unwrap();
        match name {
            "red" => state.color_red = Some(value),
            "green" => state.color_green = Some(value),
            "blue" => state.color_blue = Some(value),
            _ => unreachable!(),
        }
        Ok(())
    }

    fn unicode_control(&mut self, offset: usize, parameter: Option<i32>) -> Result<()> {
        let value = required_i32(offset, "u", parameter)?;
        let code = if value < 0 { value + 65_536 } else { value };
        let code = u16::try_from(code)
            .map_err(|_| rtf_error(offset, "Unicode control is outside UTF-16"))?;
        let state = self.states.last_mut().unwrap();
        let character = if (0xd800..=0xdbff).contains(&code) {
            if state.pending_high_surrogate.replace(code).is_some() {
                return Err(rtf_error(offset, "consecutive Unicode high surrogates"));
            }
            None
        } else if (0xdc00..=0xdfff).contains(&code) {
            let high = state
                .pending_high_surrogate
                .take()
                .ok_or_else(|| rtf_error(offset, "Unicode low surrogate has no high surrogate"))?;
            let scalar =
                0x10000 + (((u32::from(high) - 0xd800) << 10) | (u32::from(code) - 0xdc00));
            Some(char::from_u32(scalar).unwrap())
        } else {
            if state.pending_high_surrogate.is_some() {
                return Err(rtf_error(
                    offset,
                    "Unicode high surrogate has no low surrogate",
                ));
            }
            Some(char::from_u32(u32::from(code)).unwrap())
        };
        state.skip_remaining = state.uc_skip;
        if let Some(character) = character {
            self.append_unicode(offset, character)?;
        }
        Ok(())
    }

    fn raw_text_byte(&mut self, offset: usize, byte: u8) -> Result<()> {
        self.states.last_mut().unwrap().at_group_start = false;
        let destination = self.states.last().unwrap().destination.clone();
        match destination {
            Destination::Picture => self.picture_hex_byte(offset, byte),
            Destination::FontTable => {
                self.retain_output(offset, 3)?;
                self.states.last_mut().unwrap().font_name.push(byte);
                Ok(())
            }
            Destination::ListText => {
                self.retain_output(offset, 3)?;
                self.active_list_marker.push(byte);
                Ok(())
            }
            Destination::Body => {
                self.retain_output(offset, 3)?;
                let state = self.states.last_mut().unwrap();
                if state.text_bytes.is_empty() {
                    state.text_offset = offset;
                }
                state.text_bytes.push(byte);
                Ok(())
            }
            _ => Ok(()),
        }
    }

    fn text(&mut self, offset: usize, bytes: &[u8]) -> Result<()> {
        if !bytes.is_empty() {
            self.states.last_mut().unwrap().at_group_start = false;
        }
        match self.states.last().unwrap().destination.clone() {
            Destination::Body => {
                self.retain_output(offset, bytes.len().saturating_mul(3))?;
                let state = self.states.last_mut().unwrap();
                if state.pending_high_surrogate.is_some() && !bytes.is_empty() {
                    return Err(rtf_error(
                        offset,
                        "Unicode high surrogate has no low surrogate",
                    ));
                }
                if state.text_bytes.is_empty() {
                    state.text_offset = offset;
                }
                state.text_bytes.extend_from_slice(bytes);
            }
            Destination::FontTable => {
                self.retain_output(offset, bytes.len().saturating_mul(3))?;
                self.font_table_text(offset, bytes)?;
            }
            Destination::ColorTable => self.color_table_text(offset, bytes)?,
            Destination::ListText => {
                self.retain_output(offset, bytes.len().saturating_mul(3))?;
                self.active_list_marker.extend_from_slice(bytes);
            }
            Destination::Picture => self.picture_text(offset, bytes)?,
            Destination::ListTable
            | Destination::ListOverrideTable
            | Destination::Container
            | Destination::UnicodeAlternatives
            | Destination::Skip(_) => {}
        }
        Ok(())
    }

    fn binary(&mut self, offset: usize, bytes: &[u8]) -> Result<()> {
        if self
            .states
            .last()
            .is_some_and(|state| state.pending_high_surrogate.is_some())
        {
            return Err(rtf_error(
                offset,
                "RTF binary control interrupts a Unicode surrogate pair",
            ));
        }
        self.states.last_mut().unwrap().at_group_start = false;
        if self.states.last().unwrap().destination != Destination::Picture {
            self.diagnostics.push(RtfDiagnostic {
                offset,
                destination: None,
                message: "binary RTF payload outside a picture was ignored".to_owned(),
            });
            return Ok(());
        }
        self.retain_output(offset, bytes.len())?;
        let picture = self.picture_mut(offset)?;
        if picture.data.len().saturating_add(bytes.len()) > MAX_PICTURE_BYTES {
            return Err(rtf_error(offset, "RTF picture exceeds the size limit"));
        }
        picture.data.extend_from_slice(bytes);
        Ok(())
    }

    fn flush_state(&mut self, state: &mut State) -> Result<()> {
        if state.text_bytes.is_empty() {
            return Ok(());
        }
        let bytes = std::mem::take(&mut state.text_bytes);
        match state.destination {
            Destination::Body => {
                if state.pending_high_surrogate.is_some() {
                    return Err(rtf_error(
                        state.text_offset,
                        "Unicode high surrogate has no low surrogate",
                    ));
                }
                let code_page = state
                    .format
                    .font
                    .and_then(|id| self.fonts.get(&id))
                    .and_then(|font| font.code_page)
                    .unwrap_or(state.code_page);
                let text = decode(code_page, &bytes).ok_or_else(|| {
                    rtf_error(
                        state.text_offset,
                        format!("unsupported RTF code page {code_page}"),
                    )
                })?;
                if text.1 {
                    self.diagnostics.push(RtfDiagnostic {
                        offset: state.text_offset,
                        destination: None,
                        message: format!(
                            "invalid byte sequence replaced for code page {code_page}"
                        ),
                    });
                }
                self.append_text(state.text_offset, &text.0, &state.format, &state.paragraph)?;
            }
            Destination::FontTable => state.font_name.extend(bytes),
            Destination::ListText => self.active_list_marker.extend(bytes),
            _ => {}
        }
        Ok(())
    }

    fn append_unicode(&mut self, offset: usize, character: char) -> Result<()> {
        let mut state = self.states.pop().unwrap();
        self.flush_state(&mut state)?;
        let destination = state.destination.clone();
        self.states.push(state);
        match destination {
            Destination::Body => {
                self.body_started = true;
                self.retain_output(offset, character.len_utf8())?;
                let state = self.states.last().unwrap();
                let format = state.format.clone();
                let paragraph = state.paragraph.clone();
                self.append_text(offset, &character.to_string(), &format, &paragraph)?;
            }
            Destination::ListText => {
                self.retain_output(offset, character.len_utf8())?;
                self.active_list_marker_unicode.push(character);
            }
            _ => {
                self.diagnostics.push(RtfDiagnostic {
                    offset,
                    destination: Some(destination_name(&destination)),
                    message: "Unicode text in an unsupported RTF destination was ignored"
                        .to_owned(),
                });
            }
        }
        Ok(())
    }

    fn append_text(
        &mut self,
        offset: usize,
        text: &str,
        format: &RunFormat,
        paragraph: &ParagraphFormat,
    ) -> Result<()> {
        if text.is_empty() {
            return Ok(());
        }
        self.body_started = true;
        if !self.in_table && !self.table_rows.is_empty() {
            self.flush_table()?;
        }
        self.current_paragraph.format = paragraph.clone();
        if let Some(ParagraphItem::Text(last)) = self.current_paragraph.items.last_mut()
            && last.format == *format
        {
            last.text.push_str(text);
        } else {
            self.reserve_run(offset)?;
            self.current_paragraph
                .items
                .push(ParagraphItem::Text(RunData {
                    text: text.to_owned(),
                    format: format.clone(),
                }));
        }
        Ok(())
    }

    fn append_special(&mut self, item: ParagraphItem) -> Result<()> {
        self.body_started = true;
        if !self.in_table && !self.table_rows.is_empty() {
            self.flush_table()?;
        }
        self.reserve_run(0)?;
        self.current_paragraph.format = self.states.last().unwrap().paragraph.clone();
        self.current_paragraph.items.push(item);
        Ok(())
    }

    fn append_structural_character(
        &mut self,
        offset: usize,
        item: ParagraphItem,
        character: char,
    ) -> Result<()> {
        if self.in_destination(Destination::Body) {
            self.append_special(item)
        } else {
            self.append_unicode(offset, character)
        }
    }

    fn finish_paragraph(&mut self, explicit: bool) -> Result<()> {
        if explicit {
            self.body_started = true;
        }
        let state = self.states.last().unwrap();
        self.current_paragraph.format = state.paragraph.clone();
        if self.in_table {
            if explicit || self.current_paragraph.has_content() {
                self.reserve_block(0)?;
                self.current_cell
                    .push(std::mem::take(&mut self.current_paragraph));
            }
        } else {
            if !self.table_rows.is_empty() {
                self.flush_table()?;
            }
            if explicit || self.current_paragraph.has_content() {
                self.reserve_block(0)?;
                self.blocks.push(Block::Paragraph(std::mem::take(
                    &mut self.current_paragraph,
                )));
            }
        }
        Ok(())
    }

    fn start_row(&mut self, offset: usize) -> Result<()> {
        if self.in_table {
            return Err(rtf_error(offset, "nested or unfinished RTF table row"));
        }
        if self.current_paragraph.has_content() {
            self.finish_paragraph(false)?;
        }
        self.body_started = true;
        self.in_table = true;
        self.current_row.clear();
        self.current_cell.clear();
        self.current_cell_boundaries.clear();
        self.expected_cells = 0;
        Ok(())
    }

    fn finish_cell(&mut self) -> Result<()> {
        if !self.in_table {
            return Err(rtf_error(0, "table cell appears outside a row"));
        }
        if self.current_paragraph.has_content() || self.current_cell.is_empty() {
            self.reserve_block(0)?;
            self.current_cell
                .push(std::mem::take(&mut self.current_paragraph));
        }
        self.current_row
            .push(std::mem::take(&mut self.current_cell));
        if self.current_row.len() > MAX_TABLE_COLUMNS {
            return Err(rtf_error(0, "RTF table exceeds the column limit"));
        }
        Ok(())
    }

    fn finish_row(&mut self, offset: usize) -> Result<()> {
        if !self.in_table {
            return Err(rtf_error(
                offset,
                "table row terminator appears outside a row",
            ));
        }
        if self.current_paragraph.has_content() || !self.current_cell.is_empty() {
            self.finish_cell()?;
        }
        if self.current_row.len() != self.expected_cells {
            return Err(rtf_error(
                offset,
                "RTF table row cell count does not match its boundaries",
            ));
        }
        let new_cells = self.current_row.len();
        if self.total_cells.saturating_add(new_cells) > MAX_TABLE_CELLS {
            return Err(rtf_error(offset, "RTF table cell limit exceeded"));
        }
        self.total_cells += new_cells;
        if self
            .table_rows
            .first()
            .is_some_and(|row| row.boundaries != self.current_cell_boundaries)
        {
            self.diagnostics.push(RtfDiagnostic {
                offset,
                destination: Some("table".to_owned()),
                message: "RTF table row boundaries differ from the first row".to_owned(),
            });
        }
        self.table_rows.push(TableRowData {
            cells: std::mem::take(&mut self.current_row),
            boundaries: std::mem::take(&mut self.current_cell_boundaries),
        });
        if self.table_rows.len() > MAX_TABLE_ROWS {
            return Err(rtf_error(offset, "RTF table exceeds the row limit"));
        }
        self.in_table = false;
        Ok(())
    }

    fn flush_table(&mut self) -> Result<()> {
        if !self.table_rows.is_empty() {
            self.reserve_block(0)?;
            self.blocks
                .push(Block::Table(std::mem::take(&mut self.table_rows)));
        }
        Ok(())
    }

    fn finish_document(&mut self) -> Result<()> {
        if self.in_table {
            return Err(rtf_error(0, "RTF table row is not terminated"));
        }
        if self.current_paragraph.has_content()
            || (self.blocks.is_empty() && self.table_rows.is_empty())
        {
            self.reserve_block(0)?;
            self.blocks.push(Block::Paragraph(std::mem::take(
                &mut self.current_paragraph,
            )));
        }
        self.flush_table()?;
        Ok(())
    }

    fn font_table_text(&mut self, offset: usize, bytes: &[u8]) -> Result<()> {
        for &byte in bytes {
            if byte == b';' {
                let state = self.states.last_mut().unwrap();
                let id = state.font_id.take();
                let charset = state.font_charset.take();
                let explicit_code_page = state.font_code_page.take();
                let name_bytes = std::mem::take(&mut state.font_name);
                let Some(id) = id else {
                    continue;
                };
                let name = decode(state.code_page, &name_bytes)
                    .map(|value| value.0.trim().to_owned())
                    .filter(|name| !name.is_empty())
                    .ok_or_else(|| rtf_error(offset, "font table entry has no valid name"))?;
                let code_page = match (explicit_code_page, charset) {
                    (Some(code_page), _) => Some(code_page),
                    (None, Some(1)) | (None, None) => None,
                    (None, Some(130)) => {
                        return Err(rtf_error(offset, "RTF Johab font charset is unsupported"));
                    }
                    (None, Some(charset)) => charset_code_page(charset),
                };
                if self.fonts.len() >= MAX_LOOKUP_ENTRIES && !self.fonts.contains_key(&id) {
                    return Err(rtf_error(offset, "RTF font table exceeds the entry limit"));
                }
                self.fonts.insert(id, FontEntry { name, code_page });
            } else {
                self.states.last_mut().unwrap().font_name.push(byte);
            }
        }
        Ok(())
    }

    fn color_table_text(&mut self, offset: usize, bytes: &[u8]) -> Result<()> {
        for &byte in bytes {
            if byte != b';' {
                continue;
            }
            if self.colors.len() >= MAX_LOOKUP_ENTRIES {
                return Err(rtf_error(offset, "RTF color table exceeds the entry limit"));
            }
            let state = self.states.last_mut().unwrap();
            let color = match (state.color_red, state.color_green, state.color_blue) {
                (None, None, None) => None,
                (red, green, blue) => Some(format!(
                    "{:02X}{:02X}{:02X}",
                    red.unwrap_or(0),
                    green.unwrap_or(0),
                    blue.unwrap_or(0)
                )),
            };
            self.colors.push(color);
            state.color_red = None;
            state.color_green = None;
            state.color_blue = None;
        }
        Ok(())
    }

    fn picture_text(&mut self, offset: usize, bytes: &[u8]) -> Result<()> {
        for &byte in bytes {
            if byte.is_ascii_whitespace() {
                continue;
            }
            let nibble = hex_value(byte)
                .ok_or_else(|| rtf_error(offset, "invalid hexadecimal picture data"))?;
            self.picture_hex_nibble(offset, nibble)?;
        }
        Ok(())
    }

    fn picture_hex_byte(&mut self, offset: usize, byte: u8) -> Result<()> {
        self.picture_hex_nibble(offset, byte >> 4)?;
        self.picture_hex_nibble(offset, byte & 0x0f)
    }

    fn picture_hex_nibble(&mut self, offset: usize, nibble: u8) -> Result<()> {
        if self
            .active_picture
            .as_ref()
            .is_some_and(|picture| picture.high_nibble.is_some())
        {
            self.retain_output(offset, 1)?;
        }
        let picture = self.picture_mut(offset)?;
        if let Some(high) = picture.high_nibble.take() {
            if picture.data.len() >= MAX_PICTURE_BYTES {
                return Err(rtf_error(offset, "RTF picture exceeds the size limit"));
            }
            picture.data.push(high << 4 | nibble);
        } else {
            picture.high_nibble = Some(nibble);
        }
        Ok(())
    }

    fn picture_mut(&mut self, offset: usize) -> Result<&mut PictureBuilder> {
        self.active_picture
            .as_mut()
            .ok_or_else(|| rtf_error(offset, "picture property appears outside a picture"))
    }

    fn picture_kind(&mut self, kind: PictureKind) {
        if let Some(picture) = self.active_picture.as_mut() {
            picture.kind = Some(kind);
        }
    }

    fn finish_picture(&mut self, picture: Option<PictureBuilder>, offset: usize) -> Result<()> {
        self.body_started = true;
        let picture =
            picture.ok_or_else(|| rtf_error(offset, "picture destination has no state"))?;
        if picture.high_nibble.is_some() {
            return Err(rtf_error(
                offset,
                "picture data has an odd number of hex digits",
            ));
        }
        if picture.data.is_empty() {
            return Err(rtf_error(offset, "picture destination has no image data"));
        }
        let kind = picture.kind.unwrap_or(PictureKind::Unsupported);
        if kind == PictureKind::Unsupported {
            self.diagnostics.push(RtfDiagnostic {
                offset: picture.offset,
                destination: Some("pict".to_owned()),
                message: "unsupported RTF picture type skipped".to_owned(),
            });
            return Ok(());
        }
        let probed = oxml_media::probe(&picture.data)
            .ok_or_else(|| rtf_error(picture.offset, "invalid RTF image payload"))?;
        let expected = match kind {
            PictureKind::Png => oxml_media::ImageFormat::Png,
            PictureKind::Jpeg => oxml_media::ImageFormat::Jpeg,
            PictureKind::Unsupported => unreachable!(),
        };
        if probed.format != expected {
            return Err(rtf_error(
                picture.offset,
                "RTF picture type does not match its payload",
            ));
        }
        if [
            picture.crop_top,
            picture.crop_bottom,
            picture.crop_left,
            picture.crop_right,
        ]
        .into_iter()
        .any(|crop| crop != 0)
        {
            self.diagnostics.push(RtfDiagnostic {
                offset: picture.offset,
                destination: Some("pict".to_owned()),
                message: "RTF picture cropping was dropped".to_owned(),
            });
        }
        if !self.in_table && !self.table_rows.is_empty() {
            self.flush_table()?;
        }
        self.reserve_run(picture.offset)?;
        self.current_paragraph.format = self.states.last().unwrap().paragraph.clone();
        self.current_paragraph
            .items
            .push(ParagraphItem::Picture(PictureData {
                kind,
                data: picture.data,
                width_goal: picture.width_goal,
                height_goal: picture.height_goal,
                scale_x: picture.scale_x,
                scale_y: picture.scale_y,
            }));
        Ok(())
    }

    fn finish_list_marker(&mut self, code_page: u16) -> Result<()> {
        if self.active_list_marker.is_empty() && self.active_list_marker_unicode.is_empty() {
            return Ok(());
        }
        let decoded = decode(code_page, &self.active_list_marker)
            .map(|value| value.0)
            .unwrap_or_else(|| Cow::Borrowed(""));
        let format = if self.active_list_marker_unicode.contains('•')
            || self.active_list_marker_unicode.contains('·')
            || decoded.contains('•')
            || decoded.contains('·')
            || self.active_list_marker.contains(&0xb7)
        {
            ListNumberFormat::Bullet
        } else {
            ListNumberFormat::Decimal
        };
        self.current_paragraph.inferred_list = Some(format);
        self.active_list_marker.clear();
        self.active_list_marker_unicode.clear();
        Ok(())
    }

    fn insert_list(&mut self, offset: usize, id: i32, levels: Vec<ListLevelData>) -> Result<()> {
        if self.lists.len() >= MAX_LOOKUP_ENTRIES && !self.lists.contains_key(&id) {
            return Err(rtf_error(offset, "RTF list table exceeds the entry limit"));
        }
        self.lists.insert(id, levels);
        Ok(())
    }

    fn insert_override(
        &mut self,
        offset: usize,
        id: i32,
        override_data: ListOverrideData,
    ) -> Result<()> {
        if self.overrides.len() >= MAX_LOOKUP_ENTRIES && !self.overrides.contains_key(&id) {
            return Err(rtf_error(
                offset,
                "RTF list override table exceeds the entry limit",
            ));
        }
        self.overrides.insert(id, override_data);
        Ok(())
    }

    fn finish_list_override(&mut self, offset: usize) -> Result<()> {
        let override_id = self
            .active_override_id
            .take()
            .ok_or_else(|| rtf_error(offset, "RTF list override is missing its \\ls identifier"))?;
        let list_id = self.active_override_list.take().ok_or_else(|| {
            rtf_error(
                offset,
                "RTF list override is missing its \\listid reference",
            )
        })?;
        let levels = std::mem::take(&mut self.active_override_levels);
        self.insert_override(offset, override_id, ListOverrideData { list_id, levels })
    }

    fn list_format_or_diagnose(&mut self, offset: usize, value: i32) -> ListNumberFormat {
        if let Some(format) = list_format(value) {
            return format;
        }
        self.diagnostics.push(RtfDiagnostic {
            offset,
            destination: Some("listtable".to_owned()),
            message: format!("unsupported RTF list number format {value} converted to decimal"),
        });
        ListNumberFormat::Decimal
    }

    fn reserve_block(&mut self, offset: usize) -> Result<()> {
        if self.total_blocks >= MAX_BLOCKS {
            return Err(rtf_error(offset, "RTF block limit exceeded"));
        }
        self.total_blocks += 1;
        Ok(())
    }

    fn set_document_code_page(&mut self, offset: usize, code_page: u16) -> Result<()> {
        if self.states.len() != 1 || self.body_started || self.header_table_started {
            return Err(rtf_error(
                offset,
                "RTF character-set declaration appears outside the document header",
            ));
        }
        self.states.last_mut().unwrap().code_page = code_page;
        Ok(())
    }

    fn reserve_run(&mut self, offset: usize) -> Result<()> {
        if self.total_runs >= MAX_RUNS {
            return Err(rtf_error(offset, "RTF run limit exceeded"));
        }
        self.total_runs += 1;
        Ok(())
    }

    fn retain_output(&mut self, offset: usize, bytes: usize) -> Result<()> {
        let retained = self
            .retained_output_bytes
            .checked_add(bytes)
            .ok_or_else(|| rtf_error(offset, "RTF retained output byte count overflows"))?;
        if retained > MAX_RETAINED_OUTPUT_BYTES {
            return Err(rtf_error(
                offset,
                "RTF retained output exceeds the size limit",
            ));
        }
        self.retained_output_bytes = retained;
        Ok(())
    }

    fn in_destination(&self, destination: Destination) -> bool {
        self.states.last().unwrap().destination == destination
    }

    fn awaiting_unicode_destination(&self) -> bool {
        let Some(current) = self.states.last() else {
            return false;
        };
        let Some(parent) = self.states.iter().rev().nth(1) else {
            return false;
        };
        unicode_child_awaits_ud(current, parent)
    }

    fn awaiting_unicode_destination_with(&self, current: &State) -> bool {
        self.states
            .last()
            .is_some_and(|parent| unicode_child_awaits_ud(current, parent))
    }
}

fn unicode_child_awaits_ud(current: &State, parent: &State) -> bool {
    current.destination == Destination::Container
        && current.at_group_start
        && parent.destination == Destination::UnicodeAlternatives
        && parent.unicode_alternative_children == 2
}

fn token_offset(token: &Token<'_>) -> usize {
    match token {
        Token::Open { offset }
        | Token::Close { offset }
        | Token::Word { offset, .. }
        | Token::Symbol { offset, .. }
        | Token::Hex { offset, .. }
        | Token::Binary { offset, .. }
        | Token::Text { offset, .. } => *offset,
    }
}

fn toggle(parameter: Option<i32>) -> bool {
    parameter != Some(0)
}

fn required_i32(offset: usize, name: &str, parameter: Option<i32>) -> Result<i32> {
    parameter.ok_or_else(|| rtf_error(offset, format!("\\{name} requires a parameter")))
}

fn required_nonnegative(offset: usize, name: &str, parameter: Option<i32>) -> Result<usize> {
    usize::try_from(required_i32(offset, name, parameter)?)
        .map_err(|_| rtf_error(offset, format!("\\{name} requires a nonnegative parameter")))
}

fn required_u16(offset: usize, name: &str, parameter: Option<i32>) -> Result<u16> {
    u16::try_from(required_i32(offset, name, parameter)?).map_err(|_| {
        rtf_error(
            offset,
            format!("\\{name} parameter is outside 0 through 65535"),
        )
    })
}

fn charset_code_page(charset: i32) -> Option<u16> {
    match charset {
        0 => Some(1252),
        2 => Some(42),
        77 => Some(10000),
        128 => Some(932),
        129 => Some(949),
        130 => Some(1361),
        134 => Some(936),
        136 => Some(950),
        161 => Some(1253),
        162 => Some(1254),
        163 => Some(1258),
        177 => Some(1255),
        178 => Some(1256),
        186 => Some(1257),
        204 => Some(1251),
        222 => Some(874),
        238 => Some(1250),
        254 => Some(437),
        255 => Some(850),
        _ => None,
    }
}

fn is_known_font_charset(charset: i32) -> bool {
    matches!(charset, 1 | 130) || charset_code_page(charset).is_some()
}

fn encoding_for_code_page(code_page: u16) -> Option<&'static Encoding> {
    let label: &[u8] = match code_page {
        819 | 1252 => b"windows-1252",
        874 => b"windows-874",
        932 => b"shift_jis",
        936 => b"gbk",
        949 => b"euc-kr",
        950 => b"big5",
        1250 => b"windows-1250",
        1251 => b"windows-1251",
        1253 => b"windows-1253",
        1254 => b"windows-1254",
        1255 => b"windows-1255",
        1256 => b"windows-1256",
        1257 => b"windows-1257",
        1258 => b"windows-1258",
        866 => b"ibm866",
        _ => return None,
    };
    Encoding::for_label(label)
}

fn decode(code_page: u16, bytes: &[u8]) -> Option<(Cow<'_, str>, bool)> {
    if code_page == 42 {
        return Some((Cow::Owned(decode_symbol_bytes(bytes)), false));
    }
    if let Some(table) = legacy_code_page(code_page) {
        let mut decoded = String::with_capacity(bytes.len());
        for &byte in bytes {
            let scalar = if byte < 0x80 {
                u32::from(byte)
            } else {
                table[usize::from(byte - 0x80)]
            };
            decoded.push(char::from_u32(scalar).unwrap());
        }
        return Some((Cow::Owned(decoded), false));
    }
    let encoding = encoding_for_code_page(code_page)?;
    Some(encoding.decode_without_bom_handling(bytes))
}

fn supports_code_page(code_page: u16) -> bool {
    code_page == 42
        || legacy_code_page(code_page).is_some()
        || encoding_for_code_page(code_page).is_some()
}

fn decode_symbol_bytes(bytes: &[u8]) -> String {
    let mut decoded = String::with_capacity(bytes.len());
    for &byte in bytes {
        decoded.push(symbol_char(byte));
    }
    decoded
}

fn symbol_char(byte: u8) -> char {
    match byte {
        0x00..=0x7f => char::from(byte),
        0xa3 => '\u{2264}',
        0xa5 => '\u{221e}',
        0xb0 => '\u{00b0}',
        0xb1 => '\u{00b1}',
        0xb3 => '\u{2265}',
        0xb4 => '\u{00d7}',
        0xb5 => '\u{221d}',
        0xb6 => '\u{2202}',
        0xb7 => '\u{2022}',
        0xb8 => '\u{00f7}',
        0xb9 => '\u{2260}',
        0xba => '\u{2261}',
        0xbb => '\u{2248}',
        0xbc => '\u{2026}',
        0xbd => '\u{23d0}',
        0xbe => '\u{23af}',
        0xbf => '\u{21b5}',
        0xc0 => '\u{2135}',
        0xc1 => '\u{2111}',
        0xc2 => '\u{211c}',
        0xc3 => '\u{2118}',
        0xc4 => '\u{2297}',
        0xc5 => '\u{2295}',
        0xc6 => '\u{2205}',
        0xc7 => '\u{2229}',
        0xc8 => '\u{222a}',
        0xc9 => '\u{2283}',
        0xca => '\u{2287}',
        0xcb => '\u{2284}',
        0xcc => '\u{2282}',
        0xcd => '\u{2286}',
        0xce => '\u{2208}',
        0xcf => '\u{2209}',
        0xd0 => '\u{2220}',
        0xd1 => '\u{2207}',
        0xd2 => '\u{00ae}',
        0xd3 => '\u{00a9}',
        0xd4 => '\u{2122}',
        0xd5 => '\u{220f}',
        0xd6 => '\u{221a}',
        0xd7 => '\u{22c5}',
        0xd8 => '\u{00ac}',
        0xd9 => '\u{2227}',
        0xda => '\u{2228}',
        0xdb => '\u{21d4}',
        0xdc => '\u{21d0}',
        0xdd => '\u{21d1}',
        0xde => '\u{21d2}',
        0xdf => '\u{21d3}',
        0xe0 => '\u{25ca}',
        0xe1 => '\u{2329}',
        0xe2 => '\u{00ae}',
        0xe3 => '\u{00a9}',
        0xe4 => '\u{2122}',
        0xe5 => '\u{2211}',
        0xe6 => '\u{239b}',
        0xe7 => '\u{239c}',
        0xe8 => '\u{239d}',
        0xe9 => '\u{23a1}',
        0xea => '\u{23a2}',
        0xeb => '\u{23a3}',
        0xec => '\u{23a7}',
        0xed => '\u{23a8}',
        0xee => '\u{23a9}',
        0xef => '\u{23aa}',
        0xf0 => '\u{20ac}',
        0xf1 => '\u{232a}',
        0xf2 => '\u{222b}',
        0xf3 => '\u{2320}',
        0xf4 => '\u{23ae}',
        0xf5 => '\u{2321}',
        0xf6 => '\u{239e}',
        0xf7 => '\u{239f}',
        0xf8 => '\u{23a0}',
        0xf9 => '\u{23a4}',
        0xfa => '\u{23a5}',
        0xfb => '\u{23a6}',
        0xfc => '\u{23ab}',
        0xfd => '\u{23ac}',
        0xfe => '\u{23ad}',
        _ => char::from(byte),
    }
}

fn legacy_code_page(code_page: u16) -> Option<&'static [u32; 128]> {
    match code_page {
        437 => Some(&CP437),
        850 => Some(&CP850),
        10000 => Some(&MAC_ROMAN),
        _ => None,
    }
}

fn list_format(value: i32) -> Option<ListNumberFormat> {
    match value {
        0 => Some(ListNumberFormat::Decimal),
        1 => Some(ListNumberFormat::UpperRoman),
        2 => Some(ListNumberFormat::LowerRoman),
        3 => Some(ListNumberFormat::UpperLetter),
        4 => Some(ListNumberFormat::LowerLetter),
        5 => Some(ListNumberFormat::Ordinal),
        23 => Some(ListNumberFormat::Bullet),
        _ => None,
    }
}

fn is_supported_destination(name: &str) -> bool {
    matches!(
        name,
        "fonttbl"
            | "colortbl"
            | "listtable"
            | "listoverridetable"
            | "listtext"
            | "pntext"
            | "pict"
            | "shppict"
            | "nonshppict"
            | "upr"
            | "ud"
    )
}

fn is_unsupported_destination(name: &str) -> bool {
    matches!(
        name,
        "stylesheet"
            | "info"
            | "generator"
            | "filetbl"
            | "revtbl"
            | "rsidtbl"
            | "themedata"
            | "colorschememapping"
            | "datastore"
            | "xmlnstbl"
            | "header"
            | "headerl"
            | "headerr"
            | "headerf"
            | "footer"
            | "footerl"
            | "footerr"
            | "footerf"
            | "footnote"
            | "annotation"
            | "object"
            | "field"
    )
}

fn destination_name(destination: &Destination) -> String {
    match destination {
        Destination::Body => "body".to_owned(),
        Destination::FontTable => "fonttbl".to_owned(),
        Destination::ColorTable => "colortbl".to_owned(),
        Destination::ListTable => "listtable".to_owned(),
        Destination::ListOverrideTable => "listoverridetable".to_owned(),
        Destination::ListText => "listtext".to_owned(),
        Destination::Picture => "pict".to_owned(),
        Destination::Container => "container".to_owned(),
        Destination::UnicodeAlternatives => "upr".to_owned(),
        Destination::Skip(name) => name.clone(),
    }
}

const MAC_ROMAN: [u32; 128] = [
    0x00C4, 0x00C5, 0x00C7, 0x00C9, 0x00D1, 0x00D6, 0x00DC, 0x00E1, 0x00E0, 0x00E2, 0x00E4, 0x00E3,
    0x00E5, 0x00E7, 0x00E9, 0x00E8, 0x00EA, 0x00EB, 0x00ED, 0x00EC, 0x00EE, 0x00EF, 0x00F1, 0x00F3,
    0x00F2, 0x00F4, 0x00F6, 0x00F5, 0x00FA, 0x00F9, 0x00FB, 0x00FC, 0x2020, 0x00B0, 0x00A2, 0x00A3,
    0x00A7, 0x2022, 0x00B6, 0x00DF, 0x00AE, 0x00A9, 0x2122, 0x00B4, 0x00A8, 0x2260, 0x00C6, 0x00D8,
    0x221E, 0x00B1, 0x2264, 0x2265, 0x00A5, 0x00B5, 0x2202, 0x2211, 0x220F, 0x03C0, 0x222B, 0x00AA,
    0x00BA, 0x03A9, 0x00E6, 0x00F8, 0x00BF, 0x00A1, 0x00AC, 0x221A, 0x0192, 0x2248, 0x2206, 0x00AB,
    0x00BB, 0x2026, 0x00A0, 0x00C0, 0x00C3, 0x00D5, 0x0152, 0x0153, 0x2013, 0x2014, 0x201C, 0x201D,
    0x2018, 0x2019, 0x00F7, 0x25CA, 0x00FF, 0x0178, 0x2044, 0x20AC, 0x2039, 0x203A, 0xFB01, 0xFB02,
    0x2021, 0x00B7, 0x201A, 0x201E, 0x2030, 0x00C2, 0x00CA, 0x00C1, 0x00CB, 0x00C8, 0x00CD, 0x00CE,
    0x00CF, 0x00CC, 0x00D3, 0x00D4, 0xF8FF, 0x00D2, 0x00DA, 0x00DB, 0x00D9, 0x0131, 0x02C6, 0x02DC,
    0x00AF, 0x02D8, 0x02D9, 0x02DA, 0x00B8, 0x02DD, 0x02DB, 0x02C7,
];

const CP437: [u32; 128] = [
    0x00C7, 0x00FC, 0x00E9, 0x00E2, 0x00E4, 0x00E0, 0x00E5, 0x00E7, 0x00EA, 0x00EB, 0x00E8, 0x00EF,
    0x00EE, 0x00EC, 0x00C4, 0x00C5, 0x00C9, 0x00E6, 0x00C6, 0x00F4, 0x00F6, 0x00F2, 0x00FB, 0x00F9,
    0x00FF, 0x00D6, 0x00DC, 0x00A2, 0x00A3, 0x00A5, 0x20A7, 0x0192, 0x00E1, 0x00ED, 0x00F3, 0x00FA,
    0x00F1, 0x00D1, 0x00AA, 0x00BA, 0x00BF, 0x2310, 0x00AC, 0x00BD, 0x00BC, 0x00A1, 0x00AB, 0x00BB,
    0x2591, 0x2592, 0x2593, 0x2502, 0x2524, 0x2561, 0x2562, 0x2556, 0x2555, 0x2563, 0x2551, 0x2557,
    0x255D, 0x255C, 0x255B, 0x2510, 0x2514, 0x2534, 0x252C, 0x251C, 0x2500, 0x253C, 0x255E, 0x255F,
    0x255A, 0x2554, 0x2569, 0x2566, 0x2560, 0x2550, 0x256C, 0x2567, 0x2568, 0x2564, 0x2565, 0x2559,
    0x2558, 0x2552, 0x2553, 0x256B, 0x256A, 0x2518, 0x250C, 0x2588, 0x2584, 0x258C, 0x2590, 0x2580,
    0x03B1, 0x00DF, 0x0393, 0x03C0, 0x03A3, 0x03C3, 0x00B5, 0x03C4, 0x03A6, 0x0398, 0x03A9, 0x03B4,
    0x221E, 0x03C6, 0x03B5, 0x2229, 0x2261, 0x00B1, 0x2265, 0x2264, 0x2320, 0x2321, 0x00F7, 0x2248,
    0x00B0, 0x2219, 0x00B7, 0x221A, 0x207F, 0x00B2, 0x25A0, 0x00A0,
];

const CP850: [u32; 128] = [
    0x00C7, 0x00FC, 0x00E9, 0x00E2, 0x00E4, 0x00E0, 0x00E5, 0x00E7, 0x00EA, 0x00EB, 0x00E8, 0x00EF,
    0x00EE, 0x00EC, 0x00C4, 0x00C5, 0x00C9, 0x00E6, 0x00C6, 0x00F4, 0x00F6, 0x00F2, 0x00FB, 0x00F9,
    0x00FF, 0x00D6, 0x00DC, 0x00F8, 0x00A3, 0x00D8, 0x00D7, 0x0192, 0x00E1, 0x00ED, 0x00F3, 0x00FA,
    0x00F1, 0x00D1, 0x00AA, 0x00BA, 0x00BF, 0x00AE, 0x00AC, 0x00BD, 0x00BC, 0x00A1, 0x00AB, 0x00BB,
    0x2591, 0x2592, 0x2593, 0x2502, 0x2524, 0x00C1, 0x00C2, 0x00C0, 0x00A9, 0x2563, 0x2551, 0x2557,
    0x255D, 0x00A2, 0x00A5, 0x2510, 0x2514, 0x2534, 0x252C, 0x251C, 0x2500, 0x253C, 0x00E3, 0x00C3,
    0x255A, 0x2554, 0x2569, 0x2566, 0x2560, 0x2550, 0x256C, 0x00A4, 0x00F0, 0x00D0, 0x00CA, 0x00CB,
    0x00C8, 0x0131, 0x00CD, 0x00CE, 0x00CF, 0x2518, 0x250C, 0x2588, 0x2584, 0x00A6, 0x00CC, 0x2580,
    0x00D3, 0x00DF, 0x00D4, 0x00D2, 0x00F5, 0x00D5, 0x00B5, 0x00FE, 0x00DE, 0x00DA, 0x00DB, 0x00D9,
    0x00FD, 0x00DD, 0x00AF, 0x00B4, 0x00AD, 0x00B1, 0x2017, 0x00BE, 0x00B6, 0x00A7, 0x00F7, 0x00B8,
    0x00B0, 0x00A8, 0x00B7, 0x00B9, 0x00B3, 0x00B2, 0x25A0, 0x00A0,
];

fn rtf_error(offset: usize, message: impl Into<String>) -> Error {
    Error::Rtf {
        offset,
        message: message.into(),
    }
}

#[cfg(test)]
mod tests {
    use super::{Destination, State};
    use crate::Document;

    #[test]
    fn nested_groups_do_not_clone_accumulated_rtf_buffers() {
        let parent = State {
            destination: Destination::Picture,
            text_bytes: vec![0xaa; 1024 * 1024],
            font_name: vec![0xbb; 1024 * 1024],
            ..State::default()
        };

        let child = parent.child();

        assert_eq!(parent.text_bytes.len(), 1024 * 1024);
        assert_eq!(parent.font_name.len(), 1024 * 1024);
        assert_eq!(child.text_bytes.capacity(), 0);
        assert_eq!(child.font_name.capacity(), 0);
    }

    #[test]
    fn symbol_font_charset_decodes_non_ascii_symbol_bytes() {
        let parsed =
            Document::from_rtf_bytes(br"{\rtf1\ansi{\fonttbl{\f0\fcharset2 Symbol;}}\f0\'b3}")
                .unwrap();

        assert_eq!(parsed.document.paragraph(0).unwrap().text(), "\u{2265}");
    }

    #[test]
    fn root_character_set_declarations_must_precede_header_tables() {
        assert!(
            Document::from_rtf_bytes(br"{\rtf1{\fonttbl{\f0 Arial;}}\ansicpg1251 text}").is_err()
        );
    }

    #[test]
    fn cell_boundaries_outside_table_rows_are_malformed() {
        assert!(Document::from_rtf_bytes(br"{\rtf1\cellx1440 text}").is_err());
    }

    #[test]
    fn paragraph_list_override_zero_clears_numbering() {
        let parsed = Document::from_rtf_bytes(
            br"{\rtf1{\listtable{\list{\listlevel\levelnfc23}\listid10}}{\listoverridetable{\listoverride\listid10\ls5}}\ls5 numbered\par\ls0 plain}",
        )
        .unwrap();

        assert!(parsed.document.paragraph(0).unwrap().numbering().is_some());
        assert_eq!(parsed.document.paragraph(1).unwrap().numbering(), None);
    }

    #[test]
    fn binary_controls_cannot_interrupt_unicode_surrogate_pairs() {
        assert!(Document::from_rtf_bytes(br"{\rtf1\uc0\u-10179\bin1X\u-8704}").is_err());
    }

    #[test]
    fn group_boundaries_cannot_interrupt_unicode_surrogate_pairs() {
        for input in [
            br"{\rtf1\uc0\u-10179{\u-8704}\u-8704}".as_slice(),
            br"{\rtf1\uc0\u-10179{fallback}\u-8704}",
            br"{\rtf1\uc0\u-10179{\*\producer ignored}\u-8704}",
        ] {
            assert!(Document::from_rtf_bytes(input).is_err());
        }
    }

    #[test]
    fn unsupported_page_geometry_controls_are_diagnosed() {
        let parsed = Document::from_rtf_bytes(
            br"{\rtf1\paperw12240\paperh15840\margl1440\margr1440\margt720\margb720\gutter120 text}",
        )
        .unwrap();
        let messages = parsed
            .diagnostics
            .iter()
            .map(|diagnostic| diagnostic.message.as_str())
            .collect::<Vec<_>>();

        assert_eq!(
            messages,
            [
                "RTF page geometry control \\paperw was dropped",
                "RTF page geometry control \\paperh was dropped",
                "RTF page geometry control \\margl was dropped",
                "RTF page geometry control \\margr was dropped",
                "RTF page geometry control \\margt was dropped",
                "RTF page geometry control \\margb was dropped",
                "RTF page geometry control \\gutter was dropped",
            ]
        );
    }

    #[test]
    fn unsupported_visible_document_controls_are_diagnosed() {
        let input = br"{\rtf1\widowctrl\nowidctlpar\hyphauto\sectd\lang1033\langfe1033\loch\hich\dbch\af0\rtlch\ltrch\formshade\headery720\footery720\endnhere text}";
        let parsed = Document::from_rtf_bytes(input).unwrap();
        let diagnostics = parsed
            .diagnostics
            .iter()
            .map(|diagnostic| (diagnostic.offset, diagnostic.message.as_str()))
            .collect::<Vec<_>>();

        assert_eq!(
            diagnostics,
            [
                (6, "RTF document formatting control \\widowctrl was dropped"),
                (
                    16,
                    "RTF document formatting control \\nowidctlpar was dropped",
                ),
                (28, "RTF document formatting control \\hyphauto was dropped"),
                (37, "RTF document formatting control \\sectd was dropped"),
                (43, "RTF document formatting control \\lang was dropped"),
                (52, "RTF document formatting control \\langfe was dropped"),
                (63, "RTF document formatting control \\loch was dropped"),
                (68, "RTF document formatting control \\hich was dropped"),
                (73, "RTF document formatting control \\dbch was dropped"),
                (78, "RTF document formatting control \\af was dropped"),
                (82, "RTF document formatting control \\rtlch was dropped"),
                (88, "RTF document formatting control \\ltrch was dropped"),
                (
                    94,
                    "RTF document formatting control \\formshade was dropped"
                ),
                (104, "RTF document formatting control \\headery was dropped"),
                (115, "RTF document formatting control \\footery was dropped"),
                (
                    126,
                    "RTF document formatting control \\endnhere was dropped"
                ),
            ]
        );
    }

    #[test]
    fn final_picture_only_paragraph_keeps_current_paragraph_formatting() {
        let input = br"{\rtf1\qc\li720\ri360\fi-240\sb120\sa240\sl-360\slmult0{\pict\pngblip\picwgoal100\pichgoal200 89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4890000000d49444154789c6360f8cff00000040101089d1de10000000049454e44ae426082}}";
        let parsed = Document::from_rtf_bytes(input).unwrap();
        let paragraph = parsed.document.paragraph(0).unwrap();

        assert_eq!(paragraph.alignment(), Some(crate::Alignment::Center));
        assert_eq!(paragraph.indent_left(), Some(crate::Length::twips(720)));
        assert_eq!(paragraph.indent_right(), Some(crate::Length::twips(360)));
        assert_eq!(
            paragraph.first_line_indent(),
            Some(crate::Length::twips(-240))
        );
        assert_eq!(paragraph.space_before(), Some(crate::Length::twips(120)));
        assert_eq!(paragraph.space_after(), Some(crate::Length::twips(240)));
        assert_eq!(paragraph.line_spacing(), Some(crate::Length::twips(360)));
        assert!(paragraph.run(0).unwrap().inline_image().is_some());
    }
}
