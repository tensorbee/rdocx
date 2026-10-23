//! Minimal SpreadsheetML workbook writer for editable OOXML chart data.
//!
//! This crate writes one worksheet with text and numeric columns. It is not a
//! general spreadsheet library. Chart-owned range edits preserve an imported
//! workbook and its unrelated sheets and styles. General spreadsheet reading,
//! formula evaluation and arbitrary cell styling remain outside this crate.

use std::collections::{BTreeMap, BTreeSet, HashMap};
use std::io::Cursor;

use oxml_opc::relationship::rel_types;
use oxml_opc::{OpcPackage, content_types};
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Writer, XmlVersion};
use thiserror::Error;

const SML_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MAX_COLUMNS: usize = 16_384;
const MAX_ROWS: usize = 1_048_576;
const FIRST_CUSTOM_NUMBER_FORMAT_ID: u32 = 164;

/// A typed worksheet column.
#[derive(Debug, Clone, PartialEq)]
pub enum Column {
    /// A shared-string header followed by shared-string values.
    Text {
        /// The row-one header.
        header: String,
        /// Values beginning in row two.
        values: Vec<String>,
    },
    /// A shared-string header followed by numeric values.
    Number {
        /// The row-one header.
        header: String,
        /// Values beginning in row two.
        values: Vec<f64>,
        /// An optional Excel number format for every value cell.
        number_format: Option<String>,
    },
}

impl Column {
    fn header(&self) -> &str {
        match self {
            Self::Text { header, .. } | Self::Number { header, .. } => header,
        }
    }

    fn len(&self) -> usize {
        match self {
            Self::Text { values, .. } => values.len(),
            Self::Number { values, .. } => values.len(),
        }
    }
}

/// One ordinary embedded-workbook, single-column A1 chart formula range.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartRange {
    pub sheet: String,
    pub column: usize,
    pub first_row: usize,
    pub last_row: usize,
}

impl ChartRange {
    /// Parse ordinary local ranges, refusing external, compound or computed formulas.
    pub fn from_formula(formula: &str) -> Result<Self> {
        let (sheet, address) = formula
            .rsplit_once('!')
            .ok_or_else(|| chart_range_error("formula has no worksheet owner"))?;
        if formula.len() > 4096 || sheet.contains(['[', ']']) {
            return Err(chart_range_error("external workbook formula"));
        }
        let sheet = if sheet.starts_with('\'') && sheet.ends_with('\'') && sheet.len() >= 2 {
            let body = &sheet[1..sheet.len() - 1];
            let mut name = String::new();
            let mut chars = body.chars().peekable();
            while let Some(c) = chars.next() {
                if c == '\'' && chars.next() != Some('\'') {
                    return Err(chart_range_error("malformed quoted worksheet"));
                }
                name.push(c);
            }
            name
        } else {
            if sheet.is_empty()
                || sheet
                    .chars()
                    .any(|c| c.is_whitespace() || matches!(c, '\'' | ':' | ',' | '(' | ')' | '='))
            {
                return Err(chart_range_error("unsupported worksheet formula"));
            }
            sheet.to_owned()
        };
        validate_sheet_name(&sheet)?;
        let (first, last) = address.split_once(':').unwrap_or((address, address));
        let (column, first_row) = chart_cell_address(first)?;
        let (last_column, last_row) = chart_cell_address(last)?;
        if column != last_column || first_row > last_row {
            return Err(chart_range_error("chart range must be one ordered column"));
        }
        Ok(Self {
            sheet,
            column,
            first_row,
            last_row,
        })
    }

    /// Return a quoted formula retaining this worksheet and column origin.
    pub fn formula(&self) -> Result<String> {
        validate_sheet_name(&self.sheet)?;
        if self.column >= MAX_COLUMNS
            || self.first_row == 0
            || self.last_row > MAX_ROWS
            || self.first_row > self.last_row
        {
            return Err(chart_range_error("chart range exceeds worksheet bounds"));
        }
        let prefix = format!(
            "{}!${}${}",
            formula_sheet_name(&self.sheet),
            column_name(self.column),
            self.first_row
        );
        Ok(if self.first_row == self.last_row {
            prefix
        } else {
            format!("{prefix}:${}${}", column_name(self.column), self.last_row)
        })
    }
}

/// A chart-owned replacement cell, independent of unrelated worksheet content.
#[derive(Debug, Clone, PartialEq)]
pub enum ChartCellValue {
    Text(String),
    Number(f64),
}

/// One old owned range and its validated successor, or an added or removed range.
#[derive(Debug, Clone, PartialEq)]
pub struct ChartRangeUpdate {
    pub original: Option<ChartRange>,
    pub target: Option<ChartRange>,
    pub values: Vec<ChartCellValue>,
    pub number_format: Option<String>,
}

fn chart_range_error(message: impl Into<String>) -> Error {
    Error::ChartRange(message.into())
}

fn chart_cell_address(address: &str) -> Result<(usize, usize)> {
    let address = address.strip_prefix('$').unwrap_or(address);
    let split = address
        .find(|c: char| !c.is_ascii_alphabetic())
        .ok_or_else(|| chart_range_error("cell address has no row"))?;
    if split == 0 {
        return Err(chart_range_error("cell address has no column"));
    }
    let mut column = 0usize;
    for c in address[..split].bytes() {
        column = column
            .checked_mul(26)
            .and_then(|v| v.checked_add(usize::from(c.to_ascii_uppercase() - b'A') + 1))
            .ok_or_else(|| chart_range_error("cell column overflow"))?;
    }
    let row_text = address[split..]
        .strip_prefix('$')
        .unwrap_or(&address[split..]);
    if row_text.is_empty() || !row_text.bytes().all(|c| c.is_ascii_digit()) {
        return Err(chart_range_error("unsupported cell address"));
    }
    let row = row_text
        .parse::<usize>()
        .map_err(|_| chart_range_error("cell row overflow"))?;
    if column == 0 || column > MAX_COLUMNS || row == 0 || row > MAX_ROWS {
        return Err(chart_range_error("cell address exceeds worksheet bounds"));
    }
    Ok((column - 1, row))
}

#[derive(Debug)]
struct ChartXmlNode {
    local: Vec<u8>,
    qualified: bool,
    qname: String,
    parent: Option<usize>,
    start: usize,
    open_end: usize,
    close_start: usize,
    end: usize,
    empty: bool,
    attributes: Vec<(Option<String>, Vec<u8>, String)>,
}

impl ChartXmlNode {
    fn attribute(&self, name: &[u8]) -> Option<&str> {
        self.attributes
            .iter()
            .find(|(ns, local, _)| ns.is_none() && local == name)
            .map(|(_, _, value)| value.as_str())
    }
}

fn chart_xml_nodes(xml: &[u8], expected_root: &[u8]) -> Result<Vec<ChartXmlNode>> {
    use quick_xml::name::ResolveResult;
    use quick_xml::reader::NsReader;
    let mut reader = NsReader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut nodes: Vec<ChartXmlNode> = Vec::new();
    let mut stack = Vec::new();
    loop {
        let start = reader.buffer_position() as usize;
        let event = reader.read_event_into(&mut buffer)?;
        let end = reader.buffer_position() as usize;
        let empty = matches!(&event, Event::Empty(_));
        match event {
            Event::Start(element) | Event::Empty(element) => {
                if nodes.len() >= 1_000_000 {
                    return Err(chart_range_error("worksheet XML node limit"));
                }
                let (namespace, local) = reader.resolver().resolve_element(element.name());
                let qualified = matches!(namespace, ResolveResult::Bound(value) if value.as_ref() == SML_NS.as_bytes());
                let local = local.as_ref().to_vec();
                if stack.is_empty() && (!nodes.is_empty() || !qualified || local != expected_root) {
                    return Err(chart_range_error("invalid worksheet or workbook root"));
                }
                let mut attributes = Vec::new();
                for attribute in element.attributes() {
                    let attribute = attribute.map_err(quick_xml::Error::from)?;
                    let (ns, local) = reader.resolver().resolve_attribute(attribute.key);
                    let ns = match ns {
                        ResolveResult::Unbound => None,
                        ResolveResult::Bound(value) => {
                            Some(String::from_utf8_lossy(value.as_ref()).into_owned())
                        }
                        ResolveResult::Unknown(_) => {
                            return Err(chart_range_error("unbound XML attribute prefix"));
                        }
                    };
                    attributes.push((
                        ns,
                        local.as_ref().to_vec(),
                        attribute
                            .decoded_and_normalized_value(
                                XmlVersion::Implicit1_0,
                                reader.decoder(),
                            )?
                            .into_owned(),
                    ));
                }
                let qname = String::from_utf8(element.name().as_ref().to_vec())
                    .map_err(|_| chart_range_error("invalid XML name"))?;
                let index = nodes.len();
                nodes.push(ChartXmlNode {
                    local,
                    qualified,
                    qname,
                    parent: stack.last().copied(),
                    start,
                    open_end: end,
                    close_start: end,
                    end,
                    empty,
                    attributes,
                });
                if !empty {
                    stack.push(index);
                }
            }
            Event::End(_) => {
                let index = stack
                    .pop()
                    .ok_or_else(|| chart_range_error("unbalanced worksheet XML"))?;
                nodes[index].close_start = start;
                nodes[index].end = end;
            }
            Event::DocType(_) => {
                return Err(chart_range_error(
                    "DOCTYPE is unsupported in chart workbook",
                ));
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    if nodes.is_empty() || !stack.is_empty() {
        return Err(chart_range_error("truncated worksheet XML"));
    }
    Ok(nodes)
}

/// Merge only proven chart-owned cells into an imported embedded workbook.
/// Removed series retain cells also owned by table headers. Every other part
/// and relationship stays unchanged. Strings are inline, so
/// existing shared-string indexes and opaque strings need no global rewrite.
pub fn merge_chart_ranges(xlsx: &[u8], updates: &[ChartRangeUpdate]) -> Result<Vec<u8>> {
    let mut package = OpcPackage::from_reader_with_limits(
        Cursor::new(xlsx),
        oxml_opc::PackageReadLimits {
            max_entries: 4096,
            max_part_uncompressed_bytes: 64 * 1024 * 1024,
            max_total_uncompressed_bytes: 256 * 1024 * 1024,
        },
    )?;
    let sheets = chart_sheet_parts(&package)?;
    let mut formats: BTreeMap<String, BTreeMap<(usize, usize), String>> = BTreeMap::new();
    let mut edit_cells = 0usize;
    for update in updates {
        for range in [&update.original, &update.target].into_iter().flatten() {
            range.formula()?;
            edit_cells = edit_cells
                .checked_add(range.last_row - range.first_row + 1)
                .ok_or_else(|| chart_range_error("chart-owned cell edit limit"))?;
            if edit_cells > 4_194_304 {
                return Err(chart_range_error(
                    "chart-owned cell edit limit of 4194304 cells",
                ));
            }
        }
    }
    let mut owned = BTreeSet::new();
    let mut wanted: BTreeMap<String, BTreeMap<(usize, usize), Option<ChartCellValue>>> =
        BTreeMap::new();
    for update in updates {
        if let Some(range) = &update.original {
            range.formula()?;
            for row in range.first_row..=range.last_row {
                owned.insert((range.sheet.clone(), row, range.column));
                wanted
                    .entry(range.sheet.clone())
                    .or_default()
                    .entry((row, range.column))
                    .or_insert(None);
            }
        }
        if let Some(range) = &update.target {
            range.formula()?;
            if update.values.len() != range.last_row - range.first_row + 1 {
                return Err(chart_range_error(
                    "replacement values disagree with formula range",
                ));
            }
            for (offset, value) in update.values.iter().enumerate() {
                if matches!(value, ChartCellValue::Number(v) if !v.is_finite()) {
                    return Err(chart_range_error("nonfinite chart workbook cell"));
                }
                let entry = wanted
                    .entry(range.sheet.clone())
                    .or_default()
                    .entry((range.first_row + offset, range.column))
                    .or_insert(None);
                if entry.as_ref().is_some_and(|old| old != value) {
                    return Err(chart_range_error("overlapping chart ranges disagree"));
                }
                *entry = Some(value.clone());
                if matches!(value, ChartCellValue::Number(_))
                    && let Some(format) = &update.number_format
                {
                    if format.is_empty() {
                        return Err(chart_range_error("empty chart number format"));
                    }
                    let existing = formats
                        .entry(range.sheet.clone())
                        .or_default()
                        .insert((range.first_row + offset, range.column), format.clone());
                    if existing.as_ref().is_some_and(|old| old != format) {
                        return Err(chart_range_error(
                            "overlapping chart number formats disagree",
                        ));
                    }
                }
            }
        } else if !update.values.is_empty() {
            return Err(chart_range_error(
                "removed range carries replacement values",
            ));
        }
    }
    for (sheet, mut cells) in wanted {
        let part = sheets
            .get(&sheet)
            .ok_or_else(|| chart_range_error(format!("formula worksheet {sheet:?} is missing")))?;
        retain_chart_table_headers(&package, part, &mut cells)?;
        let source = package
            .get_part(part)
            .ok_or_else(|| chart_range_error("worksheet part is missing"))?;
        let result = merge_chart_worksheet(source, &sheet, cells, &owned)?;
        package.set_part(part, result);
    }
    for (sheet, cells) in formats {
        let part = sheets
            .get(&sheet)
            .ok_or_else(|| chart_range_error("number format worksheet is missing"))?;
        apply_chart_number_formats(&mut package, part, &cells)?;
    }
    let mut output = Cursor::new(Vec::new());
    package.write_to(&mut output)?;
    Ok(output.into_inner())
}

// A removed series does not exclusively own a table's header cells. Clearing
// those cells while retaining the table metadata makes Excel recover the file.
fn retain_chart_table_headers(
    package: &OpcPackage,
    worksheet: &str,
    cells: &mut BTreeMap<(usize, usize), Option<ChartCellValue>>,
) -> Result<()> {
    if !cells.values().any(Option::is_none) {
        return Ok(());
    }
    let source = package
        .get_part(worksheet)
        .ok_or_else(|| chart_range_error("worksheet part is missing"))?;
    let nodes = chart_xml_nodes(source, b"worksheet")?;
    for node in nodes.iter().filter(|node| {
        node.qualified
            && node.local == b"tablePart"
            && node.parent.is_some_and(|parent| {
                nodes[parent].qualified
                    && nodes[parent].local == b"tableParts"
                    && nodes[parent].parent == Some(0)
            })
    }) {
        let ids = node
            .attributes
            .iter()
            .filter(|(ns, local, _)| ns.as_deref() == Some(REL_NS) && local == b"id")
            .collect::<Vec<_>>();
        let [(_, _, id)] = ids.as_slice() else {
            return Err(chart_range_error(
                "table has no unique relationship identity",
            ));
        };
        let relationship = package
            .get_part_rels(worksheet)
            .into_iter()
            .flat_map(|relationships| &relationships.items)
            .find(|relationship| relationship.id == *id)
            .ok_or_else(|| chart_range_error("missing worksheet table relationship"))?;
        if relationship.rel_type != format!("{REL_NS}/table")
            || relationship.target_mode.as_deref() == Some("External")
        {
            return Err(chart_range_error(
                "external or unsupported worksheet table owner",
            ));
        }
        let part = OpcPackage::resolve_rel_target(worksheet, &relationship.target);
        let table = package
            .get_part(&part)
            .ok_or_else(|| chart_range_error("missing worksheet table part"))?;
        let table_nodes = chart_xml_nodes(table, b"table")?;
        let table = &table_nodes[0];
        let header_rows = table
            .attribute(b"headerRowCount")
            .unwrap_or("1")
            .parse::<u32>()
            .map_err(|_| chart_range_error("invalid table header row count"))?;
        if header_rows == 0 {
            continue;
        }
        if header_rows != 1 {
            return Err(chart_range_error("unsupported table header row count"));
        }
        let reference = table
            .attribute(b"ref")
            .ok_or_else(|| chart_range_error("table has no cell range"))?;
        let (first, last) = reference.split_once(':').unwrap_or((reference, reference));
        let (first_column, first_row) = chart_cell_address(first)?;
        let (last_column, last_row) = chart_cell_address(last)?;
        if first_column > last_column || first_row > last_row {
            return Err(chart_range_error("unordered table cell range"));
        }
        cells.retain(|&(row, column), value| {
            value.is_some() || row != first_row || column < first_column || column > last_column
        });
    }
    Ok(())
}

fn chart_sheet_parts(package: &OpcPackage) -> Result<BTreeMap<String, String>> {
    let main = package
        .main_document_part()
        .ok_or_else(|| chart_range_error("missing workbook relationship"))?;
    let workbook = package
        .get_part(&main)
        .ok_or_else(|| chart_range_error("missing workbook part"))?;
    let nodes = chart_xml_nodes(workbook, b"workbook")?;
    let relationships = package
        .get_part_rels(&main)
        .ok_or_else(|| chart_range_error("missing worksheet relationships"))?;
    let mut sheets = BTreeMap::new();
    for node in &nodes {
        if !node.qualified
            || node.local != b"sheet"
            || !node.parent.is_some_and(|p| {
                nodes[p].qualified && nodes[p].local == b"sheets" && nodes[p].parent == Some(0)
            })
        {
            continue;
        }
        let name = node
            .attribute(b"name")
            .ok_or_else(|| chart_range_error("worksheet has no name"))?;
        let ids = node
            .attributes
            .iter()
            .filter(|(ns, local, _)| ns.as_deref() == Some(REL_NS) && local == b"id")
            .collect::<Vec<_>>();
        let [(_, _, id)] = ids.as_slice() else {
            return Err(chart_range_error(
                "worksheet has no unique relationship identity",
            ));
        };
        let rel = relationships
            .items
            .iter()
            .find(|r| r.id == *id)
            .ok_or_else(|| chart_range_error("missing owned worksheet relationship"))?;
        if rel.rel_type != rel_types::WORKSHEET || rel.target_mode.as_deref() == Some("External") {
            return Err(chart_range_error("external or unsupported worksheet owner"));
        }
        let target = OpcPackage::resolve_rel_target(&main, &rel.target);
        if sheets.insert(name.to_owned(), target).is_some() {
            return Err(chart_range_error("duplicate worksheet name"));
        }
    }
    Ok(sheets)
}

/// Select a wholly unused worksheet column for a newly added chart series.
pub fn chart_unused_columns(xlsx: &[u8], sheet: &str, count: usize) -> Result<Vec<usize>> {
    let package = OpcPackage::from_reader_with_limits(
        Cursor::new(xlsx),
        oxml_opc::PackageReadLimits {
            max_entries: 4096,
            max_part_uncompressed_bytes: 64 * 1024 * 1024,
            max_total_uncompressed_bytes: 256 * 1024 * 1024,
        },
    )?;
    let sheets = chart_sheet_parts(&package)?;
    let part = sheets
        .get(sheet)
        .ok_or_else(|| chart_range_error("new series worksheet is missing"))?;
    let source = package
        .get_part(part)
        .ok_or_else(|| chart_range_error("new series worksheet part is missing"))?;
    let nodes = chart_xml_nodes(source, b"worksheet")?;
    let mut columns = BTreeSet::new();
    for node in &nodes {
        if node.qualified
            && node.local == b"c"
            && node.parent.is_some_and(|p| {
                nodes[p].qualified
                    && nodes[p].local == b"row"
                    && nodes[p].parent.is_some_and(|p| {
                        nodes[p].qualified
                            && nodes[p].local == b"sheetData"
                            && nodes[p].parent == Some(0)
                    })
            })
        {
            columns.insert(
                chart_cell_address(
                    node.attribute(b"r")
                        .ok_or_else(|| chart_range_error("cell has no address"))?,
                )?
                .0,
            );
        }
    }
    let free = (0..MAX_COLUMNS)
        .filter(|column| !columns.contains(column))
        .take(count)
        .collect::<Vec<_>>();
    if free.len() != count {
        return Err(chart_range_error(
            "no free worksheet columns for new chart series",
        ));
    }
    Ok(free)
}

fn chart_xml_attribute(raw: &[u8], key: &[u8], value: &str) -> Result<Vec<u8>> {
    let mut position = 1usize;
    while position < raw.len()
        && !raw[position].is_ascii_whitespace()
        && !matches!(raw[position], b'>' | b'/')
    {
        position += 1;
    }
    while position < raw.len() {
        while position < raw.len() && raw[position].is_ascii_whitespace() {
            position += 1;
        }
        if position >= raw.len() || matches!(raw[position], b'>' | b'/') {
            break;
        }
        let begin = position;
        while position < raw.len() && !raw[position].is_ascii_whitespace() && raw[position] != b'='
        {
            position += 1;
        }
        let name = &raw[begin..position];
        while position < raw.len() && raw[position].is_ascii_whitespace() {
            position += 1;
        }
        if raw.get(position) != Some(&b'=') {
            return Err(chart_range_error("invalid source attribute"));
        }
        position += 1;
        while position < raw.len() && raw[position].is_ascii_whitespace() {
            position += 1;
        }
        let quote = *raw
            .get(position)
            .ok_or_else(|| chart_range_error("truncated source attribute"))?;
        if !matches!(quote, b'\'' | b'"') {
            return Err(chart_range_error("unquoted source attribute"));
        }
        position += 1;
        let value_start = position;
        while position < raw.len() && raw[position] != quote {
            position += 1;
        }
        if position == raw.len() {
            return Err(chart_range_error("truncated source attribute value"));
        }
        if name == key {
            let mut result = raw.to_vec();
            result.splice(
                value_start..position,
                quick_xml::escape::escape(value).as_bytes().iter().copied(),
            );
            return Ok(result);
        }
        position += 1;
    }
    let ending = if raw.ends_with(b"/>") {
        raw.len() - 2
    } else if raw.ends_with(b">") {
        raw.len() - 1
    } else {
        return Err(chart_range_error("invalid XML start span"));
    };
    let mut result = raw.to_vec();
    let addition = format!(
        " {}=\"{}\"",
        String::from_utf8_lossy(key),
        quick_xml::escape::escape(value)
    );
    result.splice(ending..ending, addition.bytes());
    Ok(result)
}

fn append_chart_xml_children(
    source: &[u8],
    node: &ChartXmlNode,
    children: &[u8],
    count: usize,
) -> Result<Vec<u8>> {
    let mut start = chart_xml_attribute(
        &source[node.start..node.open_end],
        b"count",
        &count.to_string(),
    )?;
    if node.empty {
        start.truncate(start.len() - 2);
        start.push(b'>');
    } else {
        start.extend_from_slice(&source[node.open_end..node.close_start]);
    }
    start.extend_from_slice(children);
    start.extend_from_slice(format!("</{}>", node.qname).as_bytes());
    Ok(start)
}

fn chart_style_owner(package: &mut OpcPackage) -> Result<String> {
    let main = package
        .main_document_part()
        .ok_or_else(|| chart_range_error("missing workbook main part"))?;
    let existing = package
        .get_part_rels(&main)
        .into_iter()
        .flat_map(|r| &r.items)
        .filter(|r| r.rel_type == rel_types::STYLES)
        .collect::<Vec<_>>();
    if let [relationship] = existing.as_slice() {
        if relationship.target_mode.as_deref() == Some("External") {
            return Err(chart_range_error("external workbook styles"));
        }
        let part = OpcPackage::resolve_rel_target(&main, &relationship.target);
        if package.get_part(&part).is_none() {
            return Err(chart_range_error("missing workbook style part"));
        }
        return Ok(part);
    }
    if !existing.is_empty() {
        return Err(chart_range_error("ambiguous workbook style owner"));
    }
    let directory = main.rsplit_once('/').map(|(dir, _)| dir).unwrap_or("");
    let mut offset = 0usize;
    let part = loop {
        let candidate = if offset == 0 {
            format!("{directory}/styles.xml")
        } else {
            format!("{directory}/styles{offset}.xml")
        };
        if !package.contains_part(&candidate) {
            break candidate;
        }
        offset += 1;
    };
    package
        .get_or_create_part_rels(&main)
        .add(rel_types::STYLES, &part);
    package
        .content_types
        .add_override(&part, content_types::STYLES);
    package.set_part(
        &part,
        Styles {
            formats: Vec::new(),
            indexes: HashMap::new(),
        }
        .to_xml()?,
    );
    Ok(part)
}

fn apply_chart_number_formats(
    package: &mut OpcPackage,
    worksheet: &str,
    formats: &BTreeMap<(usize, usize), String>,
) -> Result<()> {
    let style_part = chart_style_owner(package)?;
    let style_source = package.get_part(&style_part).unwrap().to_vec();
    let style_nodes = chart_xml_nodes(&style_source, b"styleSheet")?;
    let cell_xfs = style_nodes
        .iter()
        .enumerate()
        .filter(|(_, n)| n.parent == Some(0) && n.qualified && n.local == b"cellXfs")
        .collect::<Vec<_>>();
    let [(xf_owner_index, xf_owner)] = cell_xfs.as_slice() else {
        return Err(chart_range_error("styles require unique cellXfs"));
    };
    let xfs = style_nodes
        .iter()
        .filter(|n| n.parent == Some(*xf_owner_index) && n.qualified && n.local == b"xf")
        .collect::<Vec<_>>();
    if xfs.is_empty() {
        return Err(chart_range_error("styles have no base cell format"));
    }
    let num_owners = style_nodes
        .iter()
        .enumerate()
        .filter(|(_, n)| n.parent == Some(0) && n.qualified && n.local == b"numFmts")
        .collect::<Vec<_>>();
    if num_owners.len() > 1 {
        return Err(chart_range_error("ambiguous custom number formats"));
    }
    let mut codes = BTreeMap::new();
    let mut ids = BTreeSet::new();
    let mut format_count = 0usize;
    for node in &style_nodes {
        if node.qualified
            && node.local == b"numFmt"
            && num_owners
                .first()
                .is_some_and(|(owner, _)| node.parent == Some(*owner))
        {
            let id = node
                .attribute(b"numFmtId")
                .ok_or_else(|| chart_range_error("number format has no identity"))?
                .parse::<u32>()
                .map_err(|_| chart_range_error("invalid number format identity"))?;
            if !ids.insert(id) {
                return Err(chart_range_error("duplicate number format identity"));
            }
            let code = node
                .attribute(b"formatCode")
                .ok_or_else(|| chart_range_error("number format has no code"))?;
            codes.insert(code.to_owned(), id);
            format_count += 1;
        }
    }
    for (code, id) in [
        ("General", 0),
        ("0", 1),
        ("0.00", 2),
        ("0%", 9),
        ("0.00%", 10),
    ] {
        codes.entry(code.to_owned()).or_insert(id);
    }
    let prefix = style_nodes[0].qname.strip_suffix("styleSheet").unwrap();
    let source = package
        .get_part(worksheet)
        .ok_or_else(|| chart_range_error("formatted worksheet missing"))?
        .to_vec();
    let nodes = chart_xml_nodes(&source, b"worksheet")?;
    let mut new_formats = Vec::new();
    let mut new_xfs = Vec::new();
    let mut cache = BTreeMap::new();
    let mut worksheet_patches = Vec::new();
    for node in &nodes {
        if !node.qualified
            || node.local != b"c"
            || !node.parent.is_some_and(|p| {
                nodes[p].qualified
                    && nodes[p].local == b"row"
                    && nodes[p].parent.is_some_and(|p| {
                        nodes[p].qualified
                            && nodes[p].local == b"sheetData"
                            && nodes[p].parent == Some(0)
                    })
            })
        {
            continue;
        }
        let (column, row) = chart_cell_address(
            node.attribute(b"r")
                .ok_or_else(|| chart_range_error("formatted cell missing address"))?,
        )?;
        let Some(code) = formats.get(&(row, column)) else {
            continue;
        };
        let base = node
            .attribute(b"s")
            .unwrap_or("0")
            .parse::<usize>()
            .map_err(|_| chart_range_error("invalid cell style identity"))?;
        let base_xf = xfs
            .get(base)
            .ok_or_else(|| chart_range_error("cell style identity is out of range"))?;
        let num_format = if let Some(&id) = codes.get(code) {
            id
        } else {
            let mut id = ids
                .last()
                .copied()
                .unwrap_or(163)
                .max(163)
                .checked_add(1)
                .ok_or_else(|| chart_range_error("number format identity overflow"))?;
            while ids.contains(&id) {
                id = id
                    .checked_add(1)
                    .ok_or_else(|| chart_range_error("number format identity overflow"))?;
            }
            ids.insert(id);
            codes.insert(code.clone(), id);
            let tag = format!("{prefix}numFmt");
            let mut writer = Writer::new(Vec::new());
            let mut entry = BytesStart::new(&tag);
            let id_text = id.to_string();
            let encoded = encode_spreadsheet_text(code);
            entry.push_attribute(("numFmtId", id_text.as_str()));
            entry.push_attribute(("formatCode", encoded.as_str()));
            writer.write_event(Event::Empty(entry))?;
            new_formats.extend(writer.into_inner());
            format_count += 1;
            id
        };
        let key = (base, code.clone());
        let index = if base_xf.attribute(b"numFmtId").unwrap_or("0") == num_format.to_string() {
            base
        } else if let Some(&index) = cache.get(&key) {
            index
        } else {
            let index = xfs.len() + cache.len();
            let original = &style_source[base_xf.start..base_xf.open_end];
            let start = chart_xml_attribute(original, b"numFmtId", &num_format.to_string())?;
            let mut start = chart_xml_attribute(&start, b"applyNumberFormat", "1")?;
            if !base_xf.empty {
                start.extend_from_slice(&style_source[base_xf.open_end..base_xf.end]);
            }
            new_xfs.extend(start);
            cache.insert(key, index);
            index
        };
        let changed =
            chart_xml_attribute(&source[node.start..node.open_end], b"s", &index.to_string())?;
        worksheet_patches.push((node.start, node.open_end, changed));
    }
    let mut changed = source;
    for (start, end, bytes) in worksheet_patches.into_iter().rev() {
        changed.splice(start..end, bytes);
    }
    chart_xml_nodes(&changed, b"worksheet")?;
    let mut style_patches = Vec::new();
    if !new_xfs.is_empty() {
        style_patches.push((
            xf_owner.start,
            xf_owner.end,
            append_chart_xml_children(&style_source, xf_owner, &new_xfs, xfs.len() + cache.len())?,
        ));
    }
    if !new_formats.is_empty() {
        if let Some((_, owner)) = num_owners.first() {
            style_patches.push((
                owner.start,
                owner.end,
                append_chart_xml_children(&style_source, owner, &new_formats, format_count)?,
            ));
        } else {
            let mut bytes = format!("<{prefix}numFmts count=\"{format_count}\">").into_bytes();
            bytes.extend(new_formats);
            bytes.extend_from_slice(format!("</{prefix}numFmts>").as_bytes());
            style_patches.push((style_nodes[0].open_end, style_nodes[0].open_end, bytes));
        }
    }
    style_patches.sort_by_key(|(start, end, _)| (*start, *end));
    let mut changed_styles = style_source;
    for (start, end, bytes) in style_patches.into_iter().rev() {
        changed_styles.splice(start..end, bytes);
    }
    chart_xml_nodes(&changed_styles, b"styleSheet")?;
    package.set_part(worksheet, changed);
    package.set_part(&style_part, changed_styles);
    Ok(())
}

fn merge_chart_worksheet(
    source: &[u8],
    sheet: &str,
    mut wanted: BTreeMap<(usize, usize), Option<ChartCellValue>>,
    owned: &BTreeSet<(String, usize, usize)>,
) -> Result<Vec<u8>> {
    let nodes = chart_xml_nodes(source, b"worksheet")?;
    let data = nodes
        .iter()
        .enumerate()
        .filter(|(_, n)| n.qualified && n.local == b"sheetData" && n.parent == Some(0))
        .collect::<Vec<_>>();
    let [(data_index, data)] = data.as_slice() else {
        return Err(chart_range_error("worksheet requires unique sheetData"));
    };
    let prefix = data.qname.strip_suffix("sheetData").unwrap();
    let mut patches = Vec::new();
    let mut rows = BTreeMap::new();
    let mut cell_keys = BTreeSet::new();
    for (index, node) in nodes.iter().enumerate() {
        if node.qualified && node.local == b"row" && node.parent == Some(*data_index) {
            let row = node
                .attribute(b"r")
                .ok_or_else(|| chart_range_error("row has no explicit address"))?
                .parse::<usize>()
                .map_err(|_| chart_range_error("invalid worksheet row"))?;
            if row == 0 || row > MAX_ROWS || rows.insert(row, index).is_some() {
                return Err(chart_range_error("duplicate or invalid worksheet row"));
            }
        }
    }
    let row_indices = rows.values().copied().collect::<BTreeSet<_>>();
    let mut children: HashMap<usize, Vec<usize>> = HashMap::new();
    for (index, node) in nodes.iter().enumerate() {
        if let Some(parent) = node.parent {
            children.entry(parent).or_default().push(index);
        }
    }
    for (index, node) in nodes.iter().enumerate() {
        if !node.qualified
            || node.local != b"c"
            || !node.parent.is_some_and(|p| row_indices.contains(&p))
        {
            continue;
        }
        let (column, row) = chart_cell_address(
            node.attribute(b"r")
                .ok_or_else(|| chart_range_error("cell has no address"))?,
        )?;
        if rows.get(&row) != node.parent.as_ref() || !cell_keys.insert((row, column)) {
            return Err(chart_range_error("duplicate or mismatched cell address"));
        }
        if let Some(value) = wanted.remove(&(row, column)) {
            let original_owned = owned.contains(&(sheet.to_owned(), row, column));
            let child_indices = children.get(&index).map(Vec::as_slice).unwrap_or_default();
            let populated = child_indices.iter().any(|&child| {
                nodes[child].qualified
                    && matches!(nodes[child].local.as_slice(), b"v" | b"is" | b"f")
            });
            if !original_owned && populated {
                return Err(chart_range_error(format!(
                    "expanded chart range would overwrite unrelated cell {}",
                    cell_address(column, row)
                )));
            }
            patches.push((
                node.start,
                node.end,
                replace_chart_cell(source, node, child_indices, &nodes, value.as_ref())?,
            ));
        }
    }
    let mut additions: BTreeMap<usize, Vec<(usize, ChartCellValue)>> = BTreeMap::new();
    for ((row, column), value) in wanted {
        if let Some(value) = value {
            additions.entry(row).or_default().push((column, value));
        }
    }
    for (row, cells) in additions {
        if let Some(&row_index) = rows.get(&row) {
            let node = &nodes[row_index];
            // The qualified row's own prefix remains SpreadsheetML here even
            // when its declarations rebind the sheetData prefix or default.
            let row_prefix = node.qname.strip_suffix("row").unwrap();
            let mut positions: BTreeMap<usize, Vec<u8>> = BTreeMap::new();
            for (column, value) in cells {
                let children = children
                    .get(&row_index)
                    .into_iter()
                    .flatten()
                    .map(|&index| &nodes[index])
                    .filter(|child| child.qualified && child.local == b"c")
                    .collect::<Vec<_>>();
                let position = children
                    .iter()
                    .find(|child| {
                        chart_cell_address(child.attribute(b"r").unwrap())
                            .is_ok_and(|(c, _)| c > column)
                    })
                    .map(|child| child.start)
                    .unwrap_or_else(|| {
                        children
                            .last()
                            .map(|child| child.end)
                            .unwrap_or(node.open_end)
                    });
                positions
                    .entry(position)
                    .or_default()
                    .extend(new_chart_cell(row_prefix, column, row, &value)?);
            }
            if node.empty {
                let content = positions.into_values().flatten().collect::<Vec<_>>();
                let mut bytes = source[node.start..node.open_end - 2].to_vec();
                bytes.extend_from_slice(b">");
                bytes.extend(content);
                bytes.extend_from_slice(format!("</{}>", node.qname).as_bytes());
                patches.push((node.start, node.end, bytes));
            } else {
                for (position, bytes) in positions {
                    patches.push((position, position, bytes));
                }
            }
        } else {
            let position = rows
                .range(row + 1..)
                .next()
                .map(|(_, index)| nodes[*index].start)
                .unwrap_or(data.close_start);
            let mut bytes = format!("<{prefix}row r=\"{row}\">").into_bytes();
            for (column, value) in cells {
                bytes.extend(new_chart_cell(prefix, column, row, &value)?);
            }
            bytes.extend_from_slice(format!("</{prefix}row>").as_bytes());
            patches.push((position, position, bytes));
        }
    }
    if data.empty && !patches.is_empty() {
        let mut content: Vec<u8> = Vec::new();
        for (_, _, bytes) in &patches {
            content.extend(bytes);
        }
        let mut bytes = source[data.start..data.open_end - 2].to_vec();
        bytes.extend_from_slice(b">");
        bytes.extend(content);
        bytes.extend_from_slice(format!("</{}>", data.qname).as_bytes());
        patches = vec![(data.start, data.end, bytes)];
    }
    patches.sort_by_key(|(start, end, _)| (*start, *end));
    for pair in patches.windows(2) {
        if pair[0].1 > pair[1].0 {
            return Err(chart_range_error("overlapping worksheet edit spans"));
        }
    }
    let mut result = source.to_vec();
    for (start, end, bytes) in patches.into_iter().rev() {
        result.splice(start..end, bytes);
    }
    chart_xml_nodes(&result, b"worksheet")?;
    Ok(result)
}

fn new_chart_cell(
    prefix: &str,
    column: usize,
    row: usize,
    value: &ChartCellValue,
) -> Result<Vec<u8>> {
    let mut writer = Writer::new(Vec::new());
    let tag = format!("{prefix}c");
    let mut start = BytesStart::new(&tag);
    let address = cell_address(column, row);
    start.push_attribute(("r", address.as_str()));
    if matches!(value, ChartCellValue::Text(_)) {
        start.push_attribute(("t", "inlineStr"));
    }
    writer.write_event(Event::Start(start))?;
    writer.get_mut().extend(chart_cell_content(prefix, value)?);
    writer.write_event(Event::End(BytesEnd::new(&tag)))?;
    Ok(writer.into_inner())
}

fn chart_cell_content(prefix: &str, value: &ChartCellValue) -> Result<Vec<u8>> {
    let mut writer = Writer::new(Vec::new());
    match value {
        ChartCellValue::Text(text) => {
            let tag = format!("{prefix}is");
            writer.write_event(Event::Start(BytesStart::new(&tag)))?;
            let t = format!("{prefix}t");
            let mut start = BytesStart::new(&t);
            start.push_attribute(("xml:space", "preserve"));
            writer.write_event(Event::Start(start))?;
            writer.write_event(Event::Text(BytesText::new(&encode_spreadsheet_text(text))))?;
            writer.write_event(Event::End(BytesEnd::new(&t)))?;
            writer.write_event(Event::End(BytesEnd::new(&tag)))?;
        }
        ChartCellValue::Number(number) => {
            let tag = format!("{prefix}v");
            writer.write_event(Event::Start(BytesStart::new(&tag)))?;
            writer.write_event(Event::Text(BytesText::new(&number.to_string())))?;
            writer.write_event(Event::End(BytesEnd::new(&tag)))?;
        }
    }
    Ok(writer.into_inner())
}

fn replace_chart_cell(
    source: &[u8],
    node: &ChartXmlNode,
    child_indices: &[usize],
    nodes: &[ChartXmlNode],
    value: Option<&ChartCellValue>,
) -> Result<Vec<u8>> {
    let prefix = node.qname.strip_suffix('c').unwrap();
    let mut bytes = chart_cell_open(&source[node.start..node.open_end], value)?;
    let content = value
        .map(|v| chart_cell_content(prefix, v))
        .transpose()?
        .unwrap_or_default();
    let children = child_indices
        .iter()
        .map(|&index| &nodes[index])
        .filter(|child| child.qualified && matches!(child.local.as_slice(), b"v" | b"is" | b"f"))
        .collect::<Vec<_>>();
    if node.empty {
        bytes.extend(content);
    } else {
        let mut cursor = node.open_end;
        let mut inserted = false;
        for child in children {
            bytes.extend_from_slice(&source[cursor..child.start]);
            if !inserted {
                bytes.extend_from_slice(&content);
                inserted = true;
            }
            cursor = child.end;
        }
        if !inserted {
            bytes.extend(content);
        }
        bytes.extend_from_slice(&source[cursor..node.close_start]);
    }
    bytes.extend_from_slice(format!("</{}>", node.qname).as_bytes());
    Ok(bytes)
}

fn chart_cell_open(raw: &[u8], value: Option<&ChartCellValue>) -> Result<Vec<u8>> {
    let mut position = 1usize;
    while position < raw.len()
        && !raw[position].is_ascii_whitespace()
        && !matches!(raw[position], b'>' | b'/')
    {
        position += 1;
    }
    let mut remove = None;
    while position < raw.len() {
        while position < raw.len() && raw[position].is_ascii_whitespace() {
            position += 1;
        }
        if position >= raw.len() || matches!(raw[position], b'>' | b'/') {
            break;
        }
        let begin = position;
        while position < raw.len() && !raw[position].is_ascii_whitespace() && raw[position] != b'='
        {
            position += 1;
        }
        let key = &raw[begin..position];
        while position < raw.len() && raw[position].is_ascii_whitespace() {
            position += 1;
        }
        if raw.get(position) != Some(&b'=') {
            return Err(chart_range_error("invalid cell attribute source"));
        }
        position += 1;
        while position < raw.len() && raw[position].is_ascii_whitespace() {
            position += 1;
        }
        let quote = *raw
            .get(position)
            .ok_or_else(|| chart_range_error("truncated cell attribute"))?;
        if !matches!(quote, b'\'' | b'"') {
            return Err(chart_range_error("unquoted cell attribute"));
        }
        position += 1;
        while position < raw.len() && raw[position] != quote {
            position += 1;
        }
        if position == raw.len() {
            return Err(chart_range_error("truncated cell attribute value"));
        }
        position += 1;
        if key == b"t" {
            remove = Some(begin..position);
        }
    }
    let mut bytes = raw.to_vec();
    if let Some(range) = remove {
        bytes.drain(range);
    }
    if bytes.ends_with(b"/>") {
        bytes.truncate(bytes.len() - 2);
    } else if bytes.ends_with(b">") {
        bytes.pop();
    } else {
        return Err(chart_range_error("invalid cell start boundary"));
    }
    if matches!(value, Some(ChartCellValue::Text(_))) {
        bytes.extend_from_slice(b" t=\"inlineStr\"");
    }
    bytes.push(b'>');
    Ok(bytes)
}

/// A validated, single-sheet chart workbook.
#[derive(Debug, Clone, PartialEq)]
pub struct Workbook {
    sheet_name: String,
    columns: Vec<Column>,
}

/// A workbook construction or serialization error.
#[derive(Debug, Error)]
pub enum Error {
    /// Worksheet names must follow Excel's one-sheet naming rules.
    #[error("invalid worksheet name: {0}")]
    InvalidSheetName(&'static str),
    /// At least one column is required.
    #[error("a workbook must contain at least one column")]
    NoColumns,
    /// SpreadsheetML addresses stop at column XFD.
    #[error("workbook has {0} columns, exceeding the SpreadsheetML limit of 16384")]
    TooManyColumns(usize),
    /// SpreadsheetML numeric cells cannot carry NaN or infinity.
    #[error("column {column}, row {row} contains a nonfinite number")]
    NonFiniteNumber {
        /// Zero-based column index.
        column: usize,
        /// One-based worksheet row, including the header row.
        row: usize,
    },
    /// The row-one header leaves room for at most 1,048,575 data values.
    #[error("column {column} has {values} values, exceeding the SpreadsheetML limit of 1048575")]
    TooManyRows {
        /// Zero-based column index.
        column: usize,
        /// Number of data values in the column.
        values: usize,
    },
    /// Shared-string indexes and counts are SpreadsheetML unsigned integers.
    #[error("workbook has {0} shared-string references, exceeding the u32 limit")]
    TooManySharedStrings(u64),
    /// A chart formula or worksheet cell lies outside the supported owned range.
    #[error("unsupported chart workbook edit: {0}")]
    ChartRange(String),
    /// XML serialization failed.
    #[error("SpreadsheetML XML error: {0}")]
    Xml(#[from] quick_xml::Error),
    /// Writing XML to its byte buffer failed.
    #[error("SpreadsheetML I/O error: {0}")]
    Io(#[from] std::io::Error),
    /// OPC package construction failed.
    #[error("SpreadsheetML package error: {0}")]
    Opc(#[from] oxml_opc::OpcError),
}

/// Result type for workbook operations.
pub type Result<T> = std::result::Result<T, Error>;

impl Workbook {
    /// Validate and construct a single-sheet workbook.
    pub fn new(sheet_name: impl Into<String>, columns: Vec<Column>) -> Result<Self> {
        let sheet_name = sheet_name.into();
        validate_sheet_name(&sheet_name)?;
        if columns.is_empty() {
            return Err(Error::NoColumns);
        }
        if columns.len() > MAX_COLUMNS {
            return Err(Error::TooManyColumns(columns.len()));
        }
        for (column, value) in columns.iter().enumerate() {
            if value.len() >= MAX_ROWS {
                return Err(Error::TooManyRows {
                    column,
                    values: value.len(),
                });
            }
            if let Column::Number { values, .. } = value {
                for (value_index, value) in values.iter().enumerate() {
                    if !value.is_finite() {
                        return Err(Error::NonFiniteNumber {
                            column,
                            row: value_index + 2,
                        });
                    }
                }
            }
        }
        let text_value_counts = columns
            .iter()
            .filter_map(|column| match column {
                Column::Text { values, .. } => Some(values.len()),
                Column::Number { .. } => None,
            })
            .collect::<Vec<_>>();
        validate_shared_string_count(columns.len(), &text_value_counts)?;
        Ok(Self {
            sheet_name,
            columns,
        })
    }

    /// Serialize this workbook as deterministic `.xlsx` bytes.
    pub fn to_xlsx_bytes(&self) -> Result<Vec<u8>> {
        let shared_strings = SharedStrings::from_columns(&self.columns)?;
        let styles = Styles::from_columns(&self.columns);
        let mut package = OpcPackage::with_main_part("xl/workbook.xml", content_types::WORKBOOK);

        package
            .content_types
            .add_override("/xl/worksheets/sheet1.xml", content_types::WORKSHEET);
        package
            .content_types
            .add_override("/xl/sharedStrings.xml", content_types::SHARED_STRINGS);
        if !styles.formats.is_empty() {
            package
                .content_types
                .add_override("/xl/styles.xml", content_types::STYLES);
        }

        let workbook_rels = package.get_or_create_part_rels("/xl/workbook.xml");
        let worksheet_rel_id = workbook_rels.add(rel_types::WORKSHEET, "worksheets/sheet1.xml");
        workbook_rels.add(rel_types::SHARED_STRINGS, "sharedStrings.xml");
        if !styles.formats.is_empty() {
            workbook_rels.add(rel_types::STYLES, "styles.xml");
        }

        package.set_part("/xl/workbook.xml", self.workbook_xml(&worksheet_rel_id)?);
        package.set_part(
            "/xl/worksheets/sheet1.xml",
            self.worksheet_xml(&shared_strings, &styles)?,
        );
        package.set_part("/xl/sharedStrings.xml", shared_strings.to_xml()?);
        if !styles.formats.is_empty() {
            package.set_part("/xl/styles.xml", styles.to_xml()?);
        }

        let mut output = Cursor::new(Vec::new());
        package.write_to(&mut output)?;
        Ok(output.into_inner())
    }

    /// Return the formula-addressable data range for a column.
    pub fn formula_range(&self, column: usize) -> Option<String> {
        let data_len = self.columns.get(column)?.len();
        if data_len == 0 {
            return None;
        }
        let name = formula_sheet_name(&self.sheet_name);
        let address = column_name(column);
        Some(format!("{name}!${address}$2:${address}${}", data_len + 1))
    }

    fn workbook_xml(&self, worksheet_rel_id: &str) -> Result<Vec<u8>> {
        let mut writer = xml_writer()?;
        let mut workbook = BytesStart::new("workbook");
        workbook.push_attribute(("xmlns", SML_NS));
        workbook.push_attribute(("xmlns:r", REL_NS));
        writer.write_event(Event::Start(workbook))?;
        writer.write_event(Event::Start(BytesStart::new("sheets")))?;
        let mut sheet = BytesStart::new("sheet");
        let sheet_name = encode_spreadsheet_text(&self.sheet_name);
        sheet.push_attribute(("name", sheet_name.as_str()));
        sheet.push_attribute(("sheetId", "1"));
        sheet.push_attribute(("r:id", worksheet_rel_id));
        writer.write_event(Event::Empty(sheet))?;
        writer.write_event(Event::End(BytesEnd::new("sheets")))?;
        writer.write_event(Event::End(BytesEnd::new("workbook")))?;
        Ok(writer.into_inner())
    }

    fn worksheet_xml(&self, strings: &SharedStrings, styles: &Styles) -> Result<Vec<u8>> {
        let mut writer = xml_writer()?;
        let mut worksheet = BytesStart::new("worksheet");
        worksheet.push_attribute(("xmlns", SML_NS));
        writer.write_event(Event::Start(worksheet))?;
        writer.write_event(Event::Start(BytesStart::new("sheetData")))?;

        write_row_start(&mut writer, 1)?;
        for (column, value) in self.columns.iter().enumerate() {
            write_shared_string_cell(
                &mut writer,
                &cell_address(column, 1),
                strings.index(value.header()),
            )?;
        }
        writer.write_event(Event::End(BytesEnd::new("row")))?;

        let row_count = self.columns.iter().map(Column::len).max().unwrap_or(0);
        for value_index in 0..row_count {
            let row = value_index + 2;
            write_row_start(&mut writer, row)?;
            for (column_index, column) in self.columns.iter().enumerate() {
                let address = cell_address(column_index, row);
                match column {
                    Column::Text { values, .. } => {
                        if let Some(value) = values.get(value_index) {
                            write_shared_string_cell(&mut writer, &address, strings.index(value))?;
                        }
                    }
                    Column::Number {
                        values,
                        number_format,
                        ..
                    } => {
                        if let Some(value) = values.get(value_index) {
                            write_number_cell(
                                &mut writer,
                                &address,
                                *value,
                                number_format
                                    .as_deref()
                                    .and_then(|format| styles.style_index(format)),
                            )?;
                        }
                    }
                }
            }
            writer.write_event(Event::End(BytesEnd::new("row")))?;
        }

        writer.write_event(Event::End(BytesEnd::new("sheetData")))?;
        writer.write_event(Event::End(BytesEnd::new("worksheet")))?;
        Ok(writer.into_inner())
    }
}

fn validate_sheet_name(name: &str) -> Result<()> {
    if name.is_empty() {
        return Err(Error::InvalidSheetName("name is empty"));
    }
    if name.encode_utf16().count() > 31 {
        return Err(Error::InvalidSheetName("name exceeds 31 UTF-16 code units"));
    }
    if name.starts_with('\'') || name.ends_with('\'') {
        return Err(Error::InvalidSheetName(
            "name begins or ends with an apostrophe",
        ));
    }
    if name.chars().any(|character| {
        matches!(character, ':' | '\\' | '/' | '?' | '*' | '[' | ']') || character.is_control()
    }) {
        return Err(Error::InvalidSheetName(
            "name contains a forbidden character",
        ));
    }
    Ok(())
}

fn formula_sheet_name(name: &str) -> String {
    if name.chars().enumerate().all(|(index, character)| {
        character.is_ascii_alphanumeric() || character == '_' || (character == '.' && index > 0)
    }) && name
        .chars()
        .next()
        .is_some_and(|character| character.is_ascii_alphabetic() || character == '_')
    {
        name.to_string()
    } else {
        format!("'{}'", name.replace('\'', "''"))
    }
}

fn encode_spreadsheet_text(value: &str) -> String {
    let mut encoded = String::with_capacity(value.len());
    for (index, character) in value.char_indices() {
        if character == '_' && starts_spreadsheet_escape(value.as_bytes(), index) {
            encoded.push_str("_x005F_");
        } else if matches!(character as u32, 0x0000..=0x001F | 0xFFFE | 0xFFFF) {
            encoded.push_str(&format!("_x{:04X}_", character as u32));
        } else {
            encoded.push(character);
        }
    }
    encoded
}

fn starts_spreadsheet_escape(bytes: &[u8], index: usize) -> bool {
    bytes.get(index..index + 7).is_some_and(|candidate| {
        candidate[0] == b'_'
            && candidate[1].eq_ignore_ascii_case(&b'x')
            && candidate[2..6].iter().all(u8::is_ascii_hexdigit)
            && candidate[6] == b'_'
    })
}

fn validate_shared_string_count(header_count: usize, text_value_counts: &[usize]) -> Result<u32> {
    let mut count = header_count as u64;
    for value_count in text_value_counts {
        count = count.saturating_add(*value_count as u64);
        if count > u32::MAX as u64 {
            return Err(Error::TooManySharedStrings(count));
        }
    }
    u32::try_from(count).map_err(|_| Error::TooManySharedStrings(count))
}

fn column_name(mut index: usize) -> String {
    let mut reversed = Vec::new();
    loop {
        reversed.push((b'A' + (index % 26) as u8) as char);
        if index < 26 {
            break;
        }
        index = index / 26 - 1;
    }
    reversed.into_iter().rev().collect()
}

fn cell_address(column: usize, row: usize) -> String {
    format!("{}{row}", column_name(column))
}

fn xml_writer() -> Result<Writer<Vec<u8>>> {
    let mut writer = Writer::new(Vec::new());
    writer.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("UTF-8"),
        Some("yes"),
    )))?;
    Ok(writer)
}

fn write_row_start(writer: &mut Writer<Vec<u8>>, row: usize) -> Result<()> {
    let row = row.to_string();
    let mut start = BytesStart::new("row");
    start.push_attribute(("r", row.as_str()));
    writer.write_event(Event::Start(start))?;
    Ok(())
}

fn write_shared_string_cell(writer: &mut Writer<Vec<u8>>, address: &str, index: u32) -> Result<()> {
    let mut cell = BytesStart::new("c");
    cell.push_attribute(("r", address));
    cell.push_attribute(("t", "s"));
    writer.write_event(Event::Start(cell))?;
    write_text_element(writer, "v", &index.to_string())?;
    writer.write_event(Event::End(BytesEnd::new("c")))?;
    Ok(())
}

fn write_number_cell(
    writer: &mut Writer<Vec<u8>>,
    address: &str,
    value: f64,
    style_index: Option<u32>,
) -> Result<()> {
    let style_index = style_index.map(|index| index.to_string());
    let mut cell = BytesStart::new("c");
    cell.push_attribute(("r", address));
    if let Some(index) = style_index.as_deref() {
        cell.push_attribute(("s", index));
    }
    writer.write_event(Event::Start(cell))?;
    write_text_element(writer, "v", &value.to_string())?;
    writer.write_event(Event::End(BytesEnd::new("c")))?;
    Ok(())
}

fn write_text_element(writer: &mut Writer<Vec<u8>>, name: &str, value: &str) -> Result<()> {
    writer.write_event(Event::Start(BytesStart::new(name)))?;
    writer.write_event(Event::Text(BytesText::new(value)))?;
    writer.write_event(Event::End(BytesEnd::new(name)))?;
    Ok(())
}

struct SharedStrings {
    values: Vec<String>,
    indexes: HashMap<String, u32>,
    count: u32,
}

impl SharedStrings {
    fn from_columns(columns: &[Column]) -> Result<Self> {
        let text_value_counts = columns
            .iter()
            .filter_map(|column| match column {
                Column::Text { values, .. } => Some(values.len()),
                Column::Number { .. } => None,
            })
            .collect::<Vec<_>>();
        let count = validate_shared_string_count(columns.len(), &text_value_counts)?;
        let mut shared = Self {
            values: Vec::new(),
            indexes: HashMap::new(),
            count,
        };
        for column in columns {
            shared.intern(column.header())?;
        }
        let row_count = columns.iter().map(Column::len).max().unwrap_or(0);
        for row in 0..row_count {
            for column in columns {
                if let Column::Text { values, .. } = column
                    && let Some(value) = values.get(row)
                {
                    shared.intern(value)?;
                }
            }
        }
        Ok(shared)
    }

    fn intern(&mut self, value: &str) -> Result<u32> {
        if let Some(index) = self.indexes.get(value) {
            return Ok(*index);
        }
        let index = u32::try_from(self.values.len()).map_err(|_| {
            Error::TooManySharedStrings((self.values.len() as u64).saturating_add(1))
        })?;
        self.values.push(value.to_string());
        self.indexes.insert(value.to_string(), index);
        Ok(index)
    }

    fn index(&self, value: &str) -> u32 {
        self.indexes[value]
    }

    fn to_xml(&self) -> Result<Vec<u8>> {
        let mut writer = xml_writer()?;
        let count = self.count.to_string();
        let unique_count = self.values.len().to_string();
        let mut table = BytesStart::new("sst");
        table.push_attribute(("xmlns", SML_NS));
        table.push_attribute(("count", count.as_str()));
        table.push_attribute(("uniqueCount", unique_count.as_str()));
        writer.write_event(Event::Start(table))?;
        for value in &self.values {
            writer.write_event(Event::Start(BytesStart::new("si")))?;
            let mut text = BytesStart::new("t");
            if value.starts_with(char::is_whitespace) || value.ends_with(char::is_whitespace) {
                text.push_attribute(("xml:space", "preserve"));
            }
            writer.write_event(Event::Start(text))?;
            let value = encode_spreadsheet_text(value);
            writer.write_event(Event::Text(BytesText::new(&value)))?;
            writer.write_event(Event::End(BytesEnd::new("t")))?;
            writer.write_event(Event::End(BytesEnd::new("si")))?;
        }
        writer.write_event(Event::End(BytesEnd::new("sst")))?;
        Ok(writer.into_inner())
    }
}

struct Styles {
    formats: Vec<String>,
    indexes: HashMap<String, u32>,
}

impl Styles {
    fn from_columns(columns: &[Column]) -> Self {
        let mut formats = Vec::new();
        let mut indexes = HashMap::new();
        for column in columns {
            if let Column::Number {
                number_format: Some(format),
                ..
            } = column
                && !indexes.contains_key(format)
            {
                let index = u32::try_from(formats.len() + 1).expect("column limit fits u32");
                indexes.insert(format.clone(), index);
                formats.push(format.clone());
            }
        }
        Self { formats, indexes }
    }

    fn style_index(&self, format: &str) -> Option<u32> {
        self.indexes.get(format).copied()
    }

    fn to_xml(&self) -> Result<Vec<u8>> {
        let mut writer = xml_writer()?;
        let mut root = BytesStart::new("styleSheet");
        root.push_attribute(("xmlns", SML_NS));
        writer.write_event(Event::Start(root))?;

        let format_count = self.formats.len().to_string();
        let mut num_formats = BytesStart::new("numFmts");
        num_formats.push_attribute(("count", format_count.as_str()));
        writer.write_event(Event::Start(num_formats))?;
        for (offset, format) in self.formats.iter().enumerate() {
            let id = (FIRST_CUSTOM_NUMBER_FORMAT_ID + offset as u32).to_string();
            let mut entry = BytesStart::new("numFmt");
            entry.push_attribute(("numFmtId", id.as_str()));
            let format = encode_spreadsheet_text(format);
            entry.push_attribute(("formatCode", format.as_str()));
            writer.write_event(Event::Empty(entry))?;
        }
        writer.write_event(Event::End(BytesEnd::new("numFmts")))?;

        write_fonts(&mut writer)?;
        write_fills(&mut writer)?;
        write_borders(&mut writer)?;
        write_base_cell_style(&mut writer)?;

        let cell_format_count = (self.formats.len() + 1).to_string();
        let mut cell_formats = BytesStart::new("cellXfs");
        cell_formats.push_attribute(("count", cell_format_count.as_str()));
        writer.write_event(Event::Start(cell_formats))?;
        write_xf(&mut writer, "0", false)?;
        for offset in 0..self.formats.len() {
            let id = (FIRST_CUSTOM_NUMBER_FORMAT_ID + offset as u32).to_string();
            write_xf(&mut writer, &id, true)?;
        }
        writer.write_event(Event::End(BytesEnd::new("cellXfs")))?;

        writer.write_event(Event::Start(counted("cellStyles", "1")))?;
        let mut normal = BytesStart::new("cellStyle");
        normal.push_attribute(("name", "Normal"));
        normal.push_attribute(("xfId", "0"));
        normal.push_attribute(("builtinId", "0"));
        writer.write_event(Event::Empty(normal))?;
        writer.write_event(Event::End(BytesEnd::new("cellStyles")))?;
        writer.write_event(Event::Empty(counted("dxfs", "0")))?;
        let mut tables = counted("tableStyles", "0");
        tables.push_attribute(("defaultTableStyle", "TableStyleMedium2"));
        tables.push_attribute(("defaultPivotStyle", "PivotStyleLight16"));
        writer.write_event(Event::Empty(tables))?;
        writer.write_event(Event::End(BytesEnd::new("styleSheet")))?;
        Ok(writer.into_inner())
    }
}

fn counted<'a>(name: &'a str, count: &'a str) -> BytesStart<'a> {
    let mut start = BytesStart::new(name);
    start.push_attribute(("count", count));
    start
}

fn write_fonts(writer: &mut Writer<Vec<u8>>) -> Result<()> {
    writer.write_event(Event::Start(counted("fonts", "1")))?;
    writer.write_event(Event::Start(BytesStart::new("font")))?;
    for (name, key, value) in [
        ("sz", "val", "11"),
        ("color", "theme", "1"),
        ("name", "val", "Calibri"),
        ("family", "val", "2"),
        ("scheme", "val", "minor"),
    ] {
        let mut child = BytesStart::new(name);
        child.push_attribute((key, value));
        writer.write_event(Event::Empty(child))?;
    }
    writer.write_event(Event::End(BytesEnd::new("font")))?;
    writer.write_event(Event::End(BytesEnd::new("fonts")))?;
    Ok(())
}

fn write_fills(writer: &mut Writer<Vec<u8>>) -> Result<()> {
    writer.write_event(Event::Start(counted("fills", "2")))?;
    for pattern in ["none", "gray125"] {
        writer.write_event(Event::Start(BytesStart::new("fill")))?;
        let mut fill = BytesStart::new("patternFill");
        fill.push_attribute(("patternType", pattern));
        writer.write_event(Event::Empty(fill))?;
        writer.write_event(Event::End(BytesEnd::new("fill")))?;
    }
    writer.write_event(Event::End(BytesEnd::new("fills")))?;
    Ok(())
}

fn write_borders(writer: &mut Writer<Vec<u8>>) -> Result<()> {
    writer.write_event(Event::Start(counted("borders", "1")))?;
    writer.write_event(Event::Start(BytesStart::new("border")))?;
    for name in ["left", "right", "top", "bottom", "diagonal"] {
        writer.write_event(Event::Empty(BytesStart::new(name)))?;
    }
    writer.write_event(Event::End(BytesEnd::new("border")))?;
    writer.write_event(Event::End(BytesEnd::new("borders")))?;
    Ok(())
}

fn write_base_cell_style(writer: &mut Writer<Vec<u8>>) -> Result<()> {
    writer.write_event(Event::Start(counted("cellStyleXfs", "1")))?;
    write_xf(writer, "0", false)?;
    writer.write_event(Event::End(BytesEnd::new("cellStyleXfs")))?;
    Ok(())
}

fn write_xf(writer: &mut Writer<Vec<u8>>, number_format: &str, apply: bool) -> Result<()> {
    let mut format = BytesStart::new("xf");
    format.push_attribute(("numFmtId", number_format));
    format.push_attribute(("fontId", "0"));
    format.push_attribute(("fillId", "0"));
    format.push_attribute(("borderId", "0"));
    format.push_attribute(("xfId", "0"));
    if apply {
        format.push_attribute(("applyNumberFormat", "1"));
    }
    writer.write_event(Event::Empty(format))?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use std::fs;
    use std::io::Cursor;
    use std::path::{Path, PathBuf};
    use std::process::Command;

    use oxml_opc::OpcPackage;

    use super::*;

    const EXCEL_VERSION: &str = "16.104";
    const EXCEL_BUILD: &str = "16.104.25121423";
    const LIBREOFFICE_VERSION: &str =
        "LibreOffice 26.2.5.2 cd7284b4cbbfeb507e630c1aac019f4157393acb";
    const VIEWER_ARTIFACT_SHA256: &str =
        "8f8d12aa4ebe94f86c8164fd251cdb23845f985090be0fb6c77242aaa0fba329";

    fn sample_workbook() -> Workbook {
        Workbook::new(
            "Sales '24",
            vec![
                Column::Text {
                    header: "Category".into(),
                    values: vec!["North & West".into(), "North & West".into()],
                },
                Column::Number {
                    header: "Revenue".into(),
                    values: vec![12.5, 20.0],
                    number_format: Some("$#,##0.00".into()),
                },
            ],
        )
        .expect("sample workbook should be valid")
    }

    #[test]
    fn formula_ranges_quote_sheet_names_and_track_column_lengths() {
        let workbook = Workbook::new(
            "Sales '24",
            vec![
                Column::Text {
                    header: "Category".into(),
                    values: vec!["North".into(), "South".into(), "West".into()],
                },
                Column::Number {
                    header: "Revenue".into(),
                    values: vec![12.5],
                    number_format: None,
                },
                Column::Text {
                    header: "Empty".into(),
                    values: Vec::new(),
                },
            ],
        )
        .unwrap();
        assert_eq!(
            workbook.formula_range(0).as_deref(),
            Some("'Sales ''24'!$A$2:$A$4")
        );
        assert_eq!(
            workbook.formula_range(1).as_deref(),
            Some("'Sales ''24'!$B$2:$B$2")
        );
        assert_eq!(workbook.formula_range(2), None);
        assert_eq!(workbook.formula_range(3), None);
        assert_eq!(
            Workbook::new("Sheet1", vec![text_column()])
                .unwrap()
                .formula_range(0)
                .as_deref(),
            Some("Sheet1!$A$2:$A$2")
        );
        assert_eq!(column_name(16_383), "XFD");
    }

    #[test]
    fn invalid_workbook_inputs_fail_before_package_construction() {
        assert!(Workbook::new("", vec![text_column()]).is_err());
        assert!(Workbook::new("bad/name", vec![text_column()]).is_err());
        assert!(Workbook::new("'bad", vec![text_column()]).is_err());
        assert!(Workbook::new("12345678901234567890123456789012", vec![text_column()]).is_err());
        assert!(Workbook::new("Data", Vec::new()).is_err());
        assert!(Workbook::new("Data", vec![text_column(); MAX_COLUMNS + 1]).is_err());
        assert!(
            Workbook::new(
                "Data",
                vec![Column::Number {
                    header: "Value".into(),
                    values: vec![f64::NAN],
                    number_format: None,
                }]
            )
            .is_err()
        );

        let maximum = Workbook::new(
            "Data",
            vec![Column::Number {
                header: "Value".into(),
                values: vec![0.0; MAX_ROWS - 1],
                number_format: None,
            }],
        )
        .expect("maximum valid data row should be accepted");
        assert_eq!(
            maximum.formula_range(0).as_deref(),
            Some("Data!$A$2:$A$1048576")
        );
        assert!(matches!(
            Workbook::new(
                "Data",
                vec![Column::Number {
                    header: "Value".into(),
                    values: vec![0.0; MAX_ROWS],
                    number_format: None,
                }]
            ),
            Err(Error::TooManyRows {
                column: 0,
                values: MAX_ROWS
            })
        ));
    }

    #[test]
    fn spreadsheet_strings_escape_xml_and_reserved_sequences_exactly() {
        let workbook = Workbook::new(
            "Data_x000D_",
            vec![
                Column::Text {
                    header: "<&\r\u{1}\u{fffe}_x000D_".into(),
                    values: vec![" value\t\n\u{b}\u{ffff}_x000A_ ".into()],
                },
                Column::Number {
                    header: "Number".into(),
                    values: vec![1.0],
                    number_format: Some("0\t\n<&\r\u{2}\u{fffe}_x000D_".into()),
                },
            ],
        )
        .unwrap();
        let package = OpcPackage::from_reader(Cursor::new(workbook.to_xlsx_bytes().unwrap()))
            .expect("escaped workbook should reopen");
        let workbook_xml =
            std::str::from_utf8(package.get_part("/xl/workbook.xml").unwrap()).unwrap();
        assert!(workbook_xml.contains("name=\"Data_x005F_x000D_\""));
        let strings =
            std::str::from_utf8(package.get_part("/xl/sharedStrings.xml").unwrap()).unwrap();
        assert!(strings.contains("&lt;&amp;_x000D__x0001__xFFFE__x005F_x000D_"));
        assert!(strings.contains(
            "<t xml:space=\"preserve\"> value_x0009__x000A__x000B__xFFFF__x005F_x000A_ </t>"
        ));
        assert!(!strings.as_bytes().contains(&0x01));
        assert!(!strings.as_bytes().contains(&0x0b));
        let styles = std::str::from_utf8(package.get_part("/xl/styles.xml").unwrap()).unwrap();
        assert!(
            styles.contains(
                "formatCode=\"0_x0009__x000A_&lt;&amp;_x000D__x0002__xFFFE__x005F_x000D_\""
            )
        );
        assert!(!styles.as_bytes().contains(&0x02));
    }

    #[test]
    fn shared_string_count_rejects_the_first_value_past_u32() {
        assert_eq!(
            validate_shared_string_count(1, &[(u32::MAX - 1) as usize]).unwrap(),
            u32::MAX
        );
        assert!(matches!(
            validate_shared_string_count(1, &[u32::MAX as usize]),
            Err(Error::TooManySharedStrings(count)) if count == u32::MAX as u64 + 1
        ));
    }

    #[test]
    fn workbook_package_has_the_minimal_editable_part_graph() {
        let bytes = sample_workbook()
            .to_xlsx_bytes()
            .expect("workbook should serialize");
        let package = OpcPackage::from_reader(Cursor::new(bytes)).expect("package should reopen");
        assert_eq!(
            package.main_document_part().as_deref(),
            Some("/xl/workbook.xml")
        );
        for part in [
            "/xl/workbook.xml",
            "/xl/worksheets/sheet1.xml",
            "/xl/sharedStrings.xml",
            "/xl/styles.xml",
        ] {
            assert!(package.get_part(part).is_some(), "missing {part}");
            assert!(package.content_types.content_type_for(part).is_some());
        }
        let rels = package
            .get_part_rels("/xl/workbook.xml")
            .expect("workbook relationships");
        let worksheet_rel = rels
            .get_by_type(rel_types::WORKSHEET)
            .expect("worksheet relationship");
        let workbook = std::str::from_utf8(package.get_part("/xl/workbook.xml").unwrap()).unwrap();
        assert!(workbook.contains(&format!("r:id=\"{}\"", worksheet_rel.id)));
        for relationship in &rels.items {
            let target = OpcPackage::resolve_rel_target("/xl/workbook.xml", &relationship.target);
            assert!(
                package.get_part(&target).is_some(),
                "missing target {target}"
            );
        }

        let sheet =
            std::str::from_utf8(package.get_part("/xl/worksheets/sheet1.xml").unwrap()).unwrap();
        assert!(sheet.contains("r=\"B2\" s=\"1\"><v>12.5</v>"));
        let shared_strings = std::str::from_utf8(
            package
                .get_part("/xl/sharedStrings.xml")
                .expect("shared strings part"),
        )
        .unwrap();
        assert!(shared_strings.contains("North &amp; West"));

        let no_styles = Workbook::new("Data", vec![text_column()])
            .unwrap()
            .to_xlsx_bytes()
            .unwrap();
        let no_styles = OpcPackage::from_reader(Cursor::new(no_styles)).unwrap();
        assert!(no_styles.get_part("/xl/styles.xml").is_none());
        assert!(
            !no_styles
                .content_types
                .overrides
                .contains_key("/xl/styles.xml")
        );
    }

    #[test]
    fn equal_strings_share_one_stable_shared_string_index() {
        let workbook = sample_workbook();
        let first = workbook.to_xlsx_bytes().expect("first serialization");
        let second = workbook.to_xlsx_bytes().expect("second serialization");
        assert_eq!(first, second);
        let package = OpcPackage::from_reader(Cursor::new(first)).expect("package should reopen");
        let shared = std::str::from_utf8(
            package
                .get_part("/xl/sharedStrings.xml")
                .expect("shared strings part"),
        )
        .expect("shared strings should be UTF-8");
        assert!(shared.contains("count=\"4\""));
        assert!(shared.contains("uniqueCount=\"3\""));
        assert_eq!(shared.matches("North &amp; West").count(), 1);
    }

    #[test]
    #[ignore = "requires Excel and LibreOffice Calc"]
    fn generated_workbook_opens_cleanly_in_excel_and_libreoffice_calc() {
        let output = std::env::var_os("OXML_SML_VIEWER_GATE_OUTPUT")
            .map(std::path::PathBuf::from)
            .expect("set OXML_SML_VIEWER_GATE_OUTPUT to the SHA-bound .xlsx path");
        let bytes = sample_workbook()
            .to_xlsx_bytes()
            .expect("workbook should serialize");
        fs::write(&output, bytes).expect("write viewer gate workbook");
        assert_eq!(sha256(&output), VIEWER_ARTIFACT_SHA256);
        assert_excel_acceptance(&output);
        assert_libreoffice_acceptance(&output);
    }

    #[test]
    fn viewer_gate_candidate_is_bound_to_recorded_sha() {
        let output =
            std::env::temp_dir().join(format!("oxml-sml-f117-sha-{}.xlsx", std::process::id()));
        fs::write(&output, sample_workbook().to_xlsx_bytes().unwrap()).unwrap();
        assert_eq!(sha256(&output), VIEWER_ARTIFACT_SHA256);
        fs::remove_file(output).unwrap();
    }

    fn text_column() -> Column {
        Column::Text {
            header: "Name".into(),
            values: vec!["A".into()],
        }
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
            .unwrap()
            .split_whitespace()
            .next()
            .unwrap()
            .to_string()
    }

    fn assert_excel_acceptance(path: &Path) {
        let application = "/Applications/Microsoft Excel.app/Contents/Info.plist";
        assert_eq!(
            plist_value(application, "CFBundleShortVersionString"),
            EXCEL_VERSION
        );
        assert_eq!(plist_value(application, "CFBundleVersion"), EXCEL_BUILD);

        let path = path.to_string_lossy();
        let script = format!(
            "with timeout of 120 seconds\ntell application \"Microsoft Excel\"\nactivate\nset gateWorkbook to open workbook workbook file name \"{path}\"\ntry\nif (count of worksheets of gateWorkbook) is not 1 then error \"worksheet count mismatch\"\nset gateSheet to worksheet 1 of gateWorkbook\nif (name of gateSheet) is not \"Sales '24\" then error \"worksheet name mismatch\"\nif ((value of range \"A1\" of gateSheet) as text) is not \"Category\" then error \"A1 mismatch\"\nif ((value of range \"B1\" of gateSheet) as text) is not \"Revenue\" then error \"B1 mismatch\"\nif ((value of range \"A2\" of gateSheet) as text) is not \"North & West\" then error \"A2 mismatch\"\nif ((value of range \"A3\" of gateSheet) as text) is not \"North & West\" then error \"A3 mismatch\"\nif ((value of range \"B2\" of gateSheet) as real) is not 12.5 then error \"B2 mismatch\"\nif ((value of range \"B3\" of gateSheet) as real) is not 20 then error \"B3 mismatch\"\nclose gateWorkbook saving no\non error errorMessage number errorNumber\ntry\nclose gateWorkbook saving no\nend try\nerror errorMessage number errorNumber\nend try\nend tell\nend timeout\n"
        );
        let result = Command::new("osascript")
            .args(["-e", &script])
            .output()
            .expect("launch Excel acceptance script");
        assert!(
            result.status.success(),
            "Excel F-117 acceptance failed: {}",
            String::from_utf8_lossy(&result.stderr)
        );
    }

    fn assert_libreoffice_acceptance(path: &Path) {
        let executable = std::env::var_os("OXML_SML_LIBREOFFICE")
            .map(PathBuf::from)
            .unwrap_or_else(|| PathBuf::from("soffice"));
        let version = Command::new(&executable)
            .arg("--version")
            .output()
            .expect("read LibreOffice version");
        assert!(version.status.success());
        assert_eq!(
            String::from_utf8(version.stdout).unwrap().trim(),
            LIBREOFFICE_VERSION
        );

        let output_dir =
            std::env::temp_dir().join(format!("oxml-sml-f117-libreoffice-{}", std::process::id()));
        let profile_dir = std::env::temp_dir().join(format!(
            "oxml-sml-f117-libreoffice-profile-{}",
            std::process::id()
        ));
        if output_dir.exists() {
            fs::remove_dir_all(&output_dir).expect("remove stale LibreOffice output");
        }
        if profile_dir.exists() {
            fs::remove_dir_all(&profile_dir).expect("remove stale LibreOffice profile");
        }
        fs::create_dir_all(&output_dir).expect("create LibreOffice output");
        let profile = format!("-env:UserInstallation=file://{}", profile_dir.display());
        let result = Command::new(&executable)
            .args(["--headless", &profile, "--convert-to", "xlsx", "--outdir"])
            .arg(&output_dir)
            .arg(path)
            .output()
            .expect("run LibreOffice Calc acceptance");
        assert!(
            result.status.success(),
            "LibreOffice Calc F-117 import or export failed: {}",
            String::from_utf8_lossy(&result.stderr)
        );
        let reopened_path = output_dir.join(path.file_name().unwrap());
        assert!(reopened_path.is_file());
        let reopened = OpcPackage::open(&reopened_path).expect("reopen LibreOffice workbook");
        let sheet = std::str::from_utf8(
            reopened
                .get_part("/xl/worksheets/sheet1.xml")
                .expect("LibreOffice worksheet"),
        )
        .unwrap();
        assert!(sheet.contains("dimension ref=\"A1:B3\""));
        assert!(sheet.contains("r=\"B2\" s=\"1\" t=\"n\"><v>12.5</v>"));
        assert!(sheet.contains("r=\"B3\" s=\"1\" t=\"n\"><v>20</v>"));
        let strings = std::str::from_utf8(
            reopened
                .get_part("/xl/sharedStrings.xml")
                .expect("LibreOffice shared strings"),
        )
        .unwrap();
        for expected in ["Category", "Revenue", "North &amp; West"] {
            assert!(strings.contains(expected), "missing {expected}");
        }
        fs::remove_dir_all(&output_dir).expect("remove LibreOffice output");
        if profile_dir.exists() {
            fs::remove_dir_all(&profile_dir).expect("remove LibreOffice profile");
        }
    }

    fn plist_value(application: &str, key: &str) -> String {
        let output = Command::new("defaults")
            .args(["read", application, key])
            .output()
            .expect("read Excel Info.plist");
        assert!(output.status.success());
        String::from_utf8(output.stdout).unwrap().trim().to_string()
    }
    #[test]
    fn fx215_chart_range_refuses_external_and_computed_formulas() {
        let range = ChartRange::from_formula("'Sales ''24'!$BC$2:$BC$4").unwrap();
        assert_eq!(range.sheet, "Sales '24");
        assert_eq!(range.column, 54);
        assert_eq!(range.formula().unwrap(), "'Sales ''24'!$BC$2:$BC$4");
        for formula in [
            "[Book.xlsx]Sheet1!$A$1",
            "Sheet1!$A$2:$B$3",
            "Sheet1!SUM(A1)",
            "Sheet1!$A$0",
            "'broken'quote'!A1",
        ] {
            assert!(ChartRange::from_formula(formula).is_err(), "{formula}");
        }
    }

    #[test]
    fn fx215_chart_cell_merge_retains_unrelated_workbook_payload() {
        let source = sample_workbook().to_xlsx_bytes().unwrap();
        let mut package = OpcPackage::from_reader(Cursor::new(&source)).unwrap();
        package.set_part("/xl/opaque.bin", b"unrelated bytes".to_vec());
        let old_sheet = String::from_utf8(
            package
                .get_part("/xl/worksheets/sheet1.xml")
                .unwrap()
                .to_vec(),
        )
        .unwrap();
        let old_sheet = old_sheet
            .replace(
                "</sheetData>",
                "<row r=\"9\"><c r=\"D9\" s=\"0\"><v>771</v></c></row></sheetData>",
            )
            .replace(
                "<c r=\"B2\"",
                "<c xmlns:f=\"urn:foreign\" f:keep=\"yes\" r=\"B2\"",
            )
            .replace(
                "<v>12.5</v>",
                "<v>12.5</v><f:opaque xmlns:f=\"urn:foreign\" data=\"keep\"/>",
            );
        package.set_part("/xl/worksheets/sheet1.xml", old_sheet.as_bytes().to_vec());
        let before = package.clone();
        let mut zip = Cursor::new(Vec::new());
        package.write_to(&mut zip).unwrap();
        let original = ChartRange::from_formula("'Sales ''24'!$B$2:$B$3").unwrap();
        let mut target = original.clone();
        target.last_row = 4;
        let updated = merge_chart_ranges(
            zip.get_ref(),
            &[ChartRangeUpdate {
                number_format: None,
                original: Some(original),
                target: Some(target),
                values: vec![
                    ChartCellValue::Number(8.0),
                    ChartCellValue::Number(9.0),
                    ChartCellValue::Number(10.0),
                ],
            }],
        )
        .unwrap();
        let after = OpcPackage::from_reader(Cursor::new(updated)).unwrap();
        for (name, bytes) in &before.parts {
            if name != "/xl/worksheets/sheet1.xml" {
                assert_eq!(after.get_part(name).unwrap(), bytes.as_slice());
            }
        }
        assert_eq!(
            before.package_rels.to_xml().unwrap(),
            after.package_rels.to_xml().unwrap()
        );
        assert_eq!(
            before
                .get_part_rels("/xl/workbook.xml")
                .unwrap()
                .to_xml()
                .unwrap(),
            after
                .get_part_rels("/xl/workbook.xml")
                .unwrap()
                .to_xml()
                .unwrap()
        );
        let sheet = String::from_utf8(
            after
                .get_part("/xl/worksheets/sheet1.xml")
                .unwrap()
                .to_vec(),
        )
        .unwrap();
        assert!(sheet.contains("<v>8</v>"));
        assert!(sheet.contains("<v>9</v>"));
        assert!(sheet.contains("<v>10</v>"));
        assert!(sheet.contains("<f:opaque xmlns:f=\"urn:foreign\" data=\"keep\"/>"));
        assert!(sheet.contains("<row r=\"9\"><c r=\"D9\" s=\"0\"><v>771</v></c></row>"));
    }

    #[test]
    fn fx215_chart_added_series_uses_destination_row_namespace() {
        for (label, sheet_prefix, row_prefix, declarations) in [
            ("rebound ancestor", "s:", "t:", "xmlns:s=\"urn:foreign\""),
            (
                "default row rebinds ancestor",
                "s:",
                "",
                "xmlns:s=\"urn:foreign\"",
            ),
            (
                "prefixed row rebinds default",
                "",
                "t:",
                "xmlns=\"urn:foreign\"",
            ),
            ("ordinary alias", "s:", "t:", ""),
            ("ordinary default", "", "", ""),
        ] {
            let untouched = format!(
                "<{row_prefix}row {declarations} r=\"9\"><{row_prefix}c r=\"D9\" s=\"1\"><{row_prefix}f>SUM(B2:B3)</{row_prefix}f><{row_prefix}v>32.5</{row_prefix}v></{row_prefix}c></{row_prefix}row>"
            );
            let retained_cell = format!(
                "<{row_prefix}c r=\"B2\" s=\"1\"><{row_prefix}f>6.25*2</{row_prefix}f><{row_prefix}v>12.5</{row_prefix}v></{row_prefix}c>"
            );
            let foreign_cell = "<foreign:c xmlns:foreign=\"urn:foreign\" r=\"C2\"><foreign:v>999</foreign:v></foreign:c>";
            let worksheet = format!(
                "<s:worksheet xmlns=\"{SML_NS}\" xmlns:s=\"{SML_NS}\" xmlns:t=\"{SML_NS}\"><{sheet_prefix}sheetData><{row_prefix}row {declarations} r=\"1\"/><{row_prefix}row {declarations} r=\"2\">{retained_cell}{foreign_cell}</{row_prefix}row><{row_prefix}row {declarations} r=\"3\"/>{untouched}</{sheet_prefix}sheetData></s:worksheet>"
            );
            let mut package =
                OpcPackage::from_reader(Cursor::new(sample_workbook().to_xlsx_bytes().unwrap()))
                    .unwrap();
            package.set_part("/xl/worksheets/sheet1.xml", worksheet.into_bytes());
            package.set_part("/xl/opaque.bin", b"unrelated bytes".to_vec());
            let before = package.clone();
            let mut original = Cursor::new(Vec::new());
            package.write_to(&mut original).unwrap();
            let updates = [
                ChartRangeUpdate {
                    original: None,
                    target: Some(ChartRange::from_formula("'Sales ''24'!$C$1").unwrap()),
                    values: vec![ChartCellValue::Text("Added & series".into())],
                    number_format: None,
                },
                ChartRangeUpdate {
                    original: None,
                    target: Some(ChartRange::from_formula("'Sales ''24'!$C$2:$C$3").unwrap()),
                    values: vec![ChartCellValue::Number(8.0), ChartCellValue::Number(9.0)],
                    number_format: None,
                },
            ];
            let saved = merge_chart_ranges(original.get_ref(), &updates).unwrap();
            let reopened = OpcPackage::from_reader(Cursor::new(saved)).unwrap();
            let sheet = reopened.get_part("/xl/worksheets/sheet1.xml").unwrap();
            let nodes = chart_xml_nodes(sheet, b"worksheet").unwrap();
            // Resolve every formula target to actual SpreadsheetML cells and
            // typed payloads. Foreign cells with the same address do not count.
            for update in &updates {
                let range = update.target.as_ref().unwrap();
                for (offset, expected) in update.values.iter().enumerate() {
                    let address = cell_address(range.column, range.first_row + offset);
                    let cells = nodes
                        .iter()
                        .enumerate()
                        .filter(|(_, n)| {
                            n.qualified && n.local == b"c" && n.attribute(b"r") == Some(&address)
                        })
                        .collect::<Vec<_>>();
                    assert_eq!(cells.len(), 1, "{label}: {address}");
                    let (cell_index, _) = cells[0];
                    match expected {
                        ChartCellValue::Number(expected) => {
                            let values = nodes
                                .iter()
                                .filter(|n| {
                                    n.qualified && n.local == b"v" && n.parent == Some(cell_index)
                                })
                                .collect::<Vec<_>>();
                            assert_eq!(values.len(), 1, "{label}: {address}");
                            let value = values[0];
                            let actual =
                                std::str::from_utf8(&sheet[value.open_end..value.close_start])
                                    .unwrap()
                                    .parse::<f64>()
                                    .unwrap();
                            assert_eq!(actual, *expected, "{label}: {address}");
                        }
                        ChartCellValue::Text(expected) => {
                            let inline = nodes
                                .iter()
                                .enumerate()
                                .find(|(_, n)| {
                                    n.qualified && n.local == b"is" && n.parent == Some(cell_index)
                                })
                                .expect("qualified inline string")
                                .0;
                            let text = nodes
                                .iter()
                                .find(|n| {
                                    n.qualified && n.local == b"t" && n.parent == Some(inline)
                                })
                                .expect("qualified text");
                            let raw = std::str::from_utf8(&sheet[text.open_end..text.close_start])
                                .unwrap();
                            assert_eq!(
                                quick_xml::escape::unescape(raw).unwrap(),
                                expected.as_str()
                            );
                        }
                    }
                }
            }
            let xml = std::str::from_utf8(sheet).unwrap();
            assert!(xml.contains(&untouched), "{label}");
            assert!(xml.contains(&retained_cell), "{label}");
            assert!(xml.contains(foreign_cell), "{label}");
            for (part, bytes) in &before.parts {
                if part != "/xl/worksheets/sheet1.xml" {
                    assert_eq!(
                        reopened.get_part(part).unwrap(),
                        bytes.as_slice(),
                        "{label}: {part}"
                    );
                }
            }
            assert_eq!(
                before.package_rels.to_xml().unwrap(),
                reopened.package_rels.to_xml().unwrap()
            );
            assert_eq!(
                before
                    .get_part_rels("/xl/workbook.xml")
                    .unwrap()
                    .to_xml()
                    .unwrap(),
                reopened
                    .get_part_rels("/xl/workbook.xml")
                    .unwrap()
                    .to_xml()
                    .unwrap()
            );
        }
    }

    #[test]
    fn fx215_removed_series_preserves_table_owned_header_without_retaining_values() {
        for (prefix, header_rows, keeps_header) in [
            ("", "", true),
            ("", "1", true),
            ("s:", "1", true),
            ("s:", "0", false),
        ] {
            let mut package =
                OpcPackage::from_reader(Cursor::new(sample_workbook().to_xlsx_bytes().unwrap()))
                    .unwrap();
            let header = format!(
                "<{prefix}c xmlns:f=\"urn:foreign\" f:keep=\"yes\" r=\"B1\" t=\"s\"><{prefix}v>0</{prefix}v></{prefix}c>"
            );
            let unrelated = format!(
                "<{prefix}c r=\"C1\" t=\"inlineStr\"><{prefix}is><{prefix}t>Unrelated header</{prefix}t></{prefix}is></{prefix}c>"
            );
            let worksheet = format!(
                "<{prefix}worksheet xmlns=\"{SML_NS}\" xmlns:s=\"{SML_NS}\" xmlns:rel=\"{REL_NS}\"><{prefix}sheetData><{prefix}row r=\"1\"><{prefix}c r=\"A1\" t=\"s\"><{prefix}v>0</{prefix}v></{prefix}c>{header}{unrelated}</{prefix}row><{prefix}row r=\"2\"><{prefix}c r=\"B2\"><{prefix}v>12.5</{prefix}v></{prefix}c></{prefix}row><{prefix}row r=\"3\"><{prefix}c r=\"B3\"><{prefix}v>20</{prefix}v></{prefix}c></{prefix}row></{prefix}sheetData><{prefix}tableParts count=\"1\"><{prefix}tablePart rel:id=\"ownedTable\"/><f:tablePart xmlns:f=\"urn:foreign\" rel:id=\"notOwned\"/></{prefix}tableParts></{prefix}worksheet>"
            );
            let header_attribute = if header_rows.is_empty() {
                String::new()
            } else {
                format!(" headerRowCount=\"{header_rows}\"")
            };
            let table = format!(
                "<t:table xmlns:t=\"{SML_NS}\" xmlns:f=\"urn:foreign\" f:keep=\"table\" id=\"1\" name=\"Table1\" displayName=\"Table1\" ref=\"B1:C3\"{header_attribute}><t:tableColumns count=\"2\"><t:tableColumn id=\"1\" name=\"Series\"><f:opaque/></t:tableColumn><t:tableColumn id=\"2\" name=\"Unrelated header\"/></t:tableColumns></t:table>"
            );
            package.set_part("/xl/worksheets/sheet1.xml", worksheet.into_bytes());
            package.set_part("/xl/tables/table1.xml", table.into_bytes());
            package
                .get_or_create_part_rels("/xl/worksheets/sheet1.xml")
                .add_with_id(
                    "ownedTable",
                    &format!("{REL_NS}/table"),
                    "../tables/table1.xml",
                );
            let before = package.clone();
            let mut source = Cursor::new(Vec::new());
            package.write_to(&mut source).unwrap();
            let updates = ["$A$1", "$B$1", "$B$2:$B$3"].map(|address| ChartRangeUpdate {
                original: Some(
                    ChartRange::from_formula(&format!("'Sales ''24'!{address}")).unwrap(),
                ),
                target: None,
                values: Vec::new(),
                number_format: None,
            });
            let result = merge_chart_ranges(source.get_ref(), &updates).unwrap();
            let after = OpcPackage::from_reader(Cursor::new(result)).unwrap();
            let sheet = after.get_part("/xl/worksheets/sheet1.xml").unwrap();
            let nodes = chart_xml_nodes(sheet, b"worksheet").unwrap();
            for (address, expected_populated) in [
                ("A1", false),
                ("B1", keeps_header),
                ("B2", false),
                ("B3", false),
            ] {
                let (index, _) = nodes
                    .iter()
                    .enumerate()
                    .find(|(_, node)| {
                        node.qualified
                            && node.local == b"c"
                            && node.attribute(b"r") == Some(address)
                    })
                    .unwrap();
                assert_eq!(
                    nodes.iter().any(|node| node.parent == Some(index)
                        && node.qualified
                        && node.local == b"v"),
                    expected_populated,
                    "{prefix} {header_rows} {address}"
                );
            }
            if keeps_header {
                assert!(
                    sheet
                        .windows(header.len())
                        .any(|bytes| bytes == header.as_bytes())
                );
            }
            assert!(
                sheet
                    .windows(unrelated.len())
                    .any(|bytes| bytes == unrelated.as_bytes())
            );
            for (name, bytes) in &before.parts {
                if name != "/xl/worksheets/sheet1.xml" {
                    assert_eq!(after.get_part(name).unwrap(), bytes.as_slice());
                }
            }
            assert_eq!(
                before
                    .get_part_rels("/xl/worksheets/sheet1.xml")
                    .unwrap()
                    .to_xml()
                    .unwrap(),
                after
                    .get_part_rels("/xl/worksheets/sheet1.xml")
                    .unwrap()
                    .to_xml()
                    .unwrap()
            );
        }
    }

    #[test]
    fn fx215_expanding_range_refuses_unowned_populated_cell() {
        let source = sample_workbook().to_xlsx_bytes().unwrap();
        let original = ChartRange::from_formula("'Sales ''24'!$A$2:$A$3").unwrap();
        let target = ChartRange::from_formula("'Sales ''24'!$B$2:$B$3").unwrap();
        assert!(
            merge_chart_ranges(
                &source,
                &[ChartRangeUpdate {
                    number_format: None,
                    original: Some(original),
                    target: Some(target),
                    values: vec![ChartCellValue::Number(8.0), ChartCellValue::Number(9.0)]
                }]
            )
            .is_err()
        );
        let unchanged = OpcPackage::from_reader(Cursor::new(&source)).unwrap();
        assert!(
            String::from_utf8(
                unchanged
                    .get_part("/xl/worksheets/sheet1.xml")
                    .unwrap()
                    .to_vec()
            )
            .unwrap()
            .contains("12.5")
        );
    }

    #[test]
    fn fx215_chart_number_format_adds_style_without_rewriting_original_styles() {
        let source = sample_workbook().to_xlsx_bytes().unwrap();
        let before = OpcPackage::from_reader(Cursor::new(&source)).unwrap();
        let old_styles =
            String::from_utf8(before.get_part("/xl/styles.xml").unwrap().to_vec()).unwrap();
        let range = ChartRange::from_formula("'Sales ''24'!$B$2:$B$3").unwrap();
        let updated = merge_chart_ranges(
            &source,
            &[ChartRangeUpdate {
                original: Some(range.clone()),
                target: Some(range),
                values: vec![ChartCellValue::Number(0.25), ChartCellValue::Number(0.5)],
                number_format: Some("0.0%".into()),
            }],
        )
        .unwrap();
        let after = OpcPackage::from_reader(Cursor::new(updated)).unwrap();
        let styles = String::from_utf8(after.get_part("/xl/styles.xml").unwrap().to_vec()).unwrap();
        assert!(styles.contains("formatCode=\"0.0%\""));
        assert!(styles.contains("formatCode=\"$#,##0.00\""));
        let old_nodes = chart_xml_nodes(old_styles.as_bytes(), b"styleSheet").unwrap();
        let new_nodes = chart_xml_nodes(styles.as_bytes(), b"styleSheet").unwrap();
        for node in old_nodes.iter().filter(|node| node.local == b"xf") {
            assert!(new_nodes.iter().any(|new| new.local == b"xf"
                && styles.as_bytes()[new.start..new.end]
                    == old_styles.as_bytes()[node.start..node.end]));
        }
        let sheet = String::from_utf8(
            after
                .get_part("/xl/worksheets/sheet1.xml")
                .unwrap()
                .to_vec(),
        )
        .unwrap();
        assert!(sheet.contains("<v>0.25</v>"));
        assert!(sheet.contains("s=\"2\""));
        assert_eq!(
            before.get_part("/xl/sharedStrings.xml"),
            after.get_part("/xl/sharedStrings.xml")
        );
    }
}
