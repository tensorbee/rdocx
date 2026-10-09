//! Read-only content-control handles and atomic value binding.

use std::collections::HashMap;

use oxml_opc::OpcPackage;
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use rdocx_oxml::content_control::{CT_DataBinding, CT_Sdt, SdtContent, SdtListItem, SdtType};
use rdocx_oxml::document::{BodyContent, CT_Body};
use rdocx_oxml::namespace::matches_local_name;
use rdocx_oxml::table::{CT_Row, CT_Tbl, CT_Tc, CellContent};
use rdocx_oxml::text::{CT_P, CT_R, CT_Text, RunContent};

use crate::{Document, Error, Result};

const CUSTOM_XML_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";
const CUSTOM_XML_PROPS_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps";
const CUSTOM_XML_PROPS_NS: &[u8] =
    b"http://schemas.openxmlformats.org/officeDocument/2006/customXml";

/// Read-only view of one structured document tag in document order.
#[derive(Debug, Clone, Copy)]
pub struct ContentControlRef<'a> {
    pub(crate) inner: &'a CT_Sdt,
}

impl ContentControlRef<'_> {
    pub fn tag(&self) -> Option<&str> {
        self.inner
            .properties
            .as_ref()
            .and_then(|properties| properties.tag.as_deref())
    }

    pub fn alias(&self) -> Option<&str> {
        self.inner
            .properties
            .as_ref()
            .and_then(|properties| properties.alias.as_deref())
    }

    pub fn id(&self) -> Option<i32> {
        self.inner
            .properties
            .as_ref()
            .and_then(|properties| properties.id)
    }

    pub fn control_type(&self) -> Option<SdtType> {
        self.inner
            .properties
            .as_ref()
            .and_then(|properties| properties.control_type)
    }

    pub fn text(&self) -> String {
        let mut text = String::new();
        collect_sdt_text(self.inner, &mut text);
        text
    }

    /// The `w:listItem` entries of a drop-down list or combo box control.
    pub fn list_items(&self) -> Vec<SdtListItem> {
        self.inner
            .properties
            .as_ref()
            .and_then(|properties| properties.list_items().ok())
            .unwrap_or_default()
    }

    /// Check that `value` suits the control as the value setters write it:
    /// an item of a drop-down list, or an ISO date (`YYYY-MM-DD`, with an
    /// optional time) for a date control.
    pub fn check_value(&self, value: &str) -> Result<()> {
        prepare_control_value(self.inner, value).map(|_| ())
    }
}

/// What a value setter writes into one control.
#[derive(Debug, Clone)]
struct ControlValue {
    /// The text the control shows.
    display: String,
    /// The text a data binding stores.
    bound: String,
    /// The `w:fullDate` of a date control.
    full_date: Option<String>,
    /// The `w:lastValue` of a drop-down list or combo box.
    last_value: Option<String>,
}

/// Turn a caller value into what one control shows and stores, refusing a
/// drop-down list value outside its items and a date that does not parse. A
/// combo box also accepts free text, as Word lets its user type one.
fn prepare_control_value(control: &CT_Sdt, value: &str) -> Result<ControlValue> {
    let plain = ControlValue {
        display: value.to_owned(),
        bound: value.to_owned(),
        full_date: None,
        last_value: None,
    };
    let Some(properties) = &control.properties else {
        return Ok(plain);
    };
    let name = properties
        .tag
        .as_deref()
        .or(properties.alias.as_deref())
        .unwrap_or("without a tag");
    match properties.control_type {
        Some(control_type @ (SdtType::DropDownList | SdtType::ComboBox)) => {
            let items = properties.list_items()?;
            let found = items.iter().find(|item| {
                item.display_text.as_deref() == Some(value) || item.value.as_deref() == Some(value)
            });
            match found {
                Some(item) => {
                    let display = item.display_text.clone().or_else(|| item.value.clone());
                    let stored = item.value.clone().or_else(|| item.display_text.clone());
                    Ok(ControlValue {
                        display: display.unwrap_or_default(),
                        bound: stored.clone().unwrap_or_default(),
                        full_date: None,
                        last_value: stored,
                    })
                }
                None if control_type == SdtType::ComboBox => Ok(ControlValue {
                    last_value: Some(value.to_owned()),
                    ..plain
                }),
                None => {
                    let allowed = items
                        .iter()
                        .filter_map(|item| item.display_text.clone().or_else(|| item.value.clone()))
                        .collect::<Vec<_>>();
                    Err(Error::Other(format!(
                        "content control {name:?} is a drop-down list, and {value:?} is not one of its items {allowed:?}"
                    )))
                }
            }
        }
        Some(SdtType::Date) => {
            let Some((year, month, day)) = parse_iso_date(value) else {
                return Err(Error::Other(format!(
                    "content control {name:?} is a date control, and {value:?} is not an ISO date such as 2026-01-31"
                )));
            };
            let iso = format!("{year:04}-{month:02}-{day:02}");
            let full_date = format!("{iso}T00:00:00Z");
            let format = properties.type_child_value("dateFormat")?;
            let language = properties.type_child_value("lid")?;
            let display = format
                .as_deref()
                .and_then(|format| format_date(format, language.as_deref(), year, month, day))
                .unwrap_or_else(|| iso.clone());
            let bound = match properties.type_child_value("storeMappedDataAs")?.as_deref() {
                Some("date") => iso,
                Some("dateTime") => full_date.clone(),
                _ => display.clone(),
            };
            Ok(ControlValue {
                display,
                bound,
                full_date: Some(full_date),
                last_value: None,
            })
        }
        _ => Ok(plain),
    }
}

/// Read `YYYY-MM-DD`, optionally followed by `T` and a time, as a valid
/// calendar date.
fn parse_iso_date(value: &str) -> Option<(u32, u32, u32)> {
    let value = value.trim();
    let date = value.split_once('T').map_or(value, |(date, _)| date);
    let bytes = date.as_bytes();
    if bytes.len() != 10 || bytes[4] != b'-' || bytes[7] != b'-' {
        return None;
    }
    let number = |range: std::ops::Range<usize>| -> Option<u32> {
        let digits = &date[range];
        digits
            .bytes()
            .all(|byte| byte.is_ascii_digit())
            .then(|| digits.parse().ok())?
    };
    let (year, month, day) = (number(0..4)?, number(5..7)?, number(8..10)?);
    let leap = (year % 4 == 0 && year % 100 != 0) || year % 400 == 0;
    let days = match month {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        2 if leap => 29,
        2 => 28,
        _ => return None,
    };
    (year >= 1 && (1..=days).contains(&day)).then_some((year, month, day))
}

/// Format a date with a Word date picture such as `M/d/yyyy` or
/// `dd MMMM yyyy`. Month and day names are English, so a picture that
/// names them in another language, or one with a time field, gives `None`.
fn format_date(
    picture: &str,
    language: Option<&str>,
    year: u32,
    month: u32,
    day: u32,
) -> Option<String> {
    const MONTHS: [&str; 12] = [
        "January",
        "February",
        "March",
        "April",
        "May",
        "June",
        "July",
        "August",
        "September",
        "October",
        "November",
        "December",
    ];
    const DAYS: [&str; 7] = [
        "Sunday",
        "Monday",
        "Tuesday",
        "Wednesday",
        "Thursday",
        "Friday",
        "Saturday",
    ];
    let english = language.is_none_or(|language| language.to_ascii_lowercase().starts_with("en"));
    // Sakamoto's day of the week, Sunday first.
    let weekday = {
        const OFFSETS: [u32; 12] = [0, 3, 2, 5, 0, 3, 5, 1, 4, 6, 2, 4];
        let year = if month < 3 { year - 1 } else { year };
        ((year + year / 4 - year / 100 + year / 400 + OFFSETS[month as usize - 1] + day) % 7)
            as usize
    };
    let mut output = String::new();
    let characters = picture.chars().collect::<Vec<_>>();
    let mut index = 0;
    while index < characters.len() {
        let character = characters[index];
        if character == '\'' {
            let end = characters[index + 1..].iter().position(|c| *c == '\'')?;
            output.extend(&characters[index + 1..index + 1 + end]);
            index += end + 2;
            continue;
        }
        let run = characters[index..]
            .iter()
            .take_while(|c| **c == character)
            .count();
        match (character, run) {
            ('d', 1) => output.push_str(&day.to_string()),
            ('d', 2) => output.push_str(&format!("{day:02}")),
            ('d', 3) if english => output.push_str(&DAYS[weekday][..3]),
            ('d', 4) if english => output.push_str(DAYS[weekday]),
            ('M', 1) => output.push_str(&month.to_string()),
            ('M', 2) => output.push_str(&format!("{month:02}")),
            ('M', 3) if english => output.push_str(&MONTHS[month as usize - 1][..3]),
            ('M', 4) if english => output.push_str(MONTHS[month as usize - 1]),
            ('y', 2) => output.push_str(&format!("{:02}", year % 100)),
            ('y', 4) => output.push_str(&format!("{year:04}")),
            _ if character.is_ascii_alphabetic() => return None,
            _ => output.extend(&characters[index..index + run]),
        }
        index += run;
    }
    Some(output)
}

/// Drop the `PlaceholderText` character style Word gives the placeholder of
/// an empty control, so a value does not keep its grey placeholder look.
fn clear_placeholder_style(control: &mut CT_Sdt) {
    fn clear_run(run: &mut CT_R) {
        if let Some(properties) = run.properties.as_mut()
            && properties.style_id.as_deref() == Some("PlaceholderText")
        {
            properties.style_id = None;
        }
    }
    fn clear_paragraph(paragraph: &mut CT_P) {
        paragraph.runs.iter_mut().for_each(clear_run);
    }
    fn clear_table(table: &mut CT_Tbl) {
        table.rows.iter_mut().for_each(clear_row);
    }
    fn clear_row(row: &mut CT_Row) {
        row.cells.iter_mut().for_each(clear_cell);
    }
    fn clear_cell(cell: &mut CT_Tc) {
        for content in &mut cell.content {
            match content {
                CellContent::Paragraph(paragraph) => clear_paragraph(paragraph),
                CellContent::Table(table) => clear_table(table),
                CellContent::ContentControl(_) => {}
            }
        }
    }
    for content in &mut control.content {
        match content {
            SdtContent::Paragraph(paragraph) => clear_paragraph(paragraph),
            SdtContent::Table(table) => clear_table(table),
            SdtContent::Row(row) => clear_row(row),
            SdtContent::Cell(cell) => clear_cell(cell),
            SdtContent::Run(run) => clear_run(run),
            SdtContent::ContentControl(_) | SdtContent::RawXml(_) => {}
        }
    }
}

impl Document {
    /// Return every typed content control in document order.
    pub fn content_controls(&self) -> Vec<ContentControlRef<'_>> {
        self.document
            .body
            .content_controls()
            .into_iter()
            .map(|inner| ContentControlRef { inner })
            .collect()
    }

    /// Return every content control whose tag equals `tag`.
    pub fn content_controls_by_tag(&self, tag: &str) -> Vec<ContentControlRef<'_>> {
        self.content_controls()
            .into_iter()
            .filter(|control| control.tag() == Some(tag))
            .collect()
    }

    /// Return every content control whose alias equals `alias`.
    pub fn content_controls_by_alias(&self, alias: &str) -> Vec<ContentControlRef<'_>> {
        self.content_controls()
            .into_iter()
            .filter(|control| control.alias() == Some(alias))
            .collect()
    }

    /// Replace the display and bound custom XML value of controls with `tag`.
    pub fn set_content_control_value_by_tag(&mut self, tag: &str, value: &str) -> Result<usize> {
        self.apply_content_control_values(|control| {
            (control.tag() == Some(tag)).then(|| value.to_owned())
        })
    }

    /// Replace the display and bound custom XML value of controls with `alias`.
    pub fn set_content_control_value_by_alias(
        &mut self,
        alias: &str,
        value: &str,
    ) -> Result<usize> {
        self.apply_content_control_values(|control| {
            (control.alias() == Some(alias)).then(|| value.to_owned())
        })
    }

    /// Apply values by tag first, then by alias when the tag has no map entry.
    pub fn bind_content_controls(&mut self, values: &HashMap<String, String>) -> Result<usize> {
        self.apply_content_control_values(|control| {
            control
                .tag()
                .and_then(|tag| values.get(tag))
                .or_else(|| control.alias().and_then(|alias| values.get(alias)))
                .cloned()
        })
    }

    fn apply_content_control_values(
        &mut self,
        select: impl Fn(ContentControlRef<'_>) -> Option<String>,
    ) -> Result<usize> {
        let selected = self
            .document
            .body
            .content_controls()
            .into_iter()
            .enumerate()
            .filter_map(|(index, inner)| {
                let value = select(ContentControlRef { inner })?;
                let binding = inner
                    .properties
                    .as_ref()
                    .and_then(|properties| properties.data_binding.clone());
                Some(prepare_control_value(inner, &value).map(|value| (index, value, binding)))
            })
            .collect::<Result<Vec<_>>>()?;
        if selected.is_empty() {
            return Ok(0);
        }

        let replacements = selected
            .iter()
            .map(|(index, value, _)| (*index, value.clone()))
            .collect::<HashMap<_, _>>();
        let mut staged_document = self.document.clone();
        replace_document_controls(&mut staged_document.body, &replacements)?;
        staged_document.to_xml()?;

        let mut staged_package = self.package.clone();
        let mut edits_by_part: HashMap<String, Vec<XmlEdit>> = HashMap::new();
        for (_, value, binding) in &selected {
            let Some(binding) = binding else {
                continue;
            };
            let (part_name, edit) =
                resolve_binding_edit(&staged_package, &self.doc_part_name, binding, &value.bound)?;
            edits_by_part.entry(part_name).or_default().push(edit);
        }
        for (part_name, edits) in edits_by_part {
            let original = staged_package
                .get_part(&part_name)
                .ok_or_else(|| Error::Other(format!("custom XML part {part_name} is missing")))?;
            let changed = apply_xml_edits(original, edits)?;
            staged_package.set_part(&part_name, changed);
        }

        self.document = staged_document;
        self.package = staged_package;
        self.invalidate_layout();
        Ok(selected.len())
    }
}

fn collect_sdt_text(sdt: &CT_Sdt, output: &mut String) {
    for content in &sdt.content {
        match content {
            SdtContent::Paragraph(paragraph) => collect_paragraph_text(paragraph, output),
            SdtContent::Table(table) => collect_table_text(table, output),
            SdtContent::Row(row) => collect_row_text(row, output),
            SdtContent::Cell(cell) => collect_cell_text(cell, output),
            SdtContent::Run(run) => output.push_str(&run.text()),
            SdtContent::ContentControl(_) | SdtContent::RawXml(_) => {}
        }
    }
}

fn collect_paragraph_text(paragraph: &CT_P, output: &mut String) {
    for run in &paragraph.runs {
        output.push_str(&run.text());
    }
}

fn collect_table_text(table: &CT_Tbl, output: &mut String) {
    for row in &table.rows {
        collect_row_text(row, output);
    }
}

fn collect_row_text(row: &CT_Row, output: &mut String) {
    for cell in &row.cells {
        collect_cell_text(cell, output);
    }
}

fn collect_cell_text(cell: &CT_Tc, output: &mut String) {
    for content in &cell.content {
        match content {
            CellContent::Paragraph(paragraph) => collect_paragraph_text(paragraph, output),
            CellContent::Table(table) => collect_table_text(table, output),
            CellContent::ContentControl(_) => {}
        }
    }
}

fn replace_document_controls(
    body: &mut CT_Body,
    replacements: &HashMap<usize, ControlValue>,
) -> Result<()> {
    let expected_count = body.content_controls().len();
    let mut index = 0usize;
    for content in &mut body.content {
        match content {
            BodyContent::Paragraph(paragraph) => {
                replace_paragraph_controls(paragraph, replacements, &mut index)?
            }
            BodyContent::Table(table) => replace_table_controls(table, replacements, &mut index)?,
            BodyContent::ContentControl(control) => {
                replace_one_control(control, replacements, &mut index)?
            }
            BodyContent::RawXml(_) => {}
        }
    }
    if index != expected_count || replacements.keys().any(|selected| *selected >= index) {
        return Err(Error::Other(
            "content-control traversal changed during staged mutation".to_owned(),
        ));
    }
    Ok(())
}

fn replace_one_control(
    control: &mut CT_Sdt,
    replacements: &HashMap<usize, ControlValue>,
    index: &mut usize,
) -> Result<()> {
    if let Some(value) = replacements.get(index) {
        replace_display_text(control, &value.display)?;
        clear_placeholder_style(control);
        if let Some(properties) = control.properties.as_mut() {
            properties.clear_showing_placeholder();
            if let Some(full_date) = &value.full_date {
                properties.set_type_attribute("fullDate", full_date);
            }
            if let Some(last_value) = &value.last_value {
                properties.set_type_attribute("lastValue", last_value);
            }
        }
    }
    *index = index
        .checked_add(1)
        .ok_or_else(|| Error::Other("content-control count overflow".to_owned()))?;
    replace_nested_controls(control, replacements, index)
}

fn replace_nested_controls(
    control: &mut CT_Sdt,
    replacements: &HashMap<usize, ControlValue>,
    index: &mut usize,
) -> Result<()> {
    for content in &mut control.content {
        match content {
            SdtContent::Paragraph(paragraph) => {
                replace_paragraph_controls(paragraph, replacements, index)?
            }
            SdtContent::Table(table) => replace_table_controls(table, replacements, index)?,
            SdtContent::Row(row) => replace_row_controls(row, replacements, index)?,
            SdtContent::Cell(cell) => replace_cell_controls(cell, replacements, index)?,
            SdtContent::ContentControl(nested) => replace_one_control(nested, replacements, index)?,
            SdtContent::Run(_) | SdtContent::RawXml(_) => {}
        }
    }
    Ok(())
}

fn replace_paragraph_controls(
    paragraph: &mut CT_P,
    replacements: &HashMap<usize, ControlValue>,
    index: &mut usize,
) -> Result<()> {
    for position in 0..=paragraph.runs.len() {
        for (_, _, _, control) in paragraph
            .content_controls
            .iter_mut()
            .filter(|(at, _, _, _)| *at == position)
        {
            replace_one_control(control, replacements, index)?;
        }
    }
    Ok(())
}

fn replace_table_controls(
    table: &mut CT_Tbl,
    replacements: &HashMap<usize, ControlValue>,
    index: &mut usize,
) -> Result<()> {
    for position in 0..=table.rows.len() {
        for (_, _, control) in table
            .content_controls
            .iter_mut()
            .filter(|(at, _, _)| *at == position)
        {
            replace_one_control(control, replacements, index)?;
        }
        if let Some(row) = table.rows.get_mut(position) {
            replace_row_controls(row, replacements, index)?;
        }
    }
    Ok(())
}

fn replace_row_controls(
    row: &mut CT_Row,
    replacements: &HashMap<usize, ControlValue>,
    index: &mut usize,
) -> Result<()> {
    for position in 0..=row.cells.len() {
        for (_, _, control) in row
            .content_controls
            .iter_mut()
            .filter(|(at, _, _)| *at == position)
        {
            replace_one_control(control, replacements, index)?;
        }
        if let Some(cell) = row.cells.get_mut(position) {
            replace_cell_controls(cell, replacements, index)?;
        }
    }
    Ok(())
}

fn replace_cell_controls(
    cell: &mut CT_Tc,
    replacements: &HashMap<usize, ControlValue>,
    index: &mut usize,
) -> Result<()> {
    for content in &mut cell.content {
        match content {
            CellContent::Paragraph(paragraph) => {
                replace_paragraph_controls(paragraph, replacements, index)?
            }
            CellContent::Table(table) => replace_table_controls(table, replacements, index)?,
            CellContent::ContentControl(control) => {
                replace_one_control(control, replacements, index)?
            }
        }
    }
    Ok(())
}

fn replace_display_text(control: &mut CT_Sdt, value: &str) -> Result<()> {
    let mut replacement = DisplayReplacement {
        value,
        wrote_value: false,
        saw_run: false,
    };
    replace_sdt_display(control, &mut replacement);
    if replacement.wrote_value {
        return Ok(());
    }
    if add_text_to_first_paragraph(control, value) {
        return Ok(());
    }
    Err(Error::Other(
        "content control has no typed text display location".to_owned(),
    ))
}

struct DisplayReplacement<'a> {
    value: &'a str,
    wrote_value: bool,
    saw_run: bool,
}

fn replace_sdt_display(control: &mut CT_Sdt, replacement: &mut DisplayReplacement<'_>) {
    for content in &mut control.content {
        match content {
            SdtContent::Paragraph(paragraph) => replace_paragraph_display(paragraph, replacement),
            SdtContent::Table(table) => replace_table_display(table, replacement),
            SdtContent::Row(row) => replace_row_display(row, replacement),
            SdtContent::Cell(cell) => replace_cell_display(cell, replacement),
            SdtContent::Run(run) => replace_run_display(run, replacement),
            SdtContent::ContentControl(_) | SdtContent::RawXml(_) => {}
        }
    }
}

fn replace_run_display(run: &mut CT_R, replacement: &mut DisplayReplacement<'_>) {
    let is_first_run = !replacement.saw_run;
    replacement.saw_run = true;
    let insert_replacement = is_first_run
        && !replacement.wrote_value
        && !run
            .content
            .iter()
            .any(|content| matches!(content, RunContent::Text(_)));
    let original_raw_positions = insert_replacement.then(|| run.extra_xml_positions.clone());
    let property_boundary = usize::from(run.properties.is_some());
    let removed = run
        .content
        .iter()
        .map(|content| {
            !matches!(
                content,
                RunContent::Text(_)
                    | RunContent::DeletedText(_)
                    | RunContent::CommentReference { .. }
            )
        })
        .collect::<Vec<_>>();
    run.remap_removed_content(&removed);
    run.content.retain(|content| {
        matches!(
            content,
            RunContent::Text(_) | RunContent::DeletedText(_) | RunContent::CommentReference { .. }
        )
    });
    for content in &mut run.content {
        if let RunContent::Text(text) = content {
            if is_first_run && !replacement.wrote_value {
                *text = CT_Text::new(replacement.value);
                replacement.wrote_value = true;
            } else {
                text.text.clear();
                text.preserve_space = false;
            }
        }
    }
    if is_first_run && !replacement.wrote_value {
        if let Some(original_raw_positions) = original_raw_positions {
            for (position, original) in run
                .extra_xml_positions
                .iter_mut()
                .zip(original_raw_positions)
            {
                if CT_R::raw_child_position(original) > property_boundary {
                    let shifted = CT_R::raw_child_position(*position) + 1;
                    CT_R::set_raw_child_position(position, shifted);
                }
            }
        }
        run.content
            .insert(0, RunContent::Text(CT_Text::new(replacement.value)));
        replacement.wrote_value = true;
    }
}

fn replace_paragraph_display(paragraph: &mut CT_P, replacement: &mut DisplayReplacement<'_>) {
    for run in &mut paragraph.runs {
        replace_run_display(run, replacement);
    }
}

fn replace_table_display(table: &mut CT_Tbl, replacement: &mut DisplayReplacement<'_>) {
    for row in &mut table.rows {
        replace_row_display(row, replacement);
    }
}

fn replace_row_display(row: &mut CT_Row, replacement: &mut DisplayReplacement<'_>) {
    for cell in &mut row.cells {
        replace_cell_display(cell, replacement);
    }
}

fn replace_cell_display(cell: &mut CT_Tc, replacement: &mut DisplayReplacement<'_>) {
    for content in &mut cell.content {
        match content {
            CellContent::Paragraph(paragraph) => replace_paragraph_display(paragraph, replacement),
            CellContent::Table(table) => replace_table_display(table, replacement),
            CellContent::ContentControl(_) => {}
        }
    }
}

fn add_text_to_first_paragraph(control: &mut CT_Sdt, value: &str) -> bool {
    for content in &mut control.content {
        match content {
            SdtContent::Paragraph(paragraph) => {
                paragraph.runs.insert(0, CT_R::new(value));
                return true;
            }
            SdtContent::Table(table) => {
                if add_text_to_table(table, value) {
                    return true;
                }
            }
            SdtContent::Row(row) => {
                if add_text_to_row(row, value) {
                    return true;
                }
            }
            SdtContent::Cell(cell) => {
                if add_text_to_cell(cell, value) {
                    return true;
                }
            }
            SdtContent::ContentControl(_) => {}
            _ => {}
        }
    }
    false
}

fn add_text_to_table(table: &mut CT_Tbl, value: &str) -> bool {
    for row in &mut table.rows {
        if add_text_to_row(row, value) {
            return true;
        }
    }
    false
}

fn add_text_to_row(row: &mut CT_Row, value: &str) -> bool {
    for cell in &mut row.cells {
        if add_text_to_cell(cell, value) {
            return true;
        }
    }
    false
}

fn add_text_to_cell(cell: &mut CT_Tc, value: &str) -> bool {
    for content in &mut cell.content {
        match content {
            CellContent::Paragraph(paragraph) => {
                paragraph.runs.insert(0, CT_R::new(value));
                return true;
            }
            CellContent::Table(table) => {
                if add_text_to_table(table, value) {
                    return true;
                }
            }
            CellContent::ContentControl(_) => {}
        }
    }
    false
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct XmlEdit {
    start: usize,
    end: usize,
    replacement: Vec<u8>,
}

fn resolve_binding_edit(
    package: &OpcPackage,
    document_part: &str,
    binding: &CT_DataBinding,
    value: &str,
) -> Result<(String, XmlEdit)> {
    let store_item_id = binding
        .store_item_id
        .as_deref()
        .ok_or_else(|| Error::Other("content-control binding has no store item id".to_owned()))?;
    let xpath = binding
        .xpath
        .as_deref()
        .ok_or_else(|| Error::Other("content-control binding has no XPath".to_owned()))?;
    let mappings = parse_prefix_mappings(binding.prefix_mappings.as_deref().unwrap_or(""))?;
    let steps = parse_xpath(xpath, &mappings)?;
    let part_name = resolve_custom_xml_part(package, document_part, store_item_id)?;
    let xml = package
        .get_part(&part_name)
        .ok_or_else(|| Error::Other(format!("custom XML part {part_name} is missing")))?;
    let selection = select_xml_element(xml, &steps)?;
    let replacement = if selection.empty {
        expand_empty_element(xml, &selection, value)?
    } else {
        quick_xml::escape::escape(value).as_bytes().to_vec()
    };
    Ok((
        part_name,
        XmlEdit {
            start: selection.replace_start,
            end: selection.replace_end,
            replacement,
        },
    ))
}

pub(crate) fn resolve_custom_xml_part(
    package: &OpcPackage,
    document_part: &str,
    store_item_id: &str,
) -> Result<String> {
    let matches = matching_custom_xml_parts(package, document_part, store_item_id)?;
    match matches.as_slice() {
        [part_name] => Ok(part_name.clone()),
        [] => Err(Error::Other(format!(
            "custom XML store item {store_item_id} was not found"
        ))),
        _ => Err(Error::Other(format!(
            "custom XML store item {store_item_id} is ambiguous"
        ))),
    }
}

pub(crate) fn custom_xml_store_item_count(
    package: &OpcPackage,
    document_part: &str,
    store_item_id: &str,
) -> Result<usize> {
    Ok(matching_custom_xml_parts(package, document_part, store_item_id)?.len())
}

fn matching_custom_xml_parts(
    package: &OpcPackage,
    document_part: &str,
    store_item_id: &str,
) -> Result<Vec<String>> {
    let expected = normalize_store_item_id(store_item_id);
    let mut matches = Vec::new();
    let Some(relationships) = package.get_part_rels(document_part) else {
        return Ok(matches);
    };
    for relationship in relationships
        .items
        .iter()
        .filter(|relationship| relationship.rel_type == CUSTOM_XML_REL_TYPE)
    {
        let item_part = OpcPackage::resolve_rel_target(document_part, &relationship.target);
        let Some(properties_relationship) = package
            .get_part_rels(&item_part)
            .and_then(|relationships| relationships.get_by_type(CUSTOM_XML_PROPS_REL_TYPE))
        else {
            continue;
        };
        let properties_part =
            OpcPackage::resolve_rel_target(&item_part, &properties_relationship.target);
        let Some(properties_xml) = package.get_part(&properties_part) else {
            continue;
        };
        if parse_store_item_id(properties_xml)?
            .is_some_and(|found| normalize_store_item_id(&found) == expected)
        {
            matches.push(item_part);
        }
    }
    Ok(matches)
}

pub(crate) fn normalize_store_item_id(value: &str) -> String {
    value
        .trim()
        .trim_start_matches('{')
        .trim_end_matches('}')
        .to_ascii_lowercase()
}

pub(crate) fn parse_store_item_id(xml: &[u8]) -> Result<Option<String>> {
    let mut reader = NsReader::from_reader(xml);
    let mut buffer = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("invalid custom XML properties: {error}")))?;
        let root_matches = namespace_matches(&namespace, Some(CUSTOM_XML_PROPS_NS));
        let event = event.into_owned();
        match event {
            Event::Start(element) | Event::Empty(element)
                if root_matches
                    && matches_local_name(element.name().as_ref(), b"datastoreItem") =>
            {
                for attribute in element.attributes() {
                    let attribute = attribute.map_err(|error| {
                        Error::Other(format!("invalid custom XML properties attribute: {error}"))
                    })?;
                    let (attribute_namespace, local_name) =
                        reader.resolver().resolve_attribute(attribute.key);
                    if namespace_matches(&attribute_namespace, Some(CUSTOM_XML_PROPS_NS))
                        && local_name.as_ref() == b"itemID"
                    {
                        return attribute
                            .decoded_and_normalized_value(
                                XmlVersion::Implicit1_0,
                                element.decoder(),
                            )
                            .map(|value| Some(value.into_owned()))
                            .map_err(|error| {
                                Error::Other(format!(
                                    "invalid custom XML properties item id: {error}"
                                ))
                            });
                    }
                }
                return Ok(None);
            }
            Event::Start(_) | Event::Empty(_) => return Ok(None),
            Event::Text(text) if is_xml_whitespace(text.as_ref()) => {}
            Event::Decl(_) | Event::Comment(_) | Event::PI(_) | Event::DocType(_) => {}
            Event::Eof => return Ok(None),
            _ => return Ok(None),
        }
        buffer.clear();
    }
}

fn is_xml_whitespace(bytes: &[u8]) -> bool {
    bytes.iter().all(u8::is_ascii_whitespace)
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct XPathStep {
    namespace: Option<Vec<u8>>,
    local_name: Vec<u8>,
    index: Option<usize>,
}

fn parse_prefix_mappings(value: &str) -> Result<HashMap<String, Vec<u8>>> {
    let bytes = value.as_bytes();
    let mut mappings = HashMap::new();
    let mut position = 0usize;
    while position < bytes.len() {
        skip_ascii_whitespace(bytes, &mut position);
        if position == bytes.len() {
            break;
        }
        if !bytes[position..].starts_with(b"xmlns") {
            return Err(Error::Other("unsupported prefixMappings syntax".to_owned()));
        }
        position += b"xmlns".len();
        let prefix = if bytes.get(position) == Some(&b':') {
            position += 1;
            let start = position;
            while bytes
                .get(position)
                .is_some_and(|byte| is_xml_name_byte(*byte))
            {
                position += 1;
            }
            if start == position {
                return Err(Error::Other(
                    "prefixMappings has an empty prefix".to_owned(),
                ));
            }
            let prefix = std::str::from_utf8(&bytes[start..position])
                .map_err(|_| Error::Other("prefixMappings prefix is not UTF-8".to_owned()))?
                .to_owned();
            if !is_xml_name(&prefix) {
                return Err(Error::Other(
                    "prefixMappings contains an unsupported prefix".to_owned(),
                ));
            }
            prefix
        } else {
            String::new()
        };
        skip_ascii_whitespace(bytes, &mut position);
        if bytes.get(position) != Some(&b'=') {
            return Err(Error::Other("prefixMappings is missing '='".to_owned()));
        }
        position += 1;
        skip_ascii_whitespace(bytes, &mut position);
        let quote = *bytes
            .get(position)
            .filter(|quote| **quote == b'\'' || **quote == b'"')
            .ok_or_else(|| Error::Other("prefixMappings URI is not quoted".to_owned()))?;
        position += 1;
        let start = position;
        while bytes.get(position).is_some_and(|byte| *byte != quote) {
            position += 1;
        }
        if bytes.get(position) != Some(&quote) {
            return Err(Error::Other(
                "prefixMappings URI has no closing quote".to_owned(),
            ));
        }
        let uri = bytes[start..position].to_vec();
        position += 1;
        if let Some(previous) = mappings.insert(prefix, uri.clone())
            && previous != uri
        {
            return Err(Error::Other(
                "prefixMappings defines one prefix more than once".to_owned(),
            ));
        }
    }
    Ok(mappings)
}

fn parse_xpath(xpath: &str, mappings: &HashMap<String, Vec<u8>>) -> Result<Vec<XPathStep>> {
    if !xpath.starts_with('/') || xpath.starts_with("//") {
        return Err(Error::Other(
            "content-control XPath must be an absolute child path".to_owned(),
        ));
    }
    let mut steps = Vec::new();
    for segment in xpath[1..].split('/') {
        if segment.is_empty() {
            return Err(Error::Other(
                "content-control XPath contains a descendant or empty step".to_owned(),
            ));
        }
        let (name, index) = parse_xpath_segment(segment)?;
        let mut names = name.split(':');
        let first = names.next().unwrap_or_default();
        let second = names.next();
        if names.next().is_some() {
            return Err(Error::Other(
                "content-control XPath contains an invalid qualified name".to_owned(),
            ));
        }
        let (namespace, local_name) = match second {
            Some(local) => {
                if !is_xml_name(first) {
                    return Err(Error::Other(
                        "content-control XPath contains an unsupported prefix".to_owned(),
                    ));
                }
                let namespace = mappings.get(first).cloned().ok_or_else(|| {
                    Error::Other(format!(
                        "content-control XPath prefix {first} is not mapped"
                    ))
                })?;
                (Some(namespace), local)
            }
            None => (mappings.get("").cloned(), first),
        };
        if !is_xml_name(local_name) {
            return Err(Error::Other(
                "content-control XPath contains an unsupported name".to_owned(),
            ));
        }
        steps.push(XPathStep {
            namespace,
            local_name: local_name.as_bytes().to_vec(),
            index,
        });
    }
    if steps.is_empty() {
        return Err(Error::Other("content-control XPath is empty".to_owned()));
    }
    Ok(steps)
}

fn parse_xpath_segment(segment: &str) -> Result<(&str, Option<usize>)> {
    let Some(open) = segment.find('[') else {
        if segment.contains(['*', '(', ')', ']']) {
            return Err(Error::Other(
                "content-control XPath uses an unsupported selector".to_owned(),
            ));
        }
        return Ok((segment, None));
    };
    if !segment.ends_with(']') || segment[open + 1..segment.len() - 1].contains(['[', ']']) {
        return Err(Error::Other(
            "content-control XPath uses an unsupported predicate".to_owned(),
        ));
    }
    let name = &segment[..open];
    let index = segment[open + 1..segment.len() - 1]
        .parse::<usize>()
        .ok()
        .filter(|index| *index > 0)
        .ok_or_else(|| Error::Other("content-control XPath index must be one-based".to_owned()))?;
    if name.is_empty() || name.contains(['*', '(', ')', '[', ']']) {
        return Err(Error::Other(
            "content-control XPath uses an unsupported selector".to_owned(),
        ));
    }
    Ok((name, Some(index)))
}

fn skip_ascii_whitespace(bytes: &[u8], position: &mut usize) {
    while bytes
        .get(*position)
        .is_some_and(|byte| byte.is_ascii_whitespace())
    {
        *position += 1;
    }
}

fn is_xml_name_byte(byte: u8) -> bool {
    byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-' | b'.')
}

fn is_xml_name(value: &str) -> bool {
    value
        .bytes()
        .next()
        .is_some_and(|byte| byte.is_ascii_alphabetic() || byte == b'_')
        && value.bytes().skip(1).all(is_xml_name_byte)
}

#[derive(Debug, Clone)]
struct XmlNode {
    parent: Option<usize>,
    namespace: Option<Vec<u8>>,
    local_name: Vec<u8>,
    occurrence: usize,
    replace_start: usize,
    replace_end: usize,
    element_start: usize,
    element_end: usize,
    qualified_name: Vec<u8>,
    empty: bool,
    simple_text: bool,
}

#[derive(Debug, Clone)]
struct XmlSelection {
    replace_start: usize,
    replace_end: usize,
    element_start: usize,
    element_end: usize,
    qualified_name: Vec<u8>,
    empty: bool,
}

fn select_xml_element(xml: &[u8], steps: &[XPathStep]) -> Result<XmlSelection> {
    let nodes = parse_xml_nodes(xml)?;
    let mut candidates = Vec::new();
    for (step_index, step) in steps.iter().enumerate() {
        let parents = candidates;
        candidates = nodes
            .iter()
            .enumerate()
            .filter(|(_, node)| {
                let parent_matches = if step_index == 0 {
                    node.parent.is_none()
                } else {
                    node.parent.is_some_and(|parent| parents.contains(&parent))
                };
                parent_matches
                    && node.namespace == step.namespace
                    && node.local_name == step.local_name
                    && step.index.is_none_or(|index| node.occurrence == index)
            })
            .map(|(index, _)| index)
            .collect();
    }
    let [selected_index] = candidates.as_slice() else {
        return if candidates.is_empty() {
            Err(Error::Other(
                "content-control XPath selected no custom XML element".to_owned(),
            ))
        } else {
            Err(Error::Other(
                "content-control XPath selected more than one custom XML element".to_owned(),
            ))
        };
    };
    let selected = &nodes[*selected_index];
    if !selected.simple_text {
        return Err(Error::Other(
            "content-control XPath selected an element with preserved markup".to_owned(),
        ));
    }
    Ok(XmlSelection {
        replace_start: selected.replace_start,
        replace_end: selected.replace_end,
        element_start: selected.element_start,
        element_end: selected.element_end,
        qualified_name: selected.qualified_name.clone(),
        empty: selected.empty,
    })
}

fn parse_xml_nodes(xml: &[u8]) -> Result<Vec<XmlNode>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut nodes: Vec<XmlNode> = Vec::new();
    let mut stack: Vec<usize> = Vec::new();
    loop {
        let before = usize::try_from(reader.buffer_position())
            .map_err(|_| Error::Other("custom XML position overflow".to_owned()))?;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("invalid custom XML part: {error}")))?;
        let namespace = namespace_bytes(&namespace)?;
        let event = event.into_owned();
        let after = usize::try_from(reader.buffer_position())
            .map_err(|_| Error::Other("custom XML position overflow".to_owned()))?;
        match event {
            Event::Start(element) => {
                if let Some(parent) = stack.last().copied() {
                    nodes[parent].simple_text = false;
                }
                let parent = stack.last().copied();
                let qualified_name = element.name().as_ref().to_vec();
                let local_name = qualified_name
                    .rsplit(|byte| *byte == b':')
                    .next()
                    .unwrap_or(&qualified_name)
                    .to_vec();
                let occurrence = nodes
                    .iter()
                    .filter(|node| {
                        node.parent == parent
                            && node.namespace == namespace
                            && node.local_name == local_name
                    })
                    .count()
                    .checked_add(1)
                    .ok_or_else(|| Error::Other("custom XML sibling count overflow".to_owned()))?;
                let index = nodes.len();
                nodes.push(XmlNode {
                    parent,
                    namespace,
                    local_name,
                    occurrence,
                    replace_start: after,
                    replace_end: 0,
                    element_start: before,
                    element_end: after,
                    qualified_name,
                    empty: false,
                    simple_text: true,
                });
                stack.push(index);
            }
            Event::Empty(element) => {
                if let Some(parent) = stack.last().copied() {
                    nodes[parent].simple_text = false;
                }
                let parent = stack.last().copied();
                let qualified_name = element.name().as_ref().to_vec();
                let local_name = qualified_name
                    .rsplit(|byte| *byte == b':')
                    .next()
                    .unwrap_or(&qualified_name)
                    .to_vec();
                let occurrence = nodes
                    .iter()
                    .filter(|node| {
                        node.parent == parent
                            && node.namespace == namespace
                            && node.local_name == local_name
                    })
                    .count()
                    .checked_add(1)
                    .ok_or_else(|| Error::Other("custom XML sibling count overflow".to_owned()))?;
                nodes.push(XmlNode {
                    parent,
                    namespace,
                    local_name,
                    occurrence,
                    replace_start: before,
                    replace_end: after,
                    element_start: before,
                    element_end: after,
                    qualified_name,
                    empty: true,
                    simple_text: true,
                });
            }
            Event::End(_) => {
                let index = stack.pop().ok_or_else(|| {
                    Error::Other("custom XML contains an unmatched end element".to_owned())
                })?;
                nodes[index].replace_end = before;
                nodes[index].element_end = after;
            }
            Event::Comment(_) | Event::PI(_) | Event::DocType(_) => {
                if let Some(index) = stack.last().copied() {
                    nodes[index].simple_text = false;
                }
            }
            Event::Eof => {
                if !stack.is_empty() {
                    return Err(Error::Other(
                        "custom XML contains an unclosed element".to_owned(),
                    ));
                }
                return Ok(nodes);
            }
            _ => {}
        }
        buffer.clear();
    }
}

fn namespace_bytes(namespace: &ResolveResult<'_>) -> Result<Option<Vec<u8>>> {
    match namespace {
        ResolveResult::Bound(Namespace(uri)) => Ok(Some(uri.to_vec())),
        ResolveResult::Unbound => Ok(None),
        ResolveResult::Unknown(prefix) => Err(Error::Other(format!(
            "custom XML uses unmapped namespace prefix {}",
            String::from_utf8_lossy(prefix)
        ))),
    }
}

fn namespace_matches(namespace: &ResolveResult<'_>, expected: Option<&[u8]>) -> bool {
    match (namespace, expected) {
        (ResolveResult::Bound(Namespace(actual)), Some(expected)) => *actual == expected,
        (ResolveResult::Unbound, None) => true,
        _ => false,
    }
}

fn expand_empty_element(xml: &[u8], selection: &XmlSelection, value: &str) -> Result<Vec<u8>> {
    let raw = xml
        .get(selection.element_start..selection.element_end)
        .ok_or_else(|| Error::Other("custom XML selection is out of bounds".to_owned()))?;
    let slash = raw
        .iter()
        .rposition(|byte| *byte == b'/')
        .filter(|slash| raw.get(slash + 1) == Some(&b'>'))
        .ok_or_else(|| Error::Other("custom XML empty element is malformed".to_owned()))?;
    let mut expanded = Vec::new();
    expanded.extend_from_slice(&raw[..slash]);
    expanded.push(b'>');
    expanded.extend_from_slice(quick_xml::escape::escape(value).as_bytes());
    expanded.extend_from_slice(b"</");
    expanded.extend_from_slice(&selection.qualified_name);
    expanded.push(b'>');
    Ok(expanded)
}

fn apply_xml_edits(xml: &[u8], mut edits: Vec<XmlEdit>) -> Result<Vec<u8>> {
    edits.sort_by_key(|edit| (edit.start, edit.end));
    let mut unique: Vec<XmlEdit> = Vec::new();
    for edit in edits {
        if let Some(previous) = unique.last() {
            if previous.start == edit.start && previous.end == edit.end {
                if previous.replacement == edit.replacement {
                    continue;
                }
                return Err(Error::Other(
                    "content controls assign different values to one custom XML element".to_owned(),
                ));
            }
            if edit.start < previous.end {
                return Err(Error::Other(
                    "content-control custom XML selections overlap".to_owned(),
                ));
            }
        }
        unique.push(edit);
    }
    let mut output = xml.to_vec();
    for edit in unique.into_iter().rev() {
        if edit.start > edit.end || edit.end > output.len() {
            return Err(Error::Other(
                "content-control custom XML selection is out of bounds".to_owned(),
            ));
        }
        output.splice(edit.start..edit.end, edit.replacement);
    }
    validate_xml(&output)?;
    Ok(output)
}

fn validate_xml(xml: &[u8]) -> Result<()> {
    let mut reader = NsReader::from_reader(xml);
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Eof) => return Ok(()),
            Ok(_) => buffer.clear(),
            Err(error) => {
                return Err(Error::Other(format!(
                    "content-control custom XML edit is invalid: {error}"
                )));
            }
        }
    }
}
