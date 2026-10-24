//! Read-only content-control handles and atomic value binding.

use std::collections::{HashMap, HashSet};
use std::ops::Range;

use oxml_opc::OpcPackage;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesRef, BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::{NsReader, Reader};
use rdocx_oxml::content_control::{CT_DataBinding, CT_Sdt, SdtContent, SdtType};
use rdocx_oxml::document::{BodyContent, CT_Body};
use rdocx_oxml::namespace::{W_NS, matches_local_name};
use rdocx_oxml::table::{CT_Row, CT_Tbl, CT_Tc, CellContent};
use rdocx_oxml::text::{CT_P, CT_R, CT_Text, RunContent};

use crate::{ContentFragment, ContentLocation, Document, Error, Result, StoryItemKind, StoryKind};

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
}

#[derive(Clone)]
struct ControlValue {
    display: String,
    bound: String,
    checkbox: Option<(bool, String)>,
    full_date: Option<String>,
}

fn prepare_control_value(control: &CT_Sdt, value: &str) -> Result<ControlValue> {
    let kind = control
        .properties
        .as_ref()
        .and_then(|properties| properties.control_type);
    if kind == Some(SdtType::CheckBox) {
        let checked = match value.trim().to_ascii_lowercase().as_str() {
            "true" | "1" | "yes" => true,
            "false" | "0" | "no" => false,
            _ => {
                return Err(Error::Other(
                    "checkbox value must be true or false (also 1/0, yes/no)".to_owned(),
                ));
            }
        };
        let (glyph, font) = control
            .properties
            .as_ref()
            .unwrap()
            .checkbox_state(checked)?;
        return Ok(ControlValue {
            display: glyph.to_string(),
            bound: checked.to_string(),
            checkbox: Some((checked, font)),
            full_date: None,
        });
    }
    if matches!(kind, Some(SdtType::Picture | SdtType::Group)) {
        return Err(Error::Other(
            "selected content-control kind does not support a text value".to_owned(),
        ));
    }
    let (display, bound, full_date) = if let Some(properties) = &control.properties {
        properties.prepare_text_value(value)?
    } else {
        (value.to_owned(), value.to_owned(), None)
    };
    Ok(ControlValue {
        display,
        bound,
        checkbox: None,
        full_date,
    })
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

    /// Fill distinct tags atomically, refusing missing tags and invalid late owners.
    /// Aliases do not participate. Counts retain the caller's order.
    pub fn fill_content_control_tags(&mut self, pairs: &[(&str, &str)]) -> Result<Vec<usize>> {
        let mut seen = std::collections::HashSet::new();
        for &(tag, value) in pairs {
            if !seen.insert(tag) {
                return Err(Error::Other(format!(
                    "duplicate content-control tag {tag:?}"
                )));
            }
            oxml_core::xml::reject_non_xml_characters("content-control value", value)?;
        }
        if pairs.is_empty() {
            return Ok(Vec::new());
        }
        let mut candidate = self.clone_for_staging();
        let mut counts = Vec::with_capacity(pairs.len());
        for &(tag, value) in pairs {
            let count = candidate.set_content_control_value_by_tag(tag, value)?;
            if count == 0 {
                return Err(Error::Other(format!("no content control has tag {tag:?}")));
            }
            counts.push(count);
        }
        self.commit_staged_mutation(candidate.prepare_and_reopen_staged()?);
        Ok(counts)
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
        let mut current = String::new();
        collect_sdt_text(control, &mut current);
        if current != value.display || value.checkbox.is_some() {
            replace_display_text(
                control,
                &value.display,
                value.checkbox.as_ref().map(|(_, font)| font.as_str()),
            )?;
        }
        if let Some(properties) = &mut control.properties {
            properties.set_value_metadata(value.full_date.as_deref())?;
        }
        if let Some((checked, _)) = &value.checkbox {
            control
                .properties
                .as_mut()
                .unwrap()
                .set_checkbox_checked(*checked)?;
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

fn replace_display_text(control: &mut CT_Sdt, value: &str, font: Option<&str>) -> Result<()> {
    let mut replacement = DisplayReplacement {
        value,
        font,
        wrote_value: false,
        saw_run: false,
    };
    replace_sdt_display(control, &mut replacement);
    if replacement.wrote_value {
        return Ok(());
    }
    if add_text_to_first_paragraph(control, value) {
        if font.is_some() {
            let mut replacement = DisplayReplacement {
                value,
                font,
                wrote_value: false,
                saw_run: false,
            };
            replace_sdt_display(control, &mut replacement);
        }
        return Ok(());
    }
    Err(Error::Other(
        "content control has no typed text display location".to_owned(),
    ))
}

struct DisplayReplacement<'a> {
    value: &'a str,
    font: Option<&'a str>,
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
    if is_first_run && let Some(font) = replacement.font {
        let properties = run.properties.get_or_insert_with(Default::default);
        properties.font_ascii = Some(font.to_owned());
        properties.font_hansi = Some(font.to_owned());
        properties.font_east_asia = Some(font.to_owned());
        properties.font_cs = Some(font.to_owned());
        properties.font_ascii_theme = None;
        properties.font_hansi_theme = None;
        properties.font_east_asia_theme = None;
        properties.font_cs_theme = None;
    }
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

const W14_NS: &str = "http://schemas.microsoft.com/office/word/2010/wordml";
const W15_NS: &str = "http://schemas.microsoft.com/office/word/2012/wordml";
const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/// Where a new content control is inserted.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ContentControlPlacement {
    Block,
    Row,
    Cell,
    Paragraph,
    Inline,
}

/// Lock stored on `w:lock`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ContentControlLock {
    Unlocked,
    SdtLocked,
    ContentLocked,
    SdtContentLocked,
}

impl ContentControlLock {
    fn as_str(self) -> &'static str {
        match self {
            Self::Unlocked => "unlocked",
            Self::SdtLocked => "sdtLocked",
            Self::ContentLocked => "contentLocked",
            Self::SdtContentLocked => "sdtContentLocked",
        }
    }
}

/// Appearance stored on `w15:appearance`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ContentControlAppearance {
    BoundingBox,
    Tags,
    Hidden,
}

impl ContentControlAppearance {
    fn as_str(self) -> &'static str {
        match self {
            Self::BoundingBox => "boundingBox",
            Self::Tags => "tags",
            Self::Hidden => "hidden",
        }
    }
}

/// Type payload written onto an existing rich or plain shell.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ContentControlTypePayload {
    RichText,
    PlainText,
    Date { format: String },
    CheckBox,
    ComboBox { items: Vec<(String, String)> },
    DropDownList { items: Vec<(String, String)> },
    Picture { relationship_id: String },
    Group,
    RepeatingSection,
    RepeatingSectionItem,
}

/// Fields for one rich-text or plain-text shell.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ContentControlSpec {
    pub placement: ContentControlPlacement,
    pub kind: SdtType,
    pub tag: Option<String>,
    pub alias: Option<String>,
    pub lock: Option<ContentControlLock>,
    pub appearance: Option<ContentControlAppearance>,
    pub placeholder: Option<bool>,
}

impl Document {
    /// Insert one rich-text or plain-text shell at `destination`.
    pub fn insert_content_control(
        &mut self,
        destination: &ContentLocation,
        spec: ContentControlSpec,
    ) -> Result<()> {
        reject_control_text("content-control tag", spec.tag.as_deref())?;
        reject_control_text("content-control alias", spec.alias.as_deref())?;
        let kind_name = plain_or_rich(spec.kind)?;
        let id = self.smallest_content_control_id()?;
        let shell = control_shell(&spec, id, kind_name)?;
        let mut candidate = self.clone_for_staging();
        candidate.flush_to_package()?;
        match spec.placement {
            ContentControlPlacement::Block | ContentControlPlacement::Paragraph => {
                require_block_parent(&candidate, destination)?;
                let control = parse_control(&shell)?;
                candidate
                    .insert_content(destination, ContentFragment::content_control(control)?)
                    .map_err(control_fault)?;
                self.commit_staged_mutation(candidate);
            }
            ContentControlPlacement::Row => {
                let located = table_parent_range(&candidate, destination)?;
                let columns = table_column_count(&located.xml, &located.range)?;
                let shell =
                    control_shell_with_body(&spec, id, kind_name, &row_shell(columns, spec.kind)?)?;
                splice_inside(&mut candidate, located, &shell)?;
                self.publish_candidate(candidate)?;
            }
            ContentControlPlacement::Cell => {
                let located = row_parent_range(&candidate, destination)?;
                let shell = control_shell_with_body(&spec, id, kind_name, &cell_shell(spec.kind)?)?;
                splice_inside(&mut candidate, located, &shell)?;
                self.publish_candidate(candidate)?;
            }
            ContentControlPlacement::Inline => {
                let located = paragraph_parent_range(&candidate, destination)?;
                if shell_contains_paragraph(&shell)? {
                    return Err(Error::Other(
                        "inline content control cannot contain a paragraph".to_owned(),
                    ));
                }
                splice_inside(&mut candidate, located, &shell)?;
                self.publish_candidate(candidate)?;
            }
        }
        Ok(())
    }

    /// Stage one custom XML item and its properties part.
    ///
    /// The returned value is the braced item id. An empty URI list writes an
    /// empty `ds:schemaRefs`. Invalid item bytes and an empty schema URI
    /// publish nothing.
    pub fn add_custom_xml_store(
        &mut self,
        item_xml: &[u8],
        schema_uris: &[&str],
    ) -> Result<String> {
        validate_store_item(item_xml)?;
        for uri in schema_uris {
            if uri.is_empty() {
                return Err(Error::Other("custom XML schema URI is empty".to_owned()));
            }
            oxml_core::xml::reject_non_xml_characters("custom XML schema URI", uri)?;
        }
        let mut candidate = self.clone_for_staging();
        candidate.prepare_staged_package()?;
        let staged =
            crate::bibliography::stage_custom_xml_store(&mut candidate, item_xml, schema_uris)?;
        let mut reopened = candidate.reopen_prepared_staged().map_err(control_fault)?;
        reopened.invalidate_layout();
        let item_id = staged.item_id;
        self.commit_staged_mutation(reopened);
        Ok(item_id)
    }

    /// Attach `w:dataBinding` to one control after the store and path resolve.
    ///
    /// `None` and an empty mapping string omit `w:prefixMappings`. A failed
    /// check publishes nothing.
    pub fn set_content_control_binding(
        &mut self,
        location: &ContentLocation,
        store_item_id: &str,
        xpath: &str,
        prefix_mappings: Option<&str>,
    ) -> Result<()> {
        let mappings_text = prefix_mappings.filter(|value| !value.is_empty());
        let mut candidate = self.clone_for_staging();
        candidate.flush_to_package()?;
        let mappings = parse_prefix_mappings(mappings_text.unwrap_or(""))?;
        let steps = parse_xpath(xpath, &mappings)?;
        let part_name =
            resolve_custom_xml_part(&candidate.package, &candidate.doc_part_name, store_item_id)?;
        {
            let item = candidate
                .package
                .get_part(&part_name)
                .ok_or_else(|| Error::Other(format!("custom XML part {part_name} is missing")))?;
            select_xml_element(item, &steps)?;
        }
        let (part, mut xml, range, _) = control_item(&candidate, location)?;
        let updated = edit_sdt_property(
            &xml,
            &range,
            PropertyEdit::DataBinding {
                prefix_mappings: mappings_text.map(str::to_owned),
                xpath: xpath.to_owned(),
                store_item_id: store_item_id.to_owned(),
            },
        )?;
        xml.splice(range, updated);
        candidate.write_story_part(&part, xml)?;
        self.publish_candidate(candidate)
    }

    /// Set or clear `w:tag` without rewriting `sdtContent`.
    pub fn set_content_control_tag(
        &mut self,
        location: &ContentLocation,
        tag: Option<&str>,
    ) -> Result<()> {
        reject_control_text("content-control tag", tag)?;
        self.edit_control_property(location, PropertyEdit::Tag(tag.map(str::to_owned)))
    }

    /// Set or clear `w:alias` without rewriting `sdtContent`.
    pub fn set_content_control_alias(
        &mut self,
        location: &ContentLocation,
        alias: Option<&str>,
    ) -> Result<()> {
        reject_control_text("content-control alias", alias)?;
        self.edit_control_property(location, PropertyEdit::Alias(alias.map(str::to_owned)))
    }

    /// Set or clear `w:lock` without rewriting `sdtContent`.
    pub fn set_content_control_lock(
        &mut self,
        location: &ContentLocation,
        lock: Option<ContentControlLock>,
    ) -> Result<()> {
        self.edit_control_property(location, PropertyEdit::Lock(lock))
    }

    /// Set or clear `w15:appearance` without rewriting `sdtContent`.
    pub fn set_content_control_appearance(
        &mut self,
        location: &ContentLocation,
        appearance: Option<ContentControlAppearance>,
    ) -> Result<()> {
        self.edit_control_property(location, PropertyEdit::Appearance(appearance))
    }

    /// Set or clear the qualified `w:showingPlcHdr` marker.
    ///
    /// A foreign-namespace lookalike stays in place.
    pub fn set_content_control_placeholder(
        &mut self,
        location: &ContentLocation,
        placeholder: Option<bool>,
    ) -> Result<()> {
        self.edit_control_property(location, PropertyEdit::Placeholder(placeholder))
    }

    /// Replace the type element on an existing shell.
    ///
    /// Tag, alias, id, lock, appearance, placeholder, and unmodelled property
    /// children stay. A refusal publishes nothing.
    pub fn set_content_control_type(
        &mut self,
        location: &ContentLocation,
        payload: ContentControlTypePayload,
    ) -> Result<()> {
        validate_type_payload(&payload)?;
        let mut candidate = self.clone_for_staging();
        candidate.flush_to_package()?;
        let (part, mut xml, mut range, _) = control_item(&candidate, location)?;
        validate_repeating_type_change(&xml, &range, &payload)?;
        match &payload {
            ContentControlTypePayload::Picture { relationship_id } => {
                if !part_owns_image(&candidate, &part, relationship_id) {
                    return Err(Error::Other(
                        "content control picture relationship is missing".to_owned(),
                    ));
                }
                if !content_embeds_relationship(&xml, &range, relationship_id)? {
                    return Err(Error::Other(
                        "content control picture content does not embed that relationship"
                            .to_owned(),
                    ));
                }
            }
            ContentControlTypePayload::RepeatingSection => match repeating_content(&xml, &range)? {
                RepeatingContent::Empty => {
                    return Err(Error::Other("repeating section has no item".to_owned()));
                }
                RepeatingContent::NotBlock => {
                    return Err(Error::Other(
                        "repeating section content is not a block sequence".to_owned(),
                    ));
                }
                RepeatingContent::AlreadyItem => {}
                RepeatingContent::Wrap => {
                    let id = candidate.smallest_content_control_id()?;
                    (xml, range) = wrap_repeating_item(&xml, &range, id)?;
                }
            },
            _ => {}
        }
        let markup = type_markup(&payload)?;
        let updated = replace_control_type_markup(&xml, &range, markup.as_bytes())?;
        xml.splice(range, updated);
        candidate.write_story_part(&part, xml)?;
        self.publish_candidate(candidate)
    }

    /// Clone the first repeating item inside a repeating section and append it.
    pub fn repeat_content_control(&mut self, location: &ContentLocation) -> Result<()> {
        let mut candidate = self.clone_for_staging();
        candidate.flush_to_package()?;
        let (part, mut xml, range, _) = control_item(&candidate, location)?;
        let elems = index_elements(&xml)?;
        let section = element_at(&elems, &range)?;
        if !sdt_is_repeating_section(&elems, section) {
            return Err(Error::Other(
                "repeat_content_control addresses a repeating section".to_owned(),
            ));
        }
        let Some(item_range) = first_repeating_item_range(&elems, section) else {
            return Err(Error::Other(
                "repeating section has no item to clone".to_owned(),
            ));
        };
        let bytes = xml
            .get(item_range)
            .ok_or_else(|| Error::Other("content control bytes are missing".to_owned()))?
            .to_vec();
        let fresh = freshen_control(&mut candidate, &bytes)?;
        let elems = index_elements(&xml)?;
        let section = element_at(&elems, &range)?;
        let content = sdt_content_elem(&elems, section)
            .ok_or_else(|| Error::Other("repeating section has no item to clone".to_owned()))?;
        let content_range = content.range.clone();
        let replacement = insert_into_element(&xml, &content_range, &fresh)?;
        xml.splice(content_range, replacement);
        candidate.write_story_part(&part, xml)?;
        self.publish_candidate(candidate)
    }

    /// Delete one control, including the content inside it.
    pub fn remove_content_control(&mut self, location: &ContentLocation) -> Result<()> {
        let (_part, xml, range, direct) = control_item(self, location)?;
        if last_repeating_item(&xml, &range)? {
            return Err(Error::Other(
                "cannot remove the last repeating section item".to_owned(),
            ));
        }
        if direct && control_parent_is_block(&xml, &range)? {
            refuse_cell_without_paragraph(self, location.story(), &xml, &range, b"")?;
            return self
                .remove_content_at(location)
                .map(|_| ())
                .map_err(control_fault);
        }
        self.rewrite_control(location, ControlRewrite::Remove)
    }

    /// Drop the shell and leave its content when the parent can hold it.
    pub fn unwrap_content_control(&mut self, location: &ContentLocation) -> Result<()> {
        self.rewrite_control(location, ControlRewrite::Unwrap)
    }

    /// Move one control inside the same story. The id stays.
    pub fn move_content_control(
        &mut self,
        source: &ContentLocation,
        destination: &ContentLocation,
    ) -> Result<()> {
        if source.story() != destination.story() {
            return Err(Error::Other(
                "content control can move only within one story".to_owned(),
            ));
        }
        let (_part, xml, range, direct) = control_item(self, source)?;
        if last_repeating_item(&xml, &range)? {
            return Err(Error::Other(
                "cannot move the last repeating section item".to_owned(),
            ));
        }
        if direct && control_parent_is_block(&xml, &range)? {
            if !control_matches_context(
                ParentSlot::Block,
                story_inside_repeating_section(self, destination.story())?,
                &xml,
                &range,
            )? {
                return Err(Error::Other(
                    "content control placement does not match its parent".to_owned(),
                ));
            }
            return self
                .move_content(source, destination)
                .map_err(control_fault);
        }
        let bytes = xml
            .get(range.clone())
            .ok_or_else(|| Error::Other("content control bytes are missing".to_owned()))?
            .to_vec();
        let mut candidate = self.clone_for_staging();
        candidate.flush_to_package()?;
        let (part, mut updated, source_range, _) = control_item(&candidate, source)?;
        let dest = nested_destination(&candidate, destination)?;
        if part != dest.part {
            return Err(Error::Other(
                "content control can move only within one story".to_owned(),
            ));
        }
        if ranges_overlap(&source_range, &dest.range) {
            return Err(Error::Other(
                "content control move destination is inside the source".to_owned(),
            ));
        }
        destination_accepts_control(&dest.xml, &dest.range, &updated, &source_range)?;
        let original_len = updated.len();
        updated.drain(source_range.clone());
        let dest_range = shift_after_removal(&dest.range, &source_range);
        let replacement = insert_into_element(&updated, &dest_range, &bytes)?;
        updated.splice(dest_range, replacement);
        ensure_cell_paragraph(&candidate, source.story(), &updated, original_len)?;
        candidate.write_story_part(&part, updated)?;
        self.publish_candidate(candidate)
    }

    /// Clone one control. The copy gets a new id and drops paragraph identity.
    pub fn clone_content_control(
        &mut self,
        source: &ContentLocation,
        destination: &ContentLocation,
    ) -> Result<()> {
        let (_part, xml, range, direct) = control_item(self, source)?;
        if direct && control_parent_is_block(&xml, &range)? {
            if !control_matches_context(
                ParentSlot::Block,
                story_inside_repeating_section(self, destination.story())?,
                &xml,
                &range,
            )? {
                return Err(Error::Other(
                    "content control placement does not match its parent".to_owned(),
                ));
            }
            return self
                .clone_content(source, destination)
                .map_err(control_fault);
        }
        let bytes = xml
            .get(range)
            .ok_or_else(|| Error::Other("content control bytes are missing".to_owned()))?
            .to_vec();
        let mut candidate = self.clone_for_staging();
        candidate.flush_to_package()?;
        let (_, source_xml, source_range, _) = control_item(&candidate, source)?;
        let dest = nested_destination(&candidate, destination)?;
        destination_accepts_control(&dest.xml, &dest.range, &source_xml, &source_range)?;
        let fresh = freshen_control(&mut candidate, &bytes)?;
        let mut dest_xml = dest.xml;
        let replacement = insert_into_element(&dest_xml, &dest.range, &fresh)?;
        dest_xml.splice(dest.range, replacement);
        candidate.write_story_part(&dest.part, dest_xml)?;
        self.publish_candidate(candidate)
    }

    fn publish_candidate(&mut self, candidate: Document) -> Result<()> {
        let reopened = candidate
            .prepare_and_reopen_staged()
            .map_err(control_fault)?;
        self.commit_staged_mutation(reopened);
        Ok(())
    }

    fn smallest_content_control_id(&self) -> Result<i32> {
        let mut used = HashSet::new();
        for xml in self.owned_story_parts()? {
            collect_control_ids(&xml, &mut used)?;
        }
        let mut id = 1i32;
        while used.contains(&id) {
            id = id
                .checked_add(1)
                .ok_or_else(|| Error::Other("content control ids are exhausted".to_owned()))?;
        }
        Ok(id)
    }

    fn edit_control_property(
        &mut self,
        location: &ContentLocation,
        edit: PropertyEdit,
    ) -> Result<()> {
        let mut candidate = self.clone_for_staging();
        candidate.flush_to_package()?;
        let (part, mut xml, range, _) = control_item(&candidate, location)?;
        let updated = edit_sdt_property(&xml, &range, edit)?;
        let _ = xml.splice(range, updated);
        candidate.write_story_part(&part, xml)?;
        self.publish_candidate(candidate)
    }

    fn rewrite_control(
        &mut self,
        location: &ContentLocation,
        rewrite: ControlRewrite,
    ) -> Result<()> {
        let mut candidate = self.clone_for_staging();
        candidate.flush_to_package()?;
        let (part, mut xml, range, _) = control_item(&candidate, location)?;
        let replacement = match rewrite {
            ControlRewrite::Remove => {
                refuse_cell_without_paragraph(&candidate, location.story(), &xml, &range, b"")?;
                Vec::new()
            }
            ControlRewrite::Unwrap => {
                let inner = sdt_content_bytes(&xml, &range)?;
                if !unwrap_keeps_repeating_context(&xml, &range)?
                    || !parent_can_hold(&xml, &range, &inner)?
                {
                    return Err(Error::Other(
                        "content control unwrap leaves content the parent cannot hold".to_owned(),
                    ));
                }
                refuse_cell_without_paragraph(&candidate, location.story(), &xml, &range, &inner)?;
                inner
            }
        };
        xml.splice(range, replacement);
        candidate.write_story_part(&part, xml)?;
        self.publish_candidate(candidate)
    }
}

enum PropertyEdit {
    Tag(Option<String>),
    Alias(Option<String>),
    Lock(Option<ContentControlLock>),
    Appearance(Option<ContentControlAppearance>),
    Placeholder(Option<bool>),
    DataBinding {
        prefix_mappings: Option<String>,
        xpath: String,
        store_item_id: String,
    },
}

enum ControlRewrite {
    Remove,
    Unwrap,
}

fn reject_control_text(what: &str, value: Option<&str>) -> Result<()> {
    if let Some(value) = value {
        oxml_core::xml::reject_non_xml_characters(what, value)?;
    }
    Ok(())
}

fn validate_store_item(xml: &[u8]) -> Result<()> {
    if xml.is_empty() {
        return Err(Error::Other("custom XML item is empty".to_owned()));
    }
    let text = std::str::from_utf8(xml)
        .map_err(|error| Error::Other(format!("custom XML item is not XML: {error}")))?;
    oxml_core::xml::reject_non_xml_characters("custom XML item is not XML", text)?;
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_comments = true;
    let mut buffer = Vec::new();
    let mut roots = 0usize;
    let mut depth = 0usize;
    let mut declaration_allowed = true;
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(start)) => {
                reject_store_attributes(&start)?;
                if depth == 0 {
                    roots += 1;
                    if roots > 1 {
                        return Err(Error::Other("custom XML item is a fragment".to_owned()));
                    }
                }
                depth += 1;
            }
            Ok(Event::Empty(start)) => {
                reject_store_attributes(&start)?;
                if depth == 0 {
                    roots += 1;
                    if roots > 1 {
                        return Err(Error::Other("custom XML item is a fragment".to_owned()));
                    }
                }
            }
            Ok(Event::End(_)) => {
                if depth == 0 {
                    return Err(Error::Other(
                        "custom XML item is not one element".to_owned(),
                    ));
                }
                depth -= 1;
            }
            Ok(Event::Text(text)) => {
                if text.windows(3).any(|bytes| bytes == b"]]>") {
                    return Err(Error::Other(
                        "custom XML item is not XML: literal CDATA terminator in text".to_owned(),
                    ));
                }
                if depth == 0 && !is_xml_whitespace(text.as_ref()) {
                    return Err(Error::Other("custom XML item is a fragment".to_owned()));
                }
            }
            Ok(Event::CData(_)) => {
                if depth == 0 {
                    return Err(Error::Other("custom XML item is a fragment".to_owned()));
                }
            }
            Ok(Event::GeneralRef(reference)) => {
                if depth == 0 {
                    return Err(Error::Other("custom XML item is a fragment".to_owned()));
                }
                decode_one_xml_reference(&reference).map_err(|detail| {
                    Error::Other(format!("custom XML item is not XML: {detail}"))
                })?;
            }
            Ok(Event::Decl(declaration)) => {
                if !declaration_allowed {
                    return Err(Error::Other(
                        "custom XML item is not XML: declaration must be first".to_owned(),
                    ));
                }
                validate_store_declaration(&declaration, text.is_ascii())?;
            }
            Ok(Event::DocType(_)) => {
                return Err(Error::Other(
                    "custom XML item is not XML: DTD declarations are unsupported".to_owned(),
                ));
            }
            Ok(Event::PI(instruction)) => {
                let target = instruction.target();
                if !store_xml_name_is_valid(target) || target.eq_ignore_ascii_case(b"xml") {
                    return Err(Error::Other(
                        "custom XML item is not XML: invalid processing instruction target"
                            .to_owned(),
                    ));
                }
            }
            Ok(Event::Comment(_)) => {} // The reader checks comment syntax.
            Ok(Event::Eof) => break,
            Err(error) => {
                return Err(Error::Other(format!("custom XML item is not XML: {error}")));
            }
        }
        declaration_allowed = false;
        buffer.clear();
    }
    if roots != 1 || depth != 0 {
        return Err(Error::Other(
            "custom XML item is not one element".to_owned(),
        ));
    }
    Ok(())
}

fn validate_store_declaration(
    declaration: &quick_xml::events::BytesDecl<'_>,
    ascii_only: bool,
) -> Result<()> {
    let invalid =
        || Error::Other("custom XML item is not XML: invalid XML 1.0 declaration".to_owned());
    let content = std::str::from_utf8(declaration.as_ref()).map_err(|_| invalid())?;
    let start = BytesStart::from_content(content, 3);
    reject_store_attribute_separators(&start)?;
    let mut encoding_seen = false;
    let mut standalone_seen = false;
    let mut version_seen = false;
    for (index, attribute) in start.attributes().enumerate() {
        let attribute = attribute.map_err(|_| invalid())?;
        let value = attribute.value.as_ref();
        match attribute.key.as_ref() {
            b"version" if index == 0 && value == b"1.0" => version_seen = true,
            b"encoding" if version_seen && !encoding_seen && !standalone_seen => {
                if !value.first().is_some_and(u8::is_ascii_alphabetic)
                    || !value.iter().all(|byte| {
                        byte.is_ascii_alphanumeric() || matches!(byte, b'.' | b'_' | b'-')
                    })
                {
                    return Err(invalid());
                }
                if !value.eq_ignore_ascii_case(b"UTF-8")
                    && !(ascii_only && value.eq_ignore_ascii_case(b"US-ASCII"))
                {
                    return Err(Error::Other(
                        "custom XML item is not XML: encoding is unsupported or contradicts UTF-8 input"
                            .to_owned(),
                    ));
                }
                encoding_seen = true;
            }
            b"standalone"
                if version_seen && !standalone_seen && matches!(value, b"yes" | b"no") =>
            {
                standalone_seen = true;
            }
            _ => return Err(invalid()),
        }
    }
    if !version_seen {
        return Err(invalid());
    }
    Ok(())
}

fn reject_store_attributes(start: &BytesStart<'_>) -> Result<()> {
    reject_store_name(start.name().as_ref())?;
    reject_store_attribute_separators(start)?;
    for attribute in start.attributes() {
        let attribute = attribute
            .map_err(|error| Error::Other(format!("custom XML item is not XML: {error}")))?;
        reject_store_name(attribute.key.as_ref())?;
        reject_store_attribute_value(attribute.value.as_ref())?;
    }
    Ok(())
}

fn reject_store_attribute_separators(start: &BytesStart<'_>) -> Result<()> {
    // quick-xml checks quoting and duplicates, but accepts adjacent attributes.
    // Inspect only their separators here. Names and values are checked above.
    let mut quote = None;
    let mut separator_required = true;
    for &byte in start.attributes_raw() {
        if let Some(delimiter) = quote {
            if byte == delimiter {
                quote = None;
                separator_required = true;
            }
        } else {
            if separator_required {
                if !matches!(byte, b' ' | b'\t' | b'\n' | b'\r') {
                    return Err(Error::Other(
                        "custom XML item is not XML: attribute requires whitespace".to_owned(),
                    ));
                }
                separator_required = false;
            }
            if matches!(byte, b'\'' | b'"') {
                quote = Some(byte);
            }
        }
    }
    Ok(())
}

fn reject_store_name(name: &[u8]) -> Result<()> {
    if store_xml_name_is_valid(name) {
        Ok(())
    } else {
        Err(Error::Other(format!(
            "custom XML item is not XML: invalid name {}",
            String::from_utf8_lossy(name)
        )))
    }
}

fn reject_store_attribute_value(value: &[u8]) -> Result<()> {
    if value.contains(&b'<') {
        return Err(Error::Other(
            "custom XML item is not XML: attribute contains a literal less-than sign".to_owned(),
        ));
    }
    let text = std::str::from_utf8(value)
        .map_err(|error| Error::Other(format!("custom XML item is not XML: {error}")))?;
    decode_xml_reference_text(text)
        .map_err(|detail| Error::Other(format!("custom XML item is not XML: {detail}")))?;
    Ok(())
}

/// Decode XML references with quick-xml and then apply the XML 1.0 character rule.
/// Numeric syntax stays in `unescape`, including its lowercase hexadecimal marker.
fn decode_xml_reference_text(raw: &str) -> std::result::Result<std::borrow::Cow<'_, str>, String> {
    let decoded = quick_xml::escape::unescape(raw).map_err(|error| error.to_string())?;
    if decoded.chars().all(oxml_core::xml::is_xml_1_0_character) {
        Ok(decoded)
    } else {
        Err("character reference is not legal in XML 1.0".to_owned())
    }
}

fn decode_one_xml_reference(reference: &BytesRef<'_>) -> std::result::Result<char, String> {
    let body = reference.decode().map_err(|error| error.to_string())?;
    let mut raw = String::with_capacity(body.len() + 2);
    raw.push('&');
    raw.push_str(body.as_ref());
    raw.push(';');
    let decoded = decode_xml_reference_text(&raw)?;
    let mut characters = decoded.chars();
    match (characters.next(), characters.next()) {
        (Some(character), None) => Ok(character),
        _ => Err("character reference is not legal in XML 1.0".to_owned()),
    }
}

fn store_xml_name_is_valid(name: &[u8]) -> bool {
    let Ok(name) = std::str::from_utf8(name) else {
        return false;
    };
    let mut characters = name.chars();
    let Some(first) = characters.next() else {
        return false;
    };
    store_name_start(first) && characters.all(store_name_character)
}

fn store_name_start(character: char) -> bool {
    character == ':'
        || matches!(
            character,
            'A'..='Z'
                | '_'
                | 'a'..='z'
                | '\u{00C0}'..='\u{00D6}'
                | '\u{00D8}'..='\u{00F6}'
                | '\u{00F8}'..='\u{02FF}'
                | '\u{0370}'..='\u{037D}'
                | '\u{037F}'..='\u{1FFF}'
                | '\u{200C}'..='\u{200D}'
                | '\u{2070}'..='\u{218F}'
                | '\u{2C00}'..='\u{2FEF}'
                | '\u{3001}'..='\u{D7FF}'
                | '\u{F900}'..='\u{FDCF}'
                | '\u{FDF0}'..='\u{FFFD}'
                | '\u{10000}'..='\u{EFFFF}'
        )
}

fn store_name_character(character: char) -> bool {
    store_name_start(character)
        || matches!(
            character,
            '-' | '.' | '0'..='9' | '\u{00B7}' | '\u{0300}'..='\u{036F}' | '\u{203F}'..='\u{2040}'
        )
}

fn control_fault(error: Error) -> Error {
    match error {
        Error::Story(inner) => Error::Other(inner.to_string()),
        other => other,
    }
}

fn plain_or_rich(kind: SdtType) -> Result<&'static str> {
    match kind {
        SdtType::RichText => Ok("richText"),
        SdtType::PlainText => Ok("text"),
        other => Err(Error::Other(format!(
            "{} content control type payload is owned by F-286",
            sdt_type_name(other)
        ))),
    }
}

fn sdt_type_name(kind: SdtType) -> &'static str {
    match kind {
        SdtType::RichText => "rich text",
        SdtType::PlainText => "plain text",
        SdtType::Picture => "picture",
        SdtType::CheckBox => "checkbox",
        SdtType::ComboBox => "combo box",
        SdtType::DropDownList => "drop-down list",
        SdtType::Date => "date",
        SdtType::DocumentPartList => "document part list",
        SdtType::DocumentPartObject => "document part object",
        SdtType::Group => "group",
        SdtType::RepeatingSection => "repeating section",
        SdtType::RepeatingSectionItem => "repeating section item",
        SdtType::Citation => "citation",
        SdtType::Equation => "equation",
        SdtType::Bibliography => "bibliography",
    }
}

fn control_shell(spec: &ContentControlSpec, id: i32, kind_name: &str) -> Result<String> {
    let body = match spec.placement {
        ContentControlPlacement::Inline => "<w:r/>".to_owned(),
        ContentControlPlacement::Row | ContentControlPlacement::Cell => String::new(),
        ContentControlPlacement::Block | ContentControlPlacement::Paragraph => {
            block_paragraph(spec.kind)?
        }
    };
    control_shell_with_body(spec, id, kind_name, &body)
}

fn control_shell_with_body(
    spec: &ContentControlSpec,
    id: i32,
    kind_name: &str,
    body: &str,
) -> Result<String> {
    let mut properties = String::new();
    if let Some(alias) = &spec.alias {
        properties.push_str(&val_element("w:alias", alias));
    }
    if let Some(tag) = &spec.tag {
        properties.push_str(&val_element("w:tag", tag));
    }
    properties.push_str(&val_element("w:id", &id.to_string()));
    if let Some(lock) = spec.lock {
        properties.push_str(&val_element("w:lock", lock.as_str()));
    }
    if spec.placeholder == Some(true) {
        properties.push_str("<w:showingPlcHdr/>");
    }
    properties.push_str(&format!("<w:{kind_name}/>"));
    if let Some(appearance) = spec.appearance {
        properties.push_str(&format!(
            "<w15:appearance w15:val=\"{}\" xmlns:w15=\"{W15_NS}\"/>",
            appearance.as_str()
        ));
    }
    Ok(format!(
        "<w:sdt xmlns:w=\"{W_NS}\"><w:sdtPr>{properties}</w:sdtPr><w:sdtContent>{body}</w:sdtContent></w:sdt>"
    ))
}

fn block_paragraph(kind: SdtType) -> Result<String> {
    match kind {
        SdtType::PlainText => Ok("<w:p><w:r><w:t></w:t></w:r></w:p>".to_owned()),
        SdtType::RichText => Ok("<w:p/>".to_owned()),
        other => Err(Error::Other(format!(
            "{} content control type payload is owned by F-286",
            sdt_type_name(other)
        ))),
    }
}

fn row_shell(columns: usize, kind: SdtType) -> Result<String> {
    let columns = columns.max(1);
    let cell = cell_shell(kind)?;
    let mut row = String::from("<w:tr>");
    for _ in 0..columns {
        row.push_str(&cell);
    }
    row.push_str("</w:tr>");
    Ok(row)
}

fn cell_shell(kind: SdtType) -> Result<String> {
    Ok(format!(
        "<w:tc><w:tcPr><w:tcW w:w=\"0\" w:type=\"auto\"/></w:tcPr>{}</w:tc>",
        block_paragraph(kind)?
    ))
}

fn val_element(name: &str, value: &str) -> String {
    format!("<{name} w:val=\"{}\"/>", xml_attr(value))
}

fn xml_attr(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('"', "&quot;")
}

fn parse_control(xml: &str) -> Result<CT_Sdt> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(start)) if start.local_name().as_ref() == b"sdt" => {
                return CT_Sdt::from_xml(&mut reader, &start).map_err(Error::from);
            }
            Ok(Event::Eof) => {
                return Err(Error::Other(
                    "content control shell did not parse".to_owned(),
                ));
            }
            Ok(_) => buffer.clear(),
            Err(error) => {
                return Err(Error::Other(format!(
                    "content control shell did not parse: {error}"
                )));
            }
        }
    }
}

fn shell_contains_paragraph(xml: &str) -> Result<bool> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut in_content = false;
    let mut depth = 0usize;
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(start)) => {
                let local = start.local_name();
                if in_content && depth == 0 && local.as_ref() == b"p" {
                    return Ok(true);
                }
                if local.as_ref() == b"sdtContent" {
                    in_content = true;
                    depth = 0;
                } else if in_content {
                    depth += 1;
                }
                buffer.clear();
            }
            Ok(Event::Empty(start)) => {
                let local = start.local_name();
                if in_content && depth == 0 && local.as_ref() == b"p" {
                    return Ok(true);
                }
                buffer.clear();
            }
            Ok(Event::End(end)) => {
                if end.local_name().as_ref() == b"sdtContent" {
                    return Ok(false);
                }
                if in_content && depth > 0 {
                    depth -= 1;
                }
                buffer.clear();
            }
            Ok(Event::Eof) => return Ok(false),
            Ok(_) => buffer.clear(),
            Err(error) => {
                return Err(Error::Other(format!(
                    "content control shell did not parse: {error}"
                )));
            }
        }
    }
}

fn require_block_parent(document: &Document, destination: &ContentLocation) -> Result<()> {
    let kind = destination.story().kind();
    if !matches!(
        kind,
        StoryKind::Body
            | StoryKind::TableCell
            | StoryKind::Header
            | StoryKind::Footer
            | StoryKind::Footnote
            | StoryKind::Endnote
            | StoryKind::Comment
            | StoryKind::TextBox
    ) {
        return Err(Error::Other(
            "content control placement does not match its parent".to_owned(),
        ));
    }
    if Document::location_is_end(destination) {
        return Ok(());
    }
    let (_, _, _, direct) = document
        .copy_indexed_story_item(destination, destination.index_path())
        .map_err(control_fault)?;
    if direct {
        Ok(())
    } else {
        Err(Error::Other(
            "content control placement does not match its parent".to_owned(),
        ))
    }
}

struct PartRange {
    part: String,
    xml: Vec<u8>,
    range: Range<usize>,
}

fn table_parent_range(document: &Document, destination: &ContentLocation) -> Result<PartRange> {
    if Document::location_is_end(destination)
        || destination.item_kind() != StoryItemKind::Table
        || destination.index_path().len() != 1
    {
        return Err(Error::Other(
            "content control placement does not match its parent".to_owned(),
        ));
    }
    let (part, xml, range, _) = document
        .copy_indexed_story_item(destination, destination.index_path())
        .map_err(control_fault)?;
    Ok(PartRange { part, xml, range })
}

fn row_parent_range(document: &Document, destination: &ContentLocation) -> Result<PartRange> {
    let path = destination.index_path();
    let [table_index, row_index] = path else {
        return Err(Error::Other(
            "content control placement does not match its parent".to_owned(),
        ));
    };
    if destination.item_kind() != StoryItemKind::Table {
        return Err(Error::Other(
            "content control placement does not match its parent".to_owned(),
        ));
    }
    let (part, xml, table_range, _) = document
        .copy_indexed_story_item(destination, &[*table_index])
        .map_err(control_fault)?;
    let range = nth_direct_child(&xml, &table_range, b"tr", *row_index)?;
    Ok(PartRange { part, xml, range })
}

fn paragraph_parent_range(document: &Document, destination: &ContentLocation) -> Result<PartRange> {
    if Document::location_is_end(destination) || destination.item_kind() != StoryItemKind::Paragraph
    {
        return Err(Error::Other(
            "content control placement does not match its parent".to_owned(),
        ));
    }
    let (part, xml, range, _) = document
        .copy_indexed_story_item(destination, destination.index_path())
        .map_err(control_fault)?;
    Ok(PartRange { part, xml, range })
}

fn splice_inside(document: &mut Document, located: PartRange, payload: &str) -> Result<()> {
    let PartRange {
        part,
        mut xml,
        range,
    } = located;
    let replacement = insert_into_element(&xml, &range, payload.as_bytes())?;
    xml.splice(range, replacement);
    document.write_story_part(&part, xml)
}

fn control_item(
    document: &Document,
    location: &ContentLocation,
) -> Result<(String, Vec<u8>, Range<usize>, bool)> {
    document
        .copy_indexed_story_item(location, location.index_path())
        .map_err(control_fault)
}

struct XmlElem {
    id: usize,
    parent: Option<usize>,
    local: Vec<u8>,
    word: bool,
    w14: bool,
    w15: bool,
    range: Range<usize>,
    open_end: usize,
    inner_end: Option<usize>,
    val: Option<String>,
}

fn index_elements(xml: &[u8]) -> Result<Vec<XmlElem>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut stack: Vec<(usize, usize, usize)> = Vec::new();
    let mut elems = Vec::new();
    let mut next_id = 0usize;
    loop {
        let before = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("content control XML is invalid: {error}")))?;
        let word = is_word(&namespace);
        let w14 = is_w14(&namespace);
        let w15 = is_w15(&namespace);
        drop(namespace);
        let after = reader.buffer_position() as usize;
        match event {
            Event::Start(element) => {
                let id = next_id;
                next_id += 1;
                let parent = stack.last().map(|frame| frame.0);
                let val = word_val(&reader, &element)?;
                stack.push((id, before, after));
                elems.push(XmlElem {
                    id,
                    parent,
                    local: element.local_name().as_ref().to_vec(),
                    word,
                    w14,
                    w15,
                    range: before..before,
                    open_end: after,
                    inner_end: None,
                    val,
                });
            }
            Event::Empty(element) => {
                let id = next_id;
                next_id += 1;
                elems.push(XmlElem {
                    id,
                    parent: stack.last().map(|frame| frame.0),
                    local: element.local_name().as_ref().to_vec(),
                    word,
                    w14,
                    w15,
                    range: before..after,
                    open_end: after,
                    inner_end: None,
                    val: word_val(&reader, &element)?,
                });
            }
            Event::End(_) => {
                let Some((id, start, open_end)) = stack.pop() else {
                    return Err(Error::Other(
                        "content control XML has an unmatched end element".to_owned(),
                    ));
                };
                // Ids number the elements in the order they are pushed.
                let Some(elem) = elems.get_mut(id) else {
                    return Err(Error::Other(
                        "content control XML lost an open element".to_owned(),
                    ));
                };
                elem.range = start..after;
                elem.open_end = open_end;
                elem.inner_end = Some(before);
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    if !stack.is_empty() {
        return Err(Error::Other(
            "content control XML ended inside an element".to_owned(),
        ));
    }
    Ok(elems)
}

fn is_word(namespace: &ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(uri)) if *uri == W_NS.as_bytes())
}

fn is_w14(namespace: &ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(uri)) if *uri == W14_NS.as_bytes())
}

fn is_w15(namespace: &ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(uri)) if *uri == W15_NS.as_bytes())
}

fn word_val(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<Option<String>> {
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| {
            Error::Other(format!("content control attribute is invalid: {error}"))
        })?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if is_word(&namespace) && local.as_ref() == b"val" {
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map_err(|error| {
                    Error::Other(format!(
                        "content control attribute value is invalid: {error}"
                    ))
                })?;
            return Ok(Some(value.into_owned()));
        }
    }
    Ok(None)
}

fn element_at<'a>(elems: &'a [XmlElem], range: &Range<usize>) -> Result<&'a XmlElem> {
    elems
        .iter()
        .find(|elem| elem.range.start == range.start)
        .ok_or_else(|| Error::Other("content control bytes are missing".to_owned()))
}

fn nth_direct_child(
    xml: &[u8],
    parent: &Range<usize>,
    local: &[u8],
    index: usize,
) -> Result<Range<usize>> {
    let elems = index_elements(xml)?;
    let parent = element_at(&elems, parent)?;
    elems
        .iter()
        .filter(|elem| elem.parent == Some(parent.id) && elem.word && elem.local == local)
        .nth(index)
        .map(|elem| elem.range.clone())
        .ok_or_else(|| {
            Error::Other("content control placement does not match its parent".to_owned())
        })
}

fn table_column_count(xml: &[u8], table_range: &Range<usize>) -> Result<usize> {
    let elems = index_elements(xml)?;
    let table = element_at(&elems, table_range)?;
    if let Some(row) = elems
        .iter()
        .find(|elem| elem.parent == Some(table.id) && elem.word && elem.local == b"tr")
    {
        let count = elems
            .iter()
            .filter(|elem| elem.parent == Some(row.id) && elem.word && elem.local == b"tc")
            .count();
        return Ok(count.max(1));
    }
    if let Some(grid) = elems
        .iter()
        .find(|elem| elem.parent == Some(table.id) && elem.word && elem.local == b"tblGrid")
    {
        let count = elems
            .iter()
            .filter(|elem| elem.parent == Some(grid.id) && elem.word && elem.local == b"gridCol")
            .count();
        if count > 0 {
            return Ok(count);
        }
    }
    Ok(1)
}

fn insert_into_element(xml: &[u8], range: &Range<usize>, payload: &[u8]) -> Result<Vec<u8>> {
    let elems = index_elements(xml)?;
    let elem = element_at(&elems, range)?;
    if elem.inner_end.is_none() {
        return expand_empty(&xml[elem.range.clone()], payload);
    }
    let inner_end = elem.inner_end.unwrap_or(elem.range.end);
    let mut updated = xml[elem.range.start..inner_end].to_vec();
    updated.extend_from_slice(payload);
    updated.extend_from_slice(&xml[inner_end..elem.range.end]);
    Ok(updated)
}

fn expand_empty(element: &[u8], payload: &[u8]) -> Result<Vec<u8>> {
    let qname = element_qname(element)?;
    let slash = element
        .windows(2)
        .rposition(|window| window == b"/>")
        .ok_or_else(|| Error::Other("content control empty element is not closed".to_owned()))?;
    let mut updated = element[..slash].to_vec();
    while updated.last().is_some_and(u8::is_ascii_whitespace) {
        updated.pop();
    }
    updated.push(b'>');
    updated.extend_from_slice(payload);
    updated.extend_from_slice(b"</");
    updated.extend_from_slice(qname.as_bytes());
    updated.push(b'>');
    Ok(updated)
}

fn element_qname(bytes: &[u8]) -> Result<&str> {
    let rest = bytes
        .strip_prefix(b"<")
        .ok_or_else(|| Error::Other("content control element is not a start tag".to_owned()))?;
    let end = rest
        .iter()
        .position(|byte| byte.is_ascii_whitespace() || *byte == b'/' || *byte == b'>')
        .ok_or_else(|| Error::Other("content control element name is missing".to_owned()))?;
    std::str::from_utf8(&rest[..end])
        .map_err(|error| Error::Other(format!("content control element name is invalid: {error}")))
}

fn collect_control_ids(xml: &[u8], used: &mut HashSet<i32>) -> Result<()> {
    let elems = index_elements(xml)?;
    for elem in &elems {
        if !(elem.word && elem.local == b"id") {
            continue;
        }
        let Some(parent) = elem.parent.and_then(|id| elems.get(id)) else {
            continue;
        };
        if !(parent.word && parent.local == b"sdtPr") {
            continue;
        }
        if let Some(value) = &elem.val
            && let Ok(id) = value.parse::<i32>()
            && id > 0
        {
            used.insert(id);
        }
    }
    Ok(())
}

fn edit_sdt_property(xml: &[u8], sdt_range: &Range<usize>, edit: PropertyEdit) -> Result<Vec<u8>> {
    let elems = index_elements(xml)?;
    let sdt = element_at(&elems, sdt_range)?;
    let sdt_pr = elems
        .iter()
        .find(|elem| elem.parent == Some(sdt.id) && elem.word && elem.local == b"sdtPr")
        .ok_or_else(|| Error::Other("content control has no properties".to_owned()))?;
    if let Some(target) = elems
        .iter()
        .find(|elem| elem.parent == Some(sdt_pr.id) && property_matches(elem, &edit))
    {
        let replacement = property_bytes(xml, target, &edit)?;
        return splice_absolute(xml, sdt_range, &target.range, &replacement);
    }
    if property_clears(&edit) {
        return Ok(xml[sdt_range.clone()].to_vec());
    }
    let markup = property_bytes_new(xml, sdt_pr, &edit)?;
    if sdt_pr.inner_end.is_none() {
        let expanded = expand_empty(&xml[sdt_pr.range.clone()], &markup)?;
        return splice_absolute(xml, sdt_range, &sdt_pr.range, &expanded);
    }
    let rank = property_insert_rank(&edit);
    let insert_at = elems
        .iter()
        .filter(|elem| elem.parent == Some(sdt_pr.id))
        .find(|elem| property_rank(elem).is_some_and(|child| child > rank))
        .map(|elem| elem.range.start)
        .unwrap_or(sdt_pr.inner_end.unwrap_or(sdt_pr.range.end));
    splice_absolute(xml, sdt_range, &(insert_at..insert_at), &markup)
}

fn splice_absolute(
    xml: &[u8],
    outer: &Range<usize>,
    target: &Range<usize>,
    replacement: &[u8],
) -> Result<Vec<u8>> {
    if target.start < outer.start || target.end > outer.end {
        return Err(Error::Other(
            "content control property is outside its shell".to_owned(),
        ));
    }
    let mut updated = xml[outer.clone()].to_vec();
    let local = (target.start - outer.start)..(target.end - outer.start);
    updated.splice(local, replacement.iter().copied());
    Ok(updated)
}

fn property_matches(elem: &XmlElem, edit: &PropertyEdit) -> bool {
    match edit {
        PropertyEdit::Tag(_) => elem.word && elem.local == b"tag",
        PropertyEdit::Alias(_) => elem.word && elem.local == b"alias",
        PropertyEdit::Lock(_) => elem.word && elem.local == b"lock",
        PropertyEdit::Appearance(_) => elem.w15 && elem.local == b"appearance",
        PropertyEdit::Placeholder(_) => elem.word && elem.local == b"showingPlcHdr",
        PropertyEdit::DataBinding { .. } => elem.word && elem.local == b"dataBinding",
    }
}

fn property_clears(edit: &PropertyEdit) -> bool {
    matches!(
        edit,
        PropertyEdit::Tag(None)
            | PropertyEdit::Alias(None)
            | PropertyEdit::Lock(None)
            | PropertyEdit::Appearance(None)
            | PropertyEdit::Placeholder(None)
            | PropertyEdit::Placeholder(Some(false))
    )
}

fn property_insert_rank(edit: &PropertyEdit) -> u8 {
    match edit {
        PropertyEdit::Alias(_) => 1,
        PropertyEdit::Tag(_) => 2,
        PropertyEdit::Lock(_) => 4,
        PropertyEdit::Placeholder(_) => 7,
        PropertyEdit::DataBinding { .. } => 8,
        PropertyEdit::Appearance(_) => 21,
    }
}

fn property_rank(elem: &XmlElem) -> Option<u8> {
    if elem.word {
        return match elem.local.as_slice() {
            b"rPr" => Some(0),
            b"alias" => Some(1),
            b"tag" => Some(2),
            b"id" => Some(3),
            b"lock" => Some(4),
            b"placeholder" => Some(5),
            b"temporary" => Some(6),
            b"showingPlcHdr" => Some(7),
            b"dataBinding" => Some(8),
            b"label" | b"tabIndex" => Some(9),
            b"richText" | b"text" | b"date" | b"picture" | b"comboBox" | b"dropDownList"
            | b"citation" | b"group" | b"bibliography" | b"equation" | b"docPartList"
            | b"docPartObj" => Some(20),
            _ => None,
        };
    }
    if elem.w15 {
        return match elem.local.as_slice() {
            b"appearance" => Some(21),
            b"repeatingSection" | b"repeatingSectionItem" => Some(20),
            _ => None,
        };
    }
    if elem.local.as_slice() == b"checkbox" {
        return Some(20);
    }
    None
}

fn property_bytes(xml: &[u8], target: &XmlElem, edit: &PropertyEdit) -> Result<Vec<u8>> {
    if property_clears(edit) {
        return Ok(Vec::new());
    }
    if matches!(edit, PropertyEdit::Appearance(_)) {
        return Ok(appearance_markup(edit)?.into_bytes());
    }
    let prefix = name_prefix(element_qname(&xml[target.range.clone()])?);
    Ok(property_markup(prefix, edit)?.into_bytes())
}

fn property_bytes_new(xml: &[u8], sdt_pr: &XmlElem, edit: &PropertyEdit) -> Result<Vec<u8>> {
    if matches!(edit, PropertyEdit::Appearance(_)) {
        return Ok(appearance_markup(edit)?.into_bytes());
    }
    let prefix = name_prefix(element_qname(&xml[sdt_pr.range.clone()])?);
    Ok(property_markup(prefix, edit)?.into_bytes())
}

fn name_prefix(qname: &str) -> &str {
    match qname.split_once(':') {
        Some((prefix, _)) => prefix,
        None => "",
    }
}

fn property_markup(prefix: &str, edit: &PropertyEdit) -> Result<String> {
    match edit {
        PropertyEdit::Tag(Some(value)) => Ok(prefixed_val(prefix, "tag", value)),
        PropertyEdit::Alias(Some(value)) => Ok(prefixed_val(prefix, "alias", value)),
        PropertyEdit::Lock(Some(lock)) => Ok(prefixed_val(prefix, "lock", lock.as_str())),
        PropertyEdit::Placeholder(Some(true)) => Ok(prefixed_marker(prefix, "showingPlcHdr")),
        PropertyEdit::Appearance(_) => appearance_markup(edit),
        PropertyEdit::DataBinding {
            prefix_mappings,
            xpath,
            store_item_id,
        } => Ok(data_binding_markup(
            prefix,
            prefix_mappings.as_deref(),
            xpath,
            store_item_id,
        )),
        _ => Err(Error::Other(
            "content control property value is missing".to_owned(),
        )),
    }
}

fn data_binding_markup(
    prefix: &str,
    mappings: Option<&str>,
    xpath: &str,
    store_item_id: &str,
) -> String {
    let qualify = |local: &str| {
        if prefix.is_empty() {
            local.to_owned()
        } else {
            format!("{prefix}:{local}")
        }
    };
    let mut attrs = String::new();
    if let Some(mappings) = mappings.filter(|value| !value.is_empty()) {
        attrs.push_str(&format!(
            " {}=\"{}\"",
            qualify("prefixMappings"),
            xml_attr(mappings)
        ));
    }
    attrs.push_str(&format!(" {}=\"{}\"", qualify("xpath"), xml_attr(xpath)));
    attrs.push_str(&format!(
        " {}=\"{}\"",
        qualify("storeItemID"),
        xml_attr(store_item_id)
    ));
    format!("<{}{attrs}/>", qualify("dataBinding"))
}

fn appearance_markup(edit: &PropertyEdit) -> Result<String> {
    let PropertyEdit::Appearance(Some(appearance)) = edit else {
        return Err(Error::Other(
            "content control appearance value is missing".to_owned(),
        ));
    };
    Ok(format!(
        "<w15:appearance w15:val=\"{}\" xmlns:w15=\"{W15_NS}\"/>",
        xml_attr(appearance.as_str())
    ))
}

fn prefixed_val(prefix: &str, local: &str, value: &str) -> String {
    if prefix.is_empty() {
        format!("<{local} val=\"{}\"/>", xml_attr(value))
    } else {
        format!("<{prefix}:{local} {prefix}:val=\"{}\"/>", xml_attr(value))
    }
}

fn prefixed_marker(prefix: &str, local: &str) -> String {
    if prefix.is_empty() {
        format!("<{local}/>")
    } else {
        format!("<{prefix}:{local}/>")
    }
}

fn sdt_content_bytes(xml: &[u8], sdt_range: &Range<usize>) -> Result<Vec<u8>> {
    let elems = index_elements(xml)?;
    let sdt = element_at(&elems, sdt_range)?;
    let Some(content) = elems
        .iter()
        .find(|elem| elem.parent == Some(sdt.id) && elem.word && elem.local == b"sdtContent")
    else {
        return Ok(Vec::new());
    };
    let Some(inner_end) = content.inner_end else {
        return Ok(Vec::new());
    };
    Ok(xml[content.open_end..inner_end].to_vec())
}

#[derive(Clone, Copy)]
enum ParentSlot {
    Block,
    Table,
    Row,
    Inline,
    Other,
}

fn parent_can_hold(xml: &[u8], sdt_range: &Range<usize>, inner: &[u8]) -> Result<bool> {
    let elems = index_elements(xml)?;
    let sdt = element_at(&elems, sdt_range)?;
    inner_matches_slot(parent_slot(&elems, sdt.id), inner)
}

fn parent_slot(elems: &[XmlElem], id: usize) -> ParentSlot {
    let Some(elem) = elems.iter().find(|elem| elem.id == id) else {
        return ParentSlot::Other;
    };
    let Some(parent_id) = elem.parent else {
        return ParentSlot::Other;
    };
    let Some(parent) = elems.iter().find(|elem| elem.id == parent_id) else {
        return ParentSlot::Other;
    };
    element_slot(elems, parent)
}

fn control_parent_is_block(xml: &[u8], range: &Range<usize>) -> Result<bool> {
    let elems = index_elements(xml)?;
    let sdt = element_at(&elems, range)?;
    Ok(matches!(parent_slot(&elems, sdt.id), ParentSlot::Block))
}

fn wrap_word_root(inner: &[u8]) -> Vec<u8> {
    let mut wrapped = Vec::with_capacity(inner.len() + W_NS.len() + 32);
    wrapped.extend_from_slice(b"<w:root xmlns:w=\"");
    wrapped.extend_from_slice(W_NS.as_bytes());
    wrapped.extend_from_slice(b"\">");
    wrapped.extend_from_slice(inner);
    wrapped.extend_from_slice(b"</w:root>");
    wrapped
}

fn inner_matches_slot(slot: ParentSlot, inner: &[u8]) -> Result<bool> {
    if matches!(slot, ParentSlot::Other) {
        return Ok(false);
    }
    if payload_has_nonspace_text(inner)? {
        return Ok(false);
    }
    let wrapped = wrap_word_root(inner);
    let elems = index_elements(&wrapped)?;
    let Some(root) = elems.iter().find(|elem| elem.parent.is_none()) else {
        return Ok(true);
    };
    transparent_children_match(slot, &wrapped, &elems, root.id)
}

fn transparent_children_match(
    slot: ParentSlot,
    xml: &[u8],
    elems: &[XmlElem],
    parent_id: usize,
) -> Result<bool> {
    let custom_parent = elems
        .iter()
        .any(|elem| elem.id == parent_id && elem.word && elem.local == b"customXml");
    for child in elems.iter().filter(|elem| elem.parent == Some(parent_id)) {
        if custom_parent && child.word && child.local == b"customXmlPr" {
            continue;
        }
        if child.word && child.local == b"customXml" {
            if let Some(end) = child.inner_end
                && (payload_has_nonspace_text(&xml[child.open_end..end])?
                    || !transparent_children_match(slot, xml, elems, child.id)?)
            {
                return Ok(false);
            }
        } else if child.word && child.local == b"sdt" {
            if !sdt_payload_matches(slot, xml, elems, child)? {
                return Ok(false);
            }
        } else if !child_allowed(slot, child.word, &child.local) {
            return Ok(false);
        }
    }
    Ok(true)
}

fn sdt_payload_matches(
    slot: ParentSlot,
    xml: &[u8],
    elems: &[XmlElem],
    sdt: &XmlElem,
) -> Result<bool> {
    if (sdt_is_repeating_section(elems, sdt) || sdt_is_repeating_item(elems, sdt))
        && !repeating_slot_allowed(slot)
    {
        return Ok(false);
    }
    let mut contents = Vec::new();
    for child in elems.iter().filter(|elem| elem.parent == Some(sdt.id)) {
        match (child.word, child.local.as_slice()) {
            (true, b"sdtPr" | b"sdtEndPr") => {}
            (true, b"sdtContent") => contents.push(child),
            _ => return Ok(false),
        }
    }
    for content in contents {
        let Some(inner_end) = content.inner_end else {
            continue;
        };
        let inner = &xml[content.open_end..inner_end];
        if payload_has_nonspace_text(inner)?
            || !transparent_children_match(slot, xml, elems, content.id)?
        {
            return Ok(false);
        }
    }
    Ok(true)
}

fn payload_has_nonspace_text(inner: &[u8]) -> Result<bool> {
    let wrapped = wrap_word_root(inner);
    let mut reader = Reader::from_reader(wrapped.as_slice());
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(_)) => depth += 1,
            Ok(Event::End(_)) => {
                if depth == 0 {
                    return Err(Error::Other(
                        "content control unwrap content is invalid".to_owned(),
                    ));
                }
                depth -= 1;
            }
            Ok(Event::Text(text)) if depth == 1 && !xml_space(text.as_ref()) => return Ok(true),
            Ok(Event::CData(text)) if depth == 1 && !xml_space(text.as_ref()) => return Ok(true),
            Ok(Event::GeneralRef(reference)) if depth == 1 => {
                let character = decode_one_xml_reference(&reference).map_err(|detail| {
                    Error::Other(format!(
                        "content control unwrap content is invalid: {detail}"
                    ))
                })?;
                if !matches!(character, ' ' | '\t' | '\n' | '\r') {
                    return Ok(true);
                }
            }
            Ok(Event::Eof) => return Ok(false),
            Ok(_) => {}
            Err(error) => {
                return Err(Error::Other(format!(
                    "content control unwrap content is invalid: {error}"
                )));
            }
        }
        buffer.clear();
    }
}

fn child_allowed(slot: ParentSlot, word: bool, local: &[u8]) -> bool {
    if !word {
        return false;
    }
    match slot {
        ParentSlot::Block => matches!(
            local,
            b"p" | b"tbl"
                | b"sdt"
                | b"bookmarkStart"
                | b"bookmarkEnd"
                | b"proofErr"
                | b"commentRangeStart"
                | b"commentRangeEnd"
                | b"permStart"
                | b"permEnd"
                | b"customXml"
        ),
        ParentSlot::Table => {
            matches!(
                local,
                b"tr" | b"sdt" | b"bookmarkStart" | b"bookmarkEnd" | b"proofErr" | b"customXml"
            )
        }
        ParentSlot::Row => {
            matches!(
                local,
                b"tc" | b"sdt" | b"bookmarkStart" | b"bookmarkEnd" | b"proofErr" | b"customXml"
            )
        }
        ParentSlot::Inline => matches!(
            local,
            b"r" | b"sdt"
                | b"hyperlink"
                | b"fldSimple"
                | b"ins"
                | b"del"
                | b"moveFrom"
                | b"moveTo"
                | b"bookmarkStart"
                | b"bookmarkEnd"
                | b"commentRangeStart"
                | b"commentRangeEnd"
                | b"proofErr"
                | b"customXml"
        ),
        ParentSlot::Other => false,
    }
}

fn xml_space(bytes: &[u8]) -> bool {
    bytes
        .iter()
        .all(|byte| matches!(byte, b' ' | b'\t' | b'\n' | b'\r'))
}

fn destination_accepts_control(
    parent_xml: &[u8],
    parent_range: &Range<usize>,
    source_xml: &[u8],
    source_range: &Range<usize>,
) -> Result<()> {
    let elems = index_elements(parent_xml)?;
    let parent = element_at(&elems, parent_range)?;
    let in_section = ancestor_repeating_section(&elems, parent.id).is_some();
    if control_matches_context(
        element_slot(&elems, parent),
        in_section,
        source_xml,
        source_range,
    )? {
        Ok(())
    } else {
        Err(Error::Other(
            "content control placement does not match its parent".to_owned(),
        ))
    }
}

fn element_slot(elems: &[XmlElem], elem: &XmlElem) -> ParentSlot {
    if !elem.word {
        return ParentSlot::Other;
    }
    match elem.local.as_slice() {
        b"sdt" | b"sdtContent" | b"customXml" => parent_slot(elems, elem.id),
        b"p" | b"hyperlink" | b"fldSimple" | b"ins" | b"del" | b"moveFrom" | b"moveTo" => {
            ParentSlot::Inline
        }
        b"tbl" => ParentSlot::Table,
        b"tr" => ParentSlot::Row,
        b"body" | b"hdr" | b"ftr" | b"footnote" | b"endnote" | b"comment" | b"txbxContent"
        | b"tc" => ParentSlot::Block,
        _ => ParentSlot::Other,
    }
}

fn nested_destination(document: &Document, destination: &ContentLocation) -> Result<PartRange> {
    if !Document::location_is_end(destination)
        && destination.item_kind() == StoryItemKind::Paragraph
    {
        return paragraph_parent_range(document, destination);
    }
    if !Document::location_is_end(destination) && destination.item_kind() == StoryItemKind::Table {
        return match destination.index_path().len() {
            1 => table_parent_range(document, destination),
            2 => row_parent_range(document, destination),
            _ => Err(Error::Other(
                "content control placement does not match its parent".to_owned(),
            )),
        };
    }
    Err(Error::Other(
        "content control placement does not match its parent".to_owned(),
    ))
}

fn ranges_overlap(left: &Range<usize>, right: &Range<usize>) -> bool {
    left.start < right.end && right.start < left.end
}

fn shift_after_removal(dest: &Range<usize>, removed: &Range<usize>) -> Range<usize> {
    if dest.start >= removed.end {
        let shift = removed.len();
        dest.start - shift..dest.end - shift
    } else {
        dest.clone()
    }
}

fn ensure_cell_paragraph(
    document: &Document,
    story: &crate::StoryId,
    edited: &[u8],
    original_len: usize,
) -> Result<()> {
    if story.kind() != StoryKind::TableCell {
        return Ok(());
    }
    let delta = edited.len() as isize - original_len as isize;
    let (_xml, span) = document.story_owner_span(story)?;
    let end = usize::try_from(span.end as isize + delta)
        .map_err(|_| Error::Other("content control cell span is invalid".to_owned()))?;
    if span.start > end || end > edited.len() {
        return Err(Error::Other(
            "content control cell span is invalid".to_owned(),
        ));
    }
    if !word_paragraph_in_span(edited, span.start..end)? {
        return Err(Error::Other(
            "cannot remove the final required table-cell paragraph".to_owned(),
        ));
    }
    Ok(())
}

fn refuse_cell_without_paragraph(
    document: &Document,
    story: &crate::StoryId,
    xml: &[u8],
    range: &Range<usize>,
    replacement: &[u8],
) -> Result<()> {
    if story.kind() != StoryKind::TableCell {
        return Ok(());
    }
    let mut edited = xml.to_vec();
    edited.splice(range.clone(), replacement.iter().copied());
    ensure_cell_paragraph(document, story, &edited, xml.len())
}

fn word_paragraph_in_span(xml: &[u8], span: Range<usize>) -> Result<bool> {
    let elems = index_elements(xml)?;
    Ok(elems.iter().any(|elem| {
        elem.word
            && elem.local == b"p"
            && elem.range.start >= span.start
            && elem.range.end <= span.end
    }))
}

fn validate_type_payload(payload: &ContentControlTypePayload) -> Result<()> {
    match payload {
        ContentControlTypePayload::Date { format } => {
            if format != "yyyy-MM-dd" && format != "dd/MM/yyyy" {
                return Err(Error::Other(
                    "content control date format must be yyyy-MM-dd or dd/MM/yyyy".to_owned(),
                ));
            }
        }
        ContentControlTypePayload::ComboBox { items }
        | ContentControlTypePayload::DropDownList { items } => validate_list_items(items)?,
        ContentControlTypePayload::Picture { relationship_id } => {
            if relationship_id.is_empty() {
                return Err(Error::Other(
                    "content control picture relationship is missing".to_owned(),
                ));
            }
            oxml_core::xml::reject_non_xml_characters(
                "content control picture relationship",
                relationship_id,
            )?;
        }
        _ => {}
    }
    Ok(())
}

fn validate_list_items(items: &[(String, String)]) -> Result<()> {
    if items.is_empty() {
        return Err(Error::Other("content control list is empty".to_owned()));
    }
    let mut seen = HashSet::new();
    for (display, value) in items {
        oxml_core::xml::reject_non_xml_characters("content control list display", display)?;
        oxml_core::xml::reject_non_xml_characters("content control list value", value)?;
        if !seen.insert(value.clone()) {
            return Err(Error::Other(
                "content control list value is repeated".to_owned(),
            ));
        }
    }
    Ok(())
}

fn type_markup(payload: &ContentControlTypePayload) -> Result<String> {
    match payload {
        ContentControlTypePayload::RichText => Ok("<w:richText/>".to_owned()),
        ContentControlTypePayload::PlainText => Ok("<w:text/>".to_owned()),
        ContentControlTypePayload::Date { format } => Ok(format!(
            "<w:date><w:dateFormat w:val=\"{}\"/><w:storeMappedDataAs w:val=\"text\"/><w:calendar w:val=\"gregorian\"/></w:date>",
            xml_attr(format)
        )),
        ContentControlTypePayload::CheckBox => Ok(format!(
            "<w14:checkbox xmlns:w14=\"{W14_NS}\"><w14:checked w14:val=\"0\"/><w14:checkedState w14:val=\"2612\" w14:font=\"MS Gothic\"/><w14:uncheckedState w14:val=\"2610\" w14:font=\"MS Gothic\"/></w14:checkbox>"
        )),
        ContentControlTypePayload::ComboBox { items } => Ok(list_markup("w:comboBox", items)),
        ContentControlTypePayload::DropDownList { items } => {
            Ok(list_markup("w:dropDownList", items))
        }
        ContentControlTypePayload::Picture { .. } => Ok("<w:picture/>".to_owned()),
        ContentControlTypePayload::Group => Ok("<w:group/>".to_owned()),
        ContentControlTypePayload::RepeatingSection => {
            Ok(format!("<w15:repeatingSection xmlns:w15=\"{W15_NS}\"/>"))
        }
        ContentControlTypePayload::RepeatingSectionItem => Ok(format!(
            "<w15:repeatingSectionItem xmlns:w15=\"{W15_NS}\"/>"
        )),
    }
}

fn list_markup(name: &str, items: &[(String, String)]) -> String {
    let mut xml = format!("<{name}>");
    for (display, value) in items {
        xml.push_str(&format!(
            "<w:listItem w:displayText=\"{}\" w:value=\"{}\"/>",
            xml_attr(display),
            xml_attr(value)
        ));
    }
    xml.push_str("</");
    xml.push_str(name);
    xml.push('>');
    xml
}

fn part_owns_image(document: &Document, part: &str, id: &str) -> bool {
    document.package.get_part_rels(part).is_some_and(|rels| {
        rels.items
            .iter()
            .any(|rel| rel.id == id && rel.rel_type == oxml_opc::relationship::rel_types::IMAGE)
    })
}

fn content_embeds_relationship(xml: &[u8], sdt_range: &Range<usize>, rel_id: &str) -> Result<bool> {
    let elems = index_elements(xml)?;
    let sdt = element_at(&elems, sdt_range)?;
    let Some(content) = sdt_content_elem(&elems, sdt) else {
        return Ok(false);
    };
    let Some(inner_end) = content.inner_end else {
        return Ok(false);
    };
    drawing_embeds(xml, content.open_end..inner_end, rel_id)
}

fn drawing_embeds(xml: &[u8], inner: Range<usize>, rel_id: &str) -> Result<bool> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut drawing_depth = 0i32;
    loop {
        let before = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| {
                Error::Other(format!(
                    "content control picture content is invalid: {error}"
                ))
            })?;
        let word = is_word(&namespace);
        drop(namespace);
        let after = reader.buffer_position() as usize;
        let in_content = before >= inner.start && after <= inner.end;
        match event {
            Event::Start(element) => {
                let drawing = word && element.local_name().as_ref() == b"drawing";
                if in_content && drawing {
                    drawing_depth += 1;
                }
                if in_content && drawing_depth > 0 && embed_matches(&reader, &element, rel_id)? {
                    return Ok(true);
                }
            }
            Event::Empty(element) => {
                let drawing = word && element.local_name().as_ref() == b"drawing";
                if in_content
                    && (drawing_depth > 0 || drawing)
                    && embed_matches(&reader, &element, rel_id)?
                {
                    return Ok(true);
                }
            }
            Event::End(element) => {
                if in_content
                    && word
                    && element.local_name().as_ref() == b"drawing"
                    && drawing_depth > 0
                {
                    drawing_depth -= 1;
                }
            }
            Event::Eof => return Ok(false),
            _ => {}
        }
        buffer.clear();
    }
}

fn embed_matches(reader: &NsReader<&[u8]>, element: &BytesStart<'_>, rel_id: &str) -> Result<bool> {
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| {
            Error::Other(format!(
                "content control picture content is invalid: {error}"
            ))
        })?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let relationships =
            matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == REL_NS.as_bytes());
        if relationships && local.as_ref() == b"embed" {
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map_err(|error| {
                    Error::Other(format!(
                        "content control picture content is invalid: {error}"
                    ))
                })?;
            if value.as_ref() == rel_id {
                return Ok(true);
            }
        }
    }
    Ok(false)
}

enum RepeatingContent {
    Empty,
    NotBlock,
    AlreadyItem,
    Wrap,
}

fn repeating_content(xml: &[u8], sdt_range: &Range<usize>) -> Result<RepeatingContent> {
    let elems = index_elements(xml)?;
    let sdt = element_at(&elems, sdt_range)?;
    let Some(content) = sdt_content_elem(&elems, sdt) else {
        return Ok(RepeatingContent::Empty);
    };
    let Some(inner_end) = content.inner_end else {
        return Ok(RepeatingContent::Empty);
    };
    let children: Vec<&XmlElem> = elems
        .iter()
        .filter(|elem| elem.parent == Some(content.id))
        .collect();
    let inner = &xml[content.open_end..inner_end];
    if inner_has_nonspace_text(inner)? {
        return Ok(RepeatingContent::NotBlock);
    }
    if children.is_empty() {
        return Ok(RepeatingContent::Empty);
    }
    if children.len() == 1
        && children[0].word
        && children[0].local == b"sdt"
        && sdt_is_repeating_item(&elems, children[0])
    {
        return Ok(RepeatingContent::AlreadyItem);
    }
    if !transparent_children_match(ParentSlot::Block, xml, &elems, content.id)
        .map_err(|error| Error::Other(format!("repeating section content is invalid: {error}")))?
    {
        return Ok(RepeatingContent::NotBlock);
    }
    Ok(RepeatingContent::Wrap)
}

fn inner_has_nonspace_text(inner: &[u8]) -> Result<bool> {
    let mut wrapped = Vec::new();
    wrapped.extend_from_slice(b"<w:root xmlns:w=\"");
    wrapped.extend_from_slice(W_NS.as_bytes());
    wrapped.extend_from_slice(b"\">");
    wrapped.extend_from_slice(inner);
    wrapped.extend_from_slice(b"</w:root>");
    let mut reader = Reader::from_reader(wrapped.as_slice());
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(_)) => depth += 1,
            Ok(Event::End(_)) => {
                if depth == 0 {
                    return Err(Error::Other(
                        "repeating section content is not a block sequence".to_owned(),
                    ));
                }
                depth -= 1;
            }
            Ok(Event::Text(text)) if depth == 1 && !xml_space(text.as_ref()) => return Ok(true),
            Ok(Event::CData(text)) if depth == 1 && !xml_space(text.as_ref()) => return Ok(true),
            Ok(Event::GeneralRef(reference)) if depth == 1 => {
                let character = decode_one_xml_reference(&reference).map_err(|detail| {
                    Error::Other(format!("repeating section content is invalid: {detail}"))
                })?;
                if !matches!(character, ' ' | '\t' | '\n' | '\r') {
                    return Ok(true);
                }
            }
            Ok(Event::Eof) => return Ok(false),
            Ok(_) => {}
            Err(error) => {
                return Err(Error::Other(format!(
                    "repeating section content is invalid: {error}"
                )));
            }
        }
        buffer.clear();
    }
}

fn wrap_repeating_item(
    xml: &[u8],
    sdt_range: &Range<usize>,
    id: i32,
) -> Result<(Vec<u8>, Range<usize>)> {
    let elems = index_elements(xml)?;
    let sdt = element_at(&elems, sdt_range)?;
    let content = sdt_content_elem(&elems, sdt)
        .ok_or_else(|| Error::Other("repeating section has no item".to_owned()))?;
    let inner_end = content
        .inner_end
        .ok_or_else(|| Error::Other("repeating section has no item".to_owned()))?;
    let inner = &xml[content.open_end..inner_end];
    let inner = std::str::from_utf8(inner)
        .map_err(|error| Error::Other(format!("repeating section content is invalid: {error}")))?;
    let wrapper = format!(
        "<w:sdt><w:sdtPr><w:id w:val=\"{id}\"/><w15:repeatingSectionItem xmlns:w15=\"{W15_NS}\"/></w:sdtPr><w:sdtContent>{inner}</w:sdtContent></w:sdt>"
    );
    let growth = wrapper.len() - (inner_end - content.open_end);
    let mut updated = xml.to_vec();
    updated.splice(content.open_end..inner_end, wrapper.into_bytes());
    Ok((updated, sdt_range.start..sdt_range.end + growth))
}

fn replace_control_type_markup(
    xml: &[u8],
    sdt_range: &Range<usize>,
    markup: &[u8],
) -> Result<Vec<u8>> {
    let elems = index_elements(xml)?;
    let sdt = element_at(&elems, sdt_range)?;
    let sdt_pr = elems
        .iter()
        .find(|elem| elem.parent == Some(sdt.id) && elem.word && elem.local == b"sdtPr")
        .ok_or_else(|| Error::Other("content control has no properties".to_owned()))?;
    let mut type_ranges: Vec<Range<usize>> = elems
        .iter()
        .filter(|elem| elem.parent == Some(sdt_pr.id) && is_type_child(elem))
        .map(|elem| elem.range.clone())
        .collect();
    type_ranges.sort_by_key(|range| range.start);
    if sdt_pr.inner_end.is_none() {
        let expanded = expand_empty(&xml[sdt_pr.range.clone()], markup)?;
        return splice_absolute(xml, sdt_range, &sdt_pr.range, &expanded);
    }
    let insert_at = type_ranges
        .first()
        .map(|range| range.start)
        .unwrap_or_else(|| {
            elems
                .iter()
                .filter(|elem| elem.parent == Some(sdt_pr.id))
                .find(|elem| property_rank(elem).is_some_and(|child| child > 20))
                .map(|elem| elem.range.start)
                .unwrap_or(sdt_pr.inner_end.unwrap_or(sdt_pr.range.end))
        });
    let mut updated = Vec::new();
    let mut cursor = sdt_range.start;
    let mut inserted = false;
    let mut skips = type_ranges.into_iter().peekable();
    while cursor < sdt_range.end {
        if !inserted && cursor == insert_at {
            updated.extend_from_slice(markup);
            inserted = true;
        }
        if skips.peek().is_some_and(|skip| cursor == skip.start) {
            cursor = skips.next().unwrap().end;
            continue;
        }
        updated.push(xml[cursor]);
        cursor += 1;
    }
    if !inserted {
        updated.extend_from_slice(markup);
    }
    Ok(updated)
}

fn is_type_child(elem: &XmlElem) -> bool {
    if elem.word {
        return matches!(
            elem.local.as_slice(),
            b"richText"
                | b"text"
                | b"date"
                | b"picture"
                | b"comboBox"
                | b"dropDownList"
                | b"citation"
                | b"group"
                | b"bibliography"
                | b"equation"
                | b"docPartList"
                | b"docPartObj"
        );
    }
    if elem.w15 {
        return matches!(
            elem.local.as_slice(),
            b"repeatingSection" | b"repeatingSectionItem"
        );
    }
    elem.w14 && elem.local.as_slice() == b"checkbox"
}

fn sdt_content_elem<'a>(elems: &'a [XmlElem], sdt: &XmlElem) -> Option<&'a XmlElem> {
    elems
        .iter()
        .find(|elem| elem.parent == Some(sdt.id) && elem.word && elem.local == b"sdtContent")
}

fn sdt_has_direct_type(
    elems: &[XmlElem],
    sdt: &XmlElem,
    local: &[u8],
    w15: bool,
    w14: bool,
) -> bool {
    let Some(properties) = elems
        .iter()
        .find(|elem| elem.parent == Some(sdt.id) && elem.word && elem.local == b"sdtPr")
    else {
        return false;
    };
    elems.iter().any(|elem| {
        elem.parent == Some(properties.id)
            && elem.local.as_slice() == local
            && if w15 {
                elem.w15
            } else if w14 {
                elem.w14
            } else {
                elem.word
            }
    })
}

fn sdt_is_repeating_section(elems: &[XmlElem], sdt: &XmlElem) -> bool {
    sdt_has_direct_type(elems, sdt, b"repeatingSection", true, false)
}

fn sdt_is_repeating_item(elems: &[XmlElem], sdt: &XmlElem) -> bool {
    sdt_has_direct_type(elems, sdt, b"repeatingSectionItem", true, false)
}

fn ancestor_repeating_section(elems: &[XmlElem], sdt_id: usize) -> Option<&XmlElem> {
    let mut current = elems.iter().find(|elem| elem.id == sdt_id)?;
    while let Some(parent_id) = current.parent {
        let parent = elems.iter().find(|elem| elem.id == parent_id)?;
        if parent.word && parent.local == b"sdt" && sdt_is_repeating_section(elems, parent) {
            return Some(parent);
        }
        current = parent;
    }
    None
}

fn direct_repeating_items<'a>(elems: &'a [XmlElem], section: &XmlElem) -> Vec<&'a XmlElem> {
    let Some(content) = sdt_content_elem(elems, section) else {
        return Vec::new();
    };
    elems
        .iter()
        .filter(|elem| {
            elem.parent == Some(content.id)
                && elem.word
                && elem.local == b"sdt"
                && sdt_is_repeating_item(elems, elem)
        })
        .collect()
}

fn first_repeating_item_range(elems: &[XmlElem], section: &XmlElem) -> Option<Range<usize>> {
    let content = sdt_content_elem(elems, section)?;
    let end = content.inner_end.unwrap_or(content.range.end);
    elems
        .iter()
        .filter(|elem| {
            elem.word
                && elem.local == b"sdt"
                && elem.range.start >= content.open_end
                && elem.range.end <= end
                && sdt_is_repeating_item(elems, elem)
        })
        .min_by_key(|elem| elem.range.start)
        .map(|elem| elem.range.clone())
}

fn repeating_slot_allowed(slot: ParentSlot) -> bool {
    matches!(
        slot,
        ParentSlot::Block | ParentSlot::Table | ParentSlot::Row
    )
}

fn direct_item_section<'a>(elems: &'a [XmlElem], item: &XmlElem) -> Option<&'a XmlElem> {
    if !sdt_is_repeating_item(elems, item) {
        return None;
    }
    let section = ancestor_repeating_section(elems, item.id)?;
    direct_repeating_items(elems, section)
        .iter()
        .any(|child| child.id == item.id)
        .then_some(section)
}

fn section_type_can_retire(elems: &[XmlElem], section: &XmlElem) -> bool {
    if !sdt_is_repeating_section(elems, section)
        || ancestor_repeating_section(elems, section.id).is_some()
    {
        return true;
    }
    !elems.iter().any(|item| {
        item.word
            && item.local == b"sdt"
            && section.range.start < item.range.start
            && item.range.end < section.range.end
            && sdt_is_repeating_item(elems, item)
            && ancestor_repeating_section(elems, item.id)
                .is_some_and(|owner| owner.id == section.id)
    })
}

fn validate_repeating_type_change(
    xml: &[u8],
    range: &Range<usize>,
    payload: &ContentControlTypePayload,
) -> Result<()> {
    let elems = index_elements(xml)?;
    let sdt = element_at(&elems, range)?;
    let item = matches!(payload, ContentControlTypePayload::RepeatingSectionItem);
    let section = matches!(payload, ContentControlTypePayload::RepeatingSection);
    if item && ancestor_repeating_section(&elems, sdt.id).is_none() {
        return Err(Error::Other(
            "repeating section item requires a repeating section parent".to_owned(),
        ));
    }
    if (item || section) && !repeating_slot_allowed(parent_slot(&elems, sdt.id)) {
        return Err(Error::Other(
            "repeating control requires block, row, or cell placement".to_owned(),
        ));
    }
    if !item && direct_item_section(&elems, sdt).is_some() {
        return Err(Error::Other(
            "repeating section must retain repeating section items".to_owned(),
        ));
    }
    if !section && !section_type_can_retire(&elems, sdt) {
        return Err(Error::Other(
            "cannot remove a repeating section type while its items require it".to_owned(),
        ));
    }
    Ok(())
}

fn control_matches_context(
    slot: ParentSlot,
    in_section: bool,
    xml: &[u8],
    range: &Range<usize>,
) -> Result<bool> {
    if matches!(slot, ParentSlot::Other) {
        return Ok(false);
    }
    let elems = index_elements(xml)?;
    let sdt = element_at(&elems, range)?;
    if !sdt_payload_matches(slot, xml, &elems, sdt)? {
        return Ok(false);
    }
    if in_section {
        return Ok(true);
    }
    Ok(!elems.iter().any(|item| {
        item.word
            && item.local == b"sdt"
            && range.start <= item.range.start
            && item.range.end <= range.end
            && sdt_is_repeating_item(&elems, item)
            && ancestor_repeating_section(&elems, item.id)
                .is_none_or(|owner| owner.range.start < range.start || owner.range.end > range.end)
    }))
}

fn story_inside_repeating_section(document: &Document, story: &crate::StoryId) -> Result<bool> {
    let (xml, range) = document.story_owner_span(story)?;
    let elems = index_elements(&xml)?;
    Ok(elems.iter().any(|section| {
        section.word
            && section.local == b"sdt"
            && section.range.start < range.start
            && range.end < section.range.end
            && sdt_is_repeating_section(&elems, section)
    }))
}

fn unwrap_keeps_repeating_context(xml: &[u8], range: &Range<usize>) -> Result<bool> {
    let elems = index_elements(xml)?;
    let sdt = element_at(&elems, range)?;
    if !section_type_can_retire(&elems, sdt) {
        return Ok(false);
    }
    let Some(owner) = direct_item_section(&elems, sdt) else {
        return Ok(true);
    };
    let Some(content) = sdt_content_elem(&elems, sdt) else {
        return Ok(direct_repeating_items(&elems, owner).len() > 1);
    };
    let children: Vec<_> = elems
        .iter()
        .filter(|elem| elem.parent == Some(content.id))
        .collect();
    if children
        .iter()
        .any(|child| !child.word || child.local != b"sdt" || !sdt_is_repeating_item(&elems, child))
    {
        return Ok(false);
    }
    Ok(!children.is_empty() || direct_repeating_items(&elems, owner).len() > 1)
}

fn last_repeating_item(xml: &[u8], sdt_range: &Range<usize>) -> Result<bool> {
    let elems = index_elements(xml)?;
    let sdt = element_at(&elems, sdt_range)?;
    let Some(section) = direct_item_section(&elems, sdt) else {
        return Ok(false);
    };
    Ok(direct_repeating_items(&elems, section).len() == 1)
}

fn freshen_control(document: &mut Document, sdt_bytes: &[u8]) -> Result<Vec<u8>> {
    let mut wrapped = format!(
        "<w:document xmlns:w=\"{W_NS}\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\"><w:body>"
    )
    .into_bytes();
    wrapped.extend_from_slice(sdt_bytes);
    wrapped.extend_from_slice(b"</w:body></w:document>");
    let updated = crate::field::freshen_content_fragment_identities(document, &wrapped)?;
    let elems = index_elements(&updated)?;
    let sdt = elems
        .iter()
        .find(|elem| elem.word && elem.local == b"sdt")
        .ok_or_else(|| Error::Other("content control clone did not keep its shell".to_owned()))?;
    Ok(updated[sdt.range.clone()].to_vec())
}
