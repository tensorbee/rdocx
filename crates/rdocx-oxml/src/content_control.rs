//! WordprocessingML structured document tags (`w:sdt`).

use std::io::Write;

use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer, XmlVersion};

use crate::error::{OxmlError, Result};
use crate::namespace::{W_NS, matches_local_name, nesting_exceeds, weighted_nesting_exceeds};
use crate::numbering::{local_namespace_overrides, namespace_bindings, word_prefixes_at};
use crate::properties::is_word_element;
use crate::raw_xml::{capture_element, capture_empty_element};
use crate::revision::CT_Revision;
use crate::table::{CT_Row, CT_Tbl, CT_Tc, MAX_RECOGNIZED_TABLE_NESTING};
use crate::text::{
    AcceptedRunPath, AcceptedRunPathSegment, CT_P, CT_R, Field, LegacyFormFieldValue, RunContent,
};

/// The bounded content-control type markers that rdocx reports.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SdtType {
    RichText,
    PlainText,
    Picture,
    CheckBox,
    ComboBox,
    DropDownList,
    Date,
    DocumentPartList,
    DocumentPartObject,
    Group,
    RepeatingSection,
    RepeatingSectionItem,
    Citation,
    Equation,
    Bibliography,
}

impl SdtType {
    fn from_element(local: &[u8]) -> Option<Self> {
        match local {
            b"richText" => Some(Self::RichText),
            b"text" => Some(Self::PlainText),
            b"picture" => Some(Self::Picture),
            b"checkbox" => Some(Self::CheckBox),
            b"comboBox" => Some(Self::ComboBox),
            b"dropDownList" => Some(Self::DropDownList),
            b"date" => Some(Self::Date),
            b"docPartList" => Some(Self::DocumentPartList),
            b"docPartObj" => Some(Self::DocumentPartObject),
            b"group" => Some(Self::Group),
            b"repeatingSection" => Some(Self::RepeatingSection),
            b"repeatingSectionItem" => Some(Self::RepeatingSectionItem),
            b"citation" => Some(Self::Citation),
            b"equation" => Some(Self::Equation),
            b"bibliography" => Some(Self::Bibliography),
            _ => None,
        }
    }

    fn element_name(self) -> &'static str {
        match self {
            Self::RichText => "w:richText",
            Self::PlainText => "w:text",
            Self::Picture => "w:picture",
            Self::CheckBox => "w14:checkbox",
            Self::ComboBox => "w:comboBox",
            Self::DropDownList => "w:dropDownList",
            Self::Date => "w:date",
            Self::DocumentPartList => "w:docPartList",
            Self::DocumentPartObject => "w:docPartObj",
            Self::Group => "w:group",
            Self::RepeatingSection => "w15:repeatingSection",
            Self::RepeatingSectionItem => "w15:repeatingSectionItem",
            Self::Citation => "w:citation",
            Self::Equation => "w:equation",
            Self::Bibliography => "w:bibliography",
        }
    }
}

/// The optional custom XML binding carried by `w:dataBinding`.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct CT_DataBinding {
    pub prefix_mappings: Option<String>,
    pub xpath: Option<String>,
    pub store_item_id: Option<String>,
    extra_attributes: Vec<(String, String)>,
}

impl CT_DataBinding {
    fn parse(start: &BytesStart<'_>, inherited: &[String]) -> Result<Self> {
        let prefixes = word_prefixes_at(start, inherited)?;
        let mut binding = Self::default();
        for attribute in start.attributes() {
            let attribute = attribute?;
            let key = attribute.key.as_ref();
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, start.decoder())?
                .into_owned();
            if is_word_attribute(key, b"prefixMappings", &prefixes) {
                binding.prefix_mappings = Some(value);
            } else if is_word_attribute(key, b"xpath", &prefixes) {
                binding.xpath = Some(value);
            } else if is_word_attribute(key, b"storeItemID", &prefixes) {
                binding.store_item_id = Some(value);
            } else {
                binding
                    .extra_attributes
                    .push((std::str::from_utf8(key)?.to_owned(), value));
            }
        }
        Ok(binding)
    }

    fn to_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut start = BytesStart::new("w:dataBinding");
        if let Some(value) = &self.prefix_mappings {
            start.push_attribute(("w:prefixMappings", value.as_str()));
        }
        if let Some(value) = &self.xpath {
            start.push_attribute(("w:xpath", value.as_str()));
        }
        if let Some(value) = &self.store_item_id {
            start.push_attribute(("w:storeItemID", value.as_str()));
        }
        for (name, value) in &self.extra_attributes {
            start.push_attribute((name.as_str(), value.as_str()));
        }
        writer.write_event(Event::Empty(start))?;
        Ok(())
    }
}

#[derive(Debug, Clone, PartialEq)]
enum PropertySlot {
    Alias,
    Tag,
    Id,
    Type,
    DataBinding,
    Raw(Vec<u8>),
}

/// Typed properties and ordered raw slots from `w:sdtPr`.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct CT_SdtPr {
    pub alias: Option<String>,
    pub tag: Option<String>,
    pub id: Option<i32>,
    pub control_type: Option<SdtType>,
    pub data_binding: Option<CT_DataBinding>,
    extra_attributes: Vec<(String, String)>,
    slots: Vec<PropertySlot>,
    type_element: Option<TypeElement>,
}

#[derive(Debug, Clone, PartialEq)]
struct TypeElement {
    control_type: SdtType,
    attributes: Vec<(String, String)>,
    children: Vec<u8>,
}

impl CT_SdtPr {
    fn from_xml_with_prefixes(
        reader: &mut Reader<&[u8]>,
        start: &BytesStart<'_>,
        inherited: &[String],
    ) -> Result<Self> {
        let prefixes = word_prefixes_at(start, inherited)?;
        let mut properties = Self {
            extra_attributes: capture_attributes(start, &prefixes, &[])?,
            ..Self::default()
        };
        let mut buffer = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer) {
                Ok(Event::Start(child)) => {
                    let child_prefixes = word_prefixes_at(&child, &prefixes)?;
                    if properties.parse_modelled(&child, &child_prefixes)? {
                        if let (Some(control_type), Some(PropertySlot::Type)) =
                            (properties.control_type, properties.slots.last())
                        {
                            let raw = capture_element(reader, &child)?;
                            let open = [b"<".as_slice(), &*child, b">"].concat();
                            let close = [b"</".as_slice(), child.name().as_ref(), b">"].concat();
                            let children = raw
                                .strip_prefix(open.as_slice())
                                .and_then(|rest| rest.strip_suffix(close.as_slice()))
                                .ok_or_else(|| {
                                    OxmlError::MissingElement(format!(
                                        "{} end",
                                        String::from_utf8_lossy(child.name().as_ref())
                                    ))
                                })?;
                            properties.type_element = Some(TypeElement {
                                control_type,
                                attributes: capture_attributes(&child, &child_prefixes, &[])?,
                                children: children.to_vec(),
                            });
                        } else {
                            reader.read_to_end_into(child.name(), &mut Vec::new())?;
                        }
                    } else {
                        properties
                            .slots
                            .push(PropertySlot::Raw(capture_element(reader, &child)?));
                    }
                }
                Ok(Event::Empty(child)) => {
                    let child_prefixes = word_prefixes_at(&child, &prefixes)?;
                    if properties.parse_modelled(&child, &child_prefixes)? {
                        if let (Some(control_type), Some(PropertySlot::Type)) =
                            (properties.control_type, properties.slots.last())
                        {
                            properties.type_element = Some(TypeElement {
                                control_type,
                                attributes: capture_attributes(&child, &child_prefixes, &[])?,
                                children: Vec::new(),
                            });
                        }
                    } else {
                        properties
                            .slots
                            .push(PropertySlot::Raw(capture_empty_element(&child)?));
                    }
                }
                Ok(Event::End(end)) if matches_local_name(end.name().as_ref(), b"sdtPr") => break,
                Ok(Event::Eof) => {
                    return Err(OxmlError::MissingElement("w:sdtPr end".to_owned()));
                }
                Ok(event) => push_event_raw(&mut properties.slots, event)?,
                Err(error) => return Err(error.into()),
            }
            buffer.clear();
        }
        Ok(properties)
    }

    fn parse_modelled(&mut self, child: &BytesStart<'_>, prefixes: &[String]) -> Result<bool> {
        let name = child.name();
        let local = name
            .as_ref()
            .rsplit(|byte| *byte == b':')
            .next()
            .unwrap_or(name.as_ref());
        if is_word_element(name.as_ref(), b"alias", prefixes) {
            self.alias = Some(required_word_attribute(child, b"val", prefixes)?);
            self.slots.push(PropertySlot::Alias);
        } else if is_word_element(name.as_ref(), b"tag", prefixes) {
            self.tag = Some(required_word_attribute(child, b"val", prefixes)?);
            self.slots.push(PropertySlot::Tag);
        } else if is_word_element(name.as_ref(), b"id", prefixes) {
            self.id = Some(required_word_attribute(child, b"val", prefixes)?.parse()?);
            self.slots.push(PropertySlot::Id);
        } else if is_word_element(name.as_ref(), b"dataBinding", prefixes) {
            self.data_binding = Some(CT_DataBinding::parse(child, prefixes)?);
            self.slots.push(PropertySlot::DataBinding);
        } else if let Some(control_type) = SdtType::from_element(local)
            && is_supported_type_element(name.as_ref(), local, prefixes)
        {
            if self.control_type.is_some() {
                return Ok(false);
            }
            self.control_type = Some(control_type);
            self.slots.push(PropertySlot::Type);
        } else {
            return Ok(false);
        }
        Ok(true)
    }

    fn to_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut start = BytesStart::new("w:sdtPr");
        for (name, value) in &self.extra_attributes {
            start.push_attribute((name.as_str(), value.as_str()));
        }
        writer.write_event(Event::Start(start))?;

        let mut alias_written = false;
        let mut tag_written = false;
        let mut id_written = false;
        let mut type_written = false;
        let mut binding_written = false;
        for slot in &self.slots {
            match slot {
                PropertySlot::Alias if !alias_written => {
                    write_val_element(writer, "w:alias", self.alias.as_deref())?;
                    alias_written = true;
                }
                PropertySlot::Tag if !tag_written => {
                    write_val_element(writer, "w:tag", self.tag.as_deref())?;
                    tag_written = true;
                }
                PropertySlot::Id if !id_written => {
                    if let Some(id) = self.id {
                        let value = id.to_string();
                        write_val_element(writer, "w:id", Some(&value))?;
                    }
                    id_written = true;
                }
                PropertySlot::Type if !type_written => {
                    self.write_type(writer)?;
                    type_written = true;
                }
                PropertySlot::DataBinding if !binding_written => {
                    if let Some(binding) = &self.data_binding {
                        binding.to_xml(writer)?;
                    }
                    binding_written = true;
                }
                PropertySlot::Raw(raw) => writer.get_mut().write_all(raw)?,
                _ => {}
            }
        }
        if !alias_written {
            write_val_element(writer, "w:alias", self.alias.as_deref())?;
        }
        if !tag_written {
            write_val_element(writer, "w:tag", self.tag.as_deref())?;
        }
        if !id_written && let Some(id) = self.id {
            let value = id.to_string();
            write_val_element(writer, "w:id", Some(&value))?;
        }
        if !type_written {
            self.write_type(writer)?;
        }
        if !binding_written && let Some(binding) = &self.data_binding {
            binding.to_xml(writer)?;
        }
        writer.write_event(Event::End(BytesEnd::new("w:sdtPr")))?;
        Ok(())
    }

    fn write_type<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let Some(control_type) = self.control_type else {
            return Ok(());
        };
        let name = control_type.element_name();
        let mut start = BytesStart::new(name);
        let children = match &self.type_element {
            Some(element) if element.control_type == control_type => {
                for (key, value) in &element.attributes {
                    start.push_attribute((key.as_str(), value.as_str()));
                }
                element.children.as_slice()
            }
            _ => &[],
        };
        if children.is_empty() {
            writer.write_event(Event::Empty(start))?;
        } else {
            let mut element = Writer::new(Vec::new());
            element.write_event(Event::Start(start))?;
            element.get_mut().write_all(children)?;
            element.write_event(Event::End(BytesEnd::new(name)))?;
            writer.get_mut().write_all(&element.into_inner())?;
        }
        Ok(())
    }
}

/// A typed or preserved child of `w:sdtContent`.
#[derive(Debug, Clone, PartialEq)]
pub enum SdtContent {
    Paragraph(CT_P),
    Table(CT_Tbl),
    Row(CT_Row),
    Cell(CT_Tc),
    Run(CT_R),
    ContentControl(CT_Sdt),
    RawXml(Vec<u8>),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum SdtOwner {
    Standalone,
    Body,
    Table,
    Row,
    Cell,
    Inline,
}

impl SdtOwner {
    fn owns(self, local: &[u8]) -> bool {
        match self {
            Self::Standalone => matches!(local, b"p" | b"tbl" | b"tr" | b"tc" | b"r"),
            Self::Body | Self::Cell => matches!(local, b"p" | b"tbl"),
            Self::Table => local == b"tr",
            Self::Row => local == b"tc",
            Self::Inline => local == b"r",
        }
    }

    fn owns_content(self, content: &SdtContent) -> bool {
        matches!(content, SdtContent::ContentControl(_))
            || matches!(
                (self, content),
                (
                    Self::Standalone,
                    SdtContent::Paragraph(_)
                        | SdtContent::Table(_)
                        | SdtContent::Row(_)
                        | SdtContent::Cell(_)
                        | SdtContent::Run(_)
                ) | (
                    Self::Body | Self::Cell,
                    SdtContent::Paragraph(_) | SdtContent::Table(_)
                ) | (Self::Table, SdtContent::Row(_))
                    | (Self::Row, SdtContent::Cell(_))
                    | (Self::Inline, SdtContent::Run(_))
            )
    }
}

#[derive(Debug, Clone, PartialEq)]
enum RootSlot {
    Properties,
    Content,
    Raw(Vec<u8>),
}

/// A recursive WordprocessingML structured document tag.
#[derive(Debug, Clone, PartialEq)]
pub struct CT_Sdt {
    pub properties: Option<CT_SdtPr>,
    pub content: Vec<SdtContent>,
    pub(crate) revisions: Vec<(usize, CT_Revision)>,
    extra_attributes: Vec<(String, String)>,
    content_attributes: Vec<(String, String)>,
    slots: Vec<RootSlot>,
    word_prefixes: Vec<String>,
    content_word_prefixes: Vec<String>,
    inline_run_sources: Vec<InlineRunSource>,
    pending_inline_field_edits: Vec<(Vec<u8>, Vec<u8>)>,
}

#[derive(Debug, Clone, PartialEq)]
struct InlineRunSource {
    content_index: usize,
    original: CT_R,
    raw_xml: Vec<u8>,
}

impl CT_Sdt {
    /// Project complex legacy fields owned by direct inline runs.
    #[doc(hidden)]
    pub fn inline_legacy_form_fields(&self) -> Result<Vec<Field>> {
        Ok(self
            .inline_legacy_form_fields_with_content_indices()?
            .into_iter()
            .map(|(_, field)| field)
            .collect())
    }

    /// Project direct-run legacy fields with their starting content boundary.
    #[doc(hidden)]
    pub fn inline_legacy_form_fields_with_content_indices(&self) -> Result<Vec<(usize, Field)>> {
        let Some(paragraph) = self.inline_runs_as_paragraph()? else {
            return Ok(Vec::new());
        };
        let mut fields = Vec::new();
        let mut source_cursor = 0usize;
        for field in paragraph.runs.iter().flat_map(|run| run.content.iter()) {
            if let RunContent::Field(field) = field {
                let mut owned = Vec::new();
                append_inline_legacy_fields(field, &mut owned)?;
                if owned.is_empty() {
                    continue;
                }
                let source = field
                    .source_replacement()?
                    .map(|(source, _)| source)
                    .ok_or_else(|| {
                        OxmlError::MissingElement("retained inline field source".to_owned())
                    })?;
                let relative = self.inline_run_sources[source_cursor..]
                    .iter()
                    .position(|run| source.starts_with(&run.raw_xml))
                    .ok_or_else(|| {
                        OxmlError::MissingElement("inline field start run".to_owned())
                    })?;
                let source_index = source_cursor + relative;
                let content_index = self.inline_run_sources[source_index].content_index;
                source_cursor = source_index + 1;
                fields.extend(owned.into_iter().map(|field| (content_index, field)));
            }
        }
        Ok(fields)
    }

    /// Replace one projected legacy value in direct inline runs.
    #[doc(hidden)]
    pub fn set_inline_legacy_form_value(
        &mut self,
        ordinal: usize,
        value: LegacyFormFieldValue,
    ) -> Result<bool> {
        let Some(mut paragraph) = self.inline_runs_as_paragraph()? else {
            return Ok(false);
        };
        let mut remaining = ordinal;
        let mut changed = false;
        for run in &mut paragraph.runs {
            for content in &mut run.content {
                if let RunContent::Field(field) = content
                    && set_nth_inline_legacy_form(field, &mut remaining, &value)?
                {
                    changed = true;
                    break;
                }
            }
            if changed {
                break;
            }
        }
        if !changed {
            return Ok(false);
        }
        let mut field_edits = Vec::new();
        for field in paragraph.runs.iter().flat_map(|run| run.content.iter()) {
            if let RunContent::Field(field) = field
                && let Some((source, replacement)) = field.source_replacement()?
                && source != replacement
            {
                field_edits.push((source.to_vec(), replacement));
            }
        }
        let mut writer = Writer::new(Vec::new());
        paragraph.to_xml(&mut writer)?;
        let xml = writer.into_inner();
        let start = xml
            .iter()
            .position(|byte| *byte == b'>')
            .map(|index| index + 1)
            .ok_or_else(|| OxmlError::MissingElement("w:p start tag".to_owned()))?;
        let end = xml
            .windows(b"</w:p>".len())
            .rposition(|window| window == b"</w:p>")
            .ok_or_else(|| OxmlError::MissingElement("w:p end tag".to_owned()))?;
        let replacements = inline_run_fragments(&xml[start..end])?;
        let run_count = self
            .content
            .iter()
            .filter(|content| matches!(content, SdtContent::Run(_)))
            .count();
        if replacements.len() != run_count {
            return Err(OxmlError::InvalidValue(format!(
                "inline content-control run count changed from {run_count} to {}",
                replacements.len()
            )));
        }
        let mut replacements = replacements.into_iter();
        let mut sources = Vec::with_capacity(run_count);
        for (content_index, content) in self.content.iter_mut().enumerate() {
            if matches!(content, SdtContent::Run(_)) {
                let raw_xml = replacements
                    .next()
                    .expect("replacement count was checked above");
                let run = parse_inline_run(&raw_xml, &self.content_word_prefixes)?;
                *content = SdtContent::Run(run.clone());
                sources.push(InlineRunSource {
                    content_index,
                    original: run,
                    raw_xml,
                });
            }
        }
        self.inline_run_sources = sources;
        self.pending_inline_field_edits = field_edits;
        Ok(true)
    }

    fn inline_runs_as_paragraph(&self) -> Result<Option<CT_P>> {
        if !self
            .content
            .iter()
            .any(|content| matches!(content, SdtContent::Run(_)))
        {
            return Ok(None);
        }
        let mut writer = Writer::new(Vec::new());
        let existing_prefix = self
            .content_word_prefixes
            .iter()
            .find(|prefix| !prefix.starts_with('\0'))
            .cloned();
        let root_name = existing_prefix.as_deref().map_or_else(
            || "w:p".to_owned(),
            |prefix| {
                if prefix.is_empty() {
                    "p".to_owned()
                } else {
                    format!("{prefix}:p")
                }
            },
        );
        let mut root = BytesStart::new(root_name.clone());
        if existing_prefix.is_none() {
            root.push_attribute(("xmlns:w", W_NS));
        }
        for (prefix, namespace) in namespace_bindings(&self.content_word_prefixes) {
            if existing_prefix.is_none() && prefix == "w" {
                continue;
            }
            let attribute = if prefix.is_empty() {
                "xmlns".to_owned()
            } else {
                format!("xmlns:{prefix}")
            };
            root.push_attribute((attribute.as_str(), namespace.as_str()));
        }
        writer.write_event(Event::Start(root))?;
        for (content_index, content) in self.content.iter().enumerate() {
            if let SdtContent::Run(run) = content {
                if let Some(source) = self.inline_run_source(content_index, run) {
                    writer.get_mut().write_all(&source.raw_xml)?;
                } else {
                    run.to_xml(&mut writer)?;
                }
            }
        }
        writer.write_event(Event::End(BytesEnd::new(root_name)))?;
        let xml = writer.into_inner();
        let mut reader = Reader::from_reader(xml.as_slice());
        let mut buffer = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer)? {
                Event::Start(start) if matches_local_name(start.name().as_ref(), b"p") => {
                    let prefixes = word_prefixes_at(&start, &self.content_word_prefixes)?;
                    return Ok(Some(CT_P::from_xml_with_prefixes_and_root(
                        &mut reader,
                        &prefixes,
                        Some(&start),
                    )?));
                }
                Event::Eof => return Ok(None),
                _ => {}
            }
            buffer.clear();
        }
    }

    fn inline_run_source(&self, content_index: usize, run: &CT_R) -> Option<&InlineRunSource> {
        self.inline_run_sources
            .iter()
            .find(|source| source.content_index == content_index && source.original == *run)
    }

    /// Return staged source edits for complex fields owned by direct runs.
    #[doc(hidden)]
    pub fn inline_field_source_replacements(&self) -> Vec<(&[u8], &[u8])> {
        self.pending_inline_field_edits
            .iter()
            .map(|(source, replacement)| (source.as_slice(), replacement.as_slice()))
            .collect()
    }

    /// Return revision wrappers at their direct content boundaries.
    #[doc(hidden)]
    pub fn revisions(&self) -> &[(usize, CT_Revision)] {
        &self.revisions
    }

    /// Mutate physical revision projection metadata on a layout-only clone.
    /// This does not publish an edit or rewrite the producer's raw XML.
    #[doc(hidden)]
    pub fn physical_revisions_mut(&mut self) -> &mut [(usize, CT_Revision)] {
        &mut self.revisions
    }

    pub(crate) fn append_accepted_run_paths(
        &self,
        prefix: &mut Vec<AcceptedRunPathSegment>,
        output: &mut Vec<AcceptedRunPath>,
    ) {
        for boundary in 0..=self.content.len() {
            for (revision_index, (_, revision)) in self
                .revisions
                .iter()
                .enumerate()
                .filter(|(_, (at, _))| *at == boundary)
            {
                prefix.push(AcceptedRunPathSegment::Revision(revision_index));
                revision.append_accepted_run_paths(prefix, output);
                prefix.pop();
            }
            match self.content.get(boundary) {
                Some(SdtContent::Run(_)) => {
                    prefix.push(AcceptedRunPathSegment::Run(boundary));
                    output.push(AcceptedRunPath {
                        segments: prefix.clone(),
                    });
                    prefix.pop();
                }
                Some(SdtContent::ContentControl(control)) => {
                    prefix.push(AcceptedRunPathSegment::ContentControl(boundary));
                    control.append_accepted_run_paths(prefix, output);
                    prefix.pop();
                }
                Some(
                    SdtContent::Paragraph(_)
                    | SdtContent::Table(_)
                    | SdtContent::Row(_)
                    | SdtContent::Cell(_)
                    | SdtContent::RawXml(_),
                )
                | None => {}
            }
        }
    }

    pub(crate) fn accepted_run_segments(&self, path: &[AcceptedRunPathSegment]) -> Option<&CT_R> {
        let (first, rest) = path.split_first()?;
        match *first {
            AcceptedRunPathSegment::Run(index) if rest.is_empty() => {
                match self.content.get(index)? {
                    SdtContent::Run(run) => Some(run),
                    _ => None,
                }
            }
            AcceptedRunPathSegment::ContentControl(index) => match self.content.get(index)? {
                SdtContent::ContentControl(control) => control.accepted_run_segments(rest),
                _ => None,
            },
            AcceptedRunPathSegment::Revision(index) => {
                self.revisions.get(index)?.1.accepted_run_segments(rest)
            }
            AcceptedRunPathSegment::Run(_) => None,
        }
    }

    pub(crate) fn replace_accepted_run_segments(
        &mut self,
        path: &[AcceptedRunPathSegment],
        replacement: CT_R,
    ) -> Result<bool> {
        let Some((first, rest)) = path.split_first() else {
            return Ok(false);
        };
        match *first {
            AcceptedRunPathSegment::Run(index) if rest.is_empty() => {
                let Some(SdtContent::Run(run)) = self.content.get_mut(index) else {
                    return Ok(false);
                };
                *run = replacement;
                Ok(true)
            }
            AcceptedRunPathSegment::ContentControl(index) => {
                let Some(SdtContent::ContentControl(control)) = self.content.get_mut(index) else {
                    return Ok(false);
                };
                control.replace_accepted_run_segments(rest, replacement)
            }
            AcceptedRunPathSegment::Revision(index) => {
                let Some((_, revision)) = self.revisions.get_mut(index) else {
                    return Ok(false);
                };
                revision.replace_accepted_run_segments(rest, replacement)
            }
            AcceptedRunPathSegment::Run(_) => Ok(false),
        }
    }

    pub(crate) fn split_accepted_run_segments(
        &mut self,
        path: &[AcceptedRunPathSegment],
        offset: usize,
    ) -> Result<bool> {
        let Some((first, rest)) = path.split_first() else {
            return Ok(false);
        };
        match *first {
            AcceptedRunPathSegment::Run(index) if rest.is_empty() => {
                let Some(SdtContent::Run(run)) = self.content.get(index) else {
                    return Ok(false);
                };
                let mut paragraph = CT_P::new();
                paragraph.runs.push(run.clone());
                paragraph
                    .split_run(0, offset)
                    .map_err(|error| OxmlError::InvalidValue(error.to_string()))?;
                let replacement = paragraph.runs.drain(..).map(SdtContent::Run);
                self.content.splice(index..=index, replacement);
                for (boundary, _) in &mut self.revisions {
                    if *boundary > index {
                        *boundary += 1;
                    }
                }
                for source in &mut self.inline_run_sources {
                    if source.content_index > index {
                        source.content_index += 1;
                    }
                }
                Ok(true)
            }
            AcceptedRunPathSegment::ContentControl(index) => {
                let Some(SdtContent::ContentControl(control)) = self.content.get_mut(index) else {
                    return Ok(false);
                };
                control.split_accepted_run_segments(rest, offset)
            }
            AcceptedRunPathSegment::Revision(index) => {
                let Some((_, revision)) = self.revisions.get_mut(index) else {
                    return Ok(false);
                };
                revision.split_accepted_run_segments(rest, offset)
            }
            AcceptedRunPathSegment::Run(_) => Ok(false),
        }
    }

    /// Remove the accepted-view run at `path`, and a tracked insertion left
    /// with nothing in it. An emptied nested control stays. Returns false for
    /// a stale path.
    pub(crate) fn remove_accepted_run_segments(
        &mut self,
        path: &[AcceptedRunPathSegment],
    ) -> Result<bool> {
        let Some((first, rest)) = path.split_first() else {
            return Ok(false);
        };
        match *first {
            AcceptedRunPathSegment::Run(index) if rest.is_empty() => {
                if !matches!(self.content.get(index), Some(SdtContent::Run(_))) {
                    return Ok(false);
                }
                self.remove_content(index);
                Ok(true)
            }
            AcceptedRunPathSegment::ContentControl(index) => {
                let Some(SdtContent::ContentControl(control)) = self.content.get_mut(index) else {
                    return Ok(false);
                };
                control.remove_accepted_run_segments(rest)
            }
            AcceptedRunPathSegment::Revision(index) => {
                let Some((_, revision)) = self.revisions.get_mut(index) else {
                    return Ok(false);
                };
                match revision.remove_accepted_run_segments(rest)? {
                    None => Ok(false),
                    Some(emptied) => {
                        if emptied {
                            // The revision is written in place of the raw
                            // child at its boundary, which goes with it.
                            let (boundary, _) = self.revisions.remove(index);
                            self.remove_content(boundary);
                        }
                        Ok(true)
                    }
                }
            }
            AcceptedRunPathSegment::Run(_) => Ok(false),
        }
    }

    /// Insert direct `w:sdtContent` children before `index`, keeping the
    /// revision and retained run source positions after it aligned.
    pub(crate) fn insert_content(&mut self, index: usize, children: Vec<SdtContent>) -> bool {
        if index > self.content.len() {
            return false;
        }
        let count = children.len();
        for (boundary, _) in &mut self.revisions {
            if *boundary >= index {
                *boundary += count;
            }
        }
        for source in &mut self.inline_run_sources {
            if source.content_index >= index {
                source.content_index += count;
            }
        }
        self.content.splice(index..index, children);
        true
    }

    /// Remove the selected comment range markers and the reference runs
    /// that are direct `w:sdtContent` children.
    ///
    /// A run loses only the selected references and is removed when nothing
    /// else remains in it. Nested controls and block content are left to the
    /// caller.
    #[doc(hidden)]
    pub fn remove_comment_anchors(&mut self, ids: &[i32]) {
        // Markers added in this session carry the fixed `w` prefix whatever
        // prefix the source document gave the Word namespace.
        let mut marker_prefixes = vec!["w".to_owned()];
        marker_prefixes.extend(self.content_word_prefixes.iter().cloned());
        let removed = self
            .content
            .iter_mut()
            .map(|child| match child {
                SdtContent::Run(run) => {
                    run.remove_comment_references(ids)
                        && run.content.is_empty()
                        && run.extra_xml.is_empty()
                        && run.alt_drawings.is_empty()
                }
                SdtContent::RawXml(raw) => {
                    crate::text::raw_comment_marker_id(raw, &marker_prefixes)
                        .is_some_and(|id| ids.contains(&id))
                }
                _ => false,
            })
            .collect::<Vec<_>>();
        if !removed.contains(&true) {
            return;
        }
        let kept_before = |index: usize| removed[..index].iter().filter(|remove| !**remove).count();
        self.revisions
            .retain(|(boundary, _)| !removed.get(*boundary).copied().unwrap_or(false));
        for (boundary, _) in &mut self.revisions {
            *boundary = kept_before(*boundary);
        }
        self.inline_run_sources
            .retain(|source| !removed.get(source.content_index).copied().unwrap_or(false));
        for source in &mut self.inline_run_sources {
            source.content_index = kept_before(source.content_index);
        }
        self.content = std::mem::take(&mut self.content)
            .into_iter()
            .zip(removed)
            .filter_map(|(child, remove)| (!remove).then_some(child))
            .collect();
    }

    /// Remap the facade-authored bookmark markers among the direct and nested
    /// inline `w:sdtContent` children.
    pub(crate) fn remap_authored_bookmark_ids(
        &mut self,
        remap: &std::collections::HashMap<i32, i32>,
    ) {
        for content in &mut self.content {
            match content {
                SdtContent::RawXml(raw) => crate::text::remap_authored_bookmark_marker(raw, remap),
                SdtContent::ContentControl(control) => control.remap_authored_bookmark_ids(remap),
                _ => {}
            }
        }
    }

    /// Remove one content child, keeping the revisions and the source bytes
    /// of the later runs at their content index.
    pub(crate) fn remove_content(&mut self, index: usize) {
        if index >= self.content.len() {
            return;
        }
        self.content.remove(index);
        for (boundary, _) in &mut self.revisions {
            if *boundary > index {
                *boundary -= 1;
            }
        }
        self.inline_run_sources
            .retain(|source| source.content_index != index);
        for source in &mut self.inline_run_sources {
            if source.content_index > index {
                source.content_index -= 1;
            }
        }
    }

    pub(crate) fn word_prefixes(&self) -> &[String] {
        &self.word_prefixes
    }

    /// Parse a content control at the reader's current `w:sdt` start.
    ///
    /// Fails when content controls nest more than 64 levels deep.
    pub fn from_xml(reader: &mut Reader<&[u8]>, start: &BytesStart<'_>) -> Result<Self> {
        let raw = capture_element(reader, start)?;
        validate_content_control_nesting(&raw)?;
        Self::parse_raw(&raw, &["w".to_owned()], SdtOwner::Standalone)
    }

    pub(crate) fn from_body_raw(raw: &[u8], inherited: &[String]) -> Result<Option<Self>> {
        Self::from_raw_with_context(raw, inherited, SdtOwner::Body)
    }

    pub(crate) fn from_table_raw(raw: &[u8], inherited: &[String]) -> Result<Option<Self>> {
        Self::from_raw_with_context(raw, inherited, SdtOwner::Table)
    }

    pub(crate) fn from_row_raw(raw: &[u8], inherited: &[String]) -> Result<Option<Self>> {
        Self::from_raw_with_context(raw, inherited, SdtOwner::Row)
    }

    pub(crate) fn from_cell_raw(raw: &[u8], inherited: &[String]) -> Result<Option<Self>> {
        Self::from_raw_with_context(raw, inherited, SdtOwner::Cell)
    }

    pub(crate) fn from_inline_raw(raw: &[u8], inherited: &[String]) -> Result<Option<Self>> {
        Self::from_raw_with_context(raw, inherited, SdtOwner::Inline)
    }

    /// Whether a preserved control is admitted by the complete typed parser.
    #[doc(hidden)]
    pub fn story_raw_is_typed(raw: &[u8], inherited: &[String], owner: StorySdtOwner) -> bool {
        matches!(
            Self::from_raw_with_context(raw, inherited, owner.into()),
            Ok(Some(_))
        )
    }

    /// Parse a captured control, or `None` when the typed model does not
    /// admit it and the caller preserves it as raw XML.
    ///
    /// Nesting beyond [`MAX_CONTENT_CONTROL_NESTING`] is an error rather than
    /// `None`: the typed parser and every later walk of the preserved XML
    /// recurse once per level, so such a control cannot be kept at all.
    fn from_raw_with_context(
        raw: &[u8],
        inherited: &[String],
        owner: SdtOwner,
    ) -> Result<Option<Self>> {
        validate_content_control_nesting(raw)?;
        Ok(Self::parse_raw(raw, inherited, owner).ok())
    }

    fn parse_raw(raw: &[u8], inherited: &[String], owner: SdtOwner) -> Result<Self> {
        let mut reader = Reader::from_reader(raw);
        reader.config_mut().trim_text(false);
        let mut buffer = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer)? {
                Event::Start(start) if matches_local_name(start.name().as_ref(), b"sdt") => {
                    return Self::from_xml_with_prefixes(&mut reader, &start, inherited, owner);
                }
                Event::Eof => return Err(OxmlError::MissingElement("w:sdt".to_owned())),
                _ => {}
            }
            buffer.clear();
        }
    }

    fn from_xml_with_prefixes(
        reader: &mut Reader<&[u8]>,
        start: &BytesStart<'_>,
        inherited: &[String],
        owner: SdtOwner,
    ) -> Result<Self> {
        let prefixes = word_prefixes_at(start, inherited)?;
        let mut sdt = Self {
            properties: None,
            content: Vec::new(),
            revisions: Vec::new(),
            extra_attributes: capture_attributes(start, &prefixes, &[])?,
            content_attributes: Vec::new(),
            slots: Vec::new(),
            word_prefixes: prefixes.clone(),
            content_word_prefixes: prefixes.clone(),
            inline_run_sources: Vec::new(),
            pending_inline_field_edits: Vec::new(),
        };
        let mut buffer = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer) {
                Ok(Event::Start(child)) => {
                    let child_prefixes = word_prefixes_at(&child, &prefixes)?;
                    if is_word_element(child.name().as_ref(), b"sdtPr", &child_prefixes)
                        && sdt.properties.is_none()
                    {
                        sdt.properties = Some(CT_SdtPr::from_xml_with_prefixes(
                            reader,
                            &child,
                            &child_prefixes,
                        )?);
                        sdt.slots.push(RootSlot::Properties);
                    } else if is_word_element(child.name().as_ref(), b"sdtContent", &child_prefixes)
                        && !sdt
                            .slots
                            .iter()
                            .any(|slot| matches!(slot, RootSlot::Content))
                    {
                        sdt.content_attributes = capture_attributes(&child, &child_prefixes, &[])?;
                        sdt.content_word_prefixes = child_prefixes.clone();
                        sdt.content = parse_content(
                            reader,
                            &child_prefixes,
                            &mut sdt.revisions,
                            &mut sdt.inline_run_sources,
                            owner,
                        )?;
                        sdt.slots.push(RootSlot::Content);
                    } else {
                        sdt.slots
                            .push(RootSlot::Raw(capture_element(reader, &child)?));
                    }
                }
                Ok(Event::Empty(child)) => {
                    sdt.slots
                        .push(RootSlot::Raw(capture_empty_element(&child)?));
                }
                Ok(Event::End(end)) if matches_local_name(end.name().as_ref(), b"sdt") => break,
                Ok(Event::Eof) => {
                    return Err(OxmlError::MissingElement("w:sdt end".to_owned()));
                }
                Ok(event) => push_root_event_raw(&mut sdt.slots, event)?,
                Err(error) => return Err(error.into()),
            }
            buffer.clear();
        }
        Ok(sdt)
    }

    /// Whether the control has a child element or visible direct text.
    pub fn has_child_content(&self) -> bool {
        self.properties.is_some()
            || !self.content.is_empty()
            || self.slots.iter().any(|slot| match slot {
                RootSlot::Properties | RootSlot::Content => true,
                RootSlot::Raw(raw) => raw_fragment_has_child_content(raw),
            })
    }

    pub(crate) fn to_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut start = BytesStart::new("w:sdt");
        for (name, value) in &self.extra_attributes {
            start.push_attribute((name.as_str(), value.as_str()));
        }
        writer.write_event(Event::Start(start))?;
        let mut properties_written = false;
        let mut content_written = false;
        for slot in &self.slots {
            match slot {
                RootSlot::Properties if !properties_written => {
                    if let Some(properties) = &self.properties {
                        properties.to_xml(writer)?;
                    }
                    properties_written = true;
                }
                RootSlot::Content if !content_written => {
                    self.write_content(writer)?;
                    content_written = true;
                }
                RootSlot::Raw(raw) => writer.get_mut().write_all(raw)?,
                _ => {}
            }
        }
        if !properties_written && let Some(properties) = &self.properties {
            properties.to_xml(writer)?;
        }
        if !content_written
            && (!self.content.is_empty()
                || self
                    .slots
                    .iter()
                    .any(|slot| matches!(slot, RootSlot::Content)))
        {
            self.write_content(writer)?;
        }
        writer.write_event(Event::End(BytesEnd::new("w:sdt")))?;
        Ok(())
    }

    fn write_content<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut start = BytesStart::new("w:sdtContent");
        for (name, value) in &self.content_attributes {
            start.push_attribute((name.as_str(), value.as_str()));
        }
        writer.write_event(Event::Start(start))?;
        for (content_index, child) in self.content.iter().enumerate() {
            match child {
                SdtContent::Paragraph(paragraph) => paragraph.to_xml(writer)?,
                SdtContent::Table(table) => table.to_xml(writer)?,
                SdtContent::Row(row) => row.to_xml(writer)?,
                SdtContent::Cell(cell) => cell.to_xml(writer)?,
                SdtContent::Run(run) => {
                    if let Some(source) = self.inline_run_source(content_index, run) {
                        writer.get_mut().write_all(&source.raw_xml)?;
                    } else {
                        run.to_xml(writer)?;
                    }
                }
                SdtContent::ContentControl(sdt) => sdt.to_xml(writer)?,
                SdtContent::RawXml(raw) => {
                    if let Some((_, revision)) = self
                        .revisions
                        .iter()
                        .find(|(boundary, _)| *boundary == content_index)
                    {
                        revision.write_xml(writer)?;
                    } else {
                        writer.get_mut().write_all(raw)?;
                    }
                }
            }
        }
        writer.write_event(Event::End(BytesEnd::new("w:sdtContent")))?;
        Ok(())
    }

    pub(crate) fn collect_controls<'a>(&'a self, owner: SdtOwner, controls: &mut Vec<&'a CT_Sdt>) {
        for child in self
            .content
            .iter()
            .filter(|child| owner.owns_content(child))
        {
            match child {
                SdtContent::Paragraph(paragraph) => paragraph.collect_controls(controls),
                SdtContent::Table(table) => table.collect_controls(controls),
                SdtContent::Row(row) => row.collect_controls(controls),
                SdtContent::Cell(cell) => cell.collect_controls(controls),
                SdtContent::Run(_) | SdtContent::RawXml(_) => {}
                SdtContent::ContentControl(sdt) => {
                    controls.push(sdt);
                    sdt.collect_controls(owner, controls);
                }
            }
        }
    }

    pub(crate) fn collect_paragraphs<'a>(
        &'a self,
        owner: SdtOwner,
        paragraphs: &mut Vec<&'a CT_P>,
    ) {
        for child in self
            .content
            .iter()
            .filter(|child| owner.owns_content(child))
        {
            match child {
                SdtContent::Paragraph(paragraph) => paragraphs.push(paragraph),
                SdtContent::Table(table) => table.collect_paragraphs(paragraphs),
                SdtContent::Row(row) => row.collect_paragraphs(paragraphs),
                SdtContent::Cell(cell) => cell.collect_paragraphs(paragraphs),
                SdtContent::ContentControl(sdt) => sdt.collect_paragraphs(owner, paragraphs),
                _ => {}
            }
        }
    }

    pub(crate) fn collect_tables<'a>(&'a self, owner: SdtOwner, tables: &mut Vec<&'a CT_Tbl>) {
        for child in self
            .content
            .iter()
            .filter(|child| owner.owns_content(child))
        {
            match child {
                SdtContent::Table(table) => tables.push(table),
                SdtContent::Cell(cell) => cell.collect_tables(tables),
                SdtContent::ContentControl(sdt) => sdt.collect_tables(owner, tables),
                _ => {}
            }
        }
    }

    pub(crate) fn collect_rows<'a>(&'a self, owner: SdtOwner, rows: &mut Vec<&'a CT_Row>) {
        for child in self
            .content
            .iter()
            .filter(|child| owner.owns_content(child))
        {
            match child {
                SdtContent::Row(row) => rows.push(row),
                SdtContent::Table(table) => table.collect_rows(rows),
                SdtContent::ContentControl(sdt) => sdt.collect_rows(owner, rows),
                _ => {}
            }
        }
    }

    pub(crate) fn collect_cells<'a>(&'a self, owner: SdtOwner, cells: &mut Vec<&'a CT_Tc>) {
        for child in self
            .content
            .iter()
            .filter(|child| owner.owns_content(child))
        {
            match child {
                SdtContent::Cell(cell) => cells.push(cell),
                SdtContent::Row(row) => row.collect_cells(cells),
                SdtContent::ContentControl(sdt) => sdt.collect_cells(owner, cells),
                _ => {}
            }
        }
    }

    pub(crate) fn collect_runs<'a>(&'a self, runs: &mut Vec<&'a CT_R>) {
        for child in self
            .content
            .iter()
            .filter(|child| SdtOwner::Inline.owns_content(child))
        {
            match child {
                SdtContent::Run(run) => runs.push(run),
                SdtContent::Paragraph(paragraph) => paragraph.collect_runs(runs),
                SdtContent::ContentControl(sdt) => sdt.collect_runs(runs),
                _ => {}
            }
        }
    }
}

/// The typed parent grammar used to admit a story content control.
#[doc(hidden)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum StorySdtOwner {
    Block,
    Table,
    Row,
    Inline,
}

impl From<StorySdtOwner> for SdtOwner {
    fn from(owner: StorySdtOwner) -> Self {
        match owner {
            StorySdtOwner::Block => Self::Body,
            StorySdtOwner::Table => Self::Table,
            StorySdtOwner::Row => Self::Row,
            StorySdtOwner::Inline => Self::Inline,
        }
    }
}

fn inline_run_fragments(xml: &[u8]) -> Result<Vec<Vec<u8>>> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut runs = Vec::new();
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if matches_local_name(start.name().as_ref(), b"r") => {
                runs.push(capture_element(&mut reader, &start)?);
            }
            Event::Empty(start) if matches_local_name(start.name().as_ref(), b"r") => {
                runs.push(capture_empty_element(&start)?);
            }
            Event::Eof => return Ok(runs),
            _ => {}
        }
        buffer.clear();
    }
}

fn parse_inline_run(xml: &[u8], inherited: &[String]) -> Result<CT_R> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(start) if matches_local_name(start.name().as_ref(), b"r") => {
                let prefixes = word_prefixes_at(&start, inherited)?;
                return CT_R::from_xml_with_prefixes_and_root(&mut reader, &prefixes, Some(&start));
            }
            Event::Empty(start) if matches_local_name(start.name().as_ref(), b"r") => {
                let prefixes = word_prefixes_at(&start, inherited)?;
                return CT_R::from_empty_root(&start, &prefixes);
            }
            Event::Eof => return Err(OxmlError::MissingElement("inline control run".to_owned())),
            _ => {}
        }
        buffer.clear();
    }
}

fn append_inline_legacy_fields(field: &Field, output: &mut Vec<Field>) -> Result<()> {
    field.validate_legacy_form_owner()?;
    if field.legacy_form.is_some() {
        output.push(field.clone());
    }
    for nested in field.nested_fields_in_source_order() {
        append_inline_legacy_fields(nested, output)?;
    }
    Ok(())
}

fn set_nth_inline_legacy_form(
    field: &mut Field,
    remaining: &mut usize,
    value: &LegacyFormFieldValue,
) -> Result<bool> {
    field.set_nth_legacy_form_value_in_source_order(remaining, value)
}

/// Deepest `w:sdt` nesting the readers accept.
///
/// Word nests content controls a handful of levels deep. The typed parser,
/// the serialiser and the story walks recurse once per level, so a bound far
/// above anything a producer writes keeps every one of them within the stack.
/// At this depth a release build loads, edits and saves a document within
/// about 0.75 MB of stack, which fits the 1 MB stacks of Windows threads and
/// of wasm32 and the 2 MB default of spawned Rust threads.
const MAX_CONTENT_CONTROL_NESTING: usize = 64;

/// Deepest `w:txbxContent` nesting the readers accept.
///
/// Word does not put a text box inside another one. Each level runs the
/// drawing, text box and body parsers and the story walks once more, which
/// costs about 40 KB of stack in a release build. At this depth a release
/// build loads, edits and saves a document within about 0.75 MB of stack,
/// the same bound as for content controls.
const MAX_TEXT_BOX_NESTING: usize = 16;

/// How many content control levels a text box level counts for when the two
/// nest inside each other, so that any mix stays within the stack bound of
/// either kind alone.
const TEXT_BOX_NESTING_WEIGHT: usize = MAX_CONTENT_CONTROL_NESTING / MAX_TEXT_BOX_NESTING;

/// Reject a story part whose content controls or text boxes nest deeper than
/// [`MAX_CONTENT_CONTROL_NESTING`] and [`MAX_TEXT_BOX_NESTING`] anywhere,
/// without recursing. A mix of the two may not nest deeper than
/// [`MAX_CONTENT_CONTROL_NESTING`], a text box counting for
/// [`TEXT_BOX_NESTING_WEIGHT`] levels.
///
/// The typed parser checks each control it reads, but a control inside
/// content kept as raw XML, such as a text box, alternate content or a
/// block-level custom XML element, is only walked later, by recursive
/// passes such as revision acceptance. Checking the whole part at open keeps
/// those passes within the same bound.
#[doc(hidden)]
pub fn validate_story_part_nesting(xml: &[u8]) -> Result<()> {
    if !weighted_nesting_exceeds(
        xml,
        &[(b"sdt", 1), (b"txbxContent", TEXT_BOX_NESTING_WEIGHT)],
        MAX_CONTENT_CONTROL_NESTING,
    ) {
        return Ok(());
    }
    if nesting_exceeds(xml, &[b"sdt"], MAX_CONTENT_CONTROL_NESTING) {
        return Err(content_control_nesting_error());
    }
    if nesting_exceeds(xml, &[b"txbxContent"], MAX_TEXT_BOX_NESTING) {
        return Err(OxmlError::InvalidValue(format!(
            "text box nesting exceeds {MAX_TEXT_BOX_NESTING} levels"
        )));
    }
    Err(OxmlError::InvalidValue(format!(
        "content control and text box nesting exceeds {MAX_CONTENT_CONTROL_NESTING} levels, \
         a text box counting for {TEXT_BOX_NESTING_WEIGHT}"
    )))
}

fn content_control_nesting_error() -> OxmlError {
    OxmlError::InvalidValue(format!(
        "content control nesting exceeds {MAX_CONTENT_CONTROL_NESTING} levels"
    ))
}

/// Reject XML whose content controls nest deeper than
/// [`MAX_CONTENT_CONTROL_NESTING`], without recursing.
///
/// Tables count too: a table parser starts its own nesting count inside each
/// control, so without this check alternating controls and tables would nest
/// tables past the table limit.
fn validate_content_control_nesting(xml: &[u8]) -> Result<()> {
    validate_story_part_nesting(xml)?;
    if nesting_exceeds(xml, &[b"tbl"], MAX_RECOGNIZED_TABLE_NESTING) {
        return Err(OxmlError::InvalidValue(
            "recognized model nesting exceeds table limit".to_owned(),
        ));
    }
    Ok(())
}

fn parse_content(
    reader: &mut Reader<&[u8]>,
    inherited: &[String],
    revisions: &mut Vec<(usize, CT_Revision)>,
    inline_run_sources: &mut Vec<InlineRunSource>,
    owner: SdtOwner,
) -> Result<Vec<SdtContent>> {
    // Every nested `w:sdt` adds this frame to the stack, so it holds no
    // `SdtContent` value of its own: the helpers build each child and push it,
    // and their large temporaries leave the stack before the next level.
    let mut content = Vec::new();
    let mut sources = ContentSources {
        content: &mut content,
        revisions,
        inline_run_sources,
    };
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(child)) => {
                let prefixes = word_prefixes_at(&child, inherited)?;
                if is_word_element(child.name().as_ref(), b"sdt", &prefixes) {
                    push_nested_control(reader, &child, &prefixes, &mut sources, owner)?;
                } else {
                    push_content_child(reader, &child, &prefixes, inherited, &mut sources, owner)?;
                }
            }
            Ok(Event::Empty(child)) => {
                push_empty_content_child(&child, inherited, &mut sources, owner)?;
            }
            Ok(Event::End(end)) if matches_local_name(end.name().as_ref(), b"sdtContent") => break,
            Ok(Event::Eof) => {
                return Err(OxmlError::MissingElement("w:sdtContent end".to_owned()));
            }
            Ok(event) => sources.push(SdtContent::RawXml(event_to_raw(event)?)),
            Err(error) => return Err(error.into()),
        }
        buffer.clear();
    }
    Ok(content)
}

/// The children of one `w:sdtContent` and the side tables that record
/// sources against their positions.
struct ContentSources<'a> {
    content: &'a mut Vec<SdtContent>,
    revisions: &'a mut Vec<(usize, CT_Revision)>,
    inline_run_sources: &'a mut Vec<InlineRunSource>,
}

// Building an `SdtContent` in these helpers keeps that large value out of
// the frames that stay on the stack while a nested control is parsed.
impl ContentSources<'_> {
    #[inline(never)]
    fn push(&mut self, item: SdtContent) {
        self.content.push(item);
    }

    #[inline(never)]
    fn push_control(&mut self, control: CT_Sdt) {
        self.content.push(SdtContent::ContentControl(control));
    }
}

/// Parse a nested control in place, or preserve it as raw XML when the typed
/// model does not admit it.
///
/// The outermost control's [`CT_Sdt::from_raw_with_context`] has already
/// bounded the nesting. Parsing from a copy of the reader, rather than from a
/// captured copy of the subtree, reads each byte once however deep the
/// controls nest, and the untouched reader still captures a rejected control.
#[inline(never)]
fn push_nested_control(
    reader: &mut Reader<&[u8]>,
    child: &BytesStart<'_>,
    prefixes: &[String],
    sources: &mut ContentSources<'_>,
    owner: SdtOwner,
) -> Result<()> {
    let mut attempt = reader.clone();
    match CT_Sdt::from_xml_with_prefixes(&mut attempt, child, prefixes, owner) {
        Ok(sdt) => {
            *reader = attempt;
            sources.push_control(sdt);
        }
        Err(_) => sources.push(SdtContent::RawXml(capture_element(reader, child)?)),
    }
    Ok(())
}

/// The Word block, row, cell and run elements a control may own.
fn owned_word_child(child: &BytesStart<'_>, prefixes: &[String]) -> Option<&'static [u8]> {
    [
        b"p".as_slice(),
        b"tbl".as_slice(),
        b"tr".as_slice(),
        b"tc".as_slice(),
        b"r".as_slice(),
    ]
    .into_iter()
    .find(|local| is_word_element(child.name().as_ref(), local, prefixes))
}

/// Parse one non-control `w:sdtContent` child that has content.
#[inline(never)]
fn push_content_child(
    reader: &mut Reader<&[u8]>,
    child: &BytesStart<'_>,
    prefixes: &[String],
    inherited: &[String],
    sources: &mut ContentSources<'_>,
    owner: SdtOwner,
) -> Result<()> {
    let owned = owned_word_child(child, prefixes);
    let index = sources.content.len();
    let item = if owned.is_some_and(|local| !owner.owns(local)) {
        SdtContent::RawXml(capture_element(reader, child)?)
    } else if owned == Some(b"p".as_slice()) {
        SdtContent::Paragraph(CT_P::from_xml_with_prefixes_and_root(
            reader,
            prefixes,
            Some(child),
        )?)
    } else if owned == Some(b"tbl".as_slice()) {
        let owner_bindings = local_namespace_overrides(child, inherited)?;
        SdtContent::Table(CT_Tbl::from_xml_with_prefixes_and_owner_bindings(
            reader,
            prefixes,
            &owner_bindings,
        )?)
    } else if owned == Some(b"tr".as_slice()) {
        let owner_bindings = local_namespace_overrides(child, inherited)?;
        SdtContent::Row(CT_Row::from_xml_with_prefixes_and_owner_bindings(
            reader,
            prefixes,
            &owner_bindings,
            Some(child),
        )?)
    } else if owned == Some(b"tc".as_slice()) {
        let owner_bindings = local_namespace_overrides(child, inherited)?;
        SdtContent::Cell(CT_Tc::from_xml_with_prefixes_and_owner_bindings(
            reader,
            prefixes,
            &owner_bindings,
        )?)
    } else if owned == Some(b"r".as_slice()) {
        let raw_xml = capture_element(reader, child)?;
        let run = parse_inline_run(&raw_xml, inherited)?;
        sources.inline_run_sources.push(InlineRunSource {
            content_index: index,
            original: run.clone(),
            raw_xml,
        });
        SdtContent::Run(run)
    } else {
        let raw = capture_element(reader, child)?;
        if let Some(revision) = CT_Revision::from_raw(raw.clone(), prefixes) {
            sources.revisions.push((index, revision));
        }
        SdtContent::RawXml(raw)
    };
    sources.content.push(item);
    Ok(())
}

/// Parse one empty `w:sdtContent` child.
#[inline(never)]
fn push_empty_content_child(
    child: &BytesStart<'_>,
    inherited: &[String],
    sources: &mut ContentSources<'_>,
    owner: SdtOwner,
) -> Result<()> {
    let prefixes = word_prefixes_at(child, inherited)?;
    let owned = owned_word_child(child, &prefixes);
    let index = sources.content.len();
    let item = if owned.is_some_and(|local| !owner.owns(local)) {
        SdtContent::RawXml(capture_empty_element(child)?)
    } else if owned == Some(b"p".as_slice()) {
        SdtContent::Paragraph(CT_P::from_empty_root(child, &prefixes)?)
    } else if owned == Some(b"tbl".as_slice()) {
        SdtContent::Table(CT_Tbl::new())
    } else if owned == Some(b"tr".as_slice()) {
        SdtContent::Row(CT_Row::from_empty_root(child, &prefixes)?)
    } else if owned == Some(b"tc".as_slice()) {
        SdtContent::Cell(CT_Tc {
            properties: None,
            content: Vec::new(),
            extra_xml: Vec::new(),
        })
    } else if owned == Some(b"r".as_slice()) {
        let run = CT_R {
            properties: None,
            content: Vec::new(),
            extra_xml: Vec::new(),
            extra_xml_positions: Vec::new(),
            alt_drawings: Vec::new(),
        };
        sources.inline_run_sources.push(InlineRunSource {
            content_index: index,
            original: run.clone(),
            raw_xml: capture_empty_element(child)?,
        });
        SdtContent::Run(run)
    } else {
        let raw = capture_empty_element(child)?;
        if let Some(revision) = CT_Revision::from_raw(raw.clone(), &prefixes) {
            sources.revisions.push((index, revision));
        }
        SdtContent::RawXml(raw)
    };
    sources.content.push(item);
    Ok(())
}

fn write_val_element<W: Write>(
    writer: &mut Writer<W>,
    name: &str,
    value: Option<&str>,
) -> Result<()> {
    if let Some(value) = value {
        let mut start = BytesStart::new(name);
        start.push_attribute(("w:val", value));
        writer.write_event(Event::Empty(start))?;
    }
    Ok(())
}

fn required_word_attribute(
    start: &BytesStart<'_>,
    local: &[u8],
    prefixes: &[String],
) -> Result<String> {
    for attribute in start.attributes() {
        let attribute = attribute?;
        if is_word_attribute(attribute.key.as_ref(), local, prefixes) {
            return Ok(attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, start.decoder())?
                .into_owned());
        }
    }
    Err(OxmlError::MissingElement(format!(
        "w:{} attribute",
        String::from_utf8_lossy(local)
    )))
}

fn is_word_attribute(key: &[u8], local: &[u8], prefixes: &[String]) -> bool {
    let Some(separator) = key.iter().position(|byte| *byte == b':') else {
        return false;
    };
    key.get(separator + 1..) == Some(local)
        && prefixes
            .iter()
            .any(|prefix| prefix.as_bytes() == &key[..separator])
}

fn is_supported_type_element(name: &[u8], local: &[u8], prefixes: &[String]) -> bool {
    if is_word_element(name, local, prefixes) {
        return true;
    }
    let Some(separator) = name.iter().position(|byte| *byte == b':') else {
        return false;
    };
    let prefix = &name[..separator];
    prefixes.iter().any(|binding| {
        let Some(rest) = binding.strip_prefix('\0') else {
            return false;
        };
        let Some((bound_prefix, namespace)) = rest.split_once('\0') else {
            return false;
        };
        bound_prefix.as_bytes() == prefix
            && matches!(
                namespace,
                "http://schemas.microsoft.com/office/word/2010/wordml"
                    | "http://schemas.microsoft.com/office/word/2012/wordml"
            )
    })
}

fn capture_attributes(
    start: &BytesStart<'_>,
    prefixes: &[String],
    modelled: &[&[u8]],
) -> Result<Vec<(String, String)>> {
    let mut attributes = Vec::new();
    for attribute in start.attributes() {
        let attribute = attribute?;
        if modelled
            .iter()
            .any(|local| is_word_attribute(attribute.key.as_ref(), local, prefixes))
        {
            continue;
        }
        attributes.push((
            std::str::from_utf8(attribute.key.as_ref())?.to_owned(),
            attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, start.decoder())?
                .into_owned(),
        ));
    }
    Ok(attributes)
}

fn event_to_raw(event: Event<'_>) -> Result<Vec<u8>> {
    let mut writer = Writer::new(Vec::new());
    writer.write_event(event.into_owned())?;
    Ok(writer.into_inner())
}

fn push_event_raw(slots: &mut Vec<PropertySlot>, event: Event<'_>) -> Result<()> {
    let raw = event_to_raw(event)?;
    if !raw.is_empty() {
        slots.push(PropertySlot::Raw(raw));
    }
    Ok(())
}

fn push_root_event_raw(slots: &mut Vec<RootSlot>, event: Event<'_>) -> Result<()> {
    let raw = event_to_raw(event)?;
    if !raw.is_empty() {
        slots.push(RootSlot::Raw(raw));
    }
    Ok(())
}

fn raw_fragment_has_child_content(raw: &[u8]) -> bool {
    let mut reader = Reader::from_reader(raw);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(_) | Event::Empty(_)) => return true,
            Ok(Event::Text(text)) if !text.as_ref().iter().all(u8::is_ascii_whitespace) => {
                return true;
            }
            Ok(Event::CData(text)) if !text.as_ref().iter().all(u8::is_ascii_whitespace) => {
                return true;
            }
            Ok(Event::GeneralRef(reference)) => match reference.resolve_char_ref() {
                Ok(Some(character)) if !character.is_ascii_whitespace() => return true,
                Ok(Some(_)) => {}
                Ok(None) => return true,
                Err(_) => return false,
            },
            Ok(Event::Eof) | Err(_) => return false,
            _ => {}
        }
        buffer.clear();
    }
}

#[cfg(test)]
mod tests {
    use quick_xml::Reader;
    use quick_xml::events::Event;

    use super::*;
    use crate::document::{BodyContent, CT_Document};

    const W_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    fn parse_document(body: &str) -> CT_Document {
        let xml = format!(r#"<w:document xmlns:w="{W_NS}"><w:body>{body}</w:body></w:document>"#);
        CT_Document::from_xml(xml.as_bytes()).expect("document parses")
    }

    fn parse_table(inner: &str) -> CT_Tbl {
        let xml = format!(r#"<w:tbl xmlns:w="{W_NS}">{inner}</w:tbl>"#);
        let mut reader = Reader::from_str(&xml);
        reader.config_mut().trim_text(false);
        let mut buffer = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer) {
                Ok(Event::Start(start)) if start.local_name().as_ref() == b"tbl" => break,
                Ok(Event::Eof) => panic!("table start is missing"),
                Ok(_) => {}
                Err(error) => panic!("table XML is malformed: {error}"),
            }
            buffer.clear();
        }
        CT_Tbl::from_xml(&mut reader).expect("table parses")
    }

    fn parse_standalone_control(xml: &str) -> CT_Sdt {
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(false);
        let mut buffer = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer) {
                Ok(Event::Start(start)) if start.local_name().as_ref() == b"sdt" => {
                    return CT_Sdt::from_xml(&mut reader, &start).expect("control parses");
                }
                Ok(Event::Eof) => panic!("control start is missing"),
                Ok(_) => {}
                Err(error) => panic!("control XML is valid: {error}"),
            }
            buffer.clear();
        }
    }

    #[test]
    fn sdt_properties_report_tag_alias_id_type_and_binding() {
        let document = parse_document(
            r#"<x:sdt xmlns:x="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><x:sdtPr><x:alias x:val="Customer"/><x:dataBinding x:storeItemID="{A}" x:xpath="/root/name" x:prefixMappings="xmlns:n='urn:x'"/><x:id x:val="42"/><x:tag x:val="customer"/><x:text/><x:temporary x:val="1"/></x:sdtPr><x:sdtContent><x:p/></x:sdtContent></x:sdt>"#,
        );
        let BodyContent::ContentControl(sdt) = &document.body.content[0] else {
            panic!("content control remains opaque");
        };
        assert_eq!(document.body.paragraphs().count(), 1);
        let properties = sdt.properties.as_ref().expect("properties");
        assert_eq!(properties.tag.as_deref(), Some("customer"));
        assert_eq!(properties.alias.as_deref(), Some("Customer"));
        assert_eq!(properties.id, Some(42));
        assert_eq!(properties.control_type, Some(SdtType::PlainText));
        let binding = properties.data_binding.as_ref().expect("binding");
        assert_eq!(binding.store_item_id.as_deref(), Some("{A}"));
        assert_eq!(binding.xpath.as_deref(), Some("/root/name"));
        assert_eq!(binding.prefix_mappings.as_deref(), Some("xmlns:n='urn:x'"));

        let xml = String::from_utf8(document.to_xml().expect("serializes")).expect("UTF-8");
        let typed = xml.find("<w:text").expect("typed marker");
        let unknown = xml.find("<x:temporary").expect("unknown property");
        assert!(typed < unknown, "unmodelled property moved");
    }

    #[test]
    fn content_control_type_payload_round_trips_until_type_changes() {
        let cases = [
            (
                r#"<w14:checkbox w14:custom="kept"><w14:checked w14:val="1"/></w14:checkbox>"#,
                SdtType::CheckBox,
                "w14:checked",
            ),
            (
                r#"<w:date w:fullDate="2026-09-15T00:00:00Z"><w:dateFormat w:val="yyyy-MM-dd"/></w:date>"#,
                SdtType::Date,
                "w:dateFormat",
            ),
            (
                r#"<w:comboBox w:lastValue="B"><w:listItem w:displayText="A" w:value="a"/></w:comboBox>"#,
                SdtType::ComboBox,
                "w:listItem",
            ),
            (
                r#"<w15:repeatingSection w15:sectionTitle="Rows"><w15:doNotAllowInsertDeleteSection w15:val="1"/></w15:repeatingSection>"#,
                SdtType::RepeatingSection,
                "w15:doNotAllowInsertDeleteSection",
            ),
            (
                r#"<w:citation w:val="source"><w:customXml w:uri="urn:citation"/></w:citation>"#,
                SdtType::Citation,
                "w:customXml",
            ),
            (
                r#"<w:equation w:val="inline"><w:proofErr w:type="spellStart"/></w:equation>"#,
                SdtType::Equation,
                "w:proofErr",
            ),
        ];

        for (type_xml, expected_type, retained_child) in cases {
            let xml = format!(
                r#"<w:sdt xmlns:w="{W_NS}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"><w:sdtPr>{type_xml}<w:tag w:val="kept"/></w:sdtPr><w:sdtContent><w:p/></w:sdtContent></w:sdt>"#,
            );
            let mut control = parse_standalone_control(&xml);
            assert_eq!(
                control.properties.as_ref().unwrap().control_type,
                Some(expected_type)
            );

            let mut preserved = Vec::new();
            control
                .to_xml(&mut Writer::new(&mut preserved))
                .expect("control writes");
            let preserved = String::from_utf8(preserved).expect("control XML is UTF-8");
            assert!(
                preserved.contains(retained_child),
                "payload was lost: {preserved}"
            );
            assert!(preserved.contains("custom=\"kept\"") || expected_type != SdtType::CheckBox);

            control.properties.as_mut().unwrap().control_type = Some(SdtType::PlainText);
            let mut changed = Vec::new();
            control
                .to_xml(&mut Writer::new(&mut changed))
                .expect("changed control writes");
            let changed = String::from_utf8(changed).expect("changed control XML is UTF-8");
            assert!(changed.contains("<w:text/>"));
            assert!(!changed.contains(retained_child));
            assert!(changed.contains("<w:tag w:val=\"kept\"/>"));
        }
    }

    #[test]
    fn controls_at_all_five_levels_round_trip_without_losing_content() {
        let body = r#"<w:sdt><w:sdtPr><w:alias w:val="Block"/><w:id w:val="1"/><w:tag w:val="block"/><w:richText/></w:sdtPr><w:sdtContent><w:tbl><w:sdt><w:sdtPr><w:alias w:val="Row"/><w:id w:val="2"/><w:tag w:val="row"/><w:text/></w:sdtPr><w:sdtContent><w:tr><w:sdt><w:sdtPr><w:alias w:val="Cell"/><w:id w:val="3"/><w:tag w:val="cell"/><w:text/></w:sdtPr><w:sdtContent><w:tc><w:sdt><w:sdtPr><w:alias w:val="Paragraph"/><w:id w:val="4"/><w:tag w:val="paragraph"/><w:text/></w:sdtPr><w:sdtContent><w:p><w:sdt><w:sdtPr><w:alias w:val="Run"/><w:id w:val="5"/><w:tag w:val="run"/><w:text/></w:sdtPr><w:sdtContent><w:r><w:t>visible</w:t></w:r></w:sdtContent></w:sdt></w:p></w:sdtContent></w:sdt></w:tc></w:sdtContent></w:sdt></w:tr></w:sdtContent></w:sdt></w:tbl></w:sdtContent></w:sdt>"#;
        let document = parse_document(body);
        let controls = document.body.content_controls();
        assert_eq!(controls.len(), 5);
        let metadata = controls
            .iter()
            .map(|sdt| {
                let properties = sdt.properties.as_ref().expect("properties");
                (
                    properties.tag.as_deref(),
                    properties.alias.as_deref(),
                    properties.id,
                    properties.control_type,
                )
            })
            .collect::<Vec<_>>();
        assert_eq!(
            metadata,
            [
                (
                    Some("block"),
                    Some("Block"),
                    Some(1),
                    Some(SdtType::RichText)
                ),
                (Some("row"), Some("Row"), Some(2), Some(SdtType::PlainText)),
                (
                    Some("cell"),
                    Some("Cell"),
                    Some(3),
                    Some(SdtType::PlainText)
                ),
                (
                    Some("paragraph"),
                    Some("Paragraph"),
                    Some(4),
                    Some(SdtType::PlainText)
                ),
                (Some("run"), Some("Run"), Some(5), Some(SdtType::PlainText)),
            ]
        );
        assert_eq!(
            document
                .body
                .tables()
                .next()
                .expect("wrapped table")
                .rows()
                .len(),
            1
        );
        assert_eq!(
            document
                .body
                .paragraphs()
                .next()
                .expect("wrapped paragraph")
                .text(),
            "visible"
        );

        let saved = document.to_xml().expect("serializes");
        let reopened = CT_Document::from_xml(&saved).expect("reopens");
        assert_eq!(reopened.body.content_controls().len(), 5);
        assert_eq!(
            reopened.body.paragraphs().next().expect("paragraph").text(),
            "visible"
        );
    }

    #[test]
    fn unmodelled_sdt_properties_and_children_remain_byte_identical() {
        let raw_property = r#"<w15:appearance xmlns:w15="urn:producer" w15:val="hidden"><w15:ext>one &amp; two</w15:ext></w15:appearance>"#;
        let raw_child =
            r#"<p:custom xmlns:p="urn:producer" p:flag="1"><p:child/><!--note--></p:custom>"#;
        let body = format!(
            r#"<w:sdt w:rsidR="00112233"><w:sdtPr><w:alias w:val="Known"/>{raw_property}<w:unsupportedType w:val="future"/></w:sdtPr><w:sdtContent><w:p/>{raw_child}</w:sdtContent></w:sdt>"#
        );
        let document = parse_document(&body);
        let saved = String::from_utf8(document.to_xml().expect("serializes")).expect("UTF-8");
        assert!(saved.contains(r#"<w:sdt w:rsidR="00112233">"#));
        assert!(saved.contains(raw_property));
        assert!(saved.contains(r#"<w:unsupportedType w:val="future"/>"#));
        assert!(saved.contains(raw_child));

        let malformed = parse_document(
            r#"<w:sdt><w:sdtPr><w:id w:val="not-an-integer"/></w:sdtPr><w:sdtContent><w:p/></w:sdtContent></w:sdt>"#,
        );
        assert!(matches!(malformed.body.content[0], BodyContent::RawXml(_)));
    }

    fn nested_controls(depth: usize, inner: &str) -> String {
        format!(
            "{}{inner}{}",
            "<w:sdt><w:sdtContent>".repeat(depth),
            "</w:sdtContent></w:sdt>".repeat(depth)
        )
    }

    fn document_error(body: &str) -> String {
        let xml = format!(r#"<w:document xmlns:w="{W_NS}"><w:body>{body}</w:body></w:document>"#);
        CT_Document::from_xml(xml.as_bytes())
            .expect_err("deep controls are rejected")
            .to_string()
    }

    #[test]
    fn controls_nested_past_the_limit_are_rejected_at_every_placement() {
        let run = "<w:r><w:t>x</w:t></w:r>";
        let paragraph = "<w:p><w:r><w:t>x</w:t></w:r></w:p>";
        let limit = MAX_CONTENT_CONTROL_NESTING;
        let body = |placement: &str, depth: usize| match placement {
            "block" => nested_controls(depth, paragraph),
            "inline" => format!("<w:p>{}</w:p>", nested_controls(depth, run)),
            "cell" => format!(
                "<w:tbl><w:tr><w:tc>{}</w:tc></w:tr></w:tbl>",
                nested_controls(depth, "<w:p/>")
            ),
            _ => format!(
                "<w:tbl>{}</w:tbl>",
                nested_controls(depth, "<w:tr><w:tc><w:p/></w:tc></w:tr>")
            ),
        };
        for name in ["block", "inline", "cell", "row"] {
            assert!(
                document_error(&body(name, limit + 1))
                    .contains("content control nesting exceeds 64 levels"),
                "{name}"
            );
            parse_document(&body(name, limit));
        }

        let mut depth = 0;
        let document = parse_document(&nested_controls(limit, paragraph));
        let BodyContent::ContentControl(outer) = &document.body.content[0] else {
            panic!("the outer control is typed");
        };
        let mut control = outer;
        loop {
            depth += 1;
            match &control.content[..] {
                [SdtContent::ContentControl(inner)] => control = inner,
                [SdtContent::Paragraph(_)] => break,
                other => panic!("unexpected content {other:?}"),
            }
        }
        assert_eq!(depth, limit);

        let xml = nested_controls(limit + 1, run).replacen(
            "<w:sdt>",
            &format!(r#"<w:sdt xmlns:w="{W_NS}">"#),
            1,
        );
        let mut reader = Reader::from_str(&xml);
        let Ok(Event::Start(start)) = reader.read_event() else {
            panic!("control start");
        };
        let start = start.into_owned();
        assert!(CT_Sdt::from_xml(&mut reader, &start).is_err());
    }

    /// A nested control the typed model rejects stays raw in place, and the
    /// parent goes on reading the siblings that follow it.
    #[test]
    fn rejected_nested_control_is_preserved_before_its_siblings() {
        let rejected = r#"<w:sdt><w:sdtPr><w:id w:val="not-an-integer"/></w:sdtPr><w:sdtContent><w:p/></w:sdtContent></w:sdt>"#;
        let document = parse_document(&format!(
            r#"<w:sdt><w:sdtContent>{rejected}<w:p><w:r><w:t>after</w:t></w:r></w:p></w:sdtContent></w:sdt>"#
        ));
        let BodyContent::ContentControl(control) = &document.body.content[0] else {
            panic!("the outer control is typed");
        };
        let [SdtContent::RawXml(raw), SdtContent::Paragraph(paragraph)] = &control.content[..]
        else {
            panic!("unexpected content {:?}", control.content);
        };
        assert_eq!(std::str::from_utf8(raw).unwrap(), rejected);
        assert_eq!(paragraph.text(), "after");
    }

    #[test]
    fn block_control_parser_types_only_children_owned_at_its_placement() {
        let body = CT_Sdt::from_body_raw(
            br#"<w:sdt><w:sdtContent><w:p/><w:tbl/><w:tr data-case="body-row"/><w:tc data-case="body-cell"/></w:sdtContent></w:sdt>"#,
            &["w".to_owned()],
        )
        .expect("body control is within the nesting limit")
        .expect("body control parses");
        assert!(matches!(body.content[0], SdtContent::Paragraph(_)));
        assert!(matches!(body.content[1], SdtContent::Table(_)));
        assert_eq!(
            body.content[2],
            SdtContent::RawXml(br#"<w:tr data-case="body-row"/>"#.to_vec())
        );
        assert_eq!(
            body.content[3],
            SdtContent::RawXml(br#"<w:tc data-case="body-cell"/>"#.to_vec())
        );

        let table = CT_Sdt::from_table_raw(
            br#"<w:sdt><w:sdtContent><w:tr/><w:p data-case="table-paragraph"/><w:tbl data-case="table-table"/><w:tc data-case="table-cell"/></w:sdtContent></w:sdt>"#,
            &["w".to_owned()],
        )
        .expect("table control is within the nesting limit")
        .expect("table control parses");
        assert!(matches!(table.content[0], SdtContent::Row(_)));
        assert_eq!(
            table.content[1],
            SdtContent::RawXml(br#"<w:p data-case="table-paragraph"/>"#.to_vec())
        );
        assert_eq!(
            table.content[2],
            SdtContent::RawXml(br#"<w:tbl data-case="table-table"/>"#.to_vec())
        );
        assert_eq!(
            table.content[3],
            SdtContent::RawXml(br#"<w:tc data-case="table-cell"/>"#.to_vec())
        );

        let row = CT_Sdt::from_row_raw(
            br#"<w:sdt><w:sdtContent><w:tc/><w:p data-case="row-paragraph"/><w:tbl data-case="row-table"/><w:tr data-case="row-row"/></w:sdtContent></w:sdt>"#,
            &["w".to_owned()],
        )
        .expect("row control is within the nesting limit")
        .expect("row control parses");
        assert!(matches!(row.content[0], SdtContent::Cell(_)));
        assert_eq!(
            row.content[1],
            SdtContent::RawXml(br#"<w:p data-case="row-paragraph"/>"#.to_vec())
        );
        assert_eq!(
            row.content[2],
            SdtContent::RawXml(br#"<w:tbl data-case="row-table"/>"#.to_vec())
        );
        assert_eq!(
            row.content[3],
            SdtContent::RawXml(br#"<w:tr data-case="row-row"/>"#.to_vec())
        );

        let cell = CT_Sdt::from_cell_raw(
            br#"<w:sdt><w:sdtContent><w:p/><w:tbl/><w:tr data-case="cell-row"/><w:tc data-case="cell-cell"/></w:sdtContent></w:sdt>"#,
            &["w".to_owned()],
        )
        .expect("cell control is within the nesting limit")
        .expect("cell control parses");
        assert!(matches!(cell.content[0], SdtContent::Paragraph(_)));
        assert!(matches!(cell.content[1], SdtContent::Table(_)));
        assert_eq!(
            cell.content[2],
            SdtContent::RawXml(br#"<w:tr data-case="cell-row"/>"#.to_vec())
        );
        assert_eq!(
            cell.content[3],
            SdtContent::RawXml(br#"<w:tc data-case="cell-cell"/>"#.to_vec())
        );
    }

    #[test]
    fn public_standalone_control_parser_keeps_the_union_child_contract() {
        let cases = [
            ("paragraph", "<w:p><w:r><w:t>paragraph</w:t></w:r></w:p>"),
            ("table", "<w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>"),
            ("row", "<w:tr><w:tc><w:p/></w:tc></w:tr>"),
            ("cell", "<w:tc><w:p/></w:tc>"),
            ("run", "<w:r><w:t>run</w:t></w:r>"),
        ];
        for (label, child) in cases {
            let xml =
                format!(r#"<w:sdt xmlns:w="{W_NS}"><w:sdtContent>{child}</w:sdtContent></w:sdt>"#);
            let control = parse_standalone_control(&xml);
            let modeled = matches!(
                (label, &control.content[0]),
                ("paragraph", SdtContent::Paragraph(_))
                    | ("table", SdtContent::Table(_))
                    | ("row", SdtContent::Row(_))
                    | ("cell", SdtContent::Cell(_))
                    | ("run", SdtContent::Run(_))
            );
            assert!(modeled, "standalone {label} child remains modeled");

            let mut writer = Writer::new(Vec::new());
            control.to_xml(&mut writer).expect("control serializes");
            assert_eq!(
                String::from_utf8(writer.into_inner()).expect("UTF-8"),
                xml,
                "standalone {label} control round-trips exactly"
            );
        }
    }

    #[test]
    fn table_traversal_sees_rows_cells_and_paragraphs_inside_controls_once() {
        let table = parse_table(
            r#"<w:sdt><w:sdtContent><w:tr><w:sdt><w:sdtContent><w:tc><w:sdt><w:sdtContent><w:p><w:sdt><w:sdtContent><w:r><w:t>once</w:t></w:r></w:sdtContent></w:sdt></w:p></w:sdtContent></w:sdt></w:tc></w:sdtContent></w:sdt></w:tr></w:sdtContent></w:sdt>"#,
        );
        let rows = table.rows();
        assert_eq!(rows.len(), 1);
        let cells = rows[0].cells();
        assert_eq!(cells.len(), 1);
        let paragraphs = cells[0].paragraphs();
        assert_eq!(paragraphs.len(), 1);
        let runs = paragraphs[0].runs();
        assert_eq!(runs.len(), 1);
        assert_eq!(runs[0].text(), "once");
    }

    #[test]
    fn content_control_cell_preserves_child_binding_declared_on_cell() {
        let table = parse_table(
            r#"<w:tr><w:sdt><w:sdtContent><w:tc xmlns:ext="urn:producer"><w:tcPr><ext:property ext:value="kept"/></w:tcPr><w:p/></w:tc></w:sdtContent></w:sdt></w:tr>"#,
        );
        let cell = table.rows()[0].cells()[0];
        let properties = cell.properties.as_ref().expect("cell properties parse");
        assert_eq!(
            properties.extra_xml,
            vec![(
                0,
                br#"<ext:property ext:value="kept" xmlns:ext="urn:producer"/>"#.to_vec(),
            )]
        );

        let mut writer = Writer::new(Vec::new());
        table.to_xml(&mut writer).expect("serializes");
        let saved = String::from_utf8(writer.into_inner()).expect("UTF-8");
        assert!(
            saved.contains(r#"<ext:property ext:value="kept" xmlns:ext="urn:producer"/>"#),
            "content-control cell child keeps its cell-local namespace binding: {saved}"
        );
    }

    #[test]
    fn run_control_keeps_comment_anchor_and_hyperlink_boundaries() {
        let document = parse_document(
            r#"<w:p><w:hyperlink w:anchor="target"><w:r><w:t>linked</w:t></w:r></w:hyperlink><w:commentRangeStart w:id="7"/><w:sdt><w:sdtContent><w:r><w:t>controlled</w:t></w:r></w:sdtContent></w:sdt><w:commentRangeEnd w:id="7"/><w:r><w:t>after</w:t></w:r></w:p>"#,
        );
        let saved = String::from_utf8(document.to_xml().expect("serializes")).expect("UTF-8");
        let hyperlink_end = saved.find("</w:hyperlink>").expect("hyperlink end");
        let anchor_start = saved.find("<w:commentRangeStart").expect("anchor start");
        let control = saved.find("<w:sdt>").expect("content control");
        let anchor_end = saved.find("<w:commentRangeEnd").expect("anchor end");
        assert!(hyperlink_end < anchor_start);
        assert!(anchor_start < control);
        assert!(control < anchor_end);

        let reopened = CT_Document::from_xml(saved.as_bytes()).expect("reopens");
        let paragraph = reopened.body.paragraphs().next().expect("paragraph");
        assert_eq!(paragraph.text(), "linkedcontrolledafter");
        assert_eq!(paragraph.content_controls.len(), 1);
        assert_eq!(paragraph.comment_ranges.len(), 2);
    }
}
