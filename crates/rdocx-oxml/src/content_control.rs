//! WordprocessingML structured document tags (`w:sdt`).

use std::io::Write;
use std::ops::Range;

use quick_xml::reader::NsReader;

use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer, XmlVersion};

use crate::error::{OxmlError, Result};
use crate::namespace::{W_NS, matches_local_name};
use crate::numbering::{local_namespace_overrides, namespace_bindings, word_prefixes_at};
use crate::properties::is_word_element;
use crate::raw_xml::{capture_element, capture_empty_element};
use crate::revision::CT_Revision;
use crate::table::{CT_Row, CT_Tbl, CT_Tc};
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
    source_prefixes: Vec<String>,
}

#[derive(Debug, Clone, PartialEq)]
struct TypeElement {
    control_type: SdtType,
    name: String,
    attributes: Vec<(String, String)>,
    children: Vec<u8>,
    namespace_bindings: Vec<(String, String)>,
}

struct ValueChild {
    name: Vec<u8>,
    value: Option<String>,
    display: Option<String>,
}

struct CheckboxChild {
    value: Option<Range<usize>>,
    font: Option<Range<usize>>,
    insert_attribute: usize,
}

impl CT_SdtPr {
    /// Resolve the bounded text, dropdown and Gregorian date value conventions.
    pub fn prepare_text_value(&self, value: &str) -> Result<(String, String, Option<String>)> {
        oxml_core::xml::reject_non_xml_characters("content-control value", value)?;
        match self.control_type {
            Some(SdtType::Date) => {
                let (year, month, day) = civil_date(value)?;
                let children = self.value_children()?;
                let metadata = |name: &[u8]| -> Result<Option<&str>> {
                    let mut matches = children.iter().filter(|child| child.name == name);
                    let first = matches.next();
                    if matches.next().is_some() {
                        return Err(OxmlError::InvalidValue(
                            "duplicate date metadata".to_owned(),
                        ));
                    }
                    Ok(first.and_then(|child| child.value.as_deref()))
                };
                if !matches!(metadata(b"calendar")?, None | Some("gregorian")) {
                    return Err(OxmlError::InvalidValue(
                        "unsupported date calendar".to_owned(),
                    ));
                }
                let display = match metadata(b"dateFormat")? {
                    Some("yyyy-MM-dd") => value.to_owned(),
                    Some("dd/MM/yyyy") => format!("{day:02}/{month:02}/{year:04}"),
                    _ => {
                        return Err(OxmlError::InvalidValue(
                            "unsupported or missing date format".to_owned(),
                        ));
                    }
                };
                let full_date = format!("{value}T00:00:00Z");
                let bound = match metadata(b"storeMappedDataAs")? {
                    None | Some("text") => display.clone(),
                    Some("date") => value.to_owned(),
                    Some("dateTime") => full_date.clone(),
                    _ => {
                        return Err(OxmlError::InvalidValue(
                            "unsupported date storage mapping".to_owned(),
                        ));
                    }
                };
                Ok((display, bound, Some(full_date)))
            }
            Some(SdtType::DropDownList | SdtType::ComboBox) => {
                let label = if self.control_type == Some(SdtType::ComboBox) {
                    "combo box"
                } else {
                    "dropdown"
                };
                let children = self.value_children()?;
                let items: Vec<_> = children
                    .iter()
                    .filter(|child| child.name == b"listItem")
                    .collect();
                let choices: Vec<_> = items
                    .iter()
                    .filter_map(|item| item.display.as_deref().or(item.value.as_deref()))
                    .collect();
                let mut selected = None;
                for item in items {
                    let display = item.display.as_deref().or(item.value.as_deref());
                    if item.value.as_deref() == Some(value) || display == Some(value) {
                        let pair = match (display, item.value.as_deref()) {
                            (Some(display), Some(bound)) => (display.to_owned(), bound.to_owned()),
                            _ => {
                                return Err(OxmlError::InvalidValue(format!(
                                    "{label} item has no value, choices: {choices:?}"
                                )));
                            }
                        };
                        if selected.as_ref().is_some_and(|previous| previous != &pair) {
                            return Err(OxmlError::InvalidValue(format!(
                                "ambiguous {label} value, choices: {choices:?}"
                            )));
                        }
                        selected = Some(pair);
                    }
                }
                let (display, bound) = selected.ok_or_else(|| {
                    OxmlError::InvalidValue(format!(
                        "unknown {label} value {value:?}, choices: {choices:?}"
                    ))
                })?;
                Ok((display, bound, None))
            }
            _ => Ok((value.to_owned(), value.to_owned(), None)),
        }
    }

    /// Remove qualified placeholder metadata and optionally set a selected date.
    pub fn set_value_metadata(&mut self, full_date: Option<&str>) -> Result<()> {
        let date_attribute = if let Some(value) = full_date {
            if self.control_type != Some(SdtType::Date)
                || value.len() != 20
                || !value.ends_with("T00:00:00Z")
            {
                return Err(OxmlError::InvalidValue(
                    "fullDate must be a selected Gregorian UTC midnight".to_owned(),
                ));
            }
            civil_date(&value[..10])?;
            let element = self.type_element.as_ref().ok_or_else(|| {
                OxmlError::InvalidValue("date type payload is missing".to_owned())
            })?;
            let mut start = BytesStart::new(element.name.as_str());
            for (name, value) in &element.attributes {
                start.push_attribute((name.as_str(), value.as_str()));
            }
            let prefixes = word_prefixes_at(&start, &self.source_prefixes)?;
            let matching: Vec<_> = element
                .attributes
                .iter()
                .enumerate()
                .filter(|(_, (name, _))| is_word_attribute(name.as_bytes(), b"fullDate", &prefixes))
                .map(|(index, _)| index)
                .collect();
            if matching.len() > 1 {
                return Err(OxmlError::InvalidValue(
                    "duplicate qualified fullDate attribute".to_owned(),
                ));
            }
            let prefix = prefixes
                .iter()
                .find(|prefix| !prefix.is_empty() && !prefix.starts_with('\0'))
                .cloned();
            Some((matching.first().copied(), prefix))
        } else {
            None
        };
        let type_name = self
            .type_element
            .as_ref()
            .map(|element| -> Result<Option<String>> {
                let mut start = BytesStart::new(element.name.as_str());
                for (name, value) in &element.attributes {
                    start.push_attribute((name.as_str(), value.as_str()));
                }
                let prefixes = word_prefixes_at(&start, &self.source_prefixes)?;
                let local = element
                    .control_type
                    .element_name()
                    .split_once(':')
                    .unwrap()
                    .1;
                if !matches!(
                    element.control_type,
                    SdtType::CheckBox | SdtType::RepeatingSection | SdtType::RepeatingSectionItem
                ) && !is_word_element(element.name.as_bytes(), local.as_bytes(), &prefixes)
                {
                    let prefix = prefixes
                        .iter()
                        .find(|prefix| !prefix.starts_with('\0'))
                        .ok_or_else(|| {
                            OxmlError::InvalidValue(
                                "selected type has no qualified Word prefix".to_owned(),
                            )
                        })?;
                    Ok(Some(if prefix.is_empty() {
                        local.to_owned()
                    } else {
                        format!("{prefix}:{local}")
                    }))
                } else {
                    Ok(None)
                }
            })
            .transpose()?
            .flatten();
        let mut remove = Vec::new();
        for (index, slot) in self.slots.iter().enumerate() {
            if let PropertySlot::Raw(raw) = slot {
                let mut reader = Reader::from_reader(raw.as_slice());
                if let Event::Start(start) | Event::Empty(start) = reader.read_event()? {
                    let prefixes = word_prefixes_at(&start, &self.source_prefixes)?;
                    if is_word_element(start.name().as_ref(), b"showingPlcHdr", &prefixes) {
                        remove.push(index);
                    }
                }
            }
        }
        if let (Some(value), Some((existing, prefix))) = (full_date, date_attribute) {
            let element = self.type_element.as_mut().unwrap();
            if let Some(index) = existing {
                element.attributes[index].1 = value.to_owned();
            } else if let Some(prefix) = prefix {
                element
                    .attributes
                    .push((format!("{prefix}:fullDate"), value.to_owned()));
            } else {
                let mut prefix = "rdocxW".to_owned();
                while element
                    .namespace_bindings
                    .iter()
                    .any(|(bound, _)| bound == &prefix)
                    || element
                        .attributes
                        .iter()
                        .any(|(name, _)| name == &format!("xmlns:{prefix}"))
                {
                    prefix.push('W');
                }
                element
                    .attributes
                    .push((format!("xmlns:{prefix}"), W_NS.to_owned()));
                element
                    .namespace_bindings
                    .push((prefix.clone(), W_NS.to_owned()));
                element
                    .attributes
                    .push((format!("{prefix}:fullDate"), value.to_owned()));
            }
        }
        if let Some(name) = type_name {
            self.type_element.as_mut().unwrap().name = name;
        }
        for index in remove.into_iter().rev() {
            self.slots.remove(index);
        }
        Ok(())
    }

    fn type_children_xml(&self) -> Result<(Vec<u8>, usize)> {
        let element = self.type_element.as_ref().ok_or_else(|| {
            OxmlError::InvalidValue("content-control type payload is missing".to_owned())
        })?;
        let mut bindings = element.namespace_bindings.clone();
        for prefix in &self.source_prefixes {
            if !prefix.starts_with('\0') && !bindings.iter().any(|(bound, _)| bound == prefix) {
                bindings.push((prefix.clone(), W_NS.to_owned()));
            }
        }
        let mut start = BytesStart::new("scope");
        for (prefix, namespace) in bindings {
            let name = if prefix.is_empty() {
                "xmlns".to_owned()
            } else {
                format!("xmlns:{prefix}")
            };
            start.push_attribute((name.as_str(), namespace.as_str()));
        }
        let mut writer = Writer::new(Vec::new());
        writer.write_event(Event::Start(start))?;
        let offset = writer.get_ref().len();
        writer.get_mut().extend_from_slice(&element.children);
        writer.write_event(Event::End(BytesEnd::new("scope")))?;
        Ok((writer.into_inner(), offset))
    }

    fn value_children(&self) -> Result<Vec<ValueChild>> {
        for slot in &self.slots {
            if let PropertySlot::Raw(raw) = slot {
                let mut reader = Reader::from_reader(raw.as_slice());
                if let Event::Start(start) | Event::Empty(start) = reader.read_event()? {
                    let prefixes = word_prefixes_at(&start, &self.source_prefixes)?;
                    let local = start.local_name();
                    if is_supported_type_element(start.name().as_ref(), local.as_ref(), &prefixes)
                        && SdtType::from_element(local.as_ref()).is_some()
                    {
                        return Err(OxmlError::InvalidValue(
                            "conflicting content-control type metadata".to_owned(),
                        ));
                    }
                }
            }
        }
        let (xml, _) = self.type_children_xml()?;
        let mut reader = NsReader::from_reader(xml.as_slice());
        let mut depth = 0usize;
        let mut children = Vec::new();
        let is_word_namespace = |namespace: quick_xml::name::ResolveResult<'_>| {
            matches!(namespace, quick_xml::name::ResolveResult::Bound(ns)
                if std::str::from_utf8(ns.as_ref()).ok()
                    .and_then(|uri| quick_xml::escape::unescape(uri).ok())
                    .is_some_and(|uri| uri == W_NS))
        };
        loop {
            let (namespace, event) = reader.read_resolved_event()?;
            let qualified = is_word_namespace(namespace);
            let (child, empty) = match event {
                Event::Start(child) => (child, false),
                Event::Empty(child) => (child, true),
                Event::End(_) => {
                    depth = depth.saturating_sub(1);
                    continue;
                }
                Event::Eof => return Ok(children),
                _ => continue,
            };
            if depth == 1 && qualified {
                let mut record = ValueChild {
                    name: child.local_name().as_ref().to_vec(),
                    value: None,
                    display: None,
                };
                for attribute in child.attributes() {
                    let attribute = attribute?;
                    let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
                    if !is_word_namespace(namespace) {
                        continue;
                    }
                    let destination = match local.as_ref() {
                        b"val" if record.name != b"listItem" => &mut record.value,
                        b"value" if record.name == b"listItem" => &mut record.value,
                        b"displayText" => &mut record.display,
                        _ => continue,
                    };
                    if destination.is_some() {
                        return Err(OxmlError::InvalidValue(
                            "duplicate qualified value metadata attribute".to_owned(),
                        ));
                    }
                    *destination = Some(
                        attribute
                            .decoded_and_normalized_value(XmlVersion::Implicit1_0, child.decoder())?
                            .into_owned(),
                    );
                }
                children.push(record);
            }
            if !empty {
                depth += 1;
            }
        }
    }

    /// The selected checkbox glyph and its physical state font travel together.
    pub fn checkbox_state(&self, checked: bool) -> Result<(char, String)> {
        let name = if checked {
            b"checkedState".as_slice()
        } else {
            b"uncheckedState".as_slice()
        };
        let child = self.checkbox_child(name)?;
        let element = self.type_element.as_ref().ok_or_else(|| {
            OxmlError::InvalidValue("checkbox type payload is missing".to_owned())
        })?;
        let read = |range: &Range<usize>| -> Result<String> {
            Ok(
                quick_xml::escape::unescape(std::str::from_utf8(&element.children[range.clone()])?)
                    .map_err(|error| {
                        OxmlError::InvalidValue(format!("checkbox state decode failed: {error}"))
                    })?
                    .into_owned(),
            )
        };
        let code = child
            .as_ref()
            .and_then(|child| child.value.as_ref())
            .map(read)
            .transpose()?
            .unwrap_or_else(|| if checked { "2612" } else { "2610" }.to_owned());
        let glyph = u32::from_str_radix(&code, 16)
            .ok()
            .and_then(char::from_u32)
            .ok_or_else(|| {
                OxmlError::InvalidValue(format!(
                    "checkbox state glyph {code:?} is not a Unicode scalar"
                ))
            })?;
        oxml_core::xml::reject_non_xml_characters("checkbox state glyph", &glyph.to_string())?;
        let font = child
            .as_ref()
            .and_then(|child| child.font.as_ref())
            .map(read)
            .transpose()?
            .unwrap_or_else(|| "MS Gothic".to_owned());
        oxml_core::xml::reject_non_xml_characters("checkbox state font", &font)?;
        if font.trim().is_empty() {
            return Err(OxmlError::InvalidValue(
                "checkbox state font is empty".to_owned(),
            ));
        }
        Ok((glyph, font))
    }

    /// Change only the qualified checked value, preserving all other raw bytes.
    pub fn set_checkbox_checked(&mut self, checked: bool) -> Result<()> {
        let child = self.checkbox_child(b"checked")?;
        let element = self.type_element.as_mut().ok_or_else(|| {
            OxmlError::InvalidValue("checkbox type payload is missing".to_owned())
        })?;
        let value = if checked {
            b"1".as_slice()
        } else {
            b"0".as_slice()
        };
        if let Some(child) = child {
            if let Some(range) = child.value {
                element.children.splice(range, value.iter().copied());
            } else {
                let mut prefix = "rdocxW14".to_owned();
                let mut suffix = 0usize;
                while element
                    .namespace_bindings
                    .iter()
                    .any(|(bound, _)| bound == &prefix)
                    || element
                        .children
                        .windows(prefix.len())
                        .any(|bytes| bytes == prefix.as_bytes())
                {
                    suffix += 1;
                    prefix = format!("rdocxW14{suffix}");
                }
                let attribute = format!(
                    " xmlns:{prefix}=\"http://schemas.microsoft.com/office/word/2010/wordml\" {prefix}:val=\"{}\"",
                    u8::from(checked)
                );
                element.children.splice(
                    child.insert_attribute..child.insert_attribute,
                    attribute.bytes(),
                );
            }
        } else {
            let checked_xml = format!(
                "<rdocxW14:checked xmlns:rdocxW14=\"http://schemas.microsoft.com/office/word/2010/wordml\" rdocxW14:val=\"{}\"/>",
                u8::from(checked)
            );
            element.children.splice(0..0, checked_xml.bytes());
        }
        Ok(())
    }

    fn checkbox_child(&self, wanted: &[u8]) -> Result<Option<CheckboxChild>> {
        const W14: &[u8] = b"http://schemas.microsoft.com/office/word/2010/wordml";
        if self.control_type != Some(SdtType::CheckBox) {
            return Err(OxmlError::InvalidValue(
                "only a checkbox has a checked state".to_owned(),
            ));
        }
        let (xml, offset) = self.type_children_xml()?;
        let mut reader = NsReader::from_reader(xml.as_slice());
        let mut depth = 0usize;
        let mut found = None;
        loop {
            let before = reader.buffer_position() as usize;
            let (namespace, event) = reader.read_resolved_event()?;
            let qualified = matches!(namespace, quick_xml::name::ResolveResult::Bound(ns) if ns.as_ref() == W14);
            match event {
                Event::Start(child) | Event::Empty(child) => {
                    let empty = xml.get(reader.buffer_position() as usize - 2) == Some(&b'/');
                    if depth == 1 && qualified && child.local_name().as_ref() == wanted {
                        if found.is_some() {
                            return Err(OxmlError::InvalidValue(
                                "duplicate qualified checkbox state child".to_owned(),
                            ));
                        }
                        let mut record = CheckboxChild {
                            value: None,
                            font: None,
                            insert_attribute: reader.buffer_position() as usize
                                - offset
                                - if empty { 2 } else { 1 },
                        };
                        for attribute in child.attributes() {
                            let attribute = attribute?;
                            let (namespace, local) =
                                reader.resolver().resolve_attribute(attribute.key);
                            if !matches!(namespace, quick_xml::name::ResolveResult::Bound(ns) if ns.as_ref() == W14)
                            {
                                continue;
                            }
                            let destination = match local.as_ref() {
                                b"val" => &mut record.value,
                                b"font" => &mut record.font,
                                _ => continue,
                            };
                            if destination.is_some() {
                                return Err(OxmlError::InvalidValue(
                                    "duplicate qualified checkbox state attribute".to_owned(),
                                ));
                            }
                            let relative = attribute.value.as_ptr() as usize
                                - child.as_ref().as_ptr() as usize;
                            let begin = before + 1 + relative - offset;
                            *destination = Some(begin..begin + attribute.value.len());
                        }
                        found = Some(record);
                    }
                    if !empty {
                        depth += 1;
                    }
                }
                Event::End(_) => depth = depth.saturating_sub(1),
                Event::Eof => return Ok(found),
                _ => {}
            }
        }
    }
    fn from_xml_with_prefixes(
        reader: &mut Reader<&[u8]>,
        start: &BytesStart<'_>,
        inherited: &[String],
    ) -> Result<Self> {
        let prefixes = word_prefixes_at(start, inherited)?;
        let mut properties = Self {
            extra_attributes: capture_attributes(start, &prefixes, &[])?,
            source_prefixes: prefixes.clone(),
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
                                name: std::str::from_utf8(child.name().as_ref())?.to_owned(),
                                attributes: capture_attributes(&child, &child_prefixes, &[])?,
                                children: children.to_vec(),
                                namespace_bindings: namespace_bindings(&child_prefixes),
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
                                name: std::str::from_utf8(child.name().as_ref())?.to_owned(),
                                attributes: capture_attributes(&child, &child_prefixes, &[])?,
                                children: Vec::new(),
                                namespace_bindings: namespace_bindings(&child_prefixes),
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
        let name = self
            .type_element
            .as_ref()
            .filter(|element| element.control_type == control_type)
            .map_or_else(
                || control_type.element_name(),
                |element| element.name.as_str(),
            );
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
#[derive(Debug, PartialEq)]
pub enum SdtContent {
    Paragraph(CT_P),
    Table(CT_Tbl),
    Row(CT_Row),
    Cell(CT_Tc),
    Run(CT_R),
    ContentControl(CT_Sdt),
    RawXml(Vec<u8>),
}

impl Clone for SdtContent {
    fn clone(&self) -> Self {
        if let Self::ContentControl(control) = self {
            return Self::ContentControl(control.clone());
        }
        clone_non_control_sdt_content(self)
    }
}

// Keep the large modeled leaf variants out of the recursive control frame.
// This helper remains exhaustive even though its current caller selects leaves.
#[inline(never)]
fn clone_non_control_sdt_content(content: &SdtContent) -> SdtContent {
    match content {
        SdtContent::Paragraph(paragraph) => SdtContent::Paragraph(paragraph.clone()),
        SdtContent::Table(table) => SdtContent::Table(table.clone()),
        SdtContent::Row(row) => SdtContent::Row(row.clone()),
        SdtContent::Cell(cell) => SdtContent::Cell(cell.clone()),
        SdtContent::Run(run) => SdtContent::Run(run.clone()),
        SdtContent::ContentControl(control) => SdtContent::ContentControl(control.clone()),
        SdtContent::RawXml(raw) => SdtContent::RawXml(raw.clone()),
    }
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
    pub(crate) fn owns(self, local: &[u8]) -> bool {
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
    pub fn from_xml(reader: &mut Reader<&[u8]>, start: &BytesStart<'_>) -> Result<Self> {
        let prefixes = word_prefixes_at(start, &["w".to_owned()])?;
        crate::document::admit_word_reader(
            reader,
            &prefixes,
            b"sdt",
            Some(crate::namespace::W_NS),
        )?;
        Self::from_xml_with_prefixes(reader, start, &["w".to_owned()], SdtOwner::Standalone)
    }

    pub(crate) fn from_body_raw(raw: &[u8], inherited: &[String]) -> Option<Self> {
        Self::from_raw_with_context(raw, inherited, SdtOwner::Body)
    }

    pub(crate) fn from_table_raw(raw: &[u8], inherited: &[String]) -> Option<Self> {
        Self::from_raw_with_context(raw, inherited, SdtOwner::Table)
    }

    pub(crate) fn from_row_raw(raw: &[u8], inherited: &[String]) -> Option<Self> {
        Self::from_raw_with_context(raw, inherited, SdtOwner::Row)
    }

    pub(crate) fn from_cell_raw(raw: &[u8], inherited: &[String]) -> Option<Self> {
        Self::from_raw_with_context(raw, inherited, SdtOwner::Cell)
    }

    pub(crate) fn from_inline_raw(raw: &[u8], inherited: &[String]) -> Option<Self> {
        Self::from_raw_with_context(raw, inherited, SdtOwner::Inline)
    }

    /// Whether a preserved control is admitted by the complete typed parser.
    #[doc(hidden)]
    pub fn story_raw_is_typed(raw: &[u8], inherited: &[String], owner: StorySdtOwner) -> bool {
        if crate::document::admit_word_xml(raw, inherited, None).is_err() {
            return false;
        }
        Self::from_raw_with_context(raw, inherited, owner.into()).is_some()
    }

    fn from_raw_with_context(raw: &[u8], inherited: &[String], owner: SdtOwner) -> Option<Self> {
        let mut reader = Reader::from_reader(raw);
        reader.config_mut().trim_text(false);
        let mut buffer = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer) {
                Ok(Event::Start(start)) if matches_local_name(start.name().as_ref(), b"sdt") => {
                    return Self::from_xml_with_prefixes(&mut reader, &start, inherited, owner)
                        .ok();
                }
                Ok(Event::Eof) | Err(_) => return None,
                Ok(_) => {}
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

    /// Stored run originals are cloned even after the current run is edited.
    pub(crate) fn original_resource_runs(&self) -> impl Iterator<Item = &CT_R> {
        self.inline_run_sources
            .iter()
            .map(|source| &source.original)
    }

    pub(crate) fn to_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        crate::document::admit_word_model(crate::document::WordOwner::Control(self))?;
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

// Keep by-value leaf model construction out of each recursive control frame.
// This boundary addresses the measured ordinary-stack table/control overflow.
#[inline(never)]
fn append_modeled_content_child(
    reader: &mut Reader<&[u8]>,
    child: &BytesStart<'_>,
    inherited: &[String],
    empty: bool,
    content: &mut Vec<SdtContent>,
    inline_run_sources: &mut Vec<InlineRunSource>,
) -> Result<()> {
    let prefixes = word_prefixes_at(child, inherited)?;
    match (child.local_name().as_ref(), empty) {
        (b"p", false) => content.push(SdtContent::Paragraph(
            CT_P::from_xml_with_prefixes_and_root(reader, &prefixes, Some(child))?,
        )),
        (b"p", true) => content.push(SdtContent::Paragraph(CT_P::from_empty_root(
            child, &prefixes,
        )?)),
        (b"tbl", false) => {
            let bindings = local_namespace_overrides(child, inherited)?;
            let mut table =
                CT_Tbl::from_xml_with_prefixes_and_owner_bindings(reader, &prefixes, &bindings)?;
            table.retain_root_attributes(child, &prefixes)?;
            content.push(SdtContent::Table(table));
        }
        (b"tbl", true) => {
            let mut table = CT_Tbl::new();
            table.retain_root_attributes(child, &prefixes)?;
            content.push(SdtContent::Table(table));
        }
        (b"tr", false) => {
            let bindings = local_namespace_overrides(child, inherited)?;
            content.push(SdtContent::Row(
                CT_Row::from_xml_with_prefixes_and_owner_bindings(
                    reader,
                    &prefixes,
                    &bindings,
                    Some(child),
                )?,
            ));
        }
        (b"tr", true) => content.push(SdtContent::Row(CT_Row::from_empty_root(child, &prefixes)?)),
        (b"tc", false) => {
            let bindings = local_namespace_overrides(child, inherited)?;
            let mut cell =
                CT_Tc::from_xml_with_prefixes_and_owner_bindings(reader, &prefixes, &bindings)?;
            cell.retain_root_attributes(child, &prefixes)?;
            content.push(SdtContent::Cell(cell));
        }
        (b"tc", true) => {
            let mut cell = CT_Tc {
                properties: None,
                content: Vec::new(),
                extra_xml: Vec::new(),
            };
            cell.retain_root_attributes(child, &prefixes)?;
            content.push(SdtContent::Cell(cell));
        }
        (b"r", _) => {
            let (run, raw_xml) = if empty {
                (
                    CT_R {
                        properties: None,
                        content: Vec::new(),
                        extra_xml: Vec::new(),
                        extra_xml_positions: Vec::new(),
                        alt_drawings: Vec::new(),
                    },
                    capture_empty_element(child)?,
                )
            } else {
                let raw = capture_element(reader, child)?;
                (parse_inline_run(&raw, inherited)?, raw)
            };
            inline_run_sources.push(InlineRunSource {
                content_index: content.len(),
                original: run.clone(),
                raw_xml,
            });
            content.push(SdtContent::Run(run));
        }
        _ => {
            return Err(OxmlError::InvalidValue(
                "unsupported modeled content-control child".into(),
            ));
        }
    }
    Ok(())
}

// Variant construction reserves space for the largest modeled leaf. Keep it
// outside the recursive control parser even for retained and nested children.
#[inline(never)]
fn append_retained_sdt_content(
    content: &mut Vec<SdtContent>,
    control: Option<CT_Sdt>,
    raw: Vec<u8>,
) {
    content.push(match control {
        Some(control) => SdtContent::ContentControl(control),
        None => SdtContent::RawXml(raw),
    });
}

fn parse_content(
    reader: &mut Reader<&[u8]>,
    inherited: &[String],
    revisions: &mut Vec<(usize, CT_Revision)>,
    inline_run_sources: &mut Vec<InlineRunSource>,
    owner: SdtOwner,
) -> Result<Vec<SdtContent>> {
    let mut content = Vec::new();
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(child)) => {
                let prefixes = word_prefixes_at(&child, inherited)?;
                let owned_word_child = [
                    b"p".as_slice(),
                    b"tbl".as_slice(),
                    b"tr".as_slice(),
                    b"tc".as_slice(),
                    b"r".as_slice(),
                ]
                .into_iter()
                .find(|local| is_word_element(child.name().as_ref(), local, &prefixes));
                if owned_word_child.is_some_and(|local| !owner.owns(local)) {
                    append_retained_sdt_content(
                        &mut content,
                        None,
                        capture_element(reader, &child)?,
                    );
                } else if owned_word_child.is_some() {
                    append_modeled_content_child(
                        reader,
                        &child,
                        inherited,
                        false,
                        &mut content,
                        inline_run_sources,
                    )?;
                } else if is_word_element(child.name().as_ref(), b"sdt", &prefixes) {
                    let raw = capture_element(reader, &child)?;
                    let parsed = CT_Sdt::from_raw_with_context(&raw, &prefixes, owner);
                    append_retained_sdt_content(&mut content, parsed, raw);
                } else {
                    let raw = capture_element(reader, &child)?;
                    if let Some(revision) = CT_Revision::from_raw(raw.clone(), &prefixes) {
                        revisions.push((content.len(), revision));
                    }
                    append_retained_sdt_content(&mut content, None, raw);
                }
            }
            Ok(Event::Empty(child)) => {
                let prefixes = word_prefixes_at(&child, inherited)?;
                let owned_word_child = [
                    b"p".as_slice(),
                    b"tbl".as_slice(),
                    b"tr".as_slice(),
                    b"tc".as_slice(),
                    b"r".as_slice(),
                ]
                .into_iter()
                .find(|local| is_word_element(child.name().as_ref(), local, &prefixes));
                if owned_word_child.is_some_and(|local| !owner.owns(local)) {
                    append_retained_sdt_content(&mut content, None, capture_empty_element(&child)?);
                } else if owned_word_child.is_some() {
                    append_modeled_content_child(
                        reader,
                        &child,
                        inherited,
                        true,
                        &mut content,
                        inline_run_sources,
                    )?;
                } else {
                    let raw = capture_empty_element(&child)?;
                    if let Some(revision) = CT_Revision::from_raw(raw.clone(), &prefixes) {
                        revisions.push((content.len(), revision));
                    }
                    append_retained_sdt_content(&mut content, None, raw);
                }
            }
            Ok(Event::End(end)) if matches_local_name(end.name().as_ref(), b"sdtContent") => break,
            Ok(Event::Eof) => {
                return Err(OxmlError::MissingElement("w:sdtContent end".to_owned()));
            }
            Ok(event) => append_retained_sdt_content(&mut content, None, event_to_raw(event)?),
            Err(error) => return Err(error.into()),
        }
        buffer.clear();
    }
    Ok(content)
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

fn civil_date(value: &str) -> Result<(u16, u8, u8)> {
    let bytes = value.as_bytes();
    let invalid = || {
        OxmlError::InvalidValue(format!(
            "expected a Gregorian YYYY-MM-DD date, got {value:?}"
        ))
    };
    if bytes.len() != 10
        || bytes[4] != b'-'
        || bytes[7] != b'-'
        || bytes
            .iter()
            .enumerate()
            .any(|(index, byte)| index != 4 && index != 7 && !byte.is_ascii_digit())
    {
        return Err(invalid());
    }
    let year: u16 = value[..4].parse().map_err(|_| invalid())?;
    let month: u8 = value[5..7].parse().map_err(|_| invalid())?;
    let day: u8 = value[8..].parse().map_err(|_| invalid())?;
    let leap = year.is_multiple_of(400) || year.is_multiple_of(4) && !year.is_multiple_of(100);
    let maximum = match month {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        2 if leap => 29,
        2 => 28,
        _ => return Err(invalid()),
    };
    if year == 0 || day == 0 || day > maximum {
        return Err(invalid());
    }
    Ok((year, month, day))
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
    if local != b"checkbox" && is_word_element(name, local, prefixes) {
        return true;
    }
    let prefix = match name.iter().position(|byte| *byte == b':') {
        Some(separator) => &name[..separator],
        None if local == b"checkbox" => b"".as_slice(),
        None => return false,
    };
    prefixes.iter().any(|binding| {
        let Some(rest) = binding.strip_prefix('\0') else {
            return false;
        };
        let Some((bound_prefix, namespace)) = rest.split_once('\0') else {
            return false;
        };
        bound_prefix.as_bytes() == prefix
            && if local == b"checkbox" {
                namespace == "http://schemas.microsoft.com/office/word/2010/wordml"
            } else {
                matches!(
                    namespace,
                    "http://schemas.microsoft.com/office/word/2010/wordml"
                        | "http://schemas.microsoft.com/office/word/2012/wordml"
                )
            }
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

/// The actual sdtPr parser models only these property leaves/attributes.
pub(crate) fn source_modeled_attributes(
    parent: &[u8],
    local: &[u8],
) -> Option<&'static [&'static [u8]]> {
    if parent != b"sdtPr" {
        return None;
    }
    match local {
        b"alias" | b"tag" | b"id" => Some(&[b"val"]),
        b"dataBinding" => Some(&[b"prefixMappings", b"xpath", b"storeItemID"]),
        // Type element attributes and payload are explicitly retained source state.
        _ if SdtType::from_element(local).is_some() => Some(&[]),
        _ => None,
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
    fn fx249_review_d1_local_namespace_entities_are_decoded_once() {
        let make = |kind: &str, children: &str| {
            parse_standalone_control(&format!(
                r#"<w:sdt xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:sdtPr><w:{kind}>{children}</w:{kind}></w:sdtPr><w:sdtContent><w:p/></w:sdtContent></w:sdt>"#
            ))
        };
        for child in [
            r#"<a:dateFormat xmlns:a="http://schemas.openxmlformats.org/wordprocessingml/2006/mai&#110;" a:val="yyyy-MM-dd"/>"#,
            r#"<w:dateFormat xmlns:a="http://schemas.openxmlformats.org/wordprocessingml/2006/mai&#110;" a:val="yyyy-MM-dd"/>"#,
            r#"<dateFormat xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/mai&#110;" xmlns:a="http://schemas.openxmlformats.org/wordprocessingml/2006/mai&#110;" a:val="yyyy-MM-dd"/>"#,
        ] {
            assert!(
                make("date", child)
                    .properties
                    .unwrap()
                    .prepare_text_value("2026-10-31")
                    .is_ok(),
                "{child}"
            );
        }
        for child in [
            r#"<a:listItem xmlns:a="http://schemas.openxmlformats.org/wordprocessingml/2006/mai&#110;" a:displayText="Two" a:value="two"/>"#,
            r#"<w:listItem xmlns:a="http://schemas.openxmlformats.org/wordprocessingml/2006/mai&#110;" a:displayText="Two" a:value="two"/>"#,
        ] {
            assert_eq!(
                make("dropDownList", child)
                    .properties
                    .unwrap()
                    .prepare_text_value("two")
                    .unwrap(),
                ("Two".into(), "two".into(), None)
            );
        }
        for child in [
            r#"<a:dateFormat xmlns:a="http://schemas.openxmlformats.org/wordprocessingml/2006/mai&amp;#110;" a:val="yyyy-MM-dd"/>"#,
            r#"<w:dateFormat xmlns:a="http://schemas.openxmlformats.org/wordprocessingml/2006/mai&amp;#110;" a:val="yyyy-MM-dd"/>"#,
        ] {
            assert!(
                make("date", child)
                    .properties
                    .unwrap()
                    .prepare_text_value("2026-10-31")
                    .is_err(),
                "decoded twice: {child}"
            );
        }
        let opaque = r#"<a:dateFormat xmlns:a="http://schemas.openxmlformats.org/wordprocessingml/2006/mai&amp;#110;" a:val="foreign"/>"#;
        let mut control = make(
            "date",
            &format!(r#"{opaque}<w:dateFormat w:val="yyyy-MM-dd"/>"#),
        );
        assert!(
            control
                .properties
                .as_ref()
                .unwrap()
                .prepare_text_value("2026-10-31")
                .is_ok()
        );
        control
            .properties
            .as_mut()
            .unwrap()
            .set_value_metadata(Some("2026-10-31T00:00:00Z"))
            .unwrap();
        let mut saved = Vec::new();
        control.to_xml(&mut Writer::new(&mut saved)).unwrap();
        assert!(String::from_utf8(saved).unwrap().contains(opaque));
    }

    #[test]
    fn imported_type_names_keep_namespace_identity_during_serialization() {
        let cases = [
            (W_NS, "date", "w", SdtType::Date),
            (
                "http://schemas.microsoft.com/office/word/2010/wordml",
                "checkbox",
                "w14",
                SdtType::CheckBox,
            ),
            (
                "http://schemas.microsoft.com/office/word/2012/wordml",
                "repeatingSection",
                "w15",
                SdtType::RepeatingSection,
            ),
            (
                "http://schemas.microsoft.com/office/word/2012/wordml",
                "repeatingSectionItem",
                "w15",
                SdtType::RepeatingSectionItem,
            ),
        ];
        let opaque = r#"<foreign:kept xmlns:foreign="urn:opaque" foreign:val="raw"/>"#;
        for (namespace, local, canonical, kind) in cases {
            for expanded in [false, true] {
                let attributes = format!(r#" xmlns:{canonical}="urn:opaque" q:custom="kept""#);
                let marker = if expanded {
                    format!("<q:{local}{attributes}>{opaque}</q:{local}>")
                } else {
                    format!("<q:{local}{attributes}/>")
                };
                let xml = format!(
                    r#"<w:sdt xmlns:w="{W_NS}" xmlns:q="{namespace}"><w:sdtPr>{marker}</w:sdtPr><w:sdtContent><w:p/></w:sdtContent></w:sdt>"#
                );
                let control = parse_standalone_control(&xml);
                assert_eq!(
                    control.properties.as_ref().unwrap().control_type,
                    Some(kind)
                );
                let mut saved = Vec::new();
                control.to_xml(&mut Writer::new(&mut saved)).unwrap();
                let mut reader = NsReader::from_reader(saved.as_slice());
                let mut found = 0;
                loop {
                    let (resolved, event) = reader.read_resolved_event().unwrap();
                    match event {
                        Event::Start(start) | Event::Empty(start)
                            if start.local_name().as_ref() == local.as_bytes() =>
                        {
                            assert!(
                                matches!(resolved, quick_xml::name::ResolveResult::Bound(value) if value.as_ref() == namespace.as_bytes()),
                                "type lost its namespace: {}",
                                String::from_utf8_lossy(&saved)
                            );
                            found += 1;
                        }
                        Event::Eof => break,
                        _ => {}
                    }
                }
                assert_eq!(found, 1);
                let saved = String::from_utf8(saved).unwrap();
                assert!(saved.contains(r#"q:custom="kept""#));
                if expanded {
                    assert!(saved.contains(opaque));
                }
                assert_eq!(
                    parse_standalone_control(&saved)
                        .properties
                        .unwrap()
                        .control_type,
                    Some(kind)
                );
            }
        }
    }

    #[test]
    fn fx249_review_d2_supported_extension_type_conflicts_refuse() {
        for kind in ["date", "dropDownList"] {
            let children = if kind == "date" {
                r#"<w:dateFormat w:val="yyyy-MM-dd"/>"#
            } else {
                r#"<w:listItem w:value="2026-10-31"/>"#
            };
            for conflict in [
                r#"<a:checkbox xmlns:a="http://schemas.microsoft.com/office/word/2010/wordml"/>"#,
                r#"<a:repeatingSection xmlns:a="http://schemas.microsoft.com/office/word/2012/wordml"/>"#,
                r#"<a:repeatingSectionItem xmlns:a="http://schemas.microsoft.com/office/word/2012/wordml"/>"#,
            ] {
                let control = parse_standalone_control(&format!(
                    r#"<w:sdt xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:sdtPr><w:{kind}>{children}</w:{kind}>{conflict}</w:sdtPr><w:sdtContent><w:p/></w:sdtContent></w:sdt>"#
                ));
                let properties = control.properties.unwrap();
                let before = properties.clone();
                assert!(
                    properties.prepare_text_value("2026-10-31").is_err(),
                    "{kind} + {conflict}"
                );
                assert_eq!(properties, before);
            }
            let foreign =
                r#"<a:checkbox xmlns:a="urn:foreign"/><a:repeatingSection xmlns:a="urn:foreign"/>"#;
            let mut control = parse_standalone_control(&format!(
                r#"<w:sdt xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:sdtPr><w:{kind}>{children}</w:{kind}>{foreign}</w:sdtPr><w:sdtContent><w:p/></w:sdtContent></w:sdt>"#
            ));
            assert!(
                control
                    .properties
                    .as_ref()
                    .unwrap()
                    .prepare_text_value("2026-10-31")
                    .is_ok()
            );
            control
                .properties
                .as_mut()
                .unwrap()
                .set_value_metadata(if kind == "date" {
                    Some("2026-10-31T00:00:00Z")
                } else {
                    None
                })
                .unwrap();
            let mut saved = Vec::new();
            control.to_xml(&mut Writer::new(&mut saved)).unwrap();
            assert!(String::from_utf8(saved).unwrap().contains(foreign));
        }
    }

    #[test]
    fn fx249_review_d3_selected_extension_metadata_preserves_kind() {
        for kind in ["repeatingSection", "repeatingSectionItem"] {
            let mut control = parse_standalone_control(&format!(
                r#"<w:sdt xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"><w:sdtPr><w:showingPlcHdr/><w15:{kind}/></w:sdtPr><w:sdtContent><w:p/></w:sdtContent></w:sdt>"#
            ));
            let expected = control.properties.as_ref().unwrap().control_type;
            control
                .properties
                .as_mut()
                .unwrap()
                .set_value_metadata(None)
                .unwrap();
            let mut saved = Vec::new();
            control.to_xml(&mut Writer::new(&mut saved)).unwrap();
            let saved = String::from_utf8(saved).unwrap();
            assert!(saved.contains(&format!("<w15:{kind}/>")), "{saved}");
            assert!(!saved.contains("showingPlcHdr"));
            assert_eq!(
                parse_standalone_control(&saved)
                    .properties
                    .unwrap()
                    .control_type,
                expected
            );
        }
    }

    #[test]
    fn fx249_qualified_value_metadata_preserves_opaque_slots() {
        let xml = r#"<w:sdt xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:p="urn:producer&amp;amp;"><w:sdtPr><p:before exact='yes'/><q:showingPlcHdr/><p:showingPlcHdr/><showingPlcHdr/><w:showingPlcHdr xmlns:w="urn:foreign"/><w:date p:fullDate="foreign" q:fullDate="old"><p:dateFormat p:val="unknown"/><q:dateFormat q:val="dd/MM/yyyy"/><q:calendar/><q:storeMappedDataAs/><p:opaque exact='yes'/></w:date><p:after/></w:sdtPr><w:sdtContent><w:p/></w:sdtContent></w:sdt>"#;
        let mut control = parse_standalone_control(xml);
        let properties = control.properties.as_mut().unwrap();
        assert_eq!(
            properties.prepare_text_value("2026-10-31").unwrap(),
            (
                "31/10/2026".into(),
                "31/10/2026".into(),
                Some("2026-10-31T00:00:00Z".into())
            )
        );
        let before = properties.clone();
        assert!(
            properties
                .set_value_metadata(Some("2026-02-29T00:00:00Z"))
                .is_err()
        );
        assert_eq!(*properties, before);
        properties
            .set_value_metadata(Some("2026-10-31T00:00:00Z"))
            .unwrap();
        let mut output = Vec::new();
        control.to_xml(&mut Writer::new(&mut output)).unwrap();
        let output = String::from_utf8(output).unwrap();
        assert!(!output.contains("<q:showingPlcHdr"));
        for raw in [
            "<p:before exact='yes'/>",
            "<p:showingPlcHdr/>",
            "<showingPlcHdr/>",
            "<w:showingPlcHdr xmlns:w=\"urn:foreign\"/>",
            "<p:opaque exact='yes'/>",
            "<p:after/>",
        ] {
            assert!(output.contains(raw), "missing {raw}");
        }
        assert!(output.contains("p:fullDate=\"foreign\" q:fullDate=\"2026-10-31T00:00:00Z\""));
        assert!(output.find("<p:before").unwrap() < output.find("<w:date ").unwrap());
        assert!(output.find("<w:date ").unwrap() < output.find("<p:after").unwrap());
        assert!(output.contains("urn:producer&amp;amp;"));
        let mut reopened = parse_standalone_control(&output);
        assert_eq!(
            reopened
                .properties
                .as_ref()
                .unwrap()
                .prepare_text_value("2026-10-31")
                .unwrap()
                .1,
            "31/10/2026"
        );
        reopened
            .properties
            .as_mut()
            .unwrap()
            .set_value_metadata(Some("2026-10-31T00:00:00Z"))
            .unwrap();
        let mut second = Vec::new();
        reopened.to_xml(&mut Writer::new(&mut second)).unwrap();
        assert_eq!(second, output.as_bytes());
    }

    #[test]
    fn fx249_selected_type_alias_survives_local_shadowing() {
        let mut control = parse_standalone_control(
            r#"<w:sdt xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:sdtPr><q:date xmlns:w="urn:foreign" q:fullDate="old"><w:dateFormat w:val="foreign"/><q:dateFormat q:val="yyyy-MM-dd"/></q:date></w:sdtPr><w:sdtContent><w:p/></w:sdtContent></w:sdt>"#,
        );
        let properties = control.properties.as_mut().unwrap();
        assert!(properties.prepare_text_value("2026-10-31").is_ok());
        properties
            .set_value_metadata(Some("2026-10-31T00:00:00Z"))
            .unwrap();
        let mut output = Vec::new();
        control.to_xml(&mut Writer::new(&mut output)).unwrap();
        let output = String::from_utf8(output).unwrap();
        assert!(
            output.contains("<q:date xmlns:w=\"urn:foreign\" q:fullDate=\"2026-10-31T00:00:00Z\"")
        );
        assert!(
            parse_standalone_control(&output)
                .properties
                .unwrap()
                .prepare_text_value("2026-10-31")
                .is_ok()
        );
    }

    #[test]
    fn fx249_default_word_type_scope_adds_only_selected_full_date_prefix() {
        let mut control = parse_standalone_control(
            r#"<w:sdt xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:sdtPr><date xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w="urn:foreign"><dateFormat xmlns:a="http://schemas.openxmlformats.org/wordprocessingml/2006/main" a:val="yyyy-MM-dd"/></date></w:sdtPr><w:sdtContent><w:p/></w:sdtContent></w:sdt>"#,
        );
        let properties = control.properties.as_mut().unwrap();
        assert!(properties.prepare_text_value("2026-10-31").is_ok());
        properties
            .set_value_metadata(Some("2026-10-31T00:00:00Z"))
            .unwrap();
        let mut output = Vec::new();
        control.to_xml(&mut Writer::new(&mut output)).unwrap();
        let output = String::from_utf8(output).unwrap();
        assert!(output.contains("rdocxW:fullDate=\"2026-10-31T00:00:00Z\""));
        assert!(
            parse_standalone_control(&output)
                .properties
                .unwrap()
                .prepare_text_value("2026-10-31")
                .is_ok()
        );
    }

    #[test]
    fn fx249_date_admission_and_storage_are_finite() {
        let make = |metadata: &str| {
            parse_standalone_control(&format!(
                r#"<w:sdt xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:sdtPr><w:date>{metadata}</w:date></w:sdtPr><w:sdtContent><w:p/></w:sdtContent></w:sdt>"#
            ))
        };
        for (mapping, expected) in [
            ("", "29/02/2000"),
            ("<w:storeMappedDataAs/>", "29/02/2000"),
            ("<w:storeMappedDataAs w:val=\"text\"/>", "29/02/2000"),
            ("<w:storeMappedDataAs w:val=\"date\"/>", "2000-02-29"),
            (
                "<w:storeMappedDataAs w:val=\"dateTime\"/>",
                "2000-02-29T00:00:00Z",
            ),
        ] {
            let control = make(&format!("<w:dateFormat w:val=\"dd/MM/yyyy\"/>{mapping}"));
            let properties = control.properties.unwrap();
            assert_eq!(
                properties.prepare_text_value("2000-02-29").unwrap().1,
                expected
            );
            for invalid in [
                "1900-02-29",
                "2026-02-29",
                "0000-01-01",
                "10000-01-01",
                "2026-13-01",
                "2026-04-31",
                "2026-01-00",
                "2026-1-01",
                " 2026-01-01",
                "2026-01-01T00:00:00Z",
                "２０２６-01-01",
            ] {
                assert!(
                    properties.prepare_text_value(invalid).is_err(),
                    "accepted {invalid}"
                );
            }
            for valid in ["0001-01-01", "9999-12-31", "2024-02-29"] {
                assert!(properties.prepare_text_value(valid).is_ok());
            }
        }
        for metadata in [
            "",
            "<w:dateFormat/>",
            "<w:dateFormat w:val=\"\"/>",
            "<w:dateFormat w:val=\"MM/dd/yyyy\"/>",
            "<w:dateFormat w:val=\"yyyy-MM-dd\"/><w:calendar w:val=\"hijri\"/>",
            "<w:dateFormat w:val=\"yyyy-MM-dd\"/><w:storeMappedDataAs w:val=\"unknown\"/>",
            "<w:dateFormat w:val=\"yyyy-MM-dd\"/><w:dateFormat w:val=\"yyyy-MM-dd\"/>",
            "<w:dateFormat w:val=\"yyyy-MM-dd\"/><w:calendar/><w:calendar/>",
        ] {
            assert!(
                make(metadata)
                    .properties
                    .unwrap()
                    .prepare_text_value("2026-10-31")
                    .is_err(),
                "accepted {metadata}"
            );
        }
        for calendar in ["", "<w:calendar/>", "<w:calendar w:val=\"gregorian\"/>"] {
            assert!(
                make(&format!("<w:dateFormat w:val=\"yyyy-MM-dd\"/>{calendar}"))
                    .properties
                    .unwrap()
                    .prepare_text_value("2026-10-31")
                    .is_ok()
            );
        }
    }

    #[test]
    fn fx249_dropdown_selection_preserves_absence_and_rejects_ambiguity() {
        let make = |items: &str, kind: &str| {
            parse_standalone_control(&format!(
                r#"<w:sdt xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:f="urn:foreign"><w:sdtPr><w:{kind} w:lastValue="producer">{items}</w:{kind}></w:sdtPr><w:sdtContent><w:p/></w:sdtContent></w:sdt>"#
            ))
        };
        let control = make(
            r#"<f:listItem w:value="foreign"/><w:listItem w:displayText="Two" w:value="two"/><w:listItem w:value="alpha"/><w:listItem w:displayText="" w:value="empty"/><w:listItem w:displayText="Missing"/>"#,
            "dropDownList",
        );
        let properties = control.properties.as_ref().unwrap();
        for input in ["Two", "two"] {
            assert_eq!(
                properties.prepare_text_value(input).unwrap(),
                ("Two".into(), "two".into(), None)
            );
        }
        assert_eq!(
            properties.prepare_text_value("alpha").unwrap(),
            ("alpha".into(), "alpha".into(), None)
        );
        for input in ["empty", ""] {
            assert_eq!(
                properties.prepare_text_value(input).unwrap(),
                ("".into(), "empty".into(), None)
            );
        }
        for input in ["Missing", "foreign", "TWO", " two", "two "] {
            assert!(properties.prepare_text_value(input).is_err());
        }
        for items in [
            r#"<w:listItem w:displayText="Same" w:value="a"/><w:listItem w:displayText="Same" w:value="b"/>"#,
            r#"<w:listItem w:displayText="Same" w:value="a"/><w:listItem w:displayText="a" w:value="b"/>"#,
        ] {
            assert!(
                make(items, "dropDownList")
                    .properties
                    .unwrap()
                    .prepare_text_value(if items.contains("displayText=\"a\"") {
                        "a"
                    } else {
                        "Same"
                    })
                    .is_err()
            );
        }
        assert!(
            make("", "comboBox")
                .properties
                .unwrap()
                .prepare_text_value("free")
                .is_err()
        );
        let combo = make(
            r#"<w:listItem w:displayText="Shown" w:value="stored"/>"#,
            "comboBox",
        );
        let combo = combo.properties.as_ref().unwrap();
        for input in ["Shown", "stored"] {
            assert_eq!(
                combo.prepare_text_value(input).unwrap(),
                ("Shown".into(), "stored".into(), None)
            );
        }
        assert!(combo.prepare_text_value("free").is_err());
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
        let typed = xml.find("<x:text").expect("retained typed marker");
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

    #[test]
    fn fx211_clone_preserves_all_standalone_content_carriers() {
        let raw = r#"<x:opaque x:key="one &amp; two"><x:child/><!--kept--></x:opaque>"#;
        let xml = format!(
            r#"<w:sdt xmlns:w="{W_NS}" xmlns:x="urn:producer" x:root="retained"><w:sdtPr><w:alias w:val="source"/><w:id w:val="42"/><w:tag w:val="kept"/></w:sdtPr><w:sdtContent x:content="retained"><w:p><w:r><w:t>paragraph</w:t></w:r></w:p><w:tbl><w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl><w:tr><w:trPr><w:cantSplit/></w:trPr><w:tc><w:p/></w:tc></w:tr><w:tc><w:tcPr><w:shd w:fill="AABBCC"/></w:tcPr><w:p/></w:tc><w:r><w:rPr><w:b/></w:rPr><w:t>run</w:t></w:r><w:sdt><w:sdtPr><w:tag w:val="nested"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>nested</w:t></w:r></w:p></w:sdtContent></w:sdt>{raw}</w:sdtContent></w:sdt>"#
        );
        let original =
            CT_Sdt::from_raw_with_context(xml.as_bytes(), &["w".to_owned()], SdtOwner::Standalone)
                .unwrap();
        assert_eq!(original.content.len(), 7);
        assert!(matches!(original.content[0], SdtContent::Paragraph(_)));
        assert!(matches!(original.content[1], SdtContent::Table(_)));
        assert!(matches!(original.content[2], SdtContent::Row(_)));
        assert!(matches!(original.content[3], SdtContent::Cell(_)));
        assert!(matches!(original.content[4], SdtContent::Run(_)));
        assert!(matches!(original.content[5], SdtContent::ContentControl(_)));
        assert!(matches!(original.content[6], SdtContent::RawXml(_)));
        let mut cloned = original.clone();
        assert_eq!(cloned, original);
        let mut before = Vec::new();
        let mut after = Vec::new();
        original.to_xml(&mut Writer::new(&mut before)).unwrap();
        cloned.to_xml(&mut Writer::new(&mut after)).unwrap();
        assert_eq!(after, before);
        assert!(String::from_utf8(after).unwrap().contains(raw));
        let SdtContent::RawXml(bytes) = &mut cloned.content[6] else {
            panic!("raw carrier")
        };
        bytes.extend_from_slice(b"<!--clone-only-->");
        assert_ne!(cloned, original);
        let SdtContent::RawXml(bytes) = &original.content[6] else {
            panic!("original raw carrier")
        };
        assert_eq!(bytes.as_slice(), raw.as_bytes());
        assert_eq!(original.properties.as_ref().unwrap().id, Some(42));
    }

    #[test]
    fn block_control_parser_types_only_children_owned_at_its_placement() {
        let body = CT_Sdt::from_body_raw(
            br#"<w:sdt><w:sdtContent><w:p/><w:tbl/><w:tr data-case="body-row"/><w:tc data-case="body-cell"/></w:sdtContent></w:sdt>"#,
            &["w".to_owned()],
        )
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
