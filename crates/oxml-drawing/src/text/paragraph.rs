use std::io::Write;

use oxml_core::OxmlError;
use oxml_core::raw_xml::{capture_element, capture_empty_element};
use oxml_core::xml::{local_name, matches_local_name};
use oxml_core::xml_text::{decode_plain, resolve_entity};
use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer, XmlVersion};

use crate::fill::Fill;
use crate::namespace::reject_conflicting_a_prefix;
use crate::order::OrderedRawChildren;

use super::body::{Result, TextError, has_forbidden_xml_char, missing_end};
use super::bullet::TextBullet;

const MAX_TEXT_MARGIN: i32 = 51_206_400;
const MAX_TEXT_POINT: i32 = 400_000;
const MAX_TEXT_SPACING_POINT: i32 = 158_400;
const MAX_TRANSITIONAL_SPACING_PERCENT: i32 = 13_200_000;

/// Whether source text explicitly requested XML whitespace preservation.
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub enum TextSpace {
    #[default]
    Default,
    Preserve,
}

/// One DrawingML `a:t` value and its source whitespace intent.
#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub struct TextValue {
    pub value: String,
    pub space: TextSpace,
    raw_attributes: Vec<(String, String)>,
}

impl TextValue {
    fn from_start(start: &BytesStart<'_>) -> Result<Self> {
        if !matches_local_name(start.name().as_ref(), b"t") {
            return Err(unexpected(start));
        }
        let space_value = text_attr(start, b"xml:space")?;
        let space = match space_value.as_deref() {
            None | Some("default") => TextSpace::Default,
            Some("preserve") => TextSpace::Preserve,
            Some(value) => return Err(invalid_attribute("t", "xml:space", value)),
        };
        Ok(Self {
            value: String::new(),
            space,
            raw_attributes: capture_raw_attributes(start, &[b"xml:space"])?,
        })
    }

    fn from_element(reader: &mut Reader<&[u8]>, start: &BytesStart<'_>) -> Result<Self> {
        let mut text = Self::from_start(start)?;
        let mut buffer = Vec::new();
        loop {
            match reader
                .read_event_into(&mut buffer)
                .map_err(OxmlError::from)?
            {
                Event::Text(value) => text.value.push_str(&decode_plain(&value)),
                Event::CData(value) => {
                    let value = std::str::from_utf8(value.as_ref()).map_err(OxmlError::from)?;
                    text.value.push_str(value);
                }
                Event::GeneralRef(value) => text.value.push_str(&resolve_entity(&value)),
                Event::End(element) if matches_local_name(element.name().as_ref(), b"t") => {
                    if text.space == TextSpace::Default
                        && (text.value.chars().next().is_some_and(char::is_whitespace)
                            || text
                                .value
                                .chars()
                                .next_back()
                                .is_some_and(char::is_whitespace))
                    {
                        text.space = TextSpace::Preserve;
                    }
                    return Ok(text);
                }
                Event::Start(element) | Event::Empty(element) => {
                    return Err(unexpected(&element));
                }
                Event::Eof => return Err(missing_end("t")),
                _ => {}
            }
            buffer.clear();
        }
    }

    fn write_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let preserve = self.space == TextSpace::Preserve
            || self.value.chars().next().is_some_and(char::is_whitespace)
            || self
                .value
                .chars()
                .next_back()
                .is_some_and(char::is_whitespace);
        let mut start = BytesStart::new("a:t");
        if preserve {
            start.push_attribute(("xml:space", "preserve"));
        }
        push_raw_attributes(&mut start, &self.raw_attributes);
        if self.value.is_empty() {
            return write_empty(writer, start);
        }
        write_start(writer, start)?;
        writer
            .write_event(Event::Text(BytesText::new(&self.value)))
            .map_err(OxmlError::from)?;
        write_end(writer, "a:t")
    }
}

/// The seven members of `ST_TextAlignType`.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum TextAlignment {
    Left,
    Center,
    Right,
    Justified,
    JustifiedLow,
    Distributed,
    ThaiDistributed,
}

impl TextAlignment {
    fn parse(value: &str) -> Option<Self> {
        match value {
            "l" => Some(Self::Left),
            "ctr" => Some(Self::Center),
            "r" => Some(Self::Right),
            "just" => Some(Self::Justified),
            "justLow" => Some(Self::JustifiedLow),
            "dist" => Some(Self::Distributed),
            "thaiDist" => Some(Self::ThaiDistributed),
            _ => None,
        }
    }

    const fn as_str(self) -> &'static str {
        match self {
            Self::Left => "l",
            Self::Center => "ctr",
            Self::Right => "r",
            Self::Justified => "just",
            Self::JustifiedLow => "justLow",
            Self::Distributed => "dist",
            Self::ThaiDistributed => "thaiDist",
        }
    }
}

/// DrawingML paragraph spacing in percentage text or centipoints.
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum TextSpacing {
    Percent(String),
    Points(i32),
}

impl TextSpacing {
    fn from_xml(xml: &[u8], wrapper: &[u8]) -> Result<Self> {
        parse_complete(
            xml,
            wrapper,
            |reader, start| Self::from_element(reader, start, wrapper),
            |_| {
                Err(TextError::UnexpectedElement(
                    String::from_utf8_lossy(wrapper).into_owned(),
                ))
            },
        )
    }

    fn from_element(
        reader: &mut Reader<&[u8]>,
        start: &BytesStart<'_>,
        wrapper: &[u8],
    ) -> Result<Self> {
        reject_conflicting_a_prefix(start)?;
        let mut spacing = None;
        let mut buffer = Vec::new();
        loop {
            match reader
                .read_event_into(&mut buffer)
                .map_err(OxmlError::from)?
            {
                Event::Empty(element) if matches_local_name(element.name().as_ref(), b"spcPct") => {
                    reject_conflicting_a_prefix(&element)?;
                    if spacing.is_some() {
                        return Err(duplicate("text spacing choice"));
                    }
                    let value = required_attr(&element, b"val")?;
                    validate_spacing_percent(&value)?;
                    spacing = Some(Self::Percent(value));
                }
                Event::Empty(element) if matches_local_name(element.name().as_ref(), b"spcPts") => {
                    reject_conflicting_a_prefix(&element)?;
                    if spacing.is_some() {
                        return Err(duplicate("text spacing choice"));
                    }
                    let value = parse_i32_attr(&element, b"val", 0, MAX_TEXT_SPACING_POINT)?;
                    spacing = Some(Self::Points(value));
                }
                Event::Start(element)
                    if matches_local_name(element.name().as_ref(), b"spcPct")
                        || matches_local_name(element.name().as_ref(), b"spcPts") =>
                {
                    reject_conflicting_a_prefix(&element)?;
                    if spacing.is_some() {
                        return Err(duplicate("text spacing choice"));
                    }
                    let name = local_name(element.name().as_ref()).to_vec();
                    let value = if name == b"spcPct" {
                        let value = required_attr(&element, b"val")?;
                        validate_spacing_percent(&value)?;
                        Self::Percent(value)
                    } else {
                        Self::Points(parse_i32_attr(&element, b"val", 0, MAX_TEXT_SPACING_POINT)?)
                    };
                    ensure_empty(reader, &name)?;
                    spacing = Some(value);
                }
                Event::Start(element) | Event::Empty(element) => return Err(unexpected(&element)),
                Event::End(element) if matches_local_name(element.name().as_ref(), wrapper) => {
                    return spacing.ok_or_else(|| {
                        TextError::UnexpectedElement(String::from_utf8_lossy(wrapper).into_owned())
                    });
                }
                Event::Eof => return Err(missing_end(&String::from_utf8_lossy(wrapper))),
                _ => {}
            }
            buffer.clear();
        }
    }

    fn write_xml<W: Write>(&self, writer: &mut Writer<W>, wrapper: &str) -> Result<()> {
        write_start(writer, BytesStart::new(wrapper))?;
        match self {
            Self::Percent(value) => {
                validate_spacing_percent(value)?;
                let mut start = BytesStart::new("a:spcPct");
                start.push_attribute(("val", value.as_str()));
                write_empty(writer, start)?;
            }
            Self::Points(value) if (0..=MAX_TEXT_SPACING_POINT).contains(value) => {
                let value = value.to_string();
                let mut start = BytesStart::new("a:spcPts");
                start.push_attribute(("val", value.as_str()));
                write_empty(writer, start)?;
            }
            Self::Points(value) => {
                return Err(invalid_attribute("spcPts", "val", &value.to_string()));
            }
        }
        write_end(writer, wrapper)
    }
}

/// Character spacing in centipoints or a strict universal measure.
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum TextPointValue {
    Centipoints(i32),
    UniversalMeasure(String),
}

impl TextPointValue {
    fn parse(value: String) -> Result<Self> {
        if let Ok(point) = value.parse::<i32>() {
            if (-MAX_TEXT_POINT..=MAX_TEXT_POINT).contains(&point) {
                return Ok(Self::Centipoints(point));
            }
            return Err(invalid_attribute("rPr", "spc", &value));
        }
        if is_universal_measure(&value) {
            Ok(Self::UniversalMeasure(value))
        } else {
            Err(invalid_attribute("rPr", "spc", &value))
        }
    }

    fn as_xml(&self) -> Result<String> {
        match self {
            Self::Centipoints(value) if (-MAX_TEXT_POINT..=MAX_TEXT_POINT).contains(value) => {
                Ok(value.to_string())
            }
            Self::Centipoints(value) => Err(invalid_attribute("rPr", "spc", &value.to_string())),
            Self::UniversalMeasure(value) if is_universal_measure(value) => Ok(value.clone()),
            Self::UniversalMeasure(value) => Err(invalid_attribute("rPr", "spc", value)),
        }
    }
}

/// DrawingML underline styles.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum TextUnderline {
    None,
    Words,
    Single,
    Double,
    Heavy,
    Dotted,
    DottedHeavy,
    Dash,
    DashHeavy,
    DashLong,
    DashLongHeavy,
    DotDash,
    DotDashHeavy,
    DotDotDash,
    DotDotDashHeavy,
    Wavy,
    WavyHeavy,
    WavyDouble,
}

impl TextUnderline {
    fn parse(value: &str) -> Option<Self> {
        Some(match value {
            "none" => Self::None,
            "words" => Self::Words,
            "sng" => Self::Single,
            "dbl" => Self::Double,
            "heavy" => Self::Heavy,
            "dotted" => Self::Dotted,
            "dottedHeavy" => Self::DottedHeavy,
            "dash" => Self::Dash,
            "dashHeavy" => Self::DashHeavy,
            "dashLong" => Self::DashLong,
            "dashLongHeavy" => Self::DashLongHeavy,
            "dotDash" => Self::DotDash,
            "dotDashHeavy" => Self::DotDashHeavy,
            "dotDotDash" => Self::DotDotDash,
            "dotDotDashHeavy" => Self::DotDotDashHeavy,
            "wavy" => Self::Wavy,
            "wavyHeavy" => Self::WavyHeavy,
            "wavyDbl" => Self::WavyDouble,
            _ => return None,
        })
    }

    const fn as_str(self) -> &'static str {
        match self {
            Self::None => "none",
            Self::Words => "words",
            Self::Single => "sng",
            Self::Double => "dbl",
            Self::Heavy => "heavy",
            Self::Dotted => "dotted",
            Self::DottedHeavy => "dottedHeavy",
            Self::Dash => "dash",
            Self::DashHeavy => "dashHeavy",
            Self::DashLong => "dashLong",
            Self::DashLongHeavy => "dashLongHeavy",
            Self::DotDash => "dotDash",
            Self::DotDashHeavy => "dotDashHeavy",
            Self::DotDotDash => "dotDotDash",
            Self::DotDotDashHeavy => "dotDotDashHeavy",
            Self::Wavy => "wavy",
            Self::WavyHeavy => "wavyHeavy",
            Self::WavyDouble => "wavyDbl",
        }
    }
}

/// DrawingML strike styles.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum TextStrike {
    None,
    Single,
    Double,
}

impl TextStrike {
    fn parse(value: &str) -> Option<Self> {
        match value {
            "noStrike" => Some(Self::None),
            "sngStrike" => Some(Self::Single),
            "dblStrike" => Some(Self::Double),
            _ => None,
        }
    }

    const fn as_str(self) -> &'static str {
        match self {
            Self::None => "noStrike",
            Self::Single => "sngStrike",
            Self::Double => "dblStrike",
        }
    }
}

/// A DrawingML text typeface reference.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct TextFont {
    pub typeface: String,
    raw_attributes: Vec<(String, String)>,
}

impl TextFont {
    pub fn new(typeface: impl Into<String>) -> Result<Self> {
        let typeface = typeface.into();
        if typeface.is_empty() || has_forbidden_xml_char(&typeface) {
            return Err(invalid_attribute("font", "typeface", &typeface));
        }
        Ok(Self {
            typeface,
            raw_attributes: Vec::new(),
        })
    }

    pub(crate) fn from_xml(xml: &[u8], expected: &[u8]) -> Result<Self> {
        parse_complete(
            xml,
            expected,
            |reader, start| {
                let font = Self::from_start(start)?;
                ensure_empty(reader, expected)?;
                Ok(font)
            },
            Self::from_start,
        )
    }

    fn from_start(start: &BytesStart<'_>) -> Result<Self> {
        reject_conflicting_a_prefix(start)?;
        Ok(Self {
            typeface: required_attr(start, b"typeface")?,
            raw_attributes: capture_raw_attributes(start, &[b"typeface"])?,
        })
    }

    pub(crate) fn write_xml<W: Write>(&self, writer: &mut Writer<W>, tag: &str) -> Result<()> {
        if self.typeface.is_empty() || has_forbidden_xml_char(&self.typeface) {
            return Err(invalid_attribute(tag, "typeface", &self.typeface));
        }
        let mut start = BytesStart::new(tag);
        start.push_attribute(("typeface", self.typeface.as_str()));
        push_raw_attributes(&mut start, &self.raw_attributes);
        write_empty(writer, start)
    }
}

/// Hyperlink attributes retained on text character properties.
#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub struct TextHyperlink {
    pub relationship_id: Option<String>,
    pub invalid_url: Option<String>,
    pub action: Option<String>,
    pub target_frame: Option<String>,
    pub tooltip: Option<String>,
    pub history: Option<bool>,
    pub highlight_click: Option<bool>,
    pub end_sound: Option<bool>,
    raw_attributes: Vec<(String, String)>,
    raw_children: OrderedRawChildren,
}

impl TextHyperlink {
    fn from_xml(xml: &[u8], expected: &[u8]) -> Result<Self> {
        parse_complete(
            xml,
            expected,
            |reader, start| Self::from_element(reader, start, expected),
            Self::from_start,
        )
    }

    fn from_start(start: &BytesStart<'_>) -> Result<Self> {
        reject_conflicting_a_prefix(start)?;
        let relationship_id = text_attr(start, b"r:id")?;
        Ok(Self {
            relationship_id,
            invalid_url: text_attr(start, b"invalidUrl")?,
            action: text_attr(start, b"action")?,
            target_frame: text_attr(start, b"tgtFrame")?,
            tooltip: text_attr(start, b"tooltip")?,
            history: parse_optional_bool(start, b"history")?,
            highlight_click: parse_optional_bool(start, b"highlightClick")?,
            end_sound: parse_optional_bool(start, b"endSnd")?,
            raw_attributes: capture_raw_attributes(
                start,
                &[
                    b"r:id",
                    b"invalidUrl",
                    b"action",
                    b"tgtFrame",
                    b"tooltip",
                    b"history",
                    b"highlightClick",
                    b"endSnd",
                ],
            )?,
            raw_children: OrderedRawChildren::default(),
        })
    }

    fn from_element(
        reader: &mut Reader<&[u8]>,
        start: &BytesStart<'_>,
        expected: &[u8],
    ) -> Result<Self> {
        let mut hyperlink = Self::from_start(start)?;
        let mut boundary = 0;
        let mut seen = [false; 2];
        let mut buffer = Vec::new();
        loop {
            match reader
                .read_event_into(&mut buffer)
                .map_err(OxmlError::from)?
            {
                Event::Start(element) => {
                    let name = local_name(element.name().as_ref()).to_vec();
                    let raw = capture_element(reader, &element)?;
                    capture_hyperlink_raw(&mut hyperlink, &name, raw, &mut boundary, &mut seen)?;
                }
                Event::Empty(element) => {
                    let name = local_name(element.name().as_ref()).to_vec();
                    let raw = capture_empty_element(&element)?;
                    capture_hyperlink_raw(&mut hyperlink, &name, raw, &mut boundary, &mut seen)?;
                }
                Event::End(element) if matches_local_name(element.name().as_ref(), expected) => {
                    return Ok(hyperlink);
                }
                Event::Eof => return Err(missing_end(&String::from_utf8_lossy(expected))),
                _ => {}
            }
            buffer.clear();
        }
    }

    fn write_xml<W: Write>(&self, writer: &mut Writer<W>, tag: &str) -> Result<()> {
        let mut start = BytesStart::new(tag);
        push_optional_attr(&mut start, "r:id", self.relationship_id.as_deref());
        push_optional_attr(&mut start, "invalidUrl", self.invalid_url.as_deref());
        push_optional_attr(&mut start, "action", self.action.as_deref());
        push_optional_attr(&mut start, "tgtFrame", self.target_frame.as_deref());
        push_optional_attr(&mut start, "tooltip", self.tooltip.as_deref());
        push_optional_bool(&mut start, "history", self.history);
        push_optional_bool(&mut start, "highlightClick", self.highlight_click);
        push_optional_bool(&mut start, "endSnd", self.end_sound);
        push_raw_attributes(&mut start, &self.raw_attributes);
        if self.raw_children.is_empty() {
            return write_empty(writer, start);
        }
        write_start(writer, start)?;
        for boundary in 0..=2 {
            emit_raw(writer, self.raw_children.at(boundary))?;
        }
        write_end(writer, tag)
    }
}

fn capture_hyperlink_raw(
    hyperlink: &mut TextHyperlink,
    name: &[u8],
    raw: Vec<u8>,
    boundary: &mut usize,
    seen: &mut [bool; 2],
) -> Result<()> {
    let slot = match name {
        b"snd" => Some(0),
        b"extLst" => Some(1),
        _ => None,
    };
    if let Some(slot) = slot {
        if seen[slot] {
            return Err(duplicate(&String::from_utf8_lossy(name)));
        }
        seen[slot] = true;
        hyperlink.raw_children.push((*boundary).max(slot), raw);
        *boundary = (*boundary).max(slot + 1);
    } else {
        hyperlink.raw_children.push(*boundary, raw);
    }
    Ok(())
}

/// DrawingML character properties shared by runs, breaks, fields, and defaults.
#[allow(non_camel_case_types)]
#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub struct CT_TextCharacterProperties {
    pub font_size: Option<i32>,
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub all_caps: Option<bool>,
    pub underline: Option<TextUnderline>,
    pub strike: Option<TextStrike>,
    pub spacing: Option<TextPointValue>,
    pub baseline: Option<String>,
    pub fill: Option<Fill>,
    pub latin: Option<TextFont>,
    pub east_asian: Option<TextFont>,
    pub complex_script: Option<TextFont>,
    pub symbol: Option<TextFont>,
    pub hyperlink_click: Option<TextHyperlink>,
    pub hyperlink_mouse_over: Option<TextHyperlink>,
    raw_attributes: Vec<(String, String)>,
    raw_children: OrderedRawChildren,
}

impl CT_TextCharacterProperties {
    /// Sets the typed capitalisation. A typed value replaces a preserved
    /// `cap`, such as `cap="small"`, for good, so clearing the typed value
    /// later leaves no capitalisation rather than the preserved one.
    pub fn set_all_caps(&mut self, all_caps: Option<bool>) {
        if all_caps.is_some() {
            self.raw_attributes.retain(|(name, _)| name != "cap");
        }
        self.all_caps = all_caps;
    }

    /// Parses one complete character-property element with any prefix.
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        let expected = root_local_name(xml)?;
        if !is_character_property_tag(&expected) {
            return Err(TextError::UnexpectedElement(
                String::from_utf8_lossy(&expected).into_owned(),
            ));
        }
        parse_complete(xml, &expected, Self::from_element, Self::from_start)
    }

    fn from_start(start: &BytesStart<'_>) -> Result<Self> {
        reject_conflicting_a_prefix(start)?;
        let element = String::from_utf8_lossy(local_name(start.name().as_ref())).into_owned();
        let font_size = text_attr(start, b"sz")?
            .map(|value| parse_i32_value(&element, "sz", &value, 100, MAX_TEXT_POINT))
            .transpose()?;
        let spacing = text_attr(start, b"spc")?
            .map(TextPointValue::parse)
            .transpose()?;
        let baseline = text_attr(start, b"baseline")?;
        if let Some(value) = baseline.as_deref() {
            validate_baseline(value)?;
        }
        let all_caps = match text_attr(start, b"cap")?.as_deref() {
            Some("all") => Some(true),
            Some("none") => Some(false),
            _ => None,
        };
        let mut raw_attributes = capture_raw_attributes(
            start,
            &[b"sz", b"b", b"i", b"u", b"strike", b"spc", b"baseline"],
        )?;
        if all_caps.is_some() {
            raw_attributes.retain(|(name, _)| name != "cap");
        }
        Ok(Self {
            font_size,
            bold: parse_optional_bool(start, b"b")?,
            italic: parse_optional_bool(start, b"i")?,
            all_caps,
            underline: parse_optional_enum(start, b"u", TextUnderline::parse)?,
            strike: parse_optional_enum(start, b"strike", TextStrike::parse)?,
            spacing,
            baseline,
            raw_attributes,
            ..Self::default()
        })
    }

    fn from_element(reader: &mut Reader<&[u8]>, start: &BytesStart<'_>) -> Result<Self> {
        let expected = local_name(start.name().as_ref()).to_vec();
        let mut properties = Self::from_start(start)?;
        let mut boundary = 0;
        let mut seen = [false; 14];
        let mut buffer = Vec::new();
        loop {
            match reader
                .read_event_into(&mut buffer)
                .map_err(OxmlError::from)?
            {
                Event::Start(element) => {
                    let name = local_name(element.name().as_ref()).to_vec();
                    let raw = capture_element(reader, &element)?;
                    properties.capture_child(&name, raw, &mut boundary, &mut seen)?;
                }
                Event::Empty(element) => {
                    let name = local_name(element.name().as_ref()).to_vec();
                    let raw = capture_empty_element(&element)?;
                    properties.capture_child(&name, raw, &mut boundary, &mut seen)?;
                }
                Event::End(element) if matches_local_name(element.name().as_ref(), &expected) => {
                    return Ok(properties);
                }
                Event::Eof => return Err(missing_end(&String::from_utf8_lossy(&expected))),
                _ => {}
            }
            buffer.clear();
        }
    }

    fn capture_child(
        &mut self,
        name: &[u8],
        raw: Vec<u8>,
        boundary: &mut usize,
        seen: &mut [bool; 14],
    ) -> Result<()> {
        let slot = character_property_slot(name);
        if let Some(slot) = slot {
            if seen[slot - 1] {
                return Err(duplicate(&String::from_utf8_lossy(name)));
            }
            seen[slot - 1] = true;
        }
        match name {
            name if is_fill(name) => self.fill = Some(Fill::from_xml(&raw)?),
            b"latin" => self.latin = Some(TextFont::from_xml(&raw, b"latin")?),
            b"ea" => self.east_asian = Some(TextFont::from_xml(&raw, b"ea")?),
            b"cs" => self.complex_script = Some(TextFont::from_xml(&raw, b"cs")?),
            b"sym" => self.symbol = Some(TextFont::from_xml(&raw, b"sym")?),
            b"hlinkClick" => {
                self.hyperlink_click = Some(TextHyperlink::from_xml(&raw, b"hlinkClick")?)
            }
            b"hlinkMouseOver" => {
                self.hyperlink_mouse_over = Some(TextHyperlink::from_xml(&raw, b"hlinkMouseOver")?)
            }
            _ => self.raw_children.push(
                slot.map_or(*boundary, |slot| (*boundary).max(slot - 1)),
                raw,
            ),
        }
        if let Some(slot) = slot {
            *boundary = (*boundary).max(slot);
        }
        Ok(())
    }

    /// Writes character properties with a caller-selected fixed `a:` tag.
    pub fn write_xml<W: Write>(&self, writer: &mut Writer<W>, tag: &str) -> Result<()> {
        let mut start = BytesStart::new(tag);
        if let Some(value) = self.font_size {
            parse_i32_value(tag, "sz", &value.to_string(), 100, MAX_TEXT_POINT)?;
            push_owned_attr(&mut start, "sz", value.to_string());
        }
        push_optional_bool(&mut start, "b", self.bold);
        push_optional_bool(&mut start, "i", self.italic);
        if let Some(value) = self.all_caps {
            start.push_attribute(("cap", if value { "all" } else { "none" }));
        }
        if let Some(value) = self.underline {
            start.push_attribute(("u", value.as_str()));
        }
        if let Some(value) = self.strike {
            start.push_attribute(("strike", value.as_str()));
        }
        if let Some(value) = &self.spacing {
            push_owned_attr(&mut start, "spc", value.as_xml()?);
        }
        if let Some(value) = self.baseline.as_deref() {
            validate_baseline(value)?;
            start.push_attribute(("baseline", value));
        }
        for (name, value) in &self.raw_attributes {
            // A typed capitalisation replaces a preserved `cap="small"`.
            if name == "cap" && self.all_caps.is_some() {
                continue;
            }
            start.push_attribute((name.as_str(), value.as_str()));
        }

        let has_modelled_children = self.fill.is_some()
            || self.latin.is_some()
            || self.east_asian.is_some()
            || self.complex_script.is_some()
            || self.symbol.is_some()
            || self.hyperlink_click.is_some()
            || self.hyperlink_mouse_over.is_some();
        if !has_modelled_children && self.raw_children.is_empty() {
            return write_empty(writer, start);
        }
        write_start(writer, start)?;
        emit_raw(writer, self.raw_children.at(0))?;
        emit_raw(writer, self.raw_children.at(1))?;
        if let Some(fill) = &self.fill {
            fill.write_xml(writer)?;
        }
        for boundary in 2..=6 {
            emit_raw(writer, self.raw_children.at(boundary))?;
        }
        if let Some(font) = &self.latin {
            font.write_xml(writer, "a:latin")?;
        }
        emit_raw(writer, self.raw_children.at(7))?;
        if let Some(font) = &self.east_asian {
            font.write_xml(writer, "a:ea")?;
        }
        emit_raw(writer, self.raw_children.at(8))?;
        if let Some(font) = &self.complex_script {
            font.write_xml(writer, "a:cs")?;
        }
        emit_raw(writer, self.raw_children.at(9))?;
        if let Some(font) = &self.symbol {
            font.write_xml(writer, "a:sym")?;
        }
        emit_raw(writer, self.raw_children.at(10))?;
        if let Some(hyperlink) = &self.hyperlink_click {
            hyperlink.write_xml(writer, "a:hlinkClick")?;
        }
        emit_raw(writer, self.raw_children.at(11))?;
        if let Some(hyperlink) = &self.hyperlink_mouse_over {
            hyperlink.write_xml(writer, "a:hlinkMouseOver")?;
        }
        for boundary in 12..=14 {
            emit_raw(writer, self.raw_children.at(boundary))?;
        }
        write_end(writer, tag)
    }

    pub fn raw_children(&self) -> &OrderedRawChildren {
        &self.raw_children
    }
}

fn character_property_slot(name: &[u8]) -> Option<usize> {
    match name {
        b"ln" => Some(1),
        name if is_fill(name) => Some(2),
        b"effectLst" | b"effectDag" => Some(3),
        b"highlight" => Some(4),
        b"uLnTx" | b"uLn" => Some(5),
        b"uFillTx" | b"uFill" => Some(6),
        b"latin" => Some(7),
        b"ea" => Some(8),
        b"cs" => Some(9),
        b"sym" => Some(10),
        b"hlinkClick" => Some(11),
        b"hlinkMouseOver" => Some(12),
        b"rtl" => Some(13),
        b"extLst" => Some(14),
        _ => None,
    }
}

fn is_character_property_tag(name: &[u8]) -> bool {
    matches!(name, b"rPr" | b"defRPr" | b"endParaRPr")
}

fn is_fill(name: &[u8]) -> bool {
    matches!(
        name,
        b"noFill" | b"solidFill" | b"gradFill" | b"pattFill" | b"blipFill"
    )
}

/// Paragraph attributes and the story-owned spacing and default-run children.
#[allow(non_camel_case_types)]
#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub struct CT_TextParagraphProperties {
    pub left_margin: Option<i32>,
    pub right_margin: Option<i32>,
    pub level: Option<u8>,
    pub indent: Option<i32>,
    pub alignment: Option<TextAlignment>,
    pub right_to_left: Option<bool>,
    pub line_spacing: Option<TextSpacing>,
    pub space_before: Option<TextSpacing>,
    pub space_after: Option<TextSpacing>,
    pub bullet: Option<TextBullet>,
    pub default_run_properties: Option<CT_TextCharacterProperties>,
    raw_attributes: Vec<(String, String)>,
    raw_children: OrderedRawChildren,
}

impl CT_TextParagraphProperties {
    /// Parses one complete paragraph-property element with any prefix.
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        let expected = root_local_name(xml)?;
        if !is_paragraph_property_tag(&expected) {
            return Err(TextError::UnexpectedElement(
                String::from_utf8_lossy(&expected).into_owned(),
            ));
        }
        parse_complete(xml, &expected, Self::from_element, Self::from_start)
    }

    fn from_start(start: &BytesStart<'_>) -> Result<Self> {
        reject_conflicting_a_prefix(start)?;
        let element = String::from_utf8_lossy(local_name(start.name().as_ref())).into_owned();
        let level = text_attr(start, b"lvl")?
            .map(|value| parse_i32_value(&element, "lvl", &value, 0, 8).map(|v| v as u8))
            .transpose()?;
        Ok(Self {
            left_margin: parse_optional_i32(start, b"marL", 0, MAX_TEXT_MARGIN)?,
            right_margin: parse_optional_i32(start, b"marR", 0, MAX_TEXT_MARGIN)?,
            level,
            indent: parse_optional_i32(start, b"indent", -MAX_TEXT_MARGIN, MAX_TEXT_MARGIN)?,
            alignment: parse_optional_enum(start, b"algn", TextAlignment::parse)?,
            right_to_left: parse_optional_bool(start, b"rtl")?,
            raw_attributes: capture_raw_attributes(
                start,
                &[b"marL", b"marR", b"lvl", b"indent", b"algn"],
            )?,
            ..Self::default()
        })
    }

    fn from_element(reader: &mut Reader<&[u8]>, start: &BytesStart<'_>) -> Result<Self> {
        let expected = local_name(start.name().as_ref()).to_vec();
        let mut properties = Self::from_start(start)?;
        let mut boundary = 0;
        let mut seen = [false; 10];
        let mut buffer = Vec::new();
        loop {
            match reader
                .read_event_into(&mut buffer)
                .map_err(OxmlError::from)?
            {
                Event::Start(element) => {
                    let name = local_name(element.name().as_ref()).to_vec();
                    let raw = capture_element(reader, &element)?;
                    properties.capture_child(&name, raw, &mut boundary, &mut seen)?;
                }
                Event::Empty(element) => {
                    let name = local_name(element.name().as_ref()).to_vec();
                    let raw = capture_empty_element(&element)?;
                    properties.capture_child(&name, raw, &mut boundary, &mut seen)?;
                }
                Event::End(element) if matches_local_name(element.name().as_ref(), &expected) => {
                    return Ok(properties);
                }
                Event::Eof => return Err(missing_end(&String::from_utf8_lossy(&expected))),
                _ => {}
            }
            buffer.clear();
        }
    }

    fn capture_child(
        &mut self,
        name: &[u8],
        raw: Vec<u8>,
        boundary: &mut usize,
        seen: &mut [bool; 10],
    ) -> Result<()> {
        let slot = paragraph_property_slot(name);
        if let Some(slot) = slot {
            if seen[slot - 1] {
                return Err(duplicate(&String::from_utf8_lossy(name)));
            }
            seen[slot - 1] = true;
        }
        match name {
            b"lnSpc" => self.line_spacing = Some(TextSpacing::from_xml(&raw, b"lnSpc")?),
            b"spcBef" => self.space_before = Some(TextSpacing::from_xml(&raw, b"spcBef")?),
            b"spcAft" => self.space_after = Some(TextSpacing::from_xml(&raw, b"spcAft")?),
            b"defRPr" => {
                self.default_run_properties = Some(CT_TextCharacterProperties::from_xml(&raw)?)
            }
            _ => {
                let mut bullet = self.bullet.take().unwrap_or_default();
                if bullet.capture_component(name, &raw)? {
                    self.bullet = Some(bullet);
                } else {
                    if !bullet.is_empty() {
                        self.bullet = Some(bullet);
                    }
                    self.raw_children.push(
                        slot.map_or(*boundary, |slot| (*boundary).max(slot - 1)),
                        raw,
                    );
                }
            }
        }
        if let Some(slot) = slot {
            *boundary = (*boundary).max(slot);
        }
        Ok(())
    }

    /// Writes paragraph properties with a caller-selected fixed `a:` tag.
    pub fn write_xml<W: Write>(&self, writer: &mut Writer<W>, tag: &str) -> Result<()> {
        let mut start = BytesStart::new(tag);
        if let Some(value) = self.left_margin {
            parse_i32_value(tag, "marL", &value.to_string(), 0, MAX_TEXT_MARGIN)?;
            push_owned_attr(&mut start, "marL", value.to_string());
        }
        if let Some(value) = self.right_margin {
            parse_i32_value(tag, "marR", &value.to_string(), 0, MAX_TEXT_MARGIN)?;
            push_owned_attr(&mut start, "marR", value.to_string());
        }
        if let Some(value) = self.level {
            if value > 8 {
                return Err(invalid_attribute(tag, "lvl", &value.to_string()));
            }
            push_owned_attr(&mut start, "lvl", value.to_string());
        }
        if let Some(value) = self.indent {
            parse_i32_value(
                tag,
                "indent",
                &value.to_string(),
                -MAX_TEXT_MARGIN,
                MAX_TEXT_MARGIN,
            )?;
            push_owned_attr(&mut start, "indent", value.to_string());
        }
        if let Some(value) = self.alignment {
            start.push_attribute(("algn", value.as_str()));
        }
        let mut wrote_right_to_left = false;
        for (name, value) in &self.raw_attributes {
            if name == "rtl" {
                if let Some(value) = self.right_to_left {
                    start.push_attribute(("rtl", if value { "1" } else { "0" }));
                    wrote_right_to_left = true;
                }
            } else {
                start.push_attribute((name.as_str(), value.as_str()));
            }
        }
        if !wrote_right_to_left {
            push_optional_bool(&mut start, "rtl", self.right_to_left);
        }

        let modelled = self.line_spacing.is_some()
            || self.space_before.is_some()
            || self.space_after.is_some()
            || self.bullet.is_some()
            || self.default_run_properties.is_some();
        if !modelled && self.raw_children.is_empty() {
            return write_empty(writer, start);
        }
        write_start(writer, start)?;
        // A typed bullet member replaces the preserved alternative of its
        // group whatever path set it, so a group never holds two members.
        let raw_at = |boundary| {
            self.raw_children
                .at(boundary)
                .filter(|raw| !self.bullet_replaces(raw_local_name(raw)))
        };
        emit_raw(writer, raw_at(0))?;
        if let Some(spacing) = &self.line_spacing {
            spacing.write_xml(writer, "a:lnSpc")?;
        }
        emit_raw(writer, raw_at(1))?;
        if let Some(spacing) = &self.space_before {
            spacing.write_xml(writer, "a:spcBef")?;
        }
        emit_raw(writer, raw_at(2))?;
        if let Some(spacing) = &self.space_after {
            spacing.write_xml(writer, "a:spcAft")?;
        }
        emit_raw(writer, raw_at(3))?;
        if let Some(bullet) = &self.bullet {
            bullet.write_color(writer)?;
        }
        emit_raw(writer, raw_at(4))?;
        if let Some(bullet) = &self.bullet {
            bullet.write_size(writer)?;
        }
        emit_raw(writer, raw_at(5))?;
        if let Some(bullet) = &self.bullet {
            bullet.write_font(writer)?;
        }
        emit_raw(writer, raw_at(6))?;
        if let Some(bullet) = &self.bullet {
            bullet.write_choice(writer)?;
        }
        emit_raw(writer, raw_at(7))?;
        emit_raw(writer, raw_at(8))?;
        if let Some(properties) = &self.default_run_properties {
            properties.write_xml(writer, "a:defRPr")?;
        }
        emit_raw(writer, raw_at(9))?;
        emit_raw(writer, raw_at(10))?;
        write_end(writer, tag)
    }

    /// Whether the typed bullet has the member of the group a preserved
    /// child named `name` belongs to.
    fn bullet_replaces(&self, name: &[u8]) -> bool {
        self.bullet.as_ref().is_some_and(|bullet| match name {
            b"buClrTx" => bullet.color.is_some(),
            b"buSzTx" => bullet.size.is_some(),
            b"buFontTx" => bullet.font.is_some(),
            b"buBlip" => bullet.choice.is_some(),
            _ => false,
        })
    }

    pub fn raw_children(&self) -> &OrderedRawChildren {
        &self.raw_children
    }

    /// Returns whether the paragraph keeps a preserved `a:buBlip` picture bullet.
    pub fn has_picture_bullet(&self) -> bool {
        !self.bullet_replaces(b"buBlip")
            && (0..=10).any(|boundary| {
                self.raw_children
                    .at(boundary)
                    .any(|raw| raw_local_name(raw) == b"buBlip")
            })
    }

    /// Replaces the direct bullet.
    ///
    /// Each bullet group holds one member, so a colour, size, font, or choice
    /// set here removes its preserved `a:buClrTx`, `a:buSzTx`, `a:buFontTx`,
    /// or `a:buBlip` alternative. Clearing the bullet removes all four.
    pub fn set_bullet(&mut self, bullet: Option<TextBullet>) {
        let replaces = |name: &[u8]| {
            let Some(bullet) = bullet.as_ref() else {
                return matches!(name, b"buClrTx" | b"buSzTx" | b"buFontTx" | b"buBlip");
            };
            match name {
                b"buClrTx" => bullet.color.is_some(),
                b"buSzTx" => bullet.size.is_some(),
                b"buFontTx" => bullet.font.is_some(),
                b"buBlip" => bullet.choice.is_some(),
                _ => false,
            }
        };
        self.raw_children
            .retain(|raw| !replaces(raw_local_name(raw)));
        self.bullet = bullet;
    }

    /// Serialises a complete paragraph-property fragment.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        let mut writer = Writer::new(Vec::new());
        self.write_xml(&mut writer, "a:pPr")?;
        Ok(writer.into_inner())
    }
}

fn paragraph_property_slot(name: &[u8]) -> Option<usize> {
    match name {
        b"lnSpc" => Some(1),
        b"spcBef" => Some(2),
        b"spcAft" => Some(3),
        b"buClrTx" | b"buClr" => Some(4),
        b"buSzTx" | b"buSzPct" | b"buSzPts" => Some(5),
        b"buFontTx" | b"buFont" => Some(6),
        b"buNone" | b"buAutoNum" | b"buChar" | b"buBlip" => Some(7),
        b"tabLst" => Some(8),
        b"defRPr" => Some(9),
        b"extLst" => Some(10),
        _ => None,
    }
}

/// Returns the local name of a captured raw element.
fn raw_local_name(raw: &[u8]) -> &[u8] {
    let tag = raw.strip_prefix(b"<").unwrap_or(raw);
    let end = tag
        .iter()
        .position(|byte| byte.is_ascii_whitespace() || matches!(byte, b'/' | b'>'))
        .unwrap_or(tag.len());
    local_name(&tag[..end])
}

fn is_paragraph_property_tag(name: &[u8]) -> bool {
    matches!(
        name,
        b"pPr"
            | b"defPPr"
            | b"lvl1pPr"
            | b"lvl2pPr"
            | b"lvl3pPr"
            | b"lvl4pPr"
            | b"lvl5pPr"
            | b"lvl6pPr"
            | b"lvl7pPr"
            | b"lvl8pPr"
            | b"lvl9pPr"
    )
}

/// A regular DrawingML run containing optional properties and required text.
#[allow(non_camel_case_types)]
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct CT_RegularTextRun {
    pub properties: Option<CT_TextCharacterProperties>,
    pub text: TextValue,
    raw_children: OrderedRawChildren,
}

impl CT_RegularTextRun {
    /// Creates a regular text run without direct character formatting.
    pub fn new(text: impl Into<String>) -> Self {
        Self {
            properties: None,
            text: TextValue {
                value: text.into(),
                ..TextValue::default()
            },
            raw_children: OrderedRawChildren::default(),
        }
    }

    /// Replaces the run text while retaining properties and unmodelled XML.
    pub fn set_text(&mut self, text: &str) {
        self.text.value = text.to_owned();
    }

    /// Parses one complete `a:r` element with any prefix.
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        parse_complete(xml, b"r", Self::from_element, |_| {
            Err(TextError::UnexpectedElement("r".to_owned()))
        })
    }

    fn from_element(reader: &mut Reader<&[u8]>, _start: &BytesStart<'_>) -> Result<Self> {
        let mut properties = None;
        let mut text = None;
        let mut raw_children = OrderedRawChildren::default();
        let mut boundary = 0;
        let mut buffer = Vec::new();
        loop {
            match reader
                .read_event_into(&mut buffer)
                .map_err(OxmlError::from)?
            {
                Event::Start(element) if matches_local_name(element.name().as_ref(), b"rPr") => {
                    if properties.is_some() {
                        return Err(duplicate("rPr"));
                    }
                    properties = Some(CT_TextCharacterProperties::from_element(reader, &element)?);
                    boundary = boundary.max(1);
                }
                Event::Empty(element) if matches_local_name(element.name().as_ref(), b"rPr") => {
                    if properties.is_some() {
                        return Err(duplicate("rPr"));
                    }
                    properties = Some(CT_TextCharacterProperties::from_start(&element)?);
                    boundary = boundary.max(1);
                }
                Event::Start(element) if matches_local_name(element.name().as_ref(), b"t") => {
                    if text.is_some() {
                        return Err(duplicate("t"));
                    }
                    text = Some(TextValue::from_element(reader, &element)?);
                    boundary = boundary.max(2);
                }
                Event::Empty(element) if matches_local_name(element.name().as_ref(), b"t") => {
                    if text.is_some() {
                        return Err(duplicate("t"));
                    }
                    text = Some(TextValue::from_start(&element)?);
                    boundary = boundary.max(2);
                }
                Event::Start(element) => {
                    raw_children.push(boundary, capture_element(reader, &element)?)
                }
                Event::Empty(element) => {
                    raw_children.push(boundary, capture_empty_element(&element)?)
                }
                Event::End(element) if matches_local_name(element.name().as_ref(), b"r") => break,
                Event::Eof => return Err(missing_end("r")),
                _ => {}
            }
            buffer.clear();
        }
        Ok(Self {
            properties,
            text: text.ok_or_else(|| {
                TextError::Xml(OxmlError::MissingElement("DrawingML run text".to_owned()))
            })?,
            raw_children,
        })
    }

    fn write_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        write_start(writer, BytesStart::new("a:r"))?;
        emit_raw(writer, self.raw_children.at(0))?;
        if let Some(properties) = &self.properties {
            properties.write_xml(writer, "a:rPr")?;
        }
        emit_raw(writer, self.raw_children.at(1))?;
        self.text.write_xml(writer)?;
        emit_raw(writer, self.raw_children.at(2))?;
        write_end(writer, "a:r")
    }
}

/// A DrawingML explicit line break and its optional run properties.
#[allow(non_camel_case_types)]
#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub struct CT_TextLineBreak {
    pub properties: Option<CT_TextCharacterProperties>,
    raw_children: OrderedRawChildren,
}

impl CT_TextLineBreak {
    /// Parses one complete `a:br` element with any prefix.
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        parse_complete(xml, b"br", Self::from_element, |_| Ok(Self::default()))
    }

    fn from_element(reader: &mut Reader<&[u8]>, _start: &BytesStart<'_>) -> Result<Self> {
        let mut line_break = Self::default();
        let mut boundary = 0;
        let mut buffer = Vec::new();
        loop {
            match reader
                .read_event_into(&mut buffer)
                .map_err(OxmlError::from)?
            {
                Event::Start(element) if matches_local_name(element.name().as_ref(), b"rPr") => {
                    if line_break.properties.is_some() {
                        return Err(duplicate("rPr"));
                    }
                    line_break.properties =
                        Some(CT_TextCharacterProperties::from_element(reader, &element)?);
                    boundary = 1;
                }
                Event::Empty(element) if matches_local_name(element.name().as_ref(), b"rPr") => {
                    if line_break.properties.is_some() {
                        return Err(duplicate("rPr"));
                    }
                    line_break.properties = Some(CT_TextCharacterProperties::from_start(&element)?);
                    boundary = 1;
                }
                Event::Start(element) => line_break
                    .raw_children
                    .push(boundary, capture_element(reader, &element)?),
                Event::Empty(element) => line_break
                    .raw_children
                    .push(boundary, capture_empty_element(&element)?),
                Event::End(element) if matches_local_name(element.name().as_ref(), b"br") => {
                    return Ok(line_break);
                }
                Event::Eof => return Err(missing_end("br")),
                _ => {}
            }
            buffer.clear();
        }
    }

    fn write_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        if self.properties.is_none() && self.raw_children.is_empty() {
            return write_empty(writer, BytesStart::new("a:br"));
        }
        write_start(writer, BytesStart::new("a:br"))?;
        emit_raw(writer, self.raw_children.at(0))?;
        if let Some(properties) = &self.properties {
            properties.write_xml(writer, "a:rPr")?;
        }
        emit_raw(writer, self.raw_children.at(1))?;
        write_end(writer, "a:br")
    }
}

/// A DrawingML field with required GUID, optional type, properties, and text.
#[allow(non_camel_case_types)]
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct CT_TextField {
    pub id: String,
    pub field_type: Option<String>,
    pub run_properties: Option<CT_TextCharacterProperties>,
    pub paragraph_properties: Option<CT_TextParagraphProperties>,
    pub text: Option<TextValue>,
    raw_children: OrderedRawChildren,
}

impl CT_TextField {
    /// Parses one complete `a:fld` element with any prefix.
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        parse_complete(xml, b"fld", Self::from_element, Self::from_empty)
    }

    fn from_start(start: &BytesStart<'_>) -> Result<Self> {
        let id = required_attr(start, b"id")?;
        if !is_guid(&id) {
            return Err(invalid_attribute("fld", "id", &id));
        }
        Ok(Self {
            id,
            field_type: text_attr(start, b"type")?,
            run_properties: None,
            paragraph_properties: None,
            text: None,
            raw_children: OrderedRawChildren::default(),
        })
    }

    fn from_empty(start: &BytesStart<'_>) -> Result<Self> {
        Self::from_start(start)
    }

    fn from_element(reader: &mut Reader<&[u8]>, start: &BytesStart<'_>) -> Result<Self> {
        let mut field = Self::from_start(start)?;
        let mut boundary = 0;
        let mut buffer = Vec::new();
        loop {
            match reader
                .read_event_into(&mut buffer)
                .map_err(OxmlError::from)?
            {
                Event::Start(element) if matches_local_name(element.name().as_ref(), b"rPr") => {
                    if field.run_properties.is_some() {
                        return Err(duplicate("rPr"));
                    }
                    field.run_properties =
                        Some(CT_TextCharacterProperties::from_element(reader, &element)?);
                    boundary = boundary.max(1);
                }
                Event::Empty(element) if matches_local_name(element.name().as_ref(), b"rPr") => {
                    if field.run_properties.is_some() {
                        return Err(duplicate("rPr"));
                    }
                    field.run_properties = Some(CT_TextCharacterProperties::from_start(&element)?);
                    boundary = boundary.max(1);
                }
                Event::Start(element) if matches_local_name(element.name().as_ref(), b"pPr") => {
                    if field.paragraph_properties.is_some() {
                        return Err(duplicate("pPr"));
                    }
                    field.paragraph_properties =
                        Some(CT_TextParagraphProperties::from_element(reader, &element)?);
                    boundary = boundary.max(2);
                }
                Event::Empty(element) if matches_local_name(element.name().as_ref(), b"pPr") => {
                    if field.paragraph_properties.is_some() {
                        return Err(duplicate("pPr"));
                    }
                    field.paragraph_properties =
                        Some(CT_TextParagraphProperties::from_start(&element)?);
                    boundary = boundary.max(2);
                }
                Event::Start(element) if matches_local_name(element.name().as_ref(), b"t") => {
                    if field.text.is_some() {
                        return Err(duplicate("t"));
                    }
                    field.text = Some(TextValue::from_element(reader, &element)?);
                    boundary = boundary.max(3);
                }
                Event::Empty(element) if matches_local_name(element.name().as_ref(), b"t") => {
                    if field.text.is_some() {
                        return Err(duplicate("t"));
                    }
                    field.text = Some(TextValue::from_start(&element)?);
                    boundary = boundary.max(3);
                }
                Event::Start(element) => field
                    .raw_children
                    .push(boundary, capture_element(reader, &element)?),
                Event::Empty(element) => field
                    .raw_children
                    .push(boundary, capture_empty_element(&element)?),
                Event::End(element) if matches_local_name(element.name().as_ref(), b"fld") => {
                    return Ok(field);
                }
                Event::Eof => return Err(missing_end("fld")),
                _ => {}
            }
            buffer.clear();
        }
    }

    fn write_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        if !is_guid(&self.id) {
            return Err(invalid_attribute("fld", "id", &self.id));
        }
        let mut start = BytesStart::new("a:fld");
        start.push_attribute(("id", self.id.as_str()));
        push_optional_attr(&mut start, "type", self.field_type.as_deref());
        let modelled = self.run_properties.is_some()
            || self.paragraph_properties.is_some()
            || self.text.is_some();
        if !modelled && self.raw_children.is_empty() {
            return write_empty(writer, start);
        }
        write_start(writer, start)?;
        emit_raw(writer, self.raw_children.at(0))?;
        if let Some(properties) = &self.run_properties {
            properties.write_xml(writer, "a:rPr")?;
        }
        emit_raw(writer, self.raw_children.at(1))?;
        if let Some(properties) = &self.paragraph_properties {
            properties.write_xml(writer, "a:pPr")?;
        }
        emit_raw(writer, self.raw_children.at(2))?;
        if let Some(text) = &self.text {
            text.write_xml(writer)?;
        }
        emit_raw(writer, self.raw_children.at(3))?;
        write_end(writer, "a:fld")
    }
}

/// One member of the paragraph text-content choice.
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum TextRun {
    Run(CT_RegularTextRun),
    Break(CT_TextLineBreak),
    Field(Box<CT_TextField>),
}

impl TextRun {
    fn from_xml(xml: &[u8], name: &[u8]) -> Result<Self> {
        match name {
            b"r" => Ok(Self::Run(CT_RegularTextRun::from_xml(xml)?)),
            b"br" => Ok(Self::Break(CT_TextLineBreak::from_xml(xml)?)),
            b"fld" => Ok(Self::Field(Box::new(CT_TextField::from_xml(xml)?))),
            _ => Err(TextError::UnexpectedElement(
                String::from_utf8_lossy(name).into_owned(),
            )),
        }
    }

    fn write_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        match self {
            Self::Run(run) => run.write_xml(writer),
            Self::Break(line_break) => line_break.write_xml(writer),
            Self::Field(field) => field.write_xml(writer),
        }
    }
}

/// A DrawingML paragraph with ordered runs, breaks, and fields.
#[allow(non_camel_case_types)]
#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub struct CT_TextParagraph {
    pub properties: Option<CT_TextParagraphProperties>,
    pub runs: Vec<TextRun>,
    pub end_properties: Option<CT_TextCharacterProperties>,
    raw_children: OrderedRawChildren,
}

impl CT_TextParagraph {
    /// Returns direct properties, inserting them before preserved content.
    pub fn properties_mut(&mut self) -> &mut CT_TextParagraphProperties {
        if self.properties.is_none() {
            let mut raw_children = OrderedRawChildren::default();
            for boundary in 0..=2 + self.runs.len() {
                let new_boundary = if boundary == 0 { 1 } else { boundary };
                for child in self.raw_children.at(boundary) {
                    raw_children.push(new_boundary, child.to_vec());
                }
            }
            self.raw_children = raw_children;
            self.properties = Some(CT_TextParagraphProperties::default());
        }
        self.properties
            .as_mut()
            .expect("paragraph properties were inserted")
    }

    /// Replaces ordered text choices with one regular run.
    ///
    /// The first existing regular run supplies direct formatting and
    /// unmodelled run content for the replacement.
    pub fn set_text(&mut self, text: &str) {
        let old_run_count = self.runs.len();
        let retained_run = self.runs.iter().find_map(|run| match run {
            TextRun::Run(run) => Some(run.clone()),
            TextRun::Break(_) | TextRun::Field(_) => None,
        });
        let mut raw_children = OrderedRawChildren::default();
        for boundary in 0..=2 + old_run_count {
            let new_boundary = if boundary <= 1 {
                boundary
            } else if boundary <= 1 + old_run_count {
                2
            } else {
                3
            };
            for child in self.raw_children.at(boundary) {
                raw_children.push(new_boundary, child.to_vec());
            }
        }
        self.raw_children = raw_children;
        let mut run = retained_run.unwrap_or_else(|| CT_RegularTextRun::new(text));
        run.set_text(text);
        self.runs = vec![TextRun::Run(run)];
    }

    /// Appends one regular run after the existing ordered text choices.
    pub fn add_run(&mut self, text: &str) -> &mut CT_RegularTextRun {
        self.raw_children.shift_boundaries_from(2 + self.runs.len());
        self.runs.push(TextRun::Run(CT_RegularTextRun::new(text)));
        let Some(TextRun::Run(run)) = self.runs.last_mut() else {
            unreachable!("the appended text choice is a regular run")
        };
        run
    }

    /// Parses one complete `a:p` element with any prefix.
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        parse_complete(xml, b"p", Self::from_element, |_| Ok(Self::default()))
    }

    fn from_element(reader: &mut Reader<&[u8]>, _start: &BytesStart<'_>) -> Result<Self> {
        let mut paragraph = Self::default();
        let mut boundary = 0;
        let mut buffer = Vec::new();
        loop {
            match reader
                .read_event_into(&mut buffer)
                .map_err(OxmlError::from)?
            {
                Event::Start(element) if matches_local_name(element.name().as_ref(), b"pPr") => {
                    if paragraph.properties.is_some() {
                        return Err(duplicate("pPr"));
                    }
                    paragraph.properties =
                        Some(CT_TextParagraphProperties::from_element(reader, &element)?);
                    boundary = boundary.max(1);
                }
                Event::Empty(element) if matches_local_name(element.name().as_ref(), b"pPr") => {
                    if paragraph.properties.is_some() {
                        return Err(duplicate("pPr"));
                    }
                    paragraph.properties = Some(CT_TextParagraphProperties::from_start(&element)?);
                    boundary = boundary.max(1);
                }
                Event::Start(element) if is_text_run(element.name().as_ref()) => {
                    let name = local_name(element.name().as_ref()).to_vec();
                    let raw = capture_element(reader, &element)?;
                    paragraph.runs.push(TextRun::from_xml(&raw, &name)?);
                    boundary = boundary.max(1 + paragraph.runs.len());
                }
                Event::Empty(element) if is_text_run(element.name().as_ref()) => {
                    let name = local_name(element.name().as_ref()).to_vec();
                    let raw = capture_empty_element(&element)?;
                    paragraph.runs.push(TextRun::from_xml(&raw, &name)?);
                    boundary = boundary.max(1 + paragraph.runs.len());
                }
                Event::Start(element)
                    if matches_local_name(element.name().as_ref(), b"endParaRPr") =>
                {
                    if paragraph.end_properties.is_some() {
                        return Err(duplicate("endParaRPr"));
                    }
                    paragraph.end_properties =
                        Some(CT_TextCharacterProperties::from_element(reader, &element)?);
                    boundary = boundary.max(2 + paragraph.runs.len());
                }
                Event::Empty(element)
                    if matches_local_name(element.name().as_ref(), b"endParaRPr") =>
                {
                    if paragraph.end_properties.is_some() {
                        return Err(duplicate("endParaRPr"));
                    }
                    paragraph.end_properties =
                        Some(CT_TextCharacterProperties::from_start(&element)?);
                    boundary = boundary.max(2 + paragraph.runs.len());
                }
                Event::Start(element) => paragraph
                    .raw_children
                    .push(boundary, capture_element(reader, &element)?),
                Event::Empty(element) => paragraph
                    .raw_children
                    .push(boundary, capture_empty_element(&element)?),
                Event::End(element) if matches_local_name(element.name().as_ref(), b"p") => {
                    return Ok(paragraph);
                }
                Event::Eof => return Err(missing_end("p")),
                _ => {}
            }
            buffer.clear();
        }
    }

    /// Writes a paragraph with fixed prefixes and schema child order.
    pub fn write_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        if self.properties.is_none()
            && self.runs.is_empty()
            && self.end_properties.is_none()
            && self.raw_children.is_empty()
        {
            return write_empty(writer, BytesStart::new("a:p"));
        }
        write_start(writer, BytesStart::new("a:p"))?;
        emit_raw(writer, self.raw_children.at(0))?;
        if let Some(properties) = &self.properties {
            properties.write_xml(writer, "a:pPr")?;
        }
        emit_raw(writer, self.raw_children.at(1))?;
        for (index, run) in self.runs.iter().enumerate() {
            run.write_xml(writer)?;
            emit_raw(writer, self.raw_children.at(2 + index))?;
        }
        if let Some(properties) = &self.end_properties {
            properties.write_xml(writer, "a:endParaRPr")?;
        }
        emit_raw(writer, self.raw_children.at(2 + self.runs.len()))?;
        write_end(writer, "a:p")
    }

    /// Serialises a complete paragraph fragment.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        let mut writer = Writer::new(Vec::new());
        self.write_xml(&mut writer)?;
        Ok(writer.into_inner())
    }

    pub fn raw_children(&self) -> &OrderedRawChildren {
        &self.raw_children
    }
}

fn is_text_run(name: &[u8]) -> bool {
    matches!(local_name(name), b"r" | b"br" | b"fld")
}

fn parse_complete<T>(
    xml: &[u8],
    expected: &[u8],
    parse_start: impl FnOnce(&mut Reader<&[u8]>, &BytesStart<'_>) -> Result<T>,
    parse_empty: impl FnOnce(&BytesStart<'_>) -> Result<T>,
) -> Result<T> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    loop {
        match reader
            .read_event_into(&mut buffer)
            .map_err(OxmlError::from)?
        {
            Event::Start(element) if matches_local_name(element.name().as_ref(), expected) => {
                reject_conflicting_a_prefix(&element)?;
                return parse_start(&mut reader, &element);
            }
            Event::Empty(element) if matches_local_name(element.name().as_ref(), expected) => {
                reject_conflicting_a_prefix(&element)?;
                return parse_empty(&element);
            }
            Event::Start(element) | Event::Empty(element) => return Err(unexpected(&element)),
            Event::Eof => {
                return Err(TextError::Xml(OxmlError::MissingElement(
                    String::from_utf8_lossy(expected).into_owned(),
                )));
            }
            _ => {}
        }
        buffer.clear();
    }
}

fn root_local_name(xml: &[u8]) -> Result<Vec<u8>> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    loop {
        match reader
            .read_event_into(&mut buffer)
            .map_err(OxmlError::from)?
        {
            Event::Start(element) | Event::Empty(element) => {
                return Ok(local_name(element.name().as_ref()).to_vec());
            }
            Event::Eof => {
                return Err(TextError::Xml(OxmlError::MissingElement(
                    "DrawingML text properties".to_owned(),
                )));
            }
            _ => {}
        }
        buffer.clear();
    }
}

fn ensure_empty(reader: &mut Reader<&[u8]>, expected: &[u8]) -> Result<()> {
    let mut buffer = Vec::new();
    loop {
        match reader
            .read_event_into(&mut buffer)
            .map_err(OxmlError::from)?
        {
            Event::End(element) if matches_local_name(element.name().as_ref(), expected) => {
                return Ok(());
            }
            Event::Text(text) if text.iter().all(u8::is_ascii_whitespace) => {}
            Event::Comment(_) => {}
            Event::Start(element) | Event::Empty(element) => return Err(unexpected(&element)),
            Event::Eof => return Err(missing_end(&String::from_utf8_lossy(expected))),
            _ => {
                return Err(TextError::UnexpectedElement(
                    String::from_utf8_lossy(expected).into_owned(),
                ));
            }
        }
        buffer.clear();
    }
}

fn text_attr(start: &BytesStart<'_>, name: &[u8]) -> Result<Option<String>> {
    for attribute in start.attributes() {
        let attribute = attribute.map_err(OxmlError::from)?;
        if attribute.key.as_ref() == name {
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, start.decoder())
                .map_err(OxmlError::from)?;
            return Ok(Some(value.into_owned()));
        }
    }
    Ok(None)
}

fn capture_raw_attributes(
    start: &BytesStart<'_>,
    modelled: &[&[u8]],
) -> Result<Vec<(String, String)>> {
    let mut raw = Vec::new();
    for attribute in start.attributes() {
        let attribute = attribute.map_err(OxmlError::from)?;
        if modelled.iter().any(|name| attribute.key.as_ref() == *name) {
            continue;
        }
        let name = std::str::from_utf8(attribute.key.as_ref())
            .map_err(OxmlError::from)?
            .to_owned();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, start.decoder())
            .map_err(OxmlError::from)?
            .into_owned();
        raw.push((name, value));
    }
    Ok(raw)
}

fn push_raw_attributes(start: &mut BytesStart<'_>, attributes: &[(String, String)]) {
    for (name, value) in attributes {
        start.push_attribute((name.as_str(), value.as_str()));
    }
}

fn required_attr(start: &BytesStart<'_>, name: &[u8]) -> Result<String> {
    text_attr(start, name)?
        .filter(|value| !value.is_empty())
        .ok_or_else(|| {
            missing_attribute(
                &String::from_utf8_lossy(local_name(start.name().as_ref())),
                &String::from_utf8_lossy(name),
            )
        })
}

fn parse_optional_i32(
    start: &BytesStart<'_>,
    name: &[u8],
    min: i32,
    max: i32,
) -> Result<Option<i32>> {
    text_attr(start, name)?
        .map(|value| {
            parse_i32_value(
                &String::from_utf8_lossy(local_name(start.name().as_ref())),
                &String::from_utf8_lossy(name),
                &value,
                min,
                max,
            )
        })
        .transpose()
}

fn parse_i32_attr(start: &BytesStart<'_>, name: &[u8], min: i32, max: i32) -> Result<i32> {
    let value = required_attr(start, name)?;
    parse_i32_value(
        &String::from_utf8_lossy(local_name(start.name().as_ref())),
        &String::from_utf8_lossy(name),
        &value,
        min,
        max,
    )
}

fn parse_i32_value(element: &str, attribute: &str, value: &str, min: i32, max: i32) -> Result<i32> {
    let parsed = value
        .parse::<i32>()
        .map_err(|_| invalid_attribute(element, attribute, value))?;
    if (min..=max).contains(&parsed) {
        Ok(parsed)
    } else {
        Err(invalid_attribute(element, attribute, value))
    }
}

fn parse_optional_bool(start: &BytesStart<'_>, name: &[u8]) -> Result<Option<bool>> {
    text_attr(start, name)?
        .map(|value| match value.as_str() {
            "true" | "1" => Ok(true),
            "false" | "0" => Ok(false),
            _ => Err(invalid_attribute(
                &String::from_utf8_lossy(local_name(start.name().as_ref())),
                &String::from_utf8_lossy(name),
                &value,
            )),
        })
        .transpose()
}

fn parse_optional_enum<T>(
    start: &BytesStart<'_>,
    name: &[u8],
    parse: impl FnOnce(&str) -> Option<T>,
) -> Result<Option<T>> {
    let Some(value) = text_attr(start, name)? else {
        return Ok(None);
    };
    parse(&value).map(Some).ok_or_else(|| {
        invalid_attribute(
            &String::from_utf8_lossy(local_name(start.name().as_ref())),
            &String::from_utf8_lossy(name),
            &value,
        )
    })
}

fn validate_spacing_percent(value: &str) -> Result<()> {
    if is_percentage_string(value) {
        return Ok(());
    }
    parse_i32_value("spcPct", "val", value, 0, MAX_TRANSITIONAL_SPACING_PERCENT)?;
    Ok(())
}

fn validate_baseline(value: &str) -> Result<()> {
    if is_percentage_string(value) {
        return Ok(());
    }
    parse_i32_value("rPr", "baseline", value, -100_000, 100_000)?;
    Ok(())
}

fn is_percentage_string(value: &str) -> bool {
    value.strip_suffix('%').is_some_and(is_signed_decimal)
}

fn is_universal_measure(value: &str) -> bool {
    ["mm", "cm", "in", "pt", "pc", "pi"]
        .iter()
        .find_map(|unit| value.strip_suffix(unit))
        .is_some_and(is_signed_decimal)
}

fn is_signed_decimal(value: &str) -> bool {
    let unsigned = value.strip_prefix('-').unwrap_or(value);
    let mut parts = unsigned.split('.');
    let Some(integer) = parts.next() else {
        return false;
    };
    if integer.is_empty() || !integer.bytes().all(|byte| byte.is_ascii_digit()) {
        return false;
    }
    if let Some(fraction) = parts.next()
        && (fraction.is_empty() || !fraction.bytes().all(|byte| byte.is_ascii_digit()))
    {
        return false;
    }
    parts.next().is_none()
}

fn is_guid(value: &str) -> bool {
    if value.len() != 38 || !value.starts_with('{') || !value.ends_with('}') {
        return false;
    }
    value.bytes().enumerate().all(|(index, byte)| match index {
        0 => byte == b'{',
        37 => byte == b'}',
        9 | 14 | 19 | 24 => byte == b'-',
        _ => byte.is_ascii_hexdigit(),
    })
}

fn push_optional_attr<'a>(start: &mut BytesStart<'a>, name: &'a str, value: Option<&'a str>) {
    if let Some(value) = value {
        start.push_attribute((name, value));
    }
}

fn push_optional_bool(start: &mut BytesStart<'_>, name: &'static str, value: Option<bool>) {
    if let Some(value) = value {
        start.push_attribute((name, if value { "1" } else { "0" }));
    }
}

fn push_owned_attr(start: &mut BytesStart<'_>, name: &'static str, value: String) {
    start.push_attribute((name, value.as_str()));
}

fn emit_raw<'a, W: Write>(
    writer: &mut Writer<W>,
    children: impl Iterator<Item = &'a [u8]>,
) -> Result<()> {
    for child in children {
        writer.get_mut().write_all(child).map_err(OxmlError::from)?;
    }
    Ok(())
}

fn write_start<W: Write>(writer: &mut Writer<W>, start: BytesStart<'_>) -> Result<()> {
    writer
        .write_event(Event::Start(start))
        .map_err(OxmlError::from)?;
    Ok(())
}

fn write_empty<W: Write>(writer: &mut Writer<W>, start: BytesStart<'_>) -> Result<()> {
    writer
        .write_event(Event::Empty(start))
        .map_err(OxmlError::from)?;
    Ok(())
}

fn write_end<W: Write>(writer: &mut Writer<W>, tag: &str) -> Result<()> {
    writer
        .write_event(Event::End(BytesEnd::new(tag)))
        .map_err(OxmlError::from)?;
    Ok(())
}

fn unexpected(element: &BytesStart<'_>) -> TextError {
    TextError::UnexpectedElement(String::from_utf8_lossy(element.name().as_ref()).into_owned())
}

fn duplicate(element: &str) -> TextError {
    TextError::DuplicateElement(element.to_owned())
}

fn missing_attribute(element: &str, attribute: &str) -> TextError {
    TextError::MissingAttribute {
        element: element.to_owned(),
        attribute: attribute.to_owned(),
    }
}

fn invalid_attribute(element: &str, attribute: &str, value: &str) -> TextError {
    TextError::InvalidAttribute {
        element: element.to_owned(),
        attribute: attribute.to_owned(),
        value: value.to_owned(),
    }
}

#[cfg(test)]
mod tests {
    use std::panic;

    use super::{
        CT_TextCharacterProperties, CT_TextParagraph, CT_TextParagraphProperties, TextRun,
        TextSpace, TextSpacing,
    };
    use crate::color::ColorChoice;
    use crate::text::CT_TextBody;
    use crate::text::{
        TextAutoNumberScheme, TextBullet, TextBulletCharacter, TextBulletChoice,
        TextBulletSizeValue,
    };

    #[test]
    fn drawingml_rtl_attribute_becomes_typed_without_reordering_unknown_content() {
        let parsed = CT_TextParagraphProperties::from_xml(
            br#"<q:pPr xmlns:q="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:x="urn:test" defTabSz="457200" rtl="1" x:keep="yes"/>"#,
        )
        .unwrap();

        assert_eq!(parsed.right_to_left, Some(true));
        let serialized = String::from_utf8(parsed.to_xml().unwrap()).unwrap();
        assert!(serialized.contains("rtl=\"1\""));
        assert!(serialized.contains("x:keep=\"yes\""));
        assert!(
            serialized.find("defTabSz").unwrap() < serialized.find("rtl").unwrap(),
            "{serialized}"
        );

        let mut changed = parsed;
        changed.right_to_left = Some(false);
        let serialized = String::from_utf8(changed.to_xml().unwrap()).unwrap();
        assert!(serialized.contains("rtl=\"0\""), "{serialized}");
        assert!(!serialized.contains("rtl=\"1\""), "{serialized}");
    }

    #[test]
    fn leading_and_trailing_text_whitespace_survives_via_xml_space_preserve() {
        let xml = br#"<q:p><q:r><q:t> leading</q:t></q:r><q:r><q:t>trailing </q:t></q:r><q:fld id="{00112233-4455-6677-8899-AABBCCDDEEFF}"><q:t xml:space="preserve">middle</q:t></q:fld></q:p>"#;
        let paragraph = CT_TextParagraph::from_xml(xml).unwrap();
        assert_eq!(paragraph.to_xml().unwrap(), br#"<a:p><a:r><a:t xml:space="preserve"> leading</a:t></a:r><a:r><a:t xml:space="preserve">trailing </a:t></a:r><a:fld id="{00112233-4455-6677-8899-AABBCCDDEEFF}"><a:t xml:space="preserve">middle</a:t></a:fld></a:p>"#);

        let TextRun::Run(first) = &paragraph.runs[0] else {
            panic!("expected regular run");
        };
        assert_eq!(first.text.value, " leading");
        assert_eq!(first.text.space, TextSpace::Preserve);

        let body = CT_TextBody::from_xml(
            br#"<q:txBody><q:bodyPr/><q:p><q:r><q:t> body </q:t></q:r></q:p></q:txBody>"#,
        )
        .unwrap();
        assert_eq!(body.to_xml().unwrap(), br#"<a:txBody><a:bodyPr/><a:p><a:r><a:t xml:space="preserve"> body </a:t></a:r></a:p></a:txBody>"#);

        let hostile = CT_TextParagraph::from_xml(
            br#"<q:p><q:r><q:t x:space="preserve">plain</q:t></q:r></q:p>"#,
        )
        .unwrap();
        let TextRun::Run(run) = &hostile.runs[0] else {
            panic!("expected regular run");
        };
        assert_eq!(run.text.space, TextSpace::Default);
        assert_eq!(
            hostile.to_xml().unwrap(),
            br#"<a:p><a:r><a:t x:space="preserve">plain</a:t></a:r></a:p>"#
        );
    }

    #[test]
    fn canonical_text_state_matches_whitespace_and_empty_hyperlinks_from_real_decks() {
        let paragraph = CT_TextParagraph::from_xml(
            br#"<q:p><q:r><q:rPr><q:hlinkClick r:id=""/></q:rPr><q:t>trailing </q:t></q:r></q:p>"#,
        )
        .unwrap();
        let TextRun::Run(run) = &paragraph.runs[0] else {
            panic!("expected regular run");
        };
        assert_eq!(run.text.space, TextSpace::Preserve);
        assert_eq!(
            run.properties
                .as_ref()
                .and_then(|properties| properties.hyperlink_click.as_ref())
                .and_then(|hyperlink| hyperlink.relationship_id.as_deref()),
            Some("")
        );
        let written = paragraph.to_xml().unwrap();
        assert_eq!(CT_TextParagraph::from_xml(&written).unwrap(), paragraph);
    }

    #[test]
    fn paragraph_runs_fields_and_breaks_round_trip_structurally() {
        let xml = br#"<q:p><x:before/><q:pPr algn="ctr"/><x:afterPPr/><q:r><q:rPr b="1"><q:hlinkClick x:id="not-a-relationship" r:id="rId7" action="ppaction://hlinksldjump" tooltip="go &amp; stay" history="0" highlightClick="1" endSnd="0"><x:sound/></q:hlinkClick></q:rPr><q:t>one</q:t></q:r><x:between/><q:br><q:rPr i="true"/></q:br><q:fld id="{00112233-4455-6677-8899-AABBCCDDEEFF}" type="slidenum"><q:rPr sz="1200"/><q:pPr lvl="1"/><q:t>2</q:t></q:fld><q:endParaRPr sz="1400"/><x:after/></q:p>"#;
        let paragraph = CT_TextParagraph::from_xml(xml).unwrap();
        assert_eq!(paragraph.runs.len(), 3);
        let written = paragraph.to_xml().unwrap();
        assert_eq!(written, br#"<a:p><x:before/><a:pPr algn="ctr"/><x:afterPPr/><a:r><a:rPr b="1"><a:hlinkClick r:id="rId7" action="ppaction://hlinksldjump" tooltip="go &amp; stay" history="0" highlightClick="1" endSnd="0" x:id="not-a-relationship"><x:sound/></a:hlinkClick></a:rPr><a:t>one</a:t></a:r><x:between/><a:br><a:rPr i="1"/></a:br><a:fld id="{00112233-4455-6677-8899-AABBCCDDEEFF}" type="slidenum"><a:rPr sz="1200"/><a:pPr lvl="1"/><a:t>2</a:t></a:fld><a:endParaRPr sz="1400"/><x:after/></a:p>"#);
        assert_eq!(CT_TextParagraph::from_xml(&written).unwrap(), paragraph);

        let hostile = CT_TextCharacterProperties::from_xml(
            br#"<q:rPr><q:hlinkClick x:id="not-a-relationship"/></q:rPr>"#,
        )
        .unwrap();
        assert!(
            hostile
                .hyperlink_click
                .as_ref()
                .unwrap()
                .relationship_id
                .is_none()
        );
        let mut writer = quick_xml::Writer::new(Vec::new());
        hostile.write_xml(&mut writer, "a:rPr").unwrap();
        assert_eq!(
            writer.into_inner(),
            br#"<a:rPr><a:hlinkClick x:id="not-a-relationship"/></a:rPr>"#
        );
    }

    #[test]
    fn character_all_caps_round_trips_while_small_caps_stays_unmodelled() {
        let all = CT_TextCharacterProperties::from_xml(
            br#"<q:rPr cap="all" sz="1800" x:producer="kept"/>"#,
        )
        .unwrap();
        assert_eq!(all.all_caps, Some(true));
        let mut writer = quick_xml::Writer::new(Vec::new());
        all.write_xml(&mut writer, "a:rPr").unwrap();
        let written = writer.into_inner();
        assert_eq!(
            written,
            br#"<a:rPr sz="1800" cap="all" x:producer="kept"/>"#
        );
        assert_eq!(CT_TextCharacterProperties::from_xml(&written).unwrap(), all);

        let none = CT_TextCharacterProperties::from_xml(br#"<q:rPr cap="none"/>"#).unwrap();
        assert_eq!(none.all_caps, Some(false));

        let small = CT_TextCharacterProperties::from_xml(br#"<q:rPr cap="small"/>"#).unwrap();
        assert_eq!(small.all_caps, None);
        let mut writer = quick_xml::Writer::new(Vec::new());
        small.write_xml(&mut writer, "a:rPr").unwrap();
        assert_eq!(writer.into_inner(), br#"<a:rPr cap="small"/>"#);
    }

    #[test]
    fn setting_all_caps_over_preserved_small_caps_writes_one_cap_attribute() {
        let mut properties =
            CT_TextCharacterProperties::from_xml(br#"<q:rPr cap="small" lang="en-US"/>"#).unwrap();
        properties.all_caps = Some(true);
        let mut writer = quick_xml::Writer::new(Vec::new());
        properties.write_xml(&mut writer, "a:rPr").unwrap();
        let written = writer.into_inner();
        assert_eq!(written, br#"<a:rPr cap="all" lang="en-US"/>"#);
        assert_eq!(
            CT_TextCharacterProperties::from_xml(&written)
                .unwrap()
                .all_caps,
            Some(true)
        );

        properties.all_caps = None;
        let mut writer = quick_xml::Writer::new(Vec::new());
        properties.write_xml(&mut writer, "a:rPr").unwrap();
        assert_eq!(writer.into_inner(), br#"<a:rPr cap="small" lang="en-US"/>"#);

        // The setter drops the preserved value, so clearing afterwards leaves
        // no capitalisation, saved or not.
        properties.set_all_caps(Some(true));
        properties.set_all_caps(None);
        let mut writer = quick_xml::Writer::new(Vec::new());
        properties.write_xml(&mut writer, "a:rPr").unwrap();
        assert_eq!(writer.into_inner(), br#"<a:rPr lang="en-US"/>"#);
    }

    #[test]
    fn spacing_percentages_above_two_lines_parse_up_to_the_schema_maximum() {
        let xml = br#"<q:pPr><q:lnSpc><q:spcPct val="250000"/></q:lnSpc><q:spcBef><q:spcPct val="13200000"/></q:spcBef></q:pPr>"#;
        let properties = CT_TextParagraphProperties::from_xml(xml).unwrap();
        assert_eq!(
            properties.line_spacing,
            Some(TextSpacing::Percent("250000".to_owned()))
        );
        assert_eq!(
            properties.space_before,
            Some(TextSpacing::Percent("13200000".to_owned()))
        );
        assert_eq!(
            properties.to_xml().unwrap(),
            br#"<a:pPr><a:lnSpc><a:spcPct val="250000"/></a:lnSpc><a:spcBef><a:spcPct val="13200000"/></a:spcBef></a:pPr>"#
        );
    }

    #[test]
    fn setting_a_bullet_part_removes_the_preserved_alternative_it_replaces() {
        let mut properties = CT_TextParagraphProperties::from_xml(
            br#"<q:pPr><q:buClrTx/><q:buSzTx/><q:buFontTx/><q:buBlip><q:blip x:embed="rId9"/></q:buBlip><q:tabLst/></q:pPr>"#,
        )
        .unwrap();
        assert!(properties.has_picture_bullet());

        properties.set_bullet(Some(TextBullet {
            choice: Some(TextBulletChoice::Character(
                TextBulletCharacter::new("*").unwrap(),
            )),
            ..TextBullet::default()
        }));
        assert!(!properties.has_picture_bullet());
        assert_eq!(
            properties.to_xml().unwrap(),
            br#"<a:pPr><q:buClrTx/><q:buSzTx/><q:buFontTx/><a:buChar char="*"/><q:tabLst/></a:pPr>"#
        );

        properties.set_bullet(None);
        assert_eq!(
            properties.to_xml().unwrap(),
            br#"<a:pPr><q:tabLst/></a:pPr>"#
        );
    }

    #[test]
    fn a_bullet_member_set_through_the_public_field_still_replaces_its_alternative() {
        // Assigning the field bypasses `set_bullet`, so the writer itself keeps
        // one member per bullet group.
        let mut properties = CT_TextParagraphProperties::from_xml(
            br#"<q:pPr><q:buFontTx/><q:buBlip><q:blip x:embed="rId9"/></q:buBlip></q:pPr>"#,
        )
        .unwrap();
        properties.bullet = Some(TextBullet {
            font: Some(super::TextFont::new("Wingdings").unwrap()),
            choice: Some(TextBulletChoice::Character(
                TextBulletCharacter::new("*").unwrap(),
            )),
            ..TextBullet::default()
        });

        assert!(!properties.has_picture_bullet());
        assert_eq!(
            properties.to_xml().unwrap(),
            br#"<a:pPr><a:buFont typeface="Wingdings"/><a:buChar char="*"/></a:pPr>"#
        );
    }

    #[test]
    fn typefaces_and_bullet_characters_refuse_characters_xml_cannot_carry() {
        assert!(super::TextFont::new("Aria\u{1}l").is_err());
        assert!(TextBulletCharacter::new("\u{b}").is_err());
        assert!(super::TextFont::new("Aptos Display").is_ok());
    }

    #[test]
    fn paragraph_and_run_properties_use_drawingml_units_and_schema_order() {
        let paragraph_properties = CT_TextParagraphProperties::from_xml(br#"<q:pPr marL="0" marR="51206400" lvl="8" indent="-51206400" algn="thaiDist" rtl="1" defTabSz="914400"><x:before/><q:lnSpc><q:spcPct val="100000"/></q:lnSpc><q:spcBef><q:spcPts val="158400"/></q:spcBef><q:spcAft><q:spcPct val="125.5%"/></q:spcAft><q:buChar char="*"/><x:afterBullet/><q:defRPr sz="400000" b="true" i="0" u="wavyDbl" strike="dblStrike" spc="-400000" baseline="25%" xmlns:b="urn:test" lang="en-US" kern="1200"><q:solidFill><q:srgbClr val="AABBCC"/></q:solidFill><q:latin typeface="Aptos" pitchFamily="34" charset="0"/><q:ea typeface="Yu Gothic"/><q:cs typeface="Arial"/><q:sym typeface="Symbol"/><q:hlinkMouseOver r:id="rId8" tgtFrame="_blank" invalidUrl="bad"/></q:defRPr><x:last/></q:pPr>"#).unwrap();
        let written = {
            let mut writer = quick_xml::Writer::new(Vec::new());
            paragraph_properties
                .write_xml(&mut writer, "a:pPr")
                .unwrap();
            writer.into_inner()
        };
        assert_eq!(written, br#"<a:pPr marL="0" marR="51206400" lvl="8" indent="-51206400" algn="thaiDist" rtl="1" defTabSz="914400"><x:before/><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPts val="158400"/></a:spcBef><a:spcAft><a:spcPct val="125.5%"/></a:spcAft><a:buChar char="*"/><x:afterBullet/><a:defRPr sz="400000" b="1" i="0" u="wavyDbl" strike="dblStrike" spc="-400000" baseline="25%" xmlns:b="urn:test" lang="en-US" kern="1200"><a:solidFill><a:srgbClr val="AABBCC"/></a:solidFill><a:latin typeface="Aptos" pitchFamily="34" charset="0"/><a:ea typeface="Yu Gothic"/><a:cs typeface="Arial"/><a:sym typeface="Symbol"/><a:hlinkMouseOver r:id="rId8" invalidUrl="bad" tgtFrame="_blank"/></a:defRPr><x:last/></a:pPr>"#);

        let run_properties = CT_TextCharacterProperties::from_xml(
            br#"<q:rPr sz="100" spc="1.25pt" baseline="-100000"/>"#,
        )
        .unwrap();
        let mut writer = quick_xml::Writer::new(Vec::new());
        run_properties.write_xml(&mut writer, "a:rPr").unwrap();
        assert_eq!(
            writer.into_inner(),
            br#"<a:rPr sz="100" spc="1.25pt" baseline="-100000"/>"#
        );

        let mut inserted_spacing =
            CT_TextParagraphProperties::from_xml(br#"<q:pPr><q:buChar char="*"/></q:pPr>"#)
                .unwrap();
        inserted_spacing.line_spacing = Some(TextSpacing::Percent("100000".to_owned()));
        let mut writer = quick_xml::Writer::new(Vec::new());
        inserted_spacing.write_xml(&mut writer, "a:pPr").unwrap();
        assert_eq!(
            writer.into_inner(),
            br#"<a:pPr><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:buChar char="*"/></a:pPr>"#
        );
    }

    #[test]
    fn every_modelled_bullet_form_round_trips_in_schema_order() {
        let cases: &[(&[u8], &[u8])] = &[
            (
                br#"<q:pPr><q:buClr><q:srgbClr val="AABBCC"/></q:buClr><q:buSzPct val="125000"/><q:buFont typeface="Wingdings"/><q:buChar char="*"/></q:pPr>"#,
                br#"<a:pPr><a:buClr><a:srgbClr val="AABBCC"/></a:buClr><a:buSzPct val="125000"/><a:buFont typeface="Wingdings"/><a:buChar char="*"/></a:pPr>"#,
            ),
            (
                br#"<q:pPr><q:buSzPts val="1800"/><q:buAutoNum type="arabicPeriod" startAt="3"/></q:pPr>"#,
                br#"<a:pPr><a:buSzPts val="1800"/><a:buAutoNum type="arabicPeriod" startAt="3"/></a:pPr>"#,
            ),
            (br#"<q:pPr><q:buNone/></q:pPr>"#, br#"<a:pPr><a:buNone/></a:pPr>"#),
        ];

        for (xml, expected) in cases {
            let properties = CT_TextParagraphProperties::from_xml(xml).unwrap();
            let mut writer = quick_xml::Writer::new(Vec::new());
            properties.write_xml(&mut writer, "a:pPr").unwrap();
            let written = writer.into_inner();
            assert_eq!(written, *expected);
            assert_eq!(
                CT_TextParagraphProperties::from_xml(&written).unwrap(),
                properties
            );
        }

        let character = CT_TextParagraphProperties::from_xml(cases[0].0).unwrap();
        let bullet = character.bullet.as_ref().unwrap();
        assert!(matches!(
            bullet.choice,
            Some(TextBulletChoice::Character(_))
        ));
        assert_eq!(bullet.font.as_ref().unwrap().typeface, "Wingdings");
        assert!(matches!(
            bullet.size.as_ref().unwrap().value,
            TextBulletSizeValue::Percent(ref value) if value == "125000"
        ));
        assert!(matches!(
            bullet.color.as_ref().unwrap().color,
            ColorChoice::Srgb { .. }
        ));

        let automatic = CT_TextParagraphProperties::from_xml(cases[1].0).unwrap();
        let Some(TextBulletChoice::AutoNumber(numbering)) =
            automatic.bullet.as_ref().unwrap().choice.as_ref()
        else {
            panic!("expected automatic-number bullet");
        };
        assert_eq!(numbering.scheme, TextAutoNumberScheme::ArabicPeriod);
        assert_eq!(numbering.start_at, Some(3));
        assert!(matches!(
            automatic
                .bullet
                .as_ref()
                .unwrap()
                .size
                .as_ref()
                .unwrap()
                .value,
            TextBulletSizeValue::Points(1800)
        ));

        let none = CT_TextParagraphProperties::from_xml(cases[2].0).unwrap();
        assert!(matches!(
            none.bullet.as_ref().unwrap().choice,
            Some(TextBulletChoice::None(_))
        ));
    }

    #[test]
    fn bullet_font_size_and_colour_keep_their_schema_positions() {
        let properties = CT_TextParagraphProperties::from_xml(
            br#"<q:pPr><q:buFont typeface="Aptos"/><q:buChar char="*"/><q:buClr><q:schemeClr val="accent1"/></q:buClr><q:buSzPct val="80%"/><q:defRPr b="1"/></q:pPr>"#,
        )
        .unwrap();
        let mut writer = quick_xml::Writer::new(Vec::new());
        properties.write_xml(&mut writer, "a:pPr").unwrap();
        assert_eq!(
            writer.into_inner(),
            br#"<a:pPr><a:buClr><a:schemeClr val="accent1"/></a:buClr><a:buSzPct val="80%"/><a:buFont typeface="Aptos"/><a:buChar char="*"/><a:defRPr b="1"/></a:pPr>"#
        );
    }

    #[test]
    fn unknown_bullet_children_round_trip_byte_for_byte() {
        let properties = CT_TextParagraphProperties::from_xml(
            br#"<q:pPr><q:buClrTx x:mode="stay"><x:nested>one &amp; two</x:nested><!--note--></q:buClrTx><q:buChar char="*"/><x:after/></q:pPr>"#,
        )
        .unwrap();
        let mut writer = quick_xml::Writer::new(Vec::new());
        properties.write_xml(&mut writer, "a:pPr").unwrap();
        assert_eq!(
            writer.into_inner(),
            br#"<a:pPr><q:buClrTx x:mode="stay"><x:nested>one &amp; two</x:nested><!--note--></q:buClrTx><a:buChar char="*"/><x:after/></a:pPr>"#
        );
    }

    #[test]
    fn malformed_bullet_values_return_errors_without_panicking() {
        let cases: &[&[u8]] = &[
            br#"<q:pPr><q:buChar/></q:pPr>"#,
            br#"<q:pPr><q:buChar char=""/></q:pPr>"#,
            br#"<q:pPr><q:buSzPct val="24999"/></q:pPr>"#,
            br#"<q:pPr><q:buSzPct val="400001"/></q:pPr>"#,
            br#"<q:pPr><q:buSzPts val="99"/></q:pPr>"#,
            br#"<q:pPr><q:buSzPts val="400001"/></q:pPr>"#,
            br#"<q:pPr><q:buAutoNum type="unknown"/></q:pPr>"#,
            br#"<q:pPr><q:buAutoNum/></q:pPr>"#,
            br#"<q:pPr><q:buAutoNum type="arabicPeriod" startAt="0"/></q:pPr>"#,
            br#"<q:pPr><q:buAutoNum type="arabicPeriod" startAt="32768"/></q:pPr>"#,
            br#"<q:pPr><q:buFont/></q:pPr>"#,
            br#"<q:pPr><q:buClr/></q:pPr>"#,
            br#"<q:pPr><q:buClr><q:srgbClr x:val="AABBCC"/></q:buClr></q:pPr>"#,
            br#"<q:pPr><q:buClr><q:srgbClr val="GGGGGG"/></q:buClr></q:pPr>"#,
            br#"<q:pPr><q:buClr><q:srgbClr val="AABBCC"/><q:schemeClr val="accent1"/></q:buClr></q:pPr>"#,
            br#"<q:pPr><q:buChar char="*"/><q:buNone/></q:pPr>"#,
            br#"<q:pPr><q:buSzPct val="100000"/><q:buSzPts val="1200"/></q:pPr>"#,
        ];

        for xml in cases {
            let result = panic::catch_unwind(|| CT_TextParagraphProperties::from_xml(xml));
            assert!(result.is_ok(), "bullet parser panicked");
            assert!(result.unwrap().is_err(), "malformed bullet parsed");
        }
    }

    #[test]
    fn malformed_text_content_returns_errors_without_panicking() {
        let cases: &[&[u8]] = &[
            br#"<q:r/>"#,
            br#"<q:r><q:t><x:nested/></q:t></q:r>"#,
            br#"<q:r><q:t xml:space="keep">x</q:t></q:r>"#,
            br#"<q:r><q:t>x</q:t><q:t>y</q:t></q:r>"#,
            br#"<q:pPr lvl="9"/>"#,
            br#"<q:pPr marL="-1"/>"#,
            br#"<q:pPr indent="51206401"/>"#,
            br#"<q:pPr algn="middle"/>"#,
            br#"<q:pPr><q:lnSpc/></q:pPr>"#,
            br#"<q:pPr><q:lnSpc><q:spcPts val="158401"/></q:lnSpc></q:pPr>"#,
            br#"<q:pPr><q:lnSpc><q:spcPct val="13200001"/></q:lnSpc></q:pPr>"#,
            br#"<q:rPr sz="99"/>"#,
            br#"<q:rPr b="yes"/>"#,
            br#"<q:rPr u="triple"/>"#,
            br#"<q:rPr spc="400001"/>"#,
            r#"<q:rPr spc="éx"/>"#.as_bytes(),
            br#"<q:rPr baseline="100001"/>"#,
            br#"<q:rPr><q:latin/></q:rPr>"#,
            br#"<q:fld type="slidenum"/>"#,
            br#"<q:fld id="not-a-guid"/>"#,
            br#"<q:p><q:endParaRPr/><q:endParaRPr/></q:p>"#,
            br#"<q:notPPr/>"#,
            br#"<q:notRPr/>"#,
        ];
        for xml in cases {
            let result = panic::catch_unwind(|| match super::root_local_name(xml).as_deref() {
                Ok(b"r") => super::CT_RegularTextRun::from_xml(xml).map(|_| ()),
                Ok(b"pPr") => CT_TextParagraphProperties::from_xml(xml).map(|_| ()),
                Ok(b"rPr") => CT_TextCharacterProperties::from_xml(xml).map(|_| ()),
                Ok(b"fld") => super::CT_TextField::from_xml(xml).map(|_| ()),
                Ok(b"p") => CT_TextParagraph::from_xml(xml).map(|_| ()),
                Ok(b"notPPr") => CT_TextParagraphProperties::from_xml(xml).map(|_| ()),
                Ok(b"notRPr") => CT_TextCharacterProperties::from_xml(xml).map(|_| ()),
                _ => unreachable!(),
            });
            assert!(result.is_ok(), "text parser panicked");
            assert!(result.unwrap().is_err(), "malformed text parsed");
        }
    }
}
