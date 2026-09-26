use std::fmt;
use std::io::Write;

use oxml_core::OxmlError;
use oxml_core::raw_xml::{capture_element, capture_empty_element};
use oxml_core::xml::{get_attr, local_name, matches_local_name};
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};

use crate::color::ColorError;
use crate::fill::FillError;
use crate::order::OrderedRawChildren;

const MAX_TEXT_SPACING_PERCENT: i32 = 13_200_000;

/// Errors produced while parsing or writing DrawingML text shells.
#[derive(Debug)]
pub enum TextError {
    Xml(OxmlError),
    Color(ColorError),
    Fill(FillError),
    UnexpectedElement(String),
    MissingAttribute {
        element: String,
        attribute: String,
    },
    MissingBodyProperties,
    MissingParagraph,
    DuplicateElement(String),
    InvalidAttribute {
        element: String,
        attribute: String,
        value: String,
    },
}

impl fmt::Display for TextError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Xml(error) => error.fmt(formatter),
            Self::Color(error) => error.fmt(formatter),
            Self::Fill(error) => error.fmt(formatter),
            Self::UnexpectedElement(element) => {
                write!(formatter, "unexpected DrawingML text element: {element}")
            }
            Self::MissingAttribute { element, attribute } => {
                write!(formatter, "DrawingML {element} requires @{attribute}")
            }
            Self::MissingBodyProperties => write!(formatter, "DrawingML txBody requires bodyPr"),
            Self::MissingParagraph => write!(formatter, "DrawingML txBody requires at least one p"),
            Self::DuplicateElement(element) => {
                write!(formatter, "DrawingML text contains duplicate {element}")
            }
            Self::InvalidAttribute {
                element,
                attribute,
                value,
            } => write!(
                formatter,
                "DrawingML {element} has invalid @{attribute}: {value}"
            ),
        }
    }
}

impl std::error::Error for TextError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Xml(error) => Some(error),
            Self::Color(error) => Some(error),
            Self::Fill(error) => Some(error),
            _ => None,
        }
    }
}

impl From<OxmlError> for TextError {
    fn from(error: OxmlError) -> Self {
        Self::Xml(error)
    }
}

impl From<ColorError> for TextError {
    fn from(error: ColorError) -> Self {
        Self::Color(error)
    }
}

impl From<FillError> for TextError {
    fn from(error: FillError) -> Self {
        Self::Fill(error)
    }
}

pub type Result<T> = std::result::Result<T, TextError>;

/// One strict or transitional lexical form of `ST_Coordinate32`.
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum Coordinate32Value {
    Emu(i32),
    UniversalMeasure(String),
}

impl Coordinate32Value {
    fn parse(element: &str, attribute: &str, value: String) -> Result<Self> {
        if let Ok(value) = value.parse::<i32>() {
            return Ok(Self::Emu(value));
        }
        if is_universal_measure(&value) {
            return Ok(Self::UniversalMeasure(value));
        }
        Err(invalid_attribute(element, attribute, value))
    }

    fn as_xml(&self) -> Result<String> {
        match self {
            Self::Emu(value) => Ok(value.to_string()),
            Self::UniversalMeasure(value) if is_universal_measure(value) => Ok(value.clone()),
            Self::UniversalMeasure(value) => {
                Err(invalid_attribute("bodyPr", "inset", value.clone()))
            }
        }
    }
}

/// Vertical anchoring inside the text rectangle.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum TextAnchor {
    Top,
    Center,
    Bottom,
    Justified,
    Distributed,
}

impl TextAnchor {
    fn parse(value: &str) -> Option<Self> {
        match value {
            "t" => Some(Self::Top),
            "ctr" => Some(Self::Center),
            "b" => Some(Self::Bottom),
            "just" => Some(Self::Justified),
            "dist" => Some(Self::Distributed),
            _ => None,
        }
    }

    const fn as_str(self) -> &'static str {
        match self {
            Self::Top => "t",
            Self::Center => "ctr",
            Self::Bottom => "b",
            Self::Justified => "just",
            Self::Distributed => "dist",
        }
    }
}

/// Text wrapping at the shape boundary.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum TextWrap {
    None,
    Square,
}

impl TextWrap {
    fn parse(value: &str) -> Option<Self> {
        match value {
            "none" => Some(Self::None),
            "square" => Some(Self::Square),
            _ => None,
        }
    }

    const fn as_str(self) -> &'static str {
        match self {
            Self::None => "none",
            Self::Square => "square",
        }
    }
}

/// The seven values of `ST_TextVerticalType`.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum TextVertical {
    Horizontal,
    Vertical,
    Vertical270,
    WordArtVertical,
    EastAsianVertical,
    MongolianVertical,
    WordArtVerticalRtl,
}

impl TextVertical {
    fn parse(value: &str) -> Option<Self> {
        match value {
            "horz" => Some(Self::Horizontal),
            "vert" => Some(Self::Vertical),
            "vert270" => Some(Self::Vertical270),
            "wordArtVert" => Some(Self::WordArtVertical),
            "eaVert" => Some(Self::EastAsianVertical),
            "mongolianVert" => Some(Self::MongolianVertical),
            "wordArtVertRtl" => Some(Self::WordArtVerticalRtl),
            _ => None,
        }
    }

    const fn as_str(self) -> &'static str {
        match self {
            Self::Horizontal => "horz",
            Self::Vertical => "vert",
            Self::Vertical270 => "vert270",
            Self::WordArtVertical => "wordArtVert",
            Self::EastAsianVertical => "eaVert",
            Self::MongolianVertical => "mongolianVert",
            Self::WordArtVerticalRtl => "wordArtVertRtl",
        }
    }
}

/// Stored values on one `a:normAutofit` choice.
#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub struct NormalAutofit {
    pub font_scale: Option<String>,
    pub line_spacing_reduction: Option<String>,
}

/// The three members of `EG_TextAutofit`.
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum TextAutofit {
    NoAutofit,
    Normal(NormalAutofit),
    ShapeAutofit,
}

impl TextAutofit {
    fn from_xml(xml: &[u8]) -> Result<Self> {
        let mut reader = Reader::from_reader(xml);
        let mut buffer = Vec::new();
        loop {
            match reader
                .read_event_into(&mut buffer)
                .map_err(OxmlError::from)?
            {
                Event::Empty(element) => return Self::from_start(&element),
                Event::Start(element) => {
                    let autofit = Self::from_start(&element)?;
                    ensure_empty_element(&mut reader, local_name(element.name().as_ref()))?;
                    return Ok(autofit);
                }
                Event::Eof => {
                    return Err(TextError::UnexpectedElement("EOF".to_owned()));
                }
                _ => {}
            }
            buffer.clear();
        }
    }

    fn from_start(start: &BytesStart<'_>) -> Result<Self> {
        match local_name(start.name().as_ref()) {
            b"noAutofit" => Ok(Self::NoAutofit),
            b"spAutoFit" => Ok(Self::ShapeAutofit),
            b"normAutofit" => {
                let font_scale = get_attr(start, b"fontScale");
                let line_spacing_reduction = get_attr(start, b"lnSpcReduction");
                if let Some(value) = font_scale.as_deref() {
                    validate_autofit_percent("fontScale", value, PercentKind::FontScale)?;
                }
                if let Some(value) = line_spacing_reduction.as_deref() {
                    validate_autofit_percent("lnSpcReduction", value, PercentKind::LineSpacing)?;
                }
                Ok(Self::Normal(NormalAutofit {
                    font_scale,
                    line_spacing_reduction,
                }))
            }
            _ => Err(TextError::UnexpectedElement(element_name(start))),
        }
    }

    fn write_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        match self {
            Self::NoAutofit => write_empty(writer, BytesStart::new("a:noAutofit")),
            Self::ShapeAutofit => write_empty(writer, BytesStart::new("a:spAutoFit")),
            Self::Normal(normal) => {
                let mut start = BytesStart::new("a:normAutofit");
                if let Some(value) = normal.font_scale.as_deref() {
                    validate_autofit_percent("fontScale", value, PercentKind::FontScale)?;
                    start.push_attribute(("fontScale", value));
                }
                if let Some(value) = normal.line_spacing_reduction.as_deref() {
                    validate_autofit_percent("lnSpcReduction", value, PercentKind::LineSpacing)?;
                    start.push_attribute(("lnSpcReduction", value));
                }
                write_empty(writer, start)
            }
        }
    }
}

/// Insets, anchoring, wrapping, vertical direction, and autofit on `a:bodyPr`.
#[allow(non_camel_case_types)]
#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub struct CT_TextBodyProperties {
    pub left_inset: Option<Coordinate32Value>,
    pub top_inset: Option<Coordinate32Value>,
    pub right_inset: Option<Coordinate32Value>,
    pub bottom_inset: Option<Coordinate32Value>,
    pub anchor: Option<TextAnchor>,
    pub wrap: Option<TextWrap>,
    pub vertical: Option<TextVertical>,
    pub space_first_last_paragraph: Option<bool>,
    pub autofit: Option<TextAutofit>,
    raw_children: OrderedRawChildren,
}

impl CT_TextBodyProperties {
    /// Parses one complete `a:bodyPr` element with any namespace prefix.
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        let mut reader = Reader::from_reader(xml);
        let mut buffer = Vec::new();
        loop {
            match reader
                .read_event_into(&mut buffer)
                .map_err(OxmlError::from)?
            {
                Event::Start(element) if matches_local_name(element.name().as_ref(), b"bodyPr") => {
                    return Self::from_element(&mut reader, &element);
                }
                Event::Empty(element) if matches_local_name(element.name().as_ref(), b"bodyPr") => {
                    return Self::from_start(&element);
                }
                Event::Start(element) | Event::Empty(element) => {
                    return Err(TextError::UnexpectedElement(element_name(&element)));
                }
                Event::Eof => {
                    return Err(TextError::Xml(OxmlError::MissingElement(
                        "DrawingML body properties".to_owned(),
                    )));
                }
                _ => {}
            }
            buffer.clear();
        }
    }

    pub(crate) fn from_element(reader: &mut Reader<&[u8]>, start: &BytesStart<'_>) -> Result<Self> {
        let mut properties = Self::from_start(start)?;
        let mut boundary = 0;
        let mut occurrences = [false; 5];
        let mut buffer = Vec::new();
        loop {
            match reader
                .read_event_into(&mut buffer)
                .map_err(OxmlError::from)?
            {
                Event::Start(element) => {
                    let name = local_name(element.name().as_ref()).to_vec();
                    let raw = capture_element(reader, &element)?;
                    properties.capture_child(&name, raw, &mut boundary, &mut occurrences)?;
                }
                Event::Empty(element) => {
                    let name = local_name(element.name().as_ref()).to_vec();
                    let raw = capture_empty_element(&element)?;
                    properties.capture_child(&name, raw, &mut boundary, &mut occurrences)?;
                }
                Event::End(element) if matches_local_name(element.name().as_ref(), b"bodyPr") => {
                    break;
                }
                Event::Eof => return Err(missing_end("bodyPr")),
                _ => {}
            }
            buffer.clear();
        }
        Ok(properties)
    }

    pub(crate) fn from_start(start: &BytesStart<'_>) -> Result<Self> {
        if !matches_local_name(start.name().as_ref(), b"bodyPr") {
            return Err(TextError::UnexpectedElement(element_name(start)));
        }
        Ok(Self {
            left_inset: parse_coordinate(start, b"lIns")?,
            top_inset: parse_coordinate(start, b"tIns")?,
            right_inset: parse_coordinate(start, b"rIns")?,
            bottom_inset: parse_coordinate(start, b"bIns")?,
            anchor: parse_enum(start, b"anchor", TextAnchor::parse)?,
            wrap: parse_enum(start, b"wrap", TextWrap::parse)?,
            vertical: parse_enum(start, b"vert", TextVertical::parse)?,
            space_first_last_paragraph: parse_optional_bool(start, b"spcFirstLastPara")?,
            ..Self::default()
        })
    }

    fn capture_child(
        &mut self,
        name: &[u8],
        raw: Vec<u8>,
        boundary: &mut usize,
        occurrences: &mut [bool; 5],
    ) -> Result<()> {
        if let Some(index) = schema_choice_index(name) {
            if occurrences[index] {
                return Err(TextError::DuplicateElement(
                    String::from_utf8_lossy(name).into_owned(),
                ));
            }
            occurrences[index] = true;
        }
        if is_autofit(name) {
            self.autofit = Some(TextAutofit::from_xml(&raw)?);
            *boundary = (*boundary).max(2);
            return Ok(());
        }
        // A known child keeps its schema slot, so an autofit choice added
        // after parsing is still written before a preserved scene or extension.
        let slot = raw_boundary_after(name).saturating_sub(1);
        self.raw_children.push((*boundary).max(slot), raw);
        *boundary = (*boundary).max(raw_boundary_after(name));
        Ok(())
    }

    /// Writes body properties with fixed prefixes and schema child order.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        let mut writer = Writer::new(Vec::new());
        self.write_xml(&mut writer)?;
        Ok(writer.into_inner())
    }

    pub(crate) fn write_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut start = BytesStart::new("a:bodyPr");
        let left_inset = self
            .left_inset
            .as_ref()
            .map(Coordinate32Value::as_xml)
            .transpose()?;
        let top_inset = self
            .top_inset
            .as_ref()
            .map(Coordinate32Value::as_xml)
            .transpose()?;
        let right_inset = self
            .right_inset
            .as_ref()
            .map(Coordinate32Value::as_xml)
            .transpose()?;
        let bottom_inset = self
            .bottom_inset
            .as_ref()
            .map(Coordinate32Value::as_xml)
            .transpose()?;
        for (name, value) in [
            ("lIns", left_inset.as_deref()),
            ("tIns", top_inset.as_deref()),
            ("rIns", right_inset.as_deref()),
            ("bIns", bottom_inset.as_deref()),
        ] {
            if let Some(value) = value {
                start.push_attribute((name, value));
            }
        }
        if let Some(anchor) = self.anchor {
            start.push_attribute(("anchor", anchor.as_str()));
        }
        if let Some(wrap) = self.wrap {
            start.push_attribute(("wrap", wrap.as_str()));
        }
        if let Some(vertical) = self.vertical {
            start.push_attribute(("vert", vertical.as_str()));
        }
        if let Some(value) = self.space_first_last_paragraph {
            start.push_attribute(("spcFirstLastPara", if value { "1" } else { "0" }));
        }

        if self.autofit.is_none() && self.raw_children.is_empty() {
            return write_empty(writer, start);
        }
        writer
            .write_event(Event::Start(start))
            .map_err(OxmlError::from)?;
        emit_raw(writer, self.raw_children.at(0))?;
        emit_raw(writer, self.raw_children.at(1))?;
        if let Some(autofit) = &self.autofit {
            autofit.write_xml(writer)?;
        }
        for boundary in 2..=5 {
            emit_raw(writer, self.raw_children.at(boundary))?;
        }
        writer
            .write_event(Event::End(BytesEnd::new("a:bodyPr")))
            .map_err(OxmlError::from)?;
        Ok(())
    }

    pub fn raw_children(&self) -> &OrderedRawChildren {
        &self.raw_children
    }
}

fn parse_coordinate(start: &BytesStart<'_>, attribute: &[u8]) -> Result<Option<Coordinate32Value>> {
    get_attr(start, attribute)
        .map(|value| Coordinate32Value::parse("bodyPr", &String::from_utf8_lossy(attribute), value))
        .transpose()
}

fn parse_enum<T>(
    start: &BytesStart<'_>,
    attribute: &[u8],
    parse: impl FnOnce(&str) -> Option<T>,
) -> Result<Option<T>> {
    let Some(value) = get_attr(start, attribute) else {
        return Ok(None);
    };
    parse(&value)
        .map(Some)
        .ok_or_else(|| invalid_attribute("bodyPr", &String::from_utf8_lossy(attribute), value))
}

fn parse_optional_bool(start: &BytesStart<'_>, attribute: &[u8]) -> Result<Option<bool>> {
    let Some(value) = get_attr(start, attribute) else {
        return Ok(None);
    };
    match value.as_str() {
        "1" | "true" => Ok(Some(true)),
        "0" | "false" => Ok(Some(false)),
        _ => Err(invalid_attribute(
            "bodyPr",
            &String::from_utf8_lossy(attribute),
            value,
        )),
    }
}

fn is_autofit(name: &[u8]) -> bool {
    matches!(name, b"noAutofit" | b"normAutofit" | b"spAutoFit")
}

fn schema_choice_index(name: &[u8]) -> Option<usize> {
    match name {
        b"prstTxWarp" => Some(0),
        name if is_autofit(name) => Some(1),
        b"scene3d" => Some(2),
        b"sp3d" | b"flatTx" => Some(3),
        b"extLst" => Some(4),
        _ => None,
    }
}

fn raw_boundary_after(name: &[u8]) -> usize {
    match name {
        b"prstTxWarp" => 1,
        name if is_autofit(name) => 2,
        b"scene3d" => 3,
        b"sp3d" | b"flatTx" => 4,
        b"extLst" => 5,
        _ => 0,
    }
}

enum PercentKind {
    FontScale,
    LineSpacing,
}

fn validate_autofit_percent(attribute: &str, value: &str, kind: PercentKind) -> Result<()> {
    if is_percentage_string(value) {
        return Ok(());
    }
    let integer = value
        .parse::<i32>()
        .map_err(|_| invalid_attribute("normAutofit", attribute, value.to_owned()))?;
    let valid = match kind {
        PercentKind::FontScale => (1_000..=100_000).contains(&integer),
        PercentKind::LineSpacing => (0..=MAX_TEXT_SPACING_PERCENT).contains(&integer),
    };
    if valid {
        Ok(())
    } else {
        Err(invalid_attribute(
            "normAutofit",
            attribute,
            value.to_owned(),
        ))
    }
}

fn is_percentage_string(value: &str) -> bool {
    let Some(number) = value.strip_suffix('%') else {
        return false;
    };
    is_signed_decimal(number)
}

fn is_universal_measure(value: &str) -> bool {
    if value.len() < 3 {
        return false;
    }
    let (number, unit) = value.split_at(value.len() - 2);
    matches!(unit, "mm" | "cm" | "in" | "pt" | "pc" | "pi") && is_signed_decimal(number)
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

fn ensure_empty_element(reader: &mut Reader<&[u8]>, expected: &[u8]) -> Result<()> {
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
            Event::Eof => return Err(missing_end(&String::from_utf8_lossy(expected))),
            Event::Start(element) | Event::Empty(element) => {
                return Err(TextError::UnexpectedElement(element_name(&element)));
            }
            _ => {
                return Err(TextError::UnexpectedElement(
                    String::from_utf8_lossy(expected).into_owned(),
                ));
            }
        }
        buffer.clear();
    }
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

fn write_empty<W: Write>(writer: &mut Writer<W>, start: BytesStart<'_>) -> Result<()> {
    writer
        .write_event(Event::Empty(start))
        .map_err(OxmlError::from)?;
    Ok(())
}

fn invalid_attribute(element: &str, attribute: &str, value: String) -> TextError {
    TextError::InvalidAttribute {
        element: element.to_owned(),
        attribute: attribute.to_owned(),
        value,
    }
}

fn element_name(element: &BytesStart<'_>) -> String {
    String::from_utf8_lossy(element.name().as_ref()).into_owned()
}

/// Whether a value holds a character XML 1.0 cannot carry. The writer only
/// escapes markup, so such a character would reach the part as it is.
pub(crate) fn has_forbidden_xml_char(value: &str) -> bool {
    value.chars().any(|character| {
        matches!(
            character,
            '\u{0}'..='\u{8}' | '\u{b}' | '\u{c}' | '\u{e}'..='\u{1f}' | '\u{fffe}' | '\u{ffff}'
        )
    })
}

pub(crate) fn missing_end(element: &str) -> TextError {
    TextError::Xml(OxmlError::MissingElement(format!(
        "closing DrawingML {element}"
    )))
}

#[cfg(test)]
mod tests {
    use std::panic;

    use super::{CT_TextBodyProperties, Coordinate32Value, TextAutofit};
    use crate::text::CT_TextBody;

    #[test]
    fn every_body_property_autofit_form_round_trips_in_schema_order() {
        let cases: &[(&[u8], &[u8])] = &[
            (
                br#"<q:bodyPr lIns="-2147483648" tIns="2147483647" rIns="1.25in" bIns="0" anchor="dist" wrap="square" vert="wordArtVertRtl" spcFirstLastPara="true"><x:warp/><q:prstTxWarp prst="textPlain"/><x:beforeFit/><q:noAutofit/></q:bodyPr>"#,
                br#"<a:bodyPr lIns="-2147483648" tIns="2147483647" rIns="1.25in" bIns="0" anchor="dist" wrap="square" vert="wordArtVertRtl" spcFirstLastPara="1"><x:warp/><q:prstTxWarp prst="textPlain"/><x:beforeFit/><a:noAutofit/></a:bodyPr>"#,
            ),
            (
                br#"<q:bodyPr><q:spAutoFit/></q:bodyPr>"#,
                br#"<a:bodyPr><a:spAutoFit/></a:bodyPr>"#,
            ),
            (
                br#"<q:bodyPr><q:normAutofit fontScale="62.500%" lnSpcReduction="20000"/></q:bodyPr>"#,
                br#"<a:bodyPr><a:normAutofit fontScale="62.500%" lnSpcReduction="20000"/></a:bodyPr>"#,
            ),
        ];

        for (xml, expected) in cases {
            let parsed = CT_TextBodyProperties::from_xml(xml).unwrap();
            let written = parsed.to_xml().unwrap();
            assert_eq!(&written, expected);
            assert_eq!(CT_TextBodyProperties::from_xml(&written).unwrap(), parsed);
        }

        let parsed = CT_TextBodyProperties::from_xml(cases[0].0).unwrap();
        assert_eq!(parsed.left_inset, Some(Coordinate32Value::Emu(i32::MIN)));
        assert_eq!(parsed.space_first_last_paragraph, Some(true));
        assert!(matches!(parsed.autofit, Some(TextAutofit::NoAutofit)));
    }

    #[test]
    fn body_properties_preserve_unknown_children_at_their_boundaries() {
        let xml = br#"<q:bodyPr><x:before/><q:prstTxWarp prst="textPlain"><x:warp/></q:prstTxWarp><x:beforeFit/><q:normAutofit fontScale="62500"/><x:afterFit/><q:scene3d><x:scene/></q:scene3d><x:afterScene/><q:sp3d><x:shape/></q:sp3d><x:after3d/><q:extLst><x:ext/></q:extLst><x:afterExt/></q:bodyPr>"#;
        let written = CT_TextBodyProperties::from_xml(xml)
            .unwrap()
            .to_xml()
            .unwrap();
        assert_eq!(written, br#"<a:bodyPr><x:before/><q:prstTxWarp prst="textPlain"><x:warp/></q:prstTxWarp><x:beforeFit/><a:normAutofit fontScale="62500"/><x:afterFit/><q:scene3d><x:scene/></q:scene3d><x:afterScene/><q:sp3d><x:shape/></q:sp3d><x:after3d/><q:extLst><x:ext/></q:extLst><x:afterExt/></a:bodyPr>"#);
    }

    #[test]
    fn an_added_autofit_is_written_before_preserved_scene_and_extension_children() {
        let mut properties = CT_TextBodyProperties::from_xml(
            br#"<q:bodyPr><x:first/><q:scene3d><x:scene/></q:scene3d><q:flatTx/><q:extLst><x:ext/></q:extLst></q:bodyPr>"#,
        )
        .unwrap();
        properties.autofit = Some(TextAutofit::ShapeAutofit);
        let written = properties.to_xml().unwrap();
        assert_eq!(
            written,
            br#"<a:bodyPr><x:first/><a:spAutoFit/><q:scene3d><x:scene/></q:scene3d><q:flatTx/><q:extLst><x:ext/></q:extLst></a:bodyPr>"#
        );
        assert_eq!(
            CT_TextBodyProperties::from_xml(&written).unwrap(),
            properties
        );
    }

    #[test]
    fn malformed_body_properties_return_errors_without_panicking() {
        let body_cases: &[&[u8]] = &[
            br#"<q:notBodyPr/>"#,
            br#"<q:bodyPr anchor="middle"/>"#,
            br#"<q:bodyPr wrap="tight"/>"#,
            br#"<q:bodyPr vert="sideways"/>"#,
            br#"<q:bodyPr lIns="2147483648"/>"#,
            br#"<q:bodyPr rIns="1.in"/>"#,
            br#"<q:bodyPr spcFirstLastPara="yes"/>"#,
            br#"<q:bodyPr><q:normAutofit fontScale="999"/></q:bodyPr>"#,
            br#"<q:bodyPr><q:normAutofit fontScale="100001"/></q:bodyPr>"#,
            br#"<q:bodyPr><q:normAutofit fontScale=".5%"/></q:bodyPr>"#,
            br#"<q:bodyPr><q:normAutofit lnSpcReduction="13200001"/></q:bodyPr>"#,
            br#"<q:bodyPr><q:noAutofit><x:child/></q:noAutofit></q:bodyPr>"#,
            br#"<q:bodyPr><q:noAutofit/><q:spAutoFit/></q:bodyPr>"#,
            br#"<q:bodyPr><q:sp3d/><q:flatTx/></q:bodyPr>"#,
        ];
        for xml in body_cases {
            let result = panic::catch_unwind(|| CT_TextBodyProperties::from_xml(xml));
            assert!(result.is_ok(), "body-property parser panicked");
            assert!(result.unwrap().is_err(), "malformed body properties parsed");
        }

        let shell_cases: &[&[u8]] = &[
            br#"<q:txBody><q:p/></q:txBody>"#,
            br#"<q:txBody/>"#,
            br#"<q:txBody><q:bodyPr/><q:bodyPr/><q:p/></q:txBody>"#,
            br#"<q:txBody><q:bodyPr/><q:lstStyle/><q:lstStyle/><q:p/></q:txBody>"#,
        ];
        for xml in shell_cases {
            let result = panic::catch_unwind(|| CT_TextBody::from_xml(xml));
            assert!(result.is_ok(), "text-body parser panicked");
            assert!(result.unwrap().is_err(), "malformed text body parsed");
        }

        let empty = CT_TextBody::from_xml(br#"<q:txBody><q:bodyPr/></q:txBody>"#).unwrap();
        assert_eq!(empty.paragraph_count(), 0);
        assert!(empty.to_xml().is_ok());

        let mut properties = CT_TextBodyProperties {
            left_inset: Some(Coordinate32Value::UniversalMeasure("NaNin".to_owned())),
            ..CT_TextBodyProperties::default()
        };
        assert!(properties.to_xml().is_err());
        properties.left_inset = Some(Coordinate32Value::Emu(i32::MAX));
        assert!(properties.to_xml().is_ok());
    }
}
