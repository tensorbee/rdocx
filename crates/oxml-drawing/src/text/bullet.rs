use std::io::Write;

use oxml_core::OxmlError;
use oxml_core::raw_xml::{capture_element, capture_empty_element};
use oxml_core::xml::{local_name, matches_local_name};
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer, XmlVersion};

use crate::color::ColorChoice;
use crate::namespace::reject_conflicting_a_prefix;
use crate::order::OrderedRawChildren;

use super::body::{Result, TextError, has_forbidden_xml_char, missing_end};
use super::paragraph::TextFont;

const MIN_BULLET_PERCENT: i32 = 25_000;
const MAX_BULLET_PERCENT: i32 = 400_000;
const MIN_BULLET_POINTS: i32 = 100;
const MAX_BULLET_POINTS: i32 = 400_000;
const MAX_BULLET_START_AT: u16 = 32_767;

/// The modelled colour, size, font, and mutually exclusive choice of one bullet.
#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub struct TextBullet {
    pub color: Option<TextBulletColor>,
    pub size: Option<TextBulletSize>,
    pub font: Option<TextFont>,
    pub choice: Option<TextBulletChoice>,
}

impl TextBullet {
    pub(crate) fn capture_component(&mut self, name: &[u8], xml: &[u8]) -> Result<bool> {
        match name {
            b"buClr" => {
                if self.color.is_some() {
                    return Err(duplicate("bullet colour choice"));
                }
                self.color = Some(TextBulletColor::from_xml(xml)?);
            }
            b"buSzPct" | b"buSzPts" => {
                if self.size.is_some() {
                    return Err(duplicate("bullet size choice"));
                }
                self.size = Some(TextBulletSize::from_xml(xml)?);
            }
            b"buFont" => {
                if self.font.is_some() {
                    return Err(duplicate("buFont"));
                }
                self.font = Some(TextFont::from_xml(xml, b"buFont")?);
            }
            b"buNone" | b"buAutoNum" | b"buChar" => {
                if self.choice.is_some() {
                    return Err(duplicate("bullet choice"));
                }
                self.choice = Some(TextBulletChoice::from_xml(xml, name)?);
            }
            _ => return Ok(false),
        }
        Ok(true)
    }

    pub(crate) fn write_color<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        if let Some(color) = &self.color {
            color.write_xml(writer)?;
        }
        Ok(())
    }

    pub(crate) fn write_size<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        if let Some(size) = &self.size {
            size.write_xml(writer)?;
        }
        Ok(())
    }

    pub(crate) fn write_font<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        if let Some(font) = &self.font {
            font.write_xml(writer, "a:buFont")?;
        }
        Ok(())
    }

    pub(crate) fn write_choice<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        if let Some(choice) = &self.choice {
            choice.write_xml(writer)?;
        }
        Ok(())
    }

    pub(crate) fn is_empty(&self) -> bool {
        self.color.is_none() && self.size.is_none() && self.font.is_none() && self.choice.is_none()
    }
}

/// An `a:buClr` wrapper containing one DrawingML colour choice.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct TextBulletColor {
    pub color: ColorChoice,
    raw_attributes: Vec<(String, String)>,
    raw_children: OrderedRawChildren,
}

impl TextBulletColor {
    pub fn new(color: ColorChoice) -> Self {
        Self {
            color,
            raw_attributes: Vec::new(),
            raw_children: OrderedRawChildren::default(),
        }
    }

    fn from_xml(xml: &[u8]) -> Result<Self> {
        parse_complete(xml, b"buClr", Self::from_element, |_| {
            Err(TextError::UnexpectedElement("buClr".to_owned()))
        })
    }

    fn from_element(reader: &mut Reader<&[u8]>, start: &BytesStart<'_>) -> Result<Self> {
        let raw_attributes = capture_raw_attributes(start, &[])?;
        let mut color = None;
        let mut raw_children = OrderedRawChildren::default();
        let mut boundary = 0;
        let mut buffer = Vec::new();
        loop {
            match reader
                .read_event_into(&mut buffer)
                .map_err(OxmlError::from)?
            {
                Event::Start(element) if is_color(element.name().as_ref()) => {
                    if color.is_some() {
                        return Err(duplicate("bullet colour"));
                    }
                    validate_color_attributes(&element)?;
                    reject_conflicting_a_prefix(&element)?;
                    color = Some(ColorChoice::from_xml(reader, &element)?);
                    boundary = 1;
                }
                Event::Empty(element) if is_color(element.name().as_ref()) => {
                    if color.is_some() {
                        return Err(duplicate("bullet colour"));
                    }
                    validate_color_attributes(&element)?;
                    reject_conflicting_a_prefix(&element)?;
                    color = Some(ColorChoice::from_empty_xml(&element)?);
                    boundary = 1;
                }
                Event::Start(element) => {
                    raw_children.push(boundary, capture_element(reader, &element)?)
                }
                Event::Empty(element) => {
                    raw_children.push(boundary, capture_empty_element(&element)?)
                }
                Event::End(element) if matches_local_name(element.name().as_ref(), b"buClr") => {
                    return Ok(Self {
                        color: color.ok_or_else(|| missing("buClr", "colour choice"))?,
                        raw_attributes,
                        raw_children,
                    });
                }
                Event::Eof => return Err(missing_end("buClr")),
                _ => {}
            }
            buffer.clear();
        }
    }

    fn write_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut start = BytesStart::new("a:buClr");
        push_raw_attributes(&mut start, &self.raw_attributes);
        write_start(writer, start)?;
        emit_raw(writer, self.raw_children.at(0))?;
        self.color.to_xml(writer)?;
        emit_raw(writer, self.raw_children.at(1))?;
        write_end(writer, "a:buClr")
    }

    pub fn raw_children(&self) -> &OrderedRawChildren {
        &self.raw_children
    }
}

/// Percentage or centipoint size for a DrawingML bullet.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct TextBulletSize {
    pub value: TextBulletSizeValue,
    raw_attributes: Vec<(String, String)>,
}

impl TextBulletSize {
    pub fn percent(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        validate_bullet_percent(&value)?;
        Ok(Self {
            value: TextBulletSizeValue::Percent(value),
            raw_attributes: Vec::new(),
        })
    }

    pub fn points(value: i32) -> Result<Self> {
        validate_range(
            "buSzPts",
            "val",
            value,
            MIN_BULLET_POINTS,
            MAX_BULLET_POINTS,
        )?;
        Ok(Self {
            value: TextBulletSizeValue::Points(value),
            raw_attributes: Vec::new(),
        })
    }

    fn from_xml(xml: &[u8]) -> Result<Self> {
        let expected = root_local_name(xml)?;
        if !matches!(expected.as_slice(), b"buSzPct" | b"buSzPts") {
            return Err(TextError::UnexpectedElement(
                String::from_utf8_lossy(&expected).into_owned(),
            ));
        }
        parse_complete(
            xml,
            &expected,
            |reader, start| {
                let size = Self::from_start(start)?;
                ensure_empty(reader, &expected)?;
                Ok(size)
            },
            Self::from_start,
        )
    }

    fn from_start(start: &BytesStart<'_>) -> Result<Self> {
        let qualified_name = start.name();
        let name = local_name(qualified_name.as_ref());
        let value = required_attr(start, b"val")?;
        let value = match name {
            b"buSzPct" => {
                validate_bullet_percent(&value)?;
                TextBulletSizeValue::Percent(value)
            }
            b"buSzPts" => TextBulletSizeValue::Points(parse_range(
                "buSzPts",
                "val",
                &value,
                MIN_BULLET_POINTS,
                MAX_BULLET_POINTS,
            )?),
            _ => return Err(unexpected(start)),
        };
        Ok(Self {
            value,
            raw_attributes: capture_raw_attributes(start, &[b"val"])?,
        })
    }

    fn write_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let (tag, value) = match &self.value {
            TextBulletSizeValue::Percent(value) => {
                validate_bullet_percent(value)?;
                ("a:buSzPct", value.clone())
            }
            TextBulletSizeValue::Points(value) => {
                validate_range(
                    "buSzPts",
                    "val",
                    *value,
                    MIN_BULLET_POINTS,
                    MAX_BULLET_POINTS,
                )?;
                ("a:buSzPts", value.to_string())
            }
        };
        let mut start = BytesStart::new(tag);
        start.push_attribute(("val", value.as_str()));
        push_raw_attributes(&mut start, &self.raw_attributes);
        write_empty(writer, start)
    }
}

/// The two explicit DrawingML bullet-size forms.
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum TextBulletSizeValue {
    Percent(String),
    Points(i32),
}

/// One member of the DrawingML paragraph bullet choice.
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum TextBulletChoice {
    Character(TextBulletCharacter),
    AutoNumber(TextAutoNumber),
    None(TextNoBullet),
}

impl TextBulletChoice {
    fn from_xml(xml: &[u8], name: &[u8]) -> Result<Self> {
        match name {
            b"buChar" => Ok(Self::Character(TextBulletCharacter::from_xml(xml)?)),
            b"buAutoNum" => Ok(Self::AutoNumber(TextAutoNumber::from_xml(xml)?)),
            b"buNone" => Ok(Self::None(TextNoBullet::from_xml(xml)?)),
            _ => Err(TextError::UnexpectedElement(
                String::from_utf8_lossy(name).into_owned(),
            )),
        }
    }

    fn write_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        match self {
            Self::Character(value) => value.write_xml(writer),
            Self::AutoNumber(value) => value.write_xml(writer),
            Self::None(value) => value.write_xml(writer),
        }
    }
}

/// A literal DrawingML bullet character kept in its wire representation.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct TextBulletCharacter {
    pub character: String,
    raw_attributes: Vec<(String, String)>,
}

impl TextBulletCharacter {
    pub fn new(character: impl Into<String>) -> Result<Self> {
        let character = character.into();
        if character.is_empty() || has_forbidden_xml_char(&character) {
            return Err(invalid("buChar", "char", &character));
        }
        Ok(Self {
            character,
            raw_attributes: Vec::new(),
        })
    }

    fn from_xml(xml: &[u8]) -> Result<Self> {
        parse_complete(
            xml,
            b"buChar",
            |reader, start| {
                let value = Self::from_start(start)?;
                ensure_empty(reader, b"buChar")?;
                Ok(value)
            },
            Self::from_start,
        )
    }

    fn from_start(start: &BytesStart<'_>) -> Result<Self> {
        Ok(Self {
            character: required_attr(start, b"char")?,
            raw_attributes: capture_raw_attributes(start, &[b"char"])?,
        })
    }

    fn write_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        if self.character.is_empty() || has_forbidden_xml_char(&self.character) {
            return Err(invalid("buChar", "char", &self.character));
        }
        let mut start = BytesStart::new("a:buChar");
        start.push_attribute(("char", self.character.as_str()));
        push_raw_attributes(&mut start, &self.raw_attributes);
        write_empty(writer, start)
    }
}

/// An automatic-number bullet with a schema-defined numbering scheme.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct TextAutoNumber {
    pub scheme: TextAutoNumberScheme,
    pub start_at: Option<u16>,
    raw_attributes: Vec<(String, String)>,
}

impl TextAutoNumber {
    pub fn new(scheme: TextAutoNumberScheme) -> Self {
        Self {
            scheme,
            start_at: None,
            raw_attributes: Vec::new(),
        }
    }

    fn from_xml(xml: &[u8]) -> Result<Self> {
        parse_complete(
            xml,
            b"buAutoNum",
            |reader, start| {
                let value = Self::from_start(start)?;
                ensure_empty(reader, b"buAutoNum")?;
                Ok(value)
            },
            Self::from_start,
        )
    }

    fn from_start(start: &BytesStart<'_>) -> Result<Self> {
        let scheme_value = required_attr(start, b"type")?;
        let scheme = TextAutoNumberScheme::parse(&scheme_value)
            .ok_or_else(|| invalid("buAutoNum", "type", &scheme_value))?;
        let start_at = text_attr(start, b"startAt")?
            .map(|value| {
                parse_range(
                    "buAutoNum",
                    "startAt",
                    &value,
                    1,
                    i32::from(MAX_BULLET_START_AT),
                )
                .map(|value| value as u16)
            })
            .transpose()?;
        Ok(Self {
            scheme,
            start_at,
            raw_attributes: capture_raw_attributes(start, &[b"type", b"startAt"])?,
        })
    }

    fn write_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut start = BytesStart::new("a:buAutoNum");
        start.push_attribute(("type", self.scheme.as_str()));
        let start_at = if let Some(start_at) = self.start_at {
            if !(1..=MAX_BULLET_START_AT).contains(&start_at) {
                return Err(invalid("buAutoNum", "startAt", &start_at.to_string()));
            }
            Some(start_at.to_string())
        } else {
            None
        };
        if let Some(value) = start_at.as_deref() {
            start.push_attribute(("startAt", value));
        }
        push_raw_attributes(&mut start, &self.raw_attributes);
        write_empty(writer, start)
    }
}

/// The 41 values of `ST_TextAutonumberScheme`.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum TextAutoNumberScheme {
    AlphaLowerParenBoth,
    AlphaUpperParenBoth,
    AlphaLowerParenRight,
    AlphaUpperParenRight,
    AlphaLowerPeriod,
    AlphaUpperPeriod,
    ArabicParenBoth,
    ArabicParenRight,
    ArabicPeriod,
    ArabicPlain,
    RomanLowerParenBoth,
    RomanUpperParenBoth,
    RomanLowerParenRight,
    RomanUpperParenRight,
    RomanLowerPeriod,
    RomanUpperPeriod,
    CircleNumberDoubleBytePlain,
    CircleNumberWingdingsBlackPlain,
    CircleNumberWingdingsWhitePlain,
    ArabicDoubleBytePeriod,
    ArabicDoubleBytePlain,
    EastAsianSimplifiedChinesePeriod,
    EastAsianSimplifiedChinesePlain,
    EastAsianTraditionalChinesePeriod,
    EastAsianTraditionalChinesePlain,
    EastAsianJapaneseDoubleBytePeriod,
    EastAsianJapaneseKoreanPlain,
    EastAsianJapaneseKoreanPeriod,
    Arabic1Minus,
    Arabic2Minus,
    Hebrew2Minus,
    ThaiAlphaPeriod,
    ThaiAlphaParenRight,
    ThaiAlphaParenBoth,
    ThaiNumberPeriod,
    ThaiNumberParenRight,
    ThaiNumberParenBoth,
    HindiAlphaPeriod,
    HindiNumberPeriod,
    HindiNumberParenRight,
    HindiAlpha1Period,
}

impl TextAutoNumberScheme {
    fn parse(value: &str) -> Option<Self> {
        Some(match value {
            "alphaLcParenBoth" => Self::AlphaLowerParenBoth,
            "alphaUcParenBoth" => Self::AlphaUpperParenBoth,
            "alphaLcParenR" => Self::AlphaLowerParenRight,
            "alphaUcParenR" => Self::AlphaUpperParenRight,
            "alphaLcPeriod" => Self::AlphaLowerPeriod,
            "alphaUcPeriod" => Self::AlphaUpperPeriod,
            "arabicParenBoth" => Self::ArabicParenBoth,
            "arabicParenR" => Self::ArabicParenRight,
            "arabicPeriod" => Self::ArabicPeriod,
            "arabicPlain" => Self::ArabicPlain,
            "romanLcParenBoth" => Self::RomanLowerParenBoth,
            "romanUcParenBoth" => Self::RomanUpperParenBoth,
            "romanLcParenR" => Self::RomanLowerParenRight,
            "romanUcParenR" => Self::RomanUpperParenRight,
            "romanLcPeriod" => Self::RomanLowerPeriod,
            "romanUcPeriod" => Self::RomanUpperPeriod,
            "circleNumDbPlain" => Self::CircleNumberDoubleBytePlain,
            "circleNumWdBlackPlain" => Self::CircleNumberWingdingsBlackPlain,
            "circleNumWdWhitePlain" => Self::CircleNumberWingdingsWhitePlain,
            "arabicDbPeriod" => Self::ArabicDoubleBytePeriod,
            "arabicDbPlain" => Self::ArabicDoubleBytePlain,
            "ea1ChsPeriod" => Self::EastAsianSimplifiedChinesePeriod,
            "ea1ChsPlain" => Self::EastAsianSimplifiedChinesePlain,
            "ea1ChtPeriod" => Self::EastAsianTraditionalChinesePeriod,
            "ea1ChtPlain" => Self::EastAsianTraditionalChinesePlain,
            "ea1JpnChsDbPeriod" => Self::EastAsianJapaneseDoubleBytePeriod,
            "ea1JpnKorPlain" => Self::EastAsianJapaneseKoreanPlain,
            "ea1JpnKorPeriod" => Self::EastAsianJapaneseKoreanPeriod,
            "arabic1Minus" => Self::Arabic1Minus,
            "arabic2Minus" => Self::Arabic2Minus,
            "hebrew2Minus" => Self::Hebrew2Minus,
            "thaiAlphaPeriod" => Self::ThaiAlphaPeriod,
            "thaiAlphaParenR" => Self::ThaiAlphaParenRight,
            "thaiAlphaParenBoth" => Self::ThaiAlphaParenBoth,
            "thaiNumPeriod" => Self::ThaiNumberPeriod,
            "thaiNumParenR" => Self::ThaiNumberParenRight,
            "thaiNumParenBoth" => Self::ThaiNumberParenBoth,
            "hindiAlphaPeriod" => Self::HindiAlphaPeriod,
            "hindiNumPeriod" => Self::HindiNumberPeriod,
            "hindiNumParenR" => Self::HindiNumberParenRight,
            "hindiAlpha1Period" => Self::HindiAlpha1Period,
            _ => return None,
        })
    }

    pub const fn as_str(self) -> &'static str {
        match self {
            Self::AlphaLowerParenBoth => "alphaLcParenBoth",
            Self::AlphaUpperParenBoth => "alphaUcParenBoth",
            Self::AlphaLowerParenRight => "alphaLcParenR",
            Self::AlphaUpperParenRight => "alphaUcParenR",
            Self::AlphaLowerPeriod => "alphaLcPeriod",
            Self::AlphaUpperPeriod => "alphaUcPeriod",
            Self::ArabicParenBoth => "arabicParenBoth",
            Self::ArabicParenRight => "arabicParenR",
            Self::ArabicPeriod => "arabicPeriod",
            Self::ArabicPlain => "arabicPlain",
            Self::RomanLowerParenBoth => "romanLcParenBoth",
            Self::RomanUpperParenBoth => "romanUcParenBoth",
            Self::RomanLowerParenRight => "romanLcParenR",
            Self::RomanUpperParenRight => "romanUcParenR",
            Self::RomanLowerPeriod => "romanLcPeriod",
            Self::RomanUpperPeriod => "romanUcPeriod",
            Self::CircleNumberDoubleBytePlain => "circleNumDbPlain",
            Self::CircleNumberWingdingsBlackPlain => "circleNumWdBlackPlain",
            Self::CircleNumberWingdingsWhitePlain => "circleNumWdWhitePlain",
            Self::ArabicDoubleBytePeriod => "arabicDbPeriod",
            Self::ArabicDoubleBytePlain => "arabicDbPlain",
            Self::EastAsianSimplifiedChinesePeriod => "ea1ChsPeriod",
            Self::EastAsianSimplifiedChinesePlain => "ea1ChsPlain",
            Self::EastAsianTraditionalChinesePeriod => "ea1ChtPeriod",
            Self::EastAsianTraditionalChinesePlain => "ea1ChtPlain",
            Self::EastAsianJapaneseDoubleBytePeriod => "ea1JpnChsDbPeriod",
            Self::EastAsianJapaneseKoreanPlain => "ea1JpnKorPlain",
            Self::EastAsianJapaneseKoreanPeriod => "ea1JpnKorPeriod",
            Self::Arabic1Minus => "arabic1Minus",
            Self::Arabic2Minus => "arabic2Minus",
            Self::Hebrew2Minus => "hebrew2Minus",
            Self::ThaiAlphaPeriod => "thaiAlphaPeriod",
            Self::ThaiAlphaParenRight => "thaiAlphaParenR",
            Self::ThaiAlphaParenBoth => "thaiAlphaParenBoth",
            Self::ThaiNumberPeriod => "thaiNumPeriod",
            Self::ThaiNumberParenRight => "thaiNumParenR",
            Self::ThaiNumberParenBoth => "thaiNumParenBoth",
            Self::HindiAlphaPeriod => "hindiAlphaPeriod",
            Self::HindiNumberPeriod => "hindiNumPeriod",
            Self::HindiNumberParenRight => "hindiNumParenR",
            Self::HindiAlpha1Period => "hindiAlpha1Period",
        }
    }
}

/// An explicit `a:buNone` choice with unmodelled attributes retained.
#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub struct TextNoBullet {
    raw_attributes: Vec<(String, String)>,
}

impl TextNoBullet {
    fn from_xml(xml: &[u8]) -> Result<Self> {
        parse_complete(
            xml,
            b"buNone",
            |reader, start| {
                let value = Self::from_start(start)?;
                ensure_empty(reader, b"buNone")?;
                Ok(value)
            },
            Self::from_start,
        )
    }

    fn from_start(start: &BytesStart<'_>) -> Result<Self> {
        Ok(Self {
            raw_attributes: capture_raw_attributes(start, &[])?,
        })
    }

    fn write_xml<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut start = BytesStart::new("a:buNone");
        push_raw_attributes(&mut start, &self.raw_attributes);
        write_empty(writer, start)
    }
}

fn validate_bullet_percent(value: &str) -> Result<()> {
    if let Ok(value) = value.parse::<i32>() {
        return validate_range(
            "buSzPct",
            "val",
            value,
            MIN_BULLET_PERCENT,
            MAX_BULLET_PERCENT,
        );
    }
    let Some(percent) = value.strip_suffix('%') else {
        return Err(invalid("buSzPct", "val", value));
    };
    if !is_decimal(percent) {
        return Err(invalid("buSzPct", "val", value));
    }
    let percent = percent
        .parse::<f64>()
        .map_err(|_| invalid("buSzPct", "val", value))?;
    if (25.0..=400.0).contains(&percent) {
        Ok(())
    } else {
        Err(invalid("buSzPct", "val", value))
    }
}

fn is_decimal(value: &str) -> bool {
    let mut parts = value.split('.');
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

fn validate_range(element: &str, attribute: &str, value: i32, min: i32, max: i32) -> Result<()> {
    if (min..=max).contains(&value) {
        Ok(())
    } else {
        Err(invalid(element, attribute, &value.to_string()))
    }
}

fn parse_range(element: &str, attribute: &str, value: &str, min: i32, max: i32) -> Result<i32> {
    let parsed = value
        .parse::<i32>()
        .map_err(|_| invalid(element, attribute, value))?;
    validate_range(element, attribute, parsed, min, max)?;
    Ok(parsed)
}

fn is_color(name: &[u8]) -> bool {
    matches!(
        local_name(name),
        b"srgbClr" | b"schemeClr" | b"sysClr" | b"prstClr"
    )
}

fn validate_color_attributes(element: &BytesStart<'_>) -> Result<()> {
    required_attr(element, b"val")?;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(OxmlError::from)?;
        let name = attribute.key.as_ref();
        if name != b"val" && name != b"lastClr" && matches!(local_name(name), b"val" | b"lastClr") {
            return Err(invalid(
                &String::from_utf8_lossy(local_name(element.name().as_ref())),
                &String::from_utf8_lossy(name),
                &String::from_utf8_lossy(attribute.value.as_ref()),
            ));
        }
    }
    Ok(())
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
                    "DrawingML bullet".to_owned(),
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

fn required_attr(start: &BytesStart<'_>, name: &[u8]) -> Result<String> {
    text_attr(start, name)?
        .filter(|value| !value.is_empty())
        .ok_or_else(|| {
            missing(
                &String::from_utf8_lossy(local_name(start.name().as_ref())),
                &String::from_utf8_lossy(name),
            )
        })
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

fn missing(element: &str, attribute: &str) -> TextError {
    TextError::MissingAttribute {
        element: element.to_owned(),
        attribute: attribute.to_owned(),
    }
}

fn invalid(element: &str, attribute: &str, value: &str) -> TextError {
    TextError::InvalidAttribute {
        element: element.to_owned(),
        attribute: attribute.to_owned(),
        value: value.to_owned(),
    }
}

#[cfg(test)]
mod tests {
    use super::TextAutoNumberScheme;

    #[test]
    fn every_auto_number_scheme_token_maps_without_a_fallback() {
        let tokens = [
            "alphaLcParenBoth",
            "alphaUcParenBoth",
            "alphaLcParenR",
            "alphaUcParenR",
            "alphaLcPeriod",
            "alphaUcPeriod",
            "arabicParenBoth",
            "arabicParenR",
            "arabicPeriod",
            "arabicPlain",
            "romanLcParenBoth",
            "romanUcParenBoth",
            "romanLcParenR",
            "romanUcParenR",
            "romanLcPeriod",
            "romanUcPeriod",
            "circleNumDbPlain",
            "circleNumWdBlackPlain",
            "circleNumWdWhitePlain",
            "arabicDbPeriod",
            "arabicDbPlain",
            "ea1ChsPeriod",
            "ea1ChsPlain",
            "ea1ChtPeriod",
            "ea1ChtPlain",
            "ea1JpnChsDbPeriod",
            "ea1JpnKorPlain",
            "ea1JpnKorPeriod",
            "arabic1Minus",
            "arabic2Minus",
            "hebrew2Minus",
            "thaiAlphaPeriod",
            "thaiAlphaParenR",
            "thaiAlphaParenBoth",
            "thaiNumPeriod",
            "thaiNumParenR",
            "thaiNumParenBoth",
            "hindiAlphaPeriod",
            "hindiNumPeriod",
            "hindiNumParenR",
            "hindiAlpha1Period",
        ];

        for token in tokens {
            assert_eq!(TextAutoNumberScheme::parse(token).unwrap().as_str(), token);
        }
    }
}
