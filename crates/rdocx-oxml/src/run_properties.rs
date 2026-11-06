//! Run properties (`CT_RPr`).

use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer, XmlVersion};

use std::borrow::Cow;

use crate::borders::CT_BorderEdge;
use crate::error::{OxmlError, Result};
use crate::namespace::matches_local_name;
use crate::properties::{
    CT_Shd, get_word_val_attr, is_word_attribute, is_word_element, parse_integer_measurement,
    parse_uchar_hex, parse_word_toggle, raw_is_modeled_attribute_carrier, raw_occurrence,
    record_modeled_toggle_candidate, remove_redundant_modeled_toggle_candidate,
    replay_modeled_toggle_raw, toggle_element_is_explicitly_empty,
    toggle_has_unsupported_attributes, word_prefixes_at, write_toggle,
};
use crate::raw_xml::{capture_element, capture_empty_element};
use crate::revision::CT_Revision;
use crate::shared::{ST_HighlightColor, ST_Underline};
use crate::units::{HalfPoint, Twips};

/// The Empty-event bold action shared by construction and source qualification.
/// Started bold elements retain their separate raw-child handling.
pub(crate) fn empty_bold_value(
    element: &BytesStart<'_>,
    prefixes: &[String],
) -> Result<Option<bool>> {
    if is_word_element(element.name().as_ref(), b"b", prefixes) {
        Ok(Some(parse_word_toggle(element, prefixes)?))
    } else {
        Ok(None)
    }
}

const RPR_STYLE_SLOT: u8 = 0;
const RPR_FONTS_SLOT: u8 = 1;
const RPR_BOLD_SLOT: u8 = 2;
const RPR_BOLD_CS_SLOT: u8 = 3;
const RPR_ITALIC_SLOT: u8 = 4;
const RPR_ITALIC_CS_SLOT: u8 = 5;
const RPR_CAPS_SLOT: u8 = 6;
const RPR_SMALL_CAPS_SLOT: u8 = 7;
const RPR_STRIKE_SLOT: u8 = 8;
const RPR_DSTRIKE_SLOT: u8 = 9;
const RPR_OUTLINE_SLOT: u8 = 10;
const RPR_SHADOW_SLOT: u8 = 11;
const RPR_EMBOSS_SLOT: u8 = 12;
const RPR_IMPRINT_SLOT: u8 = 13;
const RPR_NO_PROOF_SLOT: u8 = 14;
const RPR_SNAP_TO_GRID_SLOT: u8 = 15;
const RPR_VANISH_SLOT: u8 = 16;
const RPR_WEB_HIDDEN_SLOT: u8 = 17;
const RPR_COLOR_SLOT: u8 = 18;
const RPR_SPACING_SLOT: u8 = 19;
const RPR_WIDTH_SLOT: u8 = 20;
const RPR_KERN_SLOT: u8 = 21;
const RPR_POSITION_SLOT: u8 = 22;
const RPR_SIZE_SLOT: u8 = 23;
const RPR_SIZE_CS_SLOT: u8 = 24;
const RPR_HIGHLIGHT_SLOT: u8 = 25;
const RPR_UNDERLINE_SLOT: u8 = 26;
const RPR_EFFECT_SLOT: u8 = 27;
const RPR_BORDER_SLOT: u8 = 28;
const RPR_SHADING_SLOT: u8 = 29;
const RPR_FIT_TEXT_SLOT: u8 = 30;
const RPR_VERT_ALIGN_SLOT: u8 = 31;
const RPR_RTL_SLOT: u8 = 32;
const RPR_COMPLEX_SCRIPT_SLOT: u8 = 33;
const RPR_EMPHASIS_MARK_SLOT: u8 = 34;
const RPR_LANG_SLOT: u8 = 35;
const RPR_EAST_ASIAN_LAYOUT_SLOT: u8 = 36;
const RPR_SPEC_VANISH_SLOT: u8 = 37;
const RPR_OFFICE_MATH_SLOT: u8 = 38;
const RPR_MARKER_SLOT: u8 = 39;
const RPR_CHANGE_SLOT: u8 = 40;
const RPR_END_SLOT: u8 = 41;

/// The schema slot and typed field for a `w:rPr` toggle that carries no
/// attribute beyond `w:val`.
///
/// `w:b`, `w:i` and the other long-standing toggles keep their own arms. This
/// covers the ten `EG_RPrBase` toggles F-265 modeled, so one table decides both
/// the slot and the field rather than ten near-identical arms.
fn rpr_toggle_slot<'a>(
    rpr: &'a mut CT_RPr,
    name: &[u8],
    prefixes: &[String],
) -> Option<(u8, &'a mut Option<bool>)> {
    if is_word_element(name, b"outline", prefixes) {
        Some((RPR_OUTLINE_SLOT, &mut rpr.outline))
    } else if is_word_element(name, b"shadow", prefixes) {
        Some((RPR_SHADOW_SLOT, &mut rpr.shadow))
    } else if is_word_element(name, b"emboss", prefixes) {
        Some((RPR_EMBOSS_SLOT, &mut rpr.emboss))
    } else if is_word_element(name, b"imprint", prefixes) {
        Some((RPR_IMPRINT_SLOT, &mut rpr.imprint))
    } else if is_word_element(name, b"noProof", prefixes) {
        Some((RPR_NO_PROOF_SLOT, &mut rpr.no_proof))
    } else if is_word_element(name, b"snapToGrid", prefixes) {
        Some((RPR_SNAP_TO_GRID_SLOT, &mut rpr.snap_to_grid))
    } else if is_word_element(name, b"webHidden", prefixes) {
        Some((RPR_WEB_HIDDEN_SLOT, &mut rpr.web_hidden))
    } else if is_word_element(name, b"cs", prefixes) {
        Some((RPR_COMPLEX_SCRIPT_SLOT, &mut rpr.complex_script))
    } else if is_word_element(name, b"specVanish", prefixes) {
        Some((RPR_SPEC_VANISH_SLOT, &mut rpr.spec_vanish))
    } else if is_word_element(name, b"oMath", prefixes) {
        Some((RPR_OFFICE_MATH_SLOT, &mut rpr.office_math))
    } else {
        None
    }
}

fn record_rpr_modeled(
    rpr: &mut CT_RPr,
    pending_raw: &mut Vec<Vec<u8>>,
    occurrences: &mut [usize],
    slot: u8,
) {
    let occurrence = if slot == RPR_MARKER_SLOT {
        occurrences[slot as usize]
    } else {
        0
    };
    flush_rpr_raw(rpr, pending_raw, slot, occurrence);
    occurrences[slot as usize] += 1;
}

fn record_rpr_raw_at(rpr: &mut CT_RPr, raw: Vec<u8>, slot: u8, occurrence: usize) {
    rpr.revision_xml.push(raw);
    rpr.revision_xml_positions.push((slot, occurrence));
}

fn flush_rpr_raw(rpr: &mut CT_RPr, pending_raw: &mut Vec<Vec<u8>>, slot: u8, occurrence: usize) {
    for raw in pending_raw.drain(..) {
        rpr.revision_xml.push(raw);
        rpr.revision_xml_positions.push((slot, occurrence));
    }
}

/// `ST_TextEffect` — the animated text effect on `w:effect/@w:val`.
///
/// Producer-defined tokens are retained through `Other`, matching
/// `ST_NumberFormat`.
#[derive(Debug, Clone, PartialEq, Eq)]
#[allow(non_camel_case_types)]
pub enum ST_TextEffect {
    None,
    BlinkBackground,
    Lights,
    AntsBlack,
    AntsRed,
    Shimmer,
    SparkleText,
    /// A token outside the ECMA-376 inventory, retained verbatim.
    Other(String),
}

impl ST_TextEffect {
    pub fn from_str(value: &str) -> Self {
        match value {
            "none" => ST_TextEffect::None,
            "blinkBackground" => ST_TextEffect::BlinkBackground,
            "lights" => ST_TextEffect::Lights,
            "antsBlack" => ST_TextEffect::AntsBlack,
            "antsRed" => ST_TextEffect::AntsRed,
            "shimmer" => ST_TextEffect::Shimmer,
            "sparkleText" => ST_TextEffect::SparkleText,
            other => ST_TextEffect::Other(other.to_owned()),
        }
    }

    pub fn as_str(&self) -> &str {
        match self {
            ST_TextEffect::None => "none",
            ST_TextEffect::BlinkBackground => "blinkBackground",
            ST_TextEffect::Lights => "lights",
            ST_TextEffect::AntsBlack => "antsBlack",
            ST_TextEffect::AntsRed => "antsRed",
            ST_TextEffect::Shimmer => "shimmer",
            ST_TextEffect::SparkleText => "sparkleText",
            ST_TextEffect::Other(value) => value.as_str(),
        }
    }
}

/// `ST_Em` — the emphasis mark on `w:em/@w:val`.
#[derive(Debug, Clone, PartialEq, Eq)]
#[allow(non_camel_case_types)]
pub enum ST_Em {
    None,
    Dot,
    Comma,
    Circle,
    UnderDot,
    /// A token outside the ECMA-376 inventory, retained verbatim.
    Other(String),
}

impl ST_Em {
    pub fn from_str(value: &str) -> Self {
        match value {
            "none" => ST_Em::None,
            "dot" => ST_Em::Dot,
            "comma" => ST_Em::Comma,
            "circle" => ST_Em::Circle,
            "underDot" => ST_Em::UnderDot,
            other => ST_Em::Other(other.to_owned()),
        }
    }

    pub fn as_str(&self) -> &str {
        match self {
            ST_Em::None => "none",
            ST_Em::Dot => "dot",
            ST_Em::Comma => "comma",
            ST_Em::Circle => "circle",
            ST_Em::UnderDot => "underDot",
            ST_Em::Other(value) => value.as_str(),
        }
    }
}

/// `CT_FitText` — `w:fitText`, the twip width a run segment is fitted into.
#[derive(Debug, Clone, PartialEq, Eq)]
#[allow(non_camel_case_types)]
pub struct CT_FitText {
    /// Target width in twips (`w:val`).
    pub val: Twips,
    /// Identifier grouping the runs that share one fitted segment (`w:id`).
    pub id: Option<i32>,
    /// Attributes this type does not model, in source order.
    pub extra_attributes: Vec<(String, String)>,
}

/// `CT_EastAsianLayout` — `w:eastAsianLayout`, two-lines-in-one and vertical
/// text within a run.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
#[allow(non_camel_case_types)]
pub struct CT_EastAsianLayout {
    /// Identifier grouping the runs that share one layout (`w:id`).
    pub id: Option<i32>,
    /// Whether the run is combined into one character cell (`w:combine`).
    pub combine: Option<bool>,
    /// The bracket pair drawn around a combined run (`w:combineBrackets`).
    pub combine_brackets: Option<String>,
    /// Whether the run is rotated into vertical text (`w:vert`).
    pub vert: Option<bool>,
    /// Whether rotated vertical text is compressed (`w:vertCompress`).
    pub vert_compress: Option<bool>,
    /// Attributes this type does not model, in source order.
    pub extra_attributes: Vec<(String, String)>,
}

/// Parse an `ST_OnOff` attribute value.
fn parse_on_off(value: &str) -> bool {
    crate::shared::ST_OnOff::from_str_or_default(Some(value)).is_on()
}

/// Numeric attributes whose compound property constructors consume every occurrence.
#[derive(Clone, Copy)]
pub(crate) enum RunCompoundMeasurement {
    FitWidth(i32),
    FitId(i32),
    EastAsianId(i32),
}

impl RunCompoundMeasurement {
    pub(crate) fn for_names(element: &[u8], attribute: &[u8]) -> Option<Self> {
        match (element, attribute) {
            (b"fitText", b"val") => Some(Self::FitWidth(0)),
            (b"fitText", b"id") => Some(Self::FitId(0)),
            (b"eastAsianLayout", b"id") => Some(Self::EastAsianId(0)),
            _ => None,
        }
    }

    pub(crate) fn decode(self, element: &[u8], attribute: &[u8], value: &str) -> Result<Self> {
        Ok(match self {
            Self::FitWidth(_) => {
                Self::FitWidth(parse_integer_measurement(element, attribute, value)?)
            }
            Self::FitId(_) => Self::FitId(value.parse()?),
            Self::EastAsianId(_) => Self::EastAsianId(value.parse()?),
        })
    }
}

impl CT_FitText {
    /// Read `w:fitText`, or `None` when its required `w:val` is absent.
    fn from_xml_attrs(e: &BytesStart<'_>, prefixes: &[String]) -> Result<Option<Self>> {
        let mut fit_text = CT_FitText {
            val: Twips(0),
            id: None,
            extra_attributes: Vec::new(),
        };
        let mut width_seen = false;
        for attribute in e.attributes() {
            let attribute = attribute?;
            let key = attribute.key.as_ref();
            let value = std::str::from_utf8(&attribute.value)?;
            if let Some(measurement) =
                RunCompoundMeasurement::for_names(b"fitText", attribute.key.local_name().as_ref())
                    .filter(|_| {
                        is_word_attribute(key, attribute.key.local_name().as_ref(), prefixes)
                    })
            {
                match measurement.decode(e.name().as_ref(), key, value)? {
                    RunCompoundMeasurement::FitWidth(value) => {
                        fit_text.val = Twips(value);
                        width_seen = true;
                    }
                    RunCompoundMeasurement::FitId(value) => fit_text.id = Some(value),
                    RunCompoundMeasurement::EastAsianId(_) => unreachable!("fitText selector"),
                }
            } else {
                fit_text
                    .extra_attributes
                    .push((std::str::from_utf8(key)?.to_owned(), value.to_owned()));
            }
        }
        Ok(width_seen.then_some(fit_text))
    }

    fn write_xml<W: std::io::Write>(
        &self,
        writer: &mut Writer<W>,
        role: crate::text::XmlWriteRole,
        source_context: Option<&[(&str, &str)]>,
        emission: Option<&crate::text::FieldXmlEmission>,
    ) -> Result<()> {
        let mut buf = itoa::Buffer::new();
        let mut e = crate::text::FieldXmlStart::new("w:fitText", emission)?;
        for (name, value) in &self.extra_attributes {
            e.push_raw(name, value)?;
        }
        let word_prefix = if role == crate::text::XmlWriteRole::Source {
            Cow::Owned(e.qualify(&[], role, source_context)?)
        } else {
            Cow::Borrowed("w")
        };
        e.push_word(&word_prefix, "val", buf.format(self.val.0))?;
        if let Some(id) = self.id {
            e.push_word(&word_prefix, "id", buf.format(id))?;
        }
        #[cfg(test)]
        crate::text::WORD_COMPOUND_EVENT_MATERIALIZATIONS
            .set(crate::text::WORD_COMPOUND_EVENT_MATERIALIZATIONS.get() + 1);
        e.write(writer, true)?;
        Ok(())
    }
}

impl CT_EastAsianLayout {
    fn from_xml_attrs(e: &BytesStart<'_>, prefixes: &[String]) -> Result<Self> {
        let mut layout = CT_EastAsianLayout::default();
        for attribute in e.attributes() {
            let attribute = attribute?;
            let key = attribute.key.as_ref();
            let value = std::str::from_utf8(&attribute.value)?;
            if let Some(measurement) = RunCompoundMeasurement::for_names(
                b"eastAsianLayout",
                attribute.key.local_name().as_ref(),
            )
            .filter(|_| is_word_attribute(key, attribute.key.local_name().as_ref(), prefixes))
            {
                let RunCompoundMeasurement::EastAsianId(value) =
                    measurement.decode(e.name().as_ref(), key, value)?
                else {
                    unreachable!("eastAsianLayout selector")
                };
                layout.id = Some(value);
            } else if is_word_attribute(key, b"combine", prefixes) {
                layout.combine = Some(parse_on_off(value));
            } else if is_word_attribute(key, b"combineBrackets", prefixes) {
                layout.combine_brackets = Some(value.to_owned());
            } else if is_word_attribute(key, b"vert", prefixes) {
                layout.vert = Some(parse_on_off(value));
            } else if is_word_attribute(key, b"vertCompress", prefixes) {
                layout.vert_compress = Some(parse_on_off(value));
            } else {
                layout
                    .extra_attributes
                    .push((std::str::from_utf8(key)?.to_owned(), value.to_owned()));
            }
        }
        Ok(layout)
    }

    fn write_xml<W: std::io::Write>(
        &self,
        writer: &mut Writer<W>,
        role: crate::text::XmlWriteRole,
        source_context: Option<&[(&str, &str)]>,
        emission: Option<&crate::text::FieldXmlEmission>,
    ) -> Result<()> {
        let mut buf = itoa::Buffer::new();
        let mut e = crate::text::FieldXmlStart::new("w:eastAsianLayout", emission)?;
        for (name, value) in &self.extra_attributes {
            e.push_raw(name, value)?;
        }
        let word_prefix = if role == crate::text::XmlWriteRole::Source {
            Cow::Owned(e.qualify(&[], role, source_context)?)
        } else {
            Cow::Borrowed("w")
        };
        if let Some(id) = self.id {
            e.push_word(&word_prefix, "id", buf.format(id))?;
        }
        if let Some(combine) = self.combine {
            e.push_word(&word_prefix, "combine", if combine { "1" } else { "0" })?;
        }
        if let Some(ref brackets) = self.combine_brackets {
            e.push_word(&word_prefix, "combineBrackets", brackets.as_str())?;
        }
        if let Some(vert) = self.vert {
            e.push_word(&word_prefix, "vert", if vert { "1" } else { "0" })?;
        }
        if let Some(vert_compress) = self.vert_compress {
            e.push_word(
                &word_prefix,
                "vertCompress",
                if vert_compress { "1" } else { "0" },
            )?;
        }
        #[cfg(test)]
        crate::text::WORD_COMPOUND_EVENT_MATERIALIZATIONS
            .set(crate::text::WORD_COMPOUND_EVENT_MATERIALIZATIONS.get() + 1);
        e.write(writer, true)?;
        Ok(())
    }
}

/// `CT_RPr` — Run properties.
#[derive(Debug, Clone, Default, PartialEq)]
#[allow(non_snake_case)]
pub struct CT_RPr {
    /// Character style ID (rStyle)
    pub style_id: Option<String>,
    /// Font name for ASCII range (rFonts/@w:ascii)
    pub font_ascii: Option<String>,
    /// Font name for high-ANSI range (rFonts/@w:hAnsi)
    pub font_hansi: Option<String>,
    /// Font name for East Asian text (rFonts/@w:eastAsia)
    pub font_east_asia: Option<String>,
    /// Font name for complex script (rFonts/@w:cs)
    pub font_cs: Option<String>,
    /// Theme font for ASCII range (rFonts/@w:asciiTheme), e.g. "minorHAnsi", "majorHAnsi"
    pub font_ascii_theme: Option<String>,
    /// Theme font for hAnsi range (rFonts/@w:hAnsiTheme)
    pub font_hansi_theme: Option<String>,
    /// Theme font for East Asian text (rFonts/@w:eastAsiaTheme)
    pub font_east_asia_theme: Option<String>,
    /// Theme font for complex script (rFonts/@w:cstheme)
    pub font_cs_theme: Option<String>,
    /// Font slot hint for ambiguous characters (rFonts/@w:hint)
    pub font_hint: Option<String>,
    /// Namespace declarations and foreign attributes retained from `w:rFonts`.
    #[doc(hidden)]
    pub font_extra_attributes: Vec<(String, String)>,
    /// Bold (b)
    pub bold: Option<bool>,
    /// Bold complex script (bCs)
    pub bold_cs: Option<bool>,
    /// Italic (i)
    pub italic: Option<bool>,
    /// Italic complex script (iCs)
    pub italic_cs: Option<bool>,
    /// Underline type (u)
    pub underline: Option<ST_Underline>,
    /// Strikethrough (strike)
    pub strike: Option<bool>,
    /// Double strikethrough (dstrike)
    pub dstrike: Option<bool>,
    /// Font size in half-points (sz)
    pub sz: Option<HalfPoint>,
    /// Complex-script font size in half-points (szCs)
    pub sz_cs: Option<HalfPoint>,
    /// Text color as hex string, e.g. "FF0000" (color/@w:val)
    pub color: Option<String>,
    /// Color theme reference (color/@w:themeColor)
    pub color_theme: Option<String>,
    /// Theme colour tint, 0 to 255 (color/@w:themeTint)
    pub color_theme_tint: Option<u8>,
    /// Theme colour shade, 0 to 255 (color/@w:themeShade)
    pub color_theme_shade: Option<u8>,
    /// Namespace declarations and foreign attributes retained from `w:color`.
    #[doc(hidden)]
    pub color_extra_attributes: Vec<(String, String)>,
    /// Highlight color (highlight)
    pub highlight: Option<ST_HighlightColor>,
    /// All caps (caps)
    pub caps: Option<bool>,
    /// Small caps (smallCaps)
    pub small_caps: Option<bool>,
    /// Superscript/subscript (vertAlign)
    pub vert_align: Option<String>,
    /// Character spacing in twips (spacing/@w:val)
    pub spacing: Option<Twips>,
    /// Character width scale in percent (w/@w:val)
    pub width_scale: Option<u32>,
    /// Text position (raised/lowered) in half-points (position/@w:val)
    pub position: Option<i32>,
    /// Run shading (shd)
    pub shading: Option<Box<CT_Shd>>,
    /// Vanish/hidden text (vanish)
    pub vanish: Option<bool>,
    /// Outline, stroke-only glyphs (outline), schema slot 10.
    pub outline: Option<bool>,
    /// Drop shadow behind the glyphs (shadow), schema slot 11.
    pub shadow: Option<bool>,
    /// Raised relief (emboss), schema slot 12.
    pub emboss: Option<bool>,
    /// Sunken relief (imprint), schema slot 13.
    pub imprint: Option<bool>,
    /// Exclude from proofing (noProof), schema slot 14.
    pub no_proof: Option<bool>,
    /// Snap to the document character grid (snapToGrid), schema slot 15.
    pub snap_to_grid: Option<bool>,
    /// Hidden in web view only (webHidden), schema slot 17.
    pub web_hidden: Option<bool>,
    /// Kerning threshold in half-points (kern), schema slot 21.
    pub kern: Option<HalfPoint>,
    /// Animated text effect (effect), schema slot 27.
    pub effect: Option<ST_TextEffect>,
    /// Character border (bdr), schema slot 28.
    ///
    /// Boxed so a run property struct that models the whole `EG_RPrBase`
    /// sequence still fits the test-thread stack budget, following the
    /// `CT_PPr::borders` precedent.
    pub border: Option<Box<CT_BorderEdge>>,
    /// Fitted segment width (fitText), schema slot 30.
    pub fit_text: Option<Box<CT_FitText>>,
    /// Complex-script formatting toggle (cs), schema slot 33.
    pub complex_script: Option<bool>,
    /// East Asian emphasis mark (em), schema slot 34.
    pub emphasis_mark: Option<ST_Em>,
    /// East Asian run layout (eastAsianLayout), schema slot 36.
    pub east_asian_layout: Option<Box<CT_EastAsianLayout>>,
    /// Vanish only at the end of a numbered paragraph (specVanish), slot 37.
    pub spec_vanish: Option<bool>,
    /// Office Math run (oMath), schema slot 38.
    pub office_math: Option<bool>,
    /// Character-level right-to-left direction (rtl).
    pub rtl: Option<bool>,
    /// Language for Latin and high-ANSI text (`lang/@w:val`).
    pub language: Option<String>,
    /// Language for East Asian text (`lang/@w:eastAsia`).
    pub language_east_asia: Option<String>,
    /// Language for complex-script text (`lang/@w:bidi`).
    pub language_bidi: Option<String>,
    /// Namespace declarations and foreign attributes retained from `w:lang`.
    #[doc(hidden)]
    pub language_extra_attributes: Vec<(String, String)>,
    /// Contextual insertion and deletion markers retained in schema order.
    pub revision_markers: Vec<CT_Revision>,
    /// Prior run properties from the schema-final `w:rPrChange`.
    ///
    /// Boxed for the same stack-budget reason as `border`.
    pub change: Option<Box<CT_Revision>>,
    /// Foreign and unmodelled children retained without a typed projection.
    pub revision_xml: Vec<Vec<u8>>,
    /// Schema slots and occurrences for retained raw children.
    #[doc(hidden)]
    pub revision_xml_positions: Vec<(u8, usize)>,
}

/// Empty measurement leaves share exact scalar decoding with source qualification.
/// A missing kerning value remains raw. Other missing values retain prior state.
#[derive(Clone, Copy)]
pub(crate) enum RunMeasurement {
    Size(Option<u32>),
    SizeCs(Option<u32>),
    Spacing(Option<i32>),
    Width(Option<u32>),
    Position(Option<i32>),
    Kern(Option<u32>),
}

impl RunMeasurement {
    pub(crate) fn for_local(local: &[u8]) -> Option<Self> {
        match local {
            b"sz" => Some(Self::Size(None)),
            b"szCs" => Some(Self::SizeCs(None)),
            b"spacing" => Some(Self::Spacing(None)),
            b"w" => Some(Self::Width(None)),
            b"position" => Some(Self::Position(None)),
            b"kern" => Some(Self::Kern(None)),
            _ => None,
        }
    }

    pub(crate) fn decode(self, element: &[u8], value: Option<&str>) -> Result<Self> {
        let signed = || {
            value
                .map(|value| parse_integer_measurement(element, b"w:val", value))
                .transpose()
        };
        let unsigned = || {
            value
                .map(|value| parse_integer_measurement(element, b"w:val", value))
                .transpose()
        };
        Ok(match self {
            Self::Size(_) => Self::Size(unsigned()?),
            Self::SizeCs(_) => Self::SizeCs(unsigned()?),
            Self::Spacing(_) => Self::Spacing(signed()?),
            Self::Width(_) => Self::Width(value.map(str::parse).transpose()?),
            Self::Position(_) => Self::Position(signed()?),
            Self::Kern(_) => Self::Kern(unsigned()?),
        })
    }

    fn slot(self) -> Option<u8> {
        match self {
            Self::Size(_) => Some(RPR_SIZE_SLOT),
            Self::SizeCs(_) => Some(RPR_SIZE_CS_SLOT),
            Self::Spacing(_) => Some(RPR_SPACING_SLOT),
            Self::Width(_) => Some(RPR_WIDTH_SLOT),
            Self::Position(_) => Some(RPR_POSITION_SLOT),
            Self::Kern(Some(_)) => Some(RPR_KERN_SLOT),
            Self::Kern(None) => None,
        }
    }

    fn apply(self, rpr: &mut CT_RPr) {
        match self {
            Self::Size(Some(value)) => rpr.sz = Some(HalfPoint(value)),
            Self::SizeCs(Some(value)) => rpr.sz_cs = Some(HalfPoint(value)),
            Self::Spacing(Some(value)) => rpr.spacing = Some(Twips(value)),
            Self::Width(Some(value)) => rpr.width_scale = Some(value),
            Self::Position(Some(value)) => rpr.position = Some(value),
            Self::Kern(Some(value)) => rpr.kern = Some(HalfPoint(value)),
            _ => {}
        }
    }
}

impl CT_RPr {
    /// Retained formatting records positively selected by the containing field
    /// writer. This measures stored input, never serialized output bytes.
    pub(crate) fn selected_input_work(&self) -> Result<usize> {
        let mut bytes = 0usize;
        let mut add = |length: usize| -> Result<()> {
            bytes = bytes.checked_add(length).ok_or_else(|| {
                OxmlError::InvalidValue("normalization affected-source work overflow".into())
            })?;
            Ok(())
        };
        for raw in &self.revision_xml {
            add(raw.len())?;
        }
        for revision in &self.revision_markers {
            add(revision.retained_source_len())?;
        }
        if let Some(change) = &self.change {
            add(change.retained_source_len())?;
        }
        if let Some(raw) = self
            .border
            .as_ref()
            .and_then(|border| border.source_xml.as_ref())
        {
            add(raw.len())?;
        }
        // These are retained name/value records, not raw lexical XML and not
        // semantic current namespace facts. In particular, an unused xmlns
        // record cannot receive the affine current-value multiplier.
        for attributes in [
            Some(&self.font_extra_attributes),
            Some(&self.color_extra_attributes),
            Some(&self.language_extra_attributes),
            self.shading.as_ref().map(|value| &value.extra_attributes),
            self.border.as_ref().map(|value| &value.extra_attributes),
            self.fit_text.as_ref().map(|value| &value.extra_attributes),
            self.east_asian_layout
                .as_ref()
                .map(|value| &value.extra_attributes),
        ]
        .into_iter()
        .flatten()
        {
            for (name, value) in attributes {
                add(name.len())?;
                add(value.len())?;
            }
        }
        let current = self.selected_current_bytes()?;
        crate::text::retained_owner_work(bytes, &[], current)
    }

    fn selected_current_bytes(&self) -> Result<usize> {
        let mut total = 0usize;
        let mut add = |bytes: usize| -> Result<()> {
            total = total.checked_add(bytes).ok_or_else(|| {
                OxmlError::InvalidValue("normalization affected-source work overflow".into())
            })?;
            Ok(())
        };
        for value in [
            &self.style_id,
            &self.font_ascii,
            &self.font_hansi,
            &self.font_east_asia,
            &self.font_cs,
            &self.font_ascii_theme,
            &self.font_hansi_theme,
            &self.font_east_asia_theme,
            &self.font_cs_theme,
            &self.font_hint,
            &self.color,
            &self.color_theme,
            &self.vert_align,
            &self.language,
            &self.language_east_asia,
            &self.language_bidi,
        ]
        .into_iter()
        .flatten()
        {
            add(value.len())?;
        }
        // Toggle writers use one canonical 0/1 byte, including val-less true.
        for value in [
            self.bold,
            self.bold_cs,
            self.italic,
            self.italic_cs,
            self.strike,
            self.dstrike,
            self.caps,
            self.small_caps,
            self.vanish,
            self.outline,
            self.shadow,
            self.emboss,
            self.imprint,
            self.no_proof,
            self.snap_to_grid,
            self.web_hidden,
            self.complex_script,
            self.spec_vanish,
            self.office_math,
            self.rtl,
        ]
        .into_iter()
        .flatten()
        {
            let _ = value;
            add(1)?;
        }
        let mut number = itoa::Buffer::new();
        for value in [
            self.sz.map(|value| i64::from(value.0)),
            self.sz_cs.map(|value| i64::from(value.0)),
            self.spacing.map(|value| i64::from(value.0)),
            self.width_scale.map(i64::from),
            self.position.map(i64::from),
            self.kern.map(|value| i64::from(value.0)),
        ]
        .into_iter()
        .flatten()
        {
            add(number.format(value).len())?;
        }
        // Tint and shade writers use exactly two hexadecimal digits.
        for _ in [self.color_theme_tint, self.color_theme_shade]
            .into_iter()
            .flatten()
        {
            add(2)?;
        }
        for value in [
            self.underline.map(|value| value.to_str()),
            self.highlight.map(|value| value.to_str()),
            self.effect.as_ref().map(|value| value.as_str()),
            self.emphasis_mark.as_ref().map(|value| value.as_str()),
        ]
        .into_iter()
        .flatten()
        {
            add(value.len())?;
        }
        if let Some(shading) = &self.shading {
            add(shading.val.len())?;
            for value in [
                &shading.color,
                &shading.fill,
                &shading.theme_color,
                &shading.theme_fill,
            ]
            .into_iter()
            .flatten()
            {
                add(value.len())?;
            }
            for _ in [
                shading.theme_tint,
                shading.theme_shade,
                shading.theme_fill_tint,
                shading.theme_fill_shade,
            ]
            .into_iter()
            .flatten()
            {
                add(2)?;
            }
        }
        if let Some(border) = &self.border {
            add(border.val.to_str().len())?;
            for value in [border.sz, border.space].into_iter().flatten() {
                add(number.format(value).len())?;
            }
            if let Some(color) = &border.color {
                add(color.len())?;
            }
        }
        if let Some(fit) = &self.fit_text {
            add(number.format(fit.val.0).len())?;
            if let Some(id) = fit.id {
                add(number.format(id).len())?;
            }
        }
        if let Some(layout) = &self.east_asian_layout {
            if let Some(id) = layout.id {
                add(number.format(id).len())?;
            }
            for _ in [layout.combine, layout.vert, layout.vert_compress]
                .into_iter()
                .flatten()
            {
                add(1)?;
            }
            if let Some(brackets) = &layout.combine_brackets {
                add(brackets.len())?;
            }
        }
        Ok(total)
    }

    /// Resolve qualified solid text fill retained as opaque source.
    /// Other extension representations retain their source without being guessed.
    pub fn solid_text_fill(&self) -> Option<(String, f64)> {
        const NS: &[u8] = b"http://schemas.microsoft.com/office/word/2010/wordml";
        for raw in &self.revision_xml {
            let mut probe = quick_xml::NsReader::from_reader(raw.as_slice());
            let Ok((namespace, event)) = probe.read_resolved_event() else {
                continue;
            };
            if !matches!(namespace, quick_xml::name::ResolveResult::Bound(ns) if ns.as_ref() == NS)
                || !matches!(event, Event::Start(ref element) if element.local_name().as_ref() == b"textFill")
            {
                continue;
            }
            let mut reader = quick_xml::NsReader::from_reader(raw.as_slice());
            let mut path = Vec::new();
            let mut color = None;
            let mut alpha = 1.0;
            let mut valid = true;
            loop {
                let (namespace, event) = reader.read_resolved_event().ok()?;
                match event {
                    Event::Start(ref element) | Event::Empty(ref element) => {
                        let name = element.local_name().as_ref().to_vec();
                        let qualified = matches!(namespace, quick_xml::name::ResolveResult::Bound(ns) if ns.as_ref() == NS);
                        if !qualified {
                            valid = false;
                        }
                        path.push(name);
                        let expected: &[&[u8]] = &[b"textFill", b"solidFill", b"srgbClr", b"alpha"];
                        if path.len() > expected.len()
                            || !path
                                .iter()
                                .zip(expected)
                                .all(|(actual, expected)| actual.as_slice() == *expected)
                        {
                            valid = false;
                        }
                        if path.len() == 3 || path.len() == 4 {
                            let mut value = None;
                            for attribute in element.attributes() {
                                let attribute = attribute.ok()?;
                                let (ns, local) =
                                    reader.resolver().resolve_attribute(attribute.key);
                                if matches!(ns, quick_xml::name::ResolveResult::Bound(ns) if ns.as_ref() == NS)
                                    && local.as_ref() == b"val"
                                {
                                    if value.is_some() {
                                        valid = false;
                                    }
                                    value = Some(
                                        attribute
                                            .decoded_and_normalized_value(
                                                quick_xml::XmlVersion::Implicit1_0,
                                                reader.decoder(),
                                            )
                                            .ok()?
                                            .into_owned(),
                                    );
                                }
                            }
                            if path.len() == 3 {
                                let hex = value?;
                                if hex.len() != 6
                                    || !hex.bytes().all(|byte| byte.is_ascii_hexdigit())
                                    || color.is_some()
                                {
                                    valid = false;
                                }
                                color = Some(hex);
                            } else {
                                let amount: u32 = value?.parse().ok()?;
                                if amount > 100_000 {
                                    valid = false;
                                }
                                alpha = f64::from(amount) / 100_000.0;
                            }
                        }
                        if matches!(event, Event::Empty(_)) {
                            path.pop();
                        }
                    }
                    Event::End(_) => {
                        path.pop();
                    }
                    Event::Eof => break,
                    Event::Text(ref text) if text.as_ref().iter().all(u8::is_ascii_whitespace) => {}
                    _ => {
                        valid = false;
                    }
                }
            }
            if valid && let Some(color) = color {
                return Some((color, alpha));
            }
        }
        None
    }

    pub fn from_xml(reader: &mut Reader<&[u8]>) -> Result<Self> {
        Self::from_xml_with_prefixes(reader, &["w".to_owned()])
    }

    pub(crate) fn from_xml_with_prefixes(
        reader: &mut Reader<&[u8]>,
        word_prefixes: &[String],
    ) -> Result<Self> {
        Self::from_xml_with_prefixes_and_owner_bindings(reader, word_prefixes, &[])
    }

    pub(crate) fn from_xml_with_prefixes_and_owner_bindings(
        reader: &mut Reader<&[u8]>,
        word_prefixes: &[String],
        owner_bindings: &[(String, String)],
    ) -> Result<Self> {
        Self::from_xml_with_property_observation(reader, word_prefixes, owner_bindings, None)
    }

    pub(crate) fn from_xml_with_property_observation(
        reader: &mut Reader<&[u8]>,
        word_prefixes: &[String],
        owner_bindings: &[(String, String)],
        mut ordinary: Option<&mut Option<std::ops::Range<usize>>>,
    ) -> Result<Self> {
        if let Some(output) = &mut ordinary {
            **output = None;
        }
        let mut selected_style = None;
        let mut rpr = CT_RPr::default();
        let mut change_raw_index = 0usize;
        let mut pending_raw = Vec::new();
        let mut occurrences = [0usize; RPR_END_SLOT as usize + 1];
        let mut rtl_carrier_required = false;
        let mut buf = Vec::new();

        loop {
            let event_start = ordinary.as_ref().map(|_| reader.buffer_position() as usize);
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(ref e)) => {
                    let name = e.name();
                    let prefixes = word_prefixes_at(e, word_prefixes)?;
                    if let Some(measurement) = RunMeasurement::for_local(e.local_name().as_ref())
                        .filter(|_| {
                            is_word_element(name.as_ref(), e.local_name().as_ref(), &prefixes)
                        })
                    {
                        let value = get_word_val_attr(e, &prefixes)?;
                        let measurement = measurement.decode(name.as_ref(), value.as_deref())?;
                        if let Some(slot) = measurement.slot() {
                            record_rpr_modeled(&mut rpr, &mut pending_raw, &mut occurrences, slot);
                            measurement.apply(&mut rpr);
                        } else {
                            let raw = crate::text::raw_with_external_bindings(
                                &capture_empty_element(e)?,
                                owner_bindings,
                            )?;
                            record_rpr_raw_at(&mut rpr, raw, RPR_KERN_SLOT, 0);
                        }
                    } else if is_word_element(name.as_ref(), b"rStyle", &prefixes) {
                        record_rpr_modeled(
                            &mut rpr,
                            &mut pending_raw,
                            &mut occurrences,
                            RPR_STYLE_SLOT,
                        );
                        rpr.style_id = get_word_val_attr(e, &prefixes)?;
                        // Only the last selected style assignment survives.
                        // Decoder failure exits before any observation escapes.
                        selected_style =
                            event_start.map(|start| start..reader.buffer_position() as usize);
                    } else if is_word_element(name.as_ref(), b"rFonts", &prefixes) {
                        record_rpr_modeled(
                            &mut rpr,
                            &mut pending_raw,
                            &mut occurrences,
                            RPR_FONTS_SLOT,
                        );
                        for attr in e.attributes() {
                            let attr = attr?;
                            let key = attr.key.as_ref();
                            let val = std::str::from_utf8(&attr.value)?.to_string();
                            if is_word_attribute(key, b"ascii", &prefixes) {
                                rpr.font_ascii = Some(val);
                            } else if is_word_attribute(key, b"hAnsi", &prefixes) {
                                rpr.font_hansi = Some(val);
                            } else if is_word_attribute(key, b"eastAsia", &prefixes) {
                                rpr.font_east_asia = Some(val);
                            } else if is_word_attribute(key, b"cs", &prefixes) {
                                rpr.font_cs = Some(val);
                            } else if is_word_attribute(key, b"asciiTheme", &prefixes) {
                                rpr.font_ascii_theme = Some(val);
                            } else if is_word_attribute(key, b"hAnsiTheme", &prefixes) {
                                rpr.font_hansi_theme = Some(val);
                            } else if is_word_attribute(key, b"eastAsiaTheme", &prefixes) {
                                rpr.font_east_asia_theme = Some(val);
                            } else if is_word_attribute(key, b"cstheme", &prefixes) {
                                rpr.font_cs_theme = Some(val);
                            } else if is_word_attribute(key, b"hint", &prefixes) {
                                rpr.font_hint = Some(val);
                            } else {
                                rpr.font_extra_attributes
                                    .push((std::str::from_utf8(key)?.to_owned(), val));
                            }
                        }
                    } else if let Some(bold) = empty_bold_value(e, &prefixes)? {
                        record_rpr_modeled(
                            &mut rpr,
                            &mut pending_raw,
                            &mut occurrences,
                            RPR_BOLD_SLOT,
                        );
                        rpr.bold = Some(bold);
                    } else if is_word_element(name.as_ref(), b"bCs", &prefixes) {
                        record_rpr_modeled(
                            &mut rpr,
                            &mut pending_raw,
                            &mut occurrences,
                            RPR_BOLD_CS_SLOT,
                        );
                        rpr.bold_cs = Some(parse_word_toggle(e, &prefixes)?);
                    } else if is_word_element(name.as_ref(), b"i", &prefixes) {
                        record_rpr_modeled(
                            &mut rpr,
                            &mut pending_raw,
                            &mut occurrences,
                            RPR_ITALIC_SLOT,
                        );
                        rpr.italic = Some(parse_word_toggle(e, &prefixes)?);
                    } else if is_word_element(name.as_ref(), b"iCs", &prefixes) {
                        record_rpr_modeled(
                            &mut rpr,
                            &mut pending_raw,
                            &mut occurrences,
                            RPR_ITALIC_CS_SLOT,
                        );
                        rpr.italic_cs = Some(parse_word_toggle(e, &prefixes)?);
                    } else if is_word_element(name.as_ref(), b"u", &prefixes) {
                        record_rpr_modeled(
                            &mut rpr,
                            &mut pending_raw,
                            &mut occurrences,
                            RPR_UNDERLINE_SLOT,
                        );
                        if let Some(val) = get_word_val_attr(e, &prefixes)? {
                            rpr.underline = ST_Underline::from_str(&val).ok();
                        } else {
                            rpr.underline = Some(ST_Underline::Single);
                        }
                    } else if is_word_element(name.as_ref(), b"strike", &prefixes) {
                        record_rpr_modeled(
                            &mut rpr,
                            &mut pending_raw,
                            &mut occurrences,
                            RPR_STRIKE_SLOT,
                        );
                        rpr.strike = Some(parse_word_toggle(e, &prefixes)?);
                    } else if is_word_element(name.as_ref(), b"dstrike", &prefixes) {
                        record_rpr_modeled(
                            &mut rpr,
                            &mut pending_raw,
                            &mut occurrences,
                            RPR_DSTRIKE_SLOT,
                        );
                        rpr.dstrike = Some(parse_word_toggle(e, &prefixes)?);
                    } else if is_word_element(name.as_ref(), b"color", &prefixes) {
                        record_rpr_modeled(
                            &mut rpr,
                            &mut pending_raw,
                            &mut occurrences,
                            RPR_COLOR_SLOT,
                        );
                        for attr in e.attributes() {
                            let attr = attr?;
                            let key = attr.key.as_ref();
                            let v = std::str::from_utf8(&attr.value)?.to_string();
                            if is_word_attribute(key, b"val", &prefixes) {
                                rpr.color = Some(v);
                            } else if is_word_attribute(key, b"themeColor", &prefixes) {
                                rpr.color_theme = Some(v);
                            } else if is_word_attribute(key, b"themeTint", &prefixes) {
                                match parse_uchar_hex(&v) {
                                    Some(tint) => rpr.color_theme_tint = Some(tint),
                                    None => rpr
                                        .color_extra_attributes
                                        .push((std::str::from_utf8(key)?.to_owned(), v)),
                                }
                            } else if is_word_attribute(key, b"themeShade", &prefixes) {
                                match parse_uchar_hex(&v) {
                                    Some(shade) => rpr.color_theme_shade = Some(shade),
                                    None => rpr
                                        .color_extra_attributes
                                        .push((std::str::from_utf8(key)?.to_owned(), v)),
                                }
                            } else {
                                rpr.color_extra_attributes
                                    .push((std::str::from_utf8(key)?.to_owned(), v));
                            }
                        }
                    } else if is_word_element(name.as_ref(), b"highlight", &prefixes) {
                        record_rpr_modeled(
                            &mut rpr,
                            &mut pending_raw,
                            &mut occurrences,
                            RPR_HIGHLIGHT_SLOT,
                        );
                        if let Some(val) = get_word_val_attr(e, &prefixes)? {
                            rpr.highlight = ST_HighlightColor::from_str(&val).ok();
                        }
                    } else if is_word_element(name.as_ref(), b"caps", &prefixes) {
                        record_rpr_modeled(
                            &mut rpr,
                            &mut pending_raw,
                            &mut occurrences,
                            RPR_CAPS_SLOT,
                        );
                        rpr.caps = Some(parse_word_toggle(e, &prefixes)?);
                    } else if is_word_element(name.as_ref(), b"smallCaps", &prefixes) {
                        record_rpr_modeled(
                            &mut rpr,
                            &mut pending_raw,
                            &mut occurrences,
                            RPR_SMALL_CAPS_SLOT,
                        );
                        rpr.small_caps = Some(parse_word_toggle(e, &prefixes)?);
                    } else if is_word_element(name.as_ref(), b"vertAlign", &prefixes) {
                        record_rpr_modeled(
                            &mut rpr,
                            &mut pending_raw,
                            &mut occurrences,
                            RPR_VERT_ALIGN_SLOT,
                        );
                        rpr.vert_align = get_word_val_attr(e, &prefixes)?;
                    } else if is_word_element(name.as_ref(), b"shd", &prefixes) {
                        record_rpr_modeled(
                            &mut rpr,
                            &mut pending_raw,
                            &mut occurrences,
                            RPR_SHADING_SLOT,
                        );
                        rpr.shading = Some(Box::new(CT_Shd::from_xml_attrs_with_prefixes(
                            e, &prefixes,
                        )?));
                    } else if is_word_element(name.as_ref(), b"vanish", &prefixes) {
                        record_rpr_modeled(
                            &mut rpr,
                            &mut pending_raw,
                            &mut occurrences,
                            RPR_VANISH_SLOT,
                        );
                        rpr.vanish = Some(parse_word_toggle(e, &prefixes)?);
                    } else if let Some((slot, toggle)) =
                        rpr_toggle_slot(&mut rpr, name.as_ref(), &prefixes)
                    {
                        let value = parse_word_toggle(e, &prefixes)?;
                        *toggle = Some(value);
                        record_rpr_modeled(&mut rpr, &mut pending_raw, &mut occurrences, slot);
                    } else if is_word_element(name.as_ref(), b"effect", &prefixes) {
                        record_rpr_modeled(
                            &mut rpr,
                            &mut pending_raw,
                            &mut occurrences,
                            RPR_EFFECT_SLOT,
                        );
                        rpr.effect = Some(
                            get_word_val_attr(e, &prefixes)?
                                .map_or(ST_TextEffect::None, |val| ST_TextEffect::from_str(&val)),
                        );
                    } else if is_word_element(name.as_ref(), b"bdr", &prefixes) {
                        record_rpr_modeled(
                            &mut rpr,
                            &mut pending_raw,
                            &mut occurrences,
                            RPR_BORDER_SLOT,
                        );
                        rpr.border = Some(Box::new(CT_BorderEdge::from_xml_attrs_with_prefixes(
                            e, &prefixes,
                        )?));
                    } else if is_word_element(name.as_ref(), b"fitText", &prefixes)
                        && let Some(fit_text) = CT_FitText::from_xml_attrs(e, &prefixes)?
                    {
                        // As for `w:kern`, a `w:fitText` without its required
                        // `w:val` stays raw.
                        record_rpr_modeled(
                            &mut rpr,
                            &mut pending_raw,
                            &mut occurrences,
                            RPR_FIT_TEXT_SLOT,
                        );
                        rpr.fit_text = Some(Box::new(fit_text));
                    } else if is_word_element(name.as_ref(), b"em", &prefixes) {
                        record_rpr_modeled(
                            &mut rpr,
                            &mut pending_raw,
                            &mut occurrences,
                            RPR_EMPHASIS_MARK_SLOT,
                        );
                        rpr.emphasis_mark = Some(
                            get_word_val_attr(e, &prefixes)?
                                .map_or(ST_Em::None, |val| ST_Em::from_str(&val)),
                        );
                    } else if is_word_element(name.as_ref(), b"eastAsianLayout", &prefixes) {
                        record_rpr_modeled(
                            &mut rpr,
                            &mut pending_raw,
                            &mut occurrences,
                            RPR_EAST_ASIAN_LAYOUT_SLOT,
                        );
                        rpr.east_asian_layout =
                            Some(Box::new(CT_EastAsianLayout::from_xml_attrs(e, &prefixes)?));
                    } else if is_word_element(name.as_ref(), b"rtl", &prefixes) {
                        record_rpr_modeled(
                            &mut rpr,
                            &mut pending_raw,
                            &mut occurrences,
                            RPR_RTL_SLOT,
                        );
                        rpr.rtl = Some(parse_word_toggle(e, &prefixes)?);
                        rtl_carrier_required = toggle_has_unsupported_attributes(e, &prefixes)?;
                        let raw = crate::text::raw_with_external_bindings(
                            &capture_empty_element(e)?,
                            owner_bindings,
                        )?;
                        record_modeled_toggle_candidate(
                            &mut rpr.revision_xml,
                            &mut rpr.revision_xml_positions,
                            raw,
                            RPR_RTL_SLOT,
                            occurrences[RPR_RTL_SLOT as usize] - 1,
                        );
                    } else if is_word_element(name.as_ref(), b"lang", &prefixes) {
                        record_rpr_modeled(
                            &mut rpr,
                            &mut pending_raw,
                            &mut occurrences,
                            RPR_LANG_SLOT,
                        );
                        parse_language_attributes(&mut rpr, e, &prefixes)?;
                    } else if is_word_element(name.as_ref(), b"ins", &prefixes)
                        || is_word_element(name.as_ref(), b"del", &prefixes)
                    {
                        let raw = crate::text::raw_with_external_bindings(
                            &capture_empty_element(e)?,
                            owner_bindings,
                        )?;
                        if let Some(revision) = CT_Revision::from_raw(raw.clone(), &prefixes) {
                            record_rpr_modeled(
                                &mut rpr,
                                &mut pending_raw,
                                &mut occurrences,
                                RPR_MARKER_SLOT,
                            );
                            rpr.revision_markers.push(revision);
                        } else {
                            pending_raw.push(raw);
                        }
                    } else if is_word_element(name.as_ref(), b"rPrChange", &prefixes) {
                        let raw = crate::text::raw_with_external_bindings(
                            &capture_empty_element(e)?,
                            owner_bindings,
                        )?;
                        if let Some(revision) = CT_Revision::from_raw(raw.clone(), &prefixes) {
                            if let Some(previous) = rpr.change.take() {
                                rpr.revision_xml
                                    .insert(change_raw_index, previous.into_raw_xml());
                                rpr.revision_xml_positions
                                    .insert(change_raw_index, (RPR_CHANGE_SLOT, 0));
                            }
                            record_rpr_modeled(
                                &mut rpr,
                                &mut pending_raw,
                                &mut occurrences,
                                RPR_CHANGE_SLOT,
                            );
                            rpr.change = Some(Box::new(revision));
                            change_raw_index = rpr.revision_xml.len();
                        } else {
                            flush_rpr_raw(&mut rpr, &mut pending_raw, RPR_CHANGE_SLOT, 0);
                            record_rpr_raw_at(&mut rpr, raw, RPR_CHANGE_SLOT, 0);
                        }
                    } else {
                        let raw = crate::text::raw_with_external_bindings(
                            &capture_empty_element(e)?,
                            owner_bindings,
                        )?;
                        if let Some(slot) = rpr_schema_slot(name.as_ref(), &prefixes) {
                            record_rpr_raw_at(&mut rpr, raw, slot, 0);
                        } else {
                            pending_raw.push(raw);
                        }
                    }
                }
                Ok(Event::Start(ref e)) => {
                    let prefixes = word_prefixes_at(e, word_prefixes)?;
                    if is_word_element(e.name().as_ref(), b"rtl", &prefixes) {
                        let captured = capture_element(reader, e)?;
                        let raw =
                            crate::text::raw_with_external_bindings(&captured, owner_bindings)?;
                        if toggle_element_is_explicitly_empty(&captured)? {
                            record_rpr_modeled(
                                &mut rpr,
                                &mut pending_raw,
                                &mut occurrences,
                                RPR_RTL_SLOT,
                            );
                            rpr.rtl = Some(parse_word_toggle(e, &prefixes)?);
                            rtl_carrier_required = toggle_has_unsupported_attributes(e, &prefixes)?;
                            record_modeled_toggle_candidate(
                                &mut rpr.revision_xml,
                                &mut rpr.revision_xml_positions,
                                raw,
                                RPR_RTL_SLOT,
                                occurrences[RPR_RTL_SLOT as usize] - 1,
                            );
                        } else {
                            let occurrence = occurrences[RPR_RTL_SLOT as usize];
                            occurrences[RPR_RTL_SLOT as usize] += 1;
                            record_rpr_raw_at(&mut rpr, raw, RPR_RTL_SLOT, occurrence);
                        }
                    } else if is_word_element(e.name().as_ref(), b"lang", &prefixes) {
                        let captured = capture_element(reader, e)?;
                        let raw =
                            crate::text::raw_with_external_bindings(&captured, owner_bindings)?;
                        if language_element_is_explicitly_empty(&captured)? {
                            record_rpr_modeled(
                                &mut rpr,
                                &mut pending_raw,
                                &mut occurrences,
                                RPR_LANG_SLOT,
                            );
                            parse_language_attributes(&mut rpr, e, &prefixes)?;
                        } else {
                            record_rpr_raw_at(
                                &mut rpr,
                                raw,
                                RPR_LANG_SLOT,
                                occurrences[RPR_LANG_SLOT as usize],
                            );
                        }
                    } else if is_word_element(e.name().as_ref(), b"ins", &prefixes)
                        || is_word_element(e.name().as_ref(), b"del", &prefixes)
                    {
                        let raw = crate::text::raw_with_external_bindings(
                            &capture_element(reader, e)?,
                            owner_bindings,
                        )?;
                        if let Some(revision) = CT_Revision::from_raw(raw.clone(), &prefixes) {
                            record_rpr_modeled(
                                &mut rpr,
                                &mut pending_raw,
                                &mut occurrences,
                                RPR_MARKER_SLOT,
                            );
                            rpr.revision_markers.push(revision);
                        } else {
                            pending_raw.push(raw);
                        }
                    } else if is_word_element(e.name().as_ref(), b"rPrChange", &prefixes) {
                        let raw = crate::text::raw_with_external_bindings(
                            &capture_element(reader, e)?,
                            owner_bindings,
                        )?;
                        if let Some(revision) = CT_Revision::from_raw(raw.clone(), &prefixes) {
                            if let Some(previous) = rpr.change.take() {
                                rpr.revision_xml
                                    .insert(change_raw_index, previous.into_raw_xml());
                                rpr.revision_xml_positions
                                    .insert(change_raw_index, (RPR_CHANGE_SLOT, 0));
                            }
                            record_rpr_modeled(
                                &mut rpr,
                                &mut pending_raw,
                                &mut occurrences,
                                RPR_CHANGE_SLOT,
                            );
                            rpr.change = Some(Box::new(revision));
                            change_raw_index = rpr.revision_xml.len();
                        } else {
                            flush_rpr_raw(&mut rpr, &mut pending_raw, RPR_CHANGE_SLOT, 0);
                            record_rpr_raw_at(&mut rpr, raw, RPR_CHANGE_SLOT, 0);
                        }
                    } else {
                        let raw = crate::text::raw_with_external_bindings(
                            &capture_element(reader, e)?,
                            owner_bindings,
                        )?;
                        if let Some(slot) = rpr_schema_slot(e.name().as_ref(), &prefixes) {
                            record_rpr_raw_at(&mut rpr, raw, slot, 0);
                        } else {
                            pending_raw.push(raw);
                        }
                    }
                }
                Ok(Event::End(ref e)) if matches_local_name(e.name().as_ref(), b"rPr") => {
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        let final_slot = if rpr.change.is_some() {
            RPR_CHANGE_SLOT
        } else {
            RPR_END_SLOT
        };
        flush_rpr_raw(&mut rpr, &mut pending_raw, final_slot, 0);
        remove_redundant_modeled_toggle_candidate(
            &mut rpr.revision_xml,
            &mut rpr.revision_xml_positions,
            RPR_RTL_SLOT,
            rtl_carrier_required,
        );

        if let Some(output) = ordinary {
            *output = selected_style;
        }
        Ok(rpr)
    }

    pub fn to_xml<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        self.to_xml_with_word_override(writer, None)
    }

    pub(crate) fn to_xml_with_word_override<W: std::io::Write>(
        &self,
        writer: &mut Writer<W>,
        foreign_word_namespace: Option<&str>,
    ) -> Result<()> {
        self.write_xml_role(
            writer,
            foreign_word_namespace,
            crate::text::XmlWriteRole::Publication,
            None,
            None,
        )
    }

    pub(crate) fn write_xml_role<W: std::io::Write>(
        &self,
        writer: &mut Writer<W>,
        foreign_word_namespace: Option<&str>,
        role: crate::text::XmlWriteRole,
        source_context: Option<&[(&str, &str)]>,
        emission: Option<&crate::text::FieldXmlEmission>,
    ) -> Result<()> {
        self.write_xml_role_core(
            writer,
            foreign_word_namespace,
            role,
            source_context,
            emission,
            true,
        )
    }

    fn write_xml_role_core<W: std::io::Write>(
        &self,
        writer: &mut Writer<W>,
        foreign_word_namespace: Option<&str>,
        role: crate::text::XmlWriteRole,
        source_context: Option<&[(&str, &str)]>,
        emission: Option<&crate::text::FieldXmlEmission>,
        positioned: bool,
    ) -> Result<()> {
        if if positioned {
            self.is_empty()
        } else {
            self.is_empty_in(false)
        } {
            return Ok(());
        }
        if positioned
            && !self.revision_xml.is_empty()
            && self.revision_xml_positions.len() == self.revision_xml.len()
        {
            let generated = if let Some(emission) = emission {
                let (output, _outer) = emission.new_source_record_output();
                let mut generated = Writer::new(output);
                self.write_modeled_record(
                    &mut generated,
                    foreign_word_namespace,
                    role,
                    source_context,
                    Some(emission),
                )?;
                generated.into_inner().into_mapped_parts().0
            } else {
                let mut generated = Writer::new(Vec::new());
                self.write_modeled_record(
                    &mut generated,
                    foreign_word_namespace,
                    role,
                    source_context,
                    None,
                )?;
                generated.into_inner()
            };
            return write_rpr_with_positioned_raw(
                writer,
                &generated,
                self,
                foreign_word_namespace,
                emission,
            );
        }

        let mut buf = itoa::Buffer::new();
        let word_prefix = if let Some(emission) = emission {
            let plan = if role == crate::text::XmlWriteRole::Source {
                crate::text::generated_word_name(source_context.unwrap_or(&[]))?
            } else {
                crate::text::GeneratedWordName::fixed("w")
            };
            emission.write_word_attributes(writer, plan, "rPr", &[], false)?
        } else {
            let (start, prefix) = if role == crate::text::XmlWriteRole::Source {
                crate::text::generated_word_start(
                    BytesStart::new("w:rPr"),
                    &[],
                    role,
                    source_context,
                )?
            } else {
                (BytesStart::new("w:rPr"), "w".to_owned())
            };
            writer.write_event(Event::Start(start))?;
            prefix
        };

        if let Some(ref style_id) = self.style_id {
            write_word_value(writer, &word_prefix, "rStyle", style_id.as_str(), emission)?;
        }

        // rFonts
        if self.has_fonts() {
            if let Some(emission) = emission.filter(|_| self.font_extra_attributes.is_empty()) {
                let attributes = [
                    self.font_hint.as_deref().map(|v| ("hint", v)),
                    self.font_ascii.as_deref().map(|v| ("ascii", v)),
                    self.font_hansi.as_deref().map(|v| ("hAnsi", v)),
                    self.font_east_asia.as_deref().map(|v| ("eastAsia", v)),
                    self.font_cs.as_deref().map(|v| ("cs", v)),
                    self.font_ascii_theme.as_deref().map(|v| ("asciiTheme", v)),
                    self.font_hansi_theme.as_deref().map(|v| ("hAnsiTheme", v)),
                    self.font_east_asia_theme
                        .as_deref()
                        .map(|v| ("eastAsiaTheme", v)),
                    self.font_cs_theme.as_deref().map(|v| ("cstheme", v)),
                ];
                let plan = if role == crate::text::XmlWriteRole::Source {
                    crate::text::generated_word_name(source_context.unwrap_or(&[]))?
                } else {
                    crate::text::GeneratedWordName::fixed(&word_prefix)
                };
                emission.write_word_attributes(writer, plan, "rFonts", &attributes, true)?;
            } else {
                let mut e = crate::text::FieldXmlStart::word(&word_prefix, "rFonts", emission)?;
                for (name, value) in &self.font_extra_attributes {
                    e.push_raw(name, value)?;
                }
                let word_prefix = if role == crate::text::XmlWriteRole::Source {
                    e.qualify(&[], role, source_context)?
                } else {
                    word_prefix.clone()
                };
                if let Some(ref f) = self.font_hint {
                    e.push_word(&word_prefix, "hint", f.as_str())?;
                }
                if let Some(ref f) = self.font_ascii {
                    e.push_word(&word_prefix, "ascii", f.as_str())?;
                }
                if let Some(ref f) = self.font_hansi {
                    e.push_word(&word_prefix, "hAnsi", f.as_str())?;
                }
                if let Some(ref f) = self.font_east_asia {
                    e.push_word(&word_prefix, "eastAsia", f.as_str())?;
                }
                if let Some(ref f) = self.font_cs {
                    e.push_word(&word_prefix, "cs", f.as_str())?;
                }
                if let Some(ref f) = self.font_ascii_theme {
                    e.push_word(&word_prefix, "asciiTheme", f.as_str())?;
                }
                if let Some(ref f) = self.font_hansi_theme {
                    e.push_word(&word_prefix, "hAnsiTheme", f.as_str())?;
                }
                if let Some(ref f) = self.font_east_asia_theme {
                    e.push_word(&word_prefix, "eastAsiaTheme", f.as_str())?;
                }
                if let Some(ref f) = self.font_cs_theme {
                    e.push_word(&word_prefix, "cstheme", f.as_str())?;
                }
                #[cfg(test)]
                crate::text::WORD_FONT_EVENT_MATERIALIZATIONS
                    .set(crate::text::WORD_FONT_EVENT_MATERIALIZATIONS.get() + 1);
                e.write(writer, true)?;
            }
        }

        if let Some(bold) = self.bold {
            write_word_toggle(writer, "b", bold, role, source_context, emission)?;
        }
        if let Some(bold_cs) = self.bold_cs {
            write_word_toggle(writer, "bCs", bold_cs, role, source_context, emission)?;
        }
        if let Some(italic) = self.italic {
            write_word_toggle(writer, "i", italic, role, source_context, emission)?;
        }
        if let Some(italic_cs) = self.italic_cs {
            write_word_toggle(writer, "iCs", italic_cs, role, source_context, emission)?;
        }
        if let Some(caps) = self.caps {
            write_word_toggle(writer, "caps", caps, role, source_context, emission)?;
        }
        if let Some(small_caps) = self.small_caps {
            write_word_toggle(
                writer,
                "smallCaps",
                small_caps,
                role,
                source_context,
                emission,
            )?;
        }
        if let Some(strike) = self.strike {
            write_word_toggle(writer, "strike", strike, role, source_context, emission)?;
        }
        if let Some(dstrike) = self.dstrike {
            write_word_toggle(writer, "dstrike", dstrike, role, source_context, emission)?;
        }
        if let Some(outline) = self.outline {
            write_word_toggle(writer, "outline", outline, role, source_context, emission)?;
        }
        if let Some(shadow) = self.shadow {
            write_word_toggle(writer, "shadow", shadow, role, source_context, emission)?;
        }
        if let Some(emboss) = self.emboss {
            write_word_toggle(writer, "emboss", emboss, role, source_context, emission)?;
        }
        if let Some(imprint) = self.imprint {
            write_word_toggle(writer, "imprint", imprint, role, source_context, emission)?;
        }
        if let Some(no_proof) = self.no_proof {
            write_word_toggle(writer, "noProof", no_proof, role, source_context, emission)?;
        }
        if let Some(snap_to_grid) = self.snap_to_grid {
            write_word_toggle(
                writer,
                "snapToGrid",
                snap_to_grid,
                role,
                source_context,
                emission,
            )?;
        }
        if let Some(vanish) = self.vanish {
            write_word_toggle(writer, "vanish", vanish, role, source_context, emission)?;
        }
        if let Some(web_hidden) = self.web_hidden {
            write_word_toggle(
                writer,
                "webHidden",
                web_hidden,
                role,
                source_context,
                emission,
            )?;
        }

        if self.has_color() {
            if let Some(emission) = emission.filter(|_| self.color_extra_attributes.is_empty()) {
                let tint = self
                    .color_theme_tint
                    .map(crate::properties::uchar_hex_bytes);
                let shade = self
                    .color_theme_shade
                    .map(crate::properties::uchar_hex_bytes);
                let tint = tint
                    .as_ref()
                    .map(|bytes| std::str::from_utf8(bytes))
                    .transpose()?;
                let shade = shade
                    .as_ref()
                    .map(|bytes| std::str::from_utf8(bytes))
                    .transpose()?;
                let attributes = [
                    self.color.as_deref().map(|value| ("val", value)),
                    self.color_theme
                        .as_deref()
                        .map(|value| ("themeColor", value)),
                    tint.map(|value| ("themeTint", value)),
                    shade.map(|value| ("themeShade", value)),
                ];
                let plan = if role == crate::text::XmlWriteRole::Source {
                    crate::text::generated_word_name(source_context.unwrap_or(&[]))?
                } else {
                    crate::text::GeneratedWordName::fixed(&word_prefix)
                };
                emission.write_word_attributes(writer, plan, "color", &attributes, true)?;
            } else {
                let mut e = crate::text::FieldXmlStart::word(&word_prefix, "color", emission)?;
                for (name, value) in &self.color_extra_attributes {
                    e.push_raw(name, value)?;
                }
                let word_prefix = if role == crate::text::XmlWriteRole::Source {
                    e.qualify(&[], role, source_context)?
                } else {
                    word_prefix.clone()
                };
                if let Some(ref color) = self.color {
                    e.push_word(&word_prefix, "val", color.as_str())?;
                }
                if let Some(ref tc) = self.color_theme {
                    e.push_word(&word_prefix, "themeColor", tc.as_str())?;
                }
                if let Some(tint) = self.color_theme_tint {
                    const HEX: &[u8; 16] = b"0123456789ABCDEF";
                    let value = [HEX[usize::from(tint >> 4)], HEX[usize::from(tint & 15)]];
                    e.push_word(
                        &word_prefix,
                        "themeTint",
                        std::str::from_utf8(&value).expect("ASCII hex"),
                    )?;
                }
                if let Some(shade) = self.color_theme_shade {
                    const HEX: &[u8; 16] = b"0123456789ABCDEF";
                    let value = [HEX[usize::from(shade >> 4)], HEX[usize::from(shade & 15)]];
                    e.push_word(
                        &word_prefix,
                        "themeShade",
                        std::str::from_utf8(&value).expect("ASCII hex"),
                    )?;
                }
                #[cfg(test)]
                crate::text::WORD_COLOR_EVENT_MATERIALIZATIONS
                    .set(crate::text::WORD_COLOR_EVENT_MATERIALIZATIONS.get() + 1);
                e.write(writer, true)?;
            }
        }

        if let Some(ref spacing) = self.spacing {
            write_word_value(
                writer,
                &word_prefix,
                "spacing",
                buf.format(spacing.0),
                emission,
            )?;
        }
        if let Some(ws) = self.width_scale {
            write_word_value(writer, &word_prefix, "w", buf.format(ws), emission)?;
        }
        if let Some(ref kern) = self.kern {
            write_word_value(writer, &word_prefix, "kern", buf.format(kern.0), emission)?;
        }
        if let Some(pos) = self.position {
            write_word_value(writer, &word_prefix, "position", buf.format(pos), emission)?;
        }

        if let Some(ref sz) = self.sz {
            write_word_value(writer, &word_prefix, "sz", buf.format(sz.0), emission)?;
        }
        if let Some(ref sz_cs) = self.sz_cs {
            write_word_value(writer, &word_prefix, "szCs", buf.format(sz_cs.0), emission)?;
        }

        if let Some(ref highlight) = self.highlight {
            write_word_value(
                writer,
                &word_prefix,
                "highlight",
                highlight.to_str(),
                emission,
            )?;
        }

        if let Some(underline) = self.underline {
            write_word_value(writer, &word_prefix, "u", underline.to_str(), emission)?;
        }

        if let Some(ref effect) = self.effect {
            write_word_value(writer, &word_prefix, "effect", effect.as_str(), emission)?;
        }
        if let Some(ref border) = self.border {
            if let Some(emission) = emission
                .filter(|_| border.extra_attributes.is_empty() && border.source_xml.is_none())
            {
                let mut size_buffer = itoa::Buffer::new();
                let mut space_buffer = itoa::Buffer::new();
                let size = border.sz.map(|size| size_buffer.format(size));
                let space = border.space.map(|space| space_buffer.format(space));
                let attributes = [
                    Some(("val", border.val.to_str())),
                    size.map(|value| ("sz", value)),
                    space.map(|value| ("space", value)),
                    border.color.as_deref().map(|value| ("color", value)),
                ];
                let plan = if role == crate::text::XmlWriteRole::Source {
                    crate::text::generated_word_name(source_context.unwrap_or(&[]))?
                } else {
                    crate::text::GeneratedWordName::fixed("w")
                };
                emission.write_word_attributes(writer, plan, "bdr", &attributes, true)?;
            } else {
                if role == crate::text::XmlWriteRole::Source {
                    border.write_xml_in(writer, "bdr", role, source_context, emission)?;
                } else {
                    border.write_xml_in(writer, "w:bdr", role, source_context, emission)?;
                }
            }
        }

        if let Some(ref shd) = self.shading {
            if let Some(emission) = emission.filter(|_| shd.extra_attributes.is_empty()) {
                let tint = shd.theme_tint.map(crate::properties::uchar_hex_bytes);
                let shade = shd.theme_shade.map(crate::properties::uchar_hex_bytes);
                let fill_tint = shd.theme_fill_tint.map(crate::properties::uchar_hex_bytes);
                let fill_shade = shd.theme_fill_shade.map(crate::properties::uchar_hex_bytes);
                let tint = tint
                    .as_ref()
                    .map(|bytes| std::str::from_utf8(bytes))
                    .transpose()?;
                let shade = shade
                    .as_ref()
                    .map(|bytes| std::str::from_utf8(bytes))
                    .transpose()?;
                let fill_tint = fill_tint
                    .as_ref()
                    .map(|bytes| std::str::from_utf8(bytes))
                    .transpose()?;
                let fill_shade = fill_shade
                    .as_ref()
                    .map(|bytes| std::str::from_utf8(bytes))
                    .transpose()?;
                let attributes = [
                    Some(("val", shd.val.as_str())),
                    shd.color.as_deref().map(|value| ("color", value)),
                    shd.theme_color
                        .as_deref()
                        .map(|value| ("themeColor", value)),
                    tint.map(|value| ("themeTint", value)),
                    shade.map(|value| ("themeShade", value)),
                    shd.fill.as_deref().map(|value| ("fill", value)),
                    shd.theme_fill.as_deref().map(|value| ("themeFill", value)),
                    fill_tint.map(|value| ("themeFillTint", value)),
                    fill_shade.map(|value| ("themeFillShade", value)),
                ];
                let plan = if role == crate::text::XmlWriteRole::Source {
                    crate::text::generated_word_name(source_context.unwrap_or(&[]))?
                } else {
                    crate::text::GeneratedWordName::fixed("w")
                };
                emission.write_word_attributes(writer, plan, "shd", &attributes, true)?;
            } else {
                if role == crate::text::XmlWriteRole::Source {
                    shd.write_xml_in(writer, "shd", role, source_context, emission)?;
                } else {
                    shd.write_xml_in(writer, "w:shd", role, source_context, emission)?;
                }
            }
        }

        if let Some(ref fit_text) = self.fit_text {
            if let Some(emission) = emission.filter(|_| fit_text.extra_attributes.is_empty()) {
                let mut value_buffer = itoa::Buffer::new();
                let mut id_buffer = itoa::Buffer::new();
                let id = fit_text.id.map(|id| id_buffer.format(id));
                let attributes = [
                    Some(("val", value_buffer.format(fit_text.val.0))),
                    id.map(|value| ("id", value)),
                ];
                let plan = if role == crate::text::XmlWriteRole::Source {
                    crate::text::generated_word_name(source_context.unwrap_or(&[]))?
                } else {
                    crate::text::GeneratedWordName::fixed("w")
                };
                emission.write_word_attributes(writer, plan, "fitText", &attributes, true)?;
            } else {
                fit_text.write_xml(writer, role, source_context, emission)?;
            }
        }

        if let Some(ref vert_align) = self.vert_align {
            write_word_value(
                writer,
                &word_prefix,
                "vertAlign",
                vert_align.as_str(),
                emission,
            )?;
        }

        if let Some(rtl) = self.rtl {
            write_word_toggle(writer, "rtl", rtl, role, source_context, emission)?;
        }
        if let Some(complex_script) = self.complex_script {
            write_word_toggle(writer, "cs", complex_script, role, source_context, emission)?;
        }
        if let Some(ref emphasis_mark) = self.emphasis_mark {
            write_word_value(writer, &word_prefix, "em", emphasis_mark.as_str(), emission)?;
        }

        if self.language.is_some()
            || self.language_east_asia.is_some()
            || self.language_bidi.is_some()
            || !self.language_extra_attributes.is_empty()
        {
            if let Some(emission) = emission.filter(|_| self.language_extra_attributes.is_empty()) {
                let attributes = [
                    self.language.as_deref().map(|v| ("val", v)),
                    self.language_east_asia.as_deref().map(|v| ("eastAsia", v)),
                    self.language_bidi.as_deref().map(|v| ("bidi", v)),
                ];
                let plan = if role == crate::text::XmlWriteRole::Source {
                    crate::text::generated_word_name(source_context.unwrap_or(&[]))?
                } else {
                    crate::text::GeneratedWordName::fixed(&word_prefix)
                };
                emission.write_word_attributes(writer, plan, "lang", &attributes, true)?;
            } else {
                let mut e = crate::text::FieldXmlStart::word(&word_prefix, "lang", emission)?;
                if role == crate::text::XmlWriteRole::Source {
                    for (name, value) in &self.language_extra_attributes {
                        e.push_raw(name, value)?;
                    }
                }
                let word_prefix = if role == crate::text::XmlWriteRole::Source {
                    e.qualify(&[], role, source_context)?
                } else {
                    word_prefix.clone()
                };
                if let Some(language) = &self.language {
                    e.push_word(&word_prefix, "val", language.as_str())?;
                }
                if let Some(language) = &self.language_east_asia {
                    e.push_word(&word_prefix, "eastAsia", language.as_str())?;
                }
                if let Some(language) = &self.language_bidi {
                    e.push_word(&word_prefix, "bidi", language.as_str())?;
                }
                if role != crate::text::XmlWriteRole::Source {
                    for (name, value) in &self.language_extra_attributes {
                        e.push(name, value)?;
                    }
                }
                e.write(writer, true)?;
            }
        }

        if let Some(ref layout) = self.east_asian_layout {
            if let Some(emission) = emission.filter(|_| layout.extra_attributes.is_empty()) {
                let mut id_buffer = itoa::Buffer::new();
                let id = layout.id.map(|id| id_buffer.format(id));
                let attributes = [
                    id.map(|value| ("id", value)),
                    layout
                        .combine
                        .map(|value| ("combine", if value { "1" } else { "0" })),
                    layout
                        .combine_brackets
                        .as_deref()
                        .map(|value| ("combineBrackets", value)),
                    layout
                        .vert
                        .map(|value| ("vert", if value { "1" } else { "0" })),
                    layout
                        .vert_compress
                        .map(|value| ("vertCompress", if value { "1" } else { "0" })),
                ];
                let plan = if role == crate::text::XmlWriteRole::Source {
                    crate::text::generated_word_name(source_context.unwrap_or(&[]))?
                } else {
                    crate::text::GeneratedWordName::fixed("w")
                };
                emission.write_word_attributes(
                    writer,
                    plan,
                    "eastAsianLayout",
                    &attributes,
                    true,
                )?;
            } else {
                layout.write_xml(writer, role, source_context, emission)?;
            }
        }
        if let Some(spec_vanish) = self.spec_vanish {
            write_word_toggle(
                writer,
                "specVanish",
                spec_vanish,
                role,
                source_context,
                emission,
            )?;
        }
        if let Some(office_math) = self.office_math {
            write_word_toggle(writer, "oMath", office_math, role, source_context, emission)?;
        }

        for revision in &self.revision_markers {
            revision.write_xml_with_role_override(
                writer,
                foreign_word_namespace,
                role,
                source_context,
                emission,
            )?;
        }
        if positioned {
            for raw in &self.revision_xml {
                crate::text::write_raw_with_word_override_in(
                    writer,
                    raw,
                    foreign_word_namespace,
                    emission,
                )?;
            }
        }
        if let Some(change) = &self.change {
            change.write_xml_with_role_override(
                writer,
                foreign_word_namespace,
                role,
                source_context,
                emission,
            )?;
        }

        if let Some(emission) = emission {
            emission.write_word_end(writer, &word_prefix, "rPr")?;
        } else {
            writer.write_event(Event::End(BytesEnd::new(format!("{word_prefix}:rPr"))))?;
        }
        Ok(())
    }

    fn write_modeled_record<W: std::io::Write>(
        &self,
        writer: &mut Writer<W>,
        foreign_word_namespace: Option<&str>,
        role: crate::text::XmlWriteRole,
        source_context: Option<&[(&str, &str)]>,
        emission: Option<&crate::text::FieldXmlEmission>,
    ) -> Result<()> {
        if self.is_empty_in(false) {
            crate::text::FieldXmlStart::new("w:rPr", emission)?.write(writer, false)?;
            if let Some(emission) = emission {
                emission.write_end_name(writer, "w:rPr")?;
            } else {
                writer.write_event(Event::End(BytesEnd::new("w:rPr")))?;
            }
            Ok(())
        } else {
            self.write_xml_role_core(
                writer,
                foreign_word_namespace,
                role,
                source_context,
                emission,
                false,
            )
        }
    }

    /// Whether any `w:rFonts` attribute is set.
    fn has_fonts(&self) -> bool {
        self.font_ascii.is_some()
            || self.font_hansi.is_some()
            || self.font_east_asia.is_some()
            || self.font_cs.is_some()
            || self.font_ascii_theme.is_some()
            || self.font_hansi_theme.is_some()
            || self.font_east_asia_theme.is_some()
            || self.font_cs_theme.is_some()
            || self.font_hint.is_some()
            || !self.font_extra_attributes.is_empty()
    }

    /// Whether any `w:color` attribute is set.
    fn has_color(&self) -> bool {
        self.color.is_some()
            || self.color_theme.is_some()
            || self.color_theme_tint.is_some()
            || self.color_theme_shade.is_some()
            || !self.color_extra_attributes.is_empty()
    }

    fn is_empty(&self) -> bool {
        self.is_empty_in(true)
    }

    fn is_empty_in(&self, raw: bool) -> bool {
        self.style_id.is_none()
            && self.font_ascii.is_none()
            && self.font_hansi.is_none()
            && self.font_east_asia.is_none()
            && self.font_cs.is_none()
            && self.font_ascii_theme.is_none()
            && self.font_hansi_theme.is_none()
            && self.font_east_asia_theme.is_none()
            && self.font_cs_theme.is_none()
            && self.font_hint.is_none()
            && self.font_extra_attributes.is_empty()
            && self.bold.is_none()
            && self.bold_cs.is_none()
            && self.italic.is_none()
            && self.italic_cs.is_none()
            && self.underline.is_none()
            && self.strike.is_none()
            && self.dstrike.is_none()
            && self.sz.is_none()
            && self.sz_cs.is_none()
            && self.color.is_none()
            && self.color_theme.is_none()
            && self.color_theme_tint.is_none()
            && self.color_theme_shade.is_none()
            && self.color_extra_attributes.is_empty()
            && self.highlight.is_none()
            && self.caps.is_none()
            && self.small_caps.is_none()
            && self.vert_align.is_none()
            && self.spacing.is_none()
            && self.width_scale.is_none()
            && self.position.is_none()
            && self.shading.is_none()
            && self.vanish.is_none()
            && self.outline.is_none()
            && self.shadow.is_none()
            && self.emboss.is_none()
            && self.imprint.is_none()
            && self.no_proof.is_none()
            && self.snap_to_grid.is_none()
            && self.web_hidden.is_none()
            && self.kern.is_none()
            && self.effect.is_none()
            && self.border.is_none()
            && self.fit_text.is_none()
            && self.complex_script.is_none()
            && self.emphasis_mark.is_none()
            && self.east_asian_layout.is_none()
            && self.spec_vanish.is_none()
            && self.office_math.is_none()
            && self.rtl.is_none()
            && self.language.is_none()
            && self.language_east_asia.is_none()
            && self.language_bidi.is_none()
            && self.language_extra_attributes.is_empty()
            && self.revision_markers.is_empty()
            && self.change.is_none()
            && (!raw || self.revision_xml.is_empty())
    }

    /// Merge another CT_RPr into this one (non-None fields override).
    /// Used for style inheritance.
    pub fn merge_from(&mut self, other: &CT_RPr) {
        // Only a recognized semantic solid fill participates in style precedence.
        // Opaque extension children are not transported into an effective style.
        if let Some((color, alpha)) = other.solid_text_fill() {
            let mut retained = Vec::new();
            let mut positions = Vec::new();
            for (index, raw) in self.revision_xml.iter().enumerate() {
                let probe = CT_RPr {
                    revision_xml: vec![raw.clone()],
                    ..CT_RPr::default()
                };
                if probe.solid_text_fill().is_none() {
                    retained.push(raw.clone());
                    positions.push(
                        self.revision_xml_positions
                            .get(index)
                            .copied()
                            .unwrap_or((RPR_END_SLOT, 0)),
                    );
                }
            }
            let amount = (alpha * 100_000.0).round() as u32;
            retained.push(format!("<w14:textFill xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\"><w14:solidFill><w14:srgbClr w14:val=\"{color}\"><w14:alpha w14:val=\"{amount}\"/></w14:srgbClr></w14:solidFill></w14:textFill>").into_bytes());
            positions.push((RPR_END_SLOT, 0));
            self.revision_xml = retained;
            self.revision_xml_positions = positions;
        }
        if other.style_id.is_some() {
            self.style_id = other.style_id.clone();
        }
        if other.font_ascii.is_some() {
            self.font_ascii = other.font_ascii.clone();
        }
        if other.font_hansi.is_some() {
            self.font_hansi = other.font_hansi.clone();
        }
        if other.font_east_asia.is_some() {
            self.font_east_asia = other.font_east_asia.clone();
        }
        if other.font_cs.is_some() {
            self.font_cs = other.font_cs.clone();
        }
        if other.font_ascii_theme.is_some() {
            self.font_ascii_theme = other.font_ascii_theme.clone();
        }
        if other.font_hansi_theme.is_some() {
            self.font_hansi_theme = other.font_hansi_theme.clone();
        }
        if other.font_east_asia_theme.is_some() {
            self.font_east_asia_theme = other.font_east_asia_theme.clone();
        }
        if other.font_cs_theme.is_some() {
            self.font_cs_theme = other.font_cs_theme.clone();
        }
        if other.font_hint.is_some() {
            self.font_hint = other.font_hint.clone();
        }
        if !other.font_extra_attributes.is_empty() {
            self.font_extra_attributes = other.font_extra_attributes.clone();
        }
        if other.bold.is_some() {
            self.bold = other.bold;
        }
        if other.bold_cs.is_some() {
            self.bold_cs = other.bold_cs;
        }
        if other.italic.is_some() {
            self.italic = other.italic;
        }
        if other.italic_cs.is_some() {
            self.italic_cs = other.italic_cs;
        }
        if other.underline.is_some() {
            self.underline = other.underline;
        }
        if other.strike.is_some() {
            self.strike = other.strike;
        }
        if other.dstrike.is_some() {
            self.dstrike = other.dstrike;
        }
        if other.sz.is_some() {
            self.sz = other.sz;
        }
        if other.sz_cs.is_some() {
            self.sz_cs = other.sz_cs;
        }
        if other.color.is_some() {
            self.color = other.color.clone();
        }
        if other.color_theme.is_some() {
            self.color_theme = other.color_theme.clone();
        }
        if other.color_theme_tint.is_some() {
            self.color_theme_tint = other.color_theme_tint;
        }
        if other.color_theme_shade.is_some() {
            self.color_theme_shade = other.color_theme_shade;
        }
        if !other.color_extra_attributes.is_empty() {
            self.color_extra_attributes = other.color_extra_attributes.clone();
        }
        if other.highlight.is_some() {
            self.highlight = other.highlight;
        }
        if other.caps.is_some() {
            self.caps = other.caps;
        }
        if other.small_caps.is_some() {
            self.small_caps = other.small_caps;
        }
        if other.vert_align.is_some() {
            self.vert_align = other.vert_align.clone();
        }
        if other.spacing.is_some() {
            self.spacing = other.spacing;
        }
        if other.width_scale.is_some() {
            self.width_scale = other.width_scale;
        }
        if other.position.is_some() {
            self.position = other.position;
        }
        if other.shading.is_some() {
            self.shading = other.shading.clone();
        }
        if other.vanish.is_some() {
            self.vanish = other.vanish;
        }
        if other.outline.is_some() {
            self.outline = other.outline;
        }
        if other.shadow.is_some() {
            self.shadow = other.shadow;
        }
        if other.emboss.is_some() {
            self.emboss = other.emboss;
        }
        if other.imprint.is_some() {
            self.imprint = other.imprint;
        }
        if other.no_proof.is_some() {
            self.no_proof = other.no_proof;
        }
        if other.snap_to_grid.is_some() {
            self.snap_to_grid = other.snap_to_grid;
        }
        if other.web_hidden.is_some() {
            self.web_hidden = other.web_hidden;
        }
        if other.kern.is_some() {
            self.kern = other.kern;
        }
        if other.effect.is_some() {
            self.effect = other.effect.clone();
        }
        if other.border.is_some() {
            self.border = other.border.clone();
        }
        if other.fit_text.is_some() {
            self.fit_text = other.fit_text.clone();
        }
        if other.complex_script.is_some() {
            self.complex_script = other.complex_script;
        }
        if other.emphasis_mark.is_some() {
            self.emphasis_mark = other.emphasis_mark.clone();
        }
        if other.east_asian_layout.is_some() {
            self.east_asian_layout = other.east_asian_layout.clone();
        }
        if other.spec_vanish.is_some() {
            self.spec_vanish = other.spec_vanish;
        }
        if other.office_math.is_some() {
            self.office_math = other.office_math;
        }
        if other.rtl.is_some() {
            self.rtl = other.rtl;
        }
        if other.language.is_some() {
            self.language = other.language.clone();
        }
        if other.language_east_asia.is_some() {
            self.language_east_asia = other.language_east_asia.clone();
        }
        if other.language_bidi.is_some() {
            self.language_bidi = other.language_bidi.clone();
        }
        if !other.language_extra_attributes.is_empty() {
            self.language_extra_attributes = other.language_extra_attributes.clone();
        }
    }
}

fn parse_language_attributes(
    rpr: &mut CT_RPr,
    element: &BytesStart<'_>,
    prefixes: &[String],
) -> Result<()> {
    for attribute in element.attributes() {
        let attribute = attribute?;
        let key = attribute.key.as_ref();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, element.decoder())?
            .into_owned();
        if is_word_attribute(key, b"val", prefixes) {
            rpr.language = Some(value);
        } else if is_word_attribute(key, b"eastAsia", prefixes) {
            rpr.language_east_asia = Some(value);
        } else if is_word_attribute(key, b"bidi", prefixes) {
            rpr.language_bidi = Some(value);
        } else {
            rpr.language_extra_attributes
                .push((std::str::from_utf8(key)?.to_owned(), value));
        }
    }
    Ok(())
}

pub(crate) fn language_element_is_explicitly_empty(xml: &[u8]) -> Result<bool> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    if !matches!(reader.read_event_into(&mut buffer)?, Event::Start(_)) {
        return Ok(false);
    }
    buffer.clear();
    Ok(matches!(
        reader.read_event_into(&mut buffer)?,
        Event::End(_)
    ))
}

pub(crate) fn write_word_toggle<W: std::io::Write>(
    writer: &mut Writer<W>,
    local: &str,
    value: bool,
    role: crate::text::XmlWriteRole,
    source_context: Option<&[(&str, &str)]>,
    emission: Option<&crate::text::FieldXmlEmission>,
) -> Result<()> {
    if let Some(emission) = emission {
        let plan = if role == crate::text::XmlWriteRole::Source {
            crate::text::generated_word_name(source_context.unwrap_or(&[]))?
        } else {
            crate::text::GeneratedWordName::fixed("w")
        };
        emission.write_word_attributes(
            writer,
            plan,
            local,
            &[(!value).then_some(("val", "false"))],
            true,
        )?;
        Ok(())
    } else if role == crate::text::XmlWriteRole::Source {
        crate::properties::write_toggle_source(writer, local, value, source_context)
    } else {
        write_toggle(writer, &format!("w:{local}"), value)
    }
}

fn write_word_value<W: std::io::Write>(
    writer: &mut Writer<W>,
    prefix: &str,
    local: &str,
    value: &str,
    emission: Option<&crate::text::FieldXmlEmission>,
) -> Result<()> {
    if let Some(emission) = emission {
        return emission.write_word_value(writer, prefix, local, value);
    }
    let mut element = BytesStart::new(format!("{prefix}:{local}"));
    element.push_attribute((format!("{prefix}:val").as_str(), value));
    #[cfg(test)]
    crate::text::WORD_VALUE_EVENT_MATERIALIZATIONS
        .set(crate::text::WORD_VALUE_EVENT_MATERIALIZATIONS.get() + 1);
    writer.write_event(Event::Empty(element))?;
    Ok(())
}

fn write_rpr_with_positioned_raw<W: std::io::Write>(
    writer: &mut Writer<W>,
    generated: &[u8],
    rpr: &CT_RPr,
    foreign_word_namespace: Option<&str>,
    emission: Option<&crate::text::FieldXmlEmission>,
) -> Result<()> {
    let mut raw_order = (0..rpr.revision_xml.len())
        .filter(|index| {
            replay_modeled_toggle_raw(
                rpr.revision_xml_positions[*index],
                rpr.revision_xml_positions[*index].0 != RPR_RTL_SLOT || rpr.rtl.is_some(),
            )
        })
        .collect::<Vec<_>>();
    raw_order.sort_by_key(|index| {
        let (slot, occurrence) = effective_rpr_raw_position(rpr, *index);
        (slot, occurrence, *index)
    });
    let mut raw_index = 0usize;
    let mut occurrences = [0usize; RPR_END_SLOT as usize + 1];
    let mut reader = Reader::from_reader(generated);
    reader.config_mut().trim_text(false);
    let mut inside = false;
    loop {
        let event_start = reader.buffer_position() as usize;
        match reader.read_event()? {
            Event::Start(element) if !inside => {
                inside = true;
                if let Some(emission) = emission {
                    emission.write_copy(
                        writer,
                        &generated[event_start..reader.buffer_position() as usize],
                    )?;
                } else {
                    writer.write_event(Event::Start(element))?;
                }
            }
            Event::Start(element) => {
                let slot = rpr_slot_for_name(element.local_name().as_ref());
                write_rpr_raw_before(
                    writer,
                    rpr,
                    &raw_order,
                    &mut raw_index,
                    (slot, occurrences[slot as usize]),
                    foreign_word_namespace,
                    emission,
                )?;
                let end = crate::text::generated_subtree_end(&mut reader, element.name().as_ref())?;
                if let Some(emission) = emission {
                    emission.write_copy(writer, &generated[event_start..end])?;
                } else {
                    writer.get_mut().write_all(&generated[event_start..end])?;
                }
                occurrences[slot as usize] += 1;
            }
            Event::Empty(element) => {
                let slot = rpr_slot_for_name(element.local_name().as_ref());
                write_rpr_raw_before(
                    writer,
                    rpr,
                    &raw_order,
                    &mut raw_index,
                    (slot, occurrences[slot as usize]),
                    foreign_word_namespace,
                    emission,
                )?;
                let attributes = crate::properties::modeled_toggle_attributes(
                    &rpr.revision_xml,
                    &rpr.revision_xml_positions,
                    slot,
                )?;
                let mut element = crate::text::FieldXmlStart::from_start(element, emission)?;
                for (name, value) in attributes {
                    element.push(&name, &value)?;
                }
                element.write(writer, true)?;
                occurrences[slot as usize] += 1;
            }
            Event::End(element) if inside => {
                write_rpr_raw_before(
                    writer,
                    rpr,
                    &raw_order,
                    &mut raw_index,
                    (RPR_END_SLOT, 0),
                    foreign_word_namespace,
                    emission,
                )?;
                while let Some(index) = raw_order.get(raw_index).copied() {
                    crate::text::write_raw_with_word_override_in(
                        writer,
                        &rpr.revision_xml[index],
                        foreign_word_namespace,
                        emission,
                    )?;
                    raw_index += 1;
                }
                if let Some(emission) = emission {
                    emission.write_copy(
                        writer,
                        &generated[event_start..reader.buffer_position() as usize],
                    )?;
                } else {
                    writer.write_event(Event::End(element))?;
                }
                return Ok(());
            }
            Event::Eof => return Ok(()),
            event => {
                if let Some(emission) = emission {
                    emission.write_copy(
                        writer,
                        &generated[event_start..reader.buffer_position() as usize],
                    )?;
                } else {
                    writer.write_event(event)?;
                }
            }
        }
    }
}

fn write_rpr_raw_before<W: std::io::Write>(
    writer: &mut Writer<W>,
    rpr: &CT_RPr,
    raw_order: &[usize],
    raw_index: &mut usize,
    before: (u8, usize),
    foreign_word_namespace: Option<&str>,
    emission: Option<&crate::text::FieldXmlEmission>,
) -> Result<()> {
    while let Some(index) = raw_order.get(*raw_index).copied() {
        let position = effective_rpr_raw_position(rpr, index);
        if position > before {
            break;
        }
        crate::text::write_raw_with_word_override_in(
            writer,
            &rpr.revision_xml[index],
            foreign_word_namespace,
            emission,
        )?;
        *raw_index += 1;
    }
    Ok(())
}

fn effective_rpr_raw_position(rpr: &CT_RPr, index: usize) -> (u8, usize) {
    let mut position = rpr.revision_xml_positions[index];
    position.1 = raw_occurrence(position);
    if position.0 == RPR_RTL_SLOT
        && let Some((carrier_index, carrier_occurrence)) = rpr
            .revision_xml_positions
            .iter()
            .enumerate()
            .find(|candidate| {
                candidate.1.0 == RPR_RTL_SLOT && raw_is_modeled_attribute_carrier(*candidate.1)
            })
            .map(|(carrier_index, candidate)| (carrier_index, raw_occurrence(*candidate)))
    {
        position.1 = usize::from(
            position.1 > carrier_occurrence
                || (position.1 == carrier_occurrence && index > carrier_index),
        );
    }
    if rpr.change.is_some() && position.0 >= RPR_CHANGE_SLOT {
        (RPR_CHANGE_SLOT, 0)
    } else {
        position
    }
}

fn rpr_slot_for_name(local: &[u8]) -> u8 {
    match local {
        b"rStyle" => RPR_STYLE_SLOT,
        b"rFonts" => RPR_FONTS_SLOT,
        b"b" => RPR_BOLD_SLOT,
        b"bCs" => RPR_BOLD_CS_SLOT,
        b"i" => RPR_ITALIC_SLOT,
        b"iCs" => RPR_ITALIC_CS_SLOT,
        b"caps" => RPR_CAPS_SLOT,
        b"smallCaps" => RPR_SMALL_CAPS_SLOT,
        b"strike" => RPR_STRIKE_SLOT,
        b"dstrike" => RPR_DSTRIKE_SLOT,
        b"outline" => 10,
        b"shadow" => 11,
        b"emboss" => 12,
        b"imprint" => 13,
        b"noProof" => 14,
        b"snapToGrid" => 15,
        b"vanish" => RPR_VANISH_SLOT,
        b"webHidden" => 17,
        b"color" => RPR_COLOR_SLOT,
        b"spacing" => RPR_SPACING_SLOT,
        b"w" => RPR_WIDTH_SLOT,
        b"kern" => 21,
        b"position" => RPR_POSITION_SLOT,
        b"sz" => RPR_SIZE_SLOT,
        b"szCs" => RPR_SIZE_CS_SLOT,
        b"highlight" => RPR_HIGHLIGHT_SLOT,
        b"u" => RPR_UNDERLINE_SLOT,
        b"effect" => 27,
        b"bdr" => 28,
        b"shd" => RPR_SHADING_SLOT,
        b"fitText" => 30,
        b"vertAlign" => RPR_VERT_ALIGN_SLOT,
        b"rtl" => 32,
        b"cs" => 33,
        b"em" => 34,
        b"lang" => RPR_LANG_SLOT,
        b"eastAsianLayout" => 36,
        b"specVanish" => 37,
        b"oMath" => 38,
        b"ins" | b"del" => RPR_MARKER_SLOT,
        b"rPrChange" => RPR_CHANGE_SLOT,
        _ => RPR_END_SLOT,
    }
}

/// Validate modeled run-property leaves on checked replacement only.
pub(crate) fn validate_checked_property_child(
    element: &BytesStart<'_>,
    prefixes: &[String],
) -> Result<()> {
    let local = element.local_name();
    let name = local.as_ref();
    let value = get_word_val_attr(element, prefixes)?;
    if matches!(
        name,
        b"b" | b"bCs"
            | b"i"
            | b"iCs"
            | b"caps"
            | b"smallCaps"
            | b"strike"
            | b"dstrike"
            | b"outline"
            | b"shadow"
            | b"emboss"
            | b"imprint"
            | b"noProof"
            | b"snapToGrid"
            | b"vanish"
            | b"webHidden"
            | b"rtl"
            | b"cs"
            | b"specVanish"
            | b"oMath"
    ) {
        if value
            .as_deref()
            .is_some_and(|value| !matches!(value, "true" | "false" | "on" | "off" | "1" | "0"))
        {
            return Err(crate::error::OxmlError::InvalidValue(
                "element XML run toggle has an invalid Word boolean".into(),
            ));
        }
    } else if name == b"u" {
        if value
            .as_deref()
            .is_some_and(|value| ST_Underline::from_str(value).is_err())
        {
            return Err(crate::error::OxmlError::InvalidValue(
                "element XML underline has an invalid w:val".into(),
            ));
        }
    } else if matches!(
        name,
        b"sz"
            | b"szCs"
            | b"kern"
            | b"spacing"
            | b"position"
            | b"w"
            | b"rStyle"
            | b"highlight"
            | b"vertAlign"
            | b"color"
    ) {
        let value = value.ok_or_else(|| {
            crate::error::OxmlError::InvalidValue(format!(
                "element XML {} requires w:val",
                String::from_utf8_lossy(name)
            ))
        })?;
        let valid = match name {
            b"sz" | b"szCs" | b"kern" | b"w" => {
                parse_integer_measurement::<i32>(element.name().as_ref(), b"w:val", &value)
                    .is_ok_and(|value| value >= 0)
            }
            b"spacing" | b"position" => {
                parse_integer_measurement::<i32>(element.name().as_ref(), b"w:val", &value).is_ok()
            }
            b"highlight" => ST_HighlightColor::from_str(&value).is_ok(),
            b"vertAlign" => matches!(value.as_str(), "baseline" | "superscript" | "subscript"),
            b"color" => {
                value == "auto"
                    || (value.len() == 6 && value.bytes().all(|byte| byte.is_ascii_hexdigit()))
            }
            _ => !value.is_empty(),
        };
        if !valid {
            return Err(crate::error::OxmlError::InvalidValue(format!(
                "element XML {} has an invalid w:val",
                String::from_utf8_lossy(name)
            )));
        }
    }
    Ok(())
}

pub(crate) fn rpr_schema_slot(name: &[u8], word_prefixes: &[String]) -> Option<u8> {
    let local = name.rsplit(|byte| *byte == b':').next().unwrap_or(name);
    let slot = rpr_slot_for_name(local);
    (slot != RPR_END_SLOT && is_word_element(name, local, word_prefixes)).then_some(slot)
}

/// Attribute ownership of the full run-property reader. Numbering owns only a subset.
pub(crate) fn source_modeled_attributes(
    parent: &[u8],
    local: &[u8],
) -> Option<&'static [&'static [u8]]> {
    if parent != b"rPr" {
        return None;
    }
    rpr_schema_slot(local, &[String::new()])?;
    let shared = crate::numbering::property_attributes(crate::numbering::PropertyKind::Run, local);
    Some(match local {
        b"rFonts" => &[
            b"ascii",
            b"hAnsi",
            b"eastAsia",
            b"cs",
            b"asciiTheme",
            b"hAnsiTheme",
            b"eastAsiaTheme",
            b"cstheme",
            b"hint",
        ],
        b"color" => &[b"val", b"themeColor", b"themeTint", b"themeShade"],
        b"shd" => crate::properties::source_shading_attributes(),
        b"bdr" | b"border" => crate::properties::source_border_attributes(),
        b"lang" => &[b"val", b"eastAsia", b"bidi"],
        b"fitText" => &[b"val", b"id"],
        b"eastAsianLayout" => &[
            b"id",
            b"combine",
            b"combineBrackets",
            b"vert",
            b"vertCompress",
        ],
        b"ins" | b"del" | b"moveFrom" | b"moveTo" | b"rPrChange" => &[b"id", b"author", b"date"],
        _ if !shared.is_empty() => shared,
        _ => &[b"val"],
    })
}

pub(crate) fn source_value_is_modeled(
    parent: &[u8],
    local: &[u8],
    key: &[u8],
    value: &str,
) -> bool {
    if parent != b"rPr" {
        return true;
    }
    if key == b"val" {
        if local == b"highlight" {
            return ST_HighlightColor::from_str(value).is_ok();
        }
        if local == b"u" {
            return ST_Underline::from_str(value).is_ok();
        }
        if matches!(
            local,
            b"b" | b"bCs"
                | b"i"
                | b"iCs"
                | b"caps"
                | b"smallCaps"
                | b"strike"
                | b"dstrike"
                | b"outline"
                | b"shadow"
                | b"emboss"
                | b"imprint"
                | b"noProof"
                | b"snapToGrid"
                | b"vanish"
                | b"webHidden"
                | b"rtl"
                | b"cs"
                | b"specVanish"
                | b"oMath"
        ) {
            return crate::numbering::parse_on_off_token(value).is_some();
        }
    }
    if local == b"color" && matches!(key, b"themeTint" | b"themeShade") {
        return parse_uchar_hex(value).is_some();
    }
    if local == b"eastAsianLayout" && matches!(key, b"combine" | b"vert" | b"vertCompress") {
        return crate::numbering::parse_on_off_token(value).is_some();
    }
    true
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::namespace::W_NS;

    fn parse_rpr(xml: &str) -> CT_RPr {
        let full = format!("<w:rPr>{xml}</w:rPr>");
        let mut reader = Reader::from_str(&full);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if matches_local_name(e.name().as_ref(), b"rPr") => break,
                _ => {}
            }
            buf.clear();
        }
        CT_RPr::from_xml(&mut reader).unwrap()
    }

    #[test]
    fn shared_run_measurements_preserve_unsigned_signed_and_missing_values() {
        let rpr = parse_rpr(
            r#"<w:sz w:val="4294967295"/><w:sz/><w:szCs w:val="2.5"/><w:spacing w:val="-2.5"/><w:position w:val="-2147483648"/><w:w w:val="4294967295"/><w:kern/>"#,
        );
        assert_eq!(rpr.sz, Some(HalfPoint(u32::MAX)));
        assert_eq!(rpr.sz_cs, Some(HalfPoint(3)));
        assert_eq!(rpr.spacing, Some(Twips(-3)));
        assert_eq!(rpr.position, Some(i32::MIN));
        assert_eq!(rpr.width_scale, Some(u32::MAX));
        assert_eq!(rpr.kern, None);
        assert!(rpr.revision_xml.iter().any(|raw| raw == b"<w:kern/>"));
        assert!(
            RunMeasurement::for_local(b"sz")
                .unwrap()
                .decode(b"w:sz", Some("-1"))
                .is_err()
        );
        assert!(
            RunMeasurement::for_local(b"w")
                .unwrap()
                .decode(b"w:w", Some("1.5"))
                .is_err()
        );
        assert!(
            RunMeasurement::for_local(b"spacing")
                .unwrap()
                .decode(b"w:spacing", Some("-1"))
                .is_ok()
        );
    }

    #[test]
    fn parse_basic_rpr() {
        let rpr = parse_rpr(r#"<w:b/><w:i/><w:sz w:val="24"/><w:color w:val="FF0000"/>"#);
        assert_eq!(rpr.bold, Some(true));
        assert_eq!(rpr.italic, Some(true));
        assert_eq!(rpr.sz, Some(HalfPoint(24)));
        assert_eq!(rpr.color, Some("FF0000".to_string()));
    }

    #[test]
    fn parse_rpr_spacing() {
        let rpr = parse_rpr(r#"<w:spacing w:val="20"/><w:w w:val="150"/><w:position w:val="-4"/>"#);
        assert_eq!(rpr.spacing, Some(Twips(20)));
        assert_eq!(rpr.width_scale, Some(150));
        assert_eq!(rpr.position, Some(-4));
    }

    #[test]
    fn round_trip_rpr() {
        let original = CT_RPr {
            bold: Some(true),
            italic: Some(true),
            sz: Some(HalfPoint(24)),
            color: Some("FF0000".to_string()),
            underline: Some(ST_Underline::Single),
            spacing: Some(Twips(20)),
            ..Default::default()
        };

        let mut output = Vec::new();
        let mut writer = Writer::new(&mut output);
        original.to_xml(&mut writer).unwrap();
        let xml = String::from_utf8(output).unwrap();

        let inner = xml
            .strip_prefix("<w:rPr>")
            .unwrap()
            .strip_suffix("</w:rPr>")
            .unwrap();
        let parsed = parse_rpr(inner);
        assert_eq!(parsed.bold, original.bold);
        assert_eq!(parsed.italic, original.italic);
        assert_eq!(parsed.sz, original.sz);
        assert_eq!(parsed.color, original.color);
        assert_eq!(parsed.underline, original.underline);
        assert_eq!(parsed.spacing, original.spacing);
    }

    #[test]
    fn merge_rpr() {
        let mut base = CT_RPr {
            bold: Some(true),
            sz: Some(HalfPoint(24)),
            ..Default::default()
        };
        let override_rpr = CT_RPr {
            sz: Some(HalfPoint(28)),
            italic: Some(true),
            ..Default::default()
        };
        base.merge_from(&override_rpr);
        assert_eq!(base.bold, Some(true)); // kept
        assert_eq!(base.sz, Some(HalfPoint(28))); // overridden
        assert_eq!(base.italic, Some(true)); // added
    }

    #[test]
    fn run_language_parses_aliases_and_rejects_foreign_same_local_attributes() {
        let rpr = parse_rpr(&format!(
            r#"<q:lang xmlns:q="{}" xmlns:x="urn:foreign" q:val="en-US" q:eastAsia="zh-CN" q:bidi="ar-SA" x:val="ignored" x:kept="raw"/>"#,
            W_NS
        ));
        assert_eq!(rpr.language.as_deref(), Some("en-US"));
        assert_eq!(rpr.language_east_asia.as_deref(), Some("zh-CN"));
        assert_eq!(rpr.language_bidi.as_deref(), Some("ar-SA"));

        let mut output = Vec::new();
        rpr.to_xml(&mut Writer::new(&mut output)).unwrap();
        let output = String::from_utf8(output).unwrap();
        assert!(output.contains(r#"<w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="ar-SA""#));
        assert!(output.contains(r#"x:kept="raw""#));
        assert!(!output.contains(r#"w:val="ignored""#));

        let explicit = parse_rpr(&format!(
            r#"<q:lang xmlns:q="{}" q:val="en-GB"></q:lang>"#,
            W_NS
        ));
        assert_eq!(explicit.language.as_deref(), Some("en-GB"));
    }

    #[test]
    fn run_language_round_trips_at_its_schema_slot_and_cascades() {
        let mut base = CT_RPr {
            language: Some("fr-FR".to_owned()),
            language_east_asia: Some("ja-JP".to_owned()),
            ..Default::default()
        };
        base.merge_from(&CT_RPr {
            language: Some("de-DE".to_owned()),
            language_bidi: Some("ar-SA".to_owned()),
            ..Default::default()
        });
        assert_eq!(base.language.as_deref(), Some("de-DE"));
        assert_eq!(base.language_east_asia.as_deref(), Some("ja-JP"));
        assert_eq!(base.language_bidi.as_deref(), Some("ar-SA"));

        let mut output = Vec::new();
        base.to_xml(&mut Writer::new(&mut output)).unwrap();
        let output = String::from_utf8(output).unwrap();
        assert!(output.find("w:vertAlign").is_none());
        assert!(output.find("w:lang").unwrap() < output.rfind("</w:rPr>").unwrap());
    }

    /// Every `EG_RPrBase` child F-265 typed, in schema order, with one
    /// unmodelled producer sibling between two of them.
    const F265_EVERY_NEW_CHILD: &str = concat!(
        r#"<w:outline/><w:shadow/><w:emboss/><w:imprint/><w:noProof/><w:snapToGrid/>"#,
        r#"<w:webHidden/><w:kern w:val="16"/><w:effect w:val="antsRed"/>"#,
        r#"<w:bdr w:val="single" w:sz="4" w:space="1" w:color="FF0000"/>"#,
        r#"<w:fitText w:val="1440" w:id="3"/><w:cs/><w:em w:val="dot"/>"#,
        r#"<w:eastAsianLayout w:id="7" w:combine="1" w:combineBrackets="round" w:vert="1" w:vertCompress="1"/>"#,
        r#"<w:specVanish/><w:oMath/>"#,
    );

    #[test]
    fn every_new_run_property_parses_writes_and_merges() {
        let rpr = parse_rpr(F265_EVERY_NEW_CHILD);
        assert_eq!(rpr.outline, Some(true));
        assert_eq!(rpr.shadow, Some(true));
        assert_eq!(rpr.emboss, Some(true));
        assert_eq!(rpr.imprint, Some(true));
        assert_eq!(rpr.no_proof, Some(true));
        assert_eq!(rpr.snap_to_grid, Some(true));
        assert_eq!(rpr.web_hidden, Some(true));
        assert_eq!(rpr.kern, Some(HalfPoint(16)));
        assert_eq!(rpr.effect, Some(ST_TextEffect::AntsRed));
        assert_eq!(rpr.border.as_ref().map(|border| border.sz), Some(Some(4)));
        assert_eq!(
            rpr.fit_text.as_deref(),
            Some(&CT_FitText {
                val: Twips(1440),
                id: Some(3),
                extra_attributes: Vec::new(),
            })
        );
        assert_eq!(rpr.complex_script, Some(true));
        assert_eq!(rpr.emphasis_mark, Some(ST_Em::Dot));
        assert_eq!(
            rpr.east_asian_layout.as_deref(),
            Some(&CT_EastAsianLayout {
                id: Some(7),
                combine: Some(true),
                combine_brackets: Some("round".to_owned()),
                vert: Some(true),
                vert_compress: Some(true),
                extra_attributes: Vec::new(),
            })
        );
        assert_eq!(rpr.spec_vanish, Some(true));
        assert_eq!(rpr.office_math, Some(true));

        // A run property set carrying only one new child is not empty, and
        // every new child survives a write and a re-read in schema order.
        for single in [
            "<w:outline/>",
            "<w:webHidden/>",
            r#"<w:kern w:val="18"/>"#,
            r#"<w:effect w:val="shimmer"/>"#,
            r#"<w:bdr w:val="double"/>"#,
            r#"<w:fitText w:val="720"/>"#,
            "<w:cs/>",
            r#"<w:em w:val="circle"/>"#,
            r#"<w:eastAsianLayout w:id="1"/>"#,
            "<w:specVanish/>",
            "<w:oMath/>",
        ] {
            let single = parse_rpr(single);
            assert!(!single.is_empty(), "{single:?}");
            let mut output = Vec::new();
            single.to_xml(&mut Writer::new(&mut output)).unwrap();
            assert!(!output.is_empty());
        }

        let mut merged = CT_RPr::default();
        merged.merge_from(&rpr);
        assert_eq!(merged, rpr);
    }

    #[test]
    fn new_run_properties_write_at_their_schema_ordinals() {
        let rpr = parse_rpr(F265_EVERY_NEW_CHILD);
        let mut output = Vec::new();
        rpr.to_xml(&mut Writer::new(&mut output)).unwrap();
        let output = String::from_utf8(output).unwrap();
        let ordered = [
            "<w:outline/>",
            "<w:shadow/>",
            "<w:emboss/>",
            "<w:imprint/>",
            "<w:noProof/>",
            "<w:snapToGrid/>",
            "<w:webHidden/>",
            "<w:kern ",
            "<w:effect ",
            "<w:bdr ",
            "<w:fitText ",
            "<w:cs/>",
            "<w:em ",
            "<w:eastAsianLayout ",
            "<w:specVanish/>",
            "<w:oMath/>",
        ]
        .map(|needle| {
            output
                .find(needle)
                .unwrap_or_else(|| panic!("{needle} in {output}"))
        });
        assert!(ordered.windows(2).all(|pair| pair[0] < pair[1]), "{output}");

        let reparsed = parse_rpr(
            output
                .strip_prefix("<w:rPr>")
                .unwrap()
                .strip_suffix("</w:rPr>")
                .unwrap(),
        );
        assert_eq!(reparsed, rpr);
    }

    #[test]
    fn a_new_element_missing_its_required_value_stays_raw() {
        for source in ["<w:kern/>", r#"<w:fitText w:id="4"/>"#] {
            let rpr = parse_rpr(source);
            assert_eq!(rpr.kern, None, "{source}");
            assert_eq!(rpr.fit_text, None, "{source}");

            let mut output = Vec::new();
            rpr.to_xml(&mut Writer::new(&mut output)).unwrap();
            let output = String::from_utf8(output).unwrap();
            assert!(output.contains(source), "{source} missing from {output}");
        }
    }

    #[test]
    fn new_run_properties_read_under_an_aliased_word_prefix() {
        let rpr = parse_rpr(&format!(
            concat!(
                r#"<q:outline xmlns:q="{ns}"/><q:kern xmlns:q="{ns}" q:val="20"/>"#,
                r#"<q:effect xmlns:q="{ns}" q:val="lights"/>"#,
                r#"<q:fitText xmlns:q="{ns}" q:val="960" q:id="4"/>"#,
                r#"<q:em xmlns:q="{ns}" q:val="comma"/>"#,
                r#"<q:eastAsianLayout xmlns:q="{ns}" q:vert="0"/>"#,
            ),
            ns = W_NS
        ));
        assert_eq!(rpr.outline, Some(true));
        assert_eq!(rpr.kern, Some(HalfPoint(20)));
        assert_eq!(rpr.effect, Some(ST_TextEffect::Lights));
        assert_eq!(rpr.fit_text.as_ref().map(|fit| fit.val), Some(Twips(960)));
        assert_eq!(rpr.emphasis_mark, Some(ST_Em::Comma));
        assert_eq!(
            rpr.east_asian_layout
                .as_ref()
                .and_then(|layout| layout.vert),
            Some(false)
        );

        let mut output = Vec::new();
        rpr.to_xml(&mut Writer::new(&mut output)).unwrap();
        let output = String::from_utf8(output).unwrap();
        assert!(output.contains("<w:outline/>"), "{output}");
        assert!(output.contains(r#"<w:kern w:val="20"/>"#), "{output}");
        assert!(output.contains(r#"<w:effect w:val="lights"/>"#), "{output}");
    }

    #[test]
    fn producer_tokens_and_foreign_attributes_survive_the_new_elements() {
        let rpr = parse_rpr(concat!(
            r#"<w:rFonts xmlns:x="urn:producer" w:ascii="Arial" w:eastAsiaTheme="minorEastAsia" w:cstheme="minorBidi" w:hint="eastAsia" x:kept="rfonts"/>"#,
            r#"<w:color xmlns:x="urn:producer" w:val="4472C4" w:themeColor="accent1" w:themeTint="66" w:themeShade="zz" x:kept="color"/>"#,
            r#"<w:effect w:val="producerEffect"/><w:em w:val="producerMark"/>"#,
            r#"<w:fitText xmlns:x="urn:producer" w:val="600" x:kept="fit"/>"#,
            r#"<w:eastAsianLayout xmlns:x="urn:producer" w:id="2" x:kept="layout"/>"#,
        ));
        assert_eq!(rpr.font_east_asia_theme.as_deref(), Some("minorEastAsia"));
        assert_eq!(rpr.font_cs_theme.as_deref(), Some("minorBidi"));
        assert_eq!(rpr.font_hint.as_deref(), Some("eastAsia"));
        assert_eq!(rpr.color_theme_tint, Some(0x66));
        // "zz" is not two hex digits, so it is retained verbatim rather than
        // parsed into the typed slot.
        assert_eq!(rpr.color_theme_shade, None);
        assert_eq!(
            rpr.effect,
            Some(ST_TextEffect::Other("producerEffect".to_owned()))
        );
        assert_eq!(
            rpr.emphasis_mark,
            Some(ST_Em::Other("producerMark".to_owned()))
        );

        let mut output = Vec::new();
        rpr.to_xml(&mut Writer::new(&mut output)).unwrap();
        let output = String::from_utf8(output).unwrap();
        for retained in [
            r#"x:kept="rfonts""#,
            r#"x:kept="color""#,
            r#"x:kept="fit""#,
            r#"x:kept="layout""#,
            r#"w:themeShade="zz""#,
            r#"w:eastAsiaTheme="minorEastAsia""#,
            r#"w:cstheme="minorBidi""#,
            r#"w:hint="eastAsia""#,
            r#"w:themeTint="66""#,
            r#"w:val="producerEffect""#,
            r#"w:val="producerMark""#,
        ] {
            assert!(
                output.contains(retained),
                "{retained} missing from {output}"
            );
        }
    }

    #[test]
    fn a_theme_only_colour_reaches_the_serialized_element() {
        let rpr = parse_rpr(r#"<w:color w:themeColor="accent2" w:themeShade="BF"/>"#);
        assert_eq!(rpr.color, None);
        assert_eq!(rpr.color_theme.as_deref(), Some("accent2"));
        assert_eq!(rpr.color_theme_shade, Some(0xBF));

        let mut output = Vec::new();
        rpr.to_xml(&mut Writer::new(&mut output)).unwrap();
        let output = String::from_utf8(output).unwrap();
        assert!(
            output.contains(r#"<w:color w:themeColor="accent2" w:themeShade="BF"/>"#),
            "{output}"
        );
    }

    #[test]
    fn malformed_language_after_the_modeled_occurrence_keeps_its_relative_position() {
        let rpr = parse_rpr(r#"<w:lang w:val="en-US"/><w:lang><w:producerExtension/></w:lang>"#);

        let mut output = Vec::new();
        rpr.to_xml(&mut Writer::new(&mut output)).unwrap();
        let output_text = std::str::from_utf8(&output).unwrap();
        assert!(
            output_text.find(r#"w:val="en-US""#).unwrap()
                < output_text.find("w:producerExtension").unwrap()
        );

        let mut reader = Reader::from_reader(output.as_slice());
        let mut buffer = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer).unwrap() {
                Event::Start(element) if element.local_name().as_ref() == b"rPr" => break,
                Event::Eof => panic!("serialized run properties have no root"),
                _ => {}
            }
            buffer.clear();
        }
        let reparsed = CT_RPr::from_xml(&mut reader).unwrap();
        let mut repeated = Vec::new();
        reparsed.to_xml(&mut Writer::new(&mut repeated)).unwrap();
        assert_eq!(repeated, output);
    }
}
