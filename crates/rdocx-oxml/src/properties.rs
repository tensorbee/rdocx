//! Paragraph properties (`CT_PPr`) and run properties (`CT_RPr`).

use std::borrow::Cow;

use quick_xml::events::attributes::Attribute;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::QName;
use quick_xml::{Reader, Writer, XmlVersion};

use crate::error::{OxmlError, Result};
use crate::namespace::{W_NS, matches_local_name};
use crate::run_properties::language_element_is_explicitly_empty;
use crate::shared::ST_OnOff;

pub use crate::paragraph_properties::{CT_FramePr, CT_PPr};
pub use crate::run_properties::{CT_EastAsianLayout, CT_FitText, CT_RPr, ST_Em, ST_TextEffect};

/// `CT_Shd` — Shading/background fill.
///
/// The six theme attributes are modeled rather than dropped. `w:themeTint`,
/// `w:themeShade`, `w:themeFillTint` and `w:themeFillShade` are
/// `ST_UcharHexNumber`, a single byte written as two hex digits, and they are
/// re-serialized in the canonical upper-case spelling Word itself writes.
#[derive(Debug, Clone, PartialEq)]
pub struct CT_Shd {
    /// Shading pattern (e.g. "clear", "solid", "horzStripe")
    pub val: String,
    /// Foreground color hex
    pub color: Option<String>,
    /// Background fill color hex
    pub fill: Option<String>,
    /// Foreground theme colour reference (`w:themeColor`).
    pub theme_color: Option<String>,
    /// Foreground theme tint, 0 to 255 (`w:themeTint`).
    pub theme_tint: Option<u8>,
    /// Foreground theme shade, 0 to 255 (`w:themeShade`).
    pub theme_shade: Option<u8>,
    /// Background theme colour reference (`w:themeFill`).
    pub theme_fill: Option<String>,
    /// Background theme tint, 0 to 255 (`w:themeFillTint`).
    pub theme_fill_tint: Option<u8>,
    /// Background theme shade, 0 to 255 (`w:themeFillShade`).
    pub theme_fill_shade: Option<u8>,
    /// Attributes this type does not model, in source order.
    ///
    /// A theme tint or shade whose value is not two hex digits reaches a
    /// serialized element through here rather than being discarded. Names and
    /// values keep their stored spelling and are written ahead of the modeled
    /// attributes, so a caller storing one itself owns the escaping.
    pub extra_attributes: Vec<(String, String)>,
}

impl Default for CT_Shd {
    /// The pattern Word omits, matching what the parser assumes for a
    /// `w:shd` that carries no `w:val`.
    fn default() -> Self {
        CT_Shd {
            val: "clear".to_owned(),
            color: None,
            fill: None,
            theme_color: None,
            theme_tint: None,
            theme_shade: None,
            theme_fill: None,
            theme_fill_tint: None,
            theme_fill_shade: None,
            extra_attributes: Vec::new(),
        }
    }
}

/// Parse an `ST_UcharHexNumber` attribute value.
pub(crate) fn parse_uchar_hex(value: &str) -> Option<u8> {
    (value.len() == 2 && value.bytes().all(|byte| byte.is_ascii_hexdigit()))
        .then(|| u8::from_str_radix(value, 16).ok())
        .flatten()
}

pub(crate) fn push_uchar_hex(e: &mut BytesStart<'_>, name: &str, value: u8) {
    const DIGITS: &[u8; 16] = b"0123456789ABCDEF";
    let encoded = [
        DIGITS[usize::from(value >> 4)],
        DIGITS[usize::from(value & 0x0f)],
    ];
    e.push_attribute((name.as_bytes(), encoded.as_slice()));
}

impl CT_Shd {
    pub fn from_xml_attrs(e: &BytesStart) -> Result<Self> {
        let mut shd = CT_Shd::default();

        for attr in e.attributes() {
            let attr = attr?;
            let key = attr.key.as_ref();
            let v = std::str::from_utf8(&attr.value)?;
            if matches_local_name(key, b"val") {
                shd.val = v.to_string();
            } else if matches_local_name(key, b"color") {
                shd.color = Some(v.to_string());
            } else if matches_local_name(key, b"fill") {
                shd.fill = Some(v.to_string());
            } else if matches_local_name(key, b"themeColor") {
                shd.theme_color = Some(v.to_string());
            } else if matches_local_name(key, b"themeFill") {
                shd.theme_fill = Some(v.to_string());
            } else if let Some(slot) = shd.theme_byte_slot(key) {
                match parse_uchar_hex(v) {
                    Some(parsed) => *slot = Some(parsed),
                    None => shd
                        .extra_attributes
                        .push((std::str::from_utf8(key)?.to_owned(), v.to_owned())),
                }
            } else {
                shd.extra_attributes
                    .push((std::str::from_utf8(key)?.to_owned(), v.to_owned()));
            }
        }

        Ok(shd)
    }

    pub(crate) fn from_xml_attrs_with_prefixes(
        e: &BytesStart,
        word_prefixes: &[String],
    ) -> Result<Self> {
        let mut shd = CT_Shd::default();

        for attr in e.attributes() {
            let attr = attr?;
            let key = attr.key.as_ref();
            let v = std::str::from_utf8(&attr.value)?;
            if is_word_attribute(key, b"val", word_prefixes) {
                shd.val = v.to_string();
            } else if is_word_attribute(key, b"color", word_prefixes) {
                shd.color = Some(v.to_string());
            } else if is_word_attribute(key, b"fill", word_prefixes) {
                shd.fill = Some(v.to_string());
            } else if is_word_attribute(key, b"themeColor", word_prefixes) {
                shd.theme_color = Some(v.to_string());
            } else if is_word_attribute(key, b"themeFill", word_prefixes) {
                shd.theme_fill = Some(v.to_string());
            } else if let Some(slot) = shd.word_theme_byte_slot(key, word_prefixes) {
                match parse_uchar_hex(v) {
                    Some(parsed) => *slot = Some(parsed),
                    None => shd
                        .extra_attributes
                        .push((std::str::from_utf8(key)?.to_owned(), v.to_owned())),
                }
            } else {
                shd.extra_attributes
                    .push((std::str::from_utf8(key)?.to_owned(), v.to_owned()));
            }
        }

        Ok(shd)
    }

    /// The typed slot a prefix-agnostic theme byte attribute writes into.
    fn theme_byte_slot(&mut self, key: &[u8]) -> Option<&mut Option<u8>> {
        if matches_local_name(key, b"themeTint") {
            Some(&mut self.theme_tint)
        } else if matches_local_name(key, b"themeShade") {
            Some(&mut self.theme_shade)
        } else if matches_local_name(key, b"themeFillTint") {
            Some(&mut self.theme_fill_tint)
        } else if matches_local_name(key, b"themeFillShade") {
            Some(&mut self.theme_fill_shade)
        } else {
            None
        }
    }

    /// The typed slot a Word-prefixed theme byte attribute writes into.
    fn word_theme_byte_slot(
        &mut self,
        key: &[u8],
        word_prefixes: &[String],
    ) -> Option<&mut Option<u8>> {
        if is_word_attribute(key, b"themeTint", word_prefixes) {
            Some(&mut self.theme_tint)
        } else if is_word_attribute(key, b"themeShade", word_prefixes) {
            Some(&mut self.theme_shade)
        } else if is_word_attribute(key, b"themeFillTint", word_prefixes) {
            Some(&mut self.theme_fill_tint)
        } else if is_word_attribute(key, b"themeFillShade", word_prefixes) {
            Some(&mut self.theme_fill_shade)
        } else {
            None
        }
    }

    pub fn write_xml<W: std::io::Write>(&self, writer: &mut Writer<W>, tag: &str) -> Result<()> {
        let mut e = BytesStart::new(tag);
        for (name, value) in &self.extra_attributes {
            e.push_attribute(Attribute {
                key: QName(name.as_bytes()),
                value: Cow::Borrowed(value.as_bytes()),
            });
        }
        e.push_attribute(("w:val", self.val.as_str()));
        if let Some(ref c) = self.color {
            e.push_attribute(("w:color", c.as_str()));
        }
        if let Some(ref c) = self.theme_color {
            e.push_attribute(("w:themeColor", c.as_str()));
        }
        if let Some(tint) = self.theme_tint {
            push_uchar_hex(&mut e, "w:themeTint", tint);
        }
        if let Some(shade) = self.theme_shade {
            push_uchar_hex(&mut e, "w:themeShade", shade);
        }
        if let Some(ref f) = self.fill {
            e.push_attribute(("w:fill", f.as_str()));
        }
        if let Some(ref f) = self.theme_fill {
            e.push_attribute(("w:themeFill", f.as_str()));
        }
        if let Some(tint) = self.theme_fill_tint {
            push_uchar_hex(&mut e, "w:themeFillTint", tint);
        }
        if let Some(shade) = self.theme_fill_shade {
            push_uchar_hex(&mut e, "w:themeFillShade", shade);
        }
        writer.write_event(Event::Empty(e))?;
        Ok(())
    }
}

const RAW_MODELED_ATTRIBUTES_FLAG: usize = 1 << (usize::BITS - 1);
const RAW_MODELED_DUPLICATE_FLAG: usize = 1 << (usize::BITS - 2);
const RAW_MODELED_FLAGS: usize = RAW_MODELED_ATTRIBUTES_FLAG | RAW_MODELED_DUPLICATE_FLAG;

pub(crate) fn raw_occurrence(position: (u8, usize)) -> usize {
    position.1 & !RAW_MODELED_FLAGS
}

pub(crate) fn raw_is_modeled_attribute_carrier(position: (u8, usize)) -> bool {
    position.1 & RAW_MODELED_ATTRIBUTES_FLAG != 0
}

pub(crate) fn modeled_attribute_carrier_occurrence(occurrence: usize) -> usize {
    occurrence | RAW_MODELED_ATTRIBUTES_FLAG
}

fn raw_is_modeled_duplicate(position: (u8, usize)) -> bool {
    position.1 & RAW_MODELED_DUPLICATE_FLAG != 0
}

fn modeled_duplicate_occurrence(occurrence: usize) -> usize {
    occurrence | RAW_MODELED_DUPLICATE_FLAG
}

pub(crate) fn replay_modeled_toggle_raw(
    position: (u8, usize),
    modeled_value_is_present: bool,
) -> bool {
    !raw_is_modeled_attribute_carrier(position)
        && (!raw_is_modeled_duplicate(position) || modeled_value_is_present)
}

pub(crate) fn record_modeled_toggle_candidate(
    raw_xml: &mut Vec<Vec<u8>>,
    positions: &mut Vec<(u8, usize)>,
    raw: Vec<u8>,
    slot: u8,
    occurrence: usize,
) {
    let replacing_carrier = positions
        .iter()
        .any(|position| position.0 == slot && raw_is_modeled_attribute_carrier(*position));
    if replacing_carrier {
        for position in positions
            .iter_mut()
            .filter(|position| position.0 == slot && raw_is_modeled_attribute_carrier(**position))
        {
            position.1 = modeled_duplicate_occurrence(raw_occurrence(*position));
        }
    }
    raw_xml.push(raw);
    positions.push((slot, modeled_attribute_carrier_occurrence(occurrence)));
}

pub(crate) fn remove_redundant_modeled_toggle_candidate(
    raw_xml: &mut Vec<Vec<u8>>,
    positions: &mut Vec<(u8, usize)>,
    slot: u8,
    carrier_required: bool,
) {
    if carrier_required {
        return;
    }
    if let Some(index) = positions
        .iter()
        .rposition(|position| position.0 == slot && raw_is_modeled_attribute_carrier(*position))
    {
        let carrier_occurrence = raw_occurrence(positions[index]);
        for (candidate_index, position) in positions.iter_mut().enumerate() {
            if candidate_index == index || position.0 != slot {
                continue;
            }
            let flags = position.1 & RAW_MODELED_FLAGS;
            let occurrence = raw_occurrence(*position);
            position.1 = flags
                | usize::from(
                    occurrence > carrier_occurrence
                        || (occurrence == carrier_occurrence && candidate_index > index),
                );
        }
        raw_xml.remove(index);
        positions.remove(index);
    }
}

pub(crate) fn toggle_element_is_explicitly_empty(xml: &[u8]) -> Result<bool> {
    language_element_is_explicitly_empty(xml)
}

pub(crate) fn toggle_has_unsupported_attributes(
    element: &BytesStart<'_>,
    word_prefixes: &[String],
) -> Result<bool> {
    for attribute in element.attributes() {
        let attribute = attribute?;
        if !is_word_attribute(attribute.key.as_ref(), b"val", word_prefixes) {
            return Ok(true);
        }
    }
    Ok(false)
}

pub(crate) fn append_modeled_toggle_attributes(
    element: &mut BytesStart<'_>,
    raw_xml: &[Vec<u8>],
    positions: &[(u8, usize)],
    slot: u8,
    _occurrence: usize,
) -> Result<()> {
    for (raw, position) in raw_xml.iter().zip(positions) {
        if position.0 != slot || !raw_is_modeled_attribute_carrier(*position) {
            continue;
        }
        let mut reader = Reader::from_reader(raw.as_slice());
        let mut buffer = Vec::new();
        let source = match reader.read_event_into(&mut buffer)? {
            Event::Start(source) | Event::Empty(source) => source.into_owned(),
            _ => continue,
        };
        let prefixes = word_prefixes_at(&source, &["w".to_owned()])?;
        let mut preserved = Vec::new();
        for attribute in source.attributes() {
            let attribute = attribute?;
            if is_word_attribute(attribute.key.as_ref(), b"val", &prefixes) {
                continue;
            }
            let name = std::str::from_utf8(attribute.key.as_ref())?.to_owned();
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, source.decoder())?
                .into_owned();
            preserved.push((name, value));
        }
        for (name, value) in &preserved {
            element.push_attribute((name.as_str(), value.as_str()));
        }
    }
    Ok(())
}

pub(crate) fn word_prefixes_at(
    start: &BytesStart<'_>,
    inherited: &[String],
) -> Result<Vec<String>> {
    let mut prefixes = inherited.to_vec();
    for attribute in start.attributes() {
        let attribute = attribute?;
        let name = attribute.key.as_ref();
        let prefix = if name == b"xmlns" {
            b"".as_slice()
        } else if let Some(prefix) = name.strip_prefix(b"xmlns:") {
            prefix
        } else {
            continue;
        };
        let prefix = std::str::from_utf8(prefix)?.to_string();
        prefixes.retain(|candidate| candidate != &prefix);
        let value =
            attribute.decoded_and_normalized_value(XmlVersion::Implicit1_0, start.decoder())?;
        if value.as_bytes() == W_NS.as_bytes() {
            prefixes.push(prefix);
        }
    }
    Ok(prefixes)
}

fn is_word_name(name: &[u8], word_prefixes: &[String]) -> bool {
    let Some(separator) = name.iter().position(|byte| *byte == b':') else {
        return word_prefixes.iter().any(String::is_empty);
    };
    word_prefixes
        .iter()
        .any(|prefix| prefix.as_bytes() == &name[..separator])
}

pub(crate) fn is_word_element(name: &[u8], local: &[u8], word_prefixes: &[String]) -> bool {
    matches_local_name(name, local) && is_word_name(name, word_prefixes)
}

pub(crate) fn is_word_attribute(key: &[u8], local: &[u8], word_prefixes: &[String]) -> bool {
    let Some(separator) = key.iter().position(|byte| *byte == b':') else {
        return false;
    };
    key.get(separator + 1..) == Some(local)
        && word_prefixes
            .iter()
            .any(|prefix| prefix.as_bytes() == &key[..separator])
}

pub(crate) fn get_word_val_attr(
    e: &BytesStart,
    word_prefixes: &[String],
) -> Result<Option<String>> {
    for attr in e.attributes() {
        let attr = attr?;
        if is_word_attribute(attr.key.as_ref(), b"val", word_prefixes) {
            return Ok(Some(std::str::from_utf8(&attr.value)?.to_string()));
        }
    }
    Ok(None)
}

/// Parse an integer measurement attribute, such as an `ST_TwipsMeasure`,
/// `ST_SignedTwipsMeasure` or `ST_HpsMeasure`, tolerating a decimal.
///
/// The schemas allow only an integer, but Google Docs writes a decimal part
/// (`4320.0`, `2210.0000000000005`), and LibreOffice and Word open it. A
/// decimal is read as the nearest integer, half away from zero, decided on
/// the decimal digits themselves so no float rounding intervenes. An
/// exponent, a non-finite spelling, a missing digit, or an exact value
/// outside `T` is refused with an error naming the element, the attribute
/// and the value.
pub(crate) fn parse_integer_measurement<T: TryFrom<i64>>(
    element: &[u8],
    attribute: &[u8],
    value: &str,
) -> Result<T> {
    let error = |problem: &str| {
        OxmlError::InvalidValue(format!(
            "{}/@{} {value:?} {problem}",
            String::from_utf8_lossy(element),
            String::from_utf8_lossy(attribute),
        ))
    };
    let (negative, unsigned) = match value.as_bytes().first() {
        Some(b'-') => (true, &value[1..]),
        Some(b'+') => (false, &value[1..]),
        _ => (false, value),
    };
    let (whole, fraction) = unsigned.split_once('.').unwrap_or((unsigned, "0"));
    if whole.is_empty()
        || fraction.is_empty()
        || !whole.bytes().all(|byte| byte.is_ascii_digit())
        || !fraction.bytes().all(|byte| byte.is_ascii_digit())
    {
        return Err(error("is not an integer measurement"));
    }

    // Eighteen digits cannot overflow an i64 and exceed every target type.
    let whole = whole.trim_start_matches('0');
    if whole.len() > 18 {
        return Err(error("is out of range"));
    }
    let whole = whole
        .bytes()
        .fold(0i64, |sum, digit| sum * 10 + i64::from(digit - b'0'));
    let signed = |magnitude: i64| if negative { -magnitude } else { magnitude };
    // The exact value must be in range, not only its rounding, so a
    // fraction counts as the next integer away from zero here.
    let exceeds = whole + i64::from(fraction.bytes().any(|digit| digit != b'0'));
    if T::try_from(signed(exceeds)).is_err() {
        return Err(error("is out of range"));
    }
    let rounded = whole + i64::from(fraction.as_bytes()[0] >= b'5');
    T::try_from(signed(rounded)).map_err(|_| error("is out of range"))
}

pub(crate) fn parse_word_toggle(e: &BytesStart, word_prefixes: &[String]) -> Result<bool> {
    let val = get_word_val_attr(e, word_prefixes)?;
    Ok(ST_OnOff::from_str_or_default(val.as_deref()).is_on())
}

/// Write a toggle element.
pub(crate) fn write_toggle<W: std::io::Write>(
    writer: &mut Writer<W>,
    tag: &str,
    value: bool,
) -> Result<()> {
    if value {
        writer.write_event(Event::Empty(BytesStart::new(tag)))?;
    } else {
        let mut e = BytesStart::new(tag);
        e.push_attribute(("w:val", "false"));
        writer.write_event(Event::Empty(e))?;
    }
    Ok(())
}

// These sets are owned by the corresponding typed shading/border adapters.
pub(crate) fn source_shading_attributes() -> &'static [&'static [u8]] {
    &[
        b"val",
        b"color",
        b"fill",
        b"themeColor",
        b"themeTint",
        b"themeShade",
        b"themeFill",
        b"themeFillTint",
        b"themeFillShade",
    ]
}
pub(crate) fn source_border_attributes() -> &'static [&'static [u8]] {
    &[b"val", b"sz", b"space", b"color"]
}

type ResidualName = (Option<String>, String);
struct ResidualNode {
    name: ResidualName,
    known: bool,
    text: bool,
    attributes: Vec<String>,
    scope: std::collections::BTreeMap<String, String>,
    range: std::ops::Range<usize>,
    opening_end: usize,
    items: Vec<ResidualItem>,
}
enum ResidualItem {
    Node(usize),
    Event { raw: Vec<u8>, modeled: bool },
}

/// Return the source facts omitted from the existing typed Word model.
///
/// This read-only bridge serves comparison internals. Its strings are an opaque
/// internal identity representation, not an authoring or persistent interchange
/// format. Modeled values retain their existing parser normalization. Opaque
/// attributes, payloads, events, order and effective context remain source facts.
#[doc(hidden)]
pub fn protected_source_residual(
    xml: &[u8],
    inherited_bindings: &[(String, String)],
) -> Result<Vec<String>> {
    let mut wrapper = BytesStart::new("rdocxResidual");
    for (prefix, uri) in inherited_bindings {
        let name = if prefix == "xmlns" || prefix.starts_with("xmlns:") {
            prefix.clone()
        } else if prefix.is_empty() {
            "xmlns".to_owned()
        } else {
            format!("xmlns:{prefix}")
        };
        wrapper.push_attribute((name.as_str(), uri.as_str()));
    }
    let mut writer = Writer::new(Vec::new());
    writer.write_event(Event::Start(wrapper))?;
    writer.get_mut().extend_from_slice(xml);
    writer.write_event(Event::End(quick_xml::events::BytesEnd::new(
        "rdocxResidual",
    )))?;
    let source = writer.into_inner();
    let mut reader = quick_xml::NsReader::from_reader(source.as_slice());
    reader.config_mut().trim_text(false);
    let mut nodes = Vec::<ResidualNode>::new();
    let mut stack = Vec::<usize>::new();
    loop {
        let before = reader.buffer_position() as usize;
        let (namespace, event) = reader.read_resolved_event()?;
        let empty = matches!(event, Event::Empty(_));
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let namespace = residual_namespace(namespace)?;
                let local = std::str::from_utf8(element.local_name().as_ref())?.to_owned();
                let parent = stack.last().copied();
                let parent_local = parent
                    .filter(|index| {
                        *index != 0
                            && nodes[*index].known
                            && nodes[*index].name.0.as_deref() == Some(W_NS)
                    })
                    .map(|index| nodes[index].name.1.as_bytes())
                    .unwrap_or_default();
                let root = parent == Some(0);
                let owned = if nodes.is_empty() {
                    Some(&[][..])
                } else if namespace.as_deref() == Some(W_NS) && (root || !parent_local.is_empty()) {
                    residual_owned_attributes(parent_local, local.as_bytes(), root)
                } else {
                    None
                };
                let known = owned.is_some();
                let text = known
                    && matches!(
                        local.as_str(),
                        "t" | "delText" | "instrText" | "delInstrText"
                    );
                let mut attributes = Vec::new();
                for attribute in element.attributes() {
                    let attribute = attribute?;
                    if attribute.key.as_ref() == b"xmlns"
                        || attribute.key.as_ref().starts_with(b"xmlns:")
                    {
                        continue;
                    }
                    let (namespace, name) = reader.resolver().resolve_attribute(attribute.key);
                    let namespace = residual_namespace(namespace)?;
                    let decoded = attribute
                        .decoded_and_normalized_value(XmlVersion::Implicit1_0, element.decoder())?
                        .into_owned();
                    let modeled = if namespace.as_deref() == Some(W_NS) {
                        owned.is_some_and(|keys| keys.contains(&name.as_ref()))
                            && residual_owned_value(
                                parent_local,
                                local.as_bytes(),
                                name.as_ref(),
                                &decoded,
                            )
                    } else if namespace.as_deref() == Some("http://www.w3.org/XML/1998/namespace") {
                        name.as_ref() == b"space"
                            && matches!(local.as_str(), "t" | "delText")
                            && parent_local == b"r"
                            && matches!(decoded.as_str(), "default" | "preserve")
                    } else {
                        namespace.as_deref() == Some(crate::namespace::R_NS)
                            && name.as_ref() == b"id"
                            && matches!(
                                (parent_local, local.as_str()),
                                (b"sectPr", "headerReference" | "footerReference")
                                    | (b"p", "hyperlink")
                            )
                    };
                    if !modeled || !known {
                        let identity = format!(
                            "{namespace:?}:{:?}:{:?}:{decoded:?}",
                            name.as_ref(),
                            attribute.value.as_ref()
                        );
                        attributes.push(if known {
                            identity
                        } else {
                            // The opaque root keeps every non-declaration attribute,
                            // including its lexical name and encoded value.
                            format!("{:?}:{identity}", attribute.key.as_ref())
                        });
                    }
                }
                if known {
                    attributes.sort();
                }
                let mut scope = std::collections::BTreeMap::new();
                for (prefix, uri) in reader.resolver().bindings() {
                    let prefix = match prefix {
                        quick_xml::name::PrefixDeclaration::Default => String::new(),
                        quick_xml::name::PrefixDeclaration::Named(prefix) => {
                            std::str::from_utf8(prefix)?.to_owned()
                        }
                    };
                    let uri = quick_xml::escape::unescape(std::str::from_utf8(uri.as_ref())?)
                        .map_err(|error| OxmlError::InvalidValue(error.to_string()))?
                        .into_owned();
                    if prefix != "xml" && !(prefix.is_empty() && uri.is_empty()) {
                        scope.insert(prefix, uri);
                    }
                }
                let index = nodes.len();
                let end = reader.buffer_position() as usize;
                nodes.push(ResidualNode {
                    name: (namespace, local),
                    known,
                    text,
                    attributes,
                    scope,
                    range: before..end,
                    opening_end: end,
                    items: Vec::new(),
                });
                if let Some(parent) = parent {
                    nodes[parent].items.push(ResidualItem::Node(index));
                }
                if !empty {
                    stack.push(index);
                }
            }
            Event::End(_) => {
                let index = stack
                    .pop()
                    .ok_or_else(|| OxmlError::InvalidValue("unbalanced residual source".into()))?;
                nodes[index].range.end = reader.buffer_position() as usize;
            }
            Event::Eof => break,
            event => {
                if let Some(index) = stack.last().copied() {
                    let modeled = nodes[index].text
                        && matches!(
                            event,
                            Event::Text(_) | Event::CData(_) | Event::GeneralRef(_)
                        );
                    let raw = source[before..reader.buffer_position() as usize].to_vec();
                    nodes[index]
                        .items
                        .push(ResidualItem::Event { raw, modeled });
                }
            }
        }
    }
    if !stack.is_empty() {
        return Err(OxmlError::InvalidValue("unclosed residual source".into()));
    }
    let mut output = Vec::new();
    if !nodes.is_empty() {
        residual_emit_children(0, &nodes, &source, &[], &mut output);
    }
    Ok(output)
}

fn residual_namespace(namespace: quick_xml::name::ResolveResult<'_>) -> Result<Option<String>> {
    match namespace {
        quick_xml::name::ResolveResult::Unbound => Ok(None),
        quick_xml::name::ResolveResult::Bound(uri) => Ok(Some(
            quick_xml::escape::unescape(std::str::from_utf8(uri.as_ref())?)
                .map_err(|error| OxmlError::InvalidValue(error.to_string()))?
                .into_owned(),
        )),
        quick_xml::name::ResolveResult::Unknown(prefix) => Ok(Some(format!("unbound:{prefix:?}"))),
    }
}

fn residual_revision_root(local: &[u8]) -> bool {
    matches!(
        local,
        b"ins"
            | b"del"
            | b"moveFrom"
            | b"moveTo"
            | b"pPrChange"
            | b"rPrChange"
            | b"tblPrChange"
            | b"sectPrChange"
            | b"tblGridChange"
            | b"numberingChange"
    )
}

fn residual_owned_attributes(
    parent: &[u8],
    local: &[u8],
    root: bool,
) -> Option<&'static [&'static [u8]]> {
    // Revision subtrees remain exact source, independently of modeled formatting.
    if residual_revision_root(local) {
        return None;
    }
    if root
        && matches!(
            local,
            b"p" | b"r"
                | b"tbl"
                | b"tr"
                | b"tc"
                | b"sdt"
                | b"pPr"
                | b"rPr"
                | b"tblPr"
                | b"trPr"
                | b"tcPr"
                | b"sdtPr"
                | b"sectPr"
        )
    {
        return Some(&[]);
    }
    crate::paragraph_properties::source_modeled_attributes(parent, local)
        .or_else(|| crate::run_properties::source_modeled_attributes(parent, local))
        .or_else(|| crate::table::source_modeled_attributes(parent, local))
        .or_else(|| crate::document::source_modeled_attributes(parent, local))
        .or_else(|| crate::content_control::source_modeled_attributes(parent, local))
        .or_else(|| crate::text::source_modeled_attributes(parent, local))
        .or_else(|| match (parent, local) {
            (b"p", b"pPr" | b"r" | b"sdt")
            | (b"r", b"rPr")
            | (b"tbl", b"tblPr" | b"tblGrid" | b"tr" | b"sdt")
            | (b"tr", b"tblPrEx" | b"trPr" | b"tc" | b"sdt")
            | (b"tc", b"tcPr" | b"p" | b"tbl" | b"sdt")
            | (b"sdt", b"sdtPr" | b"sdtContent")
            | (b"sdtContent", b"p" | b"r" | b"tbl" | b"tr" | b"tc" | b"sdt")
            | (b"hyperlink", b"r" | b"sdt") => Some(&[]),
            (b"numPr", b"ilvl" | b"numId") => Some(&[b"val"]),
            (b"tabs", b"tab") => Some(crate::numbering::property_attributes(
                crate::numbering::PropertyKind::Paragraph,
                local,
            )),
            (
                b"pBdr" | b"pgBorders",
                b"top" | b"left" | b"start" | b"bottom" | b"right" | b"end" | b"between" | b"bar",
            ) => Some(source_border_attributes()),
            _ => None,
        })
}

fn residual_owned_value(parent: &[u8], local: &[u8], key: &[u8], value: &str) -> bool {
    if local == b"shd"
        && matches!(
            key,
            b"themeTint" | b"themeShade" | b"themeFillTint" | b"themeFillShade"
        )
        && parse_uchar_hex(value).is_none()
    {
        return false;
    }
    if key == b"val"
        && (matches!((parent, local), (b"rPr", b"bdr" | b"border"))
            || matches!(
                parent,
                b"pBdr" | b"pgBorders" | b"tblBorders" | b"tcBorders"
            ))
        && crate::shared::ST_Border::from_str(value).is_err()
    {
        return false;
    }
    crate::paragraph_properties::source_value_is_modeled(parent, local, key, value)
        && crate::run_properties::source_value_is_modeled(parent, local, key, value)
        && crate::table::source_value_is_modeled(parent, local, key, value)
        && crate::document::source_value_is_modeled(parent, local, key, value)
        && crate::text::source_value_is_modeled(parent, local, key, value)
}

fn residual_default_container(index: usize, nodes: &[ResidualNode]) -> bool {
    let node = &nodes[index];
    node.known && node.attributes.is_empty()
        && matches!(node.name.1.as_str(), "pPr" | "rPr" | "tblPr" | "tblPrEx" | "trPr" | "tcPr" | "numPr" | "tabs" | "pBdr" | "tblBorders" | "tcBorders" | "tcMar" | "tblCellMar" | "sdtPr" | "tblGrid")
        && node.items.iter().all(|item| matches!(item, ResidualItem::Node(child) if residual_default_container(*child, nodes)))
}

fn residual_emit_children(
    index: usize,
    nodes: &[ResidualNode],
    source: &[u8],
    path: &[String],
    output: &mut Vec<String>,
) {
    let mut occurrences = std::collections::BTreeMap::<ResidualName, usize>::new();
    let mut previous = "start".to_owned();
    let mut event_ordinal = 0usize;
    for item in &nodes[index].items {
        match item {
            ResidualItem::Node(child) => {
                let node = &nodes[*child];
                let default = residual_default_container(*child, nodes);
                let occurrence = occurrences.entry(node.name.clone()).or_default();
                let slot = format!("{:?}#{occurrence}", node.name);
                if !default {
                    *occurrence += 1;
                }
                let mut child_path = path.to_vec();
                child_path.push(slot.clone());
                if !node.known {
                    if node.name.0.as_deref() == Some(W_NS)
                        && residual_revision_root(node.name.1.as_bytes())
                    {
                        output.push(format!(
                            "revision:{child_path:?}:after:{previous}:{:?}:{:?}",
                            &source[node.range.clone()],
                            node.scope
                        ));
                    } else {
                        // Only the opaque root opening admits namespace declaration
                        // plumbing. Its actual expanded identity, all other attributes,
                        // effective context and shape remain checked. The remainder,
                        // including descendant-local declarations, is untouched source.
                        output.push(format!(
                            "opaque:{child_path:?}:after:{previous}:{:?}:{:?}:empty={}:{:?}",
                            node.attributes,
                            node.scope,
                            node.opening_end == node.range.end,
                            &source[node.opening_end..node.range.end]
                        ));
                    }
                } else {
                    if !node.attributes.is_empty() {
                        output.push(format!(
                            "attributes:{child_path:?}:{:?}:{:?}",
                            node.attributes, node.scope
                        ));
                    }
                    residual_emit_children(*child, nodes, source, &child_path, output);
                }
                if !default {
                    previous = slot;
                }
            }
            ResidualItem::Event { raw, modeled } if !modeled => {
                output.push(format!(
                    "event:{path:?}:after:{previous}:#{event_ordinal}:{raw:?}:{:?}",
                    nodes[index].scope
                ));
                event_ordinal += 1;
            }
            _ => {}
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::shared::ST_Jc;

    fn residual(xml: &str) -> Vec<String> {
        protected_source_residual(
            xml.as_bytes(),
            &[
                ("w".to_owned(), W_NS.to_owned()),
                ("v".to_owned(), "urn:producer".to_owned()),
            ],
        )
        .unwrap()
    }

    #[test]
    fn protected_residual_opaque_root_namespace_closure_checks_only_root_plumbing() {
        for (prefix, uri, local) in [("v", "urn:producer", "opaque"), ("q", W_NS, "future")] {
            let root = format!(r#"<{prefix}:{local} {prefix}:note="a&amp;b"/>"#);
            let source = format!(r#"<w:pPr xmlns:{prefix}="{uri}">{root}</w:pPr>"#);
            let moved = format!(
                r#"<w:pPr><{prefix}:{local} {prefix}:note="a&amp;b" xmlns:{prefix}="{uri}"/></w:pPr>"#
            );
            assert_eq!(residual(&source), residual(&moved), "{prefix}");
            for changed in [
                moved.replace(uri, "urn:changed"),
                moved.replace("a&amp;b", "a&#38;b"),
                moved.replace("a&amp;b", "changed"),
            ] {
                assert_ne!(residual(&source), residual(&changed));
            }
        }
        let inner = r#"<!--same--><?opaque same?><![CDATA[same]]><v:child xmlns:v="urn:shadow" v:note="v:meaning"/>"#;
        let source = format!(
            r#"<w:pPr xmlns:v="urn:producer"><v:opaque v:note="same">{inner}</v:opaque><v:opaque v:note="second"/></w:pPr>"#
        );
        let moved = format!(
            r#"<w:pPr><v:opaque v:note="same" xmlns:v="urn:producer">{inner}</v:opaque><v:opaque v:note="second" xmlns:v="urn:producer"/></w:pPr>"#
        );
        assert_eq!(residual(&source), residual(&moved));
        for changed in [
            moved.replace("<!--same-->", "<!--changed-->"),
            moved.replace("<?opaque same?>", "<?opaque changed?>"),
            moved.replace("CDATA[same]", "CDATA[changed]"),
            moved.replace("urn:shadow", "urn:other"),
            moved.replace("v:meaning", "v:other"),
            moved.replace(r#"<v:opaque v:note="second" xmlns:v="urn:producer"/>"#, ""),
            moved.replace(inner, ""),
        ] {
            assert_ne!(residual(&source), residual(&changed));
        }
        let reversed = format!(
            r#"<w:pPr xmlns:v="urn:producer"><v:opaque v:note="second"/><v:opaque v:note="same">{inner}</v:opaque></w:pPr>"#
        );
        assert_ne!(residual(&source), residual(&reversed));
        let duplicate = source.replace("</w:pPr>", r#"<v:opaque v:note="second"/></w:pPr>"#);
        assert_ne!(residual(&source), residual(&duplicate));
        assert_ne!(
            residual(&source),
            residual(&source.replace("<!--same--><?opaque same?>", "<?opaque same?><!--same-->"))
        );
        assert_ne!(
            residual("<w:pPr><v:opaque/></w:pPr>"),
            residual("<w:pPr><v:opaque></v:opaque></w:pPr>")
        );
        // A revision's declaration bytes are separately protected, even when
        // that declaration is redundant in the effective root scope.
        let revision =
            r#"<w:p><w:ins w:id="7" w:author="A"><w:r><w:t>old</w:t></w:r></w:ins></w:p>"#;
        assert_ne!(
            residual(revision),
            residual(&revision.replace("<w:ins ", &format!("<w:ins xmlns:w=\"{W_NS}\" ")))
        );
    }

    #[test]
    fn protected_residual_excludes_actual_modeled_values_across_owner_families() {
        let source = concat!(
            r#"<w:tbl><w:tblPr><w:tblW w:w="00120" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="00120"/></w:tblGrid><w:tr><w:trPr><w:trHeight w:val="00120" w:hRule="atLeast"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="00120" w:type="dxa"/></w:tcPr>"#,
            r#"<w:p><w:pPr><w:spacing w:before="00120"/><w:rPr><w:b w:val="true"/></w:rPr></w:pPr><w:r><w:rPr><w:i w:val="on"/></w:rPr><w:t v:note="held">Anchor.</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#,
        );
        let canonical = source
            .replace("00120", "120")
            .replace(r#" w:val="true""#, "")
            .replace(r#" w:val="on""#, "");
        assert_eq!(residual(source), residual(&canonical));
        let section = r#"<w:sectPr><w:pgSz w:w="00120" w:h="00120"/><w:pgMar w:top="00120"/><w:cols w:num="02"><w:col w:w="00120"/></w:cols></w:sectPr>"#;
        assert_eq!(
            residual(section),
            residual(&section.replace("00120", "120").replace("02", "2"))
        );
        let control = r#"<w:sdt><w:sdtPr><w:id w:val="0007"/><w:tag w:val="same"/></w:sdtPr><w:sdtContent><w:r><w:t>text</w:t></w:r></w:sdtContent></w:sdt>"#;
        assert_eq!(residual(control), residual(&control.replace("0007", "7")));
    }

    #[test]
    fn protected_residual_keeps_mixed_unknown_and_unsupported_values() {
        let mixed = r#"<w:p><w:pPr v:note="root"><w:spacing w:before="00120" w:unknown="word" v:note="producer"/><w:rPr><w:b w:val="true" w:unknown="word" v:note="producer"/></w:rPr></w:pPr><w:r><w:t v:note="content">Anchor.</w:t></w:r></w:p>"#;
        let canonical = mixed
            .replace("00120", "120")
            .replace(r#" w:val="true""#, "");
        assert!(!residual(mixed).is_empty());
        assert_eq!(residual(mixed), residual(&canonical));
        for (before, after) in [
            (r#"w:unknown="word""#, r#"w:unknown="changed""#),
            (r#"v:note="producer""#, r#"v:note="changed""#),
            (r#"v:note="content""#, r#"v:note="changed""#),
        ] {
            assert_ne!(residual(mixed), residual(&canonical.replace(before, after)));
        }
        let unsupported = r#"<w:pPr><w:jc w:val="futureA"/></w:pPr>"#;
        assert!(!residual(unsupported).is_empty());
        assert_ne!(
            residual(unsupported),
            residual(&unsupported.replace("futureA", "futureB"))
        );
        let partial_raw = r#"<w:sectPr><w:docGrid w:linePitch="futureA"/></w:sectPr>"#;
        assert!(!residual(partial_raw).is_empty());
        assert_ne!(
            residual(partial_raw),
            residual(&partial_raw.replace("futureA", "futureB"))
        );
        for source in [
            r#"<w:rPr><w:bdr w:val="futureA"/></w:rPr>"#,
            r#"<w:rPr><w:bdr w:themeTint="futureA"/></w:rPr>"#,
            r#"<w:rPr><w:shd w:themeTint="futureA"/></w:rPr>"#,
            r#"<w:sectPr><w:headerReference w:type="futureA"/></w:sectPr>"#,
            r#"<w:sectPr><w:cols w:equalWidth="futureA"/></w:sectPr>"#,
            r#"<w:r><w:br w:type="futureA"/></w:r>"#,
            r#"<w:r><w:footnoteReference w:id="futureA"/></w:r>"#,
            r#"<w:r><w:t xml:space="futureA">text</w:t></w:r>"#,
            r#"<w:p><w:hyperlink w:history="futureA"/></w:p>"#,
            r#"<w:r><w:fldChar w:fldCharType="begin" w:dirty="futureA"/></w:r>"#,
        ] {
            assert!(
                !residual(source).is_empty(),
                "omitted source value: {source}"
            );
            assert_ne!(
                residual(source),
                residual(&source.replace("futureA", "futureB")),
                "omitted changed source value: {source}"
            );
        }
        let encoded = r#"<w:rPr v:note="a&amp;b"/>"#;
        assert_ne!(
            residual(encoded),
            residual(&encoded.replace("&amp;", "&#38;"))
        );
        assert_ne!(
            residual(r#"<w:rPr><v:b w:val="true"/></w:rPr>"#),
            residual("<w:rPr><v:b/></w:rPr>")
        );
    }

    #[test]
    fn protected_residual_retains_full_events_context_and_multiplicity() {
        let source = r#"<w:pPr><!--one--><?opaque one?><![CDATA[opaque]]>text<w:jc w:val="left"/><v:opaque/><v:opaque/></w:pPr>"#;
        for changed in [
            source.replace("one", "two"),
            source.replace("opaque]]", "changed]]"),
            source.replace("text", "changed"),
            source.replacen("<v:opaque/>", "", 1),
            source.replace("<!--one--><?opaque one?>", "<?opaque one?><!--one-->"),
        ] {
            assert_ne!(residual(source), residual(&changed));
        }
        assert_ne!(
            residual("<w:pPr><!--same--><w:jc w:val=\"left\"/></w:pPr>"),
            residual("<w:pPr><w:jc w:val=\"left\"/><!--same--></w:pPr>")
        );
        let inherited = r#"<w:pPr v:note="v:meaning"/>"#;
        let declared = r#"<w:pPr xmlns:v="urn:producer" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" v:note="v:meaning"/>"#;
        assert_eq!(residual(inherited), residual(declared));
        assert_ne!(
            residual(declared),
            residual(&declared.replace("urn:producer", "urn:changed"))
        );
    }

    #[test]
    fn protected_residual_default_omission_keeps_actual_residual_association() {
        let before = r#"<w:p><w:pPr/><w:r><w:rPr/><w:t v:note="kept">Anchor.</w:t></w:r></w:p>"#;
        let after = r#"<w:p><w:r><w:t v:note="kept">Anchor.</w:t></w:r></w:p>"#;
        assert_eq!(residual(before), residual(after));
        assert_eq!(
            residual(&before.replace("<w:rPr/>", "<w:rPr></w:rPr>")),
            residual(after)
        );
        assert_ne!(
            residual(&before.replace("<w:rPr/>", r#"<w:rPr v:note="metadata"/>"#)),
            residual(after)
        );
        assert_ne!(
            residual(&before.replace("<w:rPr/>", "<w:rPr><!--opaque--></w:rPr>")),
            residual(after)
        );
        let duplicates = r#"<w:r><w:rPr/><w:rPr v:note="metadata"/><w:t>text</w:t></w:r>"#;
        assert_eq!(
            residual(duplicates),
            residual(&duplicates.replacen("<w:rPr/>", "", 1))
        );
    }

    #[test]
    fn integer_measurements_round_a_decimal_half_away_from_zero() {
        for (value, expected) in [
            ("4320", 4320),
            ("+7", 7),
            ("4320.0", 4320),
            ("2210.0000000000005", 2210),
            ("4319.999999999999", 4320),
            ("720.5", 721),
            ("-226.99999999999977", -227),
            ("-720.5", -721),
            ("-0.4", 0),
            ("0012.50", 13),
            ("2147483646.5", i32::MAX),
            ("-2147483647.5", i32::MIN),
        ] {
            let parsed: i32 = parse_integer_measurement(b"w:ind", b"w:left", value).unwrap();
            assert_eq!(parsed, expected, "{value}");
        }
        let size: u32 = parse_integer_measurement(b"w:sz", b"w:val", "22.0").unwrap();
        assert_eq!(size, 22);
    }

    #[test]
    fn a_refused_integer_measurement_names_element_attribute_and_value() {
        for value in [
            "", "-", "NaN", "inf", "1e3", ".5", "1.", "1.2.3", " 1", "1 ", "0x10", "--1",
        ] {
            let error = parse_integer_measurement::<i32>(b"w:gridCol", b"w:w", value)
                .unwrap_err()
                .to_string();
            assert_eq!(
                error,
                format!("invalid value: w:gridCol/@w:w {value:?} is not an integer measurement")
            );
        }
        for value in [
            "2147483648",
            "2147483647.5",
            "2147483647.0000000000000000000000000000001",
            "-2147483648.1",
            "99999999999999999999.0",
        ] {
            let error = parse_integer_measurement::<i32>(b"w:pgSz", b"w:w", value)
                .unwrap_err()
                .to_string();
            assert_eq!(
                error,
                format!("invalid value: w:pgSz/@w:w {value:?} is out of range")
            );
        }
        assert!(parse_integer_measurement::<u32>(b"w:sz", b"w:val", "-0.5").is_err());
        assert!(parse_integer_measurement::<u32>(b"w:sz", b"w:val", "-1").is_err());
    }

    fn try_parse_ppr(xml: &str) -> Result<CT_PPr> {
        let full = format!("<w:pPr>{xml}</w:pPr>");
        let mut reader = Reader::from_str(&full);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if matches_local_name(e.name().as_ref(), b"pPr") => break,
                _ => {}
            }
            buf.clear();
        }
        CT_PPr::from_xml(&mut reader)
    }

    fn parse_ppr(xml: &str) -> CT_PPr {
        try_parse_ppr(xml).unwrap()
    }

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
    fn shading_keeps_every_theme_attribute_through_a_round_trip() {
        let rpr = parse_rpr(concat!(
            r#"<w:shd xmlns:x="urn:producer" w:val="pct20" w:color="4472C4" w:themeColor="accent1""#,
            r#" w:themeTint="66" w:themeShade="BF" w:fill="ED7D31" w:themeFill="accent2""#,
            r#" w:themeFillTint="33" w:themeFillShade="80" x:kept="shd"/>"#,
        ));
        let shading = rpr.shading.as_deref().unwrap();
        assert_eq!(shading.val, "pct20");
        assert_eq!(shading.color.as_deref(), Some("4472C4"));
        assert_eq!(shading.theme_color.as_deref(), Some("accent1"));
        assert_eq!(shading.theme_tint, Some(0x66));
        assert_eq!(shading.theme_shade, Some(0xBF));
        assert_eq!(shading.fill.as_deref(), Some("ED7D31"));
        assert_eq!(shading.theme_fill.as_deref(), Some("accent2"));
        assert_eq!(shading.theme_fill_tint, Some(0x33));
        assert_eq!(shading.theme_fill_shade, Some(0x80));
        // The namespace declaration is retained alongside the foreign
        // attribute, so the prefix stays bound on the way out.
        assert_eq!(
            shading.extra_attributes,
            vec![
                ("xmlns:x".to_owned(), "urn:producer".to_owned()),
                ("x:kept".to_owned(), "shd".to_owned()),
            ]
        );

        let mut output = Vec::new();
        rpr.to_xml(&mut Writer::new(&mut output)).unwrap();
        let output = String::from_utf8(output).unwrap();
        for retained in [
            r#"w:themeColor="accent1""#,
            r#"w:themeTint="66""#,
            r#"w:themeShade="BF""#,
            r#"w:themeFill="accent2""#,
            r#"w:themeFillTint="33""#,
            r#"w:themeFillShade="80""#,
            r#"x:kept="shd""#,
        ] {
            assert!(
                output.contains(retained),
                "{retained} missing from {output}"
            );
        }

        let reparsed = parse_rpr(
            output
                .strip_prefix("<w:rPr>")
                .unwrap()
                .strip_suffix("</w:rPr>")
                .unwrap(),
        );
        assert_eq!(reparsed.shading, rpr.shading);
    }

    #[test]
    fn a_shading_theme_byte_outside_the_hex_convention_is_retained_verbatim() {
        let rpr = parse_rpr(r#"<w:shd w:val="clear" w:themeTint="not-hex"/>"#);
        let shading = rpr.shading.as_deref().unwrap();
        assert_eq!(shading.theme_tint, None);
        assert_eq!(
            shading.extra_attributes,
            vec![("w:themeTint".to_owned(), "not-hex".to_owned())]
        );

        let mut output = Vec::new();
        rpr.to_xml(&mut Writer::new(&mut output)).unwrap();
        assert!(
            String::from_utf8(output)
                .unwrap()
                .contains(r#"w:themeTint="not-hex""#)
        );
    }

    #[test]
    fn word_bidi_and_run_rtl_parse_write_parse_without_raw_loss() {
        let ppr = parse_ppr(
            r#"<x:before xmlns:x="urn:producer"/><w:bidi/><x:after xmlns:x="urn:producer"/>"#,
        );
        assert_eq!(ppr.bidi, Some(true));

        let mut paragraph_output = Vec::new();
        ppr.to_xml(&mut Writer::new(&mut paragraph_output)).unwrap();
        let paragraph_output = String::from_utf8(paragraph_output).unwrap();
        assert!(
            paragraph_output.find("x:before").unwrap() < paragraph_output.find("w:bidi").unwrap()
        );
        assert!(
            paragraph_output.find("w:bidi").unwrap() < paragraph_output.find("x:after").unwrap()
        );

        let rpr = parse_rpr(
            r#"<x:before xmlns:x="urn:producer"/><w:rtl/><x:rtl xmlns:x="urn:foreign"/><x:after xmlns:x="urn:producer"/>"#,
        );
        assert_eq!(rpr.rtl, Some(true));
        let mut run_output = Vec::new();
        rpr.to_xml(&mut Writer::new(&mut run_output)).unwrap();
        let run_output = String::from_utf8(run_output).unwrap();
        assert!(run_output.contains("<w:rtl/>"));
        assert!(run_output.contains("<x:rtl xmlns:x=\"urn:foreign\"/>"));

        let reparsed_ppr = parse_ppr(
            paragraph_output
                .strip_prefix("<w:pPr>")
                .unwrap()
                .strip_suffix("</w:pPr>")
                .unwrap(),
        );
        let reparsed_rpr = parse_rpr(
            run_output
                .strip_prefix("<w:rPr>")
                .unwrap()
                .strip_suffix("</w:rPr>")
                .unwrap(),
        );
        assert_eq!(reparsed_ppr.bidi, Some(true));
        assert_eq!(reparsed_rpr.rtl, Some(true));
    }

    #[test]
    fn word_direction_toggles_preserve_unsupported_attributes() {
        let ppr = parse_ppr(r#"<w:bidi xmlns:x="urn:producer" w:val="0" x:flag="paragraph"/>"#);
        let rpr = parse_rpr(r#"<w:rtl xmlns:x="urn:producer" w:val="1" x:flag="run"></w:rtl>"#);
        assert_eq!(ppr.bidi, Some(false));
        assert_eq!(rpr.rtl, Some(true));

        let mut paragraph_output = Vec::new();
        ppr.to_xml(&mut Writer::new(&mut paragraph_output)).unwrap();
        let paragraph_output = String::from_utf8(paragraph_output).unwrap();
        assert!(paragraph_output.contains(r#"xmlns:x="urn:producer""#));
        assert!(paragraph_output.contains(r#"x:flag="paragraph""#));

        let mut run_output = Vec::new();
        rpr.to_xml(&mut Writer::new(&mut run_output)).unwrap();
        let run_output = String::from_utf8(run_output).unwrap();
        assert!(run_output.contains(r#"xmlns:x="urn:producer""#));
        assert!(run_output.contains(r#"x:flag="run""#));
    }

    #[test]
    fn malformed_direction_duplicates_keep_their_relative_occurrence() {
        let ppr =
            parse_ppr(r#"<w:bidi/><w:bidi><w:producerExtension w:val="paragraph"/></w:bidi>"#);
        let rpr = parse_rpr(r#"<w:rtl/><w:rtl><w:producerExtension w:val="run"/></w:rtl>"#);

        let mut paragraph_output = Vec::new();
        ppr.to_xml(&mut Writer::new(&mut paragraph_output)).unwrap();
        let paragraph_output = String::from_utf8(paragraph_output).unwrap();
        assert!(
            paragraph_output.find("<w:bidi/>").unwrap()
                < paragraph_output.find("producerExtension").unwrap()
        );

        let mut run_output = Vec::new();
        rpr.to_xml(&mut Writer::new(&mut run_output)).unwrap();
        let run_output = String::from_utf8(run_output).unwrap();
        assert!(
            run_output.find("<w:rtl/>").unwrap() < run_output.find("producerExtension").unwrap()
        );

        let reparsed_ppr = parse_ppr(
            paragraph_output
                .strip_prefix("<w:pPr>")
                .unwrap()
                .strip_suffix("</w:pPr>")
                .unwrap(),
        );
        let reparsed_rpr = parse_rpr(
            run_output
                .strip_prefix("<w:rPr>")
                .unwrap()
                .strip_suffix("</w:rPr>")
                .unwrap(),
        );
        let mut repeated_ppr = Vec::new();
        reparsed_ppr
            .to_xml(&mut Writer::new(&mut repeated_ppr))
            .unwrap();
        let mut repeated_rpr = Vec::new();
        reparsed_rpr
            .to_xml(&mut Writer::new(&mut repeated_rpr))
            .unwrap();
        assert_eq!(repeated_ppr, paragraph_output.as_bytes());
        assert_eq!(repeated_rpr, run_output.as_bytes());
    }

    #[test]
    fn interleaved_direction_duplicates_keep_their_relative_occurrence() {
        let ppr = parse_ppr(
            r#"<w:bidi xmlns:x="urn:first" x:first="paragraph"/><w:bidi><w:producerExtension w:val="paragraph"/></w:bidi><w:bidi xmlns:y="urn:last" y:last="paragraph"/>"#,
        );
        let rpr = parse_rpr(
            r#"<w:rtl xmlns:x="urn:first" x:first="run"/><w:rtl><w:producerExtension w:val="run"/></w:rtl><w:rtl xmlns:y="urn:last" y:last="run"/>"#,
        );

        let mut paragraph_output = Vec::new();
        ppr.to_xml(&mut Writer::new(&mut paragraph_output)).unwrap();
        let paragraph_output = String::from_utf8(paragraph_output).unwrap();
        assert!(
            paragraph_output.find("x:first").unwrap()
                < paragraph_output.find("producerExtension").unwrap(),
            "{paragraph_output}"
        );
        assert!(
            paragraph_output.find("producerExtension").unwrap()
                < paragraph_output.find("y:last").unwrap(),
            "{paragraph_output}"
        );

        let mut run_output = Vec::new();
        rpr.to_xml(&mut Writer::new(&mut run_output)).unwrap();
        let run_output = String::from_utf8(run_output).unwrap();
        assert!(
            run_output.find("x:first").unwrap() < run_output.find("producerExtension").unwrap(),
            "{run_output}"
        );
        assert!(
            run_output.find("producerExtension").unwrap() < run_output.find("y:last").unwrap(),
            "{run_output}"
        );

        let reparsed_ppr = parse_ppr(
            paragraph_output
                .strip_prefix("<w:pPr>")
                .unwrap()
                .strip_suffix("</w:pPr>")
                .unwrap(),
        );
        let reparsed_rpr = parse_rpr(
            run_output
                .strip_prefix("<w:rPr>")
                .unwrap()
                .strip_suffix("</w:rPr>")
                .unwrap(),
        );
        let mut repeated_ppr = Vec::new();
        reparsed_ppr
            .to_xml(&mut Writer::new(&mut repeated_ppr))
            .unwrap();
        let mut repeated_rpr = Vec::new();
        reparsed_rpr
            .to_xml(&mut Writer::new(&mut repeated_rpr))
            .unwrap();
        assert_eq!(repeated_ppr, paragraph_output.as_bytes());
        assert_eq!(repeated_rpr, run_output.as_bytes());
    }

    #[test]
    fn canonical_final_direction_toggle_stays_after_an_interleaved_malformed_occurrence() {
        let ppr = parse_ppr(
            r#"<w:bidi xmlns:x="urn:first" x:first="paragraph"/><w:bidi><w:producerExtension w:val="paragraph"/></w:bidi><w:bidi/>"#,
        );
        let rpr = parse_rpr(
            r#"<w:rtl xmlns:x="urn:first" x:first="run"/><w:rtl><w:producerExtension w:val="run"/></w:rtl><w:rtl/>"#,
        );

        let mut paragraph_output = Vec::new();
        ppr.to_xml(&mut Writer::new(&mut paragraph_output)).unwrap();
        let paragraph_output = String::from_utf8(paragraph_output).unwrap();
        assert_eq!(paragraph_output.matches("<w:bidi").count(), 3);
        assert!(
            paragraph_output.find("x:first").unwrap()
                < paragraph_output.find("producerExtension").unwrap(),
            "{paragraph_output}"
        );
        assert!(
            paragraph_output.find("producerExtension").unwrap()
                < paragraph_output.rfind("<w:bidi").unwrap(),
            "{paragraph_output}"
        );

        let mut run_output = Vec::new();
        rpr.to_xml(&mut Writer::new(&mut run_output)).unwrap();
        let run_output = String::from_utf8(run_output).unwrap();
        assert_eq!(run_output.matches("<w:rtl").count(), 3);
        assert!(
            run_output.find("x:first").unwrap() < run_output.find("producerExtension").unwrap(),
            "{run_output}"
        );
        assert!(
            run_output.find("producerExtension").unwrap() < run_output.rfind("<w:rtl").unwrap(),
            "{run_output}"
        );

        let reparsed_ppr = parse_ppr(
            paragraph_output
                .strip_prefix("<w:pPr>")
                .unwrap()
                .strip_suffix("</w:pPr>")
                .unwrap(),
        );
        let reparsed_rpr = parse_rpr(
            run_output
                .strip_prefix("<w:rPr>")
                .unwrap()
                .strip_suffix("</w:rPr>")
                .unwrap(),
        );
        let mut repeated_ppr = Vec::new();
        reparsed_ppr
            .to_xml(&mut Writer::new(&mut repeated_ppr))
            .unwrap();
        let mut repeated_rpr = Vec::new();
        reparsed_rpr
            .to_xml(&mut Writer::new(&mut repeated_rpr))
            .unwrap();
        assert_eq!(repeated_ppr, paragraph_output.as_bytes());
        assert_eq!(repeated_rpr, run_output.as_bytes());
    }

    #[test]
    fn duplicate_valid_direction_toggles_keep_each_occurrences_attributes_and_order() {
        let ppr = parse_ppr(
            r#"<w:bidi xmlns:x="urn:first" w:val="0" x:first="paragraph"/><w:bidi xmlns:y="urn:second" w:val="1" y:second="paragraph"/>"#,
        );
        let rpr = parse_rpr(
            r#"<w:rtl xmlns:x="urn:first" w:val="0" x:first="run"/><w:rtl xmlns:y="urn:second" w:val="1" y:second="run"/>"#,
        );

        let mut paragraph_output = Vec::new();
        ppr.to_xml(&mut Writer::new(&mut paragraph_output)).unwrap();
        let paragraph_output = String::from_utf8(paragraph_output).unwrap();
        assert_eq!(paragraph_output.matches("<w:bidi").count(), 2);
        assert!(
            paragraph_output.find("x:first").unwrap() < paragraph_output.find("y:second").unwrap(),
            "{paragraph_output}"
        );
        assert!(paragraph_output.contains(r#"x:first="paragraph""#));
        assert!(paragraph_output.contains(r#"y:second="paragraph""#));

        let mut run_output = Vec::new();
        rpr.to_xml(&mut Writer::new(&mut run_output)).unwrap();
        let run_output = String::from_utf8(run_output).unwrap();
        assert_eq!(run_output.matches("<w:rtl").count(), 2);
        assert!(
            run_output.find("x:first").unwrap() < run_output.find("y:second").unwrap(),
            "{run_output}"
        );
        assert!(run_output.contains(r#"x:first="run""#));
        assert!(run_output.contains(r#"y:second="run""#));

        let reparsed_ppr = parse_ppr(
            paragraph_output
                .strip_prefix("<w:pPr>")
                .unwrap()
                .strip_suffix("</w:pPr>")
                .unwrap(),
        );
        let reparsed_rpr = parse_rpr(
            run_output
                .strip_prefix("<w:rPr>")
                .unwrap()
                .strip_suffix("</w:rPr>")
                .unwrap(),
        );
        assert_eq!(reparsed_ppr.bidi, Some(true));
        assert_eq!(reparsed_rpr.rtl, Some(true));

        let mut repeated_ppr = Vec::new();
        reparsed_ppr
            .to_xml(&mut Writer::new(&mut repeated_ppr))
            .unwrap();
        let mut repeated_rpr = Vec::new();
        reparsed_rpr
            .to_xml(&mut Writer::new(&mut repeated_rpr))
            .unwrap();
        assert_eq!(repeated_ppr, paragraph_output.as_bytes());
        assert_eq!(repeated_rpr, run_output.as_bytes());

        let mut cleared_ppr = reparsed_ppr;
        cleared_ppr.bidi = None;
        cleared_ppr.jc = Some(ST_Jc::Left);
        let mut cleared_ppr_output = Vec::new();
        cleared_ppr
            .to_xml(&mut Writer::new(&mut cleared_ppr_output))
            .unwrap();
        let cleared_ppr_output = String::from_utf8(cleared_ppr_output).unwrap();
        assert!(
            !cleared_ppr_output.contains("<w:bidi"),
            "{cleared_ppr_output}"
        );
        assert_eq!(
            parse_ppr(
                cleared_ppr_output
                    .strip_prefix("<w:pPr>")
                    .unwrap()
                    .strip_suffix("</w:pPr>")
                    .unwrap(),
            )
            .bidi,
            None
        );

        let mut cleared_rpr = reparsed_rpr;
        cleared_rpr.rtl = None;
        cleared_rpr.bold = Some(true);
        let mut cleared_rpr_output = Vec::new();
        cleared_rpr
            .to_xml(&mut Writer::new(&mut cleared_rpr_output))
            .unwrap();
        let cleared_rpr_output = String::from_utf8(cleared_rpr_output).unwrap();
        assert!(
            !cleared_rpr_output.contains("<w:rtl"),
            "{cleared_rpr_output}"
        );
        assert_eq!(
            parse_rpr(
                cleared_rpr_output
                    .strip_prefix("<w:rPr>")
                    .unwrap()
                    .strip_suffix("</w:rPr>")
                    .unwrap(),
            )
            .rtl,
            None
        );
    }
}
