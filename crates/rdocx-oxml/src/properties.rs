//! Paragraph properties (`CT_PPr`) and run properties (`CT_RPr`).

use std::borrow::Cow;

use quick_xml::escape::escape;
use quick_xml::events::attributes::Attribute;
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::name::QName;
use quick_xml::{Reader, Writer, XmlVersion};

use crate::borders::{CT_PBdr, CT_Tabs};
use crate::document::CT_SectPr;
use crate::error::{OxmlError, Result};
use crate::namespace::{W_NS, matches_local_name};
use crate::numbering::{
    local_namespace_overrides, merged_owner_bindings, namespace_binding, namespace_bindings,
    parse_scoped_rpr_with_owner_bindings, typed_leaf_requires_raw, typed_leaf_value,
    write_typed_leaf_raw,
};
use crate::raw_xml::{capture_element, capture_empty_element};
use crate::revision::CT_Revision;
use crate::shared::{ST_HighlightColor, ST_Jc, ST_OnOff, ST_Underline};
use crate::units::{HalfPoint, Twips};

/// `CT_Shd` — Shading/background fill.
#[derive(Debug, Clone, PartialEq)]
pub struct CT_Shd {
    /// Shading pattern (e.g. "clear", "solid", "horzStripe")
    pub val: String,
    /// Foreground color hex
    pub color: Option<String>,
    /// Background fill color hex
    pub fill: Option<String>,
}

impl CT_Shd {
    pub fn from_xml_attrs(e: &BytesStart) -> Result<Self> {
        let mut val = "clear".to_string();
        let mut color = None;
        let mut fill = None;

        for attr in e.attributes() {
            let attr = attr?;
            let key = attr.key.as_ref();
            let v = std::str::from_utf8(&attr.value)?;
            if matches_local_name(key, b"val") {
                val = v.to_string();
            } else if matches_local_name(key, b"color") {
                color = Some(v.to_string());
            } else if matches_local_name(key, b"fill") {
                fill = Some(v.to_string());
            }
        }

        Ok(CT_Shd { val, color, fill })
    }

    pub(crate) fn from_xml_attrs_with_prefixes(
        e: &BytesStart,
        word_prefixes: &[String],
    ) -> Result<Self> {
        let mut val = "clear".to_string();
        let mut color = None;
        let mut fill = None;

        for attr in e.attributes() {
            let attr = attr?;
            let key = attr.key.as_ref();
            let v = std::str::from_utf8(&attr.value)?;
            if is_word_attribute(key, b"val", word_prefixes) {
                val = v.to_string();
            } else if is_word_attribute(key, b"color", word_prefixes) {
                color = Some(v.to_string());
            } else if is_word_attribute(key, b"fill", word_prefixes) {
                fill = Some(v.to_string());
            }
        }

        Ok(CT_Shd { val, color, fill })
    }

    pub fn write_xml<W: std::io::Write>(&self, writer: &mut Writer<W>, tag: &str) -> Result<()> {
        let mut e = BytesStart::new(tag);
        e.push_attribute(("w:val", self.val.as_str()));
        if let Some(ref c) = self.color {
            e.push_attribute(("w:color", c.as_str()));
        }
        if let Some(ref f) = self.fill {
            e.push_attribute(("w:fill", f.as_str()));
        }
        writer.write_event(Event::Empty(e))?;
        Ok(())
    }
}

/// `CT_PPr` — Paragraph properties.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct CT_PPr {
    /// Paragraph style ID (pStyle)
    pub style_id: Option<String>,
    /// Justification (jc)
    pub jc: Option<ST_Jc>,
    /// Space before paragraph in twips (spacing/@w:before)
    pub space_before: Option<Twips>,
    /// Space after paragraph in twips (spacing/@w:after)
    pub space_after: Option<Twips>,
    /// Line spacing in twips (spacing/@w:line)
    pub line_spacing: Option<Twips>,
    /// Line spacing rule (spacing/@w:lineRule): "auto", "exact", "atLeast"
    pub line_rule: Option<String>,
    /// Space before auto-spacing (spacing/@w:beforeAutospacing)
    pub before_autospacing: Option<bool>,
    /// Space after auto-spacing (spacing/@w:afterAutospacing)
    pub after_autospacing: Option<bool>,
    /// Left indentation in twips (ind/@w:left)
    pub ind_left: Option<Twips>,
    /// Right indentation in twips (ind/@w:right)
    pub ind_right: Option<Twips>,
    /// Logical leading-edge indentation in twips (ind/@w:start).
    pub ind_start: Option<Twips>,
    /// Logical trailing-edge indentation in twips (ind/@w:end).
    pub ind_end: Option<Twips>,
    /// First line indent in twips (ind/@w:firstLine)
    pub ind_first_line: Option<Twips>,
    /// Hanging indent in twips (ind/@w:hanging)
    pub ind_hanging: Option<Twips>,
    /// Keep with next paragraph (keepNext)
    pub keep_next: Option<bool>,
    /// Keep lines together (keepLines)
    pub keep_lines: Option<bool>,
    /// Page break before (pageBreakBefore)
    pub page_break_before: Option<bool>,
    /// Widow/orphan control (widowControl)
    pub widow_control: Option<bool>,
    /// Suppress auto-hyphens (suppressAutoHyphens)
    pub suppress_auto_hyphens: Option<bool>,
    /// Paragraph base direction (bidi).
    pub bidi: Option<bool>,
    /// Outline level 0-8 (outlineLvl)
    pub outline_lvl: Option<u32>,
    /// Paragraph borders (pBdr)
    pub borders: Option<CT_PBdr>,
    /// Tab stops (tabs)
    pub tabs: Option<CT_Tabs>,
    /// Paragraph shading (shd)
    pub shading: Option<CT_Shd>,
    /// Run properties for the paragraph mark (rPr)
    pub rpr: Option<CT_RPr>,
    /// Numbering level (numPr/ilvl)
    pub num_ilvl: Option<u32>,
    /// Original value and XML carrier for `w:numPr/w:ilvl`.
    #[doc(hidden)]
    pub num_ilvl_raw: Option<(Option<u32>, Vec<u8>, Vec<String>)>,
    /// Numbering ID (numPr/numId)
    pub num_id: Option<u32>,
    /// Original value and XML carrier for `w:numPr/w:numId`.
    #[doc(hidden)]
    pub num_id_raw: Option<(Option<u32>, Vec<u8>, Vec<String>)>,
    /// Unmodelled attributes and namespace declarations from `w:numPr`.
    #[doc(hidden)]
    pub num_pr_extra_attributes: Vec<(String, String)>,
    /// Unmodelled children retained at their schema-child boundary in `w:numPr`.
    #[doc(hidden)]
    pub num_pr_extra_xml: Vec<(usize, usize, Vec<u8>)>,
    /// Section properties embedded in paragraph (section break)
    pub sect_pr: Option<CT_SectPr>,
    /// Tracked insertion of the numbering properties.
    pub numbering_revision: Option<CT_Revision>,
    /// Malformed or foreign numbering markers retained inside `w:numPr`.
    pub numbering_revision_xml: Vec<Vec<u8>>,
    /// Schema boundaries and source ordinals for retained numbering markers.
    #[doc(hidden)]
    pub numbering_revision_xml_positions: Vec<(usize, usize)>,
    /// Schema boundary and source ordinal for the typed numbering marker.
    #[doc(hidden)]
    pub numbering_revision_position: Option<(usize, usize)>,
    /// Prior paragraph properties from the schema-final `w:pPrChange`.
    pub change: Option<CT_Revision>,
    /// Malformed tracked-change elements retained without a typed projection.
    pub revision_xml: Vec<Vec<u8>>,
    /// Schema slots and occurrences for retained raw children.
    #[doc(hidden)]
    pub revision_xml_positions: Vec<(u8, usize)>,
}

const PPR_STYLE_SLOT: u8 = 0;
const PPR_KEEP_NEXT_SLOT: u8 = 1;
const PPR_KEEP_LINES_SLOT: u8 = 2;
const PPR_PAGE_BREAK_SLOT: u8 = 3;
const PPR_WIDOW_SLOT: u8 = 5;
const PPR_NUM_SLOT: u8 = 6;
const PPR_BORDER_SLOT: u8 = 8;
const PPR_SHADING_SLOT: u8 = 9;
const PPR_TABS_SLOT: u8 = 10;
const PPR_SUPPRESS_HYPHENS_SLOT: u8 = 11;
const PPR_BIDI_SLOT: u8 = 18;
const PPR_SPACING_SLOT: u8 = 21;
const PPR_INDENT_SLOT: u8 = 22;
const PPR_JUSTIFICATION_SLOT: u8 = 26;
const PPR_OUTLINE_SLOT: u8 = 30;
const PPR_RUN_PROPERTIES_SLOT: u8 = 33;
const PPR_SECTION_SLOT: u8 = 34;
const PPR_CHANGE_SLOT: u8 = 35;
const PPR_END_SLOT: u8 = 36;
const RAW_MODELED_ATTRIBUTES_FLAG: usize = 1 << (usize::BITS - 1);
const RAW_MODELED_DUPLICATE_FLAG: usize = 1 << (usize::BITS - 2);
const RAW_MODELED_FLAGS: usize = RAW_MODELED_ATTRIBUTES_FLAG | RAW_MODELED_DUPLICATE_FLAG;

fn parse_line_spacing(value: &str) -> Result<Twips> {
    let integer_error = match value.parse::<i32>() {
        Ok(value) => return Ok(Twips(value)),
        Err(error) => error,
    };

    // Some producers write fractional values even though w:line is integral.
    // Accept only plain decimals and normalize them to the nearest twip.
    let unsigned = value
        .strip_prefix('-')
        .or_else(|| value.strip_prefix('+'))
        .unwrap_or(value);
    let Some((whole, fraction)) = unsigned.split_once('.') else {
        return Err(integer_error.into());
    };

    if whole.is_empty()
        || fraction.is_empty()
        || !whole.bytes().all(|byte| byte.is_ascii_digit())
        || !fraction.bytes().all(|byte| byte.is_ascii_digit())
    {
        return Err(integer_error.into());
    }

    let decimal = value
        .parse::<f64>()
        .map_err(|_| OxmlError::InvalidValue(format!("invalid line spacing: {value}")))?;
    if !decimal.is_finite() || decimal < i32::MIN as f64 || decimal > i32::MAX as f64 {
        return Err(OxmlError::InvalidValue(format!(
            "line spacing is out of range: {value}"
        )));
    }

    Ok(Twips(decimal.round() as i32))
}

fn raw_occurrence(position: (u8, usize)) -> usize {
    position.1 & !RAW_MODELED_FLAGS
}

fn raw_is_modeled_attribute_carrier(position: (u8, usize)) -> bool {
    position.1 & RAW_MODELED_ATTRIBUTES_FLAG != 0
}

fn modeled_attribute_carrier_occurrence(occurrence: usize) -> usize {
    occurrence | RAW_MODELED_ATTRIBUTES_FLAG
}

fn raw_is_modeled_duplicate(position: (u8, usize)) -> bool {
    position.1 & RAW_MODELED_DUPLICATE_FLAG != 0
}

fn modeled_duplicate_occurrence(occurrence: usize) -> usize {
    occurrence | RAW_MODELED_DUPLICATE_FLAG
}

fn replay_modeled_toggle_raw(position: (u8, usize), modeled_value_is_present: bool) -> bool {
    !raw_is_modeled_attribute_carrier(position)
        && (!raw_is_modeled_duplicate(position) || modeled_value_is_present)
}

fn record_modeled_toggle_candidate(
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

fn remove_redundant_modeled_toggle_candidate(
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
const RPR_VANISH_SLOT: u8 = 16;
const RPR_COLOR_SLOT: u8 = 18;
const RPR_SPACING_SLOT: u8 = 19;
const RPR_WIDTH_SLOT: u8 = 20;
const RPR_POSITION_SLOT: u8 = 22;
const RPR_SIZE_SLOT: u8 = 23;
const RPR_SIZE_CS_SLOT: u8 = 24;
const RPR_HIGHLIGHT_SLOT: u8 = 25;
const RPR_UNDERLINE_SLOT: u8 = 26;
const RPR_SHADING_SLOT: u8 = 29;
const RPR_VERT_ALIGN_SLOT: u8 = 31;
const RPR_RTL_SLOT: u8 = 32;
const RPR_LANG_SLOT: u8 = 35;
const RPR_MARKER_SLOT: u8 = 39;
const RPR_CHANGE_SLOT: u8 = 40;
const RPR_END_SLOT: u8 = 41;

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

fn record_ppr_modeled(
    ppr: &mut CT_PPr,
    pending_raw: &mut Vec<Vec<u8>>,
    occurrences: &mut [usize],
    slot: u8,
) {
    let occurrence = occurrences[slot as usize];
    for raw in pending_raw.drain(..) {
        ppr.revision_xml.push(raw);
        ppr.revision_xml_positions.push((slot, occurrence));
    }
    occurrences[slot as usize] += 1;
}

fn record_ppr_raw_at(ppr: &mut CT_PPr, raw: Vec<u8>, slot: u8, occurrence: usize) {
    ppr.revision_xml.push(raw);
    ppr.revision_xml_positions.push((slot, occurrence));
}

fn flush_ppr_raw(ppr: &mut CT_PPr, pending_raw: &mut Vec<Vec<u8>>, slot: u8) {
    for raw in pending_raw.drain(..) {
        ppr.revision_xml.push(raw);
        ppr.revision_xml_positions.push((slot, 0));
    }
}

fn typed_num_pr_scope(
    word_prefixes: &[String],
    owner_bindings: &[(String, String)],
) -> Vec<String> {
    let mut scope = word_prefixes.to_vec();
    for binding in word_prefixes
        .iter()
        .filter(|prefix| !prefix.starts_with('\0'))
        .map(|prefix| namespace_binding(prefix, W_NS))
        .chain(
            owner_bindings
                .iter()
                .map(|(prefix, namespace)| namespace_binding(prefix, namespace)),
        )
    {
        if !scope.contains(&binding) {
            scope.push(binding);
        }
    }
    scope
}

type PreservedNumPrRoot = (
    Vec<(String, String)>,
    Vec<(String, String)>,
    Vec<(String, String)>,
);

fn preserved_num_pr_root_attributes(
    root: &BytesStart<'_>,
    word_prefixes: &[String],
    owner_bindings: &[(String, String)],
) -> Result<PreservedNumPrRoot> {
    let mut attributes = root
        .attributes()
        .map(|attribute| {
            let attribute = attribute?;
            Ok((
                std::str::from_utf8(attribute.key.as_ref())?.to_owned(),
                std::str::from_utf8(attribute.value.as_ref())?.to_owned(),
                attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, root.decoder())?
                    .into_owned(),
            ))
        })
        .collect::<Result<Vec<_>>>()?;
    let local_bindings = attributes
        .iter()
        .filter_map(|(name, _, namespace)| {
            if name == "xmlns" {
                Some((String::new(), namespace.clone()))
            } else {
                name.strip_prefix("xmlns:")
                    .map(|prefix| (prefix.to_owned(), namespace.clone()))
            }
        })
        .collect::<Vec<_>>();
    attributes.retain(|(name, _, value)| name != "xmlns:w" || value != W_NS);
    let typed_root_scope = typed_num_pr_scope(word_prefixes, owner_bindings);
    let all_bindings = namespace_bindings(&typed_root_scope);
    let root_attribute_prefixes = attributes
        .iter()
        .filter_map(|(name, _, _)| {
            let (prefix, _) = name.split_once(':')?;
            (prefix != "xmlns").then_some(prefix.to_owned())
        })
        .collect::<Vec<_>>();
    for prefix in root_attribute_prefixes {
        let declaration = format!("xmlns:{prefix}");
        if attributes.iter().any(|(name, _, _)| name == &declaration) {
            continue;
        }
        if let Some((_, namespace)) = all_bindings
            .iter()
            .find(|(candidate, _)| candidate == &prefix)
        {
            attributes.push((
                declaration,
                escape(namespace).into_owned(),
                namespace.clone(),
            ));
        }
    }
    Ok((
        attributes
            .into_iter()
            .map(|(name, raw_value, _)| (name, raw_value))
            .collect(),
        local_bindings,
        all_bindings,
    ))
}

fn num_pr_leaf_is_plain(raw: &[u8], word_prefixes: &[String]) -> Result<bool> {
    let mut reader = Reader::from_reader(raw);
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Empty(element) => {
                return Ok(!typed_leaf_requires_raw(&element, word_prefixes, false)?);
            }
            Event::Start(element) => {
                return Ok(!typed_leaf_requires_raw(&element, word_prefixes, true)?);
            }
            Event::Eof => return Ok(false),
            _ => {}
        }
        buffer.clear();
    }
}

fn write_num_pr_leaf<W: std::io::Write>(
    writer: &mut Writer<W>,
    raw: &[u8],
    word_prefixes: &[String],
    original: Option<u32>,
    current: Option<u32>,
) -> Result<()> {
    if current == original {
        if let Some(current) = current
            && num_pr_leaf_is_plain(raw, word_prefixes)?
        {
            let value = current.to_string();
            write_typed_leaf_raw(writer, raw, word_prefixes, Some(&value), "w")?;
        } else {
            writer.get_mut().write_all(raw)?;
        }
    } else if current.is_none() && num_pr_leaf_is_plain(raw, word_prefixes)? {
        return Ok(());
    } else {
        let value = current.map(|value| value.to_string());
        write_typed_leaf_raw(writer, raw, word_prefixes, value.as_deref(), "w")?;
    }
    Ok(())
}

fn raw_is_num_pr_leaf(raw: &[u8], word_prefixes: &[String], local: &[u8]) -> Result<bool> {
    let mut reader = Reader::from_reader(raw);
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Empty(element) | Event::Start(element) => {
                return Ok(is_word_element(
                    element.name().as_ref(),
                    local,
                    word_prefixes,
                ));
            }
            Event::Eof => return Ok(false),
            _ => {}
        }
        buffer.clear();
    }
}

fn write_num_pr_extras<W: std::io::Write>(
    writer: &mut Writer<W>,
    extras: &[(usize, usize, Vec<u8>)],
    boundary: usize,
    scrub_local: Option<(&[u8], &[String])>,
) -> Result<()> {
    for (_, _, raw) in extras.iter().filter(|(at, _, _)| *at == boundary) {
        if let Some((local, word_prefixes)) = scrub_local
            && raw_is_num_pr_leaf(raw, word_prefixes, local)?
        {
            if !num_pr_leaf_is_plain(raw, word_prefixes)? {
                write_typed_leaf_raw(writer, raw, word_prefixes, None, "w")?;
            }
        } else {
            writer.get_mut().write_all(raw)?;
        }
    }
    Ok(())
}

#[allow(non_snake_case)]
impl CT_PPr {
    pub fn from_xml(reader: &mut Reader<&[u8]>) -> Result<Self> {
        Self::from_xml_with_prefixes(reader, &["w".to_string()])
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
        let mut ppr = CT_PPr::default();
        let mut change_raw_index = 0usize;
        let mut numbering_change_raw_index = 0usize;
        let mut pending_raw = Vec::new();
        let mut occurrences = [0usize; PPR_END_SLOT as usize + 1];
        let mut bidi_carrier_required = false;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let name = e.name();
                    let prefixes = word_prefixes_at(e, word_prefixes)?;
                    if let Some(slot) = ppr_modeled_slot(name.as_ref(), &prefixes) {
                        record_ppr_modeled(&mut ppr, &mut pending_raw, &mut occurrences, slot);
                    }
                    if is_word_element(name.as_ref(), b"rPr", &prefixes) {
                        let raw = capture_element(reader, e)?;
                        ppr.rpr = Some(parse_scoped_rpr_with_owner_bindings(
                            &raw,
                            word_prefixes,
                            owner_bindings,
                        )?);
                    } else if is_word_element(name.as_ref(), b"numPr", &prefixes) {
                        let local_bindings = local_namespace_overrides(e, word_prefixes)?;
                        let num_pr_bindings =
                            merged_owner_bindings(owner_bindings, &local_bindings);
                        Self::parse_num_pr(
                            reader,
                            e,
                            &mut ppr,
                            &prefixes,
                            &num_pr_bindings,
                            &mut numbering_change_raw_index,
                        )?;
                    } else if is_word_element(name.as_ref(), b"pBdr", &prefixes) {
                        ppr.borders = Some(CT_PBdr::from_xml_with_prefixes(reader, &prefixes)?);
                    } else if is_word_element(name.as_ref(), b"tabs", &prefixes) {
                        ppr.tabs = Some(CT_Tabs::from_xml_with_prefixes(reader, &prefixes)?);
                    } else if is_word_element(name.as_ref(), b"sectPr", &prefixes) {
                        let local_bindings = local_namespace_overrides(e, word_prefixes)?;
                        let sect_pr_bindings =
                            merged_owner_bindings(owner_bindings, &local_bindings);
                        ppr.sect_pr = Some(CT_SectPr::from_xml_with_prefixes_and_owner_bindings(
                            reader,
                            &prefixes,
                            &sect_pr_bindings,
                        )?);
                    } else if is_word_element(name.as_ref(), b"pPrChange", &prefixes) {
                        let raw = crate::text::raw_with_external_bindings(
                            &capture_element(reader, e)?,
                            owner_bindings,
                        )?;
                        if let Some(revision) = CT_Revision::from_raw(raw.clone(), &prefixes) {
                            if let Some(previous) = ppr.change.replace(revision) {
                                ppr.revision_xml
                                    .insert(change_raw_index, previous.into_raw_xml());
                                ppr.revision_xml_positions
                                    .insert(change_raw_index, (PPR_CHANGE_SLOT, 0));
                            }
                            change_raw_index = ppr.revision_xml.len();
                        } else {
                            record_ppr_raw_at(&mut ppr, raw, PPR_CHANGE_SLOT, 0);
                        }
                    } else if matches_local_name(name.as_ref(), b"pPrChange") {
                        let raw = crate::text::raw_with_external_bindings(
                            &capture_element(reader, e)?,
                            owner_bindings,
                        )?;
                        pending_raw.push(raw);
                    } else if is_word_element(name.as_ref(), b"bidi", &prefixes) {
                        let captured = capture_element(reader, e)?;
                        let raw =
                            crate::text::raw_with_external_bindings(&captured, owner_bindings)?;
                        if toggle_element_is_explicitly_empty(&captured)? {
                            ppr.bidi = Some(parse_word_toggle(e, &prefixes)?);
                            bidi_carrier_required =
                                toggle_has_unsupported_attributes(e, &prefixes)?;
                            record_modeled_toggle_candidate(
                                &mut ppr.revision_xml,
                                &mut ppr.revision_xml_positions,
                                raw,
                                PPR_BIDI_SLOT,
                                occurrences[PPR_BIDI_SLOT as usize] - 1,
                            );
                        } else {
                            let occurrence = occurrences[PPR_BIDI_SLOT as usize] - 1;
                            record_ppr_raw_at(&mut ppr, raw, PPR_BIDI_SLOT, occurrence);
                        }
                    } else {
                        let raw = crate::text::raw_with_external_bindings(
                            &capture_element(reader, e)?,
                            owner_bindings,
                        )?;
                        if let Some(slot) = ppr_schema_slot(name.as_ref(), &prefixes) {
                            record_ppr_raw_at(&mut ppr, raw, slot, 0);
                        } else {
                            pending_raw.push(raw);
                        }
                    }
                }
                Ok(Event::Empty(ref e)) => {
                    let name = e.name();
                    let prefixes = word_prefixes_at(e, word_prefixes)?;
                    if let Some(slot) = ppr_modeled_slot(name.as_ref(), &prefixes) {
                        record_ppr_modeled(&mut ppr, &mut pending_raw, &mut occurrences, slot);
                    }
                    if is_word_element(name.as_ref(), b"pStyle", &prefixes) {
                        ppr.style_id = get_word_val_attr(e, &prefixes)?;
                    } else if is_word_element(name.as_ref(), b"jc", &prefixes) {
                        if let Some(val) = get_word_val_attr(e, &prefixes)? {
                            ppr.jc = ST_Jc::from_str(&val).ok();
                        }
                    } else if is_word_element(name.as_ref(), b"spacing", &prefixes) {
                        for attr in e.attributes() {
                            let attr = attr?;
                            let key = attr.key.as_ref();
                            let val_str = std::str::from_utf8(&attr.value)?;
                            if is_word_attribute(key, b"before", &prefixes) {
                                ppr.space_before = Some(Twips(val_str.parse()?));
                            } else if is_word_attribute(key, b"after", &prefixes) {
                                ppr.space_after = Some(Twips(val_str.parse()?));
                            } else if is_word_attribute(key, b"line", &prefixes) {
                                ppr.line_spacing = Some(parse_line_spacing(val_str)?);
                            } else if is_word_attribute(key, b"lineRule", &prefixes) {
                                ppr.line_rule = Some(val_str.to_string());
                            } else if is_word_attribute(key, b"beforeAutospacing", &prefixes) {
                                ppr.before_autospacing = Some(val_str == "1" || val_str == "true");
                            } else if is_word_attribute(key, b"afterAutospacing", &prefixes) {
                                ppr.after_autospacing = Some(val_str == "1" || val_str == "true");
                            }
                        }
                    } else if is_word_element(name.as_ref(), b"ind", &prefixes) {
                        for attr in e.attributes() {
                            let attr = attr?;
                            let key = attr.key.as_ref();
                            let val_str = std::str::from_utf8(&attr.value)?;
                            if is_word_attribute(key, b"left", &prefixes) {
                                ppr.ind_left = Some(Twips(val_str.parse()?));
                            } else if is_word_attribute(key, b"right", &prefixes) {
                                ppr.ind_right = Some(Twips(val_str.parse()?));
                            } else if is_word_attribute(key, b"start", &prefixes) {
                                ppr.ind_start = Some(Twips(val_str.parse()?));
                            } else if is_word_attribute(key, b"end", &prefixes) {
                                ppr.ind_end = Some(Twips(val_str.parse()?));
                            } else if is_word_attribute(key, b"firstLine", &prefixes) {
                                ppr.ind_first_line = Some(Twips(val_str.parse()?));
                            } else if is_word_attribute(key, b"hanging", &prefixes) {
                                ppr.ind_hanging = Some(Twips(val_str.parse()?));
                            }
                        }
                    } else if is_word_element(name.as_ref(), b"keepNext", &prefixes) {
                        ppr.keep_next = Some(parse_word_toggle(e, &prefixes)?);
                    } else if is_word_element(name.as_ref(), b"keepLines", &prefixes) {
                        ppr.keep_lines = Some(parse_word_toggle(e, &prefixes)?);
                    } else if is_word_element(name.as_ref(), b"pageBreakBefore", &prefixes) {
                        ppr.page_break_before = Some(parse_word_toggle(e, &prefixes)?);
                    } else if is_word_element(name.as_ref(), b"widowControl", &prefixes) {
                        ppr.widow_control = Some(parse_word_toggle(e, &prefixes)?);
                    } else if is_word_element(name.as_ref(), b"suppressAutoHyphens", &prefixes) {
                        ppr.suppress_auto_hyphens = Some(parse_word_toggle(e, &prefixes)?);
                    } else if is_word_element(name.as_ref(), b"bidi", &prefixes) {
                        ppr.bidi = Some(parse_word_toggle(e, &prefixes)?);
                        bidi_carrier_required = toggle_has_unsupported_attributes(e, &prefixes)?;
                        let raw = crate::text::raw_with_external_bindings(
                            &capture_empty_element(e)?,
                            owner_bindings,
                        )?;
                        record_modeled_toggle_candidate(
                            &mut ppr.revision_xml,
                            &mut ppr.revision_xml_positions,
                            raw,
                            PPR_BIDI_SLOT,
                            occurrences[PPR_BIDI_SLOT as usize] - 1,
                        );
                    } else if is_word_element(name.as_ref(), b"outlineLvl", &prefixes) {
                        if let Some(val) = get_word_val_attr(e, &prefixes)? {
                            ppr.outline_lvl = Some(val.parse()?);
                        }
                    } else if is_word_element(name.as_ref(), b"shd", &prefixes) {
                        ppr.shading = Some(CT_Shd::from_xml_attrs(e)?);
                    } else if is_word_element(name.as_ref(), b"numPr", &prefixes) {
                        ppr.num_pr_extra_attributes =
                            preserved_num_pr_root_attributes(e, &prefixes, owner_bindings)?.0;
                    } else if is_word_element(name.as_ref(), b"pPrChange", &prefixes) {
                        let raw = crate::text::raw_with_external_bindings(
                            &capture_empty_element(e)?,
                            owner_bindings,
                        )?;
                        if let Some(revision) = CT_Revision::from_raw(raw.clone(), &prefixes) {
                            if let Some(previous) = ppr.change.replace(revision) {
                                ppr.revision_xml
                                    .insert(change_raw_index, previous.into_raw_xml());
                                ppr.revision_xml_positions
                                    .insert(change_raw_index, (PPR_CHANGE_SLOT, 0));
                            }
                            change_raw_index = ppr.revision_xml.len();
                        } else {
                            record_ppr_raw_at(&mut ppr, raw, PPR_CHANGE_SLOT, 0);
                        }
                    } else if matches_local_name(name.as_ref(), b"pPrChange") {
                        pending_raw.push(crate::text::raw_with_external_bindings(
                            &capture_empty_element(e)?,
                            owner_bindings,
                        )?);
                    } else {
                        let raw = crate::text::raw_with_external_bindings(
                            &capture_empty_element(e)?,
                            owner_bindings,
                        )?;
                        if let Some(slot) = ppr_schema_slot(name.as_ref(), &prefixes) {
                            record_ppr_raw_at(&mut ppr, raw, slot, 0);
                        } else {
                            pending_raw.push(raw);
                        }
                    }
                }
                Ok(Event::End(ref e))
                    if is_word_element(e.name().as_ref(), b"pPr", word_prefixes) =>
                {
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        flush_ppr_raw(&mut ppr, &mut pending_raw, PPR_END_SLOT);
        remove_redundant_modeled_toggle_candidate(
            &mut ppr.revision_xml,
            &mut ppr.revision_xml_positions,
            PPR_BIDI_SLOT,
            bidi_carrier_required,
        );

        Ok(ppr)
    }

    fn parse_num_pr(
        reader: &mut Reader<&[u8]>,
        root: &BytesStart<'_>,
        ppr: &mut CT_PPr,
        word_prefixes: &[String],
        owner_bindings: &[(String, String)],
        change_raw_index: &mut usize,
    ) -> Result<()> {
        let (attributes, local_bindings, all_bindings) =
            preserved_num_pr_root_attributes(root, word_prefixes, owner_bindings)?;
        ppr.num_pr_extra_attributes = attributes;
        let child_external_bindings = all_bindings
            .iter()
            .filter(|binding| {
                !local_bindings.contains(binding) && !(binding.0 == "w" && binding.1 == W_NS)
            })
            .cloned()
            .collect::<Vec<_>>();
        let mut buf = Vec::new();
        let mut boundary = 0usize;
        let mut source_ordinal = 0usize;
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(ref e)) => {
                    let name = e.name();
                    let prefixes = word_prefixes_at(e, word_prefixes)?;
                    let typed_scope = typed_num_pr_scope(&prefixes, owner_bindings);
                    if is_word_element(name.as_ref(), b"ilvl", &prefixes) {
                        let parsed =
                            typed_leaf_value(e, &prefixes)?.and_then(|value| value.parse().ok());
                        let raw = crate::text::raw_with_external_bindings(
                            &capture_empty_element(e)?,
                            &child_external_bindings,
                        )?;
                        if let Some((_, previous, _)) =
                            ppr.num_ilvl_raw.replace((parsed, raw, typed_scope))
                        {
                            ppr.num_pr_extra_xml.push((0, source_ordinal, previous));
                        }
                        ppr.num_ilvl = parsed;
                        boundary = 1;
                    } else if is_word_element(name.as_ref(), b"numId", &prefixes) {
                        let parsed =
                            typed_leaf_value(e, &prefixes)?.and_then(|value| value.parse().ok());
                        let raw = crate::text::raw_with_external_bindings(
                            &capture_empty_element(e)?,
                            &child_external_bindings,
                        )?;
                        if let Some((_, previous, _)) =
                            ppr.num_id_raw.replace((parsed, raw, typed_scope))
                        {
                            ppr.num_pr_extra_xml.push((1, source_ordinal, previous));
                        }
                        ppr.num_id = parsed;
                        boundary = 2;
                    } else if is_word_element(name.as_ref(), b"numberingChange", &prefixes) {
                        ppr.num_pr_extra_xml.push((
                            3,
                            source_ordinal,
                            crate::text::raw_with_external_bindings(
                                &capture_empty_element(e)?,
                                &child_external_bindings,
                            )?,
                        ));
                        boundary = 3;
                    } else if is_word_element(name.as_ref(), b"ins", &prefixes) {
                        let raw = crate::text::raw_with_external_bindings(
                            &capture_empty_element(e)?,
                            &child_external_bindings,
                        )?;
                        if let Some(revision) = CT_Revision::from_raw(raw.clone(), &prefixes) {
                            if let Some(previous) = ppr.numbering_revision.replace(revision) {
                                ppr.numbering_revision_xml
                                    .insert(*change_raw_index, previous.into_raw_xml());
                                ppr.numbering_revision_xml_positions.insert(
                                    *change_raw_index,
                                    ppr.numbering_revision_position
                                        .unwrap_or((3, source_ordinal)),
                                );
                            }
                            *change_raw_index = ppr.numbering_revision_xml.len();
                            ppr.numbering_revision_position =
                                Some((if boundary < 4 { 3 } else { 4 }, source_ordinal));
                        } else {
                            ppr.numbering_revision_xml.push(raw);
                            ppr.numbering_revision_xml_positions
                                .push((if boundary < 4 { 3 } else { 4 }, source_ordinal));
                        }
                        boundary = 4;
                    } else if matches_local_name(name.as_ref(), b"ins") {
                        ppr.numbering_revision_xml
                            .push(crate::text::raw_with_external_bindings(
                                &capture_empty_element(e)?,
                                &child_external_bindings,
                            )?);
                        ppr.numbering_revision_xml_positions
                            .push((if boundary < 4 { 3 } else { 4 }, source_ordinal));
                        boundary = 4;
                    } else {
                        ppr.num_pr_extra_xml.push((
                            boundary,
                            source_ordinal,
                            crate::text::raw_with_external_bindings(
                                &capture_empty_element(e)?,
                                &child_external_bindings,
                            )?,
                        ));
                    }
                }
                Ok(Event::Start(ref e)) => {
                    let prefixes = word_prefixes_at(e, word_prefixes)?;
                    let typed_scope = typed_num_pr_scope(&prefixes, owner_bindings);
                    if is_word_element(e.name().as_ref(), b"ilvl", &prefixes) {
                        let parsed =
                            typed_leaf_value(e, &prefixes)?.and_then(|value| value.parse().ok());
                        let raw = crate::text::raw_with_external_bindings(
                            &capture_element(reader, e)?,
                            &child_external_bindings,
                        )?;
                        if let Some((_, previous, _)) =
                            ppr.num_ilvl_raw.replace((parsed, raw, typed_scope))
                        {
                            ppr.num_pr_extra_xml.push((0, source_ordinal, previous));
                        }
                        ppr.num_ilvl = parsed;
                        boundary = 1;
                    } else if is_word_element(e.name().as_ref(), b"numId", &prefixes) {
                        let parsed =
                            typed_leaf_value(e, &prefixes)?.and_then(|value| value.parse().ok());
                        let raw = crate::text::raw_with_external_bindings(
                            &capture_element(reader, e)?,
                            &child_external_bindings,
                        )?;
                        if let Some((_, previous, _)) =
                            ppr.num_id_raw.replace((parsed, raw, typed_scope))
                        {
                            ppr.num_pr_extra_xml.push((1, source_ordinal, previous));
                        }
                        ppr.num_id = parsed;
                        boundary = 2;
                    } else if is_word_element(e.name().as_ref(), b"numberingChange", &prefixes) {
                        ppr.num_pr_extra_xml.push((
                            3,
                            source_ordinal,
                            crate::text::raw_with_external_bindings(
                                &capture_element(reader, e)?,
                                &child_external_bindings,
                            )?,
                        ));
                        boundary = 3;
                    } else if is_word_element(e.name().as_ref(), b"ins", &prefixes) {
                        let raw = crate::text::raw_with_external_bindings(
                            &capture_element(reader, e)?,
                            &child_external_bindings,
                        )?;
                        if let Some(revision) = CT_Revision::from_raw(raw.clone(), &prefixes) {
                            if let Some(previous) = ppr.numbering_revision.replace(revision) {
                                ppr.numbering_revision_xml
                                    .insert(*change_raw_index, previous.into_raw_xml());
                                ppr.numbering_revision_xml_positions.insert(
                                    *change_raw_index,
                                    ppr.numbering_revision_position
                                        .unwrap_or((3, source_ordinal)),
                                );
                            }
                            *change_raw_index = ppr.numbering_revision_xml.len();
                            ppr.numbering_revision_position =
                                Some((if boundary < 4 { 3 } else { 4 }, source_ordinal));
                        } else {
                            ppr.numbering_revision_xml.push(raw);
                            ppr.numbering_revision_xml_positions
                                .push((if boundary < 4 { 3 } else { 4 }, source_ordinal));
                        }
                        boundary = 4;
                    } else if matches_local_name(e.name().as_ref(), b"ins") {
                        ppr.numbering_revision_xml
                            .push(crate::text::raw_with_external_bindings(
                                &capture_element(reader, e)?,
                                &child_external_bindings,
                            )?);
                        ppr.numbering_revision_xml_positions
                            .push((if boundary < 4 { 3 } else { 4 }, source_ordinal));
                        boundary = 4;
                    } else {
                        ppr.num_pr_extra_xml.push((
                            boundary,
                            source_ordinal,
                            crate::text::raw_with_external_bindings(
                                &capture_element(reader, e)?,
                                &child_external_bindings,
                            )?,
                        ));
                    }
                }
                Ok(Event::End(ref e))
                    if is_word_element(e.name().as_ref(), b"numPr", word_prefixes) =>
                {
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            source_ordinal += 1;
            buf.clear();
        }
        Ok(())
    }

    pub fn to_xml<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        if self.is_empty() {
            return Ok(());
        }

        if !self.revision_xml.is_empty()
            && self.revision_xml_positions.len() == self.revision_xml.len()
        {
            let mut modeled = self.clone();
            modeled.revision_xml.clear();
            modeled.revision_xml_positions.clear();
            let mut generated = Writer::new(Vec::new());
            modeled.to_xml(&mut generated)?;
            return write_ppr_with_positioned_raw(writer, &generated.into_inner(), self);
        }

        let mut buf = itoa::Buffer::new();
        writer.write_event(Event::Start(BytesStart::new("w:pPr")))?;

        if let Some(ref style_id) = self.style_id {
            let mut e = BytesStart::new("w:pStyle");
            e.push_attribute(("w:val", style_id.as_str()));
            writer.write_event(Event::Empty(e))?;
        }

        if let Some(keep_next) = self.keep_next {
            write_toggle(writer, "w:keepNext", keep_next)?;
        }
        if let Some(keep_lines) = self.keep_lines {
            write_toggle(writer, "w:keepLines", keep_lines)?;
        }
        if let Some(page_break) = self.page_break_before {
            write_toggle(writer, "w:pageBreakBefore", page_break)?;
        }
        if let Some(widow) = self.widow_control {
            write_toggle(writer, "w:widowControl", widow)?;
        }

        // numPr
        let ilvl_raw_retained = match &self.num_ilvl_raw {
            Some((original, raw, prefixes)) => {
                self.num_ilvl == *original
                    || self.num_ilvl.is_some()
                    || !num_pr_leaf_is_plain(raw, prefixes)?
            }
            None => false,
        };
        let num_id_raw_retained = match &self.num_id_raw {
            Some((original, raw, prefixes)) => {
                self.num_id == *original
                    || self.num_id.is_some()
                    || !num_pr_leaf_is_plain(raw, prefixes)?
            }
            None => false,
        };
        if self.num_id.is_some()
            || self.num_ilvl.is_some()
            || num_id_raw_retained
            || ilvl_raw_retained
            || !self.num_pr_extra_attributes.is_empty()
            || !self.num_pr_extra_xml.is_empty()
            || self.numbering_revision.is_some()
            || !self.numbering_revision_xml.is_empty()
        {
            let mut num_pr = BytesStart::new("w:numPr");
            for (name, value) in &self.num_pr_extra_attributes {
                if name == "xmlns:w" && value != W_NS {
                    return Err(crate::error::OxmlError::InvalidValue(
                        "w:numPr shadows the Word namespace".to_owned(),
                    ));
                }
                num_pr.push_attribute(Attribute {
                    key: QName(name.as_bytes()),
                    value: Cow::Borrowed(value.as_bytes()),
                });
            }
            writer.write_event(Event::Start(num_pr))?;
            let ilvl_scrub = self
                .num_ilvl_raw
                .as_ref()
                .and_then(|(original, _, prefixes)| {
                    (original.is_some() && self.num_ilvl.is_none())
                        .then_some((b"ilvl".as_slice(), prefixes.as_slice()))
                });
            write_num_pr_extras(writer, &self.num_pr_extra_xml, 0, ilvl_scrub)?;
            if let Some((original, raw, prefixes)) = &self.num_ilvl_raw {
                write_num_pr_leaf(writer, raw, prefixes, *original, self.num_ilvl)?;
            } else if let Some(ilvl) = self.num_ilvl {
                let mut e = BytesStart::new("w:ilvl");
                e.push_attribute(("w:val", buf.format(ilvl)));
                writer.write_event(Event::Empty(e))?;
            }
            let num_id_scrub = self
                .num_id_raw
                .as_ref()
                .and_then(|(original, _, prefixes)| {
                    (original.is_some() && self.num_id.is_none())
                        .then_some((b"numId".as_slice(), prefixes.as_slice()))
                });
            write_num_pr_extras(writer, &self.num_pr_extra_xml, 1, num_id_scrub)?;
            if let Some((original, raw, prefixes)) = &self.num_id_raw {
                write_num_pr_leaf(writer, raw, prefixes, *original, self.num_id)?;
            } else if let Some(num_id) = self.num_id {
                let mut e = BytesStart::new("w:numId");
                e.push_attribute(("w:val", buf.format(num_id)));
                writer.write_event(Event::Empty(e))?;
            }
            write_num_pr_extras(writer, &self.num_pr_extra_xml, 2, None)?;
            enum Tail<'a> {
                Raw(&'a [u8]),
                Extra(&'a [u8]),
                Revision(&'a CT_Revision),
            }
            let mut tail = Vec::<((usize, usize, usize), Tail<'_>)>::new();
            for (index, raw) in self.numbering_revision_xml.iter().enumerate() {
                let position = self
                    .numbering_revision_xml_positions
                    .get(index)
                    .copied()
                    .unwrap_or((3, index));
                tail.push(((position.0, position.1, index), Tail::Raw(raw)));
            }
            for (index, (boundary, ordinal, raw)) in self
                .num_pr_extra_xml
                .iter()
                .enumerate()
                .filter(|(_, (boundary, _, _))| *boundary >= 3)
            {
                tail.push(((*boundary, *ordinal, index), Tail::Extra(raw)));
            }
            if let Some(revision) = &self.numbering_revision {
                let position = self.numbering_revision_position.unwrap_or((3, usize::MAX));
                tail.push((
                    (position.0, position.1, usize::MAX),
                    Tail::Revision(revision),
                ));
            }
            tail.sort_by_key(|(position, _)| *position);
            for (_, child) in tail {
                match child {
                    Tail::Raw(raw) | Tail::Extra(raw) => writer.get_mut().write_all(raw)?,
                    Tail::Revision(revision) => revision.write_xml(writer)?,
                }
            }
            writer.write_event(Event::End(BytesEnd::new("w:numPr")))?;
        }

        // pBdr
        if let Some(ref borders) = self.borders
            && !borders.is_empty()
        {
            borders.to_xml(writer)?;
        }

        // shd
        if let Some(ref shd) = self.shading {
            shd.write_xml(writer, "w:shd")?;
        }

        // tabs
        if let Some(ref tabs) = self.tabs {
            tabs.to_xml(writer)?;
        }

        if let Some(suppress) = self.suppress_auto_hyphens {
            write_toggle(writer, "w:suppressAutoHyphens", suppress)?;
        }

        if let Some(bidi) = self.bidi {
            write_toggle(writer, "w:bidi", bidi)?;
        }

        // spacing
        if self.space_before.is_some()
            || self.space_after.is_some()
            || self.line_spacing.is_some()
            || self.before_autospacing.is_some()
            || self.after_autospacing.is_some()
        {
            let mut e = BytesStart::new("w:spacing");
            if let Some(before) = self.space_before {
                e.push_attribute(("w:before", buf.format(before.0)));
            }
            if let Some(after) = self.space_after {
                e.push_attribute(("w:after", buf.format(after.0)));
            }
            if let Some(line) = self.line_spacing {
                e.push_attribute(("w:line", buf.format(line.0)));
            }
            if let Some(ref rule) = self.line_rule {
                e.push_attribute(("w:lineRule", rule.as_str()));
            }
            if let Some(ba) = self.before_autospacing {
                e.push_attribute(("w:beforeAutospacing", if ba { "1" } else { "0" }));
            }
            if let Some(aa) = self.after_autospacing {
                e.push_attribute(("w:afterAutospacing", if aa { "1" } else { "0" }));
            }
            writer.write_event(Event::Empty(e))?;
        }

        // ind
        if self.ind_left.is_some()
            || self.ind_right.is_some()
            || self.ind_start.is_some()
            || self.ind_end.is_some()
            || self.ind_first_line.is_some()
            || self.ind_hanging.is_some()
        {
            let mut e = BytesStart::new("w:ind");
            if let Some(left) = self.ind_left {
                e.push_attribute(("w:left", buf.format(left.0)));
            }
            if let Some(right) = self.ind_right {
                e.push_attribute(("w:right", buf.format(right.0)));
            }
            if let Some(start) = self.ind_start {
                e.push_attribute(("w:start", buf.format(start.0)));
            }
            if let Some(end) = self.ind_end {
                e.push_attribute(("w:end", buf.format(end.0)));
            }
            if let Some(fl) = self.ind_first_line {
                e.push_attribute(("w:firstLine", buf.format(fl.0)));
            }
            if let Some(hang) = self.ind_hanging {
                e.push_attribute(("w:hanging", buf.format(hang.0)));
            }
            writer.write_event(Event::Empty(e))?;
        }

        if let Some(jc) = self.jc {
            let mut e = BytesStart::new("w:jc");
            e.push_attribute(("w:val", jc.to_str()));
            writer.write_event(Event::Empty(e))?;
        }

        if let Some(lvl) = self.outline_lvl {
            let mut e = BytesStart::new("w:outlineLvl");
            e.push_attribute(("w:val", buf.format(lvl)));
            writer.write_event(Event::Empty(e))?;
        }

        if let Some(ref rpr) = self.rpr {
            rpr.to_xml(writer)?;
        }

        if let Some(ref sect) = self.sect_pr {
            sect.to_xml(writer)?;
        }

        for raw in &self.revision_xml {
            writer.get_mut().write_all(raw)?;
        }
        if let Some(change) = &self.change {
            change.write_xml(writer)?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:pPr")))?;
        Ok(())
    }

    fn is_empty(&self) -> bool {
        self.style_id.is_none()
            && self.jc.is_none()
            && self.space_before.is_none()
            && self.space_after.is_none()
            && self.line_spacing.is_none()
            && self.before_autospacing.is_none()
            && self.after_autospacing.is_none()
            && self.ind_left.is_none()
            && self.ind_right.is_none()
            && self.ind_start.is_none()
            && self.ind_end.is_none()
            && self.ind_first_line.is_none()
            && self.ind_hanging.is_none()
            && self.keep_next.is_none()
            && self.keep_lines.is_none()
            && self.page_break_before.is_none()
            && self.widow_control.is_none()
            && self.suppress_auto_hyphens.is_none()
            && self.bidi.is_none()
            && self.outline_lvl.is_none()
            && self.borders.is_none()
            && self.tabs.is_none()
            && self.shading.is_none()
            && self.rpr.is_none()
            && self.num_id.is_none()
            && self.num_ilvl.is_none()
            && self.num_id_raw.is_none()
            && self.num_ilvl_raw.is_none()
            && self.num_pr_extra_attributes.is_empty()
            && self.num_pr_extra_xml.is_empty()
            && self.sect_pr.is_none()
            && self.numbering_revision.is_none()
            && self.numbering_revision_xml.is_empty()
            && self.change.is_none()
            && self.revision_xml.is_empty()
    }

    /// Merge another CT_PPr into this one (non-None fields override).
    /// Used for style inheritance.
    pub fn merge_from(&mut self, other: &CT_PPr) {
        if other.style_id.is_some() {
            self.style_id = other.style_id.clone();
        }
        if other.jc.is_some() {
            self.jc = other.jc;
        }
        if other.space_before.is_some() {
            self.space_before = other.space_before;
        }
        if other.space_after.is_some() {
            self.space_after = other.space_after;
        }
        if other.line_spacing.is_some() {
            self.line_spacing = other.line_spacing;
        }
        if other.line_rule.is_some() {
            self.line_rule = other.line_rule.clone();
        }
        if other.before_autospacing.is_some() {
            self.before_autospacing = other.before_autospacing;
        }
        if other.after_autospacing.is_some() {
            self.after_autospacing = other.after_autospacing;
        }
        if other.ind_left.is_some() {
            self.ind_left = other.ind_left;
        }
        if other.ind_right.is_some() {
            self.ind_right = other.ind_right;
        }
        if other.ind_start.is_some() {
            self.ind_start = other.ind_start;
        }
        if other.ind_end.is_some() {
            self.ind_end = other.ind_end;
        }
        if other.ind_first_line.is_some() {
            self.ind_first_line = other.ind_first_line;
        }
        if other.ind_hanging.is_some() {
            self.ind_hanging = other.ind_hanging;
        }
        if other.keep_next.is_some() {
            self.keep_next = other.keep_next;
        }
        if other.keep_lines.is_some() {
            self.keep_lines = other.keep_lines;
        }
        if other.page_break_before.is_some() {
            self.page_break_before = other.page_break_before;
        }
        if other.widow_control.is_some() {
            self.widow_control = other.widow_control;
        }
        if other.suppress_auto_hyphens.is_some() {
            self.suppress_auto_hyphens = other.suppress_auto_hyphens;
        }
        if other.bidi.is_some() {
            self.bidi = other.bidi;
        }
        if other.outline_lvl.is_some() {
            self.outline_lvl = other.outline_lvl;
        }
        if other.borders.is_some() {
            self.borders = other.borders.clone();
        }
        if other.tabs.is_some() {
            self.tabs = other.tabs.clone();
        }
        if other.shading.is_some() {
            self.shading = other.shading.clone();
        }
        if other.num_ilvl.is_some() {
            self.num_ilvl = other.num_ilvl;
        }
        if other.num_id.is_some() {
            self.num_id = other.num_id;
        }
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
    pub shading: Option<CT_Shd>,
    /// Vanish/hidden text (vanish)
    pub vanish: Option<bool>,
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
    pub change: Option<CT_Revision>,
    /// Foreign and unmodelled children retained without a typed projection.
    pub revision_xml: Vec<Vec<u8>>,
    /// Schema slots and occurrences for retained raw children.
    #[doc(hidden)]
    pub revision_xml_positions: Vec<(u8, usize)>,
}

#[allow(non_snake_case)]
impl CT_RPr {
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
        let mut rpr = CT_RPr::default();
        let mut change_raw_index = 0usize;
        let mut pending_raw = Vec::new();
        let mut occurrences = [0usize; RPR_END_SLOT as usize + 1];
        let mut rtl_carrier_required = false;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(ref e)) => {
                    let name = e.name();
                    let prefixes = word_prefixes_at(e, word_prefixes)?;
                    if is_word_element(name.as_ref(), b"rStyle", &prefixes) {
                        record_rpr_modeled(
                            &mut rpr,
                            &mut pending_raw,
                            &mut occurrences,
                            RPR_STYLE_SLOT,
                        );
                        rpr.style_id = get_word_val_attr(e, &prefixes)?;
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
                            }
                        }
                    } else if is_word_element(name.as_ref(), b"b", &prefixes) {
                        record_rpr_modeled(
                            &mut rpr,
                            &mut pending_raw,
                            &mut occurrences,
                            RPR_BOLD_SLOT,
                        );
                        rpr.bold = Some(parse_word_toggle(e, &prefixes)?);
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
                    } else if is_word_element(name.as_ref(), b"sz", &prefixes) {
                        record_rpr_modeled(
                            &mut rpr,
                            &mut pending_raw,
                            &mut occurrences,
                            RPR_SIZE_SLOT,
                        );
                        if let Some(val) = get_word_val_attr(e, &prefixes)? {
                            rpr.sz = Some(HalfPoint(val.parse()?));
                        }
                    } else if is_word_element(name.as_ref(), b"szCs", &prefixes) {
                        record_rpr_modeled(
                            &mut rpr,
                            &mut pending_raw,
                            &mut occurrences,
                            RPR_SIZE_CS_SLOT,
                        );
                        if let Some(val) = get_word_val_attr(e, &prefixes)? {
                            rpr.sz_cs = Some(HalfPoint(val.parse()?));
                        }
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
                    } else if is_word_element(name.as_ref(), b"spacing", &prefixes) {
                        record_rpr_modeled(
                            &mut rpr,
                            &mut pending_raw,
                            &mut occurrences,
                            RPR_SPACING_SLOT,
                        );
                        if let Some(val) = get_word_val_attr(e, &prefixes)? {
                            rpr.spacing = Some(Twips(val.parse()?));
                        }
                    } else if is_word_element(name.as_ref(), b"w", &prefixes) {
                        record_rpr_modeled(
                            &mut rpr,
                            &mut pending_raw,
                            &mut occurrences,
                            RPR_WIDTH_SLOT,
                        );
                        if let Some(val) = get_word_val_attr(e, &prefixes)? {
                            rpr.width_scale = Some(val.parse()?);
                        }
                    } else if is_word_element(name.as_ref(), b"position", &prefixes) {
                        record_rpr_modeled(
                            &mut rpr,
                            &mut pending_raw,
                            &mut occurrences,
                            RPR_POSITION_SLOT,
                        );
                        if let Some(val) = get_word_val_attr(e, &prefixes)? {
                            rpr.position = Some(val.parse()?);
                        }
                    } else if is_word_element(name.as_ref(), b"shd", &prefixes) {
                        record_rpr_modeled(
                            &mut rpr,
                            &mut pending_raw,
                            &mut occurrences,
                            RPR_SHADING_SLOT,
                        );
                        rpr.shading = Some(CT_Shd::from_xml_attrs(e)?);
                    } else if is_word_element(name.as_ref(), b"vanish", &prefixes) {
                        record_rpr_modeled(
                            &mut rpr,
                            &mut pending_raw,
                            &mut occurrences,
                            RPR_VANISH_SLOT,
                        );
                        rpr.vanish = Some(parse_word_toggle(e, &prefixes)?);
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
                            rpr.change = Some(revision);
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
                            rpr.change = Some(revision);
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
        if self.is_empty() {
            return Ok(());
        }

        if !self.revision_xml.is_empty()
            && self.revision_xml_positions.len() == self.revision_xml.len()
        {
            let mut modeled = self.clone();
            modeled.revision_xml.clear();
            modeled.revision_xml_positions.clear();
            let mut generated = Writer::new(Vec::new());
            if modeled.is_empty() {
                generated.write_event(Event::Start(BytesStart::new("w:rPr")))?;
                generated.write_event(Event::End(BytesEnd::new("w:rPr")))?;
            } else {
                modeled.to_xml_with_word_override(&mut generated, foreign_word_namespace)?;
            }
            return write_rpr_with_positioned_raw(
                writer,
                &generated.into_inner(),
                self,
                foreign_word_namespace,
            );
        }

        let mut buf = itoa::Buffer::new();
        writer.write_event(Event::Start(BytesStart::new("w:rPr")))?;

        if let Some(ref style_id) = self.style_id {
            let mut e = BytesStart::new("w:rStyle");
            e.push_attribute(("w:val", style_id.as_str()));
            writer.write_event(Event::Empty(e))?;
        }

        // rFonts
        if self.font_ascii.is_some()
            || self.font_hansi.is_some()
            || self.font_east_asia.is_some()
            || self.font_cs.is_some()
            || self.font_ascii_theme.is_some()
            || self.font_hansi_theme.is_some()
        {
            let mut e = BytesStart::new("w:rFonts");
            if let Some(ref f) = self.font_ascii {
                e.push_attribute(("w:ascii", f.as_str()));
            }
            if let Some(ref f) = self.font_hansi {
                e.push_attribute(("w:hAnsi", f.as_str()));
            }
            if let Some(ref f) = self.font_east_asia {
                e.push_attribute(("w:eastAsia", f.as_str()));
            }
            if let Some(ref f) = self.font_cs {
                e.push_attribute(("w:cs", f.as_str()));
            }
            if let Some(ref f) = self.font_ascii_theme {
                e.push_attribute(("w:asciiTheme", f.as_str()));
            }
            if let Some(ref f) = self.font_hansi_theme {
                e.push_attribute(("w:hAnsiTheme", f.as_str()));
            }
            writer.write_event(Event::Empty(e))?;
        }

        if let Some(bold) = self.bold {
            write_toggle(writer, "w:b", bold)?;
        }
        if let Some(bold_cs) = self.bold_cs {
            write_toggle(writer, "w:bCs", bold_cs)?;
        }
        if let Some(italic) = self.italic {
            write_toggle(writer, "w:i", italic)?;
        }
        if let Some(italic_cs) = self.italic_cs {
            write_toggle(writer, "w:iCs", italic_cs)?;
        }
        if let Some(caps) = self.caps {
            write_toggle(writer, "w:caps", caps)?;
        }
        if let Some(small_caps) = self.small_caps {
            write_toggle(writer, "w:smallCaps", small_caps)?;
        }
        if let Some(strike) = self.strike {
            write_toggle(writer, "w:strike", strike)?;
        }
        if let Some(dstrike) = self.dstrike {
            write_toggle(writer, "w:dstrike", dstrike)?;
        }
        if let Some(vanish) = self.vanish {
            write_toggle(writer, "w:vanish", vanish)?;
        }

        if let Some(ref color) = self.color {
            let mut e = BytesStart::new("w:color");
            e.push_attribute(("w:val", color.as_str()));
            if let Some(ref tc) = self.color_theme {
                e.push_attribute(("w:themeColor", tc.as_str()));
            }
            writer.write_event(Event::Empty(e))?;
        }

        if let Some(ref spacing) = self.spacing {
            let mut e = BytesStart::new("w:spacing");
            e.push_attribute(("w:val", buf.format(spacing.0)));
            writer.write_event(Event::Empty(e))?;
        }
        if let Some(ws) = self.width_scale {
            let mut e = BytesStart::new("w:w");
            e.push_attribute(("w:val", buf.format(ws)));
            writer.write_event(Event::Empty(e))?;
        }
        if let Some(pos) = self.position {
            let mut e = BytesStart::new("w:position");
            e.push_attribute(("w:val", buf.format(pos)));
            writer.write_event(Event::Empty(e))?;
        }

        if let Some(ref sz) = self.sz {
            let mut e = BytesStart::new("w:sz");
            e.push_attribute(("w:val", buf.format(sz.0)));
            writer.write_event(Event::Empty(e))?;
        }
        if let Some(ref sz_cs) = self.sz_cs {
            let mut e = BytesStart::new("w:szCs");
            e.push_attribute(("w:val", buf.format(sz_cs.0)));
            writer.write_event(Event::Empty(e))?;
        }

        if let Some(ref highlight) = self.highlight {
            let mut e = BytesStart::new("w:highlight");
            e.push_attribute(("w:val", highlight.to_str()));
            writer.write_event(Event::Empty(e))?;
        }

        if let Some(underline) = self.underline {
            let mut e = BytesStart::new("w:u");
            e.push_attribute(("w:val", underline.to_str()));
            writer.write_event(Event::Empty(e))?;
        }

        if let Some(ref shd) = self.shading {
            shd.write_xml(writer, "w:shd")?;
        }

        if let Some(ref vert_align) = self.vert_align {
            let mut e = BytesStart::new("w:vertAlign");
            e.push_attribute(("w:val", vert_align.as_str()));
            writer.write_event(Event::Empty(e))?;
        }

        if let Some(rtl) = self.rtl {
            write_toggle(writer, "w:rtl", rtl)?;
        }

        if self.language.is_some()
            || self.language_east_asia.is_some()
            || self.language_bidi.is_some()
            || !self.language_extra_attributes.is_empty()
        {
            let mut e = BytesStart::new("w:lang");
            if let Some(language) = &self.language {
                e.push_attribute(("w:val", language.as_str()));
            }
            if let Some(language) = &self.language_east_asia {
                e.push_attribute(("w:eastAsia", language.as_str()));
            }
            if let Some(language) = &self.language_bidi {
                e.push_attribute(("w:bidi", language.as_str()));
            }
            for (name, value) in &self.language_extra_attributes {
                e.push_attribute((name.as_str(), value.as_str()));
            }
            writer.write_event(Event::Empty(e))?;
        }

        for revision in &self.revision_markers {
            revision.write_xml_with_word_override(writer, foreign_word_namespace)?;
        }
        for raw in &self.revision_xml {
            crate::text::write_raw_with_word_override(writer, raw, foreign_word_namespace)?;
        }
        if let Some(change) = &self.change {
            change.write_xml_with_word_override(writer, foreign_word_namespace)?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:rPr")))?;
        Ok(())
    }

    fn is_empty(&self) -> bool {
        self.style_id.is_none()
            && self.font_ascii.is_none()
            && self.font_hansi.is_none()
            && self.font_east_asia.is_none()
            && self.font_cs.is_none()
            && self.font_ascii_theme.is_none()
            && self.font_hansi_theme.is_none()
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
            && self.highlight.is_none()
            && self.caps.is_none()
            && self.small_caps.is_none()
            && self.vert_align.is_none()
            && self.spacing.is_none()
            && self.width_scale.is_none()
            && self.position.is_none()
            && self.shading.is_none()
            && self.vanish.is_none()
            && self.rtl.is_none()
            && self.language.is_none()
            && self.language_east_asia.is_none()
            && self.language_bidi.is_none()
            && self.language_extra_attributes.is_empty()
            && self.revision_markers.is_empty()
            && self.change.is_none()
            && self.revision_xml.is_empty()
    }

    /// Merge another CT_RPr into this one (non-None fields override).
    /// Used for style inheritance.
    pub fn merge_from(&mut self, other: &CT_RPr) {
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

fn language_element_is_explicitly_empty(xml: &[u8]) -> Result<bool> {
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

fn toggle_element_is_explicitly_empty(xml: &[u8]) -> Result<bool> {
    language_element_is_explicitly_empty(xml)
}

fn toggle_has_unsupported_attributes(
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

fn append_modeled_toggle_attributes(
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

fn write_ppr_with_positioned_raw<W: std::io::Write>(
    writer: &mut Writer<W>,
    generated: &[u8],
    ppr: &CT_PPr,
) -> Result<()> {
    let mut raw_order = (0..ppr.revision_xml.len())
        .filter(|index| {
            replay_modeled_toggle_raw(
                ppr.revision_xml_positions[*index],
                ppr.revision_xml_positions[*index].0 != PPR_BIDI_SLOT || ppr.bidi.is_some(),
            )
        })
        .collect::<Vec<_>>();
    raw_order.sort_by_key(|index| {
        let (slot, occurrence) = effective_ppr_raw_position(ppr, *index);
        (slot, occurrence, *index)
    });
    let mut raw_index = 0usize;
    let mut occurrences = [0usize; PPR_END_SLOT as usize + 1];
    let mut reader = Reader::from_reader(generated);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut inside = false;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(element) if !inside => {
                inside = true;
                writer.write_event(Event::Start(element.into_owned()))?;
            }
            Event::Start(element) => {
                let slot = ppr_slot_for_name(element.local_name().as_ref());
                write_ppr_raw_before(
                    writer,
                    ppr,
                    &raw_order,
                    &mut raw_index,
                    slot,
                    occurrences[slot as usize],
                )?;
                let raw = capture_element(&mut reader, &element)?;
                writer.get_mut().write_all(&raw)?;
                occurrences[slot as usize] += 1;
            }
            Event::Empty(mut element) => {
                let slot = ppr_slot_for_name(element.local_name().as_ref());
                write_ppr_raw_before(
                    writer,
                    ppr,
                    &raw_order,
                    &mut raw_index,
                    slot,
                    occurrences[slot as usize],
                )?;
                append_modeled_toggle_attributes(
                    &mut element,
                    &ppr.revision_xml,
                    &ppr.revision_xml_positions,
                    slot,
                    occurrences[slot as usize],
                )?;
                writer.write_event(Event::Empty(element.into_owned()))?;
                occurrences[slot as usize] += 1;
            }
            Event::End(element) if inside => {
                write_ppr_raw_before(writer, ppr, &raw_order, &mut raw_index, PPR_END_SLOT, 0)?;
                while let Some(index) = raw_order.get(raw_index).copied() {
                    writer.get_mut().write_all(&ppr.revision_xml[index])?;
                    raw_index += 1;
                }
                writer.write_event(Event::End(element.into_owned()))?;
                return Ok(());
            }
            Event::Eof => return Ok(()),
            event => writer.write_event(event.into_owned())?,
        }
        buffer.clear();
    }
}

fn write_ppr_raw_before<W: std::io::Write>(
    writer: &mut Writer<W>,
    ppr: &CT_PPr,
    raw_order: &[usize],
    raw_index: &mut usize,
    slot: u8,
    occurrence: usize,
) -> Result<()> {
    while let Some(index) = raw_order.get(*raw_index).copied() {
        let position = effective_ppr_raw_position(ppr, index);
        if position.0 > slot || (position.0 == slot && position.1 > occurrence) {
            break;
        }
        writer.get_mut().write_all(&ppr.revision_xml[index])?;
        *raw_index += 1;
    }
    Ok(())
}

fn effective_ppr_raw_position(ppr: &CT_PPr, index: usize) -> (u8, usize) {
    let mut position = ppr.revision_xml_positions[index];
    position.1 = raw_occurrence(position);
    if position.0 == PPR_BIDI_SLOT
        && let Some((carrier_index, carrier_occurrence)) = ppr
            .revision_xml_positions
            .iter()
            .enumerate()
            .find(|candidate| {
                candidate.1.0 == PPR_BIDI_SLOT && raw_is_modeled_attribute_carrier(*candidate.1)
            })
            .map(|(carrier_index, candidate)| (carrier_index, raw_occurrence(*candidate)))
    {
        position.1 = usize::from(
            position.1 > carrier_occurrence
                || (position.1 == carrier_occurrence && index > carrier_index),
        );
    }
    if ppr.change.is_some() && position.0 >= PPR_CHANGE_SLOT {
        (PPR_CHANGE_SLOT, 0)
    } else {
        position
    }
}

fn ppr_slot_for_name(local: &[u8]) -> u8 {
    match local {
        b"pStyle" => PPR_STYLE_SLOT,
        b"keepNext" => PPR_KEEP_NEXT_SLOT,
        b"keepLines" => PPR_KEEP_LINES_SLOT,
        b"pageBreakBefore" => PPR_PAGE_BREAK_SLOT,
        b"framePr" => 4,
        b"widowControl" => PPR_WIDOW_SLOT,
        b"numPr" => PPR_NUM_SLOT,
        b"suppressLineNumbers" => 7,
        b"pBdr" => PPR_BORDER_SLOT,
        b"shd" => PPR_SHADING_SLOT,
        b"tabs" => PPR_TABS_SLOT,
        b"suppressAutoHyphens" => PPR_SUPPRESS_HYPHENS_SLOT,
        b"kinsoku" => 12,
        b"wordWrap" => 13,
        b"overflowPunct" => 14,
        b"topLinePunct" => 15,
        b"autoSpaceDE" => 16,
        b"autoSpaceDN" => 17,
        b"bidi" => PPR_BIDI_SLOT,
        b"adjustRightInd" => 19,
        b"snapToGrid" => 20,
        b"spacing" => PPR_SPACING_SLOT,
        b"ind" => PPR_INDENT_SLOT,
        b"contextualSpacing" => 23,
        b"mirrorIndents" => 24,
        b"suppressOverlap" => 25,
        b"jc" => PPR_JUSTIFICATION_SLOT,
        b"textDirection" => 27,
        b"textAlignment" => 28,
        b"textboxTightWrap" => 29,
        b"outlineLvl" => PPR_OUTLINE_SLOT,
        b"divId" => 31,
        b"cnfStyle" => 32,
        b"rPr" => PPR_RUN_PROPERTIES_SLOT,
        b"sectPr" => PPR_SECTION_SLOT,
        b"pPrChange" => PPR_CHANGE_SLOT,
        _ => PPR_END_SLOT,
    }
}

fn ppr_schema_slot(name: &[u8], word_prefixes: &[String]) -> Option<u8> {
    let local = name.rsplit(|byte| *byte == b':').next().unwrap_or(name);
    let slot = ppr_slot_for_name(local);
    (slot != PPR_END_SLOT && is_word_element(name, local, word_prefixes)).then_some(slot)
}

fn ppr_modeled_slot(name: &[u8], word_prefixes: &[String]) -> Option<u8> {
    let local = name.rsplit(|byte| *byte == b':').next().unwrap_or(name);
    if !is_word_element(name, local, word_prefixes) {
        return None;
    }
    matches!(
        local,
        b"pStyle"
            | b"keepNext"
            | b"keepLines"
            | b"pageBreakBefore"
            | b"widowControl"
            | b"numPr"
            | b"pBdr"
            | b"shd"
            | b"tabs"
            | b"suppressAutoHyphens"
            | b"bidi"
            | b"spacing"
            | b"ind"
            | b"jc"
            | b"outlineLvl"
            | b"rPr"
            | b"sectPr"
            | b"pPrChange"
    )
    .then(|| ppr_slot_for_name(local))
}

fn write_rpr_with_positioned_raw<W: std::io::Write>(
    writer: &mut Writer<W>,
    generated: &[u8],
    rpr: &CT_RPr,
    foreign_word_namespace: Option<&str>,
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
    let mut buffer = Vec::new();
    let mut inside = false;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(element) if !inside => {
                inside = true;
                writer.write_event(Event::Start(element.into_owned()))?;
            }
            Event::Start(element) => {
                let slot = rpr_slot_for_name(element.local_name().as_ref());
                write_rpr_raw_before(
                    writer,
                    rpr,
                    &raw_order,
                    &mut raw_index,
                    slot,
                    occurrences[slot as usize],
                    foreign_word_namespace,
                )?;
                let raw = capture_element(&mut reader, &element)?;
                writer.get_mut().write_all(&raw)?;
                occurrences[slot as usize] += 1;
            }
            Event::Empty(mut element) => {
                let slot = rpr_slot_for_name(element.local_name().as_ref());
                write_rpr_raw_before(
                    writer,
                    rpr,
                    &raw_order,
                    &mut raw_index,
                    slot,
                    occurrences[slot as usize],
                    foreign_word_namespace,
                )?;
                append_modeled_toggle_attributes(
                    &mut element,
                    &rpr.revision_xml,
                    &rpr.revision_xml_positions,
                    slot,
                    occurrences[slot as usize],
                )?;
                writer.write_event(Event::Empty(element.into_owned()))?;
                occurrences[slot as usize] += 1;
            }
            Event::End(element) if inside => {
                write_rpr_raw_before(
                    writer,
                    rpr,
                    &raw_order,
                    &mut raw_index,
                    RPR_END_SLOT,
                    0,
                    foreign_word_namespace,
                )?;
                while let Some(index) = raw_order.get(raw_index).copied() {
                    crate::text::write_raw_with_word_override(
                        writer,
                        &rpr.revision_xml[index],
                        foreign_word_namespace,
                    )?;
                    raw_index += 1;
                }
                writer.write_event(Event::End(element.into_owned()))?;
                return Ok(());
            }
            Event::Eof => return Ok(()),
            event => writer.write_event(event.into_owned())?,
        }
        buffer.clear();
    }
}

fn write_rpr_raw_before<W: std::io::Write>(
    writer: &mut Writer<W>,
    rpr: &CT_RPr,
    raw_order: &[usize],
    raw_index: &mut usize,
    slot: u8,
    occurrence: usize,
    foreign_word_namespace: Option<&str>,
) -> Result<()> {
    while let Some(index) = raw_order.get(*raw_index).copied() {
        let position = effective_rpr_raw_position(rpr, index);
        if position.0 > slot || (position.0 == slot && position.1 > occurrence) {
            break;
        }
        crate::text::write_raw_with_word_override(
            writer,
            &rpr.revision_xml[index],
            foreign_word_namespace,
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

fn rpr_schema_slot(name: &[u8], word_prefixes: &[String]) -> Option<u8> {
    let local = name.rsplit(|byte| *byte == b':').next().unwrap_or(name);
    let slot = rpr_slot_for_name(local);
    (slot != RPR_END_SLOT && is_word_element(name, local, word_prefixes)).then_some(slot)
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

fn parse_word_toggle(e: &BytesStart, word_prefixes: &[String]) -> Result<bool> {
    let val = get_word_val_attr(e, word_prefixes)?;
    Ok(ST_OnOff::from_str_or_default(val.as_deref()).is_on())
}

/// Extract the `w:val` attribute from an element.
pub(crate) fn get_val_attr(e: &BytesStart) -> Result<Option<String>> {
    for attr in e.attributes() {
        let attr = attr?;
        if matches_local_name(attr.key.as_ref(), b"val") {
            return Ok(Some(std::str::from_utf8(&attr.value)?.to_string()));
        }
    }
    Ok(None)
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::shared::ST_Border;

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
    fn parse_basic_ppr() {
        let ppr = parse_ppr(r#"<w:pStyle w:val="Heading1"/><w:jc w:val="center"/>"#);
        assert_eq!(ppr.style_id, Some("Heading1".to_string()));
        assert_eq!(ppr.jc, Some(ST_Jc::Center));
    }

    #[test]
    fn parse_spacing() {
        let ppr = parse_ppr(r#"<w:spacing w:before="240" w:after="120" w:line="360"/>"#);
        assert_eq!(ppr.space_before, Some(Twips(240)));
        assert_eq!(ppr.space_after, Some(Twips(120)));
        assert_eq!(ppr.line_spacing, Some(Twips(360)));
    }

    #[test]
    fn fractional_line_spacing_rounds_to_nearest_twip() {
        for (value, expected) in [
            ("257.1432", 257),
            ("320.00879999999995", 320),
            ("342.8616", 343),
            ("1.5", 2),
            ("-1.5", -2),
            ("2147483647.0", i32::MAX),
            ("-2147483648.0", i32::MIN),
        ] {
            let ppr = parse_ppr(&format!(r#"<w:spacing w:line="{value}"/>"#));
            assert_eq!(ppr.line_spacing, Some(Twips(expected)));
        }
    }

    #[test]
    fn invalid_fractional_line_spacing_remains_rejected() {
        for value in [
            "NaN",
            "inf",
            "-inf",
            "1e3",
            "1.2.3",
            "2147483648.0",
            "-2147483649.0",
        ] {
            assert!(
                try_parse_ppr(&format!(r#"<w:spacing w:line="{value}"/>"#)).is_err(),
                "accepted invalid line spacing {value}"
            );
        }
    }

    #[test]
    fn parse_borders() {
        let ppr = parse_ppr(
            r#"<w:pBdr><w:top w:val="single" w:sz="4" w:space="1" w:color="000000"/><w:bottom w:val="double" w:sz="6"/></w:pBdr>"#,
        );
        let borders = ppr.borders.unwrap();
        assert_eq!(borders.top.as_ref().unwrap().val, ST_Border::Single);
        assert_eq!(borders.top.as_ref().unwrap().sz, Some(4));
        assert_eq!(borders.bottom.as_ref().unwrap().val, ST_Border::Double);
        assert!(borders.left.is_none());
    }

    #[test]
    fn parse_tabs() {
        let ppr = parse_ppr(
            r#"<w:tabs><w:tab w:val="left" w:pos="720"/><w:tab w:val="right" w:pos="8640" w:leader="dot"/></w:tabs>"#,
        );
        let tabs = ppr.tabs.unwrap();
        assert_eq!(tabs.tabs.len(), 2);
        assert_eq!(tabs.tabs[0].pos, Twips(720));
        assert_eq!(tabs.tabs[1].leader, Some(crate::shared::ST_TabLeader::Dot));
    }

    #[test]
    fn canonical_ppr_ignores_foreign_same_local_containers() {
        let xml = r#"<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:ext="urn:producer"><w:tabs><ext:tab ext:val="right" ext:pos="100"/><w:tab w:val="left" w:pos="720"/></w:tabs><ext:tabs><ext:tab ext:val="right" ext:pos="99"/></ext:tabs><ext:rPr><ext:b/></ext:rPr></w:pPr>"#;
        let mut reader = Reader::from_str(xml);
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(_)) => break,
                Ok(Event::Eof) => panic!("missing pPr start"),
                _ => {}
            }
            buf.clear();
        }
        let ppr = CT_PPr::from_xml(&mut reader).unwrap();
        let tabs = ppr.tabs.unwrap();
        assert_eq!(tabs.tabs.len(), 1);
        assert_eq!(tabs.tabs[0].pos, Twips(720));
        assert_eq!(tabs.tabs[0].source_occurrence, Some(0));
        assert!(ppr.rpr.is_none());
    }

    #[test]
    fn parse_shading() {
        let ppr = parse_ppr(r#"<w:shd w:val="clear" w:fill="FFFF00"/>"#);
        let shd = ppr.shading.unwrap();
        assert_eq!(shd.val, "clear");
        assert_eq!(shd.fill, Some("FFFF00".to_string()));
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
    fn merge_ppr() {
        let mut base = CT_PPr {
            jc: Some(ST_Jc::Left),
            space_after: Some(Twips(200)),
            ..Default::default()
        };
        let override_ppr = CT_PPr {
            jc: Some(ST_Jc::Center),
            space_before: Some(Twips(120)),
            ..Default::default()
        };
        base.merge_from(&override_ppr);
        assert_eq!(base.jc, Some(ST_Jc::Center)); // overridden
        assert_eq!(base.space_after, Some(Twips(200))); // kept
        assert_eq!(base.space_before, Some(Twips(120))); // added
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

    #[test]
    fn extended_num_pr_payload_survives_value_overlay_and_unlink() {
        let mut ppr = parse_ppr(
            r#"<w:numPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:ext="urn:producer" ext:root="kept"><ext:before/><w:ilvl w:val="2" ext:leaf="level"><ext:level-child/></w:ilvl><ext:between/><w:numId w:val="7" ext:leaf="id"><ext:id-child/></w:numId><ext:after/></w:numPr>"#,
        );
        assert_eq!((ppr.num_id, ppr.num_ilvl), (Some(7), Some(2)));
        ppr.num_id = Some(9);
        ppr.num_ilvl = Some(1);
        let mut output = Vec::new();
        ppr.to_xml(&mut Writer::new(&mut output)).unwrap();
        let output = String::from_utf8(output).unwrap();
        assert_eq!(
            output,
            r#"<w:pPr><w:numPr xmlns:ext="urn:producer" ext:root="kept"><ext:before/><w:ilvl w:val="1" ext:leaf="level"><ext:level-child/></w:ilvl><ext:between/><w:numId w:val="9" ext:leaf="id"><ext:id-child/></w:numId><ext:after/></w:numPr></w:pPr>"#
        );
        for retained in [
            r#"ext:root="kept""#,
            "<ext:before/>",
            r#"ext:leaf="level""#,
            "<ext:level-child/>",
            "<ext:between/>",
            r#"ext:leaf="id""#,
            "<ext:id-child/>",
            "<ext:after/>",
        ] {
            assert!(output.contains(retained), "{output}");
        }
        assert!(output.contains(r#"w:val="1""#), "{output}");
        assert!(output.contains(r#"w:val="9""#), "{output}");
        assert!(output.find("<w:ilvl").unwrap() < output.find("<w:numId").unwrap());

        ppr.num_id = None;
        ppr.num_ilvl = None;
        let mut cleared = Vec::new();
        ppr.to_xml(&mut Writer::new(&mut cleared)).unwrap();
        let cleared = String::from_utf8(cleared).unwrap();
        assert_eq!(
            cleared,
            r#"<w:pPr><w:numPr xmlns:ext="urn:producer" ext:root="kept"><ext:before/><w:ilvl ext:leaf="level"><ext:level-child/></w:ilvl><ext:between/><w:numId ext:leaf="id"><ext:id-child/></w:numId><ext:after/></w:numPr></w:pPr>"#
        );
    }

    #[test]
    fn self_closing_num_pr_root_attributes_retain_inherited_bindings() {
        let raw = br#"<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:ext="urn:a&amp;b"><w:numPr ext:root="kept"/></w:pPr>"#;
        let ppr = crate::numbering::parse_scoped_ppr(raw, &["w".to_owned()]).unwrap();
        let mut output = Vec::new();
        ppr.to_xml(&mut Writer::new(&mut output)).unwrap();
        assert_eq!(
            String::from_utf8(output).unwrap(),
            r#"<w:pPr><w:numPr ext:root="kept" xmlns:ext="urn:a&amp;b"></w:numPr></w:pPr>"#
        );
    }

    #[test]
    fn num_pr_alias_duplicates_and_malformed_values_remain_ordered() {
        let mut ppr = parse_ppr(
            r#"<q:numPr xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:ext="urn:producer"><q:ilvl q:val="1" ext:mark="first"/><q:ilvl q:val="2"/><ext:between/><q:numId q:val="bad"/></q:numPr>"#,
        );
        assert_eq!(ppr.num_ilvl, Some(2));
        assert_eq!(ppr.num_id, None);

        let mut unchanged = Vec::new();
        ppr.to_xml(&mut Writer::new(&mut unchanged)).unwrap();
        let unchanged = String::from_utf8(unchanged).unwrap();
        let first = unchanged.find(r#"q:val="1" ext:mark="first""#).unwrap();
        let second = unchanged.find(r#"q:val="2""#).unwrap();
        let between = unchanged.find("<ext:between/>").unwrap();
        let malformed = unchanged.find(r#"q:val="bad""#).unwrap();
        assert!(first < second && second < between && between < malformed);

        ppr.num_ilvl = Some(3);
        ppr.num_id = Some(9);
        let mut changed = Vec::new();
        ppr.to_xml(&mut Writer::new(&mut changed)).unwrap();
        let changed = String::from_utf8(changed).unwrap();
        assert!(changed.contains(r#"q:val="1" ext:mark="first""#));
        assert!(changed.contains(r#"q:val="3""#));
        assert!(changed.contains(r#"q:val="9""#));
        assert!(!changed.contains(r#"q:val="bad""#));
    }

    #[test]
    fn foreign_num_pr_lookalikes_remain_untyped() {
        let ppr = parse_ppr(
            r#"<q:numPr xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:ext="urn:producer"><ext:ilvl ext:val="4"/><ext:numId ext:val="8"/></q:numPr>"#,
        );
        assert_eq!((ppr.num_id, ppr.num_ilvl), (None, None));
        let mut output = Vec::new();
        ppr.to_xml(&mut Writer::new(&mut output)).unwrap();
        let output = String::from_utf8(output).unwrap();
        assert!(output.contains(r#"<ext:ilvl ext:val="4"/>"#));
        assert!(output.contains(r#"<ext:numId ext:val="8"/>"#));
    }

    #[test]
    fn a_foreign_w_binding_on_num_pr_fails_closed() {
        let ppr = parse_ppr(
            r#"<q:numPr xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w="urn:producer"><q:numId q:val="7"/></q:numPr>"#,
        );
        let error = ppr
            .to_xml(&mut Writer::new(Vec::new()))
            .expect_err("a fixed w:numPr cannot retain a foreign w binding");
        assert!(error.to_string().contains("shadows the Word namespace"));
    }

    #[test]
    fn plain_num_pr_disappears_after_values_are_cleared() {
        let mut ppr = parse_ppr(r#"<w:numPr><w:ilvl w:val="2"/><w:numId w:val="7"/></w:numPr>"#);
        ppr.num_id = None;
        ppr.num_ilvl = None;
        let mut output = Vec::new();
        ppr.to_xml(&mut Writer::new(&mut output)).unwrap();
        let output = String::from_utf8(output).unwrap();
        assert!(!output.contains("numPr"), "{output}");
    }

    #[test]
    fn num_pr_revision_and_producer_children_keep_total_source_order() {
        let source = r#"<w:numPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:producer"><w:ins w:id="501" w:author="Ada"/><x:a/><x:ins x:mark="foreign"/><x:b/><w:ins w:id="502" w:author="Ada"/><x:c/></w:numPr>"#;
        let write = |ppr: &CT_PPr| {
            let mut output = Vec::new();
            ppr.to_xml(&mut Writer::new(&mut output)).unwrap();
            String::from_utf8(output).unwrap()
        };
        let assert_order = |output: &str, tokens: &[&str]| {
            let positions = tokens
                .iter()
                .map(|token| {
                    output
                        .find(token)
                        .unwrap_or_else(|| panic!("missing {token}: {output}"))
                })
                .collect::<Vec<_>>();
            assert!(
                positions.windows(2).all(|pair| pair[0] < pair[1]),
                "{output}"
            );
        };

        let mut ppr = parse_ppr(source);
        assert!(ppr.numbering_revision.is_some(), "{ppr:?}");
        let output = write(&ppr);
        assert_order(
            &output,
            &[
                r#"w:id="501""#,
                "<x:a",
                r#"x:mark="foreign""#,
                "<x:b",
                r#"w:id="502""#,
                "<x:c",
            ],
        );

        let reopened =
            parse_ppr(&output[output.find('>').unwrap() + 1..output.rfind("</w:pPr>").unwrap()]);
        assert_order(
            &write(&reopened),
            &[
                r#"w:id="501""#,
                "<x:a",
                r#"x:mark="foreign""#,
                "<x:b",
                r#"w:id="502""#,
                "<x:c",
            ],
        );

        ppr.numbering_revision = None;
        let cleared = write(&ppr);
        assert!(!cleared.contains(r#"w:id="502""#), "{cleared}");
        assert_order(
            &cleared,
            &[
                r#"w:id="501""#,
                "<x:a",
                r#"x:mark="foreign""#,
                "<x:b",
                "<x:c",
            ],
        );
    }

    #[test]
    fn authored_num_pr_values_precede_a_retained_numbering_change() {
        let mut ppr = parse_ppr(
            r#"<w:numPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:numberingChange w:id="9" w:author="Ada"><w:original w:val="kept"/></w:numberingChange></w:numPr>"#,
        );
        ppr.num_ilvl = Some(2);
        ppr.num_id = Some(7);
        let mut output = Vec::new();
        ppr.to_xml(&mut Writer::new(&mut output)).unwrap();
        let output = String::from_utf8(output).unwrap();
        let ilvl = output.find("<w:ilvl").unwrap();
        let num_id = output.find("<w:numId").unwrap();
        let change = output.find("<w:numberingChange").unwrap();
        assert!(ilvl < num_id && num_id < change, "{output}");
        assert!(output.contains(r#"<w:original w:val="kept"/>"#), "{output}");
    }
}
