//! Paragraph properties (`CT_PPr`).

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
use crate::properties::{
    CT_Shd, append_modeled_toggle_attributes, get_word_val_attr, is_word_attribute,
    is_word_element, parse_word_toggle, raw_is_modeled_attribute_carrier, raw_occurrence,
    record_modeled_toggle_candidate, remove_redundant_modeled_toggle_candidate,
    replay_modeled_toggle_raw, toggle_element_is_explicitly_empty,
    toggle_has_unsupported_attributes, word_prefixes_at, write_toggle,
};
use crate::raw_xml::{capture_element, capture_empty_element};
use crate::revision::CT_Revision;
use crate::run_properties::CT_RPr;
use crate::shared::{ST_Jc, ST_OnOff};
use crate::units::Twips;

/// `CT_FramePr` — Text frame placement for a paragraph (`w:framePr`).
///
/// Every member is an attribute. A value outside its `ST_*` enumeration stays
/// a `String` so a producer token survives unchanged.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct CT_FramePr {
    /// Drop cap placement (@w:dropCap): "none", "drop", "margin".
    pub drop_cap: Option<String>,
    /// Drop cap height in lines (@w:lines).
    pub lines: Option<u32>,
    /// Frame width in twips (@w:w).
    pub w: Option<Twips>,
    /// Frame height in twips (@w:h).
    pub h: Option<Twips>,
    /// Vertical distance from the surrounding text in twips (@w:vSpace).
    pub v_space: Option<Twips>,
    /// Horizontal distance from the surrounding text in twips (@w:hSpace).
    pub h_space: Option<Twips>,
    /// Text wrapping around the frame (@w:wrap).
    pub wrap: Option<String>,
    /// Horizontal anchor (@w:hAnchor): "margin", "page", "text".
    pub h_anchor: Option<String>,
    /// Vertical anchor (@w:vAnchor): "margin", "page", "text".
    pub v_anchor: Option<String>,
    /// Absolute horizontal position in twips (@w:x).
    pub x: Option<Twips>,
    /// Relative horizontal alignment (@w:xAlign).
    pub x_align: Option<String>,
    /// Absolute vertical position in twips (@w:y).
    pub y: Option<Twips>,
    /// Relative vertical alignment (@w:yAlign).
    pub y_align: Option<String>,
    /// Frame height rule (@w:hRule): "auto", "atLeast", "exact".
    pub h_rule: Option<String>,
    /// Whether the frame stays with the paragraph it anchors (@w:anchorLock).
    pub anchor_lock: Option<bool>,
    /// Attributes this type does not model, in source order.
    ///
    /// Names and values keep their stored spelling, and they are written back
    /// ahead of the modeled attributes so typing `w:framePr` drops nothing.
    /// The value is serialized verbatim, so a caller storing one itself is
    /// storing attribute-value syntax and owns the escaping.
    pub extra_attributes: Vec<(String, String)>,
}

impl CT_FramePr {
    fn from_xml_attrs_with_prefixes(e: &BytesStart, word_prefixes: &[String]) -> Result<Self> {
        let mut frame = CT_FramePr::default();
        for attribute in e.attributes() {
            let attribute = attribute?;
            let key = attribute.key.as_ref();
            let value = std::str::from_utf8(&attribute.value)?;
            if is_word_attribute(key, b"dropCap", word_prefixes) {
                frame.drop_cap = Some(value.to_owned());
            } else if is_word_attribute(key, b"lines", word_prefixes) {
                frame.lines = Some(value.parse()?);
            } else if is_word_attribute(key, b"w", word_prefixes) {
                frame.w = Some(Twips(value.parse()?));
            } else if is_word_attribute(key, b"h", word_prefixes) {
                frame.h = Some(Twips(value.parse()?));
            } else if is_word_attribute(key, b"vSpace", word_prefixes) {
                frame.v_space = Some(Twips(value.parse()?));
            } else if is_word_attribute(key, b"hSpace", word_prefixes) {
                frame.h_space = Some(Twips(value.parse()?));
            } else if is_word_attribute(key, b"wrap", word_prefixes) {
                frame.wrap = Some(value.to_owned());
            } else if is_word_attribute(key, b"hAnchor", word_prefixes) {
                frame.h_anchor = Some(value.to_owned());
            } else if is_word_attribute(key, b"vAnchor", word_prefixes) {
                frame.v_anchor = Some(value.to_owned());
            } else if is_word_attribute(key, b"x", word_prefixes) {
                frame.x = Some(Twips(value.parse()?));
            } else if is_word_attribute(key, b"xAlign", word_prefixes) {
                frame.x_align = Some(value.to_owned());
            } else if is_word_attribute(key, b"y", word_prefixes) {
                frame.y = Some(Twips(value.parse()?));
            } else if is_word_attribute(key, b"yAlign", word_prefixes) {
                frame.y_align = Some(value.to_owned());
            } else if is_word_attribute(key, b"hRule", word_prefixes) {
                frame.h_rule = Some(value.to_owned());
            } else if is_word_attribute(key, b"anchorLock", word_prefixes) {
                frame.anchor_lock = Some(ST_OnOff::from_str_or_default(Some(value)).is_on());
            } else {
                frame
                    .extra_attributes
                    .push((std::str::from_utf8(key)?.to_owned(), value.to_owned()));
            }
        }
        Ok(frame)
    }

    fn to_xml<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut buf = itoa::Buffer::new();
        let mut e = BytesStart::new("w:framePr");
        for (name, value) in &self.extra_attributes {
            e.push_attribute(Attribute {
                key: QName(name.as_bytes()),
                value: Cow::Borrowed(value.as_bytes()),
            });
        }
        if let Some(ref drop_cap) = self.drop_cap {
            e.push_attribute(("w:dropCap", drop_cap.as_str()));
        }
        if let Some(lines) = self.lines {
            e.push_attribute(("w:lines", buf.format(lines)));
        }
        if let Some(w) = self.w {
            e.push_attribute(("w:w", buf.format(w.0)));
        }
        if let Some(h) = self.h {
            e.push_attribute(("w:h", buf.format(h.0)));
        }
        if let Some(v_space) = self.v_space {
            e.push_attribute(("w:vSpace", buf.format(v_space.0)));
        }
        if let Some(h_space) = self.h_space {
            e.push_attribute(("w:hSpace", buf.format(h_space.0)));
        }
        if let Some(ref wrap) = self.wrap {
            e.push_attribute(("w:wrap", wrap.as_str()));
        }
        if let Some(ref h_anchor) = self.h_anchor {
            e.push_attribute(("w:hAnchor", h_anchor.as_str()));
        }
        if let Some(ref v_anchor) = self.v_anchor {
            e.push_attribute(("w:vAnchor", v_anchor.as_str()));
        }
        if let Some(x) = self.x {
            e.push_attribute(("w:x", buf.format(x.0)));
        }
        if let Some(ref x_align) = self.x_align {
            e.push_attribute(("w:xAlign", x_align.as_str()));
        }
        if let Some(y) = self.y {
            e.push_attribute(("w:y", buf.format(y.0)));
        }
        if let Some(ref y_align) = self.y_align {
            e.push_attribute(("w:yAlign", y_align.as_str()));
        }
        if let Some(ref h_rule) = self.h_rule {
            e.push_attribute(("w:hRule", h_rule.as_str()));
        }
        if let Some(anchor_lock) = self.anchor_lock {
            e.push_attribute(("w:anchorLock", if anchor_lock { "1" } else { "0" }));
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
    /// East Asian line-breaking rules (kinsoku), schema slot 12.
    ///
    /// The six East Asian line-breaking policy toggles are modeled and
    /// authored here. Their effect on break opportunities and inter-script
    /// spacing is named as the remaining boundary on DOCX-033.
    pub kinsoku: Option<bool>,
    /// Break Latin words only at a break opportunity (wordWrap), slot 13.
    pub word_wrap: Option<bool>,
    /// Let trailing punctuation hang past the measure (overflowPunct), slot 14.
    pub overflow_punct: Option<bool>,
    /// Compress punctuation at the start of a line (topLinePunct), slot 15.
    pub top_line_punct: Option<bool>,
    /// Auto space between East Asian and Latin text (autoSpaceDE), slot 16.
    pub auto_space_de: Option<bool>,
    /// Auto space between East Asian text and digits (autoSpaceDN), slot 17.
    pub auto_space_dn: Option<bool>,
    /// Snap line advance to the section character grid (snapToGrid), slot 20.
    pub snap_to_grid: Option<bool>,
    /// Paragraph base direction (bidi).
    pub bidi: Option<bool>,
    /// Text frame placement (framePr).
    ///
    /// Boxed because a frame is rare and `CT_FramePr` is large next to the
    /// rest of this struct, which sits on the stack in recursive callers.
    pub frame: Option<Box<CT_FramePr>>,
    /// Suppress line numbering for this paragraph (suppressLineNumbers)
    pub suppress_line_numbers: Option<bool>,
    /// Adjust right indentation for a document grid (adjustRightInd)
    pub adjust_right_ind: Option<bool>,
    /// Drop the space between paragraphs of the same style (contextualSpacing)
    pub contextual_spacing: Option<bool>,
    /// Swap inside and outside indentation on facing pages (mirrorIndents)
    pub mirror_indents: Option<bool>,
    /// Suppress overlap with a neighbouring frame (suppressOverlap)
    pub suppress_overlap: Option<bool>,
    /// Paragraph text flow (textDirection/@w:val), e.g. "lrTb" or "tbRlV".
    pub text_direction: Option<String>,
    /// Vertical character alignment on a line (textAlignment/@w:val).
    pub text_alignment: Option<String>,
    /// Text box tight wrap mode (textboxTightWrap/@w:val).
    pub textbox_tight_wrap: Option<String>,
    /// Web settings division this paragraph belongs to (divId/@w:val).
    pub div_id: Option<u32>,
    /// `w:cnfStyle` — which conditional parts of a table style this paragraph
    /// selects, as the twelve-bit string the schema uses.
    ///
    /// The row and cell selectors on `CT_TrPr` and `CT_TcPr` carry the same
    /// shape. A paragraph inside a styled table adds its own selection to
    /// theirs.
    pub cnf_style: Option<String>,
    /// Outline level 0-8 (outlineLvl)
    pub outline_lvl: Option<u32>,
    /// Paragraph borders (pBdr).
    ///
    /// Boxed for the same reason as [`CT_PPr::frame`]. Six typed edges are
    /// large next to the rest of this struct.
    pub borders: Option<Box<CT_PBdr>>,
    /// Tab stops (tabs)
    pub tabs: Option<CT_Tabs>,
    /// Paragraph shading (shd)
    /// Paragraph shading (shd).
    ///
    /// Boxed because `CT_Shd` now models its six theme attributes, and the
    /// paragraph property struct has no stack budget to spare. Same reason as
    /// `CT_PPr::borders`.
    pub shading: Option<Box<CT_Shd>>,
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
const PPR_FRAME_SLOT: u8 = 4;
const PPR_WIDOW_SLOT: u8 = 5;
const PPR_NUM_SLOT: u8 = 6;
const PPR_SUPPRESS_LINE_NUMBERS_SLOT: u8 = 7;
const PPR_BORDER_SLOT: u8 = 8;
const PPR_SHADING_SLOT: u8 = 9;
const PPR_TABS_SLOT: u8 = 10;
const PPR_SUPPRESS_HYPHENS_SLOT: u8 = 11;
const PPR_KINSOKU_SLOT: u8 = 12;
const PPR_WORD_WRAP_SLOT: u8 = 13;
const PPR_OVERFLOW_PUNCT_SLOT: u8 = 14;
const PPR_TOP_LINE_PUNCT_SLOT: u8 = 15;
const PPR_AUTO_SPACE_DE_SLOT: u8 = 16;
const PPR_AUTO_SPACE_DN_SLOT: u8 = 17;
const PPR_BIDI_SLOT: u8 = 18;
const PPR_ADJUST_RIGHT_IND_SLOT: u8 = 19;
const PPR_SNAP_TO_GRID_SLOT: u8 = 20;
const PPR_SPACING_SLOT: u8 = 21;
const PPR_INDENT_SLOT: u8 = 22;
const PPR_CONTEXTUAL_SPACING_SLOT: u8 = 23;
const PPR_MIRROR_INDENTS_SLOT: u8 = 24;
const PPR_SUPPRESS_OVERLAP_SLOT: u8 = 25;
const PPR_JUSTIFICATION_SLOT: u8 = 26;
const PPR_TEXT_DIRECTION_SLOT: u8 = 27;
const PPR_TEXT_ALIGNMENT_SLOT: u8 = 28;
const PPR_TEXTBOX_TIGHT_WRAP_SLOT: u8 = 29;
const PPR_OUTLINE_SLOT: u8 = 30;
const PPR_DIV_ID_SLOT: u8 = 31;
const PPR_CNF_STYLE_SLOT: u8 = 32;
const PPR_RUN_PROPERTIES_SLOT: u8 = 33;
const PPR_SECTION_SLOT: u8 = 34;
const PPR_CHANGE_SLOT: u8 = 35;
const PPR_END_SLOT: u8 = 36;

fn parse_line_spacing(value: &str) -> Result<Twips> {
    let integer_error = match value.parse::<i32>() {
        Ok(value) => return Ok(Twips(value)),
        Err(error) => error,
    };

    let (negative, unsigned) = match value.as_bytes().first() {
        Some(b'-') => (true, &value[1..]),
        Some(b'+') => (false, &value[1..]),
        _ => (false, value),
    };
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

    let significant_whole = whole.trim_start_matches('0');
    if significant_whole.len() > 10 {
        return Err(OxmlError::InvalidValue(format!(
            "line spacing is out of range: {value}"
        )));
    }
    let whole = if significant_whole.is_empty() {
        0
    } else {
        significant_whole
            .parse::<u64>()
            .map_err(|_| OxmlError::InvalidValue(format!("invalid line spacing: {value}")))?
    };
    let limit = if negative {
        i32::MAX as u64 + 1
    } else {
        i32::MAX as u64
    };
    let has_fraction = fraction.bytes().any(|byte| byte != b'0');
    if whole > limit || (whole == limit && has_fraction) {
        return Err(OxmlError::InvalidValue(format!(
            "line spacing is out of range: {value}"
        )));
    }

    let rounded = whole + u64::from(fraction.as_bytes()[0] >= b'5');
    if rounded > limit {
        return Err(OxmlError::InvalidValue(format!(
            "line spacing rounds out of range: {value}"
        )));
    }
    let signed = if negative {
        -(rounded as i64)
    } else {
        rounded as i64
    };
    Ok(Twips(signed as i32))
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
        let mut toggle_carrier_required = [false; PPR_END_SLOT as usize + 1];
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
                        ppr.borders = Some(Box::new(CT_PBdr::from_xml_with_prefixes(
                            reader, &prefixes,
                        )?));
                    } else if is_word_element(name.as_ref(), b"tabs", &prefixes) {
                        ppr.tabs = Some(CT_Tabs::from_xml_with_prefixes(reader, &prefixes)?);
                    } else if is_word_element(name.as_ref(), b"sectPr", &prefixes) {
                        let local_bindings = local_namespace_overrides(e, word_prefixes)?;
                        let sect_pr_bindings =
                            merged_owner_bindings(owner_bindings, &local_bindings);
                        ppr.sect_pr =
                            Some(CT_SectPr::from_xml_with_prefixes_owner_bindings_and_root(
                                reader,
                                &prefixes,
                                &sect_pr_bindings,
                                Some(e),
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
                    } else if let Some(slot) = ppr_toggle_slot(name.as_ref(), &prefixes) {
                        let captured = capture_element(reader, e)?;
                        let raw =
                            crate::text::raw_with_external_bindings(&captured, owner_bindings)?;
                        if toggle_element_is_explicitly_empty(&captured)? {
                            let value = parse_word_toggle(e, &prefixes)?;
                            if let Some(field) = ppr_toggle_field_mut(&mut ppr, slot) {
                                *field = Some(value);
                            }
                            toggle_carrier_required[slot as usize] =
                                toggle_has_unsupported_attributes(e, &prefixes)?;
                            record_modeled_toggle_candidate(
                                &mut ppr.revision_xml,
                                &mut ppr.revision_xml_positions,
                                raw,
                                slot,
                                occurrences[slot as usize] - 1,
                            );
                        } else {
                            let occurrence = occurrences[slot as usize] - 1;
                            record_ppr_raw_at(&mut ppr, raw, slot, occurrence);
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
                    } else if let Some(slot) = ppr_toggle_slot(name.as_ref(), &prefixes) {
                        let value = parse_word_toggle(e, &prefixes)?;
                        if let Some(field) = ppr_toggle_field_mut(&mut ppr, slot) {
                            *field = Some(value);
                        }
                        toggle_carrier_required[slot as usize] =
                            toggle_has_unsupported_attributes(e, &prefixes)?;
                        let raw = crate::text::raw_with_external_bindings(
                            &capture_empty_element(e)?,
                            owner_bindings,
                        )?;
                        record_modeled_toggle_candidate(
                            &mut ppr.revision_xml,
                            &mut ppr.revision_xml_positions,
                            raw,
                            slot,
                            occurrences[slot as usize] - 1,
                        );
                    } else if is_word_element(name.as_ref(), b"framePr", &prefixes) {
                        ppr.frame = Some(Box::new(CT_FramePr::from_xml_attrs_with_prefixes(
                            e, &prefixes,
                        )?));
                    } else if is_word_element(name.as_ref(), b"textDirection", &prefixes) {
                        ppr.text_direction = get_word_val_attr(e, &prefixes)?;
                    } else if is_word_element(name.as_ref(), b"textAlignment", &prefixes) {
                        ppr.text_alignment = get_word_val_attr(e, &prefixes)?;
                    } else if is_word_element(name.as_ref(), b"textboxTightWrap", &prefixes) {
                        ppr.textbox_tight_wrap = get_word_val_attr(e, &prefixes)?;
                    } else if is_word_element(name.as_ref(), b"divId", &prefixes) {
                        if let Some(val) = get_word_val_attr(e, &prefixes)? {
                            ppr.div_id = Some(val.parse()?);
                        }
                    } else if is_word_element(name.as_ref(), b"cnfStyle", &prefixes) {
                        // `CT_Cnf` carries a per-region attribute form beside
                        // `w:val`. The source element is retained as an
                        // attribute carrier when it holds anything but `w:val`,
                        // so that form survives being modeled, the same way the
                        // modeled toggles keep theirs.
                        let value = get_word_val_attr(e, &prefixes)?;
                        if value.is_none() || toggle_has_unsupported_attributes(e, &prefixes)? {
                            let raw = crate::text::raw_with_external_bindings(
                                &capture_empty_element(e)?,
                                owner_bindings,
                            )?;
                            let occurrence = occurrences[PPR_CNF_STYLE_SLOT as usize] - 1;
                            if value.is_some() {
                                record_modeled_toggle_candidate(
                                    &mut ppr.revision_xml,
                                    &mut ppr.revision_xml_positions,
                                    raw,
                                    PPR_CNF_STYLE_SLOT,
                                    occurrence,
                                );
                            } else {
                                // Without `w:val` there is nothing to model, so
                                // the element stays raw at its own slot.
                                record_ppr_raw_at(&mut ppr, raw, PPR_CNF_STYLE_SLOT, occurrence);
                            }
                        }
                        if value.is_some() {
                            ppr.cnf_style = value;
                        }
                    } else if is_word_element(name.as_ref(), b"outlineLvl", &prefixes) {
                        if let Some(val) = get_word_val_attr(e, &prefixes)? {
                            ppr.outline_lvl = Some(val.parse()?);
                        }
                    } else if is_word_element(name.as_ref(), b"shd", &prefixes) {
                        ppr.shading = Some(Box::new(CT_Shd::from_xml_attrs(e)?));
                    } else if is_word_element(name.as_ref(), b"numPr", &prefixes) {
                        ppr.num_pr_extra_attributes =
                            preserved_num_pr_root_attributes(e, &prefixes, owner_bindings)?.0;
                    } else if is_word_element(name.as_ref(), b"rPr", &prefixes)
                        && e.attributes().next().is_none()
                    {
                        // A bare empty paragraph mark, the same as
                        // `<w:rPr></w:rPr>`. An attributed one stays raw.
                        ppr.rpr.get_or_insert_default();
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
        for (_, slot) in PPR_MODELED_TOGGLES {
            remove_redundant_modeled_toggle_candidate(
                &mut ppr.revision_xml,
                &mut ppr.revision_xml_positions,
                slot,
                toggle_carrier_required[slot as usize],
            );
        }

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
        if let Some(ref frame) = self.frame {
            frame.to_xml(writer)?;
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

        if let Some(suppress) = self.suppress_line_numbers {
            write_toggle(writer, "w:suppressLineNumbers", suppress)?;
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

        if let Some(kinsoku) = self.kinsoku {
            write_toggle(writer, "w:kinsoku", kinsoku)?;
        }

        if let Some(word_wrap) = self.word_wrap {
            write_toggle(writer, "w:wordWrap", word_wrap)?;
        }

        if let Some(overflow_punct) = self.overflow_punct {
            write_toggle(writer, "w:overflowPunct", overflow_punct)?;
        }

        if let Some(top_line_punct) = self.top_line_punct {
            write_toggle(writer, "w:topLinePunct", top_line_punct)?;
        }

        if let Some(auto_space_de) = self.auto_space_de {
            write_toggle(writer, "w:autoSpaceDE", auto_space_de)?;
        }

        if let Some(auto_space_dn) = self.auto_space_dn {
            write_toggle(writer, "w:autoSpaceDN", auto_space_dn)?;
        }

        if let Some(bidi) = self.bidi {
            write_toggle(writer, "w:bidi", bidi)?;
        }

        if let Some(adjust) = self.adjust_right_ind {
            write_toggle(writer, "w:adjustRightInd", adjust)?;
        }

        if let Some(snap_to_grid) = self.snap_to_grid {
            write_toggle(writer, "w:snapToGrid", snap_to_grid)?;
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

        if let Some(contextual) = self.contextual_spacing {
            write_toggle(writer, "w:contextualSpacing", contextual)?;
        }
        if let Some(mirror) = self.mirror_indents {
            write_toggle(writer, "w:mirrorIndents", mirror)?;
        }
        if let Some(overlap) = self.suppress_overlap {
            write_toggle(writer, "w:suppressOverlap", overlap)?;
        }

        if let Some(jc) = self.jc {
            let mut e = BytesStart::new("w:jc");
            e.push_attribute(("w:val", jc.to_str()));
            writer.write_event(Event::Empty(e))?;
        }

        if let Some(ref direction) = self.text_direction {
            let mut e = BytesStart::new("w:textDirection");
            e.push_attribute(("w:val", direction.as_str()));
            writer.write_event(Event::Empty(e))?;
        }
        if let Some(ref alignment) = self.text_alignment {
            let mut e = BytesStart::new("w:textAlignment");
            e.push_attribute(("w:val", alignment.as_str()));
            writer.write_event(Event::Empty(e))?;
        }
        if let Some(ref tight_wrap) = self.textbox_tight_wrap {
            let mut e = BytesStart::new("w:textboxTightWrap");
            e.push_attribute(("w:val", tight_wrap.as_str()));
            writer.write_event(Event::Empty(e))?;
        }

        if let Some(lvl) = self.outline_lvl {
            let mut e = BytesStart::new("w:outlineLvl");
            e.push_attribute(("w:val", buf.format(lvl)));
            writer.write_event(Event::Empty(e))?;
        }

        if let Some(div_id) = self.div_id {
            let mut e = BytesStart::new("w:divId");
            e.push_attribute(("w:val", buf.format(div_id)));
            writer.write_event(Event::Empty(e))?;
        }

        if let Some(ref cnf_style) = self.cnf_style {
            let mut e = BytesStart::new("w:cnfStyle");
            e.push_attribute(("w:val", cnf_style.as_str()));
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
            && self.kinsoku.is_none()
            && self.word_wrap.is_none()
            && self.overflow_punct.is_none()
            && self.top_line_punct.is_none()
            && self.auto_space_de.is_none()
            && self.auto_space_dn.is_none()
            && self.snap_to_grid.is_none()
            && self.bidi.is_none()
            && self.frame.is_none()
            && self.suppress_line_numbers.is_none()
            && self.adjust_right_ind.is_none()
            && self.contextual_spacing.is_none()
            && self.mirror_indents.is_none()
            && self.suppress_overlap.is_none()
            && self.text_direction.is_none()
            && self.text_alignment.is_none()
            && self.textbox_tight_wrap.is_none()
            && self.div_id.is_none()
            && self.cnf_style.is_none()
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
        if other.kinsoku.is_some() {
            self.kinsoku = other.kinsoku;
        }
        if other.word_wrap.is_some() {
            self.word_wrap = other.word_wrap;
        }
        if other.overflow_punct.is_some() {
            self.overflow_punct = other.overflow_punct;
        }
        if other.top_line_punct.is_some() {
            self.top_line_punct = other.top_line_punct;
        }
        if other.auto_space_de.is_some() {
            self.auto_space_de = other.auto_space_de;
        }
        if other.auto_space_dn.is_some() {
            self.auto_space_dn = other.auto_space_dn;
        }
        if other.snap_to_grid.is_some() {
            self.snap_to_grid = other.snap_to_grid;
        }
        if other.bidi.is_some() {
            self.bidi = other.bidi;
        }
        if other.frame.is_some() {
            self.frame = other.frame.clone();
        }
        if other.suppress_line_numbers.is_some() {
            self.suppress_line_numbers = other.suppress_line_numbers;
        }
        if other.adjust_right_ind.is_some() {
            self.adjust_right_ind = other.adjust_right_ind;
        }
        if other.contextual_spacing.is_some() {
            self.contextual_spacing = other.contextual_spacing;
        }
        if other.mirror_indents.is_some() {
            self.mirror_indents = other.mirror_indents;
        }
        if other.suppress_overlap.is_some() {
            self.suppress_overlap = other.suppress_overlap;
        }
        if other.text_direction.is_some() {
            self.text_direction = other.text_direction.clone();
        }
        if other.text_alignment.is_some() {
            self.text_alignment = other.text_alignment.clone();
        }
        if other.textbox_tight_wrap.is_some() {
            self.textbox_tight_wrap = other.textbox_tight_wrap.clone();
        }
        if other.div_id.is_some() {
            self.div_id = other.div_id;
        }
        if other.cnf_style.is_some() {
            self.cnf_style.clone_from(&other.cnf_style);
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

fn write_ppr_with_positioned_raw<W: std::io::Write>(
    writer: &mut Writer<W>,
    generated: &[u8],
    ppr: &CT_PPr,
) -> Result<()> {
    let mut raw_order = (0..ppr.revision_xml.len())
        .filter(|index| {
            let position = ppr.revision_xml_positions[*index];
            replay_modeled_toggle_raw(position, ppr_modeled_toggle_present(ppr, position.0))
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
    if PPR_MODELED_TOGGLES
        .iter()
        .any(|(_, slot)| *slot == position.0)
        && let Some((carrier_index, carrier_occurrence)) = ppr
            .revision_xml_positions
            .iter()
            .enumerate()
            .find(|candidate| {
                candidate.1.0 == position.0 && raw_is_modeled_attribute_carrier(*candidate.1)
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
        b"framePr" => PPR_FRAME_SLOT,
        b"widowControl" => PPR_WIDOW_SLOT,
        b"numPr" => PPR_NUM_SLOT,
        b"suppressLineNumbers" => PPR_SUPPRESS_LINE_NUMBERS_SLOT,
        b"pBdr" => PPR_BORDER_SLOT,
        b"shd" => PPR_SHADING_SLOT,
        b"tabs" => PPR_TABS_SLOT,
        b"suppressAutoHyphens" => PPR_SUPPRESS_HYPHENS_SLOT,
        b"kinsoku" => PPR_KINSOKU_SLOT,
        b"wordWrap" => PPR_WORD_WRAP_SLOT,
        b"overflowPunct" => PPR_OVERFLOW_PUNCT_SLOT,
        b"topLinePunct" => PPR_TOP_LINE_PUNCT_SLOT,
        b"autoSpaceDE" => PPR_AUTO_SPACE_DE_SLOT,
        b"autoSpaceDN" => PPR_AUTO_SPACE_DN_SLOT,
        b"bidi" => PPR_BIDI_SLOT,
        b"adjustRightInd" => PPR_ADJUST_RIGHT_IND_SLOT,
        b"snapToGrid" => PPR_SNAP_TO_GRID_SLOT,
        b"spacing" => PPR_SPACING_SLOT,
        b"ind" => PPR_INDENT_SLOT,
        b"contextualSpacing" => PPR_CONTEXTUAL_SPACING_SLOT,
        b"mirrorIndents" => PPR_MIRROR_INDENTS_SLOT,
        b"suppressOverlap" => PPR_SUPPRESS_OVERLAP_SLOT,
        b"jc" => PPR_JUSTIFICATION_SLOT,
        b"textDirection" => PPR_TEXT_DIRECTION_SLOT,
        b"textAlignment" => PPR_TEXT_ALIGNMENT_SLOT,
        b"textboxTightWrap" => PPR_TEXTBOX_TIGHT_WRAP_SLOT,
        b"outlineLvl" => PPR_OUTLINE_SLOT,
        b"divId" => PPR_DIV_ID_SLOT,
        b"cnfStyle" => PPR_CNF_STYLE_SLOT,
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
            | b"framePr"
            | b"widowControl"
            | b"numPr"
            | b"suppressLineNumbers"
            | b"pBdr"
            | b"shd"
            | b"tabs"
            | b"suppressAutoHyphens"
            | b"kinsoku"
            | b"wordWrap"
            | b"overflowPunct"
            | b"topLinePunct"
            | b"autoSpaceDE"
            | b"autoSpaceDN"
            | b"snapToGrid"
            | b"bidi"
            | b"adjustRightInd"
            | b"spacing"
            | b"ind"
            | b"contextualSpacing"
            | b"mirrorIndents"
            | b"suppressOverlap"
            | b"jc"
            | b"textDirection"
            | b"textAlignment"
            | b"textboxTightWrap"
            | b"outlineLvl"
            | b"divId"
            | b"cnfStyle"
            | b"rPr"
            | b"sectPr"
            | b"pPrChange"
    )
    .then(|| ppr_slot_for_name(local))
}

/// The `w:pPr` toggles that retain their source element as an attribute
/// carrier, so an unowned attribute survives being modeled.
const PPR_MODELED_TOGGLES: [(&[u8], u8); 13] = [
    (b"kinsoku", PPR_KINSOKU_SLOT),
    (b"wordWrap", PPR_WORD_WRAP_SLOT),
    (b"overflowPunct", PPR_OVERFLOW_PUNCT_SLOT),
    (b"topLinePunct", PPR_TOP_LINE_PUNCT_SLOT),
    (b"autoSpaceDE", PPR_AUTO_SPACE_DE_SLOT),
    (b"autoSpaceDN", PPR_AUTO_SPACE_DN_SLOT),
    (b"snapToGrid", PPR_SNAP_TO_GRID_SLOT),
    (b"bidi", PPR_BIDI_SLOT),
    (b"suppressLineNumbers", PPR_SUPPRESS_LINE_NUMBERS_SLOT),
    (b"adjustRightInd", PPR_ADJUST_RIGHT_IND_SLOT),
    (b"contextualSpacing", PPR_CONTEXTUAL_SPACING_SLOT),
    (b"mirrorIndents", PPR_MIRROR_INDENTS_SLOT),
    (b"suppressOverlap", PPR_SUPPRESS_OVERLAP_SLOT),
];

/// The schema slot of a modeled `w:pPr` toggle, or `None` for any other
/// element name.
fn ppr_toggle_slot(name: &[u8], word_prefixes: &[String]) -> Option<u8> {
    let local = name.rsplit(|byte| *byte == b':').next().unwrap_or(name);
    if !is_word_element(name, local, word_prefixes) {
        return None;
    }
    PPR_MODELED_TOGGLES
        .iter()
        .find(|(candidate, _)| *candidate == local)
        .map(|(_, slot)| *slot)
}

fn ppr_toggle_field_mut(ppr: &mut CT_PPr, slot: u8) -> Option<&mut Option<bool>> {
    match slot {
        PPR_KINSOKU_SLOT => Some(&mut ppr.kinsoku),
        PPR_WORD_WRAP_SLOT => Some(&mut ppr.word_wrap),
        PPR_OVERFLOW_PUNCT_SLOT => Some(&mut ppr.overflow_punct),
        PPR_TOP_LINE_PUNCT_SLOT => Some(&mut ppr.top_line_punct),
        PPR_AUTO_SPACE_DE_SLOT => Some(&mut ppr.auto_space_de),
        PPR_AUTO_SPACE_DN_SLOT => Some(&mut ppr.auto_space_dn),
        PPR_SNAP_TO_GRID_SLOT => Some(&mut ppr.snap_to_grid),
        PPR_BIDI_SLOT => Some(&mut ppr.bidi),
        PPR_SUPPRESS_LINE_NUMBERS_SLOT => Some(&mut ppr.suppress_line_numbers),
        PPR_ADJUST_RIGHT_IND_SLOT => Some(&mut ppr.adjust_right_ind),
        PPR_CONTEXTUAL_SPACING_SLOT => Some(&mut ppr.contextual_spacing),
        PPR_MIRROR_INDENTS_SLOT => Some(&mut ppr.mirror_indents),
        PPR_SUPPRESS_OVERLAP_SLOT => Some(&mut ppr.suppress_overlap),
        _ => None,
    }
}

/// Whether the modeled toggle at `slot` still has a value, which decides
/// whether a retained duplicate carrier replays. A slot that is not a modeled
/// toggle is always present, so every other retained child is unaffected.
fn ppr_modeled_toggle_present(ppr: &CT_PPr, slot: u8) -> bool {
    match slot {
        PPR_KINSOKU_SLOT => ppr.kinsoku.is_some(),
        PPR_WORD_WRAP_SLOT => ppr.word_wrap.is_some(),
        PPR_OVERFLOW_PUNCT_SLOT => ppr.overflow_punct.is_some(),
        PPR_TOP_LINE_PUNCT_SLOT => ppr.top_line_punct.is_some(),
        PPR_AUTO_SPACE_DE_SLOT => ppr.auto_space_de.is_some(),
        PPR_AUTO_SPACE_DN_SLOT => ppr.auto_space_dn.is_some(),
        PPR_SNAP_TO_GRID_SLOT => ppr.snap_to_grid.is_some(),
        PPR_BIDI_SLOT => ppr.bidi.is_some(),
        PPR_SUPPRESS_LINE_NUMBERS_SLOT => ppr.suppress_line_numbers.is_some(),
        PPR_ADJUST_RIGHT_IND_SLOT => ppr.adjust_right_ind.is_some(),
        PPR_CONTEXTUAL_SPACING_SLOT => ppr.contextual_spacing.is_some(),
        PPR_MIRROR_INDENTS_SLOT => ppr.mirror_indents.is_some(),
        PPR_SUPPRESS_OVERLAP_SLOT => ppr.suppress_overlap.is_some(),
        PPR_CNF_STYLE_SLOT => ppr.cnf_style.is_some(),
        _ => true,
    }
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
    fn fractional_line_spacing_rounds_exactly_to_nearest_twip() {
        for (value, expected) in [
            ("257.1432", 257),
            ("320.00879999999995", 320),
            ("342.8616", 343),
            ("+1.5", 2),
            ("-1.5", -2),
            ("0.499999999999999999999999999999999999", 0),
            ("0.500000000000000000000000000000000000", 1),
            ("-0.499999999999999999999999999999999999", 0),
            ("-0.500000000000000000000000000000000000", -1),
            ("2147483646.5", i32::MAX),
            ("-2147483647.5", i32::MIN),
            ("2147483647.0", i32::MAX),
            ("-2147483648.0", i32::MIN),
        ] {
            let ppr = parse_ppr(&format!(r#"<w:spacing w:line="{value}"/>"#));
            assert_eq!(ppr.line_spacing, Some(Twips(expected)), "{value}");
        }
    }

    #[test]
    fn invalid_fractional_line_spacing_remains_rejected() {
        for value in [
            "NaN",
            "inf",
            "-inf",
            "1e3",
            ".5",
            "1.",
            "1.2.3",
            "2147483647.0000000000000000000000000000001",
            "-2147483648.0000000000000000000000000000001",
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
    fn fractional_line_spacing_accepts_aliases_and_writes_canonical_siblings() {
        let xml = r#"<x:pPr xmlns:x="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><x:spacing x:before="120" x:after="240" x:line="342.8616" x:lineRule="auto" x:beforeAutospacing="false" x:afterAutospacing="true"/></x:pPr>"#;
        let ppr = crate::numbering::parse_scoped_ppr(xml.as_bytes(), &["x".to_owned()]).unwrap();
        assert_eq!(ppr.space_before, Some(Twips(120)));
        assert_eq!(ppr.space_after, Some(Twips(240)));
        assert_eq!(ppr.line_spacing, Some(Twips(343)));
        assert_eq!(ppr.line_rule.as_deref(), Some("auto"));
        assert_eq!(ppr.before_autospacing, Some(false));
        assert_eq!(ppr.after_autospacing, Some(true));

        let mut output = Vec::new();
        ppr.to_xml(&mut Writer::new(&mut output)).unwrap();
        assert_eq!(
            String::from_utf8(output).unwrap(),
            r#"<w:pPr><w:spacing w:before="120" w:after="240" w:line="343" w:lineRule="auto" w:beforeAutospacing="0" w:afterAutospacing="1"/></w:pPr>"#
        );
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

    fn write_ppr(ppr: &CT_PPr) -> String {
        let mut output = Vec::new();
        ppr.to_xml(&mut Writer::new(&mut output)).unwrap();
        String::from_utf8(output).unwrap()
    }

    fn assert_token_order(output: &str, tokens: &[&str]) {
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
    }

    #[test]
    fn frame_properties_round_trip_every_modeled_attribute() {
        let source = concat!(
            r#"<w:framePr xmlns:ext="urn:producer" ext:first="kept" w:dropCap="drop" w:lines="3""#,
            r#" w:w="2880" w:h="1440" w:vSpace="120" w:hSpace="240" w:wrap="around""#,
            r#" w:hAnchor="margin" w:vAnchor="text" w:x="720" w:xAlign="center" w:y="360""#,
            r#" w:yAlign="top" w:hRule="exact" w:anchorLock="1" ext:last="kept"/>"#,
        );
        let ppr = parse_ppr(source);
        let frame = *ppr.frame.clone().expect("a typed w:framePr");
        assert_eq!(frame.drop_cap.as_deref(), Some("drop"));
        assert_eq!(frame.lines, Some(3));
        assert_eq!(frame.w, Some(Twips(2880)));
        assert_eq!(frame.h, Some(Twips(1440)));
        assert_eq!(frame.v_space, Some(Twips(120)));
        assert_eq!(frame.h_space, Some(Twips(240)));
        assert_eq!(frame.wrap.as_deref(), Some("around"));
        assert_eq!(frame.h_anchor.as_deref(), Some("margin"));
        assert_eq!(frame.v_anchor.as_deref(), Some("text"));
        assert_eq!(frame.x, Some(Twips(720)));
        assert_eq!(frame.x_align.as_deref(), Some("center"));
        assert_eq!(frame.y, Some(Twips(360)));
        assert_eq!(frame.y_align.as_deref(), Some("top"));
        assert_eq!(frame.h_rule.as_deref(), Some("exact"));
        assert_eq!(frame.anchor_lock, Some(true));
        assert_eq!(
            frame.extra_attributes,
            vec![
                ("xmlns:ext".to_owned(), "urn:producer".to_owned()),
                ("ext:first".to_owned(), "kept".to_owned()),
                ("ext:last".to_owned(), "kept".to_owned()),
            ]
        );

        let output = write_ppr(&ppr);
        assert_eq!(
            output,
            concat!(
                r#"<w:pPr><w:framePr xmlns:ext="urn:producer" ext:first="kept" ext:last="kept""#,
                r#" w:dropCap="drop" w:lines="3" w:w="2880" w:h="1440" w:vSpace="120""#,
                r#" w:hSpace="240" w:wrap="around" w:hAnchor="margin" w:vAnchor="text""#,
                r#" w:x="720" w:xAlign="center" w:y="360" w:yAlign="top" w:hRule="exact""#,
                r#" w:anchorLock="1"/></w:pPr>"#,
            )
        );

        let reopened = parse_ppr(
            output
                .strip_prefix("<w:pPr>")
                .unwrap()
                .strip_suffix("</w:pPr>")
                .unwrap(),
        );
        assert_eq!(reopened.frame, ppr.frame);
        assert_eq!(write_ppr(&reopened), output);
    }

    #[test]
    fn modeled_paragraph_children_serialize_in_schema_sequence_order() {
        let mut ppr = parse_ppr(
            r#"<w:kinsoku/><w:snapToGrid/><w:cnfStyle w:val="100000000000"/><w:autoSpaceDE/>"#,
        );
        ppr.frame = Some(Box::new(CT_FramePr {
            w: Some(Twips(2880)),
            ..Default::default()
        }));
        ppr.suppress_line_numbers = Some(true);
        ppr.adjust_right_ind = Some(true);
        ppr.contextual_spacing = Some(true);
        ppr.mirror_indents = Some(true);
        ppr.suppress_overlap = Some(true);
        ppr.jc = Some(ST_Jc::Center);
        ppr.text_direction = Some("tbRlV".to_owned());
        ppr.text_alignment = Some("center".to_owned());
        ppr.textbox_tight_wrap = Some("firstAndLastLine".to_owned());
        ppr.outline_lvl = Some(2);
        ppr.div_id = Some(7);

        assert_token_order(
            &write_ppr(&ppr),
            &[
                "<w:framePr",
                "<w:suppressLineNumbers",
                "<w:kinsoku",
                "<w:autoSpaceDE",
                "<w:adjustRightInd",
                "<w:snapToGrid",
                "<w:contextualSpacing",
                "<w:mirrorIndents",
                "<w:suppressOverlap",
                "<w:jc",
                "<w:textDirection",
                "<w:textAlignment",
                "<w:textboxTightWrap",
                "<w:outlineLvl",
                "<w:divId",
                "<w:cnfStyle",
            ],
        );
    }

    #[test]
    fn newly_modeled_paragraph_toggles_replay_their_source_carrier() {
        for (local, slot) in PPR_MODELED_TOGGLES {
            assert_eq!(
                ppr_modeled_slot(local, &[String::new()]),
                Some(slot),
                "a carrier toggle that is not a modeled slot underflows its occurrence"
            );
        }

        for element in [
            "suppressLineNumbers",
            "adjustRightInd",
            "contextualSpacing",
            "mirrorIndents",
            "suppressOverlap",
        ] {
            let ppr = parse_ppr(&format!(
                r#"<w:{element} xmlns:x="urn:producer" x:flag="kept"/>"#
            ));
            let value = match element {
                "suppressLineNumbers" => ppr.suppress_line_numbers,
                "adjustRightInd" => ppr.adjust_right_ind,
                "contextualSpacing" => ppr.contextual_spacing,
                "mirrorIndents" => ppr.mirror_indents,
                _ => ppr.suppress_overlap,
            };
            assert_eq!(value, Some(true), "{element}");

            let output = write_ppr(&ppr);
            assert_eq!(
                output,
                format!(r#"<w:pPr><w:{element} xmlns:x="urn:producer" x:flag="kept"/></w:pPr>"#),
                "{element}"
            );

            let reopened = parse_ppr(
                output
                    .strip_prefix("<w:pPr>")
                    .unwrap()
                    .strip_suffix("</w:pPr>")
                    .unwrap(),
            );
            assert_eq!(write_ppr(&reopened), output, "{element}");
        }
    }

    #[test]
    fn paragraph_property_parsing_accepts_alias_and_foreign_namespaces() {
        let xml = format!(
            concat!(
                r#"<x:pPr xmlns:x="{}" xmlns:ext="urn:producer"><x:framePr x:w="1440"/>"#,
                r#"<x:suppressLineNumbers/><x:contextualSpacing x:val="0"/>"#,
                r#"<x:textDirection x:val="tbRlV"/><x:textAlignment x:val="center"/>"#,
                r#"<x:textboxTightWrap x:val="lastLineOnly"/><x:divId x:val="9"/>"#,
                r#"<ext:contextualSpacing/><ext:textDirection ext:val="lrTb"/></x:pPr>"#,
            ),
            W_NS
        );
        let ppr = crate::numbering::parse_scoped_ppr(xml.as_bytes(), &["x".to_owned()]).unwrap();
        assert_eq!(
            ppr.frame.as_ref().and_then(|frame| frame.w),
            Some(Twips(1440))
        );
        assert_eq!(ppr.suppress_line_numbers, Some(true));
        assert_eq!(ppr.contextual_spacing, Some(false));
        assert_eq!(ppr.text_direction.as_deref(), Some("tbRlV"));
        assert_eq!(ppr.text_alignment.as_deref(), Some("center"));
        assert_eq!(ppr.textbox_tight_wrap.as_deref(), Some("lastLineOnly"));
        assert_eq!(ppr.div_id, Some(9));

        let output = write_ppr(&ppr);
        assert!(
            output.contains(r#"<ext:contextualSpacing xmlns:ext="urn:producer"/>"#),
            "{output}"
        );
        assert!(
            output.contains(r#"<ext:textDirection ext:val="lrTb" xmlns:ext="urn:producer"/>"#),
            "{output}"
        );
        assert!(output.contains(r#"<w:framePr w:w="1440"/>"#), "{output}");
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
