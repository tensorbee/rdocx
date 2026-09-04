//! Table elements: `CT_Tbl`, `CT_Row`, `CT_Tc` and related types.

use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};

use crate::borders::CT_BorderEdge;
use crate::content_control::{CT_Sdt, SdtOwner};
use crate::error::{OxmlError, Result};
use crate::namespace::matches_local_name;
use crate::numbering::{
    local_namespace_overrides, merged_owner_bindings, namespace_bindings, word_prefixes_at,
};
use crate::properties::{
    CT_Shd, get_val_attr, get_word_val_attr, is_word_attribute, is_word_element,
};
use crate::raw_xml::{capture_element, capture_empty_element};
use crate::revision::{CT_Revision, RevisionKind};
#[cfg(test)]
use crate::shared::ST_Border;
use crate::shared::ST_Jc;
use crate::text::CT_P;
use crate::units::Twips;

/// Write any captured raw XML that belongs immediately before position `pos`.
///
/// Table children we do not model are stored as `(position, raw)` pairs so
/// they can be put back where they were found, the same way `CT_P` handles
/// its own unknown children.
fn write_extras_at<W: std::io::Write>(
    writer: &mut Writer<W>,
    extra_xml: &[(usize, Vec<u8>)],
    pos: usize,
) -> Result<()> {
    for (at, raw) in extra_xml {
        if *at == pos {
            writer.get_mut().write_all(raw)?;
        }
    }
    Ok(())
}

fn write_boundary<W: std::io::Write>(
    writer: &mut Writer<W>,
    extra_xml: &[(usize, Vec<u8>)],
    controls: &[(usize, usize, CT_Sdt)],
    pos: usize,
) -> Result<()> {
    let extras = extra_xml
        .iter()
        .filter(|(at, _)| *at == pos)
        .map(|(_, raw)| raw)
        .collect::<Vec<_>>();
    for raw_before in 0..=extras.len() {
        for (_, _, control) in controls
            .iter()
            .filter(|(at, before, _)| *at == pos && *before == raw_before)
        {
            control.to_xml(writer)?;
        }
        if let Some(raw) = extras.get(raw_before) {
            writer.get_mut().write_all(raw)?;
        }
    }
    Ok(())
}

fn preserved_table_raw_bindings(scope: &[String]) -> Vec<(String, String)> {
    namespace_bindings(scope)
        .into_iter()
        .filter(|(prefix, namespace)| prefix != "w" || namespace != crate::namespace::W_NS)
        .collect()
}

// ---- Table border types ----

/// `CT_TblBorders` — Table-level borders.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct CT_TblBorders {
    pub top: Option<CT_BorderEdge>,
    pub bottom: Option<CT_BorderEdge>,
    pub left: Option<CT_BorderEdge>,
    pub right: Option<CT_BorderEdge>,
    pub inside_h: Option<CT_BorderEdge>,
    pub inside_v: Option<CT_BorderEdge>,
    /// Other border edges retained verbatim.
    pub extra_xml: Vec<Vec<u8>>,
}

impl CT_TblBorders {
    pub fn from_xml(reader: &mut Reader<&[u8]>) -> Result<Self> {
        Self::from_xml_with_prefixes_and_owner_bindings(reader, &["w".to_owned()], &[])
    }

    fn from_xml_with_prefixes_and_owner_bindings(
        reader: &mut Reader<&[u8]>,
        word_prefixes: &[String],
        owner_bindings: &[(String, String)],
    ) -> Result<Self> {
        let mut borders = CT_TblBorders::default();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(ref e)) => {
                    let name = e.name();
                    let prefixes = word_prefixes_at(e, word_prefixes)?;
                    if is_word_element(name.as_ref(), b"top", &prefixes) {
                        borders.top =
                            Some(CT_BorderEdge::from_xml_attrs_with_prefixes(e, &prefixes)?);
                    } else if is_word_element(name.as_ref(), b"bottom", &prefixes) {
                        borders.bottom =
                            Some(CT_BorderEdge::from_xml_attrs_with_prefixes(e, &prefixes)?);
                    } else if is_word_element(name.as_ref(), b"left", &prefixes)
                        || is_word_element(name.as_ref(), b"start", &prefixes)
                    {
                        borders.left =
                            Some(CT_BorderEdge::from_xml_attrs_with_prefixes(e, &prefixes)?);
                    } else if is_word_element(name.as_ref(), b"right", &prefixes)
                        || is_word_element(name.as_ref(), b"end", &prefixes)
                    {
                        borders.right =
                            Some(CT_BorderEdge::from_xml_attrs_with_prefixes(e, &prefixes)?);
                    } else if is_word_element(name.as_ref(), b"insideH", &prefixes) {
                        borders.inside_h =
                            Some(CT_BorderEdge::from_xml_attrs_with_prefixes(e, &prefixes)?);
                    } else if is_word_element(name.as_ref(), b"insideV", &prefixes) {
                        borders.inside_v =
                            Some(CT_BorderEdge::from_xml_attrs_with_prefixes(e, &prefixes)?);
                    } else {
                        borders
                            .extra_xml
                            .push(crate::text::raw_with_external_bindings(
                                &capture_empty_element(e)?,
                                owner_bindings,
                            )?);
                    }
                }
                Ok(Event::Start(ref e)) => {
                    borders
                        .extra_xml
                        .push(crate::text::raw_with_external_bindings(
                            &capture_element(reader, e)?,
                            owner_bindings,
                        )?)
                }
                Ok(Event::End(ref e))
                    if matches_local_name(e.name().as_ref(), b"tblBorders")
                        || matches_local_name(e.name().as_ref(), b"tcBorders") =>
                {
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(borders)
    }

    pub fn to_xml<W: std::io::Write>(&self, writer: &mut Writer<W>, tag: &str) -> Result<()> {
        writer.write_event(Event::Start(BytesStart::new(tag)))?;
        if let Some(ref e) = self.top {
            e.to_xml(writer, "w:top")?;
        }
        if let Some(ref e) = self.left {
            e.to_xml(writer, "w:left")?;
        }
        if let Some(ref e) = self.bottom {
            e.to_xml(writer, "w:bottom")?;
        }
        if let Some(ref e) = self.right {
            e.to_xml(writer, "w:right")?;
        }
        if let Some(ref e) = self.inside_h {
            e.to_xml(writer, "w:insideH")?;
        }
        if let Some(ref e) = self.inside_v {
            e.to_xml(writer, "w:insideV")?;
        }
        for raw in &self.extra_xml {
            writer.get_mut().write_all(raw)?;
        }
        writer.write_event(Event::End(BytesEnd::new(tag)))?;
        Ok(())
    }

    pub fn is_empty(&self) -> bool {
        self.top.is_none()
            && self.bottom.is_none()
            && self.left.is_none()
            && self.right.is_none()
            && self.inside_h.is_none()
            && self.inside_v.is_none()
            && self.extra_xml.is_empty()
    }
}

/// Table cell margin (a single edge width).
#[derive(Debug, Clone, Default, PartialEq)]
pub struct CT_TblCellMar {
    pub top: Option<Twips>,
    pub bottom: Option<Twips>,
    pub left: Option<Twips>,
    pub right: Option<Twips>,
}

fn parse_table_measurement(value: &str) -> Result<i32> {
    let integer = if let Some((integer, fraction)) = value.split_once('.') {
        if integer.is_empty() || fraction.is_empty() || !fraction.bytes().all(|byte| byte == b'0') {
            return Err(OxmlError::InvalidValue(format!(
                "unsupported table measurement: {value:?}"
            )));
        }
        integer
    } else {
        value
    };

    integer
        .parse::<i32>()
        .map_err(|_| OxmlError::InvalidValue(format!("invalid table measurement: {value:?}")))
}

impl CT_TblCellMar {
    fn parse_edge(e: &BytesStart, word_prefixes: &[String]) -> Result<Option<Twips>> {
        for attr in e.attributes() {
            let attr = attr?;
            if is_word_attribute(attr.key.as_ref(), b"w", word_prefixes) {
                let val = parse_table_measurement(std::str::from_utf8(&attr.value)?)?;
                return Ok(Some(Twips(val)));
            }
        }
        Ok(None)
    }

    pub fn from_xml(reader: &mut Reader<&[u8]>) -> Result<Self> {
        Self::from_xml_with_prefixes(reader, &["w".to_owned()])
    }

    fn from_xml_with_prefixes(
        reader: &mut Reader<&[u8]>,
        word_prefixes: &[String],
    ) -> Result<Self> {
        let mut mar = CT_TblCellMar::default();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(ref e)) => {
                    let name = e.name();
                    let prefixes = word_prefixes_at(e, word_prefixes)?;
                    if is_word_element(name.as_ref(), b"top", &prefixes) {
                        mar.top = Self::parse_edge(e, &prefixes)?;
                    } else if is_word_element(name.as_ref(), b"bottom", &prefixes) {
                        mar.bottom = Self::parse_edge(e, &prefixes)?;
                    } else if is_word_element(name.as_ref(), b"left", &prefixes)
                        || is_word_element(name.as_ref(), b"start", &prefixes)
                    {
                        mar.left = Self::parse_edge(e, &prefixes)?;
                    } else if is_word_element(name.as_ref(), b"right", &prefixes)
                        || is_word_element(name.as_ref(), b"end", &prefixes)
                    {
                        mar.right = Self::parse_edge(e, &prefixes)?;
                    }
                }
                Ok(Event::End(ref e)) if matches_local_name(e.name().as_ref(), b"tblCellMar") => {
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(mar)
    }

    pub fn to_xml<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_event(Event::Start(BytesStart::new("w:tblCellMar")))?;

        fn write_edge<W: std::io::Write>(
            writer: &mut Writer<W>,
            tag: &str,
            val: Twips,
        ) -> Result<()> {
            let mut buf = itoa::Buffer::new();
            let mut e = BytesStart::new(tag);
            e.push_attribute(("w:w", buf.format(val.0)));
            e.push_attribute(("w:type", "dxa"));
            writer.write_event(Event::Empty(e))?;
            Ok(())
        }

        if let Some(t) = self.top {
            write_edge(writer, "w:top", t)?;
        }
        if let Some(l) = self.left {
            write_edge(writer, "w:left", l)?;
        }
        if let Some(b) = self.bottom {
            write_edge(writer, "w:bottom", b)?;
        }
        if let Some(r) = self.right {
            write_edge(writer, "w:right", r)?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:tblCellMar")))?;
        Ok(())
    }
}

// ---- Table width ----

/// Table width specification.
#[derive(Debug, Clone, PartialEq)]
pub struct CT_TblWidth {
    /// Width value
    pub w: i32,
    /// Width type: "dxa" (twips), "pct" (50ths of a percent), "auto", "nil"
    pub width_type: String,
}

impl CT_TblWidth {
    pub fn dxa(twips: i32) -> Self {
        CT_TblWidth {
            w: twips,
            width_type: "dxa".to_string(),
        }
    }

    pub fn pct(fiftieths: i32) -> Self {
        CT_TblWidth {
            w: fiftieths,
            width_type: "pct".to_string(),
        }
    }

    pub fn auto() -> Self {
        CT_TblWidth {
            w: 0,
            width_type: "auto".to_string(),
        }
    }

    pub fn from_xml_attrs(e: &BytesStart) -> Result<Self> {
        let word_prefixes = word_prefixes_at(e, &["w".to_owned()])?;
        Self::from_xml_attrs_with_prefixes(e, &word_prefixes)
    }

    fn from_xml_attrs_with_prefixes(e: &BytesStart, word_prefixes: &[String]) -> Result<Self> {
        let mut w = 0;
        let mut width_type = "dxa".to_string();

        for attr in e.attributes() {
            let attr = attr?;
            let key = attr.key.as_ref();
            let val = std::str::from_utf8(&attr.value)?;
            if is_word_attribute(key, b"w", word_prefixes) {
                w = parse_table_measurement(val)?;
            } else if is_word_attribute(key, b"type", word_prefixes) {
                width_type = val.to_string();
            }
        }

        Ok(CT_TblWidth { w, width_type })
    }

    pub fn write_xml<W: std::io::Write>(&self, writer: &mut Writer<W>, tag: &str) -> Result<()> {
        let mut buf = itoa::Buffer::new();
        let mut e = BytesStart::new(tag);
        e.push_attribute(("w:w", buf.format(self.w)));
        e.push_attribute(("w:type", self.width_type.as_str()));
        writer.write_event(Event::Empty(e))?;
        Ok(())
    }
}

// ---- Table grid column ----

/// `CT_TblGridCol` — A column definition in the table grid.
#[derive(Debug, Clone, PartialEq)]
pub struct CT_TblGridCol {
    /// Column width in twips
    pub width: Twips,
}

// ---- Table properties ----

/// `CT_TblPr` — Table properties.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct CT_TblPr {
    /// Table style ID
    pub style_id: Option<String>,
    /// Table width
    pub width: Option<CT_TblWidth>,
    /// Table alignment
    pub jc: Option<ST_Jc>,
    /// Table borders
    pub borders: Option<CT_TblBorders>,
    /// Default cell margins
    pub cell_margin: Option<CT_TblCellMar>,
    /// Table layout: "fixed" or "autofit"
    pub layout: Option<String>,
    /// Table indent from left margin
    pub indent: Option<CT_TblWidth>,
    /// Table shading/background
    pub shading: Option<CT_Shd>,
    /// Which parts of the table style's conditional formatting apply.
    pub look: Option<CT_TblLook>,
    /// Prior table properties from the schema-final `w:tblPrChange`.
    pub change: Option<CT_Revision>,
    /// Malformed table property changes retained verbatim.
    pub revision_xml: Vec<Vec<u8>>,
    /// Unmodelled property children retained at their schema insertion slots.
    #[doc(hidden)]
    pub extra_xml: Vec<(usize, Vec<u8>)>,
}

/// `w:tblLook` — which parts of a table style's conditional formatting apply.
///
/// The style reference in `w:tblStyle` says *which* style to use. This says
/// which of its conditional parts to turn on: header row emphasis, banding,
/// first-column formatting. Dropping it leaves the style name intact and the
/// table rendered with base formatting only, which reads as the style having
/// been lost.
///
/// `w:val` is a legacy bitmask carrying the same information. Both are kept,
/// because writers disagree about which one to emit and readers disagree about
/// which one to trust.
#[derive(Debug, Clone, PartialEq, Default)]
#[allow(non_snake_case)]
pub struct CT_TblLook {
    /// Legacy bitmask form, e.g. "04A0".
    pub val: Option<String>,
    pub first_row: Option<bool>,
    pub last_row: Option<bool>,
    pub first_column: Option<bool>,
    pub last_column: Option<bool>,
    pub no_h_band: Option<bool>,
    pub no_v_band: Option<bool>,
}

/// Read an OOXML boolean attribute, which may be written as 1/0 or true/false.
fn parse_ooxml_bool(value: &str) -> Option<bool> {
    match value {
        "1" | "true" | "on" => Some(true),
        "0" | "false" | "off" => Some(false),
        _ => None,
    }
}

fn ooxml_bool_str(value: bool) -> &'static str {
    if value { "1" } else { "0" }
}

#[allow(non_snake_case)]
impl CT_TblLook {
    pub fn from_xml_attrs(e: &BytesStart) -> Result<Self> {
        let mut look = CT_TblLook::default();
        for attr in e.attributes().flatten() {
            let value = std::str::from_utf8(&attr.value)?;
            let key = attr.key.as_ref();
            if matches_local_name(key, b"val") {
                look.val = Some(value.to_string());
            } else if matches_local_name(key, b"firstRow") {
                look.first_row = parse_ooxml_bool(value);
            } else if matches_local_name(key, b"lastRow") {
                look.last_row = parse_ooxml_bool(value);
            } else if matches_local_name(key, b"firstColumn") {
                look.first_column = parse_ooxml_bool(value);
            } else if matches_local_name(key, b"lastColumn") {
                look.last_column = parse_ooxml_bool(value);
            } else if matches_local_name(key, b"noHBand") {
                look.no_h_band = parse_ooxml_bool(value);
            } else if matches_local_name(key, b"noVBand") {
                look.no_v_band = parse_ooxml_bool(value);
            }
        }
        Ok(look)
    }

    fn from_xml_attrs_with_prefixes(e: &BytesStart, word_prefixes: &[String]) -> Result<Self> {
        let mut look = CT_TblLook::default();
        for attr in e.attributes().flatten() {
            let value = std::str::from_utf8(&attr.value)?;
            let key = attr.key.as_ref();
            if is_word_attribute(key, b"val", word_prefixes) {
                look.val = Some(value.to_string());
            } else if is_word_attribute(key, b"firstRow", word_prefixes) {
                look.first_row = parse_ooxml_bool(value);
            } else if is_word_attribute(key, b"lastRow", word_prefixes) {
                look.last_row = parse_ooxml_bool(value);
            } else if is_word_attribute(key, b"firstColumn", word_prefixes) {
                look.first_column = parse_ooxml_bool(value);
            } else if is_word_attribute(key, b"lastColumn", word_prefixes) {
                look.last_column = parse_ooxml_bool(value);
            } else if is_word_attribute(key, b"noHBand", word_prefixes) {
                look.no_h_band = parse_ooxml_bool(value);
            } else if is_word_attribute(key, b"noVBand", word_prefixes) {
                look.no_v_band = parse_ooxml_bool(value);
            }
        }
        Ok(look)
    }

    pub fn to_xml<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut e = BytesStart::new("w:tblLook");
        if let Some(ref val) = self.val {
            e.push_attribute(("w:val", val.as_str()));
        }
        for (name, value) in [
            ("w:firstRow", self.first_row),
            ("w:lastRow", self.last_row),
            ("w:firstColumn", self.first_column),
            ("w:lastColumn", self.last_column),
            ("w:noHBand", self.no_h_band),
            ("w:noVBand", self.no_v_band),
        ] {
            if let Some(value) = value {
                e.push_attribute((name, ooxml_bool_str(value)));
            }
        }
        writer.write_event(Event::Empty(e))?;
        Ok(())
    }
}

#[allow(non_snake_case)]
impl CT_TblPr {
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
        let mut pr = CT_TblPr::default();
        let mut change_raw_index = 0usize;
        let mut boundary = 0usize;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(ref e)) => {
                    let name = e.name();
                    let prefixes = word_prefixes_at(e, word_prefixes)?;
                    let (at, next) = tbl_pr_raw_boundary(name.as_ref(), boundary, &prefixes);
                    if is_word_element(name.as_ref(), b"tblStyle", &prefixes) {
                        pr.style_id = get_word_val_attr(e, &prefixes)?;
                    } else if is_word_element(name.as_ref(), b"tblW", &prefixes) {
                        pr.width = Some(CT_TblWidth::from_xml_attrs_with_prefixes(e, &prefixes)?);
                    } else if is_word_element(name.as_ref(), b"jc", &prefixes) {
                        if let Some(val) = get_word_val_attr(e, &prefixes)? {
                            pr.jc = ST_Jc::from_str(&val).ok();
                        }
                    } else if is_word_element(name.as_ref(), b"tblLayout", &prefixes) {
                        for attribute in e.attributes() {
                            let attribute = attribute?;
                            if is_word_attribute(attribute.key.as_ref(), b"type", &prefixes) {
                                pr.layout =
                                    Some(std::str::from_utf8(&attribute.value)?.to_string());
                                break;
                            }
                        }
                    } else if is_word_element(name.as_ref(), b"tblInd", &prefixes) {
                        pr.indent = Some(CT_TblWidth::from_xml_attrs_with_prefixes(e, &prefixes)?);
                    } else if is_word_element(name.as_ref(), b"shd", &prefixes) {
                        pr.shading = Some(CT_Shd::from_xml_attrs_with_prefixes(e, &prefixes)?);
                    } else if is_word_element(name.as_ref(), b"tblLook", &prefixes) {
                        pr.look = Some(CT_TblLook::from_xml_attrs_with_prefixes(e, &prefixes)?);
                    } else if is_word_element(name.as_ref(), b"tblPrChange", &prefixes) {
                        let raw = crate::text::raw_with_external_bindings(
                            &capture_empty_element(e)?,
                            owner_bindings,
                        )?;
                        if let Some(revision) = CT_Revision::from_raw(raw.clone(), &prefixes) {
                            if let Some(previous) = pr.change.replace(revision) {
                                pr.revision_xml
                                    .insert(change_raw_index, previous.into_raw_xml());
                            }
                            change_raw_index = pr.revision_xml.len();
                        } else {
                            pr.revision_xml.push(raw);
                        }
                    } else if matches_local_name(name.as_ref(), b"tblPrChange") {
                        pr.revision_xml
                            .push(crate::text::raw_with_external_bindings(
                                &capture_empty_element(e)?,
                                owner_bindings,
                            )?);
                    } else {
                        pr.extra_xml.push((
                            at,
                            crate::text::raw_with_external_bindings(
                                &capture_empty_element(e)?,
                                owner_bindings,
                            )?,
                        ));
                    }
                    boundary = next;
                }
                Ok(Event::Start(ref e)) => {
                    let name = e.name();
                    let prefixes = word_prefixes_at(e, word_prefixes)?;
                    let (at, next) = tbl_pr_raw_boundary(name.as_ref(), boundary, &prefixes);
                    if is_word_element(name.as_ref(), b"tblBorders", &prefixes) {
                        let local_bindings = local_namespace_overrides(e, word_prefixes)?;
                        let border_bindings =
                            merged_owner_bindings(owner_bindings, &local_bindings);
                        pr.borders =
                            Some(CT_TblBorders::from_xml_with_prefixes_and_owner_bindings(
                                reader,
                                &prefixes,
                                &border_bindings,
                            )?);
                    } else if is_word_element(name.as_ref(), b"tblCellMar", &prefixes) {
                        pr.cell_margin =
                            Some(CT_TblCellMar::from_xml_with_prefixes(reader, &prefixes)?);
                    } else if is_word_element(name.as_ref(), b"tblPrChange", &prefixes) {
                        let raw = crate::text::raw_with_external_bindings(
                            &capture_element(reader, e)?,
                            owner_bindings,
                        )?;
                        if let Some(revision) = CT_Revision::from_raw(raw.clone(), &prefixes) {
                            if let Some(previous) = pr.change.replace(revision) {
                                pr.revision_xml
                                    .insert(change_raw_index, previous.into_raw_xml());
                            }
                            change_raw_index = pr.revision_xml.len();
                        } else {
                            pr.revision_xml.push(raw);
                        }
                    } else if matches_local_name(name.as_ref(), b"tblPrChange") {
                        pr.revision_xml
                            .push(crate::text::raw_with_external_bindings(
                                &capture_element(reader, e)?,
                                owner_bindings,
                            )?);
                    } else {
                        pr.extra_xml.push((
                            at,
                            crate::text::raw_with_external_bindings(
                                &capture_element(reader, e)?,
                                owner_bindings,
                            )?,
                        ));
                    }
                    boundary = next;
                }
                Ok(Event::End(ref e)) if matches_local_name(e.name().as_ref(), b"tblPr") => {
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(pr)
    }

    pub fn to_xml<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_event(Event::Start(BytesStart::new("w:tblPr")))?;

        write_extras_at(writer, &self.extra_xml, 0)?;
        if let Some(ref style_id) = self.style_id {
            let mut e = BytesStart::new("w:tblStyle");
            e.push_attribute(("w:val", style_id.as_str()));
            writer.write_event(Event::Empty(e))?;
        }

        for boundary in 1..=6 {
            write_extras_at(writer, &self.extra_xml, boundary)?;
        }
        if let Some(ref width) = self.width {
            width.write_xml(writer, "w:tblW")?;
        }

        write_extras_at(writer, &self.extra_xml, 7)?;
        if let Some(jc) = self.jc {
            let mut e = BytesStart::new("w:jc");
            e.push_attribute(("w:val", jc.to_str()));
            writer.write_event(Event::Empty(e))?;
        }

        write_extras_at(writer, &self.extra_xml, 8)?;
        write_extras_at(writer, &self.extra_xml, 9)?;
        if let Some(ref indent) = self.indent {
            indent.write_xml(writer, "w:tblInd")?;
        }

        write_extras_at(writer, &self.extra_xml, 10)?;
        if let Some(ref borders) = self.borders
            && !borders.is_empty()
        {
            borders.to_xml(writer, "w:tblBorders")?;
        }

        write_extras_at(writer, &self.extra_xml, 11)?;
        if let Some(ref shd) = self.shading {
            shd.write_xml(writer, "w:shd")?;
        }

        write_extras_at(writer, &self.extra_xml, 12)?;
        if let Some(ref layout) = self.layout {
            let mut e = BytesStart::new("w:tblLayout");
            e.push_attribute(("w:type", layout.as_str()));
            writer.write_event(Event::Empty(e))?;
        }

        write_extras_at(writer, &self.extra_xml, 13)?;
        if let Some(ref cell_margin) = self.cell_margin {
            cell_margin.to_xml(writer)?;
        }

        write_extras_at(writer, &self.extra_xml, 14)?;
        if let Some(ref look) = self.look {
            look.to_xml(writer)?;
        }

        for boundary in 15..=18 {
            write_extras_at(writer, &self.extra_xml, boundary)?;
        }
        for raw in &self.revision_xml {
            writer.get_mut().write_all(raw)?;
        }
        if let Some(change) = &self.change {
            change.write_xml(writer)?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:tblPr")))?;
        Ok(())
    }
}

fn tbl_pr_raw_boundary(name: &[u8], current: usize, word_prefixes: &[String]) -> (usize, usize) {
    let schema_children = [
        b"tblStyle".as_slice(),
        b"tblpPr".as_slice(),
        b"tblOverlap".as_slice(),
        b"bidiVisual".as_slice(),
        b"tblStyleRowBandSize".as_slice(),
        b"tblStyleColBandSize".as_slice(),
        b"tblW".as_slice(),
        b"jc".as_slice(),
        b"tblCellSpacing".as_slice(),
        b"tblInd".as_slice(),
        b"tblBorders".as_slice(),
        b"shd".as_slice(),
        b"tblLayout".as_slice(),
        b"tblCellMar".as_slice(),
        b"tblLook".as_slice(),
        b"tblCaption".as_slice(),
        b"tblDescription".as_slice(),
        b"tblPrChange".as_slice(),
    ];

    schema_children
        .iter()
        .position(|child| is_word_element(name, child, word_prefixes))
        .map_or((current, current), |index| (index, index + 1))
}

// ---- Table grid ----

/// `CT_TblGrid` — Defines the column structure of a table.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct CT_TblGrid {
    /// Active columns used for current table layout.
    pub columns: Vec<CT_TblGridCol>,
    /// The preserved historical `w:tblGridChange` subtree, when present.
    ///
    /// Historical columns are retained for round-trip fidelity and never
    /// participate in current table layout.
    pub grid_change_xml: Option<Vec<u8>>,
    /// Unmodelled direct children retained verbatim.
    pub extra_xml: Vec<Vec<u8>>,
}

#[allow(non_snake_case)]
impl CT_TblGrid {
    pub fn from_xml(reader: &mut Reader<&[u8]>) -> Result<Self> {
        Self::from_xml_with_prefixes(reader, &["w".to_owned()])
    }

    pub(crate) fn from_xml_with_prefixes(
        reader: &mut Reader<&[u8]>,
        word_prefixes: &[String],
    ) -> Result<Self> {
        let mut columns = Vec::new();
        let mut grid_change_xml = None;
        let mut extra_xml = Vec::new();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(ref e)) => {
                    let prefixes = word_prefixes_at(e, word_prefixes)?;
                    if is_word_element(e.name().as_ref(), b"gridCol", &prefixes) {
                        columns.push(Self::parse_grid_column(e, &prefixes)?);
                    } else if is_word_element(e.name().as_ref(), b"tblGridChange", &prefixes) {
                        if grid_change_xml.is_some() {
                            return Err(OxmlError::InvalidValue(
                                "duplicate w:tblGridChange".to_owned(),
                            ));
                        }
                        let raw = capture_empty_element(e)?;
                        grid_change_xml = Some(crate::text::raw_with_external_bindings(
                            &raw,
                            &preserved_table_raw_bindings(&prefixes),
                        )?);
                    } else {
                        let raw = capture_empty_element(e)?;
                        extra_xml.push(crate::text::raw_with_external_bindings(
                            &raw,
                            &preserved_table_raw_bindings(&prefixes),
                        )?);
                    }
                }
                Ok(Event::Start(ref e)) => {
                    let prefixes = word_prefixes_at(e, word_prefixes)?;
                    if is_word_element(e.name().as_ref(), b"gridCol", &prefixes) {
                        columns.push(Self::parse_grid_column(e, &prefixes)?);
                        capture_element(reader, e)?;
                    } else if is_word_element(e.name().as_ref(), b"tblGridChange", &prefixes) {
                        if grid_change_xml.is_some() {
                            return Err(OxmlError::InvalidValue(
                                "duplicate w:tblGridChange".to_owned(),
                            ));
                        }
                        let raw = capture_element(reader, e)?;
                        grid_change_xml = Some(crate::text::raw_with_external_bindings(
                            &raw,
                            &preserved_table_raw_bindings(&prefixes),
                        )?);
                    } else {
                        let raw = capture_element(reader, e)?;
                        extra_xml.push(crate::text::raw_with_external_bindings(
                            &raw,
                            &preserved_table_raw_bindings(&prefixes),
                        )?);
                    }
                }
                Ok(Event::End(ref e))
                    if is_word_element(e.name().as_ref(), b"tblGrid", word_prefixes) =>
                {
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(CT_TblGrid {
            columns,
            grid_change_xml,
            extra_xml,
        })
    }

    fn parse_grid_column(
        element: &BytesStart<'_>,
        word_prefixes: &[String],
    ) -> Result<CT_TblGridCol> {
        let mut width = Twips(0);
        for attr in element.attributes() {
            let attr = attr?;
            if is_word_attribute(attr.key.as_ref(), b"w", word_prefixes) {
                width = Twips(std::str::from_utf8(&attr.value)?.parse()?);
            }
        }
        Ok(CT_TblGridCol { width })
    }

    pub fn to_xml<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut buf = itoa::Buffer::new();
        writer.write_event(Event::Start(BytesStart::new("w:tblGrid")))?;

        for col in &self.columns {
            let mut e = BytesStart::new("w:gridCol");
            e.push_attribute(("w:w", buf.format(col.width.0)));
            writer.write_event(Event::Empty(e))?;
        }

        for raw in &self.extra_xml {
            writer.get_mut().write_all(raw)?;
        }
        if let Some(raw) = &self.grid_change_xml {
            writer.get_mut().write_all(raw)?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:tblGrid")))?;
        Ok(())
    }
}

// ---- Row properties ----

/// Vertical merge state for a cell.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VMerge {
    /// Start of a vertical merge group
    Restart,
    /// Continuation of the merge group above
    Continue,
}

/// `CT_TrPr` — Table row properties.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct CT_TrPr {
    /// Row height in twips
    pub height: Option<Twips>,
    /// Row height rule: "exact" or "atLeast"
    pub height_rule: Option<String>,
    /// Repeat as header row on each page
    pub header: Option<bool>,
    /// Row alignment
    pub jc: Option<ST_Jc>,
    /// Number of grid columns omitted before the first cell.
    pub grid_before: Option<u32>,
    /// Number of grid columns omitted after the final cell.
    pub grid_after: Option<u32>,
    /// Allow row to break across pages
    pub cant_split: Option<bool>,
    /// `w:cnfStyle` — which conditional parts of the table style this row is.
    ///
    /// Word writes this alongside `w:tblLook` and needs both to reproduce a
    /// styled table. Dropping it loses the header-row and banding emphasis.
    pub cnf_style: Option<String>,
    /// Contextual row insertion and deletion markers in schema order.
    pub revision_markers: Vec<CT_Revision>,
    /// Malformed row markers retained verbatim.
    pub revision_xml: Vec<Vec<u8>>,
    /// Schema slots for malformed row markers, parallel to `revision_xml`.
    #[doc(hidden)]
    pub revision_xml_positions: Vec<usize>,
    /// Other row properties retained at their schema insertion slots.
    pub extra_xml: Vec<(usize, Vec<u8>)>,
}

#[allow(non_snake_case)]
impl CT_TrPr {
    pub fn from_xml(reader: &mut Reader<&[u8]>) -> Result<Self> {
        Self::from_xml_with_prefixes_and_owner_bindings(reader, &["w".to_owned()], &[])
    }

    fn from_xml_with_prefixes_and_owner_bindings(
        reader: &mut Reader<&[u8]>,
        word_prefixes: &[String],
        owner_bindings: &[(String, String)],
    ) -> Result<Self> {
        let mut pr = CT_TrPr::default();
        let mut boundary = 0usize;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(ref e)) => {
                    let name = e.name();
                    let prefixes = word_prefixes_at(e, word_prefixes)?;
                    let (at, next) = tr_pr_raw_boundary(name.as_ref(), boundary, &prefixes);
                    if matches_local_name(name.as_ref(), b"trHeight") {
                        for attr in e.attributes() {
                            let attr = attr?;
                            let key = attr.key.as_ref();
                            let val = std::str::from_utf8(&attr.value)?;
                            if matches_local_name(key, b"val") {
                                pr.height = Some(Twips(val.parse()?));
                            } else if matches_local_name(key, b"hRule") {
                                pr.height_rule = Some(val.to_string());
                            }
                        }
                    } else if matches_local_name(name.as_ref(), b"tblHeader") {
                        pr.header = Some(true);
                    } else if matches_local_name(name.as_ref(), b"jc") {
                        if let Some(val) = get_val_attr(e)? {
                            pr.jc = ST_Jc::from_str(&val).ok();
                        }
                    } else if is_word_element(name.as_ref(), b"gridBefore", &prefixes) {
                        pr.grid_before = get_word_val_attr(e, &prefixes)?
                            .map(|value| value.parse())
                            .transpose()?;
                    } else if is_word_element(name.as_ref(), b"gridAfter", &prefixes) {
                        pr.grid_after = get_word_val_attr(e, &prefixes)?
                            .map(|value| value.parse())
                            .transpose()?;
                    } else if matches_local_name(name.as_ref(), b"cnfStyle") {
                        pr.cnf_style = get_val_attr(e)?;
                    } else if matches_local_name(name.as_ref(), b"cantSplit") {
                        pr.cant_split = Some(true);
                    } else if is_word_element(name.as_ref(), b"ins", &prefixes)
                        || is_word_element(name.as_ref(), b"del", &prefixes)
                    {
                        let raw = crate::text::raw_with_external_bindings(
                            &capture_empty_element(e)?,
                            owner_bindings,
                        )?;
                        if let Some(revision) = CT_Revision::from_raw(raw.clone(), &prefixes) {
                            pr.revision_markers.push(revision);
                        } else {
                            pr.revision_xml.push(raw);
                            pr.revision_xml_positions.push(at);
                        }
                    } else {
                        pr.extra_xml.push((
                            at,
                            crate::text::raw_with_external_bindings(
                                &capture_empty_element(e)?,
                                owner_bindings,
                            )?,
                        ));
                    }
                    boundary = next;
                }
                Ok(Event::Start(ref e)) => {
                    let prefixes = word_prefixes_at(e, word_prefixes)?;
                    let (at, next) = tr_pr_raw_boundary(e.name().as_ref(), boundary, &prefixes);
                    if is_word_element(e.name().as_ref(), b"ins", &prefixes)
                        || is_word_element(e.name().as_ref(), b"del", &prefixes)
                    {
                        let raw = crate::text::raw_with_external_bindings(
                            &capture_element(reader, e)?,
                            owner_bindings,
                        )?;
                        if let Some(revision) = CT_Revision::from_raw(raw.clone(), &prefixes) {
                            pr.revision_markers.push(revision);
                        } else {
                            pr.revision_xml.push(raw);
                            pr.revision_xml_positions.push(at);
                        }
                    } else {
                        pr.extra_xml.push((
                            at,
                            crate::text::raw_with_external_bindings(
                                &capture_element(reader, e)?,
                                owner_bindings,
                            )?,
                        ));
                    }
                    boundary = next;
                }
                Ok(Event::End(ref e)) if matches_local_name(e.name().as_ref(), b"trPr") => {
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(pr)
    }

    pub fn to_xml<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        if self.is_empty() {
            return Ok(());
        }

        writer.write_event(Event::Start(BytesStart::new("w:trPr")))?;

        // cnfStyle comes first in the schema sequence for both trPr and tcPr.
        write_extras_at(writer, &self.extra_xml, 0)?;
        if let Some(ref cnf) = self.cnf_style {
            let mut e = BytesStart::new("w:cnfStyle");
            e.push_attribute(("w:val", cnf.as_str()));
            writer.write_event(Event::Empty(e))?;
        }

        write_extras_at(writer, &self.extra_xml, 1)?;
        write_extras_at(writer, &self.extra_xml, 2)?;
        if let Some(grid_before) = self.grid_before {
            let mut buffer = itoa::Buffer::new();
            let mut e = BytesStart::new("w:gridBefore");
            e.push_attribute(("w:val", buffer.format(grid_before)));
            writer.write_event(Event::Empty(e))?;
        }

        write_extras_at(writer, &self.extra_xml, 3)?;
        if let Some(grid_after) = self.grid_after {
            let mut buffer = itoa::Buffer::new();
            let mut e = BytesStart::new("w:gridAfter");
            e.push_attribute(("w:val", buffer.format(grid_after)));
            writer.write_event(Event::Empty(e))?;
        }

        write_extras_at(writer, &self.extra_xml, 4)?;
        write_extras_at(writer, &self.extra_xml, 5)?;
        write_extras_at(writer, &self.extra_xml, 6)?;
        if let Some(ref cant_split) = self.cant_split
            && *cant_split
        {
            writer.write_event(Event::Empty(BytesStart::new("w:cantSplit")))?;
        }

        write_extras_at(writer, &self.extra_xml, 7)?;
        if let Some(height) = self.height {
            let mut buf = itoa::Buffer::new();
            let mut e = BytesStart::new("w:trHeight");
            e.push_attribute(("w:val", buf.format(height.0)));
            if let Some(ref rule) = self.height_rule {
                e.push_attribute(("w:hRule", rule.as_str()));
            }
            writer.write_event(Event::Empty(e))?;
        }

        write_extras_at(writer, &self.extra_xml, 8)?;
        if let Some(true) = self.header {
            writer.write_event(Event::Empty(BytesStart::new("w:tblHeader")))?;
        }

        write_extras_at(writer, &self.extra_xml, 9)?;
        write_extras_at(writer, &self.extra_xml, 10)?;
        if let Some(jc) = self.jc {
            let mut e = BytesStart::new("w:jc");
            e.push_attribute(("w:val", jc.to_str()));
            writer.write_event(Event::Empty(e))?;
        }

        write_extras_at(writer, &self.extra_xml, 11)?;
        write_extras_at(writer, &self.extra_xml, 12)?;
        self.write_revision_xml_at(writer, 12)?;
        for revision in self
            .revision_markers
            .iter()
            .filter(|revision| revision.kind() == RevisionKind::Insertion)
        {
            revision.write_xml(writer)?;
        }
        write_extras_at(writer, &self.extra_xml, 13)?;
        self.write_revision_xml_at(writer, 13)?;
        for revision in self
            .revision_markers
            .iter()
            .filter(|revision| revision.kind() == RevisionKind::Deletion)
        {
            revision.write_xml(writer)?;
        }
        if self.revision_xml_positions.len() != self.revision_xml.len() {
            for raw in &self.revision_xml {
                writer.get_mut().write_all(raw)?;
            }
        }
        for revision in self.revision_markers.iter().filter(|revision| {
            !matches!(
                revision.kind(),
                RevisionKind::Insertion | RevisionKind::Deletion
            )
        }) {
            revision.write_xml(writer)?;
        }
        write_extras_at(writer, &self.extra_xml, 14)?;
        write_extras_at(writer, &self.extra_xml, 15)?;

        writer.write_event(Event::End(BytesEnd::new("w:trPr")))?;
        Ok(())
    }

    fn is_empty(&self) -> bool {
        self.height.is_none()
            && self.header.is_none()
            && self.jc.is_none()
            && self.cant_split.is_none()
            && self.cnf_style.is_none()
            && self.grid_before.is_none()
            && self.grid_after.is_none()
            && self.revision_markers.is_empty()
            && self.revision_xml.is_empty()
            && self.revision_xml_positions.is_empty()
            && self.extra_xml.is_empty()
    }

    fn write_revision_xml_at<W: std::io::Write>(
        &self,
        writer: &mut Writer<W>,
        position: usize,
    ) -> Result<()> {
        if self.revision_xml_positions.len() == self.revision_xml.len() {
            for (raw, _) in self
                .revision_xml
                .iter()
                .zip(&self.revision_xml_positions)
                .filter(|(_, at)| **at == position)
            {
                writer.get_mut().write_all(raw)?;
            }
        }
        Ok(())
    }
}

fn tr_pr_raw_boundary(name: &[u8], current: usize, word_prefixes: &[String]) -> (usize, usize) {
    let schema_children = [
        b"cnfStyle".as_slice(),
        b"divId".as_slice(),
        b"gridBefore".as_slice(),
        b"gridAfter".as_slice(),
        b"wBefore".as_slice(),
        b"wAfter".as_slice(),
        b"cantSplit".as_slice(),
        b"trHeight".as_slice(),
        b"tblHeader".as_slice(),
        b"tblCellSpacing".as_slice(),
        b"jc".as_slice(),
        b"hidden".as_slice(),
        b"ins".as_slice(),
        b"del".as_slice(),
        b"trPrChange".as_slice(),
    ];

    schema_children
        .iter()
        .position(|child| is_word_element(name, child, word_prefixes))
        .map_or((current, current), |index| (index, index + 1))
}

// ---- Cell properties ----

/// Vertical alignment within a cell.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ST_VerticalJc {
    Top,
    Center,
    Bottom,
}

impl ST_VerticalJc {
    pub fn from_str(s: &str) -> Self {
        match s {
            "center" => Self::Center,
            "bottom" => Self::Bottom,
            _ => Self::Top,
        }
    }

    pub fn to_str(self) -> &'static str {
        match self {
            Self::Top => "top",
            Self::Center => "center",
            Self::Bottom => "bottom",
        }
    }
}

/// `CT_TcPr` — Table cell properties.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct CT_TcPr {
    /// Cell width
    pub width: Option<CT_TblWidth>,
    /// Horizontal merge (number of grid columns spanned)
    pub grid_span: Option<u32>,
    /// Legacy horizontal merge state.
    pub h_merge: Option<String>,
    /// Vertical merge
    pub v_merge: Option<VMerge>,
    /// Cell borders
    pub borders: Option<CT_TblBorders>,
    /// Cell shading
    pub shading: Option<CT_Shd>,
    /// Vertical alignment
    pub v_align: Option<ST_VerticalJc>,
    /// No-wrap text
    pub no_wrap: Option<bool>,
    /// Text direction
    pub text_direction: Option<String>,
    /// `w:cnfStyle` — which conditional parts of the table style this cell is.
    pub cnf_style: Option<String>,
    /// Unmodelled property children retained at their schema insertion slots.
    #[doc(hidden)]
    pub extra_xml: Vec<(usize, Vec<u8>)>,
}

#[allow(non_snake_case)]
impl CT_TcPr {
    pub fn from_xml(reader: &mut Reader<&[u8]>) -> Result<Self> {
        Self::from_xml_with_prefixes_and_owner_bindings(reader, &["w".to_owned()], &[])
    }

    pub(crate) fn from_xml_with_prefixes(
        reader: &mut Reader<&[u8]>,
        word_prefixes: &[String],
    ) -> Result<Self> {
        Self::from_xml_with_prefixes_and_owner_bindings(reader, word_prefixes, &[])
    }

    fn from_xml_with_prefixes_and_owner_bindings(
        reader: &mut Reader<&[u8]>,
        word_prefixes: &[String],
        owner_bindings: &[(String, String)],
    ) -> Result<Self> {
        let mut pr = CT_TcPr::default();
        let mut boundary = 0;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(ref e)) => {
                    let prefixes = word_prefixes_at(e, word_prefixes)?;
                    let (at, next) = tc_pr_raw_boundary(e.name().as_ref(), boundary, &prefixes);
                    if Self::parse_property_element(e, &mut pr, &prefixes)? {
                        boundary = next;
                    } else {
                        pr.extra_xml.push((
                            at,
                            crate::text::raw_with_external_bindings(
                                &capture_empty_element(e)?,
                                owner_bindings,
                            )?,
                        ));
                        boundary = next;
                    }
                }
                Ok(Event::Start(ref e)) => {
                    let name = e.name();
                    let prefixes = word_prefixes_at(e, word_prefixes)?;
                    let (at, next) = tc_pr_raw_boundary(name.as_ref(), boundary, &prefixes);
                    if is_word_element(name.as_ref(), b"tcBorders", &prefixes) {
                        let local_bindings = local_namespace_overrides(e, word_prefixes)?;
                        let border_bindings =
                            merged_owner_bindings(owner_bindings, &local_bindings);
                        pr.borders =
                            Some(CT_TblBorders::from_xml_with_prefixes_and_owner_bindings(
                                reader,
                                &prefixes,
                                &border_bindings,
                            )?);
                        boundary = next;
                    } else if Self::parse_property_element(e, &mut pr, &prefixes)? {
                        reader.read_to_end_into(name, &mut Vec::new())?;
                        boundary = next;
                    } else {
                        pr.extra_xml.push((
                            at,
                            crate::text::raw_with_external_bindings(
                                &capture_element(reader, e)?,
                                owner_bindings,
                            )?,
                        ));
                        boundary = next;
                    }
                }
                Ok(Event::End(ref e)) if matches_local_name(e.name().as_ref(), b"tcPr") => {
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(pr)
    }

    fn parse_property_element(
        e: &BytesStart<'_>,
        pr: &mut Self,
        word_prefixes: &[String],
    ) -> Result<bool> {
        let name = e.name();
        if is_word_element(name.as_ref(), b"tcW", word_prefixes) {
            pr.width = Some(CT_TblWidth::from_xml_attrs_with_prefixes(e, word_prefixes)?);
        } else if is_word_element(name.as_ref(), b"gridSpan", word_prefixes) {
            if let Some(val) = get_word_val_attr(e, word_prefixes)? {
                let span: u32 = val.parse()?;
                if span == 0 {
                    return Err(OxmlError::InvalidValue(
                        "w:gridSpan must be positive".to_owned(),
                    ));
                }
                pr.grid_span = Some(span);
            }
        } else if is_word_element(name.as_ref(), b"hMerge", word_prefixes) {
            pr.h_merge = Some(get_word_val_attr(e, word_prefixes)?.unwrap_or_default());
        } else if is_word_element(name.as_ref(), b"vMerge", word_prefixes) {
            pr.v_merge = Some(match get_word_val_attr(e, word_prefixes)?.as_deref() {
                Some("restart") => VMerge::Restart,
                Some("continue") | None => VMerge::Continue,
                Some(value) => {
                    return Err(OxmlError::InvalidValue(format!(
                        "invalid w:vMerge value {value}"
                    )));
                }
            });
        } else if is_word_element(name.as_ref(), b"vAlign", word_prefixes) {
            if let Some(val) = get_word_val_attr(e, word_prefixes)? {
                pr.v_align = Some(ST_VerticalJc::from_str(&val));
            }
        } else if is_word_element(name.as_ref(), b"shd", word_prefixes) {
            pr.shading = Some(CT_Shd::from_xml_attrs_with_prefixes(e, word_prefixes)?);
        } else if is_word_element(name.as_ref(), b"cnfStyle", word_prefixes) {
            pr.cnf_style = get_word_val_attr(e, word_prefixes)?;
        } else if is_word_element(name.as_ref(), b"noWrap", word_prefixes) {
            pr.no_wrap = Some(true);
        } else if is_word_element(name.as_ref(), b"textDirection", word_prefixes)
            && let Some(val) = get_word_val_attr(e, word_prefixes)?
        {
            pr.text_direction = Some(val);
        } else {
            return Ok(false);
        }
        Ok(true)
    }

    pub fn to_xml<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        if self.is_empty() {
            return Ok(());
        }

        writer.write_event(Event::Start(BytesStart::new("w:tcPr")))?;

        write_extras_at(writer, &self.extra_xml, 0)?;
        if let Some(ref cnf) = self.cnf_style {
            let mut e = BytesStart::new("w:cnfStyle");
            e.push_attribute(("w:val", cnf.as_str()));
            writer.write_event(Event::Empty(e))?;
        }

        write_extras_at(writer, &self.extra_xml, 1)?;
        if let Some(ref width) = self.width {
            width.write_xml(writer, "w:tcW")?;
        }

        write_extras_at(writer, &self.extra_xml, 2)?;
        if let Some(grid_span) = self.grid_span
            && grid_span > 1
        {
            let mut buf = itoa::Buffer::new();
            let mut e = BytesStart::new("w:gridSpan");
            e.push_attribute(("w:val", buf.format(grid_span)));
            writer.write_event(Event::Empty(e))?;
        }

        // Unmodelled w:hMerge is preserved at boundary 3.
        write_extras_at(writer, &self.extra_xml, 3)?;
        if let Some(h_merge) = &self.h_merge {
            let mut e = BytesStart::new("w:hMerge");
            if !h_merge.is_empty() {
                e.push_attribute(("w:val", h_merge.as_str()));
            }
            writer.write_event(Event::Empty(e))?;
        }
        write_extras_at(writer, &self.extra_xml, 4)?;
        if let Some(ref vm) = self.v_merge {
            let mut e = BytesStart::new("w:vMerge");
            match vm {
                VMerge::Restart => e.push_attribute(("w:val", "restart")),
                VMerge::Continue => {} // empty element
            }
            writer.write_event(Event::Empty(e))?;
        }

        write_extras_at(writer, &self.extra_xml, 5)?;
        if let Some(ref borders) = self.borders
            && !borders.is_empty()
        {
            borders.to_xml(writer, "w:tcBorders")?;
        }

        write_extras_at(writer, &self.extra_xml, 6)?;
        if let Some(ref shd) = self.shading {
            shd.write_xml(writer, "w:shd")?;
        }

        write_extras_at(writer, &self.extra_xml, 7)?;
        if let Some(true) = self.no_wrap {
            writer.write_event(Event::Empty(BytesStart::new("w:noWrap")))?;
        }

        // Unmodelled w:tcMar is preserved at boundary 8.
        write_extras_at(writer, &self.extra_xml, 8)?;
        write_extras_at(writer, &self.extra_xml, 9)?;
        if let Some(ref td) = self.text_direction {
            let mut e = BytesStart::new("w:textDirection");
            e.push_attribute(("w:val", td.as_str()));
            writer.write_event(Event::Empty(e))?;
        }

        // Unmodelled w:tcFitText is preserved at boundary 10.
        write_extras_at(writer, &self.extra_xml, 10)?;
        write_extras_at(writer, &self.extra_xml, 11)?;
        if let Some(ref va) = self.v_align {
            let mut e = BytesStart::new("w:vAlign");
            e.push_attribute(("w:val", va.to_str()));
            writer.write_event(Event::Empty(e))?;
        }

        // The remaining standard Word children occupy boundaries 12 to 17.
        for boundary in 12..=18 {
            write_extras_at(writer, &self.extra_xml, boundary)?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:tcPr")))?;
        Ok(())
    }

    fn is_empty(&self) -> bool {
        self.width.is_none()
            && self.grid_span.is_none()
            && self.h_merge.is_none()
            && self.v_merge.is_none()
            && self.borders.is_none()
            && self.shading.is_none()
            && self.v_align.is_none()
            && self.no_wrap.is_none()
            && self.text_direction.is_none()
            && self.cnf_style.is_none()
            && self.extra_xml.is_empty()
    }
}

fn tc_pr_raw_boundary(name: &[u8], current: usize, word_prefixes: &[String]) -> (usize, usize) {
    if is_word_element(name, b"cnfStyle", word_prefixes) {
        (0, 1)
    } else if is_word_element(name, b"tcW", word_prefixes) {
        (1, 2)
    } else if is_word_element(name, b"gridSpan", word_prefixes) {
        (2, 3)
    } else if is_word_element(name, b"hMerge", word_prefixes) {
        (3, 4)
    } else if is_word_element(name, b"vMerge", word_prefixes) {
        (4, 5)
    } else if is_word_element(name, b"tcBorders", word_prefixes) {
        (5, 6)
    } else if is_word_element(name, b"shd", word_prefixes) {
        (6, 7)
    } else if is_word_element(name, b"noWrap", word_prefixes) {
        (7, 8)
    } else if is_word_element(name, b"tcMar", word_prefixes) {
        (8, 9)
    } else if is_word_element(name, b"textDirection", word_prefixes) {
        (9, 10)
    } else if is_word_element(name, b"tcFitText", word_prefixes) {
        (10, 11)
    } else if is_word_element(name, b"vAlign", word_prefixes) {
        (11, 12)
    } else if is_word_element(name, b"hideMark", word_prefixes) {
        (12, 13)
    } else if is_word_element(name, b"headers", word_prefixes) {
        (13, 14)
    } else if is_word_element(name, b"cellIns", word_prefixes) {
        (14, 15)
    } else if is_word_element(name, b"cellDel", word_prefixes) {
        (15, 16)
    } else if is_word_element(name, b"cellMerge", word_prefixes) {
        (16, 17)
    } else if is_word_element(name, b"tcPrChange", word_prefixes) {
        (17, 18)
    } else {
        (current, current)
    }
}

// ---- Table cell ----

/// Content that can appear inside a table cell.
#[derive(Debug, Clone, PartialEq)]
pub enum CellContent {
    /// A paragraph.
    Paragraph(CT_P),
    /// A nested table.
    Table(CT_Tbl),
    /// A paragraph-level content control.
    ContentControl(CT_Sdt),
}

/// `CT_Tc` — A table cell containing paragraphs and possibly nested tables.
#[derive(Debug, Clone, PartialEq)]
pub struct CT_Tc {
    pub properties: Option<CT_TcPr>,
    /// Cell content (paragraphs and nested tables).
    pub content: Vec<CellContent>,
    /// Raw XML for children we do not model, tagged with the content index they
    /// appeared before so they can be written back in place.
    pub extra_xml: Vec<(usize, Vec<u8>)>,
}

#[allow(non_snake_case)]
impl CT_Tc {
    pub fn new() -> Self {
        CT_Tc {
            properties: None,
            // OOXML requires at least one paragraph per cell
            content: vec![CellContent::Paragraph(CT_P::new())],
            extra_xml: Vec::new(),
        }
    }

    /// Get all paragraphs in this cell (excludes nested tables).
    pub fn paragraphs(&self) -> Vec<&CT_P> {
        let mut paragraphs = Vec::new();
        self.collect_paragraphs(&mut paragraphs);
        paragraphs
    }

    /// Get mutable reference to paragraphs (backward compatibility).
    pub fn paragraphs_mut(&mut self) -> Vec<&mut CT_P> {
        self.content
            .iter_mut()
            .filter_map(|c| match c {
                CellContent::Paragraph(p) => Some(p),
                CellContent::Table(_) | CellContent::ContentControl(_) => None,
            })
            .collect()
    }

    pub fn text(&self) -> String {
        self.paragraphs()
            .iter()
            .map(|p| p.text())
            .collect::<Vec<_>>()
            .join("\n")
    }

    pub fn from_xml(reader: &mut Reader<&[u8]>) -> Result<Self> {
        Self::from_xml_with_prefixes_and_owner_bindings(reader, &["w".to_string()], &[])
    }

    pub(crate) fn from_xml_with_prefixes_and_owner_bindings(
        reader: &mut Reader<&[u8]>,
        word_prefixes: &[String],
        owner_bindings: &[(String, String)],
    ) -> Result<Self> {
        let mut properties = None;
        let mut content = Vec::new();
        let mut extra_xml = Vec::new();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let name = e.name();
                    let prefixes = word_prefixes_at(e, word_prefixes)?;
                    if is_word_element(name.as_ref(), b"tcPr", &prefixes) {
                        let local_bindings = local_namespace_overrides(e, word_prefixes)?;
                        let property_bindings =
                            merged_owner_bindings(owner_bindings, &local_bindings);
                        properties = Some(CT_TcPr::from_xml_with_prefixes_and_owner_bindings(
                            reader,
                            &prefixes,
                            &property_bindings,
                        )?);
                    } else if is_word_element(name.as_ref(), b"p", &prefixes) {
                        content.push(CellContent::Paragraph(CT_P::from_xml_with_prefixes(
                            reader, &prefixes,
                        )?));
                    } else if is_word_element(name.as_ref(), b"tbl", &prefixes) {
                        let local_bindings = local_namespace_overrides(e, word_prefixes)?;
                        let table_bindings = merged_owner_bindings(owner_bindings, &local_bindings);
                        content.push(CellContent::Table(
                            CT_Tbl::from_xml_with_prefixes_and_owner_bindings(
                                reader,
                                &prefixes,
                                &table_bindings,
                            )?,
                        ));
                    } else if is_word_element(name.as_ref(), b"sdt", &prefixes) {
                        let raw = crate::text::raw_with_external_bindings(
                            &capture_element(reader, e)?,
                            owner_bindings,
                        )?;
                        if let Some(sdt) = CT_Sdt::from_cell_raw(&raw, &prefixes) {
                            content.push(CellContent::ContentControl(sdt));
                        } else {
                            extra_xml.push((content.len(), raw));
                        }
                    } else {
                        // Content controls (w:sdt), bookmarks and revision
                        // marks live here. Keep them verbatim rather than
                        // dropping the subtree, which used to delete every
                        // paragraph wrapped in a content control.
                        extra_xml.push((
                            content.len(),
                            crate::text::raw_with_external_bindings(
                                &capture_element(reader, e)?,
                                owner_bindings,
                            )?,
                        ));
                    }
                }
                Ok(Event::Empty(ref e)) => {
                    let name = e.name();
                    let prefixes = word_prefixes_at(e, word_prefixes)?;
                    if !is_word_element(name.as_ref(), b"tcPr", &prefixes) {
                        extra_xml.push((
                            content.len(),
                            crate::text::raw_with_external_bindings(
                                &capture_empty_element(e)?,
                                owner_bindings,
                            )?,
                        ));
                    }
                }
                Ok(Event::End(ref e)) if matches_local_name(e.name().as_ref(), b"tc") => {
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(CT_Tc {
            properties,
            content,
            extra_xml,
        })
    }

    pub fn to_xml<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_event(Event::Start(BytesStart::new("w:tc")))?;

        if let Some(ref props) = self.properties {
            props.to_xml(writer)?;
        }

        for (idx, item) in self.content.iter().enumerate() {
            write_extras_at(writer, &self.extra_xml, idx)?;
            match item {
                CellContent::Paragraph(p) => p.to_xml(writer)?,
                CellContent::Table(tbl) => tbl.to_xml(writer)?,
                CellContent::ContentControl(sdt) => sdt.to_xml(writer)?,
            }
        }
        write_extras_at(writer, &self.extra_xml, self.content.len())?;

        writer.write_event(Event::End(BytesEnd::new("w:tc")))?;
        Ok(())
    }

    pub(crate) fn collect_controls<'a>(&'a self, controls: &mut Vec<&'a CT_Sdt>) {
        for content in &self.content {
            match content {
                CellContent::Paragraph(paragraph) => paragraph.collect_controls(controls),
                CellContent::Table(table) => table.collect_controls(controls),
                CellContent::ContentControl(sdt) => {
                    controls.push(sdt);
                    sdt.collect_controls(SdtOwner::Cell, controls);
                }
            }
        }
    }

    pub(crate) fn collect_paragraphs<'a>(&'a self, paragraphs: &mut Vec<&'a CT_P>) {
        for content in &self.content {
            match content {
                CellContent::Paragraph(paragraph) => paragraphs.push(paragraph),
                CellContent::ContentControl(sdt) => {
                    sdt.collect_paragraphs(SdtOwner::Cell, paragraphs)
                }
                CellContent::Table(_) => {}
            }
        }
    }

    pub(crate) fn collect_tables<'a>(&'a self, tables: &mut Vec<&'a CT_Tbl>) {
        for content in &self.content {
            match content {
                CellContent::Table(table) => tables.push(table),
                CellContent::ContentControl(sdt) => sdt.collect_tables(SdtOwner::Cell, tables),
                CellContent::Paragraph(_) => {}
            }
        }
    }
}

impl Default for CT_Tc {
    fn default() -> Self {
        Self::new()
    }
}

// ---- Table row ----

/// `CT_Row` — A table row containing cells.
#[derive(Debug, Clone, PartialEq)]
pub struct CT_Row {
    /// The optional Word table-property exception for this row.
    ///
    /// Its contents affect table presentation only. The complete raw element
    /// remains the serialization source because the layout model does not
    /// interpret every table-property variant.
    pub table_property_exception: Option<Vec<u8>>,
    pub properties: Option<CT_TrPr>,
    pub cells: Vec<CT_Tc>,
    /// Raw XML for children we do not model, tagged with the cell index they
    /// appeared before so they can be written back in place.
    pub extra_xml: Vec<(usize, Vec<u8>)>,
    /// Typed cell controls at `(cell index, raw children before, control)`.
    pub content_controls: Vec<(usize, usize, CT_Sdt)>,
}

#[allow(non_snake_case)]
impl CT_Row {
    pub fn new() -> Self {
        CT_Row {
            table_property_exception: None,
            properties: None,
            cells: Vec::new(),
            extra_xml: Vec::new(),
            content_controls: Vec::new(),
        }
    }

    pub fn from_xml(reader: &mut Reader<&[u8]>) -> Result<Self> {
        Self::from_xml_with_prefixes_and_owner_bindings(reader, &["w".to_string()], &[])
    }

    pub(crate) fn from_xml_with_prefixes_and_owner_bindings(
        reader: &mut Reader<&[u8]>,
        word_prefixes: &[String],
        owner_bindings: &[(String, String)],
    ) -> Result<Self> {
        let mut table_property_exception = None;
        let mut properties = None;
        let mut cells = Vec::new();
        let mut extra_xml = Vec::new();
        let mut content_controls = Vec::new();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let name = e.name();
                    let prefixes = word_prefixes_at(e, word_prefixes)?;
                    if is_word_element(name.as_ref(), b"tblPrEx", &prefixes)
                        && table_property_exception.is_none()
                    {
                        table_property_exception = Some(crate::text::raw_with_external_bindings(
                            &capture_element(reader, e)?,
                            owner_bindings,
                        )?);
                    } else if is_word_element(name.as_ref(), b"trPr", &prefixes) {
                        let local_bindings = local_namespace_overrides(e, word_prefixes)?;
                        let property_bindings =
                            merged_owner_bindings(owner_bindings, &local_bindings);
                        properties = Some(CT_TrPr::from_xml_with_prefixes_and_owner_bindings(
                            reader,
                            &prefixes,
                            &property_bindings,
                        )?);
                    } else if is_word_element(name.as_ref(), b"tc", &prefixes) {
                        let local_bindings = local_namespace_overrides(e, word_prefixes)?;
                        let cell_bindings = merged_owner_bindings(owner_bindings, &local_bindings);
                        cells.push(CT_Tc::from_xml_with_prefixes_and_owner_bindings(
                            reader,
                            &prefixes,
                            &cell_bindings,
                        )?);
                    } else if is_word_element(name.as_ref(), b"sdt", &prefixes) {
                        let raw = crate::text::raw_with_external_bindings(
                            &capture_element(reader, e)?,
                            owner_bindings,
                        )?;
                        if let Some(sdt) = CT_Sdt::from_row_raw(&raw, &prefixes) {
                            let raw_before = extra_xml
                                .iter()
                                .filter(|(at, _)| *at == cells.len())
                                .count();
                            content_controls.push((cells.len(), raw_before, sdt));
                        } else {
                            extra_xml.push((cells.len(), raw));
                        }
                    } else {
                        // A cell wrapped in a content control used to be
                        // dropped here, leaving a row with no cells at all.
                        extra_xml.push((
                            cells.len(),
                            crate::text::raw_with_external_bindings(
                                &capture_element(reader, e)?,
                                owner_bindings,
                            )?,
                        ));
                    }
                }
                Ok(Event::Empty(ref e)) => {
                    let name = e.name();
                    let prefixes = word_prefixes_at(e, word_prefixes)?;
                    if is_word_element(name.as_ref(), b"tblPrEx", &prefixes)
                        && table_property_exception.is_none()
                    {
                        table_property_exception = Some(crate::text::raw_with_external_bindings(
                            &capture_empty_element(e)?,
                            owner_bindings,
                        )?);
                    } else if !is_word_element(name.as_ref(), b"trPr", &prefixes) {
                        extra_xml.push((
                            cells.len(),
                            crate::text::raw_with_external_bindings(
                                &capture_empty_element(e)?,
                                owner_bindings,
                            )?,
                        ));
                    }
                }
                Ok(Event::End(ref e)) if matches_local_name(e.name().as_ref(), b"tr") => {
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(CT_Row {
            table_property_exception,
            properties,
            cells,
            extra_xml,
            content_controls,
        })
    }

    pub fn to_xml<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_event(Event::Start(BytesStart::new("w:tr")))?;

        if let Some(raw) = &self.table_property_exception {
            writer.get_mut().write_all(raw)?;
        }

        if let Some(ref props) = self.properties {
            props.to_xml(writer)?;
        }

        for (idx, cell) in self.cells.iter().enumerate() {
            write_boundary(writer, &self.extra_xml, &self.content_controls, idx)?;
            cell.to_xml(writer)?;
        }
        write_boundary(
            writer,
            &self.extra_xml,
            &self.content_controls,
            self.cells.len(),
        )?;

        writer.write_event(Event::End(BytesEnd::new("w:tr")))?;
        Ok(())
    }

    /// Return direct and content-control-wrapped cells in document order.
    pub fn cells(&self) -> Vec<&CT_Tc> {
        let mut cells = Vec::new();
        self.collect_cells(&mut cells);
        cells
    }

    pub(crate) fn collect_cells<'a>(&'a self, cells: &mut Vec<&'a CT_Tc>) {
        for index in 0..=self.cells.len() {
            for (_, _, sdt) in self
                .content_controls
                .iter()
                .filter(|(at, _, _)| *at == index)
            {
                sdt.collect_cells(SdtOwner::Row, cells);
            }
            if let Some(cell) = self.cells.get(index) {
                cells.push(cell);
            }
        }
    }

    pub(crate) fn collect_controls<'a>(&'a self, controls: &mut Vec<&'a CT_Sdt>) {
        for index in 0..=self.cells.len() {
            for (_, _, sdt) in self
                .content_controls
                .iter()
                .filter(|(at, _, _)| *at == index)
            {
                controls.push(sdt);
                sdt.collect_controls(SdtOwner::Row, controls);
            }
            if let Some(cell) = self.cells.get(index) {
                cell.collect_controls(controls);
            }
        }
    }

    pub(crate) fn collect_paragraphs<'a>(&'a self, paragraphs: &mut Vec<&'a CT_P>) {
        for index in 0..=self.cells.len() {
            for (_, _, sdt) in self
                .content_controls
                .iter()
                .filter(|(at, _, _)| *at == index)
            {
                sdt.collect_paragraphs(SdtOwner::Row, paragraphs);
            }
            if let Some(cell) = self.cells.get(index) {
                cell.collect_paragraphs(paragraphs);
            }
        }
    }
}

impl Default for CT_Row {
    fn default() -> Self {
        Self::new()
    }
}

// ---- Table ----

/// `CT_Tbl` — A table element containing rows.
#[derive(Debug, Clone, PartialEq)]
pub struct CT_Tbl {
    pub properties: Option<CT_TblPr>,
    pub grid: Option<CT_TblGrid>,
    pub rows: Vec<CT_Row>,
    /// Raw XML for children we do not model, tagged with the row index they
    /// appeared before so they can be written back in place.
    pub extra_xml: Vec<(usize, Vec<u8>)>,
    /// Typed row controls at `(row index, raw children before, control)`.
    pub content_controls: Vec<(usize, usize, CT_Sdt)>,
}

#[allow(non_snake_case)]
impl CT_Tbl {
    pub fn new() -> Self {
        CT_Tbl {
            properties: None,
            grid: None,
            rows: Vec::new(),
            extra_xml: Vec::new(),
            content_controls: Vec::new(),
        }
    }

    pub fn from_xml(reader: &mut Reader<&[u8]>) -> Result<Self> {
        Self::from_xml_with_prefixes_and_owner_bindings(reader, &["w".to_string()], &[])
    }

    #[cfg(test)]
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
        let mut properties = None;
        let mut grid = None;
        let mut rows = Vec::new();
        let mut extra_xml = Vec::new();
        let mut content_controls = Vec::new();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let name = e.name();
                    let prefixes = word_prefixes_at(e, word_prefixes)?;
                    if is_word_element(name.as_ref(), b"tblPr", &prefixes) {
                        let local_bindings = local_namespace_overrides(e, word_prefixes)?;
                        let property_bindings =
                            merged_owner_bindings(owner_bindings, &local_bindings);
                        properties = Some(CT_TblPr::from_xml_with_prefixes_and_owner_bindings(
                            reader,
                            &prefixes,
                            &property_bindings,
                        )?);
                    } else if is_word_element(name.as_ref(), b"tblGrid", &prefixes) {
                        grid = Some(CT_TblGrid::from_xml_with_prefixes(reader, &prefixes)?);
                    } else if is_word_element(name.as_ref(), b"tr", &prefixes) {
                        let local_bindings = local_namespace_overrides(e, word_prefixes)?;
                        let row_bindings = merged_owner_bindings(owner_bindings, &local_bindings);
                        rows.push(CT_Row::from_xml_with_prefixes_and_owner_bindings(
                            reader,
                            &prefixes,
                            &row_bindings,
                        )?);
                    } else if is_word_element(name.as_ref(), b"sdt", &prefixes) {
                        let raw = crate::text::raw_with_external_bindings(
                            &capture_element(reader, e)?,
                            owner_bindings,
                        )?;
                        if let Some(sdt) = CT_Sdt::from_table_raw(&raw, &prefixes) {
                            let raw_before =
                                extra_xml.iter().filter(|(at, _)| *at == rows.len()).count();
                            content_controls.push((rows.len(), raw_before, sdt));
                        } else {
                            extra_xml.push((rows.len(), raw));
                        }
                    } else if matches_local_name(name.as_ref(), b"tblGrid") {
                        let raw = capture_element(reader, e)?;
                        let raw_bindings = merged_owner_bindings(
                            owner_bindings,
                            &preserved_table_raw_bindings(&prefixes),
                        );
                        extra_xml.push((
                            rows.len(),
                            crate::text::raw_with_external_bindings(&raw, &raw_bindings)?,
                        ));
                    } else {
                        // Rows wrapped in a content control used to be dropped
                        // here, which silently deleted whole tables.
                        extra_xml.push((
                            rows.len(),
                            crate::text::raw_with_external_bindings(
                                &capture_element(reader, e)?,
                                owner_bindings,
                            )?,
                        ));
                    }
                }
                Ok(Event::Empty(ref e)) => {
                    let name = e.name();
                    let prefixes = word_prefixes_at(e, word_prefixes)?;
                    // tblPr and tblGrid have fixed positions ahead of the rows,
                    // so a self-closing one must not be re-emitted from here.
                    if is_word_element(name.as_ref(), b"tblGrid", &prefixes) {
                        grid = Some(CT_TblGrid::default());
                    } else if matches_local_name(name.as_ref(), b"tblGrid") {
                        let raw = capture_empty_element(e)?;
                        let raw_bindings = merged_owner_bindings(
                            owner_bindings,
                            &preserved_table_raw_bindings(&prefixes),
                        );
                        extra_xml.push((
                            rows.len(),
                            crate::text::raw_with_external_bindings(&raw, &raw_bindings)?,
                        ));
                    } else if !is_word_element(name.as_ref(), b"tblPr", &prefixes) {
                        extra_xml.push((
                            rows.len(),
                            crate::text::raw_with_external_bindings(
                                &capture_empty_element(e)?,
                                owner_bindings,
                            )?,
                        ));
                    }
                }
                Ok(Event::End(ref e)) if matches_local_name(e.name().as_ref(), b"tbl") => {
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(CT_Tbl {
            properties,
            grid,
            rows,
            extra_xml,
            content_controls,
        })
    }

    pub fn to_xml<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_event(Event::Start(BytesStart::new("w:tbl")))?;

        if let Some(ref props) = self.properties {
            props.to_xml(writer)?;
        }

        if let Some(ref grid) = self.grid {
            grid.to_xml(writer)?;
        }

        for (idx, row) in self.rows.iter().enumerate() {
            write_boundary(writer, &self.extra_xml, &self.content_controls, idx)?;
            row.to_xml(writer)?;
        }
        write_boundary(
            writer,
            &self.extra_xml,
            &self.content_controls,
            self.rows.len(),
        )?;

        writer.write_event(Event::End(BytesEnd::new("w:tbl")))?;
        Ok(())
    }

    /// Return direct and content-control-wrapped rows in document order.
    pub fn rows(&self) -> Vec<&CT_Row> {
        let mut rows = Vec::new();
        self.collect_rows(&mut rows);
        rows
    }

    pub(crate) fn collect_rows<'a>(&'a self, rows: &mut Vec<&'a CT_Row>) {
        for index in 0..=self.rows.len() {
            for (_, _, sdt) in self
                .content_controls
                .iter()
                .filter(|(at, _, _)| *at == index)
            {
                sdt.collect_rows(SdtOwner::Table, rows);
            }
            if let Some(row) = self.rows.get(index) {
                rows.push(row);
            }
        }
    }

    pub(crate) fn collect_controls<'a>(&'a self, controls: &mut Vec<&'a CT_Sdt>) {
        for index in 0..=self.rows.len() {
            for (_, _, sdt) in self
                .content_controls
                .iter()
                .filter(|(at, _, _)| *at == index)
            {
                controls.push(sdt);
                sdt.collect_controls(SdtOwner::Table, controls);
            }
            if let Some(row) = self.rows.get(index) {
                row.collect_controls(controls);
            }
        }
    }

    pub(crate) fn collect_paragraphs<'a>(&'a self, paragraphs: &mut Vec<&'a CT_P>) {
        for index in 0..=self.rows.len() {
            for (_, _, sdt) in self
                .content_controls
                .iter()
                .filter(|(at, _, _)| *at == index)
            {
                sdt.collect_paragraphs(SdtOwner::Table, paragraphs);
            }
            if let Some(row) = self.rows.get(index) {
                row.collect_paragraphs(paragraphs);
            }
        }
    }
}

impl Default for CT_Tbl {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn parse_table_result(xml: &str) -> Result<CT_Tbl> {
        let full = format!("<w:tbl>{xml}</w:tbl>");
        let mut reader = Reader::from_str(&full);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if matches_local_name(e.name().as_ref(), b"tbl") => break,
                _ => {}
            }
            buf.clear();
        }
        CT_Tbl::from_xml(&mut reader)
    }

    fn parse_table(xml: &str) -> CT_Tbl {
        parse_table_result(xml).unwrap()
    }

    fn parse_scoped_table(xml: &str) -> Result<CT_Tbl> {
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"tbl" => {
                    let prefixes = word_prefixes_at(element, &["w".to_owned()])?;
                    let owner_bindings = local_namespace_overrides(element, &["w".to_owned()])?;
                    return CT_Tbl::from_xml_with_prefixes_and_owner_bindings(
                        &mut reader,
                        &prefixes,
                        &owner_bindings,
                    );
                }
                Ok(Event::Eof) => {
                    return Err(OxmlError::MissingElement("table".to_string()));
                }
                Err(error) => return Err(error.into()),
                Ok(_) => {}
            }
            buf.clear();
        }
    }

    #[test]
    fn parse_simple_table() {
        let tbl = parse_table(
            r#"<w:tblPr><w:tblW w:w="5000" w:type="dxa"/></w:tblPr>
               <w:tblGrid><w:gridCol w:w="2500"/><w:gridCol w:w="2500"/></w:tblGrid>
               <w:tr>
                 <w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc>
                 <w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc>
               </w:tr>
               <w:tr>
                 <w:tc><w:p><w:r><w:t>A2</w:t></w:r></w:p></w:tc>
                 <w:tc><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc>
               </w:tr>"#,
        );
        assert_eq!(tbl.rows.len(), 2);
        assert_eq!(tbl.rows[0].cells.len(), 2);
        assert_eq!(tbl.rows[0].cells[0].text(), "A1");
        assert_eq!(tbl.rows[1].cells[1].text(), "B2");

        let grid = tbl.grid.unwrap();
        assert_eq!(grid.columns.len(), 2);
        assert_eq!(grid.columns[0].width, Twips(2500));

        let pr = tbl.properties.unwrap();
        assert_eq!(pr.width.as_ref().unwrap().w, 5000);
    }

    #[test]
    fn tracked_table_grid_change_is_namespace_aware() {
        let word_namespace = crate::namespace::W_NS;
        let canonical = parse_scoped_table(&format!(
            r#"<w:tbl xmlns:w="{word_namespace}"><w:tblGrid><w:gridCol w:w="600"/><w:tblGridChange w:id="3"><w:tblGrid/></w:tblGridChange></w:tblGrid></w:tbl>"#
        ))
        .expect("canonical Word table parses");
        assert_eq!(
            canonical.grid.as_ref().unwrap().columns,
            vec![CT_TblGridCol { width: Twips(600) }]
        );
        assert!(canonical.grid.unwrap().grid_change_xml.is_some());

        let foreign_grid =
            br#"<ext:tblGrid xmlns:ext="urn:producer"><ext:gridCol ext:w="700"/></ext:tblGrid>"#
                .as_slice();
        let foreign_column = br#"<ext:gridCol ext:w="800" xmlns:ext="urn:producer"/>"#.as_slice();
        let foreign_change =
            br#"<ext:tblGridChange ext:id="foreign" xmlns:ext="urn:producer"/>"#.as_slice();
        let table = parse_scoped_table(&format!(
            r#"<q:tbl xmlns:q="{word_namespace}" xmlns:ext="urn:producer">
                 <ext:tblGrid><ext:gridCol ext:w="700"/></ext:tblGrid>
                 <q:tblGrid>
                   <ext:gridCol ext:w="800"/>
                   <q:gridCol ext:w="999" q:w="1200"/>
                   <ext:tblGridChange ext:id="foreign"/>
                   <q:tblGridChange q:id="4"><q:tblGrid><q:gridCol q:w="9000"/></q:tblGrid></q:tblGridChange>
                 </q:tblGrid>
                 <q:tr><q:tc><q:p/></q:tc></q:tr>
               </q:tbl>"#
        ))
        .expect("aliased Word table parses");

        let grid = table.grid.as_ref().expect("Word grid is modeled");
        assert_eq!(grid.columns, vec![CT_TblGridCol { width: Twips(1200) }]);
        assert_eq!(
            grid.grid_change_xml.as_deref(),
            Some(
                format!(
                    r#"<q:tblGridChange q:id="4" xmlns:q="{word_namespace}"><q:tblGrid><q:gridCol q:w="9000"/></q:tblGrid></q:tblGridChange>"#
                )
                .as_bytes()
            )
        );
        assert!(grid.extra_xml.iter().any(|raw| raw == foreign_column));
        assert!(grid.extra_xml.iter().any(|raw| raw == foreign_change));
        assert!(table.extra_xml.iter().any(|(_, raw)| raw == foreign_grid));
    }

    #[test]
    fn tracked_table_grid_raw_xml_retains_ancestor_namespace_bindings() {
        let word_namespace = crate::namespace::W_NS;
        let table = parse_scoped_table(&format!(
            r#"<q:tbl xmlns:q="{word_namespace}" xmlns:ext="urn:producer">
                 <ext:tblGrid><ext:gridCol ext:w="700"/></ext:tblGrid>
                 <q:tblGrid>
                   <q:gridCol q:w="1200"/>
                   <ext:gridCol ext:w="800"/>
                   <q:tblGridChange q:id="4"><q:tblGrid><q:gridCol q:w="9000"/><ext:marker/></q:tblGrid></q:tblGridChange>
                 </q:tblGrid>
                 <q:tr><q:tc><q:p/></q:tc></q:tr>
               </q:tbl>"#
        ))
        .expect("ancestor-bound grid XML parses");

        let first = table_to_xml(&table);
        assert!(first.contains(&format!(r#"xmlns:q="{word_namespace}""#)));
        assert!(first.contains(r#"xmlns:ext="urn:producer""#));

        let reparsed = parse_scoped_table(&first).expect("bound grid XML reparses");
        let grid = reparsed.grid.as_ref().expect("active grid remains modeled");
        assert_eq!(grid.columns, vec![CT_TblGridCol { width: Twips(1200) }]);
        assert!(grid.grid_change_xml.is_some());
        assert_eq!(grid.extra_xml.len(), 1);
        assert_eq!(reparsed.extra_xml.len(), 1);

        let second = table_to_xml(&reparsed);
        assert_eq!(second, first, "normalized raw XML remains byte-identical");
    }

    #[test]
    fn duplicate_modeled_table_grid_changes_fail_closed() {
        let result = parse_table_result(
            r#"<w:tblGrid>
                 <w:gridCol w:w="1200"/>
                 <w:tblGridChange w:id="4"><w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid></w:tblGridChange>
                 <w:tblGridChange w:id="5"><w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid></w:tblGridChange>
               </w:tblGrid>"#,
        );
        assert!(
            matches!(result, Err(OxmlError::InvalidValue(message)) if message.contains("duplicate w:tblGridChange")),
            "duplicate modeled history must fail before publishing a partial table"
        );
    }

    #[test]
    fn tracked_table_grid_change_round_trips_after_active_columns() {
        let historical = br#"<w:tblGridChange w:id="4"><w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid></w:tblGridChange>"#;
        let table = parse_table(&format!(
            r#"<w:tblGrid><w:gridCol w:w="1200"/><w:gridCol w:w="2400"/>{}</w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr>"#,
            std::str::from_utf8(historical).unwrap()
        ));

        let output = table_to_xml(&table);
        let first = output.find(r#"<w:gridCol w:w="1200"/>"#).unwrap();
        let second = output.find(r#"<w:gridCol w:w="2400"/>"#).unwrap();
        let change = output
            .find(std::str::from_utf8(historical).unwrap())
            .unwrap();
        assert!(first < second && second < change);

        let reparsed = parse_scoped_table(&output).expect("serialized table reparses");
        assert_eq!(
            reparsed.grid.unwrap().grid_change_xml.as_deref(),
            Some(historical.as_slice())
        );
    }

    #[test]
    fn whole_valued_decimal_table_measurements_parse_exactly() {
        let table = parse_table(
            r#"<w:tblPr>
                 <w:tblW w:w="9345.0" w:type="dxa"/>
                 <w:tblInd w:w="-240.000" w:type="dxa"/>
                 <w:tblCellMar>
                   <w:top w:w="120.0" w:type="dxa"/>
                   <w:left w:w="180.00" w:type="dxa"/>
                   <w:bottom w:w="120.000" w:type="dxa"/>
                   <w:right w:w="180" w:type="dxa"/>
                 </w:tblCellMar>
               </w:tblPr>
               <w:tr><w:tc><w:tcPr><w:tcW w:w="4675.00" w:type="dxa"/></w:tcPr><w:p/></w:tc></w:tr>"#,
        );

        let properties = table.properties.as_ref().expect("table properties");
        assert_eq!(properties.width.as_ref().map(|width| width.w), Some(9345));
        assert_eq!(properties.indent.as_ref().map(|width| width.w), Some(-240));
        assert_eq!(
            properties.cell_margin,
            Some(CT_TblCellMar {
                top: Some(Twips(120)),
                bottom: Some(Twips(120)),
                left: Some(Twips(180)),
                right: Some(Twips(180)),
            })
        );
        assert_eq!(
            table.rows[0].cells[0]
                .properties
                .as_ref()
                .and_then(|properties| properties.width.as_ref())
                .map(|width| width.w),
            Some(4675)
        );
    }

    #[test]
    fn table_measurement_attributes_are_namespace_aware() {
        let word_namespace = crate::namespace::W_NS;
        let table = parse_table(&format!(
            r#"<q:tblPr xmlns:q="{word_namespace}" xmlns:ext="urn:producer">
                 <q:tblW ext:w="777" ext:type="pct" q:w="9345.0" q:type="dxa"/>
                 <q:tblCellMar><q:top ext:w="999" q:w="120.00" q:type="dxa"/></q:tblCellMar>
               </q:tblPr>"#
        ));
        let properties = table.properties.as_ref().expect("table properties");
        let width = properties.width.as_ref().expect("table width");
        assert_eq!((width.w, width.width_type.as_str()), (9345, "dxa"));
        assert_eq!(
            properties
                .cell_margin
                .as_ref()
                .and_then(|margins| margins.top),
            Some(Twips(120))
        );

        let xml = format!(
            r#"<q:tblW xmlns:q="{word_namespace}" xmlns:ext="urn:producer" ext:w="777" q:type="dxa"/>"#
        );
        let mut reader = Reader::from_str(&xml);
        let mut buffer = Vec::new();
        let parsed = loop {
            match reader.read_event_into(&mut buffer) {
                Ok(Event::Empty(element)) => break CT_TblWidth::from_xml_attrs(&element).unwrap(),
                Ok(Event::Eof) => panic!("missing table width"),
                Ok(_) => {}
                Err(error) => panic!("invalid fixture: {error}"),
            }
            buffer.clear();
        };
        assert_eq!(parsed.w, 0, "foreign ext:w must not populate Word width");
    }

    #[test]
    fn unsupported_or_malformed_table_measurements_fail() {
        let cases = [
            (
                "fractional",
                r#"<w:tblPr><w:tblW w:w="9345.5" w:type="dxa"/></w:tblPr>"#,
            ),
            (
                "exponent",
                r#"<w:tr><w:tc><w:tcPr><w:tcW w:w="1e3" w:type="dxa"/></w:tcPr><w:p/></w:tc></w:tr>"#,
            ),
            (
                "empty fraction",
                r#"<w:tblPr><w:tblInd w:w="9345." w:type="dxa"/></w:tblPr>"#,
            ),
            (
                "overflow",
                r#"<w:tblPr><w:tblCellMar><w:top w:w="2147483648.0" w:type="dxa"/></w:tblCellMar></w:tblPr>"#,
            ),
            (
                "percent",
                r#"<w:tblPr><w:tblW w:w="50%" w:type="pct"/></w:tblPr>"#,
            ),
            (
                "unit",
                r#"<w:tr><w:tc><w:tcPr><w:tcW w:w="2cm" w:type="dxa"/></w:tcPr><w:p/></w:tc></w:tr>"#,
            ),
            (
                "malformed",
                r#"<w:tblPr><w:tblCellMar><w:left w:w="--1" w:type="dxa"/></w:tblCellMar></w:tblPr>"#,
            ),
            (
                "empty",
                r#"<w:tblPr><w:tblInd w:w="" w:type="dxa"/></w:tblPr>"#,
            ),
        ];

        for (name, xml) in cases {
            assert!(parse_table_result(xml).is_err(), "{name} must fail");
        }
    }

    #[test]
    fn whole_valued_decimal_table_measurements_serialize_canonically() {
        let table = parse_table(
            r#"<w:tblPr><w:tblW w:w="9345.0" w:type="dxa"/><w:tblInd w:w="-240.00" w:type="dxa"/><w:tblCellMar><w:top w:w="120.000" w:type="dxa"/></w:tblCellMar></w:tblPr><ext:kept xmlns:ext="urn:producer"/><w:tr><w:tc><w:tcPr><w:tcW w:w="4675.0" w:type="dxa"/></w:tcPr><w:p/></w:tc></w:tr>"#,
        );
        let mut output = Vec::new();
        table
            .to_xml(&mut Writer::new(&mut output))
            .expect("table writes");
        let output = String::from_utf8(output).expect("XML is UTF-8");

        assert!(output.contains(r#"<w:tblW w:w="9345" w:type="dxa"/>"#));
        assert!(output.contains(r#"<w:tblInd w:w="-240" w:type="dxa"/>"#));
        assert!(output.contains(r#"<w:top w:w="120" w:type="dxa"/>"#));
        assert!(output.contains(r#"<w:tcW w:w="4675" w:type="dxa"/>"#));
        assert!(output.contains(r#"<ext:kept xmlns:ext="urn:producer"/>"#));
        let width = output.find("<w:tblW").expect("table width writes");
        let indent = output.find("<w:tblInd").expect("table indent writes");
        let margins = output.find("<w:tblCellMar").expect("cell margins write");
        let properties_end = output.find("</w:tblPr>").expect("table properties close");
        let extension = output.find("<ext:kept").expect("extension writes");
        let row = output.find("<w:tr>").expect("table row writes");
        assert!(width < indent && indent < margins && margins < properties_end);
        assert!(properties_end < extension && extension < row);
    }

    #[test]
    fn aliased_table_cell_paragraph_properties_keep_root_scope() {
        let xml = format!(
            r#"<q:tbl xmlns:q="{}" xmlns:ext="urn:producer"><q:tr><q:tc><ext:p><ext:pPr><ext:jc ext:val="right"/></ext:pPr></ext:p><q:p><q:pPr><ext:jc ext:val="right"/><q:jc q:val="center"/></q:pPr><q:r><q:t>Cell</q:t></q:r></q:p></q:tc></q:tr></q:tbl>"#,
            crate::namespace::W_NS
        );
        let mut reader = Reader::from_str(&xml);
        let mut buf = Vec::new();
        let table = loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"tbl" => {
                    let prefixes = word_prefixes_at(element, &[]).unwrap();
                    break CT_Tbl::from_xml_with_prefixes(&mut reader, &prefixes).unwrap();
                }
                Ok(Event::Eof) => panic!("missing table"),
                event => {
                    event.unwrap();
                }
            }
            buf.clear();
        };
        let paragraphs = table.rows[0].cells[0].paragraphs();
        assert_eq!(paragraphs.len(), 1);
        assert_eq!(paragraphs[0].text(), "Cell");
        assert_eq!(
            paragraphs[0].properties.as_ref().unwrap().jc,
            Some(ST_Jc::Center)
        );
    }

    #[test]
    fn default_namespace_table_cell_properties_keep_root_scope() {
        let xml = format!(
            r#"<tbl xmlns="{0}" xmlns:w="{0}" xmlns:ext="urn:producer"><tr><tc><ext:p><ext:pPr><ext:jc ext:val="right"/></ext:pPr></ext:p><p><pPr><ext:jc ext:val="right"/><jc w:val="center"/></pPr><r><t>Cell</t></r></p></tc></tr></tbl>"#,
            crate::namespace::W_NS
        );
        let mut reader = Reader::from_str(&xml);
        let mut buf = Vec::new();
        let table = loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"tbl" => {
                    let prefixes = word_prefixes_at(element, &[]).unwrap();
                    break CT_Tbl::from_xml_with_prefixes(&mut reader, &prefixes).unwrap();
                }
                Ok(Event::Eof) => panic!("missing table"),
                event => {
                    event.unwrap();
                }
            }
            buf.clear();
        };
        let paragraphs = table.rows[0].cells[0].paragraphs();
        assert_eq!(paragraphs.len(), 1);
        assert_eq!(paragraphs[0].text(), "Cell");
        assert_eq!(
            paragraphs[0].properties.as_ref().unwrap().jc,
            Some(ST_Jc::Center)
        );
    }

    #[test]
    fn parse_cell_merge() {
        let tbl = parse_table(
            r#"<w:tblGrid><w:gridCol w:w="2500"/><w:gridCol w:w="2500"/></w:tblGrid>
               <w:tr>
                 <w:tc>
                   <w:tcPr><w:gridSpan w:val="2"/></w:tcPr>
                   <w:p><w:r><w:t>Merged</w:t></w:r></w:p>
                 </w:tc>
               </w:tr>
               <w:tr>
                 <w:tc>
                   <w:tcPr><w:vMerge w:val="restart"/></w:tcPr>
                   <w:p><w:r><w:t>VM Start</w:t></w:r></w:p>
                 </w:tc>
                 <w:tc><w:p/></w:tc>
               </w:tr>
               <w:tr>
                 <w:tc>
                   <w:tcPr><w:vMerge/></w:tcPr>
                   <w:p/>
                 </w:tc>
                 <w:tc><w:p/></w:tc>
               </w:tr>"#,
        );

        // First row: horizontal merge
        assert_eq!(
            tbl.rows[0].cells[0].properties.as_ref().unwrap().grid_span,
            Some(2)
        );

        // Second row: vertical merge start
        assert_eq!(
            tbl.rows[1].cells[0].properties.as_ref().unwrap().v_merge,
            Some(VMerge::Restart)
        );

        // Third row: vertical merge continue
        assert_eq!(
            tbl.rows[2].cells[0].properties.as_ref().unwrap().v_merge,
            Some(VMerge::Continue)
        );
    }

    #[test]
    fn expanded_presentation_cell_properties_are_read() {
        let table = parse_table(
            r#"<w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="72" w:type="dxa"></w:tcW><w:shd w:val="clear" w:fill="FF0000"></w:shd><w:noWrap></w:noWrap><w:vAlign w:val="center"></w:vAlign><w:textDirection w:val="btLr"></w:textDirection><w:cnfStyle w:val="100000000000"></w:cnfStyle></w:tcPr><w:p/></w:tc></w:tr>"#,
        );

        let properties = table.rows[0].cells[0]
            .properties
            .as_ref()
            .expect("cell properties parse");
        assert_eq!(properties.width.as_ref().map(|width| width.w), Some(72));
        assert_eq!(
            properties
                .shading
                .as_ref()
                .and_then(|shading| shading.fill.as_deref()),
            Some("FF0000")
        );
        assert_eq!(properties.no_wrap, Some(true));
        assert_eq!(properties.v_align, Some(ST_VerticalJc::Center));
        assert_eq!(properties.text_direction.as_deref(), Some("btLr"));
        assert_eq!(properties.cnf_style.as_deref(), Some("100000000000"));
    }

    #[test]
    fn foreign_cell_width_remains_raw_and_unmodelled() {
        let table = parse_table(
            r#"<w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr><w:tc><w:tcPr><ext:tcW xmlns:ext="urn:producer" ext:w="72" ext:type="dxa"/><w:tcW w:w="90" w:type="dxa"/></w:tcPr><w:p/></w:tc></w:tr>"#,
        );
        let properties = table.rows[0].cells[0]
            .properties
            .as_ref()
            .expect("cell properties parse");

        assert_eq!(properties.width.as_ref().map(|width| width.w), Some(90));
        assert_eq!(
            properties.extra_xml,
            vec![(
                0,
                br#"<ext:tcW xmlns:ext="urn:producer" ext:w="72" ext:type="dxa"/>"#.to_vec(),
            )]
        );

        let mut output = Vec::new();
        table
            .to_xml(&mut Writer::new(&mut output))
            .expect("table writes");
        let output = String::from_utf8(output).expect("XML is UTF-8");
        let foreign = output
            .find(r#"<ext:tcW xmlns:ext="urn:producer" ext:w="72" ext:type="dxa"/>"#)
            .expect("foreign width writes");
        let typed = output.find("<w:tcW").expect("typed width writes");
        assert!(
            foreign < typed,
            "foreign width stays before typed width: {output}"
        );
    }

    #[test]
    fn aliased_cell_width_uses_in_scope_word_bindings() {
        let word_namespace = crate::namespace::W_NS;
        let xml = format!(
            r#"<q:tbl xmlns:q="{word_namespace}"><q:tblGrid><q:gridCol q:w="100"/></q:tblGrid><q:tr><q:tc><q:tcPr><q:tcW q:w="72" q:type="dxa"/></q:tcPr><q:p/></q:tc></q:tr></q:tbl>"#
        );
        let mut reader = Reader::from_str(&xml);
        let mut buffer = Vec::new();
        let table = loop {
            match reader.read_event_into(&mut buffer) {
                Ok(Event::Start(element))
                    if matches_local_name(element.name().as_ref(), b"tbl") =>
                {
                    let prefixes = word_prefixes_at(&element, &[]).expect("root bindings");
                    break CT_Tbl::from_xml_with_prefixes(&mut reader, &prefixes)
                        .expect("table parses");
                }
                Ok(Event::Eof) => panic!("missing table"),
                Ok(_) => {}
                Err(error) => panic!("invalid fixture: {error}"),
            }
            buffer.clear();
        };

        assert_eq!(
            table.rows[0].cells[0]
                .properties
                .as_ref()
                .and_then(|properties| properties.width.as_ref())
                .map(|width| width.w),
            Some(72)
        );
    }

    #[test]
    fn cell_property_preserves_child_binding_declared_on_owner() {
        let table = parse_table(
            r#"<w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr><w:tc><w:tcPr xmlns:ext="urn:producer"><ext:property ext:value="kept"/></w:tcPr><w:p/></w:tc></w:tr>"#,
        );
        let properties = table.rows[0].cells[0]
            .properties
            .as_ref()
            .expect("cell properties parse");

        assert_eq!(
            properties.extra_xml,
            vec![(
                0,
                br#"<ext:property ext:value="kept" xmlns:ext="urn:producer"/>"#.to_vec(),
            )]
        );

        let mut output = Vec::new();
        table
            .to_xml(&mut Writer::new(&mut output))
            .expect("table writes");
        let output = String::from_utf8(output).expect("XML is UTF-8");
        assert!(
            output.contains(r#"<ext:property ext:value="kept" xmlns:ext="urn:producer"/>"#),
            "preserved child keeps its owner-local namespace binding: {output}"
        );
    }

    #[test]
    fn cell_property_preserves_child_binding_declared_on_cell() {
        let table = parse_table(
            r#"<w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr><w:tc xmlns:ext="urn:producer"><w:tcPr><ext:property ext:value="kept"/></w:tcPr><w:p/></w:tc></w:tr>"#,
        );
        let properties = table.rows[0].cells[0]
            .properties
            .as_ref()
            .expect("cell properties parse");

        assert_eq!(
            properties.extra_xml,
            vec![(
                0,
                br#"<ext:property ext:value="kept" xmlns:ext="urn:producer"/>"#.to_vec(),
            )]
        );

        let mut output = Vec::new();
        table
            .to_xml(&mut Writer::new(&mut output))
            .expect("table writes");
        let output = String::from_utf8(output).expect("XML is UTF-8");
        assert!(
            output.contains(r#"<ext:property ext:value="kept" xmlns:ext="urn:producer"/>"#),
            "preserved child keeps its cell-local namespace binding: {output}"
        );
    }

    #[test]
    fn foreign_same_name_after_later_property_keeps_current_boundary() {
        let table = parse_table(
            r#"<w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:textDirection w:val="btLr"/><ext:tcW xmlns:ext="urn:producer" ext:w="72"/><w:tcFitText/><w:vAlign w:val="center"/></w:tcPr><w:p/></w:tc></w:tr>"#,
        );
        let properties = table.rows[0].cells[0]
            .properties
            .as_ref()
            .expect("cell properties parse");

        assert!(properties.width.is_none());
        assert_eq!(
            properties.extra_xml,
            vec![
                (
                    10,
                    br#"<ext:tcW xmlns:ext="urn:producer" ext:w="72"/>"#.to_vec(),
                ),
                (10, br#"<w:tcFitText/>"#.to_vec()),
            ]
        );

        let mut output = Vec::new();
        table
            .to_xml(&mut Writer::new(&mut output))
            .expect("table writes");
        let output = String::from_utf8(output).expect("XML is UTF-8");
        let direction = output
            .find("<w:textDirection")
            .expect("text direction writes");
        let foreign = output.find("<ext:tcW").expect("foreign width writes");
        let fit = output.find("<w:tcFitText").expect("fit text writes");
        let align = output.find("<w:vAlign").expect("alignment writes");
        assert!(
            direction < foreign && foreign < fit && fit < align,
            "foreign child stays at its later schema boundary: {output}"
        );
    }

    #[test]
    fn unmodelled_standard_cell_properties_keep_absolute_slots_after_typed_mutation() {
        let mut table = parse_table(
            r#"<w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:hMerge/><w:tcMar/><w:hideMark/><w:headers/><w:cellIns/><w:cellDel/><w:cellMerge/><w:tcPrChange/></w:tcPr><w:p/></w:tc></w:tr>"#,
        );
        let properties = table.rows[0].cells[0]
            .properties
            .as_ref()
            .expect("cell properties parse");
        assert_eq!(properties.h_merge.as_deref(), Some(""));
        assert_eq!(
            properties.extra_xml,
            vec![
                (8, br#"<w:tcMar/>"#.to_vec()),
                (12, br#"<w:hideMark/>"#.to_vec()),
                (13, br#"<w:headers/>"#.to_vec()),
                (14, br#"<w:cellIns/>"#.to_vec()),
                (15, br#"<w:cellDel/>"#.to_vec()),
                (16, br#"<w:cellMerge/>"#.to_vec()),
                (17, br#"<w:tcPrChange/>"#.to_vec()),
            ]
        );

        let properties = table.rows[0].cells[0]
            .properties
            .as_mut()
            .expect("cell properties parse");
        properties.grid_span = Some(2);
        properties.v_merge = Some(VMerge::Restart);
        properties.no_wrap = Some(true);
        properties.text_direction = Some("btLr".to_owned());
        properties.v_align = Some(ST_VerticalJc::Center);

        let mut output = Vec::new();
        table
            .to_xml(&mut Writer::new(&mut output))
            .expect("table writes");
        let output = String::from_utf8(output).expect("XML is UTF-8");
        let positions = [
            "<w:gridSpan",
            "<w:hMerge",
            "<w:vMerge",
            "<w:noWrap",
            "<w:tcMar",
            "<w:textDirection",
            "<w:vAlign",
            "<w:hideMark",
            "<w:headers",
            "<w:cellIns",
            "<w:cellDel",
            "<w:cellMerge",
            "<w:tcPrChange",
        ]
        .map(|element| output.find(element).expect("cell property writes"));
        assert!(
            positions.windows(2).all(|pair| pair[0] < pair[1]),
            "typed mutations retain the absolute cell-property order: {output}"
        );
    }

    #[test]
    fn invalid_structural_cell_merge_values_fail_to_parse() {
        let zero_span = r#"<w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:gridSpan w:val="0"/></w:tcPr><w:p/></w:tc></w:tr>"#;
        let bad_merge = r#"<w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:vMerge w:val="sideways"/></w:tcPr><w:p/></w:tc></w:tr>"#;

        for xml in [zero_span, bad_merge] {
            let table_xml = format!("<w:tbl>{xml}</w:tbl>");
            let mut reader = Reader::from_str(&table_xml);
            let mut buffer = Vec::new();
            loop {
                match reader.read_event_into(&mut buffer) {
                    Ok(Event::Start(element))
                        if matches_local_name(element.name().as_ref(), b"tbl") =>
                    {
                        assert!(CT_Tbl::from_xml(&mut reader).is_err(), "{xml}");
                        break;
                    }
                    Ok(Event::Eof) => panic!("missing table"),
                    Ok(_) => {}
                    Err(error) => panic!("invalid fixture: {error}"),
                }
                buffer.clear();
            }
        }
    }

    #[test]
    fn parse_table_borders() {
        let tbl = parse_table(
            r#"<w:tblPr>
                 <w:tblBorders>
                   <w:top w:val="single" w:sz="4" w:color="000000"/>
                   <w:bottom w:val="single" w:sz="4" w:color="000000"/>
                   <w:left w:val="single" w:sz="4" w:color="000000"/>
                   <w:right w:val="single" w:sz="4" w:color="000000"/>
                   <w:insideH w:val="single" w:sz="4" w:color="000000"/>
                   <w:insideV w:val="single" w:sz="4" w:color="000000"/>
                 </w:tblBorders>
               </w:tblPr>
               <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
               <w:tr><w:tc><w:p/></w:tc></w:tr>"#,
        );

        let borders = tbl.properties.unwrap().borders.unwrap();
        assert_eq!(borders.top.unwrap().val, ST_Border::Single);
        assert_eq!(borders.inside_h.unwrap().val, ST_Border::Single);
        assert_eq!(borders.inside_v.unwrap().val, ST_Border::Single);
    }

    #[test]
    fn parse_cell_shading() {
        let tbl = parse_table(
            r#"<w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
               <w:tr>
                 <w:tc>
                   <w:tcPr><w:shd w:val="clear" w:fill="FFFF00"/></w:tcPr>
                   <w:p/>
                 </w:tc>
               </w:tr>"#,
        );

        let shd = tbl.rows[0].cells[0]
            .properties
            .as_ref()
            .unwrap()
            .shading
            .as_ref()
            .unwrap();
        assert_eq!(shd.fill, Some("FFFF00".to_string()));
    }

    #[test]
    fn parse_row_properties() {
        let tbl = parse_table(
            r#"<w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
               <w:tr>
                 <w:trPr>
                   <w:trHeight w:val="720" w:hRule="exact"/>
                   <w:tblHeader/>
                   <w:gridBefore w:val="1"/>
                   <w:gridAfter w:val="2"/>
                 </w:trPr>
                 <w:tc><w:p/></w:tc>
               </w:tr>"#,
        );

        let tr_pr = tbl.rows[0].properties.as_ref().unwrap();
        assert_eq!(tr_pr.height, Some(Twips(720)));
        assert_eq!(tr_pr.height_rule, Some("exact".to_string()));
        assert_eq!(tr_pr.header, Some(true));
        assert_eq!(tr_pr.grid_before, Some(1));
        assert_eq!(tr_pr.grid_after, Some(2));
    }

    #[test]
    fn row_properties_write_grid_offsets_in_schema_order() {
        let table = parse_table(
            r#"<w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
               <w:tr>
                 <w:trPr>
                   <w:gridBefore w:val="1"/>
                   <w:gridAfter w:val="2"/>
                   <w:cantSplit/>
                   <w:trHeight w:val="720" w:hRule="exact"/>
                   <w:tblHeader/>
                   <w:jc w:val="center"/>
                 </w:trPr>
                 <w:tc><w:p/></w:tc>
               </w:tr>"#,
        );

        let xml = table_to_xml(&table);
        let children = [
            "<w:gridBefore",
            "<w:gridAfter",
            "<w:cantSplit",
            "<w:trHeight",
            "<w:tblHeader",
            "<w:jc",
        ];
        let positions = children.map(|child| xml.find(child).expect("row property writes"));

        assert!(positions.windows(2).all(|pair| pair[0] < pair[1]), "{xml}");
    }

    #[test]
    fn unmodelled_row_properties_keep_schema_slots_around_revision_markers() {
        let table = parse_table(
            r#"<w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
               <w:tr>
                 <w:trPr>
                   <w:tblCellSpacing w:w="24" w:type="dxa"/>
                   <w:ins w:id="1" w:author="Author"/>
                   <w:del w:id="2" w:author="Author"/>
                   <w:trPrChange w:id="3" w:author="Author"/>
                 </w:trPr>
                 <w:tc><w:p/></w:tc>
               </w:tr>"#,
        );

        let properties = table.rows[0]
            .properties
            .as_ref()
            .expect("row properties parse");
        assert_eq!(
            properties.extra_xml,
            vec![
                (9, br#"<w:tblCellSpacing w:w="24" w:type="dxa"/>"#.to_vec()),
                (
                    14,
                    br#"<w:trPrChange w:id="3" w:author="Author"/>"#.to_vec(),
                ),
            ]
        );

        let xml = table_to_xml(&table);
        let spacing = xml.find("<w:tblCellSpacing").expect("cell spacing writes");
        let insertion = xml.find("<w:ins").expect("insertion writes");
        let deletion = xml.find("<w:del").expect("deletion writes");
        let change = xml.find("<w:trPrChange").expect("change writes");
        assert!(
            spacing < insertion && insertion < deletion && deletion < change,
            "row properties retain schema order: {xml}"
        );
    }

    #[test]
    fn row_property_children_keep_bindings_declared_on_the_row_property_owner() {
        let table = parse_table(
            r#"<w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
               <w:tr>
                 <w:trPr xmlns:ext="urn:row-property">
                   <ext:property ext:value="kept"/>
                 </w:trPr>
                 <w:tc><w:p/></w:tc>
               </w:tr>"#,
        );

        let xml = table_to_xml(&table);
        assert!(
            xml.contains(r#"<ext:property ext:value="kept" xmlns:ext="urn:row-property"/>"#),
            "row property child keeps the owner's namespace binding: {xml}"
        );

        let inner = xml
            .strip_prefix("<w:tbl>")
            .and_then(|xml| xml.strip_suffix("</w:tbl>"))
            .expect("serialized table wrapper");
        let reopened = parse_table(inner);
        assert_eq!(
            reopened.rows[0]
                .properties
                .as_ref()
                .expect("row properties reopen")
                .extra_xml,
            vec![(
                0,
                br#"<ext:property ext:value="kept" xmlns:ext="urn:row-property"/>"#.to_vec(),
            )]
        );
    }

    #[test]
    fn malformed_row_revision_markers_keep_their_distinct_schema_slots() {
        let table = parse_table(
            r#"<w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
               <w:tr>
                 <w:trPr>
                   <w:ins w:id="bad-ins" w:author="Author"/>
                   <w:ins w:id="1" w:author="Author"/>
                   <w:del w:id="bad-del" w:author="Author"/>
                   <w:del w:id="2" w:author="Author"/>
                   <w:trPrChange w:id="3" w:author="Author"/>
                 </w:trPr>
                 <w:tc><w:p/></w:tc>
               </w:tr>"#,
        );

        let properties = table.rows[0]
            .properties
            .as_ref()
            .expect("row properties parse");
        assert_eq!(properties.revision_xml_positions, vec![12, 13]);

        let xml = table_to_xml(&table);
        let malformed_insertion = xml.find(r#"w:id="bad-ins""#).expect("raw insertion writes");
        let insertion = xml.find(r#"w:id="1""#).expect("typed insertion writes");
        let malformed_deletion = xml.find(r#"w:id="bad-del""#).expect("raw deletion writes");
        let deletion = xml.find(r#"w:id="2""#).expect("typed deletion writes");
        let change = xml.find("<w:trPrChange").expect("change writes");
        assert!(
            malformed_insertion < insertion
                && insertion < malformed_deletion
                && malformed_deletion < deletion
                && deletion < change,
            "malformed row markers stay in their own schema slots: {xml}"
        );
    }

    #[test]
    fn foreign_grid_offsets_do_not_acquire_word_semantics() {
        let table = parse_table(
            r#"<w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
               <w:tr>
                 <w:trPr>
                   <ext:gridBefore xmlns:ext="urn:producer" ext:val="8"/>
                   <ext:gridAfter xmlns:ext="urn:producer" ext:val="9"/>
                 </w:trPr>
                 <w:tc><w:p/></w:tc>
               </w:tr>"#,
        );

        let properties = table.rows[0]
            .properties
            .as_ref()
            .expect("row properties parse");
        assert_eq!(properties.grid_before, None);
        assert_eq!(properties.grid_after, None);
        assert_eq!(properties.extra_xml.len(), 2);

        let xml = table_to_xml(&table);
        assert!(xml.contains(r#"<ext:gridBefore xmlns:ext="urn:producer" ext:val="8"/>"#));
        assert!(xml.contains(r#"<ext:gridAfter xmlns:ext="urn:producer" ext:val="9"/>"#));
        assert!(!xml.contains("<w:gridBefore"), "{xml}");
        assert!(!xml.contains("<w:gridAfter"), "{xml}");
    }

    #[test]
    fn foreign_row_revision_lookalikes_keep_their_observed_boundaries() {
        let table = parse_table(
            r#"<w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
               <w:tr>
                 <w:trPr>
                   <w:tblHeader/>
                   <ext:ins xmlns:ext="urn:producer"/>
                   <w:jc w:val="center"/>
                   <ext:del xmlns:ext="urn:producer"/>
                 </w:trPr>
                 <w:tc><w:p/></w:tc>
               </w:tr>"#,
        );

        let properties = table.rows[0]
            .properties
            .as_ref()
            .expect("row properties parse");
        assert!(properties.revision_xml.is_empty());
        assert_eq!(properties.extra_xml.len(), 2);

        let xml = table_to_xml(&table);
        let header = xml.find("<w:tblHeader").expect("header writes");
        let insertion = xml.find("<ext:ins").expect("foreign insertion writes");
        let alignment = xml.find("<w:jc").expect("alignment writes");
        let deletion = xml.find("<ext:del").expect("foreign deletion writes");
        assert!(
            header < insertion && insertion < alignment && alignment < deletion,
            "foreign lookalikes keep their observed boundaries: {xml}"
        );
    }

    #[test]
    fn border_children_keep_bindings_declared_on_border_owners() {
        let table = parse_table(
            r#"<w:tblPr>
                 <w:tblBorders xmlns:ext="urn:table-border">
                   <ext:diagonal ext:style="kept"/>
                 </w:tblBorders>
               </w:tblPr>
               <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
               <w:tr>
                 <w:tc>
                   <w:tcPr>
                     <w:tcBorders xmlns:vendor="urn:cell-border">
                       <vendor:diagonal vendor:style="kept"/>
                     </w:tcBorders>
                   </w:tcPr>
                   <w:p/>
                 </w:tc>
               </w:tr>"#,
        );

        let xml = table_to_xml(&table);
        assert!(
            xml.contains(r#"<ext:diagonal ext:style="kept" xmlns:ext="urn:table-border"/>"#),
            "table border child keeps the owner's namespace binding: {xml}"
        );
        assert!(
            xml.contains(
                r#"<vendor:diagonal vendor:style="kept" xmlns:vendor="urn:cell-border"/>"#
            ),
            "cell border child keeps the owner's namespace binding: {xml}"
        );
    }

    #[test]
    fn table_border_children_keep_bindings_declared_only_on_the_table_owner() {
        let table = parse_scoped_table(
            r#"<w:tbl xmlns:ext="urn:table-owner">
                  <w:tblPr>
                    <w:tblBorders>
                      <ext:diagonal ext:style="kept"/>
                    </w:tblBorders>
                  </w:tblPr>
                  <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
                </w:tbl>"#,
        )
        .expect("table parses");

        let xml = table_to_xml(&table);
        assert!(
            xml.contains(r#"<ext:diagonal ext:style="kept" xmlns:ext="urn:table-owner"/>"#),
            "table border child keeps the outer table binding: {xml}"
        );

        let reopened = parse_scoped_table(&xml).expect("serialized table reopens");
        assert_eq!(
            reopened
                .properties
                .as_ref()
                .and_then(|properties| properties.borders.as_ref())
                .expect("table borders reopen")
                .extra_xml,
            vec![br#"<ext:diagonal ext:style="kept" xmlns:ext="urn:table-owner"/>"#.to_vec()]
        );
    }

    #[test]
    fn row_property_children_keep_bindings_declared_only_on_the_row_owner() {
        let table = parse_scoped_table(
            r#"<w:tbl>
                  <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
                  <w:tr xmlns:ext="urn:row-owner">
                    <w:trPr><ext:property ext:value="kept"/></w:trPr>
                    <w:tc><w:p/></w:tc>
                  </w:tr>
                </w:tbl>"#,
        )
        .expect("table parses");

        let xml = table_to_xml(&table);
        assert!(
            xml.contains(r#"<ext:property ext:value="kept" xmlns:ext="urn:row-owner"/>"#),
            "row property child keeps the outer row binding: {xml}"
        );

        let reopened = parse_scoped_table(&xml).expect("serialized table reopens");
        assert_eq!(
            reopened.rows[0]
                .properties
                .as_ref()
                .expect("row properties reopen")
                .extra_xml,
            vec![(
                0,
                br#"<ext:property ext:value="kept" xmlns:ext="urn:row-owner"/>"#.to_vec(),
            )]
        );
    }

    #[test]
    fn owner_bindings_survive_typed_content_controls_at_table_row_and_cell_levels() {
        let cases = [
            (
                "table",
                r#"<w:tbl xmlns:ext="urn:table-control">
                      <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
                      <w:sdt><w:sdtContent>
                        <w:tr>
                          <w:trPr><ext:property ext:value="kept"/></w:trPr>
                          <w:tc><w:p/></w:tc>
                        </w:tr>
                      </w:sdtContent></w:sdt>
                    </w:tbl>"#,
                "urn:table-control",
                "<ext:property",
            ),
            (
                "row",
                r#"<w:tbl>
                      <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
                      <w:tr xmlns:ext="urn:row-control">
                        <w:sdt><w:sdtContent>
                          <w:tc>
                            <w:tcPr><ext:property ext:value="kept"/></w:tcPr>
                            <w:p/>
                          </w:tc>
                        </w:sdtContent></w:sdt>
                      </w:tr>
                    </w:tbl>"#,
                "urn:row-control",
                "<ext:property",
            ),
            (
                "cell",
                r#"<w:tbl>
                      <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
                      <w:tr>
                        <w:tc xmlns:ext="urn:cell-control">
                          <w:sdt><w:sdtContent>
                            <w:tbl>
                              <w:tblPr><w:tblBorders><ext:diagonal/></w:tblBorders></w:tblPr>
                              <w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>
                            </w:tbl>
                          </w:sdtContent></w:sdt>
                        </w:tc>
                      </w:tr>
                    </w:tbl>"#,
                "urn:cell-control",
                "<ext:diagonal",
            ),
        ];

        for (level, source, namespace, retained_child) in cases {
            let table = parse_scoped_table(source).unwrap_or_else(|error| {
                panic!("{level} content control parses: {error}");
            });
            let xml = table_to_xml(&table);
            assert!(xml.contains(retained_child), "{level}: {xml}");
            assert!(
                xml.contains(&format!(r#"xmlns:ext="{namespace}""#)),
                "{level} control carries its outer owner binding: {xml}"
            );

            let reopened = parse_scoped_table(&xml).unwrap_or_else(|error| {
                panic!("{level} content control reopens: {error}");
            });
            let rewritten = table_to_xml(&reopened);
            assert!(rewritten.contains(retained_child), "{level}: {rewritten}");
            assert!(
                rewritten.contains(&format!(r#"xmlns:ext="{namespace}""#)),
                "{level} binding survives a second serialization: {rewritten}"
            );
        }
    }

    #[test]
    fn owner_bindings_survive_direct_raw_children_at_table_row_and_cell_levels() {
        let cases = [
            (
                "table",
                r#"<w:tbl xmlns:ext="urn:table-raw">
                      <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
                      <ext:block ext:value="kept"><ext:nested/></ext:block>
                      <ext:empty ext:value="kept"/>
                      <ext:tblGrid><ext:column/></ext:tblGrid>
                      <ext:tblGrid ext:value="empty"/>
                      <w:tr><w:tc><w:p/></w:tc></w:tr>
                    </w:tbl>"#,
                "urn:table-raw",
                vec!["<ext:block", "<ext:empty", "<ext:tblGrid"],
            ),
            (
                "row",
                r#"<w:tbl>
                      <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
                      <w:tr xmlns:ext="urn:row-raw">
                        <w:tblPrEx><ext:property ext:value="kept"/></w:tblPrEx>
                        <ext:block ext:value="kept"><ext:nested/></ext:block>
                        <ext:empty ext:value="kept"/>
                        <w:tc><w:p/></w:tc>
                      </w:tr>
                    </w:tbl>"#,
                "urn:row-raw",
                vec!["<w:tblPrEx", "<ext:block", "<ext:empty"],
            ),
            (
                "cell",
                r#"<w:tbl>
                      <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
                      <w:tr><w:tc xmlns:ext="urn:cell-raw">
                        <ext:block ext:value="kept"><ext:nested/></ext:block>
                        <ext:empty ext:value="kept"/>
                        <w:p/>
                      </w:tc></w:tr>
                    </w:tbl>"#,
                "urn:cell-raw",
                vec!["<ext:block", "<ext:empty"],
            ),
        ];

        for (level, source, namespace, retained_children) in cases {
            let table = parse_scoped_table(source).unwrap_or_else(|error| {
                panic!("{level} raw children parse: {error}");
            });
            let written = table_to_xml(&table);
            for child in &retained_children {
                assert!(
                    written.contains(child),
                    "{level} retains {child}: {written}"
                );
            }
            assert!(
                written.contains(&format!(r#"xmlns:ext="{namespace}""#)),
                "{level} raw child carries its owner binding: {written}"
            );

            let reopened = parse_scoped_table(&written).unwrap_or_else(|error| {
                panic!("{level} raw children reopen: {error}");
            });
            let rewritten = table_to_xml(&reopened);
            for child in &retained_children {
                assert!(
                    rewritten.contains(child),
                    "{level} retains {child} after a second serialization: {rewritten}"
                );
            }
            assert!(
                rewritten.contains(&format!(r#"xmlns:ext="{namespace}""#)),
                "{level} owner binding survives a second serialization: {rewritten}"
            );
        }
    }

    #[test]
    fn table_property_exception_is_preserved_as_row_formatting() {
        let table = parse_table(
            r#"<w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr><w:tblPrEx><w:shd w:val="clear" w:color="auto" w:fill="ced7e7"/></w:tblPrEx><w:trPr><w:trHeight w:val="290" w:hRule="atLeast"/></w:trPr><w:tc><w:p/></w:tc></w:tr>"#,
        );
        let row = &table.rows[0];

        assert_eq!(
            row.table_property_exception.as_deref(),
            Some(
                br#"<w:tblPrEx><w:shd w:val="clear" w:color="auto" w:fill="ced7e7"/></w:tblPrEx>"#
                    .as_slice()
            )
        );

        let xml = table_to_xml(&table);
        let exception = xml.find("<w:tblPrEx>").expect("exception writes");
        let properties = xml.find("<w:trPr>").expect("row properties write");
        assert!(exception < properties, "tblPrEx must precede trPr: {xml}");
    }

    #[test]
    fn unmodelled_table_property_groups_are_retained() {
        let table = parse_table(
            r#"<w:tblPr><w:bidiVisual/><w:tblBorders><w:diagonalDown/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr><w:trPr><w:tblCellSpacing/></w:trPr><w:tc><w:tcPr><w:fitText/></w:tcPr><w:p/></w:tc></w:tr>"#,
        );

        let table_properties = table.properties.as_ref().expect("table properties parse");
        assert_eq!(table_properties.extra_xml.len(), 1);
        assert_eq!(
            table_properties
                .borders
                .as_ref()
                .expect("table borders parse")
                .extra_xml
                .len(),
            1
        );
        assert_eq!(
            table.rows[0]
                .properties
                .as_ref()
                .expect("row properties parse")
                .extra_xml
                .len(),
            1
        );
        assert_eq!(
            table.rows[0].cells[0]
                .properties
                .as_ref()
                .expect("cell properties parse")
                .extra_xml
                .len(),
            1
        );
    }

    #[test]
    fn unmodelled_table_properties_keep_schema_slots_before_change() {
        let table = parse_table(
            r#"<w:tblPr>
                 <w:tblOverlap w:val="never"/>
                 <w:tblW w:w="5000" w:type="dxa"/>
                 <ext:between xmlns:ext="urn:producer"/>
                 <w:jc w:val="center"/>
                 <w:tblCaption w:val="caption"/>
                 <w:tblPrChange/>
               </w:tblPr>
               <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>"#,
        );

        let properties = table.properties.as_ref().expect("table properties parse");
        assert_eq!(
            properties.extra_xml,
            vec![
                (2, br#"<w:tblOverlap w:val="never"/>"#.to_vec()),
                (7, br#"<ext:between xmlns:ext="urn:producer"/>"#.to_vec(),),
                (15, br#"<w:tblCaption w:val="caption"/>"#.to_vec()),
            ]
        );

        let xml = table_to_xml(&table);
        let overlap = xml.find("<w:tblOverlap").expect("overlap writes");
        let width = xml.find("<w:tblW").expect("width writes");
        let extension = xml.find("<ext:between").expect("extension writes");
        let alignment = xml.find("<w:jc").expect("alignment writes");
        let caption = xml.find("<w:tblCaption").expect("caption writes");
        let change = xml.find("<w:tblPrChange").expect("change writes");

        assert!(
            overlap < width
                && width < extension
                && extension < alignment
                && alignment < caption
                && caption < change,
            "{xml}"
        );
    }

    #[test]
    fn table_layout_reads_the_schema_type_attribute() {
        let table = parse_table(concat!(
            r#"<w:tblPr><w:tblLayout w:type="fixed"/></w:tblPr>"#,
            r#"<w:tblGrid><w:gridCol w:w="100"/></w:tblGrid>"#,
        ));

        assert_eq!(table.properties.unwrap().layout.as_deref(), Some("fixed"));
    }

    #[test]
    fn round_trip_table() {
        let mut tbl = CT_Tbl::new();
        tbl.properties = Some(CT_TblPr {
            width: Some(CT_TblWidth::dxa(9000)),
            borders: Some(CT_TblBorders {
                top: Some(CT_BorderEdge {
                    val: ST_Border::Single,
                    sz: Some(4),
                    space: Some(0),
                    color: Some("000000".to_string()),
                }),
                bottom: Some(CT_BorderEdge {
                    val: ST_Border::Single,
                    sz: Some(4),
                    space: Some(0),
                    color: Some("000000".to_string()),
                }),
                ..Default::default()
            }),
            ..Default::default()
        });
        tbl.grid = Some(CT_TblGrid {
            columns: vec![
                CT_TblGridCol { width: Twips(4500) },
                CT_TblGridCol { width: Twips(4500) },
            ],
            ..Default::default()
        });

        let mut row = CT_Row::new();
        let mut cell1 = CT_Tc::new();
        cell1.paragraphs_mut()[0].add_run("Hello");
        let mut cell2 = CT_Tc::new();
        cell2.paragraphs_mut()[0].add_run("World");
        row.cells.push(cell1);
        row.cells.push(cell2);
        tbl.rows.push(row);

        // Serialize
        let mut output = Vec::new();
        let mut writer = Writer::new(&mut output);
        tbl.to_xml(&mut writer).unwrap();
        let xml = String::from_utf8(output).unwrap();

        // Parse back
        let parsed = parse_table(
            xml.strip_prefix("<w:tbl>")
                .unwrap()
                .strip_suffix("</w:tbl>")
                .unwrap(),
        );

        assert_eq!(parsed.rows.len(), 1);
        assert_eq!(parsed.rows[0].cells.len(), 2);
        assert_eq!(parsed.rows[0].cells[0].text(), "Hello");
        assert_eq!(parsed.rows[0].cells[1].text(), "World");

        let grid = parsed.grid.unwrap();
        assert_eq!(grid.columns.len(), 2);
        assert_eq!(grid.columns[0].width, Twips(4500));

        let borders = parsed.properties.unwrap().borders.unwrap();
        assert!(borders.top.is_some());
        assert!(borders.bottom.is_some());
    }

    #[test]
    fn nested_table_xml_round_trip() {
        use crate::text::CT_P;

        // Build a cell containing a paragraph + a nested table
        let mut outer_cell = CT_Tc::new();
        outer_cell.paragraphs_mut()[0].add_run("Before table");

        let mut nested_tbl = CT_Tbl::new();
        nested_tbl.grid = Some(CT_TblGrid {
            columns: vec![CT_TblGridCol { width: Twips(2000) }],
            ..Default::default()
        });
        let mut nested_row = CT_Row::new();
        let mut nested_cell = CT_Tc::new();
        nested_cell.paragraphs_mut()[0].add_run("Nested content");
        nested_row.cells.push(nested_cell);
        nested_tbl.rows.push(nested_row);

        outer_cell.content.push(CellContent::Table(nested_tbl));

        let mut after = CT_P::new();
        after.add_run("After table");
        outer_cell.content.push(CellContent::Paragraph(after));

        // Serialize
        let mut output = Vec::new();
        let mut writer = Writer::new(&mut output);
        outer_cell.to_xml(&mut writer).unwrap();
        let xml = String::from_utf8(output).unwrap();

        // Should contain nested <w:tbl>
        assert!(xml.contains("<w:tbl>"));
        assert!(xml.contains("Nested content"));

        // Parse back
        let inner_xml = xml
            .strip_prefix("<w:tc>")
            .unwrap()
            .strip_suffix("</w:tc>")
            .unwrap();
        let full_xml = format!(
            "<w:tc xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">{inner_xml}</w:tc>"
        );
        let mut reader = Reader::from_str(&full_xml);
        reader.config_mut().trim_text(true);
        // Skip start tag
        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) if e.local_name().as_ref() == b"tc" => break,
                _ => {}
            }
        }
        let parsed = CT_Tc::from_xml(&mut reader).unwrap();

        // Check structure: 2 paragraphs + 1 nested table
        assert_eq!(parsed.paragraphs().len(), 2);
        assert_eq!(parsed.paragraphs()[0].text(), "Before table");
        assert_eq!(parsed.paragraphs()[1].text(), "After table");

        // Check nested table
        let tables: Vec<_> = parsed
            .content
            .iter()
            .filter_map(|c| match c {
                CellContent::Table(t) => Some(t),
                _ => None,
            })
            .collect();
        assert_eq!(tables.len(), 1);
        assert_eq!(tables[0].rows.len(), 1);
        assert_eq!(tables[0].rows[0].cells[0].text(), "Nested content");
    }

    #[test]
    fn paragraphs_method_backward_compat() {
        let mut cell = CT_Tc::new();
        // Cell starts with one empty paragraph
        assert_eq!(cell.paragraphs().len(), 1);

        // Add a run to existing paragraph
        cell.paragraphs_mut()[0].add_run("First");

        // Add a nested table (should not appear in paragraphs())
        let nested = CT_Tbl::new();
        cell.content.push(CellContent::Table(nested));

        // Add another paragraph
        let mut p = CT_P::new();
        p.add_run("Second");
        cell.content.push(CellContent::Paragraph(p));

        // paragraphs() should return only the 2 CT_P items
        assert_eq!(cell.paragraphs().len(), 2);
        assert_eq!(cell.paragraphs()[0].text(), "First");
        assert_eq!(cell.paragraphs()[1].text(), "Second");

        // text() should concat paragraph text with newline separator
        assert_eq!(cell.text(), "First\nSecond");
    }

    /// Serialize a table and return the XML, for the fidelity tests below.
    fn table_to_xml(tbl: &CT_Tbl) -> String {
        let mut output = Vec::new();
        let mut writer = Writer::new(&mut output);
        tbl.to_xml(&mut writer).unwrap();
        String::from_utf8(output).unwrap()
    }

    /// Table children we do not model must survive a read and write cycle.
    ///
    /// These used to be dropped, which silently deleted whole rows, cells and
    /// paragraphs whenever they were wrapped in a content control, and lost
    /// the bookmarks that cross references and a table of figures rely on.
    #[test]
    fn unknown_table_children_round_trip() {
        const GRID: &str = r#"<w:tblGrid><w:gridCol w:w="4675"/></w:tblGrid>"#;

        for (label, inner) in [
            (
                "row wrapped in a content control",
                format!(
                    r#"{GRID}<w:sdt><w:sdtContent><w:tr><w:tc><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr></w:sdtContent></w:sdt>"#
                ),
            ),
            (
                "cell wrapped in a content control",
                format!(
                    r#"{GRID}<w:tr><w:sdt><w:sdtContent><w:tc><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:sdtContent></w:sdt></w:tr>"#
                ),
            ),
            (
                "paragraph wrapped in a content control",
                format!(
                    r#"{GRID}<w:tr><w:tc><w:sdt><w:sdtContent><w:p><w:r><w:t>x</w:t></w:r></w:p></w:sdtContent></w:sdt></w:tc></w:tr>"#
                ),
            ),
            (
                "bookmark at table level",
                format!(
                    r#"{GRID}<w:bookmarkStart w:id="1" w:name="b"/><w:tr><w:tc><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr>"#
                ),
            ),
            (
                "bookmark at row level",
                format!(
                    r#"{GRID}<w:tr><w:bookmarkStart w:id="1" w:name="b"/><w:tc><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr>"#
                ),
            ),
        ] {
            let tbl = parse_table(&inner);
            let xml = table_to_xml(&tbl);
            assert_eq!(
                xml,
                format!("<w:tbl>{inner}</w:tbl>"),
                "{label} was not preserved"
            );
        }
    }

    /// A styled table must keep the markup that says which conditional parts
    /// of its style apply.
    ///
    /// `w:tblStyle` alone is not enough. `w:tblLook` and `w:cnfStyle` are what
    /// turn on the header row, banding and first column formatting, so losing
    /// them leaves the style name intact and the table drawn with base
    /// formatting only, which reads as the style having been lost.
    #[test]
    fn table_style_conditional_formatting_round_trips() {
        let inner = concat!(
            r#"<w:tblPr><w:tblStyle w:val="GridTable4-Accent1"/>"#,
            r#"<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>"#,
            r#"</w:tblPr><w:tblGrid><w:gridCol w:w="4675"/></w:tblGrid>"#,
            r#"<w:tr><w:trPr><w:cnfStyle w:val="100000000000"/></w:trPr>"#,
            r#"<w:tc><w:tcPr><w:cnfStyle w:val="001000000000"/></w:tcPr><w:p/></w:tc>"#,
            r#"</w:tr>"#,
        );
        let tbl = parse_table(inner);

        let look = tbl
            .properties
            .as_ref()
            .and_then(|p| p.look.as_ref())
            .expect("tblLook should be parsed");
        assert_eq!(look.val.as_deref(), Some("04A0"));
        assert_eq!(look.first_row, Some(true));
        assert_eq!(look.last_row, Some(false));
        assert_eq!(look.first_column, Some(true));
        assert_eq!(look.no_v_band, Some(true));

        assert_eq!(
            tbl.rows[0]
                .properties
                .as_ref()
                .and_then(|p| p.cnf_style.as_deref()),
            Some("100000000000")
        );
        assert_eq!(
            tbl.rows[0].cells[0]
                .properties
                .as_ref()
                .and_then(|p| p.cnf_style.as_deref()),
            Some("001000000000")
        );

        assert_eq!(
            table_to_xml(&tbl),
            format!("<w:tbl>{inner}</w:tbl>"),
            "the whole thing must survive a write"
        );
    }

    /// A row or cell carrying only cnfStyle is not empty.
    ///
    /// Both types skip writing their properties when every field is unset, so
    /// a new field that is not in that check is parsed and then silently
    /// dropped on the way out.
    #[test]
    fn properties_holding_only_cnf_style_are_still_written() {
        let tbl = parse_table(concat!(
            r#"<w:tblGrid><w:gridCol w:w="100"/></w:tblGrid>"#,
            r#"<w:tr><w:trPr><w:cnfStyle w:val="100000000000"/></w:trPr>"#,
            r#"<w:tc><w:tcPr><w:cnfStyle w:val="001000000000"/></w:tcPr><w:p/></w:tc></w:tr>"#,
        ));
        let xml = table_to_xml(&tbl);
        assert!(
            xml.contains(r#"<w:trPr><w:cnfStyle w:val="100000000000"/></w:trPr>"#),
            "{xml}"
        );
        assert!(
            xml.contains(r#"<w:tcPr><w:cnfStyle w:val="001000000000"/></w:tcPr>"#),
            "{xml}"
        );
    }

    /// OOXML booleans come in both spellings.
    #[test]
    fn tbl_look_accepts_either_boolean_spelling() {
        let tbl = parse_table(concat!(
            r#"<w:tblPr><w:tblLook w:firstRow="true" w:lastRow="false" w:noVBand="1"/></w:tblPr>"#,
            r#"<w:tblGrid><w:gridCol w:w="100"/></w:tblGrid>"#,
        ));
        let look = tbl
            .properties
            .as_ref()
            .and_then(|p| p.look.as_ref())
            .unwrap();
        assert_eq!(look.first_row, Some(true));
        assert_eq!(look.last_row, Some(false));
        assert_eq!(look.no_v_band, Some(true));
    }

    /// A self-closing tblPr or tblGrid must not be captured as extra XML.
    /// Both have a fixed position ahead of the rows, and extras are written
    /// from the row positions, so capturing them would reorder the children.
    #[test]
    fn self_closing_table_properties_are_not_reordered() {
        let tbl = parse_table(
            r#"<w:tblPr/><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr>"#,
        );
        assert!(tbl.extra_xml.is_empty(), "tblPr must not be captured");
        let xml = table_to_xml(&tbl);
        assert!(
            !xml.contains("</w:tr><w:tblPr/>"),
            "tblPr must never follow the rows: {xml}"
        );
    }
}
