//! Border and tab stop types for paragraph formatting.

use std::borrow::Cow;

use quick_xml::events::attributes::Attribute;
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::name::QName;
use quick_xml::{Reader, Writer};

use crate::error::{OxmlError, Result};
use crate::namespace::{W_NS, matches_local_name};
use crate::numbering::word_prefixes_at;
use crate::properties::parse_integer_measurement;
use crate::shared::{ST_Border, ST_TabJc, ST_TabLeader};
use crate::units::Twips;

const MAX_TAB_XML_DEPTH: usize = 64;

/// A single border edge (top, bottom, left, right, between).
#[derive(Debug, Clone, PartialEq)]
pub struct CT_BorderEdge {
    /// Border style
    pub val: ST_Border,
    /// Border width in eighths of a point
    pub sz: Option<u32>,
    /// Space between border and content in points
    pub space: Option<u32>,
    /// Border color as hex, e.g. "FF0000"
    pub color: Option<String>,
    /// Attributes this type does not model, in source order.
    ///
    /// `w:shadow`, `w:frame`, `w:themeColor`, `w:themeTint` and `w:themeShade`
    /// reach a serialized edge through here, as does any producer or foreign
    /// attribute. Names and values keep their stored spelling, and they are
    /// written back ahead of the modeled attributes. The value is serialized
    /// verbatim, so a caller storing one itself is storing attribute-value
    /// syntax and owns the escaping.
    pub extra_attributes: Vec<(String, String)>,
    /// Namespace-complete started edge source, including opaque child payload.
    /// Direct field edits merge owned attributes into this physical owner.
    pub source_xml: Option<Vec<u8>>,
}

impl CT_BorderEdge {
    pub fn new(val: ST_Border) -> Self {
        CT_BorderEdge {
            val,
            sz: None,
            space: None,
            color: None,
            extra_attributes: Vec::new(),
            source_xml: None,
        }
    }

    /// Select a literal color and remove only qualified competing theme selectors.
    pub fn set_literal_color(&mut self, color: &str) -> Result<()> {
        let generated;
        let raw = if let Some(raw) = &self.source_xml {
            raw.as_slice()
        } else {
            let mut start = BytesStart::new("w:border");
            start.push_attribute(("xmlns:w", W_NS));
            self.write_xml_attrs(&mut start);
            let mut writer = Writer::new(Vec::new());
            writer.write_event(Event::Empty(start))?;
            generated = writer.into_inner();
            generated.as_slice()
        };
        let mut reader = quick_xml::NsReader::from_reader(raw);
        let (_, event) = reader.read_resolved_event()?;
        let root = match event {
            Event::Start(root) | Event::Empty(root) => root,
            _ => return Err(OxmlError::InvalidValue("border source root missing".into())),
        };
        let mut remove = Vec::new();
        for attribute in root.attributes() {
            let attribute = attribute?;
            let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
            if matches!(namespace, quick_xml::name::ResolveResult::Bound(namespace) if namespace.as_ref() == W_NS.as_bytes())
                && matches!(local.as_ref(), b"themeColor" | b"themeTint" | b"themeShade")
            {
                remove.push(std::str::from_utf8(attribute.key.as_ref())?.to_owned());
            }
        }
        self.extra_attributes
            .retain(|(key, _)| !remove.contains(key));
        self.color = Some(color.to_owned());
        Ok(())
    }

    pub fn from_xml_attrs(e: &BytesStart) -> Result<Self> {
        let mut edge = CT_BorderEdge::new(ST_Border::None);

        for attr in e.attributes() {
            let attr = attr?;
            let key = attr.key.as_ref();
            let v = std::str::from_utf8(&attr.value)?;
            if matches_local_name(key, b"val") {
                edge.val = ST_Border::from_str(v).unwrap_or(edge.val);
            } else if matches_local_name(key, b"sz") {
                edge.sz = Some(parse_integer_measurement(e.name().as_ref(), key, v)?);
            } else if matches_local_name(key, b"space") {
                edge.space = Some(parse_integer_measurement(e.name().as_ref(), key, v)?);
            } else if matches_local_name(key, b"color") {
                edge.color = Some(v.to_string());
            } else {
                edge.extra_attributes
                    .push((std::str::from_utf8(key)?.to_owned(), v.to_owned()));
            }
        }

        Ok(edge)
    }

    pub(crate) fn from_xml_attrs_with_prefixes(
        e: &BytesStart,
        word_prefixes: &[String],
    ) -> Result<Self> {
        let mut edge = CT_BorderEdge::new(ST_Border::None);

        for attr in e.attributes() {
            let attr = attr?;
            let key = attr.key.as_ref();
            let value = std::str::from_utf8(&attr.value)?;
            if is_word_attribute(key, b"val", word_prefixes) {
                edge.val = ST_Border::from_str(value).unwrap_or(edge.val);
            } else if is_word_attribute(key, b"sz", word_prefixes) {
                edge.sz = Some(parse_integer_measurement(e.name().as_ref(), key, value)?);
            } else if is_word_attribute(key, b"space", word_prefixes) {
                edge.space = Some(parse_integer_measurement(e.name().as_ref(), key, value)?);
            } else if is_word_attribute(key, b"color", word_prefixes) {
                edge.color = Some(value.to_string());
            } else {
                edge.extra_attributes
                    .push((std::str::from_utf8(key)?.to_owned(), value.to_owned()));
            }
        }

        Ok(edge)
    }

    pub fn write_xml_attrs(&self, e: &mut BytesStart) {
        let mut buf = itoa::Buffer::new();
        for (name, value) in &self.extra_attributes {
            e.push_attribute(Attribute {
                key: QName(name.as_bytes()),
                value: Cow::Borrowed(value.as_bytes()),
            });
        }
        e.push_attribute(("w:val", self.val.to_str()));
        if let Some(sz) = self.sz {
            e.push_attribute(("w:sz", buf.format(sz)));
        }
        if let Some(space) = self.space {
            e.push_attribute(("w:space", buf.format(space)));
        }
        if let Some(ref color) = self.color {
            e.push_attribute(("w:color", color.as_str()));
        }
    }

    /// Write this border edge as an empty element with the given tag name.
    pub fn to_xml<W: std::io::Write>(
        &self,
        writer: &mut Writer<W>,
        tag: &str,
    ) -> crate::error::Result<()> {
        let mut e = BytesStart::new(tag);
        self.write_xml_attrs(&mut e);
        if let Some(raw) = &self.source_xml {
            let mut reader = Reader::from_reader(raw.as_slice());
            let Event::Start(start) = reader.read_event()? else {
                return Err(OxmlError::InvalidValue(
                    "started border source root missing".to_owned(),
                ));
            };
            let scope = crate::numbering::word_prefixes_at(&start, &[])?;
            let before = Self::from_xml_attrs_with_prefixes(&start, &scope)?;
            let mut old = Writer::new(Vec::new());
            let mut old_start = BytesStart::new(tag);
            before.write_xml_attrs(&mut old_start);
            old.write_event(Event::Empty(old_start))?;
            let mut generated = Writer::new(Vec::new());
            generated.write_event(Event::Empty(e))?;
            writer
                .get_mut()
                .write_all(&crate::styles::merge_defaults_scalar_attributes(
                    raw,
                    Some(old.get_ref()),
                    Some(generated.get_ref()),
                    &scope,
                    true,
                )?)?;
        } else {
            writer.write_event(Event::Empty(e))?;
        }
        Ok(())
    }
}

/// `CT_PBdr` — Paragraph borders (top, bottom, left, right, between, bar).
#[derive(Debug, Clone, Default, PartialEq)]
pub struct CT_PBdr {
    pub top: Option<CT_BorderEdge>,
    pub bottom: Option<CT_BorderEdge>,
    pub left: Option<CT_BorderEdge>,
    pub right: Option<CT_BorderEdge>,
    pub between: Option<CT_BorderEdge>,
    pub bar: Option<CT_BorderEdge>,
}

impl CT_PBdr {
    pub fn from_xml(reader: &mut Reader<&[u8]>) -> Result<Self> {
        Self::from_xml_with_prefixes(reader, &["w".to_owned()])
    }

    pub(crate) fn from_xml_with_prefixes(
        reader: &mut Reader<&[u8]>,
        word_prefixes: &[String],
    ) -> Result<Self> {
        let mut bdr = CT_PBdr::default();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(ref e)) => {
                    let name = e.name();
                    let prefixes = word_prefixes_at(e, word_prefixes)?;
                    if is_word_element(name.as_ref(), b"top", &prefixes) {
                        bdr.top = Some(CT_BorderEdge::from_xml_attrs_with_prefixes(e, &prefixes)?);
                    } else if is_word_element(name.as_ref(), b"bottom", &prefixes) {
                        bdr.bottom =
                            Some(CT_BorderEdge::from_xml_attrs_with_prefixes(e, &prefixes)?);
                    } else if is_word_element(name.as_ref(), b"left", &prefixes)
                        || is_word_element(name.as_ref(), b"start", &prefixes)
                    {
                        bdr.left = Some(CT_BorderEdge::from_xml_attrs_with_prefixes(e, &prefixes)?);
                    } else if is_word_element(name.as_ref(), b"right", &prefixes)
                        || is_word_element(name.as_ref(), b"end", &prefixes)
                    {
                        bdr.right =
                            Some(CT_BorderEdge::from_xml_attrs_with_prefixes(e, &prefixes)?);
                    } else if is_word_element(name.as_ref(), b"between", &prefixes) {
                        bdr.between =
                            Some(CT_BorderEdge::from_xml_attrs_with_prefixes(e, &prefixes)?);
                    } else if is_word_element(name.as_ref(), b"bar", &prefixes) {
                        bdr.bar = Some(CT_BorderEdge::from_xml_attrs_with_prefixes(e, &prefixes)?);
                    }
                }
                Ok(Event::Start(ref e)) => {
                    let prefixes = word_prefixes_at(e, word_prefixes)?;
                    let name = e.name();
                    let slot = if is_word_element(name.as_ref(), b"top", &prefixes) {
                        Some(&mut bdr.top)
                    } else if is_word_element(name.as_ref(), b"bottom", &prefixes) {
                        Some(&mut bdr.bottom)
                    } else if is_word_element(name.as_ref(), b"left", &prefixes)
                        || is_word_element(name.as_ref(), b"start", &prefixes)
                    {
                        Some(&mut bdr.left)
                    } else if is_word_element(name.as_ref(), b"right", &prefixes)
                        || is_word_element(name.as_ref(), b"end", &prefixes)
                    {
                        Some(&mut bdr.right)
                    } else if is_word_element(name.as_ref(), b"between", &prefixes) {
                        Some(&mut bdr.between)
                    } else if is_word_element(name.as_ref(), b"bar", &prefixes) {
                        Some(&mut bdr.bar)
                    } else {
                        None
                    };
                    if let Some(slot) = slot {
                        let mut edge = CT_BorderEdge::from_xml_attrs_with_prefixes(e, &prefixes)?;
                        let trim_start = reader.config().trim_text_start;
                        let trim_end = reader.config().trim_text_end;
                        reader.config_mut().trim_text(false);
                        let captured = crate::raw_xml::capture_element(reader, e);
                        reader.config_mut().trim_text_start = trim_start;
                        reader.config_mut().trim_text_end = trim_end;
                        let mut bindings = crate::numbering::namespace_bindings(&prefixes);
                        for prefix in prefixes.iter().filter(|prefix| !prefix.starts_with('\0')) {
                            if !bindings.iter().any(|(name, _)| name == prefix) {
                                bindings.push((prefix.clone(), W_NS.to_owned()));
                            }
                        }
                        edge.source_xml = Some(crate::text::raw_with_external_bindings(
                            &captured?, &bindings,
                        )?);
                        *slot = Some(edge);
                    } else {
                        // The whole unknown owner is opaque. Descendants are
                        // never direct pBdr children, even with Word-like names
                        // or a locally rebound prefix. The enclosing source
                        // owner retains their bytes for style serialization.
                        crate::raw_xml::capture_element(reader, e)?;
                    }
                }
                Ok(Event::End(ref e)) if matches_local_name(e.name().as_ref(), b"pBdr") => {
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(bdr)
    }

    pub(crate) fn from_source_xml(raw: &[u8]) -> Result<Self> {
        let mut reader = Reader::from_reader(raw);
        reader.config_mut().trim_text(false);
        let mut buffer = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer)? {
                Event::Start(root) => {
                    let prefixes = word_prefixes_at(&root, &[])?;
                    if !is_word_element(root.name().as_ref(), b"pBdr", &prefixes) {
                        return Err(OxmlError::InvalidValue(
                            "expected owned Word pBdr root".to_owned(),
                        ));
                    }
                    return Self::from_xml_with_prefixes(&mut reader, &prefixes);
                }
                Event::Empty(root) => {
                    let prefixes = word_prefixes_at(&root, &[])?;
                    if is_word_element(root.name().as_ref(), b"pBdr", &prefixes) {
                        return Ok(Self::default());
                    }
                    return Err(OxmlError::InvalidValue(
                        "expected owned Word pBdr root".to_owned(),
                    ));
                }
                Event::Eof => return Err(OxmlError::InvalidValue("pBdr root missing".to_owned())),
                _ => {}
            }
            buffer.clear();
        }
    }

    pub(crate) fn merge_source(
        raw: &[u8],
        original_selected: &[u8],
        selected: Option<&Self>,
        selected_owner: bool,
    ) -> Result<Vec<u8>> {
        let before = Self::from_source_xml(if selected.is_none() {
            raw
        } else {
            original_selected
        })?;
        if selected.is_some_and(|selected| selected == &before) {
            return Ok(raw.to_vec());
        }
        let mut before_writer = Writer::new(Vec::new());
        before.to_xml(&mut before_writer)?;
        let mut after_writer = Writer::new(Vec::new());
        selected
            .cloned()
            .unwrap_or_default()
            .to_xml(&mut after_writer)?;
        crate::styles::merge_style_border_source(
            raw,
            Some(before_writer.get_ref()),
            after_writer.get_ref(),
            &[],
            selected_owner,
        )
    }

    pub fn to_xml<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_event(Event::Start(BytesStart::new("w:pBdr")))?;

        if let Some(ref edge) = self.top {
            edge.to_xml(writer, "w:top")?;
        }
        if let Some(ref edge) = self.left {
            edge.to_xml(writer, "w:left")?;
        }
        if let Some(ref edge) = self.bottom {
            edge.to_xml(writer, "w:bottom")?;
        }
        if let Some(ref edge) = self.right {
            edge.to_xml(writer, "w:right")?;
        }
        if let Some(ref edge) = self.between {
            edge.to_xml(writer, "w:between")?;
        }
        if let Some(ref edge) = self.bar {
            edge.to_xml(writer, "w:bar")?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:pBdr")))?;
        Ok(())
    }

    pub fn is_empty(&self) -> bool {
        self.top.is_none()
            && self.bottom.is_none()
            && self.left.is_none()
            && self.right.is_none()
            && self.between.is_none()
            && self.bar.is_none()
    }
}

/// A single tab stop definition.
#[derive(Debug, Clone)]
pub struct CT_TabStop {
    /// Tab stop alignment
    pub val: ST_TabJc,
    /// Position in twips
    pub pos: Twips,
    /// Leader character
    pub leader: Option<ST_TabLeader>,
    /// Original occurrence in a parsed tab collection, or `None` for a new tab.
    ///
    /// This preservation value is ignored by semantic equality. Callers moving
    /// a tab into another `CT_Tabs` collection should set it to `None`.
    pub source_occurrence: Option<usize>,
}

// The Python binding keeps the Word owners it read for a body block while
// the block compares equal. `source_occurrence` only orders `w:tabs`, which
// no owner reads, but it can change the bytes and so the fingerprint of a
// table cell or text box story holding the paragraph.
impl PartialEq for CT_TabStop {
    fn eq(&self, other: &Self) -> bool {
        self.val == other.val && self.pos == other.pos && self.leader == other.leader
    }
}

impl CT_TabStop {
    pub fn new(val: ST_TabJc, pos: Twips) -> Self {
        CT_TabStop {
            val,
            pos,
            leader: None,
            source_occurrence: None,
        }
    }

    pub fn from_xml_attrs(e: &BytesStart) -> Result<Self> {
        Self::from_xml_attrs_with_prefixes(e, &["w".to_string()])
    }

    /// Parse a tab stop using the in-scope WordprocessingML prefixes.
    pub fn from_xml_attrs_with_prefixes(e: &BytesStart, word_prefixes: &[String]) -> Result<Self> {
        let prefixes = word_prefixes_at(e, word_prefixes)?;
        let mut val = ST_TabJc::Left;
        let mut pos = Twips(0);
        let mut leader = None;

        for attr in e.attributes() {
            let attr = attr?;
            let key = attr.key.as_ref();
            let v = std::str::from_utf8(&attr.value)?;
            if is_word_attribute(key, b"val", &prefixes) {
                val = ST_TabJc::from_str(v).unwrap_or(val);
            } else if is_word_attribute(key, b"pos", &prefixes) {
                pos = Twips(parse_integer_measurement(e.name().as_ref(), key, v)?);
            } else if is_word_attribute(key, b"leader", &prefixes) {
                leader = ST_TabLeader::from_str(v).ok();
            }
        }

        Ok(CT_TabStop {
            val,
            pos,
            leader,
            source_occurrence: None,
        })
    }
}

fn is_word_name(name: &[u8], word_prefixes: &[String]) -> bool {
    let Some(separator) = name.iter().position(|byte| *byte == b':') else {
        return word_prefixes.iter().any(String::is_empty);
    };
    word_prefixes
        .iter()
        .any(|prefix| prefix.as_bytes() == &name[..separator])
}

fn is_word_element(name: &[u8], local: &[u8], word_prefixes: &[String]) -> bool {
    matches_local_name(name, local) && is_word_name(name, word_prefixes)
}

fn is_word_attribute(key: &[u8], local: &[u8], word_prefixes: &[String]) -> bool {
    let Some(separator) = key.iter().position(|byte| *byte == b':') else {
        return false;
    };
    key.get(separator + 1..) == Some(local)
        && word_prefixes
            .iter()
            .any(|prefix| prefix.as_bytes() == &key[..separator])
}

/// `CT_Tabs` — Collection of tab stop definitions.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct CT_Tabs {
    pub tabs: Vec<CT_TabStop>,
}

impl CT_Tabs {
    pub fn from_xml(reader: &mut Reader<&[u8]>) -> Result<Self> {
        Self::from_xml_with_prefixes(reader, &["w".to_string()])
    }

    /// Parse tab stops using the in-scope WordprocessingML prefixes.
    pub fn from_xml_with_prefixes(
        reader: &mut Reader<&[u8]>,
        word_prefixes: &[String],
    ) -> Result<Self> {
        let mut tabs = Vec::new();
        let mut buf = Vec::new();
        let mut scopes = vec![word_prefixes.to_vec()];

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(ref e)) => {
                    let prefixes = word_prefixes_at(e, scopes.last().expect("outer scope"))?;
                    if scopes.len() == 1 && is_word_element(e.name().as_ref(), b"tab", &prefixes) {
                        let mut tab = CT_TabStop::from_xml_attrs_with_prefixes(e, &prefixes)?;
                        tab.source_occurrence = Some(tabs.len());
                        tabs.push(tab);
                    }
                }
                Ok(Event::Start(ref e)) => {
                    if scopes.len() >= MAX_TAB_XML_DEPTH {
                        return Err(OxmlError::InvalidValue(format!(
                            "tab XML depth exceeds {MAX_TAB_XML_DEPTH}"
                        )));
                    }
                    let prefixes = word_prefixes_at(e, scopes.last().expect("outer scope"))?;
                    if scopes.len() == 1 && is_word_element(e.name().as_ref(), b"tab", &prefixes) {
                        let mut tab = CT_TabStop::from_xml_attrs_with_prefixes(e, &prefixes)?;
                        tab.source_occurrence = Some(tabs.len());
                        tabs.push(tab);
                    }
                    scopes.push(prefixes);
                }
                Ok(Event::End(ref e)) => {
                    if scopes.len() == 1 && is_word_element(e.name().as_ref(), b"tabs", &scopes[0])
                    {
                        break;
                    }
                    if scopes.len() > 1 {
                        scopes.pop();
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(CT_Tabs { tabs })
    }

    /// Parse a retained standalone tabs root with its actual namespace bindings.
    pub fn from_source_xml(raw: &[u8]) -> Result<Self> {
        let mut reader = Reader::from_reader(raw);
        reader.config_mut().trim_text(false);
        let mut buffer = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer)? {
                Event::Start(root) => {
                    let prefixes = word_prefixes_at(&root, &[])?;
                    if !is_word_element(root.name().as_ref(), b"tabs", &prefixes) {
                        return Err(OxmlError::InvalidValue(
                            "expected owned Word tabs root".to_owned(),
                        ));
                    }
                    return Self::from_xml_with_prefixes(&mut reader, &prefixes);
                }
                Event::Empty(root) => {
                    let prefixes = word_prefixes_at(&root, &[])?;
                    if is_word_element(root.name().as_ref(), b"tabs", &prefixes) {
                        return Ok(Self::default());
                    }
                    return Err(OxmlError::InvalidValue(
                        "expected owned Word tabs root".to_owned(),
                    ));
                }
                Event::Eof => return Err(OxmlError::InvalidValue("tabs root missing".to_owned())),
                _ => {}
            }
            buffer.clear();
        }
    }

    /// Project exact known alignment and leader tokens from one owned source occurrence.
    pub fn source_tokens(
        raw: &[u8],
        occurrence: usize,
    ) -> Result<(ST_TabJc, Option<ST_TabLeader>)> {
        Self::from_source_xml(raw)?;
        crate::numbering::tabs_source_tokens(raw, occurrence)
    }

    pub fn to_xml<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        if self.tabs.is_empty() {
            return Ok(());
        }

        writer.write_event(Event::Start(BytesStart::new("w:tabs")))?;

        let mut buf = itoa::Buffer::new();
        for tab in &self.tabs {
            let mut e = BytesStart::new("w:tab");
            e.push_attribute(("w:val", tab.val.to_str()));
            e.push_attribute(("w:pos", buf.format(tab.pos.0)));
            if let Some(leader) = tab.leader {
                e.push_attribute(("w:leader", leader.to_str()));
            }
            writer.write_event(Event::Empty(e))?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:tabs")))?;
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn border_edge_retains_shadow_frame_and_theme_attributes() {
        let source = concat!(
            r#"<w:pBdr><w:top xmlns:ext="urn:producer" w:val="single" w:sz="8""#,
            r#" w:shadow="1" w:frame="1" w:themeColor="accent1" w:themeTint="66""#,
            r#" w:themeShade="BF" ext:mark="kept" w:space="1" w:color="4F81BD"/></w:pBdr>"#,
        );
        let mut reader = Reader::from_str(source);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if matches_local_name(e.name().as_ref(), b"pBdr") => break,
                Ok(Event::Eof) => panic!("missing pBdr start"),
                _ => {}
            }
            buf.clear();
        }
        let borders = CT_PBdr::from_xml(&mut reader).unwrap();
        let top = borders.top.clone().expect("a typed top edge");
        assert_eq!(top.val, ST_Border::Single);
        assert_eq!(top.sz, Some(8));
        assert_eq!(top.space, Some(1));
        assert_eq!(top.color.as_deref(), Some("4F81BD"));
        assert_eq!(
            top.extra_attributes,
            vec![
                ("xmlns:ext".to_owned(), "urn:producer".to_owned()),
                ("w:shadow".to_owned(), "1".to_owned()),
                ("w:frame".to_owned(), "1".to_owned()),
                ("w:themeColor".to_owned(), "accent1".to_owned()),
                ("w:themeTint".to_owned(), "66".to_owned()),
                ("w:themeShade".to_owned(), "BF".to_owned()),
                ("ext:mark".to_owned(), "kept".to_owned()),
            ]
        );

        let mut output = Vec::new();
        borders.to_xml(&mut Writer::new(&mut output)).unwrap();
        assert_eq!(
            String::from_utf8(output).unwrap(),
            concat!(
                r#"<w:pBdr><w:top xmlns:ext="urn:producer" w:shadow="1" w:frame="1""#,
                r#" w:themeColor="accent1" w:themeTint="66" w:themeShade="BF" ext:mark="kept""#,
                r#" w:val="single" w:sz="8" w:space="1" w:color="4F81BD"/></w:pBdr>"#,
            )
        );
    }

    #[test]
    fn nil_border_spelling_survives_reserialization() {
        let source = concat!(
            r#"<q:pBdr xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main">"#,
            r#"<q:top q:val="nil"/><q:bottom q:val="none"/></q:pBdr>"#,
        );
        let mut reader = Reader::from_str(source);
        loop {
            match reader.read_event() {
                Ok(Event::Start(_)) => break,
                Ok(Event::Eof) => panic!("missing pBdr start"),
                _ => {}
            }
        }
        let mut borders = CT_PBdr::from_xml_with_prefixes(&mut reader, &["q".to_owned()]).unwrap();
        let top = borders.top.as_ref().unwrap();
        assert_eq!(top.val, ST_Border::Nil);
        assert_eq!(borders.bottom.as_ref().unwrap().val, ST_Border::None);

        let mut output = Vec::new();
        borders.to_xml(&mut Writer::new(&mut output)).unwrap();
        let output = String::from_utf8(output).unwrap();
        assert!(output.contains(r#"<w:top w:val="nil"/>"#), "{output}");
        assert!(output.contains(r#"<w:bottom w:val="none"/>"#), "{output}");

        borders.top.as_mut().unwrap().val = ST_Border::Single;
        let mut output = Vec::new();
        borders.to_xml(&mut Writer::new(&mut output)).unwrap();
        assert!(
            String::from_utf8(output)
                .unwrap()
                .contains(r#"<w:top w:val="single"/>"#)
        );
    }

    #[test]
    fn round_trip_borders() {
        let bdr = CT_PBdr {
            top: Some(CT_BorderEdge {
                val: ST_Border::Single,
                sz: Some(4),
                space: Some(1),
                color: Some("000000".to_string()),
                extra_attributes: Vec::new(),
                source_xml: None,
            }),
            bottom: Some(CT_BorderEdge {
                val: ST_Border::Double,
                sz: Some(6),
                space: Some(2),
                color: Some("FF0000".to_string()),
                extra_attributes: Vec::new(),
                source_xml: None,
            }),
            ..Default::default()
        };

        let mut output = Vec::new();
        let mut writer = Writer::new(&mut output);
        bdr.to_xml(&mut writer).unwrap();
        let xml = String::from_utf8(output).unwrap();

        let full = xml.to_string();
        let mut reader = Reader::from_str(&full);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        // Skip to pBdr start
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if matches_local_name(e.name().as_ref(), b"pBdr") => {
                    break;
                }
                _ => {}
            }
            buf.clear();
        }
        let parsed = CT_PBdr::from_xml(&mut reader).unwrap();

        assert_eq!(parsed.top.as_ref().unwrap().val, ST_Border::Single);
        assert_eq!(parsed.top.as_ref().unwrap().sz, Some(4));
        assert_eq!(parsed.bottom.as_ref().unwrap().val, ST_Border::Double);
        assert!(parsed.left.is_none());
    }

    #[test]
    fn round_trip_tabs() {
        let tabs = CT_Tabs {
            tabs: vec![
                CT_TabStop {
                    val: ST_TabJc::Left,
                    pos: Twips(720),
                    leader: None,
                    source_occurrence: None,
                },
                CT_TabStop {
                    val: ST_TabJc::Center,
                    pos: Twips(4320),
                    leader: Some(ST_TabLeader::Dot),
                    source_occurrence: None,
                },
                CT_TabStop {
                    val: ST_TabJc::Right,
                    pos: Twips(8640),
                    leader: Some(ST_TabLeader::Hyphen),
                    source_occurrence: None,
                },
            ],
        };

        let mut output = Vec::new();
        let mut writer = Writer::new(&mut output);
        tabs.to_xml(&mut writer).unwrap();
        let xml = String::from_utf8(output).unwrap();

        let mut reader = Reader::from_str(&xml);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if matches_local_name(e.name().as_ref(), b"tabs") => {
                    break;
                }
                _ => {}
            }
            buf.clear();
        }
        let parsed = CT_Tabs::from_xml(&mut reader).unwrap();

        assert_eq!(parsed.tabs.len(), 3);
        assert_eq!(parsed.tabs[0].val, ST_TabJc::Left);
        assert_eq!(parsed.tabs[0].pos, Twips(720));
        assert_eq!(parsed.tabs[1].val, ST_TabJc::Center);
        assert_eq!(parsed.tabs[1].leader, Some(ST_TabLeader::Dot));
        assert_eq!(parsed.tabs[2].val, ST_TabJc::Right);
        assert_eq!(parsed.tabs[0].source_occurrence, Some(0));
        assert_eq!(parsed.tabs[1].source_occurrence, Some(1));
    }

    #[test]
    fn namespace_aware_tabs_ignore_foreign_same_local_children() {
        let xml = r#"<q:tabs xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:ext="urn:producer"><ext:tab ext:val="right" ext:pos="99"/><q:tab q:val="left" q:pos="720"/></q:tabs>"#;
        let mut reader = Reader::from_str(xml);
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(_)) => break,
                Ok(Event::Eof) => panic!("missing tabs start"),
                _ => {}
            }
            buf.clear();
        }
        let parsed = CT_Tabs::from_xml_with_prefixes(&mut reader, &["q".to_string()]).unwrap();
        assert_eq!(parsed.tabs.len(), 1);
        assert_eq!(parsed.tabs[0].pos, Twips(720));
        assert_eq!(parsed.tabs[0].source_occurrence, Some(0));

        let mut constructed = CT_TabStop::new(ST_TabJc::Left, Twips(720));
        assert_eq!(parsed.tabs[0], constructed);
        constructed.source_occurrence = Some(99);
        assert_eq!(parsed.tabs[0], constructed);
    }

    #[test]
    fn namespace_aware_tabs_track_shadows_and_expanded_tab_elements() {
        let xml = r#"<q:tabs xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><foreign xmlns="urn:producer" xmlns:q="urn:producer"><q:tab q:val="right" q:pos="99"/><q:tabs><q:tab q:val="right" q:pos="100"/></q:tabs></foreign><q:tab q:val="left" q:pos="720"></q:tab><q:tab q:val="right" q:pos="1440"/></q:tabs>"#;
        let mut reader = Reader::from_str(xml);
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(_)) => break,
                Ok(Event::Eof) => panic!("missing tabs start"),
                event => {
                    event.unwrap();
                }
            }
            buf.clear();
        }

        let parsed = CT_Tabs::from_xml_with_prefixes(&mut reader, &["q".to_string()]).unwrap();
        assert_eq!(parsed.tabs.len(), 2);
        assert_eq!(parsed.tabs[0].pos, Twips(720));
        assert_eq!(parsed.tabs[0].source_occurrence, Some(0));
        assert_eq!(parsed.tabs[1].pos, Twips(1440));
        assert_eq!(parsed.tabs[1].source_occurrence, Some(1));

        let default_xml = r#"<tabs xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><foreign xmlns="urn:producer"><tab val="right" pos="99"/><tabs><tab val="right" pos="100"/></tabs></foreign><tab xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="center" w:pos="2160"></tab></tabs>"#;
        let mut reader = Reader::from_str(default_xml);
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(_)) => break,
                Ok(Event::Eof) => panic!("missing default tabs start"),
                event => {
                    event.unwrap();
                }
            }
            buf.clear();
        }
        let parsed = CT_Tabs::from_xml_with_prefixes(&mut reader, &[String::new()]).unwrap();
        assert_eq!(parsed.tabs.len(), 1);
        assert_eq!(parsed.tabs[0].val, ST_TabJc::Center);
        assert_eq!(parsed.tabs[0].pos, Twips(2160));
        assert_eq!(parsed.tabs[0].source_occurrence, Some(0));
    }

    #[test]
    fn namespace_aware_tabs_reject_deep_distinct_aliases_normally() {
        let mut xml = format!(r#"<q:tabs xmlns:q="{W_NS}">"#);
        for depth in 0..MAX_TAB_XML_DEPTH {
            xml.push_str(&format!(r#"<n{depth}:unknown xmlns:n{depth}="{W_NS}">"#));
        }
        for depth in (0..MAX_TAB_XML_DEPTH).rev() {
            xml.push_str(&format!("</n{depth}:unknown>"));
        }
        xml.push_str("</q:tabs>");

        let mut reader = Reader::from_str(&xml);
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(_)) => break,
                Ok(Event::Eof) => panic!("missing tabs start"),
                event => {
                    event.unwrap();
                }
            }
            buf.clear();
        }
        let error = CT_Tabs::from_xml_with_prefixes(&mut reader, &["q".to_string()])
            .expect_err("deep alias nesting must be bounded");
        assert!(error.to_string().contains("tab XML depth exceeds 64"));
    }

    #[test]
    fn border_edge_all_styles_round_trip() {
        // Test that all border styles serialize and deserialize correctly
        let styles = [
            ST_Border::None,
            ST_Border::Single,
            ST_Border::Thick,
            ST_Border::Double,
            ST_Border::Dotted,
            ST_Border::Dashed,
            ST_Border::DotDash,
            ST_Border::Wave,
        ];

        for &style in &styles {
            let bdr = CT_PBdr {
                top: Some(CT_BorderEdge {
                    val: style,
                    sz: Some(8),
                    space: Some(0),
                    color: Some("FF00FF".to_string()),
                    extra_attributes: Vec::new(),
                    source_xml: None,
                }),
                ..Default::default()
            };

            let mut output = Vec::new();
            let mut writer = Writer::new(&mut output);
            bdr.to_xml(&mut writer).unwrap();
            let xml = String::from_utf8(output).unwrap();

            let mut reader = Reader::from_str(&xml);
            reader.config_mut().trim_text(true);
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf) {
                    Ok(Event::Start(ref e)) if matches_local_name(e.name().as_ref(), b"pBdr") => {
                        break;
                    }
                    _ => {}
                }
                buf.clear();
            }
            let parsed = CT_PBdr::from_xml(&mut reader).unwrap();
            let top = parsed.top.as_ref().unwrap();
            assert_eq!(
                top.val, style,
                "Border style round-trip failed for {style:?}"
            );
            assert_eq!(top.sz, Some(8));
            assert_eq!(top.color.as_deref(), Some("FF00FF"));
        }
    }
    #[test]
    fn fx200_border_parser_never_promotes_opaque_wrapper_descendants_or_shadowed_prefixes() {
        for opaque in [
            r#"<f:opaque><q:bottom q:val='double'></q:bottom></f:opaque>"#,
            r#"<f:opaque xmlns:q='urn:shadow'><q:bottom q:val='double'></q:bottom><w:left w:val='double'/></f:opaque>"#,
        ] {
            let source = format!(
                r#"<w:pPr xmlns:w="{W_NS}" xmlns:q="{W_NS}" xmlns:f="urn:foreign"><q:pBdr>{opaque}<q:right q:val='single'></q:right></q:pBdr></w:pPr>"#
            );
            let parsed =
                crate::numbering::parse_scoped_ppr(source.as_bytes(), &["w".to_owned()]).unwrap();
            let borders = parsed.borders.as_ref().unwrap();
            assert!(borders.bottom.is_none() && borders.left.is_none() && borders.top.is_none());
            assert_eq!(borders.right.as_ref().unwrap().val, ST_Border::Single);
            assert!(borders.right.as_ref().unwrap().source_xml.is_some());
            let mut writer = Writer::new(Vec::new());
            borders.to_xml(&mut writer).unwrap();
            let wrapped = format!(
                r#"<w:pPr xmlns:w="{W_NS}">{}</w:pPr>"#,
                std::str::from_utf8(writer.get_ref()).unwrap()
            );
            let reopened =
                crate::numbering::parse_scoped_ppr(wrapped.as_bytes(), &["w".to_owned()]).unwrap();
            let actual = reopened.borders.as_ref().unwrap();
            assert!(actual.bottom.is_none() && actual.left.is_none());
            assert_eq!(actual.right.as_ref().unwrap().val, ST_Border::Single);
        }
    }
}
