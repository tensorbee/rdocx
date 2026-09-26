//! Border and tab stop types for paragraph formatting.

use std::borrow::Cow;

use quick_xml::events::attributes::Attribute;
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::name::QName;
use quick_xml::{Reader, Writer, XmlVersion};

use crate::error::{OxmlError, Result};
use crate::namespace::{W_NS, matches_local_name};
use crate::shared::{ST_Border, ST_TabJc, ST_TabLeader};
use crate::units::Twips;

const MAX_TAB_XML_DEPTH: usize = 64;

/// A single border edge (top, bottom, left, right, between).
#[derive(Debug, Clone)]
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
    /// Whether an `ST_Border::None` style is spelled `nil` rather than `none`.
    ///
    /// The two tokens mean the same thing and render the same, so this only
    /// selects the token a save writes. A parsed edge sets it from its source
    /// token, and it is ignored for any other style and by equality.
    pub nil: bool,
}

impl PartialEq for CT_BorderEdge {
    fn eq(&self, other: &Self) -> bool {
        // Destructured so that a new field cannot be left out by accident.
        let CT_BorderEdge {
            val,
            sz,
            space,
            color,
            extra_attributes,
            nil: _,
        } = self;
        *val == other.val
            && *sz == other.sz
            && *space == other.space
            && *color == other.color
            && *extra_attributes == other.extra_attributes
    }
}

impl CT_BorderEdge {
    pub fn new(val: ST_Border) -> Self {
        CT_BorderEdge {
            val,
            sz: None,
            space: None,
            color: None,
            extra_attributes: Vec::new(),
            nil: false,
        }
    }

    pub fn from_xml_attrs(e: &BytesStart) -> Result<Self> {
        let mut edge = CT_BorderEdge::new(ST_Border::None);

        for attr in e.attributes() {
            let attr = attr?;
            let key = attr.key.as_ref();
            let v = std::str::from_utf8(&attr.value)?;
            if matches_local_name(key, b"val") {
                edge.val = ST_Border::from_str(v).unwrap_or(edge.val);
                edge.nil = v == "nil";
            } else if matches_local_name(key, b"sz") {
                edge.sz = Some(v.parse()?);
            } else if matches_local_name(key, b"space") {
                edge.space = Some(v.parse()?);
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
                edge.nil = value == "nil";
            } else if is_word_attribute(key, b"sz", word_prefixes) {
                edge.sz = Some(value.parse()?);
            } else if is_word_attribute(key, b"space", word_prefixes) {
                edge.space = Some(value.parse()?);
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
        let val = if self.nil && self.val == ST_Border::None {
            "nil"
        } else {
            self.val.to_str()
        };
        e.push_attribute(("w:val", val));
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
        writer.write_event(Event::Empty(e))?;
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

    pub fn to_xml<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_event(Event::Start(BytesStart::new("w:pBdr")))?;

        if let Some(ref edge) = self.top {
            let mut e = BytesStart::new("w:top");
            edge.write_xml_attrs(&mut e);
            writer.write_event(Event::Empty(e))?;
        }
        if let Some(ref edge) = self.left {
            let mut e = BytesStart::new("w:left");
            edge.write_xml_attrs(&mut e);
            writer.write_event(Event::Empty(e))?;
        }
        if let Some(ref edge) = self.bottom {
            let mut e = BytesStart::new("w:bottom");
            edge.write_xml_attrs(&mut e);
            writer.write_event(Event::Empty(e))?;
        }
        if let Some(ref edge) = self.right {
            let mut e = BytesStart::new("w:right");
            edge.write_xml_attrs(&mut e);
            writer.write_event(Event::Empty(e))?;
        }
        if let Some(ref edge) = self.between {
            let mut e = BytesStart::new("w:between");
            edge.write_xml_attrs(&mut e);
            writer.write_event(Event::Empty(e))?;
        }
        if let Some(ref edge) = self.bar {
            let mut e = BytesStart::new("w:bar");
            edge.write_xml_attrs(&mut e);
            writer.write_event(Event::Empty(e))?;
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
                pos = Twips(v.parse()?);
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

fn word_prefixes_at(start: &BytesStart<'_>, inherited: &[String]) -> Result<Vec<String>> {
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
    fn a_nil_border_edge_is_not_rewritten_as_none() {
        let source = concat!(
            r#"<q:pBdr xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main">"#,
            r#"<q:top q:val="nil"/><q:bottom q:val="none" q:sz="4"/></q:pBdr>"#,
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
        let top = borders.top.clone().expect("a typed top edge");
        assert_eq!(top.val, ST_Border::None);
        assert!(top.nil);
        assert!(!borders.bottom.as_ref().unwrap().nil);
        // Both tokens are the same style, so equality does not see the spelling.
        assert_eq!(top, CT_BorderEdge::new(ST_Border::None));

        let serialize = |borders: &CT_PBdr| {
            let mut output = Vec::new();
            borders.to_xml(&mut Writer::new(&mut output)).unwrap();
            String::from_utf8(output).unwrap()
        };
        assert_eq!(
            serialize(&borders),
            r#"<w:pBdr><w:top w:val="nil"/><w:bottom w:val="none" w:sz="4"/></w:pBdr>"#
        );

        // An authored edge writes `none`, and a changed style writes its own token.
        borders.bottom = Some(CT_BorderEdge::new(ST_Border::None));
        borders.top.as_mut().unwrap().val = ST_Border::Single;
        assert_eq!(
            serialize(&borders),
            r#"<w:pBdr><w:top w:val="single"/><w:bottom w:val="none"/></w:pBdr>"#
        );

        let start = BytesStart::from_content(r#"w:bdr w:val="nil" w:sz="0""#, 5);
        let edge = CT_BorderEdge::from_xml_attrs(&start).unwrap();
        assert_eq!(
            (edge.val, edge.nil, edge.sz),
            (ST_Border::None, true, Some(0))
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
                nil: false,
            }),
            bottom: Some(CT_BorderEdge {
                val: ST_Border::Double,
                sz: Some(6),
                space: Some(2),
                color: Some("FF0000".to_string()),
                extra_attributes: Vec::new(),
                nil: false,
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
                    nil: false,
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
}
