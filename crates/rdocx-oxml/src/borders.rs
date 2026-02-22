//! Border and tab stop types for paragraph formatting.

use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};

use crate::error::Result;
use crate::namespace::matches_local_name;
use crate::shared::{ST_Border, ST_TabJc, ST_TabLeader};
use crate::units::Twips;

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
}

impl CT_BorderEdge {
    pub fn new(val: ST_Border) -> Self {
        CT_BorderEdge {
            val,
            sz: None,
            space: None,
            color: None,
        }
    }

    pub fn from_xml_attrs(e: &BytesStart) -> Result<Self> {
        let mut val = ST_Border::None;
        let mut sz = None;
        let mut space = None;
        let mut color = None;

        for attr in e.attributes() {
            let attr = attr?;
            let key = attr.key.as_ref();
            let v = std::str::from_utf8(&attr.value)?;
            if matches_local_name(key, b"val") {
                val = ST_Border::from_str(v)?;
            } else if matches_local_name(key, b"sz") {
                sz = Some(v.parse()?);
            } else if matches_local_name(key, b"space") {
                space = Some(v.parse()?);
            } else if matches_local_name(key, b"color") {
                color = Some(v.to_string());
            }
        }

        Ok(CT_BorderEdge {
            val,
            sz,
            space,
            color,
        })
    }

    pub fn write_xml_attrs(&self, e: &mut BytesStart) {
        let mut buf = itoa::Buffer::new();
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
        let mut bdr = CT_PBdr::default();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(ref e)) => {
                    let name = e.name();
                    let edge = CT_BorderEdge::from_xml_attrs(e)?;
                    if matches_local_name(name.as_ref(), b"top") {
                        bdr.top = Some(edge);
                    } else if matches_local_name(name.as_ref(), b"bottom") {
                        bdr.bottom = Some(edge);
                    } else if matches_local_name(name.as_ref(), b"left")
                        || matches_local_name(name.as_ref(), b"start")
                    {
                        bdr.left = Some(edge);
                    } else if matches_local_name(name.as_ref(), b"right")
                        || matches_local_name(name.as_ref(), b"end")
                    {
                        bdr.right = Some(edge);
                    } else if matches_local_name(name.as_ref(), b"between") {
                        bdr.between = Some(edge);
                    } else if matches_local_name(name.as_ref(), b"bar") {
                        bdr.bar = Some(edge);
                    }
                }
                Ok(Event::End(ref e))
                    if matches_local_name(e.name().as_ref(), b"pBdr") =>
                {
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
#[derive(Debug, Clone, PartialEq)]
pub struct CT_TabStop {
    /// Tab stop alignment
    pub val: ST_TabJc,
    /// Position in twips
    pub pos: Twips,
    /// Leader character
    pub leader: Option<ST_TabLeader>,
}

impl CT_TabStop {
    pub fn new(val: ST_TabJc, pos: Twips) -> Self {
        CT_TabStop {
            val,
            pos,
            leader: None,
        }
    }

    pub fn from_xml_attrs(e: &BytesStart) -> Result<Self> {
        let mut val = ST_TabJc::Left;
        let mut pos = Twips(0);
        let mut leader = None;

        for attr in e.attributes() {
            let attr = attr?;
            let key = attr.key.as_ref();
            let v = std::str::from_utf8(&attr.value)?;
            if matches_local_name(key, b"val") {
                val = ST_TabJc::from_str(v)?;
            } else if matches_local_name(key, b"pos") {
                pos = Twips(v.parse()?);
            } else if matches_local_name(key, b"leader") {
                leader = Some(ST_TabLeader::from_str(v)?);
            }
        }

        Ok(CT_TabStop { val, pos, leader })
    }
}

/// `CT_Tabs` — Collection of tab stop definitions.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct CT_Tabs {
    pub tabs: Vec<CT_TabStop>,
}

impl CT_Tabs {
    pub fn from_xml(reader: &mut Reader<&[u8]>) -> Result<Self> {
        let mut tabs = Vec::new();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(ref e))
                    if matches_local_name(e.name().as_ref(), b"tab") =>
                {
                    tabs.push(CT_TabStop::from_xml_attrs(e)?);
                }
                Ok(Event::End(ref e))
                    if matches_local_name(e.name().as_ref(), b"tabs") =>
                {
                    break;
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
    fn round_trip_borders() {
        let bdr = CT_PBdr {
            top: Some(CT_BorderEdge {
                val: ST_Border::Single,
                sz: Some(4),
                space: Some(1),
                color: Some("000000".to_string()),
            }),
            bottom: Some(CT_BorderEdge {
                val: ST_Border::Double,
                sz: Some(6),
                space: Some(2),
                color: Some("FF0000".to_string()),
            }),
            ..Default::default()
        };

        let mut output = Vec::new();
        let mut writer = Writer::new(&mut output);
        bdr.to_xml(&mut writer).unwrap();
        let xml = String::from_utf8(output).unwrap();

        let full = format!("{xml}");
        let mut reader = Reader::from_str(&full);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        // Skip to pBdr start
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e))
                    if matches_local_name(e.name().as_ref(), b"pBdr") =>
                {
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
                },
                CT_TabStop {
                    val: ST_TabJc::Center,
                    pos: Twips(4320),
                    leader: Some(ST_TabLeader::Dot),
                },
                CT_TabStop {
                    val: ST_TabJc::Right,
                    pos: Twips(8640),
                    leader: Some(ST_TabLeader::Hyphen),
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
                Ok(Event::Start(ref e))
                    if matches_local_name(e.name().as_ref(), b"tabs") =>
                {
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
                    Ok(Event::Start(ref e))
                        if matches_local_name(e.name().as_ref(), b"pBdr") =>
                    {
                        break;
                    }
                    _ => {}
                }
                buf.clear();
            }
            let parsed = CT_PBdr::from_xml(&mut reader).unwrap();
            let top = parsed.top.as_ref().unwrap();
            assert_eq!(top.val, style, "Border style round-trip failed for {style:?}");
            assert_eq!(top.sz, Some(8));
            assert_eq!(top.color.as_deref(), Some("FF00FF"));
        }
    }
}
