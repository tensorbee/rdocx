//! Numbering definitions: `CT_Numbering`, `CT_AbstractNum`, `CT_Num`, `CT_Lvl`.
//!
//! These types represent the content of `numbering.xml`, which defines
//! abstract numbering formats and numbering instances that paragraphs reference.

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};

use crate::error::Result;
use crate::namespace::{W_NS, matches_local_name};
use crate::properties::{CT_PPr, CT_RPr, get_val_attr};
use crate::shared::ST_Jc;

/// `ST_NumberFormat` — Numbering format type.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ST_NumberFormat {
    Decimal,
    UpperRoman,
    LowerRoman,
    UpperLetter,
    LowerLetter,
    Ordinal,
    Bullet,
    None,
}

impl ST_NumberFormat {
    pub fn from_str(s: &str) -> Self {
        match s {
            "decimal" => Self::Decimal,
            "upperRoman" => Self::UpperRoman,
            "lowerRoman" => Self::LowerRoman,
            "upperLetter" => Self::UpperLetter,
            "lowerLetter" => Self::LowerLetter,
            "ordinal" => Self::Ordinal,
            "bullet" => Self::Bullet,
            "none" => Self::None,
            _ => Self::Decimal,
        }
    }

    pub fn to_str(self) -> &'static str {
        match self {
            Self::Decimal => "decimal",
            Self::UpperRoman => "upperRoman",
            Self::LowerRoman => "lowerRoman",
            Self::UpperLetter => "upperLetter",
            Self::LowerLetter => "lowerLetter",
            Self::Ordinal => "ordinal",
            Self::Bullet => "bullet",
            Self::None => "none",
        }
    }
}

/// `CT_Lvl` — A single level (0–8) in an abstract numbering definition.
#[derive(Debug, Clone, PartialEq)]
pub struct CT_Lvl {
    /// Level index (0–8)
    pub ilvl: u32,
    /// Starting number
    pub start: Option<u32>,
    /// Number format
    pub num_fmt: Option<ST_NumberFormat>,
    /// Level text (e.g., "%1.", "%1.%2.", bullet char)
    pub lvl_text: Option<String>,
    /// Level justification
    pub lvl_jc: Option<ST_Jc>,
    /// Paragraph properties for this level (typically indentation)
    pub ppr: Option<CT_PPr>,
    /// Run properties for the numbering symbol
    pub rpr: Option<CT_RPr>,
}

#[allow(non_snake_case)]
impl CT_Lvl {
    pub fn new(ilvl: u32) -> Self {
        CT_Lvl {
            ilvl,
            start: None,
            num_fmt: None,
            lvl_text: None,
            lvl_jc: None,
            ppr: None,
            rpr: None,
        }
    }

    pub fn from_xml(reader: &mut Reader<&[u8]>, ilvl: u32) -> Result<Self> {
        let mut lvl = CT_Lvl::new(ilvl);
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let name = e.name();
                    if matches_local_name(name.as_ref(), b"pPr") {
                        lvl.ppr = Some(CT_PPr::from_xml(reader)?);
                    } else if matches_local_name(name.as_ref(), b"rPr") {
                        lvl.rpr = Some(CT_RPr::from_xml(reader)?);
                    } else {
                        reader.read_to_end_into(name, &mut Vec::new())?;
                    }
                }
                Ok(Event::Empty(ref e)) => {
                    let name = e.name();
                    if matches_local_name(name.as_ref(), b"start") {
                        if let Some(val) = get_val_attr(e)? {
                            lvl.start = Some(val.parse()?);
                        }
                    } else if matches_local_name(name.as_ref(), b"numFmt") {
                        if let Some(val) = get_val_attr(e)? {
                            lvl.num_fmt = Some(ST_NumberFormat::from_str(&val));
                        }
                    } else if matches_local_name(name.as_ref(), b"lvlText") {
                        lvl.lvl_text = get_val_attr(e)?;
                    } else if matches_local_name(name.as_ref(), b"lvlJc")
                        && let Some(val) = get_val_attr(e)?
                    {
                        lvl.lvl_jc = Some(ST_Jc::from_str(&val)?);
                    }
                }
                Ok(Event::End(ref e)) if matches_local_name(e.name().as_ref(), b"lvl") => {
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(lvl)
    }

    pub fn to_xml<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut buf = itoa::Buffer::new();
        let mut start = BytesStart::new("w:lvl");
        start.push_attribute(("w:ilvl", buf.format(self.ilvl)));
        writer.write_event(Event::Start(start))?;

        if let Some(s) = self.start {
            let mut e = BytesStart::new("w:start");
            e.push_attribute(("w:val", buf.format(s)));
            writer.write_event(Event::Empty(e))?;
        }

        if let Some(fmt) = self.num_fmt {
            let mut e = BytesStart::new("w:numFmt");
            e.push_attribute(("w:val", fmt.to_str()));
            writer.write_event(Event::Empty(e))?;
        }

        if let Some(ref text) = self.lvl_text {
            let mut e = BytesStart::new("w:lvlText");
            e.push_attribute(("w:val", text.as_str()));
            writer.write_event(Event::Empty(e))?;
        }

        if let Some(jc) = self.lvl_jc {
            let mut e = BytesStart::new("w:lvlJc");
            e.push_attribute(("w:val", jc.to_str()));
            writer.write_event(Event::Empty(e))?;
        }

        if let Some(ref ppr) = self.ppr {
            ppr.to_xml(writer)?;
        }

        if let Some(ref rpr) = self.rpr {
            rpr.to_xml(writer)?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:lvl")))?;
        Ok(())
    }
}

/// `CT_AbstractNum` — An abstract numbering definition with up to 9 levels.
#[derive(Debug, Clone, PartialEq)]
pub struct CT_AbstractNum {
    pub abstract_num_id: u32,
    pub levels: Vec<CT_Lvl>,
    /// Optional multi-level type hint
    pub multi_level_type: Option<String>,
}

#[allow(non_snake_case)]
impl CT_AbstractNum {
    pub fn new(id: u32) -> Self {
        CT_AbstractNum {
            abstract_num_id: id,
            levels: Vec::new(),
            multi_level_type: None,
        }
    }

    pub fn from_xml(reader: &mut Reader<&[u8]>, abstract_num_id: u32) -> Result<Self> {
        let mut abs = CT_AbstractNum::new(abstract_num_id);
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let name = e.name();
                    if matches_local_name(name.as_ref(), b"lvl") {
                        let mut ilvl = 0u32;
                        for attr in e.attributes() {
                            let attr = attr?;
                            if matches_local_name(attr.key.as_ref(), b"ilvl") {
                                ilvl = std::str::from_utf8(&attr.value)?.parse()?;
                            }
                        }
                        abs.levels.push(CT_Lvl::from_xml(reader, ilvl)?);
                    } else {
                        reader.read_to_end_into(name, &mut Vec::new())?;
                    }
                }
                Ok(Event::Empty(ref e)) => {
                    let name = e.name();
                    if matches_local_name(name.as_ref(), b"multiLevelType") {
                        abs.multi_level_type = get_val_attr(e)?;
                    }
                }
                Ok(Event::End(ref e)) if matches_local_name(e.name().as_ref(), b"abstractNum") => {
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(abs)
    }

    pub fn to_xml<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut buf = itoa::Buffer::new();
        let mut start = BytesStart::new("w:abstractNum");
        start.push_attribute(("w:abstractNumId", buf.format(self.abstract_num_id)));
        writer.write_event(Event::Start(start))?;

        if let Some(ref mlt) = self.multi_level_type {
            let mut e = BytesStart::new("w:multiLevelType");
            e.push_attribute(("w:val", mlt.as_str()));
            writer.write_event(Event::Empty(e))?;
        }

        for lvl in &self.levels {
            lvl.to_xml(writer)?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:abstractNum")))?;
        Ok(())
    }
}

/// `CT_Num` — A numbering instance that references an abstract numbering definition.
#[derive(Debug, Clone, PartialEq)]
pub struct CT_Num {
    pub num_id: u32,
    pub abstract_num_id: u32,
}

#[allow(non_snake_case)]
impl CT_Num {
    pub fn from_xml(reader: &mut Reader<&[u8]>, num_id: u32) -> Result<Self> {
        let mut abstract_num_id = 0;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(ref e)) => {
                    if matches_local_name(e.name().as_ref(), b"abstractNumId")
                        && let Some(val) = get_val_attr(e)?
                    {
                        abstract_num_id = val.parse()?;
                    }
                }
                Ok(Event::Start(ref e)) => {
                    // Skip lvlOverride and other nested elements
                    reader.read_to_end_into(e.name(), &mut Vec::new())?;
                }
                Ok(Event::End(ref e)) if matches_local_name(e.name().as_ref(), b"num") => {
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(CT_Num {
            num_id,
            abstract_num_id,
        })
    }

    pub fn to_xml<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut buf = itoa::Buffer::new();
        let mut start = BytesStart::new("w:num");
        start.push_attribute(("w:numId", buf.format(self.num_id)));
        writer.write_event(Event::Start(start))?;

        let mut abs_ref = BytesStart::new("w:abstractNumId");
        abs_ref.push_attribute(("w:val", buf.format(self.abstract_num_id)));
        writer.write_event(Event::Empty(abs_ref))?;

        writer.write_event(Event::End(BytesEnd::new("w:num")))?;
        Ok(())
    }
}

/// `CT_Numbering` — Root element of the numbering definitions part.
#[derive(Debug, Clone, PartialEq)]
pub struct CT_Numbering {
    pub abstract_nums: Vec<CT_AbstractNum>,
    pub nums: Vec<CT_Num>,
}

#[allow(non_snake_case)]
impl CT_Numbering {
    pub fn new() -> Self {
        CT_Numbering {
            abstract_nums: Vec::new(),
            nums: Vec::new(),
        }
    }

    /// Parse from XML bytes (the content of numbering.xml).
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        let mut reader = Reader::from_reader(xml);
        reader.config_mut().trim_text(true);

        let mut abstract_nums = Vec::new();
        let mut nums = Vec::new();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let name = e.name();
                    if matches_local_name(name.as_ref(), b"abstractNum") {
                        let mut id = 0u32;
                        for attr in e.attributes() {
                            let attr = attr?;
                            if matches_local_name(attr.key.as_ref(), b"abstractNumId") {
                                id = std::str::from_utf8(&attr.value)?.parse()?;
                            }
                        }
                        abstract_nums.push(CT_AbstractNum::from_xml(&mut reader, id)?);
                    } else if matches_local_name(name.as_ref(), b"num") {
                        let mut id = 0u32;
                        for attr in e.attributes() {
                            let attr = attr?;
                            if matches_local_name(attr.key.as_ref(), b"numId") {
                                id = std::str::from_utf8(&attr.value)?.parse()?;
                            }
                        }
                        nums.push(CT_Num::from_xml(&mut reader, id)?);
                    } else if matches_local_name(name.as_ref(), b"numbering") {
                        // root element, continue
                    } else {
                        reader.read_to_end_into(name, &mut Vec::new())?;
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(CT_Numbering {
            abstract_nums,
            nums,
        })
    }

    /// Serialize to XML bytes.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        let mut writer = Writer::new_with_indent(Vec::new(), b' ', 2);

        writer.write_event(Event::Decl(BytesDecl::new(
            "1.0",
            Some("UTF-8"),
            Some("yes"),
        )))?;

        let mut start = BytesStart::new("w:numbering");
        start.push_attribute(("xmlns:w", W_NS));
        start.push_attribute((
            "xmlns:r",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        ));
        writer.write_event(Event::Start(start))?;

        for abs in &self.abstract_nums {
            abs.to_xml(&mut writer)?;
        }

        for num in &self.nums {
            num.to_xml(&mut writer)?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:numbering")))?;

        Ok(writer.into_inner())
    }

    /// Get the next available abstract numbering ID.
    pub fn next_abstract_num_id(&self) -> u32 {
        self.abstract_nums
            .iter()
            .map(|a| a.abstract_num_id)
            .max()
            .map(|m| m + 1)
            .unwrap_or(0)
    }

    /// Get the next available numbering instance ID.
    pub fn next_num_id(&self) -> u32 {
        self.nums
            .iter()
            .map(|n| n.num_id)
            .max()
            .map(|m| m + 1)
            .unwrap_or(1)
    }

    /// Create a bullet list definition and return its numId.
    pub fn add_bullet_list(&mut self) -> u32 {
        let abs_id = self.next_abstract_num_id();
        let num_id = self.next_num_id();

        let bullet_chars = [
            "\u{2022}", // bullet •
            "\u{25E6}", // white bullet ◦
            "\u{25AA}", // black small square ▪
            "\u{2022}", // repeat pattern
            "\u{25E6}", "\u{25AA}", "\u{2022}", "\u{25E6}", "\u{25AA}",
        ];

        let mut abs = CT_AbstractNum::new(abs_id);
        abs.multi_level_type = Some("hybridMultilevel".to_string());

        for i in 0..9u32 {
            let mut lvl = CT_Lvl::new(i);
            lvl.start = Some(1);
            lvl.num_fmt = Some(ST_NumberFormat::Bullet);
            lvl.lvl_text = Some(bullet_chars[i as usize].to_string());
            lvl.lvl_jc = Some(ST_Jc::Left);

            // Standard indentation: 720tw per level
            let indent = (i + 1) as i32 * 720;
            let ppr = CT_PPr {
                ind_left: Some(crate::units::Twips(indent)),
                ind_hanging: Some(crate::units::Twips(360)),
                ..Default::default()
            };
            lvl.ppr = Some(ppr);

            abs.levels.push(lvl);
        }

        self.abstract_nums.push(abs);
        self.nums.push(CT_Num {
            num_id,
            abstract_num_id: abs_id,
        });

        num_id
    }

    /// Create a numbered (decimal) list definition and return its numId.
    pub fn add_numbered_list(&mut self) -> u32 {
        let abs_id = self.next_abstract_num_id();
        let num_id = self.next_num_id();

        let formats = [
            (ST_NumberFormat::Decimal, "%1."),
            (ST_NumberFormat::LowerLetter, "%2."),
            (ST_NumberFormat::LowerRoman, "%3."),
            (ST_NumberFormat::Decimal, "%4."),
            (ST_NumberFormat::LowerLetter, "%5."),
            (ST_NumberFormat::LowerRoman, "%6."),
            (ST_NumberFormat::Decimal, "%7."),
            (ST_NumberFormat::LowerLetter, "%8."),
            (ST_NumberFormat::LowerRoman, "%9."),
        ];

        let mut abs = CT_AbstractNum::new(abs_id);
        abs.multi_level_type = Some("hybridMultilevel".to_string());

        for (i, (fmt, text)) in formats.iter().enumerate() {
            let mut lvl = CT_Lvl::new(i as u32);
            lvl.start = Some(1);
            lvl.num_fmt = Some(*fmt);
            lvl.lvl_text = Some(text.to_string());
            lvl.lvl_jc = Some(ST_Jc::Left);

            let indent = (i as i32 + 1) * 720;
            let ppr = CT_PPr {
                ind_left: Some(crate::units::Twips(indent)),
                ind_hanging: Some(crate::units::Twips(360)),
                ..Default::default()
            };
            lvl.ppr = Some(ppr);

            abs.levels.push(lvl);
        }

        self.abstract_nums.push(abs);
        self.nums.push(CT_Num {
            num_id,
            abstract_num_id: abs_id,
        });

        num_id
    }

    /// Look up the abstract numbering definition for a given numId.
    pub fn get_abstract_num_for(&self, num_id: u32) -> Option<&CT_AbstractNum> {
        let num = self.nums.iter().find(|n| n.num_id == num_id)?;
        self.abstract_nums
            .iter()
            .find(|a| a.abstract_num_id == num.abstract_num_id)
    }
}

impl Default for CT_Numbering {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::units::Twips;

    #[test]
    fn round_trip_numbering() {
        let mut numbering = CT_Numbering::new();
        let num_id = numbering.add_bullet_list();
        assert_eq!(num_id, 1);

        let xml = numbering.to_xml().unwrap();
        let parsed = CT_Numbering::from_xml(&xml).unwrap();

        assert_eq!(parsed.abstract_nums.len(), 1);
        assert_eq!(parsed.nums.len(), 1);
        assert_eq!(parsed.nums[0].num_id, 1);
        assert_eq!(parsed.nums[0].abstract_num_id, 0);

        let abs = &parsed.abstract_nums[0];
        assert_eq!(abs.levels.len(), 9);
        assert_eq!(abs.levels[0].num_fmt, Some(ST_NumberFormat::Bullet));
        assert_eq!(abs.levels[0].lvl_text, Some("\u{2022}".to_string()));
    }

    #[test]
    fn round_trip_numbered_list() {
        let mut numbering = CT_Numbering::new();
        let num_id = numbering.add_numbered_list();
        assert_eq!(num_id, 1);

        let xml = numbering.to_xml().unwrap();
        let parsed = CT_Numbering::from_xml(&xml).unwrap();

        let abs = &parsed.abstract_nums[0];
        assert_eq!(abs.levels[0].num_fmt, Some(ST_NumberFormat::Decimal));
        assert_eq!(abs.levels[0].lvl_text, Some("%1.".to_string()));
        assert_eq!(abs.levels[1].num_fmt, Some(ST_NumberFormat::LowerLetter));
    }

    #[test]
    fn multiple_lists() {
        let mut numbering = CT_Numbering::new();
        let bullet_id = numbering.add_bullet_list();
        let num_id = numbering.add_numbered_list();

        assert_eq!(bullet_id, 1);
        assert_eq!(num_id, 2);

        let xml = numbering.to_xml().unwrap();
        let parsed = CT_Numbering::from_xml(&xml).unwrap();

        assert_eq!(parsed.abstract_nums.len(), 2);
        assert_eq!(parsed.nums.len(), 2);
    }

    #[test]
    fn level_indentation() {
        let mut numbering = CT_Numbering::new();
        numbering.add_bullet_list();

        let abs = &numbering.abstract_nums[0];
        // Level 0: 720tw indent, 360tw hanging
        assert_eq!(
            abs.levels[0].ppr.as_ref().unwrap().ind_left,
            Some(Twips(720))
        );
        assert_eq!(
            abs.levels[0].ppr.as_ref().unwrap().ind_hanging,
            Some(Twips(360))
        );
        // Level 2: 2160tw indent
        assert_eq!(
            abs.levels[2].ppr.as_ref().unwrap().ind_left,
            Some(Twips(2160))
        );
    }

    #[test]
    fn parse_numbering_xml() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="720" w:hanging="360"/>
      </w:pPr>
    </w:lvl>
    <w:lvl w:ilvl="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="lowerLetter"/>
      <w:lvlText w:val="%2."/>
      <w:lvlJc w:val="left"/>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1">
    <w:abstractNumId w:val="0"/>
  </w:num>
</w:numbering>"#;

        let numbering = CT_Numbering::from_xml(xml).unwrap();
        assert_eq!(numbering.abstract_nums.len(), 1);
        assert_eq!(numbering.nums.len(), 1);

        let abs = &numbering.abstract_nums[0];
        assert_eq!(abs.abstract_num_id, 0);
        assert_eq!(abs.multi_level_type, Some("hybridMultilevel".to_string()));
        assert_eq!(abs.levels.len(), 2);
        assert_eq!(abs.levels[0].start, Some(1));
        assert_eq!(abs.levels[0].num_fmt, Some(ST_NumberFormat::Decimal));
        assert_eq!(abs.levels[0].lvl_text, Some("%1.".to_string()));
        assert_eq!(
            abs.levels[0].ppr.as_ref().unwrap().ind_left,
            Some(Twips(720))
        );
        assert_eq!(abs.levels[1].num_fmt, Some(ST_NumberFormat::LowerLetter));

        let num = &numbering.nums[0];
        assert_eq!(num.num_id, 1);
        assert_eq!(num.abstract_num_id, 0);
    }

    #[test]
    fn get_abstract_num_for_lookup() {
        let mut numbering = CT_Numbering::new();
        numbering.add_bullet_list();
        numbering.add_numbered_list();

        let abs = numbering.get_abstract_num_for(2).unwrap();
        assert_eq!(abs.levels[0].num_fmt, Some(ST_NumberFormat::Decimal));

        assert!(numbering.get_abstract_num_for(99).is_none());
    }
}
