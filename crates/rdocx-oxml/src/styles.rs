//! Style elements: `CT_Styles`, `CT_Style`, `CT_DocDefaults`.

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};

use crate::error::{OxmlError, Result};
use crate::namespace::{W_NS, matches_local_name};
use crate::numbering::{parse_scoped_ppr, word_prefixes_at};
use crate::properties::{CT_PPr, CT_RPr, is_word_element};
use crate::raw_xml::capture_element;

/// The type of a style.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum StyleType {
    Paragraph,
    Character,
    Table,
    Numbering,
}

impl StyleType {
    pub fn from_str(s: &str) -> Result<Self> {
        match s {
            "paragraph" => Ok(StyleType::Paragraph),
            "character" => Ok(StyleType::Character),
            "table" => Ok(StyleType::Table),
            "numbering" => Ok(StyleType::Numbering),
            _ => Err(OxmlError::InvalidValue(format!("invalid style type: {s}"))),
        }
    }

    pub fn to_str(self) -> &'static str {
        match self {
            StyleType::Paragraph => "paragraph",
            StyleType::Character => "character",
            StyleType::Table => "table",
            StyleType::Numbering => "numbering",
        }
    }
}

/// `CT_Style` — A single style definition.
#[derive(Debug, Clone, PartialEq)]
#[allow(non_snake_case)]
pub struct CT_Style {
    pub style_id: String,
    pub style_type: StyleType,
    pub name: Option<String>,
    pub based_on: Option<String>,
    pub next_style: Option<String>,
    pub is_default: bool,
    pub ppr: Option<CT_PPr>,
    pub rpr: Option<CT_RPr>,
    /// Verbatim `<w:tblPr>` bytes of a table style, written back on save so
    /// table-style formatting is no longer dropped by the round trip.
    pub tbl_pr_xml: Option<Vec<u8>>,
    /// Parsed view of the table style's `w:tblBorders` (from `tbl_pr_xml`),
    /// used by layout to resolve borders for tables that reference the
    /// style instead of carrying direct borders.
    pub table_borders: Option<crate::table::CT_TblBorders>,
}

#[allow(non_snake_case)]
impl CT_Style {
    pub fn from_xml(reader: &mut Reader<&[u8]>, attrs: &BytesStart) -> Result<Self> {
        let prefixes = word_prefixes_at(attrs, &["w".to_string()])?;
        Self::from_xml_with_prefixes(reader, attrs, &prefixes)
    }

    fn from_xml_with_prefixes(
        reader: &mut Reader<&[u8]>,
        attrs: &BytesStart,
        word_prefixes: &[String],
    ) -> Result<Self> {
        let mut style_id = String::new();
        let mut style_type = StyleType::Paragraph;
        let mut is_default = false;

        for attr in attrs.attributes() {
            let attr = attr?;
            let key = attr.key.as_ref();
            if matches_local_name(key, b"styleId") {
                style_id = std::str::from_utf8(&attr.value)?.to_string();
            } else if matches_local_name(key, b"type") {
                style_type = StyleType::from_str(std::str::from_utf8(&attr.value)?)?;
            } else if matches_local_name(key, b"default") {
                is_default = std::str::from_utf8(&attr.value)? == "1"
                    || std::str::from_utf8(&attr.value)? == "true";
            }
        }

        let mut name = None;
        let mut based_on = None;
        let mut next_style = None;
        let mut ppr = None;
        let mut rpr = None;
        let mut tbl_pr_xml = None;
        let mut table_borders = None;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(ref e)) => {
                    let ename = e.name();
                    if matches_local_name(ename.as_ref(), b"name") {
                        name = get_val_attr(e)?;
                    } else if matches_local_name(ename.as_ref(), b"basedOn") {
                        based_on = get_val_attr(e)?;
                    } else if matches_local_name(ename.as_ref(), b"next") {
                        next_style = get_val_attr(e)?;
                    }
                }
                Ok(Event::Start(ref e)) => {
                    let ename = e.name();
                    let prefixes = word_prefixes_at(e, word_prefixes)?;
                    if is_word_element(ename.as_ref(), b"pPr", &prefixes) {
                        let raw = capture_element(reader, e)?;
                        ppr = Some(parse_scoped_ppr(&raw, word_prefixes)?);
                    } else if matches_local_name(ename.as_ref(), b"rPr") {
                        rpr = Some(CT_RPr::from_xml(reader)?);
                    } else if is_word_element(ename.as_ref(), b"tblPr", &prefixes) {
                        let raw = capture_element(reader, e)?;
                        table_borders = parse_style_table_borders(&raw);
                        tbl_pr_xml = Some(raw);
                    } else {
                        reader.read_to_end_into(ename, &mut Vec::new())?;
                    }
                }
                Ok(Event::End(ref e)) if matches_local_name(e.name().as_ref(), b"style") => {
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(CT_Style {
            style_id,
            style_type,
            name,
            based_on,
            next_style,
            is_default,
            ppr,
            rpr,
            tbl_pr_xml,
            table_borders,
        })
    }

    pub fn to_xml<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut e = BytesStart::new("w:style");
        e.push_attribute(("w:type", self.style_type.to_str()));
        e.push_attribute(("w:styleId", self.style_id.as_str()));
        if self.is_default {
            e.push_attribute(("w:default", "1"));
        }
        writer.write_event(Event::Start(e))?;

        if let Some(ref name) = self.name {
            let mut ne = BytesStart::new("w:name");
            ne.push_attribute(("w:val", name.as_str()));
            writer.write_event(Event::Empty(ne))?;
        }

        if let Some(ref based_on) = self.based_on {
            let mut be = BytesStart::new("w:basedOn");
            be.push_attribute(("w:val", based_on.as_str()));
            writer.write_event(Event::Empty(be))?;
        }

        if let Some(ref next) = self.next_style {
            let mut ne = BytesStart::new("w:next");
            ne.push_attribute(("w:val", next.as_str()));
            writer.write_event(Event::Empty(ne))?;
        }

        if let Some(ref ppr) = self.ppr {
            ppr.to_xml(writer)?;
        }
        if let Some(ref rpr) = self.rpr {
            rpr.to_xml(writer)?;
        }
        if let Some(ref raw) = self.tbl_pr_xml {
            writer.get_mut().write_all(raw)?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:style")))?;
        Ok(())
    }
}

/// `CT_DocDefaults` — Document-level default properties.
#[derive(Debug, Clone, Default, PartialEq)]
#[allow(non_snake_case)]
pub struct CT_DocDefaults {
    pub rpr: Option<CT_RPr>,
    pub ppr: Option<CT_PPr>,
}

#[allow(non_snake_case)]
impl CT_DocDefaults {
    pub fn from_xml(reader: &mut Reader<&[u8]>) -> Result<Self> {
        Self::from_xml_with_prefixes(reader, &["w".to_string()])
    }

    fn from_xml_with_prefixes(
        reader: &mut Reader<&[u8]>,
        word_prefixes: &[String],
    ) -> Result<Self> {
        let mut defaults = CT_DocDefaults::default();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let name = e.name();
                    let prefixes = word_prefixes_at(e, word_prefixes)?;
                    if matches_local_name(name.as_ref(), b"rPrDefault") {
                        // Read into rPrDefault, expecting rPr child
                        defaults.rpr = Self::parse_pr_default(reader, b"rPrDefault")?;
                    } else if matches_local_name(name.as_ref(), b"pPrDefault") {
                        defaults.ppr = Self::parse_ppr_default(reader, &prefixes)?;
                    } else {
                        reader.read_to_end_into(name, &mut Vec::new())?;
                    }
                }
                Ok(Event::End(ref e)) if matches_local_name(e.name().as_ref(), b"docDefaults") => {
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(defaults)
    }

    fn parse_pr_default(reader: &mut Reader<&[u8]>, end_tag: &[u8]) -> Result<Option<CT_RPr>> {
        let mut rpr = None;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let name = e.name();
                    if matches_local_name(name.as_ref(), b"rPr") {
                        rpr = Some(CT_RPr::from_xml(reader)?);
                    } else {
                        reader.read_to_end_into(name, &mut Vec::new())?;
                    }
                }
                Ok(Event::End(ref e)) if matches_local_name(e.name().as_ref(), end_tag) => {
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(rpr)
    }

    fn parse_ppr_default(
        reader: &mut Reader<&[u8]>,
        word_prefixes: &[String],
    ) -> Result<Option<CT_PPr>> {
        let mut ppr = None;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let name = e.name();
                    let prefixes = word_prefixes_at(e, word_prefixes)?;
                    if is_word_element(name.as_ref(), b"pPr", &prefixes) {
                        let raw = capture_element(reader, e)?;
                        ppr = Some(parse_scoped_ppr(&raw, word_prefixes)?);
                    } else {
                        reader.read_to_end_into(name, &mut Vec::new())?;
                    }
                }
                Ok(Event::End(ref e)) if matches_local_name(e.name().as_ref(), b"pPrDefault") => {
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(ppr)
    }

    pub fn to_xml<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_event(Event::Start(BytesStart::new("w:docDefaults")))?;

        if let Some(ref rpr) = self.rpr {
            writer.write_event(Event::Start(BytesStart::new("w:rPrDefault")))?;
            rpr.to_xml(writer)?;
            writer.write_event(Event::End(BytesEnd::new("w:rPrDefault")))?;
        }

        if let Some(ref ppr) = self.ppr {
            writer.write_event(Event::Start(BytesStart::new("w:pPrDefault")))?;
            ppr.to_xml(writer)?;
            writer.write_event(Event::End(BytesEnd::new("w:pPrDefault")))?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:docDefaults")))?;
        Ok(())
    }
}

/// `CT_Styles` — The styles part (word/styles.xml).
#[derive(Debug, Clone, PartialEq)]
#[allow(non_snake_case)]
pub struct CT_Styles {
    pub doc_defaults: Option<CT_DocDefaults>,
    pub styles: Vec<CT_Style>,
}

#[allow(non_snake_case)]
impl CT_Styles {
    pub fn new() -> Self {
        CT_Styles {
            doc_defaults: None,
            styles: Vec::new(),
        }
    }

    /// Parse from XML bytes (the content of word/styles.xml).
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        let mut reader = Reader::from_reader(xml);
        reader.config_mut().trim_text(true);

        let mut doc_defaults = None;
        let mut styles = Vec::new();
        let mut buf = Vec::new();
        let mut word_prefixes = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let name = e.name();
                    let prefixes = word_prefixes_at(e, &word_prefixes)?;
                    if matches_local_name(name.as_ref(), b"docDefaults") {
                        doc_defaults = Some(CT_DocDefaults::from_xml_with_prefixes(
                            &mut reader,
                            &prefixes,
                        )?);
                    } else if matches_local_name(name.as_ref(), b"style") {
                        styles.push(CT_Style::from_xml_with_prefixes(&mut reader, e, &prefixes)?);
                    } else if matches_local_name(name.as_ref(), b"styles") {
                        // Root element, continue
                        word_prefixes = prefixes;
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

        Ok(CT_Styles {
            doc_defaults,
            styles,
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

        let mut styles_start = BytesStart::new("w:styles");
        styles_start.push_attribute(("xmlns:w", W_NS));
        styles_start.push_attribute((
            "xmlns:r",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        ));
        writer.write_event(Event::Start(styles_start))?;

        if let Some(ref defaults) = self.doc_defaults {
            defaults.to_xml(&mut writer)?;
        }

        for style in &self.styles {
            style.to_xml(&mut writer)?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:styles")))?;

        Ok(writer.into_inner())
    }

    /// Find a style by its ID.
    pub fn get_by_id(&self, style_id: &str) -> Option<&CT_Style> {
        self.styles.iter().find(|s| s.style_id == style_id)
    }

    /// Find the default style for a given type.
    pub fn get_default(&self, style_type: StyleType) -> Option<&CT_Style> {
        self.styles
            .iter()
            .find(|s| s.style_type == style_type && s.is_default)
    }

    /// Create a minimal default styles part for a new document.
    pub fn new_default() -> Self {
        use crate::units::HalfPoint;

        let normal = CT_Style {
            style_id: "Normal".to_string(),
            style_type: StyleType::Paragraph,
            name: Some("Normal".to_string()),
            based_on: None,
            next_style: None,
            is_default: true,
            ppr: None,
            rpr: None,
            tbl_pr_xml: None,
            table_borders: None,
        };

        let heading1 = CT_Style {
            style_id: "Heading1".to_string(),
            style_type: StyleType::Paragraph,
            name: Some("heading 1".to_string()),
            based_on: Some("Normal".to_string()),
            next_style: Some("Normal".to_string()),
            is_default: false,
            ppr: Some(CT_PPr {
                keep_next: Some(true),
                keep_lines: Some(true),
                space_before: Some(crate::units::Twips(240)),
                space_after: Some(crate::units::Twips(0)),
                ..Default::default()
            }),
            rpr: Some(CT_RPr {
                sz: Some(HalfPoint(32)),
                sz_cs: Some(HalfPoint(32)),
                bold: Some(true),
                bold_cs: Some(true),
                color: Some("2F5496".to_string()),
                ..Default::default()
            }),
            tbl_pr_xml: None,
            table_borders: None,
        };

        let doc_defaults = CT_DocDefaults {
            rpr: Some(CT_RPr {
                font_ascii: Some("Calibri".to_string()),
                font_hansi: Some("Calibri".to_string()),
                font_east_asia: Some("Calibri".to_string()),
                font_cs: Some("Times New Roman".to_string()),
                sz: Some(HalfPoint(22)),
                sz_cs: Some(HalfPoint(22)),
                ..Default::default()
            }),
            ppr: Some(CT_PPr {
                space_after: Some(crate::units::Twips(160)),
                line_spacing: Some(crate::units::Twips(259)),
                line_rule: Some("auto".to_string()),
                ..Default::default()
            }),
        };

        CT_Styles {
            doc_defaults: Some(doc_defaults),
            styles: vec![normal, heading1],
        }
    }
}

impl Default for CT_Styles {
    fn default() -> Self {
        Self::new()
    }
}

/// Extract the `w:val` attribute from an element.
/// Parse the `w:tblBorders` inside a table style's raw `w:tblPr` bytes.
fn parse_style_table_borders(raw: &[u8]) -> Option<crate::table::CT_TblBorders> {
    let mut reader = Reader::from_reader(raw);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e))
                if matches_local_name(e.name().as_ref(), b"tblBorders") =>
            {
                return crate::table::CT_TblBorders::from_xml(&mut reader).ok();
            }
            Ok(Event::Eof) | Err(_) => return None,
            _ => {}
        }
        buf.clear();
    }
}

fn get_val_attr(e: &BytesStart) -> Result<Option<String>> {
    for attr in e.attributes() {
        let attr = attr?;
        if matches_local_name(attr.key.as_ref(), b"val") {
            return Ok(Some(std::str::from_utf8(&attr.value)?.to_string()));
        }
    }
    Ok(None)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn round_trip_styles() {
        let styles = CT_Styles::new_default();
        let xml = styles.to_xml().unwrap();
        let parsed = CT_Styles::from_xml(&xml).unwrap();

        assert_eq!(parsed.styles.len(), 2);
        assert!(parsed.doc_defaults.is_some());

        let normal = parsed.get_by_id("Normal").unwrap();
        assert_eq!(normal.name, Some("Normal".to_string()));
        assert!(normal.is_default);

        let h1 = parsed.get_by_id("Heading1").unwrap();
        assert_eq!(h1.based_on, Some("Normal".to_string()));
    }

    #[test]
    fn find_default_style() {
        let styles = CT_Styles::new_default();
        let default_para = styles.get_default(StyleType::Paragraph).unwrap();
        assert_eq!(default_para.style_id, "Normal");
    }

    #[test]
    fn aliased_style_paragraph_properties_use_ancestor_namespace_scope() {
        let xml = format!(
            r#"<q:styles xmlns:q="{W_NS}" xmlns:ext="urn:producer"><q:style q:type="paragraph" q:styleId="Alias"><q:pPr><ext:jc ext:val="right"/><q:jc q:val="center"/></q:pPr></q:style></q:styles>"#
        );
        let parsed = CT_Styles::from_xml(xml.as_bytes()).unwrap();
        let ppr = parsed.styles[0].ppr.as_ref().unwrap();
        assert_eq!(ppr.jc, Some(crate::shared::ST_Jc::Center));
    }

    #[test]
    fn direct_style_parser_uses_supplied_start_ancestor_scope() {
        let xml = format!(
            r#"<outer xmlns:ext="urn:producer"><q:style xmlns:q="{W_NS}" q:type="paragraph" q:styleId="Direct"><ext:pPr><ext:jc ext:val="right"/></ext:pPr><q:pPr><ext:jc ext:val="right"/><q:jc q:val="center"/></q:pPr></q:style></outer>"#
        );
        let mut reader = Reader::from_str(&xml);
        let mut buf = Vec::new();
        let parsed = loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"style" => {
                    break CT_Style::from_xml(&mut reader, element).unwrap();
                }
                Ok(Event::Eof) => panic!("missing style"),
                event => {
                    event.unwrap();
                }
            }
            buf.clear();
        };
        assert_eq!(
            parsed.ppr.as_ref().unwrap().jc,
            Some(crate::shared::ST_Jc::Center)
        );
    }

    #[test]
    fn direct_style_parser_does_not_promote_foreign_start_prefix() {
        let xml = r#"<outer><ext:style xmlns:ext="urn:producer" ext:type="paragraph" ext:styleId="Foreign"><ext:pPr><ext:jc ext:val="right"/></ext:pPr></ext:style></outer>"#;
        let mut reader = Reader::from_str(xml);
        let mut buf = Vec::new();
        let parsed = loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"style" => {
                    break CT_Style::from_xml(&mut reader, element).unwrap();
                }
                Ok(Event::Eof) => panic!("missing style"),
                event => {
                    event.unwrap();
                }
            }
            buf.clear();
        };
        assert!(parsed.ppr.is_none());
    }

    #[test]
    fn direct_style_parser_accepts_default_word_namespace() {
        let xml = format!(
            r#"<outer xmlns:ext="urn:producer"><style xmlns="{W_NS}" xmlns:w="{W_NS}" w:type="paragraph" w:styleId="Direct"><ext:pPr><ext:jc ext:val="right"/></ext:pPr><pPr><ext:jc ext:val="right"/><jc w:val="center"/></pPr></style></outer>"#
        );
        let mut reader = Reader::from_str(&xml);
        let mut buf = Vec::new();
        let parsed = loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"style" => {
                    break CT_Style::from_xml(&mut reader, element).unwrap();
                }
                Ok(Event::Eof) => panic!("missing style"),
                event => {
                    event.unwrap();
                }
            }
            buf.clear();
        };
        assert_eq!(
            parsed.ppr.as_ref().unwrap().jc,
            Some(crate::shared::ST_Jc::Center)
        );
    }
}
