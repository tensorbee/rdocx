//! Text content elements: `CT_P` (paragraph), `CT_R` (run), `CT_Text`.

use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};

use crate::drawing::CT_Drawing;
use crate::error::Result;
use crate::namespace::matches_local_name;
use crate::properties::{CT_PPr, CT_RPr};
use crate::raw_xml::{capture_element, capture_empty_element};

/// `CT_Text` — The text content of a run, with optional xml:space="preserve".
#[derive(Debug, Clone, PartialEq)]
pub struct CT_Text {
    pub text: String,
    pub preserve_space: bool,
}

impl CT_Text {
    pub fn new(text: &str) -> Self {
        CT_Text {
            text: text.to_string(),
            preserve_space: text.starts_with(' ') || text.ends_with(' '),
        }
    }
}

/// Types of simple fields.
#[derive(Debug, Clone, PartialEq)]
pub enum FieldType {
    /// Current page number (PAGE field).
    Page,
    /// Total number of pages (NUMPAGES field).
    NumPages,
    /// Any other field instruction.
    Other(String),
}

/// Content that can appear inside a run.
#[derive(Debug, Clone, PartialEq)]
pub enum RunContent {
    Text(CT_Text),
    Tab,
    Break(BreakType),
    Drawing(CT_Drawing),
    /// A simple field (from `<w:fldSimple>`).
    Field { field_type: FieldType },
    /// A footnote reference (`<w:footnoteReference w:id="..."/>`).
    FootnoteRef { id: i32 },
    /// An endnote reference (`<w:endnoteReference w:id="..."/>`).
    EndnoteRef { id: i32 },
}

/// Types of breaks.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BreakType {
    Line,
    Page,
    Column,
}

/// `CT_R` — A run of text with uniform formatting.
#[derive(Debug, Clone, PartialEq)]
#[allow(non_snake_case)]
pub struct CT_R {
    pub properties: Option<CT_RPr>,
    pub content: Vec<RunContent>,
    /// Unknown child elements captured as raw XML.
    pub extra_xml: Vec<Vec<u8>>,
}

#[allow(non_snake_case)]
impl CT_R {
    pub fn new(text: &str) -> Self {
        CT_R {
            properties: None,
            content: vec![RunContent::Text(CT_Text::new(text))],
            extra_xml: Vec::new(),
        }
    }

    /// Get the combined text of all text content in this run.
    pub fn text(&self) -> String {
        let mut result = String::new();
        for item in &self.content {
            match item {
                RunContent::Text(t) => result.push_str(&t.text),
                RunContent::Tab => result.push('\t'),
                RunContent::Break(_) => result.push('\n'),
                RunContent::Drawing(_) => {} // Drawings have no text content
                RunContent::Field { .. } => {} // Fields have no static text
                RunContent::FootnoteRef { .. } | RunContent::EndnoteRef { .. } => {}
            }
        }
        result
    }

    pub fn from_xml(reader: &mut Reader<&[u8]>) -> Result<Self> {
        let mut properties = None;
        let mut content = Vec::new();
        let mut extra_xml = Vec::new();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let name = e.name();
                    if matches_local_name(name.as_ref(), b"rPr") {
                        properties = Some(CT_RPr::from_xml(reader)?);
                    } else if matches_local_name(name.as_ref(), b"t") {
                        let preserve = e.attributes().any(|a| {
                            a.ok()
                                .map(|a| a.key.as_ref() == b"xml:space" && a.value.as_ref() == b"preserve")
                                .unwrap_or(false)
                        });
                        let text = reader
                            .read_text(name)
                            .map(|t| {
                                // Unescape XML entities (&amp; &lt; &gt; &quot; &apos;)
                                quick_xml::escape::unescape(&t)
                                    .map(|u| u.to_string())
                                    .unwrap_or_else(|_| t.to_string())
                            })
                            .unwrap_or_default();
                        content.push(RunContent::Text(CT_Text {
                            text,
                            preserve_space: preserve,
                        }));
                    } else if matches_local_name(name.as_ref(), b"drawing") {
                        content.push(RunContent::Drawing(CT_Drawing::from_xml(reader)?));
                    } else {
                        // Capture unknown child elements as raw XML
                        extra_xml.push(capture_element(reader, e)?);
                    }
                }
                Ok(Event::Empty(ref e)) => {
                    let name = e.name();
                    if matches_local_name(name.as_ref(), b"tab") {
                        content.push(RunContent::Tab);
                    } else if matches_local_name(name.as_ref(), b"br") {
                        let break_type = e
                            .attributes()
                            .filter_map(|a| a.ok())
                            .find(|a| matches_local_name(a.key.as_ref(), b"type"))
                            .map(|a| match a.value.as_ref() {
                                b"page" => BreakType::Page,
                                b"column" => BreakType::Column,
                                _ => BreakType::Line,
                            })
                            .unwrap_or(BreakType::Line);
                        content.push(RunContent::Break(break_type));
                    } else if matches_local_name(name.as_ref(), b"footnoteReference") {
                        let id = e
                            .attributes()
                            .filter_map(|a| a.ok())
                            .find(|a| matches_local_name(a.key.as_ref(), b"id"))
                            .and_then(|a| std::str::from_utf8(&a.value).ok()?.parse::<i32>().ok())
                            .unwrap_or(0);
                        content.push(RunContent::FootnoteRef { id });
                    } else if matches_local_name(name.as_ref(), b"endnoteReference") {
                        let id = e
                            .attributes()
                            .filter_map(|a| a.ok())
                            .find(|a| matches_local_name(a.key.as_ref(), b"id"))
                            .and_then(|a| std::str::from_utf8(&a.value).ok()?.parse::<i32>().ok())
                            .unwrap_or(0);
                        content.push(RunContent::EndnoteRef { id });
                    }
                }
                Ok(Event::End(ref e)) if matches_local_name(e.name().as_ref(), b"r") => {
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(CT_R {
            properties,
            content,
            extra_xml,
        })
    }

    pub fn to_xml<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_event(Event::Start(BytesStart::new("w:r")))?;

        if let Some(ref props) = self.properties {
            props.to_xml(writer)?;
        }

        for item in &self.content {
            match item {
                RunContent::Text(t) => {
                    let mut e = BytesStart::new("w:t");
                    if t.preserve_space {
                        e.push_attribute(("xml:space", "preserve"));
                    }
                    writer.write_event(Event::Start(e))?;
                    writer
                        .write_event(Event::Text(BytesText::new(&t.text)))?;
                    writer.write_event(Event::End(BytesEnd::new("w:t")))?;
                }
                RunContent::Tab => {
                    writer.write_event(Event::Empty(BytesStart::new("w:tab")))?;
                }
                RunContent::Break(bt) => {
                    let mut e = BytesStart::new("w:br");
                    match bt {
                        BreakType::Page => e.push_attribute(("w:type", "page")),
                        BreakType::Column => e.push_attribute(("w:type", "column")),
                        BreakType::Line => {}
                    }
                    writer.write_event(Event::Empty(e))?;
                }
                RunContent::Drawing(d) => {
                    d.to_xml(writer)?;
                }
                RunContent::Field { .. } => {
                    // Field runs are serialized at the paragraph level as <w:fldSimple>
                }
                RunContent::FootnoteRef { id } => {
                    let mut buf = itoa::Buffer::new();
                    let mut e = BytesStart::new("w:footnoteReference");
                    e.push_attribute(("w:id", buf.format(*id)));
                    writer.write_event(Event::Empty(e))?;
                }
                RunContent::EndnoteRef { id } => {
                    let mut buf = itoa::Buffer::new();
                    let mut e = BytesStart::new("w:endnoteReference");
                    e.push_attribute(("w:id", buf.format(*id)));
                    writer.write_event(Event::Empty(e))?;
                }
            }
        }

        // Write captured unknown child elements
        for raw in &self.extra_xml {
            writer.get_mut().write_all(raw)?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:r")))?;
        Ok(())
    }
}

/// A hyperlink span that wraps a range of runs.
#[derive(Debug, Clone, PartialEq)]
pub struct HyperlinkSpan {
    /// The relationship ID for the hyperlink target.
    pub rel_id: Option<String>,
    /// Optional anchor within the document (for internal links).
    pub anchor: Option<String>,
    /// Index of the first run in the hyperlink (inclusive).
    pub run_start: usize,
    /// Index of the last run in the hyperlink (exclusive).
    pub run_end: usize,
}

/// `CT_P` — A paragraph element containing runs and properties.
#[derive(Debug, Clone, PartialEq)]
#[allow(non_snake_case)]
pub struct CT_P {
    pub properties: Option<CT_PPr>,
    pub runs: Vec<CT_R>,
    /// Hyperlink spans referencing ranges of runs.
    pub hyperlinks: Vec<HyperlinkSpan>,
    /// Unknown child elements captured as raw XML with their insertion position (run index).
    pub extra_xml: Vec<(usize, Vec<u8>)>,
}

#[allow(non_snake_case)]
impl CT_P {
    pub fn new() -> Self {
        CT_P {
            properties: None,
            runs: Vec::new(),
            hyperlinks: Vec::new(),
            extra_xml: Vec::new(),
        }
    }

    /// Get the combined text of all runs in this paragraph.
    pub fn text(&self) -> String {
        self.runs.iter().map(|r| r.text()).collect()
    }

    /// Add a run with the given text.
    pub fn add_run(&mut self, text: &str) -> &mut CT_R {
        self.runs.push(CT_R::new(text));
        self.runs.last_mut().unwrap()
    }

    pub fn from_xml(reader: &mut Reader<&[u8]>) -> Result<Self> {
        let mut properties = None;
        let mut runs = Vec::new();
        let mut hyperlinks = Vec::new();
        let mut extra_xml = Vec::new();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let name = e.name();
                    if matches_local_name(name.as_ref(), b"pPr") {
                        properties = Some(CT_PPr::from_xml(reader)?);
                    } else if matches_local_name(name.as_ref(), b"r") {
                        runs.push(CT_R::from_xml(reader)?);
                    } else if matches_local_name(name.as_ref(), b"hyperlink") {
                        // Parse hyperlink: extract r:id and/or w:anchor, then parse child runs
                        let mut rel_id = None;
                        let mut anchor = None;
                        for attr in e.attributes().flatten() {
                            let key = attr.key.as_ref();
                            if matches_local_name(key, b"id") {
                                rel_id = Some(
                                    std::str::from_utf8(&attr.value)
                                        .unwrap_or("")
                                        .to_string(),
                                );
                            } else if matches_local_name(key, b"anchor") {
                                anchor = Some(
                                    std::str::from_utf8(&attr.value)
                                        .unwrap_or("")
                                        .to_string(),
                                );
                            }
                        }

                        let run_start = runs.len();
                        // Parse child runs within the hyperlink
                        let mut inner_buf = Vec::new();
                        loop {
                            match reader.read_event_into(&mut inner_buf) {
                                Ok(Event::Start(ref ie)) => {
                                    let iname = ie.name();
                                    if matches_local_name(iname.as_ref(), b"r") {
                                        runs.push(CT_R::from_xml(reader)?);
                                    } else {
                                        reader.read_to_end_into(iname, &mut Vec::new())?;
                                    }
                                }
                                Ok(Event::End(ref ie))
                                    if matches_local_name(ie.name().as_ref(), b"hyperlink") =>
                                {
                                    break;
                                }
                                Ok(Event::Eof) => break,
                                Err(e) => return Err(e.into()),
                                _ => {}
                            }
                            inner_buf.clear();
                        }

                        let run_end = runs.len();
                        if run_start < run_end && (rel_id.is_some() || anchor.is_some()) {
                            hyperlinks.push(HyperlinkSpan {
                                rel_id,
                                anchor,
                                run_start,
                                run_end,
                            });
                        }
                    } else if matches_local_name(name.as_ref(), b"fldSimple") {
                        // Parse simple field: extract w:instr attribute
                        let mut instr = String::new();
                        for attr in e.attributes().flatten() {
                            if matches_local_name(attr.key.as_ref(), b"instr") {
                                instr = std::str::from_utf8(&attr.value)
                                    .unwrap_or("")
                                    .to_string();
                            }
                        }

                        let field_type = parse_field_instruction(&instr);

                        // Skip child runs (they contain the default display value)
                        reader.read_to_end_into(name, &mut Vec::new())?;

                        // Add a synthetic run with the field content
                        runs.push(CT_R {
                            properties: None,
                            content: vec![RunContent::Field { field_type }],
                            extra_xml: Vec::new(),
                        });
                    } else {
                        // Capture unknown elements (bookmarks, comments, etc.) as raw XML
                        extra_xml.push((runs.len(), capture_element(reader, e)?));
                    }
                }
                Ok(Event::Empty(ref e)) => {
                    let name = e.name();
                    if !matches_local_name(name.as_ref(), b"p") {
                        extra_xml.push((runs.len(), capture_empty_element(e)?));
                    }
                }
                Ok(Event::End(ref e)) if matches_local_name(e.name().as_ref(), b"p") => {
                    break;
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        Ok(CT_P {
            properties,
            runs,
            hyperlinks,
            extra_xml,
        })
    }

    pub fn to_xml<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_event(Event::Start(BytesStart::new("w:p")))?;

        if let Some(ref props) = self.properties {
            props.to_xml(writer)?;
        }

        // Build a set of run indices that are inside hyperlinks
        let mut hyperlink_runs: std::collections::HashMap<usize, usize> = std::collections::HashMap::new();
        for (hl_idx, hl) in self.hyperlinks.iter().enumerate() {
            for run_idx in hl.run_start..hl.run_end {
                hyperlink_runs.insert(run_idx, hl_idx);
            }
        }

        // Build index of extra_xml elements by position for interleaving
        let mut extras_by_pos: std::collections::HashMap<usize, Vec<&Vec<u8>>> = std::collections::HashMap::new();
        for (pos, raw) in &self.extra_xml {
            extras_by_pos.entry(*pos).or_default().push(raw);
        }

        let mut current_hyperlink: Option<usize> = None;
        for (run_idx, run) in self.runs.iter().enumerate() {
            // Write any extras that should appear before this run
            if let Some(extras) = extras_by_pos.get(&run_idx) {
                for raw in extras {
                    writer.get_mut().write_all(raw)?;
                }
            }
            let in_hl = hyperlink_runs.get(&run_idx).copied();

            // Close hyperlink if we left it
            if current_hyperlink.is_some() && current_hyperlink != in_hl {
                writer.write_event(Event::End(BytesEnd::new("w:hyperlink")))?;
                current_hyperlink = None;
            }

            // Open hyperlink if entering one
            if let Some(hl_idx) = in_hl {
                if current_hyperlink != in_hl {
                    let hl = &self.hyperlinks[hl_idx];
                    let mut e = BytesStart::new("w:hyperlink");
                    if let Some(ref rid) = hl.rel_id {
                        e.push_attribute(("r:id", rid.as_str()));
                    }
                    if let Some(ref anchor) = hl.anchor {
                        e.push_attribute(("w:anchor", anchor.as_str()));
                    }
                    writer.write_event(Event::Start(e))?;
                    current_hyperlink = in_hl;
                }
            }

            // Check if this run is a field run
            if run.content.len() == 1 {
                if let RunContent::Field { field_type } = &run.content[0] {
                    let instr = match field_type {
                        FieldType::Page => " PAGE ",
                        FieldType::NumPages => " NUMPAGES ",
                        FieldType::Other(s) => s.as_str(),
                    };
                    let mut fld = BytesStart::new("w:fldSimple");
                    fld.push_attribute(("w:instr", instr));
                    writer.write_event(Event::Start(fld))?;
                    // Emit a default display run
                    writer.write_event(Event::Start(BytesStart::new("w:r")))?;
                    writer.write_event(Event::Start(BytesStart::new("w:t")))?;
                    writer.write_event(Event::Text(BytesText::new("1")))?;
                    writer.write_event(Event::End(BytesEnd::new("w:t")))?;
                    writer.write_event(Event::End(BytesEnd::new("w:r")))?;
                    writer.write_event(Event::End(BytesEnd::new("w:fldSimple")))?;
                    continue;
                }
            }

            run.to_xml(writer)?;
        }

        // Close any remaining open hyperlink
        if current_hyperlink.is_some() {
            writer.write_event(Event::End(BytesEnd::new("w:hyperlink")))?;
        }

        // Write any extras that come after the last run
        if let Some(extras) = extras_by_pos.get(&self.runs.len()) {
            for raw in extras {
                writer.get_mut().write_all(raw)?;
            }
        }

        writer.write_event(Event::End(BytesEnd::new("w:p")))?;
        Ok(())
    }
}

/// Parse a field instruction string into a FieldType.
fn parse_field_instruction(instr: &str) -> FieldType {
    let trimmed = instr.trim().to_uppercase();
    // Field instruction may have switches like "PAGE \* MERGEFORMAT"
    let keyword = trimmed.split_whitespace().next().unwrap_or("");
    match keyword {
        "PAGE" => FieldType::Page,
        "NUMPAGES" => FieldType::NumPages,
        _ => FieldType::Other(instr.trim().to_string()),
    }
}

impl Default for CT_P {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn parse_paragraph(xml: &str) -> CT_P {
        let full = format!("<w:p>{xml}</w:p>");
        let mut reader = Reader::from_str(&full);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e))
                    if matches_local_name(e.name().as_ref(), b"p") =>
                {
                    break
                }
                _ => {}
            }
            buf.clear();
        }
        CT_P::from_xml(&mut reader).unwrap()
    }

    #[test]
    fn parse_simple_paragraph() {
        let p = parse_paragraph(r#"<w:r><w:t>Hello World</w:t></w:r>"#);
        assert_eq!(p.text(), "Hello World");
        assert_eq!(p.runs.len(), 1);
    }

    #[test]
    fn parse_paragraph_with_properties() {
        let p = parse_paragraph(
            r#"<w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>Centered</w:t></w:r>"#,
        );
        assert_eq!(p.text(), "Centered");
        assert!(p.properties.is_some());
        assert_eq!(
            p.properties.as_ref().unwrap().jc,
            Some(crate::shared::ST_Jc::Center)
        );
    }

    #[test]
    fn parse_run_with_formatting() {
        let p = parse_paragraph(
            r#"<w:r><w:rPr><w:b/><w:i/></w:rPr><w:t>Bold Italic</w:t></w:r>"#,
        );
        let run = &p.runs[0];
        let rpr = run.properties.as_ref().unwrap();
        assert_eq!(rpr.bold, Some(true));
        assert_eq!(rpr.italic, Some(true));
    }

    #[test]
    fn parse_multiple_runs() {
        let p = parse_paragraph(
            r#"<w:r><w:t>Hello </w:t></w:r><w:r><w:t>World</w:t></w:r>"#,
        );
        assert_eq!(p.runs.len(), 2);
        assert_eq!(p.text(), "Hello World");
    }

    #[test]
    fn parse_hyperlink() {
        let p = parse_paragraph(
            r#"<w:hyperlink r:id="rId5"><w:r><w:t>Click here</w:t></w:r></w:hyperlink>"#,
        );
        assert_eq!(p.runs.len(), 1);
        assert_eq!(p.text(), "Click here");
        assert_eq!(p.hyperlinks.len(), 1);
        assert_eq!(p.hyperlinks[0].rel_id, Some("rId5".to_string()));
        assert_eq!(p.hyperlinks[0].run_start, 0);
        assert_eq!(p.hyperlinks[0].run_end, 1);
    }

    #[test]
    fn parse_hyperlink_with_anchor() {
        let p = parse_paragraph(
            r#"<w:hyperlink w:anchor="section1"><w:r><w:t>Go to section</w:t></w:r></w:hyperlink>"#,
        );
        assert_eq!(p.hyperlinks.len(), 1);
        assert_eq!(p.hyperlinks[0].anchor, Some("section1".to_string()));
        assert!(p.hyperlinks[0].rel_id.is_none());
    }

    #[test]
    fn parse_hyperlink_multiple_runs() {
        let p = parse_paragraph(
            r#"<w:r><w:t>Before </w:t></w:r><w:hyperlink r:id="rId6"><w:r><w:t>link </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>text</w:t></w:r></w:hyperlink><w:r><w:t> after</w:t></w:r>"#,
        );
        assert_eq!(p.runs.len(), 4);
        assert_eq!(p.text(), "Before link text after");
        assert_eq!(p.hyperlinks.len(), 1);
        assert_eq!(p.hyperlinks[0].run_start, 1);
        assert_eq!(p.hyperlinks[0].run_end, 3);
    }

    #[test]
    fn round_trip_hyperlink() {
        let mut p = CT_P::new();
        p.add_run("Before ");
        p.add_run("link text");
        p.add_run(" after");
        p.hyperlinks.push(HyperlinkSpan {
            rel_id: Some("rId7".to_string()),
            anchor: None,
            run_start: 1,
            run_end: 2,
        });

        let mut output = Vec::new();
        let mut writer = Writer::new(&mut output);
        p.to_xml(&mut writer).unwrap();
        let xml = String::from_utf8(output).unwrap();

        let parsed = parse_paragraph(
            xml.strip_prefix("<w:p>")
                .unwrap()
                .strip_suffix("</w:p>")
                .unwrap(),
        );
        assert_eq!(parsed.text(), "Before link text after");
        assert_eq!(parsed.hyperlinks.len(), 1);
        assert_eq!(parsed.hyperlinks[0].rel_id, Some("rId7".to_string()));
        assert_eq!(parsed.hyperlinks[0].run_start, 1);
        assert_eq!(parsed.hyperlinks[0].run_end, 2);
    }

    #[test]
    fn parse_fld_simple_page() {
        let p = parse_paragraph(
            r#"<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>"#,
        );
        assert_eq!(p.runs.len(), 1);
        assert_eq!(p.runs[0].content.len(), 1);
        assert!(matches!(
            p.runs[0].content[0],
            RunContent::Field { field_type: FieldType::Page }
        ));
    }

    #[test]
    fn parse_fld_simple_numpages() {
        let p = parse_paragraph(
            r#"<w:fldSimple w:instr=" NUMPAGES \* MERGEFORMAT "><w:r><w:t>5</w:t></w:r></w:fldSimple>"#,
        );
        assert_eq!(p.runs.len(), 1);
        assert!(matches!(
            p.runs[0].content[0],
            RunContent::Field { field_type: FieldType::NumPages }
        ));
    }

    #[test]
    fn parse_fld_simple_mixed_with_text() {
        let p = parse_paragraph(
            r#"<w:r><w:t>Page </w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple><w:r><w:t> of </w:t></w:r><w:fldSimple w:instr=" NUMPAGES "><w:r><w:t>5</w:t></w:r></w:fldSimple>"#,
        );
        assert_eq!(p.runs.len(), 4);
        assert_eq!(p.text(), "Page  of ");
        assert!(matches!(
            p.runs[1].content[0],
            RunContent::Field { field_type: FieldType::Page }
        ));
        assert!(matches!(
            p.runs[3].content[0],
            RunContent::Field { field_type: FieldType::NumPages }
        ));
    }

    #[test]
    fn round_trip_fld_simple() {
        let mut p = CT_P::new();
        p.add_run("Page ");
        p.runs.push(CT_R {
            properties: None,
            content: vec![RunContent::Field { field_type: FieldType::Page }],
            extra_xml: Vec::new(),
        });

        let mut output = Vec::new();
        let mut writer = Writer::new(&mut output);
        p.to_xml(&mut writer).unwrap();
        let xml = String::from_utf8(output).unwrap();

        let parsed = parse_paragraph(
            xml.strip_prefix("<w:p>")
                .unwrap()
                .strip_suffix("</w:p>")
                .unwrap(),
        );
        assert_eq!(parsed.runs.len(), 2);
        assert!(matches!(
            parsed.runs[1].content[0],
            RunContent::Field { field_type: FieldType::Page }
        ));
    }

    #[test]
    fn round_trip_paragraph() {
        let mut p = CT_P::new();
        p.add_run("Hello ");
        let run = p.add_run("World");
        run.properties = Some(CT_RPr {
            bold: Some(true),
            ..Default::default()
        });

        let mut output = Vec::new();
        let mut writer = Writer::new(&mut output);
        p.to_xml(&mut writer).unwrap();
        let xml = String::from_utf8(output).unwrap();

        let parsed = parse_paragraph(
            xml.strip_prefix("<w:p>")
                .unwrap()
                .strip_suffix("</w:p>")
                .unwrap(),
        );
        assert_eq!(parsed.text(), "Hello World");
        assert_eq!(parsed.runs.len(), 2);
        assert_eq!(parsed.runs[1].properties.as_ref().unwrap().bold, Some(true));
    }

    #[test]
    fn parse_footnote_reference() {
        let p = parse_paragraph(
            r#"<w:r><w:t>Some text</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r>"#,
        );
        assert_eq!(p.runs.len(), 2);
        assert_eq!(p.runs[0].text(), "Some text");
        assert_eq!(p.runs[1].content.len(), 1);
        assert!(matches!(
            p.runs[1].content[0],
            RunContent::FootnoteRef { id: 1 }
        ));
    }

    #[test]
    fn parse_endnote_reference() {
        let p = parse_paragraph(
            r#"<w:r><w:endnoteReference w:id="3"/></w:r>"#,
        );
        assert_eq!(p.runs.len(), 1);
        assert!(matches!(
            p.runs[0].content[0],
            RunContent::EndnoteRef { id: 3 }
        ));
    }

    #[test]
    fn round_trip_footnote_reference() {
        let mut p = CT_P::new();
        p.add_run("Text before");
        p.runs.push(CT_R {
            properties: None,
            content: vec![RunContent::FootnoteRef { id: 2 }],
            extra_xml: Vec::new(),
        });
        p.add_run(" text after");

        let mut output = Vec::new();
        let mut writer = Writer::new(&mut output);
        p.to_xml(&mut writer).unwrap();
        let xml = String::from_utf8(output).unwrap();

        let parsed = parse_paragraph(
            xml.strip_prefix("<w:p>")
                .unwrap()
                .strip_suffix("</w:p>")
                .unwrap(),
        );
        assert_eq!(parsed.runs.len(), 3);
        assert!(matches!(
            parsed.runs[1].content[0],
            RunContent::FootnoteRef { id: 2 }
        ));
    }
}
