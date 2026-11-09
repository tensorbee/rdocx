#![doc = include_str!("../README.md")]

mod css;
mod emitter;
mod markdown;
mod sanitize;

use std::borrow::Cow;
use std::collections::HashMap;

use rdocx_oxml::content_control::{CT_Sdt, SdtContent};
use rdocx_oxml::document::{BodyContent, CT_Body, CT_Document};
use rdocx_oxml::numbering::CT_Numbering;
use rdocx_oxml::styles::CT_Styles;
use rdocx_oxml::table::{CT_Tbl, CT_Tc, CellContent};
use rdocx_oxml::text::CT_P;

/// Options for HTML conversion.
#[derive(Debug, Clone)]
pub struct HtmlOptions {
    /// Whether to inline images as base64 data URIs (default: true).
    pub inline_images: bool,
}

impl Default for HtmlOptions {
    fn default() -> Self {
        Self {
            inline_images: true,
        }
    }
}

/// Input for HTML conversion.
pub struct HtmlInput {
    pub document: CT_Document,
    pub styles: CT_Styles,
    pub numbering: Option<CT_Numbering>,
    /// Images keyed by embed/relationship ID.
    pub images: HashMap<String, ImageData>,
    /// Hyperlink URLs keyed by relationship ID.
    pub hyperlink_urls: HashMap<String, String>,
}

/// Image data for HTML embedding.
pub struct ImageData {
    pub data: Vec<u8>,
    pub content_type: String,
}

/// A semantic chart projection from the current chart part.
pub struct ChartTable {
    pub title: Option<String>,
    pub categories: Vec<String>,
    pub series: Vec<(String, Vec<f64>)>,
}

impl ChartTable {
    /// Header followed by category rows, with absent or non-finite values blank.
    pub fn rows(&self) -> Vec<Vec<String>> {
        let mut header = vec![String::new()];
        header.extend(self.series.iter().map(|(name, _)| name.clone()));
        let count = self
            .series
            .iter()
            .map(|(_, values)| values.len())
            .chain(std::iter::once(self.categories.len()))
            .max()
            .unwrap_or(0);
        let mut rows = vec![header];
        for index in 0..count {
            let mut row = vec![self.categories.get(index).cloned().unwrap_or_default()];
            row.extend(self.series.iter().map(|(_, values)| {
                values
                    .get(index)
                    .filter(|value| value.is_finite())
                    .map(ToString::to_string)
                    .unwrap_or_default()
            }));
            rows.push(row);
        }
        rows
    }
}

/// Convert a DOCX document to a complete HTML document string.
pub fn to_html_document(input: &HtmlInput, options: &HtmlOptions) -> String {
    to_html_document_with_charts(input, options, &HashMap::new())
}

/// Convert a DOCX document with current chart tables to complete HTML.
pub fn to_html_document_with_charts(
    input: &HtmlInput,
    options: &HtmlOptions,
    charts: &HashMap<String, ChartTable>,
) -> String {
    let body = to_html_fragment_with_charts(input, options, charts);
    let css = css::generate_base_css();
    format!(
        "<!DOCTYPE html>\n<html>\n<head>\n<meta charset=\"UTF-8\">\n<style>\n{css}\n</style>\n</head>\n<body>\n{body}\n</body>\n</html>"
    )
}

/// Convert a DOCX document to an HTML fragment (body content only).
pub fn to_html_fragment(input: &HtmlInput, options: &HtmlOptions) -> String {
    to_html_fragment_with_charts(input, options, &HashMap::new())
}

/// Convert body content with current chart tables to an HTML fragment.
pub fn to_html_fragment_with_charts(
    input: &HtmlInput,
    options: &HtmlOptions,
    charts: &HashMap<String, ChartTable>,
) -> String {
    emitter::emit_body(
        &input.document.body,
        &input.styles,
        input.numbering.as_ref(),
        &input.images,
        &input.hyperlink_urls,
        options,
        charts,
    )
}

/// Convert a DOCX document to Markdown.
pub fn to_markdown(input: &HtmlInput) -> String {
    to_markdown_with_charts(input, &HashMap::new())
}

/// Convert body content with current chart tables to Markdown.
pub fn to_markdown_with_charts(input: &HtmlInput, charts: &HashMap<String, ChartTable>) -> String {
    markdown::emit_markdown(input, charts)
}

fn paragraph_chart_ids(paragraph: &CT_P) -> Vec<String> {
    let paragraph = paragraph.accepted_view();
    paragraph
        .runs
        .iter()
        .flat_map(|run| &run.content)
        .filter_map(|content| {
            if let rdocx_oxml::text::RunContent::Drawing(drawing) = content {
                Some(drawing)
            } else {
                None
            }
        })
        .flat_map(|drawing| {
            drawing
                .inline
                .as_ref()
                .and_then(|inline| inline.chart_rel_id.as_ref())
                .into_iter()
                .chain(
                    drawing
                        .anchor
                        .as_ref()
                        .and_then(|anchor| anchor.chart_rel_id.as_ref()),
                )
        })
        .cloned()
        .collect()
}

/// A paragraph or a table that the emitters write as one block.
enum Block<'a> {
    Paragraph(Box<Cow<'a, CT_P>>, Vec<usize>),
    Table(&'a CT_Tbl),
}

/// The paragraphs and tables of a body in document order. A block content
/// control is transparent: what it wraps, nested controls included, is
/// written as if the control were not there, which is what `Paragraph.text`
/// and `Document::text` read.
fn body_blocks(body: &CT_Body, preserve_field_owners: bool) -> Vec<Block<'_>> {
    let mut blocks = Vec::new();
    let mut carried: Option<(CT_P, Vec<usize>)> = None;
    for (index, item) in body.content.iter().enumerate() {
        match item {
            BodyContent::Paragraph(paragraph) => {
                let (paragraph, boundaries) = match carried.take() {
                    Some(prefix) => {
                        let (joined, boundaries) =
                            join_accepted_paragraphs(prefix, paragraph, preserve_field_owners);
                        (Cow::Owned(joined), boundaries)
                    }
                    None => (Cow::Borrowed(paragraph.as_ref()), Vec::new()),
                };
                if body.accepted_paragraph_joins_next(index) {
                    carried = Some(if preserve_field_owners {
                        let (view, mut projected) = paragraph.accepted_view_with_field_owners();
                        projected.extend(boundaries);
                        (view.into_owned(), projected)
                    } else {
                        (paragraph.accepted_view().into_owned(), Vec::new())
                    });
                } else {
                    blocks.push(Block::Paragraph(Box::new(paragraph), boundaries));
                }
            }
            BodyContent::Table(table) => blocks.push(Block::Table(table.as_ref())),
            BodyContent::ContentControl(control) => push_control_blocks(control, &mut blocks),
            BodyContent::RawXml(_) => {}
        }
    }
    blocks
}

fn join_accepted_paragraphs(
    prefix: (CT_P, Vec<usize>),
    paragraph: &CT_P,
    preserve_field_owners: bool,
) -> (CT_P, Vec<usize>) {
    let (prefix, mut boundaries) = prefix;
    let (mut result, projected) = if preserve_field_owners {
        let (view, boundaries) = paragraph.accepted_view_with_field_owners();
        (view.into_owned(), boundaries)
    } else {
        (paragraph.accepted_view().into_owned(), Vec::new())
    };
    let offset = prefix.runs.len();
    boundaries.extend(projected.into_iter().map(|at| offset + at));
    for hyperlink in &mut result.hyperlinks {
        hyperlink.run_start += offset;
        hyperlink.run_end += offset;
    }
    let mut hyperlinks = prefix.hyperlinks;
    hyperlinks.extend(result.hyperlinks);
    result.hyperlinks = hyperlinks;
    let mut runs = prefix.runs;
    runs.extend(result.runs);
    result.runs = runs;
    (result, boundaries)
}

/// The paragraphs and nested tables of a table cell, as [`body_blocks`]
/// reads a body.
fn cell_blocks(cell: &CT_Tc) -> Vec<Block<'_>> {
    let mut blocks = Vec::new();
    for item in &cell.content {
        match item {
            CellContent::Paragraph(paragraph) => blocks.push(Block::Paragraph(
                Box::new(Cow::Borrowed(paragraph)),
                Vec::new(),
            )),
            CellContent::Table(table) => blocks.push(Block::Table(table)),
            CellContent::ContentControl(control) => push_control_blocks(control, &mut blocks),
        }
    }
    blocks
}

fn push_control_blocks<'a>(control: &'a CT_Sdt, blocks: &mut Vec<Block<'a>>) {
    for item in &control.content {
        match item {
            SdtContent::Paragraph(paragraph) => blocks.push(Block::Paragraph(
                Box::new(Cow::Borrowed(paragraph)),
                Vec::new(),
            )),
            SdtContent::Table(table) => blocks.push(Block::Table(table)),
            SdtContent::ContentControl(nested) => push_control_blocks(nested, blocks),
            // A control around a paragraph or a table holds no rows, cells or
            // runs of its own. Those around rows, cells and runs are read
            // through `CT_Tbl::rows`, `CT_Row::cells` and `CT_P::accepted_view`.
            SdtContent::Row(_) | SdtContent::Cell(_) | SdtContent::Run(_) => {}
            SdtContent::RawXml(_) => {}
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use rdocx_oxml::document::{BodyContent, CT_Document};
    use rdocx_oxml::styles::CT_Styles;
    use rdocx_oxml::text::CT_P;

    fn simple_input(text: &str) -> HtmlInput {
        let mut doc = CT_Document::new();
        let mut p = CT_P::new();
        p.add_run(text);
        doc.body.add_paragraph(p);

        HtmlInput {
            document: doc,
            styles: CT_Styles::new_default(),
            numbering: None,
            images: HashMap::new(),
            hyperlink_urls: HashMap::new(),
        }
    }

    #[test]
    fn html_document_basic() {
        let input = simple_input("Hello, World!");
        let html = to_html_document(&input, &HtmlOptions::default());
        assert!(html.contains("<!DOCTYPE html>"));
        assert!(html.contains("Hello, World!"));
        assert!(html.contains("<p"));
    }

    #[test]
    fn html_fragment_basic() {
        let input = simple_input("Test paragraph");
        let html = to_html_fragment(&input, &HtmlOptions::default());
        assert!(html.contains("Test paragraph"));
        assert!(html.contains("<p"));
        assert!(!html.contains("<!DOCTYPE"));
    }

    #[test]
    fn markdown_basic() {
        let input = simple_input("Test paragraph");
        let md = to_markdown(&input);
        assert!(md.contains("Test paragraph"));
    }

    #[test]
    fn complex_field_cached_display_reaches_html_and_markdown() {
        let document = CT_Document::from_xml(
            br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>DATE</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>one</w:t><w:tab/><w:t>two</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:br/><w:t>three</w:t><w:br w:type="page"/><w:t>four</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:body></w:document>"#,
        )
        .unwrap();
        let input = HtmlInput {
            document,
            styles: CT_Styles::new_default(),
            numbering: None,
            images: HashMap::new(),
            hyperlink_urls: HashMap::new(),
        };

        let html = to_html_fragment(&input, &HtmlOptions::default());
        assert!(
            html.contains("<strong>one&emsp;two</strong><em><br>three<hr>four</em>"),
            "{html}"
        );
        let markdown = to_markdown(&input);
        assert_eq!(markdown, "**one\ttwo**  \n*three*\n\n---\n\n*four*\n\n");
    }

    #[test]
    fn controls_revisions_and_wrappers_are_transparent_to_lists_links_and_merged_rows() {
        let input = |wrap: bool| {
            let control = |content: &str| {
                if wrap {
                    format!(
                        "<w:sdt><w:sdtPr><w:tag w:val=\"c\"/></w:sdtPr><w:sdtContent>{content}</w:sdtContent></w:sdt>"
                    )
                } else {
                    content.to_owned()
                }
            };
            let tracked = |kind: &str, content: &str| {
                if wrap {
                    format!(r#"<w:{kind} w:id="1" w:author="a">{content}</w:{kind}>"#)
                } else if kind == "ins" {
                    content.to_owned()
                } else {
                    String::new()
                }
            };
            let smart_tag = |content: &str| {
                if wrap {
                    format!(r#"<w:smartTag w:uri="u" w:element="e">{content}</w:smartTag>"#)
                } else {
                    content.to_owned()
                }
            };
            let item = |text: &str| {
                format!(
                    r#"<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>"#
                )
            };
            let body = [
                item("one"),
                control(&item("two")),
                item("three"),
                format!(
                    r#"<w:p><w:hyperlink r:id="rId1"><w:r><w:t>link</w:t></w:r>{}</w:hyperlink>{}{}{}</w:p>"#,
                    tracked("ins", r#"<w:r><w:t xml:space="preserve"> more</w:t></w:r>"#),
                    control(r#"<w:r><w:t xml:space="preserve"> wrapped</w:t></w:r>"#),
                    tracked("del", r#"<w:r><w:delText> gone</w:delText></w:r>"#),
                    smart_tag(r#"<w:r><w:t xml:space="preserve"> tagged</w:t></w:r>"#)
                ),
                format!(
                    r#"<w:tbl><w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>merged</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc></w:tr>{}</w:tbl>"#,
                    control(
                        r#"<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc><w:tc><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc></w:tr>"#
                    )
                ),
            ]
            .concat();
            let document = CT_Document::from_xml(
                format!(
                    r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>{body}</w:body></w:document>"#
                )
                .as_bytes(),
            )
            .unwrap();
            let numbering = CT_Numbering::from_xml(
                br#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>"#,
            )
            .unwrap();
            HtmlInput {
                document,
                styles: CT_Styles::new_default(),
                numbering: Some(numbering),
                images: HashMap::new(),
                hyperlink_urls: HashMap::from([(
                    "rId1".to_owned(),
                    "https://example.com/".to_owned(),
                )]),
            }
        };
        let (wrapped, plain) = (input(true), input(false));

        let html = to_html_fragment(&wrapped, &HtmlOptions::default());
        assert_eq!(html, to_html_fragment(&plain, &HtmlOptions::default()));
        assert!(
            html.contains("<ol>\n<li>one</li>\n<li>two</li>\n<li>three</li>\n</ol>"),
            "{html}"
        );
        assert!(
            html.contains("<p><a href=\"https://example.com/\">link more</a> wrapped tagged</p>"),
            "{html}"
        );
        assert!(html.contains("<td rowspan=\"2\"><p>merged</p>"), "{html}");
        assert!(html.contains("<td><p>b</p>"), "{html}");

        let markdown = to_markdown(&wrapped);
        assert_eq!(markdown, to_markdown(&plain));
        assert!(
            markdown.contains("1. one\n1. two\n1. three\n"),
            "{markdown}"
        );
        assert!(
            markdown.contains("[link more](https://example.com/) wrapped tagged"),
            "{markdown}"
        );
        assert!(markdown.contains("|  | b |"), "{markdown}");
    }

    #[test]
    fn html_heading() {
        let mut doc = CT_Document::new();
        let mut p = CT_P::new();
        p.add_run("Chapter 1");
        p.properties = Some(rdocx_oxml::properties::CT_PPr {
            style_id: Some("Heading1".to_string()),
            ..Default::default()
        });
        doc.body.add_paragraph(p);

        let input = HtmlInput {
            document: doc,
            styles: CT_Styles::new_default(),
            numbering: None,
            images: HashMap::new(),
            hyperlink_urls: HashMap::new(),
        };

        let html = to_html_fragment(&input, &HtmlOptions::default());
        assert!(html.contains("<h1"));
        assert!(html.contains("Chapter 1"));
    }

    #[test]
    fn html_table() {
        let mut doc = CT_Document::new();
        let mut tbl = rdocx_oxml::table::CT_Tbl::new();
        let mut row = rdocx_oxml::table::CT_Row::new();
        let mut cell = rdocx_oxml::table::CT_Tc::new();
        let mut p = CT_P::new();
        p.add_run("Cell text");
        cell.content = vec![rdocx_oxml::table::CellContent::Paragraph(p)];
        row.cells.push(cell);
        tbl.rows.push(row);
        doc.body.content.push(BodyContent::Table(Box::new(tbl)));

        let input = HtmlInput {
            document: doc,
            styles: CT_Styles::new_default(),
            numbering: None,
            images: HashMap::new(),
            hyperlink_urls: HashMap::new(),
        };

        let html = to_html_fragment(&input, &HtmlOptions::default());
        assert!(html.contains("<table"));
        assert!(html.contains("<td"));
        assert!(html.contains("Cell text"));
    }

    #[test]
    fn markdown_heading() {
        let mut doc = CT_Document::new();
        let mut p = CT_P::new();
        p.add_run("Title");
        p.properties = Some(rdocx_oxml::properties::CT_PPr {
            style_id: Some("Heading1".to_string()),
            ..Default::default()
        });
        doc.body.add_paragraph(p);

        let input = HtmlInput {
            document: doc,
            styles: CT_Styles::new_default(),
            numbering: None,
            images: HashMap::new(),
            hyperlink_urls: HashMap::new(),
        };

        let md = to_markdown(&input);
        assert!(md.contains("# Title"));
    }

    #[test]
    fn html_bold_italic() {
        let mut doc = CT_Document::new();
        let mut p = CT_P::new();
        let mut r = rdocx_oxml::text::CT_R::new("bold text");
        r.properties = Some(rdocx_oxml::properties::CT_RPr {
            bold: Some(true),
            italic: Some(true),
            ..Default::default()
        });
        p.runs.push(r);
        doc.body.add_paragraph(p);

        let input = HtmlInput {
            document: doc,
            styles: CT_Styles::new_default(),
            numbering: None,
            images: HashMap::new(),
            hyperlink_urls: HashMap::new(),
        };

        let html = to_html_fragment(&input, &HtmlOptions::default());
        assert!(html.contains("<strong"));
        assert!(html.contains("<em"));
        assert!(html.contains("bold text"));
    }

    #[test]
    fn fx203_markdown_preserves_literal_whitespace_field_break_and_source_boundaries() {
        let cases = [
            (r#"<w:p><w:r><w:t xml:space="preserve"> \t </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve"> bold </w:t><w:br/></w:r></w:p>"#.replace("\\t", "\t"), " \t  **bold**   \n\n\n"),
            (r#"<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>left</w:t></w:r><w:fldSimple w:instr=" AUTHOR "><w:r><w:rPr><w:b/></w:rPr><w:t>field</w:t></w:r></w:fldSimple><w:r><w:rPr><w:b/></w:rPr><w:t>right</w:t></w:r></w:p>"#.into(), "**left****field****right**\n\n"),
            (r#"<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>left</w:t></w:r><w:customXml><w:fldSimple w:instr=" AUTHOR "><w:r><w:rPr><w:b/></w:rPr><w:t>field</w:t></w:r></w:fldSimple></w:customXml><w:r><w:rPr><w:b/></w:rPr><w:t>right</w:t></w:r></w:p>"#.into(), "**left****field****right**\n\n"),
            (r#"<w:p><w:hyperlink r:id="rId1"><w:r><w:rPr><w:b/></w:rPr><w:t>before</w:t><w:br w:type="page"/><w:t>after</w:t></w:r></w:hyperlink></w:p>"#.into(), "[**before**](https://example.com/a)\n\n---\n\n[**after**](https://example.com/a)\n\n"),
            (r#"<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve"> \t </w:t></w:r></w:p>"#.replace("\\t", "\t"), " \t \n\n"),
            (r#"<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>left</w:t></w:r><w:sdt><w:sdtContent><w:r><w:rPr><w:b/></w:rPr><w:t>control</w:t></w:r></w:sdtContent></w:sdt><w:ins w:id="1"><w:r><w:rPr><w:b/></w:rPr><w:t>inserted</w:t></w:r></w:ins><w:r><w:rPr><w:b/></w:rPr><w:t>right</w:t></w:r></w:p>"#.into(), "**left****control****inserted****right**\n\n"),
            (r#"<w:p><w:fldSimple w:instr=" AUTHOR "><w:r><w:rPr><w:b/></w:rPr><w:t>before</w:t><w:br/><w:t>after</w:t><w:br w:type="column"/><w:t>end</w:t></w:r></w:fldSimple></w:p>"#.into(), "**before**  \n**after**  \n**end**\n\n"),
            (r#"<w:p><w:sdt><w:sdtContent><w:r><w:rPr><w:b/></w:rPr><w:t>A</w:t></w:r><w:sdt><w:sdtContent><w:r><w:rPr><w:b/></w:rPr><w:t>B</w:t></w:r></w:sdtContent></w:sdt><w:ins w:id="2" w:author="Ada"><w:r><w:rPr><w:b/></w:rPr><w:t>C</w:t></w:r></w:ins><w:r><w:rPr><w:b/></w:rPr><w:t>D</w:t></w:r></w:sdtContent></w:sdt></w:p>"#.into(), "**A****B****C****D**\n\n"),
        ];
        for (body, expected) in cases {
            let xml = format!(
                r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:q="urn:opaque"><w:body>{body}</w:body></w:document>"#
            );
            let input = HtmlInput {
                document: CT_Document::from_xml(xml.as_bytes()).unwrap(),
                styles: CT_Styles::new_default(),
                numbering: None,
                images: HashMap::new(),
                hyperlink_urls: HashMap::from([("rId1".into(), "https://example.com/a".into())]),
            };
            let before = input.document.to_xml().unwrap();
            if body.contains("fldSimple") {
                let BodyContent::Paragraph(paragraph) = &input.document.body.content[0] else {
                    panic!("the field fixture owns one paragraph");
                };
                assert!(paragraph.accepted_view().runs.iter().all(|run| {
                    run.content
                        .iter()
                        .all(|content| !matches!(content, rdocx_oxml::text::RunContent::Field(_)))
                }));
                let (accepted, boundaries) = paragraph.accepted_view_with_field_owners();
                assert!(!boundaries.is_empty());
                assert!(boundaries.iter().all(|at| *at <= accepted.runs.len()));
                assert_eq!(accepted.accepted_text(), paragraph.accepted_text());
            }
            assert_eq!(to_markdown(&input), expected, "{body}");
            assert_eq!(input.document.to_xml().unwrap(), before);
        }
    }
}
