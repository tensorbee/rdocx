//! The HTML and Markdown emitters treat the source document as untrusted.
//!
//! A DOCX is an attacker-supplied file in most deployments, so nothing it
//! carries may become markup or a script trigger in the converted output.

use std::collections::HashMap;

use rdocx_html::{HtmlInput, HtmlOptions, ImageData, to_html_fragment, to_markdown};
use rdocx_oxml::document::CT_Document;
use rdocx_oxml::properties::{CT_RPr, CT_Shd};
use rdocx_oxml::styles::CT_Styles;
use rdocx_oxml::text::{CT_P, CT_R, HyperlinkSpan};

fn input_from(doc: CT_Document, hyperlink_urls: HashMap<String, String>) -> HtmlInput {
    HtmlInput {
        document: doc,
        styles: CT_Styles::new_default(),
        numbering: None,
        images: HashMap::new(),
        hyperlink_urls,
    }
}

fn linked_paragraph(url: &str) -> HtmlInput {
    let mut doc = CT_Document::new();
    let mut p = CT_P::new();
    p.runs.push(CT_R::new("click me"));
    p.hyperlinks.push(HyperlinkSpan {
        rel_id: Some("rId7".to_string()),
        anchor: None,
        tooltip: None,
        doc_location: None,
        run_start: 0,
        run_end: 1,
        extra_attributes: Vec::new(),
        extra_xml: Vec::new(),
        preserved_raw_before: None,
    });
    doc.body.add_paragraph(p);

    input_from(doc, HashMap::from([("rId7".to_string(), url.to_string())]))
}

#[test]
fn dangerous_url_schemes_are_dropped() {
    for url in [
        "javascript:alert(1)",
        "JaVaScRiPt:alert(1)",
        "  javascript:alert(1)",
        "data:text/html;base64,PHNjcmlwdD5hbGVydCgxKTwvc2NyaXB0Pg==",
        "vbscript:msgbox(1)",
    ] {
        let input = linked_paragraph(url);

        let html = to_html_fragment(&input, &HtmlOptions::default());
        let markdown = to_markdown(&input);

        assert!(!html.contains("<a "), "emitted a link for {url}: {html}");
        assert!(!markdown.contains(url), "markdown kept {url}: {markdown}");
        // The text itself must still render.
        assert!(html.contains("click me"));
        assert!(markdown.contains("click me"));
    }
}

#[test]
fn ordinary_url_schemes_still_link() {
    for url in [
        "https://example.com/a?b=c",
        "http://example.com",
        "mailto:someone@example.com",
        "#internal-anchor",
        "relative/page.html",
    ] {
        let input = linked_paragraph(url);
        let html = to_html_fragment(&input, &HtmlOptions::default());
        assert!(html.contains("<a href="), "no link for {url}: {html}");
    }
}

#[test]
fn font_names_cannot_escape_the_style_attribute() {
    let mut doc = CT_Document::new();
    let mut p = CT_P::new();
    let mut run = CT_R::new("styled");
    run.properties = Some(CT_RPr {
        font_ascii: Some("Arial\" onmouseover=\"alert(1)".to_string()),
        ..Default::default()
    });
    p.runs.push(run);
    doc.body.add_paragraph(p);

    let html = to_html_fragment(&input_from(doc, HashMap::new()), &HtmlOptions::default());

    assert!(!html.contains("onmouseover"), "{html}");
    assert!(html.contains("styled"));
}

#[test]
fn colours_must_be_hex_to_reach_the_output() {
    let mut doc = CT_Document::new();
    let mut p = CT_P::new();
    let mut run = CT_R::new("shaded");
    run.properties = Some(CT_RPr {
        color: Some("red;} body{display:none} .x{".to_string()),
        shading: Some(Box::new(CT_Shd {
            val: "clear".to_string(),
            color: None,
            fill: Some("\"><script>alert(1)</script>".to_string()),
            ..Default::default()
        })),
        ..Default::default()
    });
    p.runs.push(run);
    doc.body.add_paragraph(p);

    let html = to_html_fragment(&input_from(doc, HashMap::new()), &HtmlOptions::default());

    assert!(!html.contains("<script"), "{html}");
    assert!(!html.contains("display:none"), "{html}");
    assert!(html.contains("shaded"));
}

#[test]
fn valid_colours_are_preserved() {
    let mut doc = CT_Document::new();
    let mut p = CT_P::new();
    let mut run = CT_R::new("red text");
    run.properties = Some(CT_RPr {
        color: Some("FF0000".to_string()),
        ..Default::default()
    });
    p.runs.push(run);
    doc.body.add_paragraph(p);

    let html = to_html_fragment(&input_from(doc, HashMap::new()), &HtmlOptions::default());

    assert!(html.contains("color:#FF0000"), "{html}");
}

#[test]
fn image_content_type_is_escaped() {
    let mut doc = CT_Document::new();
    let mut p = CT_P::new();
    p.runs.push(CT_R::new("caption"));
    doc.body.add_paragraph(p);

    let mut input = input_from(doc, HashMap::new());
    input.images.insert(
        "rId1".to_string(),
        ImageData {
            data: vec![1, 2, 3],
            content_type: "image/png\"><script>alert(1)</script>".to_string(),
        },
    );

    let html = to_html_fragment(&input, &HtmlOptions::default());
    assert!(!html.contains("<script"), "{html}");
}

#[test]
fn text_content_is_escaped() {
    let mut doc = CT_Document::new();
    let mut p = CT_P::new();
    p.runs
        .push(CT_R::new("<script>alert('xss')</script> & more"));
    doc.body.add_paragraph(p);

    let html = to_html_fragment(&input_from(doc, HashMap::new()), &HtmlOptions::default());

    assert!(!html.contains("<script"), "{html}");
    assert!(html.contains("&lt;script&gt;"), "{html}");
    assert!(html.contains("&amp; more"), "{html}");
}

#[test]
fn fx281_chart_tables_escape_cells_keep_display_values_and_stay_out_of_cells() {
    use rdocx_html::{ChartTable, to_html_fragment_with_charts, to_markdown_with_charts};
    use rdocx_oxml::drawing::{CT_Drawing, CT_Inline};
    use rdocx_oxml::text::RunContent;
    let mut doc = CT_Document::new();
    let mut paragraph = CT_P::new();
    let mut inline = CT_Inline::new("", 914400, 914400);
    inline.chart_rel_id = Some("chart".into());
    paragraph.add_run("").content = vec![RunContent::Drawing(CT_Drawing::inline(inline))];
    doc.body.add_paragraph(paragraph.clone());
    let mut table = rdocx_oxml::table::CT_Tbl::new();
    let mut row = rdocx_oxml::table::CT_Row::new();
    let mut cell = rdocx_oxml::table::CT_Tc::new();
    cell.content = vec![rdocx_oxml::table::CellContent::Paragraph(paragraph)];
    row.cells.push(cell);
    table.rows.push(row);
    doc.body.add_table(table);
    let input = input_from(doc, HashMap::new());
    let charts = HashMap::from([(
        "chart".into(),
        ChartTable {
            title: Some("# Quarter <&>".into()),
            categories: vec!["Q|1".into()],
            series: vec![("Zeta <script>".into(), vec![1150.0, f64::NAN, 1190.5])],
        },
    )]);
    let rows = charts["chart"].rows();
    assert_eq!(
        rows,
        vec![
            vec!["", "Zeta <script>"],
            vec!["Q|1", "1150"],
            vec!["", ""],
            vec!["", "1190.5"]
        ]
    );
    let html = to_html_fragment_with_charts(&input, &HtmlOptions::default(), &charts);
    assert_eq!(html.matches("class=\"chart-data\"").count(), 1);
    assert!(html.contains("# Quarter &lt;&amp;&gt;") && html.contains("Zeta &lt;script&gt;"));
    assert!(!html.contains("<script>"));
    let markdown = to_markdown_with_charts(&input, &charts);
    assert!(
        markdown.trim_start().starts_with("\\# Quarter"),
        "{markdown}"
    );
    assert!(markdown.contains("Q\\|1") && markdown.contains("1150") && markdown.contains("1190.5"));
    assert!(!markdown.contains("NaN"));
    assert!(!to_html_fragment(&input, &HtmlOptions::default()).contains("chart-data"));
    assert!(!to_markdown(&input).contains("Zeta"));
}
