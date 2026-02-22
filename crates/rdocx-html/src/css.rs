//! CSS generation from OXML properties.

use rdocx_oxml::properties::{CT_PPr, CT_RPr};
use rdocx_oxml::shared::ST_Jc;

/// Generate base CSS styles for the HTML document.
pub(crate) fn generate_base_css() -> String {
    r#"body { font-family: 'Calibri', 'Arial', sans-serif; font-size: 11pt; line-height: 1.4; margin: 1in; color: #000; }
h1, h2, h3, h4, h5, h6 { margin-top: 0.5em; margin-bottom: 0.3em; }
h1 { font-size: 2em; }
h2 { font-size: 1.5em; }
h3 { font-size: 1.17em; }
h4 { font-size: 1em; }
h5 { font-size: 0.83em; }
h6 { font-size: 0.67em; }
p { margin: 0 0 8pt 0; }
table { border-collapse: collapse; margin: 8pt 0; }
td, th { padding: 4pt 6pt; vertical-align: top; }
ul, ol { margin: 4pt 0; padding-left: 36pt; }
a { color: #0563C1; text-decoration: underline; }"#
        .to_string()
}

/// Generate inline CSS style string from paragraph properties.
pub(crate) fn paragraph_style(ppr: Option<&CT_PPr>) -> String {
    let Some(ppr) = ppr else {
        return String::new();
    };

    let mut styles = Vec::new();

    // Text alignment
    if let Some(jc) = &ppr.jc {
        let align = match jc {
            ST_Jc::Left | ST_Jc::Start => "left",
            ST_Jc::Center => "center",
            ST_Jc::Right | ST_Jc::End => "right",
            ST_Jc::Both | ST_Jc::Distribute => "justify",
        };
        styles.push(format!("text-align:{align}"));
    }

    // Spacing before/after
    if let Some(sp) = ppr.space_before {
        let pt = sp.to_pt();
        if pt > 0.0 {
            styles.push(format!("margin-top:{pt:.1}pt"));
        }
    }
    if let Some(sp) = ppr.space_after {
        let pt = sp.to_pt();
        styles.push(format!("margin-bottom:{pt:.1}pt"));
    }

    // Indentation
    if let Some(ind) = ppr.ind_left {
        let pt = ind.to_pt();
        if pt > 0.0 {
            styles.push(format!("margin-left:{pt:.1}pt"));
        }
    }
    if let Some(ind) = ppr.ind_right {
        let pt = ind.to_pt();
        if pt > 0.0 {
            styles.push(format!("margin-right:{pt:.1}pt"));
        }
    }
    if let Some(ind) = ppr.ind_first_line {
        let pt = ind.to_pt();
        if pt > 0.0 {
            styles.push(format!("text-indent:{pt:.1}pt"));
        }
    }

    // Background shading
    if let Some(shd) = &ppr.shading
        && let Some(fill) = &shd.fill
        && fill != "auto"
        && fill != "FFFFFF"
    {
        styles.push(format!("background-color:#{fill}"));
    }

    // Line spacing
    if let Some(ls) = ppr.line_spacing {
        let rule = ppr.line_rule.as_deref().unwrap_or("auto");
        match rule {
            "auto" => {
                let factor = ls.0 as f64 / 240.0;
                if (factor - 1.0).abs() > 0.01 {
                    styles.push(format!("line-height:{factor:.2}"));
                }
            }
            "exact" | "atLeast" => {
                let pt = ls.to_pt();
                styles.push(format!("line-height:{pt:.1}pt"));
            }
            _ => {}
        }
    }

    if styles.is_empty() {
        String::new()
    } else {
        styles.join(";")
    }
}

/// Generate inline CSS style string from run properties.
pub(crate) fn run_style(rpr: Option<&CT_RPr>) -> String {
    let Some(rpr) = rpr else {
        return String::new();
    };

    let mut styles = Vec::new();

    // Font family
    if let Some(font) = &rpr.font_ascii {
        styles.push(format!("font-family:'{font}'"));
    }

    // Font size (half-points → pt)
    if let Some(sz) = rpr.sz {
        let pt = sz.0 as f64 / 2.0;
        styles.push(format!("font-size:{pt}pt"));
    }

    // Color
    if let Some(color) = &rpr.color
        && color != "auto"
        && color != "000000"
    {
        styles.push(format!("color:#{color}"));
    }

    // Background/highlight via shading
    if let Some(shd) = &rpr.shading
        && let Some(fill) = &shd.fill
        && fill != "auto"
        && fill != "FFFFFF"
    {
        styles.push(format!("background-color:#{fill}"));
    }

    // Character spacing
    if let Some(spacing) = rpr.spacing {
        let pt = spacing.to_pt();
        if pt.abs() > 0.01 {
            styles.push(format!("letter-spacing:{pt:.1}pt"));
        }
    }

    // Small caps
    if rpr.small_caps == Some(true) {
        styles.push("font-variant:small-caps".to_string());
    }

    // All caps
    if rpr.caps == Some(true) {
        styles.push("text-transform:uppercase".to_string());
    }

    if styles.is_empty() {
        String::new()
    } else {
        styles.join(";")
    }
}

/// Determine which semantic HTML tags to wrap around run content.
pub(crate) struct RunTags {
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strike: bool,
    pub superscript: bool,
    pub subscript: bool,
}

pub(crate) fn run_tags(rpr: Option<&CT_RPr>) -> RunTags {
    let Some(rpr) = rpr else {
        return RunTags {
            bold: false,
            italic: false,
            underline: false,
            strike: false,
            superscript: false,
            subscript: false,
        };
    };

    RunTags {
        bold: rpr.bold == Some(true),
        italic: rpr.italic == Some(true),
        underline: rpr.underline.is_some(),
        strike: rpr.strike == Some(true) || rpr.dstrike == Some(true),
        superscript: rpr.vert_align.as_deref() == Some("superscript"),
        subscript: rpr.vert_align.as_deref() == Some("subscript"),
    }
}
