//! Markdown generation from OXML types (GFM-compatible).

use std::collections::HashMap;

use rdocx_oxml::document::{BodyContent, CT_Body};
use rdocx_oxml::numbering::CT_Numbering;
use rdocx_oxml::properties::CT_PPr;
use rdocx_oxml::styles::CT_Styles;
use rdocx_oxml::table::{CT_Tbl, CellContent};
use rdocx_oxml::text::{BreakType, CT_P, CT_R, RunContent};

/// Emit the full body content as Markdown.
pub(crate) fn emit_markdown(
    body: &CT_Body,
    styles: &CT_Styles,
    numbering: Option<&CT_Numbering>,
    hyperlink_urls: &HashMap<String, String>,
) -> String {
    let mut out = String::new();

    for content in &body.content {
        match content {
            BodyContent::Paragraph(p) => {
                emit_paragraph(&mut out, p, styles, numbering, hyperlink_urls);
            }
            BodyContent::Table(tbl) => {
                emit_table(&mut out, tbl, hyperlink_urls);
            }
            BodyContent::RawXml(_) => {}
        }
    }

    out
}

/// Detect heading level from paragraph properties/style.
fn detect_heading_level(ppr: Option<&CT_PPr>, styles: &CT_Styles) -> Option<u32> {
    let ppr = ppr?;

    if let Some(lvl) = ppr.outline_lvl {
        return Some(lvl + 1);
    }

    let style_id = ppr.style_id.as_deref()?;
    if let Some(level) = style_id.strip_prefix("Heading")
        && let Ok(n) = level.parse::<u32>()
        && (1..=6).contains(&n)
    {
        return Some(n);
    }

    if let Some(style) = styles.get_by_id(style_id)
        && let Some(spr) = &style.ppr
        && let Some(lvl) = spr.outline_lvl
    {
        return Some(lvl + 1);
    }

    None
}

/// Detect if a paragraph is a list item. Returns (is_ordered, ilvl).
fn detect_list(para: &CT_P, numbering: Option<&CT_Numbering>) -> Option<(bool, u32)> {
    let ppr = para.properties.as_ref()?;
    let num_id = ppr.num_id?;
    let ilvl = ppr.num_ilvl.unwrap_or(0);

    if num_id == 0 {
        return None;
    }

    let numbering = numbering?;

    let abstract_id = numbering
        .nums
        .iter()
        .find(|n| n.num_id == num_id)
        .map(|n| n.abstract_num_id)?;

    let abstract_num = numbering
        .abstract_nums
        .iter()
        .find(|a| a.abstract_num_id == abstract_id)?;

    let level = abstract_num.levels.iter().find(|l| l.ilvl == ilvl)?;

    let is_ordered = !matches!(
        level.num_fmt,
        Some(rdocx_oxml::numbering::ST_NumberFormat::Bullet)
    );

    Some((is_ordered, ilvl))
}

/// Emit a paragraph as Markdown.
fn emit_paragraph(
    out: &mut String,
    para: &CT_P,
    styles: &CT_Styles,
    numbering: Option<&CT_Numbering>,
    hyperlink_urls: &HashMap<String, String>,
) {
    let heading_level = detect_heading_level(para.properties.as_ref(), styles);
    let list_info = detect_list(para, numbering);

    // Collect inline text for the paragraph
    let text = collect_paragraph_text(para, hyperlink_urls);

    // Skip empty paragraphs (but still emit blank line for spacing)
    if text.trim().is_empty() && heading_level.is_none() && list_info.is_none() {
        out.push('\n');
        return;
    }

    if let Some(level) = heading_level {
        // Heading
        let hashes = "#".repeat(level.min(6) as usize);
        out.push_str(&hashes);
        out.push(' ');
        out.push_str(text.trim());
        out.push_str("\n\n");
    } else if let Some((is_ordered, ilvl)) = list_info {
        // List item with indentation
        let indent = "  ".repeat(ilvl as usize);
        if is_ordered {
            out.push_str(&format!("{indent}1. {}\n", text.trim()));
        } else {
            out.push_str(&format!("{indent}- {}\n", text.trim()));
        }
    } else {
        // Normal paragraph
        out.push_str(&text);
        out.push_str("\n\n");
    }
}

/// Collect all text from a paragraph, applying inline formatting.
fn collect_paragraph_text(para: &CT_P, hyperlink_urls: &HashMap<String, String>) -> String {
    let mut out = String::new();

    // Build hyperlink map
    let mut hyperlink_map: HashMap<usize, &str> = HashMap::new();
    for hl in &para.hyperlinks {
        if let Some(rel_id) = &hl.rel_id
            && let Some(url) = hyperlink_urls.get(rel_id)
        {
            for i in hl.run_start..hl.run_end {
                hyperlink_map.insert(i, url);
            }
        }
    }

    // Track link state to group consecutive runs in same link
    let mut current_link: Option<&str> = None;
    let mut link_text = String::new();

    for (run_idx, run) in para.runs.iter().enumerate() {
        let in_link = hyperlink_map.get(&run_idx).copied();

        // Handle link transitions
        match (current_link, in_link) {
            (Some(url), None) => {
                // Close link
                out.push_str(&format!("[{}]({})", link_text, url));
                link_text.clear();
                current_link = None;
            }
            (Some(old_url), Some(new_url)) if old_url != new_url => {
                // Close old, start new
                out.push_str(&format!("[{}]({})", link_text, old_url));
                link_text.clear();
                current_link = Some(new_url);
            }
            (None, Some(url)) => {
                current_link = Some(url);
            }
            _ => {}
        }

        let run_text = collect_run_text(run);

        if current_link.is_some() {
            link_text.push_str(&run_text);
        } else {
            out.push_str(&run_text);
        }
    }

    // Close any remaining link
    if let Some(url) = current_link {
        out.push_str(&format!("[{}]({})", link_text, url));
    }

    out
}

/// Collect text from a single run, applying Markdown formatting.
fn collect_run_text(run: &CT_R) -> String {
    let mut raw = String::new();

    for content in &run.content {
        match content {
            RunContent::Text(t) => raw.push_str(&t.text),
            RunContent::Tab => raw.push('\t'),
            RunContent::Break(bt) => match bt {
                BreakType::Line => raw.push_str("  \n"),
                BreakType::Page => raw.push_str("\n---\n"),
                BreakType::Column => raw.push_str("  \n"),
            },
            RunContent::Drawing(_)
            | RunContent::Field { .. }
            | RunContent::FootnoteRef { .. }
            | RunContent::EndnoteRef { .. } => {}
        }
    }

    if raw.is_empty() {
        return raw;
    }

    // Apply formatting wrappers
    let rpr = run.properties.as_ref();
    let bold = rpr.is_some_and(|r| r.bold == Some(true));
    let italic = rpr.is_some_and(|r| r.italic == Some(true));
    let strike = rpr.is_some_and(|r| r.strike == Some(true) || r.dstrike == Some(true));

    if bold && italic {
        format!("***{raw}***")
    } else if bold {
        format!("**{raw}**")
    } else if italic {
        format!("*{raw}*")
    } else if strike {
        format!("~~{raw}~~")
    } else {
        raw
    }
}

/// Emit a table as a GFM pipe table.
fn emit_table(out: &mut String, tbl: &CT_Tbl, hyperlink_urls: &HashMap<String, String>) {
    if tbl.rows.is_empty() {
        return;
    }

    // Collect all rows as cell text vectors
    let mut rows: Vec<Vec<String>> = Vec::new();
    let mut max_cols = 0;

    for row in &tbl.rows {
        let mut cells: Vec<String> = Vec::new();
        for cell in &row.cells {
            let text = collect_cell_text(cell, hyperlink_urls);
            cells.push(text);
        }
        if cells.len() > max_cols {
            max_cols = cells.len();
        }
        rows.push(cells);
    }

    if rows.is_empty() || max_cols == 0 {
        return;
    }

    // Pad rows to same length
    for row in &mut rows {
        while row.len() < max_cols {
            row.push(String::new());
        }
    }

    // Emit header row (first row)
    out.push('|');
    for cell in &rows[0] {
        out.push_str(&format!(" {} |", cell.trim()));
    }
    out.push('\n');

    // Emit separator
    out.push('|');
    for _ in 0..max_cols {
        out.push_str(" --- |");
    }
    out.push('\n');

    // Emit data rows
    for row in rows.iter().skip(1) {
        out.push('|');
        for cell in row {
            out.push_str(&format!(" {} |", cell.trim()));
        }
        out.push('\n');
    }

    out.push('\n');
}

/// Collect text from a table cell.
fn collect_cell_text(
    cell: &rdocx_oxml::table::CT_Tc,
    hyperlink_urls: &HashMap<String, String>,
) -> String {
    let mut parts = Vec::new();

    for content in &cell.content {
        match content {
            CellContent::Paragraph(p) => {
                let text = collect_paragraph_text(p, hyperlink_urls);
                let trimmed = text.trim().to_string();
                if !trimmed.is_empty() {
                    parts.push(trimmed);
                }
            }
            CellContent::Table(_) => {
                parts.push("(nested table)".to_string());
            }
        }
    }

    parts.join(" ")
}
