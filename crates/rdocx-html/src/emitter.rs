//! HTML element generation from OXML types.

use std::collections::HashMap;

use rdocx_oxml::document::{BodyContent, CT_Body};
use rdocx_oxml::numbering::CT_Numbering;
use rdocx_oxml::properties::CT_PPr;
use rdocx_oxml::shared::ST_Jc;
use rdocx_oxml::styles::CT_Styles;
use rdocx_oxml::table::{CT_Tbl, CellContent, VMerge};
use rdocx_oxml::text::{BreakType, CT_P, CT_R, RunContent};

use crate::css;
use crate::{HtmlOptions, ImageData};

/// Emit the full body content as HTML.
pub(crate) fn emit_body(
    body: &CT_Body,
    styles: &CT_Styles,
    numbering: Option<&CT_Numbering>,
    images: &HashMap<String, ImageData>,
    hyperlink_urls: &HashMap<String, String>,
    options: &HtmlOptions,
) -> String {
    let mut out = String::new();
    let mut list_stack: Vec<ListState> = Vec::new();

    for content in &body.content {
        match content {
            BodyContent::Paragraph(p) => {
                let list_info = detect_list(p, numbering);

                // Close lists that are no longer active
                while let Some(top) = list_stack.last() {
                    if let Some((_, level)) = &list_info {
                        if *level <= top.level && top.level > 0 {
                            close_list(&mut out, &mut list_stack);
                        } else {
                            break;
                        }
                    } else {
                        close_list(&mut out, &mut list_stack);
                    }
                }

                if let Some((is_ordered, level)) = list_info {
                    // Open new list levels if needed
                    let current_depth = list_stack.len() as u32;
                    if current_depth <= level {
                        for _ in current_depth..=level {
                            let tag = if is_ordered { "ol" } else { "ul" };
                            out.push_str(&format!("<{tag}>\n"));
                            list_stack.push(ListState {
                                ordered: is_ordered,
                                level,
                            });
                        }
                    }
                    out.push_str("<li>");
                    emit_paragraph_content(&mut out, p, styles, images, hyperlink_urls, options);
                    out.push_str("</li>\n");
                } else {
                    // Close all remaining lists
                    while !list_stack.is_empty() {
                        close_list(&mut out, &mut list_stack);
                    }
                    emit_paragraph(&mut out, p, styles, images, hyperlink_urls, options);
                }
            }
            BodyContent::Table(tbl) => {
                while !list_stack.is_empty() {
                    close_list(&mut out, &mut list_stack);
                }
                emit_table(&mut out, tbl, styles, images, hyperlink_urls, options);
            }
            BodyContent::RawXml(_) => {}
        }
    }

    // Close remaining lists
    while !list_stack.is_empty() {
        close_list(&mut out, &mut list_stack);
    }

    out
}

struct ListState {
    ordered: bool,
    level: u32,
}

fn close_list(out: &mut String, stack: &mut Vec<ListState>) {
    if let Some(state) = stack.pop() {
        let tag = if state.ordered { "ol" } else { "ul" };
        out.push_str(&format!("</{tag}>\n"));
    }
}

/// Detect heading level from paragraph style.
fn detect_heading_level(ppr: Option<&CT_PPr>, styles: &CT_Styles) -> Option<u32> {
    let ppr = ppr?;

    // Check outline_lvl directly
    if let Some(lvl) = ppr.outline_lvl {
        return Some(lvl + 1); // outline_lvl is 0-based, headings are 1-based
    }

    // Check style_id
    let style_id = ppr.style_id.as_deref()?;
    if let Some(level) = style_id.strip_prefix("Heading")
        && let Ok(n) = level.parse::<u32>()
        && (1..=6).contains(&n)
    {
        return Some(n);
    }

    // Check style's outline level
    if let Some(style) = styles.get_by_id(style_id)
        && let Some(spr) = &style.ppr
        && let Some(lvl) = spr.outline_lvl
    {
        return Some(lvl + 1);
    }

    None
}

/// Detect if a paragraph is a list item.
fn detect_list(para: &CT_P, numbering: Option<&CT_Numbering>) -> Option<(bool, u32)> {
    let ppr = para.properties.as_ref()?;
    let num_id = ppr.num_id?;
    let ilvl = ppr.num_ilvl.unwrap_or(0);

    if num_id == 0 {
        return None;
    }

    let numbering = numbering?;

    // Look up the numbering definition to determine if ordered or bullet
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

/// Emit a complete paragraph element.
fn emit_paragraph(
    out: &mut String,
    para: &CT_P,
    styles: &CT_Styles,
    images: &HashMap<String, ImageData>,
    hyperlink_urls: &HashMap<String, String>,
    options: &HtmlOptions,
) {
    let heading_level = detect_heading_level(para.properties.as_ref(), styles);

    let tag = match heading_level {
        Some(1) => "h1",
        Some(2) => "h2",
        Some(3) => "h3",
        Some(4) => "h4",
        Some(5) => "h5",
        Some(6) => "h6",
        _ => "p",
    };

    let style = css::paragraph_style(para.properties.as_ref());
    if style.is_empty() {
        out.push_str(&format!("<{tag}>"));
    } else {
        out.push_str(&format!("<{tag} style=\"{style}\">"));
    }

    emit_paragraph_content(out, para, styles, images, hyperlink_urls, options);

    out.push_str(&format!("</{tag}>\n"));
}

/// Emit the inner content of a paragraph (runs and hyperlinks).
fn emit_paragraph_content(
    out: &mut String,
    para: &CT_P,
    _styles: &CT_Styles,
    images: &HashMap<String, ImageData>,
    hyperlink_urls: &HashMap<String, String>,
    options: &HtmlOptions,
) {
    // Build a map of which runs are inside hyperlinks
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

    let mut current_link: Option<&str> = None;

    for (run_idx, run) in para.runs.iter().enumerate() {
        let in_link = hyperlink_map.get(&run_idx).copied();

        // Open/close link tags as needed
        match (current_link, in_link) {
            (None, Some(url)) => {
                out.push_str(&format!("<a href=\"{}\">", escape_html_attr(url)));
                current_link = Some(url);
            }
            (Some(_), None) => {
                out.push_str("</a>");
                current_link = None;
            }
            (Some(old), Some(new)) if old != new => {
                out.push_str("</a>");
                out.push_str(&format!("<a href=\"{}\">", escape_html_attr(new)));
                current_link = Some(new);
            }
            _ => {}
        }

        emit_run(out, run, images, options);
    }

    // Close any open link
    if current_link.is_some() {
        out.push_str("</a>");
    }
}

/// Emit a single run.
fn emit_run(
    out: &mut String,
    run: &CT_R,
    images: &HashMap<String, ImageData>,
    options: &HtmlOptions,
) {
    let tags = css::run_tags(run.properties.as_ref());
    let style = css::run_style(run.properties.as_ref());

    // Open semantic tags
    let has_style = !style.is_empty();
    if has_style {
        out.push_str(&format!("<span style=\"{style}\">"));
    }
    if tags.bold {
        out.push_str("<strong>");
    }
    if tags.italic {
        out.push_str("<em>");
    }
    if tags.underline {
        out.push_str("<u>");
    }
    if tags.strike {
        out.push_str("<s>");
    }
    if tags.superscript {
        out.push_str("<sup>");
    }
    if tags.subscript {
        out.push_str("<sub>");
    }

    // Emit content
    for content in &run.content {
        match content {
            RunContent::Text(t) => {
                out.push_str(&escape_html(&t.text));
            }
            RunContent::Tab => {
                out.push_str("&emsp;");
            }
            RunContent::Break(bt) => match bt {
                BreakType::Line => out.push_str("<br>"),
                BreakType::Page => out.push_str("<hr>"),
                BreakType::Column => out.push_str("<br>"),
            },
            RunContent::Drawing(drawing) => {
                if options.inline_images {
                    // Try to find image data
                    let embed_id = drawing
                        .inline
                        .as_ref()
                        .map(|i| i.embed_id.as_str())
                        .or_else(|| drawing.anchor.as_ref().map(|a| a.embed_id.as_str()));

                    if let Some(eid) = embed_id
                        && let Some(img_data) = images.get(eid)
                    {
                        emit_image(out, img_data);
                    }
                }
            }
            RunContent::Field { .. }
            | RunContent::FootnoteRef { .. }
            | RunContent::EndnoteRef { .. } => {}
        }
    }

    // Close semantic tags (reverse order)
    if tags.subscript {
        out.push_str("</sub>");
    }
    if tags.superscript {
        out.push_str("</sup>");
    }
    if tags.strike {
        out.push_str("</s>");
    }
    if tags.underline {
        out.push_str("</u>");
    }
    if tags.italic {
        out.push_str("</em>");
    }
    if tags.bold {
        out.push_str("</strong>");
    }
    if has_style {
        out.push_str("</span>");
    }
}

/// Emit an inline image as a base64 data URI.
fn emit_image(out: &mut String, img: &ImageData) {
    let base64 = base64_encode(&img.data);
    out.push_str(&format!(
        "<img src=\"data:{};base64,{}\" style=\"max-width:100%\">",
        img.content_type, base64
    ));
}

/// Emit a table.
fn emit_table(
    out: &mut String,
    tbl: &CT_Tbl,
    styles: &CT_Styles,
    images: &HashMap<String, ImageData>,
    hyperlink_urls: &HashMap<String, String>,
    options: &HtmlOptions,
) {
    let mut table_style = String::new();

    if let Some(props) = &tbl.properties {
        // Table alignment
        if let Some(jc) = &props.jc {
            match jc {
                ST_Jc::Center => table_style.push_str("margin-left:auto;margin-right:auto;"),
                ST_Jc::Right => table_style.push_str("margin-left:auto;"),
                _ => {}
            }
        }

        // Table borders
        if let Some(borders) = &props.borders
            && (borders.top.is_some()
                || borders.bottom.is_some()
                || borders.left.is_some()
                || borders.right.is_some()
                || borders.inside_h.is_some()
                || borders.inside_v.is_some())
        {
            table_style.push_str("border:1px solid #000;");
        }
    }

    if table_style.is_empty() {
        out.push_str("<table>\n");
    } else {
        out.push_str(&format!("<table style=\"{table_style}\">\n"));
    }

    for row in &tbl.rows {
        out.push_str("<tr>\n");
        for cell in &row.cells {
            // Skip vmerge continue cells
            if let Some(props) = &cell.properties
                && matches!(props.v_merge, Some(VMerge::Continue))
            {
                continue;
            }

            let mut td_attrs = String::new();
            let mut td_style = String::new();

            if let Some(props) = &cell.properties {
                // Column span
                if let Some(span) = props.grid_span
                    && span > 1
                {
                    td_attrs.push_str(&format!(" colspan=\"{span}\""));
                }

                // Row span (count consecutive vmerge continue cells below)
                if matches!(props.v_merge, Some(VMerge::Restart)) {
                    let rowspan = count_vmerge_span(tbl, &row.cells, cell);
                    if rowspan > 1 {
                        td_attrs.push_str(&format!(" rowspan=\"{rowspan}\""));
                    }
                }

                // Vertical alignment
                if let Some(va) = &props.v_align {
                    let css_va = match va {
                        rdocx_oxml::table::ST_VerticalJc::Top => "top",
                        rdocx_oxml::table::ST_VerticalJc::Center => "middle",
                        rdocx_oxml::table::ST_VerticalJc::Bottom => "bottom",
                    };
                    td_style.push_str(&format!("vertical-align:{css_va};"));
                }

                // Cell shading
                if let Some(shd) = &props.shading
                    && let Some(fill) = &shd.fill
                    && fill != "auto"
                    && fill != "FFFFFF"
                {
                    td_style.push_str(&format!("background-color:#{fill};"));
                }

                // Cell borders
                if let Some(borders) = &props.borders
                    && (borders.top.is_some()
                        || borders.bottom.is_some()
                        || borders.left.is_some()
                        || borders.right.is_some())
                {
                    td_style.push_str("border:1px solid #000;");
                }
            }

            if td_style.is_empty() {
                out.push_str(&format!("<td{td_attrs}>"));
            } else {
                out.push_str(&format!("<td{td_attrs} style=\"{td_style}\">"));
            }

            for content in &cell.content {
                match content {
                    CellContent::Paragraph(p) => {
                        emit_paragraph(out, p, styles, images, hyperlink_urls, options);
                    }
                    CellContent::Table(nested) => {
                        emit_table(out, nested, styles, images, hyperlink_urls, options);
                    }
                }
            }

            out.push_str("</td>\n");
        }
        out.push_str("</tr>\n");
    }

    out.push_str("</table>\n");
}

/// Count the rowspan for a vmerge restart cell.
fn count_vmerge_span(
    tbl: &CT_Tbl,
    _current_cells: &[rdocx_oxml::table::CT_Tc],
    cell: &rdocx_oxml::table::CT_Tc,
) -> u32 {
    // Find the column index of this cell
    let col_idx = _current_cells.iter().position(|c| std::ptr::eq(c, cell));
    let Some(col_idx) = col_idx else {
        return 1;
    };

    // Find which row this cell is in
    let row_idx = tbl
        .rows
        .iter()
        .position(|r| std::ptr::eq(r.cells.as_slice(), _current_cells));
    let Some(row_idx) = row_idx else {
        return 1;
    };

    let mut span = 1;
    for row in tbl.rows.iter().skip(row_idx + 1) {
        if let Some(next_cell) = row.cells.get(col_idx)
            && let Some(props) = &next_cell.properties
            && matches!(props.v_merge, Some(VMerge::Continue))
        {
            span += 1;
            continue;
        }
        break;
    }

    span
}

/// Escape HTML special characters.
fn escape_html(text: &str) -> String {
    let mut result = String::with_capacity(text.len());
    for ch in text.chars() {
        match ch {
            '&' => result.push_str("&amp;"),
            '<' => result.push_str("&lt;"),
            '>' => result.push_str("&gt;"),
            '"' => result.push_str("&quot;"),
            _ => result.push(ch),
        }
    }
    result
}

/// Escape HTML attribute value.
fn escape_html_attr(text: &str) -> String {
    escape_html(text)
}

/// Simple base64 encoding.
fn base64_encode(data: &[u8]) -> String {
    const ALPHABET: &[u8; 64] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";

    let mut result = String::with_capacity(data.len().div_ceil(3) * 4);
    let chunks = data.chunks(3);

    for chunk in chunks {
        let b0 = chunk[0] as u32;
        let b1 = if chunk.len() > 1 { chunk[1] as u32 } else { 0 };
        let b2 = if chunk.len() > 2 { chunk[2] as u32 } else { 0 };

        let n = (b0 << 16) | (b1 << 8) | b2;

        result.push(ALPHABET[((n >> 18) & 63) as usize] as char);
        result.push(ALPHABET[((n >> 12) & 63) as usize] as char);

        if chunk.len() > 1 {
            result.push(ALPHABET[((n >> 6) & 63) as usize] as char);
        } else {
            result.push('=');
        }

        if chunk.len() > 2 {
            result.push(ALPHABET[(n & 63) as usize] as char);
        } else {
            result.push('=');
        }
    }

    result
}
