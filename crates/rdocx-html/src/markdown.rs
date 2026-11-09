//! Markdown generation from OXML types (GFM-compatible).

use std::cmp::Reverse;
use std::collections::HashMap;

use rdocx_oxml::numbering::CT_Numbering;
use rdocx_oxml::properties::CT_PPr;
use rdocx_oxml::styles::CT_Styles;
use rdocx_oxml::table::CT_Tbl;
use rdocx_oxml::text::{BreakType, CT_P, CT_R, RunContent, SpecialCharacter};

use crate::{Block, HtmlInput, ImageData, body_blocks, cell_blocks};

/// Emit the full body content as Markdown.
pub(crate) fn emit_markdown(
    input: &HtmlInput,
    charts: &HashMap<String, crate::ChartTable>,
) -> String {
    let styles = &input.styles;
    let hyperlink_urls = &input.hyperlink_urls;
    let images = &input.images;
    let mut out = String::new();
    // The open list levels, each with the column its item text starts at.
    let mut list_levels: Vec<(u32, usize, u32)> = Vec::new();

    let mut previous_quote_owner = None;
    for block in body_blocks(&input.document.body, true) {
        let style_id = match &block {
            Block::Paragraph(p, _) => p.properties.as_ref().and_then(|p| p.style_id.as_deref()),
            Block::Table(t) => t.properties.as_ref().and_then(|p| p.style_id.as_deref()),
        };
        let owner = quote_owner(style_id, styles);
        let shared_depth =
            owner
                .zip(previous_quote_owner)
                .map_or(0, |(current, previous): (&str, &str)| {
                    current
                        .split('.')
                        .zip(previous.split('.'))
                        .take_while(|(a, b)| a == b)
                        .count()
                });
        if shared_depth > 0 && out.ends_with("\n\n") {
            out.pop();
            let path = container_path(style_id, styles).unwrap_or_default();
            if !path.starts_with('q')
                && let Some((_, column, _)) = list_levels.last()
            {
                out.push_str(&" ".repeat(*column));
            }
            out.push_str(&"> ".repeat(shared_depth));
            out.push('\n');
        }
        if previous_quote_owner.is_some() && owner != previous_quote_owner && shared_depth == 0 {
            leave_list(&mut out, &mut list_levels);
            if !out.ends_with("\n\n") {
                out.push('\n');
            }
        }
        previous_quote_owner = owner;
        match block {
            Block::Paragraph(p, boundaries) => {
                emit_paragraph(&mut out, &p, &boundaries, input, &mut list_levels);
                for id in crate::paragraph_chart_ids(&p) {
                    if let Some(chart) = charts.get(&id) {
                        leave_list(&mut out, &mut list_levels);
                        emit_chart_table(&mut out, chart);
                    }
                }
            }
            Block::Table(tbl) => {
                let mut table = String::new();
                emit_table(&mut table, tbl, styles, hyperlink_urls, images);
                if !table.is_empty() {
                    let path = container_path(
                        tbl.properties.as_ref().and_then(|p| p.style_id.as_deref()),
                        styles,
                    )
                    .unwrap_or_default();
                    let outer_prefix = "> ".repeat(path.chars().take_while(|&c| c == 'q').count());
                    if path.contains('q') {
                        let prefix = "> ".repeat(path.chars().filter(|&c| c == 'q').count());
                        table = table
                            .lines()
                            .map(|line| {
                                if line.is_empty() {
                                    String::new()
                                } else {
                                    format!("{prefix}{line}")
                                }
                            })
                            .collect::<Vec<_>>()
                            .join("\n")
                            + "\n";
                    }
                    if tbl
                        .properties
                        .as_ref()
                        .and_then(|properties| properties.indent.as_ref())
                        .is_some_and(|indent| indent.width_type == "dxa" && indent.w > 0)
                        && let Some((_, column, _)) = list_levels.last()
                    {
                        if !out.ends_with("\n\n") {
                            out.push_str(&outer_prefix);
                            out.push('\n');
                        }
                        for line in table.lines() {
                            let line = line.strip_prefix(&outer_prefix).unwrap_or(line);
                            out.push_str(&format!("{outer_prefix}{}{line}\n", " ".repeat(*column)));
                        }
                        out.push_str(&outer_prefix);
                        out.push('\n');
                    } else {
                        leave_quoted_list(&mut out, &mut list_levels, &outer_prefix);
                        out.push_str(&table);
                    }
                }
            }
        }
    }

    out
}

/// End an open list before another kind of block.
///
/// A list item ends in a single newline, and a block that follows it without
/// a blank line is read as a lazy continuation of the last item.
fn leave_list(out: &mut String, list_levels: &mut Vec<(u32, usize, u32)>) {
    if !list_levels.is_empty() && !out.ends_with("\n\n") {
        out.push('\n');
    }
    list_levels.clear();
}

fn leave_quoted_list(out: &mut String, list_levels: &mut Vec<(u32, usize, u32)>, prefix: &str) {
    if !list_levels.is_empty() && !out.ends_with("\n\n") {
        out.push_str(prefix);
        out.push('\n');
    }
    list_levels.clear();
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
fn detect_list(para: &CT_P, numbering: Option<&CT_Numbering>) -> Option<(bool, u32, u32, u32)> {
    let ppr = para.properties.as_ref()?;
    let num_id = ppr.num_id?;
    let ilvl = ppr.num_ilvl.unwrap_or(0);

    if num_id == 0 {
        return None;
    }

    let numbering = numbering?;

    let instance = numbering.nums.iter().find(|n| n.num_id == num_id)?;
    let abstract_id = instance.abstract_num_id;

    let abstract_num = numbering
        .abstract_nums
        .iter()
        .find(|a| a.abstract_num_id == abstract_id)?;

    let level = abstract_num.levels.iter().find(|l| l.ilvl == ilvl)?;

    use rdocx_oxml::numbering::ST_NumberFormat;
    let is_ordered = match level.num_fmt.as_ref() {
        Some(ST_NumberFormat::Bullet) => false,
        None
        | Some(
            ST_NumberFormat::Decimal
            | ST_NumberFormat::UpperRoman
            | ST_NumberFormat::LowerRoman
            | ST_NumberFormat::UpperLetter
            | ST_NumberFormat::LowerLetter
            | ST_NumberFormat::Ordinal
            | ST_NumberFormat::None,
        ) => true,
        Some(_) => return None,
    };

    let start = instance
        .level_overrides
        .iter()
        .find(|item| item.ilvl == ilvl)
        .and_then(|item| item.start_override)
        .or(level.start)
        .unwrap_or(1);
    Some((is_ordered, ilvl, num_id, start))
}

/// Emit a paragraph as Markdown.
fn emit_paragraph(
    out: &mut String,
    para: &CT_P,
    boundaries: &[usize],
    input: &HtmlInput,
    list_levels: &mut Vec<(u32, usize, u32)>,
) {
    let styles = &input.styles;
    let numbering = input.numbering.as_ref();
    let hyperlink_urls = &input.hyperlink_urls;
    let images = &input.images;
    let mut heading_level = detect_heading_level(para.properties.as_ref(), styles);
    let list_info = detect_list(para, numbering);

    // Collect inline text for the paragraph
    let mut text = collect_paragraph_text(para, styles, hyperlink_urls, images, boundaries);

    let style_id = para
        .properties
        .as_ref()
        .and_then(|properties| properties.style_id.as_deref());
    let compound_quote = quote_owner(style_id, styles).is_some();
    let container_path = container_path(style_id, styles).unwrap_or_default();
    let outer_quotes = container_path.chars().take_while(|&c| c == 'q').count();
    let style_id = style_id.and_then(|id| {
        styles
            .get_by_id(id)
            .filter(|_| !container_path.is_empty())
            .and_then(|style| style.based_on.as_deref())
            .or(Some(id))
    });
    let style_name = style_id
        .and_then(|id| styles.get_by_id(id))
        .and_then(|style| style.name.as_deref());
    let role = |id: &str, name: &str| {
        style_id == Some(id)
            || style_name.is_some_and(|existing| existing.eq_ignore_ascii_case(name))
    };
    let quoted_pre = role("HTMLQuotedPreformatted", "HTML Quoted Preformatted");
    if role("HTMLThematicBreak", "HTML Thematic Break") {
        text = "---".to_owned();
    } else if role("HTMLPreformatted", "HTML Preformatted") || quoted_pre {
        let (accepted, _) = para.accepted_view_with_field_owners();
        let code = accepted.text();
        let code = code.strip_suffix('\n').unwrap_or(&code);
        let length = code
            .split(|character| character != '`')
            .map(str::len)
            .max()
            .unwrap_or(0);
        let fence = "`".repeat((length + 1).max(3));
        text = format!("{fence}\n{code}\n{fence}");
    }
    let inner_quotes = container_path
        .rsplit('l')
        .next()
        .unwrap_or("")
        .chars()
        .filter(|&c| c == 'q')
        .count()
        .saturating_sub(if !container_path.contains('l') {
            outer_quotes
        } else {
            0
        });
    let item_quotes =
        container_path.chars().filter(|&c| c == 'q').count() - outer_quotes - inner_quotes;
    let item_quote_prefix = "> ".repeat(item_quotes);
    if inner_quotes > 0 || ((role("Quote", "Quote") || quoted_pre) && container_path.is_empty()) {
        if let Some(level) = heading_level.take() {
            text = format!("{} {}", "#".repeat(level.min(6) as usize), text.trim());
        }
        let depth = para
            .properties
            .as_ref()
            .and_then(|properties| properties.ind_left)
            .map_or(1, |indent| (indent.0 / 360).clamp(1, 256) as usize);
        let prefix = "> ".repeat(if container_path.is_empty() {
            depth
        } else {
            inner_quotes
        });
        text = text
            .split('\n')
            .map(|line| format!("{prefix}{line}"))
            .collect::<Vec<_>>()
            .join("\n");
    }

    if let Some(level) = heading_level
        && (list_info.is_some()
            || para
                .properties
                .as_ref()
                .is_some_and(|p| p.num_id == Some(0)))
    {
        text = format!(
            "{} {}",
            "#".repeat(level.min(6) as usize),
            text.trim().replace(LINE_BREAK, "<br>")
        );
        heading_level = None;
    }
    let quote_prefix = "> ".repeat(outer_quotes);

    if para
        .properties
        .as_ref()
        .is_some_and(|properties| properties.num_id == Some(0))
        && let Some(level) = para
            .properties
            .as_ref()
            .and_then(|properties| properties.num_ilvl)
        && let Some(index) = list_levels
            .iter()
            .position(|&(current, _, _)| current == level)
    {
        let indent = list_levels[index].1;
        let blank_prefix = if quote_prefix.is_empty() && inner_quotes > 0 && compound_quote {
            format!("{}> ", " ".repeat(indent))
        } else {
            quote_prefix.clone()
        };
        list_levels.truncate(index + 1);
        if !out.ends_with("\n\n") && !out.ends_with("> \n") {
            out.push_str(&blank_prefix);
            out.push('\n');
        }
        for line in text.lines() {
            out.push_str(&format!("{quote_prefix}{}{line}\n", " ".repeat(indent)));
        }
        out.push_str(&blank_prefix);
        out.push('\n');
        return;
    }

    // Skip empty paragraphs (but still emit blank line for spacing)
    if text.is_empty() && heading_level.is_none() && list_info.is_none() {
        out.push('\n');
        return;
    }

    if let Some(level) = heading_level {
        // Heading
        leave_quoted_list(out, list_levels, &quote_prefix);
        // An ATX heading is one line, so a line break is an HTML one.
        let hashes = "#".repeat(level.min(6) as usize);
        out.push_str(&quote_prefix);
        out.push_str(&hashes);
        out.push(' ');
        out.push_str(&text.trim().replace(LINE_BREAK, "<br>"));
        out.push_str("\n\n");
    } else if let Some((is_ordered, ilvl, num_id, start)) = list_info {
        if item_quotes > 0 && !list_levels.is_empty() && !out.ends_with("\n\n") {
            out.push('\n');
        }
        // A nested item is indented to the text column of its parent item,
        // which is three columns after `1. ` and two after `- `.
        let continued = list_levels
            .iter()
            .any(|&(level, _, id)| level == ilvl && id == num_id);
        while list_levels
            .last()
            .is_some_and(|&(level, _, _)| level >= ilvl)
        {
            list_levels.pop();
        }
        let indent = list_levels.last().map_or(0, |&(_, column, _)| column);
        let marker = if is_ordered {
            format!("{}. ", if continued { 1 } else { start })
        } else {
            "- ".to_owned()
        };
        list_levels.push((
            ilvl,
            indent + item_quote_prefix.len() + marker.len(),
            num_id,
        ));
        let mut lines = text.trim_matches(|c: char| c.is_ascii_whitespace()).lines();
        out.push_str(&format!(
            "{quote_prefix}{}{item_quote_prefix}{marker}{}\n",
            " ".repeat(indent),
            lines.next().unwrap_or("")
        ));
        for line in lines {
            out.push_str(&format!(
                "{quote_prefix}{}{line}\n",
                " ".repeat(indent + item_quote_prefix.len() + marker.len())
            ));
        }
    } else {
        // Retain original edge whitespace and terminal hard-break spelling.
        leave_quoted_list(out, list_levels, &quote_prefix);
        if outer_quotes > 0 {
            for line in text.split('\n') {
                out.push_str(&quote_prefix);
                out.push_str(line);
                out.push('\n');
            }
            out.push('\n');
        } else {
            out.push_str(&text);
            out.push_str("\n\n");
        }
    }
}

fn container_path<'a>(id: Option<&str>, styles: &'a CT_Styles) -> Option<&'a str> {
    let style = styles.get_by_id(id?)?;
    if !style.style_id.starts_with("RDOCXHTMLContainers") {
        return None;
    }
    let path = style
        .name
        .as_deref()?
        .strip_prefix("RDOCX HTML containers ")?
        .split('|')
        .next()?;
    (!path.is_empty() && path.len() <= 256 && path.bytes().all(|b| matches!(b, b'q' | b'l')))
        .then_some(path)
}

fn quote_owner<'a>(id: Option<&str>, styles: &'a CT_Styles) -> Option<&'a str> {
    let style = styles.get_by_id(id?)?;
    if !style.style_id.starts_with("RDOCXHTMLContainers") {
        return None;
    }
    let owner = style
        .name
        .as_deref()?
        .strip_prefix("RDOCX HTML containers ")?
        .split_once('|')?
        .1;
    (!owner.is_empty()
        && owner
            .split('.')
            .all(|id| !id.is_empty() && id.bytes().all(|b| b.is_ascii_digit())))
    .then_some(owner)
}

fn escape_block_starts(text: &str) -> String {
    text.split('\n')
        .map(|line| {
            let offset = line.bytes().take_while(|&b| b == b' ').count().min(3);
            let rest = &line[offset..];
            let escaped = if rest.starts_with(['#', '>', '+', '-', '=']) {
                Some(offset)
            } else {
                let digits = rest.bytes().take_while(u8::is_ascii_digit).count();
                (digits > 0
                    && digits <= 9
                    && rest
                        .as_bytes()
                        .get(digits)
                        .is_some_and(|b| matches!(b, b'.' | b')')))
                .then_some(offset + digits)
            };
            if let Some(at) = escaped {
                format!("{}\\{}", &line[..at], &line[at..])
            } else {
                line.to_owned()
            }
        })
        .collect::<Vec<_>>()
        .join("\n")
}

/// A hard line break inside a paragraph.
const LINE_BREAK: &str = "  \n";
/// A page break, written as a thematic break between blank lines. Without
/// the blank line before it, `---` turns the line above into a heading.
const PAGE_BREAK: &str = "\n\n---\n\n";

/// The run formatting Markdown can spell.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Mark {
    Strike,
    Bold,
    Italic,
}

impl Mark {
    /// Every mark, outermost first when several start and end together.
    /// Strikethrough goes innermost, next to the text, because cmark-gfm
    /// reads a `~` beside a `*` run as a letter rather than punctuation.
    const ALL: [Mark; 3] = [Mark::Bold, Mark::Italic, Mark::Strike];

    fn delimiter(self) -> &'static str {
        match self {
            Mark::Strike => "~~",
            Mark::Bold => "**",
            Mark::Italic => "*",
        }
    }

    fn tag(self) -> &'static str {
        match self {
            Mark::Strike => "del",
            Mark::Bold => "strong",
            Mark::Italic => "em",
        }
    }
}

/// The set of marks a stretch of text carries.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
struct Marks {
    strike: bool,
    bold: bool,
    italic: bool,
}

impl Marks {
    fn from_rpr(rpr: Option<&rdocx_oxml::properties::CT_RPr>) -> Self {
        Marks {
            strike: rpr.is_some_and(|r| r.strike == Some(true) || r.dstrike == Some(true)),
            bold: rpr.is_some_and(|r| r.bold == Some(true)),
            italic: rpr.is_some_and(|r| r.italic == Some(true)),
        }
    }

    fn overlay(self, properties: Option<&rdocx_oxml::properties::CT_RPr>) -> Self {
        let Some(properties) = properties else {
            return self;
        };
        Marks {
            strike: properties
                .strike
                .or(properties.dstrike)
                .unwrap_or(self.strike),
            bold: properties.bold.unwrap_or(self.bold),
            italic: properties.italic.unwrap_or(self.italic),
        }
    }

    fn has(self, mark: Mark) -> bool {
        match mark {
            Mark::Strike => self.strike,
            Mark::Bold => self.bold,
            Mark::Italic => self.italic,
        }
    }
}

/// Collect all text from the accepted view of a paragraph, applying inline
/// formatting.
fn collect_paragraph_text(
    para: &CT_P,
    styles: &CT_Styles,
    hyperlink_urls: &HashMap<String, String>,
    images: &HashMap<String, ImageData>,
    boundaries: &[usize],
) -> String {
    let mut out = String::new();

    // The runs the text readers read, those of content controls, tracked
    // insertions, smart tags and custom XML included, deleted ones left out.
    let (view, projected_boundaries) = para.accepted_view_with_field_owners();
    let para = &view;

    // Build hyperlink map
    let mut hyperlink_map: HashMap<usize, &str> = HashMap::new();
    for hl in &para.hyperlinks {
        // Markdown renderers turn `[text](javascript:...)` into a live link too,
        // so the same scheme allowlist applies here.
        if let Some(rel_id) = &hl.rel_id
            && let Some(url) = hyperlink_urls
                .get(rel_id)
                .map(String::as_str)
                .and_then(crate::sanitize::safe_url)
        {
            for i in hl.run_start..hl.run_end {
                hyperlink_map.insert(i, url);
            }
        }
    }

    // Consecutive runs in the same link are written as one link, and the
    // formatting inside and outside each link is rendered as a whole.
    let mut current_link: Option<&str> = None;
    let mut pieces: Vec<(String, Marks)> = Vec::new();

    for (run_idx, run) in para.runs.iter().enumerate() {
        if boundaries.contains(&run_idx) || projected_boundaries.contains(&run_idx) {
            pieces.push((String::new(), Marks::default()));
        }
        let in_link = hyperlink_map.get(&run_idx).copied();
        if in_link != current_link {
            push_inline(&mut out, &pieces, current_link);
            pieces.clear();
            current_link = in_link;
        }
        collect_run_pieces(run, styles, images, &mut pieces);
    }
    push_inline(&mut out, &pieces, current_link);

    out
}

/// Append formatted pieces to `out`, as the text of a link when `link` is set.
fn push_inline(out: &mut String, pieces: &[(String, Marks)], link: Option<&str>) {
    let mut start = 0;
    for (index, (text, _)) in pieces.iter().enumerate() {
        if text == PAGE_BREAK {
            push_inline_segment(out, &pieces[start..index], link);
            out.push_str(PAGE_BREAK);
            start = index + 1;
        }
    }
    push_inline_segment(out, &pieces[start..], link);
}

fn push_inline_segment(out: &mut String, pieces: &[(String, Marks)], link: Option<&str>) {
    let text = escape_block_starts(&render_inline(pieces));
    if text.is_empty() {
        return;
    }
    match link {
        Some(url) => {
            // `![` would start an image.
            if out.ends_with('!') {
                out.pop();
                out.push_str("\\!");
            }
            out.push('[');
            out.push_str(&text);
            out.push_str("](");
            push_link_destination(out, url);
            out.push(')');
        }
        None => out.push_str(&text),
    }
}

/// Write a URL as a link destination. A space or control character ends a
/// destination and an unbalanced parenthesis ends the link, so those are
/// percent-encoded or escaped.
fn push_link_destination(out: &mut String, url: &str) {
    for character in url.chars() {
        match character {
            ' ' => out.push_str("%20"),
            character if character.is_ascii_control() => {
                out.push_str(&format!("%{:02X}", u32::from(character)));
            }
            '(' | ')' | '<' | '>' | '\\' => {
                out.push('\\');
                out.push(character);
            }
            character => out.push(character),
        }
    }
}

/// Append `text` with `marks`, merged into the previous piece when its marks
/// are the same.
fn push_piece(pieces: &mut Vec<(String, Marks)>, text: &str, marks: Marks) {
    if text.is_empty() {
        return;
    }
    match pieces.last_mut() {
        Some((last, last_marks)) if !last.is_empty() && *last_marks == marks => last.push_str(text),
        _ => pieces.push((text.to_owned(), marks)),
    }
}

/// Append a character of document text, escaping the punctuation that would
/// otherwise start emphasis, a code span, a link, raw HTML or strikethrough.
fn push_escaped(text: &mut String, character: char) {
    if matches!(character, '\\' | '`' | '*' | '_' | '[' | ']' | '<' | '~') {
        text.push('\\');
    }
    text.push(character);
}

/// Collect the text of a single run as formatted pieces.
fn collect_run_pieces(
    run: &CT_R,
    styles: &CT_Styles,
    images: &HashMap<String, ImageData>,
    pieces: &mut Vec<(String, Marks)>,
) {
    let marks = Marks::from_rpr(run.properties.as_ref());
    let mut raw = String::new();
    let code = run
        .properties
        .as_ref()
        .and_then(|properties| properties.style_id.as_deref())
        .is_some_and(|id| {
            id == "HTMLCode"
                || styles
                    .get_by_id(id)
                    .and_then(|style| style.name.as_deref())
                    .is_some_and(|name| name.eq_ignore_ascii_case("HTML Code"))
        });
    if code {
        for content in &run.content {
            match content {
                RunContent::Text(text) | RunContent::DeletedText(text) => raw.push_str(&text.text),
                RunContent::Tab => raw.push('\t'),
                _ => {}
            }
        }
        let delimiter = "`".repeat(
            raw.split(|character| character != '`')
                .map(str::len)
                .max()
                .unwrap_or(0)
                + 1,
        );
        let pad = raw.starts_with('`')
            || raw.ends_with('`')
            || (raw.starts_with(' ')
                && raw.ends_with(' ')
                && !raw.chars().all(|character| character == ' '));
        let padding = if pad { " " } else { "" };
        push_piece(
            pieces,
            &format!("{delimiter}{padding}{raw}{padding}{delimiter}"),
            marks,
        );
        return;
    }

    for content in &run.content {
        match content {
            RunContent::Text(t) | RunContent::DeletedText(t) => {
                for character in t.text.chars() {
                    push_escaped(&mut raw, character);
                }
            }
            RunContent::Drawing(drawing) => {
                let picture = drawing
                    .inline
                    .as_ref()
                    .map(|inline| (inline.embed_id.as_str(), inline.description.as_deref()))
                    .or_else(|| {
                        drawing
                            .anchor
                            .as_ref()
                            .map(|anchor| (anchor.embed_id.as_str(), anchor.description.as_deref()))
                    });
                if let Some((relationship, description)) = picture
                    && let Some(image) = images.get(relationship)
                {
                    let mut alt = String::new();
                    for character in description.unwrap_or("").chars() {
                        push_escaped(&mut alt, character);
                    }
                    raw.push_str(&format!(
                        "![{alt}](data:{};base64,{})",
                        image.content_type,
                        crate::emitter::base64_encode(&image.data)
                    ));
                }
            }
            RunContent::Tab => raw.push('\t'),
            RunContent::Break(bt) => match bt {
                BreakType::Line | BreakType::Column => {
                    push_piece(pieces, &std::mem::take(&mut raw), marks);
                    pieces.push((String::new(), Marks::default()));
                    pieces.push((LINE_BREAK.to_owned(), Marks::default()));
                    pieces.push((String::new(), Marks::default()));
                }
                BreakType::Page => {
                    push_piece(pieces, &std::mem::take(&mut raw), marks);
                    pieces.push((String::new(), Marks::default()));
                    pieces.push((PAGE_BREAK.to_owned(), Marks::default()));
                    pieces.push((String::new(), Marks::default()));
                }
            },
            RunContent::Field(field) => {
                push_piece(pieces, &std::mem::take(&mut raw), marks);
                pieces.push((String::new(), Marks::default()));
                for (text, properties) in field.cached_display_segments() {
                    let field_marks = marks.overlay(properties);
                    let mut display = String::new();
                    for character in text.chars() {
                        match character {
                            '\t' => display.push('\t'),
                            '\n' | '\u{000b}' => {
                                push_piece(pieces, &std::mem::take(&mut display), field_marks);
                                pieces.push((String::new(), Marks::default()));
                                pieces.push((LINE_BREAK.to_owned(), Marks::default()));
                                pieces.push((String::new(), Marks::default()));
                            }
                            '\u{000c}' => {
                                push_piece(pieces, &std::mem::take(&mut display), field_marks);
                                pieces.push((String::new(), Marks::default()));
                                pieces.push((PAGE_BREAK.to_owned(), Marks::default()));
                                pieces.push((String::new(), Marks::default()));
                            }
                            character => push_escaped(&mut display, character),
                        }
                    }
                    push_piece(pieces, &display, field_marks);
                }
                pieces.push((String::new(), Marks::default()));
            }
            RunContent::SpecialCharacter(character) => match character {
                SpecialCharacter::CarriageReturn => {
                    push_piece(pieces, &std::mem::take(&mut raw), marks);
                    pieces.push((String::new(), Marks::default()));
                    pieces.push((LINE_BREAK.to_owned(), Marks::default()));
                    pieces.push((String::new(), Marks::default()));
                }
                SpecialCharacter::NoBreakHyphen => raw.push('\u{2011}'),
                SpecialCharacter::SoftHyphen => raw.push('\u{00ad}'),
                SpecialCharacter::PositionalTab { .. } => raw.push('\t'),
            },
            // A symbol is font-encoded rather than Unicode, so there is no
            // portable Markdown spelling for it.
            RunContent::Symbol { .. }
            | RunContent::FootnoteRef { .. }
            | RunContent::EndnoteRef { .. }
            | RunContent::CommentReference { .. } => {}
        }
    }

    push_piece(pieces, &raw, marks);
}

/// One emphasis, strikethrough or bold span between two boundaries.
struct Span {
    mark: Mark,
    /// The boundaries it opens and closes at.
    open: usize,
    close: usize,
    /// Written as an HTML element because CommonMark would not read
    /// delimiters at one of its boundaries as intended.
    html: bool,
}

/// Render formatted pieces as CommonMark inline text.
///
/// Whitespace at the edge of a piece is written outside the delimiters,
/// because a delimiter next to whitespace on its inner side does not open or
/// close. A piece of whitespace alone takes no delimiters. Marks are opened
/// and closed only where they change, so neighbouring pieces with the same
/// marks share one pair of delimiters, and spans nest. A span whose
/// delimiters CommonMark would read as literal text at a boundary, next to
/// punctuation inside a word for example, is written as an HTML element.
fn render_inline(pieces: &[(String, Marks)]) -> String {
    // An empty piece is a physical field/break boundary, not empty source text.
    // Render each side independently so equal marks cannot cross that owner.
    pieces
        .split(|(text, _)| text.is_empty())
        .map(render_inline_segment)
        .collect()
}

fn render_inline_segment(pieces: &[(String, Marks)]) -> String {
    // The text between whitespace, and the whitespace before, between and
    // after it. `gaps` has one more entry than `cores`.
    let mut cores: Vec<(String, Marks)> = Vec::new();
    let mut gaps = vec![String::new()];
    for (text, marks) in pieces {
        let core = text.trim();
        if core.is_empty() {
            gaps.last_mut().expect("gaps is never empty").push_str(text);
            continue;
        }
        let start = text.len() - text.trim_start().len();
        let gap = gaps.last_mut().expect("gaps is never empty");
        gap.push_str(&text[..start]);
        match cores.last_mut() {
            Some((last, last_marks)) if gap.is_empty() && last_marks == marks => {
                last.push_str(core);
            }
            _ => {
                cores.push((core.to_owned(), *marks));
                gaps.push(String::new());
            }
        }
        gaps.last_mut()
            .expect("gaps is never empty")
            .push_str(&text[start + core.len()..]);
    }

    // Boundary `b` lies before core `b`, and the last boundary after every
    // core. Each lists the spans it closes, innermost first, those that stay
    // open across it, then those it opens, outermost first. Each mark is
    // open at most once, so none of these lists is longer than three.
    let boundaries = cores.len() + 1;
    let mut spans: Vec<Span> = Vec::new();
    let mut closes: Vec<Vec<usize>> = vec![Vec::new(); boundaries];
    let mut encloses: Vec<Vec<usize>> = vec![Vec::new(); boundaries];
    let mut opens: Vec<Vec<usize>> = vec![Vec::new(); boundaries];
    let mut stack: Vec<usize> = Vec::new();
    for boundary in 0..boundaries {
        let wanted = cores
            .get(boundary)
            .map_or(Marks::default(), |(_, marks)| *marks);
        // Spans nest, so a mark that ends closes every span opened after it.
        if let Some(depth) = stack.iter().position(|&span| !wanted.has(spans[span].mark)) {
            for span in stack.drain(depth..).rev() {
                spans[span].close = boundary;
                closes[boundary].push(span);
            }
        }
        encloses[boundary].clone_from(&stack);
        let mut starting: Vec<Mark> = Mark::ALL
            .into_iter()
            .filter(|&mark| wanted.has(mark) && !stack.iter().any(|&span| spans[span].mark == mark))
            .collect();
        // The longest lasting mark goes outermost, so it need not be closed
        // and reopened when a shorter one ends.
        starting.sort_by_key(|&mark| {
            Reverse(
                cores[boundary..]
                    .iter()
                    .take_while(|(_, marks)| marks.has(mark))
                    .count(),
            )
        });
        for mark in starting {
            stack.push(spans.len());
            opens[boundary].push(spans.len());
            spans.push(Span {
                mark,
                open: boundary,
                close: boundaries,
                html: false,
            });
        }
    }

    // A span becomes HTML when one of its boundaries fails, which changes
    // the neighbours of the delimiters at its other boundary, so that one is
    // checked again. Fewer delimiters never make another boundary fail.
    let mut pending: Vec<usize> = (0..boundaries).collect();
    while let Some(boundary) = pending.pop() {
        let before = boundary
            .checked_sub(1)
            .and_then(|core| cores[core].0.chars().last());
        let after = cores
            .get(boundary)
            .and_then(|(core, _)| core.chars().next());
        let closing: Vec<(&Span, bool)> = closes[boundary]
            .iter()
            .map(|&span| (&spans[span], false))
            .collect();
        let opening: Vec<(&Span, bool)> = opens[boundary]
            .iter()
            .map(|&span| (&spans[span], true))
            .collect();
        let enclosed = encloses[boundary]
            .iter()
            .any(|&span| !spans[span].html && spans[span].mark != Mark::Strike);
        let gap = &gaps[boundary];
        let holds = if gap.is_empty() {
            delimiters_hold(before, &[closing, opening].concat(), after, enclosed)
        } else {
            delimiters_hold(before, &closing, gap.chars().next(), enclosed)
                && delimiters_hold(gap.chars().last(), &opening, after, enclosed)
        };
        if holds {
            continue;
        }
        for &span in closes[boundary].iter().chain(&opens[boundary]) {
            if !spans[span].html {
                spans[span].html = true;
                pending.push(spans[span].open);
                pending.push(spans[span].close);
            }
        }
    }

    let mut out = String::new();
    for boundary in 0..boundaries {
        for &span in &closes[boundary] {
            let Span { mark, html, .. } = spans[span];
            if html {
                out.push_str(&format!("</{}>", mark.tag()));
            } else {
                out.push_str(mark.delimiter());
            }
        }
        out.push_str(&gaps[boundary]);
        for &span in &opens[boundary] {
            let Span { mark, html, .. } = spans[span];
            if html {
                out.push_str(&format!("<{}>", mark.tag()));
            } else {
                out.push_str(mark.delimiter());
            }
        }
        if let Some((core, _)) = cores.get(boundary) {
            out.push_str(core);
        }
    }
    out
}

/// How a delimiter run sees the character on one side of it.
#[derive(Clone, Copy, PartialEq, Eq)]
enum Flank {
    Whitespace,
    Punctuation,
    Other,
}

impl Flank {
    /// The readings of a character. `None` is the edge of the inline text,
    /// which counts as whitespace. cmark-gfm does not count `~` as
    /// punctuation, unlike the CommonMark spec, and CommonMark versions
    /// differ on which other Unicode symbols are, so those are read both ways.
    fn of(character: Option<char>) -> &'static [Flank] {
        match character {
            None => &[Flank::Whitespace],
            Some(c) if c.is_whitespace() => &[Flank::Whitespace],
            Some(c) if c.is_alphanumeric() => &[Flank::Other],
            Some(c) if c.is_ascii_punctuation() && c != '~' => &[Flank::Punctuation],
            Some(_) => &[Flank::Punctuation, Flank::Other],
        }
    }
}

/// Whether CommonMark reads each delimiter in `tokens`, written between the
/// characters `before` and `after`, as the opener or closer it is meant to
/// be. `enclosed` says whether a `*` span stays open across the boundary.
///
/// Neighbouring delimiters of the same character form one delimiter run. A
/// run that closes must be right-flanking and one that opens left-flanking.
/// A run that is both could close an enclosing `*` span instead of opening,
/// and past three characters the rule of three can pair it differently. A
/// `~` run must be exactly `~~`. Every reading of a neighbouring character
/// must hold.
fn delimiters_hold(
    before: Option<char>,
    tokens: &[(&Span, bool)],
    after: Option<char>,
    enclosed: bool,
) -> bool {
    let character = |span: &Span| if span.mark == Mark::Strike { '~' } else { '*' };
    // A neighbouring token is an HTML tag or a delimiter of the other character.
    let neighbour = |span: &Span| Flank::of(Some(if span.html { '<' } else { character(span) }));

    let mut start = 0;
    while start < tokens.len() {
        let (span, _) = tokens[start];
        if span.html {
            start += 1;
            continue;
        }
        let delimiter = character(span);
        let mut end = start;
        while end < tokens.len() && !tokens[end].0.html && character(tokens[end].0) == delimiter {
            end += 1;
        }
        let run = &tokens[start..end];
        let previous = start
            .checked_sub(1)
            .map_or(Flank::of(before), |index| neighbour(tokens[index].0));
        let next = tokens
            .get(end)
            .map_or(Flank::of(after), |(span, _)| neighbour(span));
        let length: usize = run
            .iter()
            .map(|(span, _)| span.mark.delimiter().len())
            .sum();
        let closes = run.iter().any(|&(_, opens)| !opens);
        let opens = run.iter().any(|&(_, opens)| opens);
        for &previous in previous {
            for &next in next {
                let left = next != Flank::Whitespace
                    && (next != Flank::Punctuation || previous != Flank::Other);
                let right = previous != Flank::Whitespace
                    && (previous != Flank::Punctuation || next != Flank::Other);
                if (closes && !right)
                    || (opens && !left)
                    || (opens && right && enclosed && delimiter == '*')
                    || (left && right && length > 3)
                    || (delimiter == '~' && length != 2)
                {
                    return false;
                }
            }
        }
        start = end;
    }
    true
}

/// Emit a table as a GFM pipe table.
fn emit_table(
    out: &mut String,
    tbl: &CT_Tbl,
    styles: &CT_Styles,
    hyperlink_urls: &HashMap<String, String>,
    images: &HashMap<String, ImageData>,
) {
    if tbl.rows().is_empty() {
        return;
    }

    // Collect all rows as cell text vectors
    let mut rows: Vec<Vec<String>> = Vec::new();
    let mut max_cols = 0;

    // Rows and cells that content controls wrap are written in place.
    for row in tbl
        .rows()
        .into_iter()
        .filter(|row| !row.accepted_view_removes())
    {
        let mut cells: Vec<String> = Vec::new();
        for cell in row.cells() {
            let text = collect_cell_text(cell, styles, hyperlink_urls, images);
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
    let header_cells = tbl
        .rows()
        .into_iter()
        .find(|row| !row.accepted_view_removes())
        .map(|row| row.cells())
        .unwrap_or_default();
    for column in 0..max_cols {
        let alignment = header_cells
            .get(column)
            .and_then(|cell| cell.paragraphs().first().copied())
            .and_then(|paragraph| paragraph.properties.as_ref())
            .and_then(|properties| properties.jc);
        out.push_str(match alignment {
            Some(rdocx_oxml::shared::ST_Jc::Center) => " :---: |",
            Some(rdocx_oxml::shared::ST_Jc::Right | rdocx_oxml::shared::ST_Jc::End) => " ---: |",
            Some(rdocx_oxml::shared::ST_Jc::Left | rdocx_oxml::shared::ST_Jc::Start) => " :--- |",
            _ => " --- |",
        });
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
    styles: &CT_Styles,
    hyperlink_urls: &HashMap<String, String>,
    images: &HashMap<String, ImageData>,
) -> String {
    let mut parts = Vec::new();

    for block in cell_blocks(cell) {
        match block {
            Block::Paragraph(p, boundaries) => {
                // A pipe table row is one line, so breaks inside a cell are
                // written as HTML line breaks. A pipe would end the cell, in
                // a link destination too, so every pipe is escaped, and GFM
                // removes the escape before it reads the cell's inlines.
                let text = collect_paragraph_text(&p, styles, hyperlink_urls, images, &boundaries)
                    .replace(PAGE_BREAK, " ");
                let trimmed = text.trim().replace(LINE_BREAK, "<br>").replace('|', "\\|");
                if !trimmed.is_empty() {
                    parts.push(trimmed);
                }
            }
            Block::Table(table) if table.accepted_view_removes() => {}
            Block::Table(_) => {
                parts.push("(nested table)".to_string());
            }
        }
    }

    parts.join(" ")
}

fn emit_chart_table(out: &mut String, chart: &crate::ChartTable) {
    let escaped = |text: &str| {
        let mut result = String::new();
        for ch in text.chars() {
            if ch == '|' {
                result.push('\\');
            }
            push_escaped(&mut result, ch);
        }
        result.replace('\n', "<br>").replace('\r', "")
    };
    if let Some(title) = chart.title.as_ref().filter(|title| !title.is_empty()) {
        out.push_str(&escape_block_starts(&escaped(title)));
        out.push_str("\n\n");
    }
    for (index, row) in chart.rows().iter().enumerate() {
        out.push('|');
        for cell in row {
            out.push_str(&format!(" {} |", escaped(cell)));
        }
        out.push('\n');
        if index == 0 {
            out.push('|');
            for _ in row {
                out.push_str(" --- |");
            }
            out.push('\n');
        }
    }
    out.push('\n');
}
