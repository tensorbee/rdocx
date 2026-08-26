//! Bounded HTML and CSS import into the native Word document model.

use std::cmp::Ordering;
use std::collections::HashSet;
use std::fs::File;
use std::io::Read;
use std::path::Path;

use rdocx_oxml::document::BodyContent;
use rdocx_oxml::table::{
    CT_Row, CT_Tbl, CT_TblGrid, CT_TblGridCol, CT_TblPr, CT_TblWidth, CT_Tc, VMerge,
};
use rdocx_oxml::text::CT_P;
use rdocx_oxml::units::Twips;
use scraper::{ElementRef, Html, Node, Selector};

use crate::paragraph::{Alignment, Paragraph};
use crate::run::Run;
use crate::table::Cell;
use crate::{Document, Error, Length, ListLevel, Result};

const MAX_INPUT_BYTES: usize = 64 * 1024 * 1024;

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct HtmlDiagnostic {
    pub location: String,
    pub property: Option<String>,
    pub message: String,
}

pub struct HtmlReadResult {
    pub document: Document,
    pub diagnostics: Vec<HtmlDiagnostic>,
}

#[derive(Debug, Clone, Copy)]
struct Limits {
    input_bytes: usize,
    retained_text: usize,
    depth: usize,
    nodes: usize,
    blocks: usize,
    runs: usize,
    rows: usize,
    columns: usize,
    cells: usize,
    diagnostics: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            input_bytes: MAX_INPUT_BYTES,
            retained_text: MAX_INPUT_BYTES,
            depth: 256,
            nodes: 100_000,
            blocks: 100_000,
            runs: 100_000,
            rows: 10_000,
            columns: 256,
            cells: 50_000,
            diagnostics: 10_000,
        }
    }
}

#[derive(Debug, Clone, Default)]
struct ComputedStyle {
    font: Option<String>,
    size: Option<f64>,
    bold: Option<bool>,
    italic: Option<bool>,
    underline: Option<bool>,
    strike: Option<bool>,
    color: Option<String>,
    background: Option<String>,
    alignment: Option<Alignment>,
    space_before: Option<Length>,
    space_after: Option<Length>,
    indent_left: Option<Length>,
    indent_right: Option<Length>,
    first_line_indent: Option<Length>,
    vertical: Option<VerticalText>,
}

impl ComputedStyle {
    fn inherited(&self) -> Self {
        Self {
            font: self.font.clone(),
            size: self.size,
            bold: self.bold,
            italic: self.italic,
            underline: self.underline,
            strike: self.strike,
            color: self.color.clone(),
            alignment: self.alignment,
            vertical: self.vertical,
            ..Self::default()
        }
    }
}

#[derive(Debug, Clone, Copy)]
enum VerticalText {
    Superscript,
    Subscript,
}

#[derive(Debug, Clone)]
enum StyleChange {
    Font(String),
    Size(f64),
    Bold(bool),
    Italic(bool),
    Underline(bool),
    Strike(bool),
    Color(String),
    Background(Option<String>),
    Alignment(Alignment),
    SpaceBefore(Length),
    SpaceAfter(Length),
    IndentLeft(Length),
    IndentRight(Length),
    FirstLineIndent(Length),
}

impl StyleChange {
    fn apply(&self, style: &mut ComputedStyle) {
        match self {
            Self::Font(value) => style.font = Some(value.clone()),
            Self::Size(value) => style.size = Some(*value),
            Self::Bold(value) => style.bold = Some(*value),
            Self::Italic(value) => style.italic = Some(*value),
            Self::Underline(value) => style.underline = Some(*value),
            Self::Strike(value) => style.strike = Some(*value),
            Self::Color(value) => style.color = Some(value.clone()),
            Self::Background(value) => style.background = value.clone(),
            Self::Alignment(value) => style.alignment = Some(*value),
            Self::SpaceBefore(value) => style.space_before = Some(*value),
            Self::SpaceAfter(value) => style.space_after = Some(*value),
            Self::IndentLeft(value) => style.indent_left = Some(*value),
            Self::IndentRight(value) => style.indent_right = Some(*value),
            Self::FirstLineIndent(value) => style.first_line_indent = Some(*value),
        }
    }
}

#[derive(Debug)]
struct CssRule {
    selector: Selector,
    specificity: (u32, u32, u32),
    order: usize,
    changes: Vec<StyleChange>,
}

#[derive(Debug, Clone)]
enum InlinePiece {
    Text(String, Box<ComputedStyle>, bool),
    Break,
}

#[derive(Debug)]
struct ParagraphModel {
    pieces: Vec<InlinePiece>,
    style: ComputedStyle,
    paragraph_style: Option<String>,
    numbering: Option<(u32, u32)>,
}

#[derive(Debug)]
struct TableCellModel {
    start: usize,
    span: usize,
    v_merge: Option<VMerge>,
    paragraphs: Vec<ParagraphModel>,
    header: bool,
}

#[derive(Debug)]
struct TableRowModel {
    cells: Vec<TableCellModel>,
    header: bool,
}

#[derive(Debug, Clone, Copy)]
struct ActiveSpan {
    remaining: usize,
    span: usize,
}

fn from_html_with_limits(html: &str, limits: Limits) -> Result<HtmlReadResult> {
    if html.len() > limits.input_bytes {
        return Err(html_error("input", "HTML input exceeds the 64 MiB limit"));
    }
    preflight_markup(html, limits)?;

    let is_document = contains_ascii_case_insensitive(html, b"<html")
        || contains_ascii_case_insensitive(html, b"<!doctype");
    let dom = if is_document {
        Html::parse_document(html)
    } else {
        Html::parse_fragment(html)
    };
    validate_dom(&dom, limits)?;

    let mut importer = Importer {
        dom: &dom,
        document: Document::new(),
        diagnostics: Vec::new(),
        diagnostic_keys: HashSet::new(),
        rules: Vec::new(),
        limits,
        blocks: 0,
        runs: 0,
        rows: 0,
        cells: 0,
    };
    importer.record_parser_repairs()?;
    importer.record_head_resources()?;
    importer.collect_styles()?;
    importer.project()?;
    importer.finish()
}

impl Document {
    pub fn from_html(html: &str) -> Result<HtmlReadResult> {
        from_html_with_limits(html, Limits::default())
    }

    pub fn open_html<P: AsRef<Path>>(path: P) -> Result<HtmlReadResult> {
        let mut file = File::open(path)?;
        let size = usize::try_from(file.metadata()?.len())
            .map_err(|_| html_error("input", "HTML input size is not representable"))?;
        if size > MAX_INPUT_BYTES {
            return Err(html_error("input", "HTML input exceeds the 64 MiB limit"));
        }
        let bytes = read_bounded(&mut file, size, MAX_INPUT_BYTES)?;
        let html = std::str::from_utf8(&bytes)
            .map_err(|_| html_error("input", "HTML input is not valid UTF-8"))?;
        Self::from_html(html)
    }
}

fn read_bounded<R: Read>(reader: &mut R, expected_size: usize, limit: usize) -> Result<Vec<u8>> {
    let read_limit = u64::try_from(limit)
        .ok()
        .and_then(|limit| limit.checked_add(1))
        .ok_or_else(|| html_error("input", "HTML input limit is not representable"))?;
    let mut bytes = Vec::with_capacity(expected_size.min(limit));
    reader.take(read_limit).read_to_end(&mut bytes)?;
    if bytes.len() > limit {
        return Err(html_error("input", "HTML input exceeds the 64 MiB limit"));
    }
    Ok(bytes)
}

fn html_error(location: impl Into<String>, message: impl Into<String>) -> Error {
    Error::Html {
        location: location.into(),
        message: message.into(),
    }
}

fn validate_dom(dom: &Html, limits: Limits) -> Result<()> {
    let mut nodes = 0_usize;
    let mut retained_text = 0_usize;
    for node in dom.tree.nodes() {
        nodes = nodes
            .checked_add(1)
            .ok_or_else(|| html_error("html", "DOM node count overflowed"))?;
        if nodes > limits.nodes {
            return Err(html_error("html", "HTML DOM exceeds the node limit"));
        }
        let mut depth = 0_usize;
        let mut ancestor = Some(node);
        while let Some(current) = ancestor {
            depth = depth
                .checked_add(1)
                .ok_or_else(|| html_error("html", "DOM depth overflowed"))?;
            if depth > limits.depth {
                return Err(html_error("html", "HTML DOM exceeds the depth limit"));
            }
            ancestor = current.parent();
        }
        if let Node::Text(text) = node.value() {
            retained_text = retained_text
                .checked_add(text.len())
                .ok_or_else(|| html_error("html", "retained text length overflowed"))?;
            if retained_text > limits.retained_text {
                return Err(html_error("html", "HTML retained text exceeds the limit"));
            }
        }
    }
    Ok(())
}

fn preflight_markup(html: &str, limits: Limits) -> Result<()> {
    let bytes = html.as_bytes();
    let mut estimated_nodes = 4_usize;
    let mut position = 0_usize;
    while position < bytes.len() {
        let Some(relative_open) = bytes[position..].iter().position(|byte| *byte == b'<') else {
            if bytes[position..]
                .iter()
                .any(|byte| !byte.is_ascii_whitespace())
            {
                estimated_nodes = estimated_nodes
                    .checked_add(1)
                    .ok_or_else(|| html_error("html", "DOM node estimate overflowed"))?;
                if estimated_nodes > limits.nodes {
                    return Err(html_error("html", "HTML DOM exceeds the node limit"));
                }
            }
            break;
        };
        let open = position + relative_open;
        if bytes[position..open]
            .iter()
            .any(|byte| !byte.is_ascii_whitespace())
        {
            estimated_nodes = estimated_nodes
                .checked_add(1)
                .ok_or_else(|| html_error("html", "DOM node estimate overflowed"))?;
        }
        if bytes.get(open + 1).is_some_and(|byte| *byte != b'/') {
            estimated_nodes = estimated_nodes
                .checked_add(1)
                .ok_or_else(|| html_error("html", "DOM node estimate overflowed"))?;
        }
        if estimated_nodes > limits.nodes {
            return Err(html_error("html", "HTML DOM exceeds the node limit"));
        }
        if bytes[open + 1..].starts_with(b"!--") {
            position = bytes[open + 4..]
                .windows(3)
                .position(|window| window == b"-->")
                .map_or(bytes.len(), |relative_close| open + relative_close + 7);
            continue;
        }
        position = bytes[open + 1..]
            .iter()
            .position(|byte| *byte == b'>')
            .map_or(bytes.len(), |relative_close| open + relative_close + 2);
    }
    Ok(())
}

fn contains_ascii_case_insensitive(haystack: &str, needle: &[u8]) -> bool {
    haystack
        .as_bytes()
        .windows(needle.len())
        .any(|window| window.eq_ignore_ascii_case(needle))
}

struct Importer<'a> {
    dom: &'a Html,
    document: Document,
    diagnostics: Vec<HtmlDiagnostic>,
    diagnostic_keys: HashSet<HtmlDiagnostic>,
    rules: Vec<CssRule>,
    limits: Limits,
    blocks: usize,
    runs: usize,
    rows: usize,
    cells: usize,
}

impl Importer<'_> {
    fn record_parser_repairs(&mut self) -> Result<()> {
        for error in &self.dom.errors {
            self.diagnostic("html", None, format!("HTML parser repair: {error}"))?;
        }
        Ok(())
    }

    fn record_head_resources(&mut self) -> Result<()> {
        let selector = Selector::parse("head link, head script").expect("static selector");
        let resources: Vec<_> = self.dom.select(&selector).collect();
        for element in resources {
            match element.value().name() {
                "link"
                    if element.attr("rel").is_some_and(|value| {
                        value
                            .split_whitespace()
                            .any(|item| item.eq_ignore_ascii_case("stylesheet"))
                    }) =>
                {
                    self.diagnostic(
                        &element_path(element),
                        None,
                        "dropped external HTML stylesheet".to_string(),
                    )?;
                }
                "script" => {
                    self.diagnostic(
                        &element_path(element),
                        None,
                        "dropped HTML script".to_string(),
                    )?;
                }
                _ => {}
            }
        }
        Ok(())
    }

    fn collect_styles(&mut self) -> Result<()> {
        let selector = Selector::parse("style").expect("static selector");
        let styles: Vec<_> = self
            .dom
            .select(&selector)
            .map(|element| {
                let css = element.text().collect::<String>();
                (element, css)
            })
            .collect();
        let mut order = 0_usize;
        for (element, css) in styles {
            let location = element_path(element);
            let css = strip_css_comments(&css);
            let mut remainder = css.as_str();
            while let Some(open) = remainder.find('{') {
                let selector_text = remainder[..open].trim();
                let after_open = &remainder[open + 1..];
                let Some(close) = after_open.find('}') else {
                    self.diagnostic(
                        &location,
                        None,
                        "unsupported unterminated CSS rule".to_string(),
                    )?;
                    break;
                };
                let declarations = &after_open[..close];
                remainder = &after_open[close + 1..];
                if selector_text.starts_with('@') {
                    self.diagnostic(
                        &location,
                        None,
                        format!("unsupported CSS at-rule `{selector_text}`"),
                    )?;
                    continue;
                }
                for selector_part in selector_text.split(',').map(str::trim) {
                    if !supported_selector(selector_part) {
                        self.diagnostic(
                            &location,
                            None,
                            format!("unsupported CSS selector `{selector_part}`"),
                        )?;
                        continue;
                    }
                    let selector = match Selector::parse(selector_part) {
                        Ok(selector) => selector,
                        Err(_) => {
                            self.diagnostic(
                                &location,
                                None,
                                format!("invalid CSS selector `{selector_part}`"),
                            )?;
                            continue;
                        }
                    };
                    let changes = self.parse_declarations(declarations, &location)?;
                    if !changes.is_empty() {
                        if self.rules.len() >= self.limits.nodes {
                            return Err(html_error(&location, "HTML exceeds the CSS rule limit"));
                        }
                        self.rules.push(CssRule {
                            selector,
                            specificity: specificity(selector_part),
                            order,
                            changes,
                        });
                        order = order
                            .checked_add(1)
                            .ok_or_else(|| html_error(&location, "CSS source order overflowed"))?;
                    }
                }
            }
            if !remainder.trim().is_empty() {
                self.diagnostic(
                    &location,
                    None,
                    "unsupported CSS text outside a rule".to_string(),
                )?;
            }
        }
        Ok(())
    }

    fn parse_declarations(
        &mut self,
        declarations: &str,
        location: &str,
    ) -> Result<Vec<StyleChange>> {
        let mut changes = Vec::new();
        for declaration in declarations.split(';').map(str::trim) {
            if declaration.is_empty() {
                continue;
            }
            let Some((property, value)) = declaration.split_once(':') else {
                self.diagnostic(
                    location,
                    None,
                    format!("invalid CSS declaration `{declaration}`"),
                )?;
                continue;
            };
            let property = property.trim().to_ascii_lowercase();
            let value = value.trim();
            match parse_style_change(&property, value) {
                Ok(mut parsed) => changes.append(&mut parsed),
                Err(message) => {
                    self.diagnostic(location, Some(property), message)?;
                }
            }
        }
        Ok(changes)
    }

    fn project(&mut self) -> Result<()> {
        let root = self.dom.root_element();
        let body_selector = Selector::parse("body").expect("static selector");
        let container = if root.value().name() == "html" {
            root.select(&body_selector).next().unwrap_or(root)
        } else {
            root
        };
        let base = ComputedStyle::default();
        self.project_container(container, &base)
    }

    fn project_container(
        &mut self,
        container: ElementRef<'_>,
        inherited: &ComputedStyle,
    ) -> Result<()> {
        let container_style = self.computed_style(container, inherited)?;
        let mut inline_group = Vec::new();
        for child in container.children() {
            match child.value() {
                Node::Text(text) => {
                    inline_group.push(InlinePiece::Text(
                        text.to_string(),
                        Box::new(container_style.clone()),
                        false,
                    ));
                }
                Node::Element(_) => {
                    let element = ElementRef::wrap(child).expect("element checked");
                    if is_block(element.value().name()) {
                        self.flush_inline_group(&mut inline_group, &container_style)?;
                        self.project_block(element, &container_style)?;
                    } else if !matches!(element.value().name(), "head" | "style" | "title") {
                        self.collect_inline(element, &container_style, false, &mut inline_group)?;
                    }
                }
                _ => {}
            }
        }
        self.flush_inline_group(&mut inline_group, &container_style)
    }

    fn flush_inline_group(
        &mut self,
        pieces: &mut Vec<InlinePiece>,
        style: &ComputedStyle,
    ) -> Result<()> {
        if pieces.iter().any(inline_piece_is_visible) {
            let model = ParagraphModel {
                pieces: std::mem::take(pieces),
                style: style.clone(),
                paragraph_style: None,
                numbering: None,
            };
            let paragraph = self.build_paragraph(model)?;
            self.document
                .document
                .body
                .content
                .push(BodyContent::Paragraph(paragraph));
        } else {
            pieces.clear();
        }
        Ok(())
    }

    fn project_block(&mut self, element: ElementRef<'_>, inherited: &ComputedStyle) -> Result<()> {
        let name = element.value().name();
        match name {
            "ul" | "ol" => self.project_list(element, inherited, 0, None),
            "table" => self.project_table(element, inherited),
            "div" => self.project_container(element, inherited),
            "p" | "blockquote" | "pre" | "h1" | "h2" | "h3" | "h4" | "h5" | "h6" => {
                let style = self.computed_style(element, inherited)?;
                let mut pieces = Vec::new();
                let pre = name == "pre";
                self.collect_inline_children(element, &style, pre, &mut pieces)?;
                let paragraph_style = match name {
                    "h1" => Some("Heading1".to_string()),
                    "h2" => Some("Heading2".to_string()),
                    "h3" => Some("Heading3".to_string()),
                    "h4" => Some("Heading4".to_string()),
                    "h5" => Some("Heading5".to_string()),
                    "h6" => Some("Heading6".to_string()),
                    "blockquote" => Some("Quote".to_string()),
                    _ => None,
                };
                let model = ParagraphModel {
                    pieces,
                    style,
                    paragraph_style,
                    numbering: None,
                };
                let paragraph = self.build_paragraph(model)?;
                self.document
                    .document
                    .body
                    .content
                    .push(BodyContent::Paragraph(paragraph));
                Ok(())
            }
            _ => {
                self.diagnostic(
                    &element_path(element),
                    None,
                    format!("unsupported visible HTML element `{name}` retained as content"),
                )?;
                self.project_container(element, inherited)
            }
        }
    }

    fn project_list(
        &mut self,
        list: ElementRef<'_>,
        inherited: &ComputedStyle,
        level: u32,
        list_id: Option<u32>,
    ) -> Result<()> {
        for model in self.list_models(list, inherited, level, list_id)? {
            let paragraph = self.build_paragraph(model)?;
            self.document
                .document
                .body
                .content
                .push(BodyContent::Paragraph(paragraph));
        }
        Ok(())
    }

    fn list_models(
        &mut self,
        list: ElementRef<'_>,
        inherited: &ComputedStyle,
        level: u32,
        list_id: Option<u32>,
    ) -> Result<Vec<ParagraphModel>> {
        if level > 8 {
            return Err(html_error(
                element_path(list),
                "HTML list nesting exceeds Word's nine-level limit",
            ));
        }
        let ordered = list.value().name() == "ol";
        let list_id = list_id.unwrap_or_else(|| {
            let levels: Vec<_> = (0..9)
                .map(|_| {
                    if ordered {
                        ListLevel::decimal()
                    } else {
                        ListLevel::bullet()
                    }
                })
                .collect();
            self.document.add_list_definition(&levels)
        });
        let mut spec = if ordered {
            ListLevel::decimal()
        } else {
            ListLevel::bullet()
        };
        if ordered {
            if let Some(start) = list.attr("start") {
                match start.parse::<u32>() {
                    Ok(start) if start > 0 => spec = spec.start(start),
                    _ => {
                        self.diagnostic(
                            &element_path(list),
                            Some("start".to_string()),
                            format!("unsupported ordered-list start value `{start}`"),
                        )?;
                    }
                }
            }
            if list.attr("reversed").is_some() {
                self.diagnostic(
                    &element_path(list),
                    Some("reversed".to_string()),
                    "unsupported reversed ordered list".to_string(),
                )?;
            }
        }
        if !self.document.set_list_level(list_id, level, spec) {
            return Err(html_error(
                element_path(list),
                "could not define list level",
            ));
        }
        let style = self.computed_style(list, inherited)?;
        let mut models = Vec::new();
        for item in list
            .child_elements()
            .filter(|child| child.value().name() == "li")
        {
            let item_style = self.computed_style(item, &style)?;
            let mut pieces = Vec::new();
            self.collect_inline_children(item, &item_style, false, &mut pieces)?;
            models.push(ParagraphModel {
                pieces,
                style: item_style.clone(),
                paragraph_style: None,
                numbering: Some((list_id, level)),
            });
            for nested in item
                .child_elements()
                .filter(|child| matches!(child.value().name(), "ul" | "ol"))
            {
                models.extend(self.list_models(nested, &item_style, level + 1, Some(list_id))?);
            }
        }
        Ok(models)
    }

    fn project_table(&mut self, table: ElementRef<'_>, inherited: &ComputedStyle) -> Result<()> {
        let style = self.computed_style(table, inherited)?;
        for caption in table
            .child_elements()
            .filter(|child| child.value().name() == "caption")
        {
            self.diagnostic(
                &element_path(caption),
                None,
                "unsupported HTML table caption retained before the table".to_string(),
            )?;
            let caption_style = self.computed_style(caption, &style)?;
            let mut pieces = Vec::new();
            self.collect_inline_children(caption, &caption_style, false, &mut pieces)?;
            let paragraph = self.build_paragraph(ParagraphModel {
                pieces,
                style: caption_style,
                paragraph_style: None,
                numbering: None,
            })?;
            self.document
                .document
                .body
                .content
                .push(BodyContent::Paragraph(paragraph));
        }
        let row_selector = Selector::parse("tr").expect("static selector");
        let rows: Vec<_> = table
            .select(&row_selector)
            .filter(|row| nearest_ancestor_named(*row, "table") == Some(table))
            .collect();
        self.rows = self
            .rows
            .checked_add(rows.len())
            .ok_or_else(|| html_error(element_path(table), "table row count overflowed"))?;
        if self.rows > self.limits.rows {
            return Err(html_error(
                element_path(table),
                "HTML tables exceed the row limit",
            ));
        }

        let mut models = Vec::with_capacity(rows.len());
        let mut active: Vec<Option<ActiveSpan>> = Vec::new();
        let mut column_count = 0_usize;
        for row in rows {
            let previous = std::mem::take(&mut active);
            let mut occupied = vec![false; previous.len()];
            let mut cells = Vec::new();
            for (start, span) in previous.iter().enumerate() {
                if let Some(span) = span {
                    let end = start
                        .checked_add(span.span)
                        .ok_or_else(|| html_error(element_path(row), "table span overflowed"))?;
                    if occupied.len() < end {
                        occupied.resize(end, false);
                    }
                    occupied[start..end].fill(true);
                    cells.push(TableCellModel {
                        start,
                        span: span.span,
                        v_merge: Some(VMerge::Continue),
                        paragraphs: Vec::new(),
                        header: false,
                    });
                    if span.remaining > 1 {
                        if active.len() <= start {
                            active.resize(start + 1, None);
                        }
                        active[start] = Some(ActiveSpan {
                            remaining: span.remaining - 1,
                            span: span.span,
                        });
                    }
                }
            }

            let mut cursor = 0_usize;
            for cell in row
                .child_elements()
                .filter(|cell| matches!(cell.value().name(), "td" | "th"))
            {
                let colspan =
                    parse_span(cell.attr("colspan"), "colspan", cell, self.limits.columns)?;
                let rowspan = parse_span(cell.attr("rowspan"), "rowspan", cell, self.limits.rows)?;
                loop {
                    while occupied.get(cursor).copied().unwrap_or(false) {
                        cursor = cursor.checked_add(1).ok_or_else(|| {
                            html_error(element_path(cell), "table column overflowed")
                        })?;
                    }
                    let end = cursor.checked_add(colspan).ok_or_else(|| {
                        html_error(element_path(cell), "table column span overflowed")
                    })?;
                    if end > self.limits.columns {
                        return Err(html_error(
                            element_path(cell),
                            "HTML table exceeds the column limit",
                        ));
                    }
                    if occupied.len() < end {
                        occupied.resize(end, false);
                    }
                    if occupied[cursor..end].iter().any(|value| *value) {
                        cursor = cursor.checked_add(1).ok_or_else(|| {
                            html_error(element_path(cell), "table column overflowed")
                        })?;
                        continue;
                    }
                    occupied[cursor..end].fill(true);
                    let cell_style = self.computed_style(cell, &style)?;
                    let paragraphs = self.cell_paragraphs(cell, &cell_style)?;
                    cells.push(TableCellModel {
                        start: cursor,
                        span: colspan,
                        v_merge: (rowspan > 1).then_some(VMerge::Restart),
                        paragraphs,
                        header: cell.value().name() == "th",
                    });
                    if rowspan > 1 {
                        if active.len() < end {
                            active.resize(end, None);
                        }
                        active[cursor] = Some(ActiveSpan {
                            remaining: rowspan - 1,
                            span: colspan,
                        });
                    }
                    cursor = end;
                    break;
                }
            }
            column_count = column_count.max(occupied.len());
            cells.sort_by_key(|cell| cell.start);
            self.cells = self
                .cells
                .checked_add(cells.len())
                .ok_or_else(|| html_error(element_path(row), "table cell count overflowed"))?;
            if self.cells > self.limits.cells {
                return Err(html_error(
                    element_path(table),
                    "HTML tables exceed the cell limit",
                ));
            }
            models.push(TableRowModel {
                header: cells.iter().any(|cell| cell.header),
                cells,
            });
        }
        if column_count == 0 {
            self.diagnostic(
                &element_path(table),
                None,
                "dropped empty HTML table".to_string(),
            )?;
            return Ok(());
        }
        if column_count > self.limits.columns {
            return Err(html_error(
                element_path(table),
                "HTML table exceeds the column limit",
            ));
        }

        let width = Twips(9360 / column_count.max(1) as i32);
        let mut output = CT_Tbl::new();
        output.properties = Some(CT_TblPr {
            width: Some(CT_TblWidth::dxa(width.0 * column_count as i32)),
            ..CT_TblPr::default()
        });
        output.grid = Some(CT_TblGrid {
            columns: (0..column_count).map(|_| CT_TblGridCol { width }).collect(),
            grid_change_xml: None,
        });
        for model in models {
            let mut row = CT_Row::new();
            if model.header {
                row.properties.get_or_insert_with(Default::default).header = Some(true);
            }
            for model_cell in model.cells {
                let mut cell = CT_Tc::new();
                cell.content.clear();
                {
                    let mut facade = Cell { inner: &mut cell };
                    if model_cell.span > 1 {
                        facade.set_grid_span(model_cell.span as u32);
                    }
                    match model_cell.v_merge {
                        Some(VMerge::Restart) => facade.set_v_merge_restart(),
                        Some(VMerge::Continue) => facade.set_v_merge_continue(),
                        None => {}
                    }
                    if model_cell.header {
                        facade.set_shading("D9EAF7");
                    }
                }
                for paragraph in model_cell.paragraphs {
                    cell.content.push(rdocx_oxml::table::CellContent::Paragraph(
                        self.build_paragraph(paragraph)?,
                    ));
                }
                if cell.content.is_empty() {
                    cell.content
                        .push(rdocx_oxml::table::CellContent::Paragraph(CT_P::new()));
                }
                row.cells.push(cell);
            }
            output.rows.push(row);
        }
        self.bump_blocks(&element_path(table))?;
        self.document
            .document
            .body
            .content
            .push(BodyContent::Table(output));
        Ok(())
    }

    fn cell_paragraphs(
        &mut self,
        cell: ElementRef<'_>,
        inherited: &ComputedStyle,
    ) -> Result<Vec<ParagraphModel>> {
        let mut paragraphs = Vec::new();
        let mut inline_group = Vec::new();
        for child in cell.children() {
            match child.value() {
                Node::Text(text) => inline_group.push(InlinePiece::Text(
                    text.to_string(),
                    Box::new(inherited.clone()),
                    false,
                )),
                Node::Element(_) => {
                    let element = ElementRef::wrap(child).expect("element checked");
                    if matches!(element.value().name(), "p" | "div" | "blockquote" | "pre") {
                        if inline_group.iter().any(inline_piece_is_visible) {
                            paragraphs.push(ParagraphModel {
                                pieces: std::mem::take(&mut inline_group),
                                style: inherited.clone(),
                                paragraph_style: None,
                                numbering: None,
                            });
                        }
                        let style = self.computed_style(element, inherited)?;
                        let mut pieces = Vec::new();
                        self.collect_inline_children(
                            element,
                            &style,
                            element.value().name() == "pre",
                            &mut pieces,
                        )?;
                        paragraphs.push(ParagraphModel {
                            pieces,
                            style,
                            paragraph_style: None,
                            numbering: None,
                        });
                    } else if matches!(element.value().name(), "ul" | "ol") {
                        if inline_group.iter().any(inline_piece_is_visible) {
                            paragraphs.push(ParagraphModel {
                                pieces: std::mem::take(&mut inline_group),
                                style: inherited.clone(),
                                paragraph_style: None,
                                numbering: None,
                            });
                        }
                        paragraphs.extend(self.list_models(element, inherited, 0, None)?);
                    } else if element.value().name() != "table" {
                        self.collect_inline(element, inherited, false, &mut inline_group)?;
                    } else {
                        self.diagnostic(
                            &element_path(element),
                            None,
                            "dropped nested HTML table".to_string(),
                        )?;
                    }
                }
                _ => {}
            }
        }
        if inline_group.iter().any(inline_piece_is_visible) || paragraphs.is_empty() {
            paragraphs.push(ParagraphModel {
                pieces: inline_group,
                style: inherited.clone(),
                paragraph_style: None,
                numbering: None,
            });
        }
        Ok(paragraphs)
    }

    fn collect_inline_children(
        &mut self,
        element: ElementRef<'_>,
        style: &ComputedStyle,
        pre: bool,
        pieces: &mut Vec<InlinePiece>,
    ) -> Result<()> {
        for child in element.children() {
            match child.value() {
                Node::Text(text) => {
                    pieces.push(InlinePiece::Text(
                        text.to_string(),
                        Box::new(style.clone()),
                        pre,
                    ));
                }
                Node::Element(_) => {
                    let child = ElementRef::wrap(child).expect("element checked");
                    if matches!(child.value().name(), "ul" | "ol") {
                        continue;
                    }
                    if child.value().name() == "table" {
                        self.diagnostic(
                            &element_path(child),
                            None,
                            "dropped nested HTML table".to_string(),
                        )?;
                        continue;
                    }
                    self.collect_inline(child, style, pre, pieces)?;
                }
                _ => {}
            }
        }
        Ok(())
    }

    fn collect_inline(
        &mut self,
        element: ElementRef<'_>,
        inherited: &ComputedStyle,
        pre: bool,
        pieces: &mut Vec<InlinePiece>,
    ) -> Result<()> {
        let name = element.value().name();
        let location = element_path(element);
        match name {
            "br" => {
                pieces.push(InlinePiece::Break);
                return Ok(());
            }
            "script" => {
                self.diagnostic(&location, None, "dropped HTML script".to_string())?;
                return Ok(());
            }
            "style" | "head" | "title" => return Ok(()),
            "img" => {
                self.diagnostic(
                    &location,
                    None,
                    "dropped HTML image and retained alternate text".to_string(),
                )?;
                if let Some(alt) = element.attr("alt")
                    && !alt.is_empty()
                {
                    let style = self.computed_style(element, inherited)?;
                    pieces.push(InlinePiece::Text(alt.to_string(), Box::new(style), pre));
                }
                return Ok(());
            }
            "link" => {
                self.diagnostic(
                    &location,
                    None,
                    "dropped external HTML stylesheet".to_string(),
                )?;
                return Ok(());
            }
            "iframe" | "frame" | "embed" | "object" => {
                self.diagnostic(&location, None, format!("dropped HTML {name} content"))?;
                return Ok(());
            }
            "input" | "button" | "select" | "textarea" => {
                self.diagnostic(
                    &location,
                    None,
                    format!("dropped HTML form control `{name}`"),
                )?;
                if let Some(value) = element.attr("value")
                    && !value.is_empty()
                {
                    let style = self.computed_style(element, inherited)?;
                    pieces.push(InlinePiece::Text(value.to_string(), Box::new(style), pre));
                }
                return Ok(());
            }
            "a" if element.attr("href").is_some() => {
                self.diagnostic(
                    &location,
                    None,
                    "dropped HTML link target and retained anchor text".to_string(),
                )?;
            }
            "form" => {
                self.diagnostic(
                    &location,
                    None,
                    "dropped HTML form semantics and retained supported content".to_string(),
                )?;
            }
            _ if !is_supported_inline(name) => {
                self.diagnostic(
                    &location,
                    None,
                    format!("unsupported visible HTML element `{name}` retained as text"),
                )?;
            }
            _ => {}
        }

        let style = self.computed_style(element, inherited)?;
        self.collect_inline_children(element, &style, pre, pieces)
    }

    fn computed_style(
        &mut self,
        element: ElementRef<'_>,
        inherited: &ComputedStyle,
    ) -> Result<ComputedStyle> {
        let mut style = inherited.inherited();
        match element.value().name() {
            "b" | "strong" => style.bold = Some(true),
            "i" | "em" | "cite" => style.italic = Some(true),
            "u" | "ins" => style.underline = Some(true),
            "s" | "strike" | "del" => style.strike = Some(true),
            "sup" => style.vertical = Some(VerticalText::Superscript),
            "sub" => style.vertical = Some(VerticalText::Subscript),
            "code" | "kbd" | "samp" => style.font = Some("Courier New".to_string()),
            "mark" => style.background = Some("FFFF00".to_string()),
            _ => {}
        }
        let mut matching: Vec<_> = self
            .rules
            .iter()
            .filter(|rule| rule.selector.matches(&element))
            .collect();
        matching.sort_by(|left, right| {
            let specificity = compare_specificity(left.specificity, right.specificity);
            if specificity == Ordering::Equal {
                left.order.cmp(&right.order)
            } else {
                specificity
            }
        });
        for rule in matching {
            for change in &rule.changes {
                change.apply(&mut style);
            }
        }
        if let Some(inline) = element.attr("style") {
            let location = element_path(element);
            for change in self.parse_declarations(inline, &location)? {
                change.apply(&mut style);
            }
        }
        Ok(style)
    }

    fn build_paragraph(&mut self, model: ParagraphModel) -> Result<CT_P> {
        self.bump_blocks("html")?;
        let mut output = CT_P::new();
        {
            let mut paragraph = Paragraph { inner: &mut output };
            apply_paragraph_style(&mut paragraph, &model.style);
            if let Some(style) = model.paragraph_style {
                paragraph.set_style(&style);
            }
            if let Some((list_id, level)) = model.numbering
                && !paragraph.set_numbering(list_id, level)
            {
                return Err(html_error("html", "could not attach paragraph numbering"));
            }
            let mut emitted = false;
            let mut pending_space = false;
            for piece in model.pieces {
                match piece {
                    InlinePiece::Break => {
                        self.bump_runs()?;
                        paragraph.add_line_break();
                        emitted = true;
                        pending_space = false;
                    }
                    InlinePiece::Text(text, style, true) => {
                        let mut lines = text.split('\n').peekable();
                        while let Some(line) = lines.next() {
                            if !line.is_empty() {
                                self.bump_runs()?;
                                let mut run = paragraph.add_run(line);
                                apply_run_style(&mut run, &style);
                                emitted = true;
                            }
                            if lines.peek().is_some() {
                                self.bump_runs()?;
                                paragraph.add_line_break();
                                emitted = true;
                            }
                        }
                        pending_space = false;
                    }
                    InlinePiece::Text(text, style, false) => {
                        let normalized = collapse_text(&text, &mut pending_space, emitted);
                        if !normalized.is_empty() {
                            self.bump_runs()?;
                            let mut run = paragraph.add_run(&normalized);
                            apply_run_style(&mut run, &style);
                            emitted = true;
                        }
                    }
                }
            }
        }
        Ok(output)
    }

    fn bump_blocks(&mut self, location: &str) -> Result<()> {
        self.blocks = self
            .blocks
            .checked_add(1)
            .ok_or_else(|| html_error(location, "projected block count overflowed"))?;
        if self.blocks > self.limits.blocks {
            return Err(html_error(
                location,
                "HTML exceeds the projected block limit",
            ));
        }
        Ok(())
    }

    fn bump_runs(&mut self) -> Result<()> {
        self.runs = self
            .runs
            .checked_add(1)
            .ok_or_else(|| html_error("html", "projected run count overflowed"))?;
        if self.runs > self.limits.runs {
            return Err(html_error("html", "HTML exceeds the run limit"));
        }
        Ok(())
    }

    fn diagnostic(
        &mut self,
        location: &str,
        property: Option<String>,
        message: String,
    ) -> Result<()> {
        let diagnostic = HtmlDiagnostic {
            location: location.to_string(),
            property,
            message,
        };
        if self.diagnostic_keys.contains(&diagnostic) {
            return Ok(());
        }
        if self.diagnostics.len() >= self.limits.diagnostics {
            return Err(html_error(location, "HTML exceeds the diagnostic limit"));
        }
        self.diagnostic_keys.insert(diagnostic.clone());
        self.diagnostics.push(diagnostic);
        Ok(())
    }

    fn finish(mut self) -> Result<HtmlReadResult> {
        let bytes = self.document.to_bytes()?;
        let document = Document::from_bytes(&bytes)?;
        Ok(HtmlReadResult {
            document,
            diagnostics: self.diagnostics,
        })
    }
}

fn apply_paragraph_style(paragraph: &mut Paragraph<'_>, style: &ComputedStyle) {
    if let Some(alignment) = style.alignment {
        paragraph.set_alignment(alignment);
    }
    if let Some(value) = style.space_before {
        paragraph.set_space_before(value);
    }
    if let Some(value) = style.space_after {
        paragraph.set_space_after(value);
    }
    if let Some(value) = style.indent_left {
        paragraph.set_indent_left(value);
    }
    if let Some(value) = style.indent_right {
        paragraph.set_indent_right(value);
    }
    if let Some(value) = style.first_line_indent {
        paragraph.set_signed_first_line_indent_value(Some(value));
    }
    if let Some(background) = &style.background {
        paragraph.set_shading(background);
    }
}

fn apply_run_style(run: &mut Run<'_>, style: &ComputedStyle) {
    if let Some(font) = &style.font {
        run.set_font(font);
    }
    if let Some(size) = style.size {
        run.set_size(size);
    }
    if let Some(value) = style.bold {
        run.set_bold(value);
    }
    if let Some(value) = style.italic {
        run.set_italic(value);
    }
    if let Some(value) = style.underline {
        run.set_underline(value);
    }
    if let Some(value) = style.strike {
        run.set_strike(value);
    }
    if let Some(color) = &style.color {
        run.set_color(color);
    }
    if let Some(background) = &style.background {
        run.set_highlight(background);
    }
    match style.vertical {
        Some(VerticalText::Superscript) => run.set_superscript(),
        Some(VerticalText::Subscript) => run.set_subscript(),
        None => {}
    }
}

fn collapse_text(text: &str, pending_space: &mut bool, already_emitted: bool) -> String {
    let mut output = String::new();
    let mut emitted = already_emitted;
    for character in text.chars() {
        if character.is_whitespace() {
            *pending_space = true;
        } else {
            if *pending_space && emitted {
                output.push(' ');
            }
            output.push(character);
            *pending_space = false;
            emitted = true;
        }
    }
    output
}

fn inline_piece_is_visible(piece: &InlinePiece) -> bool {
    match piece {
        InlinePiece::Text(text, _, pre) => *pre || !text.trim().is_empty(),
        InlinePiece::Break => true,
    }
}

fn is_block(name: &str) -> bool {
    matches!(
        name,
        "address"
            | "article"
            | "aside"
            | "blockquote"
            | "div"
            | "footer"
            | "h1"
            | "h2"
            | "h3"
            | "h4"
            | "h5"
            | "h6"
            | "header"
            | "main"
            | "nav"
            | "ol"
            | "p"
            | "pre"
            | "section"
            | "table"
            | "ul"
    )
}

fn is_supported_inline(name: &str) -> bool {
    matches!(
        name,
        "a" | "abbr"
            | "b"
            | "br"
            | "cite"
            | "code"
            | "del"
            | "em"
            | "form"
            | "i"
            | "ins"
            | "kbd"
            | "mark"
            | "s"
            | "samp"
            | "span"
            | "strike"
            | "strong"
            | "sub"
            | "sup"
            | "u"
    )
}

fn parse_span(
    value: Option<&str>,
    attribute: &str,
    element: ElementRef<'_>,
    maximum: usize,
) -> Result<usize> {
    let Some(value) = value else {
        return Ok(1);
    };
    let parsed = value.parse::<usize>().map_err(|_| {
        html_error(
            element_path(element),
            format!("invalid HTML {attribute} value `{value}`"),
        )
    })?;
    if parsed == 0 || parsed > maximum {
        return Err(html_error(
            element_path(element),
            format!("HTML {attribute} value is outside the supported bound"),
        ));
    }
    Ok(parsed)
}

fn nearest_ancestor_named<'a>(element: ElementRef<'a>, name: &str) -> Option<ElementRef<'a>> {
    let mut ancestor = element.parent();
    while let Some(node) = ancestor {
        if let Some(element) = ElementRef::wrap(node)
            && element.value().name() == name
        {
            return Some(element);
        }
        ancestor = node.parent();
    }
    None
}

fn element_path(element: ElementRef<'_>) -> String {
    let mut segments = Vec::new();
    let mut current = Some(*element);
    while let Some(node) = current {
        if let Some(element) = ElementRef::wrap(node) {
            let name = element.value().name();
            let mut index = 1_usize;
            let mut sibling = element.prev_sibling();
            while let Some(node) = sibling {
                if let Some(sibling_element) = ElementRef::wrap(node)
                    && sibling_element.value().name() == name
                {
                    index += 1;
                }
                sibling = node.prev_sibling();
            }
            if matches!(name, "html" | "head" | "body") {
                segments.push(name.to_string());
            } else {
                segments.push(format!("{name}[{index}]"));
            }
        }
        current = node.parent();
    }
    segments.reverse();
    if segments.first().is_none_or(|segment| segment != "html") {
        segments.insert(0, "html".to_string());
    }
    if segments
        .get(1)
        .is_none_or(|segment| !segment.starts_with("body"))
    {
        segments.insert(1, "body".to_string());
    }
    segments.join("/")
}

fn strip_css_comments(css: &str) -> String {
    let mut output = String::with_capacity(css.len());
    let mut remainder = css;
    while let Some(start) = remainder.find("/*") {
        output.push_str(&remainder[..start]);
        let after_start = &remainder[start + 2..];
        let Some(end) = after_start.find("*/") else {
            break;
        };
        remainder = &after_start[end + 2..];
    }
    output.push_str(remainder);
    output
}

fn supported_selector(selector: &str) -> bool {
    !selector.is_empty()
        && !selector
            .chars()
            .any(|character| matches!(character, '[' | ']' | ':' | '+' | '~' | '*' | '|'))
        && selector.chars().all(|character| {
            character.is_ascii_alphanumeric()
                || character.is_ascii_whitespace()
                || matches!(character, '-' | '_' | '.' | '#' | '>')
        })
        && !selector.contains(">>")
}

fn specificity(selector: &str) -> (u32, u32, u32) {
    let ids = selector.bytes().filter(|byte| *byte == b'#').count() as u32;
    let classes = selector.bytes().filter(|byte| *byte == b'.').count() as u32;
    let types = selector
        .split(|character: char| character.is_ascii_whitespace() || character == '>')
        .filter(|part| !part.is_empty())
        .filter(|part| {
            part.as_bytes()
                .first()
                .is_some_and(|byte| byte.is_ascii_alphabetic())
        })
        .count() as u32;
    (ids, classes, types)
}

fn compare_specificity(left: (u32, u32, u32), right: (u32, u32, u32)) -> Ordering {
    left.0
        .cmp(&right.0)
        .then_with(|| left.1.cmp(&right.1))
        .then_with(|| left.2.cmp(&right.2))
}

fn parse_style_change(
    property: &str,
    value: &str,
) -> std::result::Result<Vec<StyleChange>, String> {
    let lower = value.trim().to_ascii_lowercase();
    let changes = match property {
        "font-family" => {
            let family = value
                .split(',')
                .next()
                .unwrap_or_default()
                .trim()
                .trim_matches(['\'', '"']);
            if family.is_empty() {
                return Err("unsupported empty font-family value".to_string());
            }
            vec![StyleChange::Font(family.to_string())]
        }
        "font-size" => vec![StyleChange::Size(parse_points(value)?)],
        "font-weight" => vec![StyleChange::Bold(match lower.as_str() {
            "normal" | "400" | "500" => false,
            "bold" | "bolder" | "600" | "700" | "800" | "900" => true,
            _ => return Err(format!("unsupported font-weight value `{value}`")),
        })],
        "font-style" => vec![StyleChange::Italic(match lower.as_str() {
            "normal" => false,
            "italic" | "oblique" => true,
            _ => return Err(format!("unsupported font-style value `{value}`")),
        })],
        "text-decoration" | "text-decoration-line" => {
            if lower == "none" {
                vec![StyleChange::Underline(false), StyleChange::Strike(false)]
            } else {
                let mut parsed = Vec::new();
                for token in lower.split_whitespace() {
                    match token {
                        "underline" => parsed.push(StyleChange::Underline(true)),
                        "line-through" => parsed.push(StyleChange::Strike(true)),
                        _ => {
                            return Err(format!("unsupported text-decoration value `{value}`"));
                        }
                    }
                }
                parsed
            }
        }
        "color" => vec![StyleChange::Color(parse_color(value)?)],
        "background" | "background-color" => {
            vec![StyleChange::Background(if lower == "transparent" {
                None
            } else {
                Some(parse_color(value)?)
            })]
        }
        "text-align" => vec![StyleChange::Alignment(match lower.as_str() {
            "left" | "start" => Alignment::Left,
            "center" => Alignment::Center,
            "right" | "end" => Alignment::Right,
            "justify" => Alignment::Justify,
            _ => return Err(format!("unsupported text-align value `{value}`")),
        })],
        "margin-top" => vec![StyleChange::SpaceBefore(parse_length(value)?)],
        "margin-bottom" => vec![StyleChange::SpaceAfter(parse_length(value)?)],
        "margin-left" => vec![StyleChange::IndentLeft(parse_length(value)?)],
        "margin-right" => vec![StyleChange::IndentRight(parse_length(value)?)],
        "text-indent" => vec![StyleChange::FirstLineIndent(parse_signed_length(value)?)],
        _ => return Err(format!("unsupported CSS property `{property}`")),
    };
    Ok(changes)
}

fn parse_points(value: &str) -> std::result::Result<f64, String> {
    let value = value.trim().to_ascii_lowercase();
    let points = if let Some(number) = value.strip_suffix("pt") {
        number.trim().parse::<f64>()
    } else if let Some(number) = value.strip_suffix("px") {
        number.trim().parse::<f64>().map(|number| number * 0.75)
    } else {
        return Err(format!("unsupported CSS length `{value}`"));
    }
    .map_err(|_| format!("invalid CSS length `{value}`"))?;
    if !points.is_finite() || points <= 0.0 || points > 1000.0 {
        return Err(format!(
            "CSS length is outside the supported range `{value}`"
        ));
    }
    Ok(points)
}

fn parse_length(value: &str) -> std::result::Result<Length, String> {
    if value.trim() == "0" {
        return Ok(Length::pt(0.0));
    }
    parse_points(value).map(Length::pt)
}

fn parse_signed_length(value: &str) -> std::result::Result<Length, String> {
    let value = value.trim().to_ascii_lowercase();
    if value == "0" {
        return Ok(Length::pt(0.0));
    }
    let points = if let Some(number) = value.strip_suffix("pt") {
        number.trim().parse::<f64>()
    } else if let Some(number) = value.strip_suffix("px") {
        number.trim().parse::<f64>().map(|number| number * 0.75)
    } else {
        return Err(format!("unsupported CSS length `{value}`"));
    }
    .map_err(|_| format!("invalid CSS length `{value}`"))?;
    if !points.is_finite() || points.abs() > 1000.0 {
        return Err(format!(
            "CSS length is outside the supported range `{value}`"
        ));
    }
    Ok(Length::pt(points))
}

fn parse_color(value: &str) -> std::result::Result<String, String> {
    let value = value.trim().to_ascii_lowercase();
    let color = match value.as_str() {
        "black" => "000000".to_string(),
        "white" => "FFFFFF".to_string(),
        "red" => "FF0000".to_string(),
        "green" => "008000".to_string(),
        "blue" => "0000FF".to_string(),
        "yellow" => "FFFF00".to_string(),
        "gray" | "grey" => "808080".to_string(),
        value if value.len() == 7 && value.starts_with('#') => value[1..].to_ascii_uppercase(),
        value if value.len() == 4 && value.starts_with('#') => {
            let mut expanded = String::with_capacity(6);
            for character in value[1..].chars() {
                expanded.push(character.to_ascii_uppercase());
                expanded.push(character.to_ascii_uppercase());
            }
            expanded
        }
        _ => return Err(format!("unsupported CSS color `{value}`")),
    };
    if !color.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return Err(format!("invalid CSS color `{value}`"));
    }
    Ok(color)
}

#[cfg(test)]
mod tests {
    use std::io::Cursor;

    use super::{Document, Limits, from_html_with_limits, preflight_markup, read_bounded};

    #[test]
    fn html_import_rejects_each_declared_resource_limit() {
        let cases = [
            (
                "<p>x</p>",
                Limits {
                    input_bytes: 1,
                    ..Limits::default()
                },
            ),
            (
                "<p>x</p>",
                Limits {
                    retained_text: 0,
                    ..Limits::default()
                },
            ),
            (
                "<div><p>x</p></div>",
                Limits {
                    depth: 2,
                    ..Limits::default()
                },
            ),
            (
                "<p>x</p>",
                Limits {
                    nodes: 1,
                    ..Limits::default()
                },
            ),
            (
                "<p>x</p>",
                Limits {
                    blocks: 0,
                    ..Limits::default()
                },
            ),
            (
                "<p>x</p>",
                Limits {
                    runs: 0,
                    ..Limits::default()
                },
            ),
            (
                "<pre>\n\n</pre>",
                Limits {
                    runs: 0,
                    ..Limits::default()
                },
            ),
            (
                "<table><tr><td>x</td></tr></table>",
                Limits {
                    rows: 0,
                    ..Limits::default()
                },
            ),
            (
                "<table><tr><td colspan='2'>x</td></tr></table>",
                Limits {
                    columns: 1,
                    ..Limits::default()
                },
            ),
            (
                "<table><tr><td>x</td></tr></table>",
                Limits {
                    cells: 0,
                    ..Limits::default()
                },
            ),
            (
                "<p style='unknown:x'>x</p>",
                Limits {
                    diagnostics: 0,
                    ..Limits::default()
                },
            ),
        ];
        for (html, limits) in cases {
            assert!(
                from_html_with_limits(html, limits).is_err(),
                "limit accepted {html}"
            );
        }

        let oversized = "x".repeat(64 * 1024 * 1024 + 1);
        assert!(Document::from_html(&oversized).is_err());

        let mut reader = Cursor::new(b"12345".to_vec());
        assert!(read_bounded(&mut reader, 0, 4).is_err());
        assert!(
            preflight_markup(
                "<!-- one --><!-- two -->",
                Limits {
                    nodes: 4,
                    ..Limits::default()
                }
            )
            .is_err()
        );
        assert!(
            preflight_markup(
                "<!-- <x><x><x> -->",
                Limits {
                    nodes: 5,
                    ..Limits::default()
                }
            )
            .is_ok()
        );
    }
}
