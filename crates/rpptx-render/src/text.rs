use oxml_layout::{
    Align, Color, FieldKind, FontManager, GlyphRun, InlineItem, LayoutError, LayoutLine,
    LineBreakParams, LineItem, LineSpacing, MultilingualGlyphRun, Paint, Point, PositionedElement,
    Rect, TextSegment, Underline, break_multilingual_into_lines,
};
use rpptx_layout::{
    ParagraphAlignment, ResolvedAutofit, ResolvedBullet, ResolvedBulletSize, ResolvedParagraph,
    ResolvedRunStyle, ResolvedShape, ResolvedTextBody, ResolvedTextRun, ResolvedTextSpacing,
    TextAnchor, TextDirection,
};

const DEFAULT_FONT_SIZE: f64 = 18.0;

struct AutoNumberSequence {
    scheme: String,
    next: u32,
}

struct NumberingState {
    levels: Vec<Option<AutoNumberSequence>>,
}

#[derive(Default)]
struct ShapingCache {
    entries: Vec<(String, ResolvedRunStyle, TextSegment)>,
}

impl ShapingCache {
    fn shape(
        &mut self,
        font_manager: &mut FontManager,
        text: &str,
        style: &ResolvedRunStyle,
    ) -> Result<TextSegment, LayoutError> {
        if let Some((_, _, segment)) = self
            .entries
            .iter()
            .find(|(cached_text, cached_style, _)| cached_text == text && cached_style == style)
        {
            return Ok(segment.clone());
        }
        let segment = shape_run(font_manager, text, style)?;
        self.entries
            .push((text.to_owned(), style.clone(), segment.clone()));
        Ok(segment)
    }
}

impl Default for NumberingState {
    fn default() -> Self {
        Self {
            levels: std::iter::repeat_with(|| None).take(9).collect(),
        }
    }
}

impl NumberingState {
    fn marker(
        &mut self,
        font_manager: &mut FontManager,
        shaping_cache: &mut ShapingCache,
        paragraph: &ResolvedParagraph,
    ) -> Result<Option<TextSegment>, LayoutError> {
        let level = usize::from(paragraph.level.min(8));
        self.levels[level + 1..].fill_with(|| None);

        let (text, font, color, size) = match paragraph.bullet.as_ref() {
            Some(ResolvedBullet::Character {
                character,
                font,
                color,
                size,
            }) => {
                self.levels[level] = None;
                (map_bullet_characters(character), font, color, size)
            }
            Some(ResolvedBullet::AutoNumber {
                scheme,
                start_at,
                font,
                color,
                size,
            }) => {
                let value = self.levels[level]
                    .as_ref()
                    .filter(|sequence| sequence.scheme == *scheme)
                    .map_or(*start_at, |sequence| sequence.next);
                self.levels[level] = Some(AutoNumberSequence {
                    scheme: scheme.clone(),
                    next: value.saturating_add(1),
                });
                (format_auto_number(scheme, value), font, color, size)
            }
            None => {
                self.levels[level] = None;
                return Ok(None);
            }
        };

        let mut marker = shape_marker(
            font_manager,
            shaping_cache,
            paragraph,
            &text,
            font,
            *color,
            size,
        )?;
        if paragraph.indent < 0.0 {
            marker.width = -paragraph.indent;
        }
        Ok(Some(marker))
    }
}

#[cfg(test)]
pub(super) fn inline_items(
    font_manager: &mut FontManager,
    paragraph: &ResolvedParagraph,
) -> Result<Vec<InlineItem>, LayoutError> {
    inline_items_cached(font_manager, &mut ShapingCache::default(), paragraph, 1)
}

fn inline_items_cached(
    font_manager: &mut FontManager,
    shaping_cache: &mut ShapingCache,
    paragraph: &ResolvedParagraph,
    page_number: usize,
) -> Result<Vec<InlineItem>, LayoutError> {
    let mut items = Vec::with_capacity(paragraph.runs.len());
    for run in &paragraph.runs {
        match run {
            ResolvedTextRun::Text { text, style } => {
                if !text.is_empty() {
                    items.push(InlineItem::Text(shaping_cache.shape(
                        font_manager,
                        text,
                        style,
                    )?));
                }
            }
            ResolvedTextRun::Field {
                text,
                field_type,
                style,
            } => {
                let slide_number = field_type
                    .as_deref()
                    .is_some_and(|field_type| field_type.eq_ignore_ascii_case("slidenum"));
                let rendered = if slide_number {
                    page_number.to_string()
                } else {
                    text.clone()
                };
                if !rendered.is_empty() {
                    let mut segment = shaping_cache.shape(font_manager, &rendered, style)?;
                    if slide_number {
                        segment.field_kind = Some(FieldKind::Page);
                    }
                    items.push(InlineItem::Text(segment));
                }
            }
            ResolvedTextRun::Break => items.push(InlineItem::LineBreak),
        }
    }
    Ok(items)
}

fn shape_run(
    font_manager: &mut FontManager,
    text: &str,
    style: &ResolvedRunStyle,
) -> Result<TextSegment, LayoutError> {
    let rendered_text = if style.all_caps {
        text.to_uppercase()
    } else {
        text.to_owned()
    };
    let font_size = style.font_size.unwrap_or(DEFAULT_FONT_SIZE);
    let typeface = typeface_for_text(style, &rendered_text);
    let font_id =
        font_manager.resolve_font_for_text(typeface, style.bold, style.italic, &rendered_text)?;
    let metrics = font_manager.metrics(font_id, font_size)?;
    let mut shaped = font_manager.shape_text(font_id, &rendered_text, font_size)?;
    if let Some(spacing) = style.spacing {
        for advance in &mut shaped.advances {
            *advance += spacing;
        }
        shaped.width += spacing * shaped.advances.len() as f64;
    }

    Ok(TextSegment {
        text: rendered_text,
        direction: oxml_layout::TextDirection::Auto,
        source: None,
        font_id,
        font_size,
        glyph_ids: shaped.glyph_ids,
        advances: shaped.advances,
        width: shaped.width,
        ascent: metrics.ascent,
        descent: metrics.descent,
        line_gap: metrics.line_gap,
        color: text_color(style.fill.as_ref()),
        bold: style.bold,
        italic: style.italic,
        underline: style.underline.then_some(Underline::Single),
        strike: style.strike,
        dstrike: false,
        highlight: None,
        baseline_offset: style.baseline.unwrap_or(0.0) * font_size,
        hyperlink_url: style.hyperlink_url.clone(),
        field_kind: None,
        note: None,
    })
}

fn shape_marker(
    font_manager: &mut FontManager,
    shaping_cache: &mut ShapingCache,
    paragraph: &ResolvedParagraph,
    text: &str,
    font: &Option<String>,
    color: Option<Color>,
    size: &Option<ResolvedBulletSize>,
) -> Result<TextSegment, LayoutError> {
    let mut style = first_run_style(paragraph).clone();
    let base_size = style.font_size.unwrap_or(DEFAULT_FONT_SIZE);
    style.font_size = Some(match size {
        Some(ResolvedBulletSize::Percent(factor)) => base_size * factor,
        Some(ResolvedBulletSize::Points(points)) => *points,
        None => base_size,
    });
    if let Some(font) = font {
        style.latin_typeface = Some(font.clone());
        style.east_asian_typeface = Some(font.clone());
        style.complex_script_typeface = Some(font.clone());
        style.symbol_typeface = Some(font.clone());
    }
    if let Some(color) = color {
        style.fill = Some(Paint::Solid(color));
    }
    shaping_cache.shape(font_manager, text, &style)
}

fn first_run_style(paragraph: &ResolvedParagraph) -> &ResolvedRunStyle {
    paragraph
        .runs
        .iter()
        .find_map(|run| match run {
            ResolvedTextRun::Text { style, .. } | ResolvedTextRun::Field { style, .. } => {
                Some(style)
            }
            ResolvedTextRun::Break => None,
        })
        .unwrap_or(&DEFAULT_RUN_STYLE)
}

static DEFAULT_RUN_STYLE: ResolvedRunStyle = ResolvedRunStyle {
    font_size: None,
    bold: false,
    italic: false,
    all_caps: false,
    underline: false,
    strike: false,
    spacing: None,
    baseline: None,
    fill: None,
    latin_typeface: None,
    east_asian_typeface: None,
    complex_script_typeface: None,
    symbol_typeface: None,
    hyperlink_url: None,
};

fn map_bullet_characters(character: &str) -> String {
    character.replace('\u{f0b7}', "\u{2022}")
}

fn format_auto_number(scheme: &str, value: u32) -> String {
    match scheme {
        "arabicPlain" => value.to_string(),
        "arabicParenR" => format!("{value})"),
        "arabicParenBoth" => format!("({value})"),
        "alphaLcPeriod" => format!("{}.", alphabetic_number(value, false)),
        "alphaUcPeriod" => format!("{}.", alphabetic_number(value, true)),
        "romanLcPeriod" => format!("{}.", roman_number(value, false)),
        "romanUcPeriod" => format!("{}.", roman_number(value, true)),
        "arabicPeriod" => format!("{value}."),
        _ => format!("{value}."),
    }
}

fn alphabetic_number(value: u32, uppercase: bool) -> String {
    let mut value = value.max(1);
    let mut letters = Vec::new();
    while value > 0 {
        value -= 1;
        let base = if uppercase { b'A' } else { b'a' };
        letters.push(char::from(base + (value % 26) as u8));
        value /= 26;
    }
    letters.into_iter().rev().collect()
}

fn roman_number(value: u32, uppercase: bool) -> String {
    let mut value = value.max(1);
    let mut result = String::new();
    for (amount, numeral) in [
        (1000, "M"),
        (900, "CM"),
        (500, "D"),
        (400, "CD"),
        (100, "C"),
        (90, "XC"),
        (50, "L"),
        (40, "XL"),
        (10, "X"),
        (9, "IX"),
        (5, "V"),
        (4, "IV"),
        (1, "I"),
    ] {
        while value >= amount {
            result.push_str(numeral);
            value -= amount;
        }
    }
    if uppercase {
        result
    } else {
        result.to_ascii_lowercase()
    }
}

fn typeface_for_text<'a>(style: &'a ResolvedRunStyle, text: &str) -> Option<&'a str> {
    let preferred = if text.chars().any(is_symbol_character) {
        style.symbol_typeface.as_deref()
    } else if text.chars().any(is_east_asian_character) {
        style.east_asian_typeface.as_deref()
    } else if text.chars().any(is_complex_script_character) {
        style.complex_script_typeface.as_deref()
    } else {
        style.latin_typeface.as_deref()
    };
    preferred
        .or(style.latin_typeface.as_deref())
        .or(style.east_asian_typeface.as_deref())
        .or(style.complex_script_typeface.as_deref())
        .or(style.symbol_typeface.as_deref())
}

fn is_symbol_character(character: char) -> bool {
    matches!(character as u32, 0xE000..=0xF8FF)
}

fn is_east_asian_character(character: char) -> bool {
    matches!(
        character as u32,
        0x1100..=0x11FF
            | 0x2E80..=0xA4CF
            | 0xAC00..=0xD7AF
            | 0xF900..=0xFAFF
            | 0xFE30..=0xFE4F
            | 0xFF00..=0xFFEF
            | 0x20000..=0x323AF
    )
}

fn is_complex_script_character(character: char) -> bool {
    matches!(
        character as u32,
        0x0590..=0x10FF | 0x1780..=0x17FF | 0xFB1D..=0xFDFF | 0xFE70..=0xFEFF
    )
}

fn text_color(fill: Option<&Paint>) -> Color {
    match fill {
        Some(Paint::Solid(color)) => *color,
        Some(Paint::Linear { stops, .. }) | Some(Paint::Radial { stops, .. }) => {
            stops.first().map_or(Color::BLACK, |stop| stop.color)
        }
        Some(Paint::Tile { .. }) | None => Color::BLACK,
    }
}

pub(super) struct StackedText {
    pub(super) elements: Vec<PositionedElement>,
    pub(super) width: f64,
    pub(super) height: f64,
    width_fits: bool,
}

#[cfg(test)]
pub(super) fn stack_text(
    font_manager: &mut FontManager,
    content: Rect,
    text: &ResolvedTextBody,
) -> Result<StackedText, LayoutError> {
    stack_text_for_page(font_manager, content, text, 1)
}

#[cfg(test)]
pub(super) fn stack_text_for_page(
    font_manager: &mut FontManager,
    content: Rect,
    text: &ResolvedTextBody,
    page_number: usize,
) -> Result<StackedText, LayoutError> {
    stack_text_for_page_with_directions(font_manager, content, text, page_number, &[])
}

pub(super) fn stack_text_for_page_with_directions(
    font_manager: &mut FontManager,
    content: Rect,
    text: &ResolvedTextBody,
    page_number: usize,
    paragraph_directions: &[oxml_layout::TextDirection],
) -> Result<StackedText, LayoutError> {
    let shaped_paragraphs =
        shape_paragraphs(font_manager, text, page_number, paragraph_directions)?;
    match text.autofit {
        ResolvedAutofit::None | ResolvedAutofit::Shape => {
            stack_shaped_text(font_manager, content, text, &shaped_paragraphs, 1.0, None)
        }
        ResolvedAutofit::Normal {
            font_scale: None,
            line_spacing_reduction: None,
        } => {
            let mut floor =
                stack_shaped_text(font_manager, content, text, &shaped_paragraphs, 1.0, None)?;
            if floor.width_fits && floor.height <= content.height + 0.01 {
                return Ok(floor);
            }
            for step in 1..=30 {
                let scale = 1.0 - f64::from(step) * 0.025;
                let candidate = stack_shaped_text(
                    font_manager,
                    content,
                    text,
                    &shaped_paragraphs,
                    scale,
                    None,
                )?;
                if candidate.width_fits && candidate.height <= content.height + 0.01 {
                    return Ok(candidate);
                }
                floor = candidate;
            }
            Ok(floor)
        }
        ResolvedAutofit::Normal {
            font_scale,
            line_spacing_reduction,
        } => stack_shaped_text(
            font_manager,
            content,
            text,
            &shaped_paragraphs,
            font_scale.unwrap_or(1.0),
            line_spacing_reduction,
        ),
    }
}

fn shape_paragraphs(
    font_manager: &mut FontManager,
    text: &ResolvedTextBody,
    page_number: usize,
    paragraph_directions: &[oxml_layout::TextDirection],
) -> Result<Vec<(Vec<InlineItem>, oxml_layout::TextDirection)>, LayoutError> {
    let mut numbering = NumberingState::default();
    let mut shaping_cache = ShapingCache::default();
    text.paragraphs
        .iter()
        .enumerate()
        .map(|(paragraph_index, paragraph)| {
            let base_direction = paragraph_directions
                .get(paragraph_index)
                .copied()
                .unwrap_or(oxml_layout::TextDirection::Auto);
            let marker = numbering.marker(font_manager, &mut shaping_cache, paragraph)?;
            let legacy_items =
                inline_items_cached(font_manager, &mut shaping_cache, paragraph, page_number)?;
            let uses_rich_text = base_direction != oxml_layout::TextDirection::Auto
                || legacy_items.iter().any(|item| {
                    matches!(item, InlineItem::Text(segment) if needs_multilingual_layout(&segment.text))
                });
            let mut items = if uses_rich_text {
                shape_multilingual_items_preserving_controls(
                    font_manager,
                    legacy_items,
                    base_direction,
                )?
            } else {
                legacy_items
            };
            let has_visible_text = items.iter().any(|item| match item {
                InlineItem::Text(segment) => !segment.text.is_empty(),
                InlineItem::MultilingualText(segment) => !segment.text().is_empty(),
                _ => false,
            });
            if items.is_empty() {
                items.push(InlineItem::Text(shaping_cache.shape(
                    font_manager,
                    "",
                    &paragraph.end_style,
                )?));
            }
            if let Some(marker) = marker.filter(|_| has_visible_text) {
                items.insert(0, InlineItem::Marker(marker));
            }
            Ok((items, base_direction))
        })
        .collect()
}

fn shape_multilingual_items_preserving_controls(
    font_manager: &mut FontManager,
    legacy_items: Vec<InlineItem>,
    base_direction: oxml_layout::TextDirection,
) -> Result<Vec<InlineItem>, LayoutError> {
    let styled = legacy_items
        .iter()
        .filter_map(|item| match item {
            InlineItem::Text(segment) => Some((segment.clone(), None)),
            _ => None,
        })
        .collect();
    let mut shaped = std::collections::VecDeque::from(font_manager.shape_multilingual_paragraph(
        styled,
        base_direction,
        false,
    )?);
    let mut items = Vec::new();
    for item in legacy_items {
        let InlineItem::Text(segment) = item else {
            items.push(item);
            continue;
        };
        let target_bytes = segment.text.len();
        let mut consumed_bytes = 0usize;
        while consumed_bytes < target_bytes {
            let Some(span) = shaped.pop_front() else {
                return Err(LayoutError::Layout(
                    "multilingual paragraph shaping lost a styled text run".to_owned(),
                ));
            };
            consumed_bytes = consumed_bytes
                .checked_add(span.text().len())
                .filter(|consumed| *consumed <= target_bytes)
                .ok_or_else(|| {
                    LayoutError::Layout(
                        "multilingual paragraph shaping crossed a styled text boundary".to_owned(),
                    )
                })?;
            items.push(InlineItem::MultilingualText(span));
        }
    }
    if !shaped.is_empty() {
        return Err(LayoutError::Layout(
            "multilingual paragraph shaping produced an unmatched text run".to_owned(),
        ));
    }
    Ok(items)
}

fn needs_multilingual_layout(text: &str) -> bool {
    text.chars().any(|character| {
        matches!(
            character as u32,
            0x0590..=0x08ff
                | 0x0900..=0x097f
                | 0x0e00..=0x0e7f
                | 0x3000..=0x30ff
                | 0x3400..=0x9fff
                | 0xf900..=0xfaff
        )
    })
}

fn stack_shaped_text(
    font_manager: &FontManager,
    content: Rect,
    text: &ResolvedTextBody,
    shaped_paragraphs: &[(Vec<InlineItem>, oxml_layout::TextDirection)],
    font_scale: f64,
    line_spacing_reduction: Option<f64>,
) -> Result<StackedText, LayoutError> {
    let mut elements = Vec::new();
    let mut line_ranges = Vec::new();
    let mut y = content.y;
    let mut width = 0.0_f64;
    let mut width_fits = true;

    let last_paragraph = text.paragraphs.len().saturating_sub(1);
    for (index, (paragraph, (shaped_items, base_direction))) in
        text.paragraphs.iter().zip(shaped_paragraphs).enumerate()
    {
        let font_size = first_run_font_size(paragraph) * font_scale;
        if index > 0 || text.space_first_last_paragraph {
            y += paragraph_spacing(paragraph.space_before.as_ref(), font_size);
        }
        let items = scaled_inline_items(shaped_items, font_scale, paragraph.indent);
        let params = line_break_params(paragraph, content.width, text.wrap);
        let mut lines =
            break_multilingual_into_lines(&items, &params, font_manager, *base_direction)?;
        if let Some(reduction) = line_spacing_reduction {
            reduce_line_spacing(&mut lines, paragraph.line_spacing.as_ref(), reduction);
        }

        for line in &lines {
            width_fits &= line.width <= line.available_width + 0.01;
            let baseline = y + line.baseline_offset();
            let element_start = elements.len();
            let occupied_width = emit_line_items(
                line,
                paragraph.alignment,
                content.x,
                baseline,
                &mut elements,
            );
            line_ranges.push((element_start, elements.len()));
            width = width.max(occupied_width);
            y += line.height;
        }

        if index < last_paragraph || text.space_first_last_paragraph {
            y += paragraph_spacing(paragraph.space_after.as_ref(), font_size);
        }
    }

    let height = y - content.y;
    anchor_lines(
        &mut elements,
        &line_ranges,
        text.anchor,
        content.height - height,
    );

    Ok(StackedText {
        elements,
        width,
        height,
        width_fits,
    })
}

fn scaled_inline_items(items: &[InlineItem], scale: f64, indent: f64) -> Vec<InlineItem> {
    items
        .iter()
        .cloned()
        .map(|item| match item {
            InlineItem::Text(mut segment) => {
                scale_segment(&mut segment, scale);
                InlineItem::Text(segment)
            }
            InlineItem::MultilingualText(segment) => {
                let mut base = segment.base().clone();
                scale_segment(&mut base, scale);
                InlineItem::MultilingualText(
                    oxml_layout::MultilingualTextSegment::new(
                        base,
                        segment.logical_index(),
                        segment.language().map(str::to_owned),
                        segment.script(),
                        segment.direction(),
                        segment.bidi_level(),
                        segment
                            .x_advances()
                            .iter()
                            .map(|value| value * scale)
                            .collect(),
                        segment
                            .y_advances()
                            .iter()
                            .map(|value| value * scale)
                            .collect(),
                        segment
                            .x_offsets()
                            .iter()
                            .map(|value| value * scale)
                            .collect(),
                        segment
                            .y_offsets()
                            .iter()
                            .map(|value| value * scale)
                            .collect(),
                        segment.clusters().to_vec(),
                        segment.break_after(),
                    )
                    .expect("scaling preserves multilingual segment invariants"),
                )
            }
            InlineItem::Marker(mut segment) => {
                scale_segment(&mut segment, scale);
                if indent < 0.0 {
                    segment.width = -indent;
                }
                InlineItem::Marker(segment)
            }
            other => other,
        })
        .collect()
}

fn scale_segment(segment: &mut TextSegment, scale: f64) {
    segment.font_size *= scale;
    segment.width *= scale;
    segment.ascent *= scale;
    segment.descent *= scale;
    segment.line_gap *= scale;
    segment.baseline_offset *= scale;
    for advance in &mut segment.advances {
        *advance *= scale;
    }
}

fn reduce_line_spacing(
    lines: &mut [LayoutLine],
    spacing: Option<&ResolvedTextSpacing>,
    reduction: f64,
) {
    if !matches!(spacing, Some(ResolvedTextSpacing::Percent(_))) {
        return;
    }
    let factor = (1.0 - reduction).max(0.0);
    for line in lines {
        line.height *= factor;
    }
}

fn anchor_lines(
    elements: &mut [PositionedElement],
    line_ranges: &[(usize, usize)],
    anchor: TextAnchor,
    spare_height: f64,
) {
    let line_count = line_ranges.len();
    for (line_index, &(start, end)) in line_ranges.iter().enumerate() {
        let offset = match anchor {
            TextAnchor::Top => 0.0,
            TextAnchor::Center => spare_height / 2.0,
            TextAnchor::Bottom => spare_height,
            TextAnchor::Justified if spare_height > 0.0 && line_count > 1 => {
                spare_height * line_index as f64 / (line_count - 1) as f64
            }
            TextAnchor::Distributed if spare_height > 0.0 && line_count > 1 => {
                spare_height * (line_index as f64 + 0.5) / line_count as f64
            }
            TextAnchor::Justified | TextAnchor::Distributed
                if spare_height > 0.0 && line_count == 1 =>
            {
                spare_height / 2.0
            }
            TextAnchor::Justified | TextAnchor::Distributed => 0.0,
        };
        for element in &mut elements[start..end] {
            translate_element_y(element, offset);
        }
    }
}

fn translate_element_y(element: &mut PositionedElement, offset: f64) {
    match element {
        PositionedElement::Text(run) => run.origin.y += offset,
        PositionedElement::MultilingualText(run) => run.origin.y += offset,
        PositionedElement::Line { start, end, .. } => {
            start.y += offset;
            end.y += offset;
        }
        PositionedElement::FilledRect { rect, .. }
        | PositionedElement::Image { rect, .. }
        | PositionedElement::LinkAnnotation { rect, .. } => rect.y += offset,
        PositionedElement::Group(group) => group.transform.f += offset,
        PositionedElement::Path(_) => {}
        _ => {}
    }
}

fn line_break_params(
    paragraph: &ResolvedParagraph,
    available_width: f64,
    wrap: bool,
) -> LineBreakParams {
    let (ind_first_line, ind_hanging) = if paragraph.indent < 0.0 {
        (0.0, -paragraph.indent)
    } else {
        (paragraph.indent, 0.0)
    };
    LineBreakParams {
        // Slide text has nothing floating beside it to flow around.
        line_prefix_widths: Vec::new(),
        line_suffix_widths: Vec::new(),
        available_width,
        ind_left: paragraph.left_margin,
        ind_right: paragraph.right_margin,
        ind_first_line,
        ind_hanging,
        tab_stops: Vec::new(),
        line_spacing: match paragraph.line_spacing.as_ref() {
            Some(ResolvedTextSpacing::Percent(factor)) => LineSpacing::Multiple(*factor),
            Some(ResolvedTextSpacing::Points(points)) => LineSpacing::Exact(*points),
            None => LineSpacing::Single,
        },
        jc: Some(paragraph_align(paragraph.alignment)),
        wrap,
    }
}

fn paragraph_align(alignment: ParagraphAlignment) -> Align {
    match alignment {
        ParagraphAlignment::Left => Align::Start,
        ParagraphAlignment::Center => Align::Center,
        ParagraphAlignment::Right => Align::End,
        ParagraphAlignment::Justified => Align::Justify,
        ParagraphAlignment::Distributed => Align::Distribute,
    }
}

fn first_run_font_size(paragraph: &ResolvedParagraph) -> f64 {
    paragraph
        .runs
        .iter()
        .find_map(|run| match run {
            ResolvedTextRun::Text { style, .. } | ResolvedTextRun::Field { style, .. } => {
                Some(style.font_size.unwrap_or(DEFAULT_FONT_SIZE))
            }
            ResolvedTextRun::Break => None,
        })
        .unwrap_or_else(|| paragraph.end_style.font_size.unwrap_or(DEFAULT_FONT_SIZE))
}

fn paragraph_spacing(spacing: Option<&ResolvedTextSpacing>, font_size: f64) -> f64 {
    match spacing {
        Some(ResolvedTextSpacing::Percent(factor)) => font_size * factor,
        Some(ResolvedTextSpacing::Points(points)) => *points,
        None => 0.0,
    }
}

fn emit_line_items(
    line: &LayoutLine,
    alignment: ParagraphAlignment,
    content_x: f64,
    baseline: f64,
    elements: &mut Vec<PositionedElement>,
) -> f64 {
    let element_start = elements.len();
    let remaining = line.available_width - line.width;
    let distribute = match alignment {
        ParagraphAlignment::Justified if !line.is_last => {
            let gaps = word_gap_count(&line.items);
            (gaps > 0).then_some((remaining.max(0.0) / gaps as f64, false))
        }
        ParagraphAlignment::Distributed => {
            let gaps = glyph_gap_count(&line.items);
            (gaps > 0).then_some((remaining.max(0.0) / gaps as f64, true))
        }
        _ => None,
    };
    let alignment_offset = match alignment {
        ParagraphAlignment::Center => remaining / 2.0,
        ParagraphAlignment::Right => remaining,
        _ => 0.0,
    };
    let mut x = content_x + line.indent_left + alignment_offset;
    let start_x = x;
    let mut distributed_gaps = distribute
        .filter(|(_, every_glyph)| *every_glyph)
        .map_or(0, |_| glyph_gap_count(&line.items));

    for item in &line.items {
        match item {
            LineItem::Text(segment) => {
                let (advances, effective_width) =
                    distributed_advances(segment, distribute, &mut distributed_gaps);
                emit_segment(
                    segment,
                    advances,
                    effective_width,
                    x,
                    baseline,
                    line.height,
                    elements,
                );
                x += effective_width;
            }
            LineItem::MultilingualText(segment) => {
                emit_multilingual_segment(segment, x, baseline, line.height, elements);
                x += segment.width();
            }
            LineItem::Marker(segment) => {
                emit_segment(
                    segment,
                    segment.advances.clone(),
                    segment.width,
                    x,
                    baseline,
                    line.height,
                    elements,
                );
                x += segment.width;
            }
            LineItem::Tab { width, leader } => {
                if let Some(segment) = leader {
                    emit_segment(
                        segment,
                        segment.advances.clone(),
                        *width,
                        x,
                        baseline,
                        line.height,
                        elements,
                    );
                }
                x += width;
            }
            LineItem::Image { width, .. } => x += width,
            LineItem::Group { width, .. } => x += width,
            _ => x += item.width(),
        }
    }
    let rich_positions = (element_start..elements.len())
        .filter(|index| matches!(elements[*index], PositionedElement::MultilingualText(_)))
        .collect::<Vec<_>>();
    let mut logical_runs = rich_positions
        .iter()
        .filter_map(|index| match &elements[*index] {
            PositionedElement::MultilingualText(run) => Some(run.clone()),
            _ => None,
        })
        .collect::<Vec<_>>();
    logical_runs.sort_by_key(|run| run.logical_index);
    for (position, run) in rich_positions.into_iter().zip(logical_runs) {
        elements[position] = PositionedElement::MultilingualText(run);
    }
    x - start_x
}

fn emit_segment(
    segment: &TextSegment,
    advances: Vec<f64>,
    effective_width: f64,
    x: f64,
    baseline: f64,
    line_height: f64,
    elements: &mut Vec<PositionedElement>,
) {
    if segment.text.is_empty() {
        return;
    }
    let adjusted_baseline = baseline - segment.baseline_offset;
    if let Some(color) = segment.highlight {
        elements.push(PositionedElement::FilledRect {
            rect: Rect {
                x,
                y: baseline - segment.ascent,
                width: effective_width,
                height: line_height,
            },
            color,
        });
    }
    elements.push(PositionedElement::Text(GlyphRun {
        origin: Point {
            x,
            y: adjusted_baseline,
        },
        font_id: segment.font_id,
        font_size: segment.font_size,
        glyph_ids: segment.glyph_ids.clone(),
        advances,
        text: segment.text.clone(),
        source: None,
        color: segment.color,
        bold: segment.bold,
        italic: segment.italic,
        field_kind: segment.field_kind,
        note: segment.note,
    }));

    if segment.underline.is_some() {
        let y = adjusted_baseline + segment.descent * 0.3;
        elements.push(PositionedElement::Line {
            start: Point { x, y },
            end: Point {
                x: x + effective_width,
                y,
            },
            width: segment.font_size / 18.0,
            color: segment.color,
            dash_pattern: None,
        });
    }
    if segment.strike {
        let y = adjusted_baseline - segment.ascent * 0.3;
        elements.push(PositionedElement::Line {
            start: Point { x, y },
            end: Point {
                x: x + effective_width,
                y,
            },
            width: segment.font_size / 24.0,
            color: segment.color,
            dash_pattern: None,
        });
    }
    if let Some(url) = &segment.hyperlink_url {
        elements.push(PositionedElement::LinkAnnotation {
            rect: Rect {
                x,
                y: baseline - segment.ascent,
                width: effective_width,
                height: line_height,
            },
            url: url.clone(),
        });
    }
}

fn emit_multilingual_segment(
    segment: &oxml_layout::MultilingualTextSegment,
    x: f64,
    baseline: f64,
    line_height: f64,
    elements: &mut Vec<PositionedElement>,
) {
    let base = segment.base();
    if base.text.is_empty() {
        return;
    }
    let adjusted_baseline = baseline - base.baseline_offset;
    if let Some(color) = base.highlight {
        elements.push(PositionedElement::FilledRect {
            rect: Rect {
                x,
                y: baseline - base.ascent,
                width: segment.width(),
                height: line_height,
            },
            color,
        });
    }
    elements.push(PositionedElement::MultilingualText(MultilingualGlyphRun {
        origin: Point {
            x,
            y: adjusted_baseline,
        },
        font_id: segment.font_id(),
        font_size: base.font_size,
        glyph_ids: segment.glyph_ids().to_vec(),
        x_advances: segment.x_advances().to_vec(),
        y_advances: segment.y_advances().to_vec(),
        x_offsets: segment.x_offsets().to_vec(),
        y_offsets: segment.y_offsets().to_vec(),
        clusters: segment.clusters().to_vec(),
        logical_text: segment.text().to_owned(),
        logical_index: segment.logical_index(),
        source: base.source,
        script: segment.script(),
        language: segment.language().map(str::to_owned),
        direction: segment.direction(),
        bidi_level: segment.bidi_level(),
        color: base.color,
        bold: base.bold,
        italic: base.italic,
        field_kind: base.field_kind,
        note: base.note,
    }));

    if base.underline.is_some() {
        let y = adjusted_baseline + base.descent * 0.3;
        elements.push(PositionedElement::Line {
            start: Point { x, y },
            end: Point {
                x: x + segment.width(),
                y,
            },
            width: base.font_size / 18.0,
            color: base.color,
            dash_pattern: None,
        });
    }
    if base.strike {
        let y = adjusted_baseline - base.ascent * 0.3;
        elements.push(PositionedElement::Line {
            start: Point { x, y },
            end: Point {
                x: x + segment.width(),
                y,
            },
            width: base.font_size / 24.0,
            color: base.color,
            dash_pattern: None,
        });
    }
    if let Some(url) = &base.hyperlink_url {
        elements.push(PositionedElement::LinkAnnotation {
            rect: Rect {
                x,
                y: baseline - base.ascent,
                width: segment.width(),
                height: line_height,
            },
            url: url.clone(),
        });
    }
}

fn word_gap_count(items: &[LineItem]) -> usize {
    items
        .iter()
        .filter_map(text_segment)
        .map(|segment| {
            segment
                .text
                .chars()
                .filter(|character| character.is_whitespace())
                .count()
        })
        .sum()
}

fn glyph_gap_count(items: &[LineItem]) -> usize {
    items
        .iter()
        .filter_map(text_segment)
        .map(|segment| segment.advances.len())
        .sum::<usize>()
        .saturating_sub(1)
}

fn text_segment(item: &LineItem) -> Option<&TextSegment> {
    match item {
        LineItem::Text(segment) => Some(segment),
        LineItem::MultilingualText(segment) => Some(segment.base()),
        LineItem::Marker(_)
        | LineItem::Tab { .. }
        | LineItem::Image { .. }
        | LineItem::Group { .. } => None,
        _ => None,
    }
}

fn distributed_advances(
    segment: &TextSegment,
    distribution: Option<(f64, bool)>,
    distributed_gaps: &mut usize,
) -> (Vec<f64>, f64) {
    let Some((extra, every_glyph)) = distribution else {
        return (segment.advances.clone(), segment.width);
    };
    let mut advances = segment.advances.clone();
    let mut additions = 0_usize;
    if every_glyph {
        for advance in &mut advances {
            if *distributed_gaps == 0 {
                break;
            }
            *advance += extra;
            additions += 1;
            *distributed_gaps -= 1;
        }
    } else {
        let characters = segment.text.chars().collect::<Vec<_>>();
        if characters.len() == advances.len() {
            for (character, advance) in characters.into_iter().zip(&mut advances) {
                if character.is_whitespace() {
                    *advance += extra;
                    additions += 1;
                }
            }
        } else if let Some(last) = advances.last_mut() {
            let whitespace_count = characters
                .into_iter()
                .filter(|character| character.is_whitespace())
                .count();
            if whitespace_count > 0 {
                *last += extra * whitespace_count as f64;
                additions = whitespace_count;
            }
        }
    }
    (advances, segment.width + extra * additions as f64)
}

pub(super) fn content_box(shape: &ResolvedShape, text: &ResolvedTextBody) -> Rect {
    let boundary = text_boundary(shape);
    Rect {
        x: boundary.x + text.insets.left,
        y: boundary.y + text.insets.top,
        width: (boundary.width - text.insets.left - text.insets.right).max(0.0),
        height: (boundary.height - text.insets.top - text.insets.bottom).max(0.0),
    }
}

fn text_boundary(shape: &ResolvedShape) -> Rect {
    match &shape.geometry {
        rpptx_layout::ResolvedGeometry::Custom {
            text_rect: Some(rect),
            ..
        } => *rect,
        _ => Rect {
            x: 0.0,
            y: 0.0,
            width: shape.bounds.width,
            height: shape.bounds.height,
        },
    }
}

pub(super) fn oriented_content_box(
    content: Rect,
    direction: TextDirection,
) -> (Rect, Option<oxml_layout::Transform>) {
    let degrees = match direction {
        TextDirection::Horizontal => return (content, None),
        TextDirection::Vertical
        | TextDirection::EastAsianVertical
        | TextDirection::WordArtVertical => 90.0,
        TextDirection::Vertical270
        | TextDirection::MongolianVertical
        | TextDirection::WordArtVerticalRtl => -90.0,
    };
    let center_x = content.x + content.width / 2.0;
    let center_y = content.y + content.height / 2.0;
    (
        Rect {
            x: center_x - content.height / 2.0,
            y: center_y - content.width / 2.0,
            width: content.height,
            height: content.width,
        },
        Some(oxml_layout::Transform::rotate_about(
            degrees, center_x, center_y,
        )),
    )
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::collections::HashMap;

    use oxml_layout::{Color, Paint, Transform};
    use rpptx_layout::{
        ParagraphAlignment, ResolvedAutofit, ResolvedBullet, ResolvedBulletSize, ResolvedContent,
        ResolvedGeometry, ResolvedRunStyle, ResolvedSlide, ResolvedTextRun, ResolvedTextSpacing,
        TextAnchor, TextDirection, TextInsets,
    };

    use crate::{RenderInput, RenderInputError, layout_presentation, layout_slide_with_fonts};

    fn style() -> ResolvedRunStyle {
        ResolvedRunStyle {
            font_size: Some(22.0),
            bold: true,
            italic: true,
            underline: true,
            strike: true,
            spacing: Some(1.5),
            baseline: Some(0.25),
            fill: Some(Paint::Solid(Color {
                r: 0.8,
                g: 0.1,
                b: 0.2,
                a: 1.0,
            })),
            latin_typeface: Some("Carlito".to_owned()),
            ..ResolvedRunStyle::default()
        }
    }

    fn shape(bounds: Rect, geometry: ResolvedGeometry) -> ResolvedShape {
        ResolvedShape {
            group_transform: Transform::IDENTITY,
            bounds,
            rotation_deg: 0.0,
            flip_h: false,
            flip_v: false,
            geometry,
            fill: None,
            image_fill: None,
            line: None,
            head_end: None,
            tail_end: None,
            shadow: None,
            content: ResolvedContent::None,
            unsupported: None,
        }
    }

    fn text_body(insets: TextInsets) -> ResolvedTextBody {
        ResolvedTextBody {
            insets,
            anchor: TextAnchor::Top,
            wrap: true,
            vertical: TextDirection::Horizontal,
            space_first_last_paragraph: false,
            autofit: ResolvedAutofit::None,
            paragraphs: Vec::new(),
        }
    }

    #[test]
    fn preset_text_rectangle_minus_unequal_insets_produces_the_computed_content_box() {
        let shape = shape(
            Rect {
                x: 30.0,
                y: 40.0,
                width: 120.0,
                height: 70.0,
            },
            ResolvedGeometry::Custom {
                paths: Vec::new(),
                text_rect: Some(Rect {
                    x: 10.0,
                    y: 15.0,
                    width: 80.0,
                    height: 40.0,
                }),
            },
        );
        let text = text_body(TextInsets {
            left: 3.0,
            top: 5.0,
            right: 7.0,
            bottom: 11.0,
        });

        assert_eq!(
            content_box(&shape, &text),
            Rect {
                x: 13.0,
                y: 20.0,
                width: 70.0,
                height: 24.0,
            }
        );
    }

    #[test]
    fn missing_text_rectangle_falls_back_to_local_shape_bounds() {
        let text = text_body(TextInsets {
            left: 4.0,
            top: 6.0,
            right: 8.0,
            bottom: 10.0,
        });

        for geometry in [
            ResolvedGeometry::Rectangle,
            ResolvedGeometry::BoundsFallback,
        ] {
            let shape = shape(
                Rect {
                    x: 30.0,
                    y: 40.0,
                    width: 120.0,
                    height: 70.0,
                },
                geometry,
            );
            assert_eq!(
                content_box(&shape, &text),
                Rect {
                    x: 4.0,
                    y: 6.0,
                    width: 108.0,
                    height: 54.0,
                }
            );
        }
    }

    #[test]
    fn insets_larger_than_the_text_rectangle_do_not_create_negative_extents() {
        let shape = shape(
            Rect {
                x: 0.0,
                y: 0.0,
                width: 30.0,
                height: 20.0,
            },
            ResolvedGeometry::Custom {
                paths: Vec::new(),
                text_rect: Some(Rect {
                    x: 1.0,
                    y: 2.0,
                    width: 10.0,
                    height: 8.0,
                }),
            },
        );
        let text = text_body(TextInsets {
            left: 8.0,
            top: 9.0,
            right: 7.0,
            bottom: 6.0,
        });

        let content = content_box(&shape, &text);
        assert_eq!(
            content,
            Rect {
                x: 9.0,
                y: 11.0,
                width: 0.0,
                height: 0.0,
            }
        );
        assert!(content.x.is_finite());
        assert!(content.y.is_finite());
        assert!(content.width.is_finite());
        assert!(content.height.is_finite());
    }

    #[test]
    fn vertical_text_uses_a_transposed_box_and_rotated_group() {
        let content = Rect {
            x: 10.0,
            y: 20.0,
            width: 80.0,
            height: 30.0,
        };

        let (transposed, transform) = oriented_content_box(content, TextDirection::Vertical);

        assert_eq!(
            transposed,
            Rect {
                x: 35.0,
                y: -5.0,
                width: 30.0,
                height: 80.0,
            }
        );
        let transform = transform.expect("vertical transform");
        assert_close(transform.a, 0.0);
        assert_close(transform.b, 1.0);
        assert_close(transform.c, -1.0);
        assert_close(transform.d, 0.0);
        assert_close(transform.e, 85.0);
        assert_close(transform.f, -15.0);

        let text = ResolvedTextBody {
            vertical: TextDirection::Vertical,
            paragraphs: vec![ResolvedParagraph {
                runs: vec![ResolvedTextRun::Text {
                    text: "Visible".to_owned(),
                    style: ResolvedRunStyle {
                        font_size: Some(12.0),
                        latin_typeface: Some("Carlito".to_owned()),
                        ..ResolvedRunStyle::default()
                    },
                }],
                ..ResolvedParagraph::default()
            }],
            ..text_body(TextInsets::default())
        };
        let mut text_shape = shape(
            Rect {
                x: 0.0,
                y: 0.0,
                width: 80.0,
                height: 30.0,
            },
            ResolvedGeometry::Rectangle,
        );
        text_shape.content = ResolvedContent::Text(text);
        let input = RenderInput {
            slides: vec![ResolvedSlide {
                size: (80.0, 30.0),
                background: None,
                shapes: vec![text_shape],
                diagnostics: Vec::new(),
            }],
            media: HashMap::new(),
            fonts: Vec::new(),
            metadata: None,
        };
        let mut fonts = FontManager::new_deterministic().expect("deterministic fonts");

        let page = layout_slide_with_fonts(&input, 0, &mut fonts).expect("layout vertical text");
        let oxml_layout::PositionedElement::Group(shape_group) = &page.elements[0] else {
            panic!("shape should lower to a group");
        };
        let text_group = shape_group
            .children
            .iter()
            .find_map(|element| match element {
                oxml_layout::PositionedElement::Group(group) => Some(group),
                _ => None,
            })
            .expect("vertical text should lower to a nested group");
        assert_close(text_group.transform.b, 1.0);
        assert_close(text_group.transform.c, -1.0);
        assert_eq!(glyph_runs(&text_group.children)[0].text, "Visible");
    }

    #[test]
    fn vertical_270_uses_the_opposite_quarter_turn() {
        let content = Rect {
            x: 10.0,
            y: 20.0,
            width: 80.0,
            height: 30.0,
        };

        let (_, vertical) = oriented_content_box(content, TextDirection::Vertical);
        let (_, vertical_270) = oriented_content_box(content, TextDirection::Vertical270);

        let vertical = vertical.expect("vertical transform");
        let vertical_270 = vertical_270.expect("vertical-270 transform");
        assert_close(vertical.b, 1.0);
        assert_close(vertical.c, -1.0);
        assert_close(vertical_270.b, -1.0);
        assert_close(vertical_270.c, 1.0);
        assert_eq!(
            vertical.apply(Point { x: 50.0, y: 20.0 }),
            vertical_270.apply(Point { x: 50.0, y: 50.0 })
        );
    }

    #[test]
    fn unsupported_vertical_variants_use_the_documented_quarter_turns() {
        let content = Rect {
            x: 10.0,
            y: 20.0,
            width: 80.0,
            height: 30.0,
        };

        for direction in [
            TextDirection::EastAsianVertical,
            TextDirection::WordArtVertical,
        ] {
            let (_, transform) = oriented_content_box(content, direction);
            let transform = transform.expect("vertical fallback transform");
            assert_close(transform.b, 1.0);
            assert_close(transform.c, -1.0);
        }
        for direction in [
            TextDirection::MongolianVertical,
            TextDirection::WordArtVerticalRtl,
        ] {
            let (_, transform) = oriented_content_box(content, direction);
            let transform = transform.expect("vertical-270 fallback transform");
            assert_close(transform.b, -1.0);
            assert_close(transform.c, 1.0);
        }
    }

    #[test]
    fn resolved_runs_emit_glyph_items_with_concrete_style_and_break_boundaries() {
        let paragraph = ResolvedParagraph {
            runs: vec![
                ResolvedTextRun::Text {
                    text: "Alpha".to_owned(),
                    style: style(),
                },
                ResolvedTextRun::Break,
                ResolvedTextRun::Field {
                    text: "42".to_owned(),
                    field_type: Some("slidenum".to_owned()),
                    style: style(),
                },
            ],
            ..ResolvedParagraph::default()
        };
        let mut fonts = FontManager::new_deterministic().expect("deterministic fonts");

        let items = inline_items(&mut fonts, &paragraph).expect("shape resolved runs");

        assert_eq!(items.len(), 3);
        let InlineItem::Text(first) = &items[0] else {
            panic!("text run should become a text item");
        };
        assert_eq!(first.text, "Alpha");
        assert_eq!(first.font_size, 22.0);
        assert_eq!(first.color.r, 0.8);
        assert!(first.bold);
        assert!(first.italic);
        assert!(first.underline.is_some());
        assert!(first.strike);
        assert_eq!(first.baseline_offset, 5.5);
        assert!(first.glyph_ids.iter().any(|glyph| *glyph != 0));
        assert!(matches!(items[1], InlineItem::LineBreak));
        let InlineItem::Text(field) = &items[2] else {
            panic!("field should become a text item");
        };
        assert_eq!(field.text, "1");
        assert_eq!(field.font_size, 22.0);
        assert_eq!(field.field_kind, Some(FieldKind::Page));
    }

    #[test]
    fn script_specific_text_selects_its_resolved_concrete_typeface() {
        let style = ResolvedRunStyle {
            latin_typeface: Some("Latin face".to_owned()),
            east_asian_typeface: Some("East Asian face".to_owned()),
            complex_script_typeface: Some("Complex face".to_owned()),
            symbol_typeface: Some("Symbol face".to_owned()),
            ..ResolvedRunStyle::default()
        };

        assert_eq!(typeface_for_text(&style, "Latin"), Some("Latin face"));
        assert_eq!(typeface_for_text(&style, "漢字"), Some("East Asian face"));
        assert_eq!(typeface_for_text(&style, "مرحبا"), Some("Complex face"));
        assert_eq!(typeface_for_text(&style, "\u{F0B7}"), Some("Symbol face"));
    }

    #[test]
    fn missing_run_size_and_fill_use_visible_renderer_defaults() {
        let paragraph = ResolvedParagraph {
            runs: vec![ResolvedTextRun::Text {
                text: "Visible".to_owned(),
                style: ResolvedRunStyle::default(),
            }],
            ..ResolvedParagraph::default()
        };
        let mut fonts = FontManager::new_deterministic().expect("deterministic fonts");

        let items = inline_items(&mut fonts, &paragraph).expect("shape default run");

        let InlineItem::Text(segment) = &items[0] else {
            panic!("text run should become a text item");
        };
        assert_eq!(segment.font_size, 18.0);
        assert_eq!(segment.color, Color::BLACK);
    }

    #[test]
    fn text_shaping_failures_return_a_render_error_without_panicking() {
        let text = ResolvedTextBody {
            insets: TextInsets::default(),
            anchor: TextAnchor::Top,
            wrap: true,
            vertical: TextDirection::Horizontal,
            space_first_last_paragraph: false,
            autofit: ResolvedAutofit::None,
            paragraphs: vec![ResolvedParagraph {
                runs: vec![ResolvedTextRun::Text {
                    text: "No font".to_owned(),
                    style: ResolvedRunStyle::default(),
                }],
                ..ResolvedParagraph::default()
            }],
        };
        let mut text_shape = shape(
            Rect {
                x: 0.0,
                y: 0.0,
                width: 100.0,
                height: 50.0,
            },
            ResolvedGeometry::Rectangle,
        );
        text_shape.content = ResolvedContent::Text(text);
        let input = RenderInput {
            slides: vec![ResolvedSlide {
                size: (100.0, 50.0),
                background: None,
                shapes: vec![text_shape],
                diagnostics: Vec::new(),
            }],
            media: HashMap::new(),
            fonts: Vec::new(),
            metadata: None,
        };
        let mut fonts = FontManager::new_with_fonts(Vec::new());

        let error =
            layout_slide_with_fonts(&input, 0, &mut fonts).expect_err("missing font should fail");

        assert!(matches!(error, RenderInputError::TextLayout { .. }));
    }

    #[test]
    fn presentation_layout_collects_every_font_used_by_shape_text() {
        let text = ResolvedTextBody {
            insets: TextInsets::default(),
            anchor: TextAnchor::Top,
            wrap: true,
            vertical: TextDirection::Horizontal,
            space_first_last_paragraph: false,
            autofit: ResolvedAutofit::None,
            paragraphs: vec![ResolvedParagraph {
                runs: vec![
                    ResolvedTextRun::Text {
                        text: "Regular".to_owned(),
                        style: ResolvedRunStyle {
                            latin_typeface: Some("Carlito".to_owned()),
                            ..ResolvedRunStyle::default()
                        },
                    },
                    ResolvedTextRun::Text {
                        text: "Bold".to_owned(),
                        style: ResolvedRunStyle {
                            bold: true,
                            latin_typeface: Some("Carlito".to_owned()),
                            ..ResolvedRunStyle::default()
                        },
                    },
                ],
                ..ResolvedParagraph::default()
            }],
        };
        let mut text_shape = shape(
            Rect {
                x: 0.0,
                y: 0.0,
                width: 100.0,
                height: 50.0,
            },
            ResolvedGeometry::Rectangle,
        );
        text_shape.content = ResolvedContent::Text(text);
        let input = RenderInput {
            slides: vec![ResolvedSlide {
                size: (100.0, 50.0),
                background: None,
                shapes: vec![text_shape],
                diagnostics: Vec::new(),
            }],
            media: HashMap::new(),
            fonts: Vec::new(),
            metadata: None,
        };

        let layout = layout_presentation(&input).expect("layout presentation text");

        assert_eq!(layout.fonts.len(), 2);
        assert!(layout.fonts.iter().any(|font| !font.bold));
        assert!(layout.fonts.iter().any(|font| font.bold));
    }

    #[test]
    fn paragraphs_stack_wrapped_lines_with_spacing_and_alignment() {
        let mut fonts = FontManager::new_deterministic().expect("deterministic fonts");
        let body = ResolvedTextBody {
            wrap: true,
            paragraphs: vec![
                ResolvedParagraph {
                    alignment: ParagraphAlignment::Center,
                    line_spacing: Some(ResolvedTextSpacing::Points(20.0)),
                    space_after: Some(ResolvedTextSpacing::Points(5.0)),
                    runs: vec![ResolvedTextRun::Text {
                        text: "one two".to_owned(),
                        style: ResolvedRunStyle {
                            font_size: Some(12.0),
                            latin_typeface: Some("Carlito".to_owned()),
                            ..ResolvedRunStyle::default()
                        },
                    }],
                    ..ResolvedParagraph::default()
                },
                ResolvedParagraph {
                    space_before: Some(ResolvedTextSpacing::Points(3.0)),
                    line_spacing: Some(ResolvedTextSpacing::Points(20.0)),
                    runs: vec![ResolvedTextRun::Text {
                        text: "three".to_owned(),
                        style: ResolvedRunStyle {
                            font_size: Some(12.0),
                            latin_typeface: Some("Carlito".to_owned()),
                            ..ResolvedRunStyle::default()
                        },
                    }],
                    ..ResolvedParagraph::default()
                },
            ],
            ..text_body(TextInsets::default())
        };
        let content = Rect {
            x: 10.0,
            y: 7.0,
            width: 35.0,
            height: 20.0,
        };

        let stacked = stack_text(&mut fonts, content, &body).expect("stack text");
        let runs = glyph_runs(&stacked.elements);

        assert_eq!(runs.len(), 3);
        assert_close(runs[1].origin.y - runs[0].origin.y, 20.0);
        assert_close(runs[2].origin.y - runs[1].origin.y, 28.0);
        assert!(runs[0].origin.x > content.x);
        assert!(runs[1].origin.x > content.x);
        assert_close(stacked.height, 68.0);
        assert!(stacked.width <= content.width);
    }

    #[test]
    fn wrap_none_breaks_only_at_explicit_line_breaks() {
        let mut fonts = FontManager::new_deterministic().expect("deterministic fonts");
        let body = ResolvedTextBody {
            wrap: false,
            paragraphs: vec![ResolvedParagraph {
                line_spacing: Some(ResolvedTextSpacing::Points(18.0)),
                runs: vec![
                    ResolvedTextRun::Text {
                        text: "alpha beta gamma".to_owned(),
                        style: ResolvedRunStyle {
                            font_size: Some(12.0),
                            latin_typeface: Some("Carlito".to_owned()),
                            ..ResolvedRunStyle::default()
                        },
                    },
                    ResolvedTextRun::Break,
                    ResolvedTextRun::Text {
                        text: "delta".to_owned(),
                        style: ResolvedRunStyle {
                            font_size: Some(12.0),
                            latin_typeface: Some("Carlito".to_owned()),
                            ..ResolvedRunStyle::default()
                        },
                    },
                ],
                ..ResolvedParagraph::default()
            }],
            ..text_body(TextInsets::default())
        };

        let stacked = stack_text(
            &mut fonts,
            Rect {
                x: 0.0,
                y: 0.0,
                width: 10.0,
                height: 10.0,
            },
            &body,
        )
        .expect("stack unwrapped text");
        let runs = glyph_runs(&stacked.elements);

        let baselines = distinct_baselines(&runs);
        assert_eq!(baselines.len(), 2);
        assert_close(baselines[1] - baselines[0], 18.0);
        assert!(stacked.width > 10.0);
    }

    #[test]
    fn point_and_percentage_paragraph_spacing_produce_computed_baselines() {
        let mut fonts = FontManager::new_deterministic().expect("deterministic fonts");
        let run = || ResolvedTextRun::Text {
            text: "line".to_owned(),
            style: ResolvedRunStyle {
                font_size: Some(20.0),
                latin_typeface: Some("Carlito".to_owned()),
                ..ResolvedRunStyle::default()
            },
        };
        let body = ResolvedTextBody {
            paragraphs: vec![
                ResolvedParagraph {
                    line_spacing: Some(ResolvedTextSpacing::Points(24.0)),
                    space_after: Some(ResolvedTextSpacing::Percent(0.5)),
                    runs: vec![run()],
                    ..ResolvedParagraph::default()
                },
                ResolvedParagraph {
                    line_spacing: Some(ResolvedTextSpacing::Points(24.0)),
                    space_before: Some(ResolvedTextSpacing::Points(3.0)),
                    runs: vec![run()],
                    ..ResolvedParagraph::default()
                },
            ],
            ..text_body(TextInsets::default())
        };

        let stacked = stack_text(
            &mut fonts,
            Rect {
                x: 0.0,
                y: 2.0,
                width: 100.0,
                height: 10.0,
            },
            &body,
        )
        .expect("stack spaced paragraphs");
        let runs = glyph_runs(&stacked.elements);

        assert_eq!(runs.len(), 2);
        assert_close(runs[1].origin.y - runs[0].origin.y, 37.0);
        assert_close(stacked.height, 61.0);
    }

    #[test]
    fn text_and_marker_items_share_one_baseline_emitter() {
        let mut fonts = FontManager::new_deterministic().expect("deterministic fonts");
        let segment = shape_run(
            &mut fonts,
            "x",
            &ResolvedRunStyle {
                font_size: Some(12.0),
                latin_typeface: Some("Carlito".to_owned()),
                ..ResolvedRunStyle::default()
            },
        )
        .expect("shape segment");
        let mut decorated = segment.clone();
        decorated.underline = Some(Underline::Single);
        decorated.strike = true;
        let line = oxml_layout::LayoutLine {
            items: vec![
                oxml_layout::LineItem::Marker(decorated.clone()),
                oxml_layout::LineItem::Text(decorated.clone()),
            ],
            width: decorated.width * 2.0,
            ascent: decorated.ascent,
            descent: decorated.descent,
            line_gap: decorated.line_gap,
            height: decorated.ascent + decorated.descent + decorated.line_gap,
            indent_left: 0.0,
            available_width: 100.0,
            is_last: true,
            page_break_after: false,
        };
        let mut elements = Vec::new();

        emit_line_items(
            &line,
            ParagraphAlignment::Left,
            4.0,
            9.0 + line.ascent,
            &mut elements,
        );
        let runs = glyph_runs(&elements);

        assert_eq!(runs.len(), 2);
        assert_close(runs[0].origin.y, runs[1].origin.y);
        assert_close(runs[1].origin.x - runs[0].origin.x, decorated.width);
        assert_eq!(
            elements
                .iter()
                .filter(|element| matches!(element, PositionedElement::Line { .. }))
                .count(),
            4
        );
    }

    #[test]
    fn resolved_paragraph_fields_map_to_shared_line_break_parameters() {
        let paragraph = ResolvedParagraph {
            left_margin: 12.0,
            right_margin: 7.0,
            indent: -5.0,
            alignment: ParagraphAlignment::Right,
            line_spacing: Some(ResolvedTextSpacing::Percent(1.25)),
            ..ResolvedParagraph::default()
        };

        let params = line_break_params(&paragraph, 80.0, false);

        assert_eq!(params.available_width, 80.0);
        assert_eq!(params.ind_left, 12.0);
        assert_eq!(params.ind_right, 7.0);
        assert_eq!(params.ind_first_line, 0.0);
        assert_eq!(params.ind_hanging, 5.0);
        assert_eq!(params.line_spacing, LineSpacing::Multiple(1.25));
        assert_eq!(params.jc, Some(Align::End));
        assert!(!params.wrap);
    }

    #[test]
    fn percentage_line_spacing_uses_effective_first_run_font_size() {
        let mut fonts = FontManager::new_deterministic().expect("deterministic fonts");
        let body = ResolvedTextBody {
            paragraphs: vec![ResolvedParagraph {
                line_spacing: Some(ResolvedTextSpacing::Percent(1.2)),
                runs: vec![ResolvedTextRun::Text {
                    text: "line".to_owned(),
                    style: ResolvedRunStyle {
                        font_size: Some(20.0),
                        latin_typeface: Some("Carlito".to_owned()),
                        ..ResolvedRunStyle::default()
                    },
                }],
                ..ResolvedParagraph::default()
            }],
            ..text_body(TextInsets::default())
        };
        let stacked = stack_text(
            &mut fonts,
            Rect {
                x: 0.0,
                y: 0.0,
                width: 100.0,
                height: 10.0,
            },
            &body,
        )
        .expect("stack percentage-spaced line");

        assert_close(stacked.height, 24.0);
    }

    #[test]
    fn empty_paragraph_uses_end_style_metrics_and_size_for_spacing() {
        let mut fonts = FontManager::new_deterministic().expect("deterministic fonts");
        let end_style = ResolvedRunStyle {
            font_size: Some(30.0),
            latin_typeface: Some("Carlito".to_owned()),
            ..ResolvedRunStyle::default()
        };
        let expected = shape_run(&mut fonts, "", &end_style).expect("shape empty metrics");
        let body = ResolvedTextBody {
            space_first_last_paragraph: true,
            paragraphs: vec![ResolvedParagraph {
                space_before: Some(ResolvedTextSpacing::Percent(0.5)),
                end_style,
                ..ResolvedParagraph::default()
            }],
            ..text_body(TextInsets::default())
        };

        let stacked = stack_text(
            &mut fonts,
            Rect {
                x: 0.0,
                y: 0.0,
                width: 100.0,
                height: 100.0,
            },
            &body,
        )
        .expect("stack empty paragraph");

        assert!(stacked.elements.is_empty());
        assert_close(
            stacked.height,
            expected.ascent + expected.descent + expected.line_gap + 15.0,
        );
    }

    #[test]
    fn empty_paragraph_suppresses_bullet_and_keeps_end_metrics() {
        let mut fonts = FontManager::new_deterministic().expect("deterministic fonts");
        let end_style = ResolvedRunStyle {
            font_size: Some(30.0),
            latin_typeface: Some("Carlito".to_owned()),
            ..ResolvedRunStyle::default()
        };
        let expected = shape_run(&mut fonts, "", &end_style).expect("shape empty metrics");
        let body = ResolvedTextBody {
            paragraphs: vec![ResolvedParagraph {
                bullet: Some(ResolvedBullet::Character {
                    character: "•".to_owned(),
                    font: Some("Carlito".to_owned()),
                    color: None,
                    size: None,
                }),
                end_style,
                ..ResolvedParagraph::default()
            }],
            ..text_body(TextInsets::default())
        };

        let stacked = stack_text(&mut fonts, test_content_box(100.0), &body)
            .expect("stack empty bullet paragraph");
        assert!(stacked.elements.is_empty());
        assert_close(
            stacked.height,
            expected.ascent + expected.descent + expected.line_gap,
        );
    }

    #[test]
    fn all_caps_transforms_text_before_font_resolution_and_shaping() {
        let mut fonts = FontManager::new_deterministic().expect("deterministic fonts");
        let caps_style = ResolvedRunStyle {
            font_size: Some(18.0),
            all_caps: true,
            latin_typeface: Some("Carlito".to_owned()),
            ..ResolvedRunStyle::default()
        };
        let normal_style = ResolvedRunStyle {
            all_caps: false,
            ..caps_style.clone()
        };

        let caps = shape_run(&mut fonts, "Straße", &caps_style).expect("shape all caps");
        let expected = shape_run(&mut fonts, "STRASSE", &normal_style).expect("shape uppercase");

        assert_eq!(caps.text, "STRASSE");
        assert_eq!(caps.font_id, expected.font_id);
        assert_eq!(caps.glyph_ids, expected.glyph_ids);
        assert_eq!(caps.advances, expected.advances);
        assert_close(caps.width, expected.width);
    }

    #[test]
    fn horizontal_alignment_and_distribution_use_the_available_line_width() {
        let segment = TextSegment {
            text: "ab".to_owned(),
            direction: oxml_layout::TextDirection::Auto,
            source: None,
            font_id: oxml_layout::FontId(0),
            font_size: 12.0,
            glyph_ids: vec![1, 2],
            advances: vec![5.0, 5.0],
            width: 10.0,
            ascent: 8.0,
            descent: 2.0,
            line_gap: 0.0,
            color: Color::BLACK,
            bold: false,
            italic: false,
            underline: None,
            strike: false,
            dstrike: false,
            highlight: None,
            baseline_offset: 0.0,
            hyperlink_url: None,
            field_kind: None,
            note: None,
        };
        let line = oxml_layout::LayoutLine {
            items: vec![
                oxml_layout::LineItem::Text(segment.clone()),
                oxml_layout::LineItem::Text(segment.clone()),
            ],
            width: 20.0,
            ascent: 8.0,
            descent: 2.0,
            line_gap: 0.0,
            height: 10.0,
            indent_left: 10.0,
            available_width: 80.0,
            is_last: true,
            page_break_after: false,
        };

        for (alignment, expected_x) in [
            (ParagraphAlignment::Left, 14.0),
            (ParagraphAlignment::Center, 44.0),
            (ParagraphAlignment::Right, 74.0),
        ] {
            let mut elements = Vec::new();
            let width = emit_line_items(&line, alignment, 4.0, 20.0, &mut elements);
            let runs = glyph_runs(&elements);
            assert_close(runs[0].origin.x, expected_x);
            assert_close(width, 20.0);
        }

        let mut elements = Vec::new();
        let width = emit_line_items(
            &line,
            ParagraphAlignment::Distributed,
            4.0,
            20.0,
            &mut elements,
        );
        let runs = glyph_runs(&elements);
        assert_close(runs[0].origin.x, 14.0);
        assert_close(width, 80.0);
        assert_close(
            runs.iter().flat_map(|run| run.advances.iter()).sum::<f64>(),
            80.0,
        );

        let mut justified_segment = segment;
        justified_segment.text = "a b".to_owned();
        justified_segment.glyph_ids = vec![1, 3, 2];
        justified_segment.advances = vec![5.0, 5.0, 5.0];
        justified_segment.width = 15.0;
        let mut justified_line = oxml_layout::LayoutLine {
            items: vec![oxml_layout::LineItem::Text(justified_segment)],
            width: 15.0,
            ascent: 8.0,
            descent: 2.0,
            line_gap: 0.0,
            height: 10.0,
            indent_left: 0.0,
            available_width: 45.0,
            is_last: false,
            page_break_after: false,
        };
        let mut elements = Vec::new();
        let width = emit_line_items(
            &justified_line,
            ParagraphAlignment::Justified,
            0.0,
            20.0,
            &mut elements,
        );
        assert_close(width, 45.0);
        assert_close(glyph_runs(&elements)[0].advances.iter().sum(), 45.0);

        justified_line.is_last = true;
        elements.clear();
        let width = emit_line_items(
            &justified_line,
            ParagraphAlignment::Justified,
            0.0,
            20.0,
            &mut elements,
        );
        assert_close(width, 15.0);
        assert_close(glyph_runs(&elements)[0].advances.iter().sum(), 15.0);
    }

    #[test]
    fn justified_ligature_text_still_expands_its_word_gap() {
        let mut fonts = FontManager::new_deterministic().expect("deterministic fonts");
        let segment = shape_run(
            &mut fonts,
            "office space",
            &ResolvedRunStyle {
                font_size: Some(12.0),
                latin_typeface: Some("Carlito".to_owned()),
                ..ResolvedRunStyle::default()
            },
        )
        .expect("shape ligature-bearing text");
        assert_ne!(
            segment.text.chars().count(),
            segment.advances.len(),
            "fixture must exercise a non-1:1 shaped run"
        );
        let available_width = segment.width + 30.0;
        let line = LayoutLine {
            items: vec![LineItem::Text(segment)],
            width: available_width - 30.0,
            ascent: 8.0,
            descent: 2.0,
            line_gap: 0.0,
            height: 10.0,
            indent_left: 0.0,
            available_width,
            is_last: false,
            page_break_after: false,
        };
        let mut elements = Vec::new();

        let width = emit_line_items(
            &line,
            ParagraphAlignment::Justified,
            0.0,
            20.0,
            &mut elements,
        );

        assert_close(width, available_width);
        assert_close(
            glyph_runs(&elements)[0].advances.iter().sum(),
            available_width,
        );
    }

    #[test]
    fn top_center_and_bottom_anchors_use_zero_half_and_full_spare_height() {
        let mut fonts = FontManager::new_deterministic().expect("deterministic fonts");
        let paragraph = ResolvedParagraph {
            line_spacing: Some(ResolvedTextSpacing::Points(20.0)),
            runs: vec![ResolvedTextRun::Text {
                text: "anchor".to_owned(),
                style: ResolvedRunStyle {
                    font_size: Some(12.0),
                    latin_typeface: Some("Carlito".to_owned()),
                    ..ResolvedRunStyle::default()
                },
            }],
            ..ResolvedParagraph::default()
        };
        let content = Rect {
            x: 0.0,
            y: 7.0,
            width: 100.0,
            height: 100.0,
        };
        let mut baselines = Vec::new();

        for anchor in [TextAnchor::Top, TextAnchor::Center, TextAnchor::Bottom] {
            let body = ResolvedTextBody {
                anchor,
                paragraphs: vec![paragraph.clone()],
                ..text_body(TextInsets::default())
            };
            let stacked = stack_text(&mut fonts, content, &body).expect("stack anchored text");
            baselines.push(glyph_runs(&stacked.elements)[0].origin.y);
        }

        let expected = shape_run(
            &mut fonts,
            "anchor",
            &ResolvedRunStyle {
                font_size: Some(12.0),
                latin_typeface: Some("Carlito".to_owned()),
                ..ResolvedRunStyle::default()
            },
        )
        .expect("shape anchor metrics");
        let half_leading = (20.0 - expected.ascent - expected.descent) / 2.0;
        assert_close(baselines[0], content.y + expected.ascent + half_leading);
        assert_close(baselines[1] - baselines[0], 40.0);
        assert_close(baselines[2] - baselines[0], 80.0);
    }

    #[test]
    fn overflowing_anchored_text_remains_visible_without_a_clip() {
        let text = ResolvedTextBody {
            anchor: TextAnchor::Bottom,
            wrap: false,
            paragraphs: vec![ResolvedParagraph {
                line_spacing: Some(ResolvedTextSpacing::Points(20.0)),
                runs: vec![ResolvedTextRun::Text {
                    text: "visible overflow".to_owned(),
                    style: ResolvedRunStyle {
                        font_size: Some(12.0),
                        latin_typeface: Some("Carlito".to_owned()),
                        ..ResolvedRunStyle::default()
                    },
                }],
                ..ResolvedParagraph::default()
            }],
            ..text_body(TextInsets::default())
        };
        let mut text_shape = shape(
            Rect {
                x: 0.0,
                y: 0.0,
                width: 10.0,
                height: 5.0,
            },
            ResolvedGeometry::Rectangle,
        );
        text_shape.content = ResolvedContent::Text(text);
        let input = RenderInput {
            slides: vec![ResolvedSlide {
                size: (100.0, 50.0),
                background: None,
                shapes: vec![text_shape],
                diagnostics: Vec::new(),
            }],
            media: HashMap::new(),
            fonts: Vec::new(),
            metadata: None,
        };
        let mut fonts = FontManager::new_deterministic().expect("deterministic fonts");

        let page = layout_slide_with_fonts(&input, 0, &mut fonts).expect("layout overflow");
        let PositionedElement::Group(shape_group) = &page.elements[0] else {
            panic!("shape should lower to a group");
        };
        let baseline = glyph_runs(&shape_group.children)[0].origin.y;

        assert!(shape_group.clip.is_none());
        assert!(
            baseline < 0.0,
            "bottom anchor should preserve negative offset"
        );
    }

    #[test]
    fn justified_and_distributed_anchors_allocate_positive_spare_height() {
        let mut fonts = FontManager::new_deterministic().expect("deterministic fonts");
        let paragraph = || ResolvedParagraph {
            line_spacing: Some(ResolvedTextSpacing::Points(10.0)),
            runs: vec![ResolvedTextRun::Text {
                text: "line".to_owned(),
                style: ResolvedRunStyle {
                    font_size: Some(8.0),
                    latin_typeface: Some("Carlito".to_owned()),
                    ..ResolvedRunStyle::default()
                },
            }],
            ..ResolvedParagraph::default()
        };
        let content = Rect {
            x: 0.0,
            y: 0.0,
            width: 100.0,
            height: 60.0,
        };

        let baselines = |anchor, fonts: &mut FontManager| {
            let body = ResolvedTextBody {
                anchor,
                paragraphs: vec![paragraph(), paragraph(), paragraph()],
                ..text_body(TextInsets::default())
            };
            let stacked = stack_text(fonts, content, &body).expect("stack distributed text");
            distinct_baselines(&glyph_runs(&stacked.elements))
        };

        let justified = baselines(TextAnchor::Justified, &mut fonts);
        assert_close(justified[1] - justified[0], 25.0);
        assert_close(justified[2] - justified[1], 25.0);
        let distributed = baselines(TextAnchor::Distributed, &mut fonts);
        assert_close(distributed[0] - justified[0], 5.0);
        assert_close(distributed[1] - distributed[0], 20.0);
        assert_close(distributed[2] - distributed[1], 20.0);
    }

    #[test]
    fn bottom_center_text_in_an_inset_box_lands_at_the_computed_baseline() {
        let insets = TextInsets {
            left: 10.0,
            top: 8.0,
            right: 14.0,
            bottom: 12.0,
        };
        let style = ResolvedRunStyle {
            font_size: Some(12.0),
            latin_typeface: Some("Carlito".to_owned()),
            ..ResolvedRunStyle::default()
        };
        let text = ResolvedTextBody {
            anchor: TextAnchor::Bottom,
            paragraphs: vec![ResolvedParagraph {
                alignment: ParagraphAlignment::Center,
                line_spacing: Some(ResolvedTextSpacing::Points(20.0)),
                runs: vec![ResolvedTextRun::Text {
                    text: "baseline".to_owned(),
                    style: style.clone(),
                }],
                ..ResolvedParagraph::default()
            }],
            ..text_body(insets)
        };
        let mut text_shape = shape(
            Rect {
                x: 0.0,
                y: 0.0,
                width: 120.0,
                height: 80.0,
            },
            ResolvedGeometry::Rectangle,
        );
        text_shape.content = ResolvedContent::Text(text);
        let input = RenderInput {
            slides: vec![ResolvedSlide {
                size: (120.0, 80.0),
                background: None,
                shapes: vec![text_shape],
                diagnostics: Vec::new(),
            }],
            media: HashMap::new(),
            fonts: Vec::new(),
            metadata: None,
        };
        let mut expected_fonts = FontManager::new_deterministic().expect("deterministic fonts");
        let expected = shape_run(&mut expected_fonts, "baseline", &style).expect("shape baseline");
        let content_width = 120.0 - insets.left - insets.right;
        let content_height = 80.0 - insets.top - insets.bottom;
        let expected_x = insets.left + (content_width - expected.width) / 2.0;
        let half_leading = (20.0 - expected.ascent - expected.descent) / 2.0;
        let expected_y = insets.top + (content_height - 20.0) + expected.ascent + half_leading;
        let mut fonts = FontManager::new_deterministic().expect("deterministic fonts");

        let page = layout_slide_with_fonts(&input, 0, &mut fonts).expect("layout anchored text");
        let PositionedElement::Group(shape_group) = &page.elements[0] else {
            panic!("shape should lower to a group");
        };
        let run = glyph_runs(&shape_group.children)[0];

        assert!(matches!(
            shape_group.children.first(),
            Some(PositionedElement::Path(_))
        ));
        assert_close(run.origin.x, expected_x);
        assert_close(run.origin.y, expected_y);
    }

    #[test]
    fn shape_text_stays_above_the_path_and_overflows_without_a_clip() {
        let text = ResolvedTextBody {
            wrap: false,
            paragraphs: vec![ResolvedParagraph {
                runs: vec![ResolvedTextRun::Text {
                    text: "overflowing text".to_owned(),
                    style: ResolvedRunStyle {
                        font_size: Some(12.0),
                        latin_typeface: Some("Carlito".to_owned()),
                        ..ResolvedRunStyle::default()
                    },
                }],
                ..ResolvedParagraph::default()
            }],
            ..text_body(TextInsets::default())
        };
        let mut text_shape = shape(
            Rect {
                x: 0.0,
                y: 0.0,
                width: 10.0,
                height: 20.0,
            },
            ResolvedGeometry::Rectangle,
        );
        text_shape.content = ResolvedContent::Text(text);
        let input = RenderInput {
            slides: vec![ResolvedSlide {
                size: (100.0, 50.0),
                background: None,
                shapes: vec![text_shape],
                diagnostics: Vec::new(),
            }],
            media: HashMap::new(),
            fonts: Vec::new(),
            metadata: None,
        };
        let mut fonts = FontManager::new_deterministic().expect("deterministic fonts");

        let page = layout_slide_with_fonts(&input, 0, &mut fonts).expect("layout shape text");
        let PositionedElement::Group(shape_group) = &page.elements[0] else {
            panic!("shape should lower to a group");
        };

        assert!(shape_group.clip.is_none());
        assert!(matches!(
            shape_group.children.first(),
            Some(PositionedElement::Path(_))
        ));
        let text_runs = glyph_runs(&shape_group.children);
        assert!(!text_runs.is_empty());
        assert!(
            text_runs
                .iter()
                .map(|run| run.origin.x + run.advances.iter().sum::<f64>())
                .fold(0.0_f64, f64::max)
                > 10.0
        );
    }

    #[test]
    fn wingdings_f0b7_bullet_renders_as_a_visible_unicode_glyph() {
        let body = ResolvedTextBody {
            paragraphs: vec![bullet_paragraph(
                0,
                Some(ResolvedBullet::Character {
                    character: "\u{f0b7}".to_owned(),
                    font: Some("Wingdings".to_owned()),
                    color: None,
                    size: None,
                }),
                "visible",
            )],
            ..text_body(TextInsets::default())
        };
        let mut fonts = FontManager::new_deterministic().expect("deterministic fonts");
        let stacked =
            stack_text(&mut fonts, test_content_box(120.0), &body).expect("stack character bullet");
        let runs = glyph_runs(&stacked.elements);

        assert_eq!(runs[0].text, "\u{2022}");
        assert!(runs[0].glyph_ids.iter().all(|glyph| *glyph != 0));
    }

    #[test]
    fn eight_common_auto_number_schemes_format_exact_markers() {
        let schemes = [
            "arabicPlain",
            "arabicPeriod",
            "arabicParenR",
            "arabicParenBoth",
            "alphaLcPeriod",
            "alphaUcPeriod",
            "romanLcPeriod",
            "romanUcPeriod",
        ];
        let body = ResolvedTextBody {
            paragraphs: schemes
                .into_iter()
                .map(|scheme| {
                    bullet_paragraph(
                        0,
                        Some(ResolvedBullet::AutoNumber {
                            scheme: scheme.to_owned(),
                            start_at: 1,
                            font: None,
                            color: None,
                            size: None,
                        }),
                        "item",
                    )
                })
                .collect(),
            ..text_body(TextInsets::default())
        };
        let mut fonts = FontManager::new_deterministic().expect("deterministic fonts");
        let stacked = stack_text(&mut fonts, test_content_box(120.0), &body)
            .expect("stack automatic bullets");
        let markers: Vec<&str> = glyph_runs(&stacked.elements)
            .into_iter()
            .map(|run| run.text.as_str())
            .filter(|text| *text != "item")
            .collect();

        assert_eq!(markers, ["1", "1.", "1)", "(1)", "a.", "A.", "i.", "I."]);
    }

    #[test]
    fn automatic_bullets_increment_and_reset_by_level() {
        let auto = |level, scheme: &str, start_at| {
            bullet_paragraph(
                level,
                Some(ResolvedBullet::AutoNumber {
                    scheme: scheme.to_owned(),
                    start_at,
                    font: None,
                    color: None,
                    size: None,
                }),
                "item",
            )
        };
        let body = ResolvedTextBody {
            paragraphs: vec![
                auto(0, "arabicPeriod", 3),
                auto(1, "alphaLcPeriod", 1),
                auto(1, "alphaLcPeriod", 1),
                auto(0, "arabicPeriod", 3),
                auto(1, "alphaLcPeriod", 1),
                auto(0, "romanUcPeriod", 4),
                bullet_paragraph(0, None, "plain"),
                auto(0, "romanUcPeriod", 4),
                bullet_paragraph(
                    0,
                    Some(ResolvedBullet::Character {
                        character: "*".to_owned(),
                        font: None,
                        color: None,
                        size: None,
                    }),
                    "item",
                ),
                auto(0, "romanUcPeriod", 4),
            ],
            ..text_body(TextInsets::default())
        };
        let mut fonts = FontManager::new_deterministic().expect("deterministic fonts");
        let stacked =
            stack_text(&mut fonts, test_content_box(120.0), &body).expect("stack numbered outline");
        let markers: Vec<&str> = glyph_runs(&stacked.elements)
            .into_iter()
            .map(|run| run.text.as_str())
            .filter(|text| *text != "item" && *text != "plain")
            .collect();

        assert_eq!(
            markers,
            ["3.", "a.", "b.", "4.", "a.", "IV.", "IV.", "*", "IV."]
        );
    }

    #[test]
    fn bullet_style_overrides_and_fallbacks_are_independent() {
        let fallback_style = ResolvedRunStyle {
            font_size: Some(20.0),
            bold: true,
            fill: Some(Paint::Solid(Color {
                r: 0.1,
                g: 0.2,
                b: 0.8,
                a: 1.0,
            })),
            latin_typeface: Some("Carlito".to_owned()),
            ..ResolvedRunStyle::default()
        };
        let bullet = |font, color, size| ResolvedBullet::Character {
            character: "*".to_owned(),
            font,
            color,
            size,
        };
        let paragraph = |bullet| ResolvedParagraph {
            bullet: Some(bullet),
            runs: vec![ResolvedTextRun::Text {
                text: "item".to_owned(),
                style: fallback_style.clone(),
            }],
            ..ResolvedParagraph::default()
        };
        let override_color = Color {
            r: 0.8,
            g: 0.2,
            b: 0.1,
            a: 1.0,
        };
        let body = ResolvedTextBody {
            paragraphs: vec![
                paragraph(bullet(Some("Caladea".to_owned()), None, None)),
                paragraph(bullet(
                    None,
                    Some(override_color),
                    Some(ResolvedBulletSize::Points(11.0)),
                )),
                paragraph(bullet(None, None, Some(ResolvedBulletSize::Percent(1.5)))),
            ],
            ..text_body(TextInsets::default())
        };
        let mut fonts = FontManager::new_deterministic().expect("deterministic fonts");
        let stacked =
            stack_text(&mut fonts, test_content_box(120.0), &body).expect("stack styled bullets");
        let markers: Vec<_> = glyph_runs(&stacked.elements)
            .into_iter()
            .filter(|run| run.text == "*")
            .collect();
        let families: HashMap<_, _> = fonts
            .all_font_data()
            .into_iter()
            .map(|font| (font.id, font.family))
            .collect();

        assert_eq!(families[&markers[0].font_id], "Caladea");
        assert_eq!(markers[0].font_size, 20.0);
        assert_eq!(markers[0].color, text_color(fallback_style.fill.as_ref()));
        assert!(markers[0].bold);
        assert_eq!(families[&markers[1].font_id], "Carlito");
        assert_eq!(markers[1].font_size, 11.0);
        assert_eq!(markers[1].color, override_color);
        assert_eq!(markers[2].font_size, 30.0);
    }

    #[test]
    fn bullet_marker_uses_left_and_hanging_indent() {
        let mut paragraph = bullet_paragraph(
            0,
            Some(ResolvedBullet::Character {
                character: "*".to_owned(),
                font: None,
                color: None,
                size: None,
            }),
            "one two three four five",
        );
        paragraph.left_margin = 20.0;
        paragraph.indent = -10.0;
        let body = ResolvedTextBody {
            paragraphs: vec![paragraph],
            ..text_body(TextInsets::default())
        };
        let mut fonts = FontManager::new_deterministic().expect("deterministic fonts");
        let stacked =
            stack_text(&mut fonts, test_content_box(55.0), &body).expect("stack indented bullet");
        let runs = glyph_runs(&stacked.elements);
        let marker_x = runs
            .iter()
            .find(|run| run.text == "*")
            .expect("marker")
            .origin
            .x;
        let text_runs: Vec<_> = runs.into_iter().filter(|run| run.text != "*").collect();

        assert_close(marker_x, 10.0);
        assert_close(text_runs[0].origin.x, 20.0);
        assert!(
            text_runs
                .iter()
                .skip(1)
                .all(|run| (run.origin.x - 20.0).abs() < 0.001)
        );
    }

    #[test]
    fn distributed_bullet_keeps_the_fixed_hanging_slot() {
        let mut paragraph = bullet_paragraph(
            0,
            Some(ResolvedBullet::Character {
                character: "*".to_owned(),
                font: None,
                color: None,
                size: None,
            }),
            "ab",
        );
        paragraph.left_margin = 20.0;
        paragraph.indent = -10.0;
        paragraph.alignment = ParagraphAlignment::Distributed;
        let body = ResolvedTextBody {
            paragraphs: vec![paragraph],
            ..text_body(TextInsets::default())
        };
        let mut fonts = FontManager::new_deterministic().expect("deterministic fonts");

        let stacked =
            stack_text(&mut fonts, test_content_box(55.0), &body).expect("stack bullet text");
        let runs = glyph_runs(&stacked.elements);
        let marker = runs.iter().find(|run| run.text == "*").expect("marker");
        let text = runs.iter().find(|run| run.text == "ab").expect("body text");

        assert_close(marker.origin.x, 10.0);
        assert_close(text.origin.x, 20.0);
    }

    #[test]
    fn wide_auto_number_marker_keeps_text_on_the_paragraph_margin() {
        let mut paragraph = bullet_paragraph(
            0,
            Some(ResolvedBullet::AutoNumber {
                scheme: "arabicPeriod".to_owned(),
                start_at: 12_345,
                font: None,
                color: None,
                size: None,
            }),
            "one two three four five",
        );
        paragraph.left_margin = 20.0;
        paragraph.indent = -10.0;
        let body = ResolvedTextBody {
            paragraphs: vec![paragraph],
            ..text_body(TextInsets::default())
        };
        let mut fonts = FontManager::new_deterministic().expect("deterministic fonts");
        let stacked = stack_text(&mut fonts, test_content_box(55.0), &body)
            .expect("stack wide automatic bullet");
        let runs = glyph_runs(&stacked.elements);
        let marker = runs
            .iter()
            .find(|run| run.text == "12345.")
            .expect("automatic marker");
        let text_runs: Vec<_> = runs.iter().filter(|run| run.text != "12345.").collect();

        assert!(marker.advances.iter().sum::<f64>() > 10.0);
        assert_close(marker.origin.x, 10.0);
        assert_close(text_runs[0].origin.x, 20.0);
        assert!(
            text_runs
                .iter()
                .skip(1)
                .all(|run| (run.origin.x - 20.0).abs() < 0.001)
        );
    }

    #[test]
    fn unsupported_auto_number_scheme_keeps_a_visible_marker() {
        let body = ResolvedTextBody {
            paragraphs: vec![bullet_paragraph(
                0,
                Some(ResolvedBullet::AutoNumber {
                    scheme: "thaiNumPeriod".to_owned(),
                    start_at: 7,
                    font: None,
                    color: None,
                    size: None,
                }),
                "item",
            )],
            ..text_body(TextInsets::default())
        };
        let mut fonts = FontManager::new_deterministic().expect("deterministic fonts");
        let stacked =
            stack_text(&mut fonts, test_content_box(120.0), &body).expect("stack fallback bullet");
        let runs = glyph_runs(&stacked.elements);

        assert_eq!(runs[0].text, "7.");
        assert!(runs[0].glyph_ids.iter().all(|glyph| *glyph != 0));
    }

    #[test]
    fn stored_font_scale_renders_at_exactly_sixty_two_point_five_percent() {
        let mut fonts = FontManager::new_deterministic().expect("deterministic fonts");
        let mut body = autofit_body(
            ResolvedAutofit::Normal {
                font_scale: Some(0.625),
                line_spacing_reduction: None,
            },
            20.0,
            "stored scale",
        );
        body.paragraphs[0].bullet = Some(ResolvedBullet::Character {
            character: "*".to_owned(),
            font: None,
            color: None,
            size: Some(ResolvedBulletSize::Points(10.0)),
        });

        let stacked = stack_text(&mut fonts, test_content_box(200.0), &body)
            .expect("stack stored autofit text");
        let runs = glyph_runs(&stacked.elements);

        assert_close(runs[0].font_size, 6.25);
        assert_close(runs[1].font_size, 12.5);
    }

    #[test]
    fn repeated_resolved_runs_reuse_one_shaped_cache_entry() {
        let mut fonts = FontManager::new_deterministic().expect("deterministic fonts");
        let style = ResolvedRunStyle {
            font_size: Some(20.0),
            latin_typeface: Some("Carlito".to_owned()),
            ..ResolvedRunStyle::default()
        };
        let mut cache = ShapingCache::default();

        let first = cache
            .shape(&mut fonts, "cached", &style)
            .expect("shape first cached run");
        let second = cache
            .shape(&mut fonts, "cached", &style)
            .expect("reuse cached run");

        assert_eq!(cache.entries.len(), 1);
        assert_eq!(first.glyph_ids, second.glyph_ids);
        assert_eq!(first.advances, second.advances);
    }

    #[test]
    fn stored_line_spacing_reduction_applies_to_percentage_not_point_spacing() {
        let mut fonts = FontManager::new_deterministic().expect("deterministic fonts");
        let mut body = autofit_body(
            ResolvedAutofit::Normal {
                font_scale: Some(1.0),
                line_spacing_reduction: Some(0.75),
            },
            20.0,
            "leading",
        );
        let segment = shape_run(&mut fonts, "leading", first_run_style(&body.paragraphs[0]))
            .expect("shape natural line");
        let natural_height = segment.ascent + segment.descent + segment.line_gap;

        let single =
            stack_text(&mut fonts, test_content_box(200.0), &body).expect("stack natural single");
        assert_close(single.height, natural_height);

        body.paragraphs[0].line_spacing = Some(ResolvedTextSpacing::Percent(1.5));
        let percentage = stack_text(&mut fonts, test_content_box(200.0), &body)
            .expect("stack reduced percentage");
        assert_close(percentage.height, 20.0 * 1.5 * 0.25);

        body.paragraphs[0].line_spacing = Some(ResolvedTextSpacing::Points(30.0));
        let points =
            stack_text(&mut fonts, test_content_box(200.0), &body).expect("stack exact points");
        assert_close(points.height, 30.0);
    }

    #[test]
    fn first_and_last_paragraph_spacing_obeys_body_property() {
        let mut fonts = FontManager::new_deterministic().expect("deterministic fonts");
        let run = || ResolvedTextRun::Text {
            text: "line".to_owned(),
            style: ResolvedRunStyle {
                font_size: Some(20.0),
                latin_typeface: Some("Carlito".to_owned()),
                ..ResolvedRunStyle::default()
            },
        };
        let mut body = ResolvedTextBody {
            paragraphs: vec![
                ResolvedParagraph {
                    line_spacing: Some(ResolvedTextSpacing::Points(20.0)),
                    space_before: Some(ResolvedTextSpacing::Points(7.0)),
                    space_after: Some(ResolvedTextSpacing::Points(3.0)),
                    runs: vec![run()],
                    ..ResolvedParagraph::default()
                },
                ResolvedParagraph {
                    line_spacing: Some(ResolvedTextSpacing::Points(20.0)),
                    space_before: Some(ResolvedTextSpacing::Points(5.0)),
                    space_after: Some(ResolvedTextSpacing::Points(11.0)),
                    runs: vec![run()],
                    ..ResolvedParagraph::default()
                },
            ],
            ..text_body(TextInsets::default())
        };

        let default_edges = stack_text(&mut fonts, test_content_box(200.0), &body)
            .expect("stack default edge spacing");
        assert_close(default_edges.height, 48.0);

        body.space_first_last_paragraph = true;
        let explicit_edges = stack_text(&mut fonts, test_content_box(200.0), &body)
            .expect("stack explicit edge spacing");
        assert_close(explicit_edges.height, 66.0);
    }

    #[test]
    fn shape_autofit_trusts_the_stored_extent() {
        let mut fonts = FontManager::new_deterministic().expect("deterministic fonts");
        let body = autofit_body(ResolvedAutofit::Shape, 20.0, "stored shape extent");
        let content = Rect {
            width: 1.0,
            height: 1.0,
            ..test_content_box(0.0)
        };

        let stacked = stack_text(&mut fonts, content, &body).expect("stack shape autofit text");

        assert_close(glyph_runs(&stacked.elements)[0].font_size, 20.0);
        assert!(stacked.width > content.width || stacked.height > content.height);
    }

    #[test]
    fn no_autofit_overflows_without_a_clip() {
        let text = autofit_body(ResolvedAutofit::None, 20.0, "visible overflow");
        let mut text_shape = shape(
            Rect {
                x: 0.0,
                y: 0.0,
                width: 1.0,
                height: 1.0,
            },
            ResolvedGeometry::Rectangle,
        );
        text_shape.content = ResolvedContent::Text(text);
        let input = RenderInput {
            slides: vec![ResolvedSlide {
                size: (100.0, 50.0),
                background: None,
                shapes: vec![text_shape],
                diagnostics: Vec::new(),
            }],
            media: HashMap::new(),
            fonts: Vec::new(),
            metadata: None,
        };
        let mut fonts = FontManager::new_deterministic().expect("deterministic fonts");

        let page = layout_slide_with_fonts(&input, 0, &mut fonts).expect("layout visible overflow");
        let PositionedElement::Group(shape_group) = &page.elements[0] else {
            panic!("shape should lower to a group");
        };
        let run = glyph_runs(&shape_group.children)[0];

        assert!(shape_group.clip.is_none());
        assert_close(run.font_size, 20.0);
        assert!(run.advances.iter().sum::<f64>() > 1.0);
    }

    #[test]
    fn bare_normal_autofit_uses_quantised_two_point_five_percent_steps() {
        let mut fonts = FontManager::new_deterministic().expect("deterministic fonts");
        let body = autofit_body(
            ResolvedAutofit::Normal {
                font_scale: None,
                line_spacing_reduction: None,
            },
            10.0,
            "quantised ladder",
        );
        let segment = shape_run(
            &mut fonts,
            "quantised ladder",
            first_run_style(&body.paragraphs[0]),
        )
        .expect("shape ladder text");
        let content = Rect {
            width: segment.width * 0.94,
            ..test_content_box(0.0)
        };

        let stacked = stack_text(&mut fonts, content, &body).expect("stack ladder text");

        assert_close(glyph_runs(&stacked.elements)[0].font_size, 9.25);
    }

    #[test]
    fn bare_normal_autofit_measures_text_inside_paragraph_margins() {
        let mut fonts = FontManager::new_deterministic().expect("deterministic fonts");
        let mut body = autofit_body(
            ResolvedAutofit::Normal {
                font_scale: None,
                line_spacing_reduction: None,
            },
            10.0,
            "margin fit",
        );
        let segment = shape_run(
            &mut fonts,
            "margin fit",
            first_run_style(&body.paragraphs[0]),
        )
        .expect("shape margin text");
        body.paragraphs[0].left_margin = segment.width * 0.15;
        let content = Rect {
            width: segment.width * 1.1,
            ..test_content_box(0.0)
        };

        let stacked = stack_text(&mut fonts, content, &body).expect("stack margin text");

        assert_close(glyph_runs(&stacked.elements)[0].font_size, 9.5);
    }

    #[test]
    fn bare_normal_autofit_keeps_the_twenty_five_percent_floor_visible() {
        let mut fonts = FontManager::new_deterministic().expect("deterministic fonts");
        let body = autofit_body(
            ResolvedAutofit::Normal {
                font_scale: None,
                line_spacing_reduction: None,
            },
            20.0,
            "visible at the floor",
        );
        let content = Rect {
            width: 1.0,
            height: 1.0,
            ..test_content_box(0.0)
        };

        let stacked = stack_text(&mut fonts, content, &body).expect("stack floor text");
        let run = glyph_runs(&stacked.elements)[0];

        assert_close(run.font_size, 5.0);
        assert!(stacked.width > content.width || stacked.height > content.height);
    }

    #[test]
    fn rich_powerpoint_text_path_keeps_visual_runs_and_logical_text_together() {
        let source = "abc العربية कि ภาษาไทย 〈中〉、 אבג 123";
        let body = ResolvedTextBody {
            paragraphs: vec![ResolvedParagraph {
                runs: vec![ResolvedTextRun::Text {
                    text: source.to_owned(),
                    style: ResolvedRunStyle {
                        font_size: Some(18.0),
                        latin_typeface: Some("Carlito".to_owned()),
                        ..ResolvedRunStyle::default()
                    },
                }],
                ..ResolvedParagraph::default()
            }],
            ..text_body(TextInsets::default())
        };
        let mut fonts = FontManager::new_deterministic().expect("deterministic fonts");

        let stacked = stack_text(&mut fonts, test_content_box(2_000.0), &body)
            .expect("shape multilingual PowerPoint text");
        let runs = stacked
            .elements
            .iter()
            .filter_map(|element| match element {
                PositionedElement::MultilingualText(run) => Some(run),
                _ => None,
            })
            .collect::<Vec<_>>();
        assert!(runs.len() > 4);
        assert!(runs.iter().all(|run| {
            run.glyph_ids.len() == run.x_advances.len()
                && run.glyph_ids.len() == run.x_offsets.len()
                && run.glyph_ids.len() == run.y_offsets.len()
        }));
        assert!(runs.iter().any(|run| run.bidi_level % 2 == 1));
        let logical_indices = runs.iter().map(|run| run.logical_index).collect::<Vec<_>>();
        let logical = runs
            .iter()
            .map(|run| run.logical_text.as_str())
            .collect::<String>();

        assert_eq!(logical, source);
        assert_eq!(
            logical_indices,
            (0..logical_indices.len()).collect::<Vec<_>>()
        );
        assert!(
            runs.windows(2)
                .any(|pair| pair[0].origin.x > pair[1].origin.x),
            "logical extraction order must coexist with visual positions"
        );

        let layout = oxml_layout::LayoutResult::new(
            vec![oxml_layout::PageFrame::new(1, 2_000.0, 200.0, stacked.elements).into()],
            fonts.all_font_data(),
            None,
            Vec::new(),
        );
        let pdf = oxml_pdf::render_to_pdf(&layout);
        let png = oxml_pdf::render_page_to_png(&layout, 0, 72.0)
            .expect("raster consumes positioned multilingual runs");
        assert!(pdf.starts_with(b"%PDF"));
        assert!(png.starts_with(b"\x89PNG"));
    }

    #[test]
    fn explicit_rtl_styled_runs_around_a_forced_break_keep_logical_order() {
        let body = ResolvedTextBody {
            paragraphs: vec![ResolvedParagraph {
                runs: vec![
                    ResolvedTextRun::Text {
                        text: "אב".to_owned(),
                        style: ResolvedRunStyle {
                            font_size: Some(18.0),
                            ..ResolvedRunStyle::default()
                        },
                    },
                    ResolvedTextRun::Text {
                        text: "גד".to_owned(),
                        style: ResolvedRunStyle {
                            font_size: Some(18.0),
                            italic: true,
                            ..ResolvedRunStyle::default()
                        },
                    },
                    ResolvedTextRun::Break,
                    ResolvedTextRun::Text {
                        text: "הו".to_owned(),
                        style: ResolvedRunStyle {
                            font_size: Some(18.0),
                            ..ResolvedRunStyle::default()
                        },
                    },
                    ResolvedTextRun::Text {
                        text: "זח".to_owned(),
                        style: ResolvedRunStyle {
                            font_size: Some(18.0),
                            bold: true,
                            ..ResolvedRunStyle::default()
                        },
                    },
                ],
                ..ResolvedParagraph::default()
            }],
            ..text_body(TextInsets::default())
        };
        let mut fonts = FontManager::new_deterministic().expect("deterministic fonts");

        let stacked = stack_text_for_page_with_directions(
            &mut fonts,
            test_content_box(2_000.0),
            &body,
            1,
            &[oxml_layout::TextDirection::RightToLeft],
        )
        .unwrap();
        let runs = stacked
            .elements
            .iter()
            .filter_map(|element| match element {
                PositionedElement::MultilingualText(run) => Some(run),
                _ => None,
            })
            .collect::<Vec<_>>();
        assert_eq!(
            runs.iter()
                .map(|run| { (run.logical_text.as_str(), run.logical_index, run.bidi_level,) })
                .collect::<Vec<_>>(),
            [("אב", 0, 1), ("גד", 1, 1), ("הו", 2, 1), ("זח", 3, 1)]
        );
        assert!(runs[0].origin.x > runs[1].origin.x);
        assert!(runs[2].origin.x > runs[3].origin.x);
        assert!(runs[0].origin.y < runs[2].origin.y);
    }

    fn bullet_paragraph(
        level: u8,
        bullet: Option<ResolvedBullet>,
        text: &str,
    ) -> ResolvedParagraph {
        ResolvedParagraph {
            level,
            left_margin: 10.0,
            indent: -5.0,
            bullet,
            runs: vec![ResolvedTextRun::Text {
                text: text.to_owned(),
                style: ResolvedRunStyle {
                    font_size: Some(12.0),
                    latin_typeface: Some("Carlito".to_owned()),
                    ..ResolvedRunStyle::default()
                },
            }],
            ..ResolvedParagraph::default()
        }
    }

    fn autofit_body(autofit: ResolvedAutofit, font_size: f64, text: &str) -> ResolvedTextBody {
        ResolvedTextBody {
            wrap: false,
            autofit,
            paragraphs: vec![ResolvedParagraph {
                runs: vec![ResolvedTextRun::Text {
                    text: text.to_owned(),
                    style: ResolvedRunStyle {
                        font_size: Some(font_size),
                        latin_typeface: Some("Carlito".to_owned()),
                        ..ResolvedRunStyle::default()
                    },
                }],
                ..ResolvedParagraph::default()
            }],
            ..text_body(TextInsets::default())
        }
    }

    fn test_content_box(width: f64) -> Rect {
        Rect {
            x: 0.0,
            y: 0.0,
            width,
            height: 400.0,
        }
    }

    fn glyph_runs(elements: &[oxml_layout::PositionedElement]) -> Vec<&oxml_layout::GlyphRun> {
        elements
            .iter()
            .filter_map(|element| match element {
                oxml_layout::PositionedElement::Text(run) => Some(run),
                _ => None,
            })
            .collect()
    }

    fn distinct_baselines(runs: &[&oxml_layout::GlyphRun]) -> Vec<f64> {
        let mut baselines = Vec::new();
        for run in runs {
            if baselines
                .last()
                .is_none_or(|baseline: &f64| (run.origin.y - baseline).abs() > 0.001)
            {
                baselines.push(run.origin.y);
            }
        }
        baselines
    }

    fn assert_close(actual: f64, expected: f64) {
        assert!(
            (actual - expected).abs() < 0.001,
            "expected {expected}, got {actual}"
        );
    }
}
