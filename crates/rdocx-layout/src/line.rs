//! Line breaking: converts inline items into laid-out lines.
//!
//! Uses a greedy algorithm with unicode-linebreak for break opportunities.

use rdocx_oxml::borders::CT_TabStop;
use rdocx_oxml::shared::{ST_Jc, ST_TabJc, ST_Underline};
use rdocx_oxml::units::Twips;

use crate::error::Result;
use crate::font::FontManager;
use crate::output::{Color, FieldKind, FontId};

/// An inline item to be placed on a line.
#[derive(Debug, Clone)]
pub enum InlineItem {
    /// A shaped text segment.
    Text(TextSegment),
    /// A tab character.
    Tab,
    /// A forced line break.
    LineBreak,
    /// A forced page break.
    PageBreak,
    /// A forced column break.
    ColumnBreak,
    /// An inline image.
    Image {
        width: f64,
        height: f64,
        embed_id: String,
    },
    /// A numbering marker (rendered before the first line).
    Marker(TextSegment),
}

/// A shaped text segment with associated formatting.
#[derive(Debug, Clone)]
pub struct TextSegment {
    pub text: String,
    pub font_id: FontId,
    pub font_size: f64,
    pub glyph_ids: Vec<u16>,
    pub advances: Vec<f64>,
    pub width: f64,
    pub ascent: f64,
    pub descent: f64,
    pub color: Color,
    pub bold: bool,
    pub italic: bool,
    /// Underline style (None = no underline).
    pub underline: Option<ST_Underline>,
    /// Single strikethrough.
    pub strike: bool,
    /// Double strikethrough.
    pub dstrike: bool,
    /// Highlight/background color for the run.
    pub highlight: Option<Color>,
    /// Baseline offset in points (positive = raise, negative = lower).
    pub baseline_offset: f64,
    /// Hyperlink URL if this segment is inside a hyperlink.
    pub hyperlink_url: Option<String>,
    /// If this segment is a field placeholder, the kind of field.
    pub field_kind: Option<FieldKind>,
    /// If this segment is a footnote/endnote reference marker, its ID.
    pub footnote_id: Option<i32>,
}

/// A single item positioned on a line.
#[derive(Debug, Clone)]
pub enum LineItem {
    Text(TextSegment),
    Tab {
        width: f64,
        /// Pre-shaped leader text to fill the tab gap (e.g., dots, hyphens).
        leader: Option<TextSegment>,
    },
    Image {
        width: f64,
        height: f64,
        embed_id: String,
    },
    Marker(TextSegment),
}

impl LineItem {
    pub fn width(&self) -> f64 {
        match self {
            LineItem::Text(seg) => seg.width,
            LineItem::Tab { width, .. } => *width,
            LineItem::Image { width, .. } => *width,
            LineItem::Marker(seg) => seg.width,
        }
    }
}

/// A laid-out line within a paragraph.
#[derive(Debug, Clone)]
pub struct LayoutLine {
    pub items: Vec<LineItem>,
    /// Total content width of the line.
    pub width: f64,
    /// Maximum ascent on this line (above baseline).
    pub ascent: f64,
    /// Maximum descent on this line (below baseline).
    pub descent: f64,
    /// Total line height.
    pub height: f64,
    /// Left indent for this line.
    pub indent_left: f64,
    /// Available width this line was laid out against.
    pub available_width: f64,
    /// Whether this is the last line of the paragraph.
    pub is_last: bool,
}

/// Parameters for line breaking.
pub struct LineBreakParams {
    /// Total available width (page width minus margins).
    pub available_width: f64,
    /// Left indentation in points.
    pub ind_left: f64,
    /// Right indentation in points.
    pub ind_right: f64,
    /// First line indent in points (positive = indent, 0 if hanging).
    pub ind_first_line: f64,
    /// Hanging indent in points (positive = text lines indented relative to first).
    pub ind_hanging: f64,
    /// Tab stops.
    pub tab_stops: Vec<CT_TabStop>,
    /// Line spacing value.
    pub line_spacing: Option<Twips>,
    /// Line spacing rule.
    pub line_rule: Option<String>,
    /// Paragraph justification.
    pub jc: Option<ST_Jc>,
}

impl Default for LineBreakParams {
    fn default() -> Self {
        LineBreakParams {
            available_width: 468.0, // US Letter with 1" margins
            ind_left: 0.0,
            ind_right: 0.0,
            ind_first_line: 0.0,
            ind_hanging: 0.0,
            tab_stops: Vec::new(),
            line_spacing: None,
            line_rule: None,
            jc: None,
        }
    }
}

/// Break inline items into lines using a greedy algorithm.
pub fn break_into_lines(
    items: &[InlineItem],
    params: &LineBreakParams,
    fm: &FontManager,
) -> Result<Vec<LayoutLine>> {
    if items.is_empty() {
        // Empty paragraph still gets one empty line
        return Ok(vec![LayoutLine {
            items: Vec::new(),
            width: 0.0,
            ascent: 0.0,
            descent: 0.0,
            height: compute_line_height(0.0, 0.0, params),
            indent_left: params.ind_left + params.ind_first_line,
            available_width: params.available_width,
            is_last: true,
        }]);
    }

    let mut lines: Vec<LayoutLine> = Vec::new();
    let mut current_items: Vec<LineItem> = Vec::new();
    let mut current_width: f64 = 0.0;
    let mut current_ascent: f64 = 0.0;
    let mut current_descent: f64 = 0.0;
    let mut is_first_line = true;

    let first_line_width = compute_first_line_width(params);
    let subsequent_line_width = compute_subsequent_line_width(params);

    let mut line_avail = first_line_width;

    // Track the most recent font context for shaping tab leaders
    let mut font_ctx: Option<(FontId, f64)> = None;
    // Initialize from the first text segment if available
    for item in items {
        if let InlineItem::Text(seg) | InlineItem::Marker(seg) = item {
            font_ctx = Some((seg.font_id, seg.font_size));
            break;
        }
    }

    // Build breakable segments from inline items
    let segments = build_breakable_segments(items);

    for seg in &segments {
        match seg {
            BreakableSegment::Items(seg_items) => {
                let seg_width: f64 = seg_items.iter().map(inline_item_width).sum();

                if !current_items.is_empty() && current_width + seg_width > line_avail + 0.01 {
                    // Finish current line
                    let indent = if is_first_line {
                        first_line_indent(params)
                    } else {
                        subsequent_line_indent(params)
                    };
                    lines.push(LayoutLine {
                        items: std::mem::take(&mut current_items),
                        width: current_width,
                        ascent: current_ascent,
                        descent: current_descent,
                        height: compute_line_height(current_ascent, current_descent, params),
                        indent_left: indent,
                        available_width: line_avail,
                        is_last: false,
                    });
                    current_width = 0.0;
                    current_ascent = 0.0;
                    current_descent = 0.0;
                    is_first_line = false;
                    line_avail = subsequent_line_width;
                }

                // Add segment items to current line
                for item in seg_items {
                    let (w, a, d) = item_metrics(item);
                    current_width += w;
                    if a > current_ascent {
                        current_ascent = a;
                    }
                    if d > current_descent {
                        current_descent = d;
                    }
                    // Update font context from text segments
                    if let InlineItem::Text(seg) | InlineItem::Marker(seg) = item {
                        font_ctx = Some((seg.font_id, seg.font_size));
                    }
                    current_items.push(inline_to_line_item(
                        item,
                        current_width,
                        &params.tab_stops,
                        fm,
                        font_ctx,
                    ));
                }
            }
            BreakableSegment::ForcedBreak(break_type) => {
                let indent = if is_first_line {
                    first_line_indent(params)
                } else {
                    subsequent_line_indent(params)
                };
                lines.push(LayoutLine {
                    items: std::mem::take(&mut current_items),
                    width: current_width,
                    ascent: current_ascent,
                    descent: current_descent,
                    height: compute_line_height(current_ascent, current_descent, params),
                    indent_left: indent,
                    available_width: line_avail,
                    is_last: matches!(break_type, ForcedBreakType::Page | ForcedBreakType::Column),
                });
                current_width = 0.0;
                current_ascent = 0.0;
                current_descent = 0.0;
                is_first_line = false;
                line_avail = subsequent_line_width;
            }
        }
    }

    // Flush remaining items as the last line
    let indent = if is_first_line {
        first_line_indent(params)
    } else {
        subsequent_line_indent(params)
    };
    lines.push(LayoutLine {
        items: current_items,
        width: current_width,
        ascent: current_ascent,
        descent: current_descent,
        height: compute_line_height(current_ascent, current_descent, params),
        indent_left: indent,
        available_width: line_avail,
        is_last: true,
    });

    Ok(lines)
}

// ---- Internal helpers ----

#[derive(Debug)]
enum BreakableSegment {
    /// A group of items that should be kept together (word or cluster).
    Items(Vec<InlineItem>),
    /// A forced break.
    ForcedBreak(ForcedBreakType),
}

#[derive(Debug)]
enum ForcedBreakType {
    Line,
    Page,
    Column,
}

/// Build breakable segments by finding break opportunities in text.
///
/// Text items are split at unicode line-break opportunities (word boundaries,
/// hyphens, etc.). Non-text items (tabs, images, markers) are treated as
/// atomic units with break opportunities around them.
fn build_breakable_segments(items: &[InlineItem]) -> Vec<BreakableSegment> {
    let mut segments = Vec::new();
    let mut current_group: Vec<InlineItem> = Vec::new();

    for item in items {
        match item {
            InlineItem::LineBreak => {
                if !current_group.is_empty() {
                    segments.push(BreakableSegment::Items(std::mem::take(&mut current_group)));
                }
                segments.push(BreakableSegment::ForcedBreak(ForcedBreakType::Line));
            }
            InlineItem::PageBreak => {
                if !current_group.is_empty() {
                    segments.push(BreakableSegment::Items(std::mem::take(&mut current_group)));
                }
                segments.push(BreakableSegment::ForcedBreak(ForcedBreakType::Page));
            }
            InlineItem::ColumnBreak => {
                if !current_group.is_empty() {
                    segments.push(BreakableSegment::Items(std::mem::take(&mut current_group)));
                }
                segments.push(BreakableSegment::ForcedBreak(ForcedBreakType::Column));
            }
            InlineItem::Tab => {
                // Tab is a break opportunity
                if !current_group.is_empty() {
                    segments.push(BreakableSegment::Items(std::mem::take(&mut current_group)));
                }
                segments.push(BreakableSegment::Items(vec![item.clone()]));
            }
            InlineItem::Text(seg) => {
                // Use unicode-linebreak to find break opportunities within text
                let breaks = split_text_at_break_opportunities(seg);

                for tb in &breaks {
                    let chunk = &seg.text[tb.start..tb.end];
                    if chunk.is_empty() {
                        continue;
                    }

                    // If this chunk starts with whitespace, treat as a break opportunity
                    if !current_group.is_empty() && chunk.starts_with(|c: char| c.is_whitespace()) {
                        segments.push(BreakableSegment::Items(std::mem::take(&mut current_group)));
                    }

                    // Create a sub-segment for just this chunk (not the entire text)
                    let sub_item = split_text_subsegment(seg, tb.start, tb.end);
                    current_group.push(sub_item);

                    if tb.is_break {
                        segments.push(BreakableSegment::Items(std::mem::take(&mut current_group)));
                    }
                }

                // Flush any remaining
                if !current_group.is_empty() {
                    segments.push(BreakableSegment::Items(std::mem::take(&mut current_group)));
                }
            }
            InlineItem::Marker(_) | InlineItem::Image { .. } => {
                current_group.push(item.clone());
            }
        }
    }

    if !current_group.is_empty() {
        segments.push(BreakableSegment::Items(current_group));
    }

    segments
}

/// Create a sub-segment InlineItem from a byte range within a TextSegment.
///
/// Slices the text, glyph_ids, and advances to the specified byte range,
/// preserving all formatting properties from the parent segment.
fn split_text_subsegment(seg: &TextSegment, byte_start: usize, byte_end: usize) -> InlineItem {
    // If this is the full segment, just clone it
    if byte_start == 0 && byte_end == seg.text.len() {
        return InlineItem::Text(seg.clone());
    }

    let sub_text = seg.text[byte_start..byte_end].to_string();
    let total_chars = seg.text.chars().count();
    let char_start = seg.text[..byte_start].chars().count();
    let char_count = sub_text.chars().count();

    let (sub_glyphs, sub_advances, sub_width) = if seg.glyph_ids.len() == total_chars {
        // 1:1 char-to-glyph mapping (common for Latin text)
        let end = (char_start + char_count).min(seg.glyph_ids.len());
        let glyphs = seg.glyph_ids[char_start..end].to_vec();
        let advances = seg.advances[char_start..end].to_vec();
        let width: f64 = advances.iter().sum();
        (glyphs, advances, width)
    } else if seg.glyph_ids.is_empty() {
        // No glyphs (shouldn't happen, but handle gracefully)
        (vec![], vec![], 0.0)
    } else {
        // Non-1:1 mapping (ligatures, complex scripts) — proportional estimate
        let byte_frac = (byte_end - byte_start) as f64 / seg.text.len() as f64;
        let est_glyphs = (seg.glyph_ids.len() as f64 * byte_frac).round() as usize;
        let glyph_start = (seg.glyph_ids.len() as f64 * byte_start as f64 / seg.text.len() as f64)
            .round() as usize;
        let glyph_end = (glyph_start + est_glyphs).min(seg.glyph_ids.len());
        let glyphs = seg.glyph_ids[glyph_start..glyph_end].to_vec();
        let advances = seg.advances[glyph_start..glyph_end].to_vec();
        let width: f64 = advances.iter().sum();
        (glyphs, advances, width)
    };

    InlineItem::Text(TextSegment {
        text: sub_text,
        font_id: seg.font_id,
        font_size: seg.font_size,
        glyph_ids: sub_glyphs,
        advances: sub_advances,
        width: sub_width,
        ascent: seg.ascent,
        descent: seg.descent,
        color: seg.color,
        bold: seg.bold,
        italic: seg.italic,
        underline: seg.underline,
        strike: seg.strike,
        dstrike: seg.dstrike,
        highlight: seg.highlight,
        baseline_offset: seg.baseline_offset,
        hyperlink_url: seg.hyperlink_url.clone(),
        field_kind: seg.field_kind,
        footnote_id: seg.footnote_id,
    })
}

struct TextBreakInfo {
    /// Byte range within the original text.
    start: usize,
    end: usize,
    /// Whether a line break is allowed after this segment.
    is_break: bool,
}

fn split_text_at_break_opportunities(seg: &TextSegment) -> Vec<TextBreakInfo> {
    use unicode_linebreak::{BreakOpportunity, linebreaks};

    let text = &seg.text;
    if text.is_empty() {
        return vec![];
    }

    let mut breaks = Vec::new();
    let mut last_start = 0;

    for (byte_pos, opportunity) in linebreaks(text) {
        if byte_pos == 0 {
            continue;
        }

        let is_break = matches!(
            opportunity,
            BreakOpportunity::Allowed | BreakOpportunity::Mandatory
        );

        breaks.push(TextBreakInfo {
            start: last_start,
            end: byte_pos,
            is_break,
        });
        last_start = byte_pos;
    }

    // If unicode-linebreak didn't produce any breaks, treat as one chunk
    if breaks.is_empty() {
        breaks.push(TextBreakInfo {
            start: 0,
            end: text.len(),
            is_break: true,
        });
    }

    breaks
}

fn inline_item_width(item: &InlineItem) -> f64 {
    match item {
        InlineItem::Text(seg) => seg.width,
        InlineItem::Tab => 36.0, // Default tab width, will be resolved
        InlineItem::Image { width, .. } => *width,
        InlineItem::Marker(seg) => seg.width,
        InlineItem::LineBreak | InlineItem::PageBreak | InlineItem::ColumnBreak => 0.0,
    }
}

fn item_metrics(item: &InlineItem) -> (f64, f64, f64) {
    // Returns (width, ascent, descent)
    match item {
        InlineItem::Text(seg) => (seg.width, seg.ascent, seg.descent),
        InlineItem::Marker(seg) => (seg.width, seg.ascent, seg.descent),
        InlineItem::Tab => (36.0, 0.0, 0.0),
        InlineItem::Image { width, height, .. } => (*width, *height, 0.0),
        InlineItem::LineBreak | InlineItem::PageBreak | InlineItem::ColumnBreak => (0.0, 0.0, 0.0),
    }
}

fn inline_to_line_item(
    item: &InlineItem,
    current_x: f64,
    tab_stops: &[CT_TabStop],
    fm: &FontManager,
    font_ctx: Option<(FontId, f64)>,
) -> LineItem {
    match item {
        InlineItem::Text(seg) => LineItem::Text(seg.clone()),
        InlineItem::Marker(seg) => LineItem::Marker(seg.clone()),
        InlineItem::Tab => {
            let (tab_width, leader_char) = resolve_tab_width(current_x, tab_stops);
            let leader = leader_char.and_then(|ch| shape_leader(fm, font_ctx, ch, tab_width));
            LineItem::Tab {
                width: tab_width,
                leader,
            }
        }
        InlineItem::Image {
            width,
            height,
            embed_id,
        } => LineItem::Image {
            width: *width,
            height: *height,
            embed_id: embed_id.clone(),
        },
        InlineItem::LineBreak | InlineItem::PageBreak | InlineItem::ColumnBreak => LineItem::Tab {
            width: 0.0,
            leader: None,
        },
    }
}

/// Shape a leader character repeated to fill the given width.
fn shape_leader(
    fm: &FontManager,
    font_ctx: Option<(FontId, f64)>,
    leader_char: char,
    tab_width: f64,
) -> Option<TextSegment> {
    let (font_id, font_size) = font_ctx?;
    if tab_width < 1.0 {
        return None;
    }

    // Shape a single leader character to get its advance width
    let single = String::from(leader_char);
    let shaped = fm.shape_text(font_id, &single, font_size).ok()?;
    if shaped.glyph_ids.is_empty() {
        return None;
    }
    let char_advance = shaped.advances[0];
    if char_advance < 0.5 {
        return None;
    }

    // Add a small gap between leader chars (about 50% of char width for dots, less for others)
    let spacing = match leader_char {
        '.' | '\u{00B7}' => char_advance * 0.5,
        _ => char_advance * 0.15,
    };
    let step = char_advance + spacing;
    let count = ((tab_width - spacing) / step).floor() as usize;
    if count == 0 {
        return None;
    }

    // Build the repeated leader text and glyph arrays
    let leader_text: String = std::iter::repeat_n(leader_char, count).collect();
    let mut glyph_ids = Vec::with_capacity(count);
    let mut advances = Vec::with_capacity(count);
    for i in 0..count {
        glyph_ids.push(shaped.glyph_ids[0]);
        if i + 1 < count {
            advances.push(char_advance + spacing);
        } else {
            advances.push(char_advance);
        }
    }

    let metrics = fm.metrics(font_id, font_size).ok()?;

    Some(TextSegment {
        text: leader_text,
        font_id,
        font_size,
        glyph_ids,
        advances,
        width: tab_width, // fill the entire tab gap
        ascent: metrics.ascent,
        descent: metrics.descent,
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
        footnote_id: None,
    })
}

/// Resolve tab stop width and leader character based on current x position and defined stops.
fn resolve_tab_width(current_x: f64, tab_stops: &[CT_TabStop]) -> (f64, Option<char>) {
    use rdocx_oxml::shared::ST_TabLeader;

    // Find the next tab stop after the current position
    for stop in tab_stops {
        let stop_pos = stop.pos.to_pt();
        if stop_pos > current_x {
            let width = match stop.val {
                ST_TabJc::Left => stop_pos - current_x,
                ST_TabJc::Center => (stop_pos - current_x).max(0.0),
                ST_TabJc::Right => (stop_pos - current_x).max(0.0),
                _ => stop_pos - current_x,
            };
            let leader = stop.leader.and_then(|l| match l {
                ST_TabLeader::Dot => Some('.'),
                ST_TabLeader::Hyphen => Some('-'),
                ST_TabLeader::Underscore => Some('_'),
                ST_TabLeader::MiddleDot => Some('\u{00B7}'),
                ST_TabLeader::Heavy => Some('_'),
                ST_TabLeader::None => None,
            });
            return (width, leader);
        }
    }
    // Default tab stops every 0.5 inches (36pt)
    let default_interval = 36.0;
    let next_stop = ((current_x / default_interval).floor() + 1.0) * default_interval;
    (next_stop - current_x, None)
}

fn compute_first_line_width(params: &LineBreakParams) -> f64 {
    if params.ind_hanging > 0.0 {
        // Hanging indent: first line has MORE width (extends left)
        params.available_width - params.ind_left - params.ind_right + params.ind_hanging
    } else {
        params.available_width - params.ind_left - params.ind_right - params.ind_first_line
    }
}

fn compute_subsequent_line_width(params: &LineBreakParams) -> f64 {
    params.available_width - params.ind_left - params.ind_right
}

fn first_line_indent(params: &LineBreakParams) -> f64 {
    if params.ind_hanging > 0.0 {
        params.ind_left - params.ind_hanging
    } else {
        params.ind_left + params.ind_first_line
    }
}

fn subsequent_line_indent(params: &LineBreakParams) -> f64 {
    params.ind_left
}

/// Compute line height based on spacing rules.
fn compute_line_height(ascent: f64, descent: f64, params: &LineBreakParams) -> f64 {
    let natural = ascent + descent;
    let natural = if natural < 1.0 { 12.0 } else { natural }; // minimum for empty lines

    match (params.line_spacing, params.line_rule.as_deref()) {
        (Some(spacing), Some("exact")) => {
            // Exact: use the specified value
            spacing.to_pt()
        }
        (Some(spacing), Some("atLeast")) => {
            // At least: max of natural and specified
            natural.max(spacing.to_pt())
        }
        (Some(spacing), _) => {
            // "auto" or default: spacing is in 240ths of a line
            // 240 = single spacing, 480 = double, etc.
            let factor = spacing.0 as f64 / 240.0;
            natural * factor
        }
        (None, _) => {
            // Default: single spacing
            natural
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn make_text_segment(text: &str, width: f64) -> TextSegment {
        TextSegment {
            text: text.to_string(),
            font_id: FontId(0),
            font_size: 12.0,
            glyph_ids: vec![],
            advances: vec![],
            width,
            ascent: 10.0,
            descent: 3.0,
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
            footnote_id: None,
        }
    }

    #[test]
    fn empty_paragraph_gets_one_line() {
        let fm = FontManager::new();
        let lines = break_into_lines(&[], &LineBreakParams::default(), &fm).unwrap();
        assert_eq!(lines.len(), 1);
        assert!(lines[0].is_last);
        assert!(lines[0].items.is_empty());
    }

    #[test]
    fn single_word_fits_one_line() {
        let fm = FontManager::new();
        let items = vec![InlineItem::Text(make_text_segment("Hello", 50.0))];
        let lines = break_into_lines(&items, &LineBreakParams::default(), &fm).unwrap();
        assert_eq!(lines.len(), 1);
        assert!(lines[0].is_last);
    }

    #[test]
    fn words_wrap_to_multiple_lines() {
        let fm = FontManager::new();
        let mut items = Vec::new();
        // Each word is 200pt wide, line is 468pt → should wrap
        items.push(InlineItem::Text(make_text_segment("Word1", 200.0)));
        items.push(InlineItem::Text(make_text_segment("Word2", 200.0)));
        items.push(InlineItem::Text(make_text_segment("Word3", 200.0)));

        let lines = break_into_lines(&items, &LineBreakParams::default(), &fm).unwrap();
        assert!(lines.len() >= 2);
    }

    #[test]
    fn forced_line_break() {
        let fm = FontManager::new();
        let items = vec![
            InlineItem::Text(make_text_segment("Before", 50.0)),
            InlineItem::LineBreak,
            InlineItem::Text(make_text_segment("After", 50.0)),
        ];
        let lines = break_into_lines(&items, &LineBreakParams::default(), &fm).unwrap();
        assert!(lines.len() >= 2);
    }

    #[test]
    fn line_height_exact() {
        let params = LineBreakParams {
            line_spacing: Some(Twips::from_pt(24.0)),
            line_rule: Some("exact".to_string()),
            ..Default::default()
        };
        let h = compute_line_height(10.0, 3.0, &params);
        assert!((h - 24.0).abs() < 0.01);
    }

    #[test]
    fn line_height_auto() {
        let params = LineBreakParams {
            line_spacing: Some(Twips(480)), // double spacing
            line_rule: Some("auto".to_string()),
            ..Default::default()
        };
        let h = compute_line_height(10.0, 3.0, &params);
        assert!((h - 26.0).abs() < 0.01); // 13 * 2.0
    }

    #[test]
    fn first_line_indent() {
        let params = LineBreakParams {
            ind_first_line: 36.0,
            ..Default::default()
        };
        let first_w = compute_first_line_width(&params);
        let subseq_w = compute_subsequent_line_width(&params);
        assert!(first_w < subseq_w);
    }

    #[test]
    fn hanging_indent() {
        let params = LineBreakParams {
            ind_left: 36.0,
            ind_hanging: 36.0,
            ..Default::default()
        };
        let first_indent = super::first_line_indent(&params);
        let subseq_indent = super::subsequent_line_indent(&params);
        assert!(first_indent < subseq_indent);
    }

    #[test]
    fn tab_stop_resolution() {
        let stops = vec![CT_TabStop::new(ST_TabJc::Left, Twips::from_pt(72.0))];
        let (w, leader) = resolve_tab_width(36.0, &stops);
        assert!((w - 36.0).abs() < 0.01);
        assert!(leader.is_none());
    }

    #[test]
    fn default_tab_stops() {
        let (w, _) = resolve_tab_width(10.0, &[]);
        assert!((w - 26.0).abs() < 0.01); // next stop at 36pt
    }

    #[test]
    fn tab_stop_with_dot_leader() {
        use rdocx_oxml::shared::ST_TabLeader;
        let stops = vec![CT_TabStop {
            val: ST_TabJc::Right,
            pos: Twips::from_pt(400.0),
            leader: Some(ST_TabLeader::Dot),
        }];
        let (w, leader) = resolve_tab_width(100.0, &stops);
        assert!((w - 300.0).abs() < 0.01);
        assert_eq!(leader, Some('.'));
    }
}
