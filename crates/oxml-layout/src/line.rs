//! Line breaking: converts inline items into laid-out lines.
//!
//! Uses a greedy algorithm with unicode-linebreak for break opportunities.

use crate::error::{LayoutError, Result};
use crate::font::{FontManager, MultilingualTextSegment, TextDirection, explicit_direction_levels};
use crate::output::{Color, FieldKind, FontId, GroupElement, MediaId, SourceSpan, StructureId};

/// A tab stop positioned in typographic points.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct TabStop {
    pub pos_pt: f64,
    pub align: TabAlign,
    pub leader: Option<TabLeader>,
}

/// Paragraph alignment.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Align {
    Start,
    Center,
    End,
    Justify,
    Distribute,
}

/// Alignment relative to a tab stop.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TabAlign {
    Left,
    Center,
    Right,
    Decimal,
    Bar,
}

/// Leader style used to fill a tab gap.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TabLeader {
    None,
    Dot,
    Hyphen,
    Underscore,
    Heavy,
    MiddleDot,
}

/// Underline style applied to a text segment.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Underline {
    Single,
    Words,
    Double,
    Thick,
    Dotted,
    Dash,
    DotDash,
    DotDotDash,
    Wave,
}

/// Line height rule.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum LineSpacing {
    Single,
    /// A multiple of the largest text point size on the line.
    Multiple(f64),
    Exact(f64),
    AtLeast(f64),
}

/// An inline item to be placed on a line.
#[derive(Debug, Clone)]
#[non_exhaustive]
pub enum InlineItem {
    /// A shaped text segment.
    Text(TextSegment),
    /// A validated script, font, and bidi-level text span.
    MultilingualText(MultilingualTextSegment),
    /// A shaped text segment eligible for language-aware automatic hyphenation.
    HyphenatedText {
        segment: TextSegment,
        language: String,
    },
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
        media_id: MediaId,
    },
    /// A backend-neutral group with child-local coordinates.
    Group {
        width: f64,
        height: f64,
        /// Distance from the group top to its text baseline. `None` keeps the
        /// established top-aligned drawing behavior.
        baseline: Option<f64>,
        group: GroupElement,
    },
    /// An informative drawing carried to a semantic output container.
    Figure {
        item: Box<InlineItem>,
        alternate_text: String,
        structure_id: Option<StructureId>,
    },
    /// A numbering marker (rendered before the first line).
    Marker(TextSegment),
}

/// Which stream a note reference belongs to.
///
/// A reference carries only a number in the markup, and the two streams
/// number independently, so a document can hold a footnote and an endnote
/// that share a number. Without the stream the two are indistinguishable and
/// one silently shadows the other.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum NoteStream {
    /// Rendered at the foot of the page carrying the reference.
    Footnote,
    /// Rendered at the end of the document.
    Endnote,
}

/// A reference to one note, unique across both streams.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct NoteRef {
    pub stream: NoteStream,
    pub id: i32,
}

/// A shaped text segment with associated formatting.
#[derive(Debug, Clone)]
pub struct TextSegment {
    pub text: String,
    /// Requested character direction before paragraph-wide bidi resolution.
    pub direction: TextDirection,
    /// Exact source range for this segment, when directly attributable.
    pub source: Option<SourceSpan>,
    pub font_id: FontId,
    pub font_size: f64,
    pub glyph_ids: Vec<u16>,
    pub advances: Vec<f64>,
    pub width: f64,
    pub ascent: f64,
    pub descent: f64,
    /// Additional font leading included in the natural line advance.
    pub line_gap: f64,
    pub color: Color,
    pub bold: bool,
    pub italic: bool,
    /// Underline style (None = no underline).
    pub underline: Option<Underline>,
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
    /// If this segment is a note reference marker, which note it points at.
    pub note: Option<NoteRef>,
}

/// A single item positioned on a line.
#[derive(Debug, Clone)]
#[non_exhaustive]
pub enum LineItem {
    Text(TextSegment),
    /// A validated script, font, and bidi-level text span.
    MultilingualText(MultilingualTextSegment),
    Tab {
        width: f64,
        /// Pre-shaped leader text to fill the tab gap (e.g., dots, hyphens).
        leader: Option<TextSegment>,
    },
    Image {
        width: f64,
        height: f64,
        media_id: MediaId,
    },
    Group {
        width: f64,
        height: f64,
        /// Distance from the group top to its text baseline.
        baseline: Option<f64>,
        group: GroupElement,
    },
    /// An informative drawing carried to a semantic output container.
    Figure {
        item: Box<LineItem>,
        alternate_text: String,
        structure_id: Option<StructureId>,
    },
    Marker(TextSegment),
}

impl LineItem {
    pub fn width(&self) -> f64 {
        match self {
            LineItem::Text(seg) => seg.width,
            LineItem::MultilingualText(seg) => seg.width(),
            LineItem::Tab { width, .. } => *width,
            LineItem::Image { width, .. } => *width,
            LineItem::Group { width, .. } => *width,
            LineItem::Figure { item, .. } => item.width(),
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
    /// Effective leading needed to preserve the tallest run's natural advance.
    pub line_gap: f64,
    /// Total line height.
    pub height: f64,
    /// Left indent for this line.
    pub indent_left: f64,
    /// Available width this line was laid out against.
    pub available_width: f64,
    /// Whether this is the last line of the paragraph.
    pub is_last: bool,
}

impl LayoutLine {
    /// Distance from the line-box top to its text baseline.
    pub fn baseline_offset(&self) -> f64 {
        let leading = self.height - self.ascent - self.descent;
        self.ascent + if leading >= 0.0 { leading / 2.0 } else { 0.0 }
    }
}

/// Parameters for line breaking.
#[derive(Debug, Clone)]
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
    pub tab_stops: Vec<TabStop>,
    /// Line spacing rule and value.
    pub line_spacing: LineSpacing,
    /// Paragraph justification.
    pub jc: Option<Align>,
    /// Whether width overflow may create automatic line breaks.
    pub wrap: bool,
    /// Extra width kept clear at the start of individual lines, by line index.
    ///
    /// This is how a floating drawing pushes text aside. An empty vector, the
    /// default, reserves nothing and reproduces unwrapped line breaking
    /// exactly.
    pub line_prefix_widths: Vec<f64>,
    /// Extra width kept clear at the end of individual lines, by line index.
    pub line_suffix_widths: Vec<f64>,
}

impl Default for LineBreakParams {
    fn default() -> Self {
        LineBreakParams {
            available_width: 468.0, // US Letter with 1" margins
            line_prefix_widths: Vec::new(),
            line_suffix_widths: Vec::new(),
            ind_left: 0.0,
            ind_right: 0.0,
            ind_first_line: 0.0,
            ind_hanging: 0.0,
            tab_stops: Vec::new(),
            line_spacing: LineSpacing::Single,
            jc: None,
            wrap: true,
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
            line_gap: 0.0,
            height: compute_line_height(0.0, 0.0, 0.0, 0.0, params),
            indent_left: line_indent_at(params, 0, true),
            available_width: line_width_at(params, 0, true),
            is_last: true,
        }]);
    }

    let mut lines: Vec<LayoutLine> = Vec::new();
    let mut current_items: Vec<LineItem> = Vec::new();
    let mut current_width: f64 = 0.0;
    let mut current_ascent: f64 = 0.0;
    let mut current_descent: f64 = 0.0;
    let mut current_natural_height: f64 = 0.0;
    let mut current_font_size: f64 = 0.0;
    // The line index drives the per-line reservations a floating drawing
    // creates, so it is tracked rather than a plain first-or-not flag.
    let mut line_index = 0usize;
    let mut is_first_line = true;

    let first_line_width = line_width_at(params, 0, true);

    let mut line_avail = first_line_width;

    // Track the most recent font context for shaping tab leaders
    let mut font_ctx: Option<(FontId, f64)> = None;
    // Initialize from the first text segment if available
    for item in items {
        if let InlineItem::Text(seg)
        | InlineItem::HyphenatedText { segment: seg, .. }
        | InlineItem::Marker(seg) = item
        {
            font_ctx = Some((seg.font_id, seg.font_size));
            break;
        }
        if let InlineItem::MultilingualText(seg) = item {
            font_ctx = Some((seg.font_id(), seg.base().font_size));
            break;
        }
    }

    // Build breakable segments from inline items
    let mut segments = std::collections::VecDeque::from(build_breakable_segments(items, fm)?);

    while let Some(seg) = segments.pop_front() {
        match seg {
            BreakableSegment::Items(seg_items) => {
                let seg_width: f64 = seg_items.iter().map(inline_item_width).sum();

                if params.wrap
                    && !current_items.is_empty()
                    && current_width + seg_width > line_avail + 0.01
                {
                    // Finish current line
                    let indent = line_indent_at(params, line_index, is_first_line);
                    let line_gap =
                        effective_line_gap(current_ascent, current_descent, current_natural_height);
                    lines.push(LayoutLine {
                        items: std::mem::take(&mut current_items),
                        width: current_width,
                        ascent: current_ascent,
                        descent: current_descent,
                        line_gap,
                        height: compute_line_height(
                            current_ascent,
                            current_descent,
                            line_gap,
                            current_font_size,
                            params,
                        ),
                        indent_left: indent,
                        available_width: line_avail,
                        is_last: false,
                    });
                    current_width = 0.0;
                    current_ascent = 0.0;
                    current_descent = 0.0;
                    current_natural_height = 0.0;
                    current_font_size = 0.0;
                    is_first_line = false;
                    line_index += 1;
                    line_avail = line_width_at(params, line_index, false);
                }

                // Add segment items to current line
                for item in &seg_items {
                    let (w, a, d, natural_height, font_size) = item_metrics(item);
                    current_width += w;
                    if a > current_ascent {
                        current_ascent = a;
                    }
                    if d > current_descent {
                        current_descent = d;
                    }
                    current_natural_height = current_natural_height.max(natural_height);
                    current_font_size = current_font_size.max(font_size);
                    // Update font context from text segments
                    if let InlineItem::Text(seg) | InlineItem::Marker(seg) = item {
                        font_ctx = Some((seg.font_id, seg.font_size));
                    } else if let InlineItem::MultilingualText(seg) = item {
                        font_ctx = Some((seg.font_id(), seg.base().font_size));
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
            BreakableSegment::Hyphenated(boxed) => {
                let HyphenatedSegment {
                    segment,
                    break_points,
                } = *boxed;
                let fits = current_width + segment.width <= line_avail + 0.01;
                if !params.wrap || fits {
                    let item = InlineItem::Text(segment);
                    let (w, a, d, natural_height, font_size) = item_metrics(&item);
                    current_width += w;
                    current_ascent = current_ascent.max(a);
                    current_descent = current_descent.max(d);
                    current_natural_height = current_natural_height.max(natural_height);
                    current_font_size = current_font_size.max(font_size);
                    font_ctx = Some((segment_font_id(&item), segment_font_size(&item)));
                    current_items.push(inline_to_line_item(
                        &item,
                        current_width,
                        &params.tab_stops,
                        fm,
                        font_ctx,
                    ));
                    continue;
                }

                if let Some(FittingHyphenation {
                    prefix,
                    hyphen,
                    remainder,
                    remaining_points,
                }) = fitting_hyphenation(&segment, &break_points, current_width, line_avail, fm)?
                {
                    for text in [prefix, hyphen] {
                        let item = InlineItem::Text(text);
                        let (w, a, d, natural_height, font_size) = item_metrics(&item);
                        current_width += w;
                        current_ascent = current_ascent.max(a);
                        current_descent = current_descent.max(d);
                        current_natural_height = current_natural_height.max(natural_height);
                        current_font_size = current_font_size.max(font_size);
                        font_ctx = Some((segment_font_id(&item), segment_font_size(&item)));
                        current_items.push(inline_to_line_item(
                            &item,
                            current_width,
                            &params.tab_stops,
                            fm,
                            font_ctx,
                        ));
                    }

                    let indent = line_indent_at(params, line_index, is_first_line);
                    let line_gap =
                        effective_line_gap(current_ascent, current_descent, current_natural_height);
                    lines.push(LayoutLine {
                        items: std::mem::take(&mut current_items),
                        width: current_width,
                        ascent: current_ascent,
                        descent: current_descent,
                        line_gap,
                        height: compute_line_height(
                            current_ascent,
                            current_descent,
                            line_gap,
                            current_font_size,
                            params,
                        ),
                        indent_left: indent,
                        available_width: line_avail,
                        is_last: false,
                    });
                    current_width = 0.0;
                    current_ascent = 0.0;
                    current_descent = 0.0;
                    current_natural_height = 0.0;
                    current_font_size = 0.0;
                    is_first_line = false;
                    line_index += 1;
                    line_avail = line_width_at(params, line_index, false);
                    segments.push_front(BreakableSegment::Hyphenated(Box::new(
                        HyphenatedSegment {
                            segment: remainder,
                            break_points: remaining_points,
                        },
                    )));
                } else if !current_items.is_empty() {
                    let indent = line_indent_at(params, line_index, is_first_line);
                    let line_gap =
                        effective_line_gap(current_ascent, current_descent, current_natural_height);
                    lines.push(LayoutLine {
                        items: std::mem::take(&mut current_items),
                        width: current_width,
                        ascent: current_ascent,
                        descent: current_descent,
                        line_gap,
                        height: compute_line_height(
                            current_ascent,
                            current_descent,
                            line_gap,
                            current_font_size,
                            params,
                        ),
                        indent_left: indent,
                        available_width: line_avail,
                        is_last: false,
                    });
                    current_width = 0.0;
                    current_ascent = 0.0;
                    current_descent = 0.0;
                    current_natural_height = 0.0;
                    current_font_size = 0.0;
                    is_first_line = false;
                    line_index += 1;
                    line_avail = line_width_at(params, line_index, false);
                    segments.push_front(BreakableSegment::Hyphenated(Box::new(
                        HyphenatedSegment {
                            segment,
                            break_points,
                        },
                    )));
                } else {
                    let item = InlineItem::Text(segment);
                    let (w, a, d, natural_height, font_size) = item_metrics(&item);
                    current_width += w;
                    current_ascent = current_ascent.max(a);
                    current_descent = current_descent.max(d);
                    current_natural_height = current_natural_height.max(natural_height);
                    current_font_size = current_font_size.max(font_size);
                    font_ctx = Some((segment_font_id(&item), segment_font_size(&item)));
                    current_items.push(inline_to_line_item(
                        &item,
                        current_width,
                        &params.tab_stops,
                        fm,
                        font_ctx,
                    ));
                }
            }
            BreakableSegment::ForcedBreak(break_type) => {
                let indent = line_indent_at(params, line_index, is_first_line);
                let line_gap =
                    effective_line_gap(current_ascent, current_descent, current_natural_height);
                lines.push(LayoutLine {
                    items: std::mem::take(&mut current_items),
                    width: current_width,
                    ascent: current_ascent,
                    descent: current_descent,
                    line_gap,
                    height: compute_line_height(
                        current_ascent,
                        current_descent,
                        line_gap,
                        current_font_size,
                        params,
                    ),
                    indent_left: indent,
                    available_width: line_avail,
                    is_last: matches!(break_type, ForcedBreakType::Page | ForcedBreakType::Column),
                });
                current_width = 0.0;
                current_ascent = 0.0;
                current_descent = 0.0;
                current_natural_height = 0.0;
                current_font_size = 0.0;
                is_first_line = false;
                line_index += 1;
                line_avail = line_width_at(params, line_index, false);
            }
        }
    }

    // Flush remaining items as the last line
    let indent = line_indent_at(params, line_index, is_first_line);
    let line_gap = effective_line_gap(current_ascent, current_descent, current_natural_height);
    lines.push(LayoutLine {
        items: current_items,
        width: current_width,
        ascent: current_ascent,
        descent: current_descent,
        line_gap,
        height: compute_line_height(
            current_ascent,
            current_descent,
            line_gap,
            current_font_size,
            params,
        ),
        indent_left: indent,
        available_width: line_avail,
        is_last: true,
    });

    Ok(lines)
}

/// Break rich text in logical order, then reorder each completed line for painting.
pub fn break_multilingual_into_lines(
    items: &[InlineItem],
    params: &LineBreakParams,
    fm: &FontManager,
    base_direction: TextDirection,
) -> Result<Vec<LayoutLine>> {
    let mut lines = break_into_lines(items, params, fm)?;
    let mut paragraph_text = String::new();
    let mut line_maps = Vec::with_capacity(lines.len());
    let mut has_text = false;
    for line in &lines {
        let line_start = paragraph_text.len();
        let mut positions = Vec::new();
        for (index, item) in line.items.iter().enumerate() {
            let (text, direction, shaped_level) = match item {
                LineItem::Text(segment) | LineItem::Marker(segment) => {
                    (segment.text.as_str(), segment.direction, None)
                }
                LineItem::MultilingualText(segment) => (
                    segment.text(),
                    segment.base().direction,
                    Some(unicode_bidi::Level::new(segment.bidi_level()).map_err(|_| {
                        LayoutError::Layout("rich text carried an invalid bidi level".to_owned())
                    })?),
                ),
                LineItem::Tab { .. } => ("\t", TextDirection::Auto, None),
                LineItem::Image { .. } | LineItem::Group { .. } | LineItem::Figure { .. } => {
                    ("\u{fffc}", TextDirection::Auto, None)
                }
            };
            has_text |= !text.is_empty();
            let start = paragraph_text.len();
            paragraph_text.push_str(text);
            let end = paragraph_text.len();
            positions.push((
                index,
                start,
                end,
                direction,
                text.chars().all(char::is_whitespace),
                shaped_level,
            ));
        }
        line_maps.push((line_start..paragraph_text.len(), positions));
    }
    if !has_text {
        return Ok(lines);
    }
    let paragraph_level = match base_direction {
        TextDirection::Auto => None,
        TextDirection::LeftToRight => Some(unicode_bidi::Level::ltr()),
        TextDirection::RightToLeft => Some(unicode_bidi::Level::rtl()),
    };
    let bidi = unicode_bidi::BidiInfo::new(&paragraph_text, paragraph_level);
    let [paragraph] = bidi.paragraphs.as_slice() else {
        return Err(LayoutError::Layout(
            "multilingual line layout requires one bidi paragraph".to_owned(),
        ));
    };
    for (line, (line_range, mapped_positions)) in lines.iter_mut().zip(line_maps) {
        let positions = mapped_positions
            .iter()
            .map(|(index, _, _, _, _, _)| *index)
            .collect::<Vec<_>>();
        if positions.is_empty() {
            continue;
        }
        let adjusted_levels = bidi.reordered_levels(paragraph, line_range);
        let levels = mapped_positions
            .into_iter()
            .map(|(_, start, end, direction, whitespace, shaped_level)| {
                let adjusted = adjusted_levels.get(start).copied().ok_or_else(|| {
                    LayoutError::Layout(
                        "multilingual line range exceeded its bidi paragraph".to_owned(),
                    )
                })?;
                Ok(if whitespace {
                    adjusted
                } else if let Some(shaped_level) = shaped_level {
                    shaped_level
                } else if direction == TextDirection::Auto {
                    adjusted
                } else {
                    explicit_direction_levels(
                        &paragraph_text[start..end],
                        direction,
                        paragraph.level,
                    )?
                    .first()
                    .copied()
                    .unwrap_or(adjusted)
                })
            })
            .collect::<Result<Vec<_>>>()?;
        let visual_order = unicode_bidi::BidiInfo::reorder_visual(&levels);
        let logical_items = positions
            .iter()
            .zip(&levels)
            .map(|(index, level)| match &line.items[*index] {
                LineItem::MultilingualText(segment) => Ok(LineItem::MultilingualText(
                    multilingual_segment_with_level(segment, *level)?,
                )),
                item => Ok(item.clone()),
            })
            .collect::<Result<Vec<_>>>()?;
        for (visual_slot, logical_index) in positions.into_iter().zip(visual_order) {
            line.items[visual_slot] = logical_items[logical_index].clone();
        }
    }
    Ok(lines)
}

fn multilingual_segment_with_level(
    segment: &MultilingualTextSegment,
    level: unicode_bidi::Level,
) -> Result<MultilingualTextSegment> {
    let parity_changed = segment.bidi_level() % 2 != level.number() % 2;
    let cluster_order = if parity_changed {
        (0..segment.clusters().len()).rev().collect::<Vec<_>>()
    } else {
        (0..segment.clusters().len()).collect::<Vec<_>>()
    };
    let mut glyph_ids = Vec::with_capacity(segment.glyph_ids().len());
    let mut x_advances = Vec::with_capacity(segment.x_advances().len());
    let mut y_advances = Vec::with_capacity(segment.y_advances().len());
    let mut x_offsets = Vec::with_capacity(segment.x_offsets().len());
    let mut y_offsets = Vec::with_capacity(segment.y_offsets().len());
    let mut clusters = Vec::with_capacity(segment.clusters().len());
    for index in cluster_order {
        let cluster = &segment.clusters()[index];
        let glyph_range = cluster.glyph_start as usize..cluster.glyph_end as usize;
        let glyph_start = glyph_ids.len() as u32;
        glyph_ids.extend_from_slice(&segment.glyph_ids()[glyph_range.clone()]);
        x_advances.extend_from_slice(&segment.x_advances()[glyph_range.clone()]);
        y_advances.extend_from_slice(&segment.y_advances()[glyph_range.clone()]);
        x_offsets.extend_from_slice(&segment.x_offsets()[glyph_range.clone()]);
        y_offsets.extend_from_slice(&segment.y_offsets()[glyph_range]);
        clusters.push(crate::font::GlyphCluster {
            glyph_start,
            glyph_end: glyph_ids.len() as u32,
            char_start: cluster.char_start,
            char_end: cluster.char_end,
        });
    }
    let mut base = segment.base().clone();
    base.glyph_ids = glyph_ids;
    base.advances = x_advances.clone();

    MultilingualTextSegment::new(
        base,
        segment.logical_index(),
        segment.language().map(str::to_owned),
        segment.script(),
        if level.is_rtl() {
            TextDirection::RightToLeft
        } else {
            TextDirection::LeftToRight
        },
        level.number(),
        x_advances,
        y_advances,
        x_offsets,
        y_offsets,
        clusters,
        segment.break_after(),
    )
}

// ---- Internal helpers ----

#[derive(Debug)]
enum BreakableSegment {
    /// A group of items that should be kept together (word or cluster).
    Items(Vec<InlineItem>),
    /// One language-aware text chunk and its byte-index break candidates.
    Hyphenated(Box<HyphenatedSegment>),
    /// A forced break.
    ForcedBreak(ForcedBreakType),
}

#[derive(Debug)]
struct HyphenatedSegment {
    segment: TextSegment,
    break_points: Vec<usize>,
}

struct FittingHyphenation {
    prefix: TextSegment,
    hyphen: TextSegment,
    remainder: TextSegment,
    remaining_points: Vec<usize>,
}

#[derive(Debug)]
enum ForcedBreakType {
    Line,
    Page,
    Column,
}

/// Whether a complex-script line may break between two logical characters.
pub(crate) fn multilingual_break_allowed(before: char, after: char) -> bool {
    const OPENING: &[char] = &['(', '[', '{', '〈', '《', '「', '『', '【', '〔', '〖'];
    const CLOSING_OR_NONSTARTER: &[char] = &[
        ')', ']', '}', '〉', '》', '」', '』', '】', '〕', '〗', '、', '。', '，', '．', '！',
        '？', '：', '；', '％', '‰',
    ];
    !OPENING.contains(&before) && !CLOSING_OR_NONSTARTER.contains(&after)
}

/// Build breakable segments by finding break opportunities in text.
///
/// Text items are split at unicode line-break opportunities (word boundaries,
/// hyphens, etc.). Non-text items (tabs, images, markers) are treated as
/// atomic units with break opportunities around them.
fn build_breakable_segments(
    items: &[InlineItem],
    fm: &FontManager,
) -> Result<Vec<BreakableSegment>> {
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
            InlineItem::Text(seg) | InlineItem::HyphenatedText { segment: seg, .. } => {
                if seg.text.is_empty() {
                    if !current_group.is_empty() {
                        segments.push(BreakableSegment::Items(std::mem::take(&mut current_group)));
                    }
                    segments.push(BreakableSegment::Items(vec![item.clone()]));
                    continue;
                }
                // Use unicode-linebreak to find break opportunities within text
                let breaks = split_text_at_break_opportunities(seg);
                let spacing =
                    if breaks.len() == 1 && breaks[0].start == 0 && breaks[0].end == seg.text.len()
                    {
                        0.0
                    } else {
                        let original = fm.shape_text(seg.font_id, &seg.text, seg.font_size)?;
                        if original.advances.len() == seg.advances.len()
                            && !original.advances.is_empty()
                        {
                            (seg.width - original.width) / original.advances.len() as f64
                        } else {
                            0.0
                        }
                    };

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
                    let sub_item = split_text_subsegment(seg, tb.start, tb.end, spacing, fm)?;
                    let language = match item {
                        InlineItem::HyphenatedText { language, .. } => Some(language.as_str()),
                        _ => None,
                    };
                    let break_points = language.map_or_else(Vec::new, |language| {
                        hyphenation_opportunities(language, chunk)
                    });
                    if break_points.is_empty() {
                        current_group.push(sub_item);
                    } else {
                        if !current_group.is_empty() {
                            segments
                                .push(BreakableSegment::Items(std::mem::take(&mut current_group)));
                        }
                        let InlineItem::Text(segment) = sub_item else {
                            unreachable!("split text always returns text")
                        };
                        segments.push(BreakableSegment::Hyphenated(Box::new(HyphenatedSegment {
                            segment,
                            break_points,
                        })));
                    }

                    if tb.is_break {
                        segments.push(BreakableSegment::Items(std::mem::take(&mut current_group)));
                    }
                }

                // Flush any remaining
                if !current_group.is_empty() {
                    segments.push(BreakableSegment::Items(std::mem::take(&mut current_group)));
                }
            }
            InlineItem::MultilingualText(segment) => {
                current_group.push(InlineItem::MultilingualText(segment.clone()));
                if segment.break_after() {
                    segments.push(BreakableSegment::Items(std::mem::take(&mut current_group)));
                }
            }
            InlineItem::Marker(_)
            | InlineItem::Image { .. }
            | InlineItem::Group { .. }
            | InlineItem::Figure { .. } => {
                current_group.push(item.clone());
            }
        }
    }

    if !current_group.is_empty() {
        segments.push(BreakableSegment::Items(current_group));
    }

    Ok(segments)
}

/// Create a sub-segment InlineItem from a byte range within a TextSegment.
///
/// Reshapes the selected text and preserves formatting from the parent segment.
fn split_text_subsegment(
    seg: &TextSegment,
    byte_start: usize,
    byte_end: usize,
    spacing: f64,
    fm: &FontManager,
) -> Result<InlineItem> {
    // If this is the full segment, just clone it
    if byte_start == 0 && byte_end == seg.text.len() {
        return Ok(InlineItem::Text(seg.clone()));
    }

    let sub_text = seg.text[byte_start..byte_end].to_string();
    let mut shaped = fm.shape_text(seg.font_id, &sub_text, seg.font_size)?;
    for advance in &mut shaped.advances {
        *advance += spacing;
    }
    shaped.width += spacing * shaped.advances.len() as f64;

    let source = seg.source.map(|source| {
        let start = seg.text[..byte_start].chars().count() as u32;
        let end = seg.text[..byte_end].chars().count() as u32;
        SourceSpan {
            node: source.node,
            char_start: source.char_start + start,
            char_end: source.char_start + end,
        }
    });

    Ok(InlineItem::Text(TextSegment {
        text: sub_text,
        direction: seg.direction,
        source,
        font_id: seg.font_id,
        font_size: seg.font_size,
        glyph_ids: shaped.glyph_ids,
        advances: shaped.advances,
        width: shaped.width,
        ascent: seg.ascent,
        descent: seg.descent,
        line_gap: seg.line_gap,
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
        note: seg.note,
    }))
}

fn supported_hyphenation_language(language: &str) -> Option<hypher::Lang> {
    let primary = language.split('-').next()?;
    if primary.eq_ignore_ascii_case("en") {
        Some(hypher::Lang::English)
    } else if primary.eq_ignore_ascii_case("fr") {
        Some(hypher::Lang::French)
    } else if primary.eq_ignore_ascii_case("de") {
        Some(hypher::Lang::German)
    } else if primary.eq_ignore_ascii_case("es") {
        Some(hypher::Lang::Spanish)
    } else {
        None
    }
}

fn hyphenation_opportunities(language: &str, text: &str) -> Vec<usize> {
    let Some(language) = supported_hyphenation_language(language) else {
        return Vec::new();
    };
    let word_end = text
        .char_indices()
        .take_while(|(_, character)| character.is_alphabetic())
        .map(|(offset, character)| offset + character.len_utf8())
        .last()
        .unwrap_or(0);
    if word_end == 0 || text[word_end..].chars().any(char::is_alphabetic) {
        return Vec::new();
    }
    let word = &text[..word_end];
    let syllables = hypher::hyphenate(word, language).collect::<Vec<_>>();
    let mut offset = 0usize;
    syllables
        .iter()
        .take(syllables.len().saturating_sub(1))
        .map(|syllable| {
            offset += syllable.len();
            offset
        })
        .collect()
}

fn fitting_hyphenation(
    segment: &TextSegment,
    break_points: &[usize],
    current_width: f64,
    available_width: f64,
    fm: &FontManager,
) -> Result<Option<FittingHyphenation>> {
    let spacing = text_segment_spacing(segment, fm)?;
    let hyphen = generated_hyphen(segment, spacing, fm)?;
    for &break_point in break_points.iter().rev() {
        let InlineItem::Text(prefix) = split_text_subsegment(segment, 0, break_point, spacing, fm)?
        else {
            unreachable!("split text always returns text")
        };
        if current_width + prefix.width + hyphen.width > available_width + 0.01 {
            continue;
        }
        let InlineItem::Text(remainder) =
            split_text_subsegment(segment, break_point, segment.text.len(), spacing, fm)?
        else {
            unreachable!("split text always returns text")
        };
        let remaining_points = break_points
            .iter()
            .copied()
            .filter(|point| *point > break_point)
            .map(|point| point - break_point)
            .collect();
        return Ok(Some(FittingHyphenation {
            prefix,
            hyphen,
            remainder,
            remaining_points,
        }));
    }
    Ok(None)
}

fn text_segment_spacing(segment: &TextSegment, fm: &FontManager) -> Result<f64> {
    let original = fm.shape_text(segment.font_id, &segment.text, segment.font_size)?;
    Ok(
        if original.advances.len() == segment.advances.len() && !original.advances.is_empty() {
            (segment.width - original.width) / original.advances.len() as f64
        } else {
            0.0
        },
    )
}

fn generated_hyphen(segment: &TextSegment, spacing: f64, fm: &FontManager) -> Result<TextSegment> {
    let mut shaped = fm.shape_text(segment.font_id, "-", segment.font_size)?;
    for advance in &mut shaped.advances {
        *advance += spacing;
    }
    shaped.width += spacing * shaped.advances.len() as f64;
    Ok(TextSegment {
        text: "-".to_owned(),
        direction: segment.direction,
        source: None,
        font_id: segment.font_id,
        font_size: segment.font_size,
        glyph_ids: shaped.glyph_ids,
        advances: shaped.advances,
        width: shaped.width,
        ascent: segment.ascent,
        descent: segment.descent,
        line_gap: segment.line_gap,
        color: segment.color,
        bold: segment.bold,
        italic: segment.italic,
        underline: segment.underline,
        strike: segment.strike,
        dstrike: segment.dstrike,
        highlight: segment.highlight,
        baseline_offset: segment.baseline_offset,
        hyperlink_url: segment.hyperlink_url.clone(),
        field_kind: segment.field_kind,
        note: segment.note,
    })
}

fn segment_font_id(item: &InlineItem) -> FontId {
    match item {
        InlineItem::Text(segment) | InlineItem::HyphenatedText { segment, .. } => segment.font_id,
        _ => unreachable!("called only for text"),
    }
}

fn segment_font_size(item: &InlineItem) -> f64 {
    match item {
        InlineItem::Text(segment) | InlineItem::HyphenatedText { segment, .. } => segment.font_size,
        _ => unreachable!("called only for text"),
    }
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
        InlineItem::Text(seg) | InlineItem::HyphenatedText { segment: seg, .. } => seg.width,
        InlineItem::MultilingualText(seg) => seg.width(),
        InlineItem::Tab => 36.0, // Default tab width, will be resolved
        InlineItem::Image { width, .. } => *width,
        InlineItem::Group { width, .. } => *width,
        InlineItem::Figure { item, .. } => inline_item_width(item),
        InlineItem::Marker(seg) => seg.width,
        InlineItem::LineBreak | InlineItem::PageBreak | InlineItem::ColumnBreak => 0.0,
    }
}

fn item_metrics(item: &InlineItem) -> (f64, f64, f64, f64, f64) {
    // Returns (width, ascent, descent, natural height, text font size)
    match item {
        InlineItem::Text(seg) | InlineItem::HyphenatedText { segment: seg, .. } => (
            seg.width,
            seg.ascent,
            seg.descent,
            seg.ascent + seg.descent + seg.line_gap,
            seg.font_size,
        ),
        InlineItem::MultilingualText(seg) => (
            seg.width(),
            seg.base().ascent,
            seg.base().descent,
            seg.base().ascent + seg.base().descent + seg.base().line_gap,
            seg.base().font_size,
        ),
        InlineItem::Marker(seg) => (
            seg.width,
            seg.ascent,
            seg.descent,
            seg.ascent + seg.descent + seg.line_gap,
            0.0,
        ),
        InlineItem::Tab => (36.0, 0.0, 0.0, 0.0, 0.0),
        InlineItem::Image { width, height, .. } => (*width, *height, 0.0, *height, 0.0),
        InlineItem::Group {
            width,
            height,
            baseline,
            ..
        } => {
            let baseline = normalized_group_baseline(*height, *baseline).unwrap_or(*height);
            (*width, baseline, *height - baseline, *height, 0.0)
        }
        InlineItem::Figure { item, .. } => item_metrics(item),
        InlineItem::LineBreak | InlineItem::PageBreak | InlineItem::ColumnBreak => {
            (0.0, 0.0, 0.0, 0.0, 0.0)
        }
    }
}

fn normalized_group_baseline(height: f64, baseline: Option<f64>) -> Option<f64> {
    if !height.is_finite() {
        return None;
    }
    let upper = height.max(0.0);
    baseline
        .filter(|value| value.is_finite())
        .map(|value| value.clamp(0.0, upper))
}

fn inline_to_line_item(
    item: &InlineItem,
    current_x: f64,
    tab_stops: &[TabStop],
    fm: &FontManager,
    font_ctx: Option<(FontId, f64)>,
) -> LineItem {
    match item {
        InlineItem::Text(seg) | InlineItem::HyphenatedText { segment: seg, .. } => {
            LineItem::Text(seg.clone())
        }
        InlineItem::MultilingualText(seg) => LineItem::MultilingualText(seg.clone()),
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
            media_id,
        } => LineItem::Image {
            width: *width,
            height: *height,
            media_id: *media_id,
        },
        InlineItem::Group {
            width,
            height,
            baseline,
            group,
        } => LineItem::Group {
            width: *width,
            height: *height,
            baseline: normalized_group_baseline(*height, *baseline),
            group: group.clone(),
        },
        InlineItem::Figure {
            item,
            alternate_text,
            structure_id,
        } => LineItem::Figure {
            item: Box::new(inline_to_line_item(
                item, current_x, tab_stops, fm, font_ctx,
            )),
            alternate_text: alternate_text.clone(),
            structure_id: *structure_id,
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
        direction: TextDirection::Auto,
        source: None,
        font_id,
        font_size,
        glyph_ids,
        advances,
        width: tab_width, // fill the entire tab gap
        ascent: metrics.ascent,
        descent: metrics.descent,
        line_gap: metrics.line_gap,
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
    })
}

/// Resolve tab stop width and leader character based on current x position and defined stops.
fn resolve_tab_width(current_x: f64, tab_stops: &[TabStop]) -> (f64, Option<char>) {
    // Find the next tab stop after the current position
    for stop in tab_stops {
        let stop_pos = stop.pos_pt;
        if stop_pos > current_x {
            let width = match stop.align {
                TabAlign::Left => stop_pos - current_x,
                TabAlign::Center => (stop_pos - current_x).max(0.0),
                TabAlign::Right => (stop_pos - current_x).max(0.0),
                _ => stop_pos - current_x,
            };
            let leader = stop.leader.and_then(|l| match l {
                TabLeader::Dot => Some('.'),
                TabLeader::Hyphen => Some('-'),
                TabLeader::Underscore => Some('_'),
                TabLeader::MiddleDot => Some('\u{00B7}'),
                TabLeader::Heavy => Some('_'),
                TabLeader::None => None,
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

/// Width kept clear at the start of a given line.
fn line_prefix_width(params: &LineBreakParams, line_index: usize) -> f64 {
    params
        .line_prefix_widths
        .get(line_index)
        .copied()
        .unwrap_or(0.0)
}

/// Width kept clear at the end of a given line.
fn line_suffix_width(params: &LineBreakParams, line_index: usize) -> f64 {
    params
        .line_suffix_widths
        .get(line_index)
        .copied()
        .unwrap_or(0.0)
}

/// Usable width of a line, once anything floating beside it is taken out.
fn line_width_at(params: &LineBreakParams, line_index: usize, is_first_line: bool) -> f64 {
    let base = if is_first_line {
        compute_first_line_width(params)
    } else {
        compute_subsequent_line_width(params)
    };
    (base - line_prefix_width(params, line_index) - line_suffix_width(params, line_index)).max(0.0)
}

/// Where a line starts, once anything floating to its left is taken out.
fn line_indent_at(params: &LineBreakParams, line_index: usize, is_first_line: bool) -> f64 {
    let base = if is_first_line {
        first_line_indent(params)
    } else {
        subsequent_line_indent(params)
    };
    base + line_prefix_width(params, line_index)
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
fn effective_line_gap(ascent: f64, descent: f64, natural_height: f64) -> f64 {
    (natural_height - ascent - descent).max(0.0)
}

fn compute_line_height(
    ascent: f64,
    descent: f64,
    line_gap: f64,
    font_size: f64,
    params: &LineBreakParams,
) -> f64 {
    let natural = ascent + descent + line_gap;
    let natural = if natural < 1.0 { 12.0 } else { natural }; // minimum for empty lines
    let font_size = if font_size < 1.0 { 12.0 } else { font_size };

    match params.line_spacing {
        LineSpacing::Single => natural,
        LineSpacing::Multiple(factor) => font_size * factor,
        LineSpacing::Exact(points) => points,
        LineSpacing::AtLeast(points) => natural.max(points),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn make_text_segment(text: &str, width: f64) -> TextSegment {
        TextSegment {
            text: text.to_string(),
            direction: TextDirection::Auto,
            source: None,
            font_id: FontId(0),
            font_size: 12.0,
            glyph_ids: vec![],
            advances: vec![],
            width,
            ascent: 10.0,
            descent: 3.0,
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
        }
    }

    fn deterministic_font_manager() -> FontManager {
        FontManager::new_deterministic().expect("bundled fonts should load")
    }

    fn shaped_text_segment(fm: &mut FontManager, text: &str, spacing: f64) -> TextSegment {
        let font_id = fm
            .resolve_font(Some("Carlito"), false, false)
            .expect("bundled Carlito should resolve");
        let metrics = fm.metrics(font_id, 30.0).expect("Carlito metrics");
        let mut shaped = fm.shape_text(font_id, text, 30.0).expect("shape text");
        for advance in &mut shaped.advances {
            *advance += spacing;
        }
        shaped.width += spacing * shaped.advances.len() as f64;
        TextSegment {
            text: text.to_owned(),
            direction: TextDirection::Auto,
            source: None,
            font_id,
            font_size: 30.0,
            glyph_ids: shaped.glyph_ids,
            advances: shaped.advances,
            width: shaped.width,
            ascent: metrics.ascent,
            descent: metrics.descent,
            line_gap: metrics.line_gap,
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
        }
    }

    #[test]
    fn automatic_hyphenation_selects_the_farthest_fitting_break_and_has_no_source() {
        let mut fm = deterministic_font_manager();
        let node = crate::SourceNodeId::new(9).unwrap();
        let mut segment = shaped_text_segment(&mut fm, "representation", 0.0);
        segment.source = Some(SourceSpan {
            node,
            char_start: 20,
            char_end: 34,
        });
        let width = fm
            .shape_text(segment.font_id, "represen-", segment.font_size)
            .unwrap()
            .width
            + 0.01;
        let lines = break_into_lines(
            &[InlineItem::HyphenatedText {
                segment,
                language: "en-US".to_owned(),
            }],
            &LineBreakParams {
                available_width: width,
                ..Default::default()
            },
            &fm,
        )
        .unwrap();
        let first = lines[0]
            .items
            .iter()
            .filter_map(|item| match item {
                LineItem::Text(text) => Some(text),
                _ => None,
            })
            .collect::<Vec<_>>();

        assert_eq!(
            first
                .iter()
                .map(|text| text.text.as_str())
                .collect::<String>(),
            "represen-"
        );
        assert_eq!(first.last().unwrap().source, None);
        assert_eq!(first[0].source.unwrap().char_start, 20);
        assert_eq!(first[0].source.unwrap().char_end, 28);
    }

    #[test]
    fn liang_candidates_map_supported_regional_languages_only() {
        assert_eq!(
            hyphenation_opportunities("en-US", "representation"),
            vec![3, 5, 8, 10]
        );
        assert_eq!(
            hyphenation_opportunities("fr-CA", "représentation"),
            vec![2, 6, 9, 11]
        );
        assert!(!hyphenation_opportunities("de-AT", "Silbentrennung").is_empty());
        assert!(!hyphenation_opportunities("es-MX", "representación").is_empty());
        assert!(hyphenation_opportunities("it-IT", "rappresentazione").is_empty());
    }

    #[test]
    fn unwrapped_hyphenated_text_never_emits_a_conditional_hyphen() {
        let mut fm = deterministic_font_manager();
        let segment = shaped_text_segment(&mut fm, "representation", 0.0);
        let lines = break_into_lines(
            &[InlineItem::HyphenatedText {
                segment,
                language: "en-US".to_owned(),
            }],
            &LineBreakParams {
                available_width: 20.0,
                wrap: false,
                ..Default::default()
            },
            &fm,
        )
        .unwrap();
        assert_eq!(lines.len(), 1);
        let text = lines[0]
            .items
            .iter()
            .filter_map(|item| match item {
                LineItem::Text(text) => Some(text.text.as_str()),
                _ => None,
            })
            .collect::<String>();

        assert_eq!(text, "representation");
    }

    #[test]
    fn mixed_direction_line_uses_uax9_visual_order_without_changing_logical_text() {
        let mut fm = deterministic_font_manager();
        let mut segment = shaped_text_segment(&mut fm, "abc אבג 123", 0.0);
        segment.source = Some(SourceSpan {
            node: crate::SourceNodeId::new(7).unwrap(),
            char_start: 50,
            char_end: 61,
        });
        let rich = fm
            .shape_multilingual_text(segment, Some("he-IL"), TextDirection::Auto, false)
            .unwrap();
        let logical_text = rich.iter().map(|span| span.text()).collect::<String>();
        let sources = rich
            .iter()
            .map(|span| span.base().source.expect("logical span keeps source"))
            .collect::<Vec<_>>();
        assert_eq!(sources.first().unwrap().char_start, 50);
        assert_eq!(sources.last().unwrap().char_end, 61);
        assert!(
            sources
                .windows(2)
                .all(|pair| pair[0].char_end == pair[1].char_start)
        );
        let items = rich
            .into_iter()
            .map(InlineItem::MultilingualText)
            .collect::<Vec<_>>();
        let lines = break_multilingual_into_lines(
            &items,
            &LineBreakParams {
                available_width: 1_000.0,
                ..Default::default()
            },
            &fm,
            TextDirection::Auto,
        )
        .unwrap();
        let visual_text = lines[0]
            .items
            .iter()
            .filter_map(|item| match item {
                LineItem::MultilingualText(span) => Some(span.text()),
                _ => None,
            })
            .collect::<String>();
        assert_eq!(logical_text, "abc אבג 123");
        assert_eq!(visual_text, "abc 123 אבג");
        assert_eq!(
            lines[0]
                .items
                .iter()
                .filter_map(|item| match item {
                    LineItem::MultilingualText(span) => {
                        Some((span.text(), span.logical_index(), span.bidi_level()))
                    }
                    _ => None,
                })
                .collect::<Vec<_>>(),
            vec![
                ("abc", 0, 0),
                (" ", 1, 0),
                ("123", 4, 2),
                (" ", 3, 1),
                ("אבג", 2, 1)
            ]
        );
    }

    #[test]
    fn explicit_run_directions_reorder_with_the_rtl_paragraph_once() {
        let mut fm = deterministic_font_manager();
        let mut arabic = shaped_text_segment(&mut fm, "العربية ", 0.0);
        arabic.direction = TextDirection::RightToLeft;
        let mut latin = shaped_text_segment(&mut fm, "ABC 123", 0.0);
        latin.direction = TextDirection::LeftToRight;
        let rich = fm
            .shape_multilingual_paragraph(
                vec![
                    (arabic, Some("ar-SA".to_owned())),
                    (latin, Some("en-US".to_owned())),
                ],
                TextDirection::RightToLeft,
                false,
            )
            .unwrap();
        let lines = break_multilingual_into_lines(
            &rich
                .into_iter()
                .map(InlineItem::MultilingualText)
                .collect::<Vec<_>>(),
            &LineBreakParams {
                available_width: 1_000.0,
                ..Default::default()
            },
            &fm,
            TextDirection::RightToLeft,
        )
        .unwrap();

        let visual = lines[0]
            .items
            .iter()
            .filter_map(|item| match item {
                LineItem::MultilingualText(span) => {
                    Some((span.text(), span.bidi_level(), span.base().direction))
                }
                _ => None,
            })
            .collect::<Vec<_>>();
        assert_eq!(
            visual.iter().map(|span| span.0).collect::<String>(),
            "ABC 123 العربية",
            "{visual:?}"
        );
    }

    #[test]
    fn explicit_rtl_run_retains_numeric_levels_and_line_local_whitespace_reset() {
        let mut fm = deterministic_font_manager();
        let node = crate::SourceNodeId::new(13).unwrap();
        let mut segment = shaped_text_segment(&mut fm, "אבג 123   ", 0.0);
        segment.direction = TextDirection::RightToLeft;
        segment.source = Some(SourceSpan {
            node,
            char_start: 40,
            char_end: 50,
        });
        let rich = fm
            .shape_multilingual_paragraph(
                vec![(segment, Some("he-IL".to_owned()))],
                TextDirection::LeftToRight,
                false,
            )
            .unwrap();
        assert!(
            rich.iter()
                .any(|span| span.text() == "123" && span.bidi_level() == 2),
            "numeric span must retain its higher even level: {:?}",
            rich.iter()
                .map(|span| (span.text(), span.bidi_level()))
                .collect::<Vec<_>>()
        );

        let lines = break_multilingual_into_lines(
            &rich
                .into_iter()
                .map(InlineItem::MultilingualText)
                .collect::<Vec<_>>(),
            &LineBreakParams {
                available_width: 1_000.0,
                ..Default::default()
            },
            &fm,
            TextDirection::LeftToRight,
        )
        .unwrap();
        let spans = lines[0]
            .items
            .iter()
            .filter_map(|item| match item {
                LineItem::MultilingualText(span) => Some(span),
                _ => None,
            })
            .collect::<Vec<_>>();
        assert!(spans.iter().any(|span| {
            span.text().chars().all(char::is_whitespace) && span.bidi_level() == 0
        }));
        assert!(
            spans
                .iter()
                .all(|span| span.base().source.unwrap().node == node)
        );
    }

    #[test]
    fn inline_object_participates_in_rtl_visual_order_without_changing_text_source() {
        let mut fm = deterministic_font_manager();
        let node = crate::SourceNodeId::new(14).unwrap();
        let mut segment = shaped_text_segment(&mut fm, "אבג", 0.0);
        segment.source = Some(SourceSpan {
            node,
            char_start: 8,
            char_end: 11,
        });
        let mut rich = fm
            .shape_multilingual_text(segment, Some("he-IL"), TextDirection::RightToLeft, false)
            .unwrap();
        assert_eq!(rich.len(), 1);
        let items = vec![
            InlineItem::MultilingualText(rich.remove(0)),
            InlineItem::Image {
                width: 12.0,
                height: 12.0,
                media_id: MediaId(77),
            },
        ];

        let lines = break_multilingual_into_lines(
            &items,
            &LineBreakParams {
                available_width: 1_000.0,
                ..Default::default()
            },
            &fm,
            TextDirection::RightToLeft,
        )
        .unwrap();

        assert!(matches!(
            lines[0].items[0],
            LineItem::Image {
                media_id: MediaId(77),
                ..
            }
        ));
        let LineItem::MultilingualText(text) = &lines[0].items[1] else {
            panic!("RTL text must paint after the inline object")
        };
        assert_eq!(text.text(), "אבג");
        assert_eq!(
            text.base().source.unwrap(),
            SourceSpan {
                node,
                char_start: 8,
                char_end: 11,
            }
        );
    }

    #[test]
    fn explicit_rtl_hyphenatable_latin_spans_keep_their_natural_even_levels() {
        let mut fm = deterministic_font_manager();
        let mut letters = shaped_text_segment(&mut fm, "ABC ", 0.0);
        letters.direction = TextDirection::RightToLeft;
        let mut digits = shaped_text_segment(&mut fm, "123", 0.0);
        digits.direction = TextDirection::RightToLeft;
        let lines = break_multilingual_into_lines(
            &[
                InlineItem::HyphenatedText {
                    segment: letters,
                    language: "en-US".to_owned(),
                },
                InlineItem::HyphenatedText {
                    segment: digits,
                    language: "en-US".to_owned(),
                },
            ],
            &LineBreakParams {
                available_width: 1_000.0,
                ..Default::default()
            },
            &fm,
            TextDirection::LeftToRight,
        )
        .unwrap();
        let visual = lines[0]
            .items
            .iter()
            .filter_map(|item| match item {
                LineItem::Text(segment) => Some(segment.text.as_str()),
                _ => None,
            })
            .collect::<String>();
        assert_eq!(visual, "ABC 123");
    }

    #[test]
    fn one_styled_run_applies_line_local_l1_before_l2_without_losing_source() {
        let mut fm = deterministic_font_manager();
        let node = crate::SourceNodeId::new(9).unwrap();
        let mut segment = shaped_text_segment(&mut fm, "אבג   אבג", 0.0);
        segment.source = Some(SourceSpan {
            node,
            char_start: 20,
            char_end: 29,
        });
        let rich = fm
            .shape_multilingual_paragraph(vec![(segment, None)], TextDirection::LeftToRight, false)
            .unwrap();
        assert_eq!(
            rich.iter().map(|span| span.text()).collect::<Vec<_>>(),
            ["אבג", "   ", "אבג"]
        );
        let first_line_width = rich[0].width() + rich[1].width() + 0.01;
        let items = rich
            .into_iter()
            .map(InlineItem::MultilingualText)
            .collect::<Vec<_>>();

        let lines = break_multilingual_into_lines(
            &items,
            &LineBreakParams {
                available_width: first_line_width,
                ..Default::default()
            },
            &fm,
            TextDirection::LeftToRight,
        )
        .unwrap();
        let spans = lines
            .iter()
            .flat_map(|line| &line.items)
            .filter_map(|item| match item {
                LineItem::MultilingualText(span) => Some((
                    span.text(),
                    span.logical_index(),
                    span.bidi_level(),
                    span.direction(),
                    span.base().source.unwrap(),
                )),
                _ => None,
            })
            .collect::<Vec<_>>();
        assert_eq!(
            spans
                .iter()
                .map(|(text, index, level, direction, source)| (
                    *text,
                    *index,
                    *level,
                    *direction,
                    source.char_start..source.char_end,
                ))
                .collect::<Vec<_>>(),
            [
                ("אבג", 0, 1, TextDirection::RightToLeft, 20..23),
                ("   ", 1, 0, TextDirection::LeftToRight, 23..26),
                ("אבג", 2, 1, TextDirection::RightToLeft, 26..29),
            ]
        );
        assert!(spans.iter().all(|(_, _, _, _, source)| source.node == node));
        let first_line = lines[0]
            .items
            .iter()
            .filter_map(|item| match item {
                LineItem::MultilingualText(span) => Some(span.text()),
                _ => None,
            })
            .collect::<String>();

        assert_eq!(first_line, "אבג   ");
    }

    #[test]
    fn cjk_prohibited_punctuation_never_starts_or_ends_a_line() {
        let mut fm = deterministic_font_manager();
        let mut segment = shaped_text_segment(&mut fm, "〈中〉、你好世界", 0.0);
        segment.source = Some(SourceSpan {
            node: crate::SourceNodeId::new(8).unwrap(),
            char_start: 5,
            char_end: 13,
        });
        let rich = fm
            .shape_multilingual_text(segment, Some("zh-CN"), TextDirection::LeftToRight, false)
            .unwrap();
        let items = rich
            .into_iter()
            .map(InlineItem::MultilingualText)
            .collect::<Vec<_>>();
        let lines = break_multilingual_into_lines(
            &items,
            &LineBreakParams {
                available_width: 35.0,
                ..Default::default()
            },
            &fm,
            TextDirection::LeftToRight,
        )
        .unwrap();
        let line_text = lines
            .iter()
            .map(|line| {
                line.items
                    .iter()
                    .filter_map(|item| match item {
                        LineItem::MultilingualText(span) => Some(span.text()),
                        _ => None,
                    })
                    .collect::<String>()
            })
            .collect::<Vec<_>>();
        assert_eq!(line_text, ["〈中〉、", "你", "好", "世", "界"]);
        assert!(
            line_text
                .iter()
                .all(|text| { !text.starts_with(['〉', '、']) && !text.ends_with('〈') })
        );
    }

    #[test]
    fn legacy_stable_source_fixture_compiles_unchanged_public_shapes() {
        let segment = make_text_segment("legacy", 42.0);
        let _inline = InlineItem::Text(segment.clone());
        let _line = LineItem::Text(segment.clone());
        let _params = LineBreakParams {
            available_width: 468.0,
            ind_left: 0.0,
            ind_right: 0.0,
            ind_first_line: 0.0,
            ind_hanging: 0.0,
            tab_stops: Vec::new(),
            line_spacing: LineSpacing::Single,
            jc: None,
            wrap: true,
            line_prefix_widths: Vec::new(),
            line_suffix_widths: Vec::new(),
        };
        let _shaped = crate::ShapedText {
            glyph_ids: segment.glyph_ids.clone(),
            advances: segment.advances.clone(),
            width: segment.width,
        };
        let _run = crate::GlyphRun {
            origin: crate::Point { x: 0.0, y: 0.0 },
            font_id: segment.font_id,
            font_size: segment.font_size,
            glyph_ids: segment.glyph_ids,
            advances: segment.advances,
            text: segment.text,
            source: segment.source,
            color: segment.color,
            bold: segment.bold,
            italic: segment.italic,
            field_kind: segment.field_kind,
            note: segment.note,
        };
    }

    #[test]
    fn empty_paragraph_gets_one_line() {
        let fm = deterministic_font_manager();
        let lines = break_into_lines(&[], &LineBreakParams::default(), &fm).unwrap();
        assert_eq!(lines.len(), 1);
        assert!(lines[0].is_last);
        assert!(lines[0].items.is_empty());
    }

    #[test]
    fn single_word_fits_one_line() {
        let fm = deterministic_font_manager();
        let items = vec![InlineItem::Text(make_text_segment("Hello", 50.0))];
        let lines = break_into_lines(&items, &LineBreakParams::default(), &fm).unwrap();
        assert_eq!(lines.len(), 1);
        assert!(lines[0].is_last);
    }

    #[test]
    fn words_wrap_to_multiple_lines() {
        let fm = deterministic_font_manager();
        // Each word is 200pt wide, line is 468pt → should wrap
        let mut items = vec![
            InlineItem::Text(make_text_segment("Word1", 200.0)),
            InlineItem::Text(make_text_segment("Word2", 200.0)),
        ];
        items.push(InlineItem::Text(make_text_segment("Word3", 200.0)));

        let lines = break_into_lines(&items, &LineBreakParams::default(), &fm).unwrap();
        assert!(lines.len() >= 2);
    }

    #[test]
    fn ligature_runs_reshape_each_break_chunk_without_duplicate_glyphs() {
        let mut fm = deterministic_font_manager();
        let text = "by providing opportunities to crawl in cluttered spaces and handle 3-dimensional objects";
        let spacing = 0.4;
        let segment = shaped_text_segment(&mut fm, text, spacing);
        assert_ne!(segment.glyph_ids.len(), text.chars().count());

        let lines = break_into_lines(
            &[InlineItem::Text(segment)],
            &LineBreakParams {
                available_width: 260.0,
                ..LineBreakParams::default()
            },
            &fm,
        )
        .expect("wrap ligature-bearing text");
        assert!(lines.len() > 1);

        let mut rendered_text = String::new();
        for text_segment in lines.iter().flat_map(|line| {
            line.items.iter().filter_map(|item| match item {
                LineItem::Text(segment) => Some(segment),
                _ => None,
            })
        }) {
            rendered_text.push_str(&text_segment.text);
            let exact = fm
                .shape_text(
                    text_segment.font_id,
                    &text_segment.text,
                    text_segment.font_size,
                )
                .expect("reshape emitted chunk");
            assert_eq!(text_segment.glyph_ids, exact.glyph_ids);
            assert_eq!(text_segment.advances.len(), exact.advances.len());
            for (actual, unspaced) in text_segment.advances.iter().zip(exact.advances) {
                assert!((actual - (unspaced + spacing)).abs() < 1.0e-10);
            }
        }
        assert_eq!(rendered_text, text);
    }

    #[test]
    fn line_splitting_preserves_contiguous_unicode_source_ranges() {
        let mut fm = deterministic_font_manager();
        let node = crate::SourceNodeId::new(7).expect("a non-zero source id");
        let mut segment = shaped_text_segment(&mut fm, "ab 🚀界 cd", 0.0);
        segment.source = Some(crate::SourceSpan {
            node,
            char_start: 11,
            char_end: 19,
        });

        let lines = break_into_lines(
            &[InlineItem::Text(segment)],
            &LineBreakParams {
                available_width: 55.0,
                ..LineBreakParams::default()
            },
            &fm,
        )
        .expect("split mixed Unicode text");
        let sourced = lines
            .iter()
            .flat_map(|line| &line.items)
            .filter_map(|item| match item {
                LineItem::Text(segment) => segment.source,
                _ => None,
            })
            .collect::<Vec<_>>();

        assert!(sourced.len() > 1, "the fixture must cross a line boundary");
        assert_eq!(sourced.first().expect("first range").char_start, 11);
        assert_eq!(sourced.last().expect("last range").char_end, 19);
        for pair in sourced.windows(2) {
            assert_eq!(pair[0].node, node);
            assert_eq!(pair[0].char_end, pair[1].char_start);
        }
    }

    #[test]
    fn forced_line_break() {
        let fm = deterministic_font_manager();
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
            line_spacing: LineSpacing::Exact(24.0),
            ..Default::default()
        };
        let h = compute_line_height(10.0, 3.0, 5.0, 12.0, &params);
        assert!((h - 24.0).abs() < 0.01);
    }

    #[test]
    fn line_height_auto() {
        let params = LineBreakParams {
            line_spacing: LineSpacing::Multiple(2.0),
            ..Default::default()
        };
        let h = compute_line_height(10.0, 3.0, 2.0, 15.0, &params);
        assert!((h - 30.0).abs() < 0.01); // 15 * 2.0
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
        let stops = vec![TabStop {
            pos_pt: 72.0,
            align: TabAlign::Left,
            leader: None,
        }];
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
        let stops = vec![TabStop {
            pos_pt: 400.0,
            align: TabAlign::Right,
            leader: Some(TabLeader::Dot),
        }];
        let (w, leader) = resolve_tab_width(100.0, &stops);
        assert!((w - 300.0).abs() < 0.01);
        assert_eq!(leader, Some('.'));
    }

    #[test]
    fn the_eleven_line_tests_pass_with_owned_types() {
        assert!(LineBreakParams::default().wrap);
        empty_paragraph_gets_one_line();
        single_word_fits_one_line();
        words_wrap_to_multiple_lines();
        forced_line_break();
        line_height_exact();
        line_height_auto();
        first_line_indent();
        hanging_indent();
        tab_stop_resolution();
        default_tab_stops();
        tab_stop_with_dot_leader();
    }

    #[test]
    fn line_spacing_variants_preserve_existing_height_rules() {
        let height = |line_spacing| {
            compute_line_height(
                10.0,
                3.0,
                2.0,
                11.0,
                &LineBreakParams {
                    line_spacing,
                    ..Default::default()
                },
            )
        };

        assert!((height(LineSpacing::Single) - 15.0).abs() < 0.01);
        assert!((height(LineSpacing::Multiple(1.5)) - 16.5).abs() < 0.01);
        assert!((height(LineSpacing::Exact(8.25)) - 8.25).abs() < 0.01);
        assert!((height(LineSpacing::AtLeast(8.25)) - 15.0).abs() < 0.01);
        assert!((height(LineSpacing::AtLeast(18.5)) - 18.5).abs() < 0.01);
    }

    #[test]
    fn mixed_font_line_uses_tallest_full_natural_advance() {
        let fm = deterministic_font_manager();
        let mut first = make_text_segment("first", 20.0);
        first.ascent = 10.0;
        first.descent = 2.0;
        first.line_gap = 4.0;
        let mut second = make_text_segment("second", 20.0);
        second.ascent = 8.0;
        second.descent = 5.0;
        second.line_gap = 1.0;

        let lines = break_into_lines(
            &[InlineItem::Text(first), InlineItem::Text(second)],
            &LineBreakParams::default(),
            &fm,
        )
        .expect("lay out mixed-font line");

        assert_eq!(lines.len(), 1);
        assert!((lines[0].ascent - 10.0).abs() < 0.01);
        assert!((lines[0].descent - 5.0).abs() < 0.01);
        assert!((lines[0].line_gap - 1.0).abs() < 0.01);
        assert!((lines[0].height - 16.0).abs() < 0.01);
        assert!((lines[0].baseline_offset() - 10.5).abs() < 0.01);
    }

    #[test]
    fn multiple_spacing_uses_largest_text_point_size_on_each_line() {
        let fm = deterministic_font_manager();
        let mut first = make_text_segment("first", 20.0);
        first.font_size = 12.0;
        first.line_gap = 4.0;
        let mut second = make_text_segment("second", 20.0);
        second.font_size = 20.0;
        second.line_gap = 1.0;

        let lines = break_into_lines(
            &[InlineItem::Text(first), InlineItem::Text(second)],
            &LineBreakParams {
                line_spacing: LineSpacing::Multiple(1.25),
                ..LineBreakParams::default()
            },
            &fm,
        )
        .expect("lay out percentage-spaced mixed-size line");

        assert_eq!(lines.len(), 1);
        assert!((lines[0].height - 25.0).abs() < 0.01);
    }

    #[test]
    fn positive_leading_is_split_and_below_natural_exact_spacing_is_not_clamped() {
        let positive = LayoutLine {
            items: Vec::new(),
            width: 0.0,
            ascent: 10.0,
            descent: 3.0,
            line_gap: 5.0,
            height: 18.0,
            indent_left: 0.0,
            available_width: 100.0,
            is_last: true,
        };
        let below_natural = LayoutLine {
            height: 8.0,
            ..positive.clone()
        };

        assert!((positive.baseline_offset() - 12.5).abs() < 0.01);
        assert!((below_natural.height - 8.0).abs() < 0.01);
        assert!((below_natural.baseline_offset() - 10.0).abs() < 0.01);
    }

    #[test]
    fn zero_gap_and_empty_segment_preserve_natural_height_rules() {
        let fm = deterministic_font_manager();
        let zero_gap = make_text_segment("zero", 20.0);
        let mut empty = make_text_segment("", 0.0);
        empty.line_gap = 4.0;

        let zero_gap_line = break_into_lines(
            &[InlineItem::Text(zero_gap)],
            &LineBreakParams::default(),
            &fm,
        )
        .expect("lay out zero-gap line");
        let empty_line =
            break_into_lines(&[InlineItem::Text(empty)], &LineBreakParams::default(), &fm)
                .expect("lay out styled empty line");

        assert!((zero_gap_line[0].height - 13.0).abs() < 0.01);
        assert!((empty_line[0].height - 17.0).abs() < 0.01);
    }

    #[test]
    fn wrap_false_only_breaks_on_an_explicit_break() {
        let fm = deterministic_font_manager();
        let params = LineBreakParams {
            available_width: 100.0,
            wrap: false,
            ..Default::default()
        };

        for forced_break in [
            InlineItem::LineBreak,
            InlineItem::PageBreak,
            InlineItem::ColumnBreak,
        ] {
            let items = vec![
                InlineItem::Text(make_text_segment("one", 80.0)),
                InlineItem::Text(make_text_segment("two", 80.0)),
                forced_break,
                InlineItem::Text(make_text_segment("three", 80.0)),
                InlineItem::Text(make_text_segment("four", 80.0)),
            ];
            let lines = break_into_lines(&items, &params, &fm).unwrap();

            assert_eq!(lines.len(), 2);
            assert!((lines[0].width - 160.0).abs() < 0.01);
            assert!((lines[1].width - 160.0).abs() < 0.01);
        }
    }

    #[test]
    fn tab_stops_use_point_positions_and_owned_leaders() {
        let mut fm = deterministic_font_manager();
        let font_id = fm
            .resolve_font(Some("Carlito"), false, false)
            .expect("bundled Carlito should resolve");
        let stop = TabStop {
            pos_pt: 72.25,
            align: TabAlign::Decimal,
            leader: Some(TabLeader::Dot),
        };

        let item = inline_to_line_item(&InlineItem::Tab, 12.0, &[stop], &fm, Some((font_id, 12.0)));

        let LineItem::Tab {
            width,
            leader: Some(leader),
        } = item
        else {
            panic!("owned dot leader should shape into a tab line item");
        };
        assert!((width - 60.25).abs() < 0.01);
        assert!((leader.width - 60.25).abs() < 0.01);
        assert!(!leader.glyph_ids.is_empty());
        assert!(leader.text.chars().all(|ch| ch == '.'));
    }

    #[test]
    fn staged_image_types_use_media_id_instead_of_embed_id() {
        let media_id = crate::MediaId::from_bytes(b"image");
        let item = inline_to_line_item(
            &InlineItem::Image {
                width: 10.0,
                height: 20.0,
                media_id,
            },
            0.0,
            &[],
            &deterministic_font_manager(),
            None,
        );
        let LineItem::Image {
            media_id: actual, ..
        } = item
        else {
            panic!("image should remain an image");
        };
        assert_eq!(actual, media_id);
    }

    #[test]
    fn group_inline_item_breaks_and_positions_like_an_image() {
        use crate::{GroupElement, PositionedElement, Transform};

        let group = GroupElement {
            transform: Transform::IDENTITY,
            clip: None,
            opacity: 1.0,
            effects: Vec::new(),
            children: vec![PositionedElement::FilledRect {
                rect: crate::Rect {
                    x: 2.0,
                    y: 3.0,
                    width: 4.0,
                    height: 5.0,
                },
                color: crate::Color::BLACK,
            }],
        };
        let items = vec![InlineItem::Group {
            width: 80.0,
            height: 40.0,
            baseline: None,
            group: group.clone(),
        }];
        let lines = break_into_lines(
            &items,
            &LineBreakParams {
                available_width: 80.0,
                ..Default::default()
            },
            &deterministic_font_manager(),
        )
        .expect("group line breaking");

        assert_eq!(lines.len(), 1);
        assert_eq!(lines[0].width, 80.0);
        assert_eq!(lines[0].height, 40.0);
        let LineItem::Group {
            width,
            height,
            group: actual,
            ..
        } = &lines[0].items[0]
        else {
            panic!("inline group should remain a group line item");
        };
        assert_eq!((*width, *height), (80.0, 40.0));
        assert_eq!(actual, &group);
    }

    #[test]
    fn baseline_aware_inline_groups_contribute_exact_ascent_and_descent() {
        use crate::{GroupElement, Transform};

        let group = GroupElement {
            transform: Transform::IDENTITY,
            clip: None,
            opacity: 1.0,
            effects: Vec::new(),
            children: Vec::new(),
        };
        let lines = break_into_lines(
            &[
                InlineItem::Group {
                    width: 20.0,
                    height: 18.0,
                    baseline: None,
                    group: group.clone(),
                },
                InlineItem::Group {
                    width: 20.0,
                    height: 18.0,
                    baseline: Some(11.0),
                    group,
                },
            ],
            &LineBreakParams {
                available_width: 100.0,
                ..Default::default()
            },
            &deterministic_font_manager(),
        )
        .expect("group line breaking");

        assert_eq!(lines.len(), 1);
        assert_eq!(lines[0].ascent, 18.0);
        assert_eq!(lines[0].descent, 7.0);
        let LineItem::Group {
            baseline: Some(baseline),
            ..
        } = &lines[0].items[1]
        else {
            panic!("baseline should survive line breaking");
        };
        assert_eq!(*baseline, 11.0);
    }

    #[test]
    fn group_baselines_are_normalized_before_pagination() {
        use crate::{GroupElement, Transform};

        for (height, baseline, expected) in [
            (18.0, Some(-4.0), Some(0.0)),
            (18.0, Some(30.0), Some(18.0)),
            (18.0, Some(f64::NAN), None),
            (-18.0, Some(4.0), Some(0.0)),
            (f64::NAN, Some(4.0), None),
        ] {
            let item = InlineItem::Group {
                width: 20.0,
                height,
                baseline,
                group: GroupElement {
                    transform: Transform::IDENTITY,
                    clip: None,
                    opacity: 1.0,
                    effects: Vec::new(),
                    children: Vec::new(),
                },
            };
            let LineItem::Group { baseline, .. } =
                inline_to_line_item(&item, 0.0, &[], &deterministic_font_manager(), None)
            else {
                panic!("inline group should remain a group line item");
            };
            assert_eq!(baseline, expected);
        }
    }
}
