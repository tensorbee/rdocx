//! Block-level layout: paragraphs and tables as positioned blocks.

use rdocx_oxml::borders::CT_PBdr;
use rdocx_oxml::drawing::{
    AnchorAlignH, AnchorAlignV, ST_RelativeFromH, ST_RelativeFromV, WrapType,
};
use rdocx_oxml::shared::ST_Jc;

use crate::line::{InlineItem, LayoutLine, LineBreakParams, TextSegment};
use crate::output::Color;
use crate::table::TableBlock;

/// A laid-out block element (paragraph or table).
#[derive(Debug, Clone)]
pub enum LayoutBlock {
    Paragraph(ParagraphBlock),
    Table(TableBlock),
}

impl LayoutBlock {
    /// Total height including spacing.
    pub fn total_height(&self) -> f64 {
        match self {
            LayoutBlock::Paragraph(p) => p.total_height(),
            LayoutBlock::Table(t) => t.total_height(),
        }
    }

    /// Content height without spacing.
    pub fn content_height(&self) -> f64 {
        match self {
            LayoutBlock::Paragraph(p) => p.content_height(),
            LayoutBlock::Table(t) => t.content_height(),
        }
    }

    pub fn space_before(&self) -> f64 {
        match self {
            LayoutBlock::Paragraph(p) => p.space_before,
            LayoutBlock::Table(_) => 0.0,
        }
    }

    pub fn space_after(&self) -> f64 {
        match self {
            LayoutBlock::Paragraph(p) => p.space_after,
            LayoutBlock::Table(_) => 0.0,
        }
    }

    pub fn keep_next(&self) -> bool {
        match self {
            LayoutBlock::Paragraph(p) => p.keep_next,
            LayoutBlock::Table(_) => false,
        }
    }

    pub fn keep_lines(&self) -> bool {
        match self {
            LayoutBlock::Paragraph(p) => p.keep_lines,
            LayoutBlock::Table(_) => false,
        }
    }

    pub fn page_break_before(&self) -> bool {
        match self {
            LayoutBlock::Paragraph(p) => p.page_break_before,
            LayoutBlock::Table(_) => false,
        }
    }

    pub fn widow_control(&self) -> bool {
        match self {
            LayoutBlock::Paragraph(p) => p.widow_control,
            LayoutBlock::Table(_) => false,
        }
    }
}

/// A dropped capital rendered alongside the opening lines of a paragraph.
#[derive(Debug, Clone)]
pub struct DropCap {
    /// The shaped initial glyph run.
    pub segment: TextSegment,
    /// Number of text lines wrapped around the drop cap.
    pub line_count: usize,
    /// Extra gap between the drop cap and wrapped text.
    pub padding_right: f64,
}

/// A floating image anchored to a paragraph.
#[derive(Debug, Clone)]
pub struct AnchoredImage {
    pub behind_doc: bool,
    pub width: f64,
    pub height: f64,
    /// Horizontal extent from the image's left edge to the right-most visible pixel.
    pub wrap_left_extent: f64,
    /// Horizontal extent from the image's right edge to the left-most visible pixel.
    pub wrap_right_extent: f64,
    pub embed_id: String,
    pub wrap: WrapType,
    pub dist_top: f64,
    pub dist_bottom: f64,
    pub dist_left: f64,
    pub dist_right: f64,
    pub pos_h_offset: f64,
    pub pos_h_relative_from: ST_RelativeFromH,
    pub pos_h_align: Option<AnchorAlignH>,
    pub pos_v_offset: f64,
    pub pos_v_relative_from: ST_RelativeFromV,
    pub pos_v_align: Option<AnchorAlignV>,
}

/// A laid-out paragraph with its lines and spacing.
#[derive(Debug, Clone)]
pub struct ParagraphBlock {
    /// Laid-out lines.
    pub lines: Vec<LayoutLine>,
    /// Space before the paragraph in points.
    pub space_before: f64,
    /// Space after the paragraph in points.
    pub space_after: f64,
    /// Paragraph borders.
    pub borders: Option<CT_PBdr>,
    /// Background shading color.
    pub shading: Option<Color>,
    /// Left indent in points.
    pub indent_left: f64,
    /// Right indent in points.
    pub indent_right: f64,
    /// Paragraph justification.
    pub jc: Option<ST_Jc>,
    /// Keep with next paragraph.
    pub keep_next: bool,
    /// Keep all lines together on one page.
    pub keep_lines: bool,
    /// Force page break before this paragraph.
    pub page_break_before: bool,
    /// Widow/orphan control.
    pub widow_control: bool,
    /// Heading level (1-9) if this is a heading paragraph, for outline generation.
    pub heading_level: Option<u32>,
    /// Heading text for outline generation.
    pub heading_text: Option<String>,
    /// Optional dropped capital rendered alongside the first lines.
    pub drop_cap: Option<DropCap>,
    /// Vertical offset applied before the first text line.
    pub content_offset_top: f64,
    /// Footnotes referenced by this paragraph and their reserved heights.
    pub footnote_reserves: Vec<(i32, f64)>,
    /// Floating images attached to this paragraph.
    pub anchored_images: Vec<AnchoredImage>,
    /// Source inline items used to reflow anchored-image paragraphs at pagination time.
    pub inline_items: Vec<InlineItem>,
    /// Base line-break parameters before page-position-dependent anchor wrapping is applied.
    pub line_break_params: LineBreakParams,
}

impl ParagraphBlock {
    /// Total height of the paragraph lines (not including before/after spacing).
    pub fn content_height(&self) -> f64 {
        self.content_offset_top + self.lines.iter().map(|l| l.height).sum::<f64>()
    }

    /// Total height including spacing.
    pub fn total_height(&self) -> f64 {
        self.space_before + self.content_height() + self.space_after
    }

    /// Number of lines.
    pub fn line_count(&self) -> usize {
        self.lines.len()
    }
}

/// Build a ParagraphBlock from resolved properties and layout lines.
pub fn build_paragraph_block(
    lines: Vec<LayoutLine>,
    space_before: f64,
    space_after: f64,
    borders: Option<CT_PBdr>,
    shading: Option<Color>,
    indent_left: f64,
    indent_right: f64,
    jc: Option<ST_Jc>,
    keep_next: bool,
    keep_lines: bool,
    page_break_before: bool,
    widow_control: bool,
    drop_cap: Option<DropCap>,
    footnote_reserves: Vec<(i32, f64)>,
    anchored_images: Vec<AnchoredImage>,
    inline_items: Vec<InlineItem>,
    line_break_params: LineBreakParams,
) -> ParagraphBlock {
    ParagraphBlock {
        lines,
        space_before,
        space_after,
        borders,
        shading,
        indent_left,
        indent_right,
        jc,
        keep_next,
        keep_lines,
        page_break_before,
        widow_control,
        heading_level: None,
        heading_text: None,
        drop_cap,
        content_offset_top: 0.0,
        footnote_reserves,
        anchored_images,
        inline_items,
        line_break_params,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn paragraph_block_height() {
        let block = ParagraphBlock {
            lines: vec![
                LayoutLine {
                    items: vec![],
                    width: 0.0,
                    ascent: 10.0,
                    descent: 3.0,
                    height: 13.0,
                    indent_left: 0.0,
                    available_width: 468.0,
                    is_last: false,
                },
                LayoutLine {
                    items: vec![],
                    width: 0.0,
                    ascent: 10.0,
                    descent: 3.0,
                    height: 13.0,
                    indent_left: 0.0,
                    available_width: 468.0,
                    is_last: true,
                },
            ],
            space_before: 6.0,
            space_after: 8.0,
            borders: None,
            shading: None,
            indent_left: 0.0,
            indent_right: 0.0,
            jc: None,
            keep_next: false,
            keep_lines: false,
            page_break_before: false,
            widow_control: true,
            heading_level: None,
            heading_text: None,
            drop_cap: None,
            content_offset_top: 0.0,
            footnote_reserves: Vec::new(),
            anchored_images: Vec::new(),
            inline_items: Vec::new(),
            line_break_params: LineBreakParams::default(),
        };
        assert!((block.content_height() - 26.0).abs() < 0.01);
        assert!((block.total_height() - 40.0).abs() < 0.01);
    }
}
