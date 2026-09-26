//! Paragraph — a block-level container for runs of text.

use rdocx_oxml::borders::{CT_BorderEdge, CT_PBdr, CT_TabStop, CT_Tabs};
use rdocx_oxml::document::CT_SectPr;
use rdocx_oxml::math::OfficeMath;
use rdocx_oxml::properties::{CT_FramePr, CT_PPr, CT_RPr, CT_Shd};
use rdocx_oxml::ruby::CT_Ruby;
use rdocx_oxml::shared::{
    ST_Border, ST_Jc, ST_PageOrientation, ST_SectionType, ST_TabJc, ST_TabLeader, ST_Underline,
};
use rdocx_oxml::text::{
    AcceptedRunPath, AcceptedRunPathSegment, BreakType, CT_P, CT_R, CommentRangeMarker,
    HyperlinkSpan, RunContent, hyperlink_revision_index,
};
use rdocx_oxml::units::{HalfPoint, Twips};

use crate::run::{Run, RunRef};
use crate::table::TableConditionalFormatting;
use crate::{ContentControlRef, Length, RevisionRef};

/// Paragraph alignment options.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Alignment {
    Left,
    Center,
    Right,
    Justify,
}

/// One direct child of a paragraph, in source order.
#[non_exhaustive]
pub enum ParagraphItemRef<'a> {
    /// A direct run.
    Run(RunRef<'a>),
    /// A hyperlink and its direct children.
    Hyperlink(HyperlinkRef<'a>),
    /// An inline or display OfficeMath equation.
    Equation(&'a OfficeMath),
    /// A paragraph-level content control.
    ContentControl(ContentControlRef<'a>),
    /// A tracked insertion, deletion, or move.
    Revision(RevisionRef<'a>),
    /// The start of a comment range.
    CommentRangeStart {
        /// The comment ID.
        id: i32,
        /// Whether the marker wraps content instead of being empty.
        has_child_content: bool,
    },
    /// The end of a comment range.
    CommentRangeEnd {
        /// The comment ID.
        id: i32,
        /// Whether the marker wraps content instead of being empty.
        has_child_content: bool,
    },
    /// The start of a bookmark.
    BookmarkStart {
        /// The bookmark ID, when present.
        id: Option<i32>,
        /// The bookmark name, when present.
        name: Option<&'a str>,
        /// Whether the marker wraps content instead of being empty.
        has_child_content: bool,
    },
    /// The end of a bookmark.
    BookmarkEnd {
        /// The bookmark ID, when present.
        id: Option<i32>,
        /// Whether the marker wraps content instead of being empty.
        has_child_content: bool,
    },
    /// A preserved paragraph child that rdocx does not model.
    UnsupportedXml(&'a [u8]),
}

/// One direct child within a hyperlink, in source order.
#[non_exhaustive]
pub enum HyperlinkItemRef<'a> {
    /// A run in the hyperlink.
    Run(RunRef<'a>),
    /// A tracked insertion, deletion, or move in the hyperlink.
    Revision(RevisionRef<'a>),
    /// A preserved hyperlink child that rdocx does not model.
    UnsupportedXml(&'a [u8]),
}

/// An immutable paragraph-level hyperlink.
#[derive(Clone, Copy)]
pub struct HyperlinkRef<'a> {
    paragraph: &'a CT_P,
    index: usize,
}

impl<'a> HyperlinkRef<'a> {
    fn inner(&self) -> &'a HyperlinkSpan {
        &self.paragraph.hyperlinks[self.index]
    }

    /// The relationship ID for this hyperlink, when it has an external target.
    pub fn relationship_id(&self) -> Option<&'a str> {
        self.inner().rel_id.as_deref()
    }

    /// The bookmark anchor for this hyperlink, when it has an internal target.
    pub fn anchor(&self) -> Option<&'a str> {
        self.inner().anchor.as_deref()
    }

    /// The optional hyperlink tooltip.
    pub fn tooltip(&self) -> Option<&'a str> {
        self.inner().tooltip.as_deref()
    }

    /// The optional location in the hyperlink target document.
    pub fn doc_location(&self) -> Option<&'a str> {
        self.inner().doc_location.as_deref()
    }

    /// Whether the hyperlink retains attributes outside the modeled reader API.
    pub fn has_unmodeled_semantic_attributes(&self) -> bool {
        self.inner()
            .extra_attributes
            .iter()
            .any(|(name, _)| name != "xmlns" && !name.starts_with("xmlns:"))
    }

    /// Get the combined text of the hyperlink runs.
    pub fn text(&self) -> String {
        let hyperlink = self.inner();
        self.paragraph.runs[hyperlink.run_start..hyperlink.run_end]
            .iter()
            .map(CT_R::text)
            .collect()
    }

    /// Iterate over hyperlink children in source order.
    pub fn items(&self) -> impl Iterator<Item = HyperlinkItemRef<'a>> {
        let hyperlink = self.inner();
        let mut items = Vec::new();
        for relative_index in 0..=hyperlink.run_end - hyperlink.run_start {
            let run_index = hyperlink.run_start + relative_index;
            let revisions = self
                .paragraph
                .revisions
                .iter()
                .filter(|(at, slot, _)| {
                    *at == run_index && hyperlink_revision_index(*slot) == Some(self.index)
                })
                .map(|(_, _, revision)| revision)
                .collect::<Vec<_>>();
            for revision_index in 0..=revisions.len() {
                items.extend(
                    hyperlink
                        .extra_xml
                        .iter()
                        .filter(|(at, before, _)| {
                            *at == relative_index
                                && (*before).min(revisions.len()) == revision_index
                        })
                        .map(|(_, _, raw)| HyperlinkItemRef::UnsupportedXml(raw.as_slice())),
                );
                if let Some(revision) = revisions.get(revision_index) {
                    items.push(HyperlinkItemRef::Revision(RevisionRef { inner: revision }));
                }
            }
            if relative_index < hyperlink.run_end - hyperlink.run_start {
                items.push(HyperlinkItemRef::Run(RunRef {
                    inner: &self.paragraph.runs[run_index],
                }));
            }
        }
        items.into_iter()
    }
}

impl Alignment {
    fn to_st_jc(self) -> ST_Jc {
        match self {
            Alignment::Left => ST_Jc::Left,
            Alignment::Center => ST_Jc::Center,
            Alignment::Right => ST_Jc::Right,
            Alignment::Justify => ST_Jc::Both,
        }
    }

    fn from_st_jc(jc: ST_Jc) -> Self {
        match jc {
            ST_Jc::Center => Alignment::Center,
            ST_Jc::Right | ST_Jc::End => Alignment::Right,
            ST_Jc::Both | ST_Jc::Distribute => Alignment::Justify,
            _ => Alignment::Left,
        }
    }
}

/// Border style for paragraph borders.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BorderStyle {
    None,
    Single,
    Thick,
    Double,
    Dotted,
    Dashed,
    DotDash,
    Wave,
}

impl BorderStyle {
    /// Convert to the OXML `ST_Border` type. Visible to the table module,
    /// which builds cell borders from the same enum.
    pub(crate) fn to_st(self) -> ST_Border {
        match self {
            Self::None => ST_Border::None,
            Self::Single => ST_Border::Single,
            Self::Thick => ST_Border::Thick,
            Self::Double => ST_Border::Double,
            Self::Dotted => ST_Border::Dotted,
            Self::Dashed => ST_Border::Dashed,
            Self::DotDash => ST_Border::DotDash,
            Self::Wave => ST_Border::Wave,
        }
    }
}

/// An immutable paragraph border edge.
#[derive(Debug, Clone, Copy)]
pub struct ParagraphBorderRef<'a> {
    inner: &'a CT_BorderEdge,
}

impl ParagraphBorderRef<'_> {
    /// The OOXML border style name.
    pub fn style(self) -> &'static str {
        self.inner.val.to_str()
    }

    /// Border width in eighths of a point.
    pub fn size_eighths_pt(self) -> Option<u32> {
        self.inner.sz
    }

    /// Space between the border and content in points.
    pub fn space_points(self) -> Option<u32> {
        self.inner.space
    }

    /// Border color, normally a six-digit RGB hex value or `auto`.
    pub fn color(&self) -> Option<&str> {
        self.inner.color.as_deref()
    }
}

/// Tab stop alignment.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TabAlignment {
    Left,
    Center,
    Right,
    Decimal,
}

impl TabAlignment {
    fn to_st(self) -> ST_TabJc {
        match self {
            Self::Left => ST_TabJc::Left,
            Self::Center => ST_TabJc::Center,
            Self::Right => ST_TabJc::Right,
            Self::Decimal => ST_TabJc::Decimal,
        }
    }

    fn from_st(value: ST_TabJc) -> Option<Self> {
        match value {
            ST_TabJc::Left => Some(Self::Left),
            ST_TabJc::Center => Some(Self::Center),
            ST_TabJc::Right => Some(Self::Right),
            ST_TabJc::Decimal => Some(Self::Decimal),
            ST_TabJc::Bar | ST_TabJc::Clear | ST_TabJc::Num => None,
        }
    }
}

/// Tab leader character.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TabLeader {
    None,
    Dot,
    Hyphen,
    Underscore,
}

impl TabLeader {
    fn to_st(self) -> ST_TabLeader {
        match self {
            Self::None => ST_TabLeader::None,
            Self::Dot => ST_TabLeader::Dot,
            Self::Hyphen => ST_TabLeader::Hyphen,
            Self::Underscore => ST_TabLeader::Underscore,
        }
    }

    fn from_st(value: ST_TabLeader) -> Option<Self> {
        match value {
            ST_TabLeader::None => Some(Self::None),
            ST_TabLeader::Dot => Some(Self::Dot),
            ST_TabLeader::Hyphen => Some(Self::Hyphen),
            ST_TabLeader::Underscore => Some(Self::Underscore),
            ST_TabLeader::Heavy | ST_TabLeader::MiddleDot => None,
        }
    }
}

/// One edge in the paragraph border model.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ParagraphBorderEdge {
    /// The top edge.
    Top,
    /// The bottom edge.
    Bottom,
    /// The left edge.
    Left,
    /// The right edge.
    Right,
    /// The edge drawn between consecutive paragraphs that share a border.
    Between,
    /// The bar edge drawn beside the paragraph.
    Bar,
}

/// Text flow written to `w:textDirection` for a paragraph.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ParagraphTextDirection {
    /// Horizontal text from left to right, with lines from top to bottom.
    LeftToRightTopToBottom,
    /// Vertical text from top to bottom, with lines from right to left.
    TopToBottomRightToLeft,
    /// Vertical text from bottom to top, with lines from left to right.
    BottomToTopLeftToRight,
    /// Vertically oriented left-to-right text with top-to-bottom lines.
    LeftToRightTopToBottomVertical,
    /// Vertically oriented top-to-bottom text with right-to-left lines.
    TopToBottomRightToLeftVertical,
    /// Vertically oriented top-to-bottom text with left-to-right lines.
    TopToBottomLeftToRightVertical,
}

impl ParagraphTextDirection {
    fn from_str(value: &str) -> Option<Self> {
        match value {
            "lrTb" => Some(Self::LeftToRightTopToBottom),
            "tbRl" => Some(Self::TopToBottomRightToLeft),
            "btLr" => Some(Self::BottomToTopLeftToRight),
            "lrTbV" => Some(Self::LeftToRightTopToBottomVertical),
            "tbRlV" => Some(Self::TopToBottomRightToLeftVertical),
            "tbLrV" => Some(Self::TopToBottomLeftToRightVertical),
            _ => None,
        }
    }

    fn to_str(self) -> &'static str {
        match self {
            Self::LeftToRightTopToBottom => "lrTb",
            Self::TopToBottomRightToLeft => "tbRl",
            Self::BottomToTopLeftToRight => "btLr",
            Self::LeftToRightTopToBottomVertical => "lrTbV",
            Self::TopToBottomRightToLeftVertical => "tbRlV",
            Self::TopToBottomLeftToRightVertical => "tbLrV",
        }
    }
}

/// Vertical alignment of characters on a line, written to `w:textAlignment`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ParagraphTextAlignment {
    /// Align to the top of the line.
    Top,
    /// Centre on the line.
    Center,
    /// Align to the baseline.
    Baseline,
    /// Align to the bottom of the line.
    Bottom,
    /// Let the consumer choose.
    Auto,
}

impl ParagraphTextAlignment {
    fn from_str(value: &str) -> Option<Self> {
        match value {
            "top" => Some(Self::Top),
            "center" => Some(Self::Center),
            "baseline" => Some(Self::Baseline),
            "bottom" => Some(Self::Bottom),
            "auto" => Some(Self::Auto),
            _ => None,
        }
    }

    fn to_str(self) -> &'static str {
        match self {
            Self::Top => "top",
            Self::Center => "center",
            Self::Baseline => "baseline",
            Self::Bottom => "bottom",
            Self::Auto => "auto",
        }
    }
}

/// Which lines a text box tightens against, written to `w:textboxTightWrap`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TextboxTightWrap {
    /// Wrap against the full text box width.
    None,
    /// Wrap tightly against every line.
    AllLines,
    /// Wrap tightly against the first and last line only.
    FirstAndLastLine,
    /// Wrap tightly against the first line only.
    FirstLineOnly,
    /// Wrap tightly against the last line only.
    LastLineOnly,
}

impl TextboxTightWrap {
    fn from_str(value: &str) -> Option<Self> {
        match value {
            "none" => Some(Self::None),
            "allLines" => Some(Self::AllLines),
            "firstAndLastLine" => Some(Self::FirstAndLastLine),
            "firstLineOnly" => Some(Self::FirstLineOnly),
            "lastLineOnly" => Some(Self::LastLineOnly),
            _ => None,
        }
    }

    fn to_str(self) -> &'static str {
        match self {
            Self::None => "none",
            Self::AllLines => "allLines",
            Self::FirstAndLastLine => "firstAndLastLine",
            Self::FirstLineOnly => "firstLineOnly",
            Self::LastLineOnly => "lastLineOnly",
        }
    }
}

/// Drop cap placement for a text frame.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DropCap {
    /// The frame is not a drop cap.
    None,
    /// The drop cap sits inside the text margin.
    Drop,
    /// The drop cap sits in the margin.
    Margin,
}

impl DropCap {
    fn from_str(value: &str) -> Option<Self> {
        match value {
            "none" => Some(Self::None),
            "drop" => Some(Self::Drop),
            "margin" => Some(Self::Margin),
            _ => None,
        }
    }

    fn to_str(self) -> &'static str {
        match self {
            Self::None => "none",
            Self::Drop => "drop",
            Self::Margin => "margin",
        }
    }
}

/// How surrounding text wraps around a text frame.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FrameWrap {
    /// Let the consumer choose.
    Auto,
    /// Never place text beside the frame.
    NotBeside,
    /// Wrap around the frame rectangle.
    Around,
    /// Wrap tightly against the frame.
    Tight,
    /// Wrap through the frame.
    Through,
    /// Do not wrap text around the frame.
    None,
}

impl FrameWrap {
    fn from_str(value: &str) -> Option<Self> {
        match value {
            "auto" => Some(Self::Auto),
            "notBeside" => Some(Self::NotBeside),
            "around" => Some(Self::Around),
            "tight" => Some(Self::Tight),
            "through" => Some(Self::Through),
            "none" => Some(Self::None),
            _ => None,
        }
    }

    fn to_str(self) -> &'static str {
        match self {
            Self::Auto => "auto",
            Self::NotBeside => "notBeside",
            Self::Around => "around",
            Self::Tight => "tight",
            Self::Through => "through",
            Self::None => "none",
        }
    }
}

/// What a text frame position is measured from.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FrameAnchor {
    /// Measured from the page margin.
    Margin,
    /// Measured from the page edge.
    Page,
    /// Measured from the surrounding text.
    Text,
}

impl FrameAnchor {
    fn from_str(value: &str) -> Option<Self> {
        match value {
            "margin" => Some(Self::Margin),
            "page" => Some(Self::Page),
            "text" => Some(Self::Text),
            _ => None,
        }
    }

    fn to_str(self) -> &'static str {
        match self {
            Self::Margin => "margin",
            Self::Page => "page",
            Self::Text => "text",
        }
    }
}

/// Text frame placement for a paragraph, the checked mirror of `w:framePr`.
///
/// Relative alignment (`w:xAlign`, `w:yAlign`), the height rule (`w:hRule`) and
/// any producer attribute are not mirrored. Writing a frame through this type
/// leaves whatever the document already stored for them in place.
#[derive(Debug, Clone, Copy, Default, PartialEq)]
pub struct ParagraphFrame {
    /// Frame width.
    pub width: Option<Length>,
    /// Frame height.
    pub height: Option<Length>,
    /// Horizontal distance from the surrounding text.
    pub horizontal_space: Option<Length>,
    /// Vertical distance from the surrounding text.
    pub vertical_space: Option<Length>,
    /// Absolute horizontal position from the horizontal anchor.
    pub horizontal_position: Option<Length>,
    /// Absolute vertical position from the vertical anchor.
    pub vertical_position: Option<Length>,
    /// What the horizontal position is measured from.
    pub horizontal_anchor: Option<FrameAnchor>,
    /// What the vertical position is measured from.
    pub vertical_anchor: Option<FrameAnchor>,
    /// How surrounding text wraps around the frame.
    pub wrap: Option<FrameWrap>,
    /// Drop cap placement.
    pub drop_cap: Option<DropCap>,
    /// Drop cap height in lines.
    pub drop_cap_lines: Option<u32>,
    /// Whether the frame stays with the paragraph it anchors.
    pub anchor_lock: Option<bool>,
}

impl ParagraphFrame {
    fn from_ct(frame: &CT_FramePr) -> Self {
        ParagraphFrame {
            width: frame.w.map(|twips| Length::twips(twips.0)),
            height: frame.h.map(|twips| Length::twips(twips.0)),
            horizontal_space: frame.h_space.map(|twips| Length::twips(twips.0)),
            vertical_space: frame.v_space.map(|twips| Length::twips(twips.0)),
            horizontal_position: frame.x.map(|twips| Length::twips(twips.0)),
            vertical_position: frame.y.map(|twips| Length::twips(twips.0)),
            horizontal_anchor: frame.h_anchor.as_deref().and_then(FrameAnchor::from_str),
            vertical_anchor: frame.v_anchor.as_deref().and_then(FrameAnchor::from_str),
            wrap: frame.wrap.as_deref().and_then(FrameWrap::from_str),
            drop_cap: frame.drop_cap.as_deref().and_then(DropCap::from_str),
            drop_cap_lines: frame.lines,
            anchor_lock: frame.anchor_lock,
        }
    }

    fn apply_to(self, frame: &mut CT_FramePr) {
        frame.w = self.width.map(Length::as_twips);
        frame.h = self.height.map(Length::as_twips);
        frame.h_space = self.horizontal_space.map(Length::as_twips);
        frame.v_space = self.vertical_space.map(Length::as_twips);
        frame.x = self.horizontal_position.map(Length::as_twips);
        frame.y = self.vertical_position.map(Length::as_twips);
        frame.h_anchor = self
            .horizontal_anchor
            .map(|anchor| anchor.to_str().to_owned());
        frame.v_anchor = self
            .vertical_anchor
            .map(|anchor| anchor.to_str().to_owned());
        frame.wrap = self.wrap.map(|wrap| wrap.to_str().to_owned());
        frame.drop_cap = self.drop_cap.map(|drop_cap| drop_cap.to_str().to_owned());
        frame.lines = self.drop_cap_lines;
        frame.anchor_lock = self.anchor_lock;
    }
}

/// An immutable tab stop.
#[derive(Debug, Clone, Copy)]
pub struct TabStopRef<'a> {
    inner: &'a CT_TabStop,
}

impl TabStopRef<'_> {
    /// The tab stop alignment, or `None` for a token this facade does not
    /// model, such as `bar`, `clear` or `num`.
    pub fn alignment(self) -> Option<TabAlignment> {
        TabAlignment::from_st(self.inner.val)
    }

    /// The tab stop position from the leading text margin.
    pub fn position(self) -> Length {
        Length::twips(self.inner.pos.0)
    }

    /// The leader character, or `None` when unset or outside the modeled set.
    pub fn leader(self) -> Option<TabLeader> {
        self.inner.leader.and_then(TabLeader::from_st)
    }
}

/// A mutable handle on the paragraph mark's own run properties.
///
/// The paragraph mark is the pilcrow at the end of a paragraph. Its formatting
/// lives in `w:pPr/w:rPr` and is what Word applies to the mark itself.
pub struct ParagraphMark<'a> {
    inner: &'a mut CT_RPr,
}

impl ParagraphMark<'_> {
    /// Set bold formatting on the paragraph mark.
    pub fn bold(mut self, val: bool) -> Self {
        self.set_bold(val);
        self
    }

    /// Set bold formatting on the paragraph mark in place.
    pub fn set_bold(&mut self, val: bool) {
        self.set_bold_value(Some(val));
    }

    /// Set or clear direct bold formatting on the paragraph mark.
    pub fn set_bold_value(&mut self, val: Option<bool>) {
        self.inner.bold = val;
        self.inner.bold_cs = val;
    }

    /// Set italic formatting on the paragraph mark.
    pub fn italic(mut self, val: bool) -> Self {
        self.set_italic(val);
        self
    }

    /// Set italic formatting on the paragraph mark in place.
    pub fn set_italic(&mut self, val: bool) {
        self.set_italic_value(Some(val));
    }

    /// Set or clear direct italic formatting on the paragraph mark.
    pub fn set_italic_value(&mut self, val: Option<bool>) {
        self.inner.italic = val;
        self.inner.italic_cs = val;
    }

    /// Set underline formatting on the paragraph mark.
    pub fn underline(mut self, val: bool) -> Self {
        self.set_underline(val);
        self
    }

    /// Set underline formatting on the paragraph mark in place.
    pub fn set_underline(&mut self, val: bool) {
        self.set_underline_value(Some(val));
    }

    /// Set or clear direct underline formatting on the paragraph mark.
    pub fn set_underline_value(&mut self, val: Option<bool>) {
        self.inner.underline = val.map(|val| {
            if val {
                ST_Underline::Single
            } else {
                ST_Underline::None
            }
        });
    }

    /// Set strikethrough formatting on the paragraph mark.
    pub fn strike(mut self, val: bool) -> Self {
        self.set_strike(val);
        self
    }

    /// Set strikethrough formatting on the paragraph mark in place.
    pub fn set_strike(&mut self, val: bool) {
        self.set_strike_value(Some(val));
    }

    /// Set or clear direct strikethrough formatting on the paragraph mark.
    pub fn set_strike_value(&mut self, val: Option<bool>) {
        self.inner.strike = val;
    }

    /// Set the paragraph mark font size in points.
    pub fn size(mut self, pt: f64) -> Self {
        self.set_size(pt);
        self
    }

    /// Set the paragraph mark font size in points in place.
    pub fn set_size(&mut self, pt: f64) {
        self.set_size_value(Some(pt));
    }

    /// Set or clear the direct paragraph mark font size.
    pub fn set_size_value(&mut self, pt: Option<f64>) {
        let half_points = pt.map(HalfPoint::from_pt);
        self.inner.sz = half_points;
        self.inner.sz_cs = half_points;
    }

    /// Set the paragraph mark font name.
    pub fn font(mut self, name: &str) -> Self {
        self.set_font(name);
        self
    }

    /// Set the paragraph mark font name in place.
    pub fn set_font(&mut self, name: &str) {
        self.set_font_value(Some(name));
    }

    /// Set or clear the direct paragraph mark font name.
    pub fn set_font_value(&mut self, name: Option<&str>) {
        self.inner.font_ascii = name.map(str::to_owned);
        self.inner.font_hansi = name.map(str::to_owned);
        self.inner.font_east_asia = name.map(str::to_owned);
        self.inner.font_cs = name.map(str::to_owned);
    }

    /// Set the paragraph mark color as a hex string, e.g. "FF0000".
    pub fn color(mut self, hex: &str) -> Self {
        self.set_color(hex);
        self
    }

    /// Set the paragraph mark color in place.
    pub fn set_color(&mut self, hex: &str) {
        self.set_color_value(Some(hex));
    }

    /// Set or clear the direct paragraph mark color.
    pub fn set_color_value(&mut self, hex: Option<&str>) {
        self.inner.color = hex.map(str::to_owned);
    }
}

/// An immutable view of the paragraph mark's own run properties.
#[derive(Debug, Clone, Copy)]
pub struct ParagraphMarkRef<'a> {
    inner: Option<&'a CT_RPr>,
}

impl ParagraphMarkRef<'_> {
    /// Whether the paragraph carries any direct mark formatting.
    pub fn is_present(&self) -> bool {
        self.inner.is_some()
    }

    /// Direct bold formatting on the paragraph mark.
    pub fn bold_value(&self) -> Option<bool> {
        self.inner.and_then(|rpr| rpr.bold)
    }

    /// Direct italic formatting on the paragraph mark.
    pub fn italic_value(&self) -> Option<bool> {
        self.inner.and_then(|rpr| rpr.italic)
    }

    /// Direct underline formatting on the paragraph mark.
    pub fn underline_value(&self) -> Option<bool> {
        self.inner
            .and_then(|rpr| rpr.underline)
            .map(|underline| underline != ST_Underline::None)
    }

    /// Direct strikethrough formatting on the paragraph mark.
    pub fn strike_value(&self) -> Option<bool> {
        self.inner.and_then(|rpr| rpr.strike)
    }

    /// The direct paragraph mark font size in points.
    pub fn size(&self) -> Option<f64> {
        self.inner.and_then(|rpr| rpr.sz).map(HalfPoint::to_pt)
    }

    /// The direct paragraph mark font name.
    pub fn font_name(&self) -> Option<&str> {
        self.inner.and_then(|rpr| rpr.font_ascii.as_deref())
    }

    /// The direct paragraph mark color.
    pub fn color(&self) -> Option<&str> {
        self.inner.and_then(|rpr| rpr.color.as_deref())
    }
}

/// A mutable reference to a paragraph in a document.
pub struct Paragraph<'a> {
    pub(crate) inner: &'a mut CT_P,
}

impl<'a> Paragraph<'a> {
    /// Get the combined text of all runs.
    pub fn text(&self) -> String {
        self.inner
            .accepted_bookmark_runs()
            .iter()
            .map(|run| run.text())
            .collect()
    }

    /// Iterate over typed equations in paragraph source order.
    pub fn equations(&self) -> impl Iterator<Item = &OfficeMath> {
        self.inner.equations.iter().map(|(_, _, equation)| equation)
    }

    /// Get an equation by equation index.
    pub fn equation(&self, index: usize) -> Option<&OfficeMath> {
        self.inner
            .equations
            .get(index)
            .map(|(_, _, equation)| equation)
    }

    /// Get a mutable equation by equation index.
    pub fn equation_mut(&mut self, index: usize) -> Option<&mut OfficeMath> {
        self.inner
            .equations
            .get_mut(index)
            .map(|(_, _, equation)| equation)
    }

    /// Append an inline or display equation at the final run boundary.
    pub fn add_equation(&mut self, equation: OfficeMath) -> crate::Result<()> {
        let run_index = self.inner.runs.len();
        let raw_before = self
            .inner
            .extra_xml
            .iter()
            .filter(|(position, _)| *position == run_index)
            .count();
        let raw = equation.to_xml()?;
        self.inner.extra_xml.push((run_index, raw));
        self.inner.equations.push((run_index, raw_before, equation));
        Ok(())
    }

    /// Append a footnote reference run (`<w:footnoteReference w:id="..."/>`).
    /// The footnote content itself is added via `Document::add_footnote`.
    pub fn add_footnote_ref(&mut self, id: i32) {
        use rdocx_oxml::text::{CT_R, RunContent};
        let mut r = CT_R::new("");
        r.content = vec![RunContent::FootnoteRef { id }];
        self.inner.runs.push(r);
    }

    /// Add a run with the given text and return a mutable reference for chaining.
    pub fn add_run(&mut self, text: &str) -> Run<'_> {
        self.inner.runs.push(CT_R::new(text));
        Run {
            inner: self.inner.runs.last_mut().unwrap(),
        }
    }

    /// Add a ruby phonetic guide and return its index in this paragraph.
    ///
    /// The base text becomes an ordinary paragraph run, so [`Self::text`]
    /// returns it. The phonetic text stays inside the annotation and no text
    /// extraction returns it. Use [`Self::ruby_mut`] to set the geometry in
    /// `w:rubyPr`.
    pub fn add_ruby(&mut self, base: &str, phonetic: &str) -> usize {
        let base_start = self.inner.runs.len();
        self.inner.runs.push(CT_R::new(base));
        self.inner.rubies.push(CT_Ruby {
            properties: None,
            ruby_text: vec![CT_R::new(phonetic)],
            base_start,
            base_end: base_start + 1,
            raw_xml: Vec::new(),
        });
        self.inner.rubies.len() - 1
    }

    /// Iterate over the paragraph's ruby annotations in source order.
    pub fn rubies(&self) -> impl Iterator<Item = &CT_Ruby> {
        self.inner.rubies.iter()
    }

    /// Get one ruby annotation by index.
    pub fn ruby(&self, index: usize) -> Option<&CT_Ruby> {
        self.inner.rubies.get(index)
    }

    /// Get one ruby annotation mutably by index.
    pub fn ruby_mut(&mut self, index: usize) -> Option<&mut CT_Ruby> {
        self.inner.rubies.get_mut(index)
    }

    /// Add a run that clones the paragraph mark's direct run properties.
    pub fn add_run_inheriting_mark(&mut self, text: &str) -> Run<'_> {
        let properties = self
            .inner
            .properties
            .as_ref()
            .and_then(|properties| properties.rpr.clone());
        let mut run = CT_R::new(text);
        run.properties = properties;
        self.inner.runs.push(run);
        Run {
            inner: self.inner.runs.last_mut().unwrap(),
        }
    }

    /// Add a line break in its own run.
    pub fn add_line_break(&mut self) {
        let mut run = CT_R::new("");
        run.content = vec![RunContent::Break(BreakType::Line)];
        self.inner.runs.push(run);
    }

    /// Add a tab in its own run.
    pub(crate) fn add_tab(&mut self) {
        let mut run = CT_R::new("");
        run.content = vec![RunContent::Tab];
        self.inner.runs.push(run);
    }

    /// Append an inline picture run using an existing document relationship.
    pub(crate) fn add_picture(&mut self, rel_id: &str, width: Length, height: Length) -> &mut CT_R {
        use rdocx_oxml::drawing::{CT_Drawing, CT_Inline};

        let inline = CT_Inline::new(rel_id, width.to_emu(), height.to_emu());
        let drawing = CT_Drawing::inline(inline);
        let mut run = CT_R::new("");
        run.content = vec![RunContent::Drawing(drawing)];
        self.inner.runs.push(run);
        self.inner
            .runs
            .last_mut()
            .expect("picture run was appended")
    }

    /// Add a run wrapped in an external hyperlink relationship.
    ///
    /// Obtain `relationship_id` from
    /// [`crate::Document::add_hyperlink_relationship`]. Returning the run
    /// allows the hyperlink text to receive the same direct formatting as any
    /// other run.
    pub fn add_hyperlink(&mut self, text: &str, relationship_id: &str) -> Run<'_> {
        self.add_hyperlink_with_tooltip(text, relationship_id, None)
    }

    /// Add a hyperlink run and optionally set its user-facing hover tooltip.
    pub fn add_hyperlink_with_tooltip(
        &mut self,
        text: &str,
        relationship_id: &str,
        tooltip: Option<&str>,
    ) -> Run<'_> {
        let run_start = self.inner.runs.len();
        self.inner.runs.push(CT_R::new(text));
        self.inner.hyperlinks.push(HyperlinkSpan {
            rel_id: Some(relationship_id.to_string()),
            anchor: None,
            tooltip: tooltip.map(str::to_owned),
            doc_location: None,
            run_start,
            run_end: run_start + 1,
            extra_attributes: Vec::new(),
            extra_xml: Vec::new(),
            preserved_raw_before: None,
        });
        Run {
            inner: self.inner.runs.last_mut().unwrap(),
        }
    }

    /// Get the number of runs in this paragraph.
    pub fn run_count(&self) -> usize {
        self.inner.accepted_run_paths().len()
    }

    /// Get an immutable accepted-view run by index.
    pub fn run(&self, index: usize) -> Option<RunRef<'_>> {
        let path = self.inner.accepted_run_paths().get(index)?.clone();
        self.inner.accepted_run(&path).map(|inner| RunRef { inner })
    }

    /// Get a direct mutable run when the accepted index is not nested.
    ///
    /// Use [`Self::edit_run`] for a run inside a tracked insertion or inline
    /// content control.
    pub fn run_mut(&mut self, index: usize) -> Option<Run<'_>> {
        let path = self.inner.accepted_run_paths().get(index)?.clone();
        let [AcceptedRunPathSegment::Run(direct)] = path.segments() else {
            return None;
        };
        self.inner.runs.get_mut(*direct).map(|inner| Run { inner })
    }

    /// Return the recursive source path for one accepted-view run.
    #[doc(hidden)]
    pub fn run_path(&self, index: usize) -> Option<AcceptedRunPath> {
        self.inner.accepted_run_paths().get(index).cloned()
    }

    /// Apply one checked edit to an accepted-view run at its recursive owner.
    #[doc(hidden)]
    pub fn edit_run(
        &mut self,
        path: &AcceptedRunPath,
        edit: impl FnOnce(&mut Run<'_>),
    ) -> crate::Result<()> {
        let mut replacement = self
            .inner
            .accepted_run(path)
            .cloned()
            .ok_or_else(|| crate::Error::Other("accepted run path is stale".to_owned()))?;
        edit(&mut Run {
            inner: &mut replacement,
        });
        if self.inner.replace_accepted_run(path, replacement)? {
            Ok(())
        } else {
            Err(crate::Error::Other("accepted run path is stale".to_owned()))
        }
    }

    /// Split the run at `run_index` at a Unicode scalar offset of its literal
    /// text.
    ///
    /// The run keeps the text before `offset`, and a new run at
    /// `run_index + 1` receives the rest with the same properties, inside the
    /// same hyperlink. Zero returns the boundary before the run and the literal
    /// text length returns the boundary after it without changing the paragraph.
    /// Non-text children have zero width. A comment or bookmark range can then
    /// start or end at the returned boundary. On error the paragraph is
    /// unchanged.
    pub fn split_run(&mut self, run_index: usize, offset: usize) -> crate::Result<usize> {
        let mut paragraph = self.inner.clone();
        let path = paragraph
            .accepted_run_paths()
            .get(run_index)
            .cloned()
            .ok_or_else(|| crate::Error::Other("run index out of range".to_owned()))?;
        let boundary = paragraph
            .split_accepted_run(&path, run_index, offset)
            .map_err(|error| crate::Error::Other(error.to_string()))?;
        *self.inner = paragraph;
        Ok(boundary)
    }

    /// Get an iterator over immutable run references.
    pub fn runs(&self) -> impl Iterator<Item = RunRef<'_>> {
        self.inner
            .accepted_bookmark_runs()
            .into_iter()
            .map(|inner| RunRef { inner })
    }

    /// Set the paragraph alignment.
    pub fn alignment(mut self, align: Alignment) -> Self {
        self.set_alignment(align);
        self
    }

    /// Set the paragraph alignment in place.
    pub fn set_alignment(&mut self, align: Alignment) {
        self.set_alignment_value(Some(align));
    }

    /// Set or clear direct paragraph alignment.
    pub fn set_alignment_value(&mut self, align: Option<Alignment>) {
        if align.is_none() && self.inner.properties.is_none() {
            return;
        }
        self.ensure_ppr().jc = align.map(Alignment::to_st_jc);
    }

    /// Set the paragraph style by ID.
    pub fn style(mut self, style_id: &str) -> Self {
        self.set_style(style_id);
        self
    }

    /// Set the paragraph style by ID in place.
    pub fn set_style(&mut self, style_id: &str) {
        self.ensure_ppr().style_id = Some(style_id.to_string());
    }

    /// Set or clear the paragraph style ID in place.
    pub fn set_style_value(&mut self, style_id: Option<&str>) {
        if style_id.is_none() && self.inner.properties.is_none() {
            return;
        }
        self.ensure_ppr().style_id = style_id.map(str::to_owned);
    }

    /// Attach this paragraph to a list definition as a list item.
    ///
    /// `num_id` comes from [`crate::Document::add_list_definition`] (or the
    /// shared definitions behind `add_bullet_list_item` /
    /// `add_numbered_list_item`); `level` is the 0-based indentation level.
    pub fn numbering(mut self, num_id: u32, level: u32) -> Self {
        let _ = self.set_numbering(num_id, level);
        self
    }

    /// Attach this paragraph to a list definition in place.
    ///
    /// Returns `false` without mutation when `level` is outside Word's
    /// supported range of 0 through 8.
    pub fn set_numbering(&mut self, num_id: u32, level: u32) -> bool {
        self.set_numbering_value(Some((num_id, level)))
    }

    /// Set or clear this paragraph's list numbering.
    ///
    /// Returns `false` without mutation for an out-of-range level.
    pub fn set_numbering_value(&mut self, numbering: Option<(u32, u32)>) -> bool {
        if numbering.is_some_and(|(_, level)| level > 8) {
            return false;
        }
        if numbering.is_none() && self.inner.properties.is_none() {
            return true;
        }
        let ppr = self.ensure_ppr();
        ppr.num_id = numbering.map(|(num_id, _)| num_id);
        ppr.num_ilvl = numbering.map(|(_, level)| level);
        true
    }

    /// Set space before the paragraph.
    pub fn space_before(mut self, length: Length) -> Self {
        self.set_space_before(length);
        self
    }

    /// Set space before the paragraph in place.
    pub fn set_space_before(&mut self, length: Length) {
        self.set_space_before_value(Some(length));
    }

    /// Set or clear direct space before the paragraph.
    pub fn set_space_before_value(&mut self, length: Option<Length>) {
        if length.is_none() && self.inner.properties.is_none() {
            return;
        }
        self.ensure_ppr().space_before = length.map(Length::as_twips);
    }

    /// Set space after the paragraph.
    pub fn space_after(mut self, length: Length) -> Self {
        self.set_space_after(length);
        self
    }

    /// Set space after the paragraph in place.
    pub fn set_space_after(&mut self, length: Length) {
        self.set_space_after_value(Some(length));
    }

    /// Set or clear direct space after the paragraph.
    pub fn set_space_after_value(&mut self, length: Option<Length>) {
        if length.is_none() && self.inner.properties.is_none() {
            return;
        }
        self.ensure_ppr().space_after = length.map(Length::as_twips);
    }

    /// Set left indentation.
    pub fn indent_left(mut self, length: Length) -> Self {
        self.set_indent_left(length);
        self
    }

    /// Set left indentation in place.
    pub fn set_indent_left(&mut self, length: Length) {
        self.set_indent_left_value(Some(length));
    }

    /// Set or clear direct left indentation.
    pub fn set_indent_left_value(&mut self, length: Option<Length>) {
        if length.is_none() && self.inner.properties.is_none() {
            return;
        }
        self.ensure_ppr().ind_left = length.map(Length::as_twips);
    }

    /// Set right indentation.
    pub fn indent_right(mut self, length: Length) -> Self {
        self.set_indent_right(length);
        self
    }

    /// Set right indentation in place.
    pub fn set_indent_right(&mut self, length: Length) {
        self.set_indent_right_value(Some(length));
    }

    /// Set or clear direct right indentation.
    pub fn set_indent_right_value(&mut self, length: Option<Length>) {
        if length.is_none() && self.inner.properties.is_none() {
            return;
        }
        self.ensure_ppr().ind_right = length.map(Length::as_twips);
    }

    /// Set first line indent.
    pub fn first_line_indent(mut self, length: Length) -> Self {
        self.set_first_line_indent(length);
        self
    }

    /// Set first line indent in place.
    pub fn set_first_line_indent(&mut self, length: Length) {
        self.ensure_ppr().ind_first_line = Some(length.as_twips());
    }

    /// Set or clear the signed first-line indentation.
    ///
    /// Negative values are stored as a positive hanging indentation.
    pub fn set_signed_first_line_indent_value(&mut self, length: Option<Length>) {
        if length.is_none() && self.inner.properties.is_none() {
            return;
        }
        let ppr = self.ensure_ppr();
        match length.map(Length::to_twips) {
            Some(twips) if twips < 0 => {
                ppr.ind_first_line = None;
                ppr.ind_hanging = Some(Twips(twips.saturating_abs()));
            }
            Some(twips) => {
                ppr.ind_first_line = Some(Twips(twips));
                ppr.ind_hanging = None;
            }
            None => {
                ppr.ind_first_line = None;
                ppr.ind_hanging = None;
            }
        }
    }

    /// Set hanging indent.
    pub fn hanging_indent(mut self, length: Length) -> Self {
        self.set_hanging_indent(length);
        self
    }

    /// Set hanging indent in place.
    pub fn set_hanging_indent(&mut self, length: Length) {
        self.ensure_ppr().ind_hanging = Some(length.as_twips());
    }

    /// Set keep with next paragraph.
    pub fn keep_with_next(mut self, val: bool) -> Self {
        self.set_keep_with_next(val);
        self
    }

    /// Set keep with next paragraph in place.
    pub fn set_keep_with_next(&mut self, val: bool) {
        self.set_keep_with_next_value(Some(val));
    }

    /// Set or clear direct keep-with-next formatting.
    pub fn set_keep_with_next_value(&mut self, val: Option<bool>) {
        if val.is_none() && self.inner.properties.is_none() {
            return;
        }
        self.ensure_ppr().keep_next = val;
    }

    /// Set keep lines together.
    pub fn keep_together(mut self, val: bool) -> Self {
        self.set_keep_together(val);
        self
    }

    /// Set keep lines together in place.
    pub fn set_keep_together(&mut self, val: bool) {
        self.set_keep_together_value(Some(val));
    }

    /// Set or clear direct keep-together formatting.
    pub fn set_keep_together_value(&mut self, val: Option<bool>) {
        if val.is_none() && self.inner.properties.is_none() {
            return;
        }
        self.ensure_ppr().keep_lines = val;
    }

    /// Set page break before.
    pub fn page_break_before(mut self, val: bool) -> Self {
        self.set_page_break_before(val);
        self
    }

    /// Set page break before in place.
    pub fn set_page_break_before(&mut self, val: bool) {
        self.set_page_break_before_value(Some(val));
    }

    /// Set or clear direct page-break-before formatting.
    pub fn set_page_break_before_value(&mut self, val: Option<bool>) {
        if val.is_none() && self.inner.properties.is_none() {
            return;
        }
        self.ensure_ppr().page_break_before = val;
    }

    /// Set widow/orphan control.
    pub fn widow_control(mut self, val: bool) -> Self {
        self.set_widow_control(val);
        self
    }

    /// Set widow/orphan control in place.
    pub fn set_widow_control(&mut self, val: bool) {
        self.set_widow_control_value(Some(val));
    }

    /// Set or clear direct widow-control formatting.
    pub fn set_widow_control_value(&mut self, val: Option<bool>) {
        if val.is_none() && self.inner.properties.is_none() {
            return;
        }
        self.ensure_ppr().widow_control = val;
    }

    /// Set line spacing in points with "exact" rule.
    pub fn line_spacing(mut self, pt: f64) -> Self {
        self.set_line_spacing(pt);
        self
    }

    /// Set exact line spacing in place.
    pub fn set_line_spacing(&mut self, pt: f64) {
        let ppr = self.ensure_ppr();
        ppr.line_spacing = Some(Twips::from_pt(pt));
        ppr.line_rule = Some("exact".to_string());
    }

    pub(crate) fn set_line_spacing_at_least(&mut self, pt: f64) {
        let ppr = self.ensure_ppr();
        ppr.line_spacing = Some(Twips::from_pt(pt));
        ppr.line_rule = Some("atLeast".to_string());
    }

    /// Clear direct line spacing.
    pub fn clear_line_spacing(&mut self) {
        if let Some(ppr) = self.inner.properties.as_mut() {
            ppr.line_spacing = None;
            ppr.line_rule = None;
        }
    }

    /// Set line spacing with a multiplier (1.0 = single, 1.5, 2.0 = double, etc.).
    pub fn line_spacing_multiple(mut self, multiple: f64) -> Self {
        self.set_line_spacing_multiple(multiple);
        self
    }

    /// Set multiplied line spacing in place.
    pub fn set_line_spacing_multiple(&mut self, multiple: f64) {
        let ppr = self.ensure_ppr();
        // In "auto" mode, line spacing is in 240ths of a line (240 = single)
        ppr.line_spacing = Some(Twips((multiple * 240.0) as i32));
        ppr.line_rule = Some("auto".to_string());
    }

    /// Set a background/shading fill color (hex string, e.g. "FFFF00").
    pub fn shading(mut self, fill_color: &str) -> Self {
        self.set_shading(fill_color);
        self
    }

    /// Set a background/shading fill color in place.
    pub fn set_shading(&mut self, fill_color: &str) {
        self.ensure_ppr().shading = Some(Box::new(CT_Shd {
            val: "clear".to_string(),
            color: Some("auto".to_string()),
            fill: Some(fill_color.to_string()),
            ..Default::default()
        }));
    }

    /// Add a border to all sides.
    pub fn border_all(mut self, style: BorderStyle, size_eighths_pt: u32, color: &str) -> Self {
        self.set_border_all(style, size_eighths_pt, color);
        self
    }

    /// Add a border to all sides in place.
    pub fn set_border_all(&mut self, style: BorderStyle, size_eighths_pt: u32, color: &str) {
        let edge = CT_BorderEdge {
            val: style.to_st(),
            sz: Some(size_eighths_pt),
            space: Some(1),
            color: Some(color.to_string()),
            extra_attributes: Vec::new(),
            nil: false,
        };
        self.ensure_ppr().borders = Some(Box::new(CT_PBdr {
            top: Some(edge.clone()),
            bottom: Some(edge.clone()),
            left: Some(edge.clone()),
            right: Some(edge),
            between: None,
            bar: None,
        }));
    }

    /// Add a bottom border only.
    pub fn border_bottom(mut self, style: BorderStyle, size_eighths_pt: u32, color: &str) -> Self {
        self.set_border_bottom(style, size_eighths_pt, color);
        self
    }

    /// Add a bottom border with an explicit content gap in points.
    pub fn border_bottom_with_space(
        mut self,
        style: BorderStyle,
        size_eighths_pt: u32,
        space_points: u32,
        color: &str,
    ) -> Self {
        self.set_border_bottom_with_space(style, size_eighths_pt, space_points, color);
        self
    }

    /// Add a bottom border in place.
    pub fn set_border_bottom(&mut self, style: BorderStyle, size_eighths_pt: u32, color: &str) {
        self.set_border_bottom_with_space(style, size_eighths_pt, 1, color);
    }

    /// Add a bottom border with an explicit content gap in points in place.
    pub fn set_border_bottom_with_space(
        &mut self,
        style: BorderStyle,
        size_eighths_pt: u32,
        space_points: u32,
        color: &str,
    ) {
        let edge = CT_BorderEdge {
            val: style.to_st(),
            sz: Some(size_eighths_pt),
            space: Some(space_points),
            color: Some(color.to_string()),
            extra_attributes: Vec::new(),
            nil: false,
        };
        let borders = self
            .ensure_ppr()
            .borders
            .get_or_insert_with(|| Box::new(CT_PBdr::default()));
        borders.bottom = Some(edge);
    }

    /// Add a tab stop.
    pub fn add_tab_stop(mut self, alignment: TabAlignment, position: Length) -> Self {
        self.set_add_tab_stop(alignment, position);
        self
    }

    /// Add a tab stop in place.
    pub fn set_add_tab_stop(&mut self, alignment: TabAlignment, position: Length) {
        let tabs = self
            .ensure_ppr()
            .tabs
            .get_or_insert_with(|| CT_Tabs { tabs: Vec::new() });
        tabs.tabs.push(CT_TabStop {
            val: alignment.to_st(),
            pos: position.as_twips(),
            leader: None,
            source_occurrence: None,
        });
    }

    /// Add a tab stop with a leader character.
    pub fn add_tab_stop_with_leader(
        mut self,
        alignment: TabAlignment,
        position: Length,
        leader: TabLeader,
    ) -> Self {
        self.set_add_tab_stop_with_leader(alignment, position, leader);
        self
    }

    /// Add a tab stop with a leader character in place.
    pub fn set_add_tab_stop_with_leader(
        &mut self,
        alignment: TabAlignment,
        position: Length,
        leader: TabLeader,
    ) {
        let tabs = self
            .ensure_ppr()
            .tabs
            .get_or_insert_with(|| CT_Tabs { tabs: Vec::new() });
        tabs.tabs.push(CT_TabStop {
            val: alignment.to_st(),
            pos: position.as_twips(),
            leader: Some(leader.to_st()),
            source_occurrence: None,
        });
    }

    /// Set outline level (0–9, used for TOC generation).
    pub fn outline_level(mut self, level: u32) -> Self {
        self.set_outline_level(level);
        self
    }

    /// Set outline level in place.
    ///
    /// A level above 9 leaves the paragraph unchanged. See
    /// [`Self::set_outline_level_value`].
    pub fn set_outline_level(&mut self, level: u32) {
        let _ = self.set_outline_level_value(Some(level));
    }

    /// Add a section break after this paragraph.
    ///
    /// This creates a `<w:sectPr>` inside the paragraph's properties,
    /// ending the current section at this paragraph.
    pub fn section_break(mut self, break_type: SectionBreak) -> Self {
        self.set_section_break(break_type);
        self
    }

    /// Add a section break after this paragraph in place.
    pub fn set_section_break(&mut self, break_type: SectionBreak) {
        let sect = self.ensure_sect_pr();
        sect.section_type = Some(break_type.to_st());
    }

    /// Set the section ending at this paragraph to landscape orientation.
    ///
    /// Sets page dimensions to 11" x 8.5" (US Letter landscape).
    /// Must be combined with `section_break()` to create a section break.
    pub fn section_landscape(mut self) -> Self {
        self.set_section_landscape();
        self
    }

    /// Set the section ending at this paragraph to landscape in place.
    pub fn set_section_landscape(&mut self) {
        let sect = self.ensure_sect_pr();
        sect.orientation = Some(ST_PageOrientation::Landscape);
        sect.page_width = Some(Twips(15840)); // 11"
        sect.page_height = Some(Twips(12240)); // 8.5"
    }

    /// Set the section ending at this paragraph to portrait orientation.
    ///
    /// Sets page dimensions to 8.5" x 11" (US Letter portrait).
    /// Must be combined with `section_break()` to create a section break.
    pub fn section_portrait(mut self) -> Self {
        self.set_section_portrait();
        self
    }

    /// Set the section ending at this paragraph to portrait in place.
    pub fn set_section_portrait(&mut self) {
        let sect = self.ensure_sect_pr();
        sect.orientation = Some(ST_PageOrientation::Portrait);
        sect.page_width = Some(Twips(12240)); // 8.5"
        sect.page_height = Some(Twips(15840)); // 11"
    }

    /// Set custom page dimensions for the section ending at this paragraph.
    pub fn section_page_size(mut self, width: crate::Length, height: crate::Length) -> Self {
        self.set_section_page_size(width, height);
        self
    }

    /// Set custom page dimensions for the section in place.
    pub fn set_section_page_size(&mut self, width: crate::Length, height: crate::Length) {
        let sect = self.ensure_sect_pr();
        sect.page_width = Some(width.as_twips());
        sect.page_height = Some(height.as_twips());
    }

    /// Set logical leading-edge indentation.
    pub fn indent_start(mut self, length: Length) -> Self {
        self.set_indent_start(length);
        self
    }

    /// Set logical leading-edge indentation in place.
    pub fn set_indent_start(&mut self, length: Length) {
        self.set_indent_start_value(Some(length));
    }

    /// Set or clear direct logical leading-edge indentation.
    pub fn set_indent_start_value(&mut self, length: Option<Length>) {
        if length.is_none() && self.inner.properties.is_none() {
            return;
        }
        self.ensure_ppr().ind_start = length.map(Length::as_twips);
    }

    /// Set logical trailing-edge indentation.
    pub fn indent_end(mut self, length: Length) -> Self {
        self.set_indent_end(length);
        self
    }

    /// Set logical trailing-edge indentation in place.
    pub fn set_indent_end(&mut self, length: Length) {
        self.set_indent_end_value(Some(length));
    }

    /// Set or clear direct logical trailing-edge indentation.
    pub fn set_indent_end_value(&mut self, length: Option<Length>) {
        if length.is_none() && self.inner.properties.is_none() {
            return;
        }
        self.ensure_ppr().ind_end = length.map(Length::as_twips);
    }

    /// Set or clear the direct hanging indentation.
    pub fn set_hanging_indent_value(&mut self, length: Option<Length>) {
        if length.is_none() && self.inner.properties.is_none() {
            return;
        }
        self.ensure_ppr().ind_hanging = length.map(Length::as_twips);
    }

    /// Swap inside and outside indentation on facing pages.
    pub fn mirror_indents(mut self, val: bool) -> Self {
        self.set_mirror_indents(val);
        self
    }

    /// Swap inside and outside indentation on facing pages in place.
    pub fn set_mirror_indents(&mut self, val: bool) {
        self.set_mirror_indents_value(Some(val));
    }

    /// Set or clear direct mirrored indentation.
    pub fn set_mirror_indents_value(&mut self, val: Option<bool>) {
        if val.is_none() && self.inner.properties.is_none() {
            return;
        }
        self.ensure_ppr().mirror_indents = val;
    }

    /// Adjust the right indentation for a document grid.
    pub fn adjust_right_indent(mut self, val: bool) -> Self {
        self.set_adjust_right_indent(val);
        self
    }

    /// Adjust the right indentation for a document grid in place.
    pub fn set_adjust_right_indent(&mut self, val: bool) {
        self.set_adjust_right_indent_value(Some(val));
    }

    /// Set or clear the direct document-grid right indentation adjustment.
    pub fn set_adjust_right_indent_value(&mut self, val: Option<bool>) {
        if val.is_none() && self.inner.properties.is_none() {
            return;
        }
        self.ensure_ppr().adjust_right_ind = val;
    }

    // The seven East Asian paragraph toggles. Each takes the `Option` form
    // alone, because a three-state toggle needs nothing else to create, read,
    // change and remove it, and this paragraph handle is already large.

    /// Set or clear direct `w:kinsoku`, East Asian line-breaking rules.
    pub fn set_kinsoku_value(&mut self, val: Option<bool>) {
        if val.is_none() && self.inner.properties.is_none() {
            return;
        }
        self.ensure_ppr().kinsoku = val;
    }

    /// Set or clear direct `w:wordWrap`, breaking a Latin word only at a break opportunity.
    pub fn set_word_wrap_value(&mut self, val: Option<bool>) {
        if val.is_none() && self.inner.properties.is_none() {
            return;
        }
        self.ensure_ppr().word_wrap = val;
    }

    /// Set or clear direct `w:overflowPunct`, letting trailing punctuation hang past the measure.
    pub fn set_overflow_punct_value(&mut self, val: Option<bool>) {
        if val.is_none() && self.inner.properties.is_none() {
            return;
        }
        self.ensure_ppr().overflow_punct = val;
    }

    /// Set or clear direct `w:topLinePunct`, compressing punctuation at the start of a line.
    pub fn set_top_line_punct_value(&mut self, val: Option<bool>) {
        if val.is_none() && self.inner.properties.is_none() {
            return;
        }
        self.ensure_ppr().top_line_punct = val;
    }

    /// Set or clear direct `w:autoSpaceDE`, automatic spacing between East Asian and Latin text.
    pub fn set_auto_space_de_value(&mut self, val: Option<bool>) {
        if val.is_none() && self.inner.properties.is_none() {
            return;
        }
        self.ensure_ppr().auto_space_de = val;
    }

    /// Set or clear direct `w:autoSpaceDN`, automatic spacing between East Asian text and digits.
    pub fn set_auto_space_dn_value(&mut self, val: Option<bool>) {
        if val.is_none() && self.inner.properties.is_none() {
            return;
        }
        self.ensure_ppr().auto_space_dn = val;
    }

    /// Set or clear direct `w:snapToGrid`, snapping line advance to the section character grid.
    pub fn set_snap_to_grid_value(&mut self, val: Option<bool>) {
        if val.is_none() && self.inner.properties.is_none() {
            return;
        }
        self.ensure_ppr().snap_to_grid = val;
    }

    /// Let the consumer choose the space before the paragraph.
    pub fn space_before_auto(mut self, val: bool) -> Self {
        self.set_space_before_auto(val);
        self
    }

    /// Let the consumer choose the space before the paragraph in place.
    pub fn set_space_before_auto(&mut self, val: bool) {
        self.set_space_before_auto_value(Some(val));
    }

    /// Set or clear direct automatic spacing before the paragraph.
    pub fn set_space_before_auto_value(&mut self, val: Option<bool>) {
        if val.is_none() && self.inner.properties.is_none() {
            return;
        }
        self.ensure_ppr().before_autospacing = val;
    }

    /// Let the consumer choose the space after the paragraph.
    pub fn space_after_auto(mut self, val: bool) -> Self {
        self.set_space_after_auto(val);
        self
    }

    /// Let the consumer choose the space after the paragraph in place.
    pub fn set_space_after_auto(&mut self, val: bool) {
        self.set_space_after_auto_value(Some(val));
    }

    /// Set or clear direct automatic spacing after the paragraph.
    pub fn set_space_after_auto_value(&mut self, val: Option<bool>) {
        if val.is_none() && self.inner.properties.is_none() {
            return;
        }
        self.ensure_ppr().after_autospacing = val;
    }

    /// Drop the spacing between paragraphs that share this paragraph's style.
    pub fn contextual_spacing(mut self, val: bool) -> Self {
        self.set_contextual_spacing(val);
        self
    }

    /// Drop the spacing between paragraphs of the same style in place.
    pub fn set_contextual_spacing(&mut self, val: bool) {
        self.set_contextual_spacing_value(Some(val));
    }

    /// Set or clear direct contextual spacing.
    pub fn set_contextual_spacing_value(&mut self, val: Option<bool>) {
        if val.is_none() && self.inner.properties.is_none() {
            return;
        }
        self.ensure_ppr().contextual_spacing = val;
    }

    /// Set one paragraph border edge, leaving the other edges alone.
    pub fn border(
        mut self,
        edge: ParagraphBorderEdge,
        style: BorderStyle,
        size_eighths_pt: u32,
        color: &str,
    ) -> Self {
        self.set_border(edge, style, size_eighths_pt, color);
        self
    }

    /// Set one paragraph border edge in place.
    pub fn set_border(
        &mut self,
        edge: ParagraphBorderEdge,
        style: BorderStyle,
        size_eighths_pt: u32,
        color: &str,
    ) {
        self.set_border_value(edge, Some((style, size_eighths_pt, color)));
    }

    /// Set or remove one paragraph border edge.
    ///
    /// Removing the last edge removes the `w:pBdr` element with it. An edge
    /// that already exists keeps the attributes this crate does not model.
    pub fn set_border_value(
        &mut self,
        edge: ParagraphBorderEdge,
        border: Option<(BorderStyle, u32, &str)>,
    ) {
        if border.is_none() && self.inner.properties.is_none() {
            return;
        }
        let borders = self
            .ensure_ppr()
            .borders
            .get_or_insert_with(|| Box::new(CT_PBdr::default()));
        let slot = match edge {
            ParagraphBorderEdge::Top => &mut borders.top,
            ParagraphBorderEdge::Bottom => &mut borders.bottom,
            ParagraphBorderEdge::Left => &mut borders.left,
            ParagraphBorderEdge::Right => &mut borders.right,
            ParagraphBorderEdge::Between => &mut borders.between,
            ParagraphBorderEdge::Bar => &mut borders.bar,
        };
        match border {
            Some((style, size_eighths_pt, color)) => {
                let retained = slot
                    .take()
                    .map(|existing| existing.extra_attributes)
                    .unwrap_or_default();
                *slot = Some(CT_BorderEdge {
                    val: style.to_st(),
                    sz: Some(size_eighths_pt),
                    space: Some(1),
                    color: Some(color.to_owned()),
                    extra_attributes: retained,
                    nil: false,
                });
            }
            None => *slot = None,
        }
        if borders.is_empty() {
            self.ensure_ppr().borders = None;
        }
    }

    /// Remove every direct paragraph border edge.
    pub fn clear_borders(&mut self) {
        if let Some(ppr) = self.inner.properties.as_mut() {
            ppr.borders = None;
        }
    }

    /// Set a shading pattern with its own foreground and background colors.
    pub fn shading_pattern(mut self, pattern: &str, fill_color: &str, color: &str) -> Self {
        self.set_shading_pattern(pattern, fill_color, color);
        self
    }

    /// Set a shading pattern in place.
    pub fn set_shading_pattern(&mut self, pattern: &str, fill_color: &str, color: &str) {
        self.set_shading_value(Some((pattern, fill_color, color)));
    }

    /// Set or remove the direct paragraph shading.
    pub fn set_shading_value(&mut self, shading: Option<(&str, &str, &str)>) {
        if shading.is_none() && self.inner.properties.is_none() {
            return;
        }
        self.ensure_ppr().shading = shading.map(|(pattern, fill_color, color)| {
            Box::new(CT_Shd {
                val: pattern.to_owned(),
                color: Some(color.to_owned()),
                fill: Some(fill_color.to_owned()),
                ..Default::default()
            })
        });
    }

    /// Get one tab stop by index, or `None` when the index is out of range.
    pub fn tab_stop(&self, index: usize) -> Option<TabStopRef<'_>> {
        self.inner
            .properties
            .as_ref()?
            .tabs
            .as_ref()?
            .tabs
            .get(index)
            .map(|inner| TabStopRef { inner })
    }

    /// Replace one tab stop by index.
    ///
    /// Returns `false` without mutation when the index is out of range.
    pub fn set_tab_stop(
        &mut self,
        index: usize,
        alignment: TabAlignment,
        position: Length,
        leader: Option<TabLeader>,
    ) -> bool {
        let Some(tabs) = self
            .inner
            .properties
            .as_mut()
            .and_then(|ppr| ppr.tabs.as_mut())
        else {
            return false;
        };
        let Some(tab) = tabs.tabs.get_mut(index) else {
            return false;
        };
        tab.val = alignment.to_st();
        tab.pos = position.as_twips();
        tab.leader = leader.map(TabLeader::to_st);
        true
    }

    /// Remove one tab stop by index, keeping the order of the rest.
    ///
    /// Returns `false` without mutation when the index is out of range.
    pub fn remove_tab_stop(&mut self, index: usize) -> bool {
        let Some(tabs) = self
            .inner
            .properties
            .as_mut()
            .and_then(|ppr| ppr.tabs.as_mut())
        else {
            return false;
        };
        if index >= tabs.tabs.len() {
            return false;
        }
        tabs.tabs.remove(index);
        if tabs.tabs.is_empty() {
            self.ensure_ppr().tabs = None;
        }
        true
    }

    /// Remove every direct tab stop.
    pub fn clear_tab_stops(&mut self) {
        if let Some(ppr) = self.inner.properties.as_mut() {
            ppr.tabs = None;
        }
    }

    /// Suppress line numbering for this paragraph.
    pub fn suppress_line_numbers(mut self, val: bool) -> Self {
        self.set_suppress_line_numbers(val);
        self
    }

    /// Suppress line numbering for this paragraph in place.
    pub fn set_suppress_line_numbers(&mut self, val: bool) {
        self.set_suppress_line_numbers_value(Some(val));
    }

    /// Set or clear direct line-number suppression.
    pub fn set_suppress_line_numbers_value(&mut self, val: Option<bool>) {
        if val.is_none() && self.inner.properties.is_none() {
            return;
        }
        self.ensure_ppr().suppress_line_numbers = val;
    }

    /// Suppress automatic hyphenation for this paragraph.
    pub fn suppress_auto_hyphens(mut self, val: bool) -> Self {
        self.set_suppress_auto_hyphens(val);
        self
    }

    /// Suppress automatic hyphenation for this paragraph in place.
    pub fn set_suppress_auto_hyphens(&mut self, val: bool) {
        self.set_suppress_auto_hyphens_value(Some(val));
    }

    /// Set or clear direct automatic-hyphenation suppression.
    pub fn set_suppress_auto_hyphens_value(&mut self, val: Option<bool>) {
        if val.is_none() && self.inner.properties.is_none() {
            return;
        }
        self.ensure_ppr().suppress_auto_hyphens = val;
    }

    /// Place this paragraph in a text frame.
    pub fn frame(mut self, frame: ParagraphFrame) -> Self {
        self.set_frame(frame);
        self
    }

    /// Place this paragraph in a text frame in place.
    pub fn set_frame(&mut self, frame: ParagraphFrame) {
        self.set_frame_value(Some(frame));
    }

    /// Set or remove the direct text frame.
    pub fn set_frame_value(&mut self, frame: Option<ParagraphFrame>) {
        if frame.is_none() && self.inner.properties.is_none() {
            return;
        }
        let ppr = self.ensure_ppr();
        match frame {
            Some(frame) => frame.apply_to(
                ppr.frame
                    .get_or_insert_with(|| Box::new(CT_FramePr::default())),
            ),
            None => ppr.frame = None,
        }
    }

    /// Suppress overlap with a neighbouring frame.
    pub fn suppress_overlap(mut self, val: bool) -> Self {
        self.set_suppress_overlap(val);
        self
    }

    /// Suppress overlap with a neighbouring frame in place.
    pub fn set_suppress_overlap(&mut self, val: bool) {
        self.set_suppress_overlap_value(Some(val));
    }

    /// Set or clear direct frame overlap suppression.
    pub fn set_suppress_overlap_value(&mut self, val: Option<bool>) {
        if val.is_none() && self.inner.properties.is_none() {
            return;
        }
        self.ensure_ppr().suppress_overlap = val;
    }

    /// Set which lines a surrounding text box wraps tightly against.
    pub fn textbox_tight_wrap(mut self, wrap: TextboxTightWrap) -> Self {
        self.set_textbox_tight_wrap(wrap);
        self
    }

    /// Set the text box tight wrap mode in place.
    pub fn set_textbox_tight_wrap(&mut self, wrap: TextboxTightWrap) {
        self.set_textbox_tight_wrap_value(Some(wrap));
    }

    /// Set or clear the direct text box tight wrap mode.
    pub fn set_textbox_tight_wrap_value(&mut self, wrap: Option<TextboxTightWrap>) {
        if wrap.is_none() && self.inner.properties.is_none() {
            return;
        }
        self.ensure_ppr().textbox_tight_wrap = wrap.map(|wrap| wrap.to_str().to_owned());
    }

    /// Set or clear the direct outline level.
    ///
    /// Returns `false` without mutation above level 9, because Word writes 0
    /// through 8 for heading levels and 9 for body text.
    pub fn set_outline_level_value(&mut self, level: Option<u32>) -> bool {
        if level.is_some_and(|level| level > 9) {
            return false;
        }
        if level.is_none() && self.inner.properties.is_none() {
            return true;
        }
        self.ensure_ppr().outline_lvl = level;
        true
    }

    /// Set the paragraph base direction to right to left.
    pub fn right_to_left(mut self, val: bool) -> Self {
        self.set_right_to_left(val);
        self
    }

    /// Set the paragraph base direction in place.
    pub fn set_right_to_left(&mut self, val: bool) {
        self.set_right_to_left_value(Some(val));
    }

    /// Set or clear the direct paragraph base direction.
    pub fn set_right_to_left_value(&mut self, val: Option<bool>) {
        if val.is_none() && self.inner.properties.is_none() {
            return;
        }
        self.ensure_ppr().bidi = val;
    }

    /// Set the paragraph text flow.
    pub fn text_direction(mut self, direction: ParagraphTextDirection) -> Self {
        self.set_text_direction(direction);
        self
    }

    /// Set the paragraph text flow in place.
    pub fn set_text_direction(&mut self, direction: ParagraphTextDirection) {
        self.set_text_direction_value(Some(direction));
    }

    /// Set or clear the direct paragraph text flow.
    pub fn set_text_direction_value(&mut self, direction: Option<ParagraphTextDirection>) {
        if direction.is_none() && self.inner.properties.is_none() {
            return;
        }
        self.ensure_ppr().text_direction = direction.map(|direction| direction.to_str().to_owned());
    }

    /// Set the vertical alignment of characters on a line.
    pub fn text_alignment(mut self, alignment: ParagraphTextAlignment) -> Self {
        self.set_text_alignment(alignment);
        self
    }

    /// Set the vertical alignment of characters on a line in place.
    pub fn set_text_alignment(&mut self, alignment: ParagraphTextAlignment) {
        self.set_text_alignment_value(Some(alignment));
    }

    /// Set or clear the direct vertical character alignment.
    pub fn set_text_alignment_value(&mut self, alignment: Option<ParagraphTextAlignment>) {
        if alignment.is_none() && self.inner.properties.is_none() {
            return;
        }
        self.ensure_ppr().text_alignment = alignment.map(|alignment| alignment.to_str().to_owned());
    }

    /// Set or clear the web settings division this paragraph belongs to.
    pub fn set_div_id_value(&mut self, div_id: Option<u32>) {
        if div_id.is_none() && self.inner.properties.is_none() {
            return;
        }
        self.ensure_ppr().div_id = div_id;
    }

    /// Set or clear the conditional table-style regions this paragraph selects.
    ///
    /// A paragraph, a row and a cell all select regions through the same
    /// shape. A bit set on any of them selects the region for the cell.
    pub fn set_conditional_formatting(&mut self, regions: Option<TableConditionalFormatting>) {
        if regions.is_none() && self.inner.properties.is_none() {
            return;
        }
        self.ensure_ppr().cnf_style = regions.map(|regions| regions.to_value());
    }

    /// Edit the formatting of the paragraph mark itself.
    pub fn mark(&mut self) -> ParagraphMark<'_> {
        ParagraphMark {
            inner: self.ensure_ppr().rpr.get_or_insert_with(CT_RPr::default),
        }
    }

    /// Remove every direct paragraph mark property.
    pub fn clear_mark(&mut self) {
        if let Some(ppr) = self.inner.properties.as_mut() {
            ppr.rpr = None;
        }
    }

    fn ensure_ppr(&mut self) -> &mut CT_PPr {
        self.inner.properties.get_or_insert_with(CT_PPr::default)
    }

    fn ensure_sect_pr(&mut self) -> &mut CT_SectPr {
        let ppr = self.ensure_ppr();
        ppr.sect_pr.get_or_insert_with(CT_SectPr::default_letter)
    }
}

/// Section break type.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SectionBreak {
    /// Start a new section on the next page.
    NextPage,
    /// Start a new section on the same page (continuous).
    Continuous,
    /// Start a new section on the next even-numbered page.
    EvenPage,
    /// Start a new section on the next odd-numbered page.
    OddPage,
}

impl SectionBreak {
    fn to_st(self) -> ST_SectionType {
        match self {
            SectionBreak::NextPage => ST_SectionType::NextPage,
            SectionBreak::Continuous => ST_SectionType::Continuous,
            SectionBreak::EvenPage => ST_SectionType::EvenPage,
            SectionBreak::OddPage => ST_SectionType::OddPage,
        }
    }
}

/// An immutable reference to a paragraph.
pub struct ParagraphRef<'a> {
    pub(crate) inner: &'a CT_P,
}

impl<'a> ParagraphRef<'a> {
    /// Get the combined text of all runs.
    pub fn text(&self) -> String {
        self.inner
            .accepted_bookmark_runs()
            .iter()
            .map(|run| run.text())
            .collect()
    }

    /// Iterate over typed equations in paragraph source order.
    pub fn equations(&self) -> impl Iterator<Item = &'a OfficeMath> {
        self.inner.equations.iter().map(|(_, _, equation)| equation)
    }

    /// Get an equation by equation index.
    pub fn equation(&self, index: usize) -> Option<&'a OfficeMath> {
        self.inner
            .equations
            .get(index)
            .map(|(_, _, equation)| equation)
    }

    /// Iterate over direct paragraph items in source order.
    ///
    /// Unlike [`Self::runs`], this retains hyperlinks, content controls,
    /// revisions, comment ranges, bookmarks, and preserved unmodelled XML.
    pub fn items(&self) -> impl Iterator<Item = ParagraphItemRef<'_>> {
        let mut items = Vec::new();
        let mut run_index = 0;
        while run_index <= self.inner.runs.len() {
            let extras = self
                .inner
                .extra_xml
                .iter()
                .filter(|(at, _)| *at == run_index)
                .map(|(_, raw)| raw)
                .collect::<Vec<_>>();
            for raw_index in 0..=extras.len() {
                let markers = self
                    .inner
                    .comment_ranges
                    .iter()
                    .filter(|marker| match marker {
                        CommentRangeMarker::Start {
                            run_index: at,
                            raw_before,
                            ..
                        }
                        | CommentRangeMarker::End {
                            run_index: at,
                            raw_before,
                            ..
                        } => *at == run_index && (*raw_before).min(extras.len()) == raw_index,
                    })
                    .collect::<Vec<_>>();
                for marker_index in 0..=markers.len() {
                    items.extend(
                        self.inner
                            .content_controls
                            .iter()
                            .filter(|(at, raw_before, markers_before, _)| {
                                *at == run_index
                                    && (*raw_before).min(extras.len()) == raw_index
                                    && (*markers_before).min(markers.len()) == marker_index
                            })
                            .map(|(_, _, _, control)| {
                                ParagraphItemRef::ContentControl(ContentControlRef {
                                    inner: control,
                                })
                            }),
                    );
                    if let Some(marker) = markers.get(marker_index) {
                        items.push(match marker {
                            CommentRangeMarker::Start {
                                id,
                                has_child_content,
                                ..
                            } => ParagraphItemRef::CommentRangeStart {
                                id: *id,
                                has_child_content: *has_child_content,
                            },
                            CommentRangeMarker::End {
                                id,
                                has_child_content,
                                ..
                            } => ParagraphItemRef::CommentRangeEnd {
                                id: *id,
                                has_child_content: *has_child_content,
                            },
                        });
                    }
                }
                if let Some(raw) = extras.get(raw_index) {
                    let equation = self
                        .inner
                        .equations
                        .iter()
                        .find_map(|(at, slot, equation)| {
                            (*at == run_index && *slot == raw_index).then_some(equation)
                        });
                    let revision = self
                        .inner
                        .revisions
                        .iter()
                        .find_map(|(at, slot, revision)| {
                            (*at == run_index
                                && hyperlink_revision_index(*slot).is_none()
                                && *slot == raw_index)
                                .then_some(revision)
                        });
                    let bookmark = self.inner.bookmark_markers.iter().find(|bookmark| {
                        bookmark.run_index() == run_index && bookmark.raw_before() == raw_index
                    });
                    let empty_hyperlink =
                        self.inner.hyperlinks.iter().enumerate().find(|(_, link)| {
                            link.run_start == run_index
                                && link.run_end == run_index
                                && link.preserved_raw_before == Some(raw_index)
                        });
                    items.push(if let Some(equation) = equation {
                        ParagraphItemRef::Equation(equation)
                    } else if let Some(revision) = revision {
                        ParagraphItemRef::Revision(RevisionRef { inner: revision })
                    } else if let Some(bookmark) = bookmark {
                        if bookmark.is_start() {
                            ParagraphItemRef::BookmarkStart {
                                id: bookmark.id(),
                                name: bookmark.name(),
                                has_child_content: bookmark.has_child_content(),
                            }
                        } else {
                            ParagraphItemRef::BookmarkEnd {
                                id: bookmark.id(),
                                has_child_content: bookmark.has_child_content(),
                            }
                        }
                    } else if let Some((index, _)) = empty_hyperlink {
                        ParagraphItemRef::Hyperlink(HyperlinkRef {
                            paragraph: self.inner,
                            index,
                        })
                    } else {
                        ParagraphItemRef::UnsupportedXml(raw.as_slice())
                    });
                }
            }

            if run_index == self.inner.runs.len() {
                break;
            }
            if let Some((index, hyperlink)) =
                self.inner
                    .hyperlinks
                    .iter()
                    .enumerate()
                    .find(|(_, hyperlink)| {
                        hyperlink.run_start == run_index && hyperlink.run_end > run_index
                    })
            {
                items.push(ParagraphItemRef::Hyperlink(HyperlinkRef {
                    paragraph: self.inner,
                    index,
                }));
                run_index = hyperlink.run_end;
            } else {
                items.push(ParagraphItemRef::Run(RunRef {
                    inner: &self.inner.runs[run_index],
                }));
                run_index += 1;
            }
        }
        items.into_iter()
    }

    /// Get the number of runs in this paragraph.
    pub fn run_count(&self) -> usize {
        self.inner.accepted_run_paths().len()
    }

    /// Get an immutable accepted-view run by index.
    pub fn run(&self, index: usize) -> Option<RunRef<'_>> {
        let path = self.inner.accepted_run_paths().get(index)?.clone();
        self.inner.accepted_run(&path).map(|inner| RunRef { inner })
    }

    /// Return the recursive source path for one accepted-view run.
    #[doc(hidden)]
    pub fn run_path(&self, index: usize) -> Option<AcceptedRunPath> {
        self.inner.accepted_run_paths().get(index).cloned()
    }

    /// Get the paragraph style ID, if set.
    pub fn style_id(&self) -> Option<&str> {
        self.inner
            .properties
            .as_ref()
            .and_then(|ppr| ppr.style_id.as_deref())
    }

    /// Check if a page break is set before this paragraph.
    pub fn is_page_break_before(&self) -> bool {
        self.page_break_before_value().unwrap_or(false)
    }

    /// Get direct keep-with-next formatting without collapsing inheritance.
    pub fn keep_with_next_value(&self) -> Option<bool> {
        self.inner.properties.as_ref().and_then(|ppr| ppr.keep_next)
    }

    /// Get direct keep-together formatting without collapsing inheritance.
    pub fn keep_together_value(&self) -> Option<bool> {
        self.inner
            .properties
            .as_ref()
            .and_then(|ppr| ppr.keep_lines)
    }

    /// Get direct page-break-before formatting without collapsing inheritance.
    pub fn page_break_before_value(&self) -> Option<bool> {
        self.inner
            .properties
            .as_ref()
            .and_then(|ppr| ppr.page_break_before)
    }

    /// Get direct widow-control formatting without collapsing inheritance.
    pub fn widow_control_value(&self) -> Option<bool> {
        self.inner
            .properties
            .as_ref()
            .and_then(|ppr| ppr.widow_control)
    }

    /// Line spacing as a multiplier of single spacing (1.0, 1.5, 2.0…)
    /// when the spacing rule is "auto" (the w:line value counts 240ths of
    /// a line). Returns None when unset or when the rule is exact/atLeast
    /// point spacing.
    pub fn line_spacing_multiple(&self) -> Option<f64> {
        let ppr = self.inner.properties.as_ref()?;
        let line = ppr.line_spacing?;
        match ppr.line_rule.as_deref() {
            None | Some("auto") => Some(line.0 as f64 / 240.0),
            _ => None,
        }
    }

    /// Get exact or at-least line spacing as a length.
    pub fn line_spacing(&self) -> Option<Length> {
        let ppr = self.inner.properties.as_ref()?;
        let line = ppr.line_spacing?;
        match ppr.line_rule.as_deref() {
            Some("exact") | Some("atLeast") => Some(Length::twips(line.0)),
            _ => None,
        }
    }

    /// Hyperlink spans in this paragraph as (run_start, run_end, rel_id).
    /// Resolve rel_id to a URL with `Document::hyperlink_url`.
    pub fn hyperlink_spans(&self) -> Vec<(usize, usize, Option<&str>)> {
        self.inner
            .hyperlinks
            .iter()
            .map(|h| (h.run_start, h.run_end, h.rel_id.as_deref()))
            .collect()
    }

    /// Get list numbering as (num_id, level) if this paragraph is a list
    /// item. Resolve bullet-vs-numbered via `Document::numbering_is_bullet`.
    pub fn numbering(&self) -> Option<(u32, u32)> {
        let ppr = self.inner.properties.as_ref()?;
        Some((ppr.num_id?, ppr.num_ilvl.unwrap_or(0)))
    }

    /// Get the alignment, if set.
    pub fn alignment(&self) -> Option<Alignment> {
        self.inner
            .properties
            .as_ref()
            .and_then(|ppr| ppr.jc)
            .map(Alignment::from_st_jc)
    }

    /// Get direct space before the paragraph.
    pub fn space_before(&self) -> Option<Length> {
        self.inner
            .properties
            .as_ref()
            .and_then(|ppr| ppr.space_before)
            .map(|twips| Length::twips(twips.0))
    }

    /// Get direct space after the paragraph.
    pub fn space_after(&self) -> Option<Length> {
        self.inner
            .properties
            .as_ref()
            .and_then(|ppr| ppr.space_after)
            .map(|twips| Length::twips(twips.0))
    }

    /// Get direct left indentation.
    pub fn indent_left(&self) -> Option<Length> {
        self.inner
            .properties
            .as_ref()
            .and_then(|ppr| ppr.ind_left)
            .map(|twips| Length::twips(twips.0))
    }

    /// Get direct right indentation.
    pub fn indent_right(&self) -> Option<Length> {
        self.inner
            .properties
            .as_ref()
            .and_then(|ppr| ppr.ind_right)
            .map(|twips| Length::twips(twips.0))
    }

    /// Get signed first-line indentation, negative for hanging indentation.
    pub fn first_line_indent(&self) -> Option<Length> {
        let ppr = self.inner.properties.as_ref()?;
        if let Some(twips) = ppr.ind_first_line {
            return Some(Length::twips(twips.0));
        }
        ppr.ind_hanging
            .map(|twips| Length::twips(twips.0.saturating_neg()))
    }

    /// Get an iterator over immutable run references.
    pub fn runs(&self) -> impl Iterator<Item = RunRef<'_>> {
        self.inner.runs.iter().map(|r| RunRef { inner: r })
    }

    /// Iterate over the paragraph's ruby annotations in source order.
    pub fn rubies(&self) -> impl Iterator<Item = &CT_Ruby> {
        self.inner.rubies.iter()
    }

    /// Get one ruby annotation by index.
    pub fn ruby(&self, index: usize) -> Option<&CT_Ruby> {
        self.inner.rubies.get(index)
    }

    /// Check if paragraph has borders.
    pub fn has_borders(&self) -> bool {
        self.inner
            .properties
            .as_ref()
            .and_then(|ppr| ppr.borders.as_ref())
            .map(|b| !b.is_empty())
            .unwrap_or(false)
    }

    /// Get the direct bottom border, if present.
    pub fn bottom_border(&self) -> Option<ParagraphBorderRef<'_>> {
        self.inner
            .properties
            .as_ref()?
            .borders
            .as_ref()?
            .bottom
            .as_ref()
            .map(|inner| ParagraphBorderRef { inner })
    }

    /// Count direct paragraph border edges.
    pub fn border_count(&self) -> usize {
        let Some(borders) = self
            .inner
            .properties
            .as_ref()
            .and_then(|properties| properties.borders.as_ref())
        else {
            return 0;
        };

        [
            borders.top.as_ref(),
            borders.bottom.as_ref(),
            borders.left.as_ref(),
            borders.right.as_ref(),
            borders.between.as_ref(),
            borders.bar.as_ref(),
        ]
        .into_iter()
        .flatten()
        .count()
    }

    /// Get the number of tab stops defined.
    pub fn tab_stop_count(&self) -> usize {
        self.inner
            .properties
            .as_ref()
            .and_then(|ppr| ppr.tabs.as_ref())
            .map(|t| t.tabs.len())
            .unwrap_or(0)
    }

    /// Get the shading fill color, if set.
    pub fn shading_fill(&self) -> Option<&str> {
        self.inner
            .properties
            .as_ref()
            .and_then(|ppr| ppr.shading.as_ref())
            .and_then(|shd| shd.fill.as_deref())
    }

    /// Get the shading pattern, if set.
    pub fn shading_pattern(&self) -> Option<&str> {
        self.inner
            .properties
            .as_ref()
            .and_then(|ppr| ppr.shading.as_ref())
            .map(|shd| shd.val.as_str())
    }

    /// Get the shading foreground color, if set.
    pub fn shading_color(&self) -> Option<&str> {
        self.inner
            .properties
            .as_ref()
            .and_then(|ppr| ppr.shading.as_ref())
            .and_then(|shd| shd.color.as_deref())
    }

    /// Get direct logical leading-edge indentation.
    pub fn indent_start(&self) -> Option<Length> {
        self.inner
            .properties
            .as_ref()
            .and_then(|ppr| ppr.ind_start)
            .map(|twips| Length::twips(twips.0))
    }

    /// Get direct logical trailing-edge indentation.
    pub fn indent_end(&self) -> Option<Length> {
        self.inner
            .properties
            .as_ref()
            .and_then(|ppr| ppr.ind_end)
            .map(|twips| Length::twips(twips.0))
    }

    /// Get direct hanging indentation as a positive length.
    pub fn hanging_indent(&self) -> Option<Length> {
        self.inner
            .properties
            .as_ref()
            .and_then(|ppr| ppr.ind_hanging)
            .map(|twips| Length::twips(twips.0))
    }

    /// Get direct mirrored indentation without collapsing inheritance.
    pub fn mirror_indents_value(&self) -> Option<bool> {
        self.inner
            .properties
            .as_ref()
            .and_then(|ppr| ppr.mirror_indents)
    }

    /// Get the direct document-grid right indentation adjustment.
    pub fn adjust_right_indent_value(&self) -> Option<bool> {
        self.inner
            .properties
            .as_ref()
            .and_then(|ppr| ppr.adjust_right_ind)
    }

    /// Get direct `w:kinsoku`, East Asian line-breaking rules.
    pub fn kinsoku_value(&self) -> Option<bool> {
        self.inner.properties.as_ref().and_then(|ppr| ppr.kinsoku)
    }

    /// Get direct `w:wordWrap`, breaking a Latin word only at a break opportunity.
    pub fn word_wrap_value(&self) -> Option<bool> {
        self.inner.properties.as_ref().and_then(|ppr| ppr.word_wrap)
    }

    /// Get direct `w:overflowPunct`, letting trailing punctuation hang past the measure.
    pub fn overflow_punct_value(&self) -> Option<bool> {
        self.inner
            .properties
            .as_ref()
            .and_then(|ppr| ppr.overflow_punct)
    }

    /// Get direct `w:topLinePunct`, compressing punctuation at the start of a line.
    pub fn top_line_punct_value(&self) -> Option<bool> {
        self.inner
            .properties
            .as_ref()
            .and_then(|ppr| ppr.top_line_punct)
    }

    /// Get direct `w:autoSpaceDE`, automatic spacing between East Asian and Latin text.
    pub fn auto_space_de_value(&self) -> Option<bool> {
        self.inner
            .properties
            .as_ref()
            .and_then(|ppr| ppr.auto_space_de)
    }

    /// Get direct `w:autoSpaceDN`, automatic spacing between East Asian text and digits.
    pub fn auto_space_dn_value(&self) -> Option<bool> {
        self.inner
            .properties
            .as_ref()
            .and_then(|ppr| ppr.auto_space_dn)
    }

    /// Get direct `w:snapToGrid`, snapping line advance to the section character grid.
    pub fn snap_to_grid_value(&self) -> Option<bool> {
        self.inner
            .properties
            .as_ref()
            .and_then(|ppr| ppr.snap_to_grid)
    }

    /// Get direct automatic spacing before the paragraph.
    pub fn space_before_auto_value(&self) -> Option<bool> {
        self.inner
            .properties
            .as_ref()
            .and_then(|ppr| ppr.before_autospacing)
    }

    /// Get direct automatic spacing after the paragraph.
    pub fn space_after_auto_value(&self) -> Option<bool> {
        self.inner
            .properties
            .as_ref()
            .and_then(|ppr| ppr.after_autospacing)
    }

    /// Get direct contextual spacing without collapsing inheritance.
    pub fn contextual_spacing_value(&self) -> Option<bool> {
        self.inner
            .properties
            .as_ref()
            .and_then(|ppr| ppr.contextual_spacing)
    }

    /// Get the direct outline level.
    pub fn outline_level(&self) -> Option<u32> {
        self.inner
            .properties
            .as_ref()
            .and_then(|ppr| ppr.outline_lvl)
    }

    /// Get one direct paragraph border edge, if present.
    pub fn border(&self, edge: ParagraphBorderEdge) -> Option<ParagraphBorderRef<'_>> {
        let borders = self.inner.properties.as_ref()?.borders.as_ref()?;
        let inner = match edge {
            ParagraphBorderEdge::Top => borders.top.as_ref(),
            ParagraphBorderEdge::Bottom => borders.bottom.as_ref(),
            ParagraphBorderEdge::Left => borders.left.as_ref(),
            ParagraphBorderEdge::Right => borders.right.as_ref(),
            ParagraphBorderEdge::Between => borders.between.as_ref(),
            ParagraphBorderEdge::Bar => borders.bar.as_ref(),
        }?;
        Some(ParagraphBorderRef { inner })
    }

    /// Get one tab stop by index, or `None` when the index is out of range.
    pub fn tab_stop(&self, index: usize) -> Option<TabStopRef<'_>> {
        self.inner
            .properties
            .as_ref()?
            .tabs
            .as_ref()?
            .tabs
            .get(index)
            .map(|inner| TabStopRef { inner })
    }

    /// Get direct line-number suppression without collapsing inheritance.
    pub fn suppress_line_numbers_value(&self) -> Option<bool> {
        self.inner
            .properties
            .as_ref()
            .and_then(|ppr| ppr.suppress_line_numbers)
    }

    /// Get direct automatic-hyphenation suppression.
    pub fn suppress_auto_hyphens_value(&self) -> Option<bool> {
        self.inner
            .properties
            .as_ref()
            .and_then(|ppr| ppr.suppress_auto_hyphens)
    }

    /// Get the direct text frame placement, if present.
    pub fn frame(&self) -> Option<ParagraphFrame> {
        self.inner
            .properties
            .as_ref()?
            .frame
            .as_deref()
            .map(ParagraphFrame::from_ct)
    }

    /// Get direct frame overlap suppression.
    pub fn suppress_overlap_value(&self) -> Option<bool> {
        self.inner
            .properties
            .as_ref()
            .and_then(|ppr| ppr.suppress_overlap)
    }

    /// Get the direct text box tight wrap mode.
    ///
    /// A token outside the modeled set reads as `None`.
    pub fn textbox_tight_wrap(&self) -> Option<TextboxTightWrap> {
        self.inner
            .properties
            .as_ref()?
            .textbox_tight_wrap
            .as_deref()
            .and_then(TextboxTightWrap::from_str)
    }

    /// Get the direct paragraph base direction.
    pub fn right_to_left_value(&self) -> Option<bool> {
        self.inner.properties.as_ref().and_then(|ppr| ppr.bidi)
    }

    /// Get the direct paragraph text flow.
    ///
    /// A token outside the modeled set reads as `None`.
    pub fn text_direction(&self) -> Option<ParagraphTextDirection> {
        self.inner
            .properties
            .as_ref()?
            .text_direction
            .as_deref()
            .and_then(ParagraphTextDirection::from_str)
    }

    /// Get the direct vertical character alignment.
    ///
    /// A token outside the modeled set reads as `None`.
    pub fn text_alignment(&self) -> Option<ParagraphTextAlignment> {
        self.inner
            .properties
            .as_ref()?
            .text_alignment
            .as_deref()
            .and_then(ParagraphTextAlignment::from_str)
    }

    /// Get the web settings division this paragraph belongs to.
    pub fn div_id(&self) -> Option<u32> {
        self.inner.properties.as_ref().and_then(|ppr| ppr.div_id)
    }

    /// Get the conditional table-style regions this paragraph selects.
    pub fn conditional_formatting(&self) -> Option<TableConditionalFormatting> {
        self.inner
            .properties
            .as_ref()?
            .cnf_style
            .as_deref()
            .and_then(TableConditionalFormatting::from_value)
    }

    /// Read the formatting of the paragraph mark itself.
    pub fn mark(&self) -> ParagraphMarkRef<'_> {
        ParagraphMarkRef {
            inner: self
                .inner
                .properties
                .as_ref()
                .and_then(|ppr| ppr.rpr.as_ref()),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn hyperlink_reader_exposes_modeled_and_unmodeled_attributes() {
        let mut paragraph = CT_P::new();
        paragraph.add_run("linked");
        paragraph.hyperlinks.push(HyperlinkSpan {
            rel_id: Some("rId1".to_owned()),
            anchor: None,
            tooltip: Some("Open link".to_owned()),
            doc_location: Some("section-two".to_owned()),
            run_start: 0,
            run_end: 1,
            extra_attributes: vec![("w:history".to_owned(), "1".to_owned())],
            extra_xml: Vec::new(),
            preserved_raw_before: None,
        });
        let hyperlink = HyperlinkRef {
            paragraph: &paragraph,
            index: 0,
        };

        assert_eq!(hyperlink.tooltip(), Some("Open link"));
        assert_eq!(hyperlink.doc_location(), Some("section-two"));
        assert!(hyperlink.has_unmodeled_semantic_attributes());
    }

    #[test]
    fn bottom_border_accepts_an_explicit_content_gap() {
        let mut inner = CT_P::new();
        let mut paragraph = Paragraph { inner: &mut inner };
        paragraph.set_border_bottom_with_space(BorderStyle::Single, 14, 6, "auto");

        let paragraph = ParagraphRef { inner: &inner };
        let border = paragraph.bottom_border().expect("bottom border");
        assert_eq!(border.style(), "single");
        assert_eq!(border.size_eighths_pt(), Some(14));
        assert_eq!(border.space_points(), Some(6));
        assert_eq!(border.color(), Some("auto"));
    }
}
