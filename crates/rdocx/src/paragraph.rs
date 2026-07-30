//! Paragraph — a block-level container for runs of text.

use rdocx_oxml::borders::{CT_BorderEdge, CT_PBdr, CT_TabStop, CT_Tabs};
use rdocx_oxml::document::CT_SectPr;
use rdocx_oxml::properties::{CT_PPr, CT_Shd};
use rdocx_oxml::shared::{
    ST_Border, ST_Jc, ST_PageOrientation, ST_SectionType, ST_TabJc, ST_TabLeader,
};
use rdocx_oxml::text::{CT_P, CT_R};
use rdocx_oxml::units::Twips;

use crate::Length;
use crate::run::{Run, RunRef};

/// Paragraph alignment options.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Alignment {
    Left,
    Center,
    Right,
    Justify,
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
}

/// A mutable reference to a paragraph in a document.
pub struct Paragraph<'a> {
    pub(crate) inner: &'a mut CT_P,
}

impl<'a> Paragraph<'a> {
    /// Get the combined text of all runs.
    pub fn text(&self) -> String {
        self.inner.text()
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

    /// Get an iterator over immutable run references.
    pub fn runs(&self) -> impl Iterator<Item = RunRef<'_>> {
        self.inner.runs.iter().map(|r| RunRef { inner: r })
    }

    /// Set the paragraph alignment.
    pub fn alignment(mut self, align: Alignment) -> Self {
        self.set_alignment(align);
        self
    }

    /// Set the paragraph alignment in place.
    pub fn set_alignment(&mut self, align: Alignment) {
        self.ensure_ppr().jc = Some(align.to_st_jc());
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

    /// Set space before the paragraph.
    pub fn space_before(mut self, length: Length) -> Self {
        self.set_space_before(length);
        self
    }

    /// Set space before the paragraph in place.
    pub fn set_space_before(&mut self, length: Length) {
        self.ensure_ppr().space_before = Some(length.as_twips());
    }

    /// Set space after the paragraph.
    pub fn space_after(mut self, length: Length) -> Self {
        self.set_space_after(length);
        self
    }

    /// Set space after the paragraph in place.
    pub fn set_space_after(&mut self, length: Length) {
        self.ensure_ppr().space_after = Some(length.as_twips());
    }

    /// Set left indentation.
    pub fn indent_left(mut self, length: Length) -> Self {
        self.set_indent_left(length);
        self
    }

    /// Set left indentation in place.
    pub fn set_indent_left(&mut self, length: Length) {
        self.ensure_ppr().ind_left = Some(length.as_twips());
    }

    /// Set right indentation.
    pub fn indent_right(mut self, length: Length) -> Self {
        self.set_indent_right(length);
        self
    }

    /// Set right indentation in place.
    pub fn set_indent_right(&mut self, length: Length) {
        self.ensure_ppr().ind_right = Some(length.as_twips());
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
        self.ensure_ppr().keep_next = Some(val);
    }

    /// Set keep lines together.
    pub fn keep_together(mut self, val: bool) -> Self {
        self.set_keep_together(val);
        self
    }

    /// Set keep lines together in place.
    pub fn set_keep_together(&mut self, val: bool) {
        self.ensure_ppr().keep_lines = Some(val);
    }

    /// Set page break before.
    pub fn page_break_before(mut self, val: bool) -> Self {
        self.set_page_break_before(val);
        self
    }

    /// Set page break before in place.
    pub fn set_page_break_before(&mut self, val: bool) {
        self.ensure_ppr().page_break_before = Some(val);
    }

    /// Set widow/orphan control.
    pub fn widow_control(mut self, val: bool) -> Self {
        self.set_widow_control(val);
        self
    }

    /// Set widow/orphan control in place.
    pub fn set_widow_control(&mut self, val: bool) {
        self.ensure_ppr().widow_control = Some(val);
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
        self.ensure_ppr().shading = Some(CT_Shd {
            val: "clear".to_string(),
            color: Some("auto".to_string()),
            fill: Some(fill_color.to_string()),
        });
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
        };
        self.ensure_ppr().borders = Some(CT_PBdr {
            top: Some(edge.clone()),
            bottom: Some(edge.clone()),
            left: Some(edge.clone()),
            right: Some(edge),
            between: None,
            bar: None,
        });
    }

    /// Add a bottom border only.
    pub fn border_bottom(mut self, style: BorderStyle, size_eighths_pt: u32, color: &str) -> Self {
        self.set_border_bottom(style, size_eighths_pt, color);
        self
    }

    /// Add a bottom border in place.
    pub fn set_border_bottom(&mut self, style: BorderStyle, size_eighths_pt: u32, color: &str) {
        let edge = CT_BorderEdge {
            val: style.to_st(),
            sz: Some(size_eighths_pt),
            space: Some(1),
            color: Some(color.to_string()),
        };
        let borders = self
            .ensure_ppr()
            .borders
            .get_or_insert_with(CT_PBdr::default);
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
        });
    }

    /// Set outline level (0–8, used for TOC generation).
    pub fn outline_level(mut self, level: u32) -> Self {
        self.set_outline_level(level);
        self
    }

    /// Set outline level in place.
    pub fn set_outline_level(&mut self, level: u32) {
        self.ensure_ppr().outline_lvl = Some(level);
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
        self.inner.text()
    }

    /// Get the paragraph style ID, if set.
    pub fn style_id(&self) -> Option<&str> {
        self.inner
            .properties
            .as_ref()
            .and_then(|ppr| ppr.style_id.as_deref())
    }

    /// Get the alignment, if set.
    /// Check if a page break is set before this paragraph.
    pub fn is_page_break_before(&self) -> bool {
        self.inner
            .properties
            .as_ref()
            .and_then(|ppr| ppr.page_break_before)
            .unwrap_or(false)
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

    pub fn alignment(&self) -> Option<Alignment> {
        self.inner
            .properties
            .as_ref()
            .and_then(|ppr| ppr.jc)
            .map(Alignment::from_st_jc)
    }

    /// Get an iterator over immutable run references.
    pub fn runs(&self) -> impl Iterator<Item = RunRef<'_>> {
        self.inner.runs.iter().map(|r| RunRef { inner: r })
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
}
