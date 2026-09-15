//! Run — a contiguous stretch of text with uniform formatting.

use rdocx_oxml::drawing::CT_Drawing;
use rdocx_oxml::properties::{CT_RPr, CT_Shd};
use rdocx_oxml::shared::ST_Underline;
use rdocx_oxml::text::{BreakType, CT_R, CT_Text, Field, RunContent};
use rdocx_oxml::units::{HalfPoint, Twips};

use crate::Length;

/// A break embedded in a run.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BreakKind {
    /// A line break within the current paragraph.
    Line,
    /// A page break.
    Page,
    /// A column break.
    Column,
}

/// Semantic kind of drawing exposed by the reader facade.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DrawingKind {
    /// A DrawingML picture with an image relationship.
    Image,
    /// An anchored DrawingML shape.
    Shape,
    /// Any other drawing construct.
    Other,
}

/// How an image relationship obtains its content.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DrawingRelationshipKind {
    /// The drawing refers to an image part within the package.
    Embedded,
    /// The drawing refers to an external image relationship.
    Linked,
}

/// The source representation of a Word field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FieldKind {
    /// A `w:fldSimple` element.
    Simple,
    /// A `w:fldChar` begin/separate/end sequence.
    Complex,
}

/// Resolved or direct run properties returned by the reader API.
pub type RunProperties = CT_RPr;

/// An immutable drawing embedded in a run.
#[derive(Debug, Clone, Copy)]
pub struct DrawingRef<'a> {
    inner: &'a CT_Drawing,
}

impl DrawingRef<'_> {
    /// Semantic content kind.
    pub fn kind(&self) -> DrawingKind {
        if self.relationship_id().is_some() {
            DrawingKind::Image
        } else if self
            .inner
            .anchor
            .as_ref()
            .is_some_and(|anchor| anchor.shape.is_some())
        {
            DrawingKind::Shape
        } else {
            DrawingKind::Other
        }
    }

    /// Whether this drawing is inline with the surrounding text.
    pub fn is_inline(&self) -> bool {
        self.inner.inline.is_some()
    }

    /// Whether this drawing is floating or anchored.
    pub fn is_anchor(&self) -> bool {
        self.inner.anchor.is_some()
    }

    /// The relationship ID for the drawing's embedded or linked image, when present.
    pub fn relationship_id(&self) -> Option<&str> {
        self.inner
            .inline
            .as_ref()
            .and_then(|inline| {
                (!inline.embed_id.is_empty())
                    .then_some(inline.embed_id.as_str())
                    .or(inline.link_id.as_deref())
            })
            .or_else(|| {
                self.inner.anchor.as_ref().and_then(|anchor| {
                    (!anchor.embed_id.is_empty())
                        .then_some(anchor.embed_id.as_str())
                        .or(anchor.link_id.as_deref())
                })
            })
    }

    /// Whether the image relationship is embedded in the package or linked.
    pub fn relationship_kind(&self) -> Option<DrawingRelationshipKind> {
        self.inner
            .inline
            .as_ref()
            .map(|inline| inline.embed_id.is_empty() && inline.link_id.is_some())
            .or_else(|| {
                self.inner
                    .anchor
                    .as_ref()
                    .map(|anchor| anchor.embed_id.is_empty() && anchor.link_id.is_some())
            })
            .filter(|_| self.relationship_id().is_some())
            .map(|linked| {
                if linked {
                    DrawingRelationshipKind::Linked
                } else {
                    DrawingRelationshipKind::Embedded
                }
            })
    }

    /// The drawing description, commonly used as image alternative text.
    pub fn description(&self) -> Option<&str> {
        self.inner
            .inline
            .as_ref()
            .and_then(|inline| inline.description.as_deref())
            .or_else(|| {
                self.inner
                    .anchor
                    .as_ref()
                    .and_then(|anchor| anchor.description.as_deref())
            })
    }

    /// The drawing name from its non-visual properties.
    pub fn name(&self) -> Option<&str> {
        self.inner
            .inline
            .as_ref()
            .and_then(|inline| inline.name.as_deref())
            .or_else(|| {
                self.inner
                    .anchor
                    .as_ref()
                    .and_then(|anchor| anchor.name.as_deref())
            })
    }

    /// The drawing width.
    pub fn width(&self) -> Option<Length> {
        self.inner
            .inline
            .as_ref()
            .map(|inline| Length::emu(inline.extent_cx.0))
            .or_else(|| {
                self.inner
                    .anchor
                    .as_ref()
                    .map(|anchor| Length::emu(anchor.extent_cx.0))
            })
    }

    /// The drawing height.
    pub fn height(&self) -> Option<Length> {
        self.inner
            .inline
            .as_ref()
            .map(|inline| Length::emu(inline.extent_cy.0))
            .or_else(|| {
                self.inner
                    .anchor
                    .as_ref()
                    .map(|anchor| Length::emu(anchor.extent_cy.0))
            })
    }
}

/// An immutable field embedded in a run.
#[derive(Debug, Clone, Copy)]
pub struct FieldRef<'a> {
    inner: &'a Field,
}

/// One cached display segment retained from a complex field result.
#[derive(Debug, Clone, Copy)]
pub struct FieldDisplaySegmentRef<'a> {
    text: &'a str,
    properties: Option<&'a RunProperties>,
}

impl FieldDisplaySegmentRef<'_> {
    /// The visible text retained from the result run.
    pub fn text(&self) -> &str {
        self.text
    }

    /// Direct run properties retained from the result run.
    pub fn properties(&self) -> Option<&RunProperties> {
        self.properties
    }
}

/// An immutable legacy VML horizontal rule retained by a run.
#[derive(Debug, Clone, Copy)]
pub struct LegacyHorizontalRuleRef<'a> {
    raw_xml: &'a [u8],
}

impl LegacyHorizontalRuleRef<'_> {
    /// The exact preserved `<w:pict>` subtree.
    pub fn raw_xml(&self) -> &[u8] {
        self.raw_xml
    }
}

impl FieldRef<'_> {
    /// Whether the field was represented as a simple element or complex run sequence.
    pub fn kind(&self) -> FieldKind {
        if self.inner.is_complex() {
            FieldKind::Complex
        } else {
            FieldKind::Simple
        }
    }

    /// The retained field instruction text.
    pub fn instruction(&self) -> &str {
        &self.inner.instruction.raw
    }

    /// The parsed field name.
    pub fn name(&self) -> &str {
        &self.inner.instruction.name
    }

    /// The cached display result stored in the document.
    pub fn cached_result(&self) -> &str {
        &self.inner.cached_result
    }

    /// The producer's update marker, when specified.
    pub fn dirty(&self) -> Option<bool> {
        self.inner.dirty
    }

    /// Whether the retained field source carries semantic attributes outside
    /// the modeled reader projection.
    pub fn has_unmodeled_semantic_attributes(&self) -> bool {
        self.inner.has_unmodeled_semantic_attributes()
    }

    /// Cached display segments retained from the result runs of a complex field.
    pub fn cached_display_segments(&self) -> Vec<FieldDisplaySegmentRef<'_>> {
        self.inner
            .cached_display_segments()
            .into_iter()
            .map(|(text, properties)| FieldDisplaySegmentRef { text, properties })
            .collect()
    }
}

/// One direct child of a run, in source order.
#[derive(Debug, Clone, Copy)]
#[non_exhaustive]
pub enum RunItemRef<'a> {
    /// Literal text.
    Text(&'a str),
    /// Text in a deleted revision.
    DeletedText(&'a str),
    /// A tab character.
    Tab,
    /// A line, page, or column break.
    Break(BreakKind),
    /// An inline or anchored drawing.
    Drawing(DrawingRef<'a>),
    /// A simple or complex Word field.
    Field(FieldRef<'a>),
    /// A footnote reference ID.
    FootnoteReference(i32),
    /// An endnote reference ID.
    EndnoteReference(i32),
    /// A comment reference ID.
    CommentReference(i32),
    /// An unambiguous legacy VML horizontal rule.
    LegacyHorizontalRule(LegacyHorizontalRuleRef<'a>),
    /// A preserved run child that rdocx does not model.
    UnsupportedXml(&'a [u8]),
}

fn classify_raw_run_item(raw_xml: &[u8], encoded_position: Option<usize>) -> RunItemRef<'_> {
    if encoded_position.is_some_and(CT_R::raw_child_is_legacy_horizontal_rule) {
        RunItemRef::LegacyHorizontalRule(LegacyHorizontalRuleRef { raw_xml })
    } else {
        RunItemRef::UnsupportedXml(raw_xml)
    }
}

/// Underline style for runs.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum UnderlineStyle {
    None,
    Single,
    Double,
    Thick,
    Dotted,
    Dash,
    Wave,
    Words,
}

impl UnderlineStyle {
    fn to_code(self) -> i32 {
        match self {
            Self::None => 0,
            Self::Single => 1,
            Self::Words => 2,
            Self::Double => 3,
            Self::Dotted => 4,
            Self::Thick => 6,
            Self::Dash => 7,
            Self::Wave => 11,
        }
    }
}

fn underline_from_code(code: i32) -> Option<ST_Underline> {
    match code {
        0 => Some(ST_Underline::None),
        1 => Some(ST_Underline::Single),
        2 => Some(ST_Underline::Words),
        3 => Some(ST_Underline::Double),
        4 => Some(ST_Underline::Dotted),
        6 => Some(ST_Underline::Thick),
        7 => Some(ST_Underline::Dash),
        9 => Some(ST_Underline::DotDash),
        10 => Some(ST_Underline::DotDotDash),
        11 => Some(ST_Underline::Wave),
        _ => None,
    }
}

fn underline_to_code(value: ST_Underline) -> i32 {
    match value {
        ST_Underline::None => 0,
        ST_Underline::Single => 1,
        ST_Underline::Words => 2,
        ST_Underline::Double => 3,
        ST_Underline::Dotted => 4,
        ST_Underline::Thick => 6,
        ST_Underline::Dash => 7,
        ST_Underline::DotDash => 9,
        ST_Underline::DotDotDash => 10,
        ST_Underline::Wave => 11,
    }
}

/// A run of text within a paragraph.
///
/// All text in a run shares the same formatting (font, size, bold, etc.).
pub struct Run<'a> {
    pub(crate) inner: &'a mut CT_R,
}

impl<'a> Run<'a> {
    /// Get the text content of this run.
    pub fn text(&self) -> String {
        self.inner.text()
    }

    /// Set the text content, replacing all existing content.
    pub fn set_text(&mut self, text: &str) {
        self.inner
            .replace_content(vec![RunContent::Text(CT_Text::new(text))]);
    }

    /// Add text to this run.
    pub fn add_text(&mut self, text: &str) {
        self.inner
            .append_content(RunContent::Text(CT_Text::new(text)));
    }

    /// Set bold formatting.
    pub fn bold(mut self, val: bool) -> Self {
        self.set_bold(val);
        self
    }

    /// Set bold formatting in place.
    pub fn set_bold(&mut self, val: bool) {
        self.set_bold_value(Some(val));
    }

    /// Set or clear direct bold formatting in place.
    pub fn set_bold_value(&mut self, val: Option<bool>) {
        if val.is_none() && self.inner.properties.is_none() {
            return;
        }
        let rpr = self.ensure_rpr();
        rpr.bold = val;
        rpr.bold_cs = val;
    }

    /// Set italic formatting.
    pub fn italic(mut self, val: bool) -> Self {
        self.set_italic(val);
        self
    }

    /// Set italic formatting in place.
    pub fn set_italic(&mut self, val: bool) {
        self.set_italic_value(Some(val));
    }

    /// Set or clear direct italic formatting in place.
    pub fn set_italic_value(&mut self, val: Option<bool>) {
        if val.is_none() && self.inner.properties.is_none() {
            return;
        }
        let rpr = self.ensure_rpr();
        rpr.italic = val;
        rpr.italic_cs = val;
    }

    /// Set underline formatting (simple on/off).
    pub fn underline(mut self, val: bool) -> Self {
        self.set_underline(val);
        self
    }

    /// Set underline formatting in place.
    pub fn set_underline(&mut self, val: bool) {
        self.set_underline_style_value(Some(if val {
            UnderlineStyle::Single
        } else {
            UnderlineStyle::None
        }));
    }

    /// Set underline with a specific style.
    pub fn underline_style(mut self, style: UnderlineStyle) -> Self {
        self.set_underline_style(style);
        self
    }

    /// Set an underline style in place.
    pub fn set_underline_style(&mut self, style: UnderlineStyle) {
        self.set_underline_style_value(Some(style));
    }

    /// Set or clear direct underline formatting in place.
    pub fn set_underline_style_value(&mut self, style: Option<UnderlineStyle>) {
        let applied = self.set_underline_code_value(style.map(UnderlineStyle::to_code));
        debug_assert!(applied);
    }

    /// Set or clear a direct underline code used by language bindings.
    ///
    /// Returns false without mutation when `code` is not in the bounded public
    /// Python underline inventory.
    pub fn set_underline_code_value(&mut self, code: Option<i32>) -> bool {
        let underline = match code {
            Some(code) => match underline_from_code(code) {
                Some(underline) => Some(underline),
                None => return false,
            },
            None => None,
        };
        if underline.is_none() && self.inner.properties.is_none() {
            return true;
        }
        self.ensure_rpr().underline = underline;
        true
    }

    /// Set font size in points.
    pub fn size(mut self, pt: f64) -> Self {
        self.set_size(pt);
        self
    }

    /// Set font size in place.
    pub fn set_size(&mut self, pt: f64) {
        self.set_size_value(Some(pt));
    }

    /// Set or clear the direct font size in points.
    pub fn set_size_value(&mut self, pt: Option<f64>) {
        if pt.is_none() && self.inner.properties.is_none() {
            return;
        }
        let hp = pt.map(HalfPoint::from_pt);
        let rpr = self.ensure_rpr();
        rpr.sz = hp;
        rpr.sz_cs = hp;
    }

    /// Set the font name.
    pub fn font(mut self, name: &str) -> Self {
        self.set_font(name);
        self
    }

    /// Set the font name in place.
    pub fn set_font(&mut self, name: &str) {
        self.set_font_value(Some(name));
    }

    /// Set or clear the direct font name.
    pub fn set_font_value(&mut self, name: Option<&str>) {
        if name.is_none() && self.inner.properties.is_none() {
            return;
        }
        let rpr = self.ensure_rpr();
        rpr.font_ascii = name.map(str::to_owned);
        rpr.font_hansi = name.map(str::to_owned);
        rpr.font_east_asia = name.map(str::to_owned);
        rpr.font_cs = name.map(str::to_owned);
    }

    /// Set the run language used by language-aware text layout.
    pub fn language(mut self, language: &str) -> Self {
        self.set_language(language);
        self
    }

    /// Set the run language in place.
    pub fn set_language(&mut self, language: &str) {
        self.set_language_value(Some(language));
    }

    /// Set or clear the direct Latin and high-ANSI run language.
    pub fn set_language_value(&mut self, language: Option<&str>) {
        if language.is_none() && self.inner.properties.is_none() {
            return;
        }
        self.ensure_rpr().language = language.map(str::to_owned);
    }

    /// Set text color as a hex string (e.g., "FF0000" for red).
    pub fn color(mut self, hex: &str) -> Self {
        self.set_color(hex);
        self
    }

    /// Set text color in place.
    pub fn set_color(&mut self, hex: &str) {
        self.set_color_value(Some(hex));
    }

    /// Set or clear the direct text color.
    pub fn set_color_value(&mut self, hex: Option<&str>) {
        if hex.is_none() && self.inner.properties.is_none() {
            return;
        }
        self.ensure_rpr().color = hex.map(str::to_owned);
    }

    /// Set highlight color as a hex fill value.
    pub fn highlight(mut self, color: &str) -> Self {
        self.set_highlight(color);
        self
    }

    /// Set highlight color in place.
    pub fn set_highlight(&mut self, color: &str) {
        self.ensure_rpr().shading = Some(CT_Shd {
            val: "clear".to_string(),
            color: Some("auto".to_string()),
            fill: Some(color.to_string()),
        });
    }

    /// Set or clear the highlight in place.
    ///
    /// A value is written as a shading fill, like [`Run::set_highlight`].
    /// Clearing removes that shading and any `w:highlight` keyword, the two
    /// forms [`RunRef::highlight`] reads.
    pub fn set_highlight_value(&mut self, color: Option<&str>) {
        match color {
            Some(color) => self.set_highlight(color),
            None => {
                if let Some(properties) = self.inner.properties.as_mut() {
                    properties.shading = None;
                    properties.highlight = None;
                }
            }
        }
    }

    /// Set strikethrough formatting.
    pub fn strike(mut self, val: bool) -> Self {
        self.set_strike(val);
        self
    }

    /// Set strikethrough formatting in place.
    pub fn set_strike(&mut self, val: bool) {
        self.set_strike_value(Some(val));
    }

    /// Set or clear direct strikethrough formatting in place.
    pub fn set_strike_value(&mut self, val: Option<bool>) {
        if val.is_none() && self.inner.properties.is_none() {
            return;
        }
        self.ensure_rpr().strike = val;
    }

    /// Set double strikethrough.
    pub fn double_strike(mut self, val: bool) -> Self {
        self.set_double_strike(val);
        self
    }

    /// Set double strikethrough in place.
    pub fn set_double_strike(&mut self, val: bool) {
        self.ensure_rpr().dstrike = Some(val);
    }

    /// Set all caps.
    pub fn all_caps(mut self, val: bool) -> Self {
        self.set_all_caps(val);
        self
    }

    /// Set all caps in place.
    pub fn set_all_caps(&mut self, val: bool) {
        self.ensure_rpr().caps = Some(val);
    }

    /// Set small caps.
    pub fn small_caps(mut self, val: bool) -> Self {
        self.set_small_caps(val);
        self
    }

    /// Set small caps in place.
    pub fn set_small_caps(&mut self, val: bool) {
        self.ensure_rpr().small_caps = Some(val);
    }

    /// Set superscript.
    pub fn superscript(mut self) -> Self {
        self.set_superscript();
        self
    }

    /// Set superscript in place.
    pub fn set_superscript(&mut self) {
        self.ensure_rpr().vert_align = Some("superscript".to_string());
    }

    /// Set subscript.
    pub fn subscript(mut self) -> Self {
        self.set_subscript();
        self
    }

    /// Set subscript in place.
    pub fn set_subscript(&mut self) {
        self.ensure_rpr().vert_align = Some("subscript".to_string());
    }

    /// Set character spacing (positive = expanded, negative = condensed).
    pub fn character_spacing(mut self, spacing: Length) -> Self {
        self.set_character_spacing(spacing);
        self
    }

    /// Set character spacing in place.
    pub fn set_character_spacing(&mut self, spacing: Length) {
        self.ensure_rpr().spacing = Some(spacing.as_twips());
    }

    /// Set character width scale in percent (100 = normal).
    pub fn width_scale(mut self, percent: u32) -> Self {
        self.set_width_scale(percent);
        self
    }

    /// Set character width scale in place.
    pub fn set_width_scale(&mut self, percent: u32) {
        self.ensure_rpr().width_scale = Some(percent);
    }

    /// Set text position (positive = raised, negative = lowered) in half-points.
    pub fn position(mut self, half_points: i32) -> Self {
        self.set_position(half_points);
        self
    }

    /// Set text position in place.
    pub fn set_position(&mut self, half_points: i32) {
        self.ensure_rpr().position = Some(half_points);
    }

    /// Set hidden/vanish text.
    pub fn hidden(mut self, val: bool) -> Self {
        self.set_hidden(val);
        self
    }

    /// Set hidden/vanish text in place.
    pub fn set_hidden(&mut self, val: bool) {
        self.ensure_rpr().vanish = Some(val);
    }

    /// Set the character style by ID.
    pub fn style(mut self, style_id: &str) -> Self {
        self.set_style(style_id);
        self
    }

    /// Set the character style by ID in place.
    pub fn set_style(&mut self, style_id: &str) {
        self.ensure_rpr().style_id = Some(style_id.to_string());
    }

    /// Set or clear the character style ID in place.
    pub fn set_style_value(&mut self, style_id: Option<&str>) {
        if style_id.is_none() && self.inner.properties.is_none() {
            return;
        }
        self.ensure_rpr().style_id = style_id.map(str::to_owned);
    }

    fn ensure_rpr(&mut self) -> &mut CT_RPr {
        self.inner.ensure_properties()
    }
}

/// An immutable reference to a run.
pub struct RunRef<'a> {
    pub(crate) inner: &'a CT_R,
}

impl<'a> RunRef<'a> {
    /// Get the text content of this run.
    pub fn text(&self) -> String {
        self.inner.text()
    }

    /// Iterate over direct run items in source order.
    pub fn items(&self) -> impl Iterator<Item = RunItemRef<'_>> {
        let property_boundary = usize::from(self.inner.properties.is_some());
        let mut items = Vec::with_capacity(self.inner.content.len() + self.inner.extra_xml.len());
        let ordered_raw = self.inner.extra_xml_positions.len() == self.inner.extra_xml.len();
        if ordered_raw && property_boundary > 0 {
            items.extend(
                self.inner
                    .extra_xml_positions
                    .iter()
                    .zip(&self.inner.extra_xml)
                    .filter(|(position, _)| CT_R::raw_child_position(**position) == 0)
                    .map(|(position, raw)| classify_raw_run_item(raw, Some(*position))),
            );
        }
        for index in 0..=self.inner.content.len() {
            let boundary = property_boundary + index;
            if ordered_raw {
                items.extend(
                    self.inner
                        .extra_xml_positions
                        .iter()
                        .zip(&self.inner.extra_xml)
                        .filter(|(position, _)| CT_R::raw_child_position(**position) == boundary)
                        .map(|(position, raw)| classify_raw_run_item(raw, Some(*position))),
                );
            }
            if let Some(content) = self.inner.content.get(index) {
                items.push(match content {
                    RunContent::Text(text) => RunItemRef::Text(&text.text),
                    RunContent::DeletedText(text) => RunItemRef::DeletedText(&text.text),
                    RunContent::Tab => RunItemRef::Tab,
                    RunContent::Break(kind) => RunItemRef::Break(match kind {
                        BreakType::Line => BreakKind::Line,
                        BreakType::Page => BreakKind::Page,
                        BreakType::Column => BreakKind::Column,
                    }),
                    RunContent::Drawing(drawing) => {
                        RunItemRef::Drawing(DrawingRef { inner: drawing })
                    }
                    RunContent::Field(field) => RunItemRef::Field(FieldRef { inner: field }),
                    RunContent::FootnoteRef { id } => RunItemRef::FootnoteReference(*id),
                    RunContent::EndnoteRef { id } => RunItemRef::EndnoteReference(*id),
                    RunContent::CommentReference { id, .. } => RunItemRef::CommentReference(*id),
                });
            }
        }
        if !ordered_raw {
            items.extend(
                self.inner
                    .extra_xml
                    .iter()
                    .map(|raw| classify_raw_run_item(raw, None)),
            );
        }
        items.into_iter()
    }

    /// The footnote id referenced by this run, if it holds a
    /// `<w:footnoteReference/>`.
    pub fn footnote_id(&self) -> Option<i32> {
        use rdocx_oxml::text::RunContent;
        self.inner.content.iter().find_map(|c| match c {
            RunContent::FootnoteRef { id } => Some(*id),
            _ => None,
        })
    }

    /// Check if bold.
    pub fn is_bold(&self) -> bool {
        self.bold_value().unwrap_or(false)
    }

    /// Get direct bold formatting without collapsing inheritance.
    pub fn bold_value(&self) -> Option<bool> {
        self.inner.properties.as_ref().and_then(|rpr| rpr.bold)
    }

    /// Check if italic.
    pub fn is_italic(&self) -> bool {
        self.italic_value().unwrap_or(false)
    }

    /// Get direct italic formatting without collapsing inheritance.
    pub fn italic_value(&self) -> Option<bool> {
        self.inner.properties.as_ref().and_then(|rpr| rpr.italic)
    }

    /// Check if strikethrough.
    pub fn is_strike(&self) -> bool {
        self.strike_value().unwrap_or(false)
    }

    /// Get direct strikethrough formatting without collapsing inheritance.
    pub fn strike_value(&self) -> Option<bool> {
        self.inner.properties.as_ref().and_then(|rpr| rpr.strike)
    }

    /// Check if underlined (any underline style other than none).
    pub fn is_underline(&self) -> bool {
        self.underline_code_value().is_some_and(|code| code != 0)
    }

    /// Get the direct underline code used by language bindings.
    pub fn underline_code_value(&self) -> Option<i32> {
        self.inner
            .properties
            .as_ref()
            .and_then(|rpr| rpr.underline)
            .map(underline_to_code)
    }

    /// Get font size in points, if set.
    pub fn size(&self) -> Option<f64> {
        self.inner
            .properties
            .as_ref()
            .and_then(|rpr| rpr.sz)
            .map(|hp| hp.to_pt())
    }

    /// Get font name, if set.
    pub fn font_name(&self) -> Option<&str> {
        self.inner
            .properties
            .as_ref()
            .and_then(|rpr| rpr.font_ascii.as_deref())
    }

    /// Get the direct Latin and high-ANSI run language, if set.
    pub fn language(&self) -> Option<&str> {
        self.inner
            .properties
            .as_ref()
            .and_then(|rpr| rpr.language.as_deref())
    }

    /// Get text color, if set.
    pub fn color(&self) -> Option<&str> {
        self.inner
            .properties
            .as_ref()
            .and_then(|rpr| rpr.color.as_deref())
    }

    /// Get character spacing in twips, if set.
    pub fn character_spacing(&self) -> Option<Twips> {
        self.inner.properties.as_ref().and_then(|rpr| rpr.spacing)
    }

    /// Get the highlight color, if set: either the `w:highlight` keyword
    /// (e.g. "yellow") or the shading fill value the `highlight()` builder
    /// writes — OOXML has two mechanisms for highlighted text.
    pub fn highlight(&self) -> Option<String> {
        let rpr = self.inner.properties.as_ref()?;
        if let Some(h) = rpr.highlight {
            return Some(h.to_str().to_string());
        }
        rpr.shading.as_ref().and_then(|sh| sh.fill.clone())
    }

    /// If this run contains an inline image, return (rel_id, alt text).
    pub fn inline_image(&self) -> Option<(&str, Option<&str>)> {
        use rdocx_oxml::text::RunContent;
        for c in &self.inner.content {
            if let RunContent::Drawing(d) = c
                && let Some(inline) = &d.inline
            {
                return Some((inline.embed_id.as_str(), inline.description.as_deref()));
            }
        }
        None
    }

    /// Get raised/lowered text position in half-points, if set.
    /// (LibreOffice encodes super/subscript this way on HTML import.)
    pub fn position(&self) -> Option<i32> {
        self.inner.properties.as_ref().and_then(|rpr| rpr.position)
    }

    /// Get vertical alignment (superscript/subscript), if set.
    pub fn vert_align(&self) -> Option<&str> {
        self.inner
            .properties
            .as_ref()
            .and_then(|rpr| rpr.vert_align.as_deref())
    }

    /// Get the character style ID, if set.
    pub fn style_id(&self) -> Option<&str> {
        self.inner
            .properties
            .as_ref()
            .and_then(|rpr| rpr.style_id.as_deref())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use rdocx_oxml::CT_Document;
    use rdocx_oxml::document::BodyContent;

    #[test]
    fn ct_r_public_struct_literal_keeps_its_existing_shape() {
        let run = CT_R {
            properties: None,
            content: Vec::new(),
            extra_xml: Vec::new(),
            extra_xml_positions: Vec::new(),
            alt_drawings: Vec::new(),
        };

        assert!(run.content.is_empty());
    }

    fn run_with_raw(raw: &[u8]) -> CT_R {
        let raw = std::str::from_utf8(raw).unwrap();
        run_with_inner_xml(raw)
    }

    fn run_with_inner_xml(inner_xml: &str) -> CT_R {
        let xml = format!(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r>{inner_xml}</w:r></w:p></w:body></w:document>"#
        );
        let mut document = CT_Document::from_xml(xml.as_bytes()).unwrap();
        let BodyContent::Paragraph(mut paragraph) = document.body.content.remove(0) else {
            panic!("expected paragraph");
        };
        paragraph.runs.remove(0)
    }

    fn run_from_paragraph_inner_xml(inner_xml: &str) -> CT_R {
        let xml = format!(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>{inner_xml}</w:p></w:body></w:document>"#
        );
        let mut document = CT_Document::from_xml(xml.as_bytes()).unwrap();
        let BodyContent::Paragraph(mut paragraph) = document.body.content.remove(0) else {
            panic!("expected paragraph");
        };
        paragraph.runs.remove(0)
    }

    #[test]
    fn drawing_reader_classifies_image_relationships_and_shapes() {
        let linked_run = run_with_inner_xml(concat!(
            r#"<w:drawing xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" "#,
            r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" "#,
            r#"xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" "#,
            r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
            r#"<wp:inline><wp:extent cx="10" cy="20"/><wp:docPr id="1" name="Linked"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:link="rId7"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>"#,
        ));
        let linked_run = RunRef { inner: &linked_run };
        let linked_items = linked_run.items().collect::<Vec<_>>();
        let [RunItemRef::Drawing(linked)] = linked_items.as_slice() else {
            panic!("expected one linked drawing");
        };
        assert_eq!(linked.kind(), DrawingKind::Image);
        assert_eq!(linked.relationship_id(), Some("rId7"));
        assert_eq!(
            linked.relationship_kind(),
            Some(DrawingRelationshipKind::Linked)
        );

        let embedded = CT_Drawing::inline(rdocx_oxml::drawing::CT_Inline::new("rId8", 10, 20));
        let embedded = DrawingRef { inner: &embedded };
        assert_eq!(embedded.kind(), DrawingKind::Image);
        assert_eq!(
            embedded.relationship_kind(),
            Some(DrawingRelationshipKind::Embedded)
        );

        let mut anchor = rdocx_oxml::drawing::CT_Anchor::background("", 10, 20);
        anchor.shape = Some(rdocx_oxml::drawing::CT_Shape::default());
        let shape = CT_Drawing::anchor(anchor);
        assert_eq!(DrawingRef { inner: &shape }.kind(), DrawingKind::Shape);

        let filled_shape = run_with_inner_xml(concat!(
            r#"<w:drawing xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" "#,
            r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" "#,
            r#"xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" "#,
            r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
            r#"<wp:anchor><wp:docPr id="1" name="Shape"/><wps:wsp><wps:spPr><a:blipFill><a:blip r:embed="rIdFill"/></a:blipFill></wps:spPr></wps:wsp></wp:anchor></w:drawing>"#,
        ));
        let filled_shape = DrawingRef {
            inner: filled_shape
                .content
                .iter()
                .find_map(|content| match content {
                    RunContent::Drawing(drawing) => Some(drawing),
                    _ => None,
                })
                .expect("filled shape drawing"),
        };
        assert_eq!(filled_shape.kind(), DrawingKind::Shape);
        assert_eq!(filled_shape.relationship_id(), None);

        let chart = CT_Drawing::inline(rdocx_oxml::drawing::CT_Inline::new_chart("rId9", 10, 20));
        let chart = DrawingRef { inner: &chart };
        assert_eq!(chart.kind(), DrawingKind::Other);
        assert_eq!(chart.relationship_kind(), None);
    }

    #[test]
    fn legacy_horizontal_rule_classification_is_namespace_aware() {
        let cases = [
            br#"<w:pict xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><v:rect o:hr="t"/></w:pict>"#.as_slice(),
            br#"<word:pict xmlns:word="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:shape="urn:schemas-microsoft-com:vml" xmlns:office="urn:schemas-microsoft-com:office:office"><shape:rect office:hr="true"></shape:rect></word:pict>"#.as_slice(),
            br#"<pict xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:o="urn:schemas-microsoft-com:office:office"><rect xmlns="urn:schemas-microsoft-com:vml" o:hr="t"/></pict>"#.as_slice(),
            br#"<w:pict xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:foreign" xmlns:o="urn:schemas-microsoft-com:office:office"><v:rect xmlns:v="urn:schemas-microsoft-com:vml" o:hr="true"/></w:pict>"#.as_slice(),
        ];

        for raw in cases {
            let run = run_with_raw(raw);
            let run = RunRef { inner: &run };
            let item = run.items().next().unwrap();
            let RunItemRef::LegacyHorizontalRule(rule) = item else {
                panic!("namespace-qualified legacy rule was not classified");
            };
            assert_eq!(rule.raw_xml(), raw);
        }
    }

    #[test]
    fn ambiguous_or_foreign_vml_stays_unsupported() {
        let cases = [
            br#"<w:pict xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><v:rect o:hr="1"/></w:pict>"#.as_slice(),
            br#"<w:pict xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><v:rect o:hr="false"/></w:pict>"#.as_slice(),
            br#"<w:pict xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml"><v:rect/></w:pict>"#.as_slice(),
            br#"<x:pict xmlns:x="urn:foreign" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><v:rect o:hr="t"/></x:pict>"#.as_slice(),
            br#"<w:pict xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:foreign" xmlns:o="urn:schemas-microsoft-com:office:office"><x:rect o:hr="t"/></w:pict>"#.as_slice(),
            br#"<w:pict xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:x="urn:foreign"><v:rect x:hr="t"/></w:pict>"#.as_slice(),
            br#"<w:pict xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><v:rect xmlns:v="urn:foreign" o:hr="t"/></w:pict>"#.as_slice(),
            br#"<w:pict xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><v:rect o:hr="t"/><v:rect o:hr="t"/></w:pict>"#.as_slice(),
            br#"<w:pict xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><v:rect o:hr="t">visible</v:rect></w:pict>"#.as_slice(),
            br#"<w:pict xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><!-- ambiguous --><v:rect o:hr="t"/></w:pict>"#.as_slice(),
        ];

        for raw in cases {
            let run = run_with_raw(raw);
            assert!(matches!(
                RunRef { inner: &run }.items().next().unwrap(),
                RunItemRef::UnsupportedXml(bytes) if bytes == raw
            ));
        }

        let malformed = br#"<w:pict xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><v:rect o:hr="t"></w:pict>"#;
        let run = CT_R {
            properties: None,
            content: Vec::new(),
            extra_xml: vec![malformed.to_vec()],
            extra_xml_positions: vec![0],
            alt_drawings: Vec::new(),
        };
        assert!(matches!(
            RunRef { inner: &run }.items().next().unwrap(),
            RunItemRef::UnsupportedXml(bytes) if bytes == malformed
        ));
    }

    #[test]
    fn legacy_horizontal_rule_keeps_exact_raw_xml_and_item_order() {
        let raw = br#"<w:pict xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><v:rect o:hr="t"/></w:pict>"#;
        let run = run_with_inner_xml(&format!(
            "<w:t>before</w:t>{}<w:t>after</w:t>",
            std::str::from_utf8(raw).unwrap()
        ));
        let run = RunRef { inner: &run };
        let items = run.items().collect::<Vec<_>>();

        assert!(matches!(items[0], RunItemRef::Text("before")));
        let RunItemRef::LegacyHorizontalRule(rule) = items[1] else {
            panic!("legacy rule did not retain its ordered item slot");
        };
        assert_eq!(rule.raw_xml(), raw);
        assert!(matches!(items[2], RunItemRef::Text("after")));
    }

    #[test]
    fn unordered_legacy_raw_children_follow_typed_run_content() {
        let run = CT_R {
            properties: None,
            content: vec![RunContent::Text(CT_Text::new("typed"))],
            extra_xml: vec![b"<x:raw/>".to_vec()],
            extra_xml_positions: Vec::new(),
            alt_drawings: Vec::new(),
        };
        let run = RunRef { inner: &run };
        let items = run.items().collect::<Vec<_>>();
        assert!(matches!(items[0], RunItemRef::Text("typed")));
        assert!(matches!(items[1], RunItemRef::UnsupportedXml(b"<x:raw/>")));
    }

    #[test]
    fn field_reader_reports_unmodeled_semantic_attributes() {
        let supported = run_from_paragraph_inner_xml(
            r#"<w:fldSimple w:instr="PAGE" w:dirty="false"><w:r><w:t>1</w:t></w:r></w:fldSimple>"#,
        );
        let unsupported = run_from_paragraph_inner_xml(
            r#"<w:fldSimple w:instr="PAGE" w:fldLock="true"><w:r><w:t>1</w:t></w:r></w:fldSimple>"#,
        );

        let supported_run = RunRef { inner: &supported };
        let RunItemRef::Field(supported) = supported_run.items().next().expect("supported field")
        else {
            panic!("expected field");
        };
        let unsupported_run = RunRef {
            inner: &unsupported,
        };
        let RunItemRef::Field(unsupported) =
            unsupported_run.items().next().expect("unsupported field")
        else {
            panic!("expected field");
        };

        assert!(!supported.has_unmodeled_semantic_attributes());
        assert!(unsupported.has_unmodeled_semantic_attributes());
    }

    #[test]
    fn field_reader_exposes_complex_cached_display_segments() {
        let simple = run_from_paragraph_inner_xml(
            r#"<w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple>"#,
        );
        let complex = run_from_paragraph_inner_xml(
            r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>PAGE</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>1</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>2</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
        );

        let simple_run = RunRef { inner: &simple };
        let RunItemRef::Field(simple) = simple_run.items().next().unwrap() else {
            panic!("expected simple field");
        };
        let complex_run = RunRef { inner: &complex };
        let RunItemRef::Field(complex) = complex_run.items().next().unwrap() else {
            panic!("expected complex field");
        };

        assert_eq!(simple.kind(), FieldKind::Simple);
        assert_eq!(complex.kind(), FieldKind::Complex);
        let segments = complex.cached_display_segments();
        assert_eq!(segments.len(), 2);
        assert_eq!(segments[0].text(), "1");
        assert_eq!(segments[0].properties().unwrap().bold, Some(true));
        assert_eq!(segments[1].text(), "2");
        assert_eq!(segments[1].properties().unwrap().italic, Some(true));
    }

    #[test]
    fn run_language_authoring_sets_and_clears_the_complete_word_language_value() {
        let mut inner = CT_R::new("hyphenation");
        {
            let mut run = Run { inner: &mut inner };
            run.set_language_value(Some("en-US"));
        }
        assert_eq!(
            inner.properties.as_ref().unwrap().language.as_deref(),
            Some("en-US")
        );
        {
            let mut run = Run { inner: &mut inner };
            run.set_language_value(None);
        }
        assert_eq!(inner.properties.as_ref().unwrap().language, None);
    }
}
