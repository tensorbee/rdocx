//! Run — a contiguous stretch of text with uniform formatting.

use rdocx_oxml::properties::{CT_RPr, CT_Shd};
use rdocx_oxml::shared::ST_Underline;
use rdocx_oxml::text::{CT_R, CT_Text, RunContent};
use rdocx_oxml::units::{HalfPoint, Twips};

use crate::Length;

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
    fn to_st(self) -> ST_Underline {
        match self {
            Self::None => ST_Underline::None,
            Self::Single => ST_Underline::Single,
            Self::Double => ST_Underline::Double,
            Self::Thick => ST_Underline::Thick,
            Self::Dotted => ST_Underline::Dotted,
            Self::Dash => ST_Underline::Dash,
            Self::Wave => ST_Underline::Wave,
            Self::Words => ST_Underline::Words,
        }
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
        self.inner.content = vec![RunContent::Text(CT_Text::new(text))];
    }

    /// Add text to this run.
    pub fn add_text(&mut self, text: &str) {
        self.inner
            .content
            .push(RunContent::Text(CT_Text::new(text)));
    }

    /// Set bold formatting.
    pub fn bold(mut self, val: bool) -> Self {
        self.set_bold(val);
        self
    }

    /// Set bold formatting in place.
    pub fn set_bold(&mut self, val: bool) {
        let rpr = self.ensure_rpr();
        rpr.bold = Some(val);
        rpr.bold_cs = Some(val);
    }

    /// Set italic formatting.
    pub fn italic(mut self, val: bool) -> Self {
        self.set_italic(val);
        self
    }

    /// Set italic formatting in place.
    pub fn set_italic(&mut self, val: bool) {
        let rpr = self.ensure_rpr();
        rpr.italic = Some(val);
        rpr.italic_cs = Some(val);
    }

    /// Set underline formatting (simple on/off).
    pub fn underline(mut self, val: bool) -> Self {
        self.set_underline(val);
        self
    }

    /// Set underline formatting in place.
    pub fn set_underline(&mut self, val: bool) {
        self.ensure_rpr().underline = Some(if val {
            ST_Underline::Single
        } else {
            ST_Underline::None
        });
    }

    /// Set underline with a specific style.
    pub fn underline_style(mut self, style: UnderlineStyle) -> Self {
        self.set_underline_style(style);
        self
    }

    /// Set an underline style in place.
    pub fn set_underline_style(&mut self, style: UnderlineStyle) {
        self.ensure_rpr().underline = Some(style.to_st());
    }

    /// Set font size in points.
    pub fn size(mut self, pt: f64) -> Self {
        self.set_size(pt);
        self
    }

    /// Set font size in place.
    pub fn set_size(&mut self, pt: f64) {
        let hp = HalfPoint::from_pt(pt);
        let rpr = self.ensure_rpr();
        rpr.sz = Some(hp);
        rpr.sz_cs = Some(hp);
    }

    /// Set the font name.
    pub fn font(mut self, name: &str) -> Self {
        self.set_font(name);
        self
    }

    /// Set the font name in place.
    pub fn set_font(&mut self, name: &str) {
        let rpr = self.ensure_rpr();
        rpr.font_ascii = Some(name.to_string());
        rpr.font_hansi = Some(name.to_string());
        rpr.font_east_asia = Some(name.to_string());
        rpr.font_cs = Some(name.to_string());
    }

    /// Set text color as a hex string (e.g., "FF0000" for red).
    pub fn color(mut self, hex: &str) -> Self {
        self.set_color(hex);
        self
    }

    /// Set text color in place.
    pub fn set_color(&mut self, hex: &str) {
        self.ensure_rpr().color = Some(hex.to_string());
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

    /// Set strikethrough formatting.
    pub fn strike(mut self, val: bool) -> Self {
        self.set_strike(val);
        self
    }

    /// Set strikethrough formatting in place.
    pub fn set_strike(&mut self, val: bool) {
        self.ensure_rpr().strike = Some(val);
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

    fn ensure_rpr(&mut self) -> &mut CT_RPr {
        self.inner.properties.get_or_insert_with(CT_RPr::default)
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
        self.inner
            .properties
            .as_ref()
            .and_then(|rpr| rpr.bold)
            .unwrap_or(false)
    }

    /// Check if italic.
    pub fn is_italic(&self) -> bool {
        self.inner
            .properties
            .as_ref()
            .and_then(|rpr| rpr.italic)
            .unwrap_or(false)
    }

    /// Check if strikethrough.
    pub fn is_strike(&self) -> bool {
        self.inner
            .properties
            .as_ref()
            .and_then(|rpr| rpr.strike)
            .unwrap_or(false)
    }

    /// Check if underlined (any underline style other than none).
    pub fn is_underline(&self) -> bool {
        self.inner
            .properties
            .as_ref()
            .and_then(|rpr| rpr.underline.as_ref())
            .is_some_and(|u| !matches!(u, ST_Underline::None))
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
