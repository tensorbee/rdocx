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
        let rpr = self.ensure_rpr();
        rpr.bold = Some(val);
        rpr.bold_cs = Some(val);
        self
    }

    /// Set italic formatting.
    pub fn italic(mut self, val: bool) -> Self {
        let rpr = self.ensure_rpr();
        rpr.italic = Some(val);
        rpr.italic_cs = Some(val);
        self
    }

    /// Set underline formatting (simple on/off).
    pub fn underline(mut self, val: bool) -> Self {
        self.ensure_rpr().underline = Some(if val {
            ST_Underline::Single
        } else {
            ST_Underline::None
        });
        self
    }

    /// Set underline with a specific style.
    pub fn underline_style(mut self, style: UnderlineStyle) -> Self {
        self.ensure_rpr().underline = Some(style.to_st());
        self
    }

    /// Set font size in points.
    pub fn size(mut self, pt: f64) -> Self {
        let hp = HalfPoint::from_pt(pt);
        let rpr = self.ensure_rpr();
        rpr.sz = Some(hp);
        rpr.sz_cs = Some(hp);
        self
    }

    /// Set the font name.
    pub fn font(mut self, name: &str) -> Self {
        let rpr = self.ensure_rpr();
        rpr.font_ascii = Some(name.to_string());
        rpr.font_hansi = Some(name.to_string());
        rpr.font_east_asia = Some(name.to_string());
        rpr.font_cs = Some(name.to_string());
        self
    }

    /// Set text color as a hex string (e.g., "FF0000" for red).
    pub fn color(mut self, hex: &str) -> Self {
        self.ensure_rpr().color = Some(hex.to_string());
        self
    }

    /// Set highlight color as a hex fill value.
    pub fn highlight(mut self, color: &str) -> Self {
        self.ensure_rpr().shading = Some(CT_Shd {
            val: "clear".to_string(),
            color: Some("auto".to_string()),
            fill: Some(color.to_string()),
        });
        self
    }

    /// Set strikethrough formatting.
    pub fn strike(mut self, val: bool) -> Self {
        self.ensure_rpr().strike = Some(val);
        self
    }

    /// Set double strikethrough.
    pub fn double_strike(mut self, val: bool) -> Self {
        self.ensure_rpr().dstrike = Some(val);
        self
    }

    /// Set all caps.
    pub fn all_caps(mut self, val: bool) -> Self {
        self.ensure_rpr().caps = Some(val);
        self
    }

    /// Set small caps.
    pub fn small_caps(mut self, val: bool) -> Self {
        self.ensure_rpr().small_caps = Some(val);
        self
    }

    /// Set superscript.
    pub fn superscript(mut self) -> Self {
        self.ensure_rpr().vert_align = Some("superscript".to_string());
        self
    }

    /// Set subscript.
    pub fn subscript(mut self) -> Self {
        self.ensure_rpr().vert_align = Some("subscript".to_string());
        self
    }

    /// Set character spacing (positive = expanded, negative = condensed).
    pub fn character_spacing(mut self, spacing: Length) -> Self {
        self.ensure_rpr().spacing = Some(spacing.as_twips());
        self
    }

    /// Set character width scale in percent (100 = normal).
    pub fn width_scale(mut self, percent: u32) -> Self {
        self.ensure_rpr().width_scale = Some(percent);
        self
    }

    /// Set text position (positive = raised, negative = lowered) in half-points.
    pub fn position(mut self, half_points: i32) -> Self {
        self.ensure_rpr().position = Some(half_points);
        self
    }

    /// Set hidden/vanish text.
    pub fn hidden(mut self, val: bool) -> Self {
        self.ensure_rpr().vanish = Some(val);
        self
    }

    /// Set the character style by ID.
    pub fn style(mut self, style_id: &str) -> Self {
        self.ensure_rpr().style_id = Some(style_id.to_string());
        self
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
        self.inner
            .properties
            .as_ref()
            .and_then(|rpr| rpr.spacing)
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
