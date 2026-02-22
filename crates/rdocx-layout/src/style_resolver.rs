//! Style resolution: cascade styles and generate numbering markers.
//!
//! Ports the logic from `crates/rdocx/src/style.rs` since rdocx-layout
//! depends on rdocx-oxml directly (not rdocx).

use std::collections::{HashMap, HashSet};

use rdocx_oxml::numbering::{CT_Numbering, ST_NumberFormat};
use rdocx_oxml::properties::{CT_PPr, CT_RPr};
use rdocx_oxml::styles::{CT_Style, CT_Styles, StyleType};

/// A fully resolved paragraph with merged properties and numbering info.
#[derive(Debug, Clone)]
pub struct ResolvedParagraph {
    /// Merged paragraph properties (style chain + direct formatting).
    pub ppr: CT_PPr,
    /// Resolved runs with merged run properties.
    pub runs: Vec<ResolvedRun>,
    /// Numbering marker info (if paragraph is part of a list).
    pub numbering: Option<ResolvedNumbering>,
}

/// A run with fully resolved properties.
#[derive(Debug, Clone)]
pub struct ResolvedRun {
    /// Merged run properties.
    pub rpr: CT_RPr,
    /// Run content items.
    pub content: Vec<rdocx_oxml::text::RunContent>,
}

/// Resolved numbering marker for a list paragraph.
#[derive(Debug, Clone)]
pub struct ResolvedNumbering {
    /// The text of the marker (e.g., "1.", "a)", bullet char).
    pub marker_text: String,
    /// Run properties for the marker.
    pub marker_rpr: CT_RPr,
}

/// Tracks numbering counters across paragraphs.
pub struct NumberingState {
    /// (numId, ilvl) → current count
    counters: HashMap<(u32, u32), u32>,
}

impl Default for NumberingState {
    fn default() -> Self {
        Self::new()
    }
}

impl NumberingState {
    pub fn new() -> Self {
        NumberingState {
            counters: HashMap::new(),
        }
    }

    /// Advance the counter for the given numId/ilvl and return the new value.
    /// Also resets any deeper levels.
    pub fn advance(&mut self, num_id: u32, ilvl: u32, start: u32) -> u32 {
        let key = (num_id, ilvl);
        let counter = self.counters.entry(key).or_insert(start - 1);
        *counter += 1;
        let value = *counter;

        // Reset deeper levels
        for deeper in (ilvl + 1)..=8 {
            self.counters.remove(&(num_id, deeper));
        }

        value
    }

    /// Get the current count for a level (without advancing).
    pub fn current(&self, num_id: u32, ilvl: u32) -> u32 {
        self.counters.get(&(num_id, ilvl)).copied().unwrap_or(0)
    }
}

/// Resolve paragraph properties by walking the style inheritance chain.
pub fn resolve_paragraph_properties(
    style_id: Option<&str>,
    styles: &CT_Styles,
) -> CT_PPr {
    let mut effective = CT_PPr::default();

    // 1. Start from docDefaults
    if let Some(ref defaults) = styles.doc_defaults {
        if let Some(ref ppr) = defaults.ppr {
            effective.merge_from(ppr);
        }
    }

    // 2. Walk the basedOn chain
    if let Some(sid) = style_id {
        let chain = collect_style_chain(sid, styles);
        // Apply from most-base to most-derived
        for style in chain.iter().rev() {
            if let Some(ref ppr) = style.ppr {
                effective.merge_from(ppr);
            }
        }
    } else {
        // Apply the default paragraph style
        if let Some(default_style) = styles.get_default(StyleType::Paragraph) {
            if let Some(ref ppr) = default_style.ppr {
                effective.merge_from(ppr);
            }
        }
    }

    effective
}

/// Resolve run properties by walking paragraph and character style chains.
pub fn resolve_run_properties(
    para_style_id: Option<&str>,
    run_style_id: Option<&str>,
    styles: &CT_Styles,
) -> CT_RPr {
    let mut effective = CT_RPr::default();

    // 1. docDefaults run properties
    if let Some(ref defaults) = styles.doc_defaults {
        if let Some(ref rpr) = defaults.rpr {
            effective.merge_from(rpr);
        }
    }

    // 2. paragraph style's rpr (following basedOn chain)
    let para_sid = para_style_id.or_else(|| {
        styles
            .get_default(StyleType::Paragraph)
            .map(|s| s.style_id.as_str())
    });
    if let Some(sid) = para_sid {
        let chain = collect_style_chain(sid, styles);
        for style in chain.iter().rev() {
            if let Some(ref rpr) = style.rpr {
                effective.merge_from(rpr);
            }
        }
    }

    // 3. character style's rpr (following basedOn chain)
    if let Some(sid) = run_style_id {
        let chain = collect_style_chain(sid, styles);
        for style in chain.iter().rev() {
            if let Some(ref rpr) = style.rpr {
                effective.merge_from(rpr);
            }
        }
    }

    effective
}

/// Generate the marker text for a numbered/bulleted list item.
pub fn generate_marker(
    num_id: u32,
    ilvl: u32,
    numbering: &CT_Numbering,
    state: &mut NumberingState,
) -> Option<ResolvedNumbering> {
    let abs = numbering.get_abstract_num_for(num_id)?;
    let lvl = abs.levels.iter().find(|l| l.ilvl == ilvl)?;

    let num_fmt = lvl.num_fmt.unwrap_or(ST_NumberFormat::Decimal);
    let start = lvl.start.unwrap_or(1);
    let lvl_text = lvl.lvl_text.as_deref().unwrap_or("%1.");

    let marker_text = if num_fmt == ST_NumberFormat::Bullet {
        lvl_text.to_string()
    } else {
        let count = state.advance(num_id, ilvl, start);
        format_lvl_text(lvl_text, num_id, ilvl, count, numbering, state)
    };

    let marker_rpr = lvl.rpr.clone().unwrap_or_default();

    Some(ResolvedNumbering {
        marker_text,
        marker_rpr,
    })
}

/// Format level text by substituting %1, %2, etc. with formatted counters.
fn format_lvl_text(
    template: &str,
    num_id: u32,
    current_ilvl: u32,
    current_count: u32,
    numbering: &CT_Numbering,
    state: &NumberingState,
) -> String {
    let abs = match numbering.get_abstract_num_for(num_id) {
        Some(a) => a,
        None => return template.to_string(),
    };

    let mut result = template.to_string();
    for lvl_idx in 0..=8u32 {
        let placeholder = format!("%{}", lvl_idx + 1);
        if result.contains(&placeholder) {
            let count = if lvl_idx == current_ilvl {
                current_count
            } else {
                state.current(num_id, lvl_idx)
            };
            let fmt = abs
                .levels
                .iter()
                .find(|l| l.ilvl == lvl_idx)
                .and_then(|l| l.num_fmt)
                .unwrap_or(ST_NumberFormat::Decimal);
            let formatted = format_number(count, fmt);
            result = result.replace(&placeholder, &formatted);
        }
    }
    result
}

/// Format a number according to ST_NumberFormat.
fn format_number(n: u32, fmt: ST_NumberFormat) -> String {
    match fmt {
        ST_NumberFormat::Decimal => n.to_string(),
        ST_NumberFormat::UpperRoman => to_roman(n, true),
        ST_NumberFormat::LowerRoman => to_roman(n, false),
        ST_NumberFormat::UpperLetter => to_letter(n, true),
        ST_NumberFormat::LowerLetter => to_letter(n, false),
        ST_NumberFormat::Ordinal => format!("{n}"),
        ST_NumberFormat::Bullet | ST_NumberFormat::None => String::new(),
    }
}

fn to_roman(mut n: u32, upper: bool) -> String {
    let vals = [
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
    ];
    let mut result = String::new();
    for &(value, numeral) in &vals {
        while n >= value {
            result.push_str(numeral);
            n -= value;
        }
    }
    if upper {
        result
    } else {
        result.to_lowercase()
    }
}

fn to_letter(n: u32, upper: bool) -> String {
    if n == 0 {
        return String::new();
    }
    let base = if upper { b'A' } else { b'a' };
    let idx = ((n - 1) % 26) as u8;
    String::from(char::from(base + idx))
}

/// Collect the style chain from the given style up through basedOn ancestors.
fn collect_style_chain<'a>(style_id: &str, styles: &'a CT_Styles) -> Vec<&'a CT_Style> {
    let mut chain = Vec::new();
    let mut current_id = Some(style_id.to_string());
    let mut seen = HashSet::new();

    while let Some(ref sid) = current_id {
        if !seen.insert(sid.clone()) {
            break; // Prevent cycles
        }
        if let Some(style) = styles.get_by_id(sid) {
            chain.push(style);
            current_id = style.based_on.clone();
        } else {
            break;
        }
    }

    chain
}

#[cfg(test)]
mod tests {
    use super::*;
    use rdocx_oxml::units::{HalfPoint, Twips};

    fn test_styles() -> CT_Styles {
        let mut styles = CT_Styles::new_default();
        styles.styles.push(CT_Style {
            style_id: "Heading2".to_string(),
            style_type: StyleType::Paragraph,
            name: Some("heading 2".to_string()),
            based_on: Some("Heading1".to_string()),
            next_style: Some("Normal".to_string()),
            is_default: false,
            ppr: Some(CT_PPr {
                space_before: Some(Twips(40)),
                ..Default::default()
            }),
            rpr: Some(CT_RPr {
                sz: Some(HalfPoint(26)),
                color: Some("2E74B5".to_string()),
                ..Default::default()
            }),
        });
        styles
    }

    #[test]
    fn resolve_normal_paragraph() {
        let styles = test_styles();
        let ppr = resolve_paragraph_properties(Some("Normal"), &styles);
        assert_eq!(ppr.space_after, Some(Twips(160)));
    }

    #[test]
    fn resolve_heading1() {
        let styles = test_styles();
        let ppr = resolve_paragraph_properties(Some("Heading1"), &styles);
        assert_eq!(ppr.keep_next, Some(true));
        assert_eq!(ppr.space_before, Some(Twips(240)));
        assert_eq!(ppr.space_after, Some(Twips(0)));
    }

    #[test]
    fn resolve_heading2_inherits_heading1() {
        let styles = test_styles();
        let ppr = resolve_paragraph_properties(Some("Heading2"), &styles);
        assert_eq!(ppr.keep_next, Some(true));
        assert_eq!(ppr.space_before, Some(Twips(40)));
    }

    #[test]
    fn resolve_heading2_rpr() {
        let styles = test_styles();
        let rpr = resolve_run_properties(Some("Heading2"), None, &styles);
        assert_eq!(rpr.font_ascii, Some("Calibri".to_string()));
        assert_eq!(rpr.sz, Some(HalfPoint(26)));
        assert_eq!(rpr.bold, Some(true));
        assert_eq!(rpr.color, Some("2E74B5".to_string()));
    }

    #[test]
    fn numbering_decimal_marker() {
        let mut numbering = CT_Numbering::new();
        let num_id = numbering.add_numbered_list();

        let mut state = NumberingState::new();
        let marker1 = generate_marker(num_id, 0, &numbering, &mut state).unwrap();
        assert_eq!(marker1.marker_text, "1.");
        let marker2 = generate_marker(num_id, 0, &numbering, &mut state).unwrap();
        assert_eq!(marker2.marker_text, "2.");
    }

    #[test]
    fn numbering_bullet_marker() {
        let mut numbering = CT_Numbering::new();
        let num_id = numbering.add_bullet_list();

        let mut state = NumberingState::new();
        let marker = generate_marker(num_id, 0, &numbering, &mut state).unwrap();
        assert_eq!(marker.marker_text, "\u{2022}");
    }

    #[test]
    fn numbering_sub_level_reset() {
        let mut numbering = CT_Numbering::new();
        let num_id = numbering.add_numbered_list();

        let mut state = NumberingState::new();
        // Level 0: 1, 2
        generate_marker(num_id, 0, &numbering, &mut state);
        generate_marker(num_id, 0, &numbering, &mut state);
        // Level 1: a
        let sub = generate_marker(num_id, 1, &numbering, &mut state).unwrap();
        assert_eq!(sub.marker_text, "a.");
        // Back to level 0: 3 — this should reset level 1
        generate_marker(num_id, 0, &numbering, &mut state);
        let sub2 = generate_marker(num_id, 1, &numbering, &mut state).unwrap();
        assert_eq!(sub2.marker_text, "a."); // reset
    }

    #[test]
    fn roman_numeral_formatting() {
        assert_eq!(to_roman(1, true), "I");
        assert_eq!(to_roman(4, true), "IV");
        assert_eq!(to_roman(9, true), "IX");
        assert_eq!(to_roman(14, false), "xiv");
    }

    #[test]
    fn letter_formatting() {
        assert_eq!(to_letter(1, false), "a");
        assert_eq!(to_letter(26, false), "z");
        assert_eq!(to_letter(27, false), "a"); // wraps
        assert_eq!(to_letter(1, true), "A");
    }
}
