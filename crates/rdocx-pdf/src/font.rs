//! Font subsetting and ToUnicode CMap generation for PDF embedding.

use std::collections::HashMap;

use pdf_writer::types::{SystemInfo, UnicodeCmap};
use pdf_writer::{Name, Str};
use rdocx_layout::{FontData, FontId, LayoutResult, PositionedElement};
use subsetter::GlyphRemapper;

/// Per-font glyph usage collected across all pages.
pub(crate) struct FontUsage {
    /// Mapping from original glyph ID to the Unicode text it represents.
    /// Multiple characters may map to one glyph, so we store the first seen.
    pub glyph_to_unicode: HashMap<u16, char>,
    /// The GlyphRemapper for subsetting.
    pub remapper: GlyphRemapper,
}

/// Collected font info ready for PDF embedding.
pub(crate) struct PreparedFont {
    pub font_data: FontData,
    pub subset_bytes: Vec<u8>,
    pub remapper: GlyphRemapper,
    pub cmap_bytes: Vec<u8>,
    pub widths: Vec<(u16, f64)>, // (new_gid, width_in_font_units)
}

/// Collect glyph usage across all pages for each font.
pub(crate) fn collect_glyph_usage(layout: &LayoutResult) -> HashMap<FontId, FontUsage> {
    let mut usage: HashMap<FontId, FontUsage> = HashMap::new();

    for page in &layout.pages {
        for element in &page.elements {
            if let PositionedElement::Text(run) = element {
                let entry = usage.entry(run.font_id).or_insert_with(|| FontUsage {
                    glyph_to_unicode: HashMap::new(),
                    remapper: GlyphRemapper::new(),
                });

                // Map glyph IDs to unicode chars from the text
                let chars: Vec<char> = run.text.chars().collect();
                for (i, &gid) in run.glyph_ids.iter().enumerate() {
                    entry.remapper.remap(gid);
                    if let Some(&ch) = chars.get(i) {
                        entry.glyph_to_unicode.entry(gid).or_insert(ch);
                    }
                }
            }
        }
    }

    usage
}

/// Subset a font and prepare it for PDF embedding.
pub(crate) fn prepare_font(
    font_data: &FontData,
    usage: &mut FontUsage,
) -> Option<PreparedFont> {
    // Subset the font
    let subset_bytes =
        subsetter::subset(&font_data.data, font_data.face_index, &usage.remapper).ok()?;

    // Build ToUnicode CMap
    let mut cmap = UnicodeCmap::new(
        Name(b"Adobe"),
        SystemInfo {
            registry: Str(b"Adobe"),
            ordering: Str(b"Identity"),
            supplement: 0,
        },
    );
    for (&old_gid, &ch) in &usage.glyph_to_unicode {
        if let Some(new_gid) = usage.remapper.get(old_gid) {
            cmap.pair(new_gid, ch);
        }
    }
    let cmap_bytes = cmap.finish().to_vec();

    // Compute glyph widths from the original font for the CID widths array.
    // We need to parse the font to get per-glyph advance widths.
    let widths = compute_glyph_widths(font_data, usage);

    Some(PreparedFont {
        font_data: font_data.clone(),
        subset_bytes,
        remapper: usage.remapper.clone(),
        cmap_bytes,
        widths,
    })
}

/// Compute per-glyph widths in font design units (1000 units = 1 em).
fn compute_glyph_widths(font_data: &FontData, usage: &FontUsage) -> Vec<(u16, f64)> {
    let mut widths = Vec::new();

    let face = match ttf_parser::Face::parse(&font_data.data, font_data.face_index) {
        Ok(f) => f,
        Err(_) => return widths,
    };

    let units_per_em = face.units_per_em() as f64;
    let scale = 1000.0 / units_per_em;

    for &old_gid in usage.glyph_to_unicode.keys() {
        if let Some(new_gid) = usage.remapper.get(old_gid) {
            let advance = face
                .glyph_hor_advance(ttf_parser::GlyphId(old_gid))
                .unwrap_or(0) as f64;
            widths.push((new_gid, advance * scale));
        }
    }

    // Also include glyphs that may not have unicode mapping but were remapped
    // (e.g. .notdef glyph 0 is always included by subsetter)

    widths.sort_by_key(|&(gid, _)| gid);
    widths
}

/// Get font metrics from raw font data for the font descriptor.
pub(crate) struct FontMetricsInfo {
    pub ascent: f64,
    pub descent: f64,
    pub cap_height: f64,
    pub bbox: [f64; 4],
    pub italic_angle: f64,
    pub stem_v: f64,
}

pub(crate) fn get_font_metrics(font_data: &FontData) -> Option<FontMetricsInfo> {
    let face = ttf_parser::Face::parse(&font_data.data, font_data.face_index).ok()?;
    let units_per_em = face.units_per_em() as f64;
    let scale = 1000.0 / units_per_em;

    let ascent = face.ascender() as f64 * scale;
    let descent = face.descender() as f64 * scale;
    let cap_height = match face.capital_height() {
        Some(h) => h as f64 * scale,
        None => ascent,
    };
    let italic_angle = face.italic_angle() as f64;

    let bbox = face.global_bounding_box();
    let bbox = [
        bbox.x_min as f64 * scale,
        bbox.y_min as f64 * scale,
        bbox.x_max as f64 * scale,
        bbox.y_max as f64 * scale,
    ];

    // Approximate stem_v from weight class
    let stem_v = if font_data.bold { 120.0 } else { 80.0 };

    Some(FontMetricsInfo {
        ascent,
        descent,
        cap_height,
        bbox,
        italic_angle,
        stem_v,
    })
}
