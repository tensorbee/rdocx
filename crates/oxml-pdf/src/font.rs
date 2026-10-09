//! Font subsetting and ToUnicode CMap generation for PDF embedding.

use std::collections::btree_map::Entry;
use std::collections::{BTreeMap, BTreeSet, HashMap};

use oxml_layout::{FontData, FontId, LayoutResult, PositionedElement, walk};
use pdf_writer::types::{SystemInfo, UnicodeCmap};
use pdf_writer::{Name, Str};
use subsetter::GlyphRemapper;
use ttf_parser::gsub::SubstitutionSubtable;
use ttf_parser::opentype_layout::Coverage;

/// Per-font glyph usage collected across all pages.
pub(crate) struct FontUsage {
    /// Mapping from original glyph ID to the Unicode text it draws.
    ///
    /// A ligature draws several characters with one glyph, so the text is a
    /// string. The CMap holds one entry per glyph for the whole font, so when
    /// runs pair one glyph with different text, the strongest pairing wins,
    /// and the first one seen among equals.
    ///
    /// Ordered by glyph ID, because this map is iterated to emit the ToUnicode
    /// CMap. A hashed order put the same pairs in a different order on every
    /// run, which made the written PDF differ from itself byte for byte.
    pub glyph_to_unicode: BTreeMap<u16, (Pairing, String)>,
    /// Every glyph a run draws. Each declares its width, including a glyph
    /// that carries no text of its own.
    pub drawn_glyphs: BTreeSet<u16>,
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

impl FontUsage {
    /// Subset a drawn glyph, declare its width, and map it to its text unless
    /// another run paired it more strongly.
    fn draw(&mut self, glyph: u16, text: Option<(Pairing, String)>) {
        self.remapper.remap(glyph);
        self.drawn_glyphs.insert(glyph);
        let Some((pairing, text)) = text else {
            return;
        };
        match self.glyph_to_unicode.entry(glyph) {
            Entry::Vacant(entry) => {
                entry.insert((pairing, text));
            }
            Entry::Occupied(mut entry) if entry.get().0 < pairing => {
                entry.insert((pairing, text));
            }
            Entry::Occupied(_) => {}
        }
    }
}

/// How strongly a glyph was paired with the text it draws, weakest first.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub(crate) enum Pairing {
    /// A glyph of a rich run that draws none of its cluster's characters by
    /// itself, such as the dots an Arabic font draws apart from their letter.
    /// It repeats the cluster's text, which the run's `ActualText` covers, so
    /// that every glyph of a rich run maps to Unicode.
    Shared,
    /// The first of several glyphs that draw some characters between them
    /// in no order the font explains. It carries all of those characters.
    Group,
    /// Paired by position among as many unexplained glyphs as characters.
    Position,
    /// The one glyph left to draw the characters of its group that no other
    /// glyph explains.
    Whole,
    /// The glyph the font's cmap gives the character.
    Nominal,
    /// The glyph a GSUB ligature joins the glyphs of several characters into.
    /// It outranks the cmap because one glyph can be both. Carlito draws "fi"
    /// and U+FB01 with one glyph, and one literal U+FB01 must not turn every
    /// "fi" of the document into U+FB01.
    Ligature,
}

/// What a font says about the glyphs a run draws.
struct FontGlyphs<'a> {
    face: Option<ttf_parser::Face<'a>>,
    /// Each ligature glyph, with the glyph sequences GSUB joins into it.
    ligatures: HashMap<u16, Vec<Vec<u16>>>,
}

impl<'a> FontGlyphs<'a> {
    fn new(font: Option<&'a FontData>) -> Self {
        let face = font.and_then(|font| ttf_parser::Face::parse(&font.data, font.face_index).ok());
        let ligatures = face.as_ref().map(ligature_components).unwrap_or_default();
        Self { face, ligatures }
    }

    /// The glyph the cmap gives each character, `.notdef` for a character
    /// the font lacks, which is also what the shaper draws for it.
    fn nominal_glyphs(&self, characters: &[char]) -> Vec<Option<u16>> {
        characters
            .iter()
            .map(|&character| {
                let face = self.face.as_ref()?;
                Some(face.glyph_index(character).map_or(0, |glyph| glyph.0))
            })
            .collect()
    }

    /// How many characters, from the start of `nominal`, the glyph draws when
    /// it is the cmap glyph of the first or a GSUB ligature of the first few.
    fn explain(&self, glyph: u16, nominal: &[Option<u16>]) -> Option<(usize, Pairing)> {
        if nominal.first() == Some(&Some(glyph)) {
            return Some((1, Pairing::Nominal));
        }
        let length = self.ligature_length(glyph, nominal, 2)?;
        Some((length, Pairing::Ligature))
    }

    fn ligature_length(&self, glyph: u16, nominal: &[Option<u16>], depth: u8) -> Option<usize> {
        let sequences = self.ligatures.get(&glyph)?;
        sequences
            .iter()
            .filter_map(|components| {
                let mut length = 0;
                for &component in components {
                    let rest = &nominal[length..];
                    if rest.first() == Some(&Some(component)) {
                        length += 1;
                    } else if depth > 0 {
                        length += self.ligature_length(component, rest, depth - 1)?;
                    } else {
                        return None;
                    }
                }
                Some(length)
            })
            .max()
    }
}

/// Every ligature glyph of the font's GSUB table, with the glyph sequences
/// it replaces.
fn ligature_components(face: &ttf_parser::Face<'_>) -> HashMap<u16, Vec<Vec<u16>>> {
    let mut ligatures: HashMap<u16, Vec<Vec<u16>>> = HashMap::new();
    let Some(gsub) = face.tables().gsub else {
        return ligatures;
    };
    for lookup in gsub.lookups {
        for subtable in lookup.subtables.into_iter::<SubstitutionSubtable>() {
            let SubstitutionSubtable::Ligature(table) = subtable else {
                continue;
            };
            let first_glyphs: Vec<ttf_parser::GlyphId> = match table.coverage {
                Coverage::Format1 { glyphs } => glyphs.into_iter().collect(),
                Coverage::Format2 { records } => records
                    .into_iter()
                    .flat_map(|record| record.start.0..=record.end.0)
                    .map(ttf_parser::GlyphId)
                    .collect(),
            };
            for first in first_glyphs {
                let Some(set) = table
                    .coverage
                    .get(first)
                    .and_then(|index| table.ligature_sets.get(index))
                else {
                    continue;
                };
                for ligature in set {
                    let mut components = vec![first.0];
                    components.extend(ligature.components.into_iter().map(|glyph| glyph.0));
                    ligatures
                        .entry(ligature.glyph.0)
                        .or_default()
                        .push(components);
                }
            }
        }
    }
    ligatures
}

/// How far past an unexplained glyph the pairing looks for the next glyph the
/// font explains. A ligature or a cluster spans a few characters, and the
/// bound keeps a run the font explains nowhere linear.
const RESYNC_WINDOW: usize = 16;

/// Pair each glyph of a plain run with the characters it draws.
///
/// A plain run carries its glyphs and its text but not the shaper's clusters,
/// and a ligature draws several characters with one glyph, so pairing them by
/// index shifts every later glyph. The pairing is rebuilt from the font
/// instead. A glyph draws the next character when the cmap gives it that
/// character, and the next few when GSUB joins their glyphs into it. Glyphs
/// the font explains neither way share the characters up to the next glyph
/// it does explain. When there is none within `RESYNC_WINDOW`, an unexplained
/// glyph takes one character by position. A plain run has no `ActualText`, so
/// a glyph that draws none of its group's characters by itself gets no text
/// rather than repeat a character another glyph carries, and so does a glyph
/// left after the last character.
fn pair_plain_run(
    characters: &[char],
    glyphs: &[u16],
    font: &FontGlyphs<'_>,
) -> Vec<Option<(Pairing, String)>> {
    let nominal = font.nominal_glyphs(characters);
    let mut texts = vec![None; glyphs.len()];
    let (mut char_at, mut glyph_at) = (0, 0);
    while char_at < characters.len() && glyph_at < glyphs.len() {
        if let Some((length, pairing)) = font.explain(glyphs[glyph_at], &nominal[char_at..]) {
            let text = characters[char_at..char_at + length].iter().collect();
            texts[glyph_at] = Some((pairing, text));
            char_at += length;
            glyph_at += 1;
            continue;
        }

        let char_limit = characters.len().min(char_at + 1 + RESYNC_WINDOW);
        let glyph_limit = glyphs.len().min(glyph_at + 1 + RESYNC_WINDOW);
        let next_explained = (glyph_at + 1..glyph_limit).find_map(|glyph_end| {
            (char_at + 1..char_limit)
                .find(|&char_end| {
                    font.explain(glyphs[glyph_end], &nominal[char_end..])
                        .is_some()
                })
                .map(|char_end| (char_end, glyph_end))
        });
        let (char_end, glyph_end) = match next_explained {
            Some(end) => end,
            None if characters.len() - char_at <= RESYNC_WINDOW
                && glyphs.len() - glyph_at <= RESYNC_WINDOW =>
            {
                (characters.len(), glyphs.len())
            }
            None => {
                texts[glyph_at] = Some((Pairing::Position, characters[char_at].to_string()));
                char_at += 1;
                glyph_at += 1;
                continue;
            }
        };
        let group = pair_group(
            &characters[char_at..char_end],
            &glyphs[glyph_at..glyph_end],
            &nominal[char_at..char_end],
            font,
        );
        texts.splice(glyph_at..glyph_end, group);
        char_at = char_end;
        glyph_at = glyph_end;
    }
    texts
}

/// Pair the glyphs of a group with the characters they draw between them,
/// such as one shaping cluster, in which glyph order need not follow text
/// order.
///
/// A glyph the font explains draws the characters it explains, when no other
/// glyph of the group has taken them. The characters left over go to the one
/// glyph left over, or pair with the glyphs left over by position when they
/// are as many. Otherwise the first glyph left over carries them all, so that
/// none is lost. A glyph still without text draws none of the characters by
/// itself and gets no text, so that no character is paired twice.
fn pair_group(
    characters: &[char],
    glyphs: &[u16],
    nominal: &[Option<u16>],
    font: &FontGlyphs<'_>,
) -> Vec<Option<(Pairing, String)>> {
    let mut texts = vec![None; glyphs.len()];
    let mut claimed = vec![false; characters.len()];
    for (text, &glyph) in texts.iter_mut().zip(glyphs) {
        let explained = (0..characters.len()).find_map(|start| {
            let (length, pairing) = font.explain(glyph, &nominal[start..])?;
            let range = start..start + length;
            claimed[range.clone()]
                .iter()
                .all(|taken| !taken)
                .then_some((range, pairing))
        });
        if let Some((range, pairing)) = explained {
            claimed[range.clone()].fill(true);
            *text = Some((pairing, characters[range].iter().collect()));
        }
    }
    let rest_characters = characters
        .iter()
        .zip(&claimed)
        .filter(|(_, claimed)| !**claimed)
        .map(|(character, _)| *character)
        .collect::<Vec<_>>();
    let rest_glyphs = (0..glyphs.len())
        .filter(|&index| texts[index].is_none())
        .collect::<Vec<_>>();
    match rest_glyphs.as_slice() {
        _ if rest_characters.is_empty() => {}
        [] => {}
        [only] => texts[*only] = Some((Pairing::Whole, rest_characters.into_iter().collect())),
        rest if rest.len() == rest_characters.len() => {
            for (&index, character) in rest.iter().zip(rest_characters) {
                texts[index] = Some((Pairing::Position, character.to_string()));
            }
        }
        [first, ..] => {
            texts[*first] = Some((Pairing::Group, rest_characters.into_iter().collect()));
        }
    }
    texts
}

fn font_glyphs<'f, 'a>(
    fonts: &'f mut HashMap<FontId, FontGlyphs<'a>>,
    layout: &'a LayoutResult,
    font_id: FontId,
) -> &'f FontGlyphs<'a> {
    fonts
        .entry(font_id)
        .or_insert_with(|| FontGlyphs::new(layout.fonts.iter().find(|font| font.id == font_id)))
}

/// Collect glyph usage across all pages for each font.
pub(crate) fn collect_glyph_usage(layout: &LayoutResult) -> HashMap<FontId, FontUsage> {
    let mut usage: HashMap<FontId, FontUsage> = HashMap::new();
    let mut fonts: HashMap<FontId, FontGlyphs<'_>> = HashMap::new();

    for page in &layout.pages {
        walk(&page.elements, &mut |element, _| {
            if let PositionedElement::Text(run) = element
                && !(run.text.is_empty() && run.glyph_ids.is_empty())
            {
                let entry = usage.entry(run.font_id).or_insert_with(|| FontUsage {
                    glyph_to_unicode: BTreeMap::new(),
                    drawn_glyphs: BTreeSet::new(),
                    remapper: GlyphRemapper::new(),
                });

                // Map glyph IDs to the characters each one draws
                let chars: Vec<char> = run.text.chars().collect();
                let font = font_glyphs(&mut fonts, layout, run.font_id);
                let texts = pair_plain_run(&chars, &run.glyph_ids, font);
                for (&gid, text) in run.glyph_ids.iter().zip(texts) {
                    entry.draw(gid, text);
                }
            }
            if let PositionedElement::MultilingualText(run) = element
                && !(run.logical_text.is_empty() && run.glyph_ids.is_empty())
                && run.is_valid()
            {
                let entry = usage.entry(run.font_id).or_insert_with(|| FontUsage {
                    glyph_to_unicode: BTreeMap::new(),
                    drawn_glyphs: BTreeSet::new(),
                    remapper: GlyphRemapper::new(),
                });
                let chars = run.logical_text.chars().collect::<Vec<_>>();
                let font = font_glyphs(&mut fonts, layout, run.font_id);
                for cluster in &run.clusters {
                    let characters =
                        chars.get(cluster.char_start as usize..cluster.char_end as usize);
                    let glyphs = run
                        .glyph_ids
                        .get(cluster.glyph_start as usize..cluster.glyph_end as usize);
                    let (Some(characters), Some(glyphs)) = (characters, glyphs) else {
                        continue;
                    };
                    let nominal = font.nominal_glyphs(characters);
                    let texts = pair_group(characters, glyphs, &nominal, font);
                    // The run's ActualText carries its text, so a glyph that
                    // draws none of the cluster's characters by itself repeats
                    // them all, and every glyph of the run maps to Unicode.
                    let cluster_text = characters.iter().collect::<String>();
                    for (&gid, text) in glyphs.iter().zip(texts) {
                        let text = text.unwrap_or_else(|| (Pairing::Shared, cluster_text.clone()));
                        entry.draw(gid, Some(text));
                    }
                }
            }
        });
    }

    usage
}

/// Parse a face at the variation instance layout shaped it with.
pub(crate) fn instanced_face(font_data: &FontData) -> Option<ttf_parser::Face<'_>> {
    let mut face = ttf_parser::Face::parse(&font_data.data, font_data.face_index).ok()?;
    for (tag, value) in &font_data.variations {
        face.set_variation(ttf_parser::Tag::from_bytes(tag), *value);
    }
    Some(face)
}

/// Subset a font and prepare it for PDF embedding.
///
/// A variable face is instanced at its layout coordinates, so the embedded
/// outlines are the ones layout measured rather than the default instance.
/// Should instancing fail, the default outlines are embedded instead, still
/// placed at the layout advances, so the text never disappears.
pub(crate) fn prepare_font(font_data: &FontData, usage: &mut FontUsage) -> Option<PreparedFont> {
    let coordinates = font_data
        .variations
        .iter()
        .map(|(tag, value)| (subsetter::Tag::new(tag), *value))
        .collect::<Vec<_>>();
    let instanced = (!coordinates.is_empty())
        .then(|| {
            subsetter::subset_with_variations(
                &font_data.data,
                font_data.face_index,
                &coordinates,
                &usage.remapper,
            )
            .ok()
        })
        .flatten();
    let subset_bytes = match instanced {
        Some(bytes) => bytes,
        None => subsetter::subset(&font_data.data, font_data.face_index, &usage.remapper).ok()?,
    };

    // Build ToUnicode CMap
    let mut cmap = UnicodeCmap::new(
        Name(b"Adobe"),
        SystemInfo {
            registry: Str(b"Adobe"),
            ordering: Str(b"Identity"),
            supplement: 0,
        },
    );
    for (&old_gid, (_, text)) in &usage.glyph_to_unicode {
        if let Some(new_gid) = usage.remapper.get(old_gid) {
            cmap.pair_with_multiple(new_gid, text.chars());
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

    let Some(face) = instanced_face(font_data) else {
        return widths;
    };

    let units_per_em = face.units_per_em() as f64;
    let scale = 1000.0 / units_per_em;

    for &old_gid in &usage.drawn_glyphs {
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
    let face = instanced_face(font_data)?;
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

#[cfg(test)]
mod tests {
    use super::*;
    use oxml_layout::{
        Color, FontManager, GlyphRun, MultilingualGlyphRun, PageFrame, Point, TextDirection,
        TextSegment,
    };

    /// The bfchar entries of a ToUnicode CMap, by subset glyph id.
    fn to_unicode_entries(cmap: &[u8]) -> HashMap<u16, String> {
        let cmap = std::str::from_utf8(cmap).expect("the CMap is ASCII");
        let mut entries = HashMap::new();
        let mut in_bfchar = false;
        for line in cmap.lines() {
            if line.ends_with("beginbfchar") {
                in_bfchar = true;
            } else if line == "endbfchar" {
                in_bfchar = false;
            } else if in_bfchar {
                let (glyph, text) = line.split_once(' ').expect("one bfchar pair per line");
                let glyph = u16::from_str_radix(glyph.trim_matches(['<', '>']), 16).unwrap();
                let units = text.trim_matches(['<', '>']).as_bytes().chunks(4);
                let units = units
                    .map(|unit| {
                        u16::from_str_radix(std::str::from_utf8(unit).unwrap(), 16).unwrap()
                    })
                    .collect::<Vec<_>>();
                entries.insert(glyph, String::from_utf16(&units).unwrap());
            }
        }
        entries
    }

    fn plain_run(fonts: &FontManager, font_id: FontId, text: &str) -> PositionedElement {
        let shaped = fonts.shape_text(font_id, text, 11.0).unwrap();
        PositionedElement::Text(GlyphRun {
            origin: Point { x: 0.0, y: 0.0 },
            font_id,
            font_size: 11.0,
            glyph_ids: shaped.glyph_ids,
            advances: shaped.advances,
            text: text.to_owned(),
            source: None,
            color: Color::BLACK,
            bold: false,
            italic: false,
            field_kind: None,
            field_source: None,
            tab_aligned: None,
            note: None,
            note_reference_source: None,
        })
    }

    fn multilingual_runs(fonts: &mut FontManager, text: &str) -> Vec<PositionedElement> {
        let font_id = fonts.resolve_font(Some("Calibri"), false, false).unwrap();
        let segment = TextSegment {
            text: text.to_owned(),
            direction: TextDirection::Auto,
            source: None,
            font_id,
            font_size: 11.0,
            glyph_ids: Vec::new(),
            advances: Vec::new(),
            width: 0.0,
            ascent: 0.0,
            descent: 0.0,
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
            field_source: None,
            note: None,
            note_reference_source: None,
        };
        fonts
            .shape_multilingual_text(segment, None, TextDirection::Auto, false)
            .unwrap()
            .into_iter()
            .map(|span| {
                PositionedElement::MultilingualText(MultilingualGlyphRun {
                    origin: Point { x: 0.0, y: 0.0 },
                    font_id: span.font_id(),
                    font_size: 11.0,
                    glyph_ids: span.glyph_ids().to_vec(),
                    x_advances: span.x_advances().to_vec(),
                    y_advances: span.y_advances().to_vec(),
                    x_offsets: span.x_offsets().to_vec(),
                    y_offsets: span.y_offsets().to_vec(),
                    clusters: span.clusters().to_vec(),
                    logical_text: span.text().to_owned(),
                    logical_index: span.logical_index(),
                    source: None,
                    script: span.script(),
                    language: None,
                    direction: span.direction(),
                    bidi_level: span.bidi_level(),
                    color: Color::BLACK,
                    bold: false,
                    italic: false,
                    field_kind: None,
                    field_source: None,
                    note: None,
                    note_reference_source: None,
                })
            })
            .collect()
    }

    fn layout_of(fonts: &FontManager, elements: Vec<PositionedElement>) -> LayoutResult {
        LayoutResult::new(
            vec![PageFrame::new(1, 612.0, 792.0, elements).into()],
            fonts.all_font_data(),
            None,
            vec![],
        )
    }

    /// Each font's subset remapper and ToUnicode entries, as a reader sees them.
    fn embedded_text_maps(
        layout: &LayoutResult,
    ) -> HashMap<FontId, (GlyphRemapper, HashMap<u16, String>)> {
        let mut usage = collect_glyph_usage(layout);
        layout
            .fonts
            .iter()
            .filter_map(|font| {
                let prepared = prepare_font(font, usage.get_mut(&font.id)?)?;
                let entries = to_unicode_entries(&prepared.cmap_bytes);
                Some((font.id, (prepared.remapper, entries)))
            })
            .collect()
    }

    /// The text a reader extracts from each glyph through the ToUnicode CMap.
    fn glyph_texts(
        maps: &HashMap<FontId, (GlyphRemapper, HashMap<u16, String>)>,
        font_id: FontId,
        glyph_ids: &[u16],
    ) -> Vec<String> {
        let (remapper, entries) = &maps[&font_id];
        glyph_ids
            .iter()
            .map(|glyph| {
                let subset_glyph = remapper.get(*glyph).expect("every drawn glyph is subset");
                entries.get(&subset_glyph).cloned().unwrap_or_default()
            })
            .collect()
    }

    /// Issue 297: a bold run of a variable caller face embeds the static
    /// instance layout shaped, with the widths layout measured.
    #[test]
    fn a_variable_caller_face_embeds_the_instance_layout_measured() {
        let variable = include_bytes!("../../oxml-layout/fonts/NotoSansSC-FX058-subset.ttf");
        let mut fonts = FontManager::new_deterministic().unwrap();
        fonts.load_additional_fonts(&[oxml_layout::FontFile {
            family: "Caller Variable".to_owned(),
            data: variable.to_vec(),
        }]);
        let regular = fonts
            .resolve_font(Some("Caller Variable"), false, false)
            .unwrap();
        let bold = fonts
            .resolve_font(Some("Caller Variable"), true, false)
            .unwrap();
        let runs = vec![
            plain_run(&fonts, regular, "Hill"),
            plain_run(&fonts, bold, "Hill"),
        ];
        let layout = layout_of(&fonts, runs.clone());
        let mut usage = collect_glyph_usage(&layout);
        let l_glyph = ttf_parser::Face::parse(variable, 0)
            .unwrap()
            .glyph_index('l')
            .unwrap()
            .0;
        let mut stems = Vec::new();
        for (font_id, run) in [(regular, &runs[0]), (bold, &runs[1])] {
            let PositionedElement::Text(run) = run else {
                unreachable!("plain runs are glyph runs");
            };
            let font = layout.fonts.iter().find(|font| font.id == font_id).unwrap();
            let prepared = prepare_font(font, usage.get_mut(&font_id).unwrap()).unwrap();
            let subset = ttf_parser::Face::parse(&prepared.subset_bytes, 0).unwrap();
            assert!(
                !subset.is_variable(),
                "the embedded face is a static instance"
            );
            for (glyph, advance) in run.glyph_ids.iter().zip(&run.advances) {
                let subset_glyph = prepared.remapper.get(*glyph).unwrap();
                let (_, width) = prepared
                    .widths
                    .iter()
                    .find(|(glyph, _)| *glyph == subset_glyph)
                    .unwrap();
                assert!((width * 11.0 / 1000.0 - advance).abs() < 0.01);
            }
            let l = ttf_parser::GlyphId(prepared.remapper.get(l_glyph).unwrap());
            stems.push(subset.glyph_bounding_box(l).unwrap().width());
        }
        assert!(stems[1] > stems[0], "the bold stem is wider: {stems:?}");
    }

    #[test]
    fn a_ligature_glyph_maps_to_every_character_it_draws() {
        let mut fonts = FontManager::new_deterministic().unwrap();
        for bold in [false, true] {
            let font_id = fonts.resolve_font(Some("Calibri"), bold, false).unwrap();
            // Ligature runs come first. Pairing glyphs and characters by index
            // shifted every later glyph of such a run, and the font-wide map
            // then kept that wrong character in the runs without a ligature.
            let words = [
                "Location ",
                "fifteen ",
                "office ",
                "attitude ",
                "affluent ",
                "fjord ",
                "staff ",
                "shifting ",
                "Rating ",
                "Action ",
                "Observation ",
                "no ligature here",
            ];
            let layout = layout_of(
                &fonts,
                words
                    .iter()
                    .map(|word| plain_run(&fonts, font_id, word))
                    .collect(),
            );
            let maps = embedded_text_maps(&layout);
            for element in &layout.pages[0].elements {
                let PositionedElement::Text(run) = element else {
                    unreachable!("the page holds plain runs only");
                };
                assert_eq!(
                    glyph_texts(&maps, font_id, &run.glyph_ids).concat(),
                    run.text,
                    "bold: {bold}"
                );
            }
            for ligature in ["ti", "fi", "ft", "ffi", "ffl", "tti", "fj", "ff"] {
                let shaped = fonts.shape_text(font_id, ligature, 11.0).unwrap();
                assert_eq!(shaped.glyph_ids.len(), 1, "Carlito joins {ligature}");
                assert_eq!(
                    glyph_texts(&maps, font_id, &shaped.glyph_ids),
                    [ligature],
                    "bold: {bold}"
                );
            }
        }
    }

    #[test]
    fn a_ligature_outranks_the_presentation_form_sharing_its_glyph() {
        let mut fonts = FontManager::new_deterministic().unwrap();
        let font_id = fonts.resolve_font(Some("Calibri"), false, false).unwrap();
        // Carlito draws its fi ligature with the cmap glyph of U+FB01, and the
        // CMap keeps one text per glyph. A literal U+FB01 seen first used to
        // claim that glyph and turn every "fi" of the document into U+FB01.
        let fi = fonts.shape_text(font_id, "fi", 11.0).unwrap().glyph_ids;
        let form = fonts.shape_text(font_id, "\u{fb01}", 11.0).unwrap();
        assert_eq!(form.glyph_ids, fi, "Carlito shares the glyph");
        let extracted = |texts: &[&str]| {
            let runs = texts.iter().map(|text| plain_run(&fonts, font_id, text));
            let layout = layout_of(&fonts, runs.collect());
            let maps = embedded_text_maps(&layout);
            let runs = layout.pages[0].elements.iter().map(|element| {
                let PositionedElement::Text(run) = element else {
                    unreachable!("the page holds plain runs only");
                };
                glyph_texts(&maps, font_id, &run.glyph_ids).concat()
            });
            runs.collect::<Vec<_>>()
        };
        assert_eq!(extracted(&["office fifteen"]), ["office fifteen"]);
        assert_eq!(
            extracted(&["\u{fb01}", "office fifteen"]),
            ["fi", "office fifteen"]
        );
        // Without a ligature to decompose, the glyph keeps the literal.
        assert_eq!(extracted(&["\u{fb01}"]), ["\u{fb01}"]);
    }

    #[test]
    fn a_plain_run_extracts_each_character_once() {
        let mut fonts = FontManager::new_deterministic().unwrap();
        // Liberation Sans has no "≮" or "≯" and Carlito no "Ѷ", so the shaper
        // draws each as a base glyph and a combining mark, neither of them the
        // cmap glyph of the character. A plain run has no ActualText, so the
        // mark must not repeat the character its base glyph carries.
        let mut elements = Vec::new();
        for (family, text) in [("Arial", "a ≮ b ≯ c"), ("Calibri", "Ѷ")] {
            let font_id = fonts.resolve_font(Some(family), false, false).unwrap();
            elements.push(plain_run(&fonts, font_id, text));
        }
        let layout = layout_of(&fonts, elements);
        let maps = embedded_text_maps(&layout);
        for element in &layout.pages[0].elements {
            let PositionedElement::Text(run) = element else {
                unreachable!("the page holds plain runs only");
            };
            let characters = run.text.chars().count();
            assert!(
                run.glyph_ids.len() > characters,
                "{:?} draws a mark",
                run.text
            );
            let texts = glyph_texts(&maps, run.font_id, &run.glyph_ids);
            assert_eq!(texts.concat(), run.text);
        }
    }

    #[test]
    fn a_cmap_pairing_outranks_one_inferred_earlier() {
        let mut fonts = FontManager::new_deterministic().unwrap();
        let font_id = fonts.resolve_font(Some("Calibri"), false, false).unwrap();
        // Carlito has no no-break space, so the shaper draws the space glyph
        // for it. Seen first, that pairing used to claim the glyph for the
        // whole font.
        let layout = layout_of(
            &fonts,
            vec![
                plain_run(&fonts, font_id, "a\u{a0}b"),
                plain_run(&fonts, font_id, "a b"),
            ],
        );
        let space = fonts.shape_text(font_id, " ", 11.0).unwrap().glyph_ids;
        let texts = glyph_texts(&embedded_text_maps(&layout), font_id, &space);
        assert_eq!(texts, [" "]);
    }

    #[test]
    fn a_multilingual_cluster_keeps_every_character() {
        let mut fonts = FontManager::new_deterministic().unwrap();
        // The conjunct of "क्ष" draws three characters with one glyph. The
        // cluster "त्रि" draws four with two, its vowel sign first and in a
        // width variant the cmap does not give, and "कि" then pairs that
        // variant with the vowel sign alone. The Arabic font draws the dot of
        // "ب" apart from the letter, and draws contextual forms.
        let mut elements = multilingual_runs(&mut fonts, "क्षत्रिय कि");
        elements.extend(multilingual_runs(&mut fonts, "سلام ب"));
        let layout = layout_of(&fonts, elements);
        let maps = embedded_text_maps(&layout);
        let mut clusters = Vec::new();
        for element in &layout.pages[0].elements {
            let PositionedElement::MultilingualText(run) = element else {
                unreachable!("the page holds multilingual runs only");
            };
            let texts = glyph_texts(&maps, run.font_id, &run.glyph_ids);
            let characters = run.logical_text.chars().collect::<Vec<_>>();
            for cluster in &run.clusters {
                let glyphs = cluster.glyph_start as usize..cluster.glyph_end as usize;
                let source = &characters[cluster.char_start as usize..cluster.char_end as usize];
                let texts = texts[glyphs].to_vec();
                assert!(texts.iter().all(|text| !text.is_empty()), "{texts:?}");
                let mut extracted = texts.concat();
                for character in source {
                    let at = extracted.find(*character).unwrap_or_else(|| {
                        panic!("{character:?} of {source:?} is missing from {texts:?}")
                    });
                    extracted.remove(at);
                }
                clusters.push((source.iter().collect::<String>(), texts));
            }
        }
        let texts_of = |source: &str| {
            clusters
                .iter()
                .find(|(cluster, _)| cluster == source)
                .map(|(_, texts)| texts.clone())
                .unwrap_or_else(|| panic!("no cluster {source:?} in {clusters:?}"))
        };
        assert_eq!(texts_of("क्ष"), ["क्ष"]);
        assert_eq!(texts_of("त्रि"), ["ि", "त्र"]);
        assert_eq!(texts_of("कि"), ["ि", "क"]);
        assert_eq!(texts_of("ب"), ["ب", "ب"]);
    }

    #[test]
    fn a_run_without_ligatures_keeps_one_character_per_glyph() {
        let mut fonts = FontManager::new_deterministic().unwrap();
        let font_id = fonts.resolve_font(Some("Arial"), false, false).unwrap();
        let text = "Location Rating Action fifteen office";
        let layout = layout_of(&fonts, vec![plain_run(&fonts, font_id, text)]);
        let PositionedElement::Text(run) = &layout.pages[0].elements[0] else {
            unreachable!("the page holds one plain run");
        };
        assert_eq!(run.glyph_ids.len(), text.chars().count());
        let maps = embedded_text_maps(&layout);
        let texts = glyph_texts(&maps, font_id, &run.glyph_ids);
        let characters = text.chars().map(String::from).collect::<Vec<_>>();
        assert_eq!(texts, characters);

        let mut usage = collect_glyph_usage(&layout);
        let font = layout.fonts.iter().find(|font| font.id == font_id).unwrap();
        let prepared = prepare_font(font, usage.get_mut(&font_id).unwrap()).unwrap();
        let mut drawn = run.glyph_ids.clone();
        drawn.sort_unstable();
        drawn.dedup();
        assert_eq!(prepared.widths.len(), drawn.len());
    }
}
