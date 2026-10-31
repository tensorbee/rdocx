//! Font loading, resolution, shaping, and metrics.
//!
//! Uses fontdb for system font discovery, ttf-parser for metrics,
//! and HarfRust for text shaping.

use std::collections::{HashMap, HashSet, VecDeque};
use std::ops::Range;
#[cfg(feature = "system-fonts")]
use std::path::{Path, PathBuf};
#[cfg(feature = "system-fonts")]
use std::sync::OnceLock;
use std::sync::{Arc, Mutex};

use std::sync::atomic::{AtomicUsize, Ordering};

use crate::error::{LayoutError, Result};
use crate::line::TextSegment;
use crate::output::{Diagnostic, FontId, SourceSpan};

/// Font data provided by the user or extracted from an OOXML file.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FontFile {
    /// Font family name (e.g., "Calibri", "Arial").
    pub family: String,
    /// Raw font file bytes (TTF/OTF).
    pub data: Vec<u8>,
}

impl FontFile {
    /// Read supported caller font files in lexical path order.
    /// A missing directory, unreadable entry or malformed supported font fails
    /// instead of silently changing the requested fallback configuration.
    pub fn from_dir(path: &std::path::Path) -> Result<Vec<Self>> {
        let failed = |error: std::io::Error| {
            LayoutError::FontParse(format!("font directory {}: {error}", path.display()))
        };
        if !std::fs::metadata(path).map_err(failed)?.is_dir() {
            return Err(LayoutError::FontParse(format!(
                "font directory {} is not a directory",
                path.display()
            )));
        }
        let mut paths = std::fs::read_dir(path)
            .map_err(failed)?
            .map(|entry| entry.map(|entry| entry.path()).map_err(failed))
            .collect::<Result<Vec<_>>>()?;
        paths.sort();
        let mut fonts = Vec::new();
        for entry in paths {
            let supported = entry
                .extension()
                .and_then(|ext| ext.to_str())
                .is_some_and(|ext| {
                    ["ttf", "otf", "ttc"]
                        .iter()
                        .any(|kind| ext.eq_ignore_ascii_case(kind))
                });
            if !supported {
                continue;
            }
            let data = std::fs::read(&entry).map_err(|error| {
                LayoutError::FontParse(format!("font file {}: {error}", entry.display()))
            })?;
            ttf_parser::Face::parse(&data, 0).map_err(|error| {
                LayoutError::FontParse(format!("font file {}: {error}", entry.display()))
            })?;
            let family = entry
                .file_stem()
                .and_then(|stem| stem.to_str())
                .ok_or_else(|| {
                    LayoutError::FontParse(format!(
                        "font filename {} is not UTF-8",
                        entry.display()
                    ))
                })?
                .to_owned();
            fonts.push(Self { family, data });
        }
        Ok(fonts)
    }
}

/// Key for caching resolved fonts.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
struct FontKey {
    family: String,
    bold: bool,
    italic: bool,
}

/// Metrics for a font at a given size.
#[derive(Debug, Clone, Copy)]
pub struct FontMetrics {
    /// Ascent in points (positive, above baseline).
    pub ascent: f64,
    /// Descent in points (positive, below baseline).
    pub descent: f64,
    /// Line gap in points.
    pub line_gap: f64,
    /// Units per em.
    pub units_per_em: u16,
}

/// Result of shaping a text string.
#[derive(Debug, Clone)]
pub struct ShapedText {
    /// Glyph IDs from shaping.
    pub glyph_ids: Vec<u16>,
    /// Per-glyph advances in points.
    pub advances: Vec<f64>,
    /// Total width in points.
    pub width: f64,
}

/// Requested or resolved direction for one logical text span.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum TextDirection {
    #[default]
    Auto,
    LeftToRight,
    RightToLeft,
}

pub(crate) fn directional_level(
    direction: TextDirection,
    paragraph_level: unicode_bidi::Level,
) -> unicode_bidi::Level {
    let paragraph_number = paragraph_level.number();
    let number = match direction {
        TextDirection::Auto => paragraph_number,
        TextDirection::LeftToRight if paragraph_level.is_rtl() => paragraph_number + 1,
        TextDirection::LeftToRight => paragraph_number,
        TextDirection::RightToLeft if paragraph_level.is_rtl() => paragraph_number,
        TextDirection::RightToLeft => paragraph_number + 1,
    };
    unicode_bidi::Level::new(number).expect("one directional embedding fits the bidi level limit")
}

pub(crate) fn explicit_direction_levels(
    text: &str,
    direction: TextDirection,
    paragraph_level: unicode_bidi::Level,
) -> Result<Vec<unicode_bidi::Level>> {
    let local_base = match direction {
        TextDirection::Auto => paragraph_level,
        TextDirection::LeftToRight => unicode_bidi::Level::ltr(),
        TextDirection::RightToLeft => unicode_bidi::Level::rtl(),
    };
    let target_base = directional_level(direction, paragraph_level);
    let offset = target_base.number() - local_base.number();
    unicode_bidi::BidiInfo::new(text, Some(local_base))
        .levels
        .into_iter()
        .map(|level| {
            level
                .number()
                .checked_add(offset)
                .and_then(|number| unicode_bidi::Level::new(number).ok())
                .ok_or_else(|| {
                    LayoutError::Layout(
                        "run direction exceeded the Unicode bidi level limit".to_owned(),
                    )
                })
        })
        .collect()
}

/// Script identity used to select shaping behavior and deterministic fallback.
///
/// This enum keeps growing as scripts are added, and every addition is a
/// breaking change for a downstream exhaustive match. `#[non_exhaustive]`
/// takes that break once, at the pre-1.0 moment it is cheapest.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum TextScript {
    Latin,
    Arabic,
    Hebrew,
    Devanagari,
    Thai,
    Han,
    /// Hangul jamo, compatibility jamo and precomposed syllables.
    Hangul,
    /// Hiragana, Katakana and the Katakana phonetic extensions.
    Kana,
    Common,
}

/// One glyph interval mapped to an exclusive logical Unicode-scalar interval.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct GlyphCluster {
    pub glyph_start: u32,
    pub glyph_end: u32,
    pub char_start: u32,
    pub char_end: u32,
}

impl GlyphCluster {
    pub fn is_valid(&self) -> bool {
        self.glyph_start < self.glyph_end && self.char_start < self.char_end
    }

    pub fn char_range(&self) -> Range<u32> {
        self.char_start..self.char_end
    }
}

/// One script, font, and bidi-level span shaped with complete positioning data.
#[derive(Debug, Clone)]
pub struct MultilingualTextSegment {
    base: TextSegment,
    logical_index: usize,
    language: Option<String>,
    script: TextScript,
    direction: TextDirection,
    bidi_level: u8,
    x_advances: Vec<f64>,
    y_advances: Vec<f64>,
    x_offsets: Vec<f64>,
    y_offsets: Vec<f64>,
    clusters: Vec<GlyphCluster>,
    break_after: bool,
    kerning: Option<bool>,
    legacy_shaping: bool,
    hyphenation_language: Option<String>,
}

impl MultilingualTextSegment {
    #[allow(clippy::too_many_arguments)]
    pub fn new(
        base: TextSegment,
        logical_index: usize,
        language: Option<String>,
        script: TextScript,
        direction: TextDirection,
        bidi_level: u8,
        x_advances: Vec<f64>,
        y_advances: Vec<f64>,
        x_offsets: Vec<f64>,
        y_offsets: Vec<f64>,
        clusters: Vec<GlyphCluster>,
        break_after: bool,
    ) -> Result<Self> {
        Self::new_with_legacy_shaping(
            base,
            logical_index,
            language,
            script,
            direction,
            bidi_level,
            x_advances,
            y_advances,
            x_offsets,
            y_offsets,
            clusters,
            break_after,
            false,
        )
    }

    // Legacy guess-property shaping may order glyphs independently of layout
    // bidi levels. Validate full cluster coverage in its actual monotone order.
    // Public construction and genuine rich callers retain strict parity.
    #[allow(clippy::too_many_arguments)]
    pub(crate) fn new_with_legacy_shaping(
        base: TextSegment,
        logical_index: usize,
        language: Option<String>,
        script: TextScript,
        direction: TextDirection,
        bidi_level: u8,
        x_advances: Vec<f64>,
        y_advances: Vec<f64>,
        x_offsets: Vec<f64>,
        y_offsets: Vec<f64>,
        clusters: Vec<GlyphCluster>,
        break_after: bool,
        legacy_shaping: bool,
    ) -> Result<Self> {
        let glyph_count = base.glyph_ids.len();
        let char_count = base.text.chars().count();
        let valid_level = unicode_bidi::Level::new(bidi_level).is_ok();
        if x_advances.len() != glyph_count
            || y_advances.len() != glyph_count
            || x_offsets.len() != glyph_count
            || y_offsets.len() != glyph_count
            || base.advances.len() != glyph_count
            || !valid_level
            || !position_values_are_finite(
                &x_advances,
                &y_advances,
                &x_offsets,
                &y_offsets,
                base.width,
            )
            || !(cluster_ranges_are_valid(&clusters, glyph_count, char_count, bidi_level % 2 == 1)
                || (legacy_shaping
                    && cluster_ranges_are_valid(
                        &clusters,
                        glyph_count,
                        char_count,
                        bidi_level.is_multiple_of(2),
                    )))
        {
            return Err(LayoutError::Layout(
                "invalid multilingual glyph positioning or cluster range".to_owned(),
            ));
        }
        let mut span = Self {
            base,
            logical_index,
            language,
            script,
            direction,
            bidi_level,
            x_advances,
            y_advances,
            x_offsets,
            y_offsets,
            clusters,
            break_after,
            kerning: None,
            legacy_shaping,
            hyphenation_language: None,
        };
        if legacy_shaping {
            span.normalize_legacy_storage();
        }
        Ok(span)
    }

    /// Replace only compatible base/advance data while retaining shaping context.
    #[doc(hidden)]
    pub fn with_base_and_advances(
        mut self,
        base: TextSegment,
        x_advances: Vec<f64>,
    ) -> Result<Self> {
        if base.text != self.base.text
            || base.font_id != self.base.font_id
            || base.font_size.to_bits() != self.base.font_size.to_bits()
            || base.direction != self.base.direction
            || base.glyph_ids != self.base.glyph_ids
            || base.advances.len() != base.glyph_ids.len()
            || x_advances.len() != base.glyph_ids.len()
            || !base.advances.iter().all(|value| value.is_finite())
            || !position_values_are_finite(
                &x_advances,
                &self.y_advances,
                &self.x_offsets,
                &self.y_offsets,
                base.width,
            )
            || !(cluster_ranges_are_valid(
                &self.clusters,
                base.glyph_ids.len(),
                base.text.chars().count(),
                self.bidi_level % 2 == 1,
            ) || (self.legacy_shaping
                && cluster_ranges_are_valid(
                    &self.clusters,
                    base.glyph_ids.len(),
                    base.text.chars().count(),
                    self.bidi_level.is_multiple_of(2),
                )))
        {
            return Err(LayoutError::Layout(
                "incompatible multilingual base or advance replacement".into(),
            ));
        }
        self.base = base;
        self.x_advances = x_advances;
        if self.legacy_shaping {
            self.normalize_legacy_storage();
        }
        Ok(self)
    }

    #[doc(hidden)]
    pub fn with_hyphenation_language(mut self, language: Option<String>) -> Self {
        self.hyphenation_language = language;
        if self.legacy_shaping {
            self.normalize_legacy_storage();
        }
        self
    }

    /// Preserve the shaped payload while assigning its actual assembly order
    /// and the source producer's split, reflow and paint convention.
    #[doc(hidden)]
    pub fn with_logical_index_and_legacy_shaping(
        mut self,
        logical_index: usize,
        legacy_shaping: bool,
    ) -> Self {
        self.logical_index = logical_index;
        self.legacy_shaping = legacy_shaping;
        if legacy_shaping {
            self.normalize_legacy_storage();
        }
        self
    }

    #[doc(hidden)]
    pub fn uses_legacy_shaping(&self) -> bool {
        self.legacy_shaping
    }

    #[doc(hidden)]
    pub fn kerning_override(&self) -> Option<bool> {
        self.kerning
    }

    pub(crate) fn hyphenation_language(&self) -> Option<&str> {
        self.hyphenation_language.as_deref()
    }

    pub(crate) fn inherit_shaping_context(mut self, source: &Self) -> Self {
        self.kerning = source.kerning;
        self.legacy_shaping = source.legacy_shaping;
        self.hyphenation_language = source.hyphenation_language.clone();
        if self.legacy_shaping {
            self.normalize_legacy_storage();
        }
        self
    }

    fn normalize_legacy_storage(&mut self) {
        self.x_advances = std::mem::take(&mut self.x_advances)
            .into_boxed_slice()
            .into_vec();
        self.y_advances = std::mem::take(&mut self.y_advances)
            .into_boxed_slice()
            .into_vec();
        self.x_offsets = std::mem::take(&mut self.x_offsets)
            .into_boxed_slice()
            .into_vec();
        self.y_offsets = std::mem::take(&mut self.y_offsets)
            .into_boxed_slice()
            .into_vec();
        self.clusters = std::mem::take(&mut self.clusters)
            .into_boxed_slice()
            .into_vec();
        self.language = self
            .language
            .take()
            .map(|value| value.into_boxed_str().into_string());
        self.hyphenation_language = self
            .hyphenation_language
            .take()
            .map(|value| value.into_boxed_str().into_string());
    }

    pub fn text(&self) -> &str {
        &self.base.text
    }

    pub fn base(&self) -> &TextSegment {
        &self.base
    }

    pub fn font_id(&self) -> FontId {
        self.base.font_id
    }

    pub fn language(&self) -> Option<&str> {
        self.language.as_deref()
    }

    pub fn script(&self) -> TextScript {
        self.script
    }

    pub fn direction(&self) -> TextDirection {
        self.direction
    }

    pub fn bidi_level(&self) -> u8 {
        self.bidi_level
    }

    pub fn logical_index(&self) -> usize {
        self.logical_index
    }

    pub fn glyph_ids(&self) -> &[u16] {
        &self.base.glyph_ids
    }

    pub fn x_advances(&self) -> &[f64] {
        &self.x_advances
    }

    pub fn y_advances(&self) -> &[f64] {
        &self.y_advances
    }

    pub fn x_offsets(&self) -> &[f64] {
        &self.x_offsets
    }

    pub fn y_offsets(&self) -> &[f64] {
        &self.y_offsets
    }

    pub fn clusters(&self) -> &[GlyphCluster] {
        &self.clusters
    }

    pub fn width(&self) -> f64 {
        self.base.width
    }

    pub fn break_after(&self) -> bool {
        self.break_after
    }
}

pub(crate) fn position_values_are_finite(
    x_advances: &[f64],
    y_advances: &[f64],
    x_offsets: &[f64],
    y_offsets: &[f64],
    width: f64,
) -> bool {
    width.is_finite()
        && x_advances
            .iter()
            .chain(y_advances)
            .chain(x_offsets)
            .chain(y_offsets)
            .all(|value| value.is_finite())
}

pub(crate) fn cluster_ranges_are_valid(
    clusters: &[GlyphCluster],
    glyph_count: usize,
    char_count: usize,
    right_to_left: bool,
) -> bool {
    if glyph_count == 0 || char_count == 0 {
        return glyph_count == 0 && char_count == 0 && clusters.is_empty();
    }
    let mut previous_glyph_end = 0u32;
    let mut previous_char_range = None::<Range<u32>>;
    let mut covered_chars = 0usize;
    for cluster in clusters {
        if !cluster.is_valid()
            || cluster.glyph_start != previous_glyph_end
            || cluster.glyph_end as usize > glyph_count
            || cluster.char_end as usize > char_count
        {
            return false;
        }
        if let Some(previous) = previous_char_range {
            let contiguous = if right_to_left {
                cluster.char_end == previous.start
            } else {
                cluster.char_start == previous.end
            };
            if !contiguous {
                return false;
            }
        }
        previous_glyph_end = cluster.glyph_end;
        previous_char_range = Some(cluster.char_range());
        covered_chars += (cluster.char_end - cluster.char_start) as usize;
    }
    previous_glyph_end as usize == glyph_count
        && covered_chars == char_count
        && previous_char_range.is_some_and(|last| {
            if right_to_left {
                clusters
                    .first()
                    .is_some_and(|first| first.char_end as usize == char_count)
                    && last.start == 0
            } else {
                clusters.first().is_some_and(|first| first.char_start == 0)
                    && last.end as usize == char_count
            }
        })
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct ShapingKey {
    font_id: FontId,
    text: String,
    size_bits: u64,
    kerning: Option<bool>,
}

struct ShapingMemo {
    entries: VecDeque<(ShapingKey, ShapedText, usize, u64)>,
    bytes: usize,
    #[cfg(test)]
    hits: usize,
    #[cfg(test)]
    misses: usize,
}

impl ShapingMemo {
    fn new() -> Self {
        Self {
            entries: VecDeque::new(),
            bytes: 0,
            #[cfg(test)]
            hits: 0,
            #[cfg(test)]
            misses: 0,
        }
    }

    fn clear(&mut self) {
        self.entries.clear();
        self.bytes = 0;
        #[cfg(test)]
        {
            self.hits = 0;
            self.misses = 0;
        }
    }

    fn insert(&mut self, key: ShapingKey, shaped: ShapedText) {
        let entry_bytes = std::mem::size_of::<(ShapingKey, ShapedText, usize, u64)>()
            + key.text.len()
            + shaped.glyph_ids.len() * std::mem::size_of::<u16>()
            + shaped.advances.len() * std::mem::size_of::<f64>();
        if entry_bytes > SHAPING_CACHE_MAX_BYTES {
            return;
        }
        while self.entries.len() >= SHAPING_CACHE_MAX_ENTRIES
            || self.bytes.saturating_add(entry_bytes) > SHAPING_CACHE_MAX_BYTES
        {
            let Some((_, _, evicted_bytes, _)) = self.entries.pop_front() else {
                break;
            };
            self.bytes = self.bytes.saturating_sub(evicted_bytes);
        }
        self.bytes += entry_bytes;
        let fingerprint = shaping_fingerprint(key.font_id, &key.text, key.size_bits, key.kerning);
        self.entries
            .push_back((key, shaped, entry_bytes, fingerprint));
    }
}

fn shaping_fingerprint(font_id: FontId, text: &str, size_bits: u64, kerning: Option<bool>) -> u64 {
    use std::hash::{Hash, Hasher};

    let mut hasher = std::collections::hash_map::DefaultHasher::new();
    font_id.hash(&mut hasher);
    text.hash(&mut hasher);
    size_bits.hash(&mut hasher);
    kerning.hash(&mut hasher);
    hasher.finish()
}

const SHAPING_CACHE_MAX_ENTRIES: usize = 2_048;
const SHAPING_CACHE_MAX_BYTES: usize = 16 * 1024 * 1024;

#[cfg(feature = "system-fonts")]
struct FileFontCache {
    entries: VecDeque<(PathBuf, Arc<[u8]>, usize)>,
    bytes: usize,
}

#[cfg(feature = "system-fonts")]
impl FileFontCache {
    fn new() -> Self {
        Self {
            entries: VecDeque::new(),
            bytes: 0,
        }
    }

    fn clear(&mut self) {
        self.entries.clear();
        self.bytes = 0;
    }
}

#[cfg(feature = "system-fonts")]
const FILE_FONT_CACHE_MAX_ENTRIES: usize = 256;
#[cfg(feature = "system-fonts")]
const FILE_FONT_CACHE_MAX_BYTES: usize = 128 * 1024 * 1024;

#[cfg(feature = "system-fonts")]
static NORMAL_FONT_DATABASE: OnceLock<fontdb::Database> = OnceLock::new();
#[cfg(feature = "system-fonts")]
static FILE_FONT_CACHE: OnceLock<Mutex<FileFontCache>> = OnceLock::new();
#[cfg(all(test, feature = "system-fonts"))]
static SYSTEM_FONT_DISCOVERY_RUNS: AtomicUsize = AtomicUsize::new(0);

/// Internal record for a loaded font face.
struct LoadedFont {
    db_id: fontdb::ID,
    id: FontId,
    family: String,
    bold: bool,
    italic: bool,
    data: Arc<[u8]>,
    face_index: u32,
    units_per_em: u16,
    /// Vertical metrics in design units, read once when the face is loaded.
    ascender: i16,
    descender: i16,
    line_gap: i16,
    /// The line metrics Word measures this face on, in design units, or
    /// `None` for a face without Windows metrics. See
    /// [`FontManager::word_line_metrics`].
    word_line: Option<WordLineUnits>,
    /// HarfRust's per-face shaping caches. Building these is the expensive
    /// part of shaping, so it happens once per face instead of once per run.
    shaper_data: harfrust::ShaperData,
    variations: Vec<([u8; 4], f32)>,
    shaper_instance: Option<harfrust::ShaperInstance>,
    synthetic_bold: bool,
    synthetic_italic: bool,
}

/// Preserve defaults unless a style was explicitly requested. Keep restricted
/// instances distinct from synthesis rather than claiming an unavailable style.
fn face_style_adjustment(
    face: &ttf_parser::Face<'_>,
    bold: bool,
    italic: bool,
) -> (Vec<([u8; 4], f32)>, bool, bool) {
    let axis = |tag: &[u8; 4]| {
        face.variation_axes()
            .into_iter()
            .find(|axis| axis.tag == ttf_parser::Tag::from_bytes(tag))
    };
    let mut variations = Vec::new();
    let mut select = |tag, axis: ttf_parser::VariationAxis, value: f32| {
        let selected = value.clamp(axis.min_value, axis.max_value);
        if selected != axis.def_value {
            variations.push((tag, selected));
        }
        selected
    };
    let genuine_bold = if bold {
        match axis(b"wght") {
            Some(weight) => select(*b"wght", weight, 700.0) >= 700.0,
            None => face.is_bold() || face.weight().to_number() >= 700,
        }
    } else {
        false
    };
    let mut genuine_italic = face.style() != ttf_parser::Style::Normal;
    if italic && !genuine_italic {
        if let Some(axis) =
            axis(b"ital").filter(|axis| axis.min_value <= 1.0 && axis.max_value >= 1.0)
        {
            select(*b"ital", axis, 1.0);
            genuine_italic = true;
        } else if let Some(axis) = axis(b"slnt").filter(|axis| axis.min_value < 0.0) {
            genuine_italic = select(*b"slnt", axis, -12.0) < 0.0;
        }
    }
    (variations, bold && !genuine_bold, italic && !genuine_italic)
}

/// A face's line metrics as Windows reports them to Word, in design units.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct WordLineUnits {
    /// Above the baseline.
    ascent: i32,
    /// Below the baseline, positive.
    descent: i32,
    /// External leading, never negative.
    leading: i32,
    /// Actual OS/2 CJK codepage coverage, not a font-name or script guess.
    cjk: bool,
}

impl WordLineUnits {
    /// Read the Word face policy, or `None` when the existing fallback
    /// must use shared metrics.
    fn of(face: &ttf_parser::Face<'_>) -> Option<Self> {
        let os2 = face.tables().os2?;
        let table = face
            .raw_face()
            .table(ttf_parser::Tag::from_bytes(b"OS/2"))?;
        let hhea = face.tables().hhea;
        // Version 1 added both codepage-range words. A truncated range is
        // not a positive face classification. CP932/936/949/950 occupy bits 17 through 20.
        let cjk = table.len() >= 86
            && u16::from_be_bytes([table[0], table[1]]) >= 1
            && u32::from_be_bytes(table[78..82].try_into().ok()?) & 0x001e0000 != 0;
        if cjk && hhea.ascender > 0 && hhea.descender <= 0 {
            return Some(Self {
                ascent: i32::from(hhea.ascender),
                descent: -i32::from(hhea.descender),
                leading: 0,
                cjk: true,
            });
        }
        if os2.use_typographic_metrics() {
            return Some(Self {
                ascent: i32::from(os2.typographic_ascender()),
                descent: -i32::from(os2.typographic_descender()),
                leading: i32::from(os2.typographic_line_gap()).max(0),
                cjk: false,
            });
        }
        // `usWinAscent` and `usWinDescent` are unsigned. ttf-parser reads
        // them as signed and negates the descent, which overflows at 32768,
        // so the two fields are read here directly.
        let unsigned = |offset: usize| {
            table
                .get(offset..offset + 2)
                .map(|bytes| i32::from(u16::from_be_bytes([bytes[0], bytes[1]])))
        };
        let ascent = unsigned(74)?;
        let descent = unsigned(76)?;
        if ascent + descent == 0 {
            return None;
        }
        let hhea_extent = i32::from(hhea.ascender) - i32::from(hhea.descender);
        Some(Self {
            ascent,
            descent,
            leading: (i32::from(hhea.line_gap) - (ascent + descent - hhea_extent)).max(0),
            cjk: false,
        })
    }
}

/// One successfully shaped nonempty final text contribution.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FontContribution {
    pub font_id: FontId,
    pub requested_family: Option<String>,
}

/// A transactional local capture within the current bounded trace.
#[derive(Debug)]
pub struct FontContributionScope {
    generation: u64,
    event_start: usize,
    depth: usize,
    owns_trace: bool,
    diagnostic_start: usize,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum FontTraceEvent {
    Resolution(FontId),
    Contribution(FontContribution),
}

struct ParagraphFontTrace {
    events: Vec<FontTraceEvent>,
    generation: u64,
    scope_depth: usize,
    owned_family_bytes: usize,
    overflowed: bool,
}

/// Manages font discovery, loading, shaping, and metrics.
pub struct FontManager {
    db: fontdb::Database,
    /// Database before document-embedded or caller fonts are applied.
    base_db: fontdb::Database,
    /// Caller labels whose faces are retained in `base_db`.
    base_caller_aliases: HashMap<String, Vec<fontdb::ID>>,
    /// Caller embedded families whose faces are retained in `base_db`.
    base_caller_families: HashMap<String, Vec<fontdb::ID>>,
    /// Map from FontKey to loaded font info.
    cache: HashMap<FontKey, usize>,
    /// Manager-owned bytes for bundled, embedded, and caller-provided faces.
    memory_face_data: HashMap<fontdb::ID, Arc<[u8]>>,
    /// All loaded fonts.
    fonts: Vec<LoadedFont>,
    /// Next font ID counter.
    next_id: u32,
    /// Fonts already discovered as covering something the requested family
    /// could not, keyed by (bold, italic).
    ///
    /// Finding a font that covers a character means loading and inspecting
    /// faces, which is far too slow to repeat per character. Once a CJK face
    /// has been found for one character it almost always covers the rest of
    /// the run, so it is tried first next time.
    coverage_fallbacks: HashMap<(bool, bool), Vec<usize>>,
    /// Characters already searched for and not found in any available font, so
    /// the scan is not repeated for every occurrence.
    coverage_misses: HashSet<char>,
    /// Exact additional font set currently loaded into `db`.
    additional_fonts: Vec<FontFile>,
    /// One decoded body per physical embedded part and every owned alias.
    embedded_fonts: Vec<(Vec<String>, Arc<[u8]>)>,
    /// Owned embedded typeface aliases, separate from caller provenance.
    embedded_aliases: HashMap<String, Vec<fontdb::ID>>,
    /// Lowercased caller labels mapped to the exact faces loaded from their
    /// bytes. Rebuilt together with `additional_fonts`.
    caller_aliases: HashMap<String, Vec<fontdb::ID>>,
    /// Exact embedded caller families mapped to their loaded faces, so caller
    /// bytes take priority over bundled faces with the same family.
    caller_families: HashMap<String, Vec<fontdb::ID>>,
    /// Exact caller-declared alias identity for cheap change detection.
    explicit_aliases: Vec<(String, String)>,
    /// Lowercased requested family mapped to the caller-declared target.
    explicit_alias_map: HashMap<String, String>,
    /// Bounded exact-key shaping results.
    shaping_memo: Mutex<ShapingMemo>,
    /// Exact resolution events for one cache-candidate paragraph.
    paragraph_font_trace: Option<ParagraphFontTrace>,
    /// Distinct current-layout fonts in first-resolution order.
    layout_fonts: Vec<FontId>,
    // At most one current-output message per Aptos request spelling.
    // Scope capture rolls provisional messages back with their contributions.
    layout_font_diagnostics: Vec<Diagnostic>,
}

/// Families with broad non-Latin coverage, tried before scanning everything.
///
/// Ordered roughly by how likely each is to be installed. This is only a fast
/// path: if none of them is present the full font database is still searched.
const BROAD_COVERAGE_FAMILIES: &[&str] = &[
    // Deterministic complex-script fallbacks bundled by this crate
    "Noto Sans Arabic",
    "Noto Sans Devanagari",
    "Noto Sans Thai",
    "Noto Sans SC",
    // Bundled with or shipped alongside many Linux distributions
    "Noto Sans CJK SC",
    "Noto Sans CJK JP",
    "Noto Sans CJK KR",
    "Noto Sans CJK TC",
    "Noto Serif CJK SC",
    "Source Han Sans SC",
    "WenQuanYi Zen Hei",
    "WenQuanYi Micro Hei",
    // macOS
    "PingFang SC",
    "PingFang TC",
    "Hiragino Sans",
    "Hiragino Kaku Gothic ProN",
    "Apple SD Gothic Neo",
    "Songti SC",
    "STHeiti",
    // Windows
    "Microsoft YaHei",
    "Microsoft JhengHei",
    "SimSun",
    "SimHei",
    "NSimSun",
    "Yu Gothic",
    "MS Gothic",
    "Meiryo",
    "Malgun Gothic",
    // Wide-coverage generalists
    "Arial Unicode MS",
    "DejaVu Sans",
];

const RESOLUTION_CACHE_MAX_ENTRIES: usize = 256;
const COVERAGE_FALLBACK_MAX_ENTRIES: usize = 256;
const COVERAGE_MISS_MAX_ENTRIES: usize = 4_096;
const PARAGRAPH_FONT_TRACE_MAX_ENTRIES: usize = 4_096;
const PARAGRAPH_FONT_TRACE_MAX_BYTES: usize = SHAPING_CACHE_MAX_BYTES;
static NEXT_FONT_TRACE_GENERATION: AtomicUsize = AtomicUsize::new(1);

/// Maximum caller-declared aliases retained for resolution identity.
const CALLER_ALIAS_MAX_ENTRIES: usize = 256;

/// Maximum aggregate UTF-8 payload retained by the ordered caller-alias
/// identity and its lowercased lookup map.
const CALLER_ALIAS_MAX_RETAINED_BYTES: usize = 64 * 1024;

/// Return the deterministic prefix that fits both caller-alias ceilings.
///
/// The byte accounting includes requested and target strings in the ordered
/// identity plus the normalized requested key and target value in the lookup
/// map. The first entry that would exceed either ceiling and every later entry
/// are discarded.
fn bounded_caller_aliases(aliases: &[(String, String)]) -> Vec<(String, String)> {
    let mut bounded = Vec::with_capacity(aliases.len().min(CALLER_ALIAS_MAX_ENTRIES));
    let mut retained_bytes = 0usize;
    for (requested, target) in aliases {
        if bounded.len() == CALLER_ALIAS_MAX_ENTRIES {
            break;
        }
        let normalized_requested = requested.to_lowercase();
        let entry_bytes = requested
            .len()
            .saturating_add(target.len())
            .saturating_add(normalized_requested.len())
            .saturating_add(target.len());
        let next_retained_bytes = retained_bytes.saturating_add(entry_bytes);
        if next_retained_bytes > CALLER_ALIAS_MAX_RETAINED_BYTES {
            break;
        }
        bounded.push((requested.to_owned(), target.to_owned()));
        retained_bytes = next_retained_bytes;
    }
    bounded
}

impl Default for FontManager {
    fn default() -> Self {
        Self::new()
    }
}

impl FontManager {
    fn from_base_database(db: fontdb::Database) -> Self {
        Self {
            base_db: db.clone(),
            base_caller_aliases: HashMap::new(),
            base_caller_families: HashMap::new(),
            db,
            cache: HashMap::new(),
            memory_face_data: HashMap::new(),
            fonts: Vec::new(),
            next_id: 0,
            coverage_fallbacks: HashMap::new(),
            coverage_misses: HashSet::new(),
            additional_fonts: Vec::new(),
            embedded_fonts: Vec::new(),
            embedded_aliases: HashMap::new(),
            caller_aliases: HashMap::new(),
            caller_families: HashMap::new(),
            explicit_aliases: Vec::new(),
            explicit_alias_map: HashMap::new(),
            shaping_memo: Mutex::new(ShapingMemo::new()),
            paragraph_font_trace: None,
            layout_fonts: Vec::new(),
            layout_font_diagnostics: Vec::new(),
        }
    }

    /// Create a new FontManager and load system fonts.
    ///
    /// Bundled fonts (Carlito, Caladea, Liberation) are loaded as fallbacks.
    /// System fonts are discovered when the `system-fonts` feature is enabled.
    pub fn new() -> Self {
        #[cfg(feature = "system-fonts")]
        {
            let db = NORMAL_FONT_DATABASE.get_or_init(|| {
                let mut db = bundled_font_database();
                db.load_system_fonts();
                #[cfg(test)]
                SYSTEM_FONT_DISCOVERY_RUNS.fetch_add(1, Ordering::Relaxed);
                db
            });
            Self::from_base_database(db.clone())
        }

        #[cfg(not(feature = "system-fonts"))]
        Self::from_base_database(bundled_font_database())
    }

    /// Create a font manager that loads bundled fonts without discovering
    /// system fonts.
    ///
    /// This mode makes font resolution reproducible across machines.
    pub fn new_deterministic() -> Result<Self> {
        Ok(Self::from_base_database(bundled_font_database()))
    }

    /// Load or replace additional font files (user-provided or extracted from
    /// an OOXML package).
    ///
    /// An unchanged set is a no-op so a reusable engine retains resolution and
    /// shaping state. A changed set rebuilds from the isolated base database,
    /// which prevents stale face ids and bounds repeated document edits.
    pub fn load_additional_fonts(&mut self, font_files: &[FontFile]) -> bool {
        if self.additional_fonts == font_files {
            return false;
        }
        let embedded = self.embedded_fonts.clone();
        // Previously admitted immutable embedded sources remain valid. Refuse
        // atomically if rebuilding ever fails instead of clearing active state.
        self.rebuild_font_sources(&embedded, font_files).is_ok()
    }

    /// Load bounded physical embedded font sources and all their owned aliases.
    /// Caller faces retain priority. An unchanged set preserves loaded IDs and caches.
    #[doc(hidden)]
    pub fn load_embedded_fonts(&mut self, sources: &[(Vec<String>, Arc<[u8]>)]) -> Result<bool> {
        if self.embedded_fonts == sources {
            return Ok(false);
        }
        let callers = self.additional_fonts.clone();
        self.rebuild_font_sources(sources, &callers)?;
        Ok(true)
    }

    fn rebuild_font_sources(
        &mut self,
        embedded_fonts: &[(Vec<String>, Arc<[u8]>)],
        font_files: &[FontFile],
    ) -> Result<()> {
        const FACE_LIMIT: usize = 64 * 1024 * 1024;
        const TOTAL_LIMIT: usize = 256 * 1024 * 1024;
        const ALIAS_LIMIT: usize = 16_384;
        const ALIAS_BYTES_LIMIT: usize = 4 * 1024 * 1024;
        let refused = || {
            LayoutError::FontParse(
                "embedded font sources exceed the data or alias limits".to_owned(),
            )
        };
        if embedded_fonts.len() > 4096 {
            return Err(refused());
        }
        let mut total = 0usize;
        let mut aliases = 0usize;
        let mut alias_bytes = 0usize;
        for (names, data) in embedded_fonts {
            total = total.checked_add(data.len()).ok_or_else(refused)?;
            aliases = aliases.checked_add(names.len()).ok_or_else(refused)?;
            if data.len() > FACE_LIMIT
                || total > TOTAL_LIMIT
                || aliases > ALIAS_LIMIT
                || names.is_empty()
            {
                return Err(refused());
            }
            for name in names {
                alias_bytes = alias_bytes
                    .checked_add(name.len().saturating_mul(4))
                    .ok_or_else(refused)?;
                if name.is_empty() || name.len() > 4096 || alias_bytes > ALIAS_BYTES_LIMIT {
                    return Err(refused());
                }
            }
        }
        let mut face_count = 0usize;
        let mut alias_references = 0usize;
        let mut font_name_bytes = 0usize;
        for (names, data) in embedded_fonts {
            let count = ttf_parser::fonts_in_collection(data).unwrap_or(1);
            if count == 0 || count > 256 {
                return Err(refused());
            }
            face_count = face_count.checked_add(count as usize).ok_or_else(refused)?;
            alias_references = alias_references
                .checked_add(
                    names
                        .len()
                        .checked_mul(count as usize)
                        .ok_or_else(refused)?,
                )
                .ok_or_else(refused)?;
            if face_count > 4096 || alias_references > 262_144 {
                return Err(refused());
            }
            for index in 0..count {
                let face = ttf_parser::Face::parse(data, index).map_err(|error| {
                    LayoutError::FontParse(format!("embedded font face: {error:?}"))
                })?;
                if face.names().len() > 1024 {
                    return Err(refused());
                }
                for name in face.names() {
                    font_name_bytes = font_name_bytes
                        .checked_add(name.name.len())
                        .ok_or_else(refused)?;
                    if font_name_bytes > ALIAS_BYTES_LIMIT {
                        return Err(refused());
                    }
                }
            }
        }
        let mut db = self.base_db.clone();
        let mut embedded_aliases: HashMap<String, Vec<fontdb::ID>> = HashMap::new();
        for (names, data) in embedded_fonts {
            let before = db.len();
            let mut body = Vec::new();
            body.try_reserve_exact(data.len()).map_err(|error| {
                LayoutError::FontParse(format!("embedded font allocation: {error}"))
            })?;
            body.extend_from_slice(data);
            db.load_font_data(body);
            let ids = db
                .faces()
                .skip(before)
                .map(|face| face.id)
                .collect::<Vec<_>>();
            if ids.is_empty() {
                return Err(LayoutError::FontParse(
                    "embedded font contains no supported faces".to_owned(),
                ));
            }
            for name in names {
                embedded_aliases
                    .entry(name.to_lowercase())
                    .or_default()
                    .extend(ids.iter().copied());
            }
        }
        let mut caller_aliases = self.base_caller_aliases.clone();
        let mut caller_families = self.base_caller_families.clone();
        for font_file in font_files {
            Self::load_caller_font(
                &mut db,
                &mut caller_aliases,
                &mut caller_families,
                &font_file.family,
                font_file.data.clone(),
            );
        }
        // CSS ranking still compares every face. On an equal rank, retain
        // constructor provenance ahead of the additional tier. Reverse the
        // constructor block so its established first-loaded tie stays first
        // when best_caller_face visits the additional tier last-loaded first.
        for (active, base) in [
            (&mut caller_aliases, &self.base_caller_aliases),
            (&mut caller_families, &self.base_caller_families),
        ] {
            for (family, base_ids) in base {
                if let Some(ids) = active.get_mut(family) {
                    ids[..base_ids.len()].reverse();
                    ids.rotate_left(base_ids.len());
                }
            }
        }

        self.db = db;
        self.caller_aliases = caller_aliases;
        self.caller_families = caller_families;
        self.embedded_aliases = embedded_aliases;
        self.cache.clear();
        self.memory_face_data.clear();
        self.fonts.clear();
        self.next_id = 0;
        self.coverage_fallbacks.clear();
        self.coverage_misses.clear();
        self.additional_fonts = font_files.to_vec();
        self.embedded_fonts = embedded_fonts.to_vec();
        if self.shaping_memo.is_poisoned() {
            self.shaping_memo.clear_poison();
        }
        self.shaping_memo
            .get_mut()
            .expect("shaping cache poison was cleared")
            .clear();
        Ok(())
    }

    /// Create a FontManager with user-provided fonts (no system font loading).
    ///
    /// Each entry is `(family_name, font_bytes)`. This is useful in environments
    /// where system fonts are not available, such as WASM.
    pub fn new_with_fonts(fonts: Vec<(String, Vec<u8>)>) -> Self {
        let mut db = fontdb::Database::new();
        let mut caller_aliases = HashMap::new();
        let mut caller_families = HashMap::new();
        for (family, data) in &fonts {
            Self::load_caller_font(
                &mut db,
                &mut caller_aliases,
                &mut caller_families,
                family,
                data.clone(),
            );
        }
        let mut manager = Self::from_base_database(db);
        manager.base_caller_aliases = caller_aliases.clone();
        manager.base_caller_families = caller_families.clone();
        manager.caller_aliases = caller_aliases;
        manager.caller_families = caller_families;
        // Constructor-only ties retain their original first-loaded selection.
        for active in [&mut manager.caller_aliases, &mut manager.caller_families] {
            for ids in active.values_mut() {
                ids.reverse();
            }
        }
        manager
    }

    /// Replace byte-free caller aliases from requested family to loaded family.
    ///
    /// An unchanged slice is a no-op. A changed slice invalidates only name
    /// resolution and coverage state. Loaded faces and shaping entries remain
    /// valid because their `FontId` values do not change.
    pub fn set_caller_aliases(&mut self, aliases: &[(String, String)]) -> bool {
        let aliases = bounded_caller_aliases(aliases);
        if self.explicit_aliases == aliases {
            return false;
        }

        self.explicit_alias_map = aliases
            .iter()
            .map(|(requested, target)| (requested.to_lowercase(), target.clone()))
            .collect();
        self.explicit_aliases = aliases;
        self.cache.clear();
        self.coverage_fallbacks.clear();
        self.coverage_misses.clear();
        true
    }

    fn load_caller_font(
        db: &mut fontdb::Database,
        aliases: &mut HashMap<String, Vec<fontdb::ID>>,
        families: &mut HashMap<String, Vec<fontdb::ID>>,
        family: &str,
        data: Vec<u8>,
    ) {
        let before = db.len();
        db.load_font_data(data);
        let loaded_faces = db
            .faces()
            .skip(before)
            .map(|face| {
                (
                    face.id,
                    face.families
                        .iter()
                        .map(|(name, _)| name.clone())
                        .collect::<Vec<_>>(),
                )
            })
            .collect::<Vec<_>>();
        for (id, loaded_families) in &loaded_faces {
            for loaded_family in loaded_families {
                families.entry(loaded_family.clone()).or_default().push(*id);
            }
        }
        if !loaded_faces.is_empty()
            && loaded_faces
                .iter()
                .all(|(_, families)| !families.iter().any(|loaded| loaded == family))
        {
            aliases
                .entry(family.to_lowercase())
                .or_default()
                .extend(loaded_faces.into_iter().map(|(id, _)| id));
        }
    }

    /// Begin one complete layout attempt's exact font-usage trace.
    #[doc(hidden)]
    pub fn begin_layout(&mut self) {
        self.paragraph_font_trace = None;
        self.layout_fonts = Vec::new();
        self.layout_font_diagnostics = Vec::new();
    }

    /// Begin recording exact events for one bounded cache candidate.
    #[doc(hidden)]
    pub fn begin_paragraph_font_trace(&mut self) {
        let Ok(generation) = NEXT_FONT_TRACE_GENERATION.fetch_update(
            Ordering::Relaxed,
            Ordering::Relaxed,
            |generation| generation.checked_add(1),
        ) else {
            self.paragraph_font_trace = None;
            return;
        };
        let generation = generation as u64;
        self.paragraph_font_trace = Some(ParagraphFontTrace {
            events: Vec::new(),
            generation,
            scope_depth: 0,
            owned_family_bytes: 0,
            overflowed: false,
        });
    }

    /// Finish the legacy resolution-only projection without adding contribution IDs.
    #[doc(hidden)]
    pub fn finish_paragraph_font_trace(&mut self) -> Option<Vec<FontId>> {
        if self.paragraph_font_trace.as_ref()?.scope_depth != 0 {
            return None;
        }
        let trace = self.paragraph_font_trace.take()?;
        if trace.overflowed {
            return None;
        }
        let mut ids = trace
            .events
            .into_iter()
            .filter_map(|event| match event {
                FontTraceEvent::Resolution(id) => Some(id),
                FontTraceEvent::Contribution(_) => None,
            })
            .collect::<Vec<_>>();
        ids.shrink_to_fit();
        Some(ids)
    }

    /// Finish only final contributions. Resolution attempts never establish identity.
    #[doc(hidden)]
    pub fn finish_paragraph_font_contributions(&mut self) -> Option<Vec<FontContribution>> {
        if self.paragraph_font_trace.as_ref()?.scope_depth != 0 {
            return None;
        }
        let trace = self.paragraph_font_trace.take()?;
        if trace.overflowed {
            return None;
        }
        let mut contributions = trace
            .events
            .into_iter()
            .filter_map(|event| match event {
                FontTraceEvent::Contribution(contribution) => Some(contribution),
                FontTraceEvent::Resolution(_) => None,
            })
            .collect::<Vec<_>>();
        contributions.shrink_to_fit();
        Some(contributions)
    }

    /// Finish one ordered bounded cache carrier, preserving resolution inventory.
    #[doc(hidden)]
    pub fn finish_paragraph_font_events(&mut self) -> Option<Vec<FontTraceEvent>> {
        if self.paragraph_font_trace.as_ref()?.scope_depth != 0 {
            return None;
        }
        let mut trace = self.paragraph_font_trace.take()?;
        if trace.overflowed {
            return None;
        }
        trace.events.shrink_to_fit();
        Some(trace.events)
    }

    /// Capture a local shape without replacing a caller's active trace.
    #[doc(hidden)]
    pub fn begin_font_contribution_scope(&mut self) -> Option<FontContributionScope> {
        let owns_trace = self.paragraph_font_trace.is_none();
        if owns_trace {
            self.begin_paragraph_font_trace();
        }
        let trace = self.paragraph_font_trace.as_mut()?;
        if trace.overflowed {
            return None;
        }
        let depth = trace.scope_depth.checked_add(1)?;
        trace.scope_depth = depth;
        Some(FontContributionScope {
            generation: trace.generation,
            event_start: trace.events.len(),
            depth,
            owns_trace,
            diagnostic_start: self.layout_font_diagnostics.len(),
        })
    }

    /// Return bounded staged contributions, preserving caller resolution events.
    #[doc(hidden)]
    pub fn finish_font_contribution_scope(
        &mut self,
        scope: FontContributionScope,
    ) -> Result<Option<Vec<FontTraceEvent>>> {
        let trace = self.paragraph_font_trace.as_ref().ok_or_else(|| {
            LayoutError::Layout("font contribution scope no longer active".into())
        })?;
        if trace.generation != scope.generation
            || trace.scope_depth != scope.depth
            || (!trace.overflowed && scope.event_start > trace.events.len())
        {
            return Err(LayoutError::Layout(
                "stale or nested font contribution scope".into(),
            ));
        }
        if trace.overflowed {
            self.layout_font_diagnostics
                .truncate(scope.diagnostic_start);
            if scope.owns_trace {
                self.paragraph_font_trace = None;
            } else {
                self.paragraph_font_trace.as_mut().unwrap().scope_depth -= 1;
            }
            return Ok(None);
        }
        let staged = &trace.events[scope.event_start..];
        let bytes = staged.iter().fold(
            staged
                .len()
                .checked_mul(std::mem::size_of::<FontTraceEvent>()),
            |total, event| {
                total.and_then(|total| {
                    total.checked_add(match event {
                        FontTraceEvent::Resolution(_) => 0,
                        FontTraceEvent::Contribution(contribution) => contribution
                            .requested_family
                            .as_ref()
                            .map_or(0, String::capacity),
                    })
                })
            },
        );
        if bytes.is_none_or(|bytes| bytes > PARAGRAPH_FONT_TRACE_MAX_BYTES) {
            return Err(LayoutError::Layout(
                "font trace capture exceeds byte budget".into(),
            ));
        }
        let mut result = Vec::with_capacity(staged.len());
        if result
            .capacity()
            .saturating_mul(std::mem::size_of::<FontTraceEvent>())
            .saturating_add(bytes.unwrap().saturating_sub(std::mem::size_of_val(staged)))
            > PARAGRAPH_FONT_TRACE_MAX_BYTES
        {
            return Err(LayoutError::Layout(
                "font trace capture allocation exceeds budget".into(),
            ));
        }
        result.extend(staged.iter().cloned());
        self.layout_font_diagnostics
            .truncate(scope.diagnostic_start);
        if scope.owns_trace {
            self.paragraph_font_trace = None;
        } else {
            let trace = self.paragraph_font_trace.as_mut().unwrap();
            let mut index = 0usize;
            trace.events.retain(|event| {
                let keep =
                    index < scope.event_start || matches!(event, FontTraceEvent::Resolution(_));
                index += 1;
                keep
            });
            trace.owned_family_bytes = trace
                .events
                .iter()
                .filter_map(|event| match event {
                    FontTraceEvent::Contribution(contribution) => {
                        contribution.requested_family.as_ref()
                    }
                    FontTraceEvent::Resolution(_) => None,
                })
                .map(String::capacity)
                .sum();
            trace.scope_depth -= 1;
        }
        Ok(Some(result))
    }

    /// Copy private native shaping context within one bounded admission budget.
    #[doc(hidden)]
    pub fn copy_font_request_context(
        &mut self,
        family: Option<&str>,
        owned_bytes: &mut usize,
        context_entries: &mut usize,
    ) -> Option<String> {
        let family = family?;
        let total = owned_bytes.checked_add(family.len());
        if total.is_none_or(|bytes| bytes > PARAGRAPH_FONT_TRACE_MAX_BYTES)
            || *context_entries >= PARAGRAPH_FONT_TRACE_MAX_ENTRIES
        {
            if let Some(trace) = self.paragraph_font_trace.as_mut() {
                trace.events = Vec::new();
                trace.owned_family_bytes = 0;
                trace.overflowed = true;
            }
            *owned_bytes = PARAGRAPH_FONT_TRACE_MAX_BYTES;
            return None;
        }
        let copied = family.to_owned();
        *context_entries += 1;
        *owned_bytes = owned_bytes.saturating_add(copied.capacity());
        Some(copied)
    }

    /// Admit a successfully shaped final segment, retaining its original request.
    #[doc(hidden)]
    pub fn record_font_contribution(
        &mut self,
        font_id: FontId,
        requested_family: Option<&str>,
        text: &str,
    ) -> Result<()> {
        if text.is_empty() {
            return Ok(());
        }
        self.get_font(font_id)?;
        let bytes = requested_family.map_or(0, str::len);
        if let Some(trace) = self.paragraph_font_trace.as_mut() {
            let total = trace
                .owned_family_bytes
                .checked_add(bytes)
                .and_then(|bytes| {
                    bytes.checked_add(
                        trace
                            .events
                            .capacity()
                            .max(trace.events.len().saturating_add(1))
                            .saturating_mul(std::mem::size_of::<FontTraceEvent>()),
                    )
                });
            if total.is_none_or(|bytes| bytes > PARAGRAPH_FONT_TRACE_MAX_BYTES)
                || trace.events.len() >= PARAGRAPH_FONT_TRACE_MAX_ENTRIES
            {
                trace.events = Vec::new();
                trace.owned_family_bytes = 0;
                trace.overflowed = true;
            } else if !trace.overflowed {
                let requested_family = requested_family.map(str::to_owned);
                trace.owned_family_bytes += requested_family.as_ref().map_or(0, String::capacity);
                trace
                    .events
                    .push(FontTraceEvent::Contribution(FontContribution {
                        font_id,
                        requested_family,
                    }));
                if trace
                    .events
                    .capacity()
                    .saturating_mul(std::mem::size_of::<FontTraceEvent>())
                    .saturating_add(trace.owned_family_bytes)
                    > PARAGRAPH_FONT_TRACE_MAX_BYTES
                {
                    trace.events = Vec::new();
                    trace.owned_family_bytes = 0;
                    trace.overflowed = true;
                }
            }
        }
        if !self.layout_fonts.contains(&font_id) {
            self.layout_fonts.push(font_id);
        }
        self.record_font_approximation(font_id, requested_family);
        Ok(())
    }

    /// Validate the entire cached contribution before any trace or inventory mutation.
    #[doc(hidden)]
    pub fn validate_font_trace_events(&self, contributions: &[FontTraceEvent]) -> Result<()> {
        let owned_bytes = contributions.iter().try_fold(0usize, |total, event| {
            let (id, bytes) = match event {
                FontTraceEvent::Resolution(id) => (*id, 0),
                FontTraceEvent::Contribution(contribution) => (
                    contribution.font_id,
                    contribution
                        .requested_family
                        .as_ref()
                        .map_or(0, String::capacity),
                ),
            };
            self.get_font(id)?;
            total
                .checked_add(bytes)
                .ok_or_else(|| LayoutError::Layout("font trace byte count overflow".into()))
        })?;
        let (count, existing_bytes, capacity, overflowed) = self
            .paragraph_font_trace
            .as_ref()
            .map_or((0, 0, 0, false), |trace| {
                (
                    trace.events.len(),
                    trace.owned_family_bytes,
                    trace.events.capacity(),
                    trace.overflowed,
                )
            });
        let total_count = count.checked_add(contributions.len());
        let total_bytes = total_count
            .and_then(|count| {
                capacity
                    .max(count)
                    .checked_mul(std::mem::size_of::<FontTraceEvent>())
            })
            .and_then(|bytes| bytes.checked_add(existing_bytes))
            .and_then(|bytes| bytes.checked_add(owned_bytes));
        if overflowed
            || total_count.is_none_or(|count| count > PARAGRAPH_FONT_TRACE_MAX_ENTRIES)
            || total_bytes.is_none_or(|bytes| bytes > PARAGRAPH_FONT_TRACE_MAX_BYTES)
        {
            return Err(LayoutError::Layout(
                "font contribution replay exceeds trace budget".into(),
            ));
        }
        Ok(())
    }

    /// Atomically replay an admitted final contribution vector.
    #[doc(hidden)]
    pub fn replay_font_trace_events(&mut self, contributions: &[FontTraceEvent]) -> Result<()> {
        self.validate_font_trace_events(contributions)?;
        if let Some(trace) = self.paragraph_font_trace.as_mut() {
            let mut prepared = Vec::with_capacity(trace.events.len() + contributions.len());
            let owned_bytes = trace.owned_family_bytes
                + contributions
                    .iter()
                    .map(|entry| match entry {
                        FontTraceEvent::Resolution(_) => 0,
                        FontTraceEvent::Contribution(entry) => {
                            entry.requested_family.as_ref().map_or(0, String::capacity)
                        }
                    })
                    .sum::<usize>();
            if prepared
                .capacity()
                .saturating_mul(std::mem::size_of::<FontTraceEvent>())
                .saturating_add(owned_bytes)
                > PARAGRAPH_FONT_TRACE_MAX_BYTES
            {
                return Err(LayoutError::Layout(
                    "font contribution replay allocation exceeds budget".into(),
                ));
            }
            prepared.extend(trace.events.iter().cloned());
            prepared.extend(contributions.iter().cloned());
            trace.owned_family_bytes = prepared
                .iter()
                .filter_map(|event| match event {
                    FontTraceEvent::Contribution(contribution) => {
                        contribution.requested_family.as_ref()
                    }
                    FontTraceEvent::Resolution(_) => None,
                })
                .map(String::capacity)
                .sum();
            trace.events = prepared;
        }
        for event in contributions {
            let id = match event {
                FontTraceEvent::Resolution(id) => *id,
                FontTraceEvent::Contribution(contribution) => contribution.font_id,
            };
            if !self.layout_fonts.contains(&id) {
                self.layout_fonts.push(id);
            }
            if let FontTraceEvent::Contribution(contribution) = event {
                self.record_font_approximation(id, contribution.requested_family.as_deref());
            }
        }
        Ok(())
    }

    /// Replay the exact legacy resolution events attached to a cached paragraph.
    #[doc(hidden)]
    pub fn replay_layout_font_trace(&mut self, trace: &[FontId]) {
        for &font_id in trace {
            self.record_layout_font(font_id);
        }
    }

    /// Distinct current-layout fonts in first-resolution order.
    #[doc(hidden)]
    pub fn current_layout_fonts(&self) -> &[FontId] {
        &self.layout_fonts
    }

    /// Whether no historical loaded face is absent from this layout.
    #[doc(hidden)]
    pub fn every_loaded_font_is_current(&self) -> bool {
        self.fonts
            .iter()
            .map(|font| font.id)
            .eq(self.layout_fonts.iter().copied())
            && self
                .layout_fonts
                .iter()
                .enumerate()
                .all(|(index, font_id)| *font_id == FontId(index as u32))
    }

    /// Drop faces that were loaded by an older successful layout but are no
    /// longer active. Every face used by the current document is retained,
    /// even when that working set contains more than the cache ceilings.
    #[doc(hidden)]
    pub fn retain_current_fonts(&mut self) {
        let current = self.layout_fonts.iter().copied().collect::<HashSet<_>>();
        let old_index_ids = self
            .fonts
            .iter()
            .enumerate()
            .map(|(index, font)| (index, font.id))
            .collect::<HashMap<_, _>>();
        self.fonts.retain(|font| current.contains(&font.id));
        let current_order = self
            .layout_fonts
            .iter()
            .enumerate()
            .map(|(index, font_id)| (*font_id, index))
            .collect::<HashMap<_, _>>();
        self.fonts
            .sort_by_key(|font| current_order.get(&font.id).copied().unwrap_or(usize::MAX));

        let indices = self
            .fonts
            .iter()
            .enumerate()
            .map(|(index, font)| (font.id, index))
            .collect::<HashMap<_, _>>();
        let old_cache = std::mem::take(&mut self.cache);
        self.cache = old_cache
            .into_iter()
            .filter_map(|(key, old_index)| {
                let font_id = old_index_ids.get(&old_index)?;
                Some((key, *indices.get(font_id)?))
            })
            .collect();
        self.coverage_fallbacks.clear();
        self.coverage_misses.clear();
        let active_db_ids = self
            .fonts
            .iter()
            .map(|font| font.db_id)
            .collect::<HashSet<_>>();
        self.memory_face_data
            .retain(|db_id, _| active_db_ids.contains(db_id));

        let memo = self
            .shaping_memo
            .get_mut()
            .unwrap_or_else(std::sync::PoisonError::into_inner);
        memo.entries
            .retain(|(key, _, _, _)| current.contains(&key.font_id));
        memo.bytes = memo.entries.iter().map(|(_, _, bytes, _)| bytes).sum();
    }

    /// Resolve a font for `text`, falling back on glyph coverage.
    ///
    /// `resolve_font` picks by family name alone. That is enough for Latin
    /// text, but a run asking for a Chinese family on a machine without it
    /// falls down the name chain and lands on a Latin font, which has no CJK
    /// glyphs, so every character renders as a missing-glyph box. Name
    /// matching cannot detect that, because the font it chose exists and is
    /// perfectly valid, it simply cannot draw this text.
    ///
    /// So the resolved font is checked against the text, and when a character
    /// is missing another font that can draw it is looked for.
    ///
    /// This is per run rather than per character: the font that covers the
    /// first missing character is used for the whole run. Text that mixes
    /// scripts inside one run is therefore still imperfect, but it is a large
    /// improvement on drawing boxes.
    pub fn resolve_font_for_text(
        &mut self,
        family: Option<&str>,
        bold: bool,
        italic: bool,
        text: &str,
    ) -> Result<FontId> {
        let primary = self.resolve_font(family, bold, italic)?;

        let Some(idx) = self.index_of(primary) else {
            return Ok(primary);
        };
        let missing = self.uncovered(idx, text);
        if missing.is_empty() {
            return Ok(primary);
        }

        match self.font_covering(&missing, bold, italic) {
            // Nothing installed can draw it. Keep the original font so the
            // text still occupies the right space.
            None => Ok(primary),
            Some(id) => Ok(id),
        }
    }

    /// Current layout's owned approximation facts, never loaded-font history.
    #[doc(hidden)]
    pub fn current_font_diagnostics(&self) -> &[Diagnostic] {
        &self.layout_font_diagnostics
    }

    // Selection aliases name requests, not the physical family of a face.
    // Authenticate the selected collection face from its original source bytes.
    fn physical_requested_family_matches(&self, index: usize, requested: &str) -> bool {
        let Some(font) = self.fonts.get(index) else {
            return false;
        };
        let Ok(face) = ttf_parser::Face::parse(&font.data, font.face_index) else {
            return false;
        };
        face.names().into_iter().any(|name| {
            matches!(
                name.name_id,
                ttf_parser::name_id::FAMILY
                    | ttf_parser::name_id::TYPOGRAPHIC_FAMILY
                    | ttf_parser::name_id::WWS_FAMILY
            ) && name
                .to_string()
                .is_some_and(|family| family.eq_ignore_ascii_case(requested))
        })
    }

    fn record_font_approximation(&mut self, id: FontId, requested: Option<&str>) {
        let Some(requested) = requested else {
            return;
        };
        let requested = if requested.eq_ignore_ascii_case("Aptos") {
            "Aptos"
        } else if requested.eq_ignore_ascii_case("Aptos Display") {
            "Aptos Display"
        } else {
            return;
        };
        let Some(index) = self.index_of(id) else {
            return;
        };
        if self.physical_requested_family_matches(index, requested) {
            return;
        }
        let prefix = format!("Font '{requested}' is approximate: ");
        if self
            .layout_font_diagnostics
            .iter()
            .any(|diagnostic| diagnostic.message.starts_with(&prefix))
        {
            return;
        }
        self.layout_font_diagnostics.push(Diagnostic {
            message: format!(
                "{prefix}selected '{}' for contributing text",
                self.fonts[index].family
            ),
        });
    }

    /// Select an explicit symbol or bullet without changing its source owner.
    /// Genuine original-code coverage wins over the bounded semantic substitute.
    pub fn resolve_symbol_text(
        &mut self,
        family: Option<&str>,
        bold: bool,
        italic: bool,
        original: &str,
        ppt_bullet_compatibility: bool,
    ) -> Result<(String, FontId)> {
        let primary = self.resolve_font(family, bold, italic)?;
        if self.has_original_symbol_coverage(primary, family, original) {
            return Ok((original.to_owned(), primary));
        }
        let displayed: String = original
            .chars()
            .map(|character| {
                let code = character as u32;
                let code = if (0xf000..=0xf0ff).contains(&code) {
                    code - 0xf000
                } else {
                    code
                };
                let mapped = if family.is_some_and(|name| name.eq_ignore_ascii_case("Wingdings")) {
                    match code {
                        0xa7 => Some('\u{25aa}'),
                        0xd8 => Some('\u{27a2}'),
                        0xfc => Some('\u{2713}'),
                        0x76 => Some('\u{2756}'),
                        0x71 => Some('\u{2751}'),
                        0x6e => Some('\u{25a0}'),
                        0xfe => Some('\u{2611}'),
                        0xb7 if ppt_bullet_compatibility && character == '\u{f0b7}' => {
                            Some('\u{2022}')
                        }
                        _ => None,
                    }
                } else if family.is_some_and(|name| name.eq_ignore_ascii_case("Symbol"))
                    && code == 0xb7
                {
                    Some('\u{2022}')
                } else {
                    None
                };
                mapped.unwrap_or(character)
            })
            .collect();
        let selected = self.resolve_font_for_text(family, bold, italic, &displayed)?;
        Ok((displayed, selected))
    }

    /// Whether the selected physical requested face covers the original symbol codes.
    /// Caller labels alone cannot establish this coverage or symbol meaning.
    pub fn has_original_symbol_coverage(
        &self,
        selected: FontId,
        family: Option<&str>,
        original: &str,
    ) -> bool {
        self.index_of(selected).is_some_and(|index| {
            family.is_some_and(|name| self.physical_requested_family_matches(index, name))
                && self.uncovered(index, original).is_empty()
        })
    }

    /// Owned unknown semantic mappings for accepted original symbol text.
    /// This query neither records font history nor claims glyphs are missing.
    pub fn unknown_symbol_diagnostics(
        &self,
        selected: FontId,
        requested: Option<&str>,
        original: &str,
        displayed: &str,
    ) -> Vec<String> {
        if self.index_of(selected).is_none() {
            return Vec::new();
        }
        let family = match requested {
            Some(name) if name.eq_ignore_ascii_case("Wingdings") => "Wingdings",
            Some(name) if name.eq_ignore_ascii_case("Symbol") => "Symbol",
            _ => return Vec::new(),
        };
        let mut result = Vec::new();
        for (source, displayed) in original.chars().zip(displayed.chars()) {
            let code = source as u32;
            if source != displayed
                || source.is_whitespace()
                || source.is_control()
                || !(code <= 0xff || (0xf000..=0xf0ff).contains(&code))
            {
                continue;
            }
            let mut utf8 = [0; 4];
            if self.has_original_symbol_coverage(
                selected,
                Some(family),
                source.encode_utf8(&mut utf8),
            ) {
                continue;
            }
            let message = format!("Unknown {family} symbol U+{code:04X}: original code retained");
            if !result.contains(&message) {
                result.push(message);
            }
        }
        // One canonical family and at most 512 original low-byte/PUA scalars.
        result
    }

    /// The characters in `text` that the font at `idx` cannot draw.
    ///
    /// Whitespace and control characters are skipped: a font without a glyph
    /// for a space is not a reason to go looking for another one.
    fn uncovered(&self, idx: usize, text: &str) -> Vec<char> {
        let font = &self.fonts[idx];
        let Ok(face) = ttf_parser::Face::parse(&font.data, font.face_index) else {
            return Vec::new();
        };
        let mut seen = HashSet::new();
        text.chars()
            .filter(|&ch| !ch.is_whitespace() && !ch.is_control())
            .filter(|&ch| face.glyph_index(ch).is_none())
            .filter(|&ch| seen.insert(ch))
            .collect()
    }

    /// Whether the font at `idx` has a glyph for `ch`.
    fn covers(&self, idx: usize, ch: char) -> bool {
        let font = &self.fonts[idx];
        ttf_parser::Face::parse(&font.data, font.face_index)
            .map(|face| face.glyph_index(ch).is_some())
            .unwrap_or(false)
    }

    /// Find a font that can draw `missing`.
    ///
    /// A font covering every missing character wins. Failing that the one
    /// covering the most is used, because a single run gets a single font and
    /// partial coverage still beats a row of boxes. Picking on the first
    /// missing character alone is not enough: a Japanese face may have the
    /// characters shared with Chinese and not the simplified-only ones, so it
    /// would look like a fix and still leave gaps.
    fn font_covering(&mut self, missing: &[char], bold: bool, italic: bool) -> Option<FontId> {
        if missing.iter().all(|ch| self.coverage_misses.contains(ch)) {
            return None;
        }

        let mut best: Option<(usize, usize)> = None; // (covered count, font index)
        let consider = |this: &Self, idx: usize, best: &mut Option<(usize, usize)>| -> bool {
            let covered = missing.iter().filter(|&&ch| this.covers(idx, ch)).count();
            if covered == 0 {
                return false;
            }
            if best.map(|(n, _)| covered > n).unwrap_or(true) {
                *best = Some((covered, idx));
            }
            covered == missing.len()
        };

        // Fonts that already rescued an earlier run, which for a document in
        // one script is almost always the answer again.
        if let Some(known) = self.coverage_fallbacks.get(&(bold, italic)).cloned() {
            for idx in known {
                if consider(self, idx, &mut best) {
                    let id = self.fonts[idx].id;
                    self.record_layout_font(id);
                    return Some(id);
                }
            }
        }

        // Families with broad coverage, then everything else the database
        // knows about. Both go through resolve_font so loading and caching
        // stay in one place.
        let candidates: Vec<String> = BROAD_COVERAGE_FAMILIES
            .iter()
            .map(|s| s.to_string())
            .chain(
                self.db
                    .faces()
                    .filter_map(|f| f.families.first().map(|(name, _)| name.clone())),
            )
            .collect();

        for name in candidates {
            let Ok(id) = self.resolve_font(Some(&name), bold, italic) else {
                continue;
            };
            let Some(idx) = self.index_of(id) else {
                continue;
            };
            let complete = consider(self, idx, &mut best);
            if complete {
                self.remember_coverage_fallback(bold, italic, idx);
                return Some(id);
            }
        }

        match best {
            Some((_, idx)) => {
                self.remember_coverage_fallback(bold, italic, idx);
                let id = self.fonts[idx].id;
                self.record_layout_font(id);
                Some(id)
            }
            None => {
                self.remember_coverage_misses(missing);
                None
            }
        }
    }

    /// Index into `fonts` for a FontId.
    fn index_of(&self, id: FontId) -> Option<usize> {
        self.fonts.iter().position(|f| f.id == id)
    }

    /// Resolve a font by family name, bold, and italic flags.
    /// Returns a FontId. Uses fallback chain if the requested font is not found.
    pub fn resolve_font(
        &mut self,
        family: Option<&str>,
        bold: bool,
        italic: bool,
    ) -> Result<FontId> {
        self.resolve_font_inner(family, bold, italic, true)
    }

    /// Resolve a font for metrics without claiming that it emitted glyphs.
    #[doc(hidden)]
    pub fn resolve_font_for_metrics(
        &mut self,
        family: Option<&str>,
        bold: bool,
        italic: bool,
    ) -> Result<FontId> {
        self.resolve_font_inner(family, bold, italic, false)
    }

    fn resolve_font_inner(
        &mut self,
        family: Option<&str>,
        bold: bool,
        italic: bool,
        record_layout_use: bool,
    ) -> Result<FontId> {
        let family_name = family.unwrap_or("Arial");

        let key = FontKey {
            family: family_name.to_string(),
            bold,
            italic,
        };

        if let Some(idx) = self.cache.get(&key).copied() {
            let id = self.fonts[idx].id;
            if record_layout_use {
                self.record_layout_font(id);
            }
            return Ok(id);
        }

        let requested_key = family_name.to_lowercase();
        let style = if italic {
            fontdb::Style::Italic
        } else {
            fontdb::Style::Normal
        };
        let weight = if bold {
            fontdb::Weight::BOLD
        } else {
            fontdb::Weight::NORMAL
        };

        let caller_candidate = || {
            let exact = self.caller_families.get(family_name);
            let aliases = self.caller_aliases.get(&requested_key);
            match (exact, aliases) {
                (Some(exact), Some(aliases)) => {
                    // CSS compares both buckets together. Equal ranks retain
                    // constructor-first ties, then the latest additional face.
                    let (mut constructor, mut additional): (Vec<_>, Vec<_>) = self
                        .db
                        .faces()
                        .filter(|face| exact.contains(&face.id) || aliases.contains(&face.id))
                        .map(|face| face.id)
                        .partition(|id| self.base_db.face(*id).is_some());
                    constructor.reverse();
                    additional.extend(constructor);
                    best_caller_face(&self.db, &additional, weight, style)
                }
                (Some(ids), None) | (None, Some(ids)) => {
                    best_caller_face(&self.db, ids, weight, style)
                }
                (None, None) => None,
            }
        };
        let query_family = |family: &str| {
            if let Some(ids) = self.caller_families.get(family)
                && let Some(id) = best_caller_face(&self.db, ids, weight, style)
            {
                return Some(id);
            }
            if let Some(ids) = self.embedded_aliases.get(&family.to_lowercase())
                && let Some(id) = best_caller_face(&self.db, ids, weight, style)
            {
                return Some(id);
            }
            let query = fontdb::Query {
                families: &[fontdb::Family::Name(family)],
                weight,
                style,
                stretch: fontdb::Stretch::Normal,
            };
            self.db.query(&query)
        };

        let mut found_id = if self.explicit_alias_map.contains_key(&requested_key) {
            query_family(family_name)
        } else {
            caller_candidate().or_else(|| query_family(family_name))
        };
        if found_id.is_none()
            && let Some(alias) = self.explicit_alias_map.get(&requested_key)
        {
            found_id = query_family(alias);
        }
        if found_id.is_none()
            && let Some(ids) = self.caller_aliases.get(&requested_key)
        {
            found_id = best_caller_face(&self.db, ids, weight, style);
        }

        // Map common Word font names to metric-compatible alternatives, then
        // try generic fallbacks.
        let mut fallbacks: Vec<&str> = map_font_name(family_name).to_vec();
        for generic in &[
            "Carlito",
            "Arial",
            "Liberation Sans",
            "Helvetica",
            "DejaVu Sans",
            "Noto Sans",
        ] {
            if !fallbacks.contains(generic) {
                fallbacks.push(generic);
            }
        }

        if found_id.is_none() {
            for fallback in &fallbacks {
                let Some(id) = query_family(fallback) else {
                    continue;
                };
                found_id = Some(id);
                break;
            }
        }

        // Last resort: try generic families
        if found_id.is_none() {
            for generic_family in &[
                fontdb::Family::SansSerif,
                fontdb::Family::Serif,
                fontdb::Family::Monospace,
            ] {
                let query = fontdb::Query {
                    families: &[*generic_family],
                    weight,
                    style,
                    stretch: fontdb::Stretch::Normal,
                };
                if let Some(id) = self.db.query(&query) {
                    found_id = Some(id);
                    break;
                }
            }
        }

        let db_id = found_id.ok_or_else(|| {
            LayoutError::FontNotFound(format!("No font found for family '{family_name}'"))
        })?;

        // Preserve the established one-loaded-font-per-request-key behavior
        // while the bounded alias cache has room. At the ceiling, reuse the
        // exact resolved face rather than growing without limit.
        if self.cache.len() >= RESOLUTION_CACHE_MAX_ENTRIES
            && let Some(idx) = self
                .fonts
                .iter()
                .position(|font| font.db_id == db_id && font.bold == bold && font.italic == italic)
        {
            let id = self.fonts[idx].id;
            if record_layout_use {
                self.record_layout_font(id);
            }
            return Ok(id);
        }

        let font_id = FontId(self.next_id);
        self.next_id += 1;

        // Load file-backed data through the process cache. All faces in a TTC
        // carry the same source path, so their collection indices share bytes.
        let (data, face_index) = font_data_for_face(&self.db, db_id, &mut self.memory_face_data)
            .ok_or_else(|| LayoutError::FontParse("Failed to load font data".into()))?;

        let (
            units_per_em,
            ascender,
            descender,
            line_gap,
            word_line,
            variations,
            synthetic_bold,
            synthetic_italic,
        ) = {
            let mut face = ttf_parser::Face::parse(&data, face_index)
                .map_err(|e| LayoutError::FontParse(format!("ttf-parser error: {e}")))?;
            let (variations, synthetic_bold, synthetic_italic) =
                face_style_adjustment(&face, bold, italic);
            for (tag, value) in &variations {
                face.set_variation(ttf_parser::Tag::from_bytes(tag), *value)
                    .ok_or_else(|| {
                        LayoutError::FontParse("selected font axis could not be applied".into())
                    })?;
            }
            (
                face.units_per_em(),
                face.ascender(),
                face.descender(),
                face.line_gap(),
                WordLineUnits::of(&face),
                variations,
                synthetic_bold,
                synthetic_italic,
            )
        };

        // Every metric and advance is scaled by size/upem, so a zero here would
        // turn the whole layout into infinities.
        if units_per_em == 0 {
            return Err(LayoutError::FontParse(format!(
                "font '{family_name}' declares zero units per em"
            )));
        }

        let (shaper_data, shaper_instance) = {
            let face = harfrust::FontRef::from_index(&data, face_index)
                .map_err(|e| LayoutError::FontParse(format!("failed to read font face: {e}")))?;
            let instance = (!variations.is_empty()).then(|| {
                harfrust::ShaperInstance::from_variations(
                    &face,
                    variations.iter().map(|(tag, value)| harfrust::Variation {
                        tag: harfrust::Tag::new(tag),
                        value: *value,
                    }),
                )
            });
            (harfrust::ShaperData::new(&face), instance)
        };

        let actual_family = self
            .db
            .face(db_id)
            .map(|f| {
                f.families
                    .first()
                    .map(|(name, _)| name.clone())
                    .unwrap_or_else(|| family_name.to_string())
            })
            .unwrap_or_else(|| family_name.to_string());

        let idx = self.fonts.len();
        self.fonts.push(LoadedFont {
            db_id,
            id: font_id,
            family: actual_family,
            bold,
            italic,
            data,
            face_index,
            units_per_em,
            ascender,
            descender,
            line_gap,
            word_line,
            shaper_data,
            variations,
            shaper_instance,
            synthetic_bold,
            synthetic_italic,
        });
        self.remember_font_key(key, idx);
        if record_layout_use {
            self.record_layout_font(font_id);
        }

        Ok(font_id)
    }

    /// Get font metrics at a given size in points.
    pub fn metrics(&self, font_id: FontId, size_pt: f64) -> Result<FontMetrics> {
        let font = self.get_font(font_id)?;
        let scale = size_pt / font.units_per_em as f64;

        Ok(FontMetrics {
            ascent: font.ascender as f64 * scale,
            descent: -(font.descender as f64) * scale, // make positive
            line_gap: font.line_gap as f64 * scale,
            units_per_em: font.units_per_em,
        })
    }

    /// The metrics Word measures one line of a font on, at a given size.
    ///
    /// Word reads a face the way Windows reports it (`TEXTMETRIC`): `ascent`
    /// and `descent` are the OS/2 `usWinAscent` and `usWinDescent`, and
    /// `line_gap` is the external leading, the part of the `hhea` line gap
    /// that the Windows extent does not already cover. A single line is the
    /// sum of the three, so Calibri and Carlito are 2500/2048 em and Arial and
    /// Liberation Sans 2355/2048 em. A face that sets the OS/2
    /// `USE_TYPO_METRICS` flag is measured on its typographic values instead,
    /// which is what Word does with Aptos. Actual OS/2 CP932/936/949/950
    /// coverage instead selects gap-free hhea, with 15 percent of that extent
    /// added independently above and below the baseline. Missing/invalid CJK
    /// tables retain the established fallback. A face without Windows metrics
    /// uses the values [`FontManager::metrics`] reports.
    pub fn word_line_metrics(&self, font_id: FontId, size_pt: f64) -> Result<FontMetrics> {
        let font = self.get_font(font_id)?;
        let scale = size_pt / font.units_per_em as f64;
        let units = font.word_line.unwrap_or(WordLineUnits {
            ascent: i32::from(font.ascender),
            descent: -i32::from(font.descender),
            leading: i32::from(font.line_gap).max(0),
            cjk: false,
        });

        let extra = if units.cjk {
            f64::from(units.ascent + units.descent) * 0.15
        } else {
            0.0
        };
        Ok(FontMetrics {
            ascent: (f64::from(units.ascent) + extra) * scale,
            descent: (f64::from(units.descent) + extra) * scale,
            line_gap: units.leading as f64 * scale,
            units_per_em: font.units_per_em,
        })
    }

    /// Shape a text string using HarfRust. Returns glyph IDs and advances.
    pub fn shape_text(&self, font_id: FontId, text: &str, size_pt: f64) -> Result<ShapedText> {
        self.shape_text_with_override(font_id, text, size_pt, None)
    }

    /// Shape with an explicit per-request kerning feature, without changing defaults.
    #[doc(hidden)]
    pub fn shape_text_with_kerning(
        &self,
        font_id: FontId,
        text: &str,
        size_pt: f64,
        enabled: bool,
    ) -> Result<ShapedText> {
        self.shape_text_with_override(font_id, text, size_pt, Some(enabled))
    }

    fn shape_text_with_override(
        &self,
        font_id: FontId,
        text: &str,
        size_pt: f64,
        kerning: Option<bool>,
    ) -> Result<ShapedText> {
        // HarfRust cannot derive segment properties from an empty buffer, and
        // there is nothing to shape anyway.
        if text.is_empty() {
            return Ok(ShapedText {
                glyph_ids: Vec::new(),
                advances: Vec::new(),
                width: 0.0,
            });
        }

        let key = ShapingKey {
            font_id,
            text: text.to_owned(),
            size_bits: size_pt.to_bits(),
            kerning,
        };
        let fingerprint = shaping_fingerprint(key.font_id, &key.text, key.size_bits, key.kerning);
        let mut memo = match self.shaping_memo.lock() {
            Ok(memo) => memo,
            Err(poisoned) => {
                let mut memo = poisoned.into_inner();
                memo.clear();
                self.shaping_memo.clear_poison();
                memo
            }
        };
        if let Some(shaped) = memo
            .entries
            .iter()
            .rev()
            .find(|(candidate, _, _, candidate_fingerprint)| {
                *candidate_fingerprint == fingerprint && candidate == &key
            })
            .map(|(_, shaped, _, _)| shaped.clone())
        {
            #[cfg(test)]
            {
                memo.hits += 1;
            }
            return Ok(shaped);
        }
        #[cfg(test)]
        {
            memo.misses += 1;
        }

        let positioned = self.shape_legacy(font_id, text, size_pt, kerning, false)?;
        let width = positioned.x_advances.iter().sum();
        let shaped = ShapedText {
            glyph_ids: positioned.glyph_ids,
            advances: positioned.x_advances,
            width,
        };
        memo.insert(key, shaped.clone());
        Ok(shaped)
    }

    /// Shape one logical text segment into script, font, and bidi-level spans.
    ///
    /// The returned spans are in logical order. The rich line breaker applies
    /// UAX 9 visual order only after it knows the final line boundaries.
    pub fn shape_multilingual_text(
        &mut self,
        segment: TextSegment,
        language: Option<&str>,
        base_direction: TextDirection,
        no_wrap: bool,
    ) -> Result<Vec<MultilingualTextSegment>> {
        if segment.text.is_empty() {
            return Ok(Vec::new());
        }

        let paragraph_level = match base_direction {
            TextDirection::Auto => None,
            TextDirection::LeftToRight => Some(unicode_bidi::Level::ltr()),
            TextDirection::RightToLeft => Some(unicode_bidi::Level::rtl()),
        };
        let bidi = unicode_bidi::BidiInfo::new(&segment.text, paragraph_level);
        let levels = bidi.levels.clone();
        let break_offsets = if no_wrap {
            HashSet::new()
        } else {
            multilingual_break_opportunities(&segment.text)
        };
        self.shape_multilingual_with_levels(
            segment,
            language,
            (&break_offsets, &break_offsets),
            &levels,
            0,
            (false, false, None, false, &mut Vec::new()),
        )
    }

    /// Shape styled spans with one paragraph-wide bidi resolution.
    pub fn shape_multilingual_paragraph(
        &mut self,
        segments: Vec<(TextSegment, Option<String>)>,
        base_direction: TextDirection,
        no_wrap: bool,
    ) -> Result<Vec<MultilingualTextSegment>> {
        let requests = vec![(false, false); segments.len()];
        self.shape_multilingual_paragraph_with_caps(segments, base_direction, no_wrap, &requests)
            .map(|(spans, _)| spans)
    }

    /// Shape PowerPoint small-caps requests in the same paragraph-wide bidi pass.
    /// Requests correspond to input occurrences, including equal-text runs.
    /// Logical text and source clusters remain the original input. Unsupported
    /// selected faces use uppercase lowercase graphemes at 80 percent size.
    pub fn shape_multilingual_paragraph_with_caps(
        &mut self,
        segments: Vec<(TextSegment, Option<String>)>,
        base_direction: TextDirection,
        no_wrap: bool,
        requests: &[(bool, bool)],
    ) -> Result<(Vec<MultilingualTextSegment>, Vec<String>)> {
        self.shape_paragraph_with_features(segments, base_direction, no_wrap, requests, None, false)
    }

    /// Shape explicit occurrence-aligned kerning without changing neutral callers.
    #[doc(hidden)]
    pub fn shape_paragraph_with_kerning(
        &mut self,
        segments: Vec<(TextSegment, Option<String>)>,
        base_direction: TextDirection,
        no_wrap: bool,
        enabled: &[bool],
        legacy_shaping: bool,
    ) -> Result<Vec<MultilingualTextSegment>> {
        if enabled.len() != segments.len() {
            return Err(LayoutError::Shaping(
                "kerning request count differs from styled runs".into(),
            ));
        }
        let caps = vec![(false, false); segments.len()];
        self.shape_paragraph_with_features(
            segments,
            base_direction,
            no_wrap,
            &caps,
            Some(enabled),
            legacy_shaping,
        )
        .map(|(spans, _)| spans)
    }

    fn shape_paragraph_with_features(
        &mut self,
        segments: Vec<(TextSegment, Option<String>)>,
        base_direction: TextDirection,
        no_wrap: bool,
        requests: &[(bool, bool)],
        kerning: Option<&[bool]>,
        legacy_shaping: bool,
    ) -> Result<(Vec<MultilingualTextSegment>, Vec<String>)> {
        if requests.iter().any(|(small, all)| *small && *all) {
            return Err(LayoutError::Shaping(
                "conflicting all-caps and small-caps request".to_owned(),
            ));
        }
        if requests.len() != segments.len() {
            return Err(LayoutError::Shaping(
                "small-caps request count differs from styled runs".to_owned(),
            ));
        }
        let mut diagnostics = Vec::new();
        let paragraph_text = segments
            .iter()
            .map(|(segment, _)| segment.text.as_str())
            .collect::<String>();
        let mut paragraph_offset = 0usize;
        let segment_starts = segments
            .iter()
            .map(|(segment, _)| {
                let start = paragraph_offset;
                paragraph_offset += segment.text.len();
                start
            })
            .collect::<Vec<_>>();
        if paragraph_text.is_empty() && kerning.is_none() {
            return Ok((Vec::new(), diagnostics));
        }
        let paragraph_level = match base_direction {
            TextDirection::Auto => None,
            TextDirection::LeftToRight => Some(unicode_bidi::Level::ltr()),
            TextDirection::RightToLeft => Some(unicode_bidi::Level::rtl()),
        };
        let bidi = unicode_bidi::BidiInfo::new(&paragraph_text, paragraph_level);
        let paragraph_level = bidi
            .paragraphs
            .first()
            .map(|paragraph| paragraph.level)
            .unwrap_or_else(unicode_bidi::Level::ltr);
        // Opportunities come from the whole paragraph, so a word that changes
        // formatting partway through does not break where its runs meet.
        let paragraph_breaks = if no_wrap {
            HashSet::new()
        } else {
            multilingual_break_opportunities(&paragraph_text)
        };
        let mut logical_index = 0usize;
        let mut shaped = Vec::new();
        for (occurrence, (((segment, language), byte_offset), (small_caps, all_caps))) in segments
            .into_iter()
            .zip(segment_starts)
            .zip(requests.iter().copied())
            .enumerate()
        {
            let byte_end = byte_offset + segment.text.len();
            if segment.text.is_empty()
                && let Some(kerning) = kerning
            {
                let base = segment;
                let mut span = MultilingualTextSegment::new_with_legacy_shaping(
                    base,
                    logical_index,
                    language,
                    TextScript::Common,
                    base_direction,
                    paragraph_level.number(),
                    Vec::new(),
                    Vec::new(),
                    Vec::new(),
                    Vec::new(),
                    Vec::new(),
                    false,
                    legacy_shaping,
                )?;
                span.kerning = Some(kerning[occurrence]);
                logical_index += 1;
                shaped.push(span);
                continue;
            }
            if !segment.text.is_empty() {
                let forced_levels = (segment.direction != TextDirection::Auto)
                    .then(|| {
                        explicit_direction_levels(&segment.text, segment.direction, paragraph_level)
                    })
                    .transpose()?;
                let levels = forced_levels
                    .as_deref()
                    .unwrap_or(&bidi.levels[byte_offset..byte_end]);
                let shaping_offsets = if legacy_shaping {
                    unicode_linebreak::linebreaks(&segment.text)
                        .map(|(offset, _)| offset)
                        .collect()
                } else {
                    paragraph_breaks
                        .iter()
                        .filter(|offset| (byte_offset + 1..=byte_end).contains(*offset))
                        .map(|offset| offset - byte_offset)
                        .collect()
                };
                let published_offsets = if no_wrap {
                    &paragraph_breaks
                } else {
                    &shaping_offsets
                };
                let spans = self.shape_multilingual_with_levels(
                    segment,
                    language.as_deref(),
                    (&shaping_offsets, published_offsets),
                    levels,
                    logical_index,
                    (
                        small_caps,
                        all_caps,
                        kerning.map(|values| values[occurrence]),
                        legacy_shaping,
                        &mut diagnostics,
                    ),
                )?;
                logical_index += spans.len();
                shaped.extend(spans);
            }
        }
        Ok((shaped, diagnostics))
    }

    /// Shaping cuts and published line opportunities use segment byte offsets.
    /// Legacy no-wrap retains old shaping chunks without permitting soft breaks.
    fn shape_multilingual_with_levels(
        &mut self,
        segment: TextSegment,
        language: Option<&str>,
        break_offsets: (&HashSet<usize>, &HashSet<usize>),
        levels: &[unicode_bidi::Level],
        logical_index_base: usize,
        caps: (bool, bool, Option<bool>, bool, &mut Vec<String>),
    ) -> Result<Vec<MultilingualTextSegment>> {
        let (small_caps, all_caps, kerning, legacy_shaping, diagnostics) = caps;
        let (shaping_offsets, published_offsets) = break_offsets;
        let grapheme_boundaries = icu_segmenter::GraphemeClusterSegmenter::new()
            .segment_str(&segment.text)
            .collect::<Vec<_>>();
        // Spaces before a break opportunity get a span of their own, so a
        // line can let them hang past its end.
        let hanging_starts = if legacy_shaping {
            HashSet::new()
        } else {
            shaping_offsets
                .iter()
                .map(|offset| segment.text[..*offset].trim_end_matches(' ').len())
                .collect::<HashSet<_>>()
        };

        let mut logical_ranges = Vec::<(usize, usize, TextScript, unicode_bidi::Level)>::new();
        let mut start = 0usize;
        let mut current_script = TextScript::Common;
        let mut current_level = levels[0];
        for window in grapheme_boundaries.windows(2) {
            let grapheme_start = window[0];
            let grapheme_end = window[1];
            let script = script_for_grapheme(&segment.text[grapheme_start..grapheme_end]);
            let script = if script == TextScript::Common {
                current_script
            } else {
                script
            };
            let level = levels[grapheme_start];
            if grapheme_start > start
                && ((!legacy_shaping && script != current_script) || level != current_level)
            {
                logical_ranges.push((start, grapheme_start, current_script, current_level));
                start = grapheme_start;
            }
            current_script = script;
            current_level = level;
            if (shaping_offsets.contains(&grapheme_end) || hanging_starts.contains(&grapheme_end))
                && grapheme_end < segment.text.len()
            {
                logical_ranges.push((start, grapheme_end, current_script, current_level));
                start = grapheme_end;
            }
        }
        if start < segment.text.len() {
            logical_ranges.push((start, segment.text.len(), current_script, current_level));
        }

        let mut font_ranges = Vec::new();
        for (start, end, script, level) in logical_ranges {
            let boundaries = icu_segmenter::GraphemeClusterSegmenter::new()
                .segment_str(&segment.text[start..end])
                .map(|offset| start + offset)
                .collect::<Vec<_>>();
            let mut range_start = start;
            let mut range_font = None;
            for window in boundaries.windows(2) {
                let grapheme_start = window[0];
                let grapheme_end = window[1];
                let font_id = if legacy_shaping {
                    segment.font_id
                } else {
                    self.font_for_multilingual_span(
                        segment.font_id,
                        &segment.text[grapheme_start..grapheme_end],
                        segment.bold,
                        segment.italic,
                    )
                };
                if let Some(current_font) = range_font
                    && current_font != font_id
                {
                    font_ranges.push((range_start, grapheme_start, script, level, current_font));
                    range_start = grapheme_start;
                }
                range_font = Some(font_id);
            }
            if let Some(font_id) = range_font {
                font_ranges.push((range_start, end, script, level, font_id));
            }
        }

        let mut cap_ranges = Vec::new();
        for (start, end, script, level, font_id) in font_ranges {
            let direction = if level.is_rtl() {
                TextDirection::RightToLeft
            } else {
                TextDirection::LeftToRight
            };
            let text = &segment.text[start..end];
            let lowercase = text.to_lowercase();
            // Test actual feature application on this face, script and language.
            // A global GSUB tag alone does not prove a reachable substitution.
            let has_lowercase = text.chars().any(char::is_lowercase);
            let feature_applies = if small_caps && has_lowercase {
                let plain = self.shape_explicit(
                    font_id,
                    &lowercase,
                    segment.font_size,
                    script,
                    language,
                    direction,
                )?;
                let caps = self.shape_explicit_with_small_caps(
                    font_id,
                    &lowercase,
                    segment.font_size,
                    (script, language, direction),
                    true,
                    kerning,
                )?;
                plain.glyph_ids != caps.glyph_ids
            } else {
                false
            };
            if all_caps {
                cap_ranges.push((
                    start,
                    end,
                    script,
                    level,
                    font_id,
                    segment.font_size,
                    false,
                    true,
                ));
            } else if !small_caps || !has_lowercase || feature_applies {
                cap_ranges.push((
                    start,
                    end,
                    script,
                    level,
                    font_id,
                    segment.font_size,
                    feature_applies,
                    false,
                ));
            } else {
                let font = self.get_font(font_id)?;
                let message = format!(
                    "small caps: selected face '{}' lacks an applicable smcp substitution for {:?}; using 80 percent uppercase fallback",
                    font.family, script
                );
                if !diagnostics.contains(&message) {
                    diagnostics.push(message);
                }
                // Keep every original grapheme together. An expanding uppercase
                // mapping belongs to that original grapheme's logical cluster.
                let boundaries = icu_segmenter::GraphemeClusterSegmenter::new()
                    .segment_str(text)
                    .collect::<Vec<_>>();
                for window in boundaries.windows(2) {
                    let lo = start + window[0];
                    let hi = start + window[1];
                    let lower = segment.text[lo..hi].chars().any(char::is_lowercase);
                    cap_ranges.push((
                        lo,
                        hi,
                        script,
                        level,
                        font_id,
                        segment.font_size * if lower { 0.8 } else { 1.0 },
                        false,
                        lower,
                    ));
                }
            }
        }
        let mut logical = Vec::with_capacity(cap_ranges.len());
        for (
            logical_index,
            (start, end, script, level, font_id, size_pt, feature_applies, fallback),
        ) in cap_ranges.into_iter().enumerate()
        {
            let text = &segment.text[start..end];
            let metrics = self.metrics(font_id, size_pt)?;
            let direction = if level.is_rtl() {
                TextDirection::RightToLeft
            } else {
                TextDirection::LeftToRight
            };
            let rendered = if fallback {
                text.to_uppercase()
            } else {
                text.to_owned()
            };
            let mut positioned = if legacy_shaping {
                self.shape_legacy(font_id, &rendered, size_pt, kerning, true)?
            } else {
                self.shape_explicit_with_small_caps(
                    font_id,
                    &rendered,
                    size_pt,
                    (script, language, direction),
                    feature_applies,
                    kerning,
                )?
            };
            if fallback {
                let mut owners = Vec::new();
                for (index, character) in text.chars().enumerate() {
                    owners.extend(character.to_uppercase().map(|_| index as u32));
                }
                for cluster in &mut positioned.clusters {
                    let lo = cluster.char_start as usize;
                    let hi = cluster.char_end as usize;
                    cluster.char_start = *owners.get(lo).ok_or_else(|| {
                        LayoutError::Shaping(
                            "uppercase cluster starts outside its logical owner".to_owned(),
                        )
                    })?;
                    cluster.char_end = owners
                        .get(hi.checked_sub(1).ok_or_else(|| {
                            LayoutError::Shaping("empty uppercase cluster".to_owned())
                        })?)
                        .ok_or_else(|| {
                            LayoutError::Shaping(
                                "uppercase cluster ends outside its logical owner".to_owned(),
                            )
                        })?
                        + 1;
                }
                let mut merged: Vec<GlyphCluster> = Vec::new();
                for cluster in positioned.clusters {
                    if let Some(last) = merged.last_mut()
                        && last.char_start < cluster.char_end
                        && cluster.char_start < last.char_end
                    {
                        last.glyph_end = cluster.glyph_end;
                        last.char_start = last.char_start.min(cluster.char_start);
                        last.char_end = last.char_end.max(cluster.char_end);
                    } else {
                        merged.push(cluster);
                    }
                }
                positioned.clusters = merged;
            }
            let char_start = segment.text[..start].chars().count() as u32;
            let char_end = segment.text[..end].chars().count() as u32;
            let source = segment.source.map(|source| SourceSpan {
                node: source.node,
                char_start: source.char_start + char_start,
                char_end: source.char_start + char_end,
            });
            let mut base = segment.clone();
            base.text = text.to_owned();
            base.source = source;
            base.font_id = font_id;
            base.font_size = size_pt;
            base.glyph_ids = positioned.glyph_ids;
            base.advances = positioned.x_advances.clone();
            base.width = positioned.x_advances.iter().sum();
            if !legacy_shaping {
                base.ascent = metrics.ascent;
                base.descent = metrics.descent;
                base.line_gap = metrics.line_gap;
            }
            let mut span = MultilingualTextSegment::new_with_legacy_shaping(
                base,
                logical_index_base + logical_index,
                language.map(str::to_owned),
                script,
                direction,
                level.number(),
                positioned.x_advances,
                positioned.y_advances,
                positioned.x_offsets,
                positioned.y_offsets,
                positioned.clusters,
                published_offsets.contains(&end),
                legacy_shaping,
            )?;
            span.kerning = kerning;
            logical.push(span);
        }

        Ok(logical)
    }

    fn font_for_multilingual_span(
        &mut self,
        preferred: FontId,
        text: &str,
        bold: bool,
        italic: bool,
    ) -> FontId {
        let Some(index) = self.index_of(preferred) else {
            return preferred;
        };
        if self.uncovered(index, text).is_empty() {
            return preferred;
        }
        let required = text
            .chars()
            .filter(|character| !character.is_whitespace() && !character.is_control())
            .collect::<Vec<_>>();
        self.font_covering(&required, bold, italic)
            .unwrap_or(preferred)
    }

    fn shape_in_context(
        &self,
        owner: &MultilingualTextSegment,
        text: &str,
    ) -> Result<PositionedShape> {
        let base = owner.base();
        if owner.legacy_shaping {
            self.shape_legacy(base.font_id, text, base.font_size, owner.kerning, true)
        } else {
            self.shape_explicit_with_small_caps(
                base.font_id,
                text,
                base.font_size,
                (owner.script, owner.language.as_deref(), owner.direction),
                false,
                owner.kerning,
            )
        }
    }

    pub(crate) fn context_character_spacing(&self, owner: &MultilingualTextSegment) -> Result<f64> {
        let shape = self.shape_in_context(owner, owner.text())?;
        Ok(
            if shape.x_advances.len() == owner.x_advances.len() && !shape.x_advances.is_empty() {
                (owner.width() - shape.x_advances.iter().sum::<f64>())
                    / shape.x_advances.len() as f64
            } else {
                0.0
            },
        )
    }

    pub(crate) fn context_prefix_width(
        &self,
        owner: &MultilingualTextSegment,
        byte_end: usize,
    ) -> Result<f64> {
        let characters = owner.text()[..byte_end].chars().count() as u32;
        if owner
            .clusters
            .iter()
            .all(|cluster| cluster.char_end <= characters || cluster.char_start >= characters)
        {
            return Ok(owner
                .clusters
                .iter()
                .filter(|cluster| cluster.char_end <= characters)
                .flat_map(|cluster| {
                    &owner.x_advances[cluster.glyph_start as usize..cluster.glyph_end as usize]
                })
                .sum());
        }
        let shape = self.shape_in_context(owner, &owner.text()[..byte_end])?;
        Ok(shape.x_advances.iter().sum::<f64>()
            + self.context_character_spacing(owner)? * shape.x_advances.len() as f64)
    }

    pub(crate) fn reshape_context_segment(
        &self,
        owner: &MultilingualTextSegment,
        mut base: TextSegment,
        spacing: f64,
    ) -> Result<MultilingualTextSegment> {
        let positioned = if owner.legacy_shaping {
            self.shape_legacy(
                base.font_id,
                &base.text,
                base.font_size,
                owner.kerning,
                true,
            )?
        } else {
            self.shape_explicit_with_small_caps(
                base.font_id,
                &base.text,
                base.font_size,
                (owner.script, owner.language.as_deref(), owner.direction),
                false,
                owner.kerning,
            )?
        };
        let x_advances = positioned
            .x_advances
            .iter()
            .map(|advance| advance + spacing)
            .collect::<Vec<_>>();
        base.glyph_ids = positioned.glyph_ids;
        base.advances = x_advances.clone();
        base.width = x_advances.iter().sum();
        MultilingualTextSegment::new_with_legacy_shaping(
            base,
            owner.logical_index,
            owner.language.clone(),
            owner.script,
            owner.direction,
            owner.bidi_level,
            x_advances,
            positioned.y_advances,
            positioned.x_offsets,
            positioned.y_offsets,
            positioned.clusters,
            owner.break_after,
            owner.legacy_shaping,
        )
        .map(|span| span.inherit_shaping_context(owner))
    }

    fn shape_legacy(
        &self,
        font_id: FontId,
        text: &str,
        size_pt: f64,
        kerning: Option<bool>,
        rich: bool,
    ) -> Result<PositionedShape> {
        let font = self.get_font(font_id)?;
        let face = harfrust::FontRef::from_index(&font.data, font.face_index)
            .map_err(|error| LayoutError::Shaping(format!("failed to read font face: {error}")))?;
        let shaper = font
            .shaper_data
            .shaper(&face)
            .instance(font.shaper_instance.as_ref())
            .build();
        let mut buffer = harfrust::UnicodeBuffer::new();
        buffer.push_str(text);
        buffer.guess_segment_properties();
        let feature = kerning.map(|enabled| {
            harfrust::Feature::new(harfrust::Tag::new(b"kern"), u32::from(enabled), ..)
        });
        let options = match feature.as_ref() {
            Some(feature) => {
                harfrust::ShapeOptions::default().features(std::slice::from_ref(feature))
            }
            None => harfrust::ShapeOptions::default(),
        };
        let output = shaper.shape(buffer, options);
        let infos = output.glyph_infos();
        let scale = size_pt / f64::from(font.units_per_em);
        Ok(PositionedShape {
            glyph_ids: infos.iter().map(|info| info.glyph_id as u16).collect(),
            x_advances: output
                .glyph_positions()
                .iter()
                .map(|position| f64::from(position.x_advance) * scale)
                .collect(),
            y_advances: if rich {
                vec![0.0; infos.len()]
            } else {
                Vec::new()
            },
            x_offsets: if rich {
                vec![0.0; infos.len()]
            } else {
                Vec::new()
            },
            y_offsets: if rich {
                vec![0.0; infos.len()]
            } else {
                Vec::new()
            },
            clusters: if rich {
                glyph_clusters(infos, text)
            } else {
                Vec::new()
            },
        })
    }

    fn shape_explicit(
        &self,
        font_id: FontId,
        text: &str,
        size_pt: f64,
        script: TextScript,
        language: Option<&str>,
        direction: TextDirection,
    ) -> Result<PositionedShape> {
        self.shape_explicit_with_small_caps(
            font_id,
            text,
            size_pt,
            (script, language, direction),
            false,
            None,
        )
    }

    fn shape_explicit_with_small_caps(
        &self,
        font_id: FontId,
        text: &str,
        size_pt: f64,
        properties: (TextScript, Option<&str>, TextDirection),
        small_caps: bool,
        kerning: Option<bool>,
    ) -> Result<PositionedShape> {
        let (script, language, direction) = properties;
        let font = self.get_font(font_id)?;
        let face = harfrust::FontRef::from_index(&font.data, font.face_index)
            .map_err(|error| LayoutError::Shaping(format!("failed to read font face: {error}")))?;
        let shaper = font
            .shaper_data
            .shaper(&face)
            .instance(font.shaper_instance.as_ref())
            .build();
        let mut buffer = harfrust::UnicodeBuffer::new();
        buffer.push_str(text);
        buffer.set_script(harfrust_script(script));
        buffer.set_direction(match direction {
            TextDirection::RightToLeft => harfrust::Direction::RightToLeft,
            TextDirection::Auto | TextDirection::LeftToRight => harfrust::Direction::LeftToRight,
        });
        if let Some(language) = language.and_then(harfrust::Language::new) {
            buffer.set_language(language);
        }
        let mut features = Vec::with_capacity(2);
        if small_caps {
            features.push(harfrust::Feature::new(harfrust::Tag::new(b"smcp"), 1, ..));
        }
        if let Some(enabled) = kerning {
            features.push(harfrust::Feature::new(
                harfrust::Tag::new(b"kern"),
                u32::from(enabled),
                ..,
            ));
        }
        let options = harfrust::ShapeOptions::default().features(&features);
        let output = shaper.shape(buffer, options);
        let infos = output.glyph_infos();
        let positions = output.glyph_positions();
        let scale = size_pt / f64::from(font.units_per_em);
        let glyph_ids = infos
            .iter()
            .map(|info| info.glyph_id as u16)
            .collect::<Vec<_>>();
        let x_advances = positions
            .iter()
            .map(|pos| f64::from(pos.x_advance) * scale)
            .collect();
        let y_advances = positions
            .iter()
            .map(|pos| f64::from(pos.y_advance) * scale)
            .collect();
        let x_offsets = positions
            .iter()
            .map(|pos| f64::from(pos.x_offset) * scale)
            .collect();
        let y_offsets = positions
            .iter()
            .map(|pos| f64::from(pos.y_offset) * scale)
            .collect();
        let clusters = glyph_clusters(infos, text);
        Ok(PositionedShape {
            glyph_ids,
            x_advances,
            y_advances,
            x_offsets,
            y_offsets,
            clusters,
        })
    }

    /// Get font data for PDF embedding.
    pub fn font_data(&self, font_id: FontId) -> Result<crate::output::FontData> {
        let font = self.get_font(font_id)?;
        Ok(crate::output::FontData {
            id: font.id,
            family: font.family.clone(),
            data: Arc::clone(&font.data),
            face_index: font.face_index,
            bold: font.bold,
            italic: font.italic,
            variations: font.variations.clone(),
            synthetic_bold: font.synthetic_bold,
            synthetic_italic: font.synthetic_italic,
        })
    }

    /// Get all used font data.
    pub fn all_font_data(&self) -> Vec<crate::output::FontData> {
        self.fonts
            .iter()
            .map(|f| crate::output::FontData {
                id: f.id,
                family: f.family.clone(),
                data: Arc::clone(&f.data),
                face_index: f.face_index,
                bold: f.bold,
                italic: f.italic,
                variations: f.variations.clone(),
                synthetic_bold: f.synthetic_bold,
                synthetic_italic: f.synthetic_italic,
            })
            .collect()
    }

    fn get_font(&self, font_id: FontId) -> Result<&LoadedFont> {
        self.fonts
            .iter()
            .find(|f| f.id == font_id)
            .ok_or_else(|| LayoutError::FontNotFound(format!("FontId({}) not loaded", font_id.0)))
    }

    fn remember_font_key(&mut self, key: FontKey, index: usize) {
        if self.cache.len() < RESOLUTION_CACHE_MAX_ENTRIES {
            self.cache.insert(key, index);
        }
    }

    fn remember_coverage_fallback(&mut self, bold: bool, italic: bool, index: usize) {
        let known = self.coverage_fallbacks.entry((bold, italic)).or_default();
        if known.len() < COVERAGE_FALLBACK_MAX_ENTRIES && !known.contains(&index) {
            known.push(index);
        }
    }

    fn remember_coverage_misses(&mut self, missing: &[char]) {
        for &ch in missing {
            if self.coverage_misses.len() >= COVERAGE_MISS_MAX_ENTRIES {
                break;
            }
            self.coverage_misses.insert(ch);
        }
    }

    fn record_layout_font(&mut self, font_id: FontId) {
        if let Some(trace) = self.paragraph_font_trace.as_mut() {
            if !trace.overflowed && trace.events.len() < PARAGRAPH_FONT_TRACE_MAX_ENTRIES {
                trace.events.push(FontTraceEvent::Resolution(font_id));
                if trace
                    .events
                    .capacity()
                    .saturating_mul(std::mem::size_of::<FontTraceEvent>())
                    .saturating_add(trace.owned_family_bytes)
                    > PARAGRAPH_FONT_TRACE_MAX_BYTES
                {
                    trace.events = Vec::new();
                    trace.owned_family_bytes = 0;
                    trace.overflowed = true;
                }
            } else if !trace.overflowed {
                trace.events = Vec::new();
                trace.owned_family_bytes = 0;
                trace.overflowed = true;
            }
        }
        if !self.layout_fonts.contains(&font_id) {
            self.layout_fonts.push(font_id);
        }
    }

    #[cfg(test)]
    fn shaping_memo_counts(&self) -> (usize, usize, usize, usize) {
        let memo = self
            .shaping_memo
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner);
        (memo.hits, memo.misses, memo.entries.len(), memo.bytes)
    }
}

struct PositionedShape {
    glyph_ids: Vec<u16>,
    x_advances: Vec<f64>,
    y_advances: Vec<f64>,
    x_offsets: Vec<f64>,
    y_offsets: Vec<f64>,
    clusters: Vec<GlyphCluster>,
}

fn script_for_grapheme(grapheme: &str) -> TextScript {
    grapheme
        .chars()
        .map(script_for_char)
        .find(|script| *script != TextScript::Common)
        .unwrap_or(TextScript::Common)
}

fn script_for_char(character: char) -> TextScript {
    match character as u32 {
        0x0041..=0x024f | 0x1e00..=0x1eff => TextScript::Latin,
        0x0590..=0x05ff => TextScript::Hebrew,
        0x0600..=0x06ff | 0x0750..=0x077f | 0x08a0..=0x08ff => TextScript::Arabic,
        0x0900..=0x097f | 0xa8e0..=0xa8ff => TextScript::Devanagari,
        0x0e00..=0x0e7f => TextScript::Thai,
        // Hangul and Kana are deliberately not folded into Han. They need
        // different shaper feature sets, and folding them would silently
        // change how existing Han text shapes.
        //
        // Hangul is the jamo, the Extended-A and Extended-B jamo, the
        // compatibility jamo, the precomposed syllables, and the halfwidth
        // jamo. Kana is Hiragana, Katakana, the Katakana phonetic
        // extensions, and the halfwidth Katakana.
        0x1100..=0x11ff | 0x3130..=0x318f | 0xa960..=0xa97f | 0xac00..=0xd7ff | 0xffa0..=0xffdc => {
            TextScript::Hangul
        }
        0x3040..=0x30ff | 0x31f0..=0x31ff | 0xff66..=0xff9f => TextScript::Kana,
        0x3400..=0x4dbf | 0x4e00..=0x9fff | 0xf900..=0xfaff => TextScript::Han,
        _ => TextScript::Common,
    }
}

fn harfrust_script(script: TextScript) -> harfrust::Script {
    match script {
        TextScript::Latin => harfrust::script::LATIN,
        TextScript::Arabic => harfrust::script::ARABIC,
        TextScript::Hebrew => harfrust::script::HEBREW,
        TextScript::Devanagari => harfrust::script::DEVANAGARI,
        TextScript::Thai => harfrust::script::THAI,
        TextScript::Han => harfrust::script::HAN,
        TextScript::Hangul => harfrust::script::HANGUL,
        // HarfBuzz unifies Hiragana and Katakana under one shaping model, and
        // HIRAGANA is the tag it canonicalises the pair to.
        TextScript::Kana => harfrust::script::HIRAGANA,
        TextScript::Common => harfrust::script::COMMON,
    }
}

/// UAX 14 line break opportunities, plus dictionary word boundaries inside
/// Thai, Lao, Khmer and Myanmar text, which UAX 14 leaves to a dictionary.
///
/// A word boundary elsewhere is not a line break opportunity: it falls before
/// a comma, a space or the hyphen of a compound.
fn multilingual_break_opportunities(text: &str) -> HashSet<usize> {
    let complex_context = |character: Option<char>| {
        character.is_some_and(|character| {
            unicode_linebreak::break_property(character as u32)
                == unicode_linebreak::BreakClass::ComplexContext
        })
    };
    let mut opportunities = icu_segmenter::WordSegmenter::new_auto(Default::default())
        .segment_str(text)
        .filter(|offset| {
            *offset > 0
                && *offset < text.len()
                && complex_context(text[..*offset].chars().next_back())
                && complex_context(text[*offset..].chars().next())
        })
        .collect::<HashSet<_>>();
    for (offset, _) in unicode_linebreak::linebreaks(text) {
        if offset > 0 {
            opportunities.insert(offset);
        }
    }
    opportunities.retain(|offset| {
        let before = text[..*offset].chars().next_back();
        let after = text[*offset..].chars().next();
        before
            .zip(after)
            .is_none_or(|(before, after)| crate::line::multilingual_break_allowed(before, after))
    });
    opportunities
}

fn glyph_clusters(infos: &[harfrust::GlyphInfo], text: &str) -> Vec<GlyphCluster> {
    let mut byte_starts = infos
        .iter()
        .map(|info| info.cluster as usize)
        .collect::<Vec<_>>();
    byte_starts.push(text.len());
    byte_starts.sort_unstable();
    byte_starts.dedup();

    let mut clusters = Vec::new();
    let mut glyph_start = 0usize;
    while glyph_start < infos.len() {
        let cluster_byte = infos[glyph_start].cluster as usize;
        let mut glyph_end = glyph_start + 1;
        while glyph_end < infos.len() && infos[glyph_end].cluster as usize == cluster_byte {
            glyph_end += 1;
        }
        let byte_end = byte_starts
            .iter()
            .copied()
            .find(|candidate| *candidate > cluster_byte)
            .unwrap_or(text.len());
        clusters.push(GlyphCluster {
            glyph_start: glyph_start as u32,
            glyph_end: glyph_end as u32,
            char_start: text[..cluster_byte].chars().count() as u32,
            char_end: text[..byte_end].chars().count() as u32,
        });
        glyph_start = glyph_end;
    }
    clusters
}

fn best_caller_face(
    db: &fontdb::Database,
    ids: &[fontdb::ID],
    weight: fontdb::Weight,
    style: fontdb::Style,
) -> Option<fontdb::ID> {
    const CANDIDATE_FAMILY: &str = "__rdocx_caller_candidate__";

    let mut candidates = fontdb::Database::new();
    let mut candidate_ids = Vec::with_capacity(ids.len());
    for id in ids.iter().rev() {
        let mut face = db.face(*id)?.clone();
        for (family, _) in &mut face.families {
            CANDIDATE_FAMILY.clone_into(family);
        }
        let candidate_id = candidates.push_face_info(face);
        candidate_ids.push((candidate_id, *id));
    }

    let selected = candidates.query(&fontdb::Query {
        families: &[fontdb::Family::Name(CANDIDATE_FAMILY)],
        weight,
        style,
        stretch: fontdb::Stretch::Normal,
    })?;
    candidate_ids
        .into_iter()
        .find_map(|(candidate, original)| (candidate == selected).then_some(original))
}

fn bundled_font_database() -> fontdb::Database {
    let mut db = fontdb::Database::new();
    for (_family, data) in crate::bundled_fonts::bundled_font_data() {
        db.load_font_data(data.to_vec());
    }
    db
}

fn font_data_for_face(
    db: &fontdb::Database,
    id: fontdb::ID,
    memory_face_data: &mut HashMap<fontdb::ID, Arc<[u8]>>,
) -> Option<(Arc<[u8]>, u32)> {
    let face = db.face(id)?;
    let face_index = face.index;
    match &face.source {
        fontdb::Source::Binary(data) => match memory_face_data.get(&id) {
            Some(data) => Some((Arc::clone(data), face_index)),
            None => {
                // Collection faces share one original source allocation. Match
                // source identity rather than copying identical bytes per face.
                let shared = db.faces().find_map(|other| match &other.source {
                    fontdb::Source::Binary(source) if Arc::ptr_eq(source, data) => {
                        memory_face_data.get(&other.id).cloned()
                    }
                    _ => None,
                });
                let data: Arc<[u8]> =
                    shared.unwrap_or_else(|| Arc::from(data.as_ref().as_ref().to_vec()));
                memory_face_data.insert(id, Arc::clone(&data));
                Some((data, face_index))
            }
        },
        #[cfg(feature = "system-fonts")]
        fontdb::Source::File(path) => shared_file_font_bytes(path).map(|data| (data, face_index)),
        // A source that fontdb features enabled elsewhere in the build add,
        // such as `memmap`'s shared file, read through fontdb itself.
        #[allow(unreachable_patterns)]
        _ => db.with_face_data(id, |data, index| (Arc::from(data), index)),
    }
}

#[cfg(feature = "system-fonts")]
fn shared_file_font_bytes(path: &Path) -> Option<Arc<[u8]>> {
    let cache = FILE_FONT_CACHE.get_or_init(|| Mutex::new(FileFontCache::new()));
    shared_file_font_bytes_from_cache(cache, path)
}

#[cfg(feature = "system-fonts")]
fn shared_file_font_bytes_from_cache(
    cache_lock: &Mutex<FileFontCache>,
    path: &Path,
) -> Option<Arc<[u8]>> {
    let identity = std::fs::canonicalize(path).unwrap_or_else(|_| path.to_path_buf());
    let mut cache = match cache_lock.lock() {
        Ok(cache) => cache,
        Err(poisoned) => {
            let mut cache = poisoned.into_inner();
            cache.clear();
            cache_lock.clear_poison();
            cache
        }
    };
    if let Some(index) = cache
        .entries
        .iter()
        .position(|(candidate, _, _)| candidate == &identity)
    {
        let entry = cache.entries.remove(index).expect("cache index exists");
        let bytes = Arc::clone(&entry.1);
        cache.entries.push_back(entry);
        return Some(bytes);
    }

    let bytes: Arc<[u8]> = Arc::from(std::fs::read(&identity).ok()?);
    cache_file_font_bytes(&mut cache, identity, bytes)
}

#[cfg(feature = "system-fonts")]
fn cache_file_font_bytes(
    cache: &mut FileFontCache,
    identity: PathBuf,
    bytes: Arc<[u8]>,
) -> Option<Arc<[u8]>> {
    let entry_bytes = std::mem::size_of::<(PathBuf, Arc<[u8]>, usize)>()
        .saturating_add(identity.as_os_str().len())
        .saturating_add(bytes.len());
    if entry_bytes <= FILE_FONT_CACHE_MAX_BYTES {
        while cache.entries.len() >= FILE_FONT_CACHE_MAX_ENTRIES
            || cache.bytes.saturating_add(entry_bytes) > FILE_FONT_CACHE_MAX_BYTES
        {
            let Some((_, _, evicted_bytes)) = cache.entries.pop_front() else {
                break;
            };
            cache.bytes = cache.bytes.saturating_sub(evicted_bytes);
        }
        cache.bytes += entry_bytes;
        cache
            .entries
            .push_back((identity, Arc::clone(&bytes), entry_bytes));
    }
    Some(bytes)
}

/// Map common Word font names to metric-compatible alternatives.
/// Returns a list of candidate names to try (including the original).
///
/// Priority: original font → metric-compatible open-source clone → generic fallback.
/// Carlito is metric-compatible with Calibri, Caladea with Cambria,
/// Liberation Sans/Serif/Mono with Arial/Times New Roman/Courier New.
fn map_font_name(name: &str) -> &[&str] {
    match name {
        "Calibri" => &["Calibri", "Carlito"],
        "Calibri Light" => &["Calibri Light", "Carlito"],
        "Cambria" => &["Cambria", "Caladea"],
        "Cambria Math" => &["Cambria Math", "Cambria", "Caladea"],
        "Arial" => &["Arial", "Liberation Sans", "Helvetica"],
        "Times New Roman" => &["Times New Roman", "Liberation Serif", "Times"],
        "Courier New" => &["Courier New", "Liberation Mono", "Courier"],
        "Consolas" => &["Consolas", "Liberation Mono", "DejaVu Sans Mono"],
        "Segoe UI" => &["Segoe UI", "Selawik", "Carlito", "Liberation Sans"],
        "Tahoma" => &["Tahoma", "Liberation Sans", "Helvetica"],
        "Verdana" => &["Verdana", "Liberation Sans", "DejaVu Sans"],
        "Georgia" => &["Georgia", "Gelasio", "Caladea", "Liberation Serif"],
        "Palatino Linotype" => &["Palatino Linotype", "Palatino", "Liberation Serif"],
        "Book Antiqua" => &["Book Antiqua", "Palatino", "Liberation Serif"],
        "Garamond" => &["Garamond", "Caladea", "Liberation Serif"],
        "Trebuchet MS" => &["Trebuchet MS", "Liberation Sans", "DejaVu Sans"],
        "Impact" => &["Impact", "Liberation Sans", "Arial"],
        "Comic Sans MS" => &["Comic Sans MS", "Liberation Sans", "DejaVu Sans"],
        "Symbol" => &["Symbol", "DejaVu Sans"],
        "Wingdings" => &["Wingdings", "Symbol"],
        _ => &[],
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::Color;
    use crate::bundled_fonts::bundled_font_data;

    // Independent hb-subset/hb-shape 14.4.0 static instances, 1000 upem.
    // The S91 oracle receipt pins original bytes, commands and rounding limits.
    #[test]
    fn fx195_strict_directories_preserve_lexical_order_and_refuse_bad_inputs() {
        let path =
            std::env::temp_dir().join(format!("fx195-font-directory-{}", std::process::id()));
        std::fs::create_dir(&path).unwrap();
        assert!(FontFile::from_dir(&path).unwrap().is_empty());
        let bytes = include_bytes!("../fonts/Carlito-Regular.ttf");
        std::fs::write(path.join("z.ttf"), bytes).unwrap();
        std::fs::write(path.join("a.TTF"), bytes).unwrap();
        let fonts = FontFile::from_dir(&path).unwrap();
        assert_eq!(
            fonts
                .iter()
                .map(|font| font.family.as_str())
                .collect::<Vec<_>>(),
            ["a", "z"]
        );
        assert!(
            fonts
                .iter()
                .all(|font| font.data.as_slice() == bytes.as_slice())
        );
        assert!(FontFile::from_dir(&path.join("z.ttf")).is_err());
        std::fs::write(path.join("bad.otf"), b"invalid font").unwrap();
        assert!(FontFile::from_dir(&path).is_err());
        std::fs::remove_dir_all(&path).unwrap();
        assert!(FontFile::from_dir(&path).is_err());
    }

    #[test]
    fn fx195_variable_bold_matches_independent_harfbuzz_advances() {
        let mut mismatches = Vec::new();
        for (bytes, text, expected) in [
            (
                include_bytes!("../fonts/NotoSansThai.ttf").as_slice(),
                "ภาษาไทยยินดีต้อนรับ",
                8742.0,
            ),
            (
                include_bytes!("../fonts/NotoSansArabic.ttf").as_slice(),
                "العربية مرحبا بالعالم",
                8427.0,
            ),
            (
                include_bytes!("../fonts/NotoSansDevanagari.ttf").as_slice(),
                "देवनागरी नमस्ते दुनिया",
                8656.0,
            ),
        ] {
            let mut fonts =
                FontManager::new_with_fonts(vec![("FX195 caller".into(), bytes.to_vec())]);
            let bold = fonts
                .resolve_font(Some("FX195 caller"), true, false)
                .unwrap();
            let shaped = fonts.shape_text(bold, text, 1000.0).unwrap();
            if (shaped.width - expected).abs() > 1.0 {
                mismatches.push(format!(
                    "bold {text:?}: expected {expected} design units, got {}",
                    shaped.width
                ));
            }
        }
        assert!(mismatches.is_empty(), "{mismatches:?}");
    }

    #[test]
    fn fx195_regular_variable_preserves_independent_default_advances() {
        for (bytes, text, expected) in [
            (
                include_bytes!("../fonts/NotoSansThai.ttf").as_slice(),
                "ภาษาไทยยินดีต้อนรับ",
                8300.0,
            ),
            (
                include_bytes!("../fonts/NotoSansArabic.ttf").as_slice(),
                "العربية مرحبا بالعالم",
                7591.0,
            ),
            (
                include_bytes!("../fonts/NotoSansDevanagari.ttf").as_slice(),
                "देवनागरी नमस्ते दुनिया",
                7948.0,
            ),
        ] {
            let mut fonts =
                FontManager::new_with_fonts(vec![("FX195 caller".into(), bytes.to_vec())]);
            let regular = fonts
                .resolve_font(Some("FX195 caller"), false, false)
                .unwrap();
            let shaped = fonts.shape_text(regular, text, 1000.0).unwrap();
            assert!(
                (shaped.width - expected).abs() <= 1.0,
                "regular {text:?}: expected {expected} design units, got {}",
                shaped.width
            );
        }
    }

    #[test]
    fn fx195_bundled_variable_bold_is_not_exempt_from_weight_selection() {
        let mut fonts = FontManager::new_deterministic().unwrap();
        let bold = fonts
            .resolve_font(Some("Noto Sans Thai"), true, false)
            .unwrap();
        let shaped = fonts.shape_text(bold, "ภาษาไทยยินดีต้อนรับ", 1000.0).unwrap();
        assert!(
            (shaped.width - 8742.0).abs() <= 1.0,
            "bundled bold must use the independently measured 700 instance, got {}",
            shaped.width
        );
    }

    #[test]
    fn fx195_variable_ttc_keeps_face_index_and_shared_original_bytes() {
        let collection = test_ttc(&[
            include_bytes!("../fonts/NotoSansThai.ttf"),
            include_bytes!("../fonts/Carlito-Bold.ttf"),
        ]);
        let mut fonts =
            FontManager::new_with_fonts(vec![("FX195 collection".into(), collection.clone())]);
        let thai = fonts
            .resolve_font(Some("Noto Sans Thai"), true, false)
            .unwrap();
        let carlito = fonts.resolve_font(Some("Carlito"), true, false).unwrap();
        let thai_data = fonts.font_data(thai).unwrap();
        let carlito_data = fonts.font_data(carlito).unwrap();
        assert_eq!(thai_data.face_index, 0);
        assert_eq!(carlito_data.face_index, 1);
        assert_eq!(thai_data.data.as_ref(), collection);
        assert_eq!(carlito_data.data.as_ref(), collection);
        assert!(Arc::ptr_eq(&thai_data.data, &carlito_data.data));
        let static_face = ttf_parser::Face::parse(&carlito_data.data, 1).unwrap();
        assert!(!static_face.is_variable());
        assert!(static_face.is_bold());
        let shaped = fonts.shape_text(thai, "ภาษาไทยยินดีต้อนรับ", 1000.0).unwrap();
        assert!(
            (shaped.width - 8742.0).abs() <= 1.0,
            "TTC selected700 must match independent advance, got {}",
            shaped.width
        );
    }

    #[test]
    fn fx195_static_four_style_faces_keep_their_original_bytes() {
        let mut fonts = FontManager::new_deterministic().unwrap();
        for (bold, italic, expected) in [
            (
                false,
                false,
                include_bytes!("../fonts/Carlito-Regular.ttf").as_slice(),
            ),
            (
                true,
                false,
                include_bytes!("../fonts/Carlito-Bold.ttf").as_slice(),
            ),
            (
                false,
                true,
                include_bytes!("../fonts/Carlito-Italic.ttf").as_slice(),
            ),
            (
                true,
                true,
                include_bytes!("../fonts/Carlito-BoldItalic.ttf").as_slice(),
            ),
        ] {
            let id = fonts.resolve_font(Some("Carlito"), bold, italic).unwrap();
            assert_eq!(fonts.font_data(id).unwrap().data.as_ref(), expected);
        }
    }

    #[test]
    fn fx206_caps_requests_keep_occurrence_identity_expansion_and_default_projection() {
        let mut fonts = FontManager::new_deterministic().unwrap();
        let source = SourceSpan {
            node: crate::SourceNodeId::new(7).unwrap(),
            char_start: 40,
            char_end: 43,
        };
        let first = multilingual_test_segment(&mut fonts, "aBß", 24.0, Some(source));
        let second = first.clone();
        let plain = fonts
            .shape_multilingual_paragraph(
                vec![(first.clone(), Some("en".to_owned()))],
                TextDirection::Auto,
                false,
            )
            .unwrap();
        let (defaults, diagnostics) = fonts
            .shape_multilingual_paragraph_with_caps(
                vec![(first.clone(), Some("en".to_owned()))],
                TextDirection::Auto,
                false,
                &[(false, false)],
            )
            .unwrap();
        assert_eq!(defaults.len(), plain.len());
        for (actual, expected) in defaults.iter().zip(&plain) {
            assert_eq!(actual.text(), expected.text());
            assert_eq!(actual.glyph_ids(), expected.glyph_ids());
            assert_eq!(actual.x_advances(), expected.x_advances());
            assert_eq!(actual.clusters(), expected.clusters());
        }
        assert!(diagnostics.is_empty());
        let (caps, diagnostics) = fonts
            .shape_multilingual_paragraph_with_caps(
                vec![
                    (first, Some("en".to_owned())),
                    (second, Some("en".to_owned())),
                ],
                TextDirection::Auto,
                false,
                &[(true, false), (false, false)],
            )
            .unwrap();
        assert!(!diagnostics.is_empty());
        assert_eq!(
            caps.iter()
                .map(MultilingualTextSegment::text)
                .collect::<String>(),
            "aBßaBß"
        );
        let sharp = caps.iter().find(|span| span.text() == "ß").unwrap();
        assert_eq!(sharp.glyph_ids().len(), 2);
        assert_eq!(sharp.clusters().len(), 1);
        assert_eq!(sharp.base().source.unwrap().char_start, 42);
        assert_eq!(sharp.base().source.unwrap().char_end, 43);
        assert!((sharp.base().font_size - 19.2).abs() < 1e-9);
        assert!(
            caps.iter()
                .any(|span| span.text() == "aBß" && span.base().font_size == 24.0),
            "equal text in a distinct occurrence retains its own request"
        );
        assert!(
            fonts
                .shape_multilingual_paragraph_with_caps(
                    Vec::new(),
                    TextDirection::Auto,
                    false,
                    &[(true, false)]
                )
                .is_err()
        );
    }

    #[test]
    #[ignore = "required F206 physical host smcp oracle stage, run explicitly with pinned HB_SHAPE and FONT_ORACLE_PYTHON"]
    fn fx206_selected_smcp_matches_independent_harfbuzz_and_original_clusters() {
        let path = std::path::Path::new(
            "/Applications/Microsoft PowerPoint.app/Contents/Resources/DFonts/calibri.ttf",
        );
        let data = std::fs::read(path).expect("authenticated private host oracle face must exist");
        let mut fonts = FontManager::new_deterministic().unwrap();
        fonts.load_additional_fonts(&[FontFile {
            family: "F206 private smcp oracle".to_owned(),
            data,
        }]);
        let text = "Small caps Straße";
        let mut seed = multilingual_test_segment(&mut fonts, text, 24.0, None);
        seed.font_id = fonts
            .resolve_font_for_text(Some("F206 private smcp oracle"), false, false, text)
            .unwrap();
        let font = fonts.get_font(seed.font_id).unwrap();
        assert_eq!(font.family, "Calibri");
        let upem = font.units_per_em;
        let script = r#"import json,subprocess,sys
h,p,t,u=sys.argv[1:]
for feature in ['smcp=0','smcp=1']:
 d=json.loads(subprocess.check_output([h,p,t,'--no-glyph-names','--output-format=json','--script=latn','--language=en','--direction=ltr','--font-size='+u,'--features='+feature],text=True))
 print(','.join(str(x['g']) for x in d))
 print(','.join(str(x['ax']) for x in d))
"#;
        let output = std::process::Command::new(
            std::env::var("FONT_ORACLE_PYTHON").expect("explicit pinned interpreter"),
        )
        .args([
            "-c",
            script,
            &std::env::var("HB_SHAPE").expect("explicit pinned HarfBuzz"),
            path.to_str().unwrap(),
            text,
            &upem.to_string(),
        ])
        .output()
        .unwrap();
        assert!(
            output.status.success(),
            "independent oracle precondition: {}",
            String::from_utf8_lossy(&output.stderr)
        );
        let stdout = String::from_utf8(output.stdout).unwrap();
        let rows = stdout.lines().collect::<Vec<_>>();
        assert_eq!(rows.len(), 4);
        let gids = |row: &str| {
            row.split(',')
                .map(|item| item.parse::<u16>().unwrap())
                .collect::<Vec<_>>()
        };
        let advances = |row: &str| {
            row.split(',')
                .map(|item| item.parse::<f64>().unwrap() * 24.0 / f64::from(upem))
                .collect::<Vec<_>>()
        };
        assert_ne!(
            gids(rows[0]),
            gids(rows[2]),
            "real smcp application differs from features-off shape"
        );
        let (actual, diagnostics) = fonts
            .shape_multilingual_paragraph_with_caps(
                vec![(seed, Some("en".to_owned()))],
                TextDirection::LeftToRight,
                false,
                &[(true, false)],
            )
            .unwrap();
        assert!(diagnostics.is_empty());
        assert_eq!(
            actual
                .iter()
                .map(MultilingualTextSegment::text)
                .collect::<String>(),
            text
        );
        assert_eq!(
            actual
                .iter()
                .flat_map(|span| span.glyph_ids().iter().copied())
                .collect::<Vec<_>>(),
            gids(rows[2])
        );
        let actual_advances = actual
            .iter()
            .flat_map(|span| span.x_advances().iter().copied())
            .collect::<Vec<_>>();
        let expected = advances(rows[3]);
        assert_eq!(actual_advances.len(), expected.len());
        for (actual, expected) in actual_advances.iter().zip(expected) {
            assert!((actual - expected).abs() < 1e-9);
        }
    }

    fn multilingual_test_segment(
        manager: &mut FontManager,
        text: &str,
        size_pt: f64,
        source: Option<SourceSpan>,
    ) -> TextSegment {
        let font_id = manager
            .resolve_font_for_text(None, false, false, text)
            .expect("test font resolves");
        let metrics = manager
            .metrics(font_id, size_pt)
            .expect("test font metrics");
        let shaped = manager
            .shape_text(font_id, text, size_pt)
            .expect("test seed shapes");
        TextSegment {
            text: text.to_owned(),
            direction: TextDirection::Auto,
            source,
            font_id,
            font_size: size_pt,
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
            field_source: None,
            note: None,
            note_reference_source: None,
        }
    }

    /// Word's single line is the Windows extent of a face plus its external
    /// leading. Each row is a bundled face's OS/2 and `hhea` values, whose
    /// sum is the pitch Word gives the Microsoft face it stands in for:
    /// Calibri 2500/2048 em, Arial and Times New Roman 2355/2048, Cambria
    /// 1.172 and Courier New 2320/2048.
    #[test]
    fn word_line_metrics_are_the_windows_extent_and_its_external_leading() {
        let mut manager = FontManager::new_deterministic().expect("bundled fonts load");
        for (family, units_per_em, ascent, descent, leading) in [
            ("Carlito", 2048.0, 1950.0, 550.0, 0.0),
            ("Liberation Sans", 2048.0, 1854.0, 434.0, 67.0),
            ("Liberation Serif", 2048.0, 1825.0, 443.0, 87.0),
            ("Caladea", 1000.0, 950.0, 222.0, 0.0),
            ("Liberation Mono", 2048.0, 1705.0, 615.0, 0.0),
            // Sets `USE_TYPO_METRICS`, so its typographic 896 and 408 win
            // over its Windows 1348 and 558.
            ("Noto Sans Devanagari", 1000.0, 896.0, 408.0, 0.0),
        ] {
            let font_id = manager
                .resolve_font(Some(family), false, false)
                .expect("bundled face resolves");
            let metrics = manager
                .word_line_metrics(font_id, 12.0)
                .expect("loaded face has metrics");
            let scale = 12.0 / units_per_em;
            assert_eq!(metrics.units_per_em as f64, units_per_em, "{family}");
            assert!((metrics.ascent - ascent * scale).abs() < 1e-9, "{family}");
            assert!((metrics.descent - descent * scale).abs() < 1e-9, "{family}");
            assert!(
                (metrics.line_gap - leading * scale).abs() < 1e-9,
                "{family}"
            );
        }
    }

    fn word_metric_table_record(data: &[u8], tag: &[u8; 4]) -> usize {
        let count = u16::from_be_bytes(data[4..6].try_into().unwrap()) as usize;
        (0..count)
            .map(|n| 12 + n * 16)
            .find(|p| &data[*p..*p + 4] == tag)
            .unwrap()
    }

    fn synthetic_word_metric_font(codepages: u32, typo: bool) -> Vec<u8> {
        let mut data = bundled_font_data()
            .into_iter()
            .find(|(name, _)| *name == "Liberation Sans")
            .unwrap()
            .1
            .to_vec();
        let offset = |data: &[u8], tag: &[u8; 4]| {
            let r = word_metric_table_record(data, tag);
            u32::from_be_bytes(data[r + 8..r + 12].try_into().unwrap()) as usize
        };
        let os2 = offset(&data, b"OS/2");
        let hhea = offset(&data, b"hhea");
        data[os2..os2 + 2].copy_from_slice(&4_u16.to_be_bytes());
        data[os2 + 62..os2 + 64].copy_from_slice(&(if typo { 128_u16 } else { 0 }).to_be_bytes());
        for (at, value) in [(68, 1000_i16), (70, -200), (72, 300)] {
            data[os2 + at..os2 + at + 2].copy_from_slice(&value.to_be_bytes());
        }
        for (at, value) in [(74, 2300_u16), (76, 900)] {
            data[os2 + at..os2 + at + 2].copy_from_slice(&value.to_be_bytes());
        }
        data[os2 + 78..os2 + 82].copy_from_slice(&codepages.to_be_bytes());
        data[os2 + 82..os2 + 86].copy_from_slice(&0_u32.to_be_bytes());
        for (at, value) in [(4, 1300_i16), (6, -300), (8, 700)] {
            data[hhea + at..hhea + at + 2].copy_from_slice(&value.to_be_bytes());
        }
        // ttf-parser does not use SFNT checksums. Only these synthetic test
        // bytes change, never a bundled asset or an oracle font.
        assert!(ttf_parser::Face::parse(&data, 0).is_ok());
        data
    }

    #[test]
    fn cjk_word_metrics_use_face_codepages_and_symmetric_gap_free_hhea() {
        for bits in [
            0,
            1 << 16,
            1 << 17,
            1 << 18,
            1 << 19,
            1 << 20,
            1 << 21,
            0x001e0000,
        ] {
            for typo in [false, true] {
                let data = synthetic_word_metric_font(bits, typo);
                let mut manager = FontManager::new_deterministic().unwrap();
                let label = if bits == 0 {
                    "MS Gothic"
                } else {
                    "Latin label"
                };
                manager.load_additional_fonts(&[FontFile {
                    family: label.into(),
                    data: data.clone(),
                }]);
                let id = manager.resolve_font(Some(label), false, false).unwrap();
                assert_eq!(
                    manager.font_data(id).unwrap().data.as_ref(),
                    data.as_slice(),
                    "actual supplied face resolves"
                );
                for size in [10.0, 12.0, 16.0] {
                    let word = manager.word_line_metrics(id, size).unwrap();
                    let scale = size / f64::from(word.units_per_em);
                    let expected = if bits & 0x001e0000 != 0 {
                        (1540.0, 540.0, 0.0)
                    } else if typo {
                        (1000.0, 200.0, 300.0)
                    } else {
                        (2300.0, 900.0, 0.0)
                    };
                    assert!(
                        (word.ascent - expected.0 * scale).abs() < 1e-9,
                        "bits={bits:x} typo={typo} size={size} ascent={}",
                        word.ascent
                    );
                    assert!(
                        (word.descent - expected.1 * scale).abs() < 1e-9,
                        "bits={bits:x} descent={}",
                        word.descent
                    );
                    assert!((word.line_gap - expected.2 * scale).abs() < 1e-9);
                }
                let reference = synthetic_word_metric_font(0, typo);
                let face = ttf_parser::Face::parse(&reference, 0).unwrap();
                let shared = manager.metrics(id, 12.0).unwrap();
                let scale = 12.0 / f64::from(face.units_per_em());
                assert_eq!(shared.ascent, f64::from(face.ascender()) * scale);
                assert_eq!(shared.descent, -f64::from(face.descender()) * scale);
                assert_eq!(shared.line_gap, f64::from(face.line_gap()) * scale);
                let latin = manager.shape_text(id, "A", 12.0).unwrap();
                let _han = manager.shape_text(id, "漢", 12.0).unwrap();
                let again = manager.shape_text(id, "A", 12.0).unwrap();
                assert_eq!(
                    (again.glyph_ids, again.advances, again.width),
                    (latin.glyph_ids, latin.advances, latin.width),
                    "script does not change face policy or shaping"
                );
            }
        }
    }

    #[test]
    fn cjk_word_metric_table_boundaries_keep_the_existing_fallback() {
        for (version, length) in [(0, 78), (0, 86), (1, 78), (1, 80), (1, 82), (1, 84)] {
            let mut data = synthetic_word_metric_font(1 << 17, false);
            let record = word_metric_table_record(&data, b"OS/2");
            let offset =
                u32::from_be_bytes(data[record + 8..record + 12].try_into().unwrap()) as usize;
            data[offset..offset + 2].copy_from_slice(&(version as u16).to_be_bytes());
            data[record + 12..record + 16].copy_from_slice(&(length as u32).to_be_bytes());
            let face = ttf_parser::Face::parse(&data, 0).unwrap();
            if let Some(units) = WordLineUnits::of(&face) {
                assert_eq!(
                    (units.ascent, units.descent, units.leading),
                    (2300, 900, 0),
                    "version={version} length={length}"
                );
            } else {
                assert!(face.tables().os2.is_none());
            }
        }
        let mut second_word = synthetic_word_metric_font(0, false);
        let record = word_metric_table_record(&second_word, b"OS/2");
        let offset =
            u32::from_be_bytes(second_word[record + 8..record + 12].try_into().unwrap()) as usize;
        second_word[offset + 82..offset + 86].copy_from_slice(&(1_u32 << 17).to_be_bytes());
        let face = ttf_parser::Face::parse(&second_word, 0).unwrap();
        let units = WordLineUnits::of(&face).unwrap();
        assert_eq!(
            (units.ascent, units.descent),
            (2300, 900),
            "coverage word2 is not CP932"
        );
        let mut missing = synthetic_word_metric_font(1 << 17, false);
        let record = word_metric_table_record(&missing, b"OS/2");
        missing[record..record + 4].copy_from_slice(b"XXXX");
        let mut manager = FontManager::new_deterministic().unwrap();
        manager.load_additional_fonts(&[FontFile {
            family: "Missing OS2".into(),
            data: missing,
        }]);
        let id = manager
            .resolve_font(Some("Missing OS2"), false, false)
            .unwrap();
        let word = manager.word_line_metrics(id, 12.0).unwrap();
        let shared = manager.metrics(id, 12.0).unwrap();
        assert_eq!(
            (word.ascent, word.descent, word.line_gap, word.units_per_em),
            (
                shared.ascent,
                shared.descent,
                shared.line_gap,
                shared.units_per_em
            )
        );
        for (asc, desc) in [(0_i16, 0_i16), (-100, -200), (1300, 1)] {
            for typo in [false, true] {
                let mut data = synthetic_word_metric_font(1 << 17, typo);
                let record = word_metric_table_record(&data, b"hhea");
                let at =
                    u32::from_be_bytes(data[record + 8..record + 12].try_into().unwrap()) as usize;
                data[at + 4..at + 6].copy_from_slice(&asc.to_be_bytes());
                data[at + 6..at + 8].copy_from_slice(&desc.to_be_bytes());
                let face = ttf_parser::Face::parse(&data, 0).unwrap();
                let units = WordLineUnits::of(&face).unwrap();
                assert_eq!(
                    (units.ascent, units.descent),
                    if typo { (1000, 200) } else { (2300, 900) }
                );
            }
        }
    }

    #[test]
    fn cjk_word_metric_face_storage_and_bundled_coverage_are_explicit() {
        eprintln!(
            "WORD_FACE_STORAGE units={} optional={} loaded={}",
            std::mem::size_of::<WordLineUnits>(),
            std::mem::size_of::<Option<WordLineUnits>>(),
            std::mem::size_of::<LoadedFont>()
        );
        for (name, data) in bundled_font_data() {
            let face = ttf_parser::Face::parse(data, 0).unwrap();
            if let Some(table) = face.raw_face().table(ttf_parser::Tag::from_bytes(b"OS/2"))
                && table.len() >= 86
                && u16::from_be_bytes(table[..2].try_into().unwrap()) >= 1
            {
                let codepages = u32::from_be_bytes(table[78..82].try_into().unwrap());
                assert_eq!(
                    codepages & 0x001e0000,
                    0,
                    "bundled font unexpectedly activates CJK policy: {name}"
                );
            }
        }
    }

    #[test]
    fn caller_font_labels_override_bundled_families_and_keep_explicit_aliases() {
        let caladea = bundled_font_data()
            .into_iter()
            .find(|(family, _)| *family == "Caladea")
            .expect("Caladea is bundled")
            .1;
        let mut manager = FontManager::new_deterministic().expect("bundled fonts load");
        manager.load_additional_fonts(&[
            FontFile {
                family: "Document Serif".to_owned(),
                data: caladea.to_vec(),
            },
            FontFile {
                family: "Carlito".to_owned(),
                data: caladea.to_vec(),
            },
        ]);

        let aliased = manager
            .resolve_font(Some("Document Serif"), false, false)
            .expect("caller label resolves");
        assert_eq!(
            manager.fonts[manager.index_of(aliased).unwrap()].family,
            "Caladea"
        );

        let exact = manager
            .resolve_font(Some("Carlito"), false, false)
            .expect("caller label overrides the bundled family");
        assert_eq!(
            manager.fonts[manager.index_of(exact).unwrap()].family,
            "Caladea"
        );

        manager.set_caller_aliases(&[("Document Serif".to_owned(), "Carlito".to_owned())]);
        let explicit = manager
            .resolve_font(Some("Document Serif"), false, false)
            .expect("explicit alias precedes label-derived alias");
        assert_eq!(
            manager.fonts[manager.index_of(explicit).unwrap()].family,
            "Carlito"
        );

        manager.load_additional_fonts(&[FontFile {
            family: "Caladea".to_owned(),
            data: caladea.to_vec(),
        }]);
        assert!(manager.caller_aliases.is_empty());

        manager.set_caller_aliases(&[("Arial".to_owned(), "Caladea".to_owned())]);
        let caller_alias = manager
            .resolve_font(Some("Arial"), false, false)
            .expect("explicit alias resolves before mapped fallback");
        assert_eq!(
            manager.fonts[manager.index_of(caller_alias).unwrap()].family,
            "Caladea"
        );
        let generic = manager
            .resolve_font(Some("Unmapped Document Family"), false, false)
            .expect("generic fallback remains available");
        assert_eq!(
            manager.fonts[manager.index_of(generic).unwrap()].family,
            "Carlito"
        );
    }

    #[test]
    fn caller_alias_updates_preserve_bytes_and_invalidate_resolution_state() {
        let caladea = bundled_font_data()
            .into_iter()
            .find(|(family, _)| *family == "Caladea")
            .expect("Caladea is bundled")
            .1;
        let mut manager = FontManager::new_deterministic().expect("bundled fonts load");
        manager.load_additional_fonts(&[FontFile {
            family: "Caladea".to_owned(),
            data: caladea.to_vec(),
        }]);
        let aliases = vec![
            ("Document Serif A".to_owned(), "Caladea".to_owned()),
            ("Document Serif B".to_owned(), "Caladea".to_owned()),
        ];
        assert!(manager.set_caller_aliases(&aliases));

        let first = manager
            .resolve_font(Some("Document Serif A"), false, false)
            .expect("first alias resolves");
        let second = manager
            .resolve_font(Some("Document Serif B"), false, false)
            .expect("second alias resolves");
        let first_data = &manager.fonts[manager.index_of(first).unwrap()].data;
        let second_data = &manager.fonts[manager.index_of(second).unwrap()].data;
        assert!(Arc::ptr_eq(first_data, second_data));
        assert!(!manager.set_caller_aliases(&aliases));

        assert!(
            manager.set_caller_aliases(&[("Document Serif A".to_owned(), "Carlito".to_owned(),)])
        );
        let changed = manager
            .resolve_font(Some("Document Serif A"), false, false)
            .expect("changed alias resolves");
        assert_eq!(
            manager.fonts[manager.index_of(changed).unwrap()].family,
            "Carlito"
        );
        assert!(
            manager.index_of(first).is_some(),
            "loaded faces are retained"
        );
    }

    #[test]
    fn explicit_alias_state_respects_entry_and_retained_byte_ceilings() {
        let aliases = (0..CALLER_ALIAS_MAX_ENTRIES + 32)
            .map(|index| (format!("Document Serif {index}"), "Caladea".to_owned()))
            .collect::<Vec<_>>();
        let mut manager = FontManager::new_deterministic().expect("bundled fonts load");
        manager.set_caller_aliases(&aliases);
        assert_eq!(
            manager.explicit_aliases.as_slice(),
            &aliases[..CALLER_ALIAS_MAX_ENTRIES]
        );
        assert!(manager.explicit_aliases.len() <= CALLER_ALIAS_MAX_ENTRIES);
        assert!(manager.explicit_alias_map.len() <= CALLER_ALIAS_MAX_ENTRIES);

        let retained_large = ("x".repeat(32_760), String::new());
        let byte_limited = vec![
            retained_large.clone(),
            ("discarded bytes".to_owned(), "Caladea".to_owned()),
        ];
        manager.set_caller_aliases(&byte_limited);
        assert_eq!(manager.explicit_aliases, vec![retained_large]);

        let oversized = vec![("x".repeat(40_000), "Caladea".to_owned())];
        manager.set_caller_aliases(&oversized);
        let retained_bytes = manager
            .explicit_aliases
            .iter()
            .map(|(requested, target)| requested.len() + target.len())
            .sum::<usize>()
            + manager
                .explicit_alias_map
                .iter()
                .map(|(requested, target)| requested.len() + target.len())
                .sum::<usize>();
        assert!(retained_bytes <= CALLER_ALIAS_MAX_RETAINED_BYTES);
        assert!(manager.explicit_aliases.is_empty());
        assert!(manager.explicit_alias_map.is_empty());
    }

    #[test]
    fn label_alias_prefers_caller_bytes_over_bundled_same_family() {
        let bundled = bundled_font_data()
            .into_iter()
            .find(|(family, _)| *family == "Caladea")
            .expect("Caladea is bundled")
            .1;
        let mut caller = bundled.to_vec();
        caller.push(0);
        let mut manager = FontManager::new_deterministic().expect("bundled fonts load");
        manager.load_additional_fonts(&[FontFile {
            family: "Document Serif".to_owned(),
            data: caller.clone(),
        }]);

        let resolved = manager
            .resolve_font(Some("Document Serif"), false, false)
            .expect("caller label resolves");
        let loaded = &manager.fonts[manager.index_of(resolved).unwrap()];
        assert_eq!(loaded.family, "Caladea");
        assert_eq!(loaded.data.as_ref(), caller.as_slice());
        assert_ne!(loaded.data.as_ref(), bundled);
    }

    #[test]
    fn case_only_caller_labels_resolve_to_the_supplied_face() {
        let bundled = bundled_font_data()
            .into_iter()
            .find(|(family, _)| *family == "Caladea")
            .expect("Caladea is bundled")
            .1;
        let mut caller = bundled.to_vec();
        caller.push(0);
        let mut manager = FontManager::new_deterministic().expect("bundled fonts load");
        manager.load_additional_fonts(&[FontFile {
            family: "caladea".to_owned(),
            data: caller.clone(),
        }]);

        let resolved = manager
            .resolve_font(Some("caladea"), false, false)
            .expect("case-only caller label resolves");
        let loaded = &manager.fonts[manager.index_of(resolved).unwrap()];
        assert_eq!(loaded.family, "Caladea");
        assert_eq!(loaded.data.as_ref(), caller.as_slice());
    }

    #[test]
    fn constructor_label_alias_survives_additional_font_replacement() {
        let caladea = bundled_font_data()
            .into_iter()
            .find(|(family, _)| *family == "Caladea")
            .expect("Caladea is bundled")
            .1;
        let carlito = bundled_font_data()
            .into_iter()
            .find(|(family, _)| *family == "Carlito")
            .expect("Carlito is bundled")
            .1;
        let mut constructor = caladea.to_vec();
        constructor.push(0);
        let mut manager =
            FontManager::new_with_fonts(vec![("Document Serif".to_owned(), constructor.clone())]);

        manager.load_additional_fonts(&[FontFile {
            family: "Additional Sans".to_owned(),
            data: carlito.to_vec(),
        }]);

        let resolved = manager
            .resolve_font(Some("Document Serif"), false, false)
            .expect("constructor label still resolves");
        let loaded = &manager.fonts[manager.index_of(resolved).unwrap()];
        assert_eq!(loaded.family, "Caladea");
        assert_eq!(loaded.data.as_ref(), constructor.as_slice());
    }

    #[test]
    fn constructor_family_priority_survives_additional_font_replacement() {
        let caladea = bundled_font_data()
            .into_iter()
            .find(|(family, _)| *family == "Caladea")
            .expect("Caladea is bundled")
            .1;
        let mut constructor = caladea.to_vec();
        constructor.push(0);
        let mut replacement = caladea.to_vec();
        replacement.extend_from_slice(&[0, 0]);
        let mut manager =
            FontManager::new_with_fonts(vec![("Document Serif".to_owned(), constructor.clone())]);

        manager.load_additional_fonts(&[FontFile {
            family: "Caladea".to_owned(),
            data: replacement,
        }]);

        let resolved = manager
            .resolve_font(Some("Caladea"), false, false)
            .expect("constructor family still resolves");
        let loaded = &manager.fonts[manager.index_of(resolved).unwrap()];
        assert_eq!(loaded.data.as_ref(), constructor.as_slice());
    }

    #[test]
    fn caller_face_selection_matches_fontdb_css_rules() {
        let caladea = bundled_font_data()
            .into_iter()
            .find(|(family, _)| *family == "Caladea")
            .expect("Caladea is bundled")
            .1;
        let mut source = fontdb::Database::new();
        source.load_font_data(caladea.to_vec());
        let template = source.faces().next().expect("Caladea has a face").clone();
        let mut db = fontdb::Database::new();
        let mut add_face = |weight: u16, stretch: fontdb::Stretch, style: fontdb::Style| {
            let mut face = template.clone();
            face.weight = fontdb::Weight(weight);
            face.stretch = stretch;
            face.style = style;
            db.push_face_info(face)
        };

        let weight_300 = add_face(300, fontdb::Stretch::Normal, fontdb::Style::Normal);
        let weight_500 = add_face(500, fontdb::Stretch::Normal, fontdb::Style::Normal);
        let weight_600 = add_face(600, fontdb::Stretch::Normal, fontdb::Style::Normal);
        let weight_800 = add_face(800, fontdb::Stretch::Normal, fontdb::Style::Normal);
        let expanded = add_face(400, fontdb::Stretch::SemiExpanded, fontdb::Style::Normal);
        let condensed = add_face(400, fontdb::Stretch::SemiCondensed, fontdb::Style::Normal);
        let normal = add_face(400, fontdb::Stretch::Normal, fontdb::Style::Normal);
        let italic = add_face(400, fontdb::Stretch::Normal, fontdb::Style::Italic);

        assert_eq!(
            best_caller_face(
                &db,
                &[weight_300, weight_500],
                fontdb::Weight::NORMAL,
                fontdb::Style::Normal,
            ),
            Some(weight_500)
        );
        assert_eq!(
            best_caller_face(
                &db,
                &[weight_600, weight_800],
                fontdb::Weight::BOLD,
                fontdb::Style::Normal,
            ),
            Some(weight_800)
        );
        assert_eq!(
            best_caller_face(
                &db,
                &[expanded, condensed],
                fontdb::Weight::NORMAL,
                fontdb::Style::Normal,
            ),
            Some(condensed)
        );
        assert_eq!(
            best_caller_face(
                &db,
                &[normal, italic],
                fontdb::Weight::NORMAL,
                fontdb::Style::Oblique,
            ),
            Some(italic)
        );
    }

    #[test]
    fn fx220_embedded_physical_body_serves_two_thousand_aliases_atomically() {
        let body: Arc<[u8]> = Arc::from(include_bytes!("../fonts/Caladea-Regular.ttf").as_slice());
        let names = (0..2000)
            .map(|index| format!("FX220 Alias {index}"))
            .collect::<Vec<_>>();
        let sources = vec![(names, body.clone())];
        let mut manager = FontManager::new_deterministic().unwrap();
        let base = manager.db.len();
        assert!(manager.load_embedded_fonts(&sources).unwrap());
        assert_eq!(manager.db.len(), base + 1);
        let mut ids = Vec::new();
        let mut data = Vec::new();
        for index in [0, 999, 1999] {
            let id = manager
                .resolve_font(Some(&format!("FX220 Alias {index}")), false, false)
                .unwrap();
            data.push(manager.font_data(id).unwrap().data);
            ids.push(id);
        }
        assert!(Arc::ptr_eq(&data[0], &data[1]) && Arc::ptr_eq(&data[1], &data[2]));
        assert_eq!(data[0].as_ref(), body.as_ref());
        let cache = manager.cache.len();
        assert!(!manager.load_embedded_fonts(&sources).unwrap());
        assert_eq!(manager.cache.len(), cache);
        assert_eq!(
            manager
                .resolve_font(Some("FX220 Alias 0"), false, false)
                .unwrap(),
            ids[0]
        );
        let refused = vec![(vec!["too many".to_owned(); 16_385], body.clone())];
        assert!(manager.load_embedded_fonts(&refused).is_err());
        assert!(
            manager
                .load_embedded_fonts(&[(
                    vec!["malformed".to_owned()],
                    Arc::from([0_u8; 16].as_slice())
                )])
                .is_err()
        );
        let mut hostile_collection = b"ttcf\x00\x01\x00\x00".to_vec();
        hostile_collection.extend_from_slice(&u32::MAX.to_be_bytes());
        assert!(
            manager
                .load_embedded_fonts(&[(
                    vec!["hostile collection".to_owned()],
                    Arc::from(hostile_collection)
                )])
                .is_err()
        );
        assert_eq!(manager.cache.len(), cache);
        assert_eq!(manager.db.len(), base + 1);
        assert_eq!(
            manager
                .resolve_font(Some("FX220 Alias 0"), false, false)
                .unwrap(),
            ids[0]
        );
        let caller = include_bytes!("../fonts/Carlito-Regular.ttf").to_vec();
        manager.load_additional_fonts(&[FontFile {
            family: "FX220 Alias 1999".to_owned(),
            data: caller.clone(),
        }]);
        let id = manager
            .resolve_font(Some("FX220 Alias 1999"), false, false)
            .unwrap();
        assert_eq!(manager.font_data(id).unwrap().data.as_ref(), caller);
        let id = manager
            .resolve_font(Some("FX220 Alias 999"), false, false)
            .unwrap();
        assert_eq!(manager.font_data(id).unwrap().data.as_ref(), body.as_ref());
        assert!(manager.load_embedded_fonts(&[]).unwrap());
        assert!(manager.embedded_aliases.is_empty());
    }

    #[test]
    fn fx220_embedded_four_faces_preserve_css_and_caller_priority() {
        let face_data: [&[u8]; 4] = [
            include_bytes!("../fonts/Caladea-Regular.ttf"),
            include_bytes!("../fonts/Caladea-Bold.ttf"),
            include_bytes!("../fonts/Caladea-Italic.ttf"),
            include_bytes!("../fonts/Caladea-BoldItalic.ttf"),
        ];
        let sources = face_data
            .iter()
            .map(|data| (vec!["Embedded FX220".to_owned()], Arc::<[u8]>::from(*data)))
            .collect::<Vec<_>>();
        let mut manager = FontManager::new_deterministic().unwrap();
        manager.load_embedded_fonts(&sources).unwrap();
        for (index, (bold, italic)) in [(false, false), (true, false), (false, true), (true, true)]
            .into_iter()
            .enumerate()
        {
            let id = manager
                .resolve_font(Some("Embedded FX220"), bold, italic)
                .unwrap();
            assert_eq!(
                manager.font_data(id).unwrap().data.as_ref(),
                face_data[index]
            );
        }
        let caller = include_bytes!("../fonts/Carlito-Regular.ttf").to_vec();
        let mut constructor =
            FontManager::new_with_fonts(vec![("Embedded FX220".to_owned(), caller.clone())]);
        constructor.load_embedded_fonts(&sources).unwrap();
        let id = constructor
            .resolve_font(Some("Embedded FX220"), false, false)
            .unwrap();
        assert_eq!(constructor.font_data(id).unwrap().data.as_ref(), caller);
        manager.load_additional_fonts(&[FontFile {
            family: "Embedded FX220".to_owned(),
            data: caller.clone(),
        }]);
        let id = manager
            .resolve_font(Some("Embedded FX220"), false, false)
            .unwrap();
        assert_eq!(manager.font_data(id).unwrap().data.as_ref(), caller);
    }

    #[test]
    fn fx220_cross_bucket_ranking_retains_constructor_and_css_precedence() {
        let family = "FX220CX";
        let exact = font_with_family(include_bytes!("../fonts/Caladea-Regular.ttf"), family);
        let alias = include_bytes!("../fonts/Carlito-Regular.ttf").to_vec();
        let bold = include_bytes!("../fonts/Carlito-Bold.ttf").to_vec();
        let italic = include_bytes!("../fonts/Carlito-Italic.ttf").to_vec();
        let bold_italic = include_bytes!("../fonts/Carlito-BoldItalic.ttf").to_vec();
        for (constructor, additional) in [
            (exact.clone(), alias.clone()),
            (alias.clone(), exact.clone()),
        ] {
            let mut manager =
                FontManager::new_with_fonts(vec![(family.to_owned(), constructor.clone())]);
            let files = vec![FontFile {
                family: family.to_owned(),
                data: additional,
            }];
            assert!(manager.load_additional_fonts(&files));
            let id = manager.resolve_font(Some(family), false, false).unwrap();
            assert_eq!(manager.font_data(id).unwrap().data.as_ref(), constructor);
            assert!(!manager.load_additional_fonts(&files));
            assert_eq!(
                manager.resolve_font(Some(family), false, false).unwrap(),
                id
            );
            assert!(manager.load_additional_fonts(&[
                FontFile {
                    family: family.to_owned(),
                    data: bold.clone()
                },
                FontFile {
                    family: family.to_owned(),
                    data: italic.clone()
                },
                FontFile {
                    family: family.to_owned(),
                    data: bold_italic.clone()
                },
            ]));
            for (is_bold, is_italic, expected) in [
                (true, false, &bold),
                (false, true, &italic),
                (true, true, &bold_italic),
            ] {
                let id = manager
                    .resolve_font(Some(family), is_bold, is_italic)
                    .unwrap();
                assert_eq!(manager.font_data(id).unwrap().data.as_ref(), expected);
            }
        }
        let mut manager = FontManager::new_deterministic().unwrap();
        for files in [
            vec![exact.clone(), alias.clone()],
            vec![alias.clone(), exact.clone()],
        ] {
            let expected = files.last().unwrap().clone();
            manager.load_additional_fonts(
                &files
                    .into_iter()
                    .map(|data| FontFile {
                        family: family.to_owned(),
                        data,
                    })
                    .collect::<Vec<_>>(),
            );
            let id = manager.resolve_font(Some(family), false, false).unwrap();
            assert_eq!(manager.font_data(id).unwrap().data.as_ref(), expected);
        }
        manager.set_caller_aliases(&[("Explicit FX220".to_owned(), family.to_owned())]);
        let direct = manager.resolve_font(Some(family), false, false).unwrap();
        let indirect = manager
            .resolve_font(Some("Explicit FX220"), false, false)
            .unwrap();
        assert_eq!(
            manager.font_data(direct).unwrap().data,
            manager.font_data(indirect).unwrap().data
        );
        assert!(
            manager
                .resolve_font(Some("Missing FX220"), false, false)
                .is_ok()
        );
    }

    #[test]
    fn fx195_equal_rank_caller_faces_select_last_loaded_original_bytes() {
        let family = "FX195CX";
        let first = font_with_family(include_bytes!("../fonts/Carlito-Regular.ttf"), family);
        let second = font_with_family(include_bytes!("../fonts/Caladea-Regular.ttf"), family);
        assert_ne!(first, second);
        let mut constructor = FontManager::new_with_fonts(vec![
            ("Constructor Alias".to_owned(), first.clone()),
            ("Constructor Alias".to_owned(), second.clone()),
        ]);
        for requested in [family, "Constructor Alias"] {
            let selected = constructor
                .resolve_font(Some(requested), false, false)
                .unwrap();
            assert_eq!(
                constructor.font_data(selected).unwrap().data.as_ref(),
                first
            );
        }
        assert!(constructor.load_additional_fonts(&[FontFile {
            family: family.to_owned(),
            data: second.clone(),
        }]));
        for requested in [family, "Constructor Alias"] {
            let selected = constructor
                .resolve_font(Some(requested), false, false)
                .unwrap();
            assert_eq!(
                constructor.font_data(selected).unwrap().data.as_ref(),
                first
            );
        }
        let mut manager = FontManager::new_deterministic().unwrap();
        let mut files = vec![
            FontFile {
                family: family.to_owned(),
                data: first.clone(),
            },
            FontFile {
                family: family.to_owned(),
                data: second.clone(),
            },
        ];
        assert!(manager.load_additional_fonts(&files));
        let selected = manager.resolve_font(Some(family), false, false).unwrap();
        assert_eq!(manager.font_data(selected).unwrap().data.as_ref(), second);
        assert_eq!(manager.font_data(selected).unwrap().face_index, 0);
        assert!(!manager.load_additional_fonts(&files));
        assert_eq!(
            manager.resolve_font(Some(family), false, false).unwrap(),
            selected
        );

        files.reverse();
        assert!(manager.load_additional_fonts(&files));
        let selected = manager.resolve_font(Some(family), false, false).unwrap();
        assert_eq!(manager.font_data(selected).unwrap().data.as_ref(), first);
        assert_eq!(manager.font_data(selected).unwrap().face_index, 0);

        let bold = font_with_family(include_bytes!("../fonts/Carlito-Bold.ttf"), family);
        assert!(constructor.load_additional_fonts(&[FontFile {
            family: family.to_owned(),
            data: bold.clone(),
        }]));
        let selected = constructor.resolve_font(Some(family), true, false).unwrap();
        assert_eq!(constructor.font_data(selected).unwrap().data.as_ref(), bold);
        assert!(!constructor.font_data(selected).unwrap().synthetic_bold);
        assert!(manager.load_additional_fonts(&[
            FontFile {
                family: family.to_owned(),
                data: bold.clone(),
            },
            FontFile {
                family: family.to_owned(),
                data: second.clone(),
            },
        ]));
        let selected = manager.resolve_font(Some(family), true, false).unwrap();
        assert_eq!(manager.font_data(selected).unwrap().data.as_ref(), bold);
        assert!(!manager.font_data(selected).unwrap().synthetic_bold);
        let regular = manager.resolve_font(Some(family), false, false).unwrap();
        assert_eq!(manager.font_data(regular).unwrap().data.as_ref(), second);
    }

    fn font_with_family(source: &[u8], family: &str) -> Vec<u8> {
        assert!(family.is_ascii() && !family.is_empty());
        let mut font = source.to_vec();
        let table_count = u16::from_be_bytes([font[4], font[5]]) as usize;
        let record = (0..table_count)
            .map(|table| 12 + table * 16)
            .find(|record| &font[*record..*record + 4] == b"name")
            .expect("font has name table");
        let offset = u32::from_be_bytes(font[record + 8..record + 12].try_into().unwrap()) as usize;
        let length =
            u32::from_be_bytes(font[record + 12..record + 16].try_into().unwrap()) as usize;
        let mut names = font[offset..offset + length].to_vec();
        let count = u16::from_be_bytes([names[2], names[3]]) as usize;
        let strings = u16::from_be_bytes([names[4], names[5]]) as usize;
        for index in 0..count {
            let entry = 6 + index * 12;
            let platform = u16::from_be_bytes([names[entry], names[entry + 1]]);
            let name_id = u16::from_be_bytes([names[entry + 6], names[entry + 7]]);
            if !matches!(name_id, 1 | 16) {
                continue;
            }
            let encoded = match platform {
                0 | 3 => family
                    .encode_utf16()
                    .flat_map(u16::to_be_bytes)
                    .collect::<Vec<_>>(),
                1 => family.as_bytes().to_vec(),
                _ => continue,
            };
            let string_offset = u16::try_from(names.len() - strings).unwrap();
            let length = u16::try_from(encoded.len()).unwrap();
            names[entry + 8..entry + 10].copy_from_slice(&length.to_be_bytes());
            names[entry + 10..entry + 12].copy_from_slice(&string_offset.to_be_bytes());
            names.extend_from_slice(&encoded);
        }
        let checksum = |data: &[u8]| {
            data.chunks(4).fold(0u32, |sum, chunk| {
                let mut word = [0; 4];
                word[..chunk.len()].copy_from_slice(chunk);
                sum.wrapping_add(u32::from_be_bytes(word))
            })
        };
        while !font.len().is_multiple_of(4) {
            font.push(0);
        }
        let new_offset = u32::try_from(font.len()).unwrap();
        font[record + 4..record + 8].copy_from_slice(&checksum(&names).to_be_bytes());
        font[record + 8..record + 12].copy_from_slice(&new_offset.to_be_bytes());
        font[record + 12..record + 16]
            .copy_from_slice(&u32::try_from(names.len()).unwrap().to_be_bytes());
        font.extend_from_slice(&names);
        let head = (0..table_count)
            .map(|table| 12 + table * 16)
            .find(|record| &font[*record..*record + 4] == b"head")
            .unwrap();
        let head_offset =
            u32::from_be_bytes(font[head + 8..head + 12].try_into().unwrap()) as usize;
        let head_length =
            u32::from_be_bytes(font[head + 12..head + 16].try_into().unwrap()) as usize;
        font[head_offset + 8..head_offset + 12].fill(0);
        let head_checksum = checksum(&font[head_offset..head_offset + head_length]);
        font[head + 4..head + 8].copy_from_slice(&head_checksum.to_be_bytes());
        let adjustment = 0xb1b0afba_u32.wrapping_sub(checksum(&font));
        font[head_offset + 8..head_offset + 12].copy_from_slice(&adjustment.to_be_bytes());
        assert_eq!(checksum(&font), 0xb1b0afba);
        font
    }

    fn test_ttc(fonts: &[&[u8]]) -> Vec<u8> {
        let header_len = 12 + fonts.len() * 4;
        let mut collection = vec![0u8; header_len];
        collection[0..4].copy_from_slice(b"ttcf");
        collection[4..8].copy_from_slice(&0x0001_0000u32.to_be_bytes());
        collection[8..12].copy_from_slice(&(fonts.len() as u32).to_be_bytes());

        for (font_number, font) in fonts.iter().enumerate() {
            while !collection.len().is_multiple_of(4) {
                collection.push(0);
            }
            let collection_offset = collection.len();
            collection[12 + font_number * 4..16 + font_number * 4]
                .copy_from_slice(&(collection_offset as u32).to_be_bytes());

            let mut adjusted = font.to_vec();
            let table_count = u16::from_be_bytes([adjusted[4], adjusted[5]]) as usize;
            for table in 0..table_count {
                let offset_position = 12 + table * 16 + 8;
                let offset = u32::from_be_bytes(
                    adjusted[offset_position..offset_position + 4]
                        .try_into()
                        .expect("table offset"),
                );
                adjusted[offset_position..offset_position + 4]
                    .copy_from_slice(&(offset + collection_offset as u32).to_be_bytes());
            }
            collection.extend_from_slice(&adjusted);
        }
        collection
    }

    #[test]
    fn deterministic_font_manager_uses_only_bundled_fonts() {
        let mut fm = FontManager::new_deterministic().expect("bundled fonts should load");

        assert_eq!(fm.db.faces().count(), bundled_font_data().len());
        assert!(fm.resolve_font(Some("Arial"), false, false).is_ok());
    }

    #[cfg(feature = "system-fonts")]
    #[test]
    fn normal_font_discovery_initializes_once_per_process() {
        let _first = FontManager::new();
        let _second = FontManager::new();
        assert_eq!(SYSTEM_FONT_DISCOVERY_RUNS.load(Ordering::Relaxed), 1);

        let _deterministic =
            FontManager::new_deterministic().expect("bundled font manager should load");
        let _caller = FontManager::new_with_fonts(Vec::new());
        assert_eq!(SYSTEM_FONT_DISCOVERY_RUNS.load(Ordering::Relaxed), 1);
    }

    #[cfg(feature = "system-fonts")]
    #[test]
    fn file_backed_collection_faces_share_one_byte_buffer() {
        let suffix = format!("{}-{:?}", std::process::id(), std::thread::current().id());
        let first_path = std::env::temp_dir().join(format!("rdocx-font-cache-{suffix}-a.ttf"));
        let second_path = std::env::temp_dir().join(format!("rdocx-font-cache-{suffix}-b.ttf"));
        let collection = test_ttc(&[bundled_font_data()[0].1, bundled_font_data()[4].1]);
        std::fs::write(&first_path, &collection).expect("write first temporary collection");
        std::fs::write(&second_path, &collection).expect("write second temporary collection");

        let mut db = fontdb::Database::new();
        db.load_font_file(&first_path).expect("load first TTC");
        db.load_font_file(&second_path).expect("load second TTC");
        let canonical_first = std::fs::canonicalize(&first_path).unwrap();
        let canonical_second = std::fs::canonicalize(&second_path).unwrap();
        let first_ids = db
            .faces()
            .filter_map(|face| match &face.source {
                fontdb::Source::File(path) if path == &first_path || path == &canonical_first => {
                    Some(face.id)
                }
                _ => None,
            })
            .collect::<Vec<_>>();
        let second_id = db
            .faces()
            .find_map(|face| match &face.source {
                fontdb::Source::File(path) if path == &second_path || path == &canonical_second => {
                    Some(face.id)
                }
                _ => None,
            })
            .expect("second TTC face");
        assert_eq!(first_ids.len(), 2);

        let mut memory = HashMap::new();
        let (first_face, first_index) =
            font_data_for_face(&db, first_ids[0], &mut memory).expect("first TTC face bytes");
        let (second_face, second_index) =
            font_data_for_face(&db, first_ids[1], &mut memory).expect("second TTC face bytes");
        let (other_file, _) =
            font_data_for_face(&db, second_id, &mut memory).expect("other TTC bytes");
        assert_ne!(first_index, second_index);
        assert!(Arc::ptr_eq(&first_face, &second_face));
        assert!(!Arc::ptr_eq(&first_face, &other_file));

        std::fs::remove_file(first_path).expect("remove first temporary font");
        std::fs::remove_file(second_path).expect("remove second temporary font");
    }

    #[test]
    fn shaping_memo_uses_complete_text_size_and_font_identity() {
        let mut fm = FontManager::new_deterministic().expect("bundled fonts should load");
        let regular = fm.resolve_font(Some("Carlito"), false, false).unwrap();
        let bold = fm.resolve_font(Some("Carlito"), true, false).unwrap();

        let first = fm.shape_text(regular, "exact text", 11.0).unwrap();
        let repeat = fm.shape_text(regular, "exact text", 11.0).unwrap();
        assert_eq!(first.glyph_ids, repeat.glyph_ids);
        assert_eq!(fm.shaping_memo_counts().0, 1);

        fm.shape_text(regular, "different text", 11.0).unwrap();
        fm.shape_text(regular, "exact text", 12.0).unwrap();
        fm.shape_text(bold, "exact text", 11.0).unwrap();
        assert_eq!(fm.shaping_memo_counts().1, 4);

        let replacement = FontFile {
            family: "Carlito".to_owned(),
            data: bundled_font_data()[1].1.to_vec(),
        };
        fm.load_additional_fonts(&[replacement]);
        assert_eq!(fm.shaping_memo_counts(), (0, 0, 0, 0));
        let replacement_id = fm.resolve_font(Some("Carlito"), false, false).unwrap();
        fm.shape_text(replacement_id, "exact text", 11.0).unwrap();
        assert_eq!(fm.shaping_memo_counts().1, 1);
        let mut explicit = FontManager::new_deterministic().unwrap();
        let face = explicit
            .resolve_font(Some("Carlito"), false, false)
            .unwrap();
        let default = explicit.shape_text(face, "AVATAR To", 20.0).unwrap();
        let off = explicit
            .shape_text_with_kerning(face, "AVATAR To", 20.0, false)
            .unwrap();
        let on = explicit
            .shape_text_with_kerning(face, "AVATAR To", 20.0, true)
            .unwrap();
        assert_eq!(default.glyph_ids, on.glyph_ids);
        assert_eq!(default.advances, on.advances);
        assert!(off.width > on.width);
        for _ in 0..2 {
            assert_eq!(
                explicit
                    .shape_text(face, "AVATAR To", 20.0)
                    .unwrap()
                    .advances,
                default.advances
            );
            assert_eq!(
                explicit
                    .shape_text_with_kerning(face, "AVATAR To", 20.0, false)
                    .unwrap()
                    .advances,
                off.advances
            );
            assert_eq!(
                explicit
                    .shape_text_with_kerning(face, "AVATAR To", 20.0, true)
                    .unwrap()
                    .advances,
                on.advances
            );
        }
        assert_eq!(explicit.shaping_memo_counts().2, 3);
        let base = multilingual_test_segment(&mut explicit, "AVATAR", 20.0, None);
        assert!(
            explicit
                .shape_paragraph_with_kerning(
                    vec![(base.clone(), None)],
                    TextDirection::LeftToRight,
                    false,
                    &[],
                    true
                )
                .is_err()
        );
        let mut language = String::with_capacity(2048);
        language.push_str("en-US");
        let mut span = explicit
            .shape_paragraph_with_kerning(
                vec![(base, Some(language))],
                TextDirection::LeftToRight,
                false,
                &[false],
                true,
            )
            .unwrap()
            .remove(0);
        assert_eq!(span.kerning_override(), Some(false));
        assert!(span.uses_legacy_shaping());
        assert_eq!(span.x_advances.capacity(), span.x_advances.len());
        assert_eq!(span.y_advances.capacity(), span.y_advances.len());
        assert_eq!(span.x_offsets.capacity(), span.x_offsets.len());
        assert_eq!(span.y_offsets.capacity(), span.y_offsets.len());
        assert_eq!(span.clusters.capacity(), span.clusters.len());
        assert_eq!(span.language.as_ref().unwrap().capacity(), 5);
        let mut language = String::with_capacity(4096);
        language.push_str("en-US");
        span = span.with_hyphenation_language(Some(language));
        assert_eq!(span.hyphenation_language.as_ref().unwrap().capacity(), 5);
        let mut advances = Vec::with_capacity(2048);
        advances.extend_from_slice(span.x_advances());
        let replacement = span.base().clone();
        let span = span.with_base_and_advances(replacement, advances).unwrap();
        assert_eq!(span.x_advances.capacity(), span.x_advances.len());
        assert_eq!(span.hyphenation_language(), Some("en-US"));
        let mut incompatible = span.base().clone();
        incompatible.text.push('x');
        assert!(
            span.clone()
                .with_base_and_advances(incompatible, span.x_advances().to_vec())
                .is_err()
        );
        let mut invalid = span.x_advances().to_vec();
        invalid[0] = f64::NAN;
        assert!(
            span.clone()
                .with_base_and_advances(span.base().clone(), invalid)
                .is_err()
        );
        for legacy in [false, true] {
            let base = multilingual_test_segment(&mut explicit, "ab cd", 20.0, None);
            let nowrap = explicit
                .shape_paragraph_with_kerning(
                    vec![(base.clone(), None)],
                    TextDirection::LeftToRight,
                    true,
                    &[false],
                    legacy,
                )
                .unwrap();
            assert!(
                nowrap.iter().all(|span| !span.break_after()),
                "no-wrap provider exposes a soft break in legacy={legacy}"
            );
            let wrapped = explicit
                .shape_paragraph_with_kerning(
                    vec![(base, None)],
                    TextDirection::LeftToRight,
                    false,
                    &[false],
                    legacy,
                )
                .unwrap();
            let mut bytes = 0;
            assert!(wrapped.iter().any(|span| {
                bytes += span.text().len();
                bytes == 3 && span.break_after()
            }));
        }
        // Empty styled occurrences retain context without inventing a break.
        for legacy in [false, true] {
            for no_wrap in [false, true] {
                for texts in [&[][..], &[""][..], &["ab", "", "cd"][..]] {
                    let segments = texts
                        .iter()
                        .map(|text| {
                            (
                                multilingual_test_segment(&mut explicit, text, 20.0, None),
                                None,
                            )
                        })
                        .collect();
                    let spans = explicit
                        .shape_paragraph_with_kerning(
                            segments,
                            TextDirection::LeftToRight,
                            no_wrap,
                            &vec![false; texts.len()],
                            legacy,
                        )
                        .unwrap();
                    assert_eq!(
                        spans.iter().map(|span| span.text()).collect::<String>(),
                        texts.concat()
                    );
                    assert_eq!(
                        spans.iter().filter(|span| span.text().is_empty()).count(),
                        texts.iter().filter(|text| text.is_empty()).count()
                    );
                    assert!(
                        spans
                            .iter()
                            .filter(|span| span.text().is_empty())
                            .all(|span| !span.break_after()),
                        "empty occurrence invents a break in legacy={legacy}, no_wrap={no_wrap}"
                    );
                    if no_wrap {
                        assert!(spans.iter().all(|span| !span.break_after()));
                    }
                }
            }
        }
        let generic = MultilingualTextSegment::new(
            span.base().clone(),
            span.logical_index(),
            Some({
                let mut value = String::with_capacity(2048);
                value.push_str("en");
                value
            }),
            span.script(),
            span.direction(),
            span.bidi_level(),
            span.x_advances().to_vec(),
            span.y_advances().to_vec(),
            span.x_offsets().to_vec(),
            span.y_offsets().to_vec(),
            span.clusters().to_vec(),
            span.break_after(),
        )
        .unwrap();
        assert!(!generic.uses_legacy_shaping());
        assert_eq!(generic.kerning_override(), None);
        assert_eq!(generic.language.as_ref().unwrap().capacity(), 2048);
        // Legacy shaping uses the old Unicode chunks, never an extra script buffer.
        let node = crate::SourceNodeId::new(77).unwrap();
        for (text, chunks) in [
            ("ΑVТoA\u{301}ΩЖfi", vec!["ΑVТoA\u{301}ΩЖfi"]),
            ("ΑVТoA\u{301}—ΩЖfi", vec!["ΑVТoA\u{301}", "—", "ΩЖfi"]),
        ] {
            for no_wrap in [false, true] {
                for enabled in [false, true] {
                    let mut base = multilingual_test_segment(
                        &mut explicit,
                        text,
                        11.0,
                        Some(SourceSpan {
                            node,
                            char_start: 10,
                            char_end: 10 + text.chars().count() as u32,
                        }),
                    );
                    base.font_id = face;
                    let shaped = explicit
                        .shape_text_with_kerning(face, text, 11.0, enabled)
                        .unwrap();
                    base.glyph_ids = shaped.glyph_ids;
                    base.advances = shaped.advances;
                    base.width = shaped.width;
                    let spans = explicit
                        .shape_paragraph_with_kerning(
                            vec![(base, None)],
                            TextDirection::Auto,
                            no_wrap,
                            &[enabled],
                            true,
                        )
                        .unwrap();
                    assert_eq!(
                        spans.iter().map(|span| span.text()).collect::<Vec<_>>(),
                        chunks
                    );
                    let mut offset = 10;
                    for span in spans {
                        let reference = explicit
                            .shape_text_with_kerning(face, span.text(), 11.0, enabled)
                            .unwrap();
                        assert_eq!(span.glyph_ids(), reference.glyph_ids);
                        assert_eq!(span.x_advances(), reference.advances);
                        assert_eq!(span.width().to_bits(), reference.width.to_bits());
                        assert_eq!(span.break_after(), !no_wrap);
                        assert_eq!(span.kerning_override(), Some(enabled));
                        assert!(span.uses_legacy_shaping());
                        let end = offset + span.text().chars().count() as u32;
                        assert_eq!(
                            span.base().source,
                            Some(SourceSpan {
                                node,
                                char_start: offset,
                                char_end: end
                            })
                        );
                        offset = end;
                        assert_eq!(span.x_advances.capacity(), span.x_advances.len());
                        assert_eq!(span.clusters.capacity(), span.clusters.len());
                    }
                }
            }
        }
        let equal = multilingual_test_segment(&mut explicit, "AVATAR", 20.0, None);
        let empty = multilingual_test_segment(&mut explicit, "", 20.0, None);
        let occurrences = explicit
            .shape_paragraph_with_kerning(
                vec![(equal.clone(), None), (empty, None), (equal, None)],
                TextDirection::Auto,
                true,
                &[false, true, true],
                true,
            )
            .unwrap();
        assert_eq!(
            occurrences
                .iter()
                .map(|span| span.kerning_override())
                .collect::<Vec<_>>(),
            vec![Some(false), Some(true), Some(true)]
        );
        assert_eq!(
            occurrences
                .iter()
                .map(|span| span.logical_index())
                .collect::<Vec<_>>(),
            vec![0, 1, 2]
        );
        assert!(occurrences.iter().all(|span| !span.break_after()));
    }

    #[test]
    fn arabic_joining_survives_script_and_line_break_boundaries() {
        let mut fm = FontManager::new_deterministic().expect("bundled fonts should load");
        let segment = multilingual_test_segment(&mut fm, "العربية", 18.0, None);
        let shaped = fm
            .shape_multilingual_text(segment, Some("ar"), TextDirection::RightToLeft, false)
            .unwrap();
        assert_eq!(shaped.len(), 1);
        assert_eq!(
            shaped[0].glyph_ids(),
            &[288, 85, 319, 18, 317, 19, 31, 48, 72, 8]
        );
        assert_eq!(
            shaped[0]
                .clusters()
                .iter()
                .map(|cluster| (cluster.glyph_start..cluster.glyph_end, cluster.char_range()))
                .collect::<Vec<_>>(),
            vec![
                (0..2, 6..7),
                (2..4, 5..6),
                (4..6, 4..5),
                (6..7, 3..4),
                (7..8, 2..3),
                (8..9, 1..2),
                (9..10, 0..1),
            ]
        );
    }

    #[test]
    fn indic_clusters_are_never_split_or_mapped_as_independent_scalars() {
        let mut fm = FontManager::new_deterministic().expect("bundled fonts should load");
        let segment = multilingual_test_segment(&mut fm, "कि", 18.0, None);
        let shaped = fm
            .shape_multilingual_text(segment, Some("hi"), TextDirection::LeftToRight, false)
            .unwrap();

        assert_eq!(shaped[0].clusters()[0].char_range(), 0..2);
        assert_eq!(shaped[0].x_offsets().len(), shaped[0].glyph_ids().len());
        assert_eq!(shaped[0].y_offsets().len(), shaped[0].glyph_ids().len());
    }

    #[test]
    fn same_script_coverage_changes_split_only_between_graphemes() {
        let mut fm = FontManager::new_deterministic().expect("bundled fonts should load");
        let font_id = fm.resolve_font(Some("Carlito"), false, false).unwrap();
        let metrics = fm.metrics(font_id, 18.0).unwrap();
        let seed = fm.shape_text(font_id, "A☀∙", 18.0).unwrap();
        let shaped = fm
            .shape_multilingual_text(
                TextSegment {
                    text: "A☀∙".to_owned(),
                    direction: TextDirection::Auto,
                    source: None,
                    font_id,
                    font_size: 18.0,
                    glyph_ids: seed.glyph_ids,
                    advances: seed.advances,
                    width: seed.width,
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
                    field_source: None,
                    note: None,
                    note_reference_source: None,
                },
                None,
                TextDirection::LeftToRight,
                false,
            )
            .unwrap();

        assert_eq!(
            shaped.iter().map(|span| span.text()).collect::<Vec<_>>(),
            ["A", "☀", "∙"]
        );
        assert!(shaped.iter().all(|span| span.script() == TextScript::Latin));
        assert_ne!(shaped[1].font_id(), shaped[2].font_id());
        assert!(shaped.iter().all(|span| {
            let index = fm.index_of(span.font_id()).unwrap();
            fm.uncovered(index, span.text()).is_empty()
        }));
    }

    #[test]
    fn multilingual_constructor_rejects_invalid_levels_and_cluster_maps() {
        let mut fm = FontManager::new_deterministic().expect("bundled fonts should load");
        let segment = multilingual_test_segment(&mut fm, "ab", 18.0, None);
        let valid = fm
            .shape_multilingual_text(segment, None, TextDirection::LeftToRight, true)
            .unwrap()
            .remove(0);
        let rebuild = |bidi_level, clusters| {
            MultilingualTextSegment::new(
                valid.base().clone(),
                valid.logical_index(),
                valid.language().map(str::to_owned),
                valid.script(),
                valid.direction(),
                bidi_level,
                valid.x_advances().to_vec(),
                valid.y_advances().to_vec(),
                valid.x_offsets().to_vec(),
                valid.y_offsets().to_vec(),
                clusters,
                valid.break_after(),
            )
        };

        assert!(rebuild(255, valid.clusters().to_vec()).is_err());
        let mut out_of_bounds = valid.clusters().to_vec();
        out_of_bounds.last_mut().unwrap().char_end = 3;
        assert!(rebuild(valid.bidi_level(), out_of_bounds).is_err());
        let mut glyph_gap = valid.clusters().to_vec();
        glyph_gap[0].glyph_start = 1;
        assert!(rebuild(valid.bidi_level(), glyph_gap).is_err());

        let base = multilingual_test_segment(&mut fm, "AVATAR", 18.0, None);
        let mut owner = fm
            .shape_paragraph_with_kerning(
                vec![(base, Some("en-US".into()))],
                TextDirection::LeftToRight,
                false,
                &[false],
                false,
            )
            .unwrap()
            .remove(0)
            .with_hyphenation_language(Some("en-US".into()));
        owner.x_advances.reserve_exact(31);
        owner.y_advances.reserve_exact(31);
        owner.x_offsets.reserve_exact(31);
        owner.y_offsets.reserve_exact(31);
        owner.clusters.reserve_exact(31);
        owner.language.as_mut().unwrap().reserve_exact(31);
        owner
            .hyphenation_language
            .as_mut()
            .unwrap()
            .reserve_exact(31);
        owner.base.glyph_ids.reserve_exact(31);
        let base_capacity = owner.base.glyph_ids.capacity();
        let original = owner.clone();
        let rebased = owner.with_logical_index_and_legacy_shaping(123, true);
        assert_eq!(rebased.logical_index(), 123);
        assert!(rebased.uses_legacy_shaping());
        assert_eq!(
            format!("{:?}", rebased.base()),
            format!("{:?}", original.base())
        );
        assert_eq!(rebased.base.glyph_ids.capacity(), base_capacity);
        assert_eq!(rebased.kerning_override(), Some(false));
        assert_eq!(rebased.script(), original.script());
        assert_eq!(rebased.direction(), original.direction());
        assert_eq!(rebased.bidi_level(), original.bidi_level());
        assert_eq!(rebased.break_after(), original.break_after());
        for (actual, expected) in [
            (&rebased.x_advances, &original.x_advances),
            (&rebased.y_advances, &original.y_advances),
            (&rebased.x_offsets, &original.x_offsets),
            (&rebased.y_offsets, &original.y_offsets),
        ] {
            assert_eq!(actual, expected);
            assert_eq!(actual.capacity(), actual.len());
        }
        assert_eq!(rebased.clusters, original.clusters);
        assert_eq!(rebased.clusters.capacity(), rebased.clusters.len());
        for (actual, expected) in [
            (&rebased.language, &original.language),
            (
                &rebased.hyphenation_language,
                &original.hyphenation_language,
            ),
        ] {
            assert_eq!(actual, expected);
            assert_eq!(
                actual.as_ref().unwrap().capacity(),
                actual.as_ref().unwrap().len()
            );
        }
        let rich = rebased.with_logical_index_and_legacy_shaping(7, false);
        assert_eq!(rich.logical_index(), 7);
        assert!(!rich.uses_legacy_shaping());
        assert_eq!(rich.kerning_override(), Some(false));
        assert_eq!(rich.glyph_ids(), original.glyph_ids());

        // Layout bidi and guessed glyph order are independent only for a
        // positively selected legacy constructor. Public new stays strict.
        let opposite_level = valid.bidi_level() ^ 1;
        assert!(rebuild(opposite_level, valid.clusters().to_vec()).is_err());
        let legacy_rebuild = |level, clusters| {
            MultilingualTextSegment::new_with_legacy_shaping(
                valid.base().clone(),
                valid.logical_index(),
                valid.language().map(str::to_owned),
                valid.script(),
                valid.direction(),
                level,
                valid.x_advances().to_vec(),
                valid.y_advances().to_vec(),
                valid.x_offsets().to_vec(),
                valid.y_offsets().to_vec(),
                clusters,
                valid.break_after(),
                true,
            )
        };
        let legacy = legacy_rebuild(opposite_level, valid.clusters().to_vec()).unwrap();
        assert_eq!(legacy.bidi_level(), opposite_level);
        assert!(legacy.uses_legacy_shaping());
        assert_eq!(legacy.glyph_ids(), valid.glyph_ids());
        assert_eq!(legacy.clusters(), valid.clusters());
        let unchanged = legacy
            .clone()
            .with_base_and_advances(legacy.base().clone(), legacy.x_advances().to_vec())
            .unwrap();
        assert_eq!(unchanged.bidi_level(), opposite_level);
        assert_eq!(unchanged.clusters(), valid.clusters());
        assert!(legacy_rebuild(255, valid.clusters().to_vec()).is_err());
        let mut gap = valid.clusters().to_vec();
        gap[0].glyph_start = 1;
        assert!(legacy_rebuild(opposite_level, gap).is_err());
        let mut outside = valid.clusters().to_vec();
        outside.last_mut().unwrap().char_end = 3;
        assert!(legacy_rebuild(opposite_level, outside).is_err());
    }

    #[test]
    fn thai_words_offer_approved_breaks_without_losing_source_text() {
        let text = "ภาษาไทยยินดีต้อนรับ";
        let mut fm = FontManager::new_deterministic().expect("bundled fonts should load");
        let source = SourceSpan {
            node: crate::SourceNodeId::new(3).unwrap(),
            char_start: 10,
            char_end: 29,
        };
        let segment = multilingual_test_segment(&mut fm, text, 18.0, Some(source));
        let shaped = fm
            .shape_multilingual_text(segment, Some("th"), TextDirection::LeftToRight, false)
            .unwrap();
        assert_eq!(
            shaped
                .iter()
                .map(|span| {
                    let source = span.base().source.unwrap();
                    (
                        span.text(),
                        span.break_after(),
                        source.char_start..source.char_end,
                    )
                })
                .collect::<Vec<_>>(),
            vec![
                ("ภาษา", true, 10..14),
                ("ไทยยิน", true, 14..20),
                ("ดี", true, 20..22),
                ("ต้อน", true, 22..26),
                ("รับ", true, 26..29),
            ]
        );
        assert_eq!(
            shaped.iter().map(|span| span.text()).collect::<String>(),
            text
        );
    }

    #[test]
    fn latin_spans_end_at_line_break_opportunities_with_their_spaces_apart() {
        let mut fm = FontManager::new_deterministic().expect("bundled fonts should load");
        let segment = multilingual_test_segment(&mut fm, "in service, or re-coated", 18.0, None);
        let shaped = fm
            .shape_multilingual_text(segment, None, TextDirection::LeftToRight, false)
            .unwrap();

        // No opportunity before the comma, the space or the hyphen, and each
        // space before an opportunity is a span of its own.
        assert_eq!(
            shaped
                .iter()
                .map(|span| (span.text(), span.break_after()))
                .collect::<Vec<_>>(),
            vec![
                ("in", false),
                (" ", true),
                ("service,", false),
                (" ", true),
                ("or", false),
                (" ", true),
                ("re-", true),
                ("coated", true),
            ]
        );
    }

    #[test]
    fn a_word_split_across_styled_segments_has_no_break_where_they_meet() {
        let mut fm = FontManager::new_deterministic().expect("bundled fonts should load");
        let first = multilingual_test_segment(&mut fm, "Well-kno", 18.0, None);
        let second = multilingual_test_segment(&mut fm, "wn state", 18.0, None);
        let shaped = fm
            .shape_multilingual_paragraph(
                vec![(first, None), (second, None)],
                TextDirection::LeftToRight,
                false,
            )
            .unwrap();

        assert_eq!(
            shaped
                .iter()
                .map(|span| (span.text(), span.break_after()))
                .collect::<Vec<_>>(),
            vec![
                ("Well-", true),
                ("kno", false),
                ("wn", false),
                (" ", true),
                ("state", true),
            ]
        );
    }

    #[test]
    fn shaping_memo_hits_preserve_fifo_order() {
        let mut fm = FontManager::new_deterministic().expect("bundled fonts should load");
        let font = fm.resolve_font(Some("Carlito"), false, false).unwrap();
        fm.shape_text(font, "oldest exact shape", 11.0).unwrap();
        fm.shape_text(font, "newest exact shape", 11.0).unwrap();
        fm.shape_text(font, "oldest exact shape", 11.0).unwrap();

        let memo = fm.shaping_memo.lock().unwrap();
        assert_eq!(memo.hits, 1);
        assert_eq!(memo.entries.front().unwrap().0.text, "oldest exact shape");
        assert_eq!(memo.entries.back().unwrap().0.text, "newest exact shape");
    }

    #[test]
    fn shaping_memo_fingerprint_collision_requires_exact_key_equality() {
        let mut fm = FontManager::new_deterministic().expect("bundled fonts should load");
        let font = fm.resolve_font(Some("Carlito"), false, false).unwrap();
        fm.shape_text(font, "first collision candidate", 11.0)
            .unwrap();
        let forced =
            shaping_fingerprint(font, "second collision candidate", 11.0f64.to_bits(), None);
        fm.shaping_memo.lock().unwrap().entries[0].3 = forced;

        fm.shape_text(font, "second collision candidate", 11.0)
            .unwrap();

        let (hits, misses, entries, _) = fm.shaping_memo_counts();
        assert_eq!((hits, misses, entries), (0, 2, 2));
    }

    #[test]
    fn shaping_memo_is_bounded_and_recovers_from_poison() {
        let mut fm = FontManager::new_deterministic().expect("bundled fonts should load");
        let font = fm.resolve_font(Some("Carlito"), false, false).unwrap();
        for index in 0..(SHAPING_CACHE_MAX_ENTRIES + 20) {
            fm.shape_text(font, &format!("bounded shaping entry {index}"), 11.0)
                .unwrap();
        }
        let (_, _, entries, bytes) = fm.shaping_memo_counts();
        assert!(entries <= SHAPING_CACHE_MAX_ENTRIES);
        assert!(bytes <= SHAPING_CACHE_MAX_BYTES);

        let fm = Arc::new(fm);
        let poison = Arc::clone(&fm);
        assert!(
            std::thread::spawn(move || {
                let _guard = poison.shaping_memo.lock().unwrap();
                panic!("poison shaping cache for recovery coverage");
            })
            .join()
            .is_err()
        );
        let first = fm.shape_text(font, "after poison", 11.0).unwrap();
        let second = fm.shape_text(font, "after poison", 11.0).unwrap();
        assert_eq!(first.glyph_ids, second.glyph_ids);
        let (hits, misses, entries, bytes) = fm.shaping_memo_counts();
        assert_eq!((hits, misses, entries), (1, 1, 1));
        assert!(bytes > 0);
    }

    #[test]
    fn shaping_memo_enforces_its_byte_ceiling_in_production_insertion() {
        let mut memo = ShapingMemo::new();
        for suffix in ['a', 'b'] {
            memo.insert(
                ShapingKey {
                    font_id: FontId(0),
                    text: std::iter::repeat_n(suffix, 9 * 1024 * 1024).collect(),
                    size_bits: 11.0f64.to_bits(),
                    kerning: None,
                },
                ShapedText {
                    glyph_ids: Vec::new(),
                    advances: Vec::new(),
                    width: 0.0,
                },
            );
        }
        assert_eq!(memo.entries.len(), 1);
        assert!(memo.bytes <= SHAPING_CACHE_MAX_BYTES);
    }

    #[test]
    fn persistent_coverage_and_loaded_face_state_is_bounded_and_deduplicated() {
        let mut fm = FontManager::new_deterministic().expect("bundled fonts load");
        for _ in 0..(COVERAGE_FALLBACK_MAX_ENTRIES + 20) {
            fm.remember_coverage_fallback(false, false, 0);
        }
        assert_eq!(fm.coverage_fallbacks[&(false, false)], vec![0]);

        let misses = (0..(COVERAGE_MISS_MAX_ENTRIES + 20))
            .filter_map(|value| char::from_u32(0x10_000 + value as u32))
            .collect::<Vec<_>>();
        fm.remember_coverage_misses(&misses);
        assert_eq!(fm.coverage_misses.len(), COVERAGE_MISS_MAX_ENTRIES);

        for index in 0..(RESOLUTION_CACHE_MAX_ENTRIES + 20) {
            fm.resolve_font(Some(&format!("missing alias {index}")), false, false)
                .expect("bounded fallback resolves");
        }
        assert!(fm.cache.len() <= RESOLUTION_CACHE_MAX_ENTRIES);
        assert_eq!(fm.fonts.len(), RESOLUTION_CACHE_MAX_ENTRIES);
    }

    #[test]
    fn active_document_may_resolve_more_than_256_distinct_faces() {
        let source = bundled_font_data()[4].1;
        let mut db = fontdb::Database::new();
        for index in 0..257 {
            db.load_font_data(font_with_family(source, &format!("F{index:06}")));
        }
        let mut fm = FontManager::from_base_database(db);
        fm.begin_layout();
        let mut ids = HashSet::new();
        for index in 0..257 {
            let family = format!("F{index:06}");
            let id = fm
                .resolve_font(Some(&family), false, false)
                .expect("distinct active face resolves");
            assert_eq!(fm.font_data(id).unwrap().family, family);
            ids.insert(id);
        }
        assert_eq!(ids.len(), 257);
        fm.retain_current_fonts();
        assert_eq!(fm.fonts.len(), 257);
    }

    #[test]
    fn font_trace_is_bounded_to_one_candidate_and_releases_capacity() {
        let mut fm = FontManager::new_deterministic().expect("bundled fonts load");
        fm.begin_layout();
        for _ in 0..(PARAGRAPH_FONT_TRACE_MAX_ENTRIES + 20) {
            fm.resolve_font(Some("Carlito"), false, false).unwrap();
        }
        assert!(fm.paragraph_font_trace.is_none());

        fm.begin_paragraph_font_trace();
        for _ in 0..(PARAGRAPH_FONT_TRACE_MAX_ENTRIES + 20) {
            fm.resolve_font(Some("Carlito"), false, false).unwrap();
        }
        assert!(fm.finish_paragraph_font_trace().is_none());

        fm.begin_layout();
        assert_eq!(fm.layout_fonts.capacity(), 0);
        fm.begin_paragraph_font_trace();
        fm.resolve_font(Some("Carlito"), false, false).unwrap();
        let trace = fm.finish_paragraph_font_trace().expect("bounded trace");
        assert_eq!(trace.len(), 1);
        assert_eq!(trace.capacity(), trace.len());
    }

    #[test]
    fn fx195_contribution_scope_preserves_legacy_and_original_requests() {
        let mut fm = FontManager::new_deterministic().unwrap();
        fm.begin_layout();
        fm.begin_paragraph_font_trace();
        let id = fm.resolve_font(Some("Carlito"), false, false).unwrap();
        let scope = fm.begin_font_contribution_scope().unwrap();
        fm.record_font_contribution(id, Some("Aptos"), "A").unwrap();
        fm.record_font_contribution(id, Some("Calibri"), "A")
            .unwrap();
        let staged = fm.finish_font_contribution_scope(scope).unwrap().unwrap();
        assert_eq!(staged.len(), 2);
        let projected = staged
            .iter()
            .filter_map(|event| match event {
                FontTraceEvent::Contribution(contribution) => Some(contribution),
                FontTraceEvent::Resolution(_) => None,
            })
            .collect::<Vec<_>>();
        assert_eq!(projected[0].font_id, projected[1].font_id);
        assert_ne!(projected[0].requested_family, projected[1].requested_family);
        assert!(fm.finish_paragraph_font_contributions().unwrap().is_empty());
        fm.begin_paragraph_font_trace();
        fm.replay_layout_font_trace(&[id]);
        fm.replay_font_trace_events(&staged).unwrap();
        assert_eq!(fm.finish_paragraph_font_trace().unwrap(), vec![id]);
        fm.begin_paragraph_font_trace();
        fm.replay_font_trace_events(&staged).unwrap();
        assert_eq!(
            fm.finish_paragraph_font_contributions().unwrap(),
            projected.into_iter().cloned().collect::<Vec<_>>()
        );
    }

    #[test]
    fn fx195_contribution_nested_stale_and_empty_scopes_are_atomic() {
        let mut fm = FontManager::new_deterministic().unwrap();
        let id = fm.resolve_font(Some("Carlito"), false, false).unwrap();
        fm.begin_layout();
        fm.begin_paragraph_font_trace();
        fm.record_font_contribution(id, Some("outer"), "A").unwrap();
        let outer = fm.begin_font_contribution_scope().unwrap();
        let inner = fm.begin_font_contribution_scope().unwrap();
        fm.record_font_contribution(id, Some("inner"), "A").unwrap();
        let before = format!("{:?}", fm.paragraph_font_trace.as_ref().unwrap().events);
        assert!(fm.finish_font_contribution_scope(outer).is_err());
        assert_eq!(
            format!("{:?}", fm.paragraph_font_trace.as_ref().unwrap().events),
            before
        );
        let captured = fm.finish_font_contribution_scope(inner).unwrap().unwrap();
        assert_eq!(captured.len(), 1);
        let stale = fm.begin_font_contribution_scope().unwrap();
        fm.begin_paragraph_font_trace();
        fm.record_font_contribution(id, Some("replacement"), "A")
            .unwrap();
        let inventory = fm.current_layout_fonts().to_vec();
        assert!(fm.finish_font_contribution_scope(stale).is_err());
        assert_eq!(fm.current_layout_fonts(), inventory);
        let remaining = fm.finish_paragraph_font_contributions().unwrap();
        assert_eq!(remaining.len(), 1);
        assert_eq!(
            remaining[0].requested_family.as_deref(),
            Some("replacement")
        );
        fm.begin_paragraph_font_trace();
        fm.record_font_contribution(id, Some("empty"), "").unwrap();
        assert!(fm.finish_paragraph_font_contributions().unwrap().is_empty());
        fm.begin_paragraph_font_trace();
        fm.resolve_font(Some("Carlito"), false, false).unwrap();
        assert!(fm.finish_paragraph_font_contributions().unwrap().is_empty());
    }

    #[test]
    fn fx195_contribution_replay_refusal_and_capture_overflow_are_atomic() {
        let mut fm = FontManager::new_deterministic().unwrap();
        let id = fm.resolve_font(Some("Carlito"), false, false).unwrap();
        fm.begin_layout();
        fm.begin_paragraph_font_trace();
        fm.record_font_contribution(id, Some("kept"), "A").unwrap();
        let before = format!("{:?}", fm.paragraph_font_trace.as_ref().unwrap().events);
        let inventory = fm.current_layout_fonts().to_vec();
        let bad = vec![
            FontTraceEvent::Contribution(FontContribution {
                font_id: id,
                requested_family: Some("valid first".into()),
            }),
            FontTraceEvent::Contribution(FontContribution {
                font_id: FontId(u32::MAX),
                requested_family: None,
            }),
        ];
        assert!(fm.replay_font_trace_events(&bad).is_err());
        assert_eq!(fm.current_layout_fonts(), inventory);
        assert_eq!(
            format!("{:?}", fm.paragraph_font_trace.as_ref().unwrap().events),
            before
        );
        let huge = "x".repeat(PARAGRAPH_FONT_TRACE_MAX_BYTES + 1);
        assert!(
            fm.replay_font_trace_events(&[FontTraceEvent::Contribution(FontContribution {
                font_id: id,
                requested_family: Some(huge.clone()),
            })])
            .is_err()
        );
        assert_eq!(fm.current_layout_fonts(), inventory);
        assert_eq!(
            format!("{:?}", fm.paragraph_font_trace.as_ref().unwrap().events),
            before
        );
        let scope = fm.begin_font_contribution_scope().unwrap();
        fm.record_font_contribution(id, Some(&huge), "A").unwrap();
        assert!(fm.finish_font_contribution_scope(scope).unwrap().is_none());
        assert!(fm.finish_paragraph_font_contributions().is_none());
        assert!(fm.paragraph_font_trace.is_none());
        fm.begin_paragraph_font_trace();
        for _ in 0..=PARAGRAPH_FONT_TRACE_MAX_ENTRIES {
            fm.record_font_contribution(id, Some("bounded"), "A")
                .unwrap();
        }
        assert!(fm.finish_paragraph_font_contributions().is_none());
        assert!(fm.paragraph_font_trace.is_none());
    }

    #[test]
    fn fx195_contribution_scope_rejects_other_manager_and_bounds_private_context() {
        let mut first = FontManager::new_deterministic().unwrap();
        let mut second = FontManager::new_deterministic().unwrap();
        let id = second.resolve_font(Some("Carlito"), false, false).unwrap();
        let foreign_scope = first.begin_font_contribution_scope().unwrap();
        let own_scope = second.begin_font_contribution_scope().unwrap();
        second
            .record_font_contribution(id, Some("kept"), "A")
            .unwrap();
        let before = format!("{:?}", second.paragraph_font_trace.as_ref().unwrap().events);
        let inventory = second.current_layout_fonts().to_vec();
        assert!(
            second
                .finish_font_contribution_scope(foreign_scope)
                .is_err()
        );
        assert_eq!(
            format!("{:?}", second.paragraph_font_trace.as_ref().unwrap().events),
            before
        );
        assert_eq!(second.current_layout_fonts(), inventory);
        assert!(
            second
                .finish_font_contribution_scope(own_scope)
                .unwrap()
                .is_some()
        );
        second.begin_paragraph_font_trace();
        let mut bytes = PARAGRAPH_FONT_TRACE_MAX_BYTES;
        let mut entries = 0;
        assert!(
            second
                .copy_font_request_context(Some("no extra allocation"), &mut bytes, &mut entries)
                .is_none()
        );
        assert_eq!(entries, 0);
        assert!(second.finish_paragraph_font_events().is_none());
        second.begin_paragraph_font_trace();
        let mut bytes = 0;
        let mut entries = PARAGRAPH_FONT_TRACE_MAX_ENTRIES;
        assert!(
            second
                .copy_font_request_context(Some("entry overflow"), &mut bytes, &mut entries)
                .is_none()
        );
        assert_eq!(bytes, PARAGRAPH_FONT_TRACE_MAX_BYTES);
        assert!(second.finish_paragraph_font_events().is_none());
    }

    #[cfg(feature = "system-fonts")]
    #[test]
    fn file_byte_cache_is_bounded_and_recovers_from_poison() {
        let cache = Arc::new(Mutex::new(FileFontCache::new()));
        {
            let mut cache = cache
                .lock()
                .unwrap_or_else(std::sync::PoisonError::into_inner);
            cache.clear();
            let oversized: Arc<[u8]> = Arc::from(vec![0; FILE_FONT_CACHE_MAX_BYTES + 1]);
            let returned = cache_file_font_bytes(
                &mut cache,
                PathBuf::from("oversized-font.ttc"),
                Arc::clone(&oversized),
            )
            .expect("oversized bytes are returned uncached");
            assert!(Arc::ptr_eq(&oversized, &returned));
            assert!(cache.entries.is_empty());
            assert_eq!(cache.bytes, 0);
        }

        let poison = Arc::clone(&cache);
        assert!(
            std::thread::spawn(move || {
                let cache = poison;
                let _guard = cache.lock().unwrap();
                panic!("poison file byte cache for recovery coverage");
            })
            .join()
            .is_err()
        );

        let path = std::env::temp_dir().join(format!(
            "rdocx-font-cache-poison-{}-{:?}.ttf",
            std::process::id(),
            std::thread::current().id()
        ));
        std::fs::write(&path, bundled_font_data()[0].1).expect("write recovery font");
        let first = shared_file_font_bytes_from_cache(&cache, &path).expect("recover cache");
        let second = shared_file_font_bytes_from_cache(&cache, &path).expect("reuse cache");
        assert!(Arc::ptr_eq(&first, &second));
        std::fs::remove_file(path).expect("remove recovery font");
    }

    #[cfg(not(feature = "system-fonts"))]
    #[test]
    fn no_default_features_omits_system_font_discovery() {
        let fm = FontManager::new();
        assert_eq!(fm.db.faces().count(), bundled_font_data().len());
    }

    #[test]
    fn font_manager_with_no_fonts_returns_an_error() {
        let mut fm = FontManager::new_with_fonts(Vec::new());
        assert!(matches!(
            fm.resolve_font(None, false, false),
            Err(LayoutError::FontNotFound(_))
        ));
    }

    #[test]
    fn load_system_font() {
        let mut fm = FontManager::new();
        // Should be able to resolve at least one font via fallback
        let result = fm.resolve_font(None, false, false);
        // On CI or systems without fonts this might fail, so we just check it doesn't panic
        if let Ok(id) = result {
            assert_eq!(id.0, 0);
        }
    }

    #[test]
    fn font_metrics_positive() {
        let mut fm = FontManager::new();
        if let Ok(id) = fm.resolve_font(None, false, false) {
            let metrics = fm.metrics(id, 12.0).unwrap();
            assert!(metrics.ascent > 0.0);
            assert!(metrics.descent > 0.0);
            assert!(metrics.units_per_em > 0);
        }
    }

    #[test]
    fn shape_hello_world() {
        let mut fm = FontManager::new();
        if let Ok(id) = fm.resolve_font(None, false, false) {
            let shaped = fm.shape_text(id, "Hello World", 12.0).unwrap();
            assert!(!shaped.glyph_ids.is_empty());
            assert_eq!(shaped.glyph_ids.len(), shaped.advances.len());
            assert!(shaped.width > 0.0);
        }
    }

    #[test]
    fn font_caching() {
        let mut fm = FontManager::new();
        if let Ok(id1) = fm.resolve_font(Some("Arial"), false, false) {
            let id2 = fm.resolve_font(Some("Arial"), false, false).unwrap();
            assert_eq!(id1, id2);
        }
    }

    #[test]
    fn font_resolution_alias_cache_is_bounded() {
        let mut fm = FontManager::new_deterministic().expect("bundled fonts should load");
        for index in 0..(RESOLUTION_CACHE_MAX_ENTRIES + 20) {
            fm.resolve_font(Some(&format!("Missing family {index}")), false, false)
                .expect("fallback font resolves");
        }
        assert!(fm.cache.len() <= RESOLUTION_CACHE_MAX_ENTRIES);
        assert!(fm.fonts.len() <= RESOLUTION_CACHE_MAX_ENTRIES);
    }

    #[test]
    fn bold_italic_variants() {
        let mut fm = FontManager::new();
        let regular = fm.resolve_font(None, false, false);
        let bold = fm.resolve_font(None, true, false);
        if let (Ok(r), Ok(b)) = (regular, bold) {
            // Bold should get a different font ID (different variant)
            assert_ne!(r, b);
        }
    }

    /// Latin text must resolve exactly as it did before, so the coverage check
    /// cannot disturb the overwhelmingly common case.
    #[test]
    fn latin_text_resolves_the_same_as_by_name() {
        let mut fm = FontManager::new();
        let Ok(by_name) = fm.resolve_font(Some("Arial"), false, false) else {
            return;
        };
        let for_text = fm
            .resolve_font_for_text(Some("Arial"), false, false, "Hello world")
            .unwrap();
        assert_eq!(by_name, for_text);
    }

    /// Text nothing can draw must keep the requested font rather than failing.
    ///
    /// The approved deterministic fallbacks intentionally do not cover emoji,
    /// so the search is guaranteed to come up empty. The text still needs a
    /// font so it occupies the right space.
    #[test]
    fn text_no_font_can_draw_keeps_the_requested_font() {
        let mut fm = FontManager::new_deterministic().expect("bundled fonts should load");
        let primary = fm.resolve_font(Some("Carlito"), false, false).unwrap();
        let resolved = fm
            .resolve_font_for_text(Some("Carlito"), false, false, "🀄")
            .unwrap();
        assert_eq!(
            primary, resolved,
            "with no covering font available the original must be kept"
        );
    }

    /// Whitespace absent from a font is not a reason to go hunting for another.
    #[test]
    fn whitespace_does_not_trigger_a_fallback() {
        let mut fm = FontManager::new_deterministic().expect("bundled fonts should load");
        let by_name = fm.resolve_font(Some("Carlito"), false, false).unwrap();
        let idx = fm.index_of(by_name).unwrap();
        // A non-breaking space and a tab, neither of which every face carries.
        assert!(
            fm.uncovered(idx, "a\u{00a0}b\tc")
                .iter()
                .all(|c| *c != '\t'),
            "control and whitespace characters must be ignored"
        );
    }

    /// When the machine does have a CJK font, CJK text must not keep a Latin
    /// font that cannot draw it.
    ///
    /// Skipped where no such font is installed, which is why it asserts
    /// nothing about which font is chosen.
    #[test]
    fn cjk_text_moves_off_a_latin_font_when_possible() {
        let mut fm = FontManager::new();
        let Ok(latin) = fm.resolve_font(Some("Liberation Serif"), false, false) else {
            return;
        };
        let Some(idx) = fm.index_of(latin) else {
            return;
        };
        if fm.uncovered(idx, "这是中文").is_empty() {
            return; // that font somehow covers it, nothing to prove
        }
        let resolved = fm
            .resolve_font_for_text(Some("Liberation Serif"), false, false, "这是中文")
            .unwrap();
        if resolved == latin {
            return; // no covering font installed on this machine
        }
        let new_idx = fm.index_of(resolved).unwrap();
        assert!(
            fm.uncovered(new_idx, "这是中文").len() < fm.uncovered(idx, "这是中文").len(),
            "the replacement must cover more of the text than the original"
        );
    }

    #[test]
    fn hangul_and_kana_receive_their_own_script_and_shaper_tag() {
        for character in [
            '\u{1100}', '\u{1112}', '\u{3131}', '\u{318e}', '\u{a960}', '\u{a97c}', '한', '힣',
            '\u{d7b0}', '\u{d7ff}', '\u{ffa0}', '\u{ffdc}',
        ] {
            assert_eq!(
                script_for_char(character),
                TextScript::Hangul,
                "{character:?} is Hangul"
            );
        }
        for character in [
            '\u{3041}', 'こ', 'ン', '\u{30ff}', '\u{31f0}', '\u{31ff}', '\u{ff66}', '\u{ff9f}',
        ] {
            assert_eq!(
                script_for_char(character),
                TextScript::Kana,
                "{character:?} is Kana"
            );
        }
        assert_eq!(
            harfrust_script(TextScript::Hangul),
            harfrust::script::HANGUL
        );
        assert_eq!(
            harfrust_script(TextScript::Kana),
            harfrust::script::HIRAGANA
        );

        // The two new ranges must be additions, never a re-routing. Sweeping
        // the whole Basic Multilingual Plane proves no codepoint that already
        // had a script identity moved to a different one.
        let previously_mapped = |codepoint: u32| match codepoint {
            0x0041..=0x024f | 0x1e00..=0x1eff => Some(TextScript::Latin),
            0x0590..=0x05ff => Some(TextScript::Hebrew),
            0x0600..=0x06ff | 0x0750..=0x077f | 0x08a0..=0x08ff => Some(TextScript::Arabic),
            0x0900..=0x097f | 0xa8e0..=0xa8ff => Some(TextScript::Devanagari),
            0x0e00..=0x0e7f => Some(TextScript::Thai),
            0x3400..=0x4dbf | 0x4e00..=0x9fff | 0xf900..=0xfaff => Some(TextScript::Han),
            _ => None,
        };
        for codepoint in 0..=0xffff_u32 {
            let Some(character) = char::from_u32(codepoint) else {
                continue;
            };
            let Some(expected) = previously_mapped(codepoint) else {
                continue;
            };
            assert_eq!(
                script_for_char(character),
                expected,
                "U+{codepoint:04X} changed script"
            );
        }
    }

    #[test]
    fn a_mixed_script_string_shapes_every_script_with_a_covering_bundled_face() {
        let mut fm = FontManager::new_deterministic().expect("bundled fonts should load");
        for text in ["안녕하세요", "こんにちは", "カタカナ", "שלום"] {
            let font_id = fm
                .resolve_font_for_text(Some("Carlito"), false, false, text)
                .expect("a bundled face covers the fixture repertoire");
            let index = fm.index_of(font_id).expect("resolved font is indexed");
            assert!(
                fm.uncovered(index, text).is_empty(),
                "no bundled face covers {text}"
            );
        }
    }

    #[test]
    fn fx210_owned_diagnostics_roll_back_provisional_scopes_and_replay_final_events() {
        let mut fonts = FontManager::new_deterministic().unwrap();
        fonts.begin_layout();
        let id = fonts.resolve_font(Some("Calibri"), false, false).unwrap();
        fonts
            .record_font_contribution(id, Some("Aptos"), "accepted")
            .unwrap();
        let accepted = fonts.current_font_diagnostics().to_vec();
        assert_eq!(accepted.len(), 1);
        let outer = fonts.begin_font_contribution_scope().unwrap();
        fonts
            .record_font_contribution(id, Some("Aptos Display"), "provisional")
            .unwrap();
        let inner = fonts.begin_font_contribution_scope().unwrap();
        fonts
            .record_font_contribution(id, Some("Aptos"), "nested")
            .unwrap();
        let nested = fonts
            .finish_font_contribution_scope(inner)
            .unwrap()
            .unwrap();
        assert_eq!(fonts.current_font_diagnostics().len(), 2);
        let staged = fonts
            .finish_font_contribution_scope(outer)
            .unwrap()
            .unwrap();
        assert_eq!(fonts.current_font_diagnostics(), accepted);
        fonts.replay_font_trace_events(&staged).unwrap();
        assert_eq!(fonts.current_font_diagnostics().len(), 2);
        fonts.replay_font_trace_events(&nested).unwrap();
        assert_eq!(fonts.current_font_diagnostics().len(), 2);
        let final_facts = fonts.current_font_diagnostics().to_vec();
        fonts.begin_layout();
        assert!(fonts.current_font_diagnostics().is_empty());
        fonts
            .record_font_contribution(id, Some("Aptos"), "")
            .unwrap();
        fonts
            .record_font_contribution(id, Some("Calibri"), "accepted")
            .unwrap();
        assert!(fonts.current_font_diagnostics().is_empty());
        fonts.replay_font_trace_events(&staged).unwrap();
        assert_eq!(fonts.current_font_diagnostics(), &final_facts[1..]);
    }

    #[test]
    fn fx210_overflow_cannot_lose_committed_warning_or_publish_provisional_warning() {
        let mut fonts = FontManager::new_deterministic().unwrap();
        fonts.begin_layout();
        let id = fonts.resolve_font(Some("Calibri"), false, false).unwrap();
        fonts
            .record_font_contribution(id, Some("Aptos"), "accepted")
            .unwrap();
        let accepted = fonts.current_font_diagnostics().to_vec();
        let scope = fonts.begin_font_contribution_scope().unwrap();
        for _ in 0..=PARAGRAPH_FONT_TRACE_MAX_ENTRIES {
            fonts
                .record_font_contribution(id, Some("Aptos Display"), "staged")
                .unwrap();
        }
        assert!(
            fonts
                .finish_font_contribution_scope(scope)
                .unwrap()
                .is_none()
        );
        assert_eq!(fonts.current_font_diagnostics(), accepted);
        fonts.begin_paragraph_font_trace();
        for _ in 0..=PARAGRAPH_FONT_TRACE_MAX_ENTRIES {
            fonts
                .record_font_contribution(id, Some("Calibri"), "accepted")
                .unwrap();
        }
        fonts
            .record_font_contribution(id, Some("Aptos Display"), "accepted")
            .unwrap();
        assert!(fonts.finish_paragraph_font_events().is_none());
        assert_eq!(fonts.current_font_diagnostics().len(), 2);
        let before = fonts.current_font_diagnostics().to_vec();
        assert!(
            fonts
                .replay_font_trace_events(&[FontTraceEvent::Contribution(FontContribution {
                    font_id: FontId(u32::MAX),
                    requested_family: Some("Aptos".into()),
                })])
                .is_err()
        );
        assert_eq!(fonts.current_font_diagnostics(), before);
    }

    #[test]
    fn fx211c_original_symbol_coverage_authenticates_physical_family_and_glyphs() {
        let source = include_bytes!("../fonts/Carlito-Regular.ttf");
        let mut fonts = FontManager::new_deterministic().unwrap();
        fonts.load_additional_fonts(&[FontFile {
            family: "Wingdings".to_owned(),
            data: source.to_vec(),
        }]);
        let alias = fonts.resolve_font(Some("Wingdings"), false, false).unwrap();
        assert!(!fonts.has_original_symbol_coverage(alias, Some("Wingdings"), "ÿ"));
        fonts.load_additional_fonts(&[FontFile {
            family: "Wingdings".to_owned(),
            data: font_with_family(source, "Wingdings"),
        }]);
        let actual = fonts.resolve_font(Some("Wingdings"), false, false).unwrap();
        assert!(fonts.has_original_symbol_coverage(actual, Some("Wingdings"), "ÿ"));
        assert!(
            fonts
                .unknown_symbol_diagnostics(actual, Some("Wingdings"), "ÿ", "ÿ")
                .is_empty()
        );
        assert!(
            fonts
                .unknown_symbol_diagnostics(
                    FontId(u32::MAX),
                    Some("Wingdings"),
                    "\u{f0ff}",
                    "\u{f0ff}"
                )
                .is_empty()
        );
        let genuine = fonts.font_data(actual).unwrap();
        let original = ttf_parser::Face::parse(source, 0).unwrap();
        let renamed = ttf_parser::Face::parse(genuine.data.as_ref(), 0).unwrap();
        assert_eq!(renamed.glyph_index('ÿ'), original.glyph_index('ÿ'));
        let count = u16::from_be_bytes([source[4], source[5]]) as usize;
        for table in 0..count {
            let tag = ttf_parser::Tag::from_bytes(
                source[12 + table * 16..16 + table * 16].try_into().unwrap(),
            );
            if tag != ttf_parser::Tag::from_bytes(b"name")
                && tag != ttf_parser::Tag::from_bytes(b"head")
            {
                assert_eq!(
                    renamed.raw_face().table(tag),
                    original.raw_face().table(tag)
                );
            }
        }

        assert!(!fonts.has_original_symbol_coverage(actual, Some("Wingdings"), "\u{f0ff}"));
        assert!(!fonts.has_original_symbol_coverage(actual, Some("Symbol"), "ÿ"));
        let resolved = fonts
            .resolve_symbol_text(Some("Wingdings"), false, false, "ÿ", true)
            .unwrap();
        assert_eq!(
            resolved,
            ("ÿ".to_owned(), actual),
            "genuine original glyph beats semantic fallback"
        );
    }

    #[test]
    fn fx210_explicit_symbol_table_is_bounded_and_covered() {
        let mut fonts = FontManager::new_deterministic().unwrap();
        for (family, code, target) in [
            ("Wingdings", 0xa7, '\u{25aa}'),
            ("Wingdings", 0xd8, '\u{27a2}'),
            ("Wingdings", 0xfc, '\u{2713}'),
            ("Wingdings", 0x76, '\u{2756}'),
            ("Wingdings", 0x71, '\u{2751}'),
            ("Wingdings", 0x6e, '\u{25a0}'),
            ("Wingdings", 0xfe, '\u{2611}'),
            ("Symbol", 0xb7, '\u{2022}'),
        ] {
            for code in [code, code + 0xf000] {
                let original = char::from_u32(code).unwrap().to_string();
                let (display, id) = fonts
                    .resolve_symbol_text(Some(family), false, false, &original, false)
                    .unwrap();
                assert_eq!(display, target.to_string());
                let shaped = fonts.shape_text(id, &display, 12.0).unwrap();
                assert!(!shaped.glyph_ids.is_empty());
                assert!(shaped.glyph_ids.iter().all(|glyph| *glyph != 0));
            }
        }
        for (family, original) in [
            ("Wingdings 2", "\u{f0a7}"),
            ("Unrelated", "\u{f0a7}"),
            ("Wingdings", "\u{f001}"),
            ("Wingdings", "\u{f0b7}"),
        ] {
            assert_eq!(
                fonts
                    .resolve_symbol_text(Some(family), false, false, original, false)
                    .unwrap()
                    .0,
                original
            );
        }
        assert_eq!(
            fonts
                .resolve_symbol_text(Some("Wingdings"), false, false, "\u{f0b7}", true)
                .unwrap()
                .0,
            "\u{2022}"
        );
        let id = fonts.resolve_font(Some("Carlito"), false, false).unwrap();
        assert_eq!(
            fonts
                .resolve_symbol_text(Some("Carlito"), false, false, "v", false)
                .unwrap(),
            ("v".into(), id)
        );
    }
    #[test]
    fn fx210_caller_labels_cannot_authenticate_aptos_or_original_symbol_faces() {
        let source = bundled_font_data()
            .into_iter()
            .find(|(family, _)| *family == "Carlito")
            .unwrap()
            .1;
        let mut fonts = FontManager::new_deterministic().unwrap();
        fonts.load_additional_fonts(&[FontFile {
            family: "Aptos".to_owned(),
            data: source.to_vec(),
        }]);
        fonts.begin_layout();
        let selected = fonts.resolve_font(Some("Aptos"), false, false).unwrap();
        let index = fonts.index_of(selected).unwrap();
        assert_eq!(fonts.fonts[index].data.as_ref(), source);
        assert_eq!(fonts.fonts[index].face_index, 0);
        assert!(fonts.physical_requested_family_matches(index, "Carlito"));
        assert!(!fonts.physical_requested_family_matches(index, "Aptos"));
        fonts
            .record_font_contribution(selected, Some("Aptos"), "actual")
            .unwrap();
        assert_eq!(fonts.current_font_diagnostics().len(), 1);
        assert!(
            fonts.current_font_diagnostics()[0]
                .message
                .contains("Font 'Aptos' is approximate")
        );
        assert!(
            fonts.current_font_diagnostics()[0]
                .message
                .contains("Carlito")
        );
        // Accepted caller aliases still choose the same original bytes on replay.
        fonts.begin_paragraph_font_trace();
        fonts
            .record_font_contribution(selected, Some("Aptos"), "actual")
            .unwrap();
        let events = fonts.finish_paragraph_font_events().unwrap();
        fonts.begin_layout();
        fonts.replay_font_trace_events(&events).unwrap();
        assert_eq!(fonts.current_font_diagnostics().len(), 1);
        assert_eq!(fonts.font_data(selected).unwrap().data.as_ref(), source);

        fonts.load_additional_fonts(&[FontFile {
            family: "Wingdings".to_owned(),
            data: source.to_vec(),
        }]);
        let aliased = fonts.resolve_font(Some("Wingdings"), false, false).unwrap();
        let index = fonts.index_of(aliased).unwrap();
        assert_eq!(fonts.fonts[index].data.as_ref(), source);
        assert!(
            fonts.uncovered(index, "ü").is_empty(),
            "Carlito genuinely covers the low-byte operand"
        );
        assert!(!fonts.physical_requested_family_matches(index, "Wingdings"));
        let (display, id) = fonts
            .resolve_symbol_text(Some("Wingdings"), false, false, "ü", false)
            .unwrap();
        assert_eq!(
            display, "✓",
            "a caller label cannot give Carlito's Latin glyph original Wingdings precedence"
        );
        assert!(
            fonts
                .shape_text(id, &display, 12.0)
                .unwrap()
                .glyph_ids
                .iter()
                .all(|glyph| *glyph != 0)
        );
        let (original, genuine) = fonts
            .resolve_symbol_text(Some("Carlito"), false, false, "ü", false)
            .unwrap();
        assert_eq!(original, "ü");
        let index = fonts.index_of(genuine).unwrap();
        assert!(fonts.physical_requested_family_matches(index, "Carlito"));
        assert!(fonts.uncovered(index, &original).is_empty());
        assert_eq!(fonts.fonts[index].data.as_ref(), source);
    }

    #[test]
    fn fx210_physical_family_authentication_reads_the_selected_collection_face() {
        let bundled = bundled_font_data();
        let caladea = bundled
            .iter()
            .find(|(family, _)| *family == "Caladea")
            .unwrap()
            .1;
        let carlito = bundled
            .iter()
            .find(|(family, _)| *family == "Carlito")
            .unwrap()
            .1;
        let collection = test_ttc(&[caladea, carlito]);
        let mut fonts = FontManager::new_deterministic().unwrap();
        fonts.load_additional_fonts(&[FontFile {
            family: "Two physical families".to_owned(),
            data: collection.clone(),
        }]);
        let selected = fonts.resolve_font(Some("Carlito"), false, false).unwrap();
        let index = fonts.index_of(selected).unwrap();
        assert_eq!(fonts.fonts[index].data.as_ref(), collection);
        assert_eq!(fonts.fonts[index].face_index, 1);
        assert!(fonts.physical_requested_family_matches(index, "Carlito"));
        assert!(!fonts.physical_requested_family_matches(index, "Caladea"));
        assert!(!fonts.physical_requested_family_matches(index, "Two physical families"));
        let first = fonts.resolve_font(Some("Caladea"), false, false).unwrap();
        let index = fonts.index_of(first).unwrap();
        assert_eq!(fonts.fonts[index].face_index, 0);
        assert!(fonts.physical_requested_family_matches(index, "Caladea"));
        assert!(!fonts.physical_requested_family_matches(index, "Carlito"));
        assert!(!fonts.physical_requested_family_matches(usize::MAX, "Carlito"));
    }
}
