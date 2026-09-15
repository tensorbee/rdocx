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

#[cfg(all(test, feature = "system-fonts"))]
use std::sync::atomic::{AtomicUsize, Ordering};

use crate::error::{LayoutError, Result};
use crate::line::TextSegment;
use crate::output::{FontId, SourceSpan};

/// Font data provided by the user or extracted from an OOXML file.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FontFile {
    /// Font family name (e.g., "Calibri", "Arial").
    pub family: String,
    /// Raw font file bytes (TTF/OTF).
    pub data: Vec<u8>,
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
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TextScript {
    Latin,
    Arabic,
    Hebrew,
    Devanagari,
    Thai,
    Han,
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
            || !cluster_ranges_are_valid(&clusters, glyph_count, char_count, bidi_level % 2 == 1)
        {
            return Err(LayoutError::Layout(
                "invalid multilingual glyph positioning or cluster range".to_owned(),
            ));
        }
        Ok(Self {
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
        })
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
        let fingerprint = shaping_fingerprint(key.font_id, &key.text, key.size_bits);
        self.entries
            .push_back((key, shaped, entry_bytes, fingerprint));
    }
}

fn shaping_fingerprint(font_id: FontId, text: &str, size_bits: u64) -> u64 {
    use std::hash::{Hash, Hasher};

    let mut hasher = std::collections::hash_map::DefaultHasher::new();
    font_id.hash(&mut hasher);
    text.hash(&mut hasher);
    size_bits.hash(&mut hasher);
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
    /// HarfRust's per-face shaping caches. Building these is the expensive
    /// part of shaping, so it happens once per face instead of once per run.
    shaper_data: harfrust::ShaperData,
}

struct ParagraphFontTrace {
    ids: Vec<FontId>,
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
            caller_aliases: HashMap::new(),
            caller_families: HashMap::new(),
            explicit_aliases: Vec::new(),
            explicit_alias_map: HashMap::new(),
            shaping_memo: Mutex::new(ShapingMemo::new()),
            paragraph_font_trace: None,
            layout_fonts: Vec::new(),
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

        self.db = self.base_db.clone();
        self.caller_aliases.clone_from(&self.base_caller_aliases);
        self.caller_families.clone_from(&self.base_caller_families);
        for font_file in font_files {
            Self::load_caller_font(
                &mut self.db,
                &mut self.caller_aliases,
                &mut self.caller_families,
                &font_file.family,
                font_file.data.clone(),
            );
        }
        self.cache.clear();
        self.memory_face_data.clear();
        self.fonts.clear();
        self.next_id = 0;
        self.coverage_fallbacks.clear();
        self.coverage_misses.clear();
        self.additional_fonts = font_files.to_vec();
        if self.shaping_memo.is_poisoned() {
            self.shaping_memo.clear_poison();
        }
        self.shaping_memo
            .get_mut()
            .expect("shaping cache poison was cleared")
            .clear();
        true
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
    }

    /// Begin recording the exact resolution events for one cache candidate.
    #[doc(hidden)]
    pub fn begin_paragraph_font_trace(&mut self) {
        self.paragraph_font_trace = Some(ParagraphFontTrace {
            ids: Vec::new(),
            overflowed: false,
        });
    }

    /// Finish one bounded paragraph trace. An overflowed paragraph bypasses reuse.
    #[doc(hidden)]
    pub fn finish_paragraph_font_trace(&mut self) -> Option<Vec<FontId>> {
        let mut trace = self.paragraph_font_trace.take()?;
        if trace.overflowed {
            return None;
        }
        trace.ids.shrink_to_fit();
        Some(trace.ids)
    }

    /// Replay the exact resolution events attached to a cached paragraph.
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

        let query_family = |family: &str| {
            if let Some(ids) = self.caller_families.get(family)
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

        let mut found_id = query_family(family_name);
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

        let (units_per_em, ascender, descender, line_gap) = {
            let face = ttf_parser::Face::parse(&data, face_index)
                .map_err(|e| LayoutError::FontParse(format!("ttf-parser error: {e}")))?;
            (
                face.units_per_em(),
                face.ascender(),
                face.descender(),
                face.line_gap(),
            )
        };

        // Every metric and advance is scaled by size/upem, so a zero here would
        // turn the whole layout into infinities.
        if units_per_em == 0 {
            return Err(LayoutError::FontParse(format!(
                "font '{family_name}' declares zero units per em"
            )));
        }

        let shaper_data = {
            let face = harfrust::FontRef::from_index(&data, face_index)
                .map_err(|e| LayoutError::FontParse(format!("failed to read font face: {e}")))?;
            harfrust::ShaperData::new(&face)
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
            shaper_data,
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

    /// Shape a text string using HarfRust. Returns glyph IDs and advances.
    pub fn shape_text(&self, font_id: FontId, text: &str, size_pt: f64) -> Result<ShapedText> {
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
        };
        let fingerprint = shaping_fingerprint(key.font_id, &key.text, key.size_bits);
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

        let font = self.get_font(font_id)?;

        let face = harfrust::FontRef::from_index(&font.data, font.face_index)
            .map_err(|e| LayoutError::Shaping(format!("failed to read font face: {e}")))?;

        let shaper = font.shaper_data.shaper(&face).build();

        let mut buffer = harfrust::UnicodeBuffer::new();
        buffer.push_str(text);
        // Infer direction, script and language from the text. Unlike rustybuzz,
        // HarfRust does not do this implicitly and panics on an unset direction.
        buffer.guess_segment_properties();

        let output = shaper.shape(buffer, harfrust::ShapeOptions::default());
        let infos = output.glyph_infos();
        let positions = output.glyph_positions();

        let upem = font.units_per_em as f64;
        let scale = size_pt / upem;

        let mut glyph_ids = Vec::with_capacity(infos.len());
        let mut advances = Vec::with_capacity(positions.len());
        let mut total_width = 0.0;

        for (info, pos) in infos.iter().zip(positions.iter()) {
            glyph_ids.push(info.glyph_id as u16);
            let advance = pos.x_advance as f64 * scale;
            advances.push(advance);
            total_width += advance;
        }

        let shaped = ShapedText {
            glyph_ids,
            advances,
            width: total_width,
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
        self.shape_multilingual_with_levels(segment, language, no_wrap, &levels, 0)
    }

    /// Shape styled spans with one paragraph-wide bidi resolution.
    pub fn shape_multilingual_paragraph(
        &mut self,
        segments: Vec<(TextSegment, Option<String>)>,
        base_direction: TextDirection,
        no_wrap: bool,
    ) -> Result<Vec<MultilingualTextSegment>> {
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
        if paragraph_text.is_empty() {
            return Ok(Vec::new());
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
        let mut logical_index = 0usize;
        let mut shaped = Vec::new();
        for ((segment, language), byte_offset) in segments.into_iter().zip(segment_starts) {
            let byte_end = byte_offset + segment.text.len();
            if !segment.text.is_empty() {
                let forced_levels = (segment.direction != TextDirection::Auto)
                    .then(|| {
                        explicit_direction_levels(&segment.text, segment.direction, paragraph_level)
                    })
                    .transpose()?;
                let levels = forced_levels
                    .as_deref()
                    .unwrap_or(&bidi.levels[byte_offset..byte_end]);
                let spans = self.shape_multilingual_with_levels(
                    segment,
                    language.as_deref(),
                    no_wrap,
                    levels,
                    logical_index,
                )?;
                logical_index += spans.len();
                shaped.extend(spans);
            }
        }
        Ok(shaped)
    }

    fn shape_multilingual_with_levels(
        &mut self,
        segment: TextSegment,
        language: Option<&str>,
        no_wrap: bool,
        levels: &[unicode_bidi::Level],
        logical_index_base: usize,
    ) -> Result<Vec<MultilingualTextSegment>> {
        let grapheme_boundaries = icu_segmenter::GraphemeClusterSegmenter::new()
            .segment_str(&segment.text)
            .collect::<Vec<_>>();
        let break_offsets = if no_wrap {
            HashSet::new()
        } else {
            multilingual_break_opportunities(&segment.text)
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
            if grapheme_start > start && (script != current_script || level != current_level) {
                logical_ranges.push((start, grapheme_start, current_script, current_level));
                start = grapheme_start;
            }
            current_script = script;
            current_level = level;
            if break_offsets.contains(&grapheme_end) && grapheme_end < segment.text.len() {
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
                let font_id = self.font_for_multilingual_span(
                    segment.font_id,
                    &segment.text[grapheme_start..grapheme_end],
                    segment.bold,
                    segment.italic,
                );
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

        let mut logical = Vec::with_capacity(font_ranges.len());
        for (logical_index, (start, end, script, level, font_id)) in
            font_ranges.into_iter().enumerate()
        {
            let text = &segment.text[start..end];
            let metrics = self.metrics(font_id, segment.font_size)?;
            let direction = if level.is_rtl() {
                TextDirection::RightToLeft
            } else {
                TextDirection::LeftToRight
            };
            let positioned = self.shape_explicit(
                font_id,
                text,
                segment.font_size,
                script,
                language,
                direction,
            )?;
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
            base.glyph_ids = positioned.glyph_ids;
            base.advances = positioned.x_advances.clone();
            base.width = positioned.x_advances.iter().sum();
            base.ascent = metrics.ascent;
            base.descent = metrics.descent;
            base.line_gap = metrics.line_gap;
            logical.push(MultilingualTextSegment::new(
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
                break_offsets.contains(&end),
            )?);
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

    fn shape_explicit(
        &self,
        font_id: FontId,
        text: &str,
        size_pt: f64,
        script: TextScript,
        language: Option<&str>,
        direction: TextDirection,
    ) -> Result<PositionedShape> {
        let font = self.get_font(font_id)?;
        let face = harfrust::FontRef::from_index(&font.data, font.face_index)
            .map_err(|error| LayoutError::Shaping(format!("failed to read font face: {error}")))?;
        let shaper = font.shaper_data.shaper(&face).build();
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
        let output = shaper.shape(buffer, harfrust::ShapeOptions::default());
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
            if trace.ids.len() < PARAGRAPH_FONT_TRACE_MAX_ENTRIES {
                trace.ids.push(font_id);
            } else {
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
        TextScript::Common => harfrust::script::COMMON,
    }
}

fn multilingual_break_opportunities(text: &str) -> HashSet<usize> {
    let mut opportunities = icu_segmenter::WordSegmenter::new_auto(Default::default())
        .segment_str(text)
        .filter(|offset| *offset > 0 && *offset <= text.len())
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
    for id in ids {
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
                let data: Arc<[u8]> = Arc::from(data.as_ref().as_ref().to_vec());
                memory_face_data.insert(id, Arc::clone(&data));
                Some((data, face_index))
            }
        },
        #[cfg(feature = "system-fonts")]
        fontdb::Source::File(path) => shared_file_font_bytes(path).map(|data| (data, face_index)),
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
        "Segoe UI" => &["Segoe UI", "Carlito", "Liberation Sans"],
        "Tahoma" => &["Tahoma", "Liberation Sans", "Helvetica"],
        "Verdana" => &["Verdana", "Liberation Sans", "DejaVu Sans"],
        "Georgia" => &["Georgia", "Caladea", "Liberation Serif"],
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
        }
    }

    #[test]
    fn caller_font_labels_resolve_after_exact_embedded_families() {
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
            .expect("embedded family resolves exactly");
        assert_eq!(
            manager.fonts[manager.index_of(exact).unwrap()].family,
            "Carlito"
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

    fn font_with_family(source: &[u8], family: &str) -> Vec<u8> {
        assert_eq!(family.len(), 7);
        let mut font = source.to_vec();
        let table_count = u16::from_be_bytes([font[4], font[5]]) as usize;
        let name_offset = (0..table_count)
            .find_map(|table| {
                let record = 12 + table * 16;
                (&font[record..record + 4] == b"name").then(|| {
                    u32::from_be_bytes(font[record + 8..record + 12].try_into().unwrap()) as usize
                })
            })
            .expect("font has name table");
        let count = u16::from_be_bytes([font[name_offset + 2], font[name_offset + 3]]) as usize;
        let strings = name_offset
            + u16::from_be_bytes([font[name_offset + 4], font[name_offset + 5]]) as usize;
        for index in 0..count {
            let record = name_offset + 6 + index * 12;
            let platform = u16::from_be_bytes([font[record], font[record + 1]]);
            let name_id = u16::from_be_bytes([font[record + 6], font[record + 7]]);
            let length = u16::from_be_bytes([font[record + 8], font[record + 9]]) as usize;
            let offset = u16::from_be_bytes([font[record + 10], font[record + 11]]) as usize;
            if !matches!(name_id, 1 | 16) {
                continue;
            }
            let destination = &mut font[strings + offset..strings + offset + length];
            match (platform, length) {
                (0 | 3, 14) => {
                    for (bytes, ch) in destination.chunks_exact_mut(2).zip(family.bytes()) {
                        bytes.copy_from_slice(&(ch as u16).to_be_bytes());
                    }
                }
                (1, 7) => destination.copy_from_slice(family.as_bytes()),
                _ => {}
            }
        }
        font
    }

    #[cfg(feature = "system-fonts")]
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
        let forced = shaping_fingerprint(font, "second collision candidate", 11.0f64.to_bits());
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
}
