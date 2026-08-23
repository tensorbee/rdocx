//! Font loading, resolution, shaping, and metrics.
//!
//! Uses fontdb for system font discovery, ttf-parser for metrics,
//! and HarfRust for text shaping.

use std::collections::{HashMap, HashSet, VecDeque};
#[cfg(feature = "system-fonts")]
use std::path::{Path, PathBuf};
#[cfg(feature = "system-fonts")]
use std::sync::OnceLock;
use std::sync::{Arc, Mutex};

#[cfg(all(test, feature = "system-fonts"))]
use std::sync::atomic::{AtomicUsize, Ordering};

use crate::error::{LayoutError, Result};
use crate::output::FontId;

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

#[derive(Debug, Clone, PartialEq, Eq)]
struct ShapingKey {
    font_id: FontId,
    text: String,
    size_bits: u64,
}

struct ShapingMemo {
    entries: VecDeque<(ShapingKey, ShapedText, usize)>,
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
        let entry_bytes = std::mem::size_of::<(ShapingKey, ShapedText, usize)>()
            + key.text.len()
            + shaped.glyph_ids.len() * std::mem::size_of::<u16>()
            + shaped.advances.len() * std::mem::size_of::<f64>();
        if entry_bytes > SHAPING_CACHE_MAX_BYTES {
            return;
        }
        while self.entries.len() >= SHAPING_CACHE_MAX_ENTRIES
            || self.bytes.saturating_add(entry_bytes) > SHAPING_CACHE_MAX_BYTES
        {
            let Some((_, _, evicted_bytes)) = self.entries.pop_front() else {
                break;
            };
            self.bytes = self.bytes.saturating_sub(evicted_bytes);
        }
        self.bytes += entry_bytes;
        self.entries.push_back((key, shaped, entry_bytes));
    }
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
    /// Requested-name -> real-family aliases for caller-provided fonts.
    ///
    /// Callers hand fonts over as `(family_name, bytes)` pairs, but fontdb
    /// only knows a face by the family recorded inside the file. Remembering
    /// the caller's label (lowercased) lets a document that asks for that
    /// name resolve to the supplied face even though the file calls itself
    /// something else -- e.g. an open Korean serif registered as "\uBC14\uD0D5".
    /// Rebuilt together with `additional_fonts`.
    caller_aliases: HashMap<String, String>,
    /// Caller-declared aliases set via [`Self::set_caller_aliases`], exactly
    /// as passed (for cheap change detection).
    explicit_aliases: Vec<(String, String)>,
    /// Lowercased requested-name -> target-family view of
    /// `explicit_aliases`, consulted before the label-derived aliases.
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

impl Default for FontManager {
    fn default() -> Self {
        Self::new()
    }
}

impl FontManager {
    fn from_base_database(db: fontdb::Database) -> Self {
        Self {
            base_db: db.clone(),
            db,
            cache: HashMap::new(),
            memory_face_data: HashMap::new(),
            fonts: Vec::new(),
            next_id: 0,
            coverage_fallbacks: HashMap::new(),
            coverage_misses: HashSet::new(),
            additional_fonts: Vec::new(),
            caller_aliases: HashMap::new(),
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
        self.caller_aliases.clear();
        for font_file in font_files {
            Self::load_caller_font(
                &mut self.db,
                &mut self.caller_aliases,
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
        let mut aliases = HashMap::new();
        for (name, data) in &fonts {
            Self::load_caller_font(&mut db, &mut aliases, name, data.clone());
        }
        let mut manager = Self::from_base_database(db);
        manager.caller_aliases = aliases;
        manager
    }

    /// Replace the caller-declared requested-name -> family aliases.
    ///
    /// Unlike [`Self::load_additional_fonts`] this carries no font bytes, so
    /// an arbitrarily large alias set stays cheap to pass on every layout.
    /// An unchanged set is a no-op; a changed set clears resolution state,
    /// because names may now resolve to different faces.
    pub fn set_caller_aliases(&mut self, aliases: &[(String, String)]) -> bool {
        if self.explicit_aliases == aliases {
            return false;
        }
        self.explicit_aliases = aliases.to_vec();
        self.explicit_alias_map = aliases
            .iter()
            .map(|(requested, target)| (requested.to_lowercase(), target.clone()))
            .collect();
        self.cache.clear();
        self.coverage_fallbacks.clear();
        self.coverage_misses.clear();
        true
    }

    /// Load one caller-provided font and remember its requested family name
    /// as an alias when the face's own family differs.
    fn load_caller_font(
        db: &mut fontdb::Database,
        aliases: &mut HashMap<String, String>,
        family: &str,
        data: Vec<u8>,
    ) {
        let before = db.len();
        db.load_font_data(data);
        let real = db
            .faces()
            .skip(before)
            .find_map(|face| face.families.first().map(|(name, _)| name.clone()));
        if let Some(real) = real {
            if !real.eq_ignore_ascii_case(family) {
                aliases.insert(family.to_lowercase(), real);
            }
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
            .retain(|(key, _, _)| current.contains(&key.font_id));
        memo.bytes = memo.entries.iter().map(|(_, _, bytes)| bytes).sum();
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

        // Map common Word font names to metric-compatible alternatives
        let mapped = map_font_name(family_name);

        // A caller-provided font registered under this name wins right after
        // the exact match (see `caller_aliases` / `explicit_alias_map`).
        let requested_key = family_name.to_lowercase();
        let caller_alias = self
            .explicit_alias_map
            .get(&requested_key)
            .or_else(|| self.caller_aliases.get(&requested_key))
            .cloned();

        // Try the requested font, mapped alternatives, then generic fallbacks
        let mut fallbacks: Vec<&str> = Vec::with_capacity(10);
        fallbacks.push(family_name);
        if let Some(alias) = caller_alias.as_deref() {
            if alias != family_name {
                fallbacks.push(alias);
            }
        }
        for alt in mapped {
            if *alt != family_name {
                fallbacks.push(alt);
            }
        }
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

        let mut found_id = None;
        for fallback in &fallbacks {
            let query = fontdb::Query {
                families: &[fontdb::Family::Name(fallback)],
                weight,
                style,
                stretch: fontdb::Stretch::Normal,
            };

            if let Some(id) = self.db.query(&query) {
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
        let mut memo = match self.shaping_memo.lock() {
            Ok(memo) => memo,
            Err(poisoned) => {
                let mut memo = poisoned.into_inner();
                memo.clear();
                self.shaping_memo.clear_poison();
                memo
            }
        };
        if let Some(index) = memo
            .entries
            .iter()
            .position(|(candidate, _, _)| candidate == &key)
        {
            let entry = memo.entries.remove(index).expect("cache index exists");
            let shaped = entry.1.clone();
            memo.entries.push_back(entry);
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
    use crate::bundled_fonts::bundled_font_data;

    #[test]
    fn caller_font_label_becomes_an_alias() {
        // A caller registers Caladea bytes under a Korean family name the
        // face itself does not carry; resolving that name must pick the
        // supplied face instead of drifting to the generic fallbacks.
        let mut fm = FontManager::new_deterministic().unwrap();
        let caladea = bundled_font_data()
            .into_iter()
            .find(|(family, _)| *family == "Caladea")
            .expect("Caladea is bundled")
            .1;
        fm.load_additional_fonts(&[FontFile {
            family: "\u{bc14}\u{d0d5}".to_string(), // 바탕
            data: caladea.to_vec(),
        }]);
        let id = fm.resolve_font(Some("\u{bc14}\u{d0d5}"), false, false).unwrap();
        let idx = fm.index_of(id).unwrap();
        assert_eq!(fm.fonts[idx].family, "Caladea");

        // A label matching the face's own family stores no alias.
        fm.load_additional_fonts(&[FontFile {
            family: "Caladea".to_string(),
            data: caladea.to_vec(),
        }]);
        assert!(fm.caller_aliases.is_empty());

        // Explicit aliases carry no bytes and resolve the same way.
        assert!(fm.set_caller_aliases(&[(
            "\u{bc14}\u{d0d5}".to_string(),
            "Caladea".to_string(),
        )]));
        let id = fm.resolve_font(Some("\u{bc14}\u{d0d5}"), false, false).unwrap();
        let idx = fm.index_of(id).unwrap();
        assert_eq!(fm.fonts[idx].family, "Caladea");
        // Unchanged set is a no-op.
        assert!(!fm.set_caller_aliases(&[(
            "\u{bc14}\u{d0d5}".to_string(),
            "Caladea".to_string(),
        )]));
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
    /// The bundled fonts have no CJK coverage, so in deterministic mode the
    /// search is guaranteed to come up empty. The text still needs a font so
    /// it occupies the right space.
    #[test]
    fn text_no_font_can_draw_keeps_the_requested_font() {
        let mut fm = FontManager::new_deterministic().expect("bundled fonts should load");
        let primary = fm.resolve_font(Some("Carlito"), false, false).unwrap();
        let resolved = fm
            .resolve_font_for_text(Some("Carlito"), false, false, "这是中文")
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
