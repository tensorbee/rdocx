//! Style resolution: cascade styles and generate numbering markers.
//!
//! Ports the logic from `crates/rdocx/src/style.rs` since rdocx-layout
//! depends on rdocx-oxml directly (not rdocx).

use std::collections::{HashMap, HashSet};

use oxml_layout::{NoteRef, SourceNodeId};
use rdocx_oxml::borders::CT_Tabs;
use rdocx_oxml::numbering::{CT_Lvl, CT_Numbering, ST_LvlSuffix, ST_NumberFormat};
use rdocx_oxml::properties::{CT_PPr, CT_RPr};
use rdocx_oxml::shared::ST_TabJc;
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
    /// The current level's formatted counter without surrounding level text.
    pub number_current: String,
    /// The current level's complete number expression without trailing stops.
    #[doc(hidden)]
    pub number_level: String,
    /// The current level's number expression with literal text removed.
    #[doc(hidden)]
    pub number_level_without_text: String,
    /// Whether the current level expression embeds an ancestor placeholder.
    #[doc(hidden)]
    pub number_level_has_ancestor: bool,
    /// The complete visible number without trailing full stops.
    pub number_full: String,
    /// The complete contextual number with literal text removed.
    #[doc(hidden)]
    pub number_full_without_text: String,
    /// Formatted counters from the root level through the current level.
    #[doc(hidden)]
    pub number_context: Vec<String>,
    /// Contextual number suffixes beginning at each level, using source delimiters.
    #[doc(hidden)]
    pub number_suffixes_without_text: Vec<String>,
    /// Concrete numbering instance of the paragraph.
    #[doc(hidden)]
    pub num_id: u32,
    /// Abstract definition that owns this counter sequence, shared by every
    /// instance of it.
    #[doc(hidden)]
    pub abstract_num_id: u32,
    /// Run properties for the marker.
    pub marker_rpr: CT_RPr,
    /// Item that follows the marker before paragraph content begins.
    pub suffix: ST_LvlSuffix,
}

impl ResolvedNumbering {
    /// Format REF paragraph-number context and its current-level delimiter.
    /// Position suffixes and hyperlinks remain the caller's responsibility.
    #[doc(hidden)]
    pub fn numbered_reference_text(
        &self,
        instruction: &rdocx_oxml::text::FieldInstruction,
        source: Option<&Self>,
    ) -> std::result::Result<String, String> {
        use rdocx_oxml::text::FieldArgument;
        if instruction.name != "REF" {
            return Err("paragraph-number reference formatting requires REF".into());
        }
        let has = |name| {
            instruction
                .switches
                .iter()
                .any(|switch| switch.name == name)
        };
        let omit_text = has("t");
        let value = if has("w") {
            if omit_text {
                self.number_full_without_text.clone()
            } else {
                self.number_full.clone()
            }
        } else if has("r") {
            self.relative_to(source, omit_text)
        } else if has("n") {
            if omit_text {
                self.number_level_without_text.clone()
            } else {
                self.number_level.clone()
            }
        } else {
            return Err("REF paragraph-number formatting requires n, r or w".into());
        };
        let mut delimiters = instruction
            .switches
            .iter()
            .filter(|switch| switch.name == "d");
        let Some(delimiter) = delimiters.next() else {
            return Ok(value);
        };
        if delimiters.next().is_some() {
            return Err("REF delimiter switches are ambiguous".into());
        }
        let Some(FieldArgument::Text(delimiter)) = &delimiter.argument else {
            return Err("REF delimiter requires a text operand".into());
        };
        let current = if omit_text {
            &self.number_level_without_text
        } else {
            &self.number_level
        };
        if current.is_empty() {
            return if value.is_empty() {
                Ok(value)
            } else {
                Err("REF delimiter lacks a qualified current-level boundary".into())
            };
        }
        let prefix = value
            .strip_suffix(current)
            .ok_or_else(|| "REF delimiter lacks a qualified current-level boundary".to_owned())?;
        if prefix.is_empty() {
            return Ok(value);
        }
        // Actual Word controls preserve the ancestor stops and insert only at
        // this current-level boundary. An explicit empty operand inserts space.
        Ok(format!(
            "{prefix}{}{current}",
            if delimiter.is_empty() { " " } else { delimiter }
        ))
    }

    /// Format this paragraph number relative to a source paragraph.
    #[doc(hidden)]
    pub fn relative_to(&self, source: Option<&Self>, omit_text: bool) -> String {
        let mut start = source
            .filter(|source| source.abstract_num_id == self.abstract_num_id)
            .map(|source| {
                self.number_context
                    .iter()
                    .zip(&source.number_context)
                    .take_while(|(target, source)| target == source)
                    .count()
                    .min(self.number_context.len().saturating_sub(1))
            })
            .unwrap_or(0);
        if start < self.number_context.len().saturating_sub(1) && !self.number_level_has_ancestor {
            start = 0;
        }
        let suffix = if self.number_level_has_ancestor
            && start == self.number_context.len().saturating_sub(1)
        {
            self.number_level_without_text.clone()
        } else {
            self.number_suffixes_without_text
                .get(start)
                .cloned()
                .unwrap_or_else(|| self.number_full_without_text.clone())
        };
        if omit_text {
            return suffix;
        }
        match suffix.strip_suffix(&self.number_level_without_text) {
            Some(prefix) => format!("{prefix}{}", self.number_level),
            None => suffix,
        }
    }
}

/// Tracks numbering counters across paragraphs.
///
/// As in Word, the counters belong to the abstract numbering definition, so
/// every `w:num` instance of one definition continues the same sequence. An
/// instance's `w:startOverride` restarts its level at the instance's first
/// paragraph.
///
/// `Clone` exists so a note laid out at more than one section width consumes
/// its list numbers once rather than once per width.
#[derive(Clone)]
pub struct NumberingState {
    /// (abstractNumId, ilvl) → current count
    counters: HashMap<(u32, u32), u32>,
    /// Instances that have numbered a paragraph, so their start overrides
    /// apply once.
    started_instances: HashSet<u32>,
    resolved_by_source: HashMap<SourceNodeId, ResolvedNumbering>,
    bookmark_sources: HashMap<String, SourceNodeId>,
    resolved_by_bookmark: HashMap<String, (SourceNodeId, ResolvedNumbering)>,
    main_story_sources: HashSet<SourceNodeId>,
    note_labels: HashMap<NoteRef, String>,
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
            started_instances: HashSet::new(),
            resolved_by_source: HashMap::new(),
            bookmark_sources: HashMap::new(),
            resolved_by_bookmark: HashMap::new(),
            main_story_sources: HashSet::new(),
            note_labels: HashMap::new(),
        }
    }

    /// Advance the counter of an abstract definition's level and return the
    /// new value.
    pub fn advance(&mut self, abstract_num_id: u32, ilvl: u32, start: u32) -> u32 {
        let key = (abstract_num_id, ilvl);
        match self.counters.entry(key) {
            std::collections::hash_map::Entry::Vacant(entry) => {
                entry.insert(start);
                start
            }
            std::collections::hash_map::Entry::Occupied(mut entry) => {
                let value = entry.get().saturating_add(1);
                entry.insert(value);
                value
            }
        }
    }

    /// Get the current count of an abstract definition's level (without
    /// advancing).
    pub fn current(&self, abstract_num_id: u32, ilvl: u32) -> u32 {
        self.counters
            .get(&(abstract_num_id, ilvl))
            .copied()
            .unwrap_or(0)
    }

    /// Restart the levels an instance overrides with `w:startOverride` when
    /// the instance numbers its first paragraph.
    fn start_instance(&mut self, num_id: u32, abstract_num_id: u32, numbering: &CT_Numbering) {
        if !self.started_instances.insert(num_id) {
            return;
        }
        let Some(instance) = numbering.nums.iter().find(|item| item.num_id == num_id) else {
            return;
        };
        for level_override in &instance.level_overrides {
            if level_override.start_override.is_some() {
                self.counters
                    .remove(&(abstract_num_id, level_override.ilvl));
            }
        }
    }

    fn restart_deeper_levels(
        &mut self,
        num_id: u32,
        abstract_num_id: u32,
        current_level: u32,
        numbering: &CT_Numbering,
    ) {
        for deeper in (current_level + 1)..=8 {
            let restart = resolved_numbering_level(num_id, deeper, numbering)
                .and_then(|(base, _, _)| base.restart);
            let should_restart = match restart {
                None => true,
                Some(0) => false,
                Some(owner) => current_level < owner,
            };
            if should_restart {
                self.counters.remove(&(abstract_num_id, deeper));
            }
        }
    }

    /// Restart only the abstract definitions that opt in at a section
    /// boundary.
    pub fn restart_after_section_break(&mut self, numbering: &CT_Numbering) {
        self.counters.retain(|(abstract_num_id, _), _| {
            !numbering.nums.iter().any(|instance| {
                instance.abstract_num_id == *abstract_num_id
                    && numbering.restarts_after_section_break(instance.num_id)
            })
        });
    }

    pub(crate) fn record(&mut self, source: SourceNodeId, numbering: &ResolvedNumbering) {
        self.resolved_by_source.insert(source, numbering.clone());
    }

    pub(crate) fn record_bookmark(
        &mut self,
        name: &str,
        source: SourceNodeId,
        numbering: &ResolvedNumbering,
    ) {
        self.resolved_by_bookmark
            .insert(name.to_owned(), (source, numbering.clone()));
    }

    pub(crate) fn record_bookmark_source(&mut self, name: &str, source: SourceNodeId) {
        self.bookmark_sources.insert(name.to_owned(), source);
    }

    pub(crate) fn bookmark_source(&self, name: &str) -> Option<SourceNodeId> {
        self.bookmark_sources.get(name).copied()
    }

    pub(crate) fn bookmark(&self, name: &str) -> Option<(SourceNodeId, &ResolvedNumbering)> {
        self.resolved_by_bookmark
            .get(name)
            .map(|(source, numbering)| (*source, numbering))
    }

    pub(crate) fn source(&self, source: SourceNodeId) -> Option<&ResolvedNumbering> {
        self.resolved_by_source.get(&source)
    }

    pub(crate) fn record_main_story_source(&mut self, source: SourceNodeId) {
        self.main_story_sources.insert(source);
    }

    pub(crate) fn is_main_story_source(&self, source: SourceNodeId) -> bool {
        self.main_story_sources.contains(&source)
    }

    pub(crate) fn references_only(&self) -> Self {
        Self {
            counters: HashMap::new(),
            started_instances: HashSet::new(),
            resolved_by_source: HashMap::new(),
            bookmark_sources: self.bookmark_sources.clone(),
            resolved_by_bookmark: self.resolved_by_bookmark.clone(),
            main_story_sources: self.main_story_sources.clone(),
            note_labels: self.note_labels.clone(),
        }
    }

    pub(crate) fn set_note_labels(&mut self, labels: HashMap<NoteRef, String>) {
        self.note_labels = labels;
    }

    pub(crate) fn note_label(&self, note: NoteRef) -> String {
        self.note_labels
            .get(&note)
            .cloned()
            .unwrap_or_else(|| note.id.to_string())
    }

    pub(crate) fn merge_references(&mut self, other: &Self) {
        self.bookmark_sources.extend(other.bookmark_sources.clone());
        self.resolved_by_bookmark
            .extend(other.resolved_by_bookmark.clone());
        self.main_story_sources
            .extend(other.main_story_sources.iter().copied());
    }

    pub(crate) fn bookmark_numbering(
        &self,
    ) -> impl Iterator<Item = (&str, SourceNodeId, &ResolvedNumbering)> {
        self.resolved_by_bookmark
            .iter()
            .map(|(name, (source, numbering))| (name.as_str(), *source, numbering))
    }

    pub(crate) fn take_resolved(&mut self) -> HashMap<SourceNodeId, ResolvedNumbering> {
        std::mem::take(&mut self.resolved_by_source)
    }
}

/// Resolve paragraph properties by walking the style inheritance chain.
pub fn resolve_paragraph_properties(style_id: Option<&str>, styles: &CT_Styles) -> CT_PPr {
    resolve_paragraph_properties_in_table(style_id, styles, None)
}

/// Resolve paragraph properties with a table-style layer between document
/// defaults and the paragraph style.
pub fn resolve_paragraph_properties_in_table(
    style_id: Option<&str>,
    styles: &CT_Styles,
    table_properties: Option<&CT_PPr>,
) -> CT_PPr {
    let mut effective = CT_PPr::default();

    // 1. Start from docDefaults
    if let Some(ref defaults) = styles.doc_defaults
        && let Some(ref ppr) = defaults.ppr
    {
        merge_paragraph_layer(&mut effective, ppr);
    }

    if let Some(properties) = table_properties {
        merge_paragraph_layer(&mut effective, properties);
    }

    // 2. Walk the selected style's basedOn chain
    let selected_style_id = style_id.or_else(|| {
        styles
            .get_default(StyleType::Paragraph)
            .map(|style| style.style_id.as_str())
    });
    if let Some(sid) = selected_style_id {
        let chain = collect_style_chain(sid, styles);
        // Apply from most-base to most-derived
        for style in chain.iter().rev() {
            if let Some(ref ppr) = style.ppr {
                merge_paragraph_layer(&mut effective, ppr);
            }
        }
    }

    effective
}

/// Merge one layer of paragraph properties, with Word's rule for tab stops.
fn merge_paragraph_layer(effective: &mut CT_PPr, layer: &CT_PPr) {
    let inherited_tabs = effective.tabs.take();
    effective.merge_from(layer);
    effective.tabs = inherited_tabs;
    merge_tab_stops(&mut effective.tabs, layer.tabs.as_ref());
}

/// Apply one layer of `w:tabs` to the tab stops inherited so far.
///
/// Word adds a layer's stops to the ones a paragraph inherits from its
/// styles and its list level rather than replacing them. A stop replaces an
/// inherited stop at the same position, and a `clear` stop removes it. Word
/// reads the stops of a single layer in their listed order, and sorts stops
/// merged from several layers.
pub(crate) fn merge_tab_stops(inherited: &mut Option<CT_Tabs>, layer: Option<&CT_Tabs>) {
    let Some(layer) = layer else {
        return;
    };
    let merged = inherited.get_or_insert_with(CT_Tabs::default);
    let sort = !merged.tabs.is_empty();
    // The last stop a layer sets at a position is the one that counts.
    let last = layer
        .tabs
        .iter()
        .enumerate()
        .map(|(index, tab)| (tab.pos.0, index))
        .collect::<HashMap<_, _>>();
    merged
        .tabs
        .retain(|existing| !last.contains_key(&existing.pos.0));
    merged.tabs.extend(
        layer
            .tabs
            .iter()
            .enumerate()
            .filter(|(index, tab)| last[&tab.pos.0] == *index && tab.val != ST_TabJc::Clear)
            .map(|(_, tab)| tab.clone()),
    );
    if sort {
        merged.tabs.sort_by_key(|tab| tab.pos.0);
    }
}

/// Resolve run properties by walking paragraph and character style chains.
///
/// `table_properties` is the run layer a table style resolved for the cell the
/// paragraph sits in, and is `None` everywhere outside a styled table. It
/// merges after the document defaults and before the paragraph style, which is
/// the ordering `resolve_paragraph_properties_in_table` already uses.
pub fn resolve_run_properties(
    para_style_id: Option<&str>,
    run_style_id: Option<&str>,
    styles: &CT_Styles,
    table_properties: Option<&CT_RPr>,
) -> CT_RPr {
    let mut effective = CT_RPr::default();

    // 1. docDefaults run properties
    if let Some(ref defaults) = styles.doc_defaults
        && let Some(ref rpr) = defaults.rpr
    {
        effective.merge_from(rpr);
    }

    if let Some(properties) = table_properties {
        effective.merge_from(properties);
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
    if let Some(sid) = run_style_id.and_then(|style_id| character_style_id(style_id, styles)) {
        let chain = collect_style_chain(sid, styles);
        for style in chain.iter().rev() {
            if let Some(ref rpr) = style.rpr {
                effective.merge_from(rpr);
            }
        }
    }

    effective
}

fn character_style_id<'a>(style_id: &'a str, styles: &'a CT_Styles) -> Option<&'a str> {
    let style = styles.get_by_id(style_id)?;
    match style.style_type {
        StyleType::Character => Some(style.style_id.as_str()),
        StyleType::Paragraph => style.linked_style.as_deref().filter(|linked_id| {
            styles
                .get_by_id(linked_id)
                .is_some_and(|linked| linked.style_type == StyleType::Character)
        }),
        StyleType::Table | StyleType::Numbering => None,
    }
}

/// The paragraph properties a numbering level carries, mainly its indentation.
///
/// In the property chain these sit between the paragraph style and direct
/// formatting, so the indent for a list level applies unless the paragraph
/// sets its own.
pub fn level_paragraph_properties(
    num_id: u32,
    ilvl: u32,
    numbering: &CT_Numbering,
) -> Option<&CT_PPr> {
    let (_, level, _) = resolved_numbering_level(num_id, ilvl, numbering)?;
    level.ppr.as_ref()
}

/// The abstract definition whose counters a numbering instance advances.
fn instance_definition(num_id: u32, numbering: &CT_Numbering) -> Option<u32> {
    numbering
        .nums
        .iter()
        .find(|item| item.num_id == num_id)
        .map(|item| item.abstract_num_id)
}

fn resolved_numbering_level(
    num_id: u32,
    ilvl: u32,
    numbering: &CT_Numbering,
) -> Option<(&CT_Lvl, &CT_Lvl, Option<u32>)> {
    if num_id == 0 {
        return None;
    }
    let instance = numbering.nums.iter().find(|item| item.num_id == num_id)?;
    let definition = numbering
        .abstract_nums
        .iter()
        .find(|item| item.abstract_num_id == instance.abstract_num_id)?;
    let base = definition.levels.iter().find(|level| level.ilvl == ilvl)?;
    let level_override = instance
        .level_overrides
        .iter()
        .find(|value| value.ilvl == ilvl);
    let effective = level_override
        .and_then(|value| value.level.as_ref())
        .unwrap_or(base);
    Some((
        base,
        effective,
        level_override.and_then(|value| value.start_override),
    ))
}

/// Generate the marker text for a numbered/bulleted list item.
pub fn generate_marker(
    num_id: u32,
    ilvl: u32,
    numbering: &CT_Numbering,
    state: &mut NumberingState,
) -> Option<ResolvedNumbering> {
    let (base_lvl, lvl, start_override) = resolved_numbering_level(num_id, ilvl, numbering)?;
    let definition = instance_definition(num_id, numbering)?;
    let num_fmt = lvl.num_fmt.clone().unwrap_or(ST_NumberFormat::Decimal);
    let start = start_override.or(base_lvl.start).unwrap_or(1);
    let lvl_text = lvl.lvl_text.as_deref().unwrap_or("%1.");
    let legal = lvl.legal == Some(true);

    state.start_instance(num_id, definition, numbering);
    state.restart_deeper_levels(num_id, definition, ilvl, numbering);
    let count = state.advance(definition, ilvl, start);
    if matches!(num_fmt, ST_NumberFormat::None | ST_NumberFormat::Other(_)) {
        return None;
    }
    let displayed_num_fmt = if legal && num_fmt != ST_NumberFormat::Bullet {
        ST_NumberFormat::Decimal
    } else {
        num_fmt.clone()
    };
    if displayed_num_fmt != ST_NumberFormat::Bullet
        && !number_format_has_visible_renderer(&displayed_num_fmt)
    {
        return None;
    }
    for level in 0..=ilvl {
        if !lvl_text.contains(&format!("%{}", level + 1)) {
            continue;
        }
        let format = if legal {
            ST_NumberFormat::Decimal
        } else {
            resolved_numbering_level(num_id, level, numbering)
                .and_then(|(_, level, _)| level.num_fmt.clone())
                .unwrap_or(ST_NumberFormat::Decimal)
        };
        if !number_format_has_visible_renderer(&format) {
            return None;
        }
    }

    let number_current = match &displayed_num_fmt {
        ST_NumberFormat::Bullet => String::new(),
        ST_NumberFormat::None | ST_NumberFormat::Other(_) => unreachable!("returned above"),
        _ => format_number(count, displayed_num_fmt.clone()),
    };
    let marker_text = match displayed_num_fmt {
        ST_NumberFormat::Bullet => lvl_text.to_string(),
        ST_NumberFormat::None | ST_NumberFormat::Other(_) => unreachable!("returned above"),
        _ => format_lvl_text(lvl_text, num_id, ilvl, count, numbering, state, legal),
    };
    let number_context = (0..=ilvl)
        .map(|level| {
            let value = if level == ilvl {
                count
            } else {
                let configured_start = resolved_numbering_level(num_id, level, numbering)
                    .map(|(base, _, start_override)| start_override.or(base.start).unwrap_or(1))
                    .unwrap_or(0);
                let current = state.current(definition, level);
                if current == 0 {
                    configured_start
                } else {
                    current
                }
            };
            let format = if legal && level <= ilvl {
                ST_NumberFormat::Decimal
            } else {
                resolved_numbering_level(num_id, level, numbering)
                    .and_then(|(_, level, _)| level.num_fmt.clone())
                    .unwrap_or(ST_NumberFormat::Decimal)
            };
            format_number(value, format)
        })
        .collect::<Vec<_>>();
    let number_level = marker_text.trim_end_matches('.').to_owned();
    let number_level_without_text = format_lvl_text_without_literals(lvl_text, &number_context);
    let number_level_has_ancestor =
        (0..ilvl).any(|level| lvl_text.contains(&format!("%{}", level + 1)));
    let number_delimiters = contextual_number_delimiters(num_id, ilvl, numbering);
    let trailing_delimiter = trailing_number_delimiter(lvl_text);
    let embedded_start = (0..=ilvl)
        .find(|level| lvl_text.contains(&format!("%{}", level + 1)))
        .unwrap_or(ilvl) as usize;
    let number_suffixes_without_text = (0..number_context.len())
        .map(|start| {
            if embedded_start == 0 {
                let mut result = number_context[start].clone();
                for level in start + 1..number_context.len() {
                    result.push_str(&number_delimiters[level - 1]);
                    result.push_str(&number_context[level]);
                }
                result.push_str(&trailing_delimiter);
                return result.trim_end_matches('.').to_owned();
            }
            let mut result = String::new();
            if start < embedded_start {
                result.push_str(&number_context[start]);
                for level in start + 1..embedded_start {
                    result.push_str(&number_delimiters[level - 1]);
                    result.push_str(&number_context[level]);
                }
                result.push_str(&number_delimiters[embedded_start - 1]);
            }
            result.push_str(&number_level_without_text);
            result
        })
        .collect::<Vec<_>>();
    let number_full_without_text = number_suffixes_without_text
        .first()
        .cloned()
        .unwrap_or_default();
    let number_full = if (0..ilvl).all(|level| lvl_text.contains(&format!("%{}", level + 1))) {
        number_level.clone()
    } else {
        match number_full_without_text.strip_suffix(&number_level_without_text) {
            Some(prefix) => format!("{prefix}{number_level}"),
            None => number_level.clone(),
        }
    };

    let marker_rpr = lvl.rpr.clone().unwrap_or_default();

    Some(ResolvedNumbering {
        number_full,
        marker_text,
        number_current,
        number_level,
        number_level_without_text,
        number_level_has_ancestor,
        number_full_without_text,
        number_context,
        number_suffixes_without_text,
        num_id,
        abstract_num_id: definition,
        marker_rpr,
        suffix: lvl.suffix.unwrap_or(ST_LvlSuffix::Tab),
    })
}

fn format_lvl_text_without_literals(template: &str, context: &[String]) -> String {
    let mut result = String::new();
    let mut chars = template.chars().peekable();
    while let Some(character) = chars.next() {
        if character == '%'
            && let Some(digit @ '1'..='9') = chars.peek().copied()
        {
            chars.next();
            let level = digit as usize - '1' as usize;
            if let Some(value) = context.get(level) {
                result.push_str(value);
            }
        } else if is_number_delimiter(character) {
            result.push(character);
        }
    }
    result.trim_end_matches('.').to_owned()
}

fn is_number_delimiter(character: char) -> bool {
    !character.is_alphanumeric() && !character.is_whitespace()
}

fn trailing_number_delimiter(template: &str) -> String {
    let bytes = template.as_bytes();
    let last_placeholder_end = (0..bytes.len().saturating_sub(1))
        .filter(|index| bytes[*index] == b'%' && (b'1'..=b'9').contains(&bytes[*index + 1]))
        .map(|index| index + 2)
        .next_back();
    template
        .get(last_placeholder_end.unwrap_or(template.len())..)
        .unwrap_or_default()
        .chars()
        .filter(|character| is_number_delimiter(*character))
        .collect()
}

fn contextual_number_delimiters(num_id: u32, ilvl: u32, numbering: &CT_Numbering) -> Vec<String> {
    (1..=ilvl)
        .map(|right_level| {
            for candidate_level in (right_level..=ilvl).rev() {
                let Some((_, level, _)) =
                    resolved_numbering_level(num_id, candidate_level, numbering)
                else {
                    continue;
                };
                let Some(template) = level.lvl_text.as_deref() else {
                    continue;
                };
                let left = format!("%{right_level}");
                let right = format!("%{}", right_level + 1);
                let Some(left_start) = template.find(&left) else {
                    continue;
                };
                let Some(right_start) = template[left_start + left.len()..]
                    .find(&right)
                    .map(|offset| left_start + left.len() + offset)
                else {
                    continue;
                };
                let delimiter = template[left_start + left.len()..right_start]
                    .chars()
                    .filter(|character| is_number_delimiter(*character))
                    .collect::<String>();
                if !delimiter.is_empty() {
                    return delimiter;
                }
            }

            let left_level = right_level - 1;
            let Some((_, level, _)) = resolved_numbering_level(num_id, left_level, numbering)
            else {
                return String::new();
            };
            let Some(template) = level.lvl_text.as_deref() else {
                return String::new();
            };
            let placeholder = format!("%{}", left_level + 1);
            template
                .rfind(&placeholder)
                .map(|start| &template[start + placeholder.len()..])
                .unwrap_or_default()
                .chars()
                .filter(|character| is_number_delimiter(*character))
                .collect()
        })
        .collect()
}

/// Format level text by substituting %1, %2, etc. with formatted counters.
fn format_lvl_text(
    template: &str,
    num_id: u32,
    current_ilvl: u32,
    current_count: u32,
    numbering: &CT_Numbering,
    state: &NumberingState,
    legal: bool,
) -> String {
    let mut result = template.to_string();
    for lvl_idx in 0..=8u32 {
        let placeholder = format!("%{}", lvl_idx + 1);
        if result.contains(&placeholder) {
            let count = if lvl_idx == current_ilvl {
                current_count
            } else {
                let configured_start = resolved_numbering_level(num_id, lvl_idx, numbering)
                    .map(|(base, _, start_override)| start_override.or(base.start).unwrap_or(1))
                    .unwrap_or(0);
                let current = instance_definition(num_id, numbering)
                    .map_or(0, |definition| state.current(definition, lvl_idx));
                if current == 0 {
                    configured_start
                } else {
                    current
                }
            };
            let fmt = if legal && lvl_idx <= current_ilvl {
                ST_NumberFormat::Decimal
            } else {
                resolved_numbering_level(num_id, lvl_idx, numbering)
                    .and_then(|(_, level, _)| level.num_fmt.clone())
                    .unwrap_or(ST_NumberFormat::Decimal)
            };
            let formatted = format_number(count, fmt);
            result = result.replace(&placeholder, &formatted);
        }
    }
    result
}

/// Format a number according to ST_NumberFormat.
pub(crate) fn format_number(n: u32, fmt: ST_NumberFormat) -> String {
    match fmt {
        ST_NumberFormat::Decimal => n.to_string(),
        ST_NumberFormat::UpperRoman => to_roman(n, true),
        ST_NumberFormat::LowerRoman => to_roman(n, false),
        ST_NumberFormat::UpperLetter => to_letter(n, true),
        ST_NumberFormat::LowerLetter => to_letter(n, false),
        ST_NumberFormat::Bullet | ST_NumberFormat::None | ST_NumberFormat::Other(_) => {
            String::new()
        }
        // F-247 makes every standard OOXML token typed and round-trippable.
        // Rendering algorithms beyond the established decimal, letter, and
        // Roman families remain outside this package-level story.
        _ => String::new(),
    }
}

fn number_format_has_visible_renderer(format: &ST_NumberFormat) -> bool {
    matches!(
        format,
        ST_NumberFormat::Decimal
            | ST_NumberFormat::UpperRoman
            | ST_NumberFormat::LowerRoman
            | ST_NumberFormat::UpperLetter
            | ST_NumberFormat::LowerLetter
    )
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
    if upper { result } else { result.to_lowercase() }
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
            linked_style: None,
            auto_redefine: None,
            hidden: None,
            ui_priority: None,
            semi_hidden: None,
            unhide_when_used: None,
            quick_format: None,
            locked: None,
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
            table_properties: None,
            table_properties_original: None,
            table_properties_xml: None,
            table_row_properties: None,
            table_cell_properties: None,
            conditional_table_styles: Vec::new(),
            extra_attributes: Vec::new(),
            modeled_xml: Vec::new(),
            extra_xml: Vec::new(),
        });
        styles
    }

    #[test]
    fn tab_stop_layers_add_up_and_a_clear_stop_removes_one() {
        use rdocx_oxml::borders::CT_TabStop;

        let tabs = |stops: &[(ST_TabJc, i32)]| CT_Tabs {
            tabs: stops
                .iter()
                .map(|(val, pos)| CT_TabStop::new(*val, Twips(*pos)))
                .collect(),
        };
        let positions = |tabs: &Option<CT_Tabs>| -> Vec<(ST_TabJc, i32)> {
            tabs.as_ref()
                .map(|tabs| tabs.tabs.iter().map(|tab| (tab.val, tab.pos.0)).collect())
                .unwrap_or_default()
        };

        // One layer keeps its listed order, the last stop at a position wins.
        let mut merged = None;
        merge_tab_stops(
            &mut merged,
            Some(&tabs(&[
                (ST_TabJc::Right, 9500),
                (ST_TabJc::Left, 3000),
                (ST_TabJc::Left, 9500),
            ])),
        );
        assert_eq!(
            positions(&merged),
            [(ST_TabJc::Left, 3000), (ST_TabJc::Left, 9500)]
        );

        // A later layer adds, replaces and clears, and the result is sorted.
        merge_tab_stops(
            &mut merged,
            Some(&tabs(&[
                (ST_TabJc::Center, 3000),
                (ST_TabJc::Left, 2000),
                (ST_TabJc::Clear, 9500),
                (ST_TabJc::Clear, 100),
            ])),
        );
        assert_eq!(
            positions(&merged),
            [(ST_TabJc::Left, 2000), (ST_TabJc::Center, 3000)]
        );
        merge_tab_stops(&mut merged, None);
        assert_eq!(merged.as_ref().map(|tabs| tabs.tabs.len()), Some(2));
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
        let rpr = resolve_run_properties(Some("Heading2"), None, &styles, None);
        assert_eq!(rpr.font_ascii, Some("Calibri".to_string()));
        assert_eq!(rpr.sz, Some(HalfPoint(26)));
        assert_eq!(rpr.bold, Some(true));
        assert_eq!(rpr.color, Some("2E74B5".to_string()));
    }

    #[test]
    fn default_paragraph_style_resolves_its_based_on_chain() {
        let mut styles = test_styles();
        for style in &mut styles.styles {
            style.is_default = style.style_id == "Heading2";
        }
        let ppr = resolve_paragraph_properties(None, &styles);
        assert_eq!(ppr.keep_next, Some(true));
        assert_eq!(ppr.space_before, Some(Twips(40)));
        assert_eq!(ppr.space_after, Some(Twips(0)));
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
    fn numbering_marker_uses_the_level_suffix() {
        let mut numbering = CT_Numbering::new();
        let num_id = numbering.add_numbered_list();
        numbering.abstract_nums[0].levels[0].suffix = Some(ST_LvlSuffix::Nothing);

        let marker = generate_marker(num_id, 0, &numbering, &mut NumberingState::new())
            .expect("numbered marker resolves");

        assert_eq!(marker.suffix, ST_LvlSuffix::Nothing);
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

    #[test]
    fn number_text_suppression_keeps_source_delimiters_and_unicode_boundaries() {
        let context = ["2".to_owned(), "7".to_owned()];
        assert_eq!(
            format_lvl_text_without_literals("章节 %1 - Part %2)", &context),
            "2-7)"
        );
        assert_eq!(
            format_lvl_text_without_literals("Section %1.%2.", &context),
            "2.7"
        );
    }

    #[test]
    fn full_number_context_uses_delimiters_from_level_text() {
        let mut numbering = CT_Numbering::new();
        let num_id = numbering.add_numbered_list();
        for level in &mut numbering.abstract_nums[0].levels[..3] {
            level.num_fmt = Some(ST_NumberFormat::Decimal);
        }
        numbering.abstract_nums[0].levels[0].lvl_text = Some("%1)".to_owned());
        numbering.abstract_nums[0].levels[1].lvl_text = Some("%1-%2]".to_owned());
        numbering.abstract_nums[0].levels[2].lvl_text = Some("%3.".to_owned());
        let mut state = NumberingState::new();
        generate_marker(num_id, 0, &numbering, &mut state).unwrap();
        generate_marker(num_id, 1, &numbering, &mut state).unwrap();
        let marker = generate_marker(num_id, 2, &numbering, &mut state).unwrap();

        assert_eq!(marker.number_level, "1");
        assert_eq!(marker.number_full, "1-1]1");
        assert_eq!(marker.number_full_without_text, "1-1]1");
        assert_eq!(marker.number_suffixes_without_text, ["1-1]1", "1]1", "1"]);

        numbering.abstract_nums[0].levels[2].lvl_text = Some("(%3).".to_owned());
        let marker = generate_marker(num_id, 2, &numbering, &mut NumberingState::new()).unwrap();
        assert_eq!(
            marker.number_suffixes_without_text,
            ["1-1](1)", "1](1)", "(1)"]
        );
    }

    #[test]
    fn relative_numbering_omits_only_shared_source_context() {
        let target = ResolvedNumbering {
            marker_text: "Clause 4.5.2.".to_owned(),
            number_current: "2".to_owned(),
            number_level: "Clause 2".to_owned(),
            number_level_without_text: "2".to_owned(),
            number_level_has_ancestor: false,
            number_full: "4.5.Clause 2".to_owned(),
            number_full_without_text: "4.5.2".to_owned(),
            number_context: vec!["4".to_owned(), "5".to_owned(), "2".to_owned()],
            number_suffixes_without_text: vec![
                "4.5.2".to_owned(),
                "5.2".to_owned(),
                "2".to_owned(),
            ],
            num_id: 7,
            abstract_num_id: 3,
            marker_rpr: CT_RPr::default(),
            suffix: ST_LvlSuffix::Tab,
        };
        let mut source = target.clone();
        source.number_context = vec!["4".to_owned(), "3".to_owned(), "1".to_owned()];
        assert_eq!(target.relative_to(Some(&source), false), "4.5.Clause 2");
        assert_eq!(target.relative_to(Some(&source), true), "4.5.2");

        source.number_context = vec!["4".to_owned(), "5".to_owned(), "3".to_owned()];
        assert_eq!(target.relative_to(Some(&source), false), "Clause 2");
        assert_eq!(target.relative_to(Some(&source), true), "2");

        let mut same_parent_embedded = target.clone();
        same_parent_embedded.number_level = "Section 4.5.2".to_owned();
        same_parent_embedded.number_level_without_text = "4.5.2".to_owned();
        same_parent_embedded.number_level_has_ancestor = true;
        assert_eq!(
            same_parent_embedded.relative_to(Some(&source), false),
            "Section 4.5.2"
        );
        assert_eq!(
            same_parent_embedded.relative_to(Some(&source), true),
            "4.5.2"
        );

        let mut embedded = target.clone();
        embedded.number_level = "4.5.2".to_owned();
        embedded.number_level_without_text = "4.5.2".to_owned();
        embedded.number_level_has_ancestor = true;
        source.number_context = vec!["4".to_owned(), "3".to_owned(), "1".to_owned()];
        assert_eq!(embedded.relative_to(Some(&source), false), "5.2");

        // Another instance of the same definition shares its counters, so
        // its paragraphs give relative context too.
        source.num_id = 8;
        assert_eq!(embedded.relative_to(Some(&source), false), "5.2");
        source.abstract_num_id = 4;
        assert_eq!(target.relative_to(Some(&source), false), "4.5.Clause 2");
        assert_eq!(target.relative_to(None, true), "4.5.2");
    }

    /// A second instance of the abstract definition behind `first`, with
    /// `start_overrides` as `(level, start)` pairs.
    fn add_instance(
        numbering: &mut CT_Numbering,
        first: u32,
        start_overrides: &[(u32, u32)],
    ) -> u32 {
        let abstract_num_id = numbering
            .nums
            .iter()
            .find(|n| n.num_id == first)
            .unwrap()
            .abstract_num_id;
        let num_id = numbering.nums.iter().map(|n| n.num_id).max().unwrap() + 1;
        numbering.nums.push(rdocx_oxml::numbering::CT_Num {
            num_id,
            abstract_num_id,
            abstract_num_id_raw: None,
            level_overrides: start_overrides
                .iter()
                .map(|&(level, start)| {
                    let mut level_override = rdocx_oxml::numbering::CT_NumLvl::new(level);
                    level_override.start_override = Some(start);
                    level_override
                })
                .collect(),
            extra_xml: Vec::new(),
            extra_attributes: Vec::new(),
        });
        num_id
    }

    fn markers(numbering: &CT_Numbering, paragraphs: &[(u32, u32)]) -> Vec<String> {
        let mut state = NumberingState::new();
        paragraphs
            .iter()
            .map(|&(num_id, ilvl)| {
                generate_marker(num_id, ilvl, numbering, &mut state)
                    .unwrap()
                    .marker_text
            })
            .collect()
    }

    /// Instances that share an abstract definition share its counters.
    /// Word for Mac numbers two paragraphs on one instance and two on a
    /// second, plain instance of the same definition 1, 2, 3, 4 (checked for
    /// issue 321).
    #[test]
    fn instances_of_one_abstract_definition_share_its_counters() {
        let mut numbering = CT_Numbering::new();
        let first = numbering.add_numbered_list();
        let second = add_instance(&mut numbering, first, &[]);

        assert_eq!(
            markers(
                &numbering,
                &[(first, 0), (first, 0), (second, 0), (second, 0)]
            ),
            ["1.", "2.", "3.", "4."]
        );
    }

    /// An instance's `w:startOverride` restarts its level at the instance's
    /// first paragraph, as in Word, and later paragraphs of either instance
    /// continue the shared count.
    #[test]
    fn start_override_restarts_the_shared_count_at_the_instance_first_paragraph() {
        let mut numbering = CT_Numbering::new();
        let first = numbering.add_numbered_list();
        let restarted = add_instance(&mut numbering, first, &[(0, 1)]);

        assert_eq!(
            markers(
                &numbering,
                &[
                    (first, 0),
                    (first, 0),
                    (first, 0),
                    (restarted, 0),
                    (first, 0),
                ]
            ),
            // Per-instance counters would number the last paragraph "4.".
            ["1.", "2.", "3.", "1.", "2."]
        );
    }

    /// Restarting a sub-level through a new instance leaves level 0 alone:
    /// the level-0 items after it continue with "3.", as in Word and
    /// LibreOffice.
    #[test]
    fn sub_level_restart_does_not_reset_level_zero() {
        let mut numbering = CT_Numbering::new();
        let first = numbering.add_numbered_list();
        numbering.abstract_nums[0].levels[1].num_fmt = Some(ST_NumberFormat::Decimal);
        let restarted = add_instance(&mut numbering, first, &[(1, 1)]);

        assert_eq!(
            markers(
                &numbering,
                &[
                    (first, 0),
                    (first, 1),
                    (first, 0),
                    (restarted, 1),
                    (restarted, 1),
                    (restarted, 0),
                ]
            ),
            ["1.", "1.", "2.", "1.", "2.", "3."]
        );
    }

    #[test]
    fn num_id_zero_suppresses_only_the_selected_paragraph() {
        let mut numbering = CT_Numbering::new();
        let first = numbering.add_numbered_list();
        let abstract_id = numbering
            .nums
            .iter()
            .find(|item| item.num_id == first)
            .unwrap()
            .abstract_num_id;
        let second = first + 1;
        numbering.nums.push(rdocx_oxml::numbering::CT_Num {
            num_id: second,
            abstract_num_id: abstract_id,
            abstract_num_id_raw: None,
            level_overrides: Vec::new(),
            extra_xml: Vec::new(),
            extra_attributes: Vec::new(),
        });

        let mut state = NumberingState::new();
        assert_eq!(
            generate_marker(first, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "1."
        );
        assert!(generate_marker(0, 0, &numbering, &mut state).is_none());
        assert_eq!(
            generate_marker(first, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "2."
        );
        assert_eq!(
            generate_marker(second, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "3.",
            "suppression must not consume the shared definition count"
        );
    }

    #[test]
    fn counter_state_advances_in_declared_document_order() {
        let mut numbering = CT_Numbering::new();
        let num_id = numbering.add_numbered_list();
        let mut state = NumberingState::new();

        let body_before_table = generate_marker(num_id, 0, &numbering, &mut state).unwrap();
        let table_cell = generate_marker(num_id, 0, &numbering, &mut state).unwrap();
        let body_after_table = generate_marker(num_id, 0, &numbering, &mut state).unwrap();
        let next_section = generate_marker(num_id, 0, &numbering, &mut state).unwrap();

        assert_eq!(body_before_table.marker_text, "1.");
        assert_eq!(table_cell.marker_text, "2.");
        assert_eq!(body_after_table.marker_text, "3.");
        assert_eq!(next_section.marker_text, "4.");
    }

    #[test]
    fn start_override_is_reapplied_after_a_restart() {
        let mut numbering = CT_Numbering::new();
        let num_id = numbering.add_numbered_list();
        numbering.abstract_nums[0].levels[1].num_fmt = Some(ST_NumberFormat::Decimal);
        let instance = numbering
            .nums
            .iter_mut()
            .find(|item| item.num_id == num_id)
            .unwrap();
        let mut level_override = rdocx_oxml::numbering::CT_NumLvl::new(1);
        level_override.start_override = Some(5);
        instance.level_overrides.push(level_override);

        let mut state = NumberingState::new();
        assert_eq!(
            generate_marker(num_id, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "5."
        );
        assert_eq!(
            generate_marker(num_id, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "6."
        );
        generate_marker(num_id, 0, &numbering, &mut state).unwrap();
        assert_eq!(
            generate_marker(num_id, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "5."
        );
    }

    #[test]
    fn a_level_restarts_only_after_its_configured_owner() {
        let mut numbering = CT_Numbering::new();
        let num_id = numbering.add_numbered_list();
        numbering.abstract_nums[0].levels[2].num_fmt = Some(ST_NumberFormat::Decimal);
        numbering.abstract_nums[0].levels[2].restart = Some(1);
        let mut state = NumberingState::new();

        assert_eq!(
            generate_marker(num_id, 2, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "1."
        );
        generate_marker(num_id, 1, &numbering, &mut state).unwrap();
        assert_eq!(
            generate_marker(num_id, 2, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "2."
        );
        generate_marker(num_id, 0, &numbering, &mut state).unwrap();
        assert_eq!(
            generate_marker(num_id, 2, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "1."
        );
    }

    #[test]
    fn zero_restart_keeps_a_deeper_level_running() {
        let mut numbering = CT_Numbering::new();
        let num_id = numbering.add_numbered_list();
        numbering.abstract_nums[0].levels[1].num_fmt = Some(ST_NumberFormat::Decimal);
        numbering.abstract_nums[0].levels[1].restart = Some(0);
        let mut state = NumberingState::new();

        assert_eq!(
            generate_marker(num_id, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "1."
        );
        generate_marker(num_id, 0, &numbering, &mut state).unwrap();
        assert_eq!(
            generate_marker(num_id, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "2."
        );
    }

    #[test]
    fn replacement_level_restart_is_ignored_like_word() {
        let mut numbering = CT_Numbering::new();
        let num_id = numbering.add_numbered_list();
        numbering.abstract_nums[0].levels[1].num_fmt = Some(ST_NumberFormat::Decimal);
        numbering.abstract_nums[0].levels[1].restart = Some(0);
        let mut replacement = numbering.abstract_nums[0].levels[1].clone();
        replacement.restart = Some(1);
        let mut level_override = rdocx_oxml::numbering::CT_NumLvl::new(1);
        level_override.level = Some(replacement);
        numbering.nums[0].level_overrides.push(level_override);
        let mut state = NumberingState::new();

        assert_eq!(
            generate_marker(num_id, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "1."
        );
        generate_marker(num_id, 0, &numbering, &mut state).unwrap();
        assert_eq!(
            generate_marker(num_id, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "2."
        );
    }

    #[test]
    fn replacement_level_start_is_ignored_like_word() {
        let mut numbering = CT_Numbering::new();
        let num_id = numbering.add_numbered_list();
        numbering.abstract_nums[0].levels[0].start = Some(3);
        let mut replacement = numbering.abstract_nums[0].levels[0].clone();
        replacement.start = Some(9);
        let mut level_override = rdocx_oxml::numbering::CT_NumLvl::new(0);
        level_override.level = Some(replacement);
        numbering.nums[0].level_overrides.push(level_override);

        assert_eq!(
            generate_marker(num_id, 0, &numbering, &mut NumberingState::new())
                .unwrap()
                .marker_text,
            "3."
        );

        numbering.abstract_nums[0].levels[1].num_fmt = Some(ST_NumberFormat::Decimal);
        numbering.abstract_nums[0].levels[1].lvl_text = Some("%1.%2.".to_owned());
        let deeper = generate_marker(num_id, 1, &numbering, &mut NumberingState::new()).unwrap();
        assert_eq!(deeper.marker_text, "3.1.");
        assert_eq!(deeper.number_context, ["3", "1"]);
    }

    #[test]
    fn legal_numbering_converts_the_immediate_level_to_decimal() {
        let mut numbering = CT_Numbering::new();
        let num_id = numbering.add_numbered_list();
        let level = &mut numbering.abstract_nums[0].levels[0];
        level.num_fmt = Some(ST_NumberFormat::UpperRoman);
        level.legal = Some(true);

        assert_eq!(
            generate_marker(num_id, 0, &numbering, &mut NumberingState::new())
                .unwrap()
                .marker_text,
            "1."
        );
    }

    /// Separate abstract definitions are separate lists and each restarts.
    #[test]
    fn separate_abstract_definitions_count_independently() {
        let mut numbering = CT_Numbering::new();
        let first = numbering.add_numbered_list();
        let second = numbering.add_numbered_list();

        let mut state = NumberingState::new();
        let a = generate_marker(first, 0, &numbering, &mut state).unwrap();
        let b = generate_marker(second, 0, &numbering, &mut state).unwrap();
        assert_eq!(a.marker_text, "1.");
        assert_eq!(b.marker_text, "1.");
    }

    /// Each level carries its own indent, and deeper levels step further in.
    #[test]
    fn level_paragraph_properties_expose_per_level_indent() {
        let mut numbering = CT_Numbering::new();
        let num_id = numbering.add_numbered_list();

        let lvl0 = level_paragraph_properties(num_id, 0, &numbering)
            .expect("level 0 should carry paragraph properties");
        let lvl1 = level_paragraph_properties(num_id, 1, &numbering)
            .expect("level 1 should carry paragraph properties");

        let left0 = lvl0.ind_left.expect("level 0 indent").0;
        let left1 = lvl1.ind_left.expect("level 1 indent").0;
        assert!(
            left1 > left0,
            "level 1 must indent further than level 0, got {left0} then {left1}"
        );
    }

    #[test]
    fn producer_defined_number_formats_do_not_invent_layout_markers() {
        assert_eq!(
            format_number(7, ST_NumberFormat::Other("producerFormat".to_owned())),
            ""
        );
        assert_eq!(format_number(7, ST_NumberFormat::Decimal), "7");

        let mut numbering = CT_Numbering::new();
        let num_id =
            numbering.add_list(&[(ST_NumberFormat::Other("producerFormat".to_owned()), Some(1))]);
        assert!(generate_marker(num_id, 0, &numbering, &mut NumberingState::new()).is_none());
    }

    #[test]
    fn unsupported_standard_number_formats_do_not_invent_layout_markers() {
        let mut numbering = CT_Numbering::new();
        let num_id = numbering.add_numbered_list();
        numbering.abstract_nums[0].levels[0].num_fmt = Some(ST_NumberFormat::CardinalText);
        let mut state = NumberingState::new();

        let definition = numbering.abstract_nums[0].abstract_num_id;
        assert!(generate_marker(num_id, 0, &numbering, &mut state).is_none());
        assert_eq!(state.current(definition, 0), 1);
        assert!(generate_marker(num_id, 0, &numbering, &mut state).is_none());
        assert_eq!(state.current(definition, 0), 2);
    }
}
