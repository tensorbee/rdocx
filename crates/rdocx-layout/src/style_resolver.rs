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
    /// Concrete numbering instance that owns this counter sequence.
    #[doc(hidden)]
    pub num_id: u32,
    /// Abstract numbering definition that owns this counter sequence.
    #[doc(hidden)]
    pub definition_id: u32,
    /// Whether any level of this concrete instance carries a start override.
    #[doc(hidden)]
    pub has_start_override: bool,
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
        if !matches!(instruction.name.as_str(), "REF" | "STYLEREF") {
            return Err("paragraph-number reference formatting requires REF or STYLEREF".into());
        }
        // STYLEREF without a numbered caller compares the target with itself,
        // so relative context is its current level. REF still keeps full context.
        let source = source.or_else(|| (instruction.name == "STYLEREF").then_some(self));
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
        if instruction.name == "STYLEREF" {
            return Ok(value);
        }
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
        let shares_context = source.is_some_and(|source| {
            source.num_id == self.num_id
                || (source.definition_id == self.definition_id
                    && !source.has_start_override
                    && !self.has_start_override)
        });
        let mut start = source
            .filter(|_| shares_context)
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
/// `Clone` exists so a note laid out at more than one section width consumes
/// its list numbers once rather than once per width.
#[derive(Clone)]
pub struct NumberingState {
    /// (abstractNumId, ilvl) owns the visible sequence after native resolution.
    counters: HashMap<(u32, u32), u32>,
    /// Concrete IDs remain public source identities, not sequence identities.
    definition_by_instance: HashMap<u32, u32>,
    consumed_overrides: HashSet<(u32, u32)>,
    pending_sections: HashSet<u32>,
    resolved_by_source: HashMap<SourceNodeId, ResolvedNumbering>,
    pub(crate) navigation: HashMap<SourceNodeId, crate::ParagraphNavigation>,
    bookmark_sources: HashMap<String, SourceNodeId>,
    resolved_by_bookmark: HashMap<String, (SourceNodeId, ResolvedNumbering)>,
    main_story_sources: HashSet<SourceNodeId>,
    note_labels: HashMap<NoteRef, String>,
    // Set only while shaping an actual note story. Never a body reference.
    pub(crate) note_body_context: Option<NoteRef>,
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
            definition_by_instance: HashMap::new(),
            consumed_overrides: HashSet::new(),
            pending_sections: HashSet::new(),
            resolved_by_source: HashMap::new(),
            navigation: HashMap::new(),
            bookmark_sources: HashMap::new(),
            resolved_by_bookmark: HashMap::new(),
            main_story_sources: HashSet::new(),
            note_labels: HashMap::new(),
            note_body_context: None,
        }
    }

    /// Advance the counter for the given numId/ilvl and return the new value.
    pub fn advance(&mut self, num_id: u32, ilvl: u32, start: u32) -> u32 {
        let owner = self
            .definition_by_instance
            .get(&num_id)
            .copied()
            .unwrap_or(num_id);
        let key = (owner, ilvl);
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

    /// Get the current count for a level (without advancing).
    pub fn current(&self, num_id: u32, ilvl: u32) -> u32 {
        let owner = self
            .definition_by_instance
            .get(&num_id)
            .copied()
            .unwrap_or(num_id);
        self.counters.get(&(owner, ilvl)).copied().unwrap_or(0)
    }

    fn restart_deeper_levels(
        &mut self,
        num_id: u32,
        definition_id: u32,
        current_level: u32,
        numbering: &CT_Numbering,
        rearm_current_instance: bool,
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
                self.counters.remove(&(definition_id, deeper));
                if rearm_current_instance {
                    self.consumed_overrides.remove(&(num_id, deeper));
                }
            }
        }
    }

    /// Defer opted-in definition resets until their next actual numbered item.
    pub fn restart_after_section_break(&mut self, numbering: &CT_Numbering) {
        for instance in &numbering.nums {
            if numbering.restarts_after_section_break(instance.num_id)
                && self
                    .counters
                    .keys()
                    .any(|(owner, _)| *owner == instance.abstract_num_id)
            {
                self.pending_sections.insert(instance.abstract_num_id);
            }
        }
    }

    fn advance_numbering(
        &mut self,
        num_id: u32,
        definition_id: u32,
        level: u32,
        base_start: u32,
        override_start: Option<u32>,
        numbering: &CT_Numbering,
    ) -> u32 {
        self.definition_by_instance.insert(num_id, definition_id);
        let first_after_section = self.pending_sections.remove(&definition_id);
        if first_after_section {
            let root_start = resolved_numbering_level(num_id, 0, numbering)
                .and_then(|(base, _, _)| base.start)
                .unwrap_or(1);
            self.counters.insert((definition_id, 0), root_start);
            // A section event preserves concrete override-use state. A child
            // with lvlRestart=0 also retains its carried shared count.
            self.restart_deeper_levels(num_id, definition_id, 0, numbering, false);
            if level == 0 {
                return root_start;
            }
        }
        self.restart_deeper_levels(
            num_id,
            definition_id,
            level,
            numbering,
            !first_after_section,
        );
        if !first_after_section
            && let Some(start) = override_start
            && self.consumed_overrides.insert((num_id, level))
        {
            self.counters.insert((definition_id, level), start);
            return start;
        }
        self.advance(num_id, level, base_start)
    }

    pub(crate) fn record_navigation(
        &mut self,
        source: SourceNodeId,
        facts: crate::ParagraphNavigation,
    ) {
        self.navigation.insert(source, facts);
    }

    pub(crate) fn remove_navigation(
        &mut self,
        source: SourceNodeId,
    ) -> Option<crate::ParagraphNavigation> {
        self.navigation.remove(&source)
    }

    pub(crate) fn take_navigation(&mut self) -> HashMap<SourceNodeId, crate::ParagraphNavigation> {
        std::mem::take(&mut self.navigation)
    }

    pub(crate) fn resolved_for_source(&self, source: SourceNodeId) -> Option<&ResolvedNumbering> {
        self.resolved_by_source.get(&source)
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

    pub(crate) fn bookmark_sources(&self) -> impl Iterator<Item = (&str, SourceNodeId)> {
        self.bookmark_sources
            .iter()
            .map(|(name, node)| (name.as_str(), *node))
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
            definition_by_instance: HashMap::new(),
            consumed_overrides: HashSet::new(),
            pending_sections: HashSet::new(),
            resolved_by_source: HashMap::new(),
            navigation: self
                .navigation
                .iter()
                .filter(|(_, facts)| facts.suppressed_numbering)
                .map(|(node, facts)| (*node, facts.clone()))
                .collect(),
            bookmark_sources: self.bookmark_sources.clone(),
            resolved_by_bookmark: self.resolved_by_bookmark.clone(),
            main_story_sources: self.main_story_sources.clone(),
            note_labels: self.note_labels.clone(),
            note_body_context: self.note_body_context,
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

/// Generated section labels use named Word roles, independently of body defaults.
pub(crate) fn line_number_properties(
    styles: &CT_Styles,
    diagnostics: &mut Vec<oxml_layout::Diagnostic>,
) -> CT_RPr {
    let defaults = styles
        .doc_defaults
        .as_ref()
        .and_then(|defaults| defaults.rpr.as_ref());
    let normal = line_number_style_level(styles, "Normal", StyleType::Paragraph, diagnostics);
    let character =
        line_number_style_level(styles, "line number", StyleType::Character, diagnostics);
    let mut effective = defaults.cloned().unwrap_or_default();
    effective.merge_from(&normal);
    effective.merge_from(&character);
    let toggle = |default: Option<bool>, paragraph: Option<bool>, character: Option<bool>| {
        Some(default == Some(true) || (paragraph.unwrap_or(false) ^ character.unwrap_or(false)))
    };
    effective.bold = toggle(
        defaults.and_then(|rpr| rpr.bold),
        normal.bold,
        character.bold,
    );
    effective.italic = toggle(
        defaults.and_then(|rpr| rpr.italic),
        normal.italic,
        character.italic,
    );
    effective
}

fn line_number_style_level(
    styles: &CT_Styles,
    name: &str,
    kind: StyleType,
    diagnostics: &mut Vec<oxml_layout::Diagnostic>,
) -> CT_RPr {
    let named = styles
        .styles
        .iter()
        .filter(|style| {
            style
                .name
                .as_deref()
                .is_some_and(|value| value.trim() == name)
        })
        .collect::<Vec<_>>();
    if named.iter().any(|style| style.style_type != kind) {
        diagnostics.push(oxml_layout::Diagnostic {
            message: format!("line-number style role {name:?} has a wrong-type declaration"),
        });
    }
    let candidates = named
        .into_iter()
        .filter(|style| style.style_type == kind)
        .collect::<Vec<_>>();
    if candidates.len() > 1 {
        diagnostics.push(oxml_layout::Diagnostic {
            message: format!("line-number style role {name:?} is ambiguous"),
        });
        return CT_RPr::default();
    }
    let mut current = candidates.first().copied();
    let mut chain = Vec::new();
    let mut seen = HashSet::new();
    while let Some(style) = current {
        if matches!(
            style.style_id.as_str(),
            "DefaultParagraphFont" | "NoList" | "TableNormal"
        ) {
            break;
        }
        if !seen.insert(style.style_id.as_str()) {
            diagnostics.push(oxml_layout::Diagnostic {
                message: format!("line-number style role {name:?} has a basedOn cycle"),
            });
            break;
        }
        chain.push(style);
        let Some(parent) = style.based_on.as_deref() else {
            break;
        };
        match styles.get_by_id(parent) {
            Some(parent) if parent.style_type == kind => current = Some(parent),
            _ => {
                diagnostics.push(oxml_layout::Diagnostic {
                    message: format!(
                        "line-number style role {name:?} has a missing or wrong-type basedOn owner"
                    ),
                });
                break;
            }
        }
    }
    let mut properties = CT_RPr::default();
    for style in chain.into_iter().rev() {
        if let Some(rpr) = &style.rpr {
            properties.merge_from(rpr);
        }
    }
    properties
}

pub(crate) fn word_kerning_enabled(properties: &CT_RPr) -> bool {
    properties
        .kern
        .is_some_and(|threshold| threshold.0 <= properties.sz.map_or(22, |size| size.0))
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
    let num_fmt = lvl.num_fmt.clone().unwrap_or(ST_NumberFormat::Decimal);
    let numbering_instance = numbering.nums.iter().find(|item| item.num_id == num_id)?;
    let definition_id = numbering_instance.abstract_num_id;
    let has_start_override = numbering_instance
        .level_overrides
        .iter()
        .any(|level| level.start_override.is_some());
    let start = base_lvl.start.unwrap_or(1);
    let lvl_text = lvl.lvl_text.as_deref().unwrap_or("%1.");
    let legal = lvl.legal == Some(true);

    let count = state.advance_numbering(
        num_id,
        definition_id,
        ilvl,
        start,
        start_override,
        numbering,
    );
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
                    .map(|(base, _, _)| base.start.unwrap_or(1))
                    .unwrap_or(0);
                state
                    .counters
                    .get(&(definition_id, level))
                    .copied()
                    .unwrap_or(configured_start)
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
    let number_full = full_number_keeping_literals(num_id, ilvl, numbering, &number_context);

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
        definition_id,
        has_start_override,
        marker_rpr,
        suffix: lvl.suffix.unwrap_or(ST_LvlSuffix::Tab),
    })
}

fn format_lvl_text_keeping_literals(template: &str, context: &[String]) -> String {
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
        } else {
            result.push(character);
        }
    }
    result
}

/// Build `\w` by keeping each level's literal text.
///
/// A level whose template already names every ancestor placeholder replaces
/// the accumulated path. Any other level is appended. One trailing full stop
/// is removed from the finished path.
fn full_number_keeping_literals(
    num_id: u32,
    ilvl: u32,
    numbering: &CT_Numbering,
    context: &[String],
) -> String {
    let mut accumulated = String::new();
    for level in 0..=ilvl {
        let Some((_, lvl, _)) = resolved_numbering_level(num_id, level, numbering) else {
            continue;
        };
        let template = lvl.lvl_text.as_deref().unwrap_or("%1.");
        let formatted = format_lvl_text_keeping_literals(template, context);
        let contains_every_ancestor =
            (0..level).all(|ancestor| template.contains(&format!("%{}", ancestor + 1)));
        if contains_every_ancestor {
            accumulated = formatted;
        } else {
            accumulated.push_str(&formatted);
        }
    }
    if accumulated.ends_with('.') {
        accumulated.pop();
    }
    accumulated
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
                    .map(|(base, _, _)| base.start.unwrap_or(1))
                    .unwrap_or(0);
                let owner = state
                    .definition_by_instance
                    .get(&num_id)
                    .copied()
                    .unwrap_or(num_id);
                state
                    .counters
                    .get(&(owner, lvl_idx))
                    .copied()
                    .unwrap_or(configured_start)
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
    #[test]
    fn fx200_native_word_plain_markers_match_saved_and_reopened_print_oracle() {
        let numbering = CT_Numbering::from_xml(br###"<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<w:numbering xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mo="http://schemas.microsoft.com/office/mac/office/2008/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:mv="urn:schemas-microsoft-com:mac:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 wp14"><w:abstractNum w:abstractNumId="10"><w:multiLevelType w:val="singleLevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="101"><w:abstractNumId w:val="10"/></w:num><w:num w:numId="102"><w:abstractNumId w:val="10"/></w:num><w:abstractNum w:abstractNumId="11"><w:multiLevelType w:val="singleLevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="201"><w:abstractNumId w:val="11"/></w:num><w:num w:numId="202"><w:abstractNumId w:val="11"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num><w:abstractNum w:abstractNumId="12"><w:multiLevelType w:val="singleLevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="301"><w:abstractNumId w:val="12"/></w:num><w:abstractNum w:abstractNumId="13"><w:multiLevelType w:val="multilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="401"><w:abstractNumId w:val="13"/></w:num><w:num w:numId="402"><w:abstractNumId w:val="13"/></w:num><w:num w:numId="403"><w:abstractNumId w:val="13"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num></w:numbering>"###).unwrap();
        let mut state = NumberingState::new();
        assert_eq!(
            generate_marker(101, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "1.",
            "A_FIRST_1"
        );
        assert_eq!(
            generate_marker(101, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "2.",
            "A_FIRST_2"
        );
        assert_eq!(
            generate_marker(102, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "3.",
            "A_SECOND_1"
        );
        assert_eq!(
            generate_marker(102, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "4.",
            "A_SECOND_2"
        );
        assert_eq!(
            generate_marker(201, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "1.",
            "B_FIRST_1"
        );
        assert_eq!(
            generate_marker(201, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "2.",
            "B_FIRST_2"
        );
        assert_eq!(
            generate_marker(202, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "1.",
            "B_RESTART_1"
        );
        assert_eq!(
            generate_marker(202, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "2.",
            "B_RESTART_2"
        );
        assert_eq!(
            generate_marker(301, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "1.",
            "C_FIRST_1"
        );
        assert_eq!(
            generate_marker(301, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "2.",
            "C_FIRST_2"
        );
        assert_eq!(
            generate_marker(301, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "3.",
            "C_CONTINUE_1"
        );
        assert_eq!(
            generate_marker(301, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "4.",
            "C_CONTINUE_2"
        );
        assert_eq!(
            generate_marker(401, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "1.",
            "D_ROOT_1"
        );
        assert_eq!(
            generate_marker(401, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "1.1.",
            "D_CHILD_1"
        );
        assert_eq!(
            generate_marker(401, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "1.2.",
            "D_CHILD_2"
        );
        assert_eq!(
            generate_marker(401, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "2.",
            "D_ROOT_2"
        );
        assert_eq!(
            generate_marker(401, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "2.1.",
            "D_CHILD_AFTER_ROOT"
        );
        assert_eq!(
            generate_marker(402, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "3.",
            "D_FRESH_ROOT"
        );
        assert_eq!(
            generate_marker(402, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "3.1.",
            "D_FRESH_CHILD"
        );
        assert_eq!(
            generate_marker(403, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "1.",
            "D_RESTART_ROOT"
        );
        assert_eq!(
            generate_marker(403, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "1.1.",
            "D_RESTART_CHILD"
        );
    }

    #[test]
    fn fx200_native_word_extended_markers_match_saved_and_reopened_print_oracle() {
        let numbering = CT_Numbering::from_xml(br###"<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<w:numbering xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mo="http://schemas.microsoft.com/office/mac/office/2008/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:mv="urn:schemas-microsoft-com:mac:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 wp14"><w:abstractNum w:abstractNumId="20"><w:nsid w:val="B0912014"/><w:multiLevelType w:val="singleLevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="360" w:hanging="240"/></w:pPr></w:lvl></w:abstractNum><w:abstractNum w:abstractNumId="21"><w:nsid w:val="B0912015"/><w:multiLevelType w:val="singleLevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="360" w:hanging="240"/></w:pPr></w:lvl></w:abstractNum><w:abstractNum w:abstractNumId="22"><w:nsid w:val="B0912016"/><w:multiLevelType w:val="multilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="360" w:hanging="240"/></w:pPr></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="240"/></w:pPr></w:lvl></w:abstractNum><w:abstractNum w:abstractNumId="23"><w:nsid w:val="B0912017"/><w:multiLevelType w:val="singleLevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="360" w:hanging="240"/></w:pPr></w:lvl></w:abstractNum><w:abstractNum w:abstractNumId="24"><w:nsid w:val="B0912018"/><w:multiLevelType w:val="multilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="360" w:hanging="240"/></w:pPr></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlRestart w:val="0"/><w:lvlText w:val="%1.%2."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="240"/></w:pPr></w:lvl></w:abstractNum><w:abstractNum w:abstractNumId="25"><w:nsid w:val="B0912019"/><w:multiLevelType w:val="multilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="360" w:hanging="240"/></w:pPr></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="240"/></w:pPr></w:lvl><w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlRestart w:val="1"/><w:lvlText w:val="%1.%2.%3."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1080" w:hanging="240"/></w:pPr></w:lvl></w:abstractNum><w:abstractNum w:abstractNumId="26"><w:nsid w:val="B091201A"/><w:multiLevelType w:val="multilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="360" w:hanging="240"/></w:pPr></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlRestart w:val="0"/><w:lvlText w:val="%1.%2."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="240"/></w:pPr></w:lvl></w:abstractNum><w:abstractNum xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" w:abstractNumId="27" w15:restartNumberingAfterBreak="1"><w:nsid w:val="B091201B"/><w:multiLevelType w:val="singleLevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="360" w:hanging="240"/></w:pPr></w:lvl></w:abstractNum><w:abstractNum w:abstractNumId="28"><w:nsid w:val="B091201C"/><w:multiLevelType w:val="singleLevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="360" w:hanging="240"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="501"><w:abstractNumId w:val="20"/></w:num><w:num w:numId="502"><w:abstractNumId w:val="20"/></w:num><w:num w:numId="503"><w:abstractNumId w:val="20"/></w:num><w:num w:numId="601"><w:abstractNumId w:val="21"/></w:num><w:num w:numId="603"><w:abstractNumId w:val="21"/></w:num><w:num w:numId="604"><w:abstractNumId w:val="21"/></w:num><w:num w:numId="602"><w:abstractNumId w:val="21"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride></w:num><w:num w:numId="701"><w:abstractNumId w:val="22"/></w:num><w:num w:numId="702"><w:abstractNumId w:val="22"/><w:lvlOverride w:ilvl="1"><w:startOverride w:val="5"/></w:lvlOverride></w:num><w:num w:numId="703"><w:abstractNumId w:val="22"/></w:num><w:num w:numId="711"><w:abstractNumId w:val="23"/></w:num><w:num w:numId="712"><w:abstractNumId w:val="23"/><w:lvlOverride w:ilvl="0"><w:lvl w:ilvl="0"><w:start w:val="7"/><w:numFmt w:val="upperRoman"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="360" w:hanging="240"/></w:pPr></w:lvl></w:lvlOverride></w:num><w:num w:numId="713"><w:abstractNumId w:val="23"/></w:num><w:num w:numId="721"><w:abstractNumId w:val="24"/></w:num><w:num w:numId="722"><w:abstractNumId w:val="24"/></w:num><w:num w:numId="731"><w:abstractNumId w:val="25"/></w:num><w:num w:numId="732"><w:abstractNumId w:val="25"/></w:num><w:num w:numId="741"><w:abstractNumId w:val="26"/></w:num><w:num w:numId="742"><w:abstractNumId w:val="26"/><w:lvlOverride w:ilvl="1"><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlRestart w:val="1"/><w:lvlText w:val="%1.%2."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="240"/></w:pPr></w:lvl></w:lvlOverride></w:num><w:num w:numId="801"><w:abstractNumId w:val="27"/></w:num><w:num w:numId="802"><w:abstractNumId w:val="27"/></w:num><w:num w:numId="804"><w:abstractNumId w:val="27"/></w:num><w:num w:numId="803"><w:abstractNumId w:val="27"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride></w:num><w:num w:numId="901"><w:abstractNumId w:val="28"/></w:num><w:num w:numId="902"><w:abstractNumId w:val="28"/></w:num><w:num w:numId="904"><w:abstractNumId w:val="28"/></w:num><w:num w:numId="903"><w:abstractNumId w:val="28"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride></w:num></w:numbering>"###).unwrap();
        let mut state = NumberingState::new();
        assert_eq!(
            generate_marker(501, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "1.",
            "E_OLD_1"
        );
        assert_eq!(
            generate_marker(501, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "2.",
            "E_OLD_2"
        );
        assert_eq!(
            generate_marker(502, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "3.",
            "E_NEW_1"
        );
        assert_eq!(
            generate_marker(502, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "4.",
            "E_NEW_2"
        );
        assert_eq!(
            generate_marker(501, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "5.",
            "E_RETURN_1"
        );
        assert_eq!(
            generate_marker(501, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "6.",
            "E_RETURN_2"
        );
        assert_eq!(
            generate_marker(503, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "7.",
            "E_FRESH_1"
        );
        assert_eq!(
            generate_marker(502, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "8.",
            "E_SECOND_RETURN"
        );
        assert_eq!(
            generate_marker(601, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "1.",
            "F_BASE_1"
        );
        assert_eq!(
            generate_marker(601, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "2.",
            "F_BASE_2"
        );
        assert_eq!(
            generate_marker(602, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "5.",
            "F_OVERRIDE_1"
        );
        assert_eq!(
            generate_marker(602, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "6.",
            "F_OVERRIDE_2"
        );
        assert_eq!(
            generate_marker(603, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "7.",
            "F_PLAIN_AFTER"
        );
        assert_eq!(
            generate_marker(601, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "8.",
            "F_OLD_AFTER"
        );
        assert_eq!(
            generate_marker(602, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "9.",
            "F_OVERRIDE_RETURN"
        );
        assert_eq!(
            generate_marker(604, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "10.",
            "F_NEW_AFTER_RETURN"
        );
        assert_eq!(
            generate_marker(602, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "11.",
            "F_OVERRIDE_AGAIN"
        );
        assert_eq!(
            generate_marker(601, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "12.",
            "F_OLD_FINAL"
        );
        assert_eq!(
            generate_marker(701, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "1.",
            "G_ROOT"
        );
        assert_eq!(
            generate_marker(701, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "1.1.",
            "G_CHILD_1"
        );
        assert_eq!(
            generate_marker(701, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "1.2.",
            "G_CHILD_2"
        );
        assert_eq!(
            generate_marker(702, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "1.5.",
            "G_OVERRIDE_CHILD_1"
        );
        assert_eq!(
            generate_marker(702, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "1.6.",
            "G_OVERRIDE_CHILD_2"
        );
        assert_eq!(
            generate_marker(701, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "1.7.",
            "G_OLD_CHILD"
        );
        assert_eq!(
            generate_marker(703, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "1.8.",
            "G_PLAIN_CHILD"
        );
        assert_eq!(
            generate_marker(702, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "2.",
            "G_ROOT_NEXT"
        );
        assert_eq!(
            generate_marker(702, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "2.5.",
            "G_OVERRIDE_AFTER_ROOT"
        );
        assert_eq!(
            generate_marker(701, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "2.6.",
            "G_OLD_AFTER_ROOT"
        );
        assert_eq!(
            generate_marker(702, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "2.7.",
            "G_OVERRIDE_REENTER"
        );
        assert_eq!(
            generate_marker(711, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "1.",
            "H_BASE_1"
        );
        assert_eq!(
            generate_marker(711, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "2.",
            "H_BASE_2"
        );
        assert_eq!(
            generate_marker(712, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "III.",
            "H_REPLACEMENT_1"
        );
        assert_eq!(
            generate_marker(712, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "IV.",
            "H_REPLACEMENT_2"
        );
        assert_eq!(
            generate_marker(711, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "5.",
            "H_OLD_AFTER"
        );
        assert_eq!(
            generate_marker(713, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "6.",
            "H_PLAIN_AFTER"
        );
        assert_eq!(
            generate_marker(712, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "VII.",
            "H_REPLACEMENT_RETURN"
        );
        assert_eq!(
            generate_marker(721, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "1.",
            "I_ROOT_1"
        );
        assert_eq!(
            generate_marker(721, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "1.1.",
            "I_CHILD_1"
        );
        assert_eq!(
            generate_marker(721, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "1.2.",
            "I_CHILD_2"
        );
        assert_eq!(
            generate_marker(722, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "2.",
            "I_ROOT_2"
        );
        assert_eq!(
            generate_marker(722, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "2.3.",
            "I_CHILD_NEW"
        );
        assert_eq!(
            generate_marker(721, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "3.",
            "I_ROOT_3"
        );
        assert_eq!(
            generate_marker(721, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "3.4.",
            "I_CHILD_RETURN"
        );
        assert_eq!(
            generate_marker(731, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "1.",
            "J_ROOT_1"
        );
        assert_eq!(
            generate_marker(731, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "1.1.",
            "J_MIDDLE_1"
        );
        assert_eq!(
            generate_marker(731, 2, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "1.1.1.",
            "J_LOW_1"
        );
        assert_eq!(
            generate_marker(732, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "1.2.",
            "J_MIDDLE_NEW"
        );
        assert_eq!(
            generate_marker(732, 2, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "1.2.2.",
            "J_LOW_AFTER_MIDDLE"
        );
        assert_eq!(
            generate_marker(732, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "2.",
            "J_ROOT_NEW"
        );
        assert_eq!(
            generate_marker(731, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "2.1.",
            "J_MIDDLE_AFTER_ROOT"
        );
        assert_eq!(
            generate_marker(731, 2, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "2.1.1.",
            "J_LOW_AFTER_ROOT"
        );
        assert_eq!(
            generate_marker(741, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "1.",
            "K_ROOT_1"
        );
        assert_eq!(
            generate_marker(741, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "1.1.",
            "K_CHILD_BASE"
        );
        assert_eq!(
            generate_marker(742, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "1.2.",
            "K_CHILD_REPLACEMENT"
        );
        assert_eq!(
            generate_marker(742, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "2.",
            "K_ROOT_NEW"
        );
        assert_eq!(
            generate_marker(742, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "2.3.",
            "K_CHILD_AFTER_ROOT"
        );
        assert_eq!(
            generate_marker(741, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "2.4.",
            "K_BASE_RETURN"
        );
        assert_eq!(
            generate_marker(801, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "1.",
            "L_BEFORE_OLD"
        );
        assert_eq!(
            generate_marker(802, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "2.",
            "L_BEFORE_FRESH"
        );
        assert_eq!(
            generate_marker(803, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "5.",
            "L_BEFORE_OVERRIDE"
        );
        assert_eq!(
            generate_marker(901, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "1.",
            "M_BEFORE_OLD"
        );
        assert_eq!(
            generate_marker(902, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "2.",
            "M_BEFORE_FRESH"
        );
        assert_eq!(
            generate_marker(903, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "5.",
            "M_BEFORE_OVERRIDE"
        );
        state.restart_after_section_break(&numbering);
        assert_eq!(
            generate_marker(802, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "1.",
            "L_AFTER_FRESH"
        );
        assert_eq!(
            generate_marker(801, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "2.",
            "L_AFTER_OLD"
        );
        assert_eq!(
            generate_marker(803, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "3.",
            "L_AFTER_OVERRIDE"
        );
        assert_eq!(
            generate_marker(804, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "4.",
            "L_AFTER_NEW"
        );
        assert_eq!(
            generate_marker(803, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "5.",
            "L_AFTER_OVERRIDE_RETURN"
        );
        assert_eq!(
            generate_marker(902, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "6.",
            "M_AFTER_FRESH"
        );
        assert_eq!(
            generate_marker(901, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "7.",
            "M_AFTER_OLD"
        );
        assert_eq!(
            generate_marker(903, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "8.",
            "M_AFTER_OVERRIDE"
        );
        assert_eq!(
            generate_marker(904, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "9.",
            "M_AFTER_NEW"
        );
        assert_eq!(
            generate_marker(903, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "10.",
            "M_AFTER_OVERRIDE_RETURN"
        );
        state.restart_after_section_break(&numbering);
        assert_eq!(
            generate_marker(804, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "1.",
            "L_SECOND_PLAIN"
        );
        assert_eq!(
            generate_marker(803, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "2.",
            "L_SECOND_OVERRIDE"
        );
        assert_eq!(
            generate_marker(801, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "3.",
            "L_SECOND_OLD"
        );
        assert_eq!(
            generate_marker(904, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "11.",
            "M_SECOND_PLAIN"
        );
        assert_eq!(
            generate_marker(903, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "12.",
            "M_SECOND_OVERRIDE"
        );
        assert_eq!(
            generate_marker(901, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "13.",
            "M_SECOND_OLD"
        );
    }

    #[test]
    fn fx200_native_word_section_markers_match_saved_and_reopened_print_oracle() {
        let numbering = CT_Numbering::from_xml(br###"<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<w:numbering xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mo="http://schemas.microsoft.com/office/mac/office/2008/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:mv="urn:schemas-microsoft-com:mac:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 wp14"><w:abstractNum xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" w:abstractNumId="40" w15:restartNumberingAfterBreak="1"><w:nsid w:val="B0912028"/><w:multiLevelType w:val="singleLevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="360" w:hanging="240"/></w:pPr></w:lvl></w:abstractNum><w:abstractNum xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" w:abstractNumId="41" w15:restartNumberingAfterBreak="1"><w:nsid w:val="B0912029"/><w:multiLevelType w:val="singleLevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="360" w:hanging="240"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1001"><w:abstractNumId w:val="40"/></w:num><w:num w:numId="1002"><w:abstractNumId w:val="40"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride></w:num><w:num w:numId="1003"><w:abstractNumId w:val="40"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="9"/></w:lvlOverride></w:num><w:num w:numId="1101"><w:abstractNumId w:val="41"/></w:num><w:num w:numId="1102"><w:abstractNumId w:val="41"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride></w:num></w:numbering>"###).unwrap();
        let mut state = NumberingState::new();
        assert_eq!(
            generate_marker(1001, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "1.",
            "N_BEFORE_PLAIN"
        );
        assert_eq!(
            generate_marker(1002, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "5.",
            "N_BEFORE_OVERRIDE"
        );
        assert_eq!(
            generate_marker(1101, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "1.",
            "O_BEFORE_PLAIN"
        );
        state.restart_after_section_break(&numbering);
        assert_eq!(
            generate_marker(1002, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "1.",
            "N_USED_OVERRIDE_FIRST"
        );
        assert_eq!(
            generate_marker(1002, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "2.",
            "N_USED_OVERRIDE_SECOND"
        );
        assert_eq!(
            generate_marker(1001, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "3.",
            "N_PLAIN_AFTER"
        );
        assert_eq!(
            generate_marker(1003, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "9.",
            "N_UNUSED_OVERRIDE_AFTER"
        );
        assert_eq!(
            generate_marker(1102, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "1.",
            "O_UNUSED_OVERRIDE_FIRST"
        );
        assert_eq!(
            generate_marker(1102, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "5.",
            "O_UNUSED_OVERRIDE_SECOND"
        );
        assert_eq!(
            generate_marker(1101, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "6.",
            "O_PLAIN_AFTER"
        );
        state.restart_after_section_break(&numbering);
        assert_eq!(
            generate_marker(1003, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "1.",
            "N_PREVIOUSLY_UNUSED_OVERRIDE_FIRST"
        );
        assert_eq!(
            generate_marker(1002, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "2.",
            "N_ORIGINAL_OVERRIDE_AFTER"
        );
        assert_eq!(
            generate_marker(1102, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "1.",
            "O_NOW_USED_OVERRIDE_FIRST"
        );
    }

    #[test]
    fn fx200_native_word_level_section_markers_match_saved_and_reopened_print_oracle() {
        let numbering = CT_Numbering::from_xml(br###"<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<w:numbering xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mo="http://schemas.microsoft.com/office/mac/office/2008/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:mv="urn:schemas-microsoft-com:mac:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 wp14"><w:abstractNum xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" w:abstractNumId="50" w15:restartNumberingAfterBreak="1"><w:nsid w:val="B0912032"/><w:multiLevelType w:val="multilevel"/><w:lvl w:ilvl="0"><w:start w:val="3"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="360" w:hanging="240"/></w:pPr></w:lvl><w:lvl w:ilvl="1"><w:start w:val="4"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="240"/></w:pPr></w:lvl></w:abstractNum><w:abstractNum xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" w:abstractNumId="51" w15:restartNumberingAfterBreak="1"><w:nsid w:val="B0912033"/><w:multiLevelType w:val="multilevel"/><w:lvl w:ilvl="0"><w:start w:val="3"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="360" w:hanging="240"/></w:pPr></w:lvl><w:lvl w:ilvl="1"><w:start w:val="4"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="240"/></w:pPr></w:lvl></w:abstractNum><w:abstractNum xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" w:abstractNumId="52" w15:restartNumberingAfterBreak="1"><w:nsid w:val="B0912034"/><w:multiLevelType w:val="multilevel"/><w:lvl w:ilvl="0"><w:start w:val="3"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="360" w:hanging="240"/></w:pPr></w:lvl><w:lvl w:ilvl="1"><w:start w:val="4"/><w:numFmt w:val="decimal"/><w:lvlRestart w:val="0"/><w:lvlText w:val="%1.%2."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="240"/></w:pPr></w:lvl></w:abstractNum><w:abstractNum w:abstractNumId="53"><w:nsid w:val="B0912035"/><w:multiLevelType w:val="multilevel"/><w:lvl w:ilvl="0"><w:start w:val="3"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="360" w:hanging="240"/></w:pPr></w:lvl><w:lvl w:ilvl="1"><w:start w:val="4"/><w:numFmt w:val="decimal"/><w:lvlRestart w:val="0"/><w:lvlText w:val="%1.%2."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="240"/></w:pPr></w:lvl></w:abstractNum><w:abstractNum xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" w:abstractNumId="54" w15:restartNumberingAfterBreak="1"><w:nsid w:val="B0912036"/><w:multiLevelType w:val="singleLevel"/><w:lvl w:ilvl="0"><w:start w:val="3"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="360" w:hanging="240"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1201"><w:abstractNumId w:val="50"/></w:num><w:num w:numId="1202"><w:abstractNumId w:val="50"/><w:lvlOverride w:ilvl="1"><w:startOverride w:val="7"/></w:lvlOverride></w:num><w:num w:numId="1211"><w:abstractNumId w:val="51"/></w:num><w:num w:numId="1212"><w:abstractNumId w:val="51"/><w:lvlOverride w:ilvl="1"><w:startOverride w:val="9"/></w:lvlOverride></w:num><w:num w:numId="1301"><w:abstractNumId w:val="52"/></w:num><w:num w:numId="1302"><w:abstractNumId w:val="52"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="8"/></w:lvlOverride></w:num><w:num w:numId="1401"><w:abstractNumId w:val="53"/></w:num><w:num w:numId="1402"><w:abstractNumId w:val="53"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="8"/></w:lvlOverride></w:num><w:num w:numId="1501"><w:abstractNumId w:val="54"/></w:num><w:num w:numId="1502"><w:abstractNumId w:val="54"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="8"/></w:lvlOverride></w:num></w:numbering>"###).unwrap();
        let mut state = NumberingState::new();
        assert_eq!(
            generate_marker(1201, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "3.",
            "P_ROOT_BEFORE"
        );
        assert_eq!(
            generate_marker(1202, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "3.7.",
            "P_OVERRIDE_BEFORE"
        );
        assert_eq!(
            generate_marker(1211, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "3.",
            "Q_ROOT_BEFORE"
        );
        assert_eq!(
            generate_marker(1211, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "3.4.",
            "Q_CHILD_BEFORE"
        );
        assert_eq!(
            generate_marker(1301, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "3.",
            "R_ROOT_BEFORE"
        );
        assert_eq!(
            generate_marker(1301, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "3.4.",
            "R_CHILD_BEFORE"
        );
        assert_eq!(
            generate_marker(1302, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "8.",
            "R_ROOT_OVERRIDE_BEFORE"
        );
        assert_eq!(
            generate_marker(1302, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "8.5.",
            "R_CHILD_CARRY_BEFORE"
        );
        assert_eq!(
            generate_marker(1401, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "3.",
            "S_ROOT_BEFORE"
        );
        assert_eq!(
            generate_marker(1401, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "3.4.",
            "S_CHILD_BEFORE"
        );
        assert_eq!(
            generate_marker(1402, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "8.",
            "S_ROOT_OVERRIDE_BEFORE"
        );
        assert_eq!(
            generate_marker(1402, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "8.5.",
            "S_CHILD_CARRY_BEFORE"
        );
        assert_eq!(
            generate_marker(1501, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "3.",
            "T_BASE_BEFORE"
        );
        assert_eq!(
            generate_marker(1502, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "8.",
            "T_OVERRIDE_BEFORE"
        );
        state.restart_after_section_break(&numbering);
        assert_eq!(
            generate_marker(1202, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "3.4.",
            "P_USED_CHILD_FIRST"
        );
        assert_eq!(
            generate_marker(1202, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "3.5.",
            "P_USED_CHILD_SECOND"
        );
        assert_eq!(
            generate_marker(1201, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "3.6.",
            "P_PLAIN_CHILD_AFTER"
        );
        assert_eq!(
            generate_marker(1201, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "4.",
            "P_ROOT_AFTER"
        );
        assert_eq!(
            generate_marker(1202, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "4.4.",
            "P_OVERRIDE_AFTER_ROOT"
        );
        assert_eq!(
            generate_marker(1212, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "3.4.",
            "Q_UNUSED_CHILD_FIRST"
        );
        assert_eq!(
            generate_marker(1212, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "3.9.",
            "Q_UNUSED_CHILD_SECOND"
        );
        assert_eq!(
            generate_marker(1211, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "3.10.",
            "Q_PLAIN_CHILD_AFTER"
        );
        assert_eq!(
            generate_marker(1211, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "4.",
            "Q_ROOT_AFTER"
        );
        assert_eq!(
            generate_marker(1212, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "4.4.",
            "Q_OVERRIDE_AFTER_ROOT"
        );
        assert_eq!(
            generate_marker(1302, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "3.",
            "R_USED_ROOT_FIRST"
        );
        assert_eq!(
            generate_marker(1302, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "3.6.",
            "R_CHILD_FIRST"
        );
        assert_eq!(
            generate_marker(1301, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "4.",
            "R_PLAIN_ROOT_AFTER"
        );
        assert_eq!(
            generate_marker(1301, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "4.7.",
            "R_CHILD_CARRY_AFTER"
        );
        assert_eq!(
            generate_marker(1402, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "9.",
            "S_USED_ROOT_FIRST"
        );
        assert_eq!(
            generate_marker(1402, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "9.6.",
            "S_CHILD_FIRST"
        );
        assert_eq!(
            generate_marker(1401, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "10.",
            "S_PLAIN_ROOT_AFTER"
        );
        assert_eq!(
            generate_marker(1401, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "10.7.",
            "S_CHILD_CARRY_AFTER"
        );
        assert_eq!(
            generate_marker(1502, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "3.",
            "T_USED_OVERRIDE_FIRST"
        );
        assert_eq!(
            generate_marker(1502, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "4.",
            "T_USED_OVERRIDE_SECOND"
        );
        assert_eq!(
            generate_marker(1501, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "5.",
            "T_PLAIN_AFTER"
        );
    }

    #[test]
    fn fx200_native_word_ancestor_markers_match_saved_and_reopened_print_oracle() {
        let numbering = CT_Numbering::from_xml(br###"<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<w:numbering xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mo="http://schemas.microsoft.com/office/mac/office/2008/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:mv="urn:schemas-microsoft-com:mac:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 wp14"><w:abstractNum w:abstractNumId="60"><w:nsid w:val="B091203C"/><w:multiLevelType w:val="multilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="360" w:hanging="240"/></w:pPr></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="240"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1601"><w:abstractNumId w:val="60"/></w:num><w:num w:numId="1602"><w:abstractNumId w:val="60"/><w:lvlOverride w:ilvl="1"><w:startOverride w:val="7"/></w:lvlOverride></w:num></w:numbering>"###).unwrap();
        let mut state = NumberingState::new();
        assert_eq!(
            generate_marker(1601, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "1.",
            "U_PLAIN_ROOT_1"
        );
        assert_eq!(
            generate_marker(1602, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "1.7.",
            "U_OVERRIDE_CHILD_1"
        );
        assert_eq!(
            generate_marker(1601, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "2.",
            "U_PLAIN_ROOT_2"
        );
        assert_eq!(
            generate_marker(1602, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "2.1.",
            "U_OVERRIDE_CHILD_AFTER_PLAIN_ROOT"
        );
        assert_eq!(
            generate_marker(1602, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "3.",
            "U_OVERRIDE_ROOT_3"
        );
        assert_eq!(
            generate_marker(1602, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "3.7.",
            "U_OVERRIDE_CHILD_AFTER_SAME_ROOT"
        );
        assert_eq!(
            generate_marker(1601, 0, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "4.",
            "U_PLAIN_ROOT_4"
        );
        assert_eq!(
            generate_marker(1602, 1, &numbering, &mut state)
                .unwrap()
                .marker_text,
            "4.1.",
            "U_OVERRIDE_CHILD_AFTER_OTHER_ROOT_AGAIN"
        );
    }

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

        let parse = |content: &str| {
            CT_Styles::from_xml(format!(r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{content}</w:styles>"#).as_bytes()).unwrap()
        };
        let mut generated = parse(
            r#"<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/><w:rFonts w:ascii="Arial"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:styleId="PARENT"><w:name w:val="parent"/><w:rPr><w:b/><w:i/><w:sz w:val="32"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="not-a-built-in-id"><w:name w:val=" Normal "/><w:basedOn w:val="PARENT"/><w:rPr><w:b w:val="0"/></w:rPr></w:style><w:style w:type="character" w:styleId="character-parent"><w:name w:val="ordinary"/><w:rPr><w:b/><w:i/><w:kern w:val="32"/></w:rPr></w:style><w:style w:type="character" w:styleId="arbitrary"><w:name w:val="line number"/><w:basedOn w:val="character-parent"/></w:style>"#,
        );
        let original_xml = generated.to_xml().unwrap();
        let mut diagnostics = Vec::new();
        let ink = line_number_properties(&generated, &mut diagnostics);
        assert!(diagnostics.is_empty());
        assert_eq!(
            (
                ink.font_ascii.as_deref(),
                ink.sz,
                ink.bold,
                ink.italic,
                ink.kern
            ),
            (
                Some("Arial"),
                Some(HalfPoint(32)),
                Some(true),
                Some(false),
                Some(HalfPoint(32))
            )
        );
        assert!(
            word_kerning_enabled(&ink),
            "the nominal half-point tie enables kerning"
        );
        assert_eq!(generated.to_xml().unwrap(), original_xml);
        for threshold in [
            None,
            Some(HalfPoint(0)),
            Some(HalfPoint(31)),
            Some(HalfPoint(32)),
            Some(HalfPoint(33)),
        ] {
            let properties = CT_RPr {
                kern: threshold,
                sz: Some(HalfPoint(32)),
                ..Default::default()
            };
            assert_eq!(
                word_kerning_enabled(&properties),
                threshold.is_some_and(|value| value.0 <= 32)
            );
        }
        // Document-default true is dominant, otherwise paragraph/character toggles combine.
        for (default, normal, character, expected) in [
            (false, true, true, false),
            (true, true, true, true),
            (false, true, false, true),
            (false, false, false, false),
        ] {
            let xml = format!(
                r#"<w:docDefaults><w:rPrDefault><w:rPr><w:b w:val="{}"/><w:i w:val="{}"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:styleId="normal"><w:name w:val="Normal"/><w:rPr><w:b w:val="{}"/><w:i w:val="{}"/></w:rPr></w:style><w:style w:type="character" w:styleId="line"><w:name w:val="line number"/><w:rPr><w:b w:val="{}"/><w:i w:val="{}"/></w:rPr></w:style>"#,
                u8::from(default),
                u8::from(default),
                u8::from(normal),
                u8::from(normal),
                u8::from(character),
                u8::from(character)
            );
            let properties = line_number_properties(&parse(&xml), &mut Vec::new());
            assert_eq!(
                (properties.bold, properties.italic),
                (Some(expected), Some(expected))
            );
        }
        // Ignored IDs stop before their own properties and their basedOn edge.
        generated
            .styles
            .iter_mut()
            .find(|style| style.style_id == "character-parent")
            .unwrap()
            .style_id = "DefaultParagraphFont".into();
        generated
            .styles
            .iter_mut()
            .find(|style| style.name.as_deref() == Some("line number"))
            .unwrap()
            .based_on = Some("DefaultParagraphFont".into());
        let properties = line_number_properties(&generated, &mut Vec::new());
        assert_eq!(
            (properties.bold, properties.italic, properties.kern),
            (Some(false), Some(true), None)
        );
        // Duplicate primary names omit only that role, not document defaults or the other role.
        let duplicate = generated
            .styles
            .iter()
            .find(|style| style.name.as_deref() == Some("line number"))
            .unwrap()
            .clone();
        generated.styles.push(duplicate);
        let mut diagnostics = Vec::new();
        let properties = line_number_properties(&generated, &mut diagnostics);
        assert_eq!(properties.sz, Some(HalfPoint(32)));
        assert!(
            diagnostics
                .iter()
                .any(|item| item.message.contains("ambiguous"))
        );
        let unrelated = parse(
            r#"<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="unrelated"/><w:rPr><w:sz w:val="72"/></w:rPr></w:style><w:style w:type="character" w:styleId="wrong"><w:name w:val="Normal"/></w:style>"#,
        );
        let mut diagnostics = Vec::new();
        assert_eq!(
            line_number_properties(&unrelated, &mut diagnostics).sz,
            Some(HalfPoint(22))
        );
        assert!(
            diagnostics
                .iter()
                .any(|item| item.message.contains("wrong-type"))
        );
        let cycle = parse(
            r#"<w:style w:type="paragraph" w:styleId="a"><w:name w:val="Normal"/><w:basedOn w:val="b"/><w:rPr><w:sz w:val="32"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="b"><w:name w:val="parent"/><w:basedOn w:val="a"/><w:rPr><w:sz w:val="22"/></w:rPr></w:style>"#,
        );
        let mut diagnostics = Vec::new();
        assert_eq!(
            line_number_properties(&cycle, &mut diagnostics).sz,
            Some(HalfPoint(32))
        );
        assert_eq!(diagnostics.len(), 1);
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
            definition_id: 1,
            has_start_override: false,
            marker_rpr: CT_RPr::default(),
            suffix: ST_LvlSuffix::Tab,
        };
        for (instruction, expected) in [
            (r"REF target \r", "4.5.Clause 2"),
            (r"REF target \r \t", "4.5.2"),
            (r#"STYLEREF "Heading 3" \r"#, "Clause 2"),
            (r#"STYLEREF "Heading 3" \r \t"#, "2"),
        ] {
            let instruction =
                rdocx_oxml::text::Field::new(instruction, "old").effective_instruction();
            assert_eq!(
                target.numbered_reference_text(&instruction, None).unwrap(),
                expected
            );
        }
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

        source.num_id = 8;
        source.definition_id = 99;
        source.number_context = vec!["4".to_owned(), "5".to_owned(), "9".to_owned()];
        assert_eq!(target.relative_to(Some(&source), false), "4.5.Clause 2");
        source.definition_id = target.definition_id;
        source.has_start_override = true;
        assert_eq!(target.relative_to(Some(&source), false), "4.5.Clause 2");
        source.has_start_override = false;
        assert_eq!(target.relative_to(Some(&source), false), "Clause 2");
        assert_eq!(target.relative_to(None, true), "4.5.2");
    }

    /// Plain instances of one abstract definition continue its authenticated
    /// sequence unless an explicit override or section reset selects a restart.
    #[test]
    fn shared_abstract_definition_plain_instances_continue_definition_sequence() {
        let mut numbering = CT_Numbering::new();
        let first = numbering.add_numbered_list();
        let abstract_id = numbering
            .nums
            .iter()
            .find(|n| n.num_id == first)
            .unwrap()
            .abstract_num_id;

        // A second instance pointing at the same abstract definition.
        let second = numbering.nums.iter().map(|n| n.num_id).max().unwrap() + 1;
        numbering.nums.push(rdocx_oxml::numbering::CT_Num {
            num_id: second,
            abstract_num_id: abstract_id,
            abstract_num_id_raw: None,
            level_overrides: Vec::new(),
            extra_xml: Vec::new(),
            extra_attributes: Vec::new(),
        });

        let mut state = NumberingState::new();
        let a = generate_marker(first, 0, &numbering, &mut state).unwrap();
        let b = generate_marker(second, 0, &numbering, &mut state).unwrap();
        assert_eq!(a.marker_text, "1.");
        assert_eq!(b.marker_text, "2.");
        let c = generate_marker(first, 0, &numbering, &mut state).unwrap();
        assert_eq!(c.marker_text, "3.");
        let mut repeated = NumberingState::new();
        assert_eq!(
            generate_marker(first, 0, &numbering, &mut repeated)
                .unwrap()
                .marker_text,
            "1."
        );
        assert_eq!(
            generate_marker(second, 0, &numbering, &mut repeated)
                .unwrap()
                .marker_text,
            "2."
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
            "suppression must not advance or reset the shared definition sequence"
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

        assert!(generate_marker(num_id, 0, &numbering, &mut state).is_none());
        assert_eq!(state.current(num_id, 0), 1);
        assert!(generate_marker(num_id, 0, &numbering, &mut state).is_none());
        assert_eq!(state.current(num_id, 0), 2);
    }
}
