//! Style access and manipulation for documents.

use std::collections::{HashMap, HashSet};

use rdocx_oxml::properties::{CT_PPr, CT_RPr};
use rdocx_oxml::styles::{CT_Style, CT_Styles, CT_TblStylePr, StyleType, TableStyleRegion};
use rdocx_oxml::table::{CT_TblPr, CT_TcPr, CT_TrPr};

use crate::{Error, Result};

/// An immutable reference to a style definition.
pub struct Style<'a> {
    pub(crate) inner: &'a CT_Style,
}

impl<'a> Style<'a> {
    /// The style ID (used to reference this style).
    pub fn style_id(&self) -> &str {
        &self.inner.style_id
    }

    /// The display name of the style.
    pub fn name(&self) -> Option<&str> {
        self.inner.name.as_deref()
    }

    /// The style ID this style is based on.
    pub fn based_on(&self) -> Option<&str> {
        self.inner.based_on.as_deref()
    }

    /// The kind of content this style formats.
    pub fn style_type(&self) -> StyleType {
        self.inner.style_type
    }

    /// The reciprocal paragraph or character style linked to this style.
    pub fn linked_style(&self) -> Option<&str> {
        self.inner.linked_style.as_deref()
    }

    /// The style selected for the paragraph following this paragraph style.
    pub fn next_style(&self) -> Option<&str> {
        self.inner.next_style.as_deref()
    }

    /// The user-interface ordering priority for this style.
    pub fn priority(&self) -> Option<u32> {
        self.inner.ui_priority
    }

    /// Whether Word may redefine the style from direct formatting.
    pub fn auto_redefine(&self) -> Option<bool> {
        self.inner.auto_redefine
    }

    /// Whether the style is hidden from all user-interface lists.
    pub fn hidden(&self) -> Option<bool> {
        self.inner.hidden
    }

    /// Whether the style is hidden from the main style gallery.
    pub fn semi_hidden(&self) -> Option<bool> {
        self.inner.semi_hidden
    }

    /// Whether applying the style makes it visible in the user interface.
    pub fn unhide_when_used(&self) -> Option<bool> {
        self.inner.unhide_when_used
    }

    /// Whether the style appears in the quick style gallery.
    pub fn quick_format(&self) -> Option<bool> {
        self.inner.quick_format
    }

    /// Whether the style is locked against application in the user interface.
    pub fn locked(&self) -> Option<bool> {
        self.inner.locked
    }

    /// The style's paragraph properties.
    pub fn paragraph_properties(&self) -> Option<&CT_PPr> {
        self.inner.ppr.as_ref()
    }

    /// The style's run properties.
    pub fn run_properties(&self) -> Option<&CT_RPr> {
        self.inner.rpr.as_ref()
    }

    /// The style's base table properties.
    pub fn table_properties(&self) -> Option<&CT_TblPr> {
        self.inner.table_properties.as_ref()
    }

    /// The style's base table row properties.
    pub fn table_row_properties(&self) -> Option<&CT_TrPr> {
        self.inner.table_row_properties.as_deref()
    }

    /// The style's base table cell properties.
    pub fn table_cell_properties(&self) -> Option<&CT_TcPr> {
        self.inner.table_cell_properties.as_deref()
    }

    /// Conditional table regions in source order.
    ///
    /// A region whose `w:type` this workspace does not recognise reports
    /// `None` from [`ConditionalTableStyle::region`]. It round-trips from its
    /// preserved bytes and takes no part in resolution.
    pub fn conditional_table_styles(&self) -> Vec<ConditionalTableStyle<'a>> {
        self.inner
            .conditional_table_styles
            .iter()
            .map(|inner| ConditionalTableStyle { inner })
            .collect()
    }

    /// Whether this is the default style for its type.
    pub fn is_default(&self) -> bool {
        self.inner.is_default
    }
}

/// One conditional table-style region read back through the facade.
pub struct ConditionalTableStyle<'a> {
    inner: &'a CT_TblStylePr,
}

impl<'a> ConditionalTableStyle<'a> {
    /// The region this layer formats, or `None` for an unrecognised `w:type`.
    pub fn region(&self) -> Option<TableStyleRegion> {
        self.inner.region
    }

    /// The region's paragraph properties.
    pub fn paragraph_properties(&self) -> Option<&'a CT_PPr> {
        self.inner.paragraph_properties.as_ref()
    }

    /// The region's run properties.
    pub fn run_properties(&self) -> Option<&'a CT_RPr> {
        self.inner.run_properties.as_deref()
    }

    /// The region's table properties.
    pub fn table_properties(&self) -> Option<&'a CT_TblPr> {
        self.inner.table_properties.as_ref()
    }

    /// The region's row properties.
    ///
    /// Modeled and round-tripped. Its layout application belongs to F-268a.
    pub fn row_properties(&self) -> Option<&'a CT_TrPr> {
        self.inner.row_properties.as_deref()
    }

    /// The region's cell properties.
    pub fn cell_properties(&self) -> Option<&'a CT_TcPr> {
        self.inner.cell_properties.as_ref()
    }
}

/// Builder for creating a new paragraph style.
pub struct StyleBuilder {
    style: CT_Style,
    cleared: u16,
    removed_regions: Vec<TableStyleRegion>,
}

pub(crate) const CLEAR_BASED_ON: u16 = 1 << 0;
pub(crate) const CLEAR_NEXT_STYLE: u16 = 1 << 1;
pub(crate) const CLEAR_LINKED_STYLE: u16 = 1 << 2;
pub(crate) const CLEAR_PRIORITY: u16 = 1 << 3;
pub(crate) const CLEAR_AUTO_REDEFINE: u16 = 1 << 4;
pub(crate) const CLEAR_HIDDEN: u16 = 1 << 5;
pub(crate) const CLEAR_SEMI_HIDDEN: u16 = 1 << 6;
pub(crate) const CLEAR_UNHIDE_WHEN_USED: u16 = 1 << 7;
pub(crate) const CLEAR_QUICK_FORMAT: u16 = 1 << 8;
pub(crate) const CLEAR_LOCKED: u16 = 1 << 9;
pub(crate) const CLEAR_PARAGRAPH_PROPERTIES: u16 = 1 << 10;
pub(crate) const CLEAR_RUN_PROPERTIES: u16 = 1 << 11;
pub(crate) const CLEAR_TABLE_PROPERTIES: u16 = 1 << 12;
pub(crate) const CLEAR_CONDITIONAL_TABLE_STYLES: u16 = 1 << 13;
pub(crate) const CLEAR_TABLE_ROW_PROPERTIES: u16 = 1 << 14;
pub(crate) const CLEAR_TABLE_CELL_PROPERTIES: u16 = 1 << 15;

impl StyleBuilder {
    /// Create a new paragraph style builder.
    pub fn paragraph(style_id: &str, name: &str) -> Self {
        StyleBuilder {
            style: CT_Style {
                style_id: style_id.to_string(),
                style_type: StyleType::Paragraph,
                name: Some(name.to_string()),
                based_on: None,
                next_style: None,
                linked_style: None,
                auto_redefine: None,
                hidden: None,
                ui_priority: None,
                semi_hidden: None,
                unhide_when_used: None,
                quick_format: None,
                locked: None,
                is_default: false,
                ppr: None,
                rpr: None,
                table_properties: None,
                table_properties_original: None,
                table_properties_xml: None,
                table_row_properties: None,
                table_cell_properties: None,
                conditional_table_styles: Vec::new(),
                extra_attributes: Vec::new(),
                modeled_xml: Vec::new(),
                extra_xml: Vec::new(),
            },
            cleared: 0,
            removed_regions: Vec::new(),
        }
    }

    /// Create a new character style builder.
    pub fn character(style_id: &str, name: &str) -> Self {
        StyleBuilder {
            style: CT_Style {
                style_id: style_id.to_string(),
                style_type: StyleType::Character,
                name: Some(name.to_string()),
                based_on: None,
                next_style: None,
                linked_style: None,
                auto_redefine: None,
                hidden: None,
                ui_priority: None,
                semi_hidden: None,
                unhide_when_used: None,
                quick_format: None,
                locked: None,
                is_default: false,
                ppr: None,
                rpr: None,
                table_properties: None,
                table_properties_original: None,
                table_properties_xml: None,
                table_row_properties: None,
                table_cell_properties: None,
                conditional_table_styles: Vec::new(),
                extra_attributes: Vec::new(),
                modeled_xml: Vec::new(),
                extra_xml: Vec::new(),
            },
            cleared: 0,
            removed_regions: Vec::new(),
        }
    }

    /// Create a new table style builder.
    pub fn table(style_id: &str, name: &str) -> Self {
        let mut builder = Self::paragraph(style_id, name);
        builder.style.style_type = StyleType::Table;
        builder
    }

    /// Set the parent style this one inherits from.
    pub fn based_on(mut self, style_id: &str) -> Self {
        self.style.based_on = Some(style_id.to_string());
        self.cleared &= !CLEAR_BASED_ON;
        self
    }

    /// Remove the parent style during an update.
    pub fn clear_based_on(mut self) -> Self {
        self.style.based_on = None;
        self.cleared |= CLEAR_BASED_ON;
        self
    }

    /// Set the next style (applied to the following paragraph after pressing Enter).
    pub fn next_style(mut self, style_id: &str) -> Self {
        self.style.next_style = Some(style_id.to_string());
        self.cleared &= !CLEAR_NEXT_STYLE;
        self
    }

    /// Remove the following-paragraph style during an update.
    pub fn clear_next_style(mut self) -> Self {
        self.style.next_style = None;
        self.cleared |= CLEAR_NEXT_STYLE;
        self
    }

    /// Link this paragraph style to a character style, or the reverse.
    pub fn linked_style(mut self, style_id: &str) -> Self {
        self.style.linked_style = Some(style_id.to_string());
        self.cleared &= !CLEAR_LINKED_STYLE;
        self
    }

    /// Remove the reciprocal paragraph or character link during an update.
    pub fn clear_linked_style(mut self) -> Self {
        self.style.linked_style = None;
        self.cleared |= CLEAR_LINKED_STYLE;
        self
    }

    /// Set the style's user-interface ordering priority.
    pub fn priority(mut self, priority: u32) -> Self {
        self.style.ui_priority = Some(priority);
        self.cleared &= !CLEAR_PRIORITY;
        self
    }

    /// Remove the user-interface ordering priority during an update.
    pub fn clear_priority(mut self) -> Self {
        self.style.ui_priority = None;
        self.cleared |= CLEAR_PRIORITY;
        self
    }

    /// Set whether Word may redefine the style from direct formatting.
    pub fn auto_redefine(mut self, value: bool) -> Self {
        self.style.auto_redefine = Some(value);
        self.cleared &= !CLEAR_AUTO_REDEFINE;
        self
    }

    /// Remove the automatic-redefinition setting during an update.
    pub fn clear_auto_redefine(mut self) -> Self {
        self.style.auto_redefine = None;
        self.cleared |= CLEAR_AUTO_REDEFINE;
        self
    }

    /// Set whether the style is hidden from all user-interface lists.
    pub fn hidden(mut self, value: bool) -> Self {
        self.style.hidden = Some(value);
        self.cleared &= !CLEAR_HIDDEN;
        self
    }

    /// Remove the hidden setting during an update.
    pub fn clear_hidden(mut self) -> Self {
        self.style.hidden = None;
        self.cleared |= CLEAR_HIDDEN;
        self
    }

    /// Set whether the style is hidden from the main style gallery.
    pub fn semi_hidden(mut self, value: bool) -> Self {
        self.style.semi_hidden = Some(value);
        self.cleared &= !CLEAR_SEMI_HIDDEN;
        self
    }

    /// Remove the semi-hidden setting during an update.
    pub fn clear_semi_hidden(mut self) -> Self {
        self.style.semi_hidden = None;
        self.cleared |= CLEAR_SEMI_HIDDEN;
        self
    }

    /// Set whether applying the style makes it visible in the user interface.
    pub fn unhide_when_used(mut self, value: bool) -> Self {
        self.style.unhide_when_used = Some(value);
        self.cleared &= !CLEAR_UNHIDE_WHEN_USED;
        self
    }

    /// Remove the unhide-when-used setting during an update.
    pub fn clear_unhide_when_used(mut self) -> Self {
        self.style.unhide_when_used = None;
        self.cleared |= CLEAR_UNHIDE_WHEN_USED;
        self
    }

    /// Set whether the style appears in the quick style gallery.
    pub fn quick_format(mut self, value: bool) -> Self {
        self.style.quick_format = Some(value);
        self.cleared &= !CLEAR_QUICK_FORMAT;
        self
    }

    /// Remove the quick-format setting during an update.
    pub fn clear_quick_format(mut self) -> Self {
        self.style.quick_format = None;
        self.cleared |= CLEAR_QUICK_FORMAT;
        self
    }

    /// Set whether the style is locked against application in the user interface.
    pub fn locked(mut self, value: bool) -> Self {
        self.style.locked = Some(value);
        self.cleared &= !CLEAR_LOCKED;
        self
    }

    /// Remove the locked setting during an update.
    pub fn clear_locked(mut self) -> Self {
        self.style.locked = None;
        self.cleared |= CLEAR_LOCKED;
        self
    }

    /// Set paragraph properties for this style.
    pub fn paragraph_properties(mut self, ppr: CT_PPr) -> Self {
        self.style.ppr = Some(ppr);
        self.cleared &= !CLEAR_PARAGRAPH_PROPERTIES;
        self
    }

    /// Remove all paragraph properties during an update.
    pub fn clear_paragraph_properties(mut self) -> Self {
        self.style.ppr = None;
        self.cleared |= CLEAR_PARAGRAPH_PROPERTIES;
        self
    }

    /// Set run properties for this style.
    pub fn run_properties(mut self, rpr: CT_RPr) -> Self {
        self.style.rpr = Some(rpr);
        self.cleared &= !CLEAR_RUN_PROPERTIES;
        self
    }

    /// Remove all run properties during an update.
    pub fn clear_run_properties(mut self) -> Self {
        self.style.rpr = None;
        self.cleared |= CLEAR_RUN_PROPERTIES;
        self
    }

    /// Set base table properties for a table style.
    pub fn table_properties(mut self, properties: CT_TblPr) -> Self {
        self.style.table_properties = Some(properties);
        self.cleared &= !CLEAR_TABLE_PROPERTIES;
        self
    }

    /// Remove all base table properties during an update.
    pub fn clear_table_properties(mut self) -> Self {
        self.style.table_properties = None;
        self.cleared |= CLEAR_TABLE_PROPERTIES;
        self
    }

    /// Set base table row properties for a table style.
    pub fn table_row_properties(mut self, properties: CT_TrPr) -> Self {
        self.style.table_row_properties = Some(Box::new(properties));
        self.cleared &= !CLEAR_TABLE_ROW_PROPERTIES;
        self
    }

    /// Remove all base table row properties during an update.
    pub fn clear_table_row_properties(mut self) -> Self {
        self.style.table_row_properties = None;
        self.cleared |= CLEAR_TABLE_ROW_PROPERTIES;
        self
    }

    /// Set base table cell properties for a table style.
    pub fn table_cell_properties(mut self, properties: CT_TcPr) -> Self {
        self.style.table_cell_properties = Some(Box::new(properties));
        self.cleared &= !CLEAR_TABLE_CELL_PROPERTIES;
        self
    }

    /// Remove all base table cell properties during an update.
    pub fn clear_table_cell_properties(mut self) -> Self {
        self.style.table_cell_properties = None;
        self.cleared |= CLEAR_TABLE_CELL_PROPERTIES;
        self
    }

    /// Remove all conditional table regions during an update.
    pub fn clear_conditional_table_styles(mut self) -> Self {
        self.style.conditional_table_styles.clear();
        self.cleared |= CLEAR_CONDITIONAL_TABLE_STYLES;
        self
    }

    /// Remove one conditional table region during an update.
    ///
    /// Every other region survives, which is what separates this from
    /// [`StyleBuilder::clear_conditional_table_styles`].
    pub fn remove_conditional_table_style(mut self, region: TableStyleRegion) -> Self {
        self.style
            .conditional_table_styles
            .retain(|conditional| conditional.region != Some(region));
        if !self.removed_regions.contains(&region) {
            self.removed_regions.push(region);
        }
        self
    }

    /// Add one conditional table style region in source order.
    ///
    /// The row layer is modeled and round-tripped. Its layout application
    /// belongs to F-268a.
    pub fn conditional_table_style(
        mut self,
        region: TableStyleRegion,
        paragraph_properties: Option<CT_PPr>,
        run_properties: Option<CT_RPr>,
        table_properties: Option<CT_TblPr>,
        row_properties: Option<CT_TrPr>,
        cell_properties: Option<CT_TcPr>,
    ) -> Self {
        self.removed_regions.retain(|removed| *removed != region);
        self.style.conditional_table_styles.push(CT_TblStylePr {
            region: Some(region),
            paragraph_properties,
            run_properties: run_properties.map(Box::new),
            table_properties,
            row_properties: row_properties.map(Box::new),
            cell_properties,
            extra_attributes: Vec::new(),
            raw_xml: Vec::new(),
        });
        self
    }

    /// Build the style (consumed by Document::add_style).
    pub(crate) fn build(self) -> (CT_Style, u16, Vec<TableStyleRegion>) {
        (self.style, self.cleared, self.removed_regions)
    }
}

pub(crate) fn validate_style_graph(styles: &CT_Styles) -> Result<()> {
    match style_graph_defects(styles).into_iter().next() {
        Some(defect) => Err(style_graph_error(defect)),
        None => Ok(()),
    }
}

/// Validate a staged change to a style graph a producer may already have
/// left invalid.
///
/// Each defect `staged` shares with `source` is returned once, in check order,
/// for the caller to report. A defect only `staged` has was introduced by the
/// change and rejects it.
pub(crate) fn validate_style_graph_change(
    source: &CT_Styles,
    staged: &CT_Styles,
) -> Result<Vec<String>> {
    let source_defects = style_graph_defects(source);
    let mut retained = Vec::new();
    for defect in style_graph_defects(staged) {
        if !source_defects.contains(&defect) {
            return Err(style_graph_error(defect));
        }
        if !retained.contains(&defect) {
            retained.push(defect);
        }
    }
    Ok(retained)
}

/// Every defect in check order. The first is the one `validate_style_graph`
/// reports.
fn style_graph_defects(styles: &CT_Styles) -> Vec<String> {
    let mut defects = Vec::new();
    let mut by_id = HashMap::new();
    for style in &styles.styles {
        if style.style_id.is_empty() {
            defects.push("style IDs cannot be empty".to_owned());
        } else if by_id.contains_key(style.style_id.as_str()) {
            defects.push(format!("duplicate style ID '{}'", style.style_id));
        } else {
            by_id.insert(style.style_id.as_str(), style);
        }
    }

    for style_type in [
        StyleType::Paragraph,
        StyleType::Character,
        StyleType::Table,
        StyleType::Numbering,
    ] {
        let defaults = styles
            .styles
            .iter()
            .filter(|style| style.style_type == style_type && style.is_default)
            .count();
        if defaults > 1 {
            defects.push(format!(
                "style type '{}' has more than one default",
                style_type.to_str()
            ));
        }
    }

    for style in &styles.styles {
        if style.style_type == StyleType::Character && style.ppr.is_some() {
            defects.push(format!(
                "character style '{}' cannot contain paragraph properties",
                style.style_id
            ));
        }
        if style.style_type != StyleType::Table
            && (style.table_properties.is_some()
                || style.table_row_properties.is_some()
                || style.table_cell_properties.is_some()
                || !style.conditional_table_styles.is_empty())
        {
            defects.push(format!(
                "{} style '{}' cannot contain table properties",
                style.style_type.to_str(),
                style.style_id
            ));
        }
        // Region validity is unrepresentable rather than checked. An
        // unrecognised `w:type` parses as `None`, round-trips from its
        // preserved bytes and is never a reason to refuse a file Word wrote,
        // so only recognised regions take part in the duplicate check.
        let mut conditional_regions = HashSet::new();
        for region in style
            .conditional_table_styles
            .iter()
            .filter_map(|conditional| conditional.region)
        {
            if !conditional_regions.insert(region) {
                defects.push(format!(
                    "table style '{}' repeats conditional region '{}'",
                    style.style_id,
                    region.to_str()
                ));
            }
        }

        if let Some(parent_id) = style.based_on.as_deref() {
            match by_id.get(parent_id) {
                None => defects.push(format!(
                    "style '{}' is based on missing style '{parent_id}'",
                    style.style_id
                )),
                Some(parent) if parent.style_type != style.style_type => {
                    defects.push(format!(
                        "style '{}' cannot be based on {} style '{parent_id}'",
                        style.style_id,
                        parent.style_type.to_str()
                    ));
                }
                Some(_) => {}
            }
        }

        if let Some(next_id) = style.next_style.as_deref() {
            if style.style_type != StyleType::Paragraph {
                defects.push(format!(
                    "{} style '{}' cannot declare a next style",
                    style.style_type.to_str(),
                    style.style_id
                ));
            } else {
                match by_id.get(next_id) {
                    None => defects.push(format!(
                        "style '{}' names missing next style '{next_id}'",
                        style.style_id
                    )),
                    Some(next) if next.style_type != StyleType::Paragraph => {
                        defects.push(format!(
                            "paragraph style '{}' has non-paragraph next style '{next_id}'",
                            style.style_id
                        ));
                    }
                    Some(_) => {}
                }
            }
        }

        if let Some(linked_id) = style.linked_style.as_deref() {
            match by_id.get(linked_id) {
                None => defects.push(format!(
                    "style '{}' links to missing style '{linked_id}'",
                    style.style_id
                )),
                Some(linked) => {
                    let legal_types = matches!(
                        (style.style_type, linked.style_type),
                        (StyleType::Paragraph, StyleType::Character)
                            | (StyleType::Character, StyleType::Paragraph)
                    );
                    if !legal_types {
                        defects.push(format!(
                            "{} style '{}' cannot link to {} style '{linked_id}'",
                            style.style_type.to_str(),
                            style.style_id,
                            linked.style_type.to_str()
                        ));
                    } else if linked.linked_style.as_deref() != Some(style.style_id.as_str()) {
                        defects.push(format!(
                            "linked styles '{}' and '{linked_id}' are not reciprocal",
                            style.style_id
                        ));
                    }
                }
            }
        }
    }

    for style in &styles.styles {
        let mut seen = HashSet::new();
        let mut current = Some(style.style_id.as_str());
        while let Some(style_id) = current {
            if !seen.insert(style_id) {
                defects.push(format!("based-on cycle contains style '{style_id}'"));
                break;
            }
            current = by_id
                .get(style_id)
                .and_then(|current_style| current_style.based_on.as_deref());
        }
    }

    defects
}

fn style_graph_error(message: impl Into<String>) -> Error {
    Error::Other(format!("invalid style graph: {}", message.into()))
}

/// Resolve the effective paragraph properties by walking the style inheritance chain.
pub fn resolve_paragraph_properties(style_id: Option<&str>, styles: &CT_Styles) -> CT_PPr {
    let mut effective = CT_PPr::default();

    // Start from docDefaults
    if let Some(ref defaults) = styles.doc_defaults
        && let Some(ref ppr) = defaults.ppr
    {
        effective.merge_from(ppr);
    }

    // Walk the selected style's basedOn chain.
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
                effective.merge_from(ppr);
            }
        }
    }

    effective
}

/// Resolve the effective run properties by walking the style inheritance chain.
pub fn resolve_run_properties(
    para_style_id: Option<&str>,
    run_style_id: Option<&str>,
    styles: &CT_Styles,
) -> CT_RPr {
    let mut effective = CT_RPr::default();

    // Start from docDefaults
    if let Some(ref defaults) = styles.doc_defaults
        && let Some(ref rpr) = defaults.rpr
    {
        effective.merge_from(rpr);
    }

    // Apply paragraph style's rpr
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

    // Apply character style's rpr
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

/// Collect the chain of styles from the given style up through basedOn ancestors.
fn collect_style_chain<'a>(style_id: &str, styles: &'a CT_Styles) -> Vec<&'a CT_Style> {
    let mut chain = Vec::new();
    let mut current_id = Some(style_id.to_string());
    let mut seen = std::collections::HashSet::new();

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

        // Add a Heading2 based on Heading1
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
                space_before: Some(Twips(40)), // Override Heading1's 240
                ..Default::default()
            }),
            rpr: Some(CT_RPr {
                sz: Some(HalfPoint(26)), // Override Heading1's 32
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
    fn resolve_normal_paragraph() {
        let styles = test_styles();
        let ppr = resolve_paragraph_properties(Some("Normal"), &styles);
        // Should have docDefaults' spacing
        assert_eq!(ppr.space_after, Some(Twips(160)));
    }

    #[test]
    fn resolve_heading1() {
        let styles = test_styles();
        let ppr = resolve_paragraph_properties(Some("Heading1"), &styles);
        // keepNext from Heading1
        assert_eq!(ppr.keep_next, Some(true));
        // spaceBefore from Heading1 (overrides docDefaults which has none)
        assert_eq!(ppr.space_before, Some(Twips(240)));
        // spaceAfter from Heading1 overrides docDefaults
        assert_eq!(ppr.space_after, Some(Twips(0)));
    }

    #[test]
    fn resolve_heading2_inherits_heading1() {
        let styles = test_styles();
        let ppr = resolve_paragraph_properties(Some("Heading2"), &styles);
        // keepNext inherited from Heading1
        assert_eq!(ppr.keep_next, Some(true));
        // spaceBefore overridden by Heading2
        assert_eq!(ppr.space_before, Some(Twips(40)));
    }

    #[test]
    fn resolve_heading2_rpr() {
        let styles = test_styles();
        let rpr = resolve_run_properties(Some("Heading2"), None, &styles);
        // Font from docDefaults
        assert_eq!(rpr.font_ascii, Some("Calibri".to_string()));
        // Size overridden by Heading2 (not Heading1's 32)
        assert_eq!(rpr.sz, Some(HalfPoint(26)));
        // Bold inherited from Heading1
        assert_eq!(rpr.bold, Some(true));
        // Color from Heading2
        assert_eq!(rpr.color, Some("2E74B5".to_string()));
    }

    #[test]
    fn new_paragraph_properties_inherit_through_the_style_chain() {
        use rdocx_oxml::properties::CT_FramePr;

        let mut styles = test_styles();
        let defaults = styles
            .doc_defaults
            .as_mut()
            .expect("docDefaults")
            .ppr
            .get_or_insert_with(CT_PPr::default);
        defaults.suppress_line_numbers = Some(true);
        defaults.text_alignment = Some("baseline".to_owned());
        for style in &mut styles.styles {
            let ppr = style.ppr.get_or_insert_with(CT_PPr::default);
            if style.style_id == "Heading1" {
                ppr.contextual_spacing = Some(true);
                ppr.text_direction = Some("tbRlV".to_owned());
                ppr.text_alignment = Some("center".to_owned());
                ppr.frame = Some(Box::new(CT_FramePr {
                    w: Some(Twips(2880)),
                    ..Default::default()
                }));
                ppr.div_id = Some(4);
            }
            if style.style_id == "Heading2" {
                ppr.mirror_indents = Some(true);
                ppr.text_direction = Some("lrTb".to_owned());
            }
        }

        let ppr = resolve_paragraph_properties(Some("Heading2"), &styles);
        assert_eq!(ppr.suppress_line_numbers, Some(true));
        assert_eq!(ppr.contextual_spacing, Some(true));
        assert_eq!(ppr.mirror_indents, Some(true));
        assert_eq!(ppr.text_alignment.as_deref(), Some("center"));
        assert_eq!(ppr.text_direction.as_deref(), Some("lrTb"));
        assert_eq!(ppr.frame.and_then(|frame| frame.w), Some(Twips(2880)));
        assert_eq!(ppr.div_id, Some(4));
    }

    #[test]
    fn resolve_default_when_no_style() {
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
    fn style_graph_change_retains_producer_defects_and_rejects_introduced_ones() {
        let mut source = test_styles();
        for (builder, is_default) in [
            (
                StyleBuilder::paragraph("Orphan", "Orphan").based_on("Missing"),
                false,
            ),
            (StyleBuilder::paragraph("SecondDefault", "Second"), true),
            (
                StyleBuilder::character("OneWay", "One Way").linked_style("Normal"),
                false,
            ),
        ] {
            let (mut style, _, _) = builder.build();
            style.is_default = is_default;
            source.styles.push(style);
        }
        // Strict validation still fails on the first defect in check order.
        assert_eq!(
            validate_style_graph(&source).unwrap_err().to_string(),
            "invalid style graph: style type 'paragraph' has more than one default"
        );
        assert_eq!(
            validate_style_graph_change(&source, &source).unwrap(),
            [
                "style type 'paragraph' has more than one default",
                "style 'Orphan' is based on missing style 'Missing'",
                "linked styles 'OneWay' and 'Normal' are not reciprocal",
            ]
        );

        // Resolving the dangling parent closes a cycle the source did not have.
        let mut staged = source.clone();
        let (parent, _, _) = StyleBuilder::paragraph("Missing", "Missing")
            .based_on("Orphan")
            .build();
        staged.styles.push(parent);
        assert_eq!(
            validate_style_graph_change(&source, &staged)
                .unwrap_err()
                .to_string(),
            "invalid style graph: based-on cycle contains style 'Orphan'"
        );
    }
}
